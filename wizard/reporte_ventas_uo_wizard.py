# -*- coding: utf-8 -*-
import io
import base64
from collections import defaultdict

from odoo import models, fields, _
from odoo.exceptions import UserError

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None

# Orden y etiqueta de los tres grupos del reporte de contado.
GRUPOS = (
    ('contado', 'CONTADO'),
    ('ctacte', 'CUENTA CORRIENTE'),
    ('sin', 'SIN CLASIFICAR'),
)


def _primer_dia_mes(self):
    return fields.Date.context_today(self).replace(day=1)


class ReporteVentasUoWizard(models.TransientModel):
    _name = 'lupatini.reporte.ventas.uo.wizard'
    _description = 'Reporte de Ventas por Unidad Operativa'

    # Rango libre entre fechas. Default: del primer día del mes en curso a hoy
    # (cubre el mes); el usuario lo acota a un día o a una semana a mano.
    date_from = fields.Date(
        string='Desde',
        required=True,
        default=_primer_dia_mes,
    )
    date_to = fields.Date(
        string='Hasta',
        required=True,
        default=fields.Date.context_today,
    )
    company_id = fields.Many2one(
        'res.company',
        string='Empresa',
        required=True,
        default=lambda self: self.env.company,
    )

    # -------------------------------------------------------------------------
    # Motor: venta neta sin IVA por Unidad Operativa en el rango
    # Reusa el criterio del reporte diario (_get_ventas_por_ou): FA + ND − NC,
    # generalizado a un rango de fechas (BETWEEN sobre invoice_date).
    # SQL directo para evitar N+1 al agregar por OU.
    # -------------------------------------------------------------------------
    def _ventas_por_ou(self, date_from, date_to):
        self.env.cr.execute("""
            SELECT
                COALESCE(ou.name, '(Sin sucursal)') AS ou_name,
                COALESCE(SUM(
                    CASE WHEN am.move_type = 'out_refund'
                         THEN -am.amount_untaxed
                         ELSE  am.amount_untaxed
                    END
                ), 0.0) AS total
            FROM account_move am
            LEFT JOIN operating_unit ou ON ou.id = am.operating_unit_id
            WHERE am.move_type   IN ('out_invoice', 'out_refund')
              AND am.state        = 'posted'
              AND am.invoice_date BETWEEN %(df)s AND %(dt)s
              AND am.company_id   = %(cid)s
            GROUP BY ou.id, ou.name
            ORDER BY ou.name NULLS LAST
        """, {'df': date_from, 'dt': date_to, 'cid': self.company_id.id})
        return self.env.cr.dictfetchall()

    # -------------------------------------------------------------------------
    # Acción del botón "Imprimir PDF"
    # -------------------------------------------------------------------------
    def action_print_pdf(self):
        self.ensure_one()
        if self.date_from > self.date_to:
            raise UserError(_('La fecha "Desde" no puede ser posterior a "Hasta".'))
        filas = self._ventas_por_ou(self.date_from, self.date_to)
        total = sum(f['total'] for f in filas)
        data = {
            'date_from': self.date_from.strftime('%d/%m/%Y'),
            'date_to': self.date_to.strftime('%d/%m/%Y'),
            'company_name': self.company_id.name,
            'filas': filas,
            'total': total,
        }
        return self.env.ref(
            'lupatini_reporte_diario.action_report_ventas_uo'
        ).report_action(self, data=data)

    # =========================================================================
    # Contado vs cuenta corriente (Excel)
    # =========================================================================

    # -------------------------------------------------------------------------
    # Clasificación por condición de pago, NO por medio de cobro: una venta de
    # contado cobrada con tarjeta sigue siendo de contado.
    #
    # Las notas de crédito NO heredan la condición de pago de la factura que
    # revierten: quedan con el campo vacío. Si no se resolviera por
    # reversed_entry_id, cada NC restaría del grupo equivocado y "sin clasificar"
    # llegaría a dar negativo. Por eso el doble LEFT JOIN al comprobante original.
    #
    # Los tickets se cuentan sólo sobre las facturas: una NC no es un ticket
    # generado, aunque sí resta del importe.
    # -------------------------------------------------------------------------
    def _ventas_contado(self, date_from, date_to):
        self.env.cr.execute("""
            WITH clasif AS (
                SELECT
                    am.invoice_date                     AS fecha,
                    COALESCE(ou.name, '(Sin sucursal)') AS ou_name,
                    am.move_type                        AS move_type,
                    CASE WHEN am.move_type = 'out_refund' THEN -1 ELSE 1 END AS signo,
                    am.amount_total                     AS amount_total,
                    am.amount_untaxed                   AS amount_untaxed,
                    CASE
                        WHEN COALESCE(pt.lupatini_es_contado,
                                      pto.lupatini_es_contado) IS TRUE
                            THEN 'contado'
                        WHEN COALESCE(am.invoice_payment_term_id,
                                      orig.invoice_payment_term_id) IS NOT NULL
                            THEN 'ctacte'
                        ELSE 'sin'
                    END                                 AS tipo
                FROM account_move am
                LEFT JOIN operating_unit ou
                       ON ou.id = am.operating_unit_id
                LEFT JOIN account_payment_term pt
                       ON pt.id = am.invoice_payment_term_id
                LEFT JOIN account_move orig
                       ON orig.id = am.reversed_entry_id
                LEFT JOIN account_payment_term pto
                       ON pto.id = orig.invoice_payment_term_id
                WHERE am.move_type   IN ('out_invoice', 'out_refund')
                  AND am.state        = 'posted'
                  AND am.invoice_date BETWEEN %(df)s AND %(dt)s
                  AND am.company_id   = %(cid)s
            )
            SELECT
                fecha,
                ou_name,
                tipo,
                COUNT(*) FILTER (WHERE move_type = 'out_invoice') AS tickets,
                COALESCE(SUM(signo * amount_total), 0.0)          AS total,
                COALESCE(SUM(signo * amount_untaxed), 0.0)        AS neto
            FROM clasif
            GROUP BY fecha, ou_name, tipo
            ORDER BY fecha, ou_name
        """, {'df': date_from, 'dt': date_to, 'cid': self.company_id.id})
        return self.env.cr.dictfetchall()

    def _sin_clasificar(self, date_from, date_to):
        """Detalle de los comprobantes que no se pudieron clasificar, para que el
        usuario vea de qué se trata en vez de recibir un número suelto."""
        self.env.cr.execute("""
            SELECT
                am.invoice_date                     AS fecha,
                COALESCE(ou.name, '(Sin sucursal)') AS ou_name,
                am.name                             AS comprobante,
                COALESCE(rp.name, '')               AS cliente,
                COALESCE(aj.name, '')               AS diario,
                CASE WHEN am.move_type = 'out_refund'
                     THEN -am.amount_total ELSE am.amount_total END AS total
            FROM account_move am
            LEFT JOIN operating_unit ou ON ou.id = am.operating_unit_id
            LEFT JOIN res_partner rp    ON rp.id = am.partner_id
            LEFT JOIN account_journal aj ON aj.id = am.journal_id
            LEFT JOIN account_move orig ON orig.id = am.reversed_entry_id
            WHERE am.move_type   IN ('out_invoice', 'out_refund')
              AND am.state        = 'posted'
              AND am.invoice_date BETWEEN %(df)s AND %(dt)s
              AND am.company_id   = %(cid)s
              AND COALESCE(am.invoice_payment_term_id,
                           orig.invoice_payment_term_id) IS NULL
            ORDER BY am.invoice_date, ou_name
        """, {'df': date_from, 'dt': date_to, 'cid': self.company_id.id})
        return self.env.cr.dictfetchall()

    # -------------------------------------------------------------------------
    # Acción del botón "Exportar Excel"
    # -------------------------------------------------------------------------
    def action_export_excel(self):
        self.ensure_one()
        if not xlsxwriter:
            raise UserError(_('Se requiere la librería xlsxwriter para exportar a Excel.'))
        if self.date_from > self.date_to:
            raise UserError(_('La fecha "Desde" no puede ser posterior a "Hasta".'))
        # Sin este dato el reporte saldría con todo en cuenta corriente y sin
        # explicación: se avisa dónde se configura en vez de exportar en vacío.
        if not self.env['account.payment.term'].search_count(
                [('lupatini_es_contado', '=', True)]):
            raise UserError(_(
                'Todavía no hay ninguna condición de pago marcada como venta de contado, '
                'así que el reporte no puede separar contado de cuenta corriente.\n\n'
                'Se marcan en: Contabilidad → Reportes Lupatini → '
                'Configurar Ventas de Contado.'
            ))

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        self._write_excel_contado(workbook)
        workbook.close()

        filename = 'ventas_contado_%s_a_%s.xlsx' % (
            self.date_from.strftime('%Y%m%d'), self.date_to.strftime('%Y%m%d'))
        attachment = self.env['ir.attachment'].create({
            'name': filename,
            'type': 'binary',
            'datas': base64.b64encode(output.getvalue()),
            'mimetype': (
                'application/vnd.openxmlformats-officedocument'
                '.spreadsheetml.sheet'
            ),
        })
        return {
            'type': 'ir.actions.act_url',
            'url': '/web/content/%d?download=true' % attachment.id,
            'target': 'new',
        }

    # -------------------------------------------------------------------------
    # Armado del Excel
    # -------------------------------------------------------------------------
    def _build_formats(self, wb):
        return {
            'title':   wb.add_format({'bold': True, 'font_size': 14,
                                      'font_color': '#1F6B75'}),
            'sub':     wb.add_format({'bold': True, 'font_size': 11}),
            'nota':    wb.add_format({'italic': True, 'font_color': '#555555',
                                      'text_wrap': True, 'valign': 'top'}),
            'grupo':   wb.add_format({'bold': True, 'align': 'center', 'border': 1}),
            'header':  wb.add_format({'bold': True, 'bg_color': '#2E8B8B',
                                      'font_color': 'white', 'border': 1,
                                      'align': 'center', 'text_wrap': True}),
            'text':    wb.add_format({'border': 1}),
            'date':    wb.add_format({'num_format': 'dd/mm/yyyy', 'border': 1}),
            'int':     wb.add_format({'num_format': '#,##0', 'border': 1}),
            'money':   wb.add_format({'num_format': '#,##0.00', 'border': 1}),
            'tot_lbl': wb.add_format({'bold': True, 'border': 1, 'bg_color': '#D6EAF8'}),
            'tot_int': wb.add_format({'bold': True, 'num_format': '#,##0',
                                      'border': 1, 'bg_color': '#D6EAF8'}),
            'tot_mon': wb.add_format({'bold': True, 'num_format': '#,##0.00',
                                      'border': 1, 'bg_color': '#D6EAF8'}),
        }

    def _write_excel_contado(self, wb):
        fmt = self._build_formats(wb)
        filas = self._ventas_contado(self.date_from, self.date_to)

        # celda[(fecha, sucursal)][tipo] = [tickets, total, neto]
        celda = defaultdict(lambda: defaultdict(lambda: [0, 0.0, 0.0]))
        por_suc = defaultdict(lambda: defaultdict(lambda: [0, 0.0, 0.0]))
        por_dia = defaultdict(lambda: defaultdict(lambda: [0, 0.0, 0.0]))
        for f in filas:
            v = (f['tickets'], f['total'], f['neto'])
            for dst, k in ((celda, (f['fecha'], f['ou_name'])),
                           (por_suc, f['ou_name']),
                           (por_dia, f['fecha'])):
                a = dst[k][f['tipo']]
                a[0] += v[0]
                a[1] += v[1]
                a[2] += v[2]

        self._hoja_guia(wb, fmt)
        self._hoja_datos(wb, fmt, 'Por día y sucursal', ['Fecha', 'Sucursal'],
                         [(k, celda[k]) for k in sorted(celda)], doble=True)
        self._hoja_datos(wb, fmt, 'Resumen por sucursal', ['Sucursal'],
                         [(k, por_suc[k]) for k in sorted(por_suc)])
        self._hoja_datos(wb, fmt, 'Resumen por día', ['Fecha'],
                         [(k, por_dia[k]) for k in sorted(por_dia)], fecha=True)
        self._hoja_sin_clasificar(wb, fmt)

    def _hoja_guia(self, wb, fmt):
        ws = wb.add_worksheet('Cómo leerlo')
        ws.set_column(0, 0, 110)
        lineas = [
            ('Ventas de contado y cuenta corriente — %s al %s' % (
                self.date_from.strftime('%d/%m/%Y'),
                self.date_to.strftime('%d/%m/%Y')), 'title'),
            ('', None),
            ('Cómo separamos contado de cuenta corriente', 'sub'),
            ('Por la condición de pago del comprobante, no por el medio con que se cobró. '
             'Una venta de contado pagada con tarjeta sigue siendo de contado.', 'nota'),
            ('   • Contado: las condiciones marcadas en Contabilidad → Reportes Lupatini → '
             'Configurar Ventas de Contado.', 'nota'),
            ('   • Cuenta corriente: el resto de las condiciones.', 'nota'),
            ('   • Sin clasificar: comprobantes sin condición de pago cargada. No los '
             'asignamos a ninguno de los dos: van en su propia columna.', 'nota'),
            ('', None),
            ('Qué es cada número', 'sub'),
            ('   • Tickets: cantidad de facturas emitidas. Las notas de crédito no cuentan '
             'como ticket, pero sí restan del importe.', 'nota'),
            ('   • $: importe facturado con IVA, ya restadas las notas de crédito.', 'nota'),
            ('   • $ total sin IVA: el mismo importe neto de impuestos, para comparar contra '
             'contabilidad.', 'nota'),
        ]
        for i, (txt, estilo) in enumerate(lineas):
            ws.write(i, 0, txt, fmt[estilo] if estilo else None)

    # Formato de cada una de las 9 columnas numéricas, en orden:
    # contado (tickets, $) · cta cte (tickets, $) · sin clasificar (tickets, $)
    # · totales (tickets, $, $ sin IVA)
    _COLS = ('int', 'money', 'int', 'money', 'int', 'money', 'int', 'money', 'money')

    def _hoja_datos(self, wb, fmt, titulo, cabec_izq, datos, doble=False, fecha=False):
        ws = wb.add_worksheet(titulo)
        n = len(cabec_izq)
        ws.set_column(0, 0, 14 if not doble else 12)
        if doble:
            ws.set_column(1, 1, 24)
        elif cabec_izq[0] == 'Sucursal':
            ws.set_column(0, 0, 26)
        ws.set_column(n, n + 8, 15)

        ws.merge_range(0, 0, 0, n + 8, _(
            'Contado son los comprobantes cuya condición de pago está marcada como venta de '
            'contado, sin importar con qué se pagaron. Los tickets son las facturas emitidas; '
            'los importes ya están netos de notas de crédito.'), fmt['nota'])

        # fila de grupos
        r = 2
        c = n
        for _clave, etq in GRUPOS:
            ws.merge_range(r, c, r, c + 1, etq, fmt['grupo'])
            c += 2
        ws.merge_range(r, c, r, c + 2, 'TOTAL', fmt['grupo'])

        r += 1
        cab = list(cabec_izq) + ['Tickets', '$ contado', 'Tickets', '$ cta cte',
                                 'Tickets', '$ sin clasificar',
                                 'Tickets totales', '$ total', '$ total sin IVA']
        for c, h in enumerate(cab):
            ws.write(r, c, h, fmt['header'])

        acum = [0, 0.0, 0.0, 0, 0.0, 0.0, 0, 0.0, 0.0]
        for clave, d in datos:
            r += 1
            if doble:
                ws.write_datetime(r, 0, clave[0], fmt['date'])
                ws.write(r, 1, clave[1], fmt['text'])
            elif fecha:
                ws.write_datetime(r, 0, clave, fmt['date'])
            else:
                ws.write(r, 0, clave, fmt['text'])
            vals = []
            for gclave, _etq in GRUPOS:
                vals += [d[gclave][0], d[gclave][1]]
            vals += [sum(d[g][0] for g, _e in GRUPOS),
                     sum(d[g][1] for g, _e in GRUPOS),
                     sum(d[g][2] for g, _e in GRUPOS)]
            for i, v in enumerate(vals):
                ws.write(r, n + i, v, fmt[self._COLS[i]])
                acum[i] += v

        r += 1
        ws.write(r, 0, 'TOTAL', fmt['tot_lbl'])
        for c in range(1, n):
            ws.write(r, c, '', fmt['tot_lbl'])
        for i, v in enumerate(acum):
            ws.write(r, n + i, v, fmt['tot_int' if self._COLS[i] == 'int' else 'tot_mon'])
        ws.freeze_panes(4, n)

    def _hoja_sin_clasificar(self, wb, fmt):
        ws = wb.add_worksheet('Sin clasificar')
        filas = self._sin_clasificar(self.date_from, self.date_to)
        anchos = (12, 24, 22, 34, 30, 16)
        for c, a in enumerate(anchos):
            ws.set_column(c, c, a)
        ws.merge_range(0, 0, 0, 5, _(
            'Comprobantes sin condición de pago cargada. No entran en contado ni en cuenta '
            'corriente: quedan acá para que se vea de qué se trata.'), fmt['nota'])
        for c, h in enumerate(('Fecha', 'Sucursal', 'Comprobante', 'Cliente',
                               'Diario', '$ con IVA')):
            ws.write(2, c, h, fmt['header'])
        r = 2
        for f in filas:
            r += 1
            ws.write_datetime(r, 0, f['fecha'], fmt['date'])
            ws.write(r, 1, f['ou_name'], fmt['text'])
            ws.write(r, 2, f['comprobante'], fmt['text'])
            ws.write(r, 3, f['cliente'], fmt['text'])
            ws.write(r, 4, f['diario'], fmt['text'])
            ws.write(r, 5, f['total'], fmt['money'])
        r += 1
        ws.write(r, 0, 'TOTAL', fmt['tot_lbl'])
        for c in range(1, 5):
            ws.write(r, c, '', fmt['tot_lbl'])
        ws.write(r, 5, sum(f['total'] for f in filas), fmt['tot_mon'])
        ws.freeze_panes(3, 0)
