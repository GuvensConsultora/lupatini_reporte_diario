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

INGRESOS_TIPOS = ('cash', 'card', 'mp')


class ReporteDiarioWizard(models.TransientModel):
    _name = 'lupatini.reporte.diario.wizard'
    _description = 'Reporte Diario de Ingresos'

    date = fields.Date(
        string='Fecha',
        required=True,
        default=fields.Date.today,
    )
    company_id = fields.Many2one(
        'res.company',
        string='Empresa',
        required=True,
        default=lambda self: self.env.company,
    )

    # -------------------------------------------------------------------------
    # Acción principal
    # -------------------------------------------------------------------------

    def action_export_excel(self):
        self.ensure_one()
        if not xlsxwriter:
            raise UserError(_('Se requiere la librería xlsxwriter para exportar a Excel.'))

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        self._write_excel(workbook)
        workbook.close()

        filename = 'reporte_diario_%s.xlsx' % self.date.strftime('%Y%m%d')
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
    # Consultas SQL
    # Por qué SQL directo: evitar N+1 queries para aggregados por OU.
    # Por qué join a account_move: en Odoo 17 journal_id no es columna directa
    # en account_payment, vive en el account_move relacionado.
    # -------------------------------------------------------------------------

    def _params(self):
        return {'date': self.date, 'cid': self.company_id.id}

    def _get_ventas_por_ou(self):
        """Facturas de cliente agrupadas por OU, sin IVA."""
        self.env.cr.execute("""
            SELECT
                COALESCE(ou.name, '(Sin sucursal)') AS ou_name,
                COALESCE(SUM(am.amount_untaxed), 0.0)  AS total
            FROM account_move am
            LEFT JOIN operating_unit ou ON ou.id = am.operating_unit_id
            WHERE am.move_type = 'out_invoice'
              AND am.state      = 'posted'
              AND am.invoice_date = %(date)s
              AND am.company_id   = %(cid)s
            GROUP BY ou.id, ou.name
            ORDER BY ou.name NULLS LAST
        """, self._params())
        return {r['ou_name']: r['total'] for r in self.env.cr.dictfetchall()}

    def _get_ingresos_por_ou(self):
        """Un solo query para cash/card/mp agrupado por tipo y OU.
        Devuelve dict[tipo][ou_name] = total."""
        self.env.cr.execute("""
            SELECT
                aj.lupatini_ingreso_tipo               AS tipo,
                COALESCE(ou.name, '(Sin sucursal)')    AS ou_name,
                COALESCE(SUM(ap.amount), 0.0)          AS total
            FROM account_payment ap
            JOIN account_move    mv ON mv.id  = ap.move_id
            JOIN account_journal aj ON aj.id  = mv.journal_id
            LEFT JOIN operating_unit ou ON ou.id = ap.operating_unit_id
            WHERE ap.payment_type          = 'inbound'
              AND mv.state                 = 'posted'
              AND mv.date                  = %(date)s
              AND mv.company_id            = %(cid)s
              AND aj.lupatini_ingreso_tipo IN %(tipos)s
            GROUP BY aj.lupatini_ingreso_tipo, ou.id, ou.name
            ORDER BY ou.name NULLS LAST
        """, {**self._params(), 'tipos': INGRESOS_TIPOS})
        result = {t: {} for t in INGRESOS_TIPOS}
        for r in self.env.cr.dictfetchall():
            result[r['tipo']][r['ou_name']] = r['total']
        return result

    def _get_cheques(self):
        """Cheques de terceros recibidos en el día (OCA account_check)."""
        self.env.cr.execute("""
            SELECT
                rp.name                                  AS cliente,
                COALESCE(ou.name, '(Sin sucursal)')      AS sucursal,
                ac.number                                AS nro_valor,
                mv.date                                  AS fecha_ingreso,
                ac.payment_date                          AS vencimiento,
                ap.amount                                AS importe
            FROM account_payment ap
            JOIN account_move     mv ON mv.id = ap.move_id
            JOIN account_check    ac ON ac.id = ap.check_id
            LEFT JOIN res_partner    rp ON rp.id = ap.partner_id
            LEFT JOIN operating_unit ou ON ou.id = ap.operating_unit_id
            WHERE ap.payment_type = 'inbound'
              AND mv.state        = 'posted'
              AND mv.date         = %(date)s
              AND mv.company_id   = %(cid)s
              AND ac.type         = 'third_check'
            ORDER BY ou.name NULLS LAST, ap.id
        """, self._params())
        return self.env.cr.dictfetchall()

    def _get_transferencias(self):
        """Transferencias bancarias agrupadas por banco."""
        self.env.cr.execute("""
            SELECT
                aj.name                           AS banco,
                COALESCE(SUM(ap.amount), 0.0)     AS total
            FROM account_payment ap
            JOIN account_move    mv ON mv.id = ap.move_id
            JOIN account_journal aj ON aj.id = mv.journal_id
            WHERE ap.payment_type          = 'inbound'
              AND mv.state                 = 'posted'
              AND mv.date                  = %(date)s
              AND mv.company_id            = %(cid)s
              AND aj.lupatini_ingreso_tipo = 'bank'
            GROUP BY aj.id, aj.name
            ORDER BY aj.name
        """, self._params())
        return self.env.cr.dictfetchall()

    # -------------------------------------------------------------------------
    # Excel
    # -------------------------------------------------------------------------

    def _build_formats(self, wb):
        return {
            'title': wb.add_format({
                'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter',
            }),
            'label':    wb.add_format({'bold': True}),
            'section':  wb.add_format({
                'bold': True, 'bg_color': '#1F6B75', 'font_color': 'white',
                'border': 1, 'font_size': 11,
            }),
            'header':   wb.add_format({
                'bold': True, 'bg_color': '#2E8B8B', 'font_color': 'white',
                'border': 1, 'align': 'center', 'text_wrap': True,
            }),
            'text':      wb.add_format({'border': 1}),
            'money':     wb.add_format({'num_format': '#,##0.00', 'border': 1}),
            'date':      wb.add_format({'num_format': 'dd/mm/yyyy', 'border': 1}),
            'total_lbl': wb.add_format({'bold': True, 'border': 1, 'bg_color': '#D6EAF8'}),
            'total':     wb.add_format({
                'bold': True, 'num_format': '#,##0.00', 'border': 1, 'bg_color': '#D6EAF8',
            }),
        }

    def _write_excel(self, wb):
        ws = wb.add_worksheet('Reporte Diario')
        fmt = self._build_formats(wb)

        ws.set_column(0, 0, 32)
        ws.set_column(1, 5, 18)

        # Encabezado
        ws.merge_range('A1:F1', 'REPORTE DIARIO DE INGRESOS', fmt['title'])
        ws.write('A2', 'Fecha:',   fmt['label'])
        ws.write('B2', self.date.strftime('%d/%m/%Y'))
        ws.write('A3', 'Empresa:', fmt['label'])
        ws.write('B3', self.company_id.name)

        row = 4

        # ── Sección 1: Ventas ─────────────────────────────────────────────
        ws.merge_range(row, 0, row, 4, 'Ventas', fmt['section'])
        row += 1
        for col, h in enumerate(['Sucursal', 'Ventas Sin IVA', 'Efectivo', 'Tarjetas', 'Mercado Pago']):
            ws.write(row, col, h, fmt['header'])
        row += 1

        ventas   = self._get_ventas_por_ou()
        ingresos = self._get_ingresos_por_ou()
        efectivo, tarjetas, mp = ingresos['cash'], ingresos['card'], ingresos['mp']

        all_ous = sorted(ventas.keys() | efectivo.keys() | tarjetas.keys() | mp.keys())
        totales = defaultdict(float)

        for ou in all_ous:
            cols = [ventas.get(ou, 0.0), efectivo.get(ou, 0.0),
                    tarjetas.get(ou, 0.0), mp.get(ou, 0.0)]
            ws.write(row, 0, ou, fmt['text'])
            for col, val in enumerate(cols, start=1):
                ws.write_number(row, col, val, fmt['money'])
                totales[col] += val
            row += 1

        ws.write(row, 0, 'TOTAL', fmt['total_lbl'])
        for col in range(1, 5):
            ws.write_number(row, col, totales[col], fmt['total'])
        row += 2

        # ── Sección 2: Cheques ────────────────────────────────────────────
        ws.merge_range(row, 0, row, 5, 'Cheques', fmt['section'])
        row += 1
        for col, h in enumerate(['Cliente', 'Sucursal', 'N° Valor', 'Fecha de Ingreso', 'VT', 'Ingresos']):
            ws.write(row, col, h, fmt['header'])
        row += 1

        tot_ch = 0.0
        for ch in self._get_cheques():
            ws.write(row, 0, ch['cliente'] or '', fmt['text'])
            ws.write(row, 1, ch['sucursal'] or '', fmt['text'])
            ws.write(row, 2, ch['nro_valor'] or '', fmt['text'])
            ws.write_datetime(row, 3, ch['fecha_ingreso'], fmt['date']) if ch['fecha_ingreso'] else ws.write(row, 3, '', fmt['text'])
            ws.write_datetime(row, 4, ch['vencimiento'], fmt['date'])   if ch['vencimiento']    else ws.write(row, 4, '', fmt['text'])
            importe = ch['importe'] or 0.0
            ws.write_number(row, 5, importe, fmt['money'])
            tot_ch += importe
            row += 1

        ws.write(row, 4, 'TOTAL', fmt['total_lbl'])
        ws.write_number(row, 5, tot_ch, fmt['total'])
        row += 2

        # ── Sección 3: Transferencias ─────────────────────────────────────
        ws.merge_range(row, 0, row, 1, 'Transferencias', fmt['section'])
        row += 1
        ws.write(row, 0, 'Banco', fmt['header'])
        ws.write(row, 1, 'Total', fmt['header'])
        row += 1

        tot_tr = 0.0
        for tr in self._get_transferencias():
            ws.write(row, 0, tr['banco'], fmt['text'])
            ws.write_number(row, 1, tr['total'], fmt['money'])
            tot_tr += tr['total']
            row += 1

        ws.write(row, 0, 'TOTAL', fmt['total_lbl'])
        ws.write_number(row, 1, tot_tr, fmt['total'])
