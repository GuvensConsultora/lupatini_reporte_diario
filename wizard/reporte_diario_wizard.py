# -*- coding: utf-8 -*-
import io
import base64
from datetime import date as date_type

from odoo import models, fields, _
from odoo.exceptions import UserError

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None


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
            raise UserError(_(
                'Se requiere la librería xlsxwriter para exportar a Excel.'
            ))

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
    # -------------------------------------------------------------------------

    def _get_ventas_por_ou(self):
        """Facturas de cliente (sin NC) agrupadas por OU, sin IVA."""
        self.env.cr.execute("""
            SELECT
                COALESCE(ou.name, '(Sin sucursal)') AS ou_name,
                COALESCE(SUM(am.amount_untaxed), 0.0) AS ventas_sin_iva
            FROM account_move am
            LEFT JOIN operating_unit ou ON ou.id = am.operating_unit_id
            WHERE am.move_type = 'out_invoice'
              AND am.state = 'posted'
              AND am.invoice_date = %(date)s
              AND am.company_id = %(cid)s
            GROUP BY ou.id, ou.name
            ORDER BY ou.name NULLS LAST
        """, {'date': self.date, 'cid': self.company_id.id})
        return {r['ou_name']: r['ventas_sin_iva'] for r in self.env.cr.dictfetchall()}

    def _get_ingresos_por_ou(self, ingreso_tipo):
        """Pagos entrantes por tipo de diario (cash/card/mp), agrupados por OU.
        Por qué filtramos payment_type='inbound': solo cobros, no pagos a proveedores."""
        self.env.cr.execute("""
            SELECT
                COALESCE(ou.name, '(Sin sucursal)') AS ou_name,
                COALESCE(SUM(ap.amount), 0.0) AS total
            FROM account_payment ap
            JOIN account_journal aj ON aj.id = ap.journal_id
            LEFT JOIN operating_unit ou ON ou.id = ap.operating_unit_id
            WHERE ap.state = 'posted'
              AND ap.payment_type = 'inbound'
              AND ap.date = %(date)s
              AND ap.company_id = %(cid)s
              AND aj.lupatini_ingreso_tipo = %(tipo)s
            GROUP BY ou.id, ou.name
            ORDER BY ou.name NULLS LAST
        """, {'date': self.date, 'cid': self.company_id.id, 'tipo': ingreso_tipo})
        return {r['ou_name']: r['total'] for r in self.env.cr.dictfetchall()}

    def _get_cheques(self):
        """Cheques de terceros recibidos en el día.
        Detectados por l10n_latam_use_check = True en el diario."""
        self.env.cr.execute("""
            SELECT
                rp.name AS cliente,
                COALESCE(ou.name, '(Sin sucursal)') AS sucursal,
                ap.l10n_latam_check_number AS nro_valor,
                ap.date AS fecha_ingreso,
                ap.l10n_latam_check_payment_date AS vencimiento,
                ap.amount AS importe
            FROM account_payment ap
            JOIN account_journal aj ON aj.id = ap.journal_id
            LEFT JOIN res_partner rp ON rp.id = ap.partner_id
            LEFT JOIN operating_unit ou ON ou.id = ap.operating_unit_id
            WHERE ap.state = 'posted'
              AND ap.payment_type = 'inbound'
              AND ap.date = %(date)s
              AND ap.company_id = %(cid)s
              AND aj.l10n_latam_use_check = TRUE
            ORDER BY ou.name NULLS LAST, ap.id
        """, {'date': self.date, 'cid': self.company_id.id})
        return self.env.cr.dictfetchall()

    def _get_transferencias(self):
        """Transferencias bancarias agrupadas por banco (diarios tipo 'bank')."""
        self.env.cr.execute("""
            SELECT
                aj.name AS banco,
                COALESCE(SUM(ap.amount), 0.0) AS total
            FROM account_payment ap
            JOIN account_journal aj ON aj.id = ap.journal_id
            WHERE ap.state = 'posted'
              AND ap.payment_type = 'inbound'
              AND ap.date = %(date)s
              AND ap.company_id = %(cid)s
              AND aj.lupatini_ingreso_tipo = 'bank'
            GROUP BY aj.id, aj.name
            ORDER BY aj.name
        """, {'date': self.date, 'cid': self.company_id.id})
        return self.env.cr.dictfetchall()

    # -------------------------------------------------------------------------
    # Excel
    # -------------------------------------------------------------------------

    def _build_formats(self, wb):
        """Devuelve un dict con todos los formatos del workbook."""
        return {
            'title': wb.add_format({
                'bold': True, 'font_size': 14, 'align': 'center',
                'valign': 'vcenter',
            }),
            'label': wb.add_format({'bold': True}),
            'section': wb.add_format({
                'bold': True, 'bg_color': '#1F6B75', 'font_color': 'white',
                'border': 1, 'font_size': 11,
            }),
            'header': wb.add_format({
                'bold': True, 'bg_color': '#2E8B8B', 'font_color': 'white',
                'border': 1, 'align': 'center', 'text_wrap': True,
            }),
            'text': wb.add_format({'border': 1}),
            'money': wb.add_format({'num_format': '#,##0.00', 'border': 1}),
            'date': wb.add_format({'num_format': 'dd/mm/yyyy', 'border': 1}),
            'total_lbl': wb.add_format({
                'bold': True, 'border': 1, 'bg_color': '#D6EAF8',
            }),
            'total': wb.add_format({
                'bold': True, 'num_format': '#,##0.00', 'border': 1,
                'bg_color': '#D6EAF8',
            }),
        }

    def _write_excel(self, wb):
        ws = wb.add_worksheet('Reporte Diario')
        fmt = self._build_formats(wb)

        # Anchos de columna
        ws.set_column(0, 0, 32)   # Sucursal / concepto
        ws.set_column(1, 5, 18)   # Valores

        # ── Encabezado del reporte ─────────────────────────────────────────
        ws.merge_range('A1:F1', 'REPORTE DIARIO DE INGRESOS', fmt['title'])
        ws.write('A2', 'Fecha:', fmt['label'])
        ws.write('B2', self.date.strftime('%d/%m/%Y'))
        ws.write('A3', 'Empresa:', fmt['label'])
        ws.write('B3', self.company_id.name)

        row = 4

        # ── SECCIÓN 1: VENTAS POR SUCURSAL ─────────────────────────────────
        ws.merge_range(row, 0, row, 4, 'Ventas', fmt['section'])
        row += 1

        for col, h in enumerate(['Sucursal', 'Ventas Sin IVA', 'Ingresos Efectivo',
                                  'Ingresos Tarjetas', 'Ingresos en MP']):
            ws.write(row, col, h, fmt['header'])
        row += 1

        ventas = self._get_ventas_por_ou()
        efectivo = self._get_ingresos_por_ou('cash')
        tarjetas = self._get_ingresos_por_ou('card')
        mp = self._get_ingresos_por_ou('mp')

        all_ous = sorted(set(list(ventas) + list(efectivo) + list(tarjetas) + list(mp)))

        tot_v = tot_e = tot_t = tot_m = 0.0
        for ou in all_ous:
            v = ventas.get(ou, 0.0)
            e = efectivo.get(ou, 0.0)
            t = tarjetas.get(ou, 0.0)
            m = mp.get(ou, 0.0)
            ws.write(row, 0, ou, fmt['text'])
            ws.write_number(row, 1, v, fmt['money'])
            ws.write_number(row, 2, e, fmt['money'])
            ws.write_number(row, 3, t, fmt['money'])
            ws.write_number(row, 4, m, fmt['money'])
            tot_v += v
            tot_e += e
            tot_t += t
            tot_m += m
            row += 1

        ws.write(row, 0, 'TOTAL', fmt['total_lbl'])
        ws.write_number(row, 1, tot_v, fmt['total'])
        ws.write_number(row, 2, tot_e, fmt['total'])
        ws.write_number(row, 3, tot_t, fmt['total'])
        ws.write_number(row, 4, tot_m, fmt['total'])
        row += 2

        # ── SECCIÓN 2: CHEQUES ─────────────────────────────────────────────
        ws.merge_range(row, 0, row, 5, 'Cheques', fmt['section'])
        row += 1

        for col, h in enumerate(['Cliente', 'Sucursal', 'N° Valor',
                                  'Fecha de Ingreso', 'VT', 'Ingresos']):
            ws.write(row, col, h, fmt['header'])
        row += 1

        cheques = self._get_cheques()
        tot_ch = 0.0
        for ch in cheques:
            ws.write(row, 0, ch['cliente'] or '', fmt['text'])
            ws.write(row, 1, ch['sucursal'] or '', fmt['text'])
            ws.write(row, 2, ch['nro_valor'] or '', fmt['text'])
            if ch['fecha_ingreso']:
                ws.write_datetime(row, 3, ch['fecha_ingreso'], fmt['date'])
            else:
                ws.write(row, 3, '', fmt['text'])
            if ch['vencimiento']:
                ws.write_datetime(row, 4, ch['vencimiento'], fmt['date'])
            else:
                ws.write(row, 4, '', fmt['text'])
            importe = ch['importe'] or 0.0
            ws.write_number(row, 5, importe, fmt['money'])
            tot_ch += importe
            row += 1

        ws.write(row, 4, 'TOTAL', fmt['total_lbl'])
        ws.write_number(row, 5, tot_ch, fmt['total'])
        row += 2

        # ── SECCIÓN 3: TRANSFERENCIAS ──────────────────────────────────────
        ws.merge_range(row, 0, row, 1, 'Transferencias', fmt['section'])
        row += 1

        ws.write(row, 0, 'Banco', fmt['header'])
        ws.write(row, 1, 'Total', fmt['header'])
        row += 1

        transferencias = self._get_transferencias()
        tot_tr = 0.0
        for tr in transferencias:
            ws.write(row, 0, tr['banco'], fmt['text'])
            ws.write_number(row, 1, tr['total'], fmt['money'])
            tot_tr += tr['total']
            row += 1

        ws.write(row, 0, 'TOTAL', fmt['total_lbl'])
        ws.write_number(row, 1, tot_tr, fmt['total'])
