# -*- coding: utf-8 -*-
import io
import base64
from collections import defaultdict
from dateutil.relativedelta import relativedelta

from odoo import models, fields, _
from odoo.exceptions import UserError

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None

# Orden y etiqueta de cada medio en el reporte. La clave es `lupatini_ingreso_tipo` del
# diario; un diario sin tipo cae en 'other' para que su plata no desaparezca del total.
MEDIOS = (
    ('cash', 'Efectivo'),
    ('bank', 'Transferencias'),
    ('mp', 'Mercado Pago'),
    ('card', 'Tarjetas'),
    ('check', 'Cheques'),
    ('echeq', 'eCheq'),
    ('other', 'Otros'),
)
# Las retenciones que nos hacen los clientes se cobran como un pago más, pero no es
# plata: se informan aparte y no suman a lo que entró.
RETENCION = 'retention'
SIN_SUCURSAL = 'Sin sucursal'


def _mes_anterior_desde(self):
    return fields.Date.context_today(self).replace(day=1) - relativedelta(months=1)


def _mes_anterior_hasta(self):
    return fields.Date.context_today(self).replace(day=1) - relativedelta(days=1)


class IngresosPeriodoWizard(models.TransientModel):
    _name = 'lupatini.ingresos.periodo.wizard'
    _description = 'Ingresos por período'

    date_from = fields.Date(string='Desde', required=True, default=_mes_anterior_desde)
    date_to = fields.Date(string='Hasta', required=True, default=_mes_anterior_hasta)
    company_id = fields.Many2one(
        'res.company',
        string='Empresa',
        required=True,
        default=lambda self: self.env.company,
    )
    operating_unit_ids = fields.Many2many(
        'operating.unit',
        string='Sucursales',
        help='Vacío = todas las sucursales.',
    )

    # -------------------------------------------------------------------------
    # Motor
    # Qué entra: pagos publicados de CLIENTES. Los pagos a proveedores y las
    # transferencias entre cajas quedan afuera por construcción (no son de
    # clientes), que es justamente lo que el libro mayor no permite separar.
    # Los pagos salientes a clientes son las «Reversiones» (NC devueltas):
    # restan del medio por el que se devolvió.
    # -------------------------------------------------------------------------

    def _pagos(self):
        dom = [
            ('company_id', '=', self.company_id.id),
            ('state', '=', 'posted'),
            ('partner_type', '=', 'customer'),
            ('is_internal_transfer', '=', False),
            ('date', '>=', self.date_from),
            ('date', '<=', self.date_to),
        ]
        if self.operating_unit_ids:
            # La sucursal es la del pago; si el pago no la tiene, la de su diario.
            ous = self.operating_unit_ids.ids
            dom += ['|', ('operating_unit_id', 'in', ous),
                    '&', ('operating_unit_id', '=', False),
                    ('journal_id.operating_unit_id', 'in', ous)]
        return self.env['account.payment'].search(dom, order='date, id')

    def _datos(self):
        self.ensure_one()
        # Todo el cálculo corre en la empresa elegida, no en la activa del usuario: con la CIA
        # activa y la S.H. elegida, las reglas multiempresa ocultaban sus pagos y daba cero.
        self = self.with_company(self.company_id)
        if self.date_from > self.date_to:
            raise UserError(_('La fecha «Desde» es posterior a «Hasta».'))
        etiquetas = dict(MEDIOS)
        por_medio = defaultdict(lambda: {'cobro': 0.0, 'reversion': 0.0, 'n': 0})
        pivot = defaultdict(lambda: defaultdict(float))
        detalle = []
        for p in self._pagos():
            tipo = p.journal_id.lupatini_ingreso_tipo or 'other'
            ou = p.operating_unit_id or p.journal_id.operating_unit_id
            sucursal = ou.name if ou else SIN_SUCURSAL
            entra = p.payment_type == 'inbound'
            importe = p.amount if entra else -p.amount
            por_medio[tipo]['cobro' if entra else 'reversion'] += importe
            por_medio[tipo]['n'] += 1
            if tipo != RETENCION:
                pivot[sucursal][tipo] += importe
            detalle.append({
                'fecha': p.date,
                'recibo': p.payment_group_id.display_name or '',
                'pago': p.name,
                'cliente': p.partner_id.display_name or '',
                'sucursal': sucursal,
                'medio': 'Retenciones' if tipo == RETENCION else etiquetas[tipo],
                'diario': p.journal_id.name,
                'tipo': 'Cobro' if entra else 'Reversión',
                'importe': importe,
            })

        medios = []
        for clave, etiqueta in MEDIOS:
            if clave in por_medio:
                d = por_medio[clave]
                medios.append({'clave': clave, 'medio': etiqueta, 'cobro': d['cobro'],
                               'reversion': d['reversion'], 'neto': d['cobro'] + d['reversion'],
                               'n': d['n']})
        claves = [m['clave'] for m in medios]
        sucursales = []
        for nombre in sorted(pivot, key=lambda s: (s == SIN_SUCURSAL, s)):
            vals = [pivot[nombre].get(c, 0.0) for c in claves]
            sucursales.append({'sucursal': nombre, 'valores': vals, 'total': sum(vals)})

        ret = por_medio.get(RETENCION, {'cobro': 0.0, 'n': 0})
        nc = self.env['account.move'].read_group(
            [('company_id', '=', self.company_id.id), ('move_type', '=', 'out_refund'),
             ('state', '=', 'posted'),
             ('invoice_date', '>=', self.date_from), ('invoice_date', '<=', self.date_to)],
            ['amount_total_signed:sum'], [])[0]
        prov = self.env['account.payment'].read_group(
            [('company_id', '=', self.company_id.id), ('state', '=', 'posted'),
             ('partner_type', '=', 'supplier'),
             ('date', '>=', self.date_from), ('date', '<=', self.date_to)],
            ['amount:sum'], [])[0]
        return {
            'empresa': self.company_id.name,
            'desde': self.date_from.strftime('%d/%m/%Y'),
            'hasta': self.date_to.strftime('%d/%m/%Y'),
            'filtro_sucursales': ', '.join(self.operating_unit_ids.mapped('name')),
            'medios': medios,
            'total_cobro': sum(m['cobro'] for m in medios),
            'total_reversion': sum(m['reversion'] for m in medios),
            'total_neto': sum(m['neto'] for m in medios),
            'columnas': [m['medio'] for m in medios],
            'sucursales': sucursales,
            'totales_sucursal': [sum(s['valores'][i] for s in sucursales)
                                 for i in range(len(claves))],
            'retenciones': ret['cobro'], 'retenciones_n': ret['n'],
            'nc': nc['amount_total_signed'] or 0.0, 'nc_n': nc['__count'],
            'proveedores': prov['amount'] or 0.0, 'proveedores_n': prov['__count'],
            'detalle': detalle,
        }

    # -------------------------------------------------------------------------
    # Acciones
    # -------------------------------------------------------------------------

    def action_print_pdf(self):
        self.ensure_one()
        return self.env.ref(
            'lupatini_reporte_diario.action_report_ingresos_periodo').report_action(self)

    def action_export_excel(self):
        self.ensure_one()
        if not xlsxwriter:
            raise UserError(_('Se requiere la librería xlsxwriter para exportar a Excel.'))
        datos = self._datos()
        output = io.BytesIO()
        wb = xlsxwriter.Workbook(output, {'in_memory': True})
        self._write_excel(wb, datos)
        wb.close()
        filename = 'ingresos_%s_%s.xlsx' % (self.date_from.strftime('%Y%m%d'),
                                            self.date_to.strftime('%Y%m%d'))
        attachment = self.env['ir.attachment'].create({
            'name': filename,
            'type': 'binary',
            'datas': base64.b64encode(output.getvalue()),
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })
        return {
            'type': 'ir.actions.act_url',
            'url': '/web/content/%d?download=true' % attachment.id,
            'target': 'new',
        }

    # -------------------------------------------------------------------------
    # Excel: Resumen, Por sucursal y Detalle (con filtro, para procesarlo a gusto)
    # -------------------------------------------------------------------------

    def _write_excel(self, wb, d):
        titulo = wb.add_format({'bold': True, 'font_size': 14})
        head = wb.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#1F4E79',
                              'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        num = wb.add_format({'num_format': '#,##0.00;-#,##0.00'})
        tot_txt = wb.add_format({'bold': True, 'bg_color': '#DDEBF7'})
        tot_num = wb.add_format({'bold': True, 'bg_color': '#DDEBF7',
                                 'num_format': '#,##0.00;-#,##0.00'})
        fecha = wb.add_format({'num_format': 'dd/mm/yyyy'})

        ws = wb.add_worksheet('Resumen')
        ws.write(0, 0, 'Lo que entró del %s al %s — %s' % (d['desde'], d['hasta'], d['empresa']),
                 titulo)
        if d['filtro_sucursales']:
            ws.write(1, 0, 'Sucursales: %s' % d['filtro_sucursales'])
        for c, (txt, ancho) in enumerate([('Medio', 22), ('Cobrado', 18), ('Reversiones (NC)', 18),
                                          ('Neto que entró', 18), ('Cantidad', 10)]):
            ws.write(3, c, txt, head)
            ws.set_column(c, c, ancho)
        r = 4
        for m in d['medios']:
            ws.write(r, 0, m['medio'])
            ws.write_number(r, 1, m['cobro'], num)
            ws.write_number(r, 2, m['reversion'], num)
            ws.write_number(r, 3, m['neto'], num)
            ws.write_number(r, 4, m['n'])
            r += 1
        ws.write(r, 0, 'TOTAL', tot_txt)
        ws.write_number(r, 1, d['total_cobro'], tot_num)
        ws.write_number(r, 2, d['total_reversion'], tot_num)
        ws.write_number(r, 3, d['total_neto'], tot_num)
        ws.write(r, 4, '', tot_txt)
        r += 2
        for txt, imp, n in [
            ('Retenciones que nos hicieron (no es plata)', d['retenciones'], d['retenciones_n']),
            ('Notas de crédito a clientes del período', d['nc'], d['nc_n']),
            ('Pagos a proveedores (NO incluidos)', d['proveedores'], d['proveedores_n']),
        ]:
            ws.write(r, 0, txt)
            ws.write_number(r, 1, imp, num)
            ws.write_number(r, 4, n)
            r += 1

        ws2 = wb.add_worksheet('Por sucursal')
        cols = ['Sucursal'] + d['columnas'] + ['Total']
        for c, txt in enumerate(cols):
            ws2.write(0, c, txt, head)
            ws2.set_column(c, c, 22 if c == 0 else 16)
        r = 1
        for s in d['sucursales']:
            ws2.write(r, 0, s['sucursal'])
            for c, v in enumerate(s['valores'] + [s['total']], 1):
                ws2.write_number(r, c, v, num)
            r += 1
        ws2.write(r, 0, 'TOTAL', tot_txt)
        for c, v in enumerate(d['totales_sucursal'] + [d['total_neto']], 1):
            ws2.write_number(r, c, v, tot_num)
        ws2.freeze_panes(1, 1)

        ws3 = wb.add_worksheet('Detalle')
        cols = [('Fecha', 11), ('Recibo', 16), ('Pago', 22), ('Cliente', 36), ('Sucursal', 20),
                ('Medio', 15), ('Diario', 28), ('Tipo', 11), ('Importe', 16)]
        for c, (txt, ancho) in enumerate(cols):
            ws3.write(0, c, txt, head)
            ws3.set_column(c, c, ancho)
        for r, f in enumerate(d['detalle'], 1):
            ws3.write_datetime(r, 0, fields.Datetime.to_datetime(f['fecha']), fecha)
            for c, k in enumerate(['recibo', 'pago', 'cliente', 'sucursal', 'medio', 'diario',
                                   'tipo'], 1):
                ws3.write(r, c, f[k])
            ws3.write_number(r, 8, f['importe'], num)
        ws3.freeze_panes(1, 0)
        ws3.autofilter(0, 0, max(len(d['detalle']), 1), len(cols) - 1)
