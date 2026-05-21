# -*- coding: utf-8 -*-
from odoo import models, fields, _
from odoo.exceptions import UserError


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
