# -*- coding: utf-8 -*-
from odoo import fields, models


class AccountPaymentTerm(models.Model):
    _inherit = 'account.payment.term'

    lupatini_es_contado = fields.Boolean(
        string='Es venta de contado',
        default=False,
        help=(
            'Marcar las condiciones de pago que corresponden a ventas de contado.\n'
            'El Reporte de Ventas por Unidad Operativa separa por este campo, sin '
            'importar con qué medio se haya cobrado: una venta de contado pagada '
            'con tarjeta sigue siendo de contado.\n'
            'Las condiciones sin marcar se informan como cuenta corriente. Los '
            'comprobantes que no tienen ninguna condición cargada se informan '
            'aparte, en su propia columna.'
        ),
    )
