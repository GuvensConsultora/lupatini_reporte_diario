# -*- coding: utf-8 -*-
from odoo import fields, models


class AccountJournal(models.Model):
    _inherit = 'account.journal'

    lupatini_ingreso_tipo = fields.Selection([
        ('cash', 'Efectivo'),
        ('card', 'Tarjetas'),
        ('mp', 'Mercado Pago'),
        ('bank', 'Banco (Transferencias)'),
    ],
        string='Tipo en Reporte Diario',
        help=(
            'Categoría que usa este diario en el Reporte Diario de Ingresos.\n'
            'Dejar vacío para excluirlo del reporte.\n'
            'Los diarios de cheques se detectan automáticamente por '
            '"Usar cheques".'
        ),
    )
