# -*- coding: utf-8 -*-
from odoo import fields, models


class AccountJournal(models.Model):
    _inherit = 'account.journal'

    lupatini_ingreso_tipo = fields.Selection([
        ('cash', 'Efectivo'),
        ('card', 'Tarjetas'),
        ('mp', 'Mercado Pago'),
        ('bank', 'Banco (Transferencias)'),
        # Agregados en 17.0.1.4.0 para «Ingresos por período». El reporte diario sigue
        # leyendo sólo cash/card/mp/bank, así que estas opciones no lo alteran.
        ('check', 'Cheques'),
        ('echeq', 'eCheq'),
        ('retention', 'Retenciones (no es dinero)'),
        ('other', 'Otros'),
    ],
        string='Tipo en Reporte Diario',
        help=(
            'Categoría que usa este diario en el Reporte Diario de Ingresos y en '
            'Ingresos por período.\n'
            'Dejar vacío para excluirlo del reporte diario; en Ingresos por período '
            'un diario sin tipo aparece en «Otros» para que no se pierda.\n'
            'Retenciones se informa aparte: no suma a lo que entró.'
        ),
    )
