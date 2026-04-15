# -*- coding: utf-8 -*-
{
    'name': 'Lupatini - Reporte Diario de Ingresos',
    'summary': 'Reporte diario: ventas por sucursal, cheques y transferencias',
    'version': '17.0.1.1.0',
    'author': 'Guvens',
    'category': 'Accounting',
    'license': 'LGPL-3',
    # account_operating_unit: agrega operating_unit_id a account.payment y account.move
    # account_check (OCA): agrega account.check con number y payment_date; ap.check_id
    'depends': [
        'account_accountant',
        'account_operating_unit',
        'account_check',
    ],
    'data': [
        'security/ir.model.access.csv',
        'views/account_journal_view.xml',
        'views/reporte_diario_views.xml',
    ],
    'installable': True,
    'application': False,
    'auto_install': False,
}
