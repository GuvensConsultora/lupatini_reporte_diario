# -*- coding: utf-8 -*-
{
    'name': 'Lupatini - Reporte Diario de Ingresos',
    'summary': 'Reporte diario: ventas por sucursal, cheques y transferencias',
    'version': '17.0.1.0.0',
    'author': 'Guvens',
    'category': 'Accounting',
    'license': 'LGPL-3',
    # account_operating_unit: agrega operating_unit_id a account.payment y account.move
    # l10n_latam_check: agrega l10n_latam_check_number y l10n_latam_check_payment_date
    'depends': [
        'account_accountant',
        'account_operating_unit',
        'l10n_latam_check',
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
