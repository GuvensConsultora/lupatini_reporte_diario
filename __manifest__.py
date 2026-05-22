# -*- coding: utf-8 -*-
{
    'name': 'Lupatini - Reporte Diario de Ingresos',
    'summary': 'Reporte diario (Excel) + reporte PDF de ventas por unidad operativa con selector de período',
    'version': '17.0.1.2.1',
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
        'security/security.xml',
        'security/ir.model.access.csv',
        'views/account_journal_view.xml',
        'views/reporte_diario_views.xml',
        'wizard/reporte_ventas_uo_wizard.xml',
        'report/report_ventas_uo_templates.xml',
        'views/menu_ventas_uo.xml',
    ],
    'installable': True,
    'application': False,
    'auto_install': False,
}
