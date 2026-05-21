# -*- coding: utf-8 -*-
from odoo import models


class ReportVentasUo(models.AbstractModel):
    # _name corto a propósito: la tabla derivada
    # 'report_lupatini_reporte_diario_ventas_uo_doc' (44 chars) queda dentro
    # del límite de 63 chars de Postgres para nombres de modelo.
    _name = 'report.lupatini_reporte_diario.ventas_uo_doc'
    _description = 'Reporte Ventas por Unidad Operativa'

    def _get_report_values(self, docids, data=None):
        data = data or {}
        return {
            'doc_ids': docids,
            'doc_model': 'lupatini.reporte.ventas.uo.wizard',
            'docs': self.env['lupatini.reporte.ventas.uo.wizard'].browse(docids),
            'date_from': data.get('date_from', ''),
            'date_to': data.get('date_to', ''),
            'company_name': data.get('company_name', ''),
            'filas': data.get('filas', []),
            'total': data.get('total', 0.0),
        }
