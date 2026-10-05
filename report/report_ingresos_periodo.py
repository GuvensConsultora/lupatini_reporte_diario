# -*- coding: utf-8 -*-
from odoo import models


class ReportIngresosPeriodo(models.AbstractModel):
    # Tabla derivada 'report_lupatini_reporte_diario_ingresos_periodo_doc' (52 chars):
    # dentro del límite de 63 de Postgres.
    _name = 'report.lupatini_reporte_diario.ingresos_periodo_doc'
    _description = 'Ingresos por período'

    def _get_report_values(self, docids, data=None):
        docs = self.env['lupatini.ingresos.periodo.wizard'].browse(docids)
        return {
            'doc_ids': docids,
            'doc_model': 'lupatini.ingresos.periodo.wizard',
            'docs': docs,
            # Los datos se calculan en el wizard: el PDF y el Excel salen del mismo cálculo.
            'datos': docs[:1]._datos(),
        }
