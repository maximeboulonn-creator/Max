import sys; sys.path.insert(0,'.')
from openpyxl import Workbook
import params, tables1, tables2, preview
wb=Workbook(); wb.remove(wb.active)
params.build_params(wb); tables1.build_limites(wb); tables1.build_liquidite(wb); tables1.build_scenarios(wb)
tables2.build_fonds(wb); tables2.build_recos(wb); tables2.build_data(wb)
preview.build_preview(wb)
for ws in wb.worksheets:
    ws.sheet_view.showGridLines=False; ws.freeze_panes=None
wb.active=0
out='Risk_Report_Openstone_Infraworld_light.xlsx'
wb.save(out); print('saved',out)
