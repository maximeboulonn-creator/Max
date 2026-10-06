import sys; sys.path.insert(0,'.')
from openpyxl import Workbook
import params, tables1, tables2, preview, engine, scen, priips
wb=Workbook(); wb.remove(wb.active)
params.build_params(wb); params.build_params_stress(wb); priips.build_params_priips(wb); tables1.build_limites(wb); tables1.build_liquidite(wb); SM=engine.build_engine(wb); scen.build_scenarios(wb,SM)
tables2.build_fonds(wb); priips.build_priips_sheet(wb); tables2.build_recos(wb); tables2.build_data(wb)
preview.build_preview(wb)
# ordre des onglets : PREVIEW, 0, 1, 2, 3, 3b, 4, 5, DATA
order=['PREVIEW','0 Paramètres','1 Limites','2 Liquidité','3 Scénarios','3b Moteur','4 Fonds cibles','4b PRIIPs','5 Recommandations','DATA']
wb._sheets=[wb[n] for n in order]
for ws in wb.worksheets:
    ws.sheet_view.showGridLines=False; ws.freeze_panes=None
wb.active=0
out='/tmp/claude-0/-home-user/fbfd82fa-9176-5404-b298-69879606bee1/scratchpad/build/Risk_Report_Openstone_Infraworld_light.xlsx'
wb.save(out); print('saved',out)
