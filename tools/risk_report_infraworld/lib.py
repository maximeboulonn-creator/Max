# -*- coding: utf-8 -*-
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break

NAVY='01313E'; TEAL='0287AA'; GREY='6A7A85'; LINE='DCE5EA'; PALE='B8C2C6'; WHITE='FFFFFF'; INK='1A2027'
NB=' '  # espace insécable
def F(b=False,c=INK,i=False): return Font(name='Calibri',size=9,bold=b,italic=i,color=c)
def fill(c): return PatternFill('solid',fgColor=c)
def al(h=None,v='center',wrap=True): return Alignment(horizontal=h,vertical=v,wrap_text=wrap)
thin=Side(style='thin',color=LINE); med=Side(style='medium',color=NAVY)
B_THIN=Border(bottom=thin); B_MED=Border(bottom=med)

def setup(ws, widths):
    ws.sheet_view.showGridLines=False
    for k,v in widths.items(): ws.column_dimensions[k].width=v
    ws.page_setup.orientation='portrait'; ws.page_setup.paperSize=9
    ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.print_options.horizontalCentered=True
    ws.page_margins.left=0.4; ws.page_margins.right=0.4; ws.page_margins.top=0.5; ws.page_margins.bottom=0.5

class Preview:
    """Écrit le rapport paginé, colonnes B..G, étiquette en H."""
    def __init__(s, ws, fund, header):
        s.ws=ws; s.r=1; s.fund=fund; s.header=header
        setup(ws, {'A':2,'B':30,'C':15,'D':15,'E':15,'F':15,'G':46,'H':2})
    def tag(s,t):
        c=s.ws.cell(s.r,8,t); c.font=F(c=WHITE)
    def _h(s,h):
        s.ws.row_dimensions[s.r].height=h
    def band(s):
        for col in range(2,8):
            c=s.ws.cell(s.r,col); c.fill=fill(NAVY)
        a=s.ws.cell(s.r,2,s.fund); a.font=F(True,WHITE); a.alignment=al('left')
        g=s.ws.cell(s.r,7,s.header); g.font=F(False,PALE); g.alignment=al('right')
        s.tag('BAND'); s._h(20); s.r+=1
    def cover(s, lines):
        # lines: list of (text, bold, color)
        for text,b,col,h in lines:
            for cc in range(2,8): s.ws.cell(s.r,cc).fill=fill(NAVY)
            s.ws.merge_cells(start_row=s.r,start_column=2,end_row=s.r,end_column=7)
            c=s.ws.cell(s.r,2,text); c.font=F(b,col); c.alignment=al('left',wrap=True); s._h(h)
            s.tag('COVER'); s.r+=1
    def title(s,text):
        for col in range(2,8): s.ws.cell(s.r,col).border=B_MED
        c=s.ws.cell(s.r,2,text); c.font=F(False,GREY); c.alignment=al('left',wrap=False); s._h(22); s.tag('TITLE'); s.r+=1
    def sub(s,text):
        for col in range(2,8): s.ws.cell(s.r,col).border=B_MED
        c=s.ws.cell(s.r,2,text); c.font=F(True,NAVY); c.alignment=al('left',wrap=False); s._h(16); s.tag('SUB'); s.r+=1
        s.r+=1
    def sec(s,text,note=None,h=None):
        s.ws.merge_cells(start_row=s.r,start_column=2,end_row=s.r,end_column=7)
        for col in range(2,8): s.ws.cell(s.r,col).border=B_THIN
        v=text if not note else text+'\n'+note
        c=s.ws.cell(s.r,2,v); c.font=F(True,TEAL); c.alignment=al('left',wrap=True)
        s._h(h or (26 if note else 18)); s.tag('SEC'); s.r+=1
    def hdr(s,cols,aligns=None):
        aligns=aligns or ['left']+['center']*(len(cols)-2)+['left']
        for i,v in enumerate(cols):
            c=s.ws.cell(s.r,2+i,v); c.font=F(True,WHITE); c.fill=fill(NAVY); c.alignment=al(aligns[i])
        for col in range(2,8):
            if s.ws.cell(s.r,col).value is None: s.ws.cell(s.r,col).fill=fill(NAVY)
        s._h(18); s.tag('HDR'); s.r+=1
    def row(s,vals,fmts=None,aligns=None,h=None,bold_first=True,merge=None):
        aligns=aligns or ['left']+['center']*(len(vals)-2)+['left']
        for i,v in enumerate(vals):
            if v is None: continue
            c=s.ws.cell(s.r,2+i,v); c.font=F(bold_first and i==0,NAVY if i==0 else INK)
            c.alignment=al(aligns[i] if i<len(aligns) else 'left')
            if fmts and i<len(fmts) and fmts[i]: c.number_format=fmts[i]
        for col in range(2,8): s.ws.cell(s.r,col).border=B_THIN
        if merge:
            for a,b in merge: s.ws.merge_cells(start_row=s.r,start_column=a,end_row=s.r,end_column=b)
        s._h(h or 24); s.tag('ROW'); s.r+=1
    def kpi(s,labels,values,subs,fmts=None):
        n=len(labels)
        for i in range(n):
            c=s.ws.cell(s.r,2+i,labels[i]); c.font=F(False,GREY); c.alignment=al('left')
        s._h(24); s.tag('KPI1'); s.r+=1
        for i in range(n):
            c=s.ws.cell(s.r,2+i,values[i]); c.font=F(True,NAVY); c.alignment=al('left',wrap=False)
            if fmts and fmts[i]: c.number_format=fmts[i]
        s._h(18); s.tag('KPI2'); s.r+=1
        for i in range(n):
            c=s.ws.cell(s.r,2+i,subs[i]); c.font=F(False,GREY); c.alignment=al('left')
        s._h(34); s.tag('KPI3'); s.r+=1
    def note(s,text,h=None,lines=None):
        s.ws.merge_cells(start_row=s.r,start_column=2,end_row=s.r,end_column=7)
        c=s.ws.cell(s.r,2,text); c.font=F(False,GREY); c.alignment=al('left',v='top',wrap=True)
        if lines is None:
            lines=max(1,len(str(text))//150+1)
        s._h(h or 13*lines); s.tag('NOTE'); s.r+=1
    def blank(s,n=1): s.r+=n
    def pagebreak(s):
        s.ws.row_breaks.append(Break(id=s.r-1)); s.r+=1

class Sheet:
    """Onglet de calcul : B2 titre, B3 sous-titre gris, tableaux avec en-tête navy."""
    def __init__(s, ws, title, subtitle, widths):
        s.ws=ws; setup(ws,widths); s.r=2
        c=ws.cell(2,2,title); c.font=Font(name='Calibri',size=10,bold=True,color=NAVY)
        c=ws.cell(3,2,subtitle); c.font=F(False,GREY); c.alignment=al('left',wrap=False)
        s.r=5
    def block(s,text):
        c=s.ws.cell(s.r,2,text); c.font=F(True,WHITE); c.fill=fill(NAVY); c.alignment=al('left')
        for col in range(3,s.ncols+2): s.ws.cell(s.r,col).fill=fill(NAVY)
        s.r+=1
    def hdr(s,cols):
        s.ncols=len(cols)
        for i,v in enumerate(cols):
            c=s.ws.cell(s.r,2+i,v); c.font=F(True,WHITE); c.fill=fill(TEAL); c.alignment=al('left')
        s.ws.row_dimensions[s.r].height=18; s.r+=1
    def row(s,vals,fmts=None,h=None,input_cols=(),key_cols=()):
        for i,v in enumerate(vals):
            if v is None: continue
            c=s.ws.cell(s.r,2+i,v); c.font=F(False,INK); c.alignment=al('left',v='top')
            if fmts and i<len(fmts) and fmts[i]: c.number_format=fmts[i]
            if i in input_cols: c.fill=fill('FFF8F2')
            if i in key_cols: c.fill=fill('D9F785'); c.font=F(True,INK)
        for col in range(2,2+len(vals)): s.ws.cell(s.r,col).border=B_THIN
        if h: s.ws.row_dimensions[s.r].height=h
        s.r+=1; return s.r-1
    def note(s,text,ncols=None):
        n=ncols or getattr(s,'ncols',6)
        s.ws.merge_cells(start_row=s.r,start_column=2,end_row=s.r,end_column=1+n)
        c=s.ws.cell(s.r,2,text); c.font=F(False,GREY); c.alignment=al('left',v='top',wrap=True)
        s.ws.row_dimensions[s.r].height=13*max(1,len(text)//140+1); s.r+=1
    def blank(s,n=1): s.r+=n
