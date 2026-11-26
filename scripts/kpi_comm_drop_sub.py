"""Drop supplier / non-supplier sub-rows for DMP and Online+DMP (Instore digital keeps them); target and 'vs target' move to the direction row (whole direction vs supplier target)."""
import sys
from pptx import Presentation
from pptx.oxml.ns import qn
SRC, OUT = sys.argv[1:3]
# exact values: (без не РН, с не РН, таргет) — 10М из F.pkl, 10+2 = 10М + прогноз мастера, таргеты из mc.xlsx
V = {0: {'DMP': (40.3, 40.8, 73.0), 'Online+DMP': (399.3, 789.4, 780.0)},
     1: {'DMP': (52.0, 52.6, 110.0), 'Online+DMP': (509.8, 1015.2, 1020.0)}}
fmt = lambda v: f'{round(v):,}'.replace(',', ' ')
def pct(a, b):
    p = round((a / b - 1) * 100)
    return ('0%' if p == 0 else f'{p:+d}%'), ('CCFFCC' if p > 5 else 'FFEB9C' if p >= -5 else 'FFC7CE')
def put(cell, text, fill=None):
    r = cell.text_frame.paragraphs[0].runs[0]; r.text = text; r._r.find(qn('a:rPr')).set('lang', 'en-US')
    if fill:
        sf = cell._tc.find(qn('a:tcPr')).find(qn('a:solidFill')); sf[0].tag; 
        for c in list(sf): sf.remove(c)
        from lxml import etree
        etree.SubElement(sf, qn('a:srgbClr')).set('val', fill)
p = Presentation(SRC)
for si, s in enumerate(p.slides):
    gf = [sh for sh in s.shapes if sh.has_table][0]; t = gf.table
    trs = list(t._tbl.tr_lst)
    labels = [t.cell(i, 0).text.strip('\xa0 ') for i in range(len(trs))]
    drop = []
    for name, (f0, f1, tg) in V[si].items():
        i = labels.index(name)
        assert labels[i + 1] == 'Поставщики' and labels[i + 2] == 'Не поставщики'
        assert t.cell(i, 1).text == fmt(f0) and t.cell(i, 2).text == fmt(f1), (name, t.cell(i, 1).text, t.cell(i, 2).text)
        assert t.cell(i + 1, 4).text == fmt(tg), (name, t.cell(i + 1, 4).text)
        put(t.cell(i, 4), fmt(tg))
        put(t.cell(i, 8), *pct(f0, tg)); put(t.cell(i, 9), *pct(f1, tg))
        print(si, name, fmt(tg), t.cell(i, 8).text, t.cell(i, 9).text)
        drop += [trs[i + 1], trs[i + 2]]
    h = sum(int(tr.get('h')) for tr in drop)
    for tr in drop: t._tbl.remove(tr)
    gf.height = gf.height - h
p.save(OUT)
