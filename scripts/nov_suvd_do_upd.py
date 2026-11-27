"""KPI_Comm_102_1127 slide 3: ДО column = ДО operations not overlapping with СУВД (user's table, 27.11), totals and % of forecast."""
import sys, os, re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import oct_data
from pptx import Presentation
from pptx.oxml.ns import qn
SRC, OUT, MASTER = sys.argv[1:4]
SUVD = {'E-com': 223529768 + 106285640, 'AdTech': 235533379 + 210192297, 'М+': 48717135 + 123227080}
DO = {'E-com': 189248512 + 98258928, 'AdTech': 38435372 + 63270605, 'М+': 17537917 - 11492341}
assert sum(SUVD.values()) == 947485299 and abs(sum(DO.values()) - 395258994) <= 2   # screenshot amounts are rounded to 1 rub
SUVD = {k: v / 1e6 for k, v in SUVD.items()}; DO = {k: v / 1e6 for k, v in DO.items()}
S, PL = oct_data.load(MASTER)
FC = {'E-com': S[10]['fact'][10], 'M+': PL[15]['fact'][10]}
FC['total'] = PL[16]['fact'][10] - FC['M+']; FC['AdTech'] = FC['total'] - FC['E-com']
num = lambda v: f'{round(v):,}'.replace(',', ' ')
rows = {1: ('total', SUVD['E-com'] + SUVD['AdTech'], DO['E-com'] + DO['AdTech']), 2: ('E-com', SUVD['E-com'], DO['E-com']),
        3: ('AdTech', SUVD['AdTech'], DO['AdTech']), 4: ('M+', SUVD['М+'], DO['М+'])}
p = Presentation(SRC); s = p.slides[2]
t = [sh for sh in s.shapes if sh.has_table][0].table
def put(cell, text):
    runs = [r for pp in cell.text_frame.paragraphs for r in pp.runs]
    runs[0].text = text
    for r in runs[1:]: r._r.getparent().remove(r._r)
    runs[0]._r.find(qn('a:rPr')).set('lang', 'ru-RU' if re.search('[А-Яа-яЁё]', text) else 'en-US')
res = {}
for ri, (k, su, do) in rows.items():
    fc = FC[k]; tot = su + do
    assert t.cell(ri, 1).text == num(fc) and t.cell(ri, 2).text == num(su), (ri, t.cell(ri, 1).text, t.cell(ri, 2).text)
    old = [t.cell(ri, c).text for c in (3, 4, 5, 6)]
    put(t.cell(ri, 3), num(do)); put(t.cell(ri, 4), num(tot)); put(t.cell(ri, 5), f'{round(tot / fc * 100)}%')
    put(t.cell(ri, 6), '-' + num(fc - tot) if fc > tot else num(tot - fc))
    res[k] = tot / fc
    print(t.cell(ri, 0).text, '|', old, '->', [t.cell(ri, c).text for c in (3, 4, 5, 6)])
TOT = rows[1][1] + rows[1][2]
title = [sh for sh in s.shapes if sh.name == 'Текст 1'][0]
txt = title.text_frame.text
new = f'Прогноз по заведённым операциям ноябрь. В СУВД и ДО на 27.11 = {num(TOT)} млн руб. = {round(res["total"] * 100)}% прогноза'
pr = title.text_frame.paragraphs[0]
for r in pr.runs[1:]: r._r.getparent().remove(r._r)
pr.runs[0].text = new
print('title:', txt, '->', new)
p.save(OUT)
