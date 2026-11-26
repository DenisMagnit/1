"""Update total / suppliers / non-suppliers rows of the commerce KPI slides (10M and 10+2) to master segments (12 Сегменты), as in KPI triggers."""
import sys, os, pickle, openpyxl
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import oct_data
from pptx import Presentation
from pptx.oxml.ns import qn
SRC, OUT, FPK, MASTER, MC = sys.argv[1:6]
S, PL = oct_data.load(MASTER)
F = pickle.load(open(FPK, 'rb'))                       # 10M risk (non-RN AdTech) from УО, unchanged
wbm = openpyxl.load_workbook(MASTER, read_only=True, data_only=True)
RK = {i: [(x or 0) / 1e6 for x in r] for i, r in enumerate(wbm['10 PL'].iter_rows(min_row=72, max_row=83, min_col=5, max_col=16, values_only=True), 72)}
SG = {i: [x or 0 for x in r[2:14]] for i, r in enumerate(wbm['12 Сегменты'].iter_rows(min_row=59, max_row=67, max_col=16, values_only=True), 59)}
sm = lambda a, n: sum(a[:n])
res = {}
for per, n in (('10M', 10), ('YR', 12)):
    mplus = sm(PL[15]['fact'], n)
    base = {'T': sm(PL[16]['fact'], n) - mplus,
            'S': sum(sm(S[k]['fact'], n) for k in (13, 14, 17, 18)) - mplus,
            'N': sum(sm(S[k]['fact'], n) for k in (15, 16, 19, 20))}
    assert abs(base['S'] + base['N'] - base['T']) < 0.05
    rS, rN = F[('total', 'S')][1], F[('total', 'N')][1]
    if per == 'YR':
        fc = lambda a: sum(a[10:12])
        rS += sum(fc(SG[k]) for k in (60, 61, 64, 65)) - fc(RK[82]); rN += sum(fc(SG[k]) for k in (62, 63, 66, 67))
    risk = {'S': rS, 'N': rN, 'T': rS + rN}
    bu = lambda rows: sum((sum(PL[i]['bu'][:n]) if n == 10 else PL[i]['bu_year']) for i in rows)
    sg = lambda ks: sum((S[k]['ytd'][0] or 0) if n == 10 else sum(S[k]['bu']) for k in ks)
    BU = {'T': bu([16]) - bu([15]), 'S': sg((13, 14, 17, 18)), 'N': sg((15, 16, 19, 20))}
    rows = list(openpyxl.load_workbook(MC, read_only=True, data_only=True)['2026 Plan AdTech'].iter_rows(min_row=6, max_row=11, max_col=16, values_only=True))
    TG = {r[1].strip(): sum(float(x) for x in r[2:2 + n]) for r in rows}
    TS = TG['E-com'] + TG['Digital Instore'] + TG['Online+DMP'] + TG['DMP']; TN = TG['Внешние рекламодатели'] + TG['Монетизаторы e-com']
    TGT = {'S': TS, 'N': TN, 'T': TS + TN}
    res[per] = {g: (base[g], base[g] + risk[g], BU[g], TGT[g]) for g in 'TSN'}
    for g in 'TSN': print(per, g, tuple(round(x, 1) for x in res[per][g]))

fmt = lambda v: f'{round(v):,}'.replace(',', ' ')
def pct(a, b):
    p = round((a / b - 1) * 100)
    return ('0%' if p == 0 else f'{p:+d}%'), ('CCFFCC' if p > 5 else 'FFEB9C' if p >= -5 else 'FFC7CE')
def put(cell, text, fill=None):
    r = cell.text_frame.paragraphs[0].runs[0]; r.text = text
    r._r.find(qn('a:rPr')).set('lang', 'en-US')
    if fill:
        cell._tc.find(qn('a:tcPr')).find(qn('a:solidFill'))[0].set('val', fill)
p = Presentation(SRC)
log = []
for s, per in zip(p.slides, ('10M', 'YR')):
    t = [sh for sh in s.shapes if sh.has_table][0].table
    for ri, g in ((2, 'T'), (3, 'S'), (4, 'N')):
        f0, f1, bu, tg = res[per][g]
        old = [t.cell(ri, c).text for c in range(10)]
        assert round(bu) == int(old[3].replace(' ', '')) and round(tg) == int(old[4].replace(' ', '')), (per, g, old, bu, tg)
        put(t.cell(ri, 1), fmt(f0)); put(t.cell(ri, 2), fmt(f1))
        for c, (a, b) in zip((6, 7, 8, 9), ((f0, bu), (f1, bu), (f0, tg), (f1, tg))):
            put(t.cell(ri, c), *pct(a, b))
        log.append((per, old[0].strip('\xa0'), old[1:3] + old[6:10], [t.cell(ri, c).text for c in (1, 2, 6, 7, 8, 9)]))
p.save(OUT)
for l in log: print(l)
