"""AT FIN: new slide 'key commerce KPI metrics' (10M 2026), after the KPI slide."""
import sys, copy, pickle
sys.path.insert(0, '/tmp/claude-0/-home-user-1/201a7458-fa29-5237-8125-45e111dcdac4/scratchpad/scripts')
import openpyxl
from pptx import Presentation
from pptx.util import Pt, Inches, Emu
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from oct_data import load
from oct_lib import num, pct, set_cell, rebuild_rows
from slide_dup import dup_slide, set_title, title_shape

SRC, OUT, FPK, MASTER, MC = sys.argv[1:6]
YR = len(sys.argv) > 6 and sys.argv[6] == 'yr'
F = pickle.load(open(FPK, 'rb'))
S, PL = load(MASTER)
bu10 = lambda rows: sum(sum(PL[i]['bu'][:10]) for i in rows)
seg10 = lambda ks: sum(S[k]['ytd'][0] or 0 for k in ks)
BU = {('total', 'T'): bu10([16]) - bu10([15]), ('total', 'S'): seg10((13, 14, 17, 18)), ('total', 'N'): seg10((15, 16, 19, 20)),
      ('ecom', 'T'): S[10]['ytd'][0], ('ecom', 'S'): seg10((13, 14)), ('ecom', 'N'): seg10((15, 16)),
      ('instore', 'T'): bu10([10, 11]), ('dmp', 'T'): bu10([8]), ('online', 'T'): bu10([9])}
assert abs(BU[('total', 'S')] + BU[('total', 'N')] - BU[('total', 'T')]) < 0.1
assert abs(BU[('ecom', 'S')] + BU[('ecom', 'N')] - BU[('ecom', 'T')]) < 0.1
if YR:
    buy = lambda rows: sum(PL[i]['bu_year'] for i in rows)
    sgy = lambda ks: sum(sum(S[k]['bu']) for k in ks)
    BU = {('total', 'T'): buy([16]) - buy([15]), ('total', 'S'): sgy((13, 14, 17, 18)), ('total', 'N'): sgy((15, 16, 19, 20)),
          ('ecom', 'T'): sum(S[10]['bu']), ('ecom', 'S'): sgy((13, 14)), ('ecom', 'N'): sgy((15, 16)),
          ('instore', 'T'): buy([10, 11]), ('dmp', 'T'): buy([8]), ('online', 'T'): buy([9])}
    assert abs(BU[('total', 'S')] + BU[('total', 'N')] - BU[('total', 'T')]) < 0.1
    wbm = openpyxl.load_workbook(MASTER, read_only=True, data_only=True)
    RK = {i: [(x or 0) / 1e6 for x in r] for i, r in enumerate(wbm['10 PL'].iter_rows(min_row=72, max_row=83, min_col=5, max_col=16, values_only=True), 72)}
    SG = {i: [x or 0 for x in r[2:14]] for i, r in enumerate(wbm['12 Сегменты'].iter_rows(min_row=59, max_row=67, max_col=16, values_only=True), 59)}
    fc = lambda a: sum(a[10:12])
    FC = {('total', 'T'): (fc(PL[16]['fact']) - fc(PL[15]['fact']), sum(fc(RK[k]) for k in RK) - fc(RK[82])),
          ('total', 'S'): (sum(fc(S[k]['fact']) for k in (13, 14, 17, 18)) - fc(PL[15]['fact']), sum(fc(SG[k]) for k in (60, 61, 64, 65)) - fc(RK[82])),
          ('total', 'N'): (sum(fc(S[k]['fact']) for k in (15, 16, 19, 20)), sum(fc(SG[k]) for k in (62, 63, 66, 67))),
          ('ecom', 'T'): (fc(S[10]['fact']), fc(RK[72]))}
    es = sum(fc(SG[k]) for k in (60, 61)); en = sum(fc(SG[k]) for k in (62, 63))
    FC[('ecom', 'S')] = (sum(fc(S[k]['fact']) for k in (13, 14)), fc(RK[72]) * es / (es + en))
    FC[('ecom', 'N')] = (sum(fc(S[k]['fact']) for k in (15, 16)), fc(RK[72]) * en / (es + en))
    for k, pr, rr in (('instore', [10, 11], [77, 78]), ('dmp', [8], [75]), ('online', [9], [76])):
        tb, tr_ = sum(fc(PL[i]['fact']) for i in pr), sum(fc(RK[i]) for i in rr)
        FC[(k, 'T')] = (tb, tr_)
        for sg in 'SN':
            b0, r0 = F[(k, sg)]; B0, R0 = F[(k, 'T')]
            FC[(k, sg)] = (tb * b0 / B0 if B0 else 0, tr_ * r0 / R0 if R0 else 0)
    for k in ('total', 'ecom'):
        for i in range(2): assert abs(FC[(k, 'S')][i] + FC[(k, 'N')][i] - FC[(k, 'T')][i]) < 0.05, (k, i)
    F = {key: (F[key][0] + FC[key][0], F[key][1] + FC[key][1]) for key in F}
    print('FC', {k: tuple(round(x, 1) for x in v) for k, v in FC.items()})
rows = list(openpyxl.load_workbook(MC, read_only=True, data_only=True)['2026 Plan AdTech'].iter_rows(min_row=6, max_row=11, max_col=16, values_only=True))
TG = {r[1].strip(): sum(float(x) for x in r[2:(14 if YR else 12)]) for r in rows}           # Jan..Oct or year
assert set(TG) == {'E-com', 'Digital Instore', 'Online+DMP', 'DMP', 'Внешние рекламодатели', 'Монетизаторы e-com'}, TG
TGT = {('total', 'S'): TG['E-com'] + TG['Digital Instore'] + TG['Online+DMP'] + TG['DMP'],
       ('total', 'N'): TG['Внешние рекламодатели'] + TG['Монетизаторы e-com'],
       ('ecom', 'S'): TG['E-com'], ('instore', 'S'): TG['Digital Instore'], ('dmp', 'S'): TG['DMP'], ('online', 'S'): TG['Online+DMP']}
TGT[('total', 'T')] = TGT[('total', 'S')] + TGT[('total', 'N')]

NAMES = [('total', 'Общая выручка'), ('ecom', 'РН E-com'), ('instore', 'Instore digital'), ('dmp', 'DMP'), ('online', 'Online+DMP')]
SEGN = {'T': None, 'S': '   Поставщики', 'N': '   Не поставщики'}

p = Presentation(SRC)
sl = list(p.slides)
src = [s for s in sl if title_shape(s).text_frame.text.strip() == 'Выручка по направлениям: прогноз Ноя-Дек и 2026'][0]
kpi_i = [i for i, s in enumerate(sl) if (title_shape(s).text_frame.text.strip() == 'Ключевые метрики КПЭ коммерции: 10М 2026' if YR else title_shape(s).text_frame.text.strip().startswith('Статус выполнения КПЭ'))][0]
s = dup_slide(p, src)
set_title(s, 'Ключевые метрики КПЭ коммерции: 10+2 2026' if YR else 'Ключевые метрики КПЭ коммерции: 10М 2026')
# move right after the KPI slide
lst = p.slides._sldIdLst
new_i = [i for i, x in enumerate(p.slides) if x.part is s.part][0]
item = list(lst)[new_i]; lst.remove(item); lst.insert(kpi_i + 1, item)
# drop the lower table
low = [sh for sh in s.shapes if sh.has_table and sh.name == 'Revenue forecast by direction'][0]
low._element.getparent().remove(low._element)
gf = [sh for sh in s.shapes if sh.has_table and sh.name == 'Google Shape;326;p48'][0]
gf.name = 'Commerce KPI 10M'
t0 = gf.table
assert t0.cell(12, 0).text_frame.text == 'Gross Revenue' and t0.cell(1, 0).text_frame.text == 'APP' and t0.cell(2, 0).text_frame.text == 'MIP'
A_ = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
_trs = t0._tbl.findall(A_ + 'tr')
_tc = lambda r, c: _trs[r].findall(A_ + 'tc')[c].find(A_ + 'tcPr')
LINE = dict(hdr=copy.deepcopy(_tc(0, 1).find(A_ + 'lnB')), dot=copy.deepcopy(_tc(2, 1).find(A_ + 'lnB')),
            white=copy.deepcopy(_tc(2, 0).find(A_ + 'lnB')))
spec = [(0, ['10+2 2026, млн руб.', 'Факт + прогноз', '', 'План', '', '', '10+2 vs BU', '', '10+2 vs таргет', '']) if YR else
        (0, ['10М 2026, млн руб.', 'Факт', '', 'План', '', '', 'Факт vs BU', '', 'Факт vs таргет', '']),
        (0, ['', 'без не РН AdTech', 'с не РН AdTech', 'BU', 'Таргет', '', 'без не РН AdTech', 'с не РН AdTech', 'без не РН AdTech', 'с не РН AdTech'])]
out_rows = []
for k, name in NAMES:
    for seg in 'TSN':
        base, risk = F[(k, seg)]; withr = base + risk
        bu = BU.get((k, seg)); tg = TGT.get((k, seg))
        cells = [name if seg == 'T' else SEGN[seg], num(base, ''), num(withr, ''),
                 num(bu, '') if bu is not None else '-', num(tg, '') if tg is not None else '-', '',
                 pct(base - bu, bu, plus=True) if bu is not None else '-', pct(withr - bu, bu, plus=True) if bu is not None else '-',
                 pct(base - tg, tg, plus=True) if tg is not None else '-', pct(withr - tg, tg, plus=True) if tg is not None else '-']
        spec.append((12 if seg == 'T' else 2, cells))
        out_rows.append((name, seg, round(base, 1), round(withr, 1), bu and round(bu, 1), tg and round(tg, 1)))
rebuild_rows(gf, spec)
t = gf.table
for c0, c1 in ((1, 2), (3, 4), (6, 7), (8, 9)):
    t.cell(0, c0).merge(t.cell(0, c1))
t.cell(0, 0).merge(t.cell(1, 0))
for r in range(len(t.rows)):
    for c in range(len(t.columns)):
        tf = t.cell(r, c).text_frame
        if not tf.text.strip():
            for para in tf.paragraphs:
                for run in list(para.runs): run._r.getparent().remove(run._r)
# fit height: 17 rows into the space of the original table area
H = Inches(5.5)
rh = int(H / len(t.rows))
for tr in t._tbl.findall(qn('a:tr')): tr.set('h', str(rh))
gf.height = rh * len(t.rows)
# ---------------- borders: no vertical lines; header line under the header; dotted between data rows
from lxml import etree
SIDES = ('lnL', 'lnR', 'lnT', 'lnB')
ORDER = ['lnL', 'lnR', 'lnT', 'lnB', 'lnTlToBr', 'lnBlToTr', 'cell3D', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill', 'headers', 'extLst']
def set_line(tc, side, kind):
    pr = tc.find(A_ + 'tcPr')
    old = pr.find(A_ + side)
    if kind == 'none':
        el = etree.SubElement(pr, A_ + side); pr.remove(el)
        el = etree.Element(A_ + side, w='0'); etree.SubElement(el, A_ + 'noFill')
    else:
        el = copy.deepcopy(LINE[kind]); el.tag = A_ + side
    if old is not None:
        old.addprevious(el); pr.remove(old)
    else:
        idx = ORDER.index(side)
        nxt = [c for c in pr if c.tag.split('}')[1] in ORDER[idx + 1:]]
        if nxt: nxt[0].addprevious(el)
        else: pr.append(el)
trs = t._tbl.findall(A_ + 'tr')
NR, NC, SP = len(trs), len(t.columns), 5
cell = lambda r, c: trs[r].findall(A_ + 'tc')[c]
kinds = [s_[0] for s_ in spec]             # template index per row: 0 header, 12 total, 2 sub-row
for r in range(NR):
    for c in range(NC):
        tc = cell(r, c)
        set_line(tc, 'lnL', 'none'); set_line(tc, 'lnR', 'none')
        if c == SP:
            set_line(tc, 'lnT', 'none'); set_line(tc, 'lnB', 'none'); continue
        if r == 0:
            set_line(tc, 'lnT', 'none'); set_line(tc, 'lnB', 'none' if c == 0 else 'hdr')
        elif r == 1:
            set_line(tc, 'lnT', 'none' if c == 0 else 'hdr'); set_line(tc, 'lnB', 'hdr')
        else:
            inner = 'white' if (c == 0 and kinds[r] == 2) else 'dot'
            set_line(tc, 'lnT', 'hdr' if r == 2 else ('dot' if c else ('white' if kinds[r] == 2 and kinds[r - 1] == 2 else 'dot')))
            set_line(tc, 'lnB', inner if r < NR - 1 else 'dot')
            if c == 0 and r < NR - 1 and kinds[r] == 2 and kinds[r + 1] != 2:
                set_line(tc, 'lnB', 'dot')
# ---------------- colour indication: cell fill in deviation columns (>+5% green, -5..+5% yellow, <-5% red)
FILL = {'g': RGBColor(0xCC, 0xFF, 0xCC), 'y': RGBColor(0xFF, 0xEB, 0x9C), 'r': RGBColor(0xFF, 0xC7, 0xCE)}
for r in range(2, NR):
    for c in (6, 7, 8, 9):
        tf = t.cell(r, c).text_frame
        txt = tf.text.strip()
        if not txt.endswith('%'): continue
        v = int(txt.replace('%', '').replace('+', '').replace('\u00a0', '').replace(' ', ''))
        k = 'g' if v > 5 else ('r' if v < -5 else 'y')
        t.cell(r, c).fill.solid(); t.cell(r, c).fill.fore_color.rgb = FILL[k]
# proofing language per run: Cyrillic -> ru-RU, otherwise en-US
import re as _re
for shp in [gf, title_shape(s)]:
    paras = [pa for r_ in (gf.table.rows if shp is gf else []) for c_ in r_.cells for pa in c_.text_frame.paragraphs] if shp is gf else list(shp.text_frame.paragraphs)
    for pa in paras:
        for run in pa.runs:
            run._r.get_or_add_rPr().set('lang', 'ru-RU' if _re.search('[А-Яа-яЁё]', run.text) else 'en-US')
# checks
for slide in p.slides:
    for g in [sh for sh in slide.shapes if sh.has_table]:
        tbl = g.table._tbl
        ncol = len(tbl.find(qn('a:tblGrid')).findall(qn('a:gridCol')))
        for tr in tbl.findall(qn('a:tr')):
            k = [c.tag.split('}')[1] for c in tr]
            assert k.count('tc') == ncol and not ('extLst' in k and 'tc' in k[k.index('extLst'):])
p.save(OUT)
for r in out_rows: print(r)
print('saved', OUT, len(p.slides))
