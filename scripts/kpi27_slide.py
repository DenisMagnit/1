import sys
from lxml import etree
from pptx import Presentation
from pptx.util import Emu
import os; sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from slide_dup import dup_slide, title_shape

SRC, OUT = sys.argv[1:3]
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
p = Presentation(SRC)
src = [s for s in p.slides if title_shape(s).text_frame.text.startswith('Управленческий P&L 2027')][0]
s = dup_slide(p, src)
for sh in list(s.shapes):
    if sh.has_table or sh.name == 'Shape 201':
        sh._element.getparent().remove(sh._element)
tp = title_shape(s).text_frame.paragraphs[0]
for r in tp.runs[1:]:
    r._r.getparent().remove(r._r)
tp.runs[0].text = 'Корпоративные КПЭ 2027 (AdTech)'

def ln(tag, spec):
    if spec is None:
        return f'<a:{tag} w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:{tag}>'
    w, col, dash = spec
    return (f'<a:{tag} w="{w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="{col}"/></a:solidFill>'
            f'<a:prstDash val="{dash}"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:{tag}>')

def tc(text='', sz=1300, b=False, col='000000', algn='ctr', fill=None, lnT=None, lnB=None, marL=73152, span=None):
    f = '<a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Arial" pitchFamily="34" charset="-122"/><a:cs typeface="Arial" pitchFamily="34" charset="-120"/>'
    run = (f'<a:r><a:rPr lang="ru-RU" sz="{sz}" b="{int(b)}"><a:solidFill><a:srgbClr val="{col}"/></a:solidFill>{f}</a:rPr>'
           f'<a:t>{text}</a:t></a:r>') if text else ''
    fl = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else ''
    sp = f' gridSpan="{span}"' if span else ''
    return (f'<a:tc xmlns:a="{A}"{sp}><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr marL="0" indent="0" algn="{algn}"><a:buNone/></a:pPr>{run}'
            f'<a:endParaRPr lang="ru-RU" sz="{sz}">{f}</a:endParaRPr></a:p></a:txBody>'
            f'<a:tcPr marL="{marL}" marR="73152" marT="0" marB="0" anchor="ctr">{ln("lnL", None)}{ln("lnR", None)}{ln("lnT", lnT)}{ln("lnB", lnB)}{fl}</a:tcPr></a:tc>')

def hmerge():
    return f'<a:tc xmlns:a="{A}" hMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ru-RU"/></a:p></a:txBody><a:tcPr/></a:tc>'

RED1, RED2 = ('28575', 'FF452B', 'solid'), ('25400', 'F04E37', 'solid')
WH, DASH = ('25400', 'FFFFFF', 'solid'), ('9525', 'BFBFBF', 'sysDash')
GREY, LG = 'A6A6A6', 'F2F2F2'
# columns: КПЭ, Вес, g, 3 знач, g, 3 % от цели, g, 3 оценка, g, Бюджет, Комментарий
W = [2300000, 650000, 80000] + [700000] * 3 + [80000] + [650000] * 3 + [80000] + [650000] * 3 + [80000, 850000]
LEFT, WIDTH = 368300, 11455400
W.append(WIDTH - sum(W))
G = {2, 6, 10, 14}
KPI = [  # из финальной дэки BU27 (слайд «Корпоративные КПЭ на 2027 год – предложение»), без FCF
    ('Выручка, млрд руб.', '25%', ('11,7', '14,6', '17,5'), '14,6'),
    ('EBITDA, млрд руб.', '35%', ('8,7', '10,8', '13,0'), '10,8'),
    ('Деньги от новых инициатив, млрд руб.', '15%', ('1,8', '2,3', '2,7'), '2,3'),
    ('Прирост инвентаря, %', '10%', ('+24%', '+30%', '+36%'), '+30%'),
]
PCT = ('80%', '100%', '120%')       # % выполнения бюджета — из финальной дэки
SCORE = ('50%', '100%', '150%')     # оценка КПЭ — шкала ОМНИ

rows = []
r0 = []
for i in range(len(W)):
    if i in (3, 7, 11):
        r0.append(tc({3: 'Значение КПЭ', 7: '% от целевого', 11: 'Оценка КПЭ'}[i], 1400, True, lnB=RED1, span=3))
    elif i in (4, 5, 8, 9, 12, 13):
        r0.append(hmerge())
    else:
        r0.append(tc())
rows.append((400000, r0))
hdr = {0: 'КПЭ', 1: 'Вес', 15: 'Бюджет', 16: 'Комментарий'}
r1 = []
for i in range(len(W)):
    if i in G:
        r1.append(tc())
    elif i in hdr:
        r1.append(tc(hdr[i], 1300, True, algn='l' if i in (0, 16) else 'ctr', lnB=RED2))
    else:
        r1.append(tc(('Мин', 'Цель', 'Макс')[(i - 3) % 4], 1300, True, lnT=RED1, lnB=RED2))
rows.append((400000, r1))
for k, (name, wt, vals, bud) in enumerate(KPI):
    lt = RED2 if k == 0 else WH
    r = [tc(name, 1300, True, algn='l', fill=LG, lnT=lt, lnB=WH, marL=155448),
         tc(wt, 1300, True, fill=LG, lnT=lt, lnB=WH), tc()]
    dt = RED2 if k == 0 else DASH
    r += [tc(v, 1300, True, lnT=dt, lnB=DASH) for v in vals] + [tc()]
    r += [tc(v, 1300, False, GREY, lnT=dt, lnB=DASH) for v in PCT] + [tc()]
    r += [tc(v, 1300, False, GREY, lnT=dt, lnB=DASH) for v in SCORE] + [tc()]
    r += [tc(bud, 1300, True, lnT=dt, lnB=DASH), tc('', 1200, algn='l', lnT=dt, lnB=DASH)]
    rows.append((700000, r))

gf = s.shapes.add_table(len(rows), len(W), Emu(LEFT), Emu(1041400), Emu(WIDTH), Emu(sum(h for h, _ in rows)))
gf.name = 'Table KPI 2027'
tbl = gf._element.find(f'.//{{{A}}}tbl')
tp_ = tbl.find(f'{{{A}}}tblPr')
for a in list(tp_.attrib): del tp_.attrib[a]
for c in list(tp_): tp_.remove(c)
for gc, w in zip(tbl.find(f'{{{A}}}tblGrid'), W):
    gc.set('w', str(w))
for tr in tbl.findall(f'{{{A}}}tr'):
    tbl.remove(tr)
for h, cells in rows:
    tr = etree.SubElement(tbl, f'{{{A}}}tr'); tr.set('h', str(h))
    for c in cells:
        tr.append(etree.fromstring(c))
p.save(OUT)
print('saved', OUT, len(p.slides), [i for i, x in enumerate(p.slides, 1) if x is s or x.part is s.part])
