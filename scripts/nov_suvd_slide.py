"""Slide 3 of KPI_Comm_102_1127: November 2026 forecast vs operations entered in СУВД (pivot by RN, 27.11) + ДО data for E-com."""
import sys, os, re
from lxml import etree
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import oct_data
from slide_dup import title_shape
from pptx import Presentation
from pptx.util import Emu
from pptx.oxml.ns import qn
SRC, OUT, MASTER = sys.argv[1:4]
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
# СУВД pivot 27.11 (screenshot from user), руб. без НДС
SUVD = {'М+': 48717135 + 123227080, 'AdTech': 235533379 + 210192297, 'E-com': 223529768 + 106285640}
assert sum(SUVD.values()) == 947485299 and SUVD['E-com'] == 329815408
SUVD = {k: v / 1e6 for k, v in SUVD.items()}
DO_ECOM = 311.0                                        # данные ДО по E-com (от пользователя)
S, PL = oct_data.load(MASTER)
f = lambda i, m: PL[i]['fact'][m]
NOV, OCT = 10, 9
fc = {'total': f(16, NOV) - f(15, NOV), 'ecom': S[10]['fact'][NOV], 'instore': f(10, NOV) + f(11, NOV), 'dmp': f(8, NOV), 'online': f(9, NOV), 'mplus': f(15, NOV)}
oc = {'total': f(16, OCT) - f(15, OCT), 'ecom': S[10]['fact'][OCT], 'instore': f(10, OCT) + f(11, OCT), 'dmp': f(8, OCT), 'online': f(9, OCT), 'mplus': f(15, OCT)}
for d in (fc, oc):
    d['nonecom'] = d['total'] - d['ecom']; d['other'] = d['nonecom'] - d['instore'] - d['dmp'] - d['online']
print({k: round(v, 1) for k, v in fc.items()}); print({k: round(v, 1) for k, v in oc.items()})
num = lambda v: f'{round(v):,}'.replace(',', ' ')
pc = lambda v: f'{round(v * 100)}%'
def ln(tag, spec):
    if spec is None:
        return f'<a:{tag} w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:{tag}>'
    w, col, dash = spec
    return (f'<a:{tag} w="{w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="{col}"/></a:solidFill>'
            f'<a:prstDash val="{dash}"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:{tag}>')
def tc(text='', sz=1200, b=False, col='000000', algn='ctr', fill=None, lnT=None, lnB=None, marL=73152):
    f_ = '<a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Arial" pitchFamily="34" charset="-122"/><a:cs typeface="Arial" pitchFamily="34" charset="-120"/>'
    lang = 'ru-RU' if re.search('[А-Яа-яЁё]', text) else 'en-US'
    run = (f'<a:r><a:rPr lang="{lang}" sz="{sz}" b="{int(b)}"><a:solidFill><a:srgbClr val="{col}"/></a:solidFill>{f_}</a:rPr>'
           f'<a:t xml:space="preserve">{text}</a:t></a:r>') if text else ''
    fl = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else ''
    return (f'<a:tc xmlns:a="{A}"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr marL="0" indent="0" algn="{algn}"><a:buNone/></a:pPr>{run}'
            f'<a:endParaRPr lang="ru-RU" sz="{sz}">{f_}</a:endParaRPr></a:p></a:txBody>'
            f'<a:tcPr marL="{marL}" marR="73152" marT="0" marB="0" anchor="ctr">{ln("lnL", None)}{ln("lnR", None)}{ln("lnT", lnT)}{ln("lnB", lnB)}{fl}</a:tcPr></a:tc>')
RED2, WH, DASH = ('25400', 'F04E37', 'solid'), ('25400', 'FFFFFF', 'solid'), ('9525', 'BFBFBF', 'sysDash')
LG, DG, GREY = 'F2F2F2', 'D8D8D8', '7F7F7F'
IND = '   '
TOT = SUVD['AdTech'] + SUVD['E-com'] + DO_ECOM
def row(label, f0, suvd, do, oct_, style):
    bold = style in ('total', 'main'); fill = DG if style == 'total' else LG; colr = GREY if style in ('sub', 'ref') else '000000'
    cell = lambda t, b=bold, c=colr: tc(t, 1200, b, c, fill=DG if style == 'total' else None, lnB=WH if style == 'total' else DASH)
    if suvd is None:
        vals = [num(f0), '-', '-', '-', '-', '-']
    else:
        tot = suvd + (do or 0)
        vals = [num(f0), num(suvd), num(do) if do else '-', num(tot), pc(tot / f0), num(f0 - tot)]
    return (380000, [tc(label, 1200, bold, colr, algn='l', fill=fill, lnB=WH, marL=155448)] + [cell(v) for v in vals] + [cell(num(oct_), False, GREY)])
HDR = ['Ноябрь 2026, млн руб.', 'Прогноз', 'СУВД на 27.11', 'Данные ДО', 'Итого заведено', '% прогноза', 'Не заведено до прогноза', 'Справочно: факт октябрь']
R = [(520000, [tc(h, 1200, True, algn='l' if i == 0 else 'ctr', lnB=RED2) for i, h in enumerate(HDR)])]
R.append(row('Общая выручка (без М+)', fc['total'], SUVD['AdTech'] + SUVD['E-com'], DO_ECOM, oc['total'], 'total'))
R.append(row('E-com', fc['ecom'], SUVD['E-com'], DO_ECOM, oc['ecom'], 'main'))
R.append(row('Не E-com: РН AdTech_Реклама и продвижение', fc['nonecom'], SUVD['AdTech'], None, oc['nonecom'], 'main'))
for lab, k in (('Instore digital', 'instore'), ('DMP', 'dmp'), ('Online+DMP', 'online'), ('Прочее', 'other')):
    R.append(row(IND + 'в т.ч. ' + lab if lab != 'Прочее' else IND + 'в т.ч. прочее', fc[k], None, None, oc[k], 'sub'))
R.append(row('Справочно: М+ (вне общей выручки)', fc['mplus'], SUVD['М+'], None, oc['mplus'], 'ref'))
W = [3300000, 1000000, 1150000, 1050000, 1150000, 1000000, 1350000]; W.append(11455400 - sum(W))
p = Presentation(SRC); s = p.slides[2]
old = [sh for sh in s.shapes if sh.has_table][0]
x, y = old.left, old.top
old._element.getparent().remove(old._element)
gf = s.shapes.add_table(len(R), len(W), Emu(368300), Emu(y), Emu(sum(W)), Emu(sum(h for h, _ in R))); gf.name = 'Table Nov plan vs FC'
tbl = gf._element.find(f'.//{{{A}}}tbl'); tp = tbl.find(f'{{{A}}}tblPr')
for a in list(tp.attrib): del tp.attrib[a]
for c in list(tp): tp.remove(c)
for g, w in zip(tbl.find(f'{{{A}}}tblGrid'), W): g.set('w', str(w))
for tr in tbl.findall(f'{{{A}}}tr'): tbl.remove(tr)
for h, cells in R:
    tr = etree.SubElement(tbl, f'{{{A}}}tr'); tr.set('h', str(h))
    for c in cells: tr.append(etree.fromstring(c))
def set_text(shape, text):
    para = shape.text_frame.paragraphs[0]
    for r in para.runs[1:]: r._r.getparent().remove(r._r)
    para.runs[0].text = text; para.runs[0]._r.find(qn('a:rPr')).set('lang', 'ru-RU')
    for extra in shape.text_frame.paragraphs[1:]: extra._p.getparent().remove(extra._p)
tot_fc = fc['total']
set_text(title_shape(s), f'Прогноз по заведённым операциям ноябрь. В СУВД и ДО на 27.11 = {num(TOT)} млн руб. = {pc(TOT / tot_fc)} прогноза')
sub = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.startswith('Операции в СУВД')][0]
set_text(sub, 'Операции в СУВД и ДО против прогноза')
p.save(OUT)
print('TOT', round(TOT, 1), pc(TOT / tot_fc))
