"""Slide: E-com November 2026 by operations entered in the system as planned (export 27.11) vs master forecast."""
import sys, os, re, copy, collections, openpyxl
from lxml import etree
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from slide_dup import dup_slide, title_shape
import oct_data
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.oxml.ns import qn
SRC, OUT, PLAN, MASTER = sys.argv[1:5]
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
# ---- data
rows = [r for r in list(openpyxl.load_workbook(PLAN, read_only=True, data_only=True).worksheets[0].iter_rows(values_only=True))[3:] if r[5] is not None]
assert {r[3] for r in rows} <= {'E-Commerce_Adtech', 'E-Commerce_Adtech (РА)'} and {r[1] for r in rows} == {'БОНУС', 'ШС'}
kind = lambda r: 'ШС' if r[1] == 'ШС' else ('Бонус через РА' if r[3].endswith('(РА)') else 'Бонус')
amt = collections.defaultdict(float); cnt = collections.Counter(); wk = collections.defaultdict(float); wd = {}
for r in rows:
    k = kind(r); v = r[5] / 1e6; amt[k] += v; cnt[k] += 1
    w = r[2].isocalendar()[1]; wk[(w, 'ШС' if k == 'ШС' else 'Бонус')] += v
    wd.setdefault(w, [r[2], r[2]]); wd[w][0] = min(wd[w][0], r[2]); wd[w][1] = max(wd[w][1], r[2])
TOT = sum(amt.values()); N = sum(cnt.values())
S, PL = oct_data.load(MASTER)
FC_NOV, ACT_OCT = S[10]['fact'][10], S[10]['fact'][9]
print('total', round(TOT, 1), N, 'FC Nov', round(FC_NOV, 1), 'Oct', round(ACT_OCT, 1), {k: round(v, 1) for k, v in amt.items()})
last = max(r[2] for r in rows).strftime('%d.%m')
# ---- cell builder (same style as KPI slide)
def ln(tag, spec):
    if spec is None:
        return f'<a:{tag} w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:{tag}>'
    w, col, dash = spec
    return (f'<a:{tag} w="{w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="{col}"/></a:solidFill>'
            f'<a:prstDash val="{dash}"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:{tag}>')
def tc(text='', sz=1300, b=False, col='000000', algn='ctr', fill=None, lnT=None, lnB=None, marL=73152):
    f = '<a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Arial" pitchFamily="34" charset="-122"/><a:cs typeface="Arial" pitchFamily="34" charset="-120"/>'
    lang = 'ru-RU' if re.search('[А-Яа-яЁё]', text) else 'en-US'
    run = (f'<a:r><a:rPr lang="{lang}" sz="{sz}" b="{int(b)}"><a:solidFill><a:srgbClr val="{col}"/></a:solidFill>{f}</a:rPr>'
           f'<a:t xml:space="preserve">{text}</a:t></a:r>') if text else ''
    fl = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else ''
    return (f'<a:tc xmlns:a="{A}"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr marL="0" indent="0" algn="{algn}"><a:buNone/></a:pPr>{run}'
            f'<a:endParaRPr lang="ru-RU" sz="{sz}">{f}</a:endParaRPr></a:p></a:txBody>'
            f'<a:tcPr marL="{marL}" marR="73152" marT="0" marB="0" anchor="ctr">{ln("lnL", None)}{ln("lnR", None)}{ln("lnT", lnT)}{ln("lnB", lnB)}{fl}</a:tcPr></a:tc>')
RED2, WH, DASH = ('25400', 'F04E37', 'solid'), ('25400', 'FFFFFF', 'solid'), ('9525', 'BFBFBF', 'sysDash')
LG, GREY = 'F2F2F2', '7F7F7F'
num = lambda v: f'{round(v):,}'.replace(',', ' ')
pc = lambda v: f'{round(v * 100)}%'
def add_table(s, x, y, widths, rows_, name):
    gf = s.shapes.add_table(len(rows_), len(widths), Emu(x), Emu(y), Emu(sum(widths)), Emu(sum(h for h, _ in rows_)))
    gf.name = name
    tbl = gf._element.find(f'.//{{{A}}}tbl'); tp = tbl.find(f'{{{A}}}tblPr')
    for a in list(tp.attrib): del tp.attrib[a]
    for c in list(tp): tp.remove(c)
    for gc, w in zip(tbl.find(f'{{{A}}}tblGrid'), widths): gc.set('w', str(w))
    for tr in tbl.findall(f'{{{A}}}tr'): tbl.remove(tr)
    for h, cells in rows_:
        tr = etree.SubElement(tbl, f'{{{A}}}tr'); tr.set('h', str(h))
        for c in cells: tr.append(etree.fromstring(c))
def add_text(s, x, y, w, h, text, sz=14, b=True, col='000000'):
    tb = s.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h)); tf = tb.text_frame; tf.word_wrap = True
    for side in ('left', 'right', 'top', 'bottom'): setattr(tf, f'margin_{side}', 0)
    r = tf.paragraphs[0].add_run(); r.text = text; r.font.size = Pt(sz); r.font.bold = b; r.font.name = 'Arial'
    from pptx.dml.color import RGBColor; r.font.color.rgb = RGBColor.from_string(col)
    r._r.find(qn('a:rPr')).set('lang', 'ru-RU' if re.search('[А-Яа-яЁё]', text) else 'en-US')
    return tb
# ---- slide
p = Presentation(SRC)
src = p.slides[0]
s = dup_slide(p, src)
for sh in list(s.shapes):
    if sh.has_table: sh._element.getparent().remove(sh._element)
tp = title_shape(s).text_frame.paragraphs[0]
for r in tp.runs[1:]: r._r.getparent().remove(r._r)
tp.runs[0].text = f'E-com ноябрь 2026: в системе заведено {num(TOT)} млн руб. – {pc(TOT / FC_NOV)} прогноза'
tp.runs[0]._r.find(qn('a:rPr')).set('lang', 'ru-RU')
X0, X1, Y0 = 368300, 6400000, 1041400
add_text(s, X0, Y0, 5600000, 300000, 'Операции в системе против прогноза')
add_text(s, X1, Y0, 5423700, 300000, 'Заведение операций по неделям, млн руб.')
# left table
W1 = [2700000, 900000, 1000000, 1000000]
H = 380000
L = [(H, [tc('E-com ноябрь 2026, млн руб.', 1200, True, algn='l', lnB=RED2), tc('Сумма', 1200, True, lnB=RED2),
          tc('% прогноза', 1200, True, lnB=RED2), tc('Кол-во, шт.', 1200, True, lnB=RED2)])]
L.append((H, [tc('Прогноз мастера на ноябрь', 1300, False, algn='l', fill=LG, lnB=WH, marL=155448), tc(num(FC_NOV), 1300, lnB=DASH), tc('100%', 1300, col=GREY, lnB=DASH), tc('-', 1300, col=GREY, lnB=DASH)]))
L.append((H, [tc(f'Заведено в системе на {last}', 1300, True, algn='l', fill='D8D8D8', lnB=WH, marL=155448), tc(num(TOT), 1300, True, fill='D8D8D8', lnB=WH), tc(pc(TOT / FC_NOV), 1300, True, fill='D8D8D8', lnB=WH), tc(str(N), 1300, True, fill='D8D8D8', lnB=WH)]))
for k in ('Бонус', 'Бонус через РА', 'ШС'):
    L.append((H, [tc('   ' + k, 1300, False, algn='l', fill=LG, lnB=WH, marL=155448), tc(num(amt[k]), 1300, lnB=DASH), tc(pc(amt[k] / FC_NOV), 1300, col=GREY, lnB=DASH), tc(str(cnt[k]), 1300, lnB=DASH)]))
L.append((H, [tc('Ещё не заведено до прогноза', 1300, True, algn='l', fill=LG, lnB=WH, marL=155448), tc(num(FC_NOV - TOT), 1300, True, lnB=DASH), tc(pc((FC_NOV - TOT) / FC_NOV), 1300, True, lnB=DASH), tc('-', 1300, col=GREY, lnB=DASH)]))
L.append((H, [tc('Справочно: факт октябрь', 1300, False, GREY, algn='l', fill=LG, lnB=WH, marL=155448), tc(num(ACT_OCT), 1300, False, GREY, lnB=DASH), tc('-', 1300, col=GREY, lnB=DASH), tc('-', 1300, col=GREY, lnB=DASH)]))
add_table(s, X0, Y0 + 420000, W1, L, 'Table E-com plan vs FC')
# right table
W2 = [800000, 1300000, 750000, 700000, 750000, 1123700]
R = [(H, [tc('Неделя', 1200, True, lnB=RED2), tc('Даты создания', 1200, True, lnB=RED2), tc('Бонус', 1200, True, lnB=RED2),
          tc('ШС', 1200, True, lnB=RED2), tc('Итого', 1200, True, lnB=RED2), tc('Накоплено', 1200, True, lnB=RED2)])]
acc = 0; tb = ts = 0
for w in sorted(wd):
    b_, s_ = wk[(w, 'Бонус')], wk[(w, 'ШС')]; acc += b_ + s_; tb += b_; ts += s_
    R.append((H, [tc(str(w), 1300, fill=LG, lnB=WH), tc(f"{wd[w][0]:%d.%m}–{wd[w][1]:%d.%m}", 1300, lnB=DASH), tc(num(b_), 1300, lnB=DASH),
                  tc(num(s_), 1300, lnB=DASH), tc(num(b_ + s_), 1300, True, lnB=DASH), tc(num(acc), 1300, col=GREY, lnB=DASH)]))
R.append((H, [tc('Итого', 1300, True, fill='D8D8D8', lnB=WH), tc('', 1300, fill='D8D8D8', lnB=WH), tc(num(tb), 1300, True, fill='D8D8D8', lnB=WH),
              tc(num(ts), 1300, True, fill='D8D8D8', lnB=WH), tc(num(tb + ts), 1300, True, fill='D8D8D8', lnB=WH), tc('', 1300, fill='D8D8D8', lnB=WH)]))
add_table(s, X1, Y0 + 420000, W2, R, 'Table E-com plan by week')
assert abs(acc - TOT) < 1e-6
weekly = ' → '.join(num(wk[(w, 'Бонус')] + wk[(w, 'ШС')]) for w in sorted(wd))
add_text(s, X1, Y0 + 420000 + H * len(R) + 250000, 5423700, 600000,
         f'Темп заведения растёт: {weekly} млн руб. в неделю. Сверху добавятся операции, которые ещё не заведены в систему', 12, False)
add_text(s, 2019516, 6238902, 8800000, 400000,
         f'Операции на РН E-Commerce_Adtech и E-Commerce_Adtech (РА), заведённые в систему как плановые, по дате создания {wd[min(wd)][0]:%d.%m}–{last}.2026 (выгрузка {last}). '
         'Прогноз мастера – E-comm Revenue (APP) на ноябрь, включая непоставщиков. ШС – как в выгрузке', 9, False, GREY)
p.save(OUT)
print('saved', OUT, len(p.slides))
