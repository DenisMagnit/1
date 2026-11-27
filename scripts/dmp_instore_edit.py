"""AdTech_DMP_to_Instore: exec cards (+DPP, +refusal threshold), DMP margin footnote, argument 1, chart title (sl.3),
titles of sl.8-9, conservative-assumption note (sl.4), spelling language on edited slides."""
import sys, re, copy
from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
SRC, OUT = sys.argv[1:3]
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
p = Presentation(SRC)
sl = list(p.slides)
shp = lambda s, name: [x for x in s.shapes if x.name == name][0]

def run_xml(text, sz, bold, color, baseline=None):
    lang = 'ru-RU' if re.search('[А-Яа-яЁё]', text) else 'en-US'
    bl = f' baseline="{baseline}"' if baseline else ''
    return (f'<a:r xmlns:a="{A}"><a:rPr lang="{lang}" sz="{sz}" b="{int(bold)}"{bl} dirty="0"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:rPr>'
            f'<a:t xml:space="preserve">{text}</a:t></a:r>')
def para(runs, algn='ctr', spc=None):
    sp = f'<a:spcBef><a:spcPts val="{spc}"/></a:spcBef>' if spc else ''
    return etree.fromstring(f'<a:p xmlns:a="{A}"><a:pPr algn="{algn}">{sp}<a:buNone/></a:pPr>{"".join(runs)}</a:p>')
def set_paras(shape, paras):
    tb = shape.text_frame._txBody
    for p_ in tb.findall(qn('a:p')): tb.remove(p_)
    for p_ in paras: tb.append(p_)
def set_title(s, text):
    ph = [x for x in s.placeholders if x.placeholder_format.idx == 10][0]
    pr = ph.text_frame.paragraphs[0]
    for r in pr.runs[1:]: r._r.getparent().remove(r._r)
    pr.runs[0].text = text
def set_footnote(s, lines):
    ph = [x for x in s.placeholders if x.placeholder_format.idx == 17][0]
    tb = ph.text_frame._txBody
    for p_ in tb.findall(qn('a:p')): tb.remove(p_)
    for ln in lines:
        runs = []
        for txt, sup in ln:
            lang = 'ru-RU' if re.search('[А-Яа-яЁё]', txt) else 'en-US'
            bl = ' baseline="30000"' if sup else ''
            runs.append(f'<a:r xmlns:a="{A}"><a:rPr lang="{lang}"{bl} dirty="0"/><a:t xml:space="preserve">{txt}</a:t></a:r>')
        tb.append(etree.fromstring(f'<a:p xmlns:a="{A}">{"".join(runs)}</a:p>'))

# ---------- slide 2: exec
s = sl[1]
k1, k4, k5 = shp(s, 'KPI 1'), shp(s, 'KPI 4'), shp(s, 'KPI 5')
L, R = k1.left, k5.left + k5.width
GAP = 182880; W = (R - L - 4 * GAP) // 5; H = k1.height + 91440
maxid = max(int(e.get('id')) for e in s._element.iter(qn('p:cNvPr')))
cards = [k1, k4, k5]
for i in range(2):
    el = copy.deepcopy(k1._element); maxid += 1
    el.find('.//' + qn('p:cNvPr')).set('id', str(maxid)); el.find('.//' + qn('p:cNvPr')).set('name', f'KPI {6 + i}')
    k5._element.addnext(el) if i == 0 else cards[-1]._element.addnext(el)
    cards.append([x for x in s.shapes if x.name == f'KPI {6 + i}'][0])
GREY, RED = '595959', 'FF452C'
val_col = lambda shape: shape.text_frame.paragraphs[1].runs[0]._r.find(qn('a:rPr')).find(qn('a:solidFill'))[0].get('val')
spec = [  # label, value, unit, sub, value colour
    ('NPV VS DMP', '+135,5', ' млн', 'при WACC 20%', val_col(k1)),
    ('EBITDA VS DMP', '+1 035,3', ' млн', '2027–2032', val_col(k4)),
    ('CAPEX', '-404,7', ' млн', '90% – старт 2027', val_col(k5)),
    ('Окупаемость, дисконт.', '4,0', ' года', 'начало 2031', RED),
    ('Порог отказа КА', '41,7%', '', 'от переноса КУ, база 15%', RED),
]
for i, (c, (lab, v, u, sub, col)) in enumerate(zip(cards, spec)):
    c.left, c.top, c.width, c.height = L + i * (W + GAP), k1.top, W, H
    runs_v = [run_xml(v, 2600, True, col)] + ([run_xml(u, 1400, True, col)] if u else [])
    set_paras(c, [para([run_xml(lab, 1200, False, GREY)]), para(runs_v, spc=200), para([run_xml(sub, 1100, False, GREY)], spc=200)])
# argument 1 (meaning per user: since 2027 a separate KU AdTech -> increment for suppliers no longer applies; new money = non-suppliers)
a1 = shp(s, 'Аргумент 1')
r0, r1 = [p_.runs[0] for p_ in a1.text_frame.paragraphs]
r0.text = 'Инкремент по поставщикам больше неприменим'
r1.text = ('С 2027 г. рекламные деньги поставщиков идут через отдельное КУ AdTech: за весь поток отвечаем мы, '
           'выделять прирост не нужно. Новые деньги с рынка – доход от непоставщиков и вне КУ: 286,6 млн')
# re-space arguments: arg 1 grows to 4 lines
E = 914400
y1 = shp(s, 'Аргумент 1').top
for n, y in ((1, y1), (2, y1 + int(1.36 * E)), (3, y1 + int(2.50 * E))):
    shp(s, f'Аргумент {n}').top = y; shp(s, f'Номер {n}').top = y
shp(s, 'Аргумент 1').height = int(1.30 * E)
for n in (2, 3): shp(s, f'Аргумент {n}').height = int(1.05 * E)
# margin chart header with footnote markers
hdr = shp(s, 'Заголовок маржи')
set_paras(hdr, [para([run_xml('Маржа DMP', 1600, True, '000000'), run_xml('1', 1600, True, '000000', 30000),
                      run_xml(' и Instore', 1600, True, '000000'), run_xml('2', 1600, True, '000000', 30000)], algn='l')])
set_footnote(s, [
    [('1', True), (' Маржа Online + DMP – после закупки трафика, бюджет 2027 г.: 1 000 млн трафика на 1 600 млн выручки; затраты команды DMP не вычтены', False)],
    [('2', True), (' Маржа Instore – после всех затрат проектов: ФОТ, подбор персонала, электроснабжение, ремонт и ТО; для чистой прибыли – также амортизация CAPEX и налог 25%', False)],
])
# ---------- slide 3: chart title
t = shp(sl[2], 'Заголовок графика 2')
set_paras(t, [para([run_xml('Накопленный дисконтированный FCF VS DMP, млн руб.', 1400, True, '000000')], algn='l')])
# ---------- slide 4: conservative assumptions (point 9) as footnote under "Ключевые допущения модели"
set_footnote(sl[3], [[('Оценка консервативна: остаточная стоимость оборудования после 60 мес. не учитывается; маржа Online + DMP не уменьшена на затраты команды DMP – с их учётом эффект Instore выше', False)]])
# right tab sat exactly on the text edge -> LibreOffice wrapped values onto a new line; 3.2M EMU renders on one line
for n in (1, 2, 3):
    for tab in shp(sl[3], f'Допущения {n}')._element.iter(qn('a:tab')):
        tab.set('pos', '3200000')
# ---------- slides 8-9: titles
set_title(sl[7], 'FCF VS DMP положителен с 2028, окупаемость в начале 2031')
set_title(sl[8], 'NPV эффекта положителен во всех рассмотренных сценариях')
# ---------- spelling language on edited slides (CLAUDE.md): Cyrillic -> ru-RU, else en-US
n = 0
for s in (sl[1], sl[2], sl[3], sl[7], sl[8]):
    for r in s._element.iter(qn('a:r')):
        t_ = r.find(qn('a:t')); rp = r.find(qn('a:rPr'))
        if t_ is None or rp is None or not (t_.text or '').strip(): continue
        lang = 'ru-RU' if re.search('[А-Яа-яЁё]', t_.text) else 'en-US'
        if rp.get('lang') != lang: rp.set('lang', lang); n += 1
p.save(OUT)
print('saved', OUT, 'lang fixed', n, 'card width in', round(W / E, 2))
