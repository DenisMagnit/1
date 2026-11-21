"""Таймшит AdTech за ноябрь 2026: распределение часов по задачам каталога (исправленного по замечаниям МСФО).

Логика — как в октябре (fill_timesheet_oct.py, правила R1–R8), отличия ноября:
  * норма 159 ч: 20 рабочих дней, 4.11 — праздник, 3.11 — сокращенный день (7 ч);
  * каталог — «Каталог_задач_AdTech_2026-2027_МСФО.xlsx», задача определяется строкой каталога;
  * номер задачи в Jira неизвестен (переименованные задачи заводятся заново, октябрьские номера Jira к каталогу
    не сопоставляются) — в колонке «Задача» только код проекта ADTHPL;
  * преемственности с октябрем нет: загруженный октябрьский табель (выгрузка МСФО) не совпадает с нашим
    и не содержит наименований задач;
  * табель — строго в колонках шаблона ноября: «Признак», «Вид работ», «Задача», «Период списания»
    (срок задачи ∩ ноябрь, со сдвигом на отпуск), «Комментарий» (только из исходника); ширины, фильтр и
    условное форматирование шаблона не меняются;
  * наименования задач — в отдельном реестре для Jira (4-й аргумент): лист «Задачи для Jira» (что завести,
    колонка для номера) и «Строки табеля» (какой строке табеля какая задача);
  * Гряда — 100% OPEX на сопровождение отчетности Эдтех (МСФО за октябрь: «проект вне ОМНИ», в «Бизнес ИТМ» — FINANCE);
  * часы, которые не помещаются в задачи своей должности и профессии (сроки задач в ноябре), уходят на OPEX-задачи
    родственной должности (FALLBACK_OPEX) — открытый вопрос.
Запуск: python3 -I fill_timesheet_nov.py <исходник ноября.xlsx> <каталог.xlsx> <результат.xlsx> [<реестр для Jira.xlsx>]
"""
import sys, copy, math
from datetime import date, datetime, timedelta
from collections import defaultdict, OrderedDict
import pandas as pd
import openpyxl

SRC, CATALOG, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
REG = sys.argv[4] if len(sys.argv) > 4 else None

M_START, M_END = date(2026, 11, 1), date(2026, 11, 30)
HOLIDAYS = {date(2026, 11, 4)}
SHORT = {date(2026, 11, 3): 7}            # предпраздничный день
CAP_TASK = 350
JIRA = "ADTHPL"


def workdays(a, b):
    d, out = a, []
    while d <= b:
        if d.weekday() < 5 and d not in HOLIDAYS:
            out.append(d)
        d += timedelta(days=1)
    return out


def day_h(d):
    return SHORT.get(d, 8)


MONTH_WD = workdays(M_START, M_END)
assert sum(day_h(d) for d in MONTH_WD) == 159


def to_date(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return datetime.strptime(str(v).strip(), "%d.%m.%Y").date()


def parse_abs(s, e):
    if s is None:
        return []
    ss = str(s).split("/") if isinstance(s, str) else [s]
    ee = str(e).split("/") if isinstance(e, str) else [e]
    return [(to_date(a), to_date(b)) for a, b in zip(ss, ee)]


# ---------- каталог ----------
cat = pd.read_excel(CATALOG, sheet_name="Задачи")
cat["key"] = [f"R{i + 2}" for i in cat.index]                    # строка листа каталога
cat["s"] = pd.to_datetime(cat["Дата начала работ"]).dt.date
cat["e"] = pd.to_datetime(cat["Дата окончания работ"]).dt.date
TASK = {r["key"]: dict(role=r["Должность"], name=r["Наименование задачи"].strip(), flag=r["Признак"],
                       kind=r["Вид работ"], epic=r["Эпик"], s=r["s"], e=r["e"], row=i + 2)
        for i, r in cat.iterrows()}
active = [k for k, t in TASK.items() if t["s"] <= M_END and t["e"] >= M_START]

ROLE_MAP = {"Начальник отдела": "Руководитель команды", "Специалист ML": "Ведущий специалист ML"}
OPEX_ONLY = {"Директор департамента", "Руководитель управления", "Руководитель направления",
             "Менеджер проектов", "Руководитель проектов", "Ведущий инженер сопровождения",
             "Инженер сопровождения"}
CAPEX_SHARE = {"Руководитель команды": 0.3, "Менеджер продукта": 0.6, "Ведущий менеджер продукта": 0.6}
FORCE_OPEX = {"Гряда Игорь Леонидович"}                          # отчетность Эдтех, вне продукта AdTech
# остаток часов, для которого в ноябре нет задач своей должности и профессии (по срокам каталога), —
# на текущую OPEX-задачу по смыслу работы (открытый вопрос)
FALLBACK_OPEX = {"Главный инженер по поддержке оборудования": "Сопровождение промышленной эксплуатации показа рекламы",
                 "Дизайнер продукта": "Консультирование смежных команд и разбор обращений",
                 "Исследователь пользовательского опыта": "Консультирование смежных команд и разбор обращений"}

role_tasks = defaultdict(list)
for k in active:
    role_tasks[(TASK[k]["role"], TASK[k]["flag"])].append(k)

# ---------- исходник ноября ----------
wb = openpyxl.load_workbook(SRC)
ws = wb["Лист1"]
hdr = [c.value for c in ws[1]]
assert hdr[17:] == ["Признак", "Вид работ", "Задача", "Период списания", "Комментарий"], hdr[17:]
src_rows = [list(r) for r in ws.iter_rows(min_row=2)]
people = []
for cells in src_rows:
    v = [c.value for c in cells]
    if not v[3]:
        continue
    absdays = set()
    for a, b in parse_abs(v[15], v[16]):
        absdays |= set(workdays(a, b))
    people.append(dict(cells=cells, fio=v[3], pos=v[4], role=ROLE_MAP.get(v[4], v[4]), H=float(v[13]),
                       rate=float(v[12] or 1), absdays=absdays,
                       avail=[d for d in MONTH_WD if d not in absdays], comment=v[21]))


def task_window(p, k):
    t = TASK[k]
    s, e = max(t["s"], M_START), min(t["e"], M_END)
    av = [d for d in p["avail"] if s <= d <= e]
    if not av:
        return None
    if any(s <= d < av[0] for d in p["absdays"]):
        s = av[0]
    if any(av[-1] < d <= e for d in p["absdays"]):
        e = av[-1]
    return s, e, av


def hcap(p, days):
    return p["rate"] * sum(day_h(d) for d in days)


def pcap(p, k):
    w = task_window(p, k)
    return 0 if w is None else hcap(p, w[2])


# ================= РАСПРЕДЕЛЕНИЕ =================
SECOND_SHARE, SECOND_MIN, ROW_MIN = 0.3, 16, 8
load = defaultdict(float)
alloc = defaultdict(lambda: OrderedDict())


def feas(p, n):
    """Запас по задаче n с учетом пересечения сроков: на любом отрезке [a,b] сумма часов задач, чей период
    лежит внутри отрезка, не больше емкости сотрудника на этом отрезке."""
    w = task_window(p, n)
    if w is None:
        return 0
    s0, e0, _ = w
    mine = [(task_window(p, m), h) for m, h in alloc[p["fio"]].items() if m != n]
    starts = {s0} | {x[0][0] for x in mine if x[0][0] <= s0}
    ends = {e0} | {x[0][1] for x in mine if x[0][1] >= e0}
    best = float("inf")
    for a in starts:
        for b in ends:
            cap = hcap(p, [d for d in p["avail"] if a <= d <= b])
            used = alloc[p["fio"]].get(n, 0) + sum(h for (ws_, we, _), h in mine if a <= ws_ and we <= b)
            best = min(best, cap - used)
    return best


def room(p, n):
    return min(pcap(p, n) - alloc[p["fio"]].get(n, 0), CAP_TASK - load[n], feas(p, n))


def give(p, n, h):
    h = max(0, min(h, room(p, n)))
    h = float(int(h + 1e-9)) if p["rate"] >= 1 else math.floor(h * 10 + 1e-6) / 10
    if h <= 0:
        return 0
    alloc[p["fio"]][n] = alloc[p["fio"]].get(n, 0) + h
    load[n] += h
    return h


def targets(p):
    H = p["H"]
    if p["role"] in OPEX_ONLY or p["fio"] in FORCE_OPEX:
        return {"CAPEX": 0, "OPEX": H}
    sh = CAPEX_SHARE.get(p["role"])
    if sh is None:
        return {"CAPEX": H, "OPEX": 0}
    c = round(H * sh) if p["rate"] >= 1 else round(H * sh, 1)
    return {"CAPEX": c, "OPEX": H - c}


def second_size(C):
    return min(C / 2, max(SECOND_MIN, int(round(C * SECOND_SHARE / 8)) * 8))


FAMILY = [{"Специалист по тестированию", "Ведущий специалист по тестированию", "Главный специалист по тестированию"},
          {"Аналитик данных", "Ведущий аналитик данных"},
          {"Разработчик", "Ведущий разработчик", "Главный разработчик", "Ведущий разработчик проекта Эдтех"},
          {"Системный аналитик", "Ведущий системный аналитик", "Главный системный аналитик"},
          {"Ведущий специалист ML"},
          {"Менеджер продукта", "Ведущий менеджер продукта"}]
GRADE = {"Специалист по тестированию": 0, "Ведущий специалист по тестированию": 1,
         "Главный специалист по тестированию": 2, "Аналитик данных": 0, "Ведущий аналитик данных": 1,
         "Разработчик": 0, "Ведущий разработчик": 1, "Главный разработчик": 2,
         "Ведущий разработчик проекта Эдтех": 1, "Системный аналитик": 0, "Ведущий системный аналитик": 1,
         "Главный системный аналитик": 2, "Менеджер продукта": 0, "Ведущий менеджер продукта": 1}


def grade_gap(p, n):
    r = TASK[n]["role"]
    if r == p["role"]:
        return 0
    return abs(GRADE.get(r, 9) - GRADE.get(p["role"], 9)) or 0.5


def family_pool(p, flag):
    fam = next((f for f in FAMILY if p["role"] in f), {p["role"]})
    kinds = {TASK[n]["kind"] for n in role_tasks[(p["role"], flag)]}
    return [k for k in active if TASK[k]["flag"] == flag and TASK[k]["role"] in fam
            and TASK[k]["role"] != p["role"] and TASK[k]["kind"] in kinds]


def pool(p, flag):
    own = list(role_tasks[(p["role"], flag)])
    if flag == "CAPEX" and len(own) < 2:
        own += family_pool(p, flag)
    return own


def move(p, src, dst, h):
    a = alloc[p["fio"]]
    h = min(h, a[src])
    a[src] -= h; load[src] -= h
    got = give(p, dst, h)
    a[src] += h - got; load[src] += h - got
    return got


def taken(p, flag):
    return [n for n in alloc[p["fio"]] if TASK[n]["flag"] == flag]


def pick(p, flag, need, main=None, new_only=False, src=None):
    cands = [n for n in (src if src is not None else pool(p, flag))
             if (not new_only or n not in alloc[p["fio"]]) and room(p, n) >= min(need, ROW_MIN)]
    if not cands:
        return None
    ep = TASK[main]["epic"] if main else None
    return min(cands, key=lambda n: (n not in alloc[p["fio"]], grade_gap(p, n), room(p, n) < min(need, 40),
                                     ep is not None and TASK[n]["epic"] != ep, room(p, n) < need,
                                     load[n] > 0, load[n], -room(p, n), TASK[n]["row"]))


for p in people:
    p["tgt"] = targets(p)
    p["left"] = dict(p["tgt"])

# Фаза 2 — основная CAPEX-задача, вторая и далее до закрытия часов; OPEX — доливка
order = sorted(people, key=lambda p: p["fio"])
for p in order:
    C = p["tgt"]["CAPEX"]
    if C > 0:
        n = pick(p, "CAPEX", C - second_size(C), new_only=True)
        if n:
            p["left"]["CAPEX"] -= give(p, n, C - second_size(C))
        guard = 0
        while (p["left"]["CAPEX"] > 0.001 or len(taken(p, "CAPEX")) < 2) and guard < 8:
            guard += 1
            main = max(taken(p, "CAPEX"), key=lambda m: alloc[p["fio"]][m]) if taken(p, "CAPEX") else None
            need = p["left"]["CAPEX"]
            if need <= 0.001:                      # часы закрыты, но задача одна — выделяем вторую из основной
                need = second_size(C)
                n = pick(p, "CAPEX", need, main, new_only=True)
                if not n or move(p, main, n, need) <= 0:
                    break
                continue
            n = pick(p, "CAPEX", need, main, new_only=len(taken(p, "CAPEX")) < 2)
            if not n:
                break
            p["left"]["CAPEX"] -= give(p, n, need)
    guard = 0
    while p["left"]["OPEX"] > 0.001 and guard < 8:
        guard += 1
        n = pick(p, "OPEX", p["left"]["OPEX"])
        if not n:
            break
        p["left"]["OPEX"] -= give(p, n, p["left"]["OPEX"])

# Фаза 3 — своих задач не хватило по срокам: CAPEX → та же профессия соседнего грейда (тот же вид работ),
#           затем задачи другого признака своей должности, затем OPEX родственной должности
for p in order:
    guard = 0
    while p["left"]["CAPEX"] > 0.001 and guard < 8:
        guard += 1
        n = pick(p, "CAPEX", p["left"]["CAPEX"], src=family_pool(p, "CAPEX") + taken(p, "CAPEX"))
        if not n:
            break
        p["left"]["CAPEX"] -= give(p, n, p["left"]["CAPEX"])
for p in order:
    if p["fio"] in FORCE_OPEX:
        continue
    for f, g in (("CAPEX", "OPEX"), ("OPEX", "CAPEX")):
        guard = 0
        while p["left"][f] > 0.001 and guard < 8:
            guard += 1
            n = pick(p, g, p["left"][f])
            if not n:
                break
            p["left"][f] -= give(p, n, p["left"][f])
for p in order:
    key = FALLBACK_OPEX.get(p["role"])
    fb = [k for k in active if key and TASK[k]["flag"] == "OPEX" and TASK[k]["name"].startswith(key)]
    for f in ("CAPEX", "OPEX"):
        guard = 0
        while fb and p["left"][f] > 0.001 and guard < 8:
            guard += 1
            n = pick(p, "OPEX", p["left"][f], src=fb + taken(p, "OPEX"))
            if not n:
                break
            p["left"][f] -= give(p, n, p["left"][f])

# Фаза 4a — лишние CAPEX-строки < 16 ч (при 3+ CAPEX-задачах) вливаем в другие CAPEX-задачи
for p in people:
    a = alloc[p["fio"]]
    for n in sorted(list(a), key=lambda m: a[m]):
        if n not in a or TASK[n]["flag"] != "CAPEX" or a[n] >= SECOND_MIN or len(taken(p, "CAPEX")) < 3:
            continue
        h = a.pop(n); load[n] -= h
        for m in sorted(taken(p, "CAPEX"), key=lambda m: -a[m]):
            h -= give(p, m, h)
            if h <= 0.001:
                break
        if h > 0.001:
            a[n] = a.get(n, 0) + h; load[n] += h

# Фаза 4b — строки < 8 ч добиваем до 8 ч за счет основной задачи того же признака
for p in people:
    a = alloc[p["fio"]]
    for n, h in list(a.items()):
        if 0 < h < ROW_MIN and p["rate"] >= 1:
            same = [m for m in a if m != n and TASK[m]["flag"] == TASK[n]["flag"]]
            if same:
                move(p, max(same, key=a.get), n, ROW_MIN - h)

# ---------- диагностика ----------
borrowed = {p["fio"]: [(n, TASK[n]["role"], alloc[p["fio"]][n]) for n in alloc[p["fio"]] if TASK[n]["role"] != p["role"]]
            for p in people}
print("ЗАДАЧИ ЧУЖОЙ ДОЛЖНОСТИ:", {k: v for k, v in borrowed.items() if v})
print("<2 CAPEX-ЗАДАЧ ПРИ CAPEX:", [p["fio"] for p in people if p["tgt"]["CAPEX"] > 0 and len(taken(p, "CAPEX")) < 2])
dev = [(p["fio"], p["tgt"]["CAPEX"], sum(alloc[p["fio"]][n] for n in taken(p, "CAPEX"))) for p in people]
print("ОТКЛОНЕНИЕ ДОЛЕЙ (факт CAPEX ≠ таргет):", [d for d in dev if abs(d[1] - d[2]) > 0.01])
print("НЕ РАСПРЕДЕЛЕНО:", [(p["fio"], p["left"]) for p in people if sum(p["left"].values()) > 0.001])
print("ПРЕВЫШЕНИЕ 350:", {n: v for n, v in load.items() if v > CAP_TASK})

# ---------- запись: строго колонки шаблона ноября ----------
src_style = {p["fio"]: [copy.copy(c._style) for c in p["cells"]] for p in people}
hdr_cells = [(c.value, copy.copy(c._style)) for c in ws[1]]
tmpl = [copy.copy(c._style) for c in src_rows[0]]
ws.delete_rows(1, ws.max_row)
for j, (h, st) in enumerate(hdr_cells):
    ws.cell(row=1, column=j + 1, value=h)._style = st
row, reg = 2, []
for p in people:
    v = [c.value for c in p["cells"]]
    items = sorted(alloc[p["fio"]].items(), key=lambda x: (-x[1], TASK[x[0]]["flag"] != "CAPEX", TASK[x[0]]["row"]))
    for k, h in items:
        t = TASK[k]
        s, e, _ = task_window(p, k)
        hv = int(h) if p["rate"] >= 1 else h
        period = f"{s:%d.%m.%Y} - {e:%d.%m.%Y}"
        vals = v[:13] + [hv] + v[14:17] + [t["flag"], t["kind"], JIRA, period, p["comment"]]
        for j, val in enumerate(vals):
            ws.cell(row=row, column=j + 1, value=val)._style = copy.copy(src_style[p["fio"]][j] if j < 17 else tmpl[j])
        reg.append([row, p["fio"], p["pos"], t["flag"], t["kind"], t["name"], period, hv, k])
        row += 1
ws.auto_filter.ref = f"A1:V{row - 1}"

# лист «Отсутствуют в списках ноября»: формулы XLOOKUP ссылаются на внешнюю книгу — оставляем значения
vals_wb = openpyxl.load_workbook(SRC, data_only=True)
for sh in wb.sheetnames[1:]:
    for r1, r2 in zip(wb[sh].iter_rows(), vals_wb[sh].iter_rows()):
        for c1, c2 in zip(r1, r2):
            if isinstance(c1.value, str) and c1.value.startswith("="):
                c1.value = c2.value
wb._external_links = []
wb.save(OUT)

# ---------- реестр задач для Jira (отдельный файл, в табель не входит) ----------
if REG:
    src_cat = pd.read_excel(CATALOG.replace("_МСФО", "_исходный"), sheet_name="Задачи") if "_МСФО" in CATALOG else None
    from openpyxl.styles import Font, PatternFill, Alignment
    rb = openpyxl.Workbook()
    B, Y = Font(name="Calibri", bold=True, size=10), PatternFill("solid", fgColor="FFFF00")
    F10 = Font(name="Calibri", size=10)
    w1 = rb.active; w1.title = "Задачи для Jira"
    h1 = ["Наименование задачи", "Признак", "Вид работ", "Модуль платформы", "Эпик", "Срок задачи (каталог)",
          "Часов в ноябре", "Сотрудников", "Наименование в каталоге до правки МСФО", "Номер задачи в Jira"]
    w1.append(h1)
    used = OrderedDict()
    for r in reg:
        used.setdefault(r[8], [0.0, set()]); used[r[8]][0] += r[7]; used[r[8]][1].add(r[1])
    for k in sorted(used, key=lambda k: TASK[k]["row"]):
        t = TASK[k]
        old = src_cat["Наименование задачи"].iat[t["row"] - 2].strip() if src_cat is not None else ""
        w1.append([t["name"], t["flag"], t["kind"], cat["Модуль платформы"].iat[t["row"] - 2], t["epic"],
                   f"{t['s']:%d.%m.%Y} - {t['e']:%d.%m.%Y}", round(used[k][0], 1), len(used[k][1]),
                   old if old != t["name"] else "", None])
    w2 = rb.create_sheet("Строки табеля")
    w2.append(["Строка табеля", "ФИО сотрудника", "Должность сотрудника", "Признак", "Вид работ",
               "Наименование задачи", "Период списания", "Списать часов"])
    for r in reg:
        w2.append(r[:8])
    for w, widths in ((w1, [90, 9, 30, 30, 50, 24, 10, 10, 90, 16]), (w2, [9, 34, 34, 9, 30, 90, 24, 10])):
        for c in w[1]:
            c.font, c.alignment = B, Alignment(wrap_text=True, vertical="top")
        for rr in w.iter_rows(min_row=2):
            for c in rr:
                c.font = F10
        for jj, wd in enumerate(widths):
            w.column_dimensions[openpyxl.utils.get_column_letter(jj + 1)].width = wd
        w.auto_filter.ref = w.dimensions
        w.freeze_panes = "A2"
    w1.cell(1, 10).fill = Y
    rb.save(REG)

tot = sum(p["H"] for p in people)
capx = sum(h for a in alloc.values() for n, h in a.items() if TASK[n]["flag"] == "CAPEX")
print(f"людей {len(people)}, строк {row - 2}, задач {len(load)}, часов {tot:g}, CAPEX {capx:g} ({capx / tot:.1%})")
