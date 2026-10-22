import sys, re, copy
from datetime import date, datetime, timedelta
from collections import defaultdict, OrderedDict
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

IN = sys.argv[1]
OUT = sys.argv[2]

M_START, M_END = date(2026, 10, 1), date(2026, 10, 31)
HOLIDAYS = set()  # в октябре 2026 праздников нет, норма 22 дня / 176 ч
CAP_TASK = 350

def workdays(a, b):
    d, out = a, []
    while d <= b:
        if d.weekday() < 5 and d not in HOLIDAYS:
            out.append(d)
        d += timedelta(days=1)
    return out

MONTH_WD = workdays(M_START, M_END)

def to_date(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return datetime.strptime(v.strip(), "%d.%m.%Y").date()

def parse_abs(s, e):
    if s is None:
        return []
    ss = str(s).split("/") if isinstance(s, str) else [s]
    ee = str(e).split("/") if isinstance(e, str) else [e]
    res = []
    for a, b in zip(ss, ee):
        a, b = to_date(a), to_date(b)
        res.append((a, b))
    return res

# ---------- каталог задач ----------
cat = pd.read_excel(f"{IN}/3_tasks.xlsx", sheet_name="Задачи")
cat["num"] = [f"ADTHPL-{i + 101}" for i in cat.index]   # номер = строка листа + 99
cat["s"] = pd.to_datetime(cat["Дата начала работ"]).dt.date
cat["e"] = pd.to_datetime(cat["Дата окончания работ"]).dt.date

rep = pd.read_excel(f"{IN}/4_replace.xlsx")
rep = rep.dropna(subset=[rep.columns[0]])
REPL = {str(a).strip(): str(c).strip() for a, c in zip(rep.iloc[:, 0], rep.iloc[:, 2])}
cat["name"] = [REPL.get(n.strip(), n.strip()) for n in cat["Наименование задачи"]]
cat["renamed"] = [n.strip() in REPL for n in cat["Наименование задачи"]]

TASK = {r.num: r for r in cat.itertuples()}
by_name = {r["Наименование задачи"].strip(): r.num for _, r in cat.iterrows()}
active = cat[(cat.s <= M_END) & (cat.e >= M_START)]

ROLE_MAP = {"Начальник отдела": "Руководитель команды", "Специалист ML": "Ведущий специалист ML"}
OPEX_ONLY = {"Директор департамента", "Руководитель управления", "Руководитель направления",
             "Менеджер проектов", "Руководитель проектов", "Ведущий инженер сопровождения",
             "Инженер сопровождения", "Ведущий разработчик проекта Эдтех"}
CAPEX_SHARE = {"Руководитель команды": 0.3, "Менеджер продукта": 0.6, "Ведущий менеджер продукта": 0.6}

role_tasks = defaultdict(list)
for r in active.itertuples():
    role_tasks[(r.Должность, r.Признак)].append(r.num)

# ---------- сентябрь (для преемственности) ----------
sep = pd.read_excel(f"{IN}/2_sep_done.xlsx", sheet_name="Лист1", header=None, skiprows=1)
sep_tasks = defaultdict(list)  # fio -> [(num_oct, hours)]
for _, r in sep.iterrows():
    nm = str(r[19]).strip()
    q4 = nm.replace("(3 кв. 2026)", "(4 кв. 2026)")
    num = by_name.get(q4) or by_name.get(nm)
    sep_tasks[r[3]].append((num, r[13]))
sep_comment = {r[3]: r[23] for _, r in sep.iterrows() if isinstance(r[23], str)}

# ---------- октябрь: исходник ----------
wb = openpyxl.load_workbook(f"{IN}/5_oct_src.xlsx")
ws = wb["Лист1"]
hdr = [c.value for c in ws[1]]
src_rows = [list(r) for r in ws.iter_rows(min_row=2)]
people = []
for cells in src_rows:
    v = [c.value for c in cells]
    if not v[3]:
        continue
    pos = v[4]
    role = ROLE_MAP.get(pos, pos)
    H = float(v[13])
    rate = float(v[11] or 1)
    ab = parse_abs(v[15], v[16])
    absdays = set()
    for a, b in ab:
        absdays |= set(workdays(a, b))
    avail = [d for d in MONTH_WD if d not in absdays]
    people.append(dict(cells=cells, fio=v[3], pos=pos, role=role, H=H, rate=rate,
                       absdays=absdays, avail=avail, comment=v[20]))

def task_window(p, num):
    t = TASK[num]
    s, e = max(t.s, M_START), min(t.e, M_END)
    av = [d for d in p["avail"] if s <= d <= e]
    if not av:
        return None
    fa, la = av[0], av[-1]
    # отсутствие в начале / в конце периода сдвигает даты работ (как в сентябре)
    if any(s <= d < fa for d in p["absdays"]):
        s = fa
    if any(la < d <= e for d in p["absdays"]):
        e = la
    return s, e, len(av)

def pcap(p, num):
    w = task_window(p, num)
    return 0 if w is None else 8 * p["rate"] * w[2]

load = defaultdict(float)
alloc = defaultdict(lambda: OrderedDict())  # fio -> {num: hours}

def give(p, num, h):
    h = min(h, pcap(p, num) - alloc[p["fio"]].get(num, 0), CAP_TASK - load[num])
    h = max(0, h)
    if p["rate"] >= 1:
        h = float(int(h))
    if h <= 0:
        return 0
    alloc[p["fio"]][num] = alloc[p["fio"]].get(num, 0) + h
    load[num] += h
    return h

def targets(p):
    H = p["H"]
    if p["role"] in OPEX_ONLY:
        return {"CAPEX": 0, "OPEX": H}
    sh = CAPEX_SHARE.get(p["role"])
    if sh is None:
        return {"CAPEX": H, "OPEX": 0}
    c = round(H * sh) if p["rate"] >= 1 else round(H * sh, 1)
    return {"CAPEX": c, "OPEX": H - c}

for p in people:
    p["tgt"] = targets(p)
    p["left"] = dict(p["tgt"])

# 1) преемственность: задачи сентября, которые еще идут в октябре (квартальные OPEX -> 4 кв.)
for p in people:
    cont = [(n, h) for n, h in sep_tasks.get(p["fio"], []) if n and n in active.num.values
            and TASK[n].Должность == p["role"]]
    for flag in ("CAPEX", "OPEX"):
        lst = [(n, h) for n, h in cont if TASK[n].Признак == flag]
        if not lst or p["left"][flag] <= 0:
            continue
        tot = sum(h for _, h in lst)
        want = p["left"][flag]
        rest = want
        for i, (n, h) in enumerate(lst):
            share = rest if i == len(lst) - 1 else round(want * h / tot)
            got = give(p, n, share)
            rest -= got
        p["left"][flag] = rest

# 2) добор по свободным задачам должности: сначала наименее загруженные
def fill(p, flag, pool):
    guard = 0
    while p["left"][flag] > 0.001 and guard < 10:
        guard += 1
        cands = [n for n in pool if pcap(p, n) - alloc[p["fio"]].get(n, 0) > 0 and load[n] < CAP_TASK]
        if not cands:
            return
        room = lambda n: min(pcap(p, n) - alloc[p["fio"]].get(n, 0), CAP_TASK - load[n])
        cands.sort(key=lambda n: (room(n) < min(p["left"][flag], 80), load[n] > 0, load[n], -room(n), n))
        got = give(p, cands[0], p["left"][flag])
        p["left"][flag] -= got
        if got == 0:
            return

order = sorted(people, key=lambda p: (len(sep_tasks.get(p["fio"], [])) > 0, p["fio"]))
for p in order:
    for flag in ("CAPEX", "OPEX"):
        fill(p, flag, role_tasks[(p["role"], flag)])
# 3) остаток, если CAPEX-емкости не хватило, — на OPEX задачи должности (и наоборот)
for p in order:
    for flag, other in (("CAPEX", "OPEX"), ("OPEX", "CAPEX")):
        if p["left"][flag] > 0.001:
            p["left"][other] += p["left"][flag]
            p["left"][flag] = 0
            fill(p, other, role_tasks[(p["role"], other)])

# 4) «огрызки» < 8 ч (после лимита 350) добиваем до 8 ч за счет основной задачи сотрудника
for p in people:
    a = alloc[p["fio"]]
    if p["rate"] < 1 or len(a) < 2:
        continue
    for n, h in list(a.items()):
        if h < 8:
            big = max(a, key=a.get)
            need = min(8 - h, CAP_TASK - load[n], pcap(p, n) - h)
            if big != n and need > 0:
                a[n] += need; load[n] += need
                a[big] -= need; load[big] -= need

# 5) правило «по капитализируемому ФОТ — не менее 2 CAPEX-задач»: у кого CAPEX на одной
#    задаче — ~30% CAPEX-часов (кратно 8, мин. 8) переносим на вторую CAPEX-задачу своей
#    должности (сначала тот же эпик, затем наименее загруженные). 100% OPEX — без изменений.
SECOND_SHARE = 0.3
# исключение: у должности одна CAPEX-задача на октябрь — берем ближайшую по смыслу CAPEX
# другой должности (приемка и подготовка к запуску = вывод в промышленную среду)
EXCEPT = {"Главный инженер по поддержке оборудования": ["ADTHPL-202"]}
borrowed = {}
for p in people:
    a = alloc[p["fio"]]
    capt = [n for n in a if TASK[n].Признак == "CAPEX"]
    if len(capt) != 1:
        continue
    main = capt[0]
    tm = TASK[main]
    hc = a[main]
    if p["rate"] >= 1:
        want = max(8, int(round(hc * SECOND_SHARE / 8)) * 8)
    else:
        want = round(hc * SECOND_SHARE, 1)
    own = [n for n in role_tasks[(p["role"], "CAPEX")] if n != main]
    exc = [n for n in EXCEPT.get(p["role"], []) if n != main]
    for pool in (own, exc):
        room = lambda n: min(pcap(p, n), CAP_TASK - load[n])
        cands = [n for n in pool if room(n) >= min(want, 8)]
        if not cands:
            continue
        cands.sort(key=lambda n: (TASK[n].Эпик != tm.Эпик, room(n) < want, load[n], n))
        n2 = cands[0]
        h2 = min(want, room(n2))
        if p["rate"] >= 1:
            h2 = float(int(h2))
        a[n2] = h2; load[n2] += h2
        a[main] -= h2; load[main] -= h2
        if pool is exc:
            borrowed[p["fio"]] = n2
        break
print("ВТОРАЯ ЗАДАЧА ИЗ ДРУГОЙ ДОЛЖНОСТИ:", borrowed)
print("ОДНА CAPEX-ЗАДАЧА:", [p["fio"] for p in people if sum(TASK[n].Признак == "CAPEX" for n in alloc[p["fio"]]) == 1])

problems = [(p["fio"], p["left"]) for p in people if sum(p["left"].values()) > 0.001]
print("НЕ РАСПРЕДЕЛЕНО:", problems)
over = {n: v for n, v in load.items() if v > CAP_TASK}
print("ПРЕВЫШЕНИЕ 350:", over)

# ---------- запись ----------
NEW_HDR = hdr[:18] + ["Вид работ", "Наименование задачи", "Задача",
                      "Дата начала работ над задачей", "Дата окончания работ над задачей", "Комментарий"]
styles = [copy.copy(c._style) for c in src_rows[0]]
hstyles = [copy.copy(c._style) for c in ws[1]]
widths = {k: v.width for k, v in ws.column_dimensions.items()}
ws.delete_rows(1, ws.max_row)

col_src = list(range(18)) + [18, 19, 19, 19, 19, 20]   # стиль для новых колонок берем у соседних
for j, h in enumerate(NEW_HDR):
    c = ws.cell(row=1, column=j + 1, value=h)
    c._style = copy.copy(hstyles[col_src[j]])
row = 2
DATEFMT = "dd.mm.yyyy"
for p in people:
    v = [c.value for c in p["cells"]]
    items = list(alloc[p["fio"]].items())
    items.sort(key=lambda x: (TASK[x[0]].Признак != "CAPEX", -x[1]))  # CAPEX первым, как в сентябре
    for num, h in items:
        t = TASK[num]
        s, e, _ = task_window(p, num)
        comment = p["comment"] or sep_comment.get(p["fio"])
        vals = v[:13] + [h] + v[14:17] + [t.Признак, t._5, t.name, num, s, e, comment]
        for j, val in enumerate(vals):
            c = ws.cell(row=row, column=j + 1, value=val)
            c._style = copy.copy(styles[col_src[j]])
        ws.cell(row=row, column=22).number_format = DATEFMT
        ws.cell(row=row, column=23).number_format = DATEFMT
        row += 1
for k, w in widths.items():
    ws.column_dimensions[k].width = w
ws.column_dimensions["U"].width = 15
ws.column_dimensions["T"].width = 95
ws.column_dimensions["S"].width = 30
ws.column_dimensions["V"].width = 14
ws.column_dimensions["W"].width = 14
ws.column_dimensions["X"].width = 28
ws.auto_filter.ref = f"A1:X{row - 1}"

# ---------- лист контроля ----------
ck = wb.create_sheet("Контроль")
bold = Font(bold=True)
ck.append(["Контроль по сотрудникам"]); ck["A1"].font = bold
ck.append(["ФИО", "Должность", "Списать часов", "Распределено", "CAPEX ч", "OPEX ч", "Доля CAPEX", "Задач", "Расхождение"])
for c in ck[2]: c.font = bold
for p in people:
    a = alloc[p["fio"]]
    cap = sum(h for n, h in a.items() if TASK[n].Признак == "CAPEX")
    op = sum(h for n, h in a.items() if TASK[n].Признак == "OPEX")
    ck.append([p["fio"], p["pos"], p["H"], cap + op, cap, op, round(cap / p["H"], 3) if p["H"] else 0, len(a), round(p["H"] - cap - op, 2)])
r0 = ck.max_row + 2
ck.cell(row=r0, column=1, value="Контроль по задачам (лимит 350 ч на задачу в месяц)").font = bold
ck.cell(row=r0 + 1, column=1, value="Задача").font = bold
for j, h in enumerate(["Наименование", "Признак", "Должность в каталоге", "Часов всего", "Сотрудников", "Остаток до 350"]):
    ck.cell(row=r0 + 1, column=j + 2, value=h).font = bold
cnt = defaultdict(int)
for fio, a in alloc.items():
    for n in a:
        cnt[n] += 1
rr = r0 + 2
for n in sorted(load, key=lambda x: int(x.split("-")[1])):
    t = TASK[n]
    for j, val in enumerate([n, t.name, t.Признак, t.Должность, load[n], cnt[n], CAP_TASK - load[n]]):
        ck.cell(row=rr, column=j + 1, value=val)
    rr += 1
for col, w in zip("ABCDEFGHI", [38, 40, 14, 22, 12, 12, 12, 8, 12]):
    ck.column_dimensions[col].width = w
ck.column_dimensions["B"].width = 40

wb.save(OUT)

tot = sum(p["H"] for p in people)
capx = sum(h for a in alloc.values() for n, h in a.items() if TASK[n].Признак == "CAPEX")
print(f"людей {len(people)}, строк {row - 2}, часов {tot}, CAPEX {capx} ({capx / tot:.1%})")
used_renamed = [n for n in load if TASK[n].renamed]
print("использованы переименованные задачи:", used_renamed)
