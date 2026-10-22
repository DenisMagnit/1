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
             "Инженер сопровождения"}
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

# ================= РАСПРЕДЕЛЕНИЕ (единый набор правил) =================
# Правила (по сентябрю + согласованные уточнения):
#  R1  задачи — только своей должности по каталогу, активные в октябре; исключение — Клименко (EXCEPT)
#  R2  доли: продакты 60% CAPEX / 40% OPEX; руководители команд 30/70; менеджмент, сопровождение,
#      остальные (включая ведущего разработчика проекта Эдтех) — 100% CAPEX
#  R3  у кого есть CAPEX — не менее 2 CAPEX-задач: основная ~70%, вторая ~30% (кратно 8, не менее 16 ч);
#      100% OPEX — одна задача (как в сентябре), вторая только если упираемся в лимит
#  R4  ≤ 350 ч на задачу в месяц; ≤ 8 ч × ставка × раб. дни сотрудника в периоде работ по задаче
#  R5  преемственность: задачи сентября, которые идут в октябре, сохраняются (OPEX 3 кв. → 4 кв.)
#  R6  новые задачи — сначала незанятые/наименее загруженные (покрытие каталога), вторая — из того же эпика
#  R7  строк < 8 ч не оставляем
SECOND_SHARE, SECOND_MIN, ROW_MIN = 0.3, 16, 8
EXCEPT = {"Главный инженер по поддержке оборудования": ["ADTHPL-202"]}  # приемка и подготовка к запуску

load = defaultdict(float)
alloc = defaultdict(lambda: OrderedDict())  # fio -> {num: hours}

def feas(p, n):
    """Сколько еще часов можно добавить в задачу n, чтобы на любом отрезке [a,b] сумма часов задач,
    чей период работ лежит внутри отрезка, не превышала 8 ч × ставка × раб. дни без отпуска."""
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
            cap = 8 * p["rate"] * len([d for d in p["avail"] if a <= d <= b])
            used = alloc[p["fio"]].get(n, 0) + sum(h for (ws, we, _), h in mine if a <= ws and we <= b)
            best = min(best, cap - used)
    return best

def room(p, n):
    return min(pcap(p, n) - alloc[p["fio"]].get(n, 0), CAP_TASK - load[n], feas(p, n))

def give(p, num, h):
    h = max(0, min(h, room(p, num)))
    if p["rate"] >= 1:
        h = float(int(h))
    else:
        h = round(h, 1)
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

def second_size(C):
    return min(C / 2, max(SECOND_MIN, int(round(C * SECOND_SHARE / 8)) * 8))

FAMILY = [{"Специалист по тестированию", "Ведущий специалист по тестированию", "Главный специалист по тестированию"},
          {"Аналитик данных", "Ведущий аналитик данных"},
          {"Разработчик", "Ведущий разработчик", "Главный разработчик", "Ведущий разработчик проекта Эдтех"},
          {"Системный аналитик", "Ведущий системный аналитик", "Главный системный аналитик"},
          {"Ведущий специалист ML"},
          {"Менеджер продукта", "Ведущий менеджер продукта"}]

def pool(p, flag):
    own = list(role_tasks[(p["role"], flag)])
    if flag == "CAPEX":
        own += EXCEPT.get(p["role"], [])
        if len(own) < 2:          # для второй CAPEX-задачи — та же профессия соседнего грейда
            own += family_pool(p, flag)
    return own

def family_pool(p, flag):
    fam = next((f for f in FAMILY if p["role"] in f), {p["role"]})
    kinds = {TASK[n]._5 for n in role_tasks[(p["role"], flag)]}
    return [r.num for r in active.itertuples() if r.Признак == flag and r.Должность in fam
            and r.Должность != p["role"] and r._5 in kinds]

def taken(p, flag):
    return [n for n in alloc[p["fio"]] if TASK[n].Признак == flag]

for p in people:
    p["tgt"] = targets(p)
    p["left"] = dict(p["tgt"])
    cont = [(n, h) for n, h in sep_tasks.get(p["fio"], []) if n and n in active.num.values
            and TASK[n].Должность == p["role"]]
    p["cont"] = {f: sorted([(n, h) for n, h in cont if TASK[n].Признак == f], key=lambda x: -x[1])
                 for f in ("CAPEX", "OPEX")}

# Фаза 1 — преемственность
for p in people:
    C, O = p["tgt"]["CAPEX"], p["tgt"]["OPEX"]
    cc = p["cont"]["CAPEX"]
    if C > 0 and cc:
        if len(cc) >= 2 and min(h for _, h in cc) >= SECOND_MIN:
            tot = sum(h for _, h in cc)          # сентябрьская разбивка содержательная — сохраняем пропорции
            for n, h in cc:
                p["left"]["CAPEX"] -= give(p, n, round(C * h / tot))
        else:
            p["left"]["CAPEX"] -= give(p, cc[0][0], C - second_size(C))
    if O > 0 and p["cont"]["OPEX"]:
        p["left"]["OPEX"] -= give(p, p["cont"]["OPEX"][0][0], O)

# Фаза 2 — добор: основная (если нет), затем вторая и далее до закрытия часов
def pick(p, flag, need, main=None, prefer=(), new_only=False, src=None):
    # сначала доливаем уже назначенные задачи с запасом емкости, затем берем новые
    cands = [n for n in (src if src is not None else pool(p, flag)) if (not new_only or n not in alloc[p["fio"]])
             and room(p, n) >= min(need, ROW_MIN)]
    if not cands:
        return None
    ep = TASK[main].Эпик if main else None
    return min(cands, key=lambda n: (n not in alloc[p["fio"]], n not in prefer, room(p, n) < min(need, 40),
                                     ep is not None and TASK[n].Эпик != ep,
                                     room(p, n) < need, load[n] > 0, load[n], -room(p, n), n))

order = sorted(people, key=lambda p: (not p["cont"]["CAPEX"] and not p["cont"]["OPEX"], p["fio"]))
for p in order:
    C = p["tgt"]["CAPEX"]
    if C > 0:
        if not taken(p, "CAPEX"):
            n = pick(p, "CAPEX", C - second_size(C), new_only=True)
            if n:
                p["left"]["CAPEX"] -= give(p, n, C - second_size(C))
        prefer = [n for n, _ in p["cont"]["CAPEX"][1:]]
        guard = 0
        while (p["left"]["CAPEX"] > 0.001 or len(taken(p, "CAPEX")) < 2) and guard < 8:
            guard += 1
            main = max(taken(p, "CAPEX"), key=lambda n: alloc[p["fio"]][n]) if taken(p, "CAPEX") else None
            need = p["left"]["CAPEX"]
            if need <= 0.001:                     # часы закрыты, но задача одна — выделяем вторую из основной
                need = second_size(C)
                n = pick(p, "CAPEX", need, main, prefer, new_only=True)
                if not n:
                    break
                h = give(p, n, need)
                alloc[p["fio"]][main] -= h; load[main] -= h
                continue
            n = pick(p, "CAPEX", need, main, prefer, new_only=len(taken(p, "CAPEX")) < 2)
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

# Фаза 3 — если своих задач не хватило по срокам: CAPEX → задачи соседнего грейда той же профессии
#           (тот же вид работ), затем остаток — на задачи другого признака своей должности
for p in order:
    guard = 0
    while p["left"]["CAPEX"] > 0.001 and guard < 8:
        guard += 1
        n = pick(p, "CAPEX", p["left"]["CAPEX"], src=family_pool(p, "CAPEX") + taken(p, "CAPEX"))
        if not n:
            break
        p["left"]["CAPEX"] -= give(p, n, p["left"]["CAPEX"])
for p in order:
    for f, g in (("CAPEX", "OPEX"), ("OPEX", "CAPEX")):
        guard = 0
        while p["left"][f] > 0.001 and guard < 8:
            guard += 1
            n = pick(p, g, p["left"][f])
            if not n:
                break
            p["left"][f] -= give(p, n, p["left"][f])

# Фаза 4 — строки < 8 ч добиваем до 8 ч за счет основной задачи того же признака
for p in people:
    a = alloc[p["fio"]]
    for n, h in list(a.items()):
        if 0 < h < ROW_MIN and p["rate"] >= 1:
            same = [m for m in a if m != n and TASK[m].Признак == TASK[n].Признак]
            if not same:
                continue
            big = max(same, key=a.get)
            need = min(ROW_MIN - h, CAP_TASK - load[n], pcap(p, n) - h)
            if need > 0:
                a[n] += need; load[n] += need; a[big] -= need; load[big] -= need

borrowed = {p["fio"]: [n for n in alloc[p["fio"]] if TASK[n].Должность != p["role"]] for p in people}
borrowed = {k: v for k, v in borrowed.items() if v}
print("ЗАДАЧИ ЧУЖОЙ ДОЛЖНОСТИ:", borrowed)
print("<2 CAPEX-ЗАДАЧ ПРИ CAPEX:", [p["fio"] for p in people if p["tgt"]["CAPEX"] > 0 and len(taken(p, "CAPEX")) < 2])
print("ОТКЛОНЕНИЕ ДОЛЕЙ:", [(p["fio"], p["tgt"]) for p in people
                          if abs(sum(alloc[p["fio"]].get(n, 0) for n in taken(p, "CAPEX")) - p["tgt"]["CAPEX"]) > 0.01])

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
        comment = p["comment"]          # только из октябрьского исходника
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

