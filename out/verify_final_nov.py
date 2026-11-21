"""Независимая финальная проверка таймшита AdTech за ноябрь 2026 (код не связан со скриптом распределения).
Наименования задач берутся из реестра для Jira (лист «Строки табеля»): в табеле по шаблону только код ADTHPL.
Запуск: python3 -I verify_final_nov.py <исходник ноября.xlsx> <каталог.xlsx> <итоговый файл.xlsx> <реестр для Jira.xlsx> [HR-группировки ...]"""
import sys, re
from collections import defaultdict
from datetime import date, datetime, timedelta
import openpyxl
import pandas as pd

SRC, CATALOG, F, REG, HR = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5:]
M0, M1 = date(2026, 11, 1), date(2026, 11, 30)
HOL, SHORT = {date(2026, 11, 4)}, {date(2026, 11, 3): 7}
errors = []


def report(title, errs, info=None):
    errors.append((title, len(errs)))
    print(("OK   " if not errs else "FAIL ") + title + (f" — {info}" if info else ""))
    for e in errs[:15]:
        print("        ", e)


def to_d(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.date()
    if isinstance(v, date):
        return v
    return datetime.strptime(str(v).strip(), "%d.%m.%Y").date()


def wdays(a, b):
    x, out = a, []
    while x <= b:
        if x.weekday() < 5 and x not in HOL:
            out.append(x)
        x += timedelta(1)
    return out


def dh(d):
    return SHORT.get(d, 8)


wb, wb_src = openpyxl.load_workbook(F), openpyxl.load_workbook(SRC)
wb_srcv = openpyxl.load_workbook(SRC, data_only=True)
ws, ws_src = wb["Лист1"], wb_src["Лист1"]
H, Hs = [c.value for c in ws[1]], [c.value for c in ws_src[1]]
rows = [[c.value for c in r] for r in ws.iter_rows(min_row=2)]
src_rows = [[c.value for c in r] for r in ws_src.iter_rows(min_row=2) if r[3].value]
ix, ixs = {h: i for i, h in enumerate(H)}, {h: i for i, h in enumerate(Hs)}
# реестр: строка табеля -> наименование задачи
reg = pd.read_excel(REG, sheet_name="Строки табеля")
regt = pd.read_excel(REG, sheet_name="Задачи для Jira")
reg_bad = []
if len(reg) != len(rows):
    reg_bad.append(f"строк в реестре {len(reg)} ≠ {len(rows)}")
for (_, g), r in zip(reg.iterrows(), rows):
    same = (g["ФИО сотрудника"] == r[ix["ФИО сотрудника"]] and g["Признак"] == r[ix["Признак"]] and g["Вид работ"] == r[ix["Вид работ"]]
            and g["Период списания"] == r[ix["Период списания"]] and abs(g["Списать часов"] - r[ix["Списать часов"]]) < 1e-9)
    if not same:
        reg_bad.append(f"строка {g['Строка табеля']}: {g['ФИО сотрудника']}")
ix["Наименование задачи"] = len(H)
rows = [r + [str(g).strip()] for r, g in zip(rows, reg["Наименование задачи"])]

cat = pd.read_excel(CATALOG, sheet_name="Задачи")
epi = pd.read_excel(CATALOG, sheet_name="Эпики").set_index("Эпик")
CAT = {str(r["Наименование задачи"]).strip(): r for _, r in cat.iterrows()}
assert len(CAT) == len(cat)

print("=== A. Формат (эталон — шаблон ноября) ===")
report("A1. листы — как в исходнике", [] if wb.sheetnames == wb_src.sheetnames else [f"{wb.sheetnames}"], ", ".join(wb.sheetnames))
report("A2. колонки и заголовки — строго как в шаблоне (22)", [] if H == Hs else [f"{H}"])
report("A2a. ширины колонок и условное форматирование — как в шаблоне",
       ([] if {k: v.width for k, v in ws.column_dimensions.items()} == {k: v.width for k, v in ws_src.column_dimensions.items()} else ["ширины"]) +
       ([] if [str(c.sqref) for c in ws.conditional_formatting] == [str(c.sqref) for c in ws_src.conditional_formatting] else ["условное форматирование"]))
srcc = {r[3].value: r for r in ws_src.iter_rows(min_row=2) if r[3].value}
bad = [c.coordinate for r in ws.iter_rows(min_row=2) for c in r[:17]
       if c.number_format != srcc[r[3].value][c.column - 1].number_format]
report("A3. форматы ячеек исходника (в т.ч. даты отсутствия) сохранены", bad)
extra = [c.coordinate for r in ws.iter_rows() for c in r if c.column > 22 and c.value is not None]
report("A4. нет лишних колонок и пустых строк", extra + [f"пустая строка {i + 2}" for i, r in enumerate(rows) if not r[3]])
report("A5. автофильтр по всему диапазону", [] if ws.auto_filter.ref == f"A1:V{ws.max_row}" else [ws.auto_filter.ref])
bad = []
for sh in wb_src.sheetnames[1:]:
    a = [list(r) for r in wb[sh].iter_rows(values_only=True)]
    b = [list(r) for r in wb_srcv[sh].iter_rows(values_only=True)]
    if a != b:
        bad.append(sh)
    bad += [f"{sh}!{c.coordinate} формула" for r in wb[sh].iter_rows() for c in r if isinstance(c.value, str) and c.value.startswith("=")]
bad += ["внешние связи" for _ in wb._external_links]
report("A6. остальные листы — значения как в исходнике, без формул на внешнюю книгу", bad)

print("=== B. Перенос данных исходника ===")
src_by = {r[3]: r for r in src_rows}
fio_out = [r[ix["ФИО сотрудника"]] for r in rows]
order_out = list(dict.fromkeys(fio_out))
runs = sum(1 for i in range(len(fio_out)) if i == 0 or fio_out[i] != fio_out[i - 1])
report("B1. сотрудники = исходник, порядок как в исходнике, строки сотрудника подряд",
       ([] if order_out == [r[3] for r in src_rows] else ["состав/порядок"]) + ([] if runs == len(order_out) else ["не подряд"]),
       f"{len(order_out)} чел., {len(rows)} строк")
bad = []
for r in rows:
    s = src_by[r[ix["ФИО сотрудника"]]]
    for h in Hs[:17]:
        if h != "Списать часов" and r[ix[h]] != s[ixs[h]]:
            bad.append(f"{r[3]}: {h}")
    if r[ix["Комментарий"]] != s[ixs["Комментарий"]]:
        bad.append(f"{r[3]}: комментарий")
report("B2. колонки исходника и комментарий перенесены без изменений", bad)
tot = defaultdict(float)
for r in rows:
    tot[r[3]] += r[ix["Списать часов"]]
report("B3. сумма часов сотрудника = «Списать часов» исходника",
       [f"{n}: {tot[n]} != {s[13]}" for n, s in src_by.items() if abs(tot[n] - s[13]) > 1e-6])
absd = {}
for s in src_rows:
    days = set()
    if s[15] is not None:
        p1 = str(s[15]).split("/") if isinstance(s[15], str) else [s[15]]
        p2 = str(s[16]).split("/") if isinstance(s[16], str) else [s[16]]
        for x, y in zip(p1, p2):
            days |= set(wdays(to_d(x), to_d(y)))
    absd[s[3]] = days
bad = [f"{s[3]}: {s[13]} ≠ {s[12]} × (159 − {sum(dh(d) for d in absd[s[3]])})" for s in src_rows
       if abs(s[13] - (s[12] or 1) * (159 - sum(dh(d) for d in absd[s[3]]))) > 0.05]
report("B4. исходник: часы = ставка × (норма 159 ч − отсутствие; 3.11 — 7 ч, 4.11 — праздник)", bad)

print("=== R. Реестр задач для Jira ===")
tot_reg = reg.groupby("Наименование задачи")["Списать часов"].sum()
rb = []
if set(regt["Наименование задачи"]) != set(tot_reg.index):
    rb.append("состав задач листа «Задачи для Jira» ≠ строкам табеля")
else:
    rb += [n for n, h in zip(regt["Наименование задачи"], regt["Часов в ноябре"]) if abs(h - tot_reg[n]) > 0.05]
report("R1. реестр построчно совпадает с табелем (ФИО, признак, вид работ, период, часы)", reg_bad)
report("R2. лист «Задачи для Jira»: каждая задача один раз, часы = сумме строк", rb, f"{len(regt)} задач")
print("=== C. Задачи — по каталогу ===")
bad, act, phase, epw, code = [], [], [], [], []
for r in rows:
    nm = r[ix["Наименование задачи"]]
    if nm not in CAT:
        bad.append(f"нет в каталоге: {nm}"); continue
    k = CAT[nm]
    if r[ix["Признак"]] != k["Признак"] or r[ix["Вид работ"]] != k["Вид работ"]:
        bad.append(f"{r[3]}: {nm}")
    if r[ix["Задача"]] != "ADTHPL":
        code.append(f"{r[3]}: {r[ix['Задача']]}")
    if k["Дата окончания работ"].date() < M0 or k["Дата начала работ"].date() > M1:
        act.append(nm)
    if (k["Признак"] == "CAPEX") != (k["Фаза (МСФО)"] == "Разработка — капитализация"):
        phase.append(nm)
    if k["Признак"] == "CAPEX":
        e = epi.loc[k["Эпик"]]
        a, b = [to_d(x) for x in r[ix["Период списания"]].split(" - ")]
        if not (e["Дата начала капитализации"].date() <= a and b <= e["Дата готовности к использованию"].date()):
            epw.append(f"{nm}: {a}–{b}")
report("C1. наименование, признак, вид работ — по каталогу", bad)
report("C2. «Задача» — код проекта ADTHPL без номера", code)
report("C3. все задачи активны в ноябре", act)
report("C4. CAPEX ⇔ фаза «Разработка — капитализация»", phase)
report("C5. CAPEX-работы внутри окна капитализации эпика", epw)

print("=== D. Период списания ===")
bad = []
for r in rows:
    fio, k = r[3], CAT[r[ix["Наименование задачи"]]]
    s, e = max(k["Дата начала работ"].date(), M0), min(k["Дата окончания работ"].date(), M1)
    av = [d for d in wdays(s, e) if d not in absd[fio]]
    if av and any(s <= d < av[0] for d in absd[fio]):
        s = av[0]
    if av and any(av[-1] < d <= e for d in absd[fio]):
        e = av[-1]
    if r[ix["Период списания"]] != f"{s:%d.%m.%Y} - {e:%d.%m.%Y}":
        bad.append(f"{fio}: {r[ix['Период списания']]} ≠ {s:%d.%m.%Y} - {e:%d.%m.%Y}")
report("D1. период = срок задачи ∩ ноябрь, отпуск в начале/конце сдвигает даты; формат «дд.мм.гггг - дд.мм.гггг»", bad)

print("=== E. Емкость ===")
load = defaultdict(float)
for r in rows:
    load[r[ix["Наименование задачи"]]] += r[ix["Списать часов"]]
report("E1. ≤ 350 ч на задачу в месяц", [f"{n}: {v}" for n, v in load.items() if v > 350], f"максимум {max(load.values()):g} ч")
bad, bad2, per = [], [], defaultdict(list)
for r in rows:
    fio, rate = r[3], r[ix["Занятые ставки"]]
    a, b = [to_d(x) for x in r[ix["Период списания"]].split(" - ")]
    cap = rate * sum(dh(d) for d in wdays(a, b) if d not in absd[fio])
    if r[ix["Списать часов"]] > cap + 1e-6:
        bad.append(f"{fio}: {r[ix['Списать часов']]} > {cap:g}")
    per[fio].append([a, b, r[ix["Списать часов"]], rate])
report("E2. часы строки ≤ емкость сотрудника в периоде (8 ч; 3.11 — 7 ч) × ставка", bad)
for fio, jobs in per.items():
    left = [list(j) for j in jobs]
    for d in wdays(M0, M1):
        if d in absd[fio]:
            continue
        capd = dh(d) * jobs[0][3]
        for j in sorted([j for j in left if j[0] <= d <= j[1] and j[2] > 1e-9], key=lambda j: j[1]):
            take = min(capd, j[2]); j[2] -= take; capd -= take
            if capd <= 1e-9:
                break
    if sum(j[2] for j in left) > 1e-6:
        bad2.append(f"{fio}: не помещается {sum(j[2] for j in left):g} ч")
report("E3. расписание по дням выполнимо (EDF)", bad2)

print("=== F. Роли и доли CAPEX/OPEX ===")
MAP = {"Начальник отдела": "Руководитель команды", "Специалист ML": "Ведущий специалист ML"}
FAM = [{"Специалист по тестированию", "Ведущий специалист по тестированию", "Главный специалист по тестированию"},
       {"Аналитик данных", "Ведущий аналитик данных"},
       {"Разработчик", "Ведущий разработчик", "Главный разработчик", "Ведущий разработчик проекта Эдтех"},
       {"Системный аналитик", "Ведущий системный аналитик", "Главный системный аналитик"},
       {"Менеджер продукта", "Ведущий менеджер продукта"}]
SH = {"Руководитель команды": 0.3, "Менеджер продукта": 0.6, "Ведущий менеджер продукта": 0.6}
OPEX_ROLES = {"Директор департамента", "Руководитель управления", "Руководитель направления", "Менеджер проектов",
              "Руководитель проектов", "Ведущий инженер сопровождения", "Инженер сопровождения"}
FORCE_OPEX = {"Гряда Игорь Леонидович"}
act_cat = cat[(cat["Дата начала работ"].dt.date <= M1) & (cat["Дата окончания работ"].dt.date >= M0)]


def fam_of(role):
    return next((f for f in FAM if role in f), {role})


info = defaultdict(lambda: {"C": 0.0, "O": 0.0, "tasks": []})
for r in rows:
    x = info[r[3]]
    x["pos"], x["rate"] = r[ix["Должность сотрудника"]], r[ix["Занятые ставки"]]
    x["C" if r[ix["Признак"]] == "CAPEX" else "O"] += r[ix["Списать часов"]]
    x["tasks"].append((r[ix["Наименование задачи"]], r[ix["Признак"]], r[ix["Списать часов"]]))
bad, exc = [], []
for fio, x in info.items():
    role = MAP.get(x["pos"], x["pos"])
    Ht = x["C"] + x["O"]
    if role in OPEX_ROLES or fio in FORCE_OPEX:
        exp = 0
    else:
        exp = round(Ht * SH[role]) if role in SH else Ht
    if abs(x["C"] - exp) <= 1e-6:
        continue
    # меньше таргета допустимо, только если в остальные дни нет активных CAPEX-задач своей профессии
    own = act_cat[(act_cat["Должность"].isin(fam_of(role))) & (act_cat["Признак"] == "CAPEX")]
    days = {d for _, k in own.iterrows() for d in wdays(max(k["Дата начала работ"].date(), M0), min(k["Дата окончания работ"].date(), M1))}
    coverable = x["rate"] * sum(dh(d) for d in days if d not in absd[fio])
    if x["C"] < exp and abs(x["C"] - min(exp, coverable)) <= 1e-6:
        exc.append(f"{fio} ({x['pos']}): CAPEX {x['C']:g} из {Ht:g} — CAPEX-задачи профессии есть только на {coverable:g} ч")
    else:
        bad.append(f"{fio} ({x['pos']}): CAPEX {x['C']:g} из {Ht:g}, ожидалось {exp:g}")
report("F1. доли: управленцы и сопровождение — 100% OPEX; продакты 60%; руководители команд 30%; остальные 100%; Гряда — OPEX",
       bad, f"исключений по срокам задач: {len(exc)}")
for e in exc:
    print("          INFO", e)
bad, exc2 = [], []
for f, x in info.items():
    nc = sum(1 for t in x["tasks"] if t[1] == "CAPEX")
    if x["C"] > 0 and nc < 2:
        role = MAP.get(x["pos"], x["pos"])
        avail = act_cat[(act_cat["Должность"].isin(fam_of(role))) & (act_cat["Признак"] == "CAPEX")]
        (exc2 if len(avail) < 2 else bad).append(f"{f}: {nc} CAPEX-задача, в ноябре у профессии {len(avail)}")
report("F2. у каждого с CAPEX — не менее 2 CAPEX-задач", bad, f"исключения: {'; '.join(exc2)}" if exc2 else None)
bad = []
for f, x in info.items():
    if x["C"] == 0 and len(x["tasks"]) > 1:
        main = max(x["tasks"], key=lambda t: t[2])[0]
        if load[main] + sum(t[2] for t in x["tasks"] if t[0] != main) <= 350:
            bad.append(f)
report("F3. 100% OPEX — одна задача; вторая только при лимите 350 ч", bad)

print("=== G. Задачи чужой должности ===")
FALLBACK = {"Главный инженер по поддержке оборудования": "Сопровождение промышленной эксплуатации",
            "Дизайнер продукта": "Консультирование смежных команд", "Исследователь пользовательского опыта": "Консультирование смежных команд"}
bad, lst = [], []
for r in rows:
    fio, nm = r[3], r[ix["Наименование задачи"]]
    role, trole = MAP.get(r[ix["Должность сотрудника"]], r[ix["Должность сотрудника"]]), CAT[nm]["Должность"]
    if role == trole:
        continue
    lst.append(f"{fio.split()[0]} ({r[ix['Должность сотрудника']]}) → [{trole}] {r[ix['Списать часов']]:g} ч: {nm}")
    own = act_cat[(act_cat["Должность"] == role) & (act_cat["Признак"] == "CAPEX")]
    fam_ok = trole in fam_of(role) and r[ix["Признак"]] == "CAPEX" and (len(own) < 2 or (own["Дата начала работ"].dt.date > M0).all())
    fb_ok = r[ix["Признак"]] == "OPEX" and nm.startswith(FALLBACK.get(role, "\0"))
    if not (fam_ok or fb_ok):
        bad.append(f"{fio}: {nm}")
report("G1. чужие задачи — та же профессия (своих CAPEX-задач < 2 или начинаются позже) или OPEX-остаток по списку", bad)
for x in lst:
    print("         ", x)

print("=== I. Строки ===")
seen, bad = set(), []
for r in rows:
    key = (r[3], r[ix["Наименование задачи"]])
    if key in seen:
        bad.append(f"дубль {key}")
    seen.add(key)
    h = r[ix["Списать часов"]]
    if not h or (h < 8 and r[ix["Занятые ставки"]] >= 1):
        bad.append(f"мелкая строка {key}: {h}")
    if isinstance(h, float) and h.is_integer():
        bad.append(f"«.0» {key}")
report("I1. нет дублей, нулевых и < 8 ч строк, целые часы без «.0»", bad)
bad = [f for f in order_out if [r[ix["Списать часов"]] for r in rows if r[3] == f] !=
       sorted([r[ix["Списать часов"]] for r in rows if r[3] == f], reverse=True)]
report("I2. строки сотрудника — по убыванию часов", bad)

T = sum(tot.values()); C = sum(r[ix["Списать часов"]] for r in rows if r[ix["Признак"]] == "CAPEX")
print(f"\nИТОГО: {len(order_out)} чел., {len(rows)} строк, {len(load)} задач, {T:g} ч; CAPEX {C:g} ч ({C / T:.1%}), OPEX {T - C:g} ч")

if HR:   # справочно: таргет капитализации по HR-группировкам (середина вилки зарплатной группы)
    sal, grp = {}, {}
    for f in HR:
        w = openpyxl.load_workbook(f, read_only=True, data_only=True)
        rr = list(w.worksheets[0].iter_rows(values_only=True))
        bands = {r[10]: (r[11], r[12]) for r in rr[14:30] if isinstance(r[10], int)}
        for r in rr[14:]:
            if r[0] and "AdTech" in str(r[0]) and r[8] in bands:
                sal[r[2]] = sum(bands[r[8]]) / 2
        calc = list(w.worksheets[1].iter_rows(values_only=True))
        tgt = next(r[5] for r in calc if r[1] == "CAPEX")
    BOSS = {"Директор департамента", "Руководитель управления", "Руководитель направления"}
    money = defaultdict(lambda: [0.0, 0.0, 0.0, 0.0])       # CAPEX, OPEX, CAPEX без руководителей, OPEX без них
    for r in rows:
        fio = r[3]
        if fio not in sal:
            continue
        m = sal[fio] * r[ix["Списать часов"]] / 159
        g = src_by[fio][0]
        j = 0 if r[ix["Признак"]] == "CAPEX" else 1
        money[g][j] += m
        if r[ix["Должность сотрудника"]] not in BOSS:
            money[g][2 + j] += m
    print("\nСПРАВОЧНО — доля CAPEX в ФОТ по часам (середина вилки HR-группы), таргет МСФО-модели: CAPEX = "
          f"{tgt:.0%} списаний (70% всего ФОТ с руководителями и отпусками):")
    print("   группа      все      без директора и руководителей управлений/направлений")
    for g, v in money.items():
        print(f"   {g:10s} {v[0] / (v[0] + v[1]):6.1%}   {v[2] / (v[2] + v[3]):6.1%}")
    t = [sum(v[i] for v in money.values()) for i in range(4)]
    print(f"   ИТОГО      {t[0] / (t[0] + t[1]):6.1%}   {t[2] / (t[2] + t[3]):6.1%}"
          f"   (без {', '.join(sorted(set(src_by) - set(sal)))} — нет в HR-группировках)")
print("Проверок:", len(errors), "| с ошибками:", sum(1 for _, k in errors if k))
