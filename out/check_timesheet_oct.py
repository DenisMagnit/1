import sys
from datetime import date, datetime, timedelta
from collections import defaultdict
import pandas as pd

IN, OUT = sys.argv[1], sys.argv[2]
M0, M1 = date(2026, 10, 1), date(2026, 10, 31)

def d(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NaT:
        return None
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.date()
    return datetime.strptime(str(v).strip(), "%d.%m.%Y").date()

def wd(a, b):
    out, x = [], a
    while x <= b:
        if x.weekday() < 5:
            out.append(x)
        x += timedelta(1)
    return out

o = pd.read_excel(OUT, sheet_name="Лист1")
src = pd.read_excel(f"{IN}/5_oct_src.xlsx", sheet_name="Лист1")
cat = pd.read_excel(f"{IN}/3_tasks.xlsx", sheet_name="Задачи")
epi = pd.read_excel(f"{IN}/3_tasks.xlsx", sheet_name="Эпики").set_index("Эпик")
rep = pd.read_excel(f"{IN}/4_replace.xlsx").dropna(subset=["Наименование задачи старое"])
REPL = {a.strip(): c.strip() for a, c in zip(rep.iloc[:, 0], rep.iloc[:, 2])}
cat["num"] = [f"ADTHPL-{i + 101}" for i in cat.index]
cat = cat.set_index("num")

res = []
def check(name, errs):
    res.append((name, len(errs)))
    print(("OK   " if not errs else "FAIL ") + name + ("" if not errs else f": {len(errs)}"))
    for e in errs[:10]:
        print("      ", e)

# 1. состав сотрудников
so, oo = set(src["ФИО сотрудника"]), set(o["ФИО сотрудника"])
check("1. все 105 сотрудников исходника на месте, лишних нет",
      [f"нет: {x}" for x in so - oo] + [f"лишний: {x}" for x in oo - so])

# 2. исходные колонки A–Q не изменены
cols = list(src.columns[:17])
s_idx = src.set_index("ФИО сотрудника")
errs = []
for _, r in o.iterrows():
    sr = s_idx.loc[r["ФИО сотрудника"]]
    for c in cols:
        if c in ("ФИО сотрудника", "Списать часов"):
            continue
        a, b = r[c], sr[c]
        if not ((pd.isna(a) and pd.isna(b)) or a == b):
            errs.append(f"{r['ФИО сотрудника']}: {c} {a!r} != {b!r}")
check("2. данные исходника (колонки A–Q) перенесены без изменений", errs)

# 3. сумма часов = «Списать часов»
g = o.groupby("ФИО сотрудника")["Списать часов"].sum()
errs = [f"{n}: {g[n]} vs {s_idx.loc[n, 'Списать часов']}" for n in g.index
        if abs(g[n] - s_idx.loc[n, "Списать часов"]) > 0.01]
check("3. сумма часов по строкам = «Списать часов» по каждому сотруднику", errs)

# 4. не менее 2 задач — для сотрудников с CAPEX (капитализируемый ФОТ)
c = o.groupby("ФИО сотрудника")["Задача"].nunique()
hascx = set(o.loc[o["Признак"] == "CAPEX", "ФИО сотрудника"])
check("4. у сотрудников с CAPEX (%d чел.) не менее 2 задач" % len(hascx),
      [f"{n}: {c[n]}" for n in c.index if n in hascx and c[n] < 2])
print("INFO    100%% OPEX: %d чел., из них с одной задачей: %d" % (len(c) - len(hascx), sum(1 for n in c.index if n not in hascx and c[n] < 2)))

# 5. дубликаты сотрудник+задача, пустые/нулевые часы
dup = o[o.duplicated(["ФИО сотрудника", "Задача"], keep=False)]
z = o[~(o["Списать часов"] > 0)]
check("5. нет дублей «сотрудник + задача» и строк с нулём часов",
      [f"дубль {r['ФИО сотрудника']} {r['Задача']}" for _, r in dup.iterrows()] +
      [f"ноль {r['ФИО сотрудника']} {r['Задача']}" for _, r in z.iterrows()])

# 6. лимит 350 ч на задачу
t = o.groupby("Задача")["Списать часов"].sum()
check("6. не более 350 ч на задачу в месяц (макс. %.0f ч)" % t.max(), [f"{n}: {t[n]}" for n in t.index if t[n] > 350])

# 7. реквизиты задачи = каталог (номер = строка + 99, с учетом замен)
errs = []
for _, r in o.iterrows():
    n = r["Задача"]
    if n not in cat.index:
        errs.append(f"нет в каталоге {n}"); continue
    k = cat.loc[n]
    nm = REPL.get(k["Наименование задачи"].strip(), k["Наименование задачи"].strip())
    for col, exp in (("Наименование задачи", nm), ("Признак", k["Признак"]), ("Вид работ", k["Вид работ"])):
        if str(r[col]).strip() != str(exp).strip():
            errs.append(f"{r['ФИО сотрудника']} {n}: {col}")
check("7. номер, наименование, признак, вид работ совпадают с каталогом (+ замены)", errs)

# 8. даты работ внутри срока задачи и октября, начало <= конец, не в отпуске
errs = []
absd = {}
for _, sr in src.iterrows():
    s1, s2 = sr.iloc[15], sr.iloc[16]
    days = set()
    if not pd.isna(s1):
        ss = str(s1).split("/") if isinstance(s1, str) else [s1]
        ee = str(s2).split("/") if isinstance(s2, str) else [s2]
        for a, b in zip(ss, ee):
            days |= set(wd(d(a), d(b)))
    absd[sr["ФИО сотрудника"]] = days
for _, r in o.iterrows():
    k = cat.loc[r["Задача"]]
    a, b = d(r["Дата начала работ над задачей"]), d(r["Дата окончания работ над задачей"])
    ks, ke = d(k["Дата начала работ"]), d(k["Дата окончания работ"])
    if not (a and b and a <= b and M0 <= a and b <= M1 and ks <= a and b <= ke):
        errs.append(f"{r['ФИО сотрудника']} {r['Задача']}: {a}–{b}, задача {ks}–{ke}")
    if a in absd[r["ФИО сотрудника"]] or b in absd[r["ФИО сотрудника"]]:
        errs.append(f"{r['ФИО сотрудника']} {r['Задача']}: дата работ попадает на отпуск")
check("8. даты работ в пределах срока задачи и октября, не начинаются/не кончаются в отпуске", errs)

# 9. часы на задачу не больше 8 ч × ставка × рабочих дней в периоде работ
errs = []
for _, r in o.iterrows():
    a, b = d(r["Дата начала работ над задачей"]), d(r["Дата окончания работ над задачей"])
    days = [x for x in wd(a, b) if x not in absd[r["ФИО сотрудника"]]]
    cap = 8 * r["Занятые ставки"] * len(days)
    if r["Списать часов"] > cap + 0.01:
        errs.append(f"{r['ФИО сотрудника']} {r['Задача']}: {r['Списать часов']} > {cap}")
check("9. часы по задаче ≤ 8 ч × ставка × раб. дни сотрудника в периоде работ", errs)

# 10. по каждому сотруднику общий объем ≤ 8 ч × рабочие дни без отпуска
errs = []
for n, h in g.items():
    sr = s_idx.loc[n]
    cap = 8 * sr["Занятые ставки"] * len([x for x in wd(M0, M1) if x not in absd[n]])
    if h > cap + 0.01:
        errs.append(f"{n}: {h} > {cap}")
check("10. часы сотрудника ≤ норме октября за вычетом отпуска (176 ч = 22 дня)", errs)

# 11. CAPEX — только внутри окна капитализации эпика
errs = []
for _, r in o[o["Признак"] == "CAPEX"].iterrows():
    k = cat.loc[r["Задача"]]
    e = epi.loc[k["Эпик"]]
    a, b = d(r["Дата начала работ над задачей"]), d(r["Дата окончания работ над задачей"])
    if not (d(e["Дата начала капитализации"]) <= a and b <= d(e["Дата готовности к использованию"])):
        errs.append(f"{r['Задача']} {k['Эпик']}")
check("11. CAPEX-работы внутри окна «начало капитализации — готовность» эпика (IAS 38)", errs)

# 12. доли CAPEX/OPEX по ролям
OPEX_ONLY = {"Директор департамента", "Руководитель управления", "Руководитель направления",
             "Менеджер проектов", "Руководитель проектов", "Ведущий инженер сопровождения",
             "Инженер сопровождения", "Ведущий разработчик проекта Эдтех"}
SH = {"Руководитель команды": 0.3, "Начальник отдела": 0.3, "Менеджер продукта": 0.6, "Ведущий менеджер продукта": 0.6}
OPEX_ONLY_2ND = {"Главный инженер по поддержке оборудования"}
errs = []
for n, grp in o.groupby("ФИО сотрудника"):
    pos = grp["Должность сотрудника"].iloc[0]
    H = grp["Списать часов"].sum()
    cx = grp.loc[grp["Признак"] == "CAPEX", "Списать часов"].sum()
    exp = 0 if pos in OPEX_ONLY else SH.get(pos, 1.0)
    if pos in OPEX_ONLY_2ND:   # вторая задача — OPEX, т.к. CAPEX-задач должности на октябрь одна
        exp = cx / H if 0.6 <= cx / H <= 0.8 else exp
    if abs(cx - H * exp) > 1:
        errs.append(f"{n} ({pos}): CAPEX {cx} из {H}, ожидалось {exp:.0%}")
check("12. доли CAPEX по ролям: продакты 60%, руководители команд 30%, менеджмент/сопровождение 0%, остальные 100%", errs)

# 13. должность сотрудника соответствует должности задачи в каталоге
MAP = {"Начальник отдела": "Руководитель команды", "Специалист ML": "Ведущий специалист ML"}
mism = o[[MAP.get(p, p) != cat.loc[n, "Должность"] for p, n in zip(o["Должность сотрудника"], o["Задача"])]]
check("13. задачи чужой должности каталога — только OPEX (CAPEX не заимствуем)",
      [f"{r['ФИО сотрудника']} {r['Задача']}" for _, r in mism.iterrows() if r["Признак"] == "CAPEX"])
print(f"INFO    строк с задачей другой должности каталога (все OPEX): {len(mism)}")
for _, r in mism.iterrows():
    print(f"       {r['ФИО сотрудника']} ({r['Должность сотрудника']}) → {r['Задача']} "
          f"[{cat.loc[r['Задача'], 'Должность']}] {r['Списать часов']} ч")

tot = o["Списать часов"].sum()
cx = o.loc[o["Признак"] == "CAPEX", "Списать часов"].sum()
print(f"\nИТОГО: {o['ФИО сотрудника'].nunique()} чел., {len(o)} строк, {o['Задача'].nunique()} задач, "
      f"{tot:.1f} ч, CAPEX {cx:.0f} ч ({cx / tot:.1%}), OPEX {tot - cx:.1f} ч")
print("Распределение числа задач на человека:", o.groupby("ФИО сотрудника")["Задача"].nunique().value_counts().sort_index().to_dict())
print("Проверок с ошибками:", sum(1 for _, k in res if k))
