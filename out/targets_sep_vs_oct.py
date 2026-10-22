"""Сравнение выполнения таргета капитализации: сентябрь (факт) vs октябрь (файл).
ФОТ оценивается по середине вилки зарплатной группы из HR-группировок;
CAPEX-деньги = оклад × ставка × CAPEX-часы / норма месяца (176 ч), отпуск = OPEX."""
import sys
import openpyxl
import pandas as pd

IN, HR, OCT = sys.argv[1], sys.argv[2], sys.argv[3]

info = {}
for f, lab in [("360ad8ac", "ИТ"), ("1c730ac8", "Аналитика"), ("025c54e7", "Продукт")]:
    ws = openpyxl.load_workbook(f"{HR}/{f}.xlsx", read_only=True, data_only=True).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    b = {r[10]: (r[11], r[12]) for r in rows[14:30] if isinstance(r[10], int)}
    for r in rows[14:]:
        if r[0] and "AdTech" in str(r[0]):
            lo, hi = b[r[8]]
            info[r[2]] = (lab, (lo + hi) / 2)

OPEX_ONLY = {"Директор департамента", "Руководитель управления", "Руководитель направления",
             "Менеджер проектов", "Руководитель проектов", "Ведущий инженер сопровождения",
             "Инженер сопровождения", "Ведущий разработчик проекта Эдтех"}

def role_grp(pos):
    if pos in ("Руководитель команды", "Начальник отдела"):
        return "Руководители команд"
    if pos in ("Менеджер продукта", "Ведущий менеджер продукта"):
        return "Продакты"
    if pos in OPEX_ONLY:
        return "Менеджмент и сопровождение"
    return "Разработка, тест, аналитика, дизайн"

def frame(rows_df, fio, pos, rate, h, flag):
    g = rows_df.groupby(fio)
    out = []
    for n, x in g:
        out.append(dict(fio=n, pos=x[pos].iloc[0], rate=float(x[rate].iloc[0]),
                        H=x[h].sum(), C=x.loc[x[flag] == "CAPEX", h].sum()))
    d = pd.DataFrame(out)
    d["norm"] = 176 * d.rate
    d["blk"] = d.fio.map(lambda n: info.get(n, ("нет в группировках", None))[0])
    d["mid"] = d.fio.map(lambda n: info.get(n, (None, None))[1])
    d["grp"] = d.pos.map(role_grp)
    return d

sep = pd.read_excel(f"{IN}/2_sep_done.xlsx", sheet_name="Лист1", header=None, skiprows=1)
S = frame(sep, 3, 4, 12, 13, 17)
o = pd.read_excel(OCT, sheet_name="Лист1")
O = frame(o, "ФИО сотрудника", "Должность сотрудника", "Занятые ставки", "Списать часов", "Признак")

def metrics(d):
    m = d.dropna(subset=["mid"])
    fot = (m.mid * m.rate).sum()
    capx = (m.mid * m.rate * m.C / m.norm).sum()
    capx_worked = (m.mid * m.rate * m.C / m.H).sum()
    return {
        "Сотрудников": len(d),
        "из них в HR-группировках": len(m),
        "CAPEX, % списанных часов (все)": 100 * d.C.sum() / d.H.sum(),
        "CAPEX, % списанных часов (группировки; таргет 77%)": 100 * m.C.sum() / m.H.sum(),
        "Отсутствия, % нормы": 100 * (1 - d.H.sum() / d.norm.sum()),
        "CAPEX, % ФОТ без учета отпусков": 100 * capx_worked / fot,
        "CAPEX, % ФОТ, отпуск = OPEX (таргет 70%)": 100 * capx / fot,
        "CAPEX, млн руб./мес": capx / 1e6,
        "Недобор до 70% ФОТ, млн руб./мес": max(0, 0.7 * fot - capx) / 1e6,
    }

res = pd.DataFrame({"Сентябрь": metrics(S), "Октябрь": metrics(O)}).round(1)
print(res.to_string())

def by(d, key):
    m = d.dropna(subset=["mid"]).copy()
    m["fot"] = m.mid * m.rate
    m["capx"] = m.mid * m.rate * m.C / m.norm
    g = m.groupby(key)[["fot", "capx", "C", "H"]].sum()
    return pd.DataFrame({"% ФОТ": 100 * g.capx / g.fot, "% часов": 100 * g.C / g.H})

for key in ("blk", "grp"):
    t = pd.concat({"Сентябрь": by(S, key), "Октябрь": by(O, key)}, axis=1).round(1)
    print("\n", t.to_string())

left = sorted(set(S.fio) - set(O.fio)); new = sorted(set(O.fio) - set(S.fio))
print("\nУшли после сентября:", left)
print("Пришли в октябре:", new)
print("Сентябрь: людей без вилки (нет в группировках):", sorted(S[S.mid.isna()].fio))
