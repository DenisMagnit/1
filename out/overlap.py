"""Проверка выполнимости по пересекающимся периодам: для любого отрезка [a,b] из дат начала/окончания
сумма часов задач, чей период работ целиком внутри [a,b], не больше 8 ч × ставка × рабочие дни
сотрудника без отпуска в [a,b]."""
from datetime import timedelta, datetime, date
import pandas as pd

def wd(a, b):
    out, x = [], a
    while x <= b:
        if x.weekday() < 5:
            out.append(x)
        x += timedelta(1)
    return out

def violations(rows, absd):
    """rows: [(fio, rate, h, start, end)] -> [(fio, a, b, hours, cap)]"""
    by = {}
    for fio, rate, h, s, e in rows:
        by.setdefault(fio, []).append((rate, h, s, e))
    out = []
    for fio, lst in by.items():
        rate = lst[0][0]
        starts = sorted({s for _, _, s, _ in lst}); ends = sorted({e for _, _, _, e in lst})
        worst = None
        for a in starts:
            for b in ends:
                if b < a:
                    continue
                h = sum(x[1] for x in lst if a <= x[2] and x[3] <= b)
                cap = 8 * rate * len([d for d in wd(a, b) if d not in absd.get(fio, set())])
                if h > cap + 0.01 and (worst is None or h - cap > worst[3] - worst[4]):
                    worst = (fio, a, b, h, cap)
        if worst:
            out.append(worst)
    return out
