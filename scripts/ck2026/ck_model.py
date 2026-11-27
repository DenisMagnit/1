"""Independent recompute of the СК 2026 loss (Тандер + Маркет) from 'Данные 2026' values captured in the exported session (26.11.2026)."""
import math
M = 1e6
def qrate(B, cl, steps='prop'):
    if B < 2 * M: return 0.0
    over = max(0, B - 5 * M) / M
    if steps == 'int': over = math.floor(over)
    return min(0.33, 0.18 + over * 0.006) + min(0.05, max(0, cl - 1) * 0.01)
def yrate(Y):
    if Y < 120 * M: return 0.0
    return min(0.15, 0.05 + math.floor((Y - 120 * M) / (20 * M)) * 0.01)
# agency: (Тандер 9M, Тандер Q4, Маркет Q4, клиенты Q4 Тандера)
D = {'Оккам Си Эс': (1285745.01, 32673553.19, 60974277.00, 3),
     'РАРС Медиа': (96891710.98, 34136889.17, 9000000.00, 5),
     'Изи Коммерс': (7873673.53, 6505264.84, 19254062.92, 2),
     'МедиаСистем': (1200000.00, 10176000.00, 17008000.00, 1),
     'ЛМК': (3304997.99, 773385.49, 3120431.00, 2)}
def loss(D, steps='prop', lmk_market=None):
    tq = ty = 0; out = {}
    for k, (t9, tq4, mq4, cl) in D.items():
        if k == 'ЛМК' and lmk_market: mq4 = lmk_market
        q = qrate(tq4 + mq4, cl, steps) * (tq4 + mq4) - qrate(tq4, cl, steps) * tq4 - qrate(mq4, cl, steps) * mq4
        Y_t = t9 + tq4; Y = Y_t + mq4
        y = yrate(Y) * Y - yrate(Y_t) * Y_t - yrate(mq4) * mq4
        out[k] = (q, y); tq += q; ty += y
    return tq, ty, out
for steps in ('prop', 'int'):
    for lmk in (None, 4781590.0):
        q, y, out = loss(D, steps, lmk)
        print(f'{steps:4} ЛМК Маркет {"3.12" if lmk is None else "4.78"}: Q4 {q/M:6.3f}  год {y/M:6.3f}  итого {(q+y)/M:6.3f}', {k: round((a+b)/M, 2) for k, (a, b) in out.items()})
