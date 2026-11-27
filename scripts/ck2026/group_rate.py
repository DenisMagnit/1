import sys
sys.path.insert(0, '.')
from ck_model import qrate, yrate, D, M
def split(D, lmk_market=None):
    T = Mk = 0; rows = {}
    for k, (t9, tq, mq, cl) in D.items():
        if k == 'ЛМК' and lmk_market: mq = lmk_market
        r = qrate(tq + mq, cl); rT = qrate(tq, cl); rM = qrate(mq, cl)
        Yt = t9 + tq; Y = Yt + mq; y = yrate(Y)
        dT = (r - rT) * tq + (y - yrate(Yt)) * Yt
        dM = (r - rM) * mq + (y - yrate(mq)) * mq
        rows[k] = (round(dT / M, 3), round(dM / M, 3)); T += dT; Mk += dM
    return T / M, Mk / M, rows
for lmk in (None, 4781590.0):
    t, m, rows = split(D, lmk)
    print('ЛМК', lmk, 'Тандер +%.2f  Маркет +%.2f  итого %.2f' % (t, m, t + m), rows)
