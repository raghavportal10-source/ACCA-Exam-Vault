def df(r, n):
    return round((1 + r) ** -n, 3)


def q31(cost, scrap, r=.11, tax=.25, ovh=0):
    """Q31 NPV with every line item rounded to the nearest $ before summing."""
    vol = [60000, 85000, 95000, 70000]
    price = [48 * 1.03 ** t for t in range(1, 5)]
    vc = [27 * 1.05 ** t for t in range(1, 5)]
    sales = [round(v * p) for v, p in zip(vol, price)]
    vcs = [round(v * c) for v, c in zip(vol, vc)]
    fc = [round(520000 * 1.04 ** t + ovh * 1.04 ** t) for t in range(1, 5)]
    op = [s - c - f for s, c, f in zip(sales, vcs, fc)]
    taxop = [-round(o * tax) for o in op]
    w = cost
    tad = []
    for _ in range(3):
        a = round(w * .25)
        tad.append(a)
        w -= a
    tad.append(w - scrap)
    taxtad = [round(a * tax) for a in tad]
    wcl = [round(.10 * s) for s in sales]
    wcf = [-wcl[0]] + [-(wcl[i] - wcl[i - 1]) for i in range(1, 4)] + [wcl[3]]
    rows = {'sales': sales, 'vc': vcs, 'fc': fc, 'op': op, 'taxop': taxop, 'tad': tad,
            'taxtad': taxtad, 'wc': wcf, 'wcl': wcl}
    cf = [0] * 6
    cf[0] = -cost
    for t in range(1, 5):
        cf[t] += op[t - 1]
    for t in range(2, 6):
        cf[t] += taxop[t - 2] + taxtad[t - 2]
    cf[4] += scrap
    for t in range(5):
        cf[t] += wcf[t]
    pv = [round(cf[t] * (df(r, t) if t else 1)) for t in range(6)]
    return sum(pv), rows, cf, pv, w
