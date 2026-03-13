#!/usr/bin/env python3
"""
kpi_ndvi_senslope.py — Calcula Sen's slope + Mann-Kendall por tramo
Uso: python3 kpi_ndvi_senslope.py KPI_NDVI_CostaUruguay_2017_2024.csv
Salida: KPI_NDVI_tendencia.csv (misma carpeta)
"""
import sys, csv, math
from pathlib import Path
from itertools import combinations

ANOS = list(range(2017, 2025))
MIN_PIXELS = 5      # descartar año si NDVI_count < MIN_PIXELS
MIN_ANOS   = 4      # mínimo años válidos para calcular tendencia
ALPHA      = 0.05   # umbral p-value
NDVI_UMBRAL = 0.02  # cambio mínimo absoluto para clasificar como mejora/degradación

# ── Mann-Kendall ───────────────────────────────────────────────
def mann_kendall(y):
    n = len(y)
    s = 0
    for i in range(n-1):
        for j in range(i+1, n):
            s += (1 if y[j]>y[i] else -1 if y[j]<y[i] else 0)
    var_s = n*(n-1)*(2*n+5)/18.0
    if var_s == 0: return 0, 1.0
    z = (s-1)/math.sqrt(var_s) if s>0 else (s+1)/math.sqrt(var_s) if s<0 else 0
    # p-value aproximado (distribución normal)
    p = 2*(1 - _norm_cdf(abs(z)))
    return s, p

def _norm_cdf(x):
    # Aproximación Abramowitz & Stegun
    t = 1/(1+0.2316419*x)
    poly = t*(0.319381530 + t*(-0.356563782 + t*(1.781477937 + t*(-1.821255978 + t*1.330274429))))
    return 1 - (1/math.sqrt(2*math.pi))*math.exp(-0.5*x*x)*poly

# ── Sen's slope ────────────────────────────────────────────────
def sens_slope(x, y):
    slopes = []
    for (i,xi),(j,xj) in combinations(enumerate(x), 2):
        if xj != xi:
            slopes.append((y[j]-y[i])/(xj-xi))
    if not slopes: return None
    slopes.sort()
    n = len(slopes)
    return slopes[n//2] if n%2==1 else (slopes[n//2-1]+slopes[n//2])/2

# ── Main ───────────────────────────────────────────────────────
def main():
    if len(sys.argv) < 2:
        print(__doc__); sys.exit(1)

    fpath = Path(sys.argv[1])
    print(f"Leyendo: {fpath}")

    # Leer CSV y organizar por tramo
    tramos = {}
    with open(fpath, newline='', encoding='utf-8-sig') as f:
        for r in csv.DictReader(f):
            idx = int(r['t'])
            if idx not in tramos:
                tramos[idx] = {'d':r['d'],'v':r['v'],'a':r['a'],'l':r['l'],'serie':{}}
            anio = int(r['anio'])
            med  = r.get('NDVI_median','')
            cnt  = r.get('NDVI_count','')
            if med not in ('','None','nan') and cnt not in ('','None','nan'):
                m = float(med)
                c = int(float(cnt))
                if m != -9999 and c >= MIN_PIXELS:
                    tramos[idx]['serie'][anio] = m

    print(f"Tramos cargados: {len(tramos)}")

    # Calcular tendencia por tramo
    resultados = []
    for idx in sorted(tramos):
        d = tramos[idx]
        serie = d['serie']
        anos_val = sorted(a for a in ANOS if a in serie)
        vals_val = [serie[a] for a in anos_val]
        n_val = len(anos_val)

        # Valores NDVI por año (None si sin dato)
        ndvi_por_ano = {a: serie.get(a, None) for a in ANOS}

        if n_val < MIN_ANOS:
            row = {'t':idx,'d':d['d'],'v':d['v'],'a':d['a'],'l':d['l'],
                   'sens_slope':'','mk_p':'','mk_S':'',
                   'tendencia':'insuficiente','n_anios_val':n_val}
        else:
            slope = sens_slope(anos_val, vals_val)
            mk_s, mk_p = mann_kendall(vals_val)
            cambio = slope * (anos_val[-1] - anos_val[0]) if slope else 0

            if mk_p < ALPHA and abs(cambio) >= NDVI_UMBRAL:
                tend = 'mejora' if slope > 0 else 'degradacion'
            else:
                tend = 'estable'

            row = {'t':idx,'d':d['d'],'v':d['v'],'a':d['a'],'l':d['l'],
                   'sens_slope':round(slope,6) if slope else '',
                   'mk_p':round(mk_p,5),'mk_S':int(mk_s),
                   'tendencia':tend,'n_anios_val':n_val}

        for a in ANOS:
            row[f'NDVI_{a}'] = round(ndvi_por_ano[a],5) if ndvi_por_ano[a] is not None else ''
        resultados.append(row)

    # Escribir CSV
    outpath = fpath.parent / 'KPI_NDVI_tendencia.csv'
    cols = ['t','d','v','a','l','sens_slope','mk_p','mk_S','tendencia','n_anios_val'] + \
           [f'NDVI_{a}' for a in ANOS]
    with open(outpath, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(resultados)

    # Resumen
    from collections import Counter
    tends = Counter(r['tendencia'] for r in resultados)
    print(f"\nResultados → {outpath}")
    print(f"  Total tramos: {len(resultados)}")
    for k in ['mejora','degradacion','estable','insuficiente']:
        n = tends.get(k,0)
        print(f"  {k:14s}: {n:3d} ({100*n/len(resultados):.1f}%)")

if __name__ == '__main__':
    main()
