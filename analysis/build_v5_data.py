"""Comparison data for deck v5: top 3 models per brand per class (Jan 2023 - 14 Aug 2026) next to the closest F.lli Ferrari model.
Usage: python3 analysis/build_v5_data.py <Crane_Market_GVW_Pairing_v11_Noir.xlsx>  ->  slides/v5_compare.json

Price basis (user rules, 29 Sep 2026):
- Competitor = average import unit price x JISDOR x (1 + duty: 0% China, 5% Europe) x (1 + PPN 11% + PPh 22 2.5%).
- Ferrari   = crane price (K - J) of the TSP price list Rev1 June'26 (data/ferrari_price_list_rev1.json), minus TSP margin 13.6%.
"""
import sys, json, math
import pandas as pd

FX, MARGIN = 17803, 0.136
EXCL_BRANDS = ['ZOOMLION', 'HYVA']
EXCL_MODELS = ['PALFINGER PK 33002', 'PALFINGER PK 36050']      # USD 2.000-2.150 per unit: not a complete crane
ORIGIN = {'SANY PALFINGER': 'china', 'XCMG': 'china'}             # all other brands: Europe
FAMILY = {'SPK32080B': 'SPK32080', 'SPK32080C': 'SPK32080', 'SPK36080C': 'SPK36080', 'SPK42502C': 'SPK42502',
          'SPK42502E': 'SPK42502', 'PK 41002 EH C': 'PK 41002', 'PK 41002 MH C': 'PK 41002',
          'PK 53002 SH B': 'PK 53002', 'PK 53002 SH C': 'PK 53002'}
K, T = 'Kategori Max Lifting Moment', 'Max Lifting Moment asli (tm)'

d = pd.read_excel(sys.argv[1], sheet_name='8. Pivot Source')
d = d[~d.Brand.isin(EXCL_BRANDS) & ~d.Model.isin(EXCL_MODELS)]
d = d[~(d[K].isin(['Medium', 'Heavy']) & (d['Unit Price (USD)'] < 10000))]   # below USD 10.000: not a complete Medium/Heavy crane
d['m'] = [FAMILY.get(m.replace(b + ' ', ''), m.replace(b + ' ', '')) for b, m in zip(d.Brand, d.Model)]
# SPK38080 (2023, 3 units) is not a Sany Palfinger production model; same importer and unit price as SPK36080 -> treated as SPK36080
d.loc[d.m == 'SPK38080', ['m', 'Max Lifting Moment asli (tm)']] = ['SPK36080', 36.0]
taxed = lambda usd, b: usd * FX * (1.0 if ORIGIN.get(b) == 'china' else 1.05) * (1 + 0.11 + 0.025)

def band(tm, cls):
    if cls == 'Heavy' and tm > 55: return '>55'
    hi = math.ceil(tm / 5) * 5
    return f'>{hi - 5}–{hi}'
BANDS = {'Medium': ['>25–30', '>30–35', '>35–40', '>40–45'], 'Heavy': ['>45–50', '>50–55', '>55']}

# ---- Ferrari line-up from the Rev1 price list (June'26): crane price = K - J, then less TSP margin
cat = {m['model'].replace(' ', ''): m['max_lifting_moment_tm'] for m in json.load(open('data/ferrari_catalogue.json'))['models']}
pl = json.load(open('data/ferrari_price_list_rev1.json'))['models']
PICK = ['268 A4', 'FBR350R A4', '746 A4', '749R A4', 'FBR450R A4', 'FBR600R A4', 'FBR660R A4']   # R/A4 variant per model
fer = []
for r in pl:
    if r['ferrari_model'] in PICK:
        fer.append(dict(model=r['ferrari_model'], tm=cat[r['catalogue_model']], price_idr=r['crane_price_idr'] * (1 - MARGIN), estimate=False,
                        leadtime_weeks=r['leadtime_weeks'], desc=r['desc'], origin=r['incoterm'].replace('EXW ', '').title(), photo=None))
# 7441C: not in the price list -> interpolate the crane price (K - J) between FBR350R A4 and 746 A4 by tm
a = next(r for r in pl if r['ferrari_model'] == 'FBR350R A4'); b = next(r for r in pl if r['ferrari_model'] == '746 A4')
ta, tb = cat[a['catalogue_model']], cat[b['catalogue_model']]
est = a['crane_price_idr'] + (b['crane_price_idr'] - a['crane_price_idr']) / (tb - ta) * (cat['7441C'] - ta)
fer.append(dict(model='7441C', tm=cat['7441C'], price_idr=est * (1 - MARGIN), estimate=True, leadtime_weeks=None, desc=None, origin='Italy', photo=None))
PHOTO_F = {'268 A4': '268', 'FBR350R A4': 'fbr350r', '7441C': '7441c', '746 A4': '746r', '749R A4': '749r',
           'FBR450R A4': 'fbr450r', 'FBR600R A4': 'fbr600r', 'FBR660R A4': 'fbr660r'}
for f in fer: f['photo'] = PHOTO_F[f['model']]
fer.sort(key=lambda f: f['tm'])

PHOTO_C = {'SPK36080': 'spk36080', 'SPK32080': 'spk32080', 'SPK42502': 'spk42502', 'V825': 'v825', 'V933': 'v933', 'V946B': 'v946b',
           'KSQZ300.3': 'ksqz300', 'GSQZ330.4': 'gsqz330', 'SQZ325.4': 'sqz325', 'PK 41002': 'pk41002', 'PK 32080 C': 'pk32080c',
           'X-CLX 388': 'xclx388', '477 EP-5': '477ep5', 'X-CLX 328': 'xclx328', 'PK 53002': 'pk53002', 'PK 76002 EH D': 'pk76002',
           'PK 88002 EH C': 'pk88002', 'V950': 'v950', 'GSQZ460.4': 'gsqz460', 'GSQZ880.6': 'gsqz880', 'GSQZ860.6': 'gsqz860',
           'X-HIPRO B58': 'xhipro_b58', 'EFFER 525H': 'effer525h', 'F485RA.2.23': 'f485ra', 'SPK61502': 'spk61502',
           'KSQZ400.4': 'ksqz400', 'GSQZ400.4': 'gsqz400', 'SQZ365.4': 'sqz365', 'PK 48002': 'pk48002',
           'SQ12ZK3Q': 'sq12zk3q', 'KSQZ300.4': 'ksqz3004', 'KSQZ365.4': 'ksqz3654', 'SQZ400': 'sqz400'}
# one F.lli Ferrari model per tm class (user choice, 30 Sep 2026); >50–55 has no model -> nearest FBR450R A4
BAND_FER = {'>25–30': '268 A4', '>30–35': 'FBR350R A4', '>35–40': '7441C', '>40–45': '746 A4',
            '>45–50': 'FBR450R A4', '>50–55': 'FBR450R A4', '>55': 'FBR660R A4'}
FERD = {f['model']: f for f in fer}

out = {'margin': MARGIN, 'fx': FX, 'ferrari': fer, 'classes': {}}
for cls, lo, hi in [('Medium', 25, 45), ('Heavy', 45, 999)]:
    s = d[d[K] == cls]
    g = s.groupby(['Brand', 'm']).agg(tm=(T, 'max'), q=('Quantity', 'sum'), v=('Total Value (USD)', 'sum')).reset_index()
    line = [f for f in fer if lo < f['tm'] <= hi]
    band_units = {k: 0 for k in BANDS[cls]}
    for _, r in g.iterrows(): band_units[band(r.tm, cls)] += int(r.q)
    brands = []
    for br, x in sorted(g.groupby('Brand'), key=lambda t: (-t[1].q.sum(), -t[1].v.sum())):
        x = x.sort_values(['q', 'v'], ascending=False)
        models = []
        for rank, (_, r) in enumerate(x.iterrows(), 1):              # all models of the brand
            bd = band(r.tm, cls)
            f = FERD[BAND_FER[bd]]; same = [f] if band(f['tm'], cls) == bd else []
            p = taxed(r.v / r.q, br)
            models.append(dict(rank=rank, model=r.m, tm=float(r.tm), units=int(r.q), value_idr=float(r.v * FX), band=bd, price_idr=p, photo=PHOTO_C.get(r.m),
                               ferrari=f['model'], ferrari_same_band=bool(same), diff=f['price_idr'] / p - 1))
        bb = {}
        for _, r in x.iterrows(): bb[band(r.tm, cls)] = bb.get(band(r.tm, cls), 0) + int(r.q)
        brands.append(dict(brand=br, units=int(x.q.sum()), value_idr=float(x.v.sum() * FX), n_models=len(x), models=models, band_units=bb))
    out['classes'][cls] = dict(units=int(g.q.sum()), bands=BANDS[cls], band_units=band_units, brands=brands,
                               lineup=[f['model'] for f in line], band_ferrari={b: BAND_FER[b] for b in BANDS[cls]})
json.dump(out, open('slides/v5_compare.json', 'w'), indent=1)
fmt = lambda v: f'Rp {round(v / 1e7) * 1e7:,.0f}'.replace(',', '.')
for f in fer: print(f"{f['model']:11} {f['tm']:5} {fmt(f['price_idr'])}{' (est.)' if f['estimate'] else ''}")
for cls, c in out['classes'].items():
    print('==', cls, c['units'], c['band_units'])
    for b in c['brands']:
        print(' ', b['brand'], b['units'], [(m['model'], m['band'], fmt(m['price_idr']), m['ferrari'], f"{m['diff']:+.0%}") for m in b['models']])
