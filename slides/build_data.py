# Builds data.json for the deck from the Excel source. Run: python build_data.py <xlsx>
import sys, json, pandas as pd
d = pd.read_excel(sys.argv[1], sheet_name='8. Pivot Source')
d = d[~d.Brand.isin(['ZOOMLION', 'HYVA'])]         # excluded brands (scope rule)
K = 'Kategori Max Lifting Moment'
CLS = ['Light', 'Small', 'Medium', 'Heavy']
last = d[d.Year == 2026].Date.max()
f26 = 12 / (last.month - 1 + last.day / 31)        # 2026 annualisation factor
out = {'last2026': last.strftime('%d %b %Y'), 'f26': f26,
       'ranges': {c: d[d[K] == c]['Rentang Kategori'].iloc[0] for c in CLS}}
u = d.pivot_table(index=K, columns='Year', values='Quantity', aggfunc='sum').reindex(CLS).fillna(0)
v = d.groupby(K)['Total Value (USD)'].sum().reindex(CLS)
out['unitsYear'] = {c: [int(x) for x in u.loc[c]] for c in CLS}
out['totUnits'] = {c: int(u.loc[c].sum()) for c in CLS}
out['totValueUSD'] = {c: float(v[c]) for c in CLS}
ann = lambda cs: sum(u.loc[c, 2023] + u.loc[c, 2024] + u.loc[c, 2025] + u.loc[c, 2026] * f26 for c in cs) / 4
out['avgTarget'], out['avgOther'] = ann(['Medium', 'Heavy']), ann(['Light', 'Small'])
t = d[d[K].isin(['Medium', 'Heavy'])]
out['gvw24Share'] = float(t[t['GVW Rekomendasi (ton)'] >= 24].Quantity.sum() / t.Quantity.sum())
y = d[d.Year == 2025]
for c in ['Medium', 'Heavy']:
    s = y[y[K] == c]
    b = s.groupby('Brand').agg(q=('Quantity', 'sum'), v=('Total Value (USD)', 'sum')).reset_index()
    out['brands' + c] = b.sort_values(['q', 'v']).to_dict('records')
    top = b.sort_values('v', ascending=False).head(3)
    out['top' + c] = top.assign(share=top.v / b.v.sum()).to_dict('records')
    m = {}
    for br in top.Brand:
        mm = s[s.Brand == br].groupby('Model').agg(q=('Quantity', 'sum'), v=('Total Value (USD)', 'sum'),
              tm=('Max Lifting Moment asli (tm)', 'first')).reset_index()
        m[br] = mm.sort_values(['q', 'v'], ascending=False).head(3).to_dict('records')
    out['models' + c] = m
json.dump(out, open('data.json', 'w'), indent=1, default=float)
print(json.dumps(out, indent=1, default=float)[:3500])
