"""Row filter shared by slides/build_data.py and analysis/build_v5_data.py (user rule, 30 Sep 2026).

A transaction is kept only when its goods description (sheet '7. Raw Data', column 'Description') names the model
it is counted under. Rows whose description is generic ("HIAB CRANES WITH ACCESSORIES", "NEW FASSI HYDRAULIC CRANE",
"USED CRANE ... AND PARTS : PALFINGER", "HIAB 19000KG KNUCKLE CRANE") cannot prove the model or its tm and are removed.
A row whose description names a different model is also removed (the one case, HIAB 6000XG, repeats a Multicrane row
with the same date, price and description).
Used cranes ("USED" in the description) are removed too: they are not head to head with a new F.lli Ferrari crane
(user, 7 Oct 2026; the 4 Hiab 477 EP-5 rows).
"""
import re
import pandas as pd

EXCL_BRANDS = ['ZOOMLION', 'HYVA']
# model code as written in the description, where it differs from the model name (checked by hand)
ALIAS = {'AMCO VEBA V950': ['950A4S'], 'AMCO VEBA V933': ['CR9334S'], 'AMCO VEBA V806N': ['V8063S'], 'FASSI F85B.24': ['F85B024']}
N = lambda s: re.sub(r'[^A-Z0-9]', '', str(s).upper())


def code(brand, model):
    c = model.upper().replace(brand.upper() + ' ', '')
    return N(re.sub(r'\s+(EH|MH|SH)\s*[A-D]$', '', c))          # PK 41002 EH C -> PK41002


def names(brand, model, desc):
    d = N(desc)
    return any(a in d for a in [code(brand, model)] + ALIAS.get(model, []))


def load(xlsx):
    """Pivot Source rows with Zoomlion/Hyva removed and the description rule applied. Returns (kept, removed)."""
    p = pd.read_excel(xlsx, sheet_name='8. Pivot Source')
    r = pd.read_excel(xlsx, sheet_name='7. Raw Data', header=1)
    assert len(p) == len(r) and (p.Model.values == r.Model.values).all() and (p.Importer.values == r.Importer.values).all(), 'sheets 7 and 8 are not row-aligned'
    p['Description'] = r.Description.values
    p = p[~p.Brand.isin(EXCL_BRANDS)].copy()
    used = p.Description.astype(str).str.upper().str.contains(r'\bUSED\b').values
    ok = [names(b, m, d) and not u for b, m, d, u in zip(p.Brand, p.Model, p.Description, used)]
    kept, bad = p[ok], p[[not o for o in ok]].copy()
    ref = kept.drop_duplicates('Model')
    other = lambda row: [m for b, m in zip(ref.Brand, ref.Model) if b == row.Brand and m != row.Model and names(b, m, row.Description)]
    bad['Reason'] = ['Used crane' if names(r.Brand, r.Model, r.Description) else
                     f'Description names {other(r)[0]}, not {r.Model}' if other(r) else 'Generic description: no model code'
                     for _, r in bad.iterrows()]
    return kept, bad
