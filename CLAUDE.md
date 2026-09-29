# F.lli Ferrari crane market project (PT Triatra Sinergia Pratama)

Deck analysing Indonesia's truck-mounted knuckle boom crane imports (HS 84269100 + 84264900, Jan 2023 – 14 Aug 2026) to decide which F.lli Ferrari cranes TSP should stock and how to position them.

## Writing rules (from the user)
- Metric is "Max Lifting Moment (tm)". Never "tonnage/tonase".
- Money in Rupiah, full digits, no "M": `Rp 1.610.000.000`. Product prices rounded to Rp 10.000.000.
- Formal, simple English on slides; the user chats in Indonesian.
- Ask before assuming; explain the plan before building a PPT.
- Deck logic follows the Minto Pyramid Principle (book in repo root): answer first, S-C-Q opening, slide and exhibit titles are message sentences, groups of max ~5, MECE, no blank labels ("Why…", "Background", "Findings").
- Label for Ferrari's advantages: "F.lli Ferrari crane advantage". No executive summary slide.

## Data rules
- Market data: Excel `Crane_Market_GVW_Pairing_v11_Noir.xlsx`, sheet `8. Pivot Source`. Competitor Max Lifting Moment comes from its column `Max Lifting Moment asli (tm)`.
- Zoomlion excluded everywhere. XCMG kept.
- Classes (Excel `Kategori Max Lifting Moment`): Light ≤ 8, Small > 8–25, Medium > 25–45, Heavy > 45 tm.
- tm breakdown inside Medium and Heavy must use the same step for both classes.
- 2026 data is Jan – 14 Aug; annualise ×1.61 when comparing years. Year-trend commentary was removed at the user's request.
- Competitor price = average import unit price × JISDOR (USD 1 = Rp 17.803, 23 Sep 2026) × (1 + import duty: 0% China/ACFTA, 5% EU assumed) × (1 + 11% PPN + 2.5% PPh 22). This is import cost before distributor margin.
- Ferrari price = TSP selling price to customer, crane only, incl. GP, warranty and taxes: `data/ferrari_price_list.json` (field `price_crane_only_idr`). If a Ferrari model is not in the price list, ask the user.
- Ferrari Max Lifting Moment comes from the F.lli Ferrari pocket catalogue (pending); FBR450R ≈ 45.5 tm per the user.
- Buyer profile: 89% of Medium + Heavy cranes need a truck with GVW ≥ 24 t (heavy-duty use: mining, construction, heavy logistics). Triatra does not sell trucks.

## Build
- `slides/build_data.py <xlsx>` → `slides/data.json`; `cd slides && node build_deck.js` (pptxgenjs) → latest `Flli_Ferrari_Market_Strategy_v4.pptx`.
- Theme "Pit Lane Noir": bg 0F141A, Ferrari red E30613 (Ferrari only), Medium amber F2A33A, Heavy blue 4F9CF9, font Arial.
