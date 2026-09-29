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
- Zoomlion and Hyva excluded everywhere (Hyva: user rule 29 Sep 2026; replace with the next brand/model). XCMG kept.
- Classes (Excel `Kategori Max Lifting Moment`): Light ≤ 8, Small > 8–25, Medium > 25–45, Heavy > 45 tm.
- tm breakdown inside Medium and Heavy must use the same step for both classes (proposed: 5-tm bands >25–30, >30–35 … on `Max Lifting Moment asli (tm)`).
- 2026 data is Jan – 14 Aug; annualise ×1.61 when comparing years. Year-trend commentary was removed at the user's request.
- Competitor price = average import unit price × JISDOR (USD 1 = Rp 17.803, 23 Sep 2026) × (1 + import duty: 0% China/ACFTA, 5% EU assumed) × (1 + 11% PPN + 2.5% PPh 22). This is import cost before distributor margin.
- Ferrari price = TSP selling price to customer, crane only, incl. GP, warranty and taxes: `data/ferrari_price_list.json` (field `price_crane_only_idr`). If a Ferrari model is not in the price list, write "price not yet known" (user instruction). Medium head-to-head offer = FBR350R A4 (Rp 1.436.737.906).
- Ferrari Max Lifting Moment comes from `data/ferrari_catalogue.json` (84 models, pocket catalogue 2026). FBR 350 = 32.8 tm (not 35), FBR 450 = 45.5, 746R = 43.4, 749R = 44.7, FBR 600 = 57.4, FBR 660 = 58.8. 934/934R is in the price list but not the catalogue.
- 7441C (37.7 tm) is not in the price list: estimate Rp 1.630.000.000 = FBR350R A4 + 4.9 tm × Rp 40.311.952/tm (price step FBR350R A4 → 746 A4). Always label "estimate".
- tm positioning slides (`slides/build_tm_slide.js`, one slide per class): 1 best-selling model per 5-tm class, period Jan 2023 – 14 Aug 2026; Heavy classes above 55 tm merged into one ">55 tm" row on the slide (detail stays in Excel).
- Ferrari price for comparisons (user, 30 Sep 2026): source = `data/ferrari_price_list_rev1.json` (TSP "Price List Crane Ferarri as of June'26 (Rev1)", PDF in `data/`). Crane price = K − J = (Cargo Deck + Mounting + Aksesoris + Crane) − (Cargo Deck + Mounting + Aksesoris + Install); then less TSP margin 13.6% (× 0.864) so both sides are at import-cost level. Check: FBR450R A4 = 1.780.960.730 − 167.884.657 = 1.613.076.073 → Rp 1.390.000.000. 7441C: crane price interpolated between FBR350R A4 and 746 A4 by tm (label estimate). The older price list PDF (`price_crane_only_idr`) is superseded. Competitor distributor margins are unknown.
- Excluded rows: PALFINGER PK 33002 and PK 36050 (USD 2.000–2.150 per unit, not complete cranes). Hiab X-HIPRO B58 price (Rp 3.370.000.000) confirmed correct.
- Next deck (v5): follow the deck v4 flow (not the Minto storyline). Focus = F.lli Ferrari advantage in every 5-tm class of Medium and Heavy (price, quality, leadtime, features); highlight cells where Ferrari is cheaper. Top 3 models per brand per class, option B layout (brand rows #1–#3, closest Ferrari with its own photo), up to 2 slides per class.
- Deck v5 structure (user-approved 30 Sep 2026): cover → market size 2023–26 (with 2026 partial-year note) → divider "2025 Focus" → brand leaders 2025 → models 2025 (tm per model) → divider "F.lli Ferrari Crane Positioning" → 4 comparison pages (period in kicker, "+ n units in other models", green = Ferrari cheaper / red = higher, "smaller crane" flag when Ferrari tm < rival tm − 2, "No F.lli Ferrari model" for >50–55) → F.lli Ferrari crane advantage (like-for-like units won) → What TSP should do. Titles must be honest (e.g. "costs more than 8 of 9").
- Buyer profile: 89% of Medium + Heavy cranes need a truck with GVW ≥ 24 t (heavy-duty use: mining, construction, heavy logistics). Triatra does not sell trucks.

## Build
- `slides/build_data.py <xlsx>` → `slides/data.json`; `analysis/build_v5_data.py <xlsx>` (run from repo root) → `slides/v5_compare.json`; `cd slides && node build_deck_v5.js` → `Flli_Ferrari_Market_Strategy_v5.pptx`, the only deck kept in the repo (user wants one PPT; v4 builder `build_deck.js` kept for reference).
- Theme "Pit Lane Noir": bg 0F141A, Ferrari red E30613 (Ferrari only), Medium amber F2A33A, Heavy blue 4F9CF9, font Arial.
