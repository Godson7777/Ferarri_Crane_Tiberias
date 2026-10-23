# F.lli Ferrari Market Strategy deck

    python slides/build_data.py <Crane_Market_GVW_Pairing_v11_Noir.xlsx>   # -> slides/data.json
    cd slides && node build_deck.js                                        # -> Flli_Ferrari_Market_Strategy_v4.pptx

Assumptions (FX, import taxes, prices, qualitative text) sit at the top of `build_deck.js`.

`Flli_Ferrari_Market_Strategy.pptx` is v1 (3 benchmarks per class). v2 adds two models per class to the head-to-heads:
Medium SPK42502 + SQZ325.4 (runner-ups of Sany Palfinger and XCMG), Heavy Hyva HC501X + XCMG GSQZ460.4.
v3: slide 1 shows unit trend (2026 annualised) and buyer profile; all product prices rounded to Rp 10.000.000;
price basis clarified (competitors = import cost before distributor margin, Ferrari = selling price to customers).
v4: trend columns removed; slide 1 describes heavy-duty use by sector; slide 6 label "F.lli Ferrari crane advantage".
