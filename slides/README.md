# F.lli Ferrari Market Strategy deck

    python slides/build_data.py <Crane_Market_GVW_Pairing_v11_Noir.xlsx>   # -> slides/data.json
    cd slides && node build_deck.js                                        # -> Flli_Ferrari_Market_Strategy_v2.pptx

Assumptions (FX, import taxes, prices, qualitative text) sit at the top of `build_deck.js`.

`Flli_Ferrari_Market_Strategy.pptx` is v1 (3 benchmarks per class). v2 adds two models per class to the head-to-heads:
Medium SPK42502 + SQZ325.4 (runner-ups of Sany Palfinger and XCMG), Heavy Hyva HC501X + XCMG GSQZ460.4.
