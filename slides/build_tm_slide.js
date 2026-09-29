// Review slides (one per class): best-selling model per 5-tm class vs F.lli Ferrari (theme: Pit Lane Noir).
// 1) python3 analysis/build_pareto.py <xlsx>   2) node build_tm_slide.js
const pptxgen = require('pptxgenjs');
const ALL = require('../analysis/tm_grid.json');

const C = {bg:'0F141A', card:'161D25', line:'26303A', txt:'EEF1F4', mut:'9AA6B2', dim:'5C6773',
  red:'E30613', redT:'3A0E14', med:'F2A33A', heavy:'4F9CF9'};
const CLS = {Medium:{c:C.med, label:'MEDIUM', rng:'> 25 – 45 tm'}, Heavy:{c:C.heavy, label:'HEAVY', rng:'> 45 tm'}};
const F = 'Arial';
const LOGO = {'SANY PALFINGER':['logo_sanypalfinger', 4.694], PALFINGER:['logo_palfinger', 4.516],
  'AMCO VEBA':['logo_amcoveba', 4.469], XCMG:['logo_xcmg', 4.587], HYVA:['logo_hyva', 2.110]};
const rpx = v => 'Rp ' + (Math.round(v / 1e7) * 1e7).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
const pct = v => Math.round(v * 100) + '%';

const pres = new pptxgen(); pres.layout = 'LAYOUT_WIDE'; pres.title = 'F.lli Ferrari tm Positioning';
function T(s, t, o) { s.addText(t, Object.assign({isTextBox:true, fontFace:F, color:C.txt, margin:0, valign:'top'}, o)); }
function slash(s, x, y, h = 0.34) {
  for (let i = 0; i < 3; i++) s.addShape(pres.shapes.PARALLELOGRAM, {x:x + i * 0.19, y, w:0.14, h, fill:{color:C.red}, line:{color:C.red, width:0}, adjustPoint:0.45});
}

['Medium', 'Heavy'].forEach(k => {
  const G = ALL.filter(r => r.cls === k);
  const tot = G.reduce((a, r) => a + r.units, 0);
  const priced = G.filter(r => r.ferrari_price_idr).reduce((a, r) => a + r.units, 0);
  const gap = G.filter(r => !r.ferrari_price_idr).sort((a, b) => b.units - a.units)[0];
  const why = gap.ferrari ? `has no Ferrari price yet (${gap.ferrari})` : `has no Ferrari model (${gap.model} leads)`;

  const s = pres.addSlide(); s.background = {color:C.bg};
  T(s, 'F.LLI FERRARI POSITIONING BY MAX LIFTING MOMENT', {x:0.45, y:0.3, w:10, h:0.28, fontSize:10.5, bold:true, color:C.mut, charSpacing:2});
  T(s, `In ${k}, F.lli Ferrari has a priced crane for ${pct(priced / tot)} of units; the ${gap.band} tm class (${pct(gap.units / tot)}) ${why}`,
    {x:0.45, y:0.6, w:11.7, h:0.65, fontSize:18, bold:true});
  slash(s, 12.33, 0.3);
  T(s, `${k} class (${CLS[k].rng}): best-selling import model in each 5-tm class vs the F.lli Ferrari offer, Jan 2023 – 14 Aug 2026`,
    {x:0.45, y:1.35, w:11.8, h:0.25, fontSize:10, bold:true});

  // ---------- table ----------
  const X0 = 0.95, Y0 = 1.68, HH = 0.5, RH = Math.min(0.8, 4.1 / G.length), FS = 11;
  const W = [0.95, 0.75, 1.45, 1.5, 0.8, 1.0, 1.55, 1.45, 0.8, 1.68];
  const H = (t, o = {}) => ({text:t, options:Object.assign({fill:{color:C.line}, color:C.mut, bold:true, fontSize:9, fontFace:F, align:'center'}, o)});
  const V = (t, o = {}) => ({text:t, options:Object.assign({fill:{color:C.card}, color:C.txt, fontSize:FS, fontFace:F, align:'center'}, o)});
  const FR = {fill:{color:C.redT}};
  const rows = [[H('tm class'), H('Units in class'), H('Brand'), H('Best-selling model'), H('Max Lifting Moment (tm)'),
    H('Units · share of class'), H('Price per unit, crane only (IDR, incl. tax)'),
    H('F.lli Ferrari model', {fill:{color:C.red}, color:'FFFFFF'}), H('Max Lifting Moment (tm)', {fill:{color:C.red}, color:'FFFFFF'}),
    H('Price per unit, crane only (IDR, incl. tax)', {fill:{color:C.red}, color:'FFFFFF'})]];
  G.forEach(r => {
    const band = V(r.band + ' tm', {bold:true, color:CLS[r.cls].c});
    const fer = !r.ferrari
      ? [V('No Ferrari model', {bold:true, color:C.red, colspan:3})]
      : [V(r.ferrari, Object.assign({bold:true}, r.ferrari_price_idr ? FR : {})), V(String(r.ferrari_tm), r.ferrari_price_idr ? FR : {}),
         r.ferrari_price_idr ? V(rpx(r.ferrari_price_idr), Object.assign({bold:true}, FR)) : V('Price not yet known', {color:C.mut, italic:true})];
    rows.push([band, V(String(r.units)), V(''), V(r.model, {bold:true}), V(String(r.tm)),
      V(`${r.top_units} · ${pct(r.share)}`), V(rpx(r.price_idr)), ...fer]);
  });
  s.addTable(rows, {x:X0, y:Y0, w:W.reduce((a, b) => a + b), colW:W, rowH:[HH, ...G.map(() => RH)],
    border:{type:'solid', pt:1, color:C.bg}, margin:[0.02, 0.05, 0.02, 0.05], valign:'middle'});

  // brand logos inside the Brand column
  const xB = X0 + W[0] + W[1];
  G.forEach((r, i) => {
    const [f, ar] = LOGO[r.brand], y = Y0 + HH + i * RH, tw = W[2] - 0.2, th = Math.min(0.36, RH - 0.14);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {x:xB + 0.1, y:y + (RH - th) / 2, w:tw, h:th, fill:{color:'FFFFFF'}, line:{color:'FFFFFF', width:0}, rectRadius:0.04});
    const lw = Math.min(tw - 0.1, (th - 0.08) * ar), lh = lw / ar;
    s.addImage({path:`assets/${f}.png`, x:xB + 0.1 + (tw - lw) / 2, y:y + (RH - th) / 2 + (th - lh) / 2, w:lw, h:lh});
  });

  // class bars on the left
  {
    const n = G.length, y = Y0 + HH;
    s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:y + 0.02, w:0.45, h:n * RH - 0.04, fill:{color:CLS[k].c}, line:{color:CLS[k].c, width:0}});
    T(s, CLS[k].label, {x:0.45 - (n * RH - 0.45) / 2, y:y + (n * RH - 0.45) / 2, w:n * RH, h:0.45, rotate:270,
      fontSize:12, bold:true, color:C.bg, align:'center', valign:'middle', charSpacing:2});
  }

  // legend
  const yL = Y0 + HH + G.length * RH + 0.14;
  [[C.redT, 'Priced Ferrari offer (TSP price list)', G.some(r => r.ferrari_price_idr)],
   [C.card, 'Ferrari model exists, price not yet known', G.some(r => r.ferrari && !r.ferrari_price_idr)],
   [C.bg, 'No Ferrari model = gap', G.some(r => !r.ferrari)]].filter(x => x[2])
    .forEach(([c, t], i) => {
      const x = 0.95 + i * 3.4;
      s.addShape(pres.shapes.RECTANGLE, {x, y:yL + 0.03, w:0.18, h:0.14, fill:{color:c}, line:{color:c === C.bg ? C.red : c, width:1}});
      T(s, t, {x:x + 0.26, y:yL, w:3.1, h:0.22, fontSize:8.5, color:C.mut});
    });

  T(s, 'Source: Indonesia import records, HS 84269100 + 84264900, Jan 2023 – 14 Aug 2026; truck-mounted knuckle boom cranes only, Zoomlion excluded. ' +
    'Best-selling model = most units in the class (variants merged). Competitor price = average import unit price × Rp 17.803/USD (JISDOR 23 Sep 2026) + import duty ' +
    '(0% China – ACFTA; 5% Europe, assumed), PPN 11%, PPh 22 2.5% — import cost before distributor margin. Ferrari = catalogue model closest in tm to the best seller; ' +
    'price = TSP selling price to customers, crane only, incl. GP, warranty and tax. Prices rounded to Rp 10.000.000.',
    {x:0.45, y:6.78, w:12.4, h:0.55, fontSize:7.5, color:C.dim});
});

pres.writeFile({fileName:'Slide_tm_Positioning.pptx'}).then(() => console.log('done'));
