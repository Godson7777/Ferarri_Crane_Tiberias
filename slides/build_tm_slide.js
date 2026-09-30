// Review slides (one per class): best-selling model per 5-tm class vs F.lli Ferrari (theme: Pit Lane Noir).
// 1) python3 analysis/build_pareto.py <xlsx>   2) node build_tm_slide.js
const pptxgen = require('pptxgenjs');
const ALL = require('../analysis/tm_grid.json');

const C = {bg:'0F141A', card:'161D25', line:'26303A', txt:'EEF1F4', mut:'9AA6B2', dim:'5C6773',
  red:'E30613', redT:'3A0E14', med:'F2A33A', heavy:'4F9CF9'};
const CLS = {Medium:{c:C.med, label:'MEDIUM', rng:'> 25 – 45 tm'}, Heavy:{c:C.heavy, label:'HEAVY', rng:'> 45 tm'}};
const F = 'Arial';
const LOGO = {'SANY PALFINGER':['logo_sanypalfinger', 4.694], PALFINGER:['logo_palfinger', 4.516],
  'AMCO VEBA':['logo_amcoveba', 4.469], XCMG:['logo_xcmg', 4.587]};
const PHOTO = {V825:'photo_v825', SPK32080:'photo_spk32080', SPK36080:'photo_spk36080', SPK42502:'photo_spk42502', V950:'photo_v950',
  'PK 53002':'photo_pk53002', '268 A4':'photo_268', 'FBR350R A4':'photo_fbr350r', '7441C':'photo_7441c', '746 A4':'photo_746r',
  'FBR450R A4':'photo_fbr450r', 'FBR660R A4':'photo_fbr660r', 'PK 76002 EH D':'photo_pk76002'};
const rpx = v => 'Rp ' + (Math.round(v / 1e7) * 1e7).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
const BRAND = {'SANY PALFINGER':'Sany Palfinger', PALFINGER:'Palfinger', 'AMCO VEBA':'Amco Veba', XCMG:'XCMG', HIAB:'Hiab'};
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
  const lead = G.reduce((a, r) => (a[r.brand] = (a[r.brand] || 0) + 1, a), {});
  const top = Object.entries(lead).sort((a, b) => b[1] - a[1])[0];
  const NUM = ['zero', 'one', 'two', 'three', 'four', 'five'];
  const head = gap
    ? `In ${k}, F.lli Ferrari has a priced crane for ${pct(priced / tot)} of units; the ${gap.band} tm class (${pct(gap.units / tot)}) ` +
      (gap.ferrari ? `has no Ferrari price yet (${gap.ferrari})` : `has no Ferrari model (${gap.model} leads)`)
    : `In ${k}, F.lli Ferrari has a crane in every tm class, but ${BRAND[top[0]]} leads ${NUM[top[1]]} of the ${NUM[G.length]} classes`;

  const s = pres.addSlide(); s.background = {color:C.bg};
  T(s, 'F.LLI FERRARI POSITIONING BY MAX LIFTING MOMENT', {x:0.45, y:0.3, w:10, h:0.28, fontSize:10.5, bold:true, color:C.mut, charSpacing:2});
  T(s, head,
    {x:0.45, y:0.6, w:11.7, h:0.65, fontSize:18, bold:true});
  slash(s, 12.33, 0.3);
  T(s, `${k} class (${CLS[k].rng}): best-selling import model in each 5-tm class vs the F.lli Ferrari offer, Jan 2023 – 14 Aug 2026`,
    {x:0.45, y:1.35, w:11.8, h:0.25, fontSize:10, bold:true});

  // ---------- table ----------
  const X0 = 0.95, Y0 = 1.68, HH = 0.45, RH = Math.min(1.05, 4.2 / G.length), FS = 11;
  const W = [0.95, 0.7, 1.5, 1.75, 0.9, 1.5, 1.5, 1.45, 1.68];
  const XC = W.reduce((a, w, i) => (a.push(a[i] + w), a), [X0]);   // column x positions
  const H = (t, o = {}) => ({text:t, options:Object.assign({fill:{color:C.line}, color:C.mut, bold:true, fontSize:9, fontFace:F, align:'center'}, o)});
  const V = (t, o = {}) => ({text:t, options:Object.assign({fill:{color:C.card}, color:C.txt, fontSize:FS, fontFace:F, align:'center'}, o)});
  const FR = {fill:{color:C.redT}}, RH_ = {fill:{color:C.red}, color:'FFFFFF'};
  const rows = [[H('tm class'), H('Total qty in class'), H('Best-selling model', {colspan:2}), H('Units · share of class'),
    H('Price per unit, crane only (IDR, incl. tax)'), H('F.lli Ferrari model', Object.assign({colspan:2}, RH_)),
    H('Price per unit, crane only (IDR, incl. tax)', RH_)]];
  G.forEach(r => {
    const band = r.merged_models
      ? {text:[{text:r.band + ' tm', options:{bold:true, breakLine:true}}, {text:`${r.merged_models} models`, options:{fontSize:8, color:C.mut}}],
         options:{fill:{color:C.card}, color:CLS[r.cls].c, fontSize:FS, fontFace:F, align:'center'}}
      : V(r.band + ' tm', {bold:true, color:CLS[r.cls].c});
    const fPrice = !r.ferrari_price_idr ? V('Price not yet known', {color:C.mut, italic:true})
      : r.ferrari_price_estimate
        ? {text:[{text:rpx(r.ferrari_price_idr), options:{bold:true, breakLine:true}}, {text:'estimate', options:{fontSize:8, italic:true, color:C.mut}}],
           options:Object.assign({color:C.txt, fontSize:FS, fontFace:F, align:'center'}, FR)}
        : V(rpx(r.ferrari_price_idr), Object.assign({bold:true}, FR));
    const fer = !r.ferrari ? [V('No Ferrari model', {bold:true, color:C.red, colspan:3})]
      : [V('', r.ferrari_price_idr ? FR : {}), V('', r.ferrari_price_idr ? FR : {}), fPrice];
    rows.push([band, V(String(r.units)), V(''), V(''), V(`${r.top_units} · ${pct(r.share)}`), V(rpx(r.price_idr)), ...fer]);
  });
  s.addTable(rows, {x:X0, y:Y0, w:XC[W.length] - X0, colW:W, rowH:[HH, ...G.map(() => RH)],
    border:{type:'solid', pt:1, color:C.bg}, margin:[0.02, 0.05, 0.02, 0.05], valign:'middle'});

  // photos, logos and model names drawn over the empty cells
  const photo = (m, x, y) => {
    const w = W[2] - 0.12, h = RH - 0.12, f = PHOTO[m];
    if (f) s.addImage({path:`assets/${f}.jpg`, x:x + 0.06, y:y + 0.06, w, h, sizing:{type:'cover', w, h}});
    else T(s, 'Photo to follow', {x:x + 0.06, y:y + 0.06, w, h, fontSize:8, italic:true, color:C.dim, align:'center', valign:'middle'});
  };
  G.forEach((r, i) => {
    const y = Y0 + HH + i * RH;
    photo(r.model, XC[2], y);
    const [f, ar] = LOGO[r.brand], tw = W[3] - 0.3, th = 0.3, ty = y + RH / 2 - 0.42;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {x:XC[3] + 0.15, y:ty, w:tw, h:th, fill:{color:'FFFFFF'}, line:{color:'FFFFFF', width:0}, rectRadius:0.04});
    const lw = Math.min(tw - 0.1, (th - 0.08) * ar), lh = lw / ar;
    s.addImage({path:`assets/${f}.png`, x:XC[3] + 0.15 + (tw - lw) / 2, y:ty + (th - lh) / 2, w:lw, h:lh});
    T(s, [{text:r.model, options:{bold:true, fontSize:FS, breakLine:true}}, {text:`${r.tm} tm`, options:{fontSize:9, color:C.mut}}],
      {x:XC[3], y:ty + th + 0.06, w:W[3], h:0.45, align:'center'});
    if (r.ferrari) {
      photo(r.ferrari, XC[6], y);
      T(s, [{text:r.ferrari, options:{bold:true, fontSize:FS, breakLine:true}}, {text:`${r.ferrari_tm} tm`, options:{fontSize:9, color:C.mut}}],
        {x:XC[7], y, w:W[7], h:RH, align:'center', valign:'middle'});
    }
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

  T(s, 'Source: Indonesia import records, HS 84269100 + 84264900, Jan 2023 – 14 Aug 2026; truck-mounted knuckle boom cranes only, Zoomlion and Hyva excluded. ' +
    'Best-selling model = most units in the class (variants merged). Competitor price = average import unit price × Rp 17.803/USD (JISDOR 23 Sep 2026) + import duty ' +
    '(0% China – ACFTA; 5% Europe, assumed), PPN 11%, PPh 22 2.5% — import cost before distributor margin. Ferrari = catalogue model closest in tm to the best seller; ' +
    'price = TSP selling price to customers, crane only, incl. GP, warranty and tax. ' +
    (G.some(r => r.ferrari_price_estimate) ? '7441C estimate = FBR350R A4 price + 4.9 tm × Rp 40.311.952 per tm (price step from FBR350R A4 to 746 A4 in the price list). ' : '') +
    (G.some(r => r.merged_models) ? 'Classes above 55 tm merged (few units each); tied best sellers ranked by import value.' : '') +
    'Ferrari photos: pocket catalogue 2026 (application photos). Prices rounded to Rp 10.000.000.',
    {x:0.45, y:6.78, w:12.4, h:0.55, fontSize:7.5, color:C.dim});
});

pres.writeFile({fileName:'Slide_tm_Positioning.pptx'}).then(() => console.log('done'));
