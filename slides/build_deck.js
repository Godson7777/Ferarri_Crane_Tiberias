// F.lli Ferrari Market Strategy — deck builder (theme: Pit Lane Noir).
// 1) python build_data.py <xlsx>   2) node build_deck.js
const pptxgen = require('pptxgenjs');
const D = require('./data.json');

// ---------- assumptions (edit here) ----------
const USD = 17803;                        // JISDOR 23 Sep 2026
const PPN = 0.11, PPH22 = 0.025;          // PPN 12% x DPP 11/12; PPh 22 with API
const BM = {china: 0, eu: 0.05};          // ACFTA 0%; MFN 5% assumed — confirm with BTKI
const taxed = (usd, o) => usd * USD * (1 + BM[o]) * (1 + PPN + PPH22);

// ---------- design tokens ----------
const C = {bg:'0F141A', card:'161D25', line:'26303A', txt:'EEF1F4', mut:'9AA6B2', dim:'5C6773',
  red:'E30613', redT:'3A0E14', med:'F2A33A', heavy:'4F9CF9', grey1:'4E5A66', grey2:'7D8894'};
const CLS = {Medium:{c:C.med, pale:'FBE3BD', label:'MEDIUM', rng:'> 25 – 45 tm'},
             Heavy:{c:C.heavy, pale:'CFE3FE', label:'HEAVY', rng:'> 45 tm'}};
const F = 'Arial';
const LOGO = {FERRARI:['logo_ferrari',3.798], 'SANY PALFINGER':['logo_sanypalfinger',4.694], PALFINGER:['logo_palfinger',4.516],
  'AMCO VEBA':['logo_amcoveba',4.469], XCMG:['logo_xcmg',4.587]};
const SRC = `Source: Indonesia import records, HS 84269100 + 84264900, Jan 2023 – ${D.last2026}; truck-mounted knuckle boom cranes only, Zoomlion excluded.`;

const rp = v => 'Rp ' + (Math.round(v / 1e6) * 1e6).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
const rpx = v => 'Rp ' + (Math.round(v / 1000) * 1000).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
const pct = v => Math.round(v * 100) + '%';
const title = s => s.toLowerCase().replace(/\b\w/g, m => m.toUpperCase());
const brandName = b => ({'SANY PALFINGER':'Sany Palfinger','AMCO VEBA':'Amco Veba','XCMG':'XCMG','HIAB':'Hiab'}[b] || title(b));
const model = (m, b) => m.replace(b + ' ', '');
const lerp = (a, b, t) => [0, 2, 4].map(i => Math.round(parseInt(a.substr(i, 2), 16) * (1 - t) + parseInt(b.substr(i, 2), 16) * t)
  .toString(16).padStart(2, '0')).join('').toUpperCase();

const pres = new pptxgen(); pres.layout = 'LAYOUT_WIDE'; pres.title = 'F.lli Ferrari Market Strategy';
let page = -1;

function T(s, t, o) { s.addText(t, Object.assign({isTextBox:true, fontFace:F, color:C.txt, margin:0, valign:'top'}, o)); }
function slash(s, x, y, h = 0.34) { // motif: three slanted bars from the Ferrari mark
  for (let i = 0; i < 3; i++) s.addShape(pres.shapes.PARALLELOGRAM, {x:x + i * 0.19, y, w:0.14, h, fill:{color:C.red}, line:{color:C.red, width:0}, adjustPoint:0.45});
}
function classTab(s, x, y, k, w = 2.6, h = 0.34, fs = 12) {
  s.addShape(pres.shapes.PARALLELOGRAM, {x, y, w, h, fill:{color:CLS[k].c}, line:{color:CLS[k].c, width:0}, adjustPoint:0.3});
  T(s, [{text:CLS[k].label, options:{bold:true}}, {text:'  ·  ' + CLS[k].rng}], {x, y, w, h, fontSize:fs, color:C.bg, align:'center', valign:'middle'});
}
function slide(kicker, head) {
  const s = pres.addSlide(); s.background = {color:C.bg}; page++;
  T(s, kicker, {x:0.45, y:0.3, w:10, h:0.28, fontSize:10.5, bold:true, color:C.mut, charSpacing:2});
  T(s, head, {x:0.45, y:0.6, w:11.7, h:0.45, fontSize:18, bold:true});
  slash(s, 12.33, 0.3);
  T(s, String(page), {x:12.4, y:7.05, w:0.48, h:0.25, fontSize:9, color:C.dim, align:'right'});
  return s;
}
function foot(s, t) { T(s, t, {x:0.45, y:6.93, w:11.8, h:0.4, fontSize:7.5, color:C.dim}); }
function logoTile(s, brand, x, y, w, h) {
  const [f, ar] = LOGO[brand];
  if (brand !== 'FERRARI') s.addShape(pres.shapes.ROUNDED_RECTANGLE, {x, y, w, h, fill:{color:'FFFFFF'}, line:{color:'FFFFFF', width:0}, rectRadius:0.05});
  const pad = brand === 'FERRARI' ? 0 : 0.06, lw = Math.min(w - 2 * pad, (h - 2 * pad) * ar), lh = lw / ar;
  s.addImage({path:`assets/${f}.png`, x:x + (w - lw) / 2, y:y + (h - lh) / 2, w:lw, h:lh});
}
function tableTitle(s, t, x, y, w) { T(s, t, {x, y, w, h:0.25, fontSize:10, bold:true, color:C.txt}); }
const H = (t, o = {}) => ({text:t, options:Object.assign({fill:{color:C.line}, color:C.mut, bold:true, fontSize:8.5, fontFace:F}, o)});
const V = (t, o = {}) => ({text:t, options:Object.assign({fill:{color:C.card}, color:C.txt, fontSize:9.5, fontFace:F}, o)});
const TB = {border:{type:'solid', pt:1, color:C.bg}, margin:[0.03, 0.08, 0.03, 0.08], valign:'middle'};

// =============== 0. COVER ===============
{
  const s = pres.addSlide(); s.background = {color:C.bg}; page++;
  s.addImage({path:'assets/cover_fbr450r.jpg', x:5.9, y:0, w:7.433, h:7.5, sizing:{type:'cover', w:7.433, h:7.5}});
  s.addShape(pres.shapes.PARALLELOGRAM, {x:4.9, y:0, w:2.2, h:7.5, fill:{color:C.bg}, line:{color:C.bg, width:0}, adjustPoint:0.45});
  s.addShape(pres.shapes.RECTANGLE, {x:0, y:0, w:5.95, h:7.5, fill:{color:C.bg}, line:{color:C.bg, width:0}});
  logoTile(s, 'FERRARI', 0.6, 0.7, 2.6, 0.685);
  slash(s, 0.6, 2.45, 0.4);
  T(s, 'F.lli Ferrari\nMarket Strategy', {x:0.6, y:3.0, w:5.6, h:1.5, fontSize:40, bold:true});
  T(s, 'Indonesia truck-mounted knuckle boom crane import market\nJan 2023 – Aug 2026', {x:0.6, y:4.6, w:5.2, h:0.7, fontSize:14, color:C.mut});
  T(s, 'PT Triatra Sinergia Pratama  ·  September 2026', {x:0.6, y:6.6, w:5.2, h:0.3, fontSize:10.5, color:C.dim});
}

// =============== 1. MARKET BY CLASS ===============
{
  const K = ['Light', 'Small', 'Medium', 'Heavy'];
  const vT = D.totValueUSD.Medium + D.totValueUSD.Heavy, vAll = K.reduce((a, k) => a + D.totValueUSD[k], 0);
  const uT = D.totUnits.Medium + D.totUnits.Heavy, uAll = K.reduce((a, k) => a + D.totUnits[k], 0);
  const s = slide('1 · WHICH CLASS SHOULD WE ENTER?',
    `Medium and Heavy: ${pct(uT / uAll)} of units but ${pct(vT / vAll)} of import value — our target classes`);
  tableTitle(s, `Units Imported per Year by Class, 2023 – ${D.last2026}`, 0.45, 1.3, 7.6);
  s.addChart(pres.charts.BAR, K.map(k => ({name:`${k} (${D.ranges[k]})`, labels:['2023', '2024', '2025', '2026 (Jan – 14 Aug)'], values:D.unitsYear[k]})),
    {x:0.35, y:1.55, w:7.8, h:2.85, barDir:'col', barGrouping:'clustered', barGapWidthPct:60,
     chartColors:[C.grey1, C.grey2, C.med, C.heavy], showValue:true, dataLabelPosition:'outEnd', dataLabelFontSize:8, dataLabelColor:C.mut, dataLabelFontFace:F,
     showLegend:true, legendPos:'t', legendFontSize:9, legendColor:C.txt, legendFontFace:F,
     catAxisLabelColor:C.mut, catAxisLabelFontSize:9, catAxisLabelFontFace:F, catAxisLineShow:false,
     valAxisHidden:true, valGridLine:{style:'none'}, catGridLine:{style:'none'}});
  tableTitle(s, `Total Units and Import Value (IDR, before tax) by Class, 2023 – ${D.last2026}`, 0.45, 4.55, 7.6);
  const rows = [[H('Class'), H('Max Lifting Moment'), H('Units', {align:'right'}), H('Import value (IDR, before tax)', {align:'right'}), H('Share of value', {align:'right'})]];
  K.forEach(k => {
    const t = k === 'Medium' || k === 'Heavy', o = t ? {bold:true, color:CLS[k].c} : {};
    rows.push([V(k, o), V(D.ranges[k], o), V(String(D.totUnits[k]), Object.assign({align:'right'}, o)),
      V(rp(D.totValueUSD[k] * USD), Object.assign({align:'right'}, o)), V(pct(D.totValueUSD[k] / vAll), Object.assign({align:'right'}, o))]);
  });
  rows.push([V('Total', {bold:true}), V(''), V(String(uAll), {bold:true, align:'right'}), V(rp(vAll * USD), {bold:true, align:'right'}), V('100%', {bold:true, align:'right'})]);
  s.addTable(rows, Object.assign({x:0.45, y:4.82, w:7.6, colW:[1.2, 1.7, 1.0, 2.5, 1.2], rowH:0.28}, TB));
  // KPI tiles
  tableTitle(s, 'Average Units Sold per Year', 8.45, 1.3, 4.4);
  [[D.avgTarget, 'Medium + Heavy', 'target classes', C.red], [D.avgOther, 'Light + Small', 'non-target classes', C.grey1]].forEach(([v, a, b, col], i) => {
    const x = 8.45 + i * 2.25;
    s.addShape(pres.shapes.RECTANGLE, {x, y:1.58, w:2.1, h:1.35, fill:{color:i ? C.card : C.redT}, line:{color:col, width:1}});
    T(s, String(Math.round(v)), {x:x + 0.15, y:1.68, w:1.8, h:0.6, fontSize:32, bold:true});
    T(s, [{text:a, options:{bold:true, breakLine:true}}, {text:b + ' · units / year'}], {x:x + 0.15, y:2.33, w:1.85, h:0.5, fontSize:9, color:C.mut});
  });
  T(s, `2026 annualised (Jan – 14 Aug × ${D.f26.toFixed(2)}); average of 4 years.`, {x:8.45, y:3.0, w:4.4, h:0.22, fontSize:7.5, color:C.dim});
  s.addShape(pres.shapes.RECTANGLE, {x:8.45, y:3.35, w:4.4, h:3.3, fill:{color:C.card}, line:{color:C.card, width:0}});
  T(s, 'WHY MEDIUM AND HEAVY', {x:8.65, y:3.5, w:4.0, h:0.25, fontSize:10, bold:true, color:C.red});
  const avgT = vT / uT * USD, avgO = (vAll - vT) / (uAll - uT) * USD;
  T(s, [
    {text:'Bigger value per deal', options:{bold:true, breakLine:true}},
    {text:`Fewer units per year, but ${pct(vT / vAll)} of the money. Average import value per unit is ${rp(avgT)} vs ${rp(avgO)} for Light + Small.`, options:{color:C.mut, breakLine:true}},
    {text:' ', options:{fontSize:6, breakLine:true}},
    {text:'Fits TSP\'s core business', options:{bold:true, breakLine:true}},
    {text:`Both classes need rear outriggers and heavy trucks: ${pct(D.gvw24Share)} of their units need a truck with GVW of 24 t or more — the mining-duty trucks TSP already serves.`, options:{color:C.mut}}],
    {x:8.65, y:3.85, w:4.0, h:2.7, fontSize:11.5});
  foot(s, SRC + ` Values converted at USD 1 = Rp ${USD.toLocaleString('id-ID')} (JISDOR, 23 Sep 2026).`);
}

// =============== 2. BRANDS 2025 (waterfall) ===============
function waterfall(s, k, x, y, w, h) {
  const b = D['brands' + k], tot = b.reduce((a, r) => a + r.q, 0), n = b.length + 1;
  const slot = w / n, bw = slot * 0.62, top = y + 0.55, base = y + h - 0.38, sc = (base - top) / tot;
  let cum = 0;
  b.concat([{Brand:'TOTAL', q:tot, v:b.reduce((a, r) => a + r.v, 0)}]).forEach((r, i) => {
    const isT = r.Brand === 'TOTAL', x0 = x + i * slot + (slot - bw) / 2;
    const y1 = isT ? base - tot * sc : base - (cum + r.q) * sc, hh = Math.max(r.q * sc, 0.02);
    const col = isT ? CLS[k].c : lerp(CLS[k].pale, CLS[k].c, i / (b.length - 1));
    s.addShape(pres.shapes.RECTANGLE, {x:x0, y:y1, w:bw, h:hh, fill:{color:isT ? C.bg : col}, line:{color:col, width:isT ? 1.5 : 0}});
    if (!isT && i < b.length - 1) s.addShape(pres.shapes.LINE, {x:x0 + bw, y:y1, w:slot - bw, h:0, line:{color:C.dim, width:0.75, dashType:'dash'}});
    T(s, [{text:`${r.q} ${r.q === 1 ? 'unit' : 'units'}`, options:{bold:true, fontSize:9.5, color:C.txt, breakLine:true}}, {text:rp(r.v * USD), options:{fontSize:7, color:C.mut}}],
      {x:x0 - (slot - bw) / 2, y:y1 - 0.42, w:slot, h:0.4, align:'center', valign:'bottom'});
    T(s, isT ? `Total ${k}` : brandName(r.Brand), {x:x0 - (slot - bw) / 2 - 0.05, y:base + 0.06, w:slot + 0.1, h:0.32, fontSize:8, align:'center', color:isT ? CLS[k].c : C.mut, bold:isT});
    if (!isT) cum += r.q;
  });
  s.addShape(pres.shapes.LINE, {x, y:base, w, h:0, line:{color:C.line, width:1}});
}
{
  const tm = D.topMedium[0], th = D.topHeavy[0];
  const s = slide('2 · WHO LEADS THE TARGET CLASSES?  (2025 ONLY)',
    `2025: ${brandName(tm.Brand)} leads Medium (${pct(tm.share)} of value), ${brandName(th.Brand)} leads Heavy (${pct(th.share)})`);
  s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:1.2, w:12.43, h:0.3, fill:{color:C.card}, line:{color:C.card, width:0}});
  T(s, 'Scope: 2025 only (one full year). Figures are therefore smaller than the 2023 – 2026 totals on slide 1.', {x:0.6, y:1.2, w:12, h:0.3, fontSize:9.5, color:C.mut, valign:'middle'});
  ['Medium', 'Heavy'].forEach((k, j) => {
    const x = 0.45 + j * 6.43, w = 6.0;
    classTab(s, x, 1.68, k);
    tableTitle(s, `${k} Class: Units and Import Value by Brand, 2025`, x, 2.12, w);
    T(s, 'Bars = units (smallest → largest). Label = units and import value (IDR, before tax).', {x, y:2.37, w, h:0.2, fontSize:7.5, color:C.dim});
    waterfall(s, k, x, 2.55, w, 2.35);
    tableTitle(s, `${k} Class: Top 3 Brands by Import Value, 2025`, x, 5.05, w);
    const rows = [[H('#'), H('Brand'), H('Units', {align:'right'}), H('Import value (IDR, before tax)', {align:'right'}), H('Share of class value', {align:'right'})]];
    D['top' + k].forEach((r, i) => rows.push([V(String(i + 1)), V(brandName(r.Brand), {bold:true}), V(String(r.q), {align:'right'}), V(rp(r.v * USD), {align:'right'}), V(pct(r.share), {align:'right', bold:true, color:CLS[k].c})]));
    s.addTable(rows, Object.assign({x, y:5.32, w, colW:[0.35, 1.6, 0.75, 2.1, 1.2], rowH:0.28}, TB));
  });
  foot(s, SRC + ` Values converted at USD 1 = Rp ${USD.toLocaleString('id-ID')}.`);
}

// =============== 3. WINNER MODELS ===============
{
  const s = slide('3 · WHICH MODELS DO THE LEADERS SELL?  (2025)',
    'Each leading brand relies on one winner model — these are Ferrari\'s benchmarks');
  const notes = {
    'Heavy|PALFINGER':'One model = 100% of Palfinger\'s Heavy units.',
    'Heavy|AMCO VEBA':'One model = 100% of Amco Veba\'s Heavy units.',
    'Heavy|XCMG':'Tie on units; GSQZ880.6 highlighted (higher value). Both are 86 – 88 tm.'};
  ['Medium', 'Heavy'].forEach((k, j) => {
    const y = 1.25 + j * 2.8;
    classTab(s, 0.45, y, k);
    Object.entries(D['models' + k]).forEach(([br, ms], i) => {
      const x = 0.45 + i * 4.21, w = 4.01, cy = y + 0.45;
      s.addShape(pres.shapes.RECTANGLE, {x, y:cy, w, h:2.2, fill:{color:C.card}, line:{color:C.card, width:0}});
      logoTile(s, br, x + 0.15, cy + 0.12, 1.3, 0.29);
      T(s, `${brandName(br)}: Units Sold by Model, ${k}, 2025`, {x:x + 1.6, y:cy + 0.12, w:w - 1.7, h:0.3, fontSize:9, bold:true, valign:'middle'});
      const labels = ms.map(m => model(m.Model, br));
      s.addChart(pres.charts.BAR, [
        {name:'Winner', labels, values:ms.map((m, t) => t ? 0 : m.q)},
        {name:'Other', labels, values:ms.map((m, t) => t ? m.q : 0)}],
        {x:x + 0.05, y:cy + 0.45, w:w - 0.15, h:ms.length === 1 ? 0.75 : 1.4, barDir:'bar', barGrouping:'stacked', barGapWidthPct:35,
         chartColors:[CLS[k].c, C.grey2], showValue:true, dataLabelPosition:'inEnd', dataLabelFormatCode:'0;;;', dataLabelFontSize:9,
         dataLabelFontBold:true, dataLabelColor:C.bg, dataLabelFontFace:F, catAxisOrientation:'maxMin', catAxisLabelColor:C.txt,
         catAxisLabelFontSize:9, catAxisLabelFontFace:F, catAxisLineShow:false, valAxisHidden:true, valAxisMinVal:0,
         valAxisMaxVal:Math.max(...ms.map(m => m.q)), valGridLine:{style:'none'}, catGridLine:{style:'none'}, showLegend:false});
      const n = notes[k + '|' + br];
      if (n) T(s, n, {x:x + 0.15, y:cy + 1.8, w:w - 0.3, h:0.3, fontSize:8, color:C.mut, italic:true});
    });
  });
  foot(s, SRC + ' Top 3 brands by 2025 import value per class (slide 2); up to 3 models shown per brand; highlighted bar = winner model (most units).');
}

// =============== 4–5. HEAD-TO-HEAD ===============
const FER = {quality:'EU design, EN 12999, high-strength steel + one-piece cast base', lead:'Fast — China factory and stock',
  support:'Good — parts stock in Malaysia, 24/7 team, fly-in mechanics'};
const CN = {pos:'Low — China mass production', quality:'China design, heavier, no EN 12999', lead:'Fast — China factory and stock', support:'Poor — no after-sales support'};
const H2H = {
  Medium:{fer:'FBR350 R', head:'Medium: FBR350 R is priced on par with Palfinger, with better build and faster supply',
    cols:[
      {b:'SANY PALFINGER', m:'SPK36080', ph:'spk36080', tm:'34.8 tm', p:taxed(21896, 'china'), ...CN},
      {b:'XCMG', m:'GSQZ330.4', ph:'gsqz330', tm:'33.8 tm', p:taxed(34431, 'china'), ...CN},
      {b:'PALFINGER', m:'PK 32080 C', ph:'pk32080c', tm:'30.4 tm', p:taxed(51379, 'eu'), pos:'High — Europe-made',
       quality:'EU design, EN 12999, high-strength steel', lead:'Europe and China factories', support:'Good — many service points in Indonesia'},
      {b:'FERRARI', m:'FBR350 R', ph:'fbr350r', tm:'~35 tm*', p:63000 * USD, pos:'On par with Palfinger (+3%)', f:1, ...FER}],
    why:'Chinese rivals are cheaper but have no EN 12999 and no after-sales support. Against Palfinger, Ferrari costs about the same yet adds special steel, a one-piece cast base, China stock and 24/7 support — what mining fleets value more than price.',
    note:'*FBR350 R lifting moment inferred from the model name, pending spec sheet.'},
  Heavy:{fer:'FBR450 R E4', head:'Heavy: FBR450 R E4 undercuts Palfinger and Amco Veba at equal European quality',
    cols:[
      {b:'PALFINGER', m:'PK 53002 SH B', ph:'pk53002', tm:'50.1 tm', p:taxed(75926, 'eu'), pos:'High — Austrian brand premium',
       quality:'EU design, EN 12999, high-strength steel', lead:'Europe and China factories', support:'Good — many service points in Indonesia'},
      {b:'AMCO VEBA', m:'V950', ph:'v950', tm:'46 tm', p:taxed(88792, 'eu'), pos:'Highest — made in Italy',
       quality:'EU (Italy) design, EN 12999, high-strength steel', lead:'Built in Italy — longer leadtime', support:'Moderate — Hyva dealer network'},
      {b:'FERRARI', m:'FBR450 R E4', ph:'fbr450r', tm:'45.5 tm', p:1600000000, pos:'Lowest of the three', f:1, ...FER}],
    why:'At Rp 1.600.000.000, Ferrari is below Palfinger and Amco Veba while meeting the same European standard (EN 12999). It adds two things neither rival offers — 900 N/mm² steel and a one-piece cast base — plus faster supply from China stock.',
    note:''}};
function h2h(k, n) {
  const K = H2H[k], nc = K.cols.length, lab = 1.55, cw = (12.43 - lab) / nc;
  const s = slide(`${n} · HEAD-TO-HEAD — ${CLS[k].label} CLASS`, K.head);
  classTab(s, 0.45, 1.2, k, 3.2, 0.4, 14);
  tableTitle(s, `${k} Class (${CLS[k].rng}): Ferrari ${K.fer} vs 2025 Winner Models`, 3.85, 1.28, 8.9);
  K.cols.forEach((c, i) => {
    const x = 0.45 + lab + i * cw, y = 1.75, hh = 1.18;
    s.addShape(pres.shapes.RECTANGLE, {x:x + 0.01, y, w:cw - 0.02, h:hh, fill:{color:c.f ? C.red : C.line}, line:{color:c.f ? C.red : C.line, width:0}});
    s.addImage({path:`assets/photo_${c.ph}.jpg`, x:x + 0.06, y:y + 0.05, w:hh - 0.1, h:hh - 0.1});
    const tx = x + hh + 0.02, tw = cw - hh - 0.1;
    logoTile(s, c.b, tx, y + 0.12, Math.min(tw, 1.55), Math.min(tw, 1.55) / (c.b === 'FERRARI' ? 3.798 : 4.2));
    T(s, c.m, {x:tx, y:y + 0.55, w:tw, h:0.3, fontSize:c.f ? 13 : 12, bold:true, color:'FFFFFF'});
    T(s, c.f ? 'Ferrari offer' : 'Winner model, 2025', {x:tx, y:y + 0.85, w:tw, h:0.22, fontSize:8, color:c.f ? 'FFD7D9' : C.mut});
  });
  const L = t => V(t, {bold:true, color:C.mut, fontSize:8.5});
  const cell = (c, t, o = {}) => V(t, Object.assign(c.f ? {fill:{color:C.redT}, bold:true, color:'FFFFFF'} : {}, o));
  const rows = [];
  const add = (t, f, o) => rows.push([L(t), ...K.cols.map(c => cell(c, f(c), o))]);
  add('Max Lifting Moment', c => c.tm);
  add('Price per unit, crane only (IDR, incl. tax)', c => rpx(c.p), {fontSize:11});
  add('Price position', c => c.pos); add('Quality', c => c.quality); add('Leadtime', c => c.lead); add('Support', c => c.support);
  s.addTable(rows, Object.assign({x:0.45, y:2.97, w:12.43, colW:[lab, ...K.cols.map(() => cw)], rowH:[0.3, 0.4, 0.3, 0.42, 0.3, 0.42]}, TB));
  let by = 5.28;
  if (k === 'Heavy') {
    T(s, [{text:'Why not XCMG?  ', options:{bold:true, color:C.heavy}}, {text:'XCMG is #3 in Heavy (slide 2), but its winners GSQZ880.6 and GSQZ860.6 are 86 – 88 tm — almost double the 45 – 50 tm of the cranes compared here. Not a like-for-like rival, so this comparison focuses on Palfinger and Amco Veba.', options:{color:C.mut}}],
      {x:0.45, y:5.2, w:12.43, h:0.4, fontSize:9});
    by = 5.62;
  }
  const bh = 6.82 - by;
  s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:by, w:6.0, h:bh, fill:{color:C.card}, line:{color:C.card, width:0}});
  T(s, 'FERRARI-ONLY FEATURES — NO COMPETITOR EQUIVALENT', {x:0.62, y:by + 0.1, w:5.7, h:0.22, fontSize:9.5, bold:true, color:C.red});
  T(s, [{text:'High-strength special steel (900 N/mm²) — no competitor publishes a comparable figure.', options:{bullet:true, breakLine:true}},
        {text:'One-piece cast base — +50% strength vs welded (Ferrari data); Palfinger uses KTL surface coating instead.', options:{bullet:true}}],
    {x:0.62, y:by + 0.38, w:5.7, h:bh - 0.45, fontSize:k === 'Heavy' ? 9.5 : 10.5, paraSpaceAfter:k === 'Heavy' ? 1 : 4});
  s.addShape(pres.shapes.RECTANGLE, {x:6.88, y:by, w:6.0, h:bh, fill:{color:C.redT}, line:{color:C.red, width:1}});
  T(s, 'WHY FERRARI', {x:7.05, y:by + 0.1, w:5.7, h:0.22, fontSize:9.5, bold:true, color:'FFFFFF'});
  T(s, K.why, {x:7.05, y:by + 0.38, w:5.7, h:bh - 0.45, fontSize:k === 'Heavy' ? 9.5 : 10.5});
  foot(s, `Competitor prices: 2025 average import unit price (Excel) × Rp ${USD.toLocaleString('id-ID')}/USD (JISDOR 23 Sep 2026), plus import duty (0% China – ACFTA; 5% Europe – MFN, assumed), PPN 11% and PPh 22 2.5%. Ferrari prices already include tax. ${K.note}`);
}
h2h('Medium', 4);
h2h('Heavy', 5);

// =============== 6. STOCK RECOMMENDATION ===============
{
  const s = slide('6 · STOCK RECOMMENDATION', 'Stock FBR350 R and FBR450 R E4 for buyers who already pay for European quality');
  const bm = D.brandsMedium, bh = D.brandsHeavy, sum = (a, f) => a.filter(f).reduce((t, r) => t + r.q, 0);
  const mTot = sum(bm, () => 1), mEU = sum(bm, r => ['PALFINGER', 'AMCO VEBA', 'HIAB'].includes(r.Brand)), mCN = sum(bm, r => ['SANY PALFINGER', 'XCMG'].includes(r.Brand));
  const hTot = sum(bh, () => 1), hEU = sum(bh, r => ['PALFINGER', 'AMCO VEBA'].includes(r.Brand));
  const pM = H2H.Medium.cols, pH = H2H.Heavy.cols;
  const cards = [
    {k:'Medium', ph:'fbr350r', m:'FBR350 R', stat:`${mEU} of ${mTot}`, statL:'Medium units in 2025 bought from European brands (Palfinger, Amco Veba, Hiab)',
     pts:[['Target buyer', `Palfinger PK 32080 C and other European-brand customers.`],
          ['Why we win', `Same price level as Palfinger (${rpx(pM[3].p)} vs ${rpx(pM[2].p)}), plus special steel, cast base, China stock and 24/7 support.`],
          ['Do not chase', `Sany Palfinger and XCMG buyers (${mCN} of ${mTot} units) on price — Ferrari cannot match their cost.`]]},
    {k:'Heavy', ph:'fbr450r', m:'FBR450 R E4', stat:`${hEU} of ${hTot}`, statL:'Heavy units in 2025 bought from Palfinger and Amco Veba',
     pts:[['Target buyer', `Palfinger PK 53002 SH B (${D.brandsHeavy.find(r => r.Brand === 'PALFINGER').q} units) and Amco Veba V950 (${D.brandsHeavy.find(r => r.Brand === 'AMCO VEBA').q} units) customers.`],
          ['Why we win', `Lower price than both (${rpx(pH[2].p)} vs ${rpx(pH[0].p)} and ${rpx(pH[1].p)}) with the same EN 12999 standard.`],
          ['Watch out', `45.5 tm is below PK 53002 SH B (50.1 tm) — lead with price, material and support, not lifting moment.`]]}];
  cards.forEach((c, i) => {
    const x = 0.45 + i * 6.43, y = 1.25, w = 6.0, h = 5.5;
    s.addShape(pres.shapes.RECTANGLE, {x, y, w, h, fill:{color:C.card}, line:{color:C.card, width:0}});
    classTab(s, x + 0.2, y + 0.2, c.k, 2.6);
    s.addImage({path:`assets/photo_${c.ph}.jpg`, x:x + 0.2, y:y + 0.7, w:1.9, h:1.9});
    T(s, 'STOCK', {x:x + 2.35, y:y + 0.75, w:3.4, h:0.25, fontSize:10, bold:true, color:C.mut, charSpacing:2});
    T(s, 'Ferrari ' + c.m, {x:x + 2.35, y:y + 1.0, w:3.5, h:0.45, fontSize:22, bold:true});
    T(s, c.stat, {x:x + 2.35, y:y + 1.6, w:3.5, h:0.5, fontSize:26, bold:true, color:CLS[c.k].c});
    T(s, c.statL, {x:x + 2.35, y:y + 2.1, w:3.5, h:0.5, fontSize:9, color:C.mut});
    c.pts.forEach(([a, b], j) => {
      const yy = y + 2.85 + j * 0.85;
      s.addShape(pres.shapes.LINE, {x:x + 0.2, y:yy, w:w - 0.4, h:0, line:{color:C.line, width:1}});
      T(s, a, {x:x + 0.2, y:yy + 0.1, w:1.5, h:0.6, fontSize:10, bold:true, color:j === 2 ? C.mut : C.txt});
      T(s, b, {x:x + 1.75, y:yy + 0.1, w:w - 1.95, h:0.7, fontSize:10});
    });
  });
  foot(s, SRC + ' Prices as on slides 4 – 5 (IDR, incl. tax).');
}

pres.writeFile({fileName:'Flli_Ferrari_Market_Strategy.pptx'}).then(() => console.log('done'));
