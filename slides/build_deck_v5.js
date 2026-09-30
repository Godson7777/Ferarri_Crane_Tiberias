// F.lli Ferrari Market Strategy v5 — deck builder (theme: Pit Lane Noir). Flow of v4; comparison slides rebuilt.
// 1) python build_data.py <xlsx>  2) python ../analysis/build_v5_data.py <xlsx> (from repo root)  3) node build_deck_v5.js
const pptxgen = require('pptxgenjs');
const D = require('./data.json');
const V5 = require('./v5_compare.json');

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
const F = 'Calibri', FH = 'Cambria';   // body / titles (user choice, 30 Sep 2026)
const LOGO = {FERRARI:['logo_ferrari',3.798], 'SANY PALFINGER':['logo_sanypalfinger',4.694], PALFINGER:['logo_palfinger',4.516],
  'AMCO VEBA':['logo_amcoveba',4.469], XCMG:['logo_xcmg',4.587], HYVA:['logo_hyva',2.110],
  HIAB:['logo_hiab',2.894], FASSI:['logo_fassi',3.947]};
const SRC = `Source: Indonesia import records, HS 84269100 + 84264900, Jan 2023 – ${D.last2026}; truck-mounted knuckle boom cranes only, Zoomlion and Hyva excluded.`;

const rp = v => 'Rp ' + (Math.round(v / 1e6) * 1e6).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
// rpx: prices rounded to Rp 10.000.000
const rpx = v => 'Rp ' + (Math.round(v / 1e7) * 1e7).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
const pct = v => Math.round(v * 100) + '%';
const title = s => s.toLowerCase().replace(/\b\w/g, m => m.toUpperCase());
const brandName = b => ({'SANY PALFINGER':'Sany Palfinger','AMCO VEBA':'Amco Veba','XCMG':'XCMG','HIAB':'Hiab'}[b] || title(b));
const model = (m, b) => m.replace(b + ' ', '');
const lerp = (a, b, t) => [0, 2, 4].map(i => Math.round(parseInt(a.substr(i, 2), 16) * (1 - t) + parseInt(b.substr(i, 2), 16) * t)
  .toString(16).padStart(2, '0')).join('').toUpperCase();

const pres = new pptxgen(); pres.layout = 'LAYOUT_WIDE'; pres.title = 'F.lli Ferrari Market Strategy v5';
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
  T(s, head, {x:0.45, y:0.56, w:11.7, h:0.7, fontSize:21, bold:true, fontFace:FH, valign:'top'});
  slash(s, 12.33, 0.3);
  T(s, String(page), {x:12.4, y:7.05, w:0.48, h:0.25, fontSize:9, color:C.dim, align:'right'});
  return s;
}
function foot(s, t) { T(s, t, {x:0.45, y:6.93, w:11.8, h:0.4, fontSize:7.5, color:C.dim}); }
// Section divider: text left, F.lli Ferrari photo right (same geometry as the cover)
function divider(title, sub, ph, note) {
  const s = pres.addSlide(); s.background = {color:C.bg}; page++;
  s.addImage({path:`assets/${ph}`, x:6.6, y:0, w:6.733, h:7.5, sizing:{type:'cover', w:6.733, h:7.5}});
  s.addShape(pres.shapes.PARALLELOGRAM, {x:5.6, y:0, w:2.2, h:7.5, fill:{color:C.bg}, line:{color:C.bg, width:0}, adjustPoint:0.45});
  s.addShape(pres.shapes.RECTANGLE, {x:0, y:0, w:6.65, h:7.5, fill:{color:C.bg}, line:{color:C.bg, width:0}});
  slash(s, 0.6, 2.3, 0.4);
  T(s, title, {x:0.6, y:2.9, w:6.2, h:1.7, fontSize:42, bold:true, fontFace:FH, valign:'top'});
  T(s, sub, {x:0.6, y:4.65, w:5.6, h:0.9, fontSize:14, color:C.mut});
  if (note) T(s, note, {x:0.6, y:5.65, w:5.4, h:0.6, fontSize:11, color:C.dim});
  T(s, String(page), {x:12.4, y:7.05, w:0.48, h:0.25, fontSize:9, color:C.dim, align:'right'});
}
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
  T(s, 'F.lli Ferrari\nMarket Strategy', {x:0.6, y:3.0, w:5.6, h:1.5, fontSize:44, bold:true, fontFace:FH});
  T(s, 'Indonesia truck-mounted knuckle boom crane import market\nJan 2023 – Aug 2026', {x:0.6, y:4.6, w:5.2, h:0.7, fontSize:14, color:C.mut});
  T(s, 'PT Triatra Sinergia Pratama  ·  September 2026', {x:0.6, y:6.6, w:5.2, h:0.3, fontSize:10.5, color:C.dim});
}

// =============== 1. MARKET BY CLASS ===============
{
  const K = ['Light', 'Small', 'Medium', 'Heavy'];
  const vT = D.totValueUSD.Medium + D.totValueUSD.Heavy, vAll = K.reduce((a, k) => a + D.totValueUSD[k], 0);
  const uT = D.totUnits.Medium + D.totUnits.Heavy, uAll = K.reduce((a, k) => a + D.totUnits[k], 0);
  const s = slide('MARKET SIZE BY CLASS · 2023 – 2026',
    `Medium and Heavy cranes are ${pct(uT / uAll)} of units but ${pct(vT / vAll)} of import value`);
  tableTitle(s, 'Units Imported per Year by Class, 2023 – 2026 (2026 annualised)', 0.45, 1.3, 7.6);
  const ann = k => [...D.unitsYear[k].slice(0, 3), Math.round(D.unitsYear[k][3] * D.f26)];
  s.addChart(pres.charts.BAR, K.map(k => ({name:`${k} (${D.ranges[k]})`, labels:['2023', '2024', '2025', '2026 (annualised)*'], values:ann(k)})),
    {x:0.35, y:1.55, w:7.8, h:2.85, barDir:'col', barGrouping:'clustered', barGapWidthPct:60,
     chartColors:[C.grey1, C.grey2, C.med, C.heavy], showValue:true, dataLabelPosition:'outEnd', dataLabelFontSize:8, dataLabelColor:C.mut, dataLabelFontFace:F,
     showLegend:true, legendPos:'t', legendFontSize:9, legendColor:C.txt, legendFontFace:F,
     catAxisLabelColor:C.mut, catAxisLabelFontSize:9, catAxisLabelFontFace:F, catAxisLineShow:false,
     valAxisHidden:true, valGridLine:{style:'none'}, catGridLine:{style:'none'}});
  tableTitle(s, `Total Units and Import Value by Class, 2023 – ${D.last2026}`, 0.45, 4.55, 7.6);
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
  T(s, `*2026 annualised (Jan – 14 Aug × ${D.f26.toFixed(2)}); average of 4 years.`, {x:8.45, y:3.0, w:4.4, h:0.22, fontSize:7.5, color:C.dim});
  s.addShape(pres.shapes.RECTANGLE, {x:8.45, y:3.35, w:4.4, h:3.3, fill:{color:C.card}, line:{color:C.card, width:0}});
  T(s, 'MEDIUM AND HEAVY CARRY THE VALUE', {x:8.65, y:3.5, w:4.0, h:0.25, fontSize:10, bold:true, color:C.red});
  const avgT = vT / uT * USD, avgO = (vAll - vT) / (uAll - uT) * USD;
  T(s, [
    {text:'Bigger value per deal', options:{bold:true, breakLine:true}},
    {text:`Fewer units per year, but ${pct(vT / vAll)} of the money. Average import value per unit is ${rp(avgT)} vs ${rp(avgO)} for Light + Small.`, options:{color:C.mut, breakLine:true}},
    {text:' ', options:{fontSize:6, breakLine:true}},
    {text:'Built for heavy-duty sectors', options:{bold:true, breakLine:true}},
    {text:`${pct(D.gvw24Share)} of Medium and Heavy cranes need a truck with GVW of 24 t or more. They work in mining, construction and heavy logistics.`, options:{color:C.mut, breakLine:true}},
    {text:' ', options:{fontSize:6, breakLine:true}},
    {text:'2026 is a partial year', options:{bold:true, breakLine:true}},
    {text:`Jan – 14 Aug 2026 shows only ${D.unitsYear.Medium[3]} Medium and ${D.unitsYear.Heavy[3]} Heavy units, well below the 2025 pace. The cause is not yet known; import records may still be incomplete.`, options:{color:C.mut}}],
    {x:8.65, y:3.8, w:4.0, h:2.8, fontSize:10.5});
  foot(s, SRC + ` Values converted at USD 1 = Rp ${USD.toLocaleString('id-ID')} (JISDOR, 23 Sep 2026). *2026 annualised: Jan – 14 Aug units × ${D.f26.toFixed(2)}.`);
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
divider('2025 Focus', 'Who leads the Medium and Heavy classes in 2025, and with which models', 'photo_749r.jpg', '2025 is the latest full year in the data.');
{
  const tm = D.topMedium[0], th = D.topHeavy[0];
  const s = slide('BRAND LEADERS · 2025',
    `In 2025, ${brandName(tm.Brand)} leads Medium with ${pct(tm.share)} of value and ${brandName(th.Brand)} leads Heavy with ${pct(th.share)}`);
  s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:1.2, w:12.43, h:0.3, fill:{color:C.card}, line:{color:C.card, width:0}});
  T(s, 'Scope: 2025 only (one full year), so figures are smaller than the 2023 – 2026 totals. Sany Palfinger is the Sany–Palfinger joint venture; its cranes are made in China.', {x:0.6, y:1.2, w:12, h:0.3, fontSize:9.5, color:C.mut, valign:'middle'});
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
  const s = slide('MODELS OF THE BRAND LEADERS · 2025',
    `In 2025, each leader sells one main model; PK 53002 alone is ${pct(D.modelsHeavy.PALFINGER[0].q / D.brandsHeavy.reduce((a, r) => a + r.q, 0))} of Heavy units`);
  const notes = {
    'Heavy|PALFINGER':'One model = 100% of Palfinger\'s Heavy units.',
    'Heavy|AMCO VEBA':'One model = 100% of Amco Veba\'s Heavy units.',
    'Heavy|XCMG':'Tie on units; GSQZ880.6 highlighted (higher value). Both are 86 – 88 tm.',
    'Medium|SANY PALFINGER':'SPK42502 ties SPK32080 on units; ranked higher on value.',
    'Medium|XCMG':'SQZ325.4 ties KSQZ300.3 on units; ranked higher on value.'};
  ['Medium', 'Heavy'].forEach((k, j) => {
    const y = 1.25 + j * 2.8;
    classTab(s, 0.45, y, k);
    [[CLS[k].c, 'Top model'], [C.grey2, 'Other models']].forEach(([col, t], i) => {
      const lx = 3.3 + i * 2.0;
      s.addShape(pres.shapes.RECTANGLE, {x:lx, y:y + 0.11, w:0.14, h:0.14, fill:{color:col}, line:{color:col, width:0}});
      T(s, t, {x:lx + 0.2, y:y + 0.06, w:3.0, h:0.24, fontSize:8.5, color:C.mut, valign:'middle'});
    });
    Object.entries(D['models' + k]).forEach(([br, ms], i) => {
      const x = 0.45 + i * 4.21, w = 4.01, cy = y + 0.45;
      s.addShape(pres.shapes.RECTANGLE, {x, y:cy, w, h:2.2, fill:{color:C.card}, line:{color:C.card, width:0}});
      logoTile(s, br, x + 0.15, cy + 0.12, 1.3, 0.29);
      T(s, `${brandName(br)}: Units Sold by Model, ${k}, 2025`, {x:x + 1.6, y:cy + 0.12, w:w - 1.7, h:0.3, fontSize:9, bold:true, valign:'middle'});
      const labels = ms.map(m => `${model(m.Model, br)} · ${m.tm} tm`);
      s.addChart(pres.charts.BAR, [
        {name:'Top model', labels, values:ms.map((m, t) => t ? 0 : m.q)},
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
  foot(s, SRC + ' Top 3 brands by 2025 import value per class (previous slide); up to 3 models shown per brand; top model = most units, ties broken by import value.');
}

// =============== PRICE COMPARISON BY tm CLASS (2023 – 2026) ===============
// Grid per the user's Excel sketch: rows = 5-tm classes, columns = brand | F.lli Ferrari pairs (3 brands per page).
const G = {good:'2FBF71', goodT:'12301F', bad:'FF8A8F'};
const CHINA = ['SANY PALFINGER', 'XCMG'];
const FERD = Object.fromEntries(V5.ferrari.map(f => [f.model, f]));
const nm = b => brandName(b);
function brandTile(s, b, x, y, w, h) {
  if (LOGO[b]) return logoTile(s, b, x, y, w, h);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {x, y, w, h, fill:{color:'FFFFFF'}, line:{color:'FFFFFF', width:0}, rectRadius:0.05});
  T(s, nm(b).toUpperCase(), {x, y, w, h, fontSize:10, bold:true, color:'1A1A1A', align:'center', valign:'middle', charSpacing:1});
}
const photo = (s, f, x, y, w, h) => s.addImage({path:`assets/photo_${f}.jpg`, x, y, w, h, sizing:{type:'cover', w, h}});
const listNames = a => a.length === 1 ? a[0] : a.slice(0, -1).join(', ') + ' and ' + a[a.length - 1];
const signed = d => Math.round(Math.abs(d) * 100) === 0 ? '±0%' : (d <= 0 ? '−' : '+') + Math.round(Math.abs(d) * 100) + '%';
const smaller = m => FERD[m.ferrari].tm < m.tm - 2;          // F.lli Ferrari model is clearly smaller than the rival
const h2hWin = m => m.diff !== null && m.diff <= 0 && !smaller(m);
const PRICE_NOTE = 'Competitors: highest import unit price of the model, Jan 2023 – 14 Aug 2026, × Rp 17.803/USD (JISDOR 23 Sep 2026) + import duty (0% China – ACFTA; 5% Europe, assumed) + PPN 11% + PPh 22 2.5%, before distributor margin. ' +
  'F.lli Ferrari: crane price from the TSP price list Rev1, June 2026 (column K minus column J), less the 13.6% TSP margin and 3% warranty. 990R: TSP price less Rp 168.000.000 cargo deck and install, then the same deductions. ' +
  '9601CR A8: UTPE quote of 15 Sep 2026, crane only, no deduction; 50.7 tm per its load chart. 7441C is an estimate. Prices rounded to Rp 10.000.000.';

divider('F.lli Ferrari Crane Positioning', 'Every brand and model in each 5-tm class, next to the F.lli Ferrari crane we would offer',
  'photo_fbr660r.jpg', 'Jan 2023 – 14 Aug 2026 · all prices at import-cost level');

// Native PowerPoint table (editable), grid layout per the user's sketch (30 Sep 2026):
// rows = 5-tm classes, first column F.lli Ferrari, then brands from most units to fewest; one model per sub-row,
// so each model cell carries its own green/red fill. Brands with no model on the slide are left off.
function comparePage(k, bands, pg) {
  const Cc = V5.classes[k], brands = Cc.brands.filter(b => b.models.some(m => bands.includes(m.band)));
  const inBand = (b, bd) => b.models.filter(m => m.band === bd).sort((a, c) => c.units - a.units || c.value_idr - a.value_idr);
  const all = brands.flatMap(b => b.models.filter(m => bands.includes(m.band)));
  const units = all.reduce((a, m) => a + m.units, 0), win = all.filter(h2hWin), wu = win.reduce((a, m) => a + m.units, 0);
  const lo = bands[0].replace('>', '').split('–')[0], hi = bands[bands.length - 1].split('–')[1];
  const isGap = (bd, f) => f.tm <= +bd.replace('>', '').split('–')[0] || f.tm > +bd.split('–')[1];
  const gapBand = bands.find(bd => isGap(bd, FERD[Cc.band_ferrari[bd]]));
  const head = win.length ? `${k}, ${lo}–${hi} tm: F.lli Ferrari is cheaper head to head on ${listNames(win.map(m => m.model))}, ${wu} of ${units} units`
    : `${k}, ${lo}–${hi} tm: F.lli Ferrari costs more than every model it matches in size` + (gapBand ? `, and has no model above ${Math.floor(FERD[Cc.band_ferrari[gapBand]].tm / 5) * 5 + 5} tm` : '');
  const s = slide(`PRICE COMPARISON BY tm CLASS · JAN 2023 – 14 AUG 2026 · ${CLS[k].label}${pg ? ' · ' + pg : ''}`, head);
  const dense = bands.length > 4, fs = dense ? 9.5 : 12.5;
  const X = 0.45, Y = 1.45, W = 12.43, tw = dense ? 0.95 : 1.05, fw = dense ? 1.6 : 2.0, bw = (W - tw - fw) / brands.length, HH = 0.5;
  const colW = [tw, fw, ...brands.map(() => bw)];
  const nsub = bands.map(bd => Math.max(1, ...brands.map(b => inBand(b, bd).length)));
  const lines = m => 3 + (m.diff !== null && m.diff <= 0 && smaller(m) ? 1 : 0);
  const lh = fs / 72 * 1.2, pad = 0.08;
  const subH = bands.map((bd, i) => Array.from({length:nsub[i]}, (_, j) => Math.max(brands.some(b => inBand(b, bd).length) ? 0.5 : 0.3, ...brands.map(b => { const m = inBand(b, bd)[j]; return m ? lines(m) * lh + pad : 0; }))));
  let flat = subH.flat(); const avail = 6.3 - Y - HH, tot = flat.reduce((a, v) => a + v, 0);
  if (tot < avail) { const add = Math.min(0.35, (avail - tot) / flat.length); flat = flat.map(v => v + add); }
  const base = {fontFace:F, fontSize:fs, color:C.txt, valign:'middle', margin:[0.02, 0.07, 0.02, 0.07]};
  const cell = (t, o = {}) => ({text:t, options:Object.assign({}, base, {fill:{color:C.card}}, o)});
  const rows = [[cell('tm class', {fill:{color:C.line}, color:C.mut, bold:true, align:'center', fontSize:9}),
    cell('F.lli Ferrari', {fill:{color:C.red}, color:'FFFFFF', bold:true, align:'center', fontSize:11, fontFace:FH}),
    ...brands.map(b => cell(`${b.units} ${b.units === 1 ? 'unit' : 'units'}`, {fill:{color:C.line}, color:C.mut, align:'center', valign:'bottom', fontSize:dense ? 9.5 : 10.5, margin:[0, 0.04, 0.04, 0.04]}))]];
  bands.forEach((bd, i) => {
    const f = FERD[Cc.band_ferrari[bd]], gap = isGap(bd, f), n = nsub[i], span = {rowspan:n};
    const anyImport = brands.some(b => inBand(b, bd).length);
    for (let j = 0; j < n; j++) {
      const r = [];
      if (!j) {
        r.push({text:[{text:bd + ' tm', options:{bold:true, color:CLS[k].c, fontFace:FH, fontSize:fs + 1.5, breakLine:true}}, {text:`${Cc.band_units[bd]} ${Cc.band_units[bd] === 1 ? 'unit' : 'units'}`, options:{color:C.mut, fontSize:fs - 1}}],
          options:Object.assign({}, base, {fill:{color:C.card}, align:'center'}, span)});
        r.push(gap ? {text:[{text:'No F.lli Ferrari model', options:{bold:true, color:G.bad, breakLine:true}}, {text:`largest is ${f.model}, ${f.tm} tm`, options:{color:C.mut, fontSize:fs - 1}}],
                      options:Object.assign({}, base, {fill:{color:C.card}}, span)}
          : {text:[{text:f.model, options:{bold:true, fontFace:FH, fontSize:fs + 1.5, breakLine:true}}, {text:`${f.tm} tm${f.estimate ? ' · estimate' : ''}`, options:{color:'C9A9AC', fontSize:fs - 1, breakLine:true}},
                   f.price_idr === null ? {text:'Price not yet known', options:{italic:true, color:C.mut}} : {text:rpx(f.price_idr) + ' /unit', options:{bold:true}}],
             options:Object.assign({}, base, {fill:{color:C.redT}}, span)});
        if (!anyImport) { r.push(cell('No imports of this size in 2023 – 2026', {colspan:brands.length, rowspan:n, italic:true, color:C.dim, align:'center'})); rows.push(r); continue; }
      } else if (!anyImport) { rows.push(r); continue; }
      brands.forEach(b => {
        const ms = inBand(b, bd), m = ms[j];
        if (m) {
          const ok = m.diff !== null && m.diff <= 0, fill = m.diff === null ? '1B222B' : ok ? G.goodT : '24151A';
          const flag = ok && smaller(m);
          r.push({text:[{text:m.model + '  ', options:{bold:true, fontSize:fs + 0.5}},
              {text:m.diff === null ? 'no price' : signed(m.diff), options:{bold:m.diff !== null, italic:m.diff === null, fontSize:fs + 0.5, color:m.diff === null ? C.mut : ok ? G.good : G.bad, breakLine:true}},
              {text:`${m.tm} tm · ${m.units} ${m.units === 1 ? 'unit' : 'units'}`, options:{color:C.mut, breakLine:true}},
              {text:rpx(m.price_idr) + ' /unit', options:{color:C.mut, breakLine:flag}},
              ...(flag ? [{text:'smaller crane', options:{color:'F5C26B', fontSize:fs - 1}}] : [])],
            options:Object.assign({}, base, {fill:{color:fill}})});
        } else if (j === ms.length) r.push(cell('', {rowspan:n - ms.length}));
      });
      rows.push(r);
    }
  });
  s.addTable(rows, {x:X, y:Y, w:W, colW, rowH:[HH, ...flat], border:{type:'solid', pt:1.5, color:C.bg}});
  // brand logos in the header cells
  brands.forEach((b, i) => {
    const x = X + tw + fw + i * bw, lw = Math.min(1.35, bw - 0.2), lh = 0.26, y = Y + 0.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {x:x + (bw - lw) / 2, y, w:lw, h:lh, fill:{color:'FFFFFF'}, line:{color:'FFFFFF', width:0}, rectRadius:0.04});
    const [fl, ar] = LOGO[b.brand], iw = Math.min(lw - 0.14, (lh - 0.07) * ar), ih = iw / ar;
    s.addImage({path:`assets/${fl}.png`, x:x + (bw - iw) / 2, y:y + (lh - ih) / 2, w:iw, h:ih});
  });
  [[G.good, 'F.lli Ferrari is cheaper'], [G.bad, 'F.lli Ferrari costs more']].forEach(([col, t], i) => {
    s.addShape(pres.shapes.RECTANGLE, {x:0.45 + i * 2.3, y:6.62, w:0.16, h:0.13, fill:{color:col}, line:{color:col, width:0}});
    T(s, t, {x:0.68 + i * 2.3, y:6.57, w:2.0, h:0.22, fontSize:10, color:C.mut, valign:'middle'});
  });
  T(s, [{text:'Smaller crane: ', options:{bold:true, color:'F5C26B'}}, {text:'cheaper, but the F.lli Ferrari model has at least 2 tm less, so it is not a head-to-head match.', options:{color:C.mut}}],
    {x:5.1, y:6.57, w:7.8, h:0.22, fontSize:10, valign:'middle'});
  foot(s, 'Each model: F.lli Ferrari price difference; tm and units Jan 2023 – 14 Aug 2026; price per unit. Zoomlion and Hyva excluded. ' + PRICE_NOTE);
}
comparePage('Medium', ['>25–30', '>30–35'], '1 OF 2');
comparePage('Medium', ['>35–40', '>40–45'], '2 OF 2');
comparePage('Heavy', V5.classes.Heavy.bands, '');

// =============== F.LLI FERRARI CRANE ADVANTAGE BY tm CLASS ===============
const ALL = ['Medium', 'Heavy'].flatMap(k => V5.classes[k].brands.flatMap(b => b.models.map(m => ({...m, brand:b.brand, k}))));
const TOTAL_UNITS = V5.classes.Medium.units + V5.classes.Heavy.units;
{
  const feat = d => !d ? 'Specification as 7000 New Age series' :
    d.replace('includes remote control, ', 'Remote control, ').replace('(piston) pump', 'piston pump').replace(', incl oil cooler, rear stab', ', oil cooler, rear stabilisers')
     .replace(', incl oil cooler', ', oil cooler').replace(/^piston/, 'Piston');
  const rows = [];
  ['Medium', 'Heavy'].forEach(k => {
    const Cc = V5.classes[k];
    Cc.bands.forEach(bd => {
      const ms = ALL.filter(m => m.k === k && m.band === bd); if (!ms.length) return;
      const f = FERD[Cc.band_ferrari[bd]], gap = f.tm <= +bd.replace('>', '').split('–')[0] || f.tm > +bd.split('–')[1];
      const win = ms.filter(m => m.diff !== null && m.diff <= 0);
      rows.push({k, bd, units:Cc.band_units[bd], fer:gap ? 'No model (largest ' + f.model + ')' : f.model,
        price:gap ? '–' : f.price_idr === null ? 'Not yet known' : rpx(f.price_idr) + (f.estimate ? ' (est.)' : ''),
        won:win.filter(m => !smaller(m)).reduce((a, m) => a + m.units, 0), wonSmall:win.filter(m => smaller(m)).length,
        vs:win.length ? win.map(m => `${m.model} ${signed(m.diff)}${smaller(m) ? ' (smaller crane)' : ''}`).join(', ') : f.price_idr === null ? 'No price yet' : 'None',
        feat:gap ? '–' : feat(f.desc), lead:gap ? '–' : f.leadtime_weeks ? `${f.leadtime_weeks} wk${f.origin && f.origin !== 'Quote' ? ' · ' + f.origin : ''}` : '–'});
    });
  });
  const won = rows.reduce((a, r) => a + r.won, 0);
  const s = slide('F.LLI FERRARI CRANE ADVANTAGE BY tm CLASS · JAN 2023 – 14 AUG 2026',
    `Head to head, F.lli Ferrari is cheaper on ${won} of ${TOTAL_UNITS} units, all against European brands`);
  const hd = [H('Class'), H('tm class'), H('Units', {align:'right'}), H('F.lli Ferrari model'), H('F.lli Ferrari price', {align:'right'}),
    H('Units won head to head', {align:'right'}), H('Cheaper than'), H('Standard features'), H('Leadtime')];
  const tb = [hd];
  rows.forEach(r => {
    const o = {fontSize:9.5}, g = r.won ? {fill:{color:G.goodT}, color:G.good, bold:true} : {color:C.dim};
    tb.push([V(r.k, Object.assign({bold:true, color:CLS[r.k].c}, o)), V(r.bd + ' tm', Object.assign({bold:true}, o)), V(String(r.units), Object.assign({align:'right'}, o)),
      V(r.fer, Object.assign({bold:true}, o)), V(r.price, Object.assign({align:'right'}, o)),
      V(`${r.won} of ${r.units}`, Object.assign({align:'right'}, o, g)), V(r.vs, Object.assign({}, o, r.won ? {color:G.good} : r.wonSmall ? {color:'F5C26B'} : {color:C.dim})), V(r.feat, o), V(r.lead, o)]);
  });
  const rh = Math.min(0.34, 3.6 / rows.length);
  s.addTable(tb, Object.assign({x:0.45, y:1.25, w:12.43, colW:[0.7, 0.8, 0.55, 1.75, 1.45, 1.05, 2.6, 2.63, 0.9], rowH:[0.4, ...rows.map(() => rh)]}, TB));
  const by = 1.25 + 0.4 + rh * rows.length + 0.25, bh = Math.min(1.35, 6.8 - by);
  s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:by, w:12.43, h:bh, fill:{color:C.card}, line:{color:C.card, width:0}});
  T(s, 'F.LLI FERRARI CRANE ADVANTAGE IN EVERY tm CLASS', {x:0.65, y:by + 0.1, w:12, h:0.25, fontSize:10, bold:true, color:C.red});
  const adv = [['Designed to EN 12999', 'The European standard for loader cranes, the same one Palfinger, Hiab and Amco Veba follow.'],
    ['900 N/mm² special steel', 'High-strength steel: a lighter crane that leaves more payload on the truck.'],
    ['One-piece cast base', 'About 50% stronger than a welded base (F.lli Ferrari data).'],
    ['Rich standard kit', 'Piston pump and oil cooler on FBR, 700 and 9000 models; remote control on FBR450R, FBR660R, 749R, 9601CR and 990R.']];
  const cw4 = 12.03 / 4;
  adv.forEach(([a, b], i) => T(s, [{text:a, options:{bold:true, fontSize:11, breakLine:true}}, {text:b, options:{fontSize:9.5, color:C.mut}}],
    {x:0.65 + i * cw4, y:by + 0.4, w:cw4 - 0.25, h:bh - 0.45}));
  foot(s, 'Head to head = the F.lli Ferrari model is about the same size or bigger. Smaller crane = cheaper, but not counted. ' + PRICE_NOTE);
}

// =============== RECOMMENDATION ===============
{
  const like = ALL.filter(h2hWin).sort((a, b) => b.units - a.units);
  const cn = ALL.filter(m => CHINA.includes(m.brand) && m.diff !== null), cnUnits = ['Medium', 'Heavy'].reduce((a, k) => a + V5.classes[k].brands.filter(b => CHINA.includes(b.brand)).reduce((t, b) => t + b.units, 0), 0);
  const lo = Math.min(...cn.map(m => m.diff)), hi = Math.max(...cn.map(m => m.diff));
  const pk53 = ALL.find(m => m.model === 'PK 53002'), pk76 = ALL.find(m => m.model === 'PK 76002 EH D');
  const s = slide('WHAT TSP SHOULD DO',
    'Sell 268 A4, FBR350R A4 and FBR450R A4 against European cranes, lead with quality against Chinese brands, and rework the 9601CR A8 price');
  const cards = [
    {c:G.good, t:'Win on price against European cranes', b:
      like.map(m => `${m.ferrari}${FERD[m.ferrari].estimate ? ' (est.)' : ''} vs ${nm(m.brand)} ${m.model}: ${signed(m.diff)}, ${m.units} ${m.units === 1 ? 'unit' : 'units'}`)},
    {c:C.med, t:'Sell quality against Chinese brands', b:[
      `Sany Palfinger and XCMG sell for less on ${cn.filter(m => m.diff > 0).length} of ${cn.length} models (F.lli Ferrari ${signed(lo)} to ${signed(hi)}).`,
      `They hold ${cnUnits} of ${TOTAL_UNITS} Medium and Heavy units (${pct(cnUnits / TOTAL_UNITS)}).`,
      'Lead with EN 12999, special steel, the cast base and the standard kit.']},
    {c:C.heavy, t:'Heavy above 50 tm', b:[
      `9601CR A8 (50.7 tm) matches PK 53002 on size, but costs ${signed(pk53.diff)}. PK 53002 sold ${pk53.units} units, the most of any Heavy model.`,
      `990R (74 tm) costs ${signed(pk76.diff)} against PK 76002 EH D. Above 75 tm F.lli Ferrari has no model.`,
      '9661C (63.2 tm) has no price yet.']}];
  const cw = (12.43 - 0.5) / 3;
  cards.forEach((c, i) => {
    const x = 0.45 + i * (cw + 0.25), y = 1.35, h = 4.75;
    s.addShape(pres.shapes.RECTANGLE, {x, y, w:cw, h, fill:{color:C.card}, line:{color:C.line, width:1}});
    s.addShape(pres.shapes.OVAL, {x:x + 0.3, y:y + 0.3, w:0.5, h:0.5, fill:{color:c.c}, line:{color:c.c, width:0}});
    T(s, String(i + 1), {x:x + 0.3, y:y + 0.3, w:0.5, h:0.5, fontSize:16, bold:true, color:C.bg, align:'center', valign:'middle'});
    T(s, c.t, {x:x + 0.95, y:y + 0.25, w:cw - 1.15, h:0.6, fontSize:15, bold:true, fontFace:FH, valign:'middle'});
    T(s, c.b.map((t, j) => ({text:t, options:{bullet:true, breakLine:j < c.b.length - 1}})), {x:x + 0.3, y:y + 1.05, w:cw - 0.55, h:h - 1.25, fontSize:12, color:C.txt, paraSpaceAfter:8});
  });
  s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:6.25, w:12.43, h:0.45, fill:{color:C.redT}, line:{color:C.red, width:1}});
  T(s, [{text:'Before launch, confirm:  ', options:{bold:true, color:'FFFFFF'}}, {text:'the 7441C and 9661C prices, the 5% import duty on European cranes, and TSP\'s own freight and clearance cost per unit.', options:{color:C.txt}}],
    {x:0.65, y:6.25, w:12.1, h:0.45, fontSize:11, valign:'middle'});
  foot(s, 'Units = Jan 2023 – 14 Aug 2026 imports, Medium + Heavy, all models. ' + PRICE_NOTE);
}

// =============== APPENDIX: MODEL PHOTOS ===============
function photoPage(title, head, cards) {
  const s = slide(title, head);
  const cols = 7, gx = 0.13, cw = (12.43 - gx * (cols - 1)) / cols, ph = 1.05, chh = 0.42, rh = ph + chh + 0.14;
  cards.forEach((c, i) => {
    const x = 0.45 + (i % cols) * (cw + gx), y = 1.3 + Math.floor(i / cols) * rh;
    s.addShape(pres.shapes.RECTANGLE, {x, y, w:cw, h:ph + chh, fill:{color:c.f ? C.redT : C.card}, line:{color:c.f ? C.red : C.card, width:c.f ? 1 : 0}});
    if (c.photo) photo(s, c.photo, x, y, cw, ph);
    else T(s, 'Photo not yet available', {x, y, w:cw, h:ph, fontSize:8, italic:true, color:C.dim, align:'center', valign:'middle'});
    T(s, [{text:c.name, options:{bold:true, fontSize:8.5, breakLine:true}}, {text:c.sub, options:{fontSize:7.5, color:C.mut}}], {x:x + 0.08, y:y + ph + 0.02, w:cw - 0.16, h:chh - 0.04, valign:'middle'});
  });
  foot(s, 'F.lli Ferrari photos: TSP and pocket catalogue 2026. Competitor photos: supplied by TSP; for identification only.');
}
{
  const comp = k => V5.classes[k].brands.flatMap(b => b.models.slice().sort((a, c) => a.tm - c.tm).map(m => ({name:`${nm(b.brand)} ${m.model}`, sub:`${m.tm} tm · ${m.units} ${m.units === 1 ? 'unit' : 'units'}`, photo:m.photo})));
  const fer = [...new Set(Object.values(V5.classes.Medium.band_ferrari).concat(Object.values(V5.classes.Heavy.band_ferrari)))].map(n => FERD[n])
    .filter(f => f.photo).map(f => ({name:`F.lli Ferrari ${f.model}`, sub:`${f.tm} tm`, photo:f.photo, f:1}));
  photoPage('APPENDIX · MODEL PHOTOS · MEDIUM', 'Medium: the competitor models in the price comparison', comp('Medium'));
  photoPage('APPENDIX · MODEL PHOTOS · HEAVY AND F.LLI FERRARI', 'Heavy competitor models and the F.lli Ferrari models in the comparison', comp('Heavy').concat(fer));
}

pres.writeFile({fileName:'Flli_Ferrari_Market_Strategy_v5.pptx'}).then(() => console.log('done'));
