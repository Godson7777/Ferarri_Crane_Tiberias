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
const F = 'Arial';
const LOGO = {FERRARI:['logo_ferrari',3.798], 'SANY PALFINGER':['logo_sanypalfinger',4.694], PALFINGER:['logo_palfinger',4.516],
  'AMCO VEBA':['logo_amcoveba',4.469], XCMG:['logo_xcmg',4.587], HYVA:['logo_hyva',2.110]};
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
  const s = slide('1 · MARKET SIZE BY CLASS, 2023 – 2026',
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
    {text:`${pct(D.gvw24Share)} of Medium and Heavy cranes need a truck with GVW of 24 t or more — heavy-duty lifting for mining, construction and heavy logistics.`, options:{color:C.mut}}],
    {x:8.65, y:3.85, w:4.0, h:2.7, fontSize:11.5});
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
{
  const tm = D.topMedium[0], th = D.topHeavy[0];
  const s = slide('2 · BRAND LEADERS, 2025',
    `In 2025, ${brandName(tm.Brand)} leads Medium with ${pct(tm.share)} of value and ${brandName(th.Brand)} leads Heavy with ${pct(th.share)}`);
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
  const s = slide('3 · MODELS OF THE BRAND LEADERS, 2025',
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
      const labels = ms.map(m => model(m.Model, br));
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
  foot(s, SRC + ' Top 3 brands by 2025 import value per class (slide 2); up to 3 models shown per brand; top model = most units, ties broken by import value.');
}

// =============== 4–7. PRICE COMPARISON BY tm CLASS (2023 – 2026) ===============
// Grid per the user's Excel sketch: rows = 5-tm classes, columns = brand | F.lli Ferrari pairs (3 brands per page).
const G = {good:'2FBF71', goodT:'12301F'};
const FERD = Object.fromEntries(V5.ferrari.map(f => [f.model, f]));
const nm = b => brandName(b);
function brandTile(s, b, x, y, w, h) {
  if (LOGO[b]) return logoTile(s, b, x, y, w, h);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {x, y, w, h, fill:{color:'FFFFFF'}, line:{color:'FFFFFF', width:0}, rectRadius:0.05});
  T(s, nm(b).toUpperCase(), {x, y, w, h, fontSize:10, bold:true, color:'1A1A1A', align:'center', valign:'middle', charSpacing:1});
}
const photo = (s, f, x, y, w, h) => s.addImage({path:`assets/photo_${f}.jpg`, x, y, w, h, sizing:{type:'cover', w, h}});
const listNames = a => a.length === 1 ? a[0] : a.slice(0, -1).join(', ') + ' and ' + a[a.length - 1];
const PRICE_NOTE = 'Competitors: average import unit price, Jan 2023 – 14 Aug 2026, × Rp 17.803/USD (JISDOR 23 Sep 2026) + import duty (0% China – ACFTA; 5% Europe, assumed) + PPN 11% + PPh 22 2.5% = import cost before distributor margin. ' +
  'F.lli Ferrari: crane price from the TSP price list Rev1, June 2026 (column K minus column J), less the TSP margin of 13.6%, so both sides are at import-cost level. 7441C = estimate between FBR350R A4 and 746 A4. Prices rounded to Rp 10.000.000.';

let kno = 3;
function comparePage(k, brands, pg, npg) {
  const Cc = V5.classes[k], ms = brands.flatMap(b => b.models), cheaper = ms.filter(m => m.diff <= 0).length;
  const head = `${k}: F.lli Ferrari costs less than ${cheaper === ms.length ? 'all ' : cheaper + ' of '}${ms.length} models from ${listNames(brands.map(b => nm(b.brand)))}`;
  const s = slide(`${++kno} · PRICE COMPARISON BY tm CLASS — ${CLS[k].label} (${pg}/${npg})`, head);
  classTab(s, 0.45, 1.2, k, 2.6, 0.32, 11);
  T(s, 'Top 3 models of each brand by units, Jan 2023 – 14 Aug 2026, next to the closest F.lli Ferrari model', {x:3.2, y:1.2, w:9.6, h:0.32, fontSize:11, bold:true, valign:'middle'});
  const X = 0.45, W = 12.43, tw = 0.95, cw = (W - tw) / (brands.length * 2), Y = 1.65, HH = 0.5, bot = 6.42;
  const RH = (bot - Y - HH) / Cc.bands.length;
  // header row
  s.addShape(pres.shapes.RECTANGLE, {x:X, y:Y, w:tw - 0.04, h:HH - 0.04, fill:{color:C.line}, line:{color:C.line, width:0}});
  T(s, 'tm class', {x:X, y:Y, w:tw - 0.04, h:HH - 0.04, fontSize:9, bold:true, color:C.mut, align:'center', valign:'middle'});
  brands.forEach((b, i) => {
    const x = X + tw + i * 2 * cw;
    s.addShape(pres.shapes.RECTANGLE, {x, y:Y, w:cw - 0.04, h:HH - 0.04, fill:{color:C.line}, line:{color:C.line, width:0}});
    brandTile(s, b.brand, x + 0.12, Y + 0.05, cw - 0.28, 0.24);
    T(s, `${b.units} units · ${b.n_models} models`, {x, y:Y + 0.3, w:cw - 0.04, h:0.15, fontSize:7.5, color:C.mut, align:'center'});
    s.addShape(pres.shapes.RECTANGLE, {x:x + cw, y:Y, w:cw - 0.04, h:HH - 0.04, fill:{color:C.red}, line:{color:C.red, width:0}});
    T(s, 'F.lli Ferrari', {x:x + cw, y:Y, w:cw - 0.04, h:HH - 0.04, fontSize:10.5, bold:true, color:'FFFFFF', align:'center', valign:'middle'});
  });
  // body
  Cc.bands.forEach((bd, r) => {
    const y = Y + HH + r * RH;
    s.addShape(pres.shapes.RECTANGLE, {x:X, y, w:tw - 0.04, h:RH - 0.04, fill:{color:C.card}, line:{color:C.card, width:0}});
    T(s, [{text:bd + ' tm', options:{bold:true, fontSize:12, color:CLS[k].c, breakLine:true}}, {text:`${Cc.band_units[bd]} units`, options:{fontSize:8.5, color:C.mut}}],
      {x:X, y, w:tw - 0.04, h:RH - 0.04, align:'center', valign:'middle'});
    brands.forEach((b, i) => {
      const x = X + tw + i * 2 * cw, inb = b.models.filter(m => m.band === bd);
      s.addShape(pres.shapes.RECTANGLE, {x, y, w:cw - 0.04, h:RH - 0.04, fill:{color:C.card}, line:{color:C.card, width:0}});
      s.addShape(pres.shapes.RECTANGLE, {x:x + cw, y, w:cw - 0.04, h:RH - 0.04, fill:{color:C.card}, line:{color:C.card, width:0}});
      const eh = (RH - 0.04) / Math.max(inb.length, 1);
      inb.forEach((m, j) => {
        const ey = y + j * eh, cmp = eh < 0.8, ph = Math.min(eh - 0.12, 0.95), pw = Math.min(ph * 1.33, cw * (cmp ? 0.3 : 0.38));
        const tx = x + pw + 0.13, txw = cw - pw - 0.18, fs = cmp ? 7 : 8.5;
        // competitor entry
        photo(s, m.photo, x + 0.06, ey + (eh - ph) / 2, pw, ph);
        T(s, [{text:`#${m.rank} `, options:{bold:true, color:CLS[k].c, fontSize:fs}}, {text:m.model, options:{bold:true, fontSize:fs + 1, breakLine:true}},
              {text:`${m.tm} tm · ${m.units} ${m.units === 1 ? 'unit' : 'units'}`, options:{fontSize:fs - 0.5, color:C.mut, breakLine:true}},
              {text:rpx(m.price_idr), options:{fontSize:fs}}],
          {x:tx, y:ey + 0.03, w:txw, h:eh - 0.06, valign:'middle'});
        // Ferrari entry, same slot
        const f = FERD[m.ferrari], ok = m.diff <= 0, fx = x + cw, pct2 = Math.round(Math.abs(m.diff) * 100);
        s.addShape(pres.shapes.RECTANGLE, {x:fx + 0.03, y:ey + 0.03, w:cw - 0.1, h:eh - 0.1, fill:{color:ok ? G.goodT : C.redT}, line:{color:ok ? G.good : C.redT, width:1}});
        photo(s, f.photo, fx + 0.08, ey + (eh - ph) / 2, pw, ph);
        const tmLine = `${f.tm} tm${m.ferrari_same_band ? '' : ' · nearest'}`;
        T(s, [{text:m.ferrari, options:{bold:true, fontSize:fs + 1, color:'FFFFFF', breakLine:true}},
              ...(cmp ? [] : [{text:tmLine, options:{fontSize:fs - 0.5, color:C.mut, breakLine:true}}]),
              {text:rpx(f.price_idr) + (f.estimate ? ' (est.)' : ''), options:{fontSize:fs, color:'FFFFFF', breakLine:true}},
              {text:ok ? `${pct2}% cheaper` : `${pct2}% higher`, options:{fontSize:fs, bold:true, color:ok ? G.good : 'FF8A8F'}}],
          {x:fx + pw + 0.14, y:ey + 0.03, w:txw - 0.02, h:eh - 0.06, valign:'middle'});
      });
    });
  });
  // legend + footnote
  [[G.goodT, G.good, 'F.lli Ferrari is cheaper'], [C.redT, C.redT, 'F.lli Ferrari is higher']].forEach(([f, l, t], i) => {
    s.addShape(pres.shapes.RECTANGLE, {x:0.45 + i * 2.6, y:6.5, w:0.2, h:0.14, fill:{color:f}, line:{color:l, width:1}});
    T(s, t, {x:0.72 + i * 2.6, y:6.46, w:2.3, h:0.22, fontSize:8.5, color:C.mut, valign:'middle'});
  });
  T(s, '"nearest" = no F.lli Ferrari model in this tm class; the closest model is shown.', {x:5.75, y:6.46, w:7.1, h:0.22, fontSize:8.5, color:C.mut, valign:'middle'});
  foot(s, SRC + ' ' + PRICE_NOTE);
}
['Medium', 'Heavy'].forEach(k => {
  const b = V5.classes[k].brands;
  comparePage(k, b.slice(0, 3), 1, 2);
  comparePage(k, b.slice(3, 6), 2, 2);
});

// =============== 8. F.LLI FERRARI CRANE ADVANTAGE BY tm CLASS ===============
{
  const rows = [], feat = d => !d ? 'Estimate; spec as 7000 New Age series' :
    d.replace('includes remote control, ', 'Remote control, ').replace('(piston) pump', 'piston pump').replace(', incl oil cooler, rear stab', ', oil cooler, rear stabilisers').replace(/^piston/, 'Piston');
  let pairs = 0, cheap = 0; const cheapBrands = new Set();
  ['Medium', 'Heavy'].forEach(k => {
    const Cc = V5.classes[k], all = Cc.brands.flatMap(b => b.models.map(m => ({...m, brand:b.brand})));
    Cc.bands.forEach(bd => {
      const ms = all.filter(m => m.band === bd); if (!ms.length) return;
      const top = ms.slice().sort((a, b) => b.units - a.units)[0];
      const fs = [...new Set(ms.map(m => m.ferrari))].map(n => FERD[n]);
      const same = ms.some(m => m.ferrari_same_band), ch = ms.filter(m => m.diff <= 0).length;
      pairs += ms.length; cheap += ch; ms.filter(m => m.diff <= 0).forEach(m => cheapBrands.add(m.brand));
      rows.push({k, bd, units:Cc.band_units[bd], top:`${nm(top.brand)} ${top.model}`, fer:fs.map(f => f.model).join(' / ') + (same ? '' : ' (nearest)'),
        price:fs.map(f => rpx(f.price_idr) + (f.estimate ? ' (est.)' : '')).join(' / '), ch:`${ch} of ${ms.length}`, all:ch > 0,
        feat:feat(fs[0].desc), lead:[...new Set(fs.map(f => f.leadtime_weeks ? `${f.leadtime_weeks} wk · ${f.origin}` : '–'))].join(' / ')});
    });
  });
  const s = slide(`${++kno} · F.LLI FERRARI CRANE ADVANTAGE BY tm CLASS`,
    [...cheapBrands].every(b => !['SANY PALFINGER', 'XCMG'].includes(b)) ? `F.lli Ferrari costs less than ${cheap} of ${pairs} top models, all of them European; against Chinese brands it competes on EU build quality` : `F.lli Ferrari costs less than ${cheap} of ${pairs} top models and adds EU build quality in every tm class`);
  const hd = [H('Class'), H('tm class'), H('Units', {align:'right'}), H('Top model in class'), H('F.lli Ferrari model'), H('F.lli Ferrari price (import-cost level)', {align:'right'}),
    H('Cheaper than top models'), H('Standard features'), H('Leadtime · made in')];
  const tb = [hd];
  rows.forEach(r => {
    const o = {fontSize:8.5}, g = r.all ? {fill:{color:G.goodT}, color:G.good, bold:true} : {};
    tb.push([V(r.k, Object.assign({bold:true, color:CLS[r.k].c}, o)), V(r.bd + ' tm', Object.assign({bold:true}, o)), V(String(r.units), Object.assign({align:'right'}, o)),
      V(r.top, o), V(r.fer, Object.assign({bold:true}, o)), V(r.price, Object.assign({align:'right'}, o)), V(r.ch, Object.assign({}, o, g)), V(r.feat, o), V(r.lead, o)]);
  });
  s.addTable(tb, Object.assign({x:0.45, y:1.25, w:12.43, colW:[0.75, 0.8, 0.55, 1.8, 1.7, 1.75, 1.05, 2.53, 1.5], rowH:0.34}, TB));
  const by = 1.25 + 0.34 * tb.length + 0.35, bh = 1.3;
  s.addShape(pres.shapes.RECTANGLE, {x:0.45, y:by, w:12.43, h:bh, fill:{color:C.card}, line:{color:C.card, width:0}});
  T(s, 'F.LLI FERRARI CRANE ADVANTAGE IN EVERY tm CLASS', {x:0.65, y:by + 0.1, w:12, h:0.25, fontSize:10, bold:true, color:C.red});
  const adv = [['EU design to EN 12999', 'Same safety standard as Palfinger, Hiab and Amco Veba; Chinese rivals do not state it.'],
    ['900 N/mm² special steel', 'High-strength steel for a lighter crane with more payload.'],
    ['One-piece cast base', 'About 50% stronger than a welded base (Ferrari data).'],
    ['Support in the region', 'Parts stock in Malaysia, 24/7 team and fly-in mechanics.']];
  const cw4 = 12.03 / 4;
  adv.forEach(([a, b], i) => T(s, [{text:a, options:{bold:true, fontSize:10.5, breakLine:true}}, {text:b, options:{fontSize:9, color:C.mut}}],
    {x:0.65 + i * cw4, y:by + 0.42, w:cw4 - 0.2, h:bh - 0.5}));
  foot(s, 'Green = F.lli Ferrari costs less than at least one top model in the class (pages 4 – 7). ' + PRICE_NOTE);
}

pres.writeFile({fileName:'Flli_Ferrari_Market_Strategy_v5.pptx'}).then(() => console.log('done'));
