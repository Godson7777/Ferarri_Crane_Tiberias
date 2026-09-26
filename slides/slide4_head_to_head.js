// Slides 4-5 — Head-to-head (Medium, Big Cap). Run: node slide4_head_to_head.js
const pptxgen = require('pptxgenjs');
const pres = new pptxgen(); pres.layout = 'LAYOUT_WIDE';
const USD = 17803;                        // JISDOR 23 Sep 2026
// Import taxes on competitor prices (Excel = import unit price). Ferrari prices already tax-inclusive.
const PPN = 0.11, PPH22 = 0.025;          // PPN 12% x DPP 11/12 = 11% effective; PPh 22 with API
const BM = {china: 0, eu: 0.05};          // ACFTA 0%; MFN assumed 5% — to confirm with BTKI
const taxed = (usd, o) => usd * USD * (1 + BM[o]) * (1 + PPN + PPH22);
const rp = v => 'Rp ' + (Math.round(v / 1000) * 1000).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
const C = {bg:'0F141A',card:'131A21',line:'26303A',txt:'EEF1F4',mut:'9AA6B2',yel:'FFD21F',red:'E30613',redT:'3A0E14'};
const F = 'Arial';
const FER = {
  quality:'Europe design, EN 12999, high-strength special steel, less material, higher working load',
  lead:'Besides Europe/Brazil factory, also has China factory and stock — fast leadtime',
  support:'Good — spare parts stock in Malaysia, 24/7 support team, Ferrari mechanics can fly in when needed'};
const CN = {quality:'China design, heavier, no EN 12999, standard material', lead:'Fast, stock and factory in China',
  support:'Poor, sold without after-sales support', price:'Cheaper, mass production, China factory'};
const slides = [
 {cls:'MEDIUM CLASS', rng:'25 ≤ X < 45 tm',
  title:'Medium: Ferrari FBR350 R sits in Palfinger\'s price range, with Europe-grade build and faster leadtime',
  cols:[
   {n:'Sany Palfinger SPK36080', tm:'34.8 tm', p:taxed(21896,'china'), ...CN, price:CN.price},
   {n:'XCMG GSQZ330.4', tm:'33.8 tm', p:taxed(34431,'china'), ...CN},
   {n:'Palfinger PK32080C', tm:'30.4 tm', p:taxed(51379,'eu'), price:'Expensive, Europe (Italy) manufacturing only',
    quality:'Europe design, EN 12999, high-strength special steel, less material, higher working load',
    lead:'Besides Italy factory, also has China factory', support:'Good, many support points in Indonesia'},
   {n:'Ferrari FBR350 R', tm:'~35 tm*', p:63000*USD, price:'Medium-priced positioning for this class', f:1, ...FER}]},
 {cls:'BIG CAP CLASS', rng:'X > 45 tm · min GVW 42 t',
  title:'Big Cap: Ferrari FBR450 R E4 costs less than Palfinger and Amco Veba, with equal European quality',
  cols:[
   {n:'Palfinger PK53002 SH B', tm:'50.1 tm', p:taxed(75926,'eu'), price:'Expensive, Europe (Austria) brand premium',
    quality:'Europe design, EN 12999, high-strength steel, Power Link Plus',
    lead:'Europe factory, also has China factory', support:'Good, many support points in Indonesia'},
   {n:'Amco Veba V950', tm:'46 tm', p:taxed(88792,'eu'), price:'Most expensive, Italy manufacturing (Hyva group)',
    quality:'Europe (Italy) design, EN 12999, high-strength steel',
    lead:'Built in Italy (Poviglio) — longer leadtime', support:'Moderate, via Hyva group dealer network'},
   {n:'XCMG GSQZ880.6', tm:'88 tm', p:taxed(65289,'china'), ...CN},
   {n:'Ferrari FBR450 R E4', tm:'~45 tm*', p:1600000000, price:'Mid-range — below Palfinger and Amco Veba', f:1, ...FER}]},
];
for (const k of slides) {
  const s = pres.addSlide(); s.background = {color:C.bg};
  const T = (t,o) => s.addText(t, Object.assign({isTextBox:true,fontFace:F,color:C.txt,margin:0,valign:'top'},o));
  T(`HEAD-TO-HEAD — ${k.cls}  (${k.rng})`,{x:0.45,y:0.3,w:12.4,h:0.3,fontSize:11,color:C.yel,bold:true,charSpacing:1});
  T(k.title,{x:0.45,y:0.62,w:12.4,h:0.4,fontSize:18,bold:true});
  const H = (t,f) => ({text:t,options:{fill:{color:f?C.red:C.line},color:f?'FFFFFF':C.txt,bold:true,fontSize:10.5,fontFace:F}});
  const L = t => ({text:t,options:{fill:{color:C.card},color:C.mut,bold:true,fontSize:10,fontFace:F}});
  const V = (t,f,o={}) => ({text:t,options:Object.assign({fill:{color:f?C.redT:C.card},color:f?'FFFFFF':C.txt,bold:!!f,fontSize:10,fontFace:F},o)});
  const rows = [[{text:'',options:{fill:{color:C.line}}}, ...k.cols.map(c=>H(c.n,c.f))]];
  const add = (lab,key,o) => rows.push([L(lab), ...k.cols.map(c=>V(typeof key=='function'?key(c):c[key],c.f,o))]);
  add('Max Lifting Moment','tm');
  add('Price (crane only, IDR, incl. tax)', c=>rp(c.p), {fontSize:11.5});
  add('Price (qualitative)','price'); add('Quality','quality'); add('Leadtime','lead'); add('Support','support');
  s.addTable(rows,{x:0.45,y:1.2,w:12.43,colW:[1.63,2.7,2.7,2.7,2.7],rowH:[0.32,0.3,0.34,0.45,0.62,0.5,0.62],
    border:{type:'solid',pt:0.75,color:C.bg},margin:[0.03,0.08,0.03,0.08],valign:'middle'});
  const by = 4.8;
  s.addShape(pres.shapes.RECTANGLE,{x:0.45,y:by,w:6.0,h:1.75,fill:{color:C.card}});
  T('FERRARI-ONLY FEATURES — NO COMPETITOR EQUIVALENT',{x:0.65,y:by+0.15,w:5.6,h:0.25,fontSize:10,bold:true,color:C.yel});
  T([{text:'High-strength special steel (900 N/mm²) — no competitor publishes a comparable figure.',options:{bullet:true,breakLine:true}},
     {text:'Casted-base, one-piece construction — +50% strength vs welded (Ferrari claim); Palfinger relies on KTL surface coating instead.',options:{bullet:true}}],
    {x:0.65,y:by+0.5,w:5.6,h:1.1,fontSize:11,paraSpaceAfter:6});
  s.addShape(pres.shapes.RECTANGLE,{x:6.88,y:by,w:6.0,h:1.75,fill:{color:C.redT},line:{color:C.red,width:1}});
  T('WHY FERRARI',{x:7.08,y:by+0.15,w:5.6,h:0.25,fontSize:10,bold:true,color:C.yel});
  T(k.cls.startsWith('MEDIUM')
    ? 'Chinese rivals are cheaper but carry no EN 12999 and no after-sales support. Against Palfinger, Ferrari is priced in the same range but adds special steel, a one-piece cast base, China stock for fast delivery and 24/7 support, which matters more to mining fleets than the price difference.'
    : 'At Rp 1,6 billion, Ferrari costs less than Palfinger and Amco Veba with the same European design and EN 12999 standard. XCMG is cheaper but has no EN 12999 and weak support. Ferrari gives the best quality for the price in the Big Cap class.',
    {x:7.08,y:by+0.5,w:5.6,h:1.15,fontSize:11});
  T(`Competitor prices = 2025 average import unit price (Excel, HS 84269100+84264900) converted at USD 1 = Rp ${USD.toLocaleString('id-ID')} (JISDOR, 23 Sep 2026), then taxed: import duty 0% China (ACFTA) / 5% Europe (MFN, assumed), PPN 11%, PPh 22 2.5%. Ferrari prices already include tax. *Ferrari lifting moment inferred from model name, pending spec sheet.`,
    {x:0.45,y:6.7,w:12.43,h:0.5,fontSize:8,color:C.mut});
}
pres.writeFile({fileName:'Slide4-5_HeadToHead.pptx'}).then(()=>console.log('done'));
