const pptxgen = require('pptxgenjs');
const pres = new pptxgen(); pres.layout = 'LAYOUT_WIDE';
// FX — PLACEHOLDER, not confirmed by user. Change here and rebuild.
const USD = 16500, EUR = 19000;
const C = {bg:'0F141A',card:'131A21',line:'26303A',txt:'EEF1F4',mut:'9AA6B2',yel:'FFD21F',red:'E30613',redT:'3A0E14'};
const F = 'Arial';
const rp = v => 'Rp' + Math.round(v/1e6).toLocaleString('en-US') + 'M';
const s = pres.addSlide(); s.background = {color:C.bg};
const T = (t,o) => s.addText(t, Object.assign({isTextBox:true,fontFace:F,color:C.txt,margin:0,valign:'top'},o));
T('HEAD-TO-HEAD — FERRARI vs TOP-3 COMPETITORS BY 2025 REVENUE',{x:0.45,y:0.3,w:12.4,h:0.3,fontSize:11,color:C.yel,bold:true,charSpacing:1});
T('Ferrari costs more than every rival, so it must win on quality, material and TSP support',{x:0.45,y:0.62,w:12.4,h:0.45,fontSize:18,bold:true});

const classes = [
 {x:0.45, name:'MEDIUM CLASS', rng:'25 ≤ X < 45 tm', fer:['Ferrari','FBR350 R','~35 tm*',63000*USD],
  comp:[['Sany Palfinger','SPK36080','34.8 tm',21896*USD],['XCMG','GSQZ330.4','33.8 tm',34431*USD],['Palfinger','PK32080C','30.4 tm',51379*USD]],
  q:{Quality:['Italian build; 900 N/mm² steel; one-piece cast base','China JV build','China build','Austrian build; premium'],
     Leadtime:['Ready stock at TSP (planned)','TBC','TBC','TBC'],
     Support:['TSP nationwide network','Dealer network (TBC)','Dealer network (TBC)','Dealer network (TBC)']}},
 {x:6.88, name:'BIG CAP CLASS', rng:'X > 45 tm · min GVW 42 t', fer:['Ferrari','9601C A8','57.5 tm',160000*EUR],
  comp:[['Palfinger','PK53002 SH B','50 tm',75926*USD],['Amco Veba','V950','46 tm',88792*USD],['XCMG','GSQZ880.6','88 tm',65289*USD]],
  q:{Quality:['Italian build; 900 N/mm² steel; one-piece cast base','Austrian build; premium','Italian build (Hyva group)','China build'],
     Leadtime:['Ready stock at TSP (planned)','TBC','TBC','TBC'],
     Support:['TSP nationwide network','Dealer network (TBC)','Dealer network (TBC)','Dealer network (TBC)']}},
];
const W = 6.0;
const hdr = {fill:{color:C.line},color:C.mut,bold:true,fontSize:9,fontFace:F};
const cell = (t,o={}) => ({text:t,options:Object.assign({fill:{color:C.card},color:C.txt,fontSize:10,fontFace:F},o)});
const fcell = (t,o={}) => cell(t,Object.assign({fill:{color:C.redT},bold:true,color:'FFFFFF'},o));
for (const k of classes){
  T([{text:k.name,options:{bold:true,color:C.txt}},{text:'   '+k.rng,options:{color:C.mut}}],{x:k.x,y:1.3,w:W,h:0.28,fontSize:12});
  // price table
  const all=[k.fer,...k.comp], rel=all.map(r=>r[3]/k.fer[3]);
  const rows=[[{text:'Brand',options:hdr},{text:'Model / Type',options:hdr},{text:'Max Lifting Moment',options:hdr},{text:'Price — crane only',options:Object.assign({align:'right'},hdr)}]];
  all.forEach((r,i)=>{const f=i?cell:fcell; rows.push([f(r[0]),f(r[1]),f(r[2]),f(rp(r[3]),{align:'right'})]);});
  s.addTable(rows,{x:k.x,y:1.65,w:W,colW:[1.5,1.6,1.4,1.5],rowH:0.27,border:{type:'solid',pt:0.75,color:C.bg},margin:[0,0.07,0,0.07],valign:'middle'});
  // qualitative table
  const br=[k.fer[0],...k.comp.map(c=>c[0])];
  const q=[[{text:'',options:hdr},...br.map((b,i)=>({text:b,options:Object.assign({},hdr,i?{}:{fill:{color:C.red},color:'FFFFFF'})}))]];
  const lvl=all.map(r=>r[3]); const srt=[...lvl].sort((a,b)=>a-b);
  const pw=v=>['Lowest','Low','High','Highest'][srt.indexOf(v)];
  q.push([cell('Price',{bold:true,color:C.mut,fontSize:9}),...lvl.map((v,i)=>(i?cell:fcell)(pw(v),{fontSize:9}))]);
  for (const [lab,v] of Object.entries(k.q)) q.push([cell(lab,{bold:true,color:C.mut,fontSize:9}),...v.map((t,i)=>(i?cell:fcell)(t,{fontSize:9,bold:i==0}))]);
  s.addTable(q,{x:k.x,y:3.12,w:W,colW:[0.8,1.9,1.1,1.1,1.1],rowH:[0.27,0.27,0.52,0.27,0.4],border:{type:'solid',pt:0.75,color:C.bg},margin:[0.02,0.06,0.02,0.06],valign:'middle'});
}
// bottom band
const by=5.2;
s.addShape(pres.shapes.RECTANGLE,{x:0.45,y:by,w:6.0,h:1.55,fill:{color:C.card}});
T('FERRARI-ONLY FEATURES — NO COMPETITOR EQUIVALENT',{x:0.65,y:by+0.15,w:5.6,h:0.25,fontSize:10,bold:true,color:C.yel});
T([{text:'High-strength special steel (900 N/mm²) — no competitor publishes a comparable figure.',options:{bullet:true,breakLine:true}},
   {text:'Casted-base, one-piece construction — +50% strength vs welded (Ferrari claim); Palfinger relies on KTL surface coating instead.',options:{bullet:true}}],
  {x:0.65,y:by+0.48,w:5.6,h:0.95,fontSize:11,paraSpaceAfter:6});
s.addShape(pres.shapes.RECTANGLE,{x:6.88,y:by,w:6.0,h:1.55,fill:{color:C.redT},line:{color:C.red,width:1}});
T('WHY FERRARI DESPITE THE PRICE GAP',{x:7.08,y:by+0.15,w:5.6,h:0.25,fontSize:10,bold:true,color:C.yel});
T('Every competitor matches Ferrari on standard crane technology, but none matches its steel grade or one-piece cast base. For mining customers running 24 t+ GVW trucks, that means a longer structural life for the price. TSP stock and nationwide support also cut downtime, which matters more to these customers than the purchase price.',
  {x:7.08,y:by+0.48,w:5.6,h:1.0,fontSize:11});
T(`Price basis: crane only. Competitors = average 2025 import unit price (HS 84269100+84264900); Ferrari = TSP price. FX used: USD 1 = Rp${USD.toLocaleString('en-US')}, EUR 1 = Rp${EUR.toLocaleString('en-US')} (ASSUMED — to be confirmed). *FBR350 R lifting moment inferred from model name, pending spec sheet. Quality/Leadtime/Support = draft TSP assessment, TBC to be validated.`,
  {x:0.45,y:6.88,w:12.43,h:0.45,fontSize:8,color:C.mut});
pres.writeFile({fileName:'Slide4_HeadToHead.pptx'}).then(()=>console.log('done'));
