const text=value=>String(value??"").trim();
const overall=record=>{const values=Object.values(record.scores||{}).filter(Number.isFinite);return values.length?values.reduce((sum,value)=>sum+value,0)/values.length:-Infinity;};
export const pageGroups=(items,size)=>{const groups=[];for(let start=0;start<items.length;start+=size)groups.push(items.slice(start,start+size));return groups.length?groups:[[]];};
export const participantChartPageSize=records=>8;
export const participantComparisonPageSize=6;
export const sectionIntroPages=(title,role)=>[
  {title,role,variant:"gradient"},
  {title,role,variant:"photo"},
  {title,role,variant:"wash"}
];
export const behaviorPageGroups=rows=>{const pages=[];let page=[],used=0;for(const row of rows){const units=Math.max(1,Math.ceil(text(row.behavior).length/92));if(page.length&&used+units>8){pages.push(page);page=[];used=0;}page.push(row);used+=units;}if(page.length)pages.push(page);return pages.length?pages:[[]];};
export function behaviorInsights(rows){const grouped=new Map();rows.forEach((row,index)=>{if(!grouped.has(row.competency))grouped.set(row.competency,[]);grouped.get(row.competency).push({...row,sourceIndex:index});});return[...grouped.entries()].map(([competency,items])=>{const strengths=items.slice().sort((a,b)=>b.pct2-a.pct2||b.mean-a.mean||a.pct0-b.pct0||a.sourceIndex-b.sourceIndex).slice(0,3),selected=new Set(strengths.map(item=>item.sourceIndex)),development=items.filter(item=>!selected.has(item.sourceIndex)).sort((a,b)=>b.pct0-a.pct0||a.mean-b.mean||a.pct2-b.pct2||a.sourceIndex-b.sourceIndex).slice(0,3);return{competency,strengths,development};});}
const splitBullet=value=>{const words=text(value).split(/\s+/u),parts=[],limit=142;let part="";for(const word of words){if(`${part} ${word}`.trim().length>limit&&part){parts.push(part);part=word;}else part=`${part} ${word}`.trim();}if(part)parts.push(part);return parts;};
const typicalTeamSize=sizes=>{const counts=new Map();sizes.filter(size=>size>=2).forEach(size=>counts.set(size,(counts.get(size)||0)+1));return[...counts.entries()].sort((a,b)=>b[1]-a[1]||b[0]-a[0])[0]?.[0]||2;};
export function methodologyColumns(payload){const facts={evaluators:new Set(),dates:new Set(),teamSizes:[]};for(const schema of payload.schemas||[])for(const [key,values] of Object.entries(schema.methodology||{})){if(key==="teamSizes")facts.teamSizes.push(...(values||[]));else for(const value of values||[])facts[key]?.add(value);}const metadata=payload.metadata||{},label=text(metadata.clientName)||text(metadata.projectName),days=facts.dates.size,consultants=facts.evaluators.size||typicalTeamSize(facts.teamSizes),team=typicalTeamSize(facts.teamSizes);
  const left=["CONTEXTUL EVALUĂRII",`${payload.participantCounts.included} participanți ${label}`.trim(),`${consultants} consultanți Trend implicați`,days?`${days} ${days===1?"zi de evaluare organizată":"zile de evaluare organizate"}`:"",`${payload.behaviorAggregates.length} comportamente specifice observate`,text(metadata.exercises)?`Exerciții concepute pentru a evidenția nivelul competențelor observate: ${metadata.exercises}`:"",text(metadata.otherInstruments)?`Alte instrumente / platforme folosite: ${metadata.otherInstruments}`:""].filter(Boolean).flatMap((item,index)=>index?splitBullet(item):[item]);
  const right=["PROCESUL DE EVALUARE","Fiecare participant a fost observat de o echipă de consultanți.","Fiecare comportament a fost evaluat în faza de analiză a competențelor, folosind o grilă comună.","Observațiile au fost calibrate și integrate în platforma BHB Profiler.","Platforma a generat rezultatele pentru fiecare competență și rapoartele individuale.","Scala de evaluare utilizată: de la 1 la 5, unde 1 reprezintă nivelul minim și 5 reprezintă nivelul maxim."].flatMap((item,index)=>index?splitBullet(item):[item]);
  return{left,right,missing:{exercises:!text(metadata.exercises),otherInstruments:!text(metadata.otherInstruments)},facts:{days,consultants,team}};
}
export function methodologyPages(payload){const columns=methodologyColumns(payload),limit=6,needed=Math.max(Math.ceil((columns.left.length-1)/limit),Math.ceil((columns.right.length-1)/limit),1),pages=[];for(let index=0;index<needed;index++)pages.push({left:[index?"CONTEXTUL EVALUĂRII — CONTINUARE":columns.left[0],...columns.left.slice(1+index*limit,1+(index+1)*limit)],right:[index?"PROCESUL DE EVALUARE — CONTINUARE":columns.right[0],...columns.right.slice(1+index*limit,1+(index+1)*limit)],missing:columns.missing,facts:columns.facts});return pages;}
const percent=(value,total)=>total?Math.round(value/total*100):0;
export function executiveSummary(payload,ranked=payload.calculations.filter(item=>item.mean!==null).slice().sort((a,b)=>b.mean-a.mean)){
  const insights=behaviorInsights(payload.behaviorAggregates||[]),strengths=insights.flatMap(item=>item.strengths.map(row=>({...row,competency:item.competency}))).sort((a,b)=>b.pct2-a.pct2||b.mean-a.mean).slice(0,2),development=insights.flatMap(item=>item.development.map(row=>({...row,competency:item.competency}))).sort((a,b)=>b.pct0-a.pct0||a.mean-b.mean).slice(0,2),bands=payload.bands||{above:0,typical:0,below:0,n:0};
  return{population:payload.participantCounts?.included||0,distribution:[{label:`Peste ${bands.high}`,value:percent(bands.above,bands.n)},{label:`În intervalul benchmark (${bands.low}–${bands.high})`,value:percent(bands.typical,bands.n)},{label:`Sub ${bands.low}`,value:percent(bands.below,bands.n)}],strongest:ranked[0]||null,weakest:ranked.at(-1)||null,strengths,development};
}
export function competencyFindings(payload,ranked=payload.calculations.filter(item=>item.mean!==null).slice().sort((a,b)=>b.mean-a.mean)){const insights=new Map(behaviorInsights(payload.behaviorAggregates||[]).map(item=>[item.competency,item]));return ranked.map(item=>({item,insight:insights.get(item.competency)||{competency:item.competency,strengths:[],development:[]}}));}
export function reportPlan(payload,{scope="whole"}={}){
  const ranked=payload.calculations.filter(item=>item.mean!==null).slice().sort((a,b)=>b.mean-a.mean),records=payload.records.slice(),slides=[];const add=(family,title,data={},deliverable="main")=>slides.push({family,title,deliverable,...data});
  add("cover",payload.metadata.clientName||payload.metadata.projectName);
  methodologyPages(payload).forEach((page,index)=>add("methodology",index?"Context general raport – metodologie (continuare)":"Context general raport – metodologie",{page}));
  add("executive-summary","Executive Summary",{summary:executiveSummary(payload,ranked)});
  sectionIntroPages("Distribuția rezultatelor","distribution").forEach(page=>add("divider",page.title,{role:page.role,variant:page.variant}));
  const competencySize=Math.max(...ranked.map(item=>text(item.competency).length),0)>26?5:7;
  pageGroups(ranked,competencySize).forEach((items,index)=>add("range","Distribuția rezultatelor – pe competențe (mediană)",{items,index,total:Math.ceil(ranked.length/competencySize)}));
  pageGroups(ranked,competencySize).forEach((items,index)=>add("ranking","Distribuția rezultatelor – media pe competențe",{items,index,total:Math.ceil(ranked.length/competencySize)}));
  add("benchmark","Distribuția pe media competențelor – toată populația");
  sectionIntroPages("Analiza observațiilor","observations").forEach(page=>add("divider",page.title,{role:page.role,variant:page.variant}));
  if(payload.metadata.includeCompetencySummary)add("competency-summary","Puncte forte și recomandări de grup",{ranked});
  competencyFindings(payload,ranked).forEach(finding=>add("key-findings",`Key Findings – ${finding.item.competency}`,finding));
  if(payload.behaviorAggregates.length){sectionIntroPages("Profil comportamental","behavior").forEach(page=>add("divider",page.title,{role:page.role,variant:page.variant}));behaviorInsights(payload.behaviorAggregates).forEach(insight=>add("behavior",`Comportamente cheie – ${insight.competency}`,{insight}));}
  if(text(payload.metadata.conclusions))pageGroups(text(payload.metadata.conclusions).match(/.{1,560}(?:\s|$)|\S+?(?:\s|$)/gu)||[text(payload.metadata.conclusions)],1).forEach((items,index)=>add("conclusions",index?"Concluzii și recomandări (continuare)":"Concluzii și recomandări",{copy:items.join("")}));
  add("appendix-divider","Anexă · distribuții și rezultate individuale",{},"appendix");
  const scores=records.map(record=>({name:record.name,score:overall(record)})).sort((a,b)=>b.score-a.score);pageGroups(scores,12).forEach((items,index)=>add("benchmark-summary","Distribuția pe media competențelor",{items,index,total:Math.ceil(scores.length/12)},"appendix"));
  const comparison=records.sort((a,b)=>overall(b)-overall(a));pageGroups(comparison,participantComparisonPageSize).forEach((items,index)=>add("participant-comparison","Rezultate pe competențe / participant",{items,index,total:Math.ceil(comparison.length/participantComparisonPageSize),ranked},"appendix"));
  ranked.forEach(item=>{const rows=records.filter(record=>Number.isFinite(record.scores[item.competency])).sort((a,b)=>b.scores[item.competency]-a.scores[item.competency]),size=participantChartPageSize(rows);pageGroups(rows,size).forEach((items,index)=>add("competency-distribution",`Distribuția pe competențe – ${item.competency}`,{item,items,index,total:Math.ceil(rows.length/size)},"appendix"));});
  add("close","Business Health Bar");
  let selected=slides;if(scope==="main")selected=slides.filter(slide=>slide.deliverable==="main");if(scope==="appendix")selected=[...slides.filter(slide=>slide.deliverable==="appendix"),slides.find(slide=>slide.family==="close")];
  return selected.map((slide,index)=>({...slide,number:index+1,total:selected.length}));
}
