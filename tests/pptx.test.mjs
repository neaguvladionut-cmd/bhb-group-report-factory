import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { mkdtemp, mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import { basename, dirname, join, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import test from "node:test";
import vm from "node:vm";
import PptxGenJS from "/Users/vladneagu/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/pptxgenjs/dist/pptxgen.es.js";
import JSZip from "/Users/vladneagu/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/jszip/lib/index.js";
import { buildPayload } from "../src/core.js";
import { downloadPptx } from "../src/pptx.js";
import { reportPlan } from "../src/report-plan.js";

const root=resolve(dirname(fileURLToPath(import.meta.url)),"..");
const require=createRequire(import.meta.url);
const {PNG}=require("/Users/vladneagu/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/pngjs");
const runtimeBin=process.env.RUNTIME_BIN_DIR||"/Users/vladneagu/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override";
const vendor=await readFile(resolve(root,"src/assets/vendor/xlsx.full.min.js"),"utf8");
const sandbox={exports:{},module:{exports:{}},Buffer,process}; vm.runInNewContext(vendor,sandbox); const XLSX=sandbox.exports;
function workbook(rows){const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),"Simple");return XLSX.write(wb,{type:"buffer",bookType:"xlsx"})}
async function trendAssets(){const names={"trend-cover-frame-asset":"trend-template-slide-1.png","trend-methodology-frame-asset":"trend-template-slide-2.png","trend-divider-frame-asset":"trend-template-slide-3.png","trend-range-frame-asset":"trend-template-slide-4.png","trend-ranking-frame-asset":"trend-template-slide-5.png","trend-benchmark-summary-frame-asset":"trend-template-slide-7.png","trend-benchmark-frame-asset":"trend-template-slide-8.png","trend-participant-frame-asset":"trend-template-slide-9.png","trend-observations-frame-asset":"trend-template-slide-11.png","trend-key-findings-frame-asset":"trend-template-slide-12.png","trend-competency-frame-asset":"trend-template-slide-13.png","trend-summary-divider-frame-asset":"trend-template-slide-18.png","trend-behavior-frame-asset":"trend-template-slide-19.png","trend-conclusions-divider-frame-asset":"trend-template-slide-20.png","trend-conclusions-frame-asset":"trend-template-slide-21.png","trend-closing-frame-asset":"trend-template-slide-22.png"};return Object.fromEntries(await Promise.all(Object.entries(names).map(async([id,name])=>[id,`data:image/png;base64,${(await readFile(resolve(root,"src/assets/images",name))).toString("base64")}`])));}
async function renderPptx(file,outDir){await mkdir(outDir,{recursive:true});const profile=join(outDir,"libreoffice-profile");await mkdir(profile,{recursive:true});execFileSync(join(runtimeBin,"soffice"),[`-env:UserInstallation=${pathToFileURL(profile).href}`,"--headless","--convert-to","pdf","--outdir",outDir,file],{stdio:"pipe"});const pdf=join(outDir,`${basename(file,".pptx")}.pdf`),prefix=join(outDir,"slide");execFileSync(join(runtimeBin,"pdftoppm"),["-png","-r","120",pdf,prefix],{stdio:"pipe"});return execFileSync("find",[outDir,"-maxdepth","1","-name","slide-*.png"],{encoding:"utf8"}).trim().split("\n").filter(Boolean).sort((a,b)=>Number(a.match(/-(\d+)\.png$/u)?.[1])-Number(b.match(/-(\d+)\.png$/u)?.[1]));}
async function png(path){return PNG.sync.read(await readFile(path));}
function changedPixels(actual,reference,{x,y,w,h}){assert.equal(actual.width,reference.width);assert.equal(actual.height,reference.height);const sx=value=>Math.max(0,Math.min(actual.width,Math.round(value/13.333*actual.width))),sy=value=>Math.max(0,Math.min(actual.height,Math.round(value/7.5*actual.height))),left=sx(x),top=sy(y),right=sx(x+w),bottom=sy(y+h);let changed=0,total=0;for(let py=top;py<bottom;py+=1){for(let px=left;px<right;px+=1){const offset=(py*actual.width+px)*4,delta=Math.abs(actual.data[offset]-reference.data[offset])+Math.abs(actual.data[offset+1]-reference.data[offset+1])+Math.abs(actual.data[offset+2]-reference.data[offset+2]);if(delta>12)changed+=1;total+=1;}}return{changed,total,ratio:changed/total};}
function slideTexts(xml){return[...xml.matchAll(/<a:t>([^<]*)<\/a:t>/gu)].map(match=>match[1].replaceAll("&amp;","&").replaceAll("&lt;","<").replaceAll("&gt;",">"));}

test("PPTX uses PptxGenJS native chart and workbook parts",async()=>{
  const summary=workbook([["code","name","cod cp","People Management","Colaborare","Gestionarea schimbării"],["","D210 Ana","A-1",4,3,2],["","D210 Bogdan","A-2",2,5,4]]);
  const detail=workbook([["code","name the person evaluated","cod ac","Competente","People Management","Colaborare"],["","","","Subcompetente","PM","COL"],["","","","behavior","Descrie clar obiectivul complet","Colaborează activ cu colegii"],["","D210 Ana","A-1","",2,1],["","D210 Bogdan","A-2","",0,2]]);
  const payload=buildPayload(XLSX,[{name:"summary.xlsx",bytes:summary},{name:"detail.xlsx",bytes:detail}],{projectName:"D210 Test",clientName:"Client Test",reportDate:"septembrie 2026"});
  const dir=await mkdtemp(join(tmpdir(),"grf-d210-pptx-")); const file=join(dir,"report.pptx");
  const assetNames={
    "pptx-cover-asset":"bhb-cover-template.png",
    "pptx-content-asset":"bhb-content-white.png",
    "pptx-closing-asset":"bhb-closing-template.png",
    "pptx-divider-asset":"bhb-intro-divider.png",
    "pptx-report-background-asset":"bhb-report-background.png",
    "pptx-section-gradient-asset":"bhb-section-gradient.png",
    "pptx-observations-background-asset":"bhb-observations-background.png",
    "pptx-behavior-background-asset":"bhb-behavior-background.png",
    "pptx-observer-photo-asset":"bhb-observer-photo.jpg",
    "pptx-logo-asset":"bhb-logo.svg",
    "pptx-symbol-asset":"bhb-symbol.svg",
    "pptx-report-illustration-asset":"bhb-competency-profiler-report-illustration.webp",
    "pptx-range-legend-asset":"bhb-range-legend-template.png"
  };
  const assets=Object.fromEntries(await Promise.all(Object.entries(assetNames).map(async([id,name])=>{const ext=name.split(".").pop();const mime=ext==="svg"?"image/svg+xml":ext==="jpg"?"image/jpeg":"image/png";return [id,`data:${mime};base64,${(await readFile(resolve(root,"src/assets/images",name))).toString("base64")}`]})));
  globalThis.window={PptxGenJS,JSZip,__bhbAssets:assets,__writeFile:writeFile}; await downloadPptx(payload,file);
  const files=execFileSync("unzip",["-Z1",file],{encoding:"utf8"}).split("\n");
  const charts=files.filter(name=>/^ppt\/charts\/chart\d+\.xml$/.test(name));
  const workbooks=files.filter(name=>/^ppt\/embeddings\/Microsoft_Excel_Worksheet\d+\.xlsx$/.test(name));
  assert.equal(charts.length,7);
  assert.equal(workbooks.length,charts.length);
  assert(files.filter(name=>/^ppt\/slides\/slide\d+\.xml$/.test(name)).length>=14);
  const allText=execFileSync("unzip",["-p",file,"ppt/slides/slide*.xml"],{encoding:"utf8"});
  assert.match(allText,/METODOLOGIE/iu);
  assert.match(allText,/3\.00/);
  assert.doesNotMatch(allText,/Puncte forte și recomandări de grup/iu);
  assert.doesNotMatch(allText,/DE COMPLETAT DE CONSULTANT|proveniență|consultant-owned/iu);
  const pptxSource=await readFile(resolve(root,"src/pptx.js"),"utf8");
  assert.match(pptxSource,/pptx-range-legend-asset/u,"range slide uses the canonical AC legend asset");
  assert.doesNotMatch(pptxSource,/Arial|360 Profiler|fit:\s*["']shrink/iu);
  const visibleXml=execFileSync("unzip",["-p",file,...files.filter(name=>/^ppt\/(slides\/slide\d+|charts\/chart\d+)\.xml$/u.test(name))],{encoding:"utf8"});
  assert.doesNotMatch(visibleXml,/Arial|360 Profiler/iu);
  assert.match(pptxSource,/pptx-content-asset/);
  assert.match(pptxSource,/pptx-closing-asset/);
  const coverXml=execFileSync("unzip",["-p",file,"ppt/slides/slide1.xml"],{encoding:"utf8"});
  assert.match(coverXml,/sz="5000"/u,"cover title uses the requested 50pt BHB title hierarchy");
  assert.match(coverXml,/sz="2400"/u,"cover subtitle/context uses the requested 24pt hierarchy");
  assert.equal((coverXml.match(/D210 Test/gu)||[]).length,1,"cover does not repeat the project name when the client title is present");
  const dividerXml=files.filter(name=>/^ppt\/slides\/slide\d+\.xml$/u.test(name)).map(name=>execFileSync("unzip",["-p",file,name],{encoding:"utf8"}));assert.equal(dividerXml.length,27,"BHB whole output includes the new main-plus-appendix structure");assert(dividerXml.some(xml=>xml.includes("Executive Summary")),"BHB whole output includes Executive Summary");assert.equal(dividerXml.filter(xml=>xml.includes("Key Findings")).length,3,"BHB creates one Key Findings slide per competency");assert(dividerXml.some(xml=>xml.includes("CONFIDENȚIAL")),"BHB cover marks the report confidential");assert(dividerXml.some(xml=>xml.includes("Anexă")),"BHB whole output includes the appendix");
  const media=files.filter(name=>/^ppt\/media\/.*\.png$/u.test(name));
  const mediaBytes=media.map(name=>execFileSync("unzip",["-p",file,name],{maxBuffer:20*1024*1024}));
  for(const name of ["bhb-report-background.png","bhb-section-gradient.png","bhb-observations-background.png","bhb-behavior-background.png","bhb-range-legend-template.png"]){const reference=await readFile(resolve(root,"src/assets/images",name));assert(mediaBytes.some(bytes=>bytes.equals(reference)),`${name} must be embedded as an authentic full-slide BHB asset`);}
  const jpgMedia=files.filter(name=>/^ppt\/media\/.*\.jpeg$/u.test(name));const jpgBytes=jpgMedia.map(name=>execFileSync("unzip",["-p",file,name],{maxBuffer:20*1024*1024}));const observerPhoto=await readFile(resolve(root,"src/assets/images","bhb-observer-photo.jpg"));assert(jpgBytes.length>0,"observer photo must be embedded in the reference opener");
  const chartXml=charts.map(name=>execFileSync("unzip",["-p",file,name],{encoding:"utf8"}));
  const rangeXml=chartXml.find(xml=>xml.includes("Cel mai mic scor")&&xml.includes("Mediana sau scorul median"));
  assert.match(rangeXml,/<c:stockChart>/u);
  assert.match(rangeXml,/<c:hiLowLines>/u);
  assert.doesNotMatch(rangeXml,/<c:legend>/u,"range chart relies on the canonical right-side legend, not a redundant native legend");
  assert.match(rangeXml,/<c:symbol val="dash"\/><c:size val="30"\/>/u,"range endpoints use the source-sized native markers");
  assert.match(rangeXml,/<c:symbol val="diamond"\/><c:size val="36"\/>/u,"range median uses the larger source-sized native marker");
  const stockBody=rangeXml.match(/<c:stockChart>([\s\S]*?)<\/c:stockChart>/u)?.[1];
  assert(stockBody,"range chart contains one native stockChart body");
  assert.equal((stockBody.match(/<c:ser>/gu)||[]).length,3,"stock chart retains MIN, MAX and MEDIAN series");
  assert.equal((stockBody.match(/<c:axId /gu)||[]).length,2,"stock chart has only the canonical category and value axes");
  assert(stockBody.indexOf("<c:dLbls>")<stockBody.indexOf("<c:hiLowLines>")&&stockBody.indexOf("<c:hiLowLines>")<stockBody.indexOf("<c:axId "),"stock chart follows canonical dLbls → hiLowLines → axIds order");
  assert.doesNotMatch(stockBody,/<c:lineChart>|<c:multiLvlStrRef>|<c:marker val=/u);
  [...stockBody.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/gu)].forEach((match,index)=>{assert.match(match[1],/<c:strRef>[\s\S]*?<c:strCache>[\s\S]*?<c:ptCount val="[1-9]\d*"/u,`range series ${index+1} has non-empty category cache`);assert.match(match[1],/<c:numCache>[\s\S]*?<c:ptCount val="[1-9]\d*"[\s\S]*?<c:v>[^<]+<\/c:v>/u,`range series ${index+1} has non-empty value cache`);});
  const rankingXml=chartXml.find(xml=>xml.includes("Colaborare")&&xml.includes("People Management")&&xml.includes("Gestionarea schimbării")&&xml.includes('<c:legendPos val="b"/>')&&xml.includes('<c:barDir val="bar"/>'));
  assert.match(rankingXml,/<c:legend><c:legendPos val="b"\/>/u,"competency-average chart has a native legend");
  assert.match(rankingXml,/<c:catAx>[\s\S]*?<c:delete val="1"\/>/u,"competency-average chart hides its misleading one-category axis");
  assert.match(rankingXml,/<a:defRPr b="1"[^>]*sz="2000"/u,"competency-average labels use template-sized bold text");
  assert.match(rankingXml,/<c:dLblPos val="inEnd"\/>/u,"competency-average labels remain inside their color-coded bars");
  const rankingSeries=[...rankingXml.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/gu)].map(match=>match[1]);
  assert.equal(rankingSeries.length,3,"every competency is an explicit legend-bearing native series");
  [["Colaborare","0A375B"],["People Management","27808F"],["Gestionarea schimbării","4A8C61"]].forEach(([name,color])=>{const series=rankingSeries.find(item=>item.includes(`<c:v>${name}<\/c:v>`));assert.match(series,new RegExp(`<c:v>${name}<\\/c:v>`),`${name} keeps its full legend name`);assert.match(series,new RegExp(`<a:srgbClr val="${color}"`),`${name} keeps its stable competency color`);});
  const benchmarkXml=chartXml.find(xml=>xml.includes("Mai mică de 2.75")&&xml.includes("Între 2.75 – 3.5")&&xml.includes("Peste 3.5"));
  assert.match(benchmarkXml,/<c:grouping val="percentStacked"\/>/u);assert.match(benchmarkXml,/<c:numFmt formatCode="0%"/u);assert.match(benchmarkXml,/<c:max val="1"\/>/u);
  const benchmarkValues=[...benchmarkXml.matchAll(/<c:numCache>[\s\S]*?<c:v>([^<]+)<\/c:v>/gu)].map(match=>Number(match[1]));assert.equal(benchmarkValues.length,3);assert(Math.abs(benchmarkValues.reduce((sum,value)=>sum+value,0)-1)<1e-9,"benchmark workbook series must sum to 100%");
  const benchmarkSeries=[...benchmarkXml.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/gu)].map(match=>match[1]);
  [["Mai mică de 2.75","FFFFFF"],["Între 2.75 – 3.5","003057"],["Peste 3.5","003057"]].forEach(([name,color])=>{const series=benchmarkSeries.find(item=>item.includes(`<c:v>${name}</c:v>`));assert.match(series,/<a:defRPr b="1"[^>]*sz="1800"/u,`${name} percentage label is bold and 18pt`);assert.match(series,new RegExp(`<c:dLbls>[\\s\\S]*?<a:srgbClr val="${color}"`),`${name} percentage label has its contrast color`);});
  const rangeSlides=files.filter(name=>/^ppt\/slides\/slide\d+\.xml$/u.test(name)).map(name=>execFileSync("unzip",["-p",file,name],{encoding:"utf8"})).filter(xml=>xml.includes("Distribuția rezultatelor – pe competențe (mediană)"));
  assert(rangeSlides.length>0);assert(rangeSlides.every(xml=>!xml.includes("<a:tbl>")),"range slides must use only their native chart category labels, never a duplicate label table");
  assert(chartXml.some(xml=>/<c:cat>[\s\S]*Ana[\s\S]*Bogdan[\s\S]*<\/c:cat>/u.test(xml)&&/<c:ptCount val="2"\/>/u.test(xml)),"a competency chart must contain participant categories and participant-level values");
  const opted=buildPayload(XLSX,[{name:"summary.xlsx",bytes:summary},{name:"detail.xlsx",bytes:detail}],{projectName:"D210 Test"});
  opted.metadata.includeCompetencySummary=true;
  const optedFile=join(dir,"report-opted.pptx"); await downloadPptx(opted,optedFile);
  assert.match(execFileSync("unzip",["-p",optedFile,"ppt/slides/slide*.xml"],{encoding:"utf8"}),/Puncte forte și recomandări de grup/iu);
  const largeRows=[["code","name","cod cp","People Management","Colaborare","Orientare rezultate"],...Array.from({length:45},(_,index)=>["",`Participant ${String(index+1).padStart(2,"0")}`,`A-${index+1}`,1+index%5,1+(index+1)%5,1+(index+2)%5])];
  const largePayload=buildPayload(XLSX,[{name:"summary-45.xlsx",bytes:workbook(largeRows)}],{projectName:"AC 45 participanți"});
  const largeFile=join(dir,"report-45.pptx"); await downloadPptx(largePayload,largeFile);
  const largeFiles=execFileSync("unzip",["-Z1",largeFile],{encoding:"utf8"}).split("\n");
  assert.equal(largeFiles.filter(name=>/^ppt\/charts\/chart\d+\.xml$/.test(name)).length,29);
  assert.equal(largeFiles.filter(name=>/^ppt\/embeddings\/Microsoft_Excel_Worksheet\d+\.xlsx$/.test(name)).length,29);
  const largeSlides=largeFiles.filter(name=>/^ppt\/slides\/slide\d+\.xml$/.test(name)); assert.equal(largeSlides.length,47);
  const largeText=execFileSync("unzip",["-p",largeFile,"ppt/slides/slide*.xml"],{encoding:"utf8"});
  assert.match(largeText,/Participant 45/);
  assert.match(largeText,/45 din 45/);
  assert.doesNotMatch(largeText,/Puncte forte și recomandări de grup/iu);
  await rm(dir,{recursive:true,force:true});
});

test("TREND renderer supports whole, main-only and appendix-only delivery",async()=>{
  const payload={metadata:{projectName:"Trend fixture",clientName:"Client fixture",reportDate:"septembrie 2026",conclusions:"Concluzii aprobate."},calculations:[{competency:"People Management",mean:3.5,median:3.5,min:2,max:5,n:2},{competency:"Colaborare",mean:3,median:3,min:1,max:4,n:2}],records:[{name:"Participant 1",scores:{"People Management":4,Colaborare:3}},{name:"Participant 2",scores:{"People Management":3,Colaborare:3}}],bands:{low:2.75,high:3.5,below:0,typical:2,above:0,n:2},participantCounts:{included:2,total:2},behaviorAggregates:[{competency:"People Management",behavior:"Descrie clar obiectivul complet",pct0:0,pct2:1,mean:3}],schemas:[]};
  const dir=await mkdtemp(join(tmpdir(),"grf-trend-pptx-"));globalThis.window={PptxGenJS,JSZip,__bhbAssets:{},__writeFile:writeFile};
  for(const scope of ["whole","main","appendix"]){const file=join(dir,`${scope}.pptx`);await downloadPptx(payload,file,{renderer:"trend",scope});const names=execFileSync("unzip",["-Z1",file],{encoding:"utf8"}).split("\n"),slideNames=names.filter(name=>/^ppt\/slides\/slide\d+\.xml$/u.test(name)),slideXml=slideNames.map(name=>execFileSync("unzip",["-p",file,name],{encoding:"utf8"})),text=slideXml.join(""),charts=names.filter(name=>/^ppt\/charts\/chart\d+\.xml$/u.test(name)),workbooks=names.filter(name=>/^ppt\/embeddings\/Microsoft_Excel_Worksheet\d+\.xlsx$/u.test(name));assert(slideNames.length>0);assert.equal(charts.length,workbooks.length,`${scope} keeps one embedded workbook per native chart`);const probeTitle=scope==="appendix"?"Distribuția pe media competențelor":"Executive Summary",probe=slideXml.find(xml=>xml.includes(probeTitle));assert(probe,`${scope} includes a framed content route`);assert.equal((probe.match(new RegExp(probeTitle,"gu"))||[]).length,1,`${scope} generated heading appears once in editable slide text`);if(scope!=="appendix"){assert.match(text,/Executive Summary/);assert.match(text,/CONFIDENȚIAL/);}else{assert.match(text,/Anexă/);assert.doesNotMatch(text,/Executive Summary|CONFIDENȚIAL/);}if(scope==="main")assert.doesNotMatch(text,/Anexă · rezultate detaliate/);}
  await rm(dir,{recursive:true,force:true});
});

test("both PPTX renderers omit every disabled family across all delivery scopes",async()=>{
  const payload={metadata:{projectName:"Secțiuni oprite",clientName:"Client fixture",reportDate:"septembrie 2026",conclusions:"Concluzii aprobate.",sections:{distribution:false,benchmark:false,keyFindings:false,behavior:false,conclusions:false,appendixBenchmark:false,appendixComparison:false,appendixDistribution:false}},calculations:[{competency:"People Management",mean:3.5,median:3.5,min:2,max:5,n:2},{competency:"Colaborare",mean:3,median:3,min:1,max:4,n:2}],records:[{name:"Participant 1",scores:{"People Management":4,Colaborare:3}},{name:"Participant 2",scores:{"People Management":3,Colaborare:3}}],bands:{low:2.75,high:3.5,below:0,typical:2,above:0,n:2},participantCounts:{included:2,total:2},behaviorAggregates:[{competency:"People Management",behavior:"Descrie clar obiectivul complet",pct0:0,pct2:1,mean:3}],schemas:[]};
  assert.deepEqual([reportPlan(payload,{scope:"whole"}).length,reportPlan(payload,{scope:"main"}).length,reportPlan(payload,{scope:"appendix"}).length],[4,4,1]);
  const dir=await mkdtemp(join(tmpdir(),"grf-disabled-sections-"));
  try{for(const renderer of ["bhb","trend"]){for(const scope of ["whole","main","appendix"]){const file=join(dir,`${renderer}-${scope}.pptx`);globalThis.window={PptxGenJS,JSZip,__bhbAssets:{},__writeFile:writeFile};await downloadPptx(payload,file,{renderer,scope});const names=execFileSync("unzip",["-Z1",file],{encoding:"utf8"}).split("\n"),slides=names.filter(name=>/^ppt\/slides\/slide\d+\.xml$/u.test(name)),xml=slides.map(name=>execFileSync("unzip",["-p",file,name],{encoding:"utf8"})),text=xml.join(""),charts=names.filter(name=>/^ppt\/charts\/chart\d+\.xml$/u.test(name)),workbooks=names.filter(name=>/^ppt\/embeddings\/Microsoft_Excel_Worksheet\d+\.xlsx$/u.test(name));assert.equal(slides.length,scope==="appendix"?1:4,`${renderer}/${scope} retains only always-on slides`);assert.equal(charts.length,0,`${renderer}/${scope} has no disabled native charts`);assert.equal(workbooks.length,charts.length,`${renderer}/${scope} keeps chart/workbook pairing`);if(scope!=="appendix"){assert.match(text,/Executive Summary/);assert.match(text,/CONFIDENȚIAL/);}assert.doesNotMatch(text,/Distribuția rezultatelor|Distribuția pe media competențelor|Key Findings|Analiza observațiilor|Profil comportamental|Comportamente cheie|Concluzii|Anexă/iu,`${renderer}/${scope} contains no disabled family or divider`);}}}finally{await rm(dir,{recursive:true,force:true});}
});

test("TREND rendered routes clear baked template pixels while retaining frame chrome",async()=>{
  const payload={metadata:{projectName:"Trend fixture",clientName:"Client fixture",reportDate:"septembrie 2026",conclusions:"Concluzii aprobate."},calculations:[{competency:"People Management",mean:3.5,median:3.5,min:2,max:5,n:2},{competency:"Colaborare",mean:3,median:3,min:1,max:4,n:2}],records:[{name:"Participant 1",scores:{"People Management":4,Colaborare:3}},{name:"Participant 2",scores:{"People Management":3,Colaborare:3}}],bands:{low:2.75,high:3.5,below:0,typical:2,above:0,n:2},participantCounts:{included:2,total:2},behaviorAggregates:[{competency:"People Management",behavior:"Descrie clar obiectivul complet",pct0:0,pct2:1,mean:3}],schemas:[]};
  const dir=await mkdtemp(join(tmpdir(),"grf-trend-render-")),assets=await trendAssets(),dividerTitles=new Set(["Distribuția rezultatelor","Analiza observațiilor","Profil comportamental","Anexă · rezultate detaliate"]);let inspectedContent=0,inspectedDividers=0;
  try{for(const scope of ["whole","main","appendix"]){const actual=join(dir,`${scope}-assets.pptx`),reference=join(dir,`${scope}-reference.pptx`);globalThis.window={PptxGenJS,JSZip,__bhbAssets:assets,__writeFile:writeFile};await downloadPptx(payload,actual,{renderer:"trend",scope});globalThis.window={PptxGenJS,JSZip,__bhbAssets:{},__writeFile:writeFile};await downloadPptx(payload,reference,{renderer:"trend",scope});const actualPngs=await renderPptx(actual,join(dir,`${scope}-assets`)),referencePngs=await renderPptx(reference,join(dir,`${scope}-reference`)),slideNames=execFileSync("unzip",["-Z1",actual],{encoding:"utf8"}).split("\n").filter(name=>/^ppt\/slides\/slide\d+\.xml$/u.test(name));assert.equal(actualPngs.length,slideNames.length,`${scope} asset-backed route rendered every slide`);assert.equal(referencePngs.length,slideNames.length,`${scope} clean reference rendered every slide`);for(let index=0;index<slideNames.length;index+=1){const xml=execFileSync("unzip",["-p",actual,slideNames[index]],{encoding:"utf8"}),texts=slideTexts(xml),actualPng=await png(actualPngs[index]),referencePng=await png(referencePngs[index]),isDivider=texts.some(text=>dividerTitles.has(text)),isCover=texts.includes("CONFIDENȚIAL"),isClosing=texts.length===0;if(isDivider){const reset=changedPixels(actualPng,referencePng,{x:7.7,y:2.65,w:5.633,h:2.4});assert.equal(reset.changed,0,`${scope} slide ${index+1} has no baked divider-title or subtitle pixels across the old title-panel footprint`);const photo=changedPixels(actualPng,referencePng,{x:1.0,y:1.0,w:5.3,h:5.2});assert(photo.ratio>.05,`${scope} slide ${index+1} retains authentic divider photography/chrome`);inspectedDividers+=1;}else if(!isCover&&!isClosing){const reset=changedPixels(actualPng,referencePng,{x:.73,y:.01,w:12.593,h:7.16});assert.equal(reset.changed,0,`${scope} slide ${index+1} has no baked table, legend, axis, box, side-panel, or heading pixels in the rendered content canvas`);const rail=changedPixels(actualPng,referencePng,{x:0,y:0,w:.7,h:7.5});assert(rail.ratio>.01,`${scope} slide ${index+1} retains the authentic TREND rail and page marker`);inspectedContent+=1;}}}assert(inspectedContent>=20,"render regression must exercise content frames across whole, main, and appendix routes");assert(inspectedDividers>=7,"render regression must exercise section and appendix divider frames across all routes");}finally{await rm(dir,{recursive:true,force:true});}
});
