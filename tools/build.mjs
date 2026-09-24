import { cp, mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const source = resolve(root, "src");
const deploy = resolve(root, "deploy");
await rm(deploy, { recursive: true, force: true });
await mkdir(deploy, { recursive: true });
await cp(source, deploy, { recursive: true });
await rm(resolve(deploy, "preview.js"), { force: true });
await rm(resolve(deploy, "cover-preview.css"), { force: true });
const deployIndex = await readFile(resolve(deploy, "index.html"), "utf8");
await writeFile(resolve(deploy, "index.html"), deployIndex.replace(/\s*<script>if\(location\.pathname\.endsWith\("\/src\/index\.html"\)\)location\.replace\("\.\.\/deploy\/index\.html"\);<\/script>/u, ""));
const code=async file=>(await readFile(resolve(source,file),"utf8")).replace(/^import[^;\r\n]+;\r?\n/gmu,"").replace(/^export /gmu,"");
const wrap=(body,exports,imports="")=>`(()=>{${imports}${body}\nObject.assign(window.__grf||(window.__grf={}),{${exports}});})();`;
const bundle=[
  wrap(await code("core.js"),"buildPayload,createAuditWorkbook,mergeSelectedFiles"),
  wrap(await code("report-plan.js"),"behaviorInsights,competencyFindings,executiveSummary,methodologyPages,participantChartPageSize,participantComparisonPageSize,sectionIntroPages,reportPlan"),
  wrap(await code("pptx.js"),"downloadPptx","const {behaviorInsights,competencyFindings,executiveSummary,methodologyPages,participantChartPageSize,participantComparisonPageSize,sectionIntroPages}=window.__grf;\n"),
  wrap(await code("pdf.js"),"createPdf,pdfPageSize","const {reportPlan}=window.__grf;\n"),
  `(()=>{const {buildPayload,createAuditWorkbook,mergeSelectedFiles,downloadPptx,createPdf,reportPlan}=window.__grf;\n${await code("app.js")}\n})();`
].join("\n\n");
await writeFile(resolve(deploy,"app.js"),bundle);
console.log("group-report-factory deploy built");
