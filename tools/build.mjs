import { cp, mkdir, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const source = resolve(root, "src");
const deploy = resolve(root, "deploy");
await rm(deploy, { recursive: true, force: true });
await mkdir(deploy, { recursive: true });
await cp(source, deploy, { recursive: true });
await cp(resolve(source, "rebuild-index.html"), resolve(deploy, "index.html"));
const deployIndex = await readFile(resolve(deploy, "index.html"), "utf8");
await writeFile(resolve(deploy, "index.html"), deployIndex.replace(/\s*<script>if\(location\.pathname\.endsWith\("\/src\/index\.html"\)\)location\.replace\("\.\.\/deploy\/index\.html"\);<\/script>/u, ""));
const code=async file=>(await readFile(resolve(source,file),"utf8")).replace(/^import[^;\r\n]+;\r?\n/gmu,"").replace(/^export /gmu,"");
const wrap=(body,exports,imports="")=>`(()=>{${imports}${body}\nObject.assign(window.__grf||(window.__grf={}),{${exports}});})();`;
let templateBootstrap = "";
try {
  const template = await readFile(resolve(source, "assets/trend/template-raport-de-grup-RO.pptx"));
  templateBootstrap = `window.__GRF_TEMPLATE_BASE64__=${JSON.stringify(template.toString("base64"))};`;
} catch {
  console.warn("cleaned Trend template missing; deploy will require /src/assets/trend/template-raport-de-grup-RO.pptx");
}
const bundle=[
  templateBootstrap,
  wrap(await code("rebuild-core.js"),"buildPayload,createAuditWorkbook,createEvaluationSheetTemplate,detectSchema,mergeSelectedFiles"),
  wrap(await code("rebuild-report-plan.js"),"HOW_TO_READ,behaviorInsights,competencyFindings,executiveSummary,methodologyColumns,methodologyPages,pageGroups,participantChartPageSize,participantComparisonPageSize,rankBehaviors,reportPlan"),
  wrap(await code("template-pptx.js"),"downloadBundle,downloadTrendPptx,generateBundle,generateTrendPptx","const {reportPlan}=window.__grf;\n"),
  wrap(await code("preview.js"),"mountPreview","const {reportPlan}=window.__grf;\n"),
  `(()=>{const {buildPayload,createAuditWorkbook,createEvaluationSheetTemplate,downloadBundle,downloadTrendPptx:downloadPptx,mergeSelectedFiles,mountPreview,reportPlan}=window.__grf;\n${await code("rebuild-app.js")}\n})();`
].join("\n\n");
await writeFile(resolve(deploy,"app.js"),bundle);
const servedEntries = ["app.js", "core.js", "cover-preview.css", "index.html", "pptx.js", "preview.js", "report-plan.js", "styles.css", "assets"];
for (const entry of servedEntries) await rm(resolve(root, entry), { recursive: true, force: true });
for (const entry of await readdir(deploy, { withFileTypes: true })) {
  await cp(join(deploy, entry.name), resolve(root, entry.name), { recursive: true });
}
console.log("group-report-factory deploy built");
