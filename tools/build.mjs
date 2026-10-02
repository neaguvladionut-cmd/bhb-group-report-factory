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
// Each ES module becomes one IIFE: its `import { a, b } from "./x.js";` lines read from window.__grf and its
// top-level `export` declarations are published there, so the served page needs no module loader.
const moduleSource=async file=>{
  const source=await readFile(resolve(source_,file),"utf8");
  const imports=[...source.matchAll(/^import\s*\{([^}]+)\}\s*from\s*"[^"]+";\r?$/gmu)].map(match=>`const {${match[1].split(",").map(name=>name.trim()).filter(Boolean).join(",")}}=window.__grf;`).join("\n");
  const exports=[...source.matchAll(/^export\s+(?:async\s+)?(?:function|const|let|class)\s+([A-Za-z0-9_$]+)/gmu)].map(match=>match[1]);
  const body=source.replace(/^import[^;\r\n]+;\r?\n/gmu,"").replace(/^export /gmu,"");
  return `(()=>{${imports}\n${body}\nObject.assign(window.__grf||(window.__grf={}),{${exports.join(",")}});})();`;
};
const source_=source;
let templateBootstrap = "";
try {
  const template = await readFile(resolve(source, "assets/trend/template-raport-de-grup-RO.pptx"));
  templateBootstrap = `window.__GRF_TEMPLATE_BASE64__=${JSON.stringify(template.toString("base64"))};`;
} catch {
  console.warn("cleaned Trend template missing; deploy will require /src/assets/trend/template-raport-de-grup-RO.pptx");
}
try {
  // Embedded like the template so the BHB bundle also works when the page is opened as a local file (Safari blocks file:// fetch).
  const font = await readFile(resolve(source, "assets/vendor/Poppins-Regular.ttf"));
  templateBootstrap += `window.__GRF_FONT_BASE64__=${JSON.stringify(font.toString("base64"))};`;
} catch {
  console.warn("Poppins font missing; the BHB bundle will need to fetch it");
}
const appSource=(await readFile(resolve(source,"rebuild-app.js"),"utf8")).replace(/^import[^;\r\n]+;\r?\n/gmu,"").replace(/^export /gmu,"");
const bundle=[
  templateBootstrap,
  await moduleSource("rebuild-core.js"),
  await moduleSource("rebuild-report-plan.js"),
  await moduleSource("trend-fill.js"),
  await moduleSource("template-pptx.js"),
  await moduleSource("preview.js"),
  await moduleSource("rebuild-issues.js"),
  `(()=>{const {buildPayload,createAuditWorkbook,createEvaluationSheetTemplate,downloadBundle,downloadTrendPptx:downloadPptx,mergeSelectedFiles,methodologyColumns,mountPreview,renderIssueGroup,reportPlan,unrankedCompetencies}=window.__grf;\n${appSource}\n})();`
].join("\n\n");
await writeFile(resolve(deploy,"app.js"),bundle);
const servedEntries = ["app.js", "core.js", "cover-preview.css", "index.html", "pptx.js", "preview.js", "report-plan.js", "styles.css", "assets"];
for (const entry of servedEntries) await rm(resolve(root, entry), { recursive: true, force: true });
for (const entry of await readdir(deploy, { withFileTypes: true })) {
  await cp(join(deploy, entry.name), resolve(root, entry.name), { recursive: true });
}
console.log("group-report-factory deploy built");
