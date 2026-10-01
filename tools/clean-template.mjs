import { execFile } from "node:child_process";
import { mkdir, mkdtemp, readdir, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join, posix, relative, resolve } from "node:path";
import { promisify } from "node:util";

const run = promisify(execFile);
const root = resolve(new URL("..", import.meta.url).pathname);
const [sourceArg, outputArg] = process.argv.slice(2).filter((arg) => !arg.startsWith("--"));
const source = resolve(sourceArg || "50-source-materials/11. Competency Profiler/2. Rapoarte/Template raport de grup/Template - raport de grup - RO.pptx");
const output = resolve(outputArg || join(root, "src/assets/trend/template-raport-de-grup-RO.pptx"));
// The generated output is intentionally git-ignored until Vlad confirms H1.

if (!sourceArg && !process.argv.includes("--source-default-ok")) {
  console.error("Usage: node tools/clean-template.mjs <source.pptx> [output.pptx]");
  process.exitCode = 2;
  process.exit();
}

const xmlFiles = [];
const removedParts = [];
const textParts = [];

async function walk(directory) {
  const entries = await readdir(directory, { withFileTypes: true });
  const result = [];
  for (const entry of entries) {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) result.push(...await walk(path));
    else result.push(path);
  }
  return result;
}

function stripXmlPartReferences(xml) {
  return xml
    .replace(/<Override\b[^>]*(?:notesSlides|notesMasters|comments|commentAuthors|customXml|changesInfos|docProps\/custom\.xml|revisionInfo\.xml)[^>]*\/>/giu, "")
    .replace(/<Default\b[^>]*(?:comments|vml)[^>]*\/>/giu, "")
    .replace(/<Relationship\b[^>]*(?:notesSlides|notesMasters|comments|commentAuthors|customXml|changesInfos|docProps\/custom\.xml|revisionInfo\.xml)[^>]*\/>/giu, "")
    .replace(/<Relationship\b[^>]*Target="[^"]*docProps\/custom\.xml"[^>]*\/>/giu, "")
    .replace(/<Relationship\b[^>]*Target="[^"]*revisionInfo\.xml"[^>]*\/>/giu, "")
    .replace(/<p:section\b[^>]*r:id="[^"]+"[^>]*\/>/giu, "");
}

function scrubMetadata(xml) {
  return xml
    .replace(/<(?:[\w.-]+:)?(?:creator|lastModifiedBy|Company|Manager)\b[^>]*>[\s\S]*?<\/(?:[\w.-]+:)?(?:creator|lastModifiedBy|Company|Manager)>/giu, (match) => match.replace(/>([\s\S]*)</u, "><"))
    .replace(/<(?:dc:title|dc:subject|dc:description|cp:keywords|cp:category)\b[^>]*>[\s\S]*?<\/(?:dc:title|dc:subject|dc:description|cp:keywords|cp:category)>/giu, (match) => match.replace(/>([\s\S]*)</u, "><"))
    .replace(/<cp:revision\b[^>]*>[\s\S]*?<\/cp:revision>/giu, "")
    .replace(/<dcterms:created[^>]*>[\s\S]*?<\/dcterms:created>/giu, "")
    .replace(/<dcterms:modified[^>]*>[\s\S]*?<\/dcterms:modified>/giu, "")
    .replace(/<(?:TotalTime|Words|Paragraphs|Notes|TitlesOfParts|HeadingPairs)\b[^>]*>[\s\S]*?<\/(?:TotalTime|Words|Paragraphs|Notes|TitlesOfParts|HeadingPairs)>/giu, "");
}

function removeNotesMasterList(xml) {
  return xml.replace(/<p:notesMasterIdLst\b[^>]*>[\s\S]*?<\/p:notesMasterIdLst>/giu, "");
}

function clearSlideText(xml) {
  return xml.replace(/(<a:t(?:\s[^>]*)?>)[\s\S]*?(<\/a:t>)/gu, "$1$2");
}

function clearChartCaches(xml) {
  return xml
    .replace(/(<c:f(?:\s[^>]*)?>)[\s\S]*?(<\/c:f>)/gu, "$1$2")
    .replace(/(<c:v(?:\s[^>]*)?>)[\s\S]*?(<\/c:v>)/gu, "$1$2");
}

function clearWorkbookCells(xml) {
  return xml
    .replace(/(<v(?:\s[^>]*)?>)[\s\S]*?(<\/v>)/gu, "$1$2")
    .replace(/(<t(?:\s[^>]*)?>)[\s\S]*?(<\/t>)/gu, "$1$2");
}

async function clean(sourceDir) {
  const removable = [
    "customXml",
    "[trash]",
    "ppt/notesSlides",
    "ppt/notesMasters",
    "ppt/changesInfos",
    "ppt/comments",
    "ppt/commentAuthors",
    "docProps/custom.xml",
    "ppt/revisionInfo.xml"
  ];
  for (const item of removable) {
    const path = join(sourceDir, item);
    try {
      await stat(path);
      await rm(path, { recursive: true, force: true });
      removedParts.push(item);
    } catch (error) {
      if (error.code !== "ENOENT") throw error;
    }
  }

  const files = await walk(sourceDir);
  for (const file of files) {
    const part = relative(sourceDir, file).replaceAll("\\", "/");
    if (part.endsWith(".xml") || part.endsWith(".rels")) {
      xmlFiles.push(part);
      let xml = await readFile(file, "utf8");
      xml = stripXmlPartReferences(xml);
      xml = scrubMetadata(xml);
      if (part === "ppt/presentation.xml") xml = removeNotesMasterList(xml);
      if (/^ppt\/slides\/slide\d+\.xml$/u.test(part)) {
        xml = clearSlideText(xml);
        textParts.push(part);
      }
      if (/^ppt\/charts\/chart\d+\.xml$/u.test(part)) xml = clearChartCaches(xml);
      if (part === "docProps/app.xml") {
        xml = xml.replace(/<Application>[^<]*<\/Application>/giu, "<Application>Microsoft PowerPoint</Application>");
      }
      await writeFile(file, xml);
    } else if (part.startsWith("ppt/embeddings/") && part.endsWith(".xlsx")) {
      const nested = await mkdtemp(join(tmpdir(), "grf-clean-xlsx-"));
      const nestedOutput = join(nested, "package");
      await mkdir(nestedOutput, { recursive: true });
      await run("unzip", ["-q", file, "-d", nestedOutput]);
      for (const nestedFile of await walk(nestedOutput)) {
        const nestedPart = relative(nestedOutput, nestedFile).replaceAll("\\", "/");
        if (nestedPart.endsWith(".xml") || nestedPart.endsWith(".rels")) {
          let xml = clearWorkbookCells(await readFile(nestedFile, "utf8"));
          xml = scrubMetadata(xml);
          if (nestedPart === "docProps/app.xml") xml = xml.replace(/<Application>[^<]*<\/Application>/giu, "<Application>Microsoft Excel</Application>");
          await writeFile(nestedFile, stripXmlPartReferences(xml));
        }
      }
      const old = process.cwd();
      process.chdir(nestedOutput);
      await run("zip", ["-q", "-X", "-r", file, "."]);
      process.chdir(old);
      await rm(nested, { recursive: true, force: true });
    }
  }
}

function relationshipPart(xmlPart) {
  const directory = posix.dirname(xmlPart);
  const file = posix.basename(xmlPart);
  return directory === "." ? `_rels/${file}.rels` : `${directory}/_rels/${file}.rels`;
}

function relationshipTarget(sourcePart, target) {
  if (!target || target.startsWith("#") || /^[a-z]+:/iu.test(target)) return null;
  const decoded = decodeURIComponent(target.split("#")[0]);
  return posix.normalize(posix.join(posix.dirname(sourcePart), decoded)).replace(/^\.\//u, "");
}

function relationshipsById(xml) {
  return new Map([...xml.matchAll(/<Relationship\b[^>]*>/giu)].map(([tag]) => {
    const id = tag.match(/\bId="([^"]+)"/iu)?.[1];
    const target = tag.match(/\bTarget="([^"]+)"/iu)?.[1];
    return [id, target];
  }).filter(([id, target]) => id && target));
}

async function selfCheck(packageDir) {
  const parts = new Set((await walk(packageDir)).map((file) => relative(packageDir, file).replaceAll("\\", "/")));
  const failures = [];
  const xmlParts = [...parts].filter((part) => part.endsWith(".xml") || part.endsWith(".rels"));
  for (const part of [...parts].filter((candidate) => candidate.endsWith(".rels"))) {
    const xml = await readFile(join(packageDir, part), "utf8");
    const sourcePart = part === "_rels/.rels" ? "" : part.replace(/\/_rels\/([^/]+)\.rels$/u, "/$1");
    for (const [id, target] of relationshipsById(xml)) {
      const targetPart = relationshipTarget(sourcePart, target);
      if (targetPart && !parts.has(targetPart)) failures.push(`${part} ${id} -> missing ${targetPart}`);
    }
  }
  const contentTypes = parts.has("[Content_Types].xml") ? await readFile(join(packageDir, "[Content_Types].xml"), "utf8") : "";
  for (const [tag] of contentTypes.matchAll(/<Override\b[^>]*>/giu)) {
    const partName = tag.match(/\bPartName="([^"]+)"/iu)?.[1];
    if (!partName) continue;
    const target = partName.replace(/^\//u, "");
    if (!parts.has(target)) failures.push(`[Content_Types].xml Override -> missing ${target}`);
  }
  for (const part of xmlParts.filter((candidate) => candidate.endsWith(".xml"))) {
    const xml = await readFile(join(packageDir, part), "utf8");
    if (/<(?:[\w.-]+:)?embeddedFontLst\b/iu.test(xml)) failures.push(`${part} contains embeddedFontLst`);
    for (const list of xml.matchAll(/<((?:[\w.-]+:)?[\w.-]*IdLst)\b[^>]*>[\s\S]*?<\/\1>/giu)) {
      const relsPath = relationshipPart(part);
      const rels = parts.has(relsPath) ? relationshipsById(await readFile(join(packageDir, relsPath), "utf8")) : new Map();
      for (const id of [...list[0].matchAll(/r:id="([^"]+)"/giu)].map((match) => match[1])) {
        if (!rels.has(id)) failures.push(`${part} ${list[1]} ${id} -> missing relationship`);
        else {
          const target = relationshipTarget(part, rels.get(id));
          if (target && !parts.has(target)) failures.push(`${part} ${list[1]} ${id} -> missing ${target}`);
        }
      }
    }
    for (const field of xml.matchAll(/<(?:[\w.-]+:)?(?:creator|lastModifiedBy|Company|Manager)\b[^>]*>([^<]*)<\/(?:[\w.-]+:)?(?:creator|lastModifiedBy|Company|Manager)>/giu)) {
      if (field[1].trim()) failures.push(`${part} contains non-empty personal metadata`);
    }
  }
  for (const part of parts) if (/^ppt\/fonts\//iu.test(part)) failures.push(`${part} is a forbidden ppt/fonts part`);
  if (failures.length) throw new Error(`clean-template self-check failed:\n${failures.join("\n")}`);
  return { danglingReferences: "PASS", metadata: "PASS", fonts: "PASS" };
}

await stat(source);
const work = await mkdtemp(join(tmpdir(), "grf-clean-template-"));
try {
  await run("unzip", ["-q", source, "-d", work]);
  await clean(work);
  const checks = await selfCheck(work);
  await mkdir(dirname(output), { recursive: true });
  await rm(output, { force: true });
  const old = process.cwd();
  process.chdir(work);
  await run("zip", ["-q", "-X", "-r", output, "."]);
  process.chdir(old);
  const parts = (await walk(work)).map((file) => relative(work, file).replaceAll("\\", "/")).sort();
  console.log(`clean-template source=${source}`);
  console.log(`clean-template output=${output}`);
  console.log(`removed=${removedParts.join(",") || "none"}`);
  console.log(`cleared-slide-text=${textParts.length}`);
  console.log(`xml-parts=${xmlFiles.length}`);
  console.log(`kept-parts=${parts.length}`);
  console.log(`self-check dangling-references=${checks.danglingReferences}`);
  console.log(`self-check metadata=${checks.metadata}`);
  console.log(`self-check fonts=${checks.fonts}`);
  for (const part of parts) {
    const info = await stat(join(work, part));
    console.log(`${String(info.size).padStart(9, " ")}  ${part}`);
  }
} finally {
  await rm(work, { recursive: true, force: true });
}
