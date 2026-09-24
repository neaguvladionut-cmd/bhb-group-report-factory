// Carried over from bhb-platform master 40-standalone/group-report-factory/tests/launch.test.mjs (37f9a7e), test 3 of 4.
// 72655b1 rewrote launch.test.mjs without it. Only change: the helper-assignment regex tolerates the redesign's spaced source style (\s*=\s*).
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import test from "node:test";

const root=resolve(dirname(fileURLToPath(import.meta.url)),"..");
test("deploy provides a visible download recovery link",async()=>{const index=await readFile(resolve(root,"deploy/index.html"),"utf8"),app=await readFile(resolve(root,"deploy/app.js"),"utf8"),pptx=await readFile(resolve(root,"deploy/pptx.js"),"utf8");assert.match(index,/id="download-fallback"/u,"the action area includes a manual download fallback");assert.match(app,/reportDownloadError/u,"export errors are surfaced in the action area");assert.match(app,/window\.__grfDownload\s*=\s*download/u,"PPTX and XLSX share the same browser download helper");assert.match(app,/link\.click\(\)/u,"the attached fallback link is triggered from the button action");assert.match(pptx,/window\.__grfDownload/u,"PPTX uses the shared download helper when the app is running");});
