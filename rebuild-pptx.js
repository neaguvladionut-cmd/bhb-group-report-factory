import { downloadBundle, downloadTrendPptx } from "./template-pptx.js";

export async function downloadPptx(payload, filename, options = {}) {
  return downloadTrendPptx(payload, filename, options);
}

export { downloadBundle };
