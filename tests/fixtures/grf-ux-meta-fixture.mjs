/* Synthetic GRF-UX methodology fixture. No client data. */
import { createFixture } from "./grf-r-acceptance-fixture.mjs";

const serial = (year, month, day, fraction = 0) => (Date.UTC(year, month - 1, day) - Date.UTC(1899, 11, 30)) / 86400000 + fraction;
const meta = (index) => ({
  job: index < 12 ? "Manager vânzări" : "Specialist suport",
  location: "Hotel Sintetic, Brașov",
  invitedAt: [serial(2026, 3, 10, 0.4), serial(2026, 3, 11, 0.6), serial(2026, 3, 12, 0.35)][index % 3],
  principal: index % 2 ? "Ana Pop" : "Ion Ionescu",
  secondary: index % 2 ? "Maria Dan" : "Radu Vlad",
  system: "Evaluator Sistem"
});

const workbook = (XLSX, rows) => {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), "Sheet1");
  return XLSX.write(book, { type: "buffer", bookType: "xlsx" });
};

export function createMetaFixture(XLSX, variant = "d") {
  const base = createFixture(XLSX);
  const summaryRows = XLSX.utils.sheet_to_json(XLSX.read(base.summary, { type: "buffer" }).Sheets.Sheet1, { header: 1, raw: true, defval: "" });
  const detailedRows = XLSX.utils.sheet_to_json(XLSX.read(base.detailed, { type: "buffer" }).Sheets.Sheet1, { header: 1, raw: true, defval: "" });
  const summary = variant === "s"
    ? [
      ["CODE", "name", "job", "email", "certification location", "invited at", "principal evaluator", "secondary evaluator", "evaluator 3", "cod cp", ...summaryRows[0].slice(3)],
      ...summaryRows.slice(1).map((row, index) => { const item = meta(index); return [row[0], row[1], item.job, "", item.location, item.invitedAt, item.principal, item.secondary, item.system, row[2], ...row.slice(3)]; })
    ]
    : summaryRows;
  const detailed = variant === "d"
    ? detailedRows.map((row, index) => {
      if (index === 0) return [...row.slice(0, 4), "job", "date", "invited at", "certification location", "principal evaluator", "secondary evaluator", "evaluator 3", ...row.slice(4)];
      if (index < 3) return [...row.slice(0, 4), "", "", "", "", "", "", "", ...row.slice(4)];
      const item = meta(index - 3);
      return [...row.slice(0, 4), item.job, ["10.03.2026", "11.03.2026", "12.03.2026"][(index - 3) % 3], item.invitedAt, item.location, item.principal, item.secondary, item.system, ...row.slice(4)];
    })
    : detailedRows;
  return { summary: workbook(XLSX, summary), detailed: workbook(XLSX, detailed), metadata: base.metadata };
}
