// Node-only XLSX import + fs + codepages
import * as XLSX from "xlsx/xlsx.mjs";
import * as fs from "fs";
import * as path from "path";
import * as cpexcel from "xlsx/dist/cpexcel.full.mjs";

// NodeJS မှာ readFile / writeFile အလုပ်လုပ်စေ
XLSX.set_fs(fs);
XLSX.set_cptable(cpexcel);

const SRC_XLS = path.resolve("data/source.xls");
const OUT_JSON = path.resolve("src/data/data.json");

if (!fs.existsSync(SRC_XLS)) {
  console.error(`[convert] Missing ${SRC_XLS}`);
  process.exit(1);
}

const wb = XLSX.readFile(SRC_XLS, { cellDates: false });
const sheetName = wb.SheetNames[0];
const ws = wb.Sheets[sheetName];

const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

// Header စစ်
const headerWords = new Set(["name","east","easting","north","northing","height","elevation","h"]);
const looksLikeHeader = (r) => [0,1,2,3].some(i => headerWords.has(String(r?.[i]||"").trim().toLowerCase()));
const start = aoa.length && looksLikeHeader(aoa[0]) ? 1 : 0;

const out = [];
for (let i = start; i < aoa.length; i++) {
  const r = aoa[i] || [];
  const name   = String(r[0] ?? "").trim();
  const east   = String(r[1] ?? "").trim();
  const north  = String(r[2] ?? "").trim();
  const height = String(r[3] ?? "").trim();
  if (!name) continue;
  out.push({ name, east, north, height });
}

fs.mkdirSync(path.dirname(OUT_JSON), { recursive: true });
fs.writeFileSync(OUT_JSON, JSON.stringify(out, null, 2));
console.log(`[convert] Wrote ${out.length} rows → ${OUT_JSON}`);
