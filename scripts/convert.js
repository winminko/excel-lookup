// Merge all .xls/.xlsx in /data → src/data/data.json
import * as XLSX from "xlsx/xlsx.mjs";
import * as fs from "fs";
import * as path from "path";
import * as cpexcel from "xlsx/dist/cpexcel.full.mjs";

XLSX.set_fs(fs);
XLSX.set_cptable(cpexcel);

const DATA_DIR = path.resolve("data");
const OUT_JSON = path.resolve("src/data/data.json");

// header detection (optional)
const headerWords = new Set(["name", "east", "easting", "north", "northing", "height", "elevation", "h"]);
const looksLikeHeader = (r) => {
  const s = (v) => String(v || "").trim().toLowerCase();
  return [0, 1, 2, 3].some((i) => headerWords.has(s(r?.[i])));
};

function readOne(filePath) {
  const wb = XLSX.readFile(filePath, { cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]]; // first sheet
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  const start = aoa.length && looksLikeHeader(aoa[0]) ? 1 : 0;

  const rows = [];
  for (let i = start; i < aoa.length; i++) {
    const r = aoa[i] || [];
    const name   = String(r[0] ?? "").trim();   // A = Name
    const east   = String(r[1] ?? "").trim();   // B = East
    const north  = String(r[2] ?? "").trim();   // C = North
    const height = String(r[3] ?? "").trim();   // D = Height
    if (!name) continue;
    rows.push({ name, east, north, height, __source: path.basename(filePath) });
  }
  return rows;
}

if (!fs.existsSync(DATA_DIR)) {
  console.error(`[convert] Missing folder: ${DATA_DIR}`);
  process.exit(1);
}

const files = fs
  .readdirSync(DATA_DIR)
  .filter((f) => /\.xlsx?$/i.test(f))
  .map((f) => path.join(DATA_DIR, f))
  .sort(); // ensure deterministic order

if (files.length === 0) {
  console.error("[convert] No .xls/.xlsx files in /data");
  process.exit(1);
}

let merged = [];
for (const f of files) merged = merged.concat(readOne(f));

// Dedupe by Name → keep the LAST occurrence (later file overrides earlier)
const map = new Map();
for (const row of merged) map.set(row.name, row);
const result = Array.from(map.values()).map(({ __source, ...rest }) => rest);

// Write JSON
fs.mkdirSync(path.dirname(OUT_JSON), { recursive: true });
fs.writeFileSync(OUT_JSON, JSON.stringify(result, null, 2));
console.log(`[convert] Merged ${files.length} files, ${result.length} rows → ${OUT_JSON}`);
