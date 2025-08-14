// Convert XLS -> JSON (A=Name, B=East, C=North, D=Height)
// Writes: src/data/data.json
import fs from "fs";
import path from "path";
import * as XLSX from "xlsx";

const SRC_XLS = path.resolve("data/source.xls");
const OUT_JSON = path.resolve("src/data/data.json");

if (!fs.existsSync(SRC_XLS)) {
  console.error(`[convert] Missing ${SRC_XLS}. Please add your .xls as data/source.xls`);
  process.exit(1);
}

const wb = XLSX.readFile(SRC_XLS, { cellDates: false });
const sheetName = wb.SheetNames[0];
const ws = wb.Sheets[sheetName];

// Get all rows as 2D array (no headers)
const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

// Heuristic: skip header row if it looks like labels
function looksLikeHeader(r) {
  const s = (v) => String(v || "").trim().toLowerCase();
  const keys = new Set(["name", "east", "easting", "north", "northing", "height", "elevation", "h"]);
  return [0,1,2,3].some((i) => keys.has(s(r?.[i])));
}

const start = aoa.length && looksLikeHeader(aoa[0]) ? 1 : 0;

const out = [];
for (let i = start; i < aoa.length; i++) {
  const row = aoa[i] || [];
  const name   = String(row[0] ?? "").trim();
  const east   = String(row[1] ?? "").trim();
  const north  = String(row[2] ?? "").trim();
  const height = String(row[3] ?? "").trim();
  if (!name) continue;
  out.push({ name, east, north, height });
}

// Ensure folder exists
fs.mkdirSync(path.dirname(OUT_JSON), { recursive: true });
fs.writeFileSync(OUT_JSON, JSON.stringify(out, null, 2), "utf-8");
console.log(`[convert] Wrote ${out.length} rows to ${OUT_JSON}`);
                              
