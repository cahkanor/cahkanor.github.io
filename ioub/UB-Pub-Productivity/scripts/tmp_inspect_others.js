const fs = require("fs");
const path = require("path");
const cp = require("child_process");

const root = path.resolve(__dirname, "..");
const xlsx = path.join(root, "data", "derived", "faculty_year_attribution_report.xlsx");
const tmpCopy = path.join(root, "data", "derived", "tmp_attr_inspect_copy.xlsx");
const tmpDir = path.join(root, "data", "derived", "tmp_attr_inspect_dir");

function normalize(value) {
  return String(value ?? "").trim();
}

function decode(value) {
  return normalize(value)
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function remove(target) {
  if (fs.existsSync(target)) {
    fs.rmSync(target, { recursive: true, force: true });
  }
}

function extractSharedStrings(sharedStringsPath) {
  if (!fs.existsSync(sharedStringsPath)) return [];
  const xml = fs.readFileSync(sharedStringsPath, "utf8");
  const out = [];
  for (const m of xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)) {
    const texts = [...m[1].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((x) =>
      decode(x[1])
    );
    out.push(texts.join(""));
  }
  return out;
}

function getSheets(extractedDir) {
  const workbook = fs.readFileSync(path.join(extractedDir, "xl", "workbook.xml"), "utf8");
  const rels = fs.readFileSync(path.join(extractedDir, "xl", "_rels", "workbook.xml.rels"), "utf8");
  const relMap = new Map(
    [...rels.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)].map((m) => [
      m[1],
      m[2],
    ])
  );
  return [...workbook.matchAll(/<sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)].map((m) => ({
    name: decode(m[1]),
    path: path.join(extractedDir, "xl", relMap.get(m[2]).replace(/\//g, path.sep)),
  }));
}

function colIndex(ref) {
  const letters = (ref.match(/^[A-Z]+/) || [""])[0];
  let idx = 0;
  for (const ch of letters) idx = idx * 26 + (ch.charCodeAt(0) - 64);
  return idx;
}

function parseCell(type, inner, shared) {
  if (type === "s") {
    const m = inner.match(/<v>([\s\S]*?)<\/v>/);
    const idx = m ? Number(m[1]) : -1;
    return idx >= 0 && idx < shared.length ? shared[idx] : "";
  }
  if (type === "inlineStr") {
    return [...inner.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((x) => decode(x[1])).join("");
  }
  const m = inner.match(/<v>([\s\S]*?)<\/v>/);
  return m ? decode(m[1]) : "";
}

function parseSheet(sheetPath, shared) {
  const xml = fs.readFileSync(sheetPath, "utf8");
  const rows = [...xml.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)];
  const cellRegex = /<c\b([^>]*)>([\s\S]*?)<\/c>/g;
  const headers = new Map();
  for (const c of rows[0][1].matchAll(cellRegex)) {
    const ref = (c[1].match(/\br="([^"]+)"/) || [])[1];
    const type = (c[1].match(/\bt="([^"]+)"/) || [])[1] || "";
    if (!ref) continue;
    headers.set(parseCell(type, c[2], shared), colIndex(ref));
  }
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const cellMap = new Map();
    for (const c of rows[i][1].matchAll(cellRegex)) {
      const ref = (c[1].match(/\br="([^"]+)"/) || [])[1];
      const type = (c[1].match(/\bt="([^"]+)"/) || [])[1] || "";
      if (!ref) continue;
      cellMap.set(colIndex(ref), parseCell(type, c[2], shared));
    }
    const row = {};
    for (const [h, col] of headers.entries()) row[h] = cellMap.get(col) ?? "";
    if (Object.values(row).some((v) => normalize(v))) out.push(row);
  }
  return out;
}

remove(tmpCopy);
remove(tmpDir);
fs.copyFileSync(xlsx, tmpCopy);
cp.execFileSync("powershell", [
  "-Command",
  `Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::ExtractToDirectory('${tmpCopy.replace(/'/g, "''")}','${tmpDir.replace(/'/g, "''")}')`,
]);

try {
  const shared = extractSharedStrings(path.join(tmpDir, "xl", "sharedStrings.xml"));
  const sheets = getSheets(tmpDir);
  const assignmentSheet = sheets.find((sheet) => sheet.name === "Assignments");
  const rows = parseSheet(assignmentSheet.path, shared);
  const others = rows.filter((row) => normalize(row.faculty_code) === "OTHERS");
  const byYear = {};
  const byReason = {};
  for (const row of others) {
    byYear[row.year] = (byYear[row.year] || 0) + 1;
    byReason[row.assignment_reason] = (byReason[row.assignment_reason] || 0) + 1;
  }
  const samples = others.slice(0, 12).map((row) => ({
    year: row.year,
    title: row.title,
    reason: row.assignment_reason,
    first_author_ids: row.first_author_ids,
    corresponding_author_ids: row.corresponding_author_ids,
  }));
  console.log(JSON.stringify({ count: others.length, byYear, byReason, samples }, null, 2));
} finally {
  remove(tmpCopy);
  remove(tmpDir);
}
