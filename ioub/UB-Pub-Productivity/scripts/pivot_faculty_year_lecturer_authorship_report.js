#!/usr/bin/env node

const fs = require("fs");
const path = require("path");
const cp = require("child_process");

const rootDir = path.resolve(__dirname, "..");
const sourceXlsx = path.join(
  rootDir,
  "data",
  "derived",
  "faculty_year_lecturer_authorship_report.xlsx"
);
const outputXlsx = path.join(
  rootDir,
  "data",
  "derived",
  "faculty_year_lecturer_authorship_report_pivot.xlsx"
);

function normalize(value) {
  return String(value ?? "").trim();
}

function escapeXml(value) {
  return normalize(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function decodeXml(value) {
  return normalize(value)
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function removeIfExists(targetPath) {
  if (fs.existsSync(targetPath)) {
    fs.rmSync(targetPath, { recursive: true, force: true });
  }
}

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function extractSharedStrings(sharedStringsPath) {
  if (!fs.existsSync(sharedStringsPath)) return [];
  const xml = fs.readFileSync(sharedStringsPath, "utf8");
  const values = [];
  for (const m of xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)) {
    const texts = [...m[1].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((x) =>
      decodeXml(x[1])
    );
    values.push(texts.join(""));
  }
  return values;
}

function getWorksheetPaths(extractedDir) {
  const workbookXml = fs.readFileSync(path.join(extractedDir, "xl", "workbook.xml"), "utf8");
  const relsXml = fs.readFileSync(path.join(extractedDir, "xl", "_rels", "workbook.xml.rels"), "utf8");

  const relMap = new Map();
  for (const rel of relsXml.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
    relMap.set(rel[1], rel[2]);
  }

  const sheets = [];
  for (const match of workbookXml.matchAll(/<sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)) {
    const name = decodeXml(match[1]);
    const relId = match[2];
    const target = relMap.get(relId);
    if (target) {
      sheets.push({
        name,
        path: path.join(extractedDir, "xl", target.replace(/\//g, path.sep)),
      });
    }
  }
  return sheets;
}

function columnIndex(cellRef) {
  const letters = (cellRef.match(/^[A-Z]+/) || [""])[0];
  let index = 0;
  for (const ch of letters) index = index * 26 + (ch.charCodeAt(0) - 64);
  return index;
}

function columnName(index) {
  let name = "";
  let current = index;
  while (current > 0) {
    current--;
    name = String.fromCharCode(65 + (current % 26)) + name;
    current = Math.floor(current / 26);
  }
  return name;
}

function parseCellValue(type, innerXml, sharedStrings) {
  if (type === "s") {
    const m = innerXml.match(/<v>([\s\S]*?)<\/v>/);
    const idx = m ? Number(m[1]) : -1;
    return idx >= 0 && idx < sharedStrings.length ? sharedStrings[idx] : "";
  }
  if (type === "inlineStr") {
    return [...innerXml.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((x) =>
      decodeXml(x[1])
    ).join("");
  }
  const m = innerXml.match(/<v>([\s\S]*?)<\/v>/);
  return m ? decodeXml(m[1]) : "";
}

function parseWorksheetRows(worksheetPath, sharedStrings) {
  const xml = fs.readFileSync(worksheetPath, "utf8");
  const rows = [...xml.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)];
  if (!rows.length) return [];

  const cellRegex = /<c\b([^>]*)>([\s\S]*?)<\/c>/g;
  const headers = new Map();
  for (const c of rows[0][1].matchAll(cellRegex)) {
    const attrs = c[1];
    const inner = c[2];
    const ref = (attrs.match(/\br="([^"]+)"/) || [])[1];
    const type = (attrs.match(/\bt="([^"]+)"/) || [])[1] || "";
    if (!ref) continue;
    headers.set(parseCellValue(type, inner, sharedStrings), columnIndex(ref));
  }

  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const cellMap = new Map();
    for (const c of rows[i][1].matchAll(cellRegex)) {
      const attrs = c[1];
      const inner = c[2];
      const ref = (attrs.match(/\br="([^"]+)"/) || [])[1];
      const type = (attrs.match(/\bt="([^"]+)"/) || [])[1] || "";
      if (!ref) continue;
      cellMap.set(columnIndex(ref), parseCellValue(type, inner, sharedStrings));
    }
    const row = {};
    for (const [header, col] of headers.entries()) {
      row[header] = cellMap.get(col) ?? "";
    }
    if (Object.values(row).some((value) => normalize(value))) out.push(row);
  }
  return out;
}

function buildGroupedSheetXml(headers, dataRows, headerRows, mergeRefs) {
  const allRows = [...headerRows, ...dataRows.map((row) => headers.map((header) => row[header] ?? ""))];
  const rowXml = allRows
    .map((cells, rowIdx) => {
      const cellXml = cells
        .map((value, colIdx) => {
          const ref = `${columnName(colIdx + 1)}${rowIdx + 1}`;
          const text = normalize(value);
          if (/^-?\d+(\.\d+)?$/.test(text)) {
            return `<c r="${ref}"><v>${text}</v></c>`;
          }
          return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(
            text
          )}</t></is></c>`;
        })
        .join("");
      return `<row r="${rowIdx + 1}">${cellXml}</row>`;
    })
    .join("");

  const dimension = `A1:${columnName(headers.length)}${allRows.length}`;
  const mergeXml = mergeRefs.length
    ? `<mergeCells count="${mergeRefs.length}">${mergeRefs
        .map((ref) => `<mergeCell ref="${ref}"/>`)
        .join("")}</mergeCells>`
    : "";

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="${dimension}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>${rowXml}</sheetData>
  ${mergeXml}
</worksheet>`;
}

function writeWorkbook(outputPath, sheets) {
  const buildDir = path.join(path.dirname(outputPath), "tmp_faculty_year_lecturer_pivot_xlsx");
  removeIfExists(buildDir);
  ensureDir(path.join(buildDir, "_rels"));
  ensureDir(path.join(buildDir, "xl", "_rels"));
  ensureDir(path.join(buildDir, "xl", "worksheets"));

  const worksheetOverrides = [];
  const sheetEntries = [];
  const sheetRels = [];

  sheets.forEach((sheet, idx) => {
    const sheetIndex = idx + 1;
    fs.writeFileSync(
      path.join(buildDir, "xl", "worksheets", `sheet${sheetIndex}.xml`),
      sheet.xml,
      "utf8"
    );
    worksheetOverrides.push(
      `<Override PartName="/xl/worksheets/sheet${sheetIndex}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    );
    sheetEntries.push(`<sheet name="${escapeXml(sheet.name)}" sheetId="${sheetIndex}" r:id="rId${sheetIndex}"/>`);
    sheetRels.push(
      `<Relationship Id="rId${sheetIndex}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${sheetIndex}.xml"/>`
    );
  });

  fs.writeFileSync(
    path.join(buildDir, "[Content_Types].xml"),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${worksheetOverrides.join("\n  ")}
</Types>`,
    "utf8"
  );

  fs.writeFileSync(
    path.join(buildDir, "_rels", ".rels"),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
    "utf8"
  );

  fs.writeFileSync(
    path.join(buildDir, "xl", "workbook.xml"),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    ${sheetEntries.join("\n    ")}
  </sheets>
</workbook>`,
    "utf8"
  );

  fs.writeFileSync(
    path.join(buildDir, "xl", "_rels", "workbook.xml.rels"),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${sheetRels.join("\n  ")}
</Relationships>`,
    "utf8"
  );

  removeIfExists(outputPath);
  const escapedBuildDir = buildDir.replace(/'/g, "''");
  const escapedOutput = outputPath.replace(/'/g, "''");
  cp.execFileSync("powershell", [
    "-Command",
    `Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::CreateFromDirectory('${escapedBuildDir}', '${escapedOutput}')`,
  ]);
  removeIfExists(buildDir);
}

function main() {
  if (!fs.existsSync(sourceXlsx)) {
    throw new Error(`Source workbook not found: ${sourceXlsx}`);
  }

  ensureDir(path.dirname(outputXlsx));
  const tmpCopy = path.join(path.dirname(outputXlsx), "tmp_lecturer_authorship_source_copy.xlsx");
  const tmpDir = path.join(path.dirname(outputXlsx), "tmp_lecturer_authorship_source_dir");
  removeIfExists(tmpCopy);
  removeIfExists(tmpDir);
  fs.copyFileSync(sourceXlsx, tmpCopy);

  const escapedCopy = tmpCopy.replace(/'/g, "''");
  const escapedTmpDir = tmpDir.replace(/'/g, "''");
  cp.execFileSync("powershell", [
    "-Command",
    `Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::ExtractToDirectory('${escapedCopy}', '${escapedTmpDir}')`,
  ]);

  try {
    const sharedStrings = extractSharedStrings(path.join(tmpDir, "xl", "sharedStrings.xml"));
    const worksheetDefs = getWorksheetPaths(tmpDir);

    const facultyMap = new Map();
    for (const sheet of worksheetDefs) {
      const match = sheet.name.match(/^(.+)_([0-9]{4})$/);
      if (!match) continue;
      const facultyCode = match[1];
      const year = match[2];
      const rows = parseWorksheetRows(sheet.path, sharedStrings);
      if (!facultyMap.has(facultyCode)) facultyMap.set(facultyCode, new Map());
      facultyMap.get(facultyCode).set(year, rows);
    }

    const sheetOutputs = [];
    const facultyCodes = [...facultyMap.keys()].sort((a, b) => a.localeCompare(b));
    for (const facultyCode of facultyCodes) {
      const yearMap = facultyMap.get(facultyCode);
      const years = [...yearMap.keys()].sort((a, b) => Number(a) - Number(b));
      const lecturerMap = new Map();

      for (const year of years) {
        for (const row of yearMap.get(year)) {
          const lecturerName = normalize(row.lecturer_name);
          const scopusAuthorId = normalize(row.scopus_author_id);
          const key = scopusAuthorId || lecturerName;
          if (!lecturerMap.has(key)) {
            lecturerMap.set(key, {
              lecturer_name: lecturerName,
              scopus_author_id: scopusAuthorId,
            });
          }
          const target = lecturerMap.get(key);
          target[`${year}__first`] = normalize(row.first_author_papers) || "0";
          target[`${year}__corr`] = normalize(row.corresponding_author_papers) || "0";
          target[`${year}__co`] = normalize(row.co_author_papers) || "0";
          target[`${year}__total`] = normalize(row.total_papers) || "0";
        }
      }

      const headers = ["lecturer_name", "scopus_author_id"];
      const headerRow1 = ["Lecturer Name", "Scopus Author ID"];
      const headerRow2 = ["", ""];
      const mergeRefs = ["A1:A2", "B1:B2"];
      let currentColumn = 3;
      for (const year of years) {
        headerRow1.push(year, "", "", "");
        headerRow2.push("First Author", "Corresponding Author", "Co-author", "Total");
        headers.push(`${year} First Author`);
        headers.push(`${year} Corresponding Author`);
        headers.push(`${year} Co-author`);
        headers.push(`${year} Total`);
        mergeRefs.push(`${columnName(currentColumn)}1:${columnName(currentColumn + 3)}1`);
        currentColumn += 4;
      }

      const dataRows = [...lecturerMap.values()]
        .sort((a, b) => a.lecturer_name.localeCompare(b.lecturer_name))
        .map((row) => {
          const out = {
            lecturer_name: row.lecturer_name,
            scopus_author_id: row.scopus_author_id,
          };
          for (const year of years) {
            out[`${year} First Author`] = row[`${year}__first`] ?? "0";
            out[`${year} Corresponding Author`] = row[`${year}__corr`] ?? "0";
            out[`${year} Co-author`] = row[`${year}__co`] ?? "0";
            out[`${year} Total`] = row[`${year}__total`] ?? "0";
          }
          return out;
        });

      sheetOutputs.push({
        name: facultyCode.slice(0, 31),
        xml: buildGroupedSheetXml(headers, dataRows, [headerRow1, headerRow2], mergeRefs),
      });
    }

    writeWorkbook(outputXlsx, sheetOutputs);
    console.log(`Created ${outputXlsx}`);
  } finally {
    removeIfExists(tmpCopy);
    removeIfExists(tmpDir);
  }
}

main();
