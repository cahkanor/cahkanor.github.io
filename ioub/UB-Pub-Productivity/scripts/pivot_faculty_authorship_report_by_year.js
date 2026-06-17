#!/usr/bin/env node

const fs = require("fs");
const path = require("path");
const cp = require("child_process");

const rootDir = path.resolve(__dirname, "..");
const sourceXlsx = path.join(rootDir, "data", "derived", "faculty_authorship_report_by_year.xlsx");
const outputXlsx = path.join(
  rootDir,
  "data",
  "derived",
  "faculty_authorship_report_by_year_pivot.xlsx"
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

function getFirstWorksheetPath(extractedDir) {
  const workbookXml = fs.readFileSync(path.join(extractedDir, "xl", "workbook.xml"), "utf8");
  const relsXml = fs.readFileSync(path.join(extractedDir, "xl", "_rels", "workbook.xml.rels"), "utf8");
  const sheetMatch = workbookXml.match(/<sheet\b[^>]*r:id="([^"]+)"/);
  if (!sheetMatch) throw new Error("Could not find worksheet relationship id.");
  const relId = sheetMatch[1];
  const relRegex = new RegExp(`<Relationship\\b[^>]*Id="${relId}"[^>]*Target="([^"]+)"`);
  const relMatch = relsXml.match(relRegex);
  if (!relMatch) throw new Error("Could not resolve worksheet path.");
  return path.join(extractedDir, "xl", relMatch[1].replace(/\//g, path.sep));
}

function columnIndex(cellRef) {
  const letters = (cellRef.match(/^[A-Z]+/) || [""])[0];
  let index = 0;
  for (const ch of letters) {
    index = index * 26 + (ch.charCodeAt(0) - 64);
  }
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
    if (Object.values(row).some((value) => normalize(value))) {
      out.push(row);
    }
  }
  return out;
}

function buildPivotRows(sourceRows, metricName) {
  const years = [...new Set(sourceRows.map((row) => normalize(row.year)).filter(Boolean))].sort();
  const facultyMap = new Map();

  for (const row of sourceRows) {
    const facultyCode = normalize(row.faculty_code);
    const facultyName = normalize(row.faculty_name);
    const year = normalize(row.year);
    const metricValue = normalize(row[metricName]) || "0";
    if (!facultyMap.has(facultyCode)) {
      facultyMap.set(facultyCode, {
        faculty_code: facultyCode,
        faculty_name: facultyName,
      });
    }
    facultyMap.get(facultyCode)[year] = metricValue;
  }

  const rows = [...facultyMap.values()]
    .sort((a, b) => a.faculty_code.localeCompare(b.faculty_code))
    .map((row) => {
      const out = {
        faculty_code: row.faculty_code,
        faculty_name: row.faculty_name,
      };
      for (const year of years) {
        out[year] = row[year] ?? "0";
      }
      return out;
    });

  return {
    headers: ["faculty_code", "faculty_name", ...years],
    rows,
  };
}

function writeMultiSheetXlsx(outputPath, sheets) {
  const buildDir = path.join(path.dirname(outputPath), "tmp_faculty_authorship_pivot_xlsx");
  removeIfExists(buildDir);
  ensureDir(path.join(buildDir, "_rels"));
  ensureDir(path.join(buildDir, "xl", "_rels"));
  ensureDir(path.join(buildDir, "xl", "worksheets"));

  const worksheetOverrides = [];
  const sheetEntries = [];
  const sheetRels = [];

  sheets.forEach((sheet, idx) => {
    const sheetIndex = idx + 1;
    const allRows = [sheet.headers, ...sheet.rows.map((row) => sheet.headers.map((header) => row[header] ?? ""))];
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

    const dimension = `A1:${columnName(sheet.headers.length)}${allRows.length}`;
    fs.writeFileSync(
      path.join(buildDir, "xl", "worksheets", `sheet${sheetIndex}.xml`),
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="${dimension}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>${rowXml}</sheetData>
</worksheet>`,
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
  const tmpCopy = path.join(path.dirname(outputXlsx), "tmp_faculty_authorship_by_year_copy.xlsx");
  const tmpDir = path.join(path.dirname(outputXlsx), "tmp_faculty_authorship_by_year_dir");
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
    const worksheetPath = getFirstWorksheetPath(tmpDir);
    const sourceRows = parseWorksheetRows(worksheetPath, sharedStrings);

    const sheets = [
      {
        name: "FirstAuthor",
        ...buildPivotRows(sourceRows, "paper_count_first_author"),
      },
      {
        name: "CorrespondingAuthor",
        ...buildPivotRows(sourceRows, "paper_count_corresponding_author"),
      },
      {
        name: "FirstAndCorresponding",
        ...buildPivotRows(sourceRows, "paper_count_both_first_and_corresponding_author"),
      },
      {
        name: "FirstOrCorresponding",
        ...buildPivotRows(sourceRows, "paper_count_first_or_corresponding_author"),
      },
    ];

    writeMultiSheetXlsx(outputXlsx, sheets);
    console.log(`Created ${outputXlsx}`);
  } finally {
    removeIfExists(tmpCopy);
    removeIfExists(tmpDir);
  }
}

main();
