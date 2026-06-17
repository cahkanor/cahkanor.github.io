#!/usr/bin/env node

const fs = require("fs");
const path = require("path");
const cp = require("child_process");

const rootDir = path.resolve(__dirname, "..");
const sourceXlsx = path.join(rootDir, "data", "derived", "faculty_year_attribution_report.xlsx");
const outputXlsx = path.join(
  rootDir,
  "data",
  "derived",
  "faculty_year_attribution_report_pivot.xlsx"
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

function buildSheetXmlFromRows(headers, rows, options = {}) {
  const headerRows = options.headerRows ?? [headers];
  const mergeRefs = options.mergeRefs ?? [];
  const allRows = [...headerRows, ...rows.map((row) => headers.map((header) => row[header] ?? ""))];
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

function buildPivotSummaryRows(summaryRows) {
  const years = [...new Set(summaryRows.map((row) => normalize(row.year)).filter(Boolean))].sort(
    (a, b) => Number(a) - Number(b)
  );
  const facultyMap = new Map();

  for (const row of summaryRows) {
    const facultyName = normalize(row.faculty_name);
    const facultyCode = normalize(row.faculty_code);
    const year = normalize(row.year);
    if (!facultyMap.has(facultyName)) {
      facultyMap.set(facultyName, {
        faculty_name: facultyName,
        faculty_code: facultyCode,
      });
    }
    const target = facultyMap.get(facultyName);
    target[`${year}__first`] = normalize(row.first_author_papers) || "0";
    target[`${year}__corresponding`] = normalize(row.corresponding_author_papers) || "0";
    target[`${year}__coauthor`] = normalize(row.co_author_papers) || "0";
    target[`${year}__total`] = normalize(row.total_papers) || "0";
  }

  const headers = ["faculty_name"];
  const headerRow1 = ["Faculty Name"];
  const headerRow2 = [""];
  const mergeRefs = ["A1:A2"];
  let currentColumn = 2;
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

  const rows = [...facultyMap.values()]
    .sort((a, b) => a.faculty_name.localeCompare(b.faculty_name))
    .map((row) => {
      const out = { faculty_name: row.faculty_name };
      for (const year of years) {
        out[`${year} First Author`] = row[`${year}__first`] ?? "0";
        out[`${year} Corresponding Author`] = row[`${year}__corresponding`] ?? "0";
        out[`${year} Co-author`] = row[`${year}__coauthor`] ?? "0";
        out[`${year} Total`] = row[`${year}__total`] ?? "0";
      }
      return out;
    });

  return { headers, rows, headerRows: [headerRow1, headerRow2], mergeRefs };
}

function writeWorkbook(outputPath, sheets) {
  const buildDir = path.join(path.dirname(outputPath), "tmp_faculty_year_attribution_pivot_xlsx");
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
  const tmpCopy = path.join(path.dirname(outputXlsx), "tmp_attr_source_copy.xlsx");
  const tmpDir = path.join(path.dirname(outputXlsx), "tmp_attr_source_dir");
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
    const sheets = getWorksheetPaths(tmpDir);
    const summarySheet = sheets.find((sheet) => sheet.name === "Summary");
    const assignmentsSheet = sheets.find((sheet) => sheet.name === "Assignments");
    if (!summarySheet || !assignmentsSheet) {
      throw new Error("Expected Summary and Assignments sheets in source workbook.");
    }

    const summaryRows = parseWorksheetRows(summarySheet.path, sharedStrings);
    const assignmentsRows = parseWorksheetRows(assignmentsSheet.path, sharedStrings);
    const pivot = buildPivotSummaryRows(summaryRows);

    const assignmentsHeaders = assignmentsRows.length
      ? Object.keys(assignmentsRows[0])
      : [
          "year",
          "faculty_code",
          "faculty_name",
          "category",
          "matched_lecturer_name",
          "matched_scopus_author_id",
          "assignment_reason",
          "title",
          "eid",
          "doi",
          "first_author_ids",
          "corresponding_author_ids",
          "all_author_ids",
        ];

    writeWorkbook(outputXlsx, [
      {
        name: "Summary",
        xml: buildSheetXmlFromRows(pivot.headers, pivot.rows, {
          headerRows: pivot.headerRows,
          mergeRefs: pivot.mergeRefs,
        }),
      },
      {
        name: "Assignments",
        xml: buildSheetXmlFromRows(assignmentsHeaders, assignmentsRows),
      },
    ]);

    console.log(`Created ${outputXlsx}`);
  } finally {
    removeIfExists(tmpCopy);
    removeIfExists(tmpDir);
  }
}

main();
