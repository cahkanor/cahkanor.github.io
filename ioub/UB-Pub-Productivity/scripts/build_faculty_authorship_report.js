#!/usr/bin/env node

const fs = require("fs");
const path = require("path");
const cp = require("child_process");

const rootDir = path.resolve(__dirname, "..");
const lecturersCsv = path.join(rootDir, "data", "master", "lecturers.csv");
const facultiesCsv = path.join(rootDir, "data", "master", "faculties.csv");
const publicationsXlsx = path.join(
  rootDir,
  "data",
  "master",
  "Publications_at_Brawijaya_University.xlsx"
);
const outputXlsx = path.join(rootDir, "data", "derived", "faculty_authorship_report.xlsx");

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

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const next = text[i + 1];

    if (ch === '"') {
      if (inQuotes && next === '"') {
        cell += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (ch === "," && !inQuotes) {
      row.push(cell);
      cell = "";
      continue;
    }

    if ((ch === "\n" || ch === "\r") && !inQuotes) {
      if (ch === "\r" && next === "\n") i++;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
      continue;
    }

    cell += ch;
  }

  if (cell.length || row.length) {
    row.push(cell);
    rows.push(row);
  }

  if (!rows.length) return [];
  const headers = rows[0].map((h) => normalize(h));
  return rows.slice(1).filter((r) => r.some((v) => normalize(v))).map((r) => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = r[index] ?? "";
    });
    return obj;
  });
}

function getScopusIds(value) {
  const matches = normalize(value).match(/\d{6,}/g) || [];
  return [...new Set(matches)];
}

function extractSharedStrings(sharedStringsXml) {
  if (!fs.existsSync(sharedStringsXml)) return [];
  const xml = fs.readFileSync(sharedStringsXml, "utf8");
  const values = [];
  const siRegex = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
  let match;
  while ((match = siRegex.exec(xml))) {
    const block = match[1];
    const texts = [];
    const tRegex = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g;
    let textMatch;
    while ((textMatch = tRegex.exec(block))) {
      texts.push(decodeXml(textMatch[1]));
    }
    values.push(texts.join(""));
  }
  return values;
}

function getFirstWorksheetPath(extractedDir) {
  const workbookXml = fs.readFileSync(path.join(extractedDir, "xl", "workbook.xml"), "utf8");
  const relsXml = fs.readFileSync(path.join(extractedDir, "xl", "_rels", "workbook.xml.rels"), "utf8");
  const sheetMatch = workbookXml.match(/<sheet\b[^>]*r:id="([^"]+)"/);
  if (!sheetMatch) throw new Error("Could not find first worksheet relationship id.");
  const relId = sheetMatch[1];
  const relRegex = new RegExp(`<Relationship\\b[^>]*Id="${relId}"[^>]*Target="([^"]+)"`);
  const relMatch = relsXml.match(relRegex);
  if (!relMatch) throw new Error("Could not resolve first worksheet path.");
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

function parseCellValue(cellType, innerXml, sharedStrings) {
  if (cellType === "s") {
    const valueMatch = innerXml.match(/<v>([\s\S]*?)<\/v>/);
    const idx = valueMatch ? Number(valueMatch[1]) : -1;
    return idx >= 0 && idx < sharedStrings.length ? sharedStrings[idx] : "";
  }
  if (cellType === "inlineStr") {
    const texts = [...innerXml.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((m) =>
      decodeXml(m[1])
    );
    return texts.join("");
  }
  const valueMatch = innerXml.match(/<v>([\s\S]*?)<\/v>/);
  return valueMatch ? decodeXml(valueMatch[1]) : "";
}

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function removeIfExists(targetPath) {
  if (fs.existsSync(targetPath)) {
    fs.rmSync(targetPath, { recursive: true, force: true });
  }
}

function writeMinimalXlsx(outputPath, headers, rows) {
  const buildDir = path.join(path.dirname(outputPath), "tmp_faculty_authorship_xlsx");
  removeIfExists(buildDir);
  ensureDir(path.join(buildDir, "_rels"));
  ensureDir(path.join(buildDir, "xl", "_rels"));
  ensureDir(path.join(buildDir, "xl", "worksheets"));

  const allRows = [headers, ...rows.map((row) => headers.map((header) => row[header] ?? ""))];
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

  fs.writeFileSync(
    path.join(buildDir, "[Content_Types].xml"),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
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
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
    "utf8"
  );

  fs.writeFileSync(
    path.join(buildDir, "xl", "_rels", "workbook.xml.rels"),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
    "utf8"
  );

  fs.writeFileSync(
    path.join(buildDir, "xl", "worksheets", "sheet1.xml"),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="${dimension}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>${rowXml}</sheetData>
</worksheet>`,
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
  ensureDir(path.dirname(outputXlsx));

  const lecturers = parseCsv(fs.readFileSync(lecturersCsv, "utf8"));
  const faculties = parseCsv(fs.readFileSync(facultiesCsv, "utf8"));

  const facultyByScopusId = new Map();
  for (const lecturer of lecturers) {
    const scopusId = normalize(lecturer.scopus_author_id);
    const facultyCode = normalize(lecturer.faculty_code);
    const facultyName = normalize(lecturer.faculty_name);
    if (!scopusId || !facultyCode) continue;
    if (!facultyByScopusId.has(scopusId)) facultyByScopusId.set(scopusId, new Map());
    facultyByScopusId.get(scopusId).set(facultyCode, facultyName);
  }

  const summary = new Map();
  for (const faculty of faculties) {
    const code = normalize(faculty.faculty_code);
    const name = normalize(faculty.faculty_name);
    summary.set(code, {
      faculty_code: code,
      faculty_name: name,
      paper_count_first_author: 0,
      paper_count_corresponding_author: 0,
      paper_count_first_or_corresponding_author: 0,
      paper_count_both_first_and_corresponding_author: 0,
    });
  }

  const tempCopy = path.join(path.dirname(outputXlsx), "tmp_publications_working_copy.xlsx");
  const tempDir = path.join(path.dirname(outputXlsx), "tmp_publications_working_dir");
  removeIfExists(tempCopy);
  removeIfExists(tempDir);
  fs.copyFileSync(publicationsXlsx, tempCopy);

  const escapedCopy = tempCopy.replace(/'/g, "''");
  const escapedTempDir = tempDir.replace(/'/g, "''");
  cp.execFileSync("powershell", [
    "-Command",
    `Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::ExtractToDirectory('${escapedCopy}', '${escapedTempDir}')`,
  ]);

  try {
    const sharedStrings = extractSharedStrings(path.join(tempDir, "xl", "sharedStrings.xml"));
    const worksheetPath = getFirstWorksheetPath(tempDir);
    const sheetXml = fs.readFileSync(worksheetPath, "utf8");

    const rowRegex = /<row\b[^>]*>([\s\S]*?)<\/row>/g;
    const rows = [...sheetXml.matchAll(rowRegex)];
    if (!rows.length) throw new Error("No worksheet rows found in publication workbook.");

    const headers = new Map();
    const cellRegex = /<c\b([^>]*)>([\s\S]*?)<\/c>/g;
    for (const cellMatch of rows[0][1].matchAll(cellRegex)) {
      const attrs = cellMatch[1];
      const innerXml = cellMatch[2];
      const refMatch = attrs.match(/\br="([^"]+)"/);
      const typeMatch = attrs.match(/\bt="([^"]+)"/);
      if (!refMatch) continue;
      headers.set(
        parseCellValue(typeMatch ? typeMatch[1] : "", innerXml, sharedStrings),
        columnIndex(refMatch[1])
      );
    }

    const required = [
      "Title",
      "Year",
      "DOI",
      "EID",
      "Scopus Author ID First Author",
      "Scopus Author ID Corresponding Author",
    ];
    for (const name of required) {
      if (!headers.has(name)) throw new Error(`Missing required header: ${name}`);
    }

    const neededColumns = new Set(required.map((name) => headers.get(name)));
    for (let i = 1; i < rows.length; i++) {
      const cellMap = new Map();
      for (const cellMatch of rows[i][1].matchAll(cellRegex)) {
        const attrs = cellMatch[1];
        const innerXml = cellMatch[2];
        const refMatch = attrs.match(/\br="([^"]+)"/);
        const typeMatch = attrs.match(/\bt="([^"]+)"/);
        if (!refMatch) continue;
        const col = columnIndex(refMatch[1]);
        if (!neededColumns.has(col)) continue;
        cellMap.set(col, parseCellValue(typeMatch ? typeMatch[1] : "", innerXml, sharedStrings));
      }

      const firstIds = getScopusIds(cellMap.get(headers.get("Scopus Author ID First Author")));
      const corrIds = getScopusIds(cellMap.get(headers.get("Scopus Author ID Corresponding Author")));

      const firstFaculties = new Map();
      for (const id of firstIds) {
        const map = facultyByScopusId.get(id);
        if (!map) continue;
        for (const [code, name] of map.entries()) firstFaculties.set(code, name);
      }

      const corrFaculties = new Map();
      for (const id of corrIds) {
        const map = facultyByScopusId.get(id);
        if (!map) continue;
        for (const [code, name] of map.entries()) corrFaculties.set(code, name);
      }

      for (const [code, name] of firstFaculties.entries()) {
        if (!summary.has(code)) {
          summary.set(code, {
            faculty_code: code,
            faculty_name: name,
            paper_count_first_author: 0,
            paper_count_corresponding_author: 0,
            paper_count_first_or_corresponding_author: 0,
            paper_count_both_first_and_corresponding_author: 0,
          });
        }
        summary.get(code).paper_count_first_author += 1;
      }

      for (const [code, name] of corrFaculties.entries()) {
        if (!summary.has(code)) {
          summary.set(code, {
            faculty_code: code,
            faculty_name: name,
            paper_count_first_author: 0,
            paper_count_corresponding_author: 0,
            paper_count_first_or_corresponding_author: 0,
            paper_count_both_first_and_corresponding_author: 0,
          });
        }
        summary.get(code).paper_count_corresponding_author += 1;
      }

      const eitherFaculties = new Map([...firstFaculties, ...corrFaculties]);
      for (const [code, name] of eitherFaculties.entries()) {
        if (!summary.has(code)) {
          summary.set(code, {
            faculty_code: code,
            faculty_name: name,
            paper_count_first_author: 0,
            paper_count_corresponding_author: 0,
            paper_count_first_or_corresponding_author: 0,
            paper_count_both_first_and_corresponding_author: 0,
          });
        }
        summary.get(code).paper_count_first_or_corresponding_author += 1;
        if (firstFaculties.has(code) && corrFaculties.has(code)) {
          summary.get(code).paper_count_both_first_and_corresponding_author += 1;
        }
      }
    }

    const headersOut = [
      "faculty_code",
      "faculty_name",
      "paper_count_first_author",
      "paper_count_corresponding_author",
      "paper_count_first_or_corresponding_author",
      "paper_count_both_first_and_corresponding_author",
    ];
    const summaryRows = [...summary.values()].sort((a, b) =>
      a.faculty_code.localeCompare(b.faculty_code)
    );

    writeMinimalXlsx(outputXlsx, headersOut, summaryRows);
    console.log(`Created ${outputXlsx}`);
  } finally {
    removeIfExists(tempCopy);
    removeIfExists(tempDir);
  }
}

main();
