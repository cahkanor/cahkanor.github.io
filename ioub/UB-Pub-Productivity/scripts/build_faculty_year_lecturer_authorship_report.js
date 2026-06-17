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
const outputXlsx = path.join(
  rootDir,
  "data",
  "derived",
  "faculty_year_lecturer_authorship_report.xlsx"
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
  return rows
    .slice(1)
    .filter((r) => r.some((v) => normalize(v)))
    .map((r) => Object.fromEntries(headers.map((h, i) => [h, r[i] ?? ""])));
}

function getScopusIds(value) {
  return [...new Set((normalize(value).match(/\d{6,}/g) || []))];
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

function buildSheetXml(headers, rows) {
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
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="${dimension}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>${rowXml}</sheetData>
</worksheet>`;
}

function writeWorkbook(outputPath, sheets) {
  const buildDir = path.join(path.dirname(outputPath), "tmp_faculty_year_lecturer_authorship_xlsx");
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
  ensureDir(path.dirname(outputXlsx));

  const lecturers = parseCsv(fs.readFileSync(lecturersCsv, "utf8"));
  const faculties = parseCsv(fs.readFileSync(facultiesCsv, "utf8"));

  const lecturersByFaculty = new Map();
  const lecturerByScopusId = new Map();
  for (const lecturer of lecturers) {
    const facultyCode = normalize(lecturer.faculty_code);
    const facultyName = normalize(lecturer.faculty_name);
    const scopusId = normalize(lecturer.scopus_author_id);
    const lecturerRow = {
      faculty_code: facultyCode,
      faculty_name: facultyName,
      lecturer_id: normalize(lecturer.lecturer_id),
      lecturer_name: normalize(lecturer.lecturer_name),
      scopus_author_id: scopusId,
    };
    if (!lecturersByFaculty.has(facultyCode)) lecturersByFaculty.set(facultyCode, []);
    lecturersByFaculty.get(facultyCode).push(lecturerRow);
    if (scopusId && !lecturerByScopusId.has(scopusId)) {
      lecturerByScopusId.set(scopusId, lecturerRow);
    }
  }

  for (const rows of lecturersByFaculty.values()) {
    rows.sort((a, b) => a.lecturer_name.localeCompare(b.lecturer_name));
  }

  const facultyList = faculties
    .map((faculty) => ({
      faculty_code: normalize(faculty.faculty_code),
      faculty_name: normalize(faculty.faculty_name),
    }))
    .filter((faculty) => faculty.faculty_code);

  const tmpCopy = path.join(path.dirname(outputXlsx), "tmp_publications_lecturer_copy.xlsx");
  const tmpDir = path.join(path.dirname(outputXlsx), "tmp_publications_lecturer_dir");
  removeIfExists(tmpCopy);
  removeIfExists(tmpDir);
  fs.copyFileSync(publicationsXlsx, tmpCopy);

  const escapedCopy = tmpCopy.replace(/'/g, "''");
  const escapedTmpDir = tmpDir.replace(/'/g, "''");
  cp.execFileSync("powershell", [
    "-Command",
    `Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::ExtractToDirectory('${escapedCopy}', '${escapedTmpDir}')`,
  ]);

  try {
    const sharedStrings = extractSharedStrings(path.join(tmpDir, "xl", "sharedStrings.xml"));
    const worksheetPath = getFirstWorksheetPath(tmpDir);
    const publicationRows = parseWorksheetRows(worksheetPath, sharedStrings);

    const years = [...new Set(publicationRows.map((row) => normalize(row.Year)).filter((year) => /^\d{4}$/.test(year)))]
      .sort((a, b) => Number(a) - Number(b));

    const countsBySheet = new Map();
    function ensureCounter(facultyCode, year, lecturerRow) {
      const key = `${facultyCode}||${year}`;
      if (!countsBySheet.has(key)) countsBySheet.set(key, new Map());
      const rowMap = countsBySheet.get(key);
      const lecturerKey = lecturerRow.lecturer_id || lecturerRow.scopus_author_id || lecturerRow.lecturer_name;
      if (!rowMap.has(lecturerKey)) {
        rowMap.set(lecturerKey, {
          lecturer_id: lecturerRow.lecturer_id,
          lecturer_name: lecturerRow.lecturer_name,
          scopus_author_id: lecturerRow.scopus_author_id,
          first_author_papers: 0,
          corresponding_author_papers: 0,
          co_author_papers: 0,
          total_papers: 0,
        });
      }
      return rowMap.get(lecturerKey);
    }

    for (const faculty of facultyList) {
      const facultyLecturers = lecturersByFaculty.get(faculty.faculty_code) || [];
      for (const year of years) {
        for (const lecturerRow of facultyLecturers) {
          ensureCounter(faculty.faculty_code, year, lecturerRow);
        }
      }
    }

    for (const row of publicationRows) {
      const year = normalize(row.Year);
      if (!/^\d{4}$/.test(year)) continue;

      const firstAuthorIds = getScopusIds(row["Scopus Author ID First Author"]);
      const corrAuthorIds = getScopusIds(row["Scopus Author ID Corresponding Author"]);
      const singleAuthorIds = getScopusIds(row["Scopus Author ID Single Author"]);
      const allAuthorIds = getScopusIds(row["Scopus Author Ids"]);

      const lecturerIdsInPaper = [...new Set(allAuthorIds.filter((id) => lecturerByScopusId.has(id)))];
      for (const scopusId of lecturerIdsInPaper) {
        const lecturerRow = lecturerByScopusId.get(scopusId);
        const counter = ensureCounter(lecturerRow.faculty_code, year, lecturerRow);

        let category = "co_author";
        if (singleAuthorIds.includes(scopusId) || firstAuthorIds.includes(scopusId)) {
          category = "first_author";
        } else if (corrAuthorIds.includes(scopusId)) {
          category = "corresponding_author";
        }

        if (category === "first_author") counter.first_author_papers += 1;
        else if (category === "corresponding_author") counter.corresponding_author_papers += 1;
        else counter.co_author_papers += 1;

        counter.total_papers += 1;
      }
    }

    const sheetDefs = [];
    for (const faculty of facultyList) {
      const facultyLecturers = lecturersByFaculty.get(faculty.faculty_code) || [];
      for (const year of years) {
        const key = `${faculty.faculty_code}||${year}`;
        const counterMap = countsBySheet.get(key) || new Map();
        const rows = facultyLecturers.map((lecturerRow) => {
          const lecturerKey =
            lecturerRow.lecturer_id || lecturerRow.scopus_author_id || lecturerRow.lecturer_name;
          return counterMap.get(lecturerKey) || {
            lecturer_id: lecturerRow.lecturer_id,
            lecturer_name: lecturerRow.lecturer_name,
            scopus_author_id: lecturerRow.scopus_author_id,
            first_author_papers: 0,
            corresponding_author_papers: 0,
            co_author_papers: 0,
            total_papers: 0,
          };
        });

        sheetDefs.push({
          name: `${faculty.faculty_code}_${year}`.slice(0, 31),
          xml: buildSheetXml(
            [
              "lecturer_id",
              "lecturer_name",
              "scopus_author_id",
              "first_author_papers",
              "corresponding_author_papers",
              "co_author_papers",
              "total_papers",
            ],
            rows
          ),
        });
      }
    }

    writeWorkbook(outputXlsx, sheetDefs);
    console.log(`Created ${outputXlsx}`);
  } finally {
    removeIfExists(tmpCopy);
    removeIfExists(tmpDir);
  }
}

main();
