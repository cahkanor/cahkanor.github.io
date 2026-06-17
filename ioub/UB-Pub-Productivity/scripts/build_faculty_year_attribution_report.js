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
  "faculty_year_attribution_report.xlsx"
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
  return rows.slice(1).filter((r) => r.some((v) => normalize(v))).map((r) => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = r[index] ?? "";
    });
    return obj;
  });
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

function buildWorkbook(outputPath, sheets) {
  const buildDir = path.join(path.dirname(outputPath), "tmp_faculty_year_attribution_xlsx");
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
  ensureDir(path.dirname(outputXlsx));

  const lecturers = parseCsv(fs.readFileSync(lecturersCsv, "utf8"));
  const faculties = parseCsv(fs.readFileSync(facultiesCsv, "utf8"));

  const facultyByScopusId = new Map();
  for (const lecturer of lecturers) {
    const id = normalize(lecturer.scopus_author_id);
    if (!id || facultyByScopusId.has(id)) continue;
    facultyByScopusId.set(id, {
      faculty_code: normalize(lecturer.faculty_code),
      faculty_name: normalize(lecturer.faculty_name),
      lecturer_name: normalize(lecturer.lecturer_name),
    });
  }

  const facultyList = faculties
    .map((faculty) => ({
      faculty_code: normalize(faculty.faculty_code),
      faculty_name: normalize(faculty.faculty_name),
    }))
    .filter((faculty) => faculty.faculty_code);

  facultyList.push({ faculty_code: "OTHERS", faculty_name: "Others" });

  const tmpCopy = path.join(path.dirname(outputXlsx), "tmp_publications_attr_copy.xlsx");
  const tmpDir = path.join(path.dirname(outputXlsx), "tmp_publications_attr_dir");
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
      .sort();

    const summaryMap = new Map();
    const assignmentRows = [];

    function summaryKey(year, facultyCode) {
      return `${year}||${facultyCode}`;
    }

    function ensureSummary(year, facultyCode, facultyName) {
      const key = summaryKey(year, facultyCode);
      if (!summaryMap.has(key)) {
        summaryMap.set(key, {
          year,
          faculty_code: facultyCode,
          faculty_name: facultyName,
          first_author_papers: 0,
          corresponding_author_papers: 0,
          co_author_papers: 0,
          total_papers: 0,
        });
      }
      return summaryMap.get(key);
    }

    for (const faculty of facultyList) {
      for (const year of years) {
        ensureSummary(year, faculty.faculty_code, faculty.faculty_name);
      }
    }

    for (const row of publicationRows) {
      const year = normalize(row.Year);
      if (!/^\d{4}$/.test(year)) continue;

      const title = normalize(row.Title);
      const eid = normalize(row.EID);
      const doi = normalize(row.DOI);
      const firstAuthorIds = getScopusIds(row["Scopus Author ID First Author"]);
      const corrAuthorIds = getScopusIds(row["Scopus Author ID Corresponding Author"]);
      const allAuthorIds = getScopusIds(row["Scopus Author Ids"]);

      let assignedFaculty = null;
      let category = "";
      let assignmentReason = "";
      let matchedAuthorId = "";
      let matchedLecturerName = "";

      for (const id of firstAuthorIds) {
        if (facultyByScopusId.has(id)) {
          assignedFaculty = facultyByScopusId.get(id);
          category = "first_author";
          assignmentReason = "UB first author";
          matchedAuthorId = id;
          matchedLecturerName = assignedFaculty.lecturer_name;
          break;
        }
      }

      if (!assignedFaculty) {
        for (const id of corrAuthorIds) {
          if (facultyByScopusId.has(id)) {
            assignedFaculty = facultyByScopusId.get(id);
            category = "corresponding_author";
            assignmentReason = "UB corresponding author";
            matchedAuthorId = id;
            matchedLecturerName = assignedFaculty.lecturer_name;
            break;
          }
        }
      }

      if (!assignedFaculty) {
        for (const id of allAuthorIds) {
          if (facultyByScopusId.has(id)) {
            assignedFaculty = facultyByScopusId.get(id);
            category = "co_author";
            assignmentReason = "First UB-affiliated author in author order";
            matchedAuthorId = id;
            matchedLecturerName = assignedFaculty.lecturer_name;
            break;
          }
        }
      }

      if (!assignedFaculty) {
        assignedFaculty = {
          faculty_code: "OTHERS",
          faculty_name: "Others",
          lecturer_name: "",
        };
        category = "co_author";
        assignmentReason = "No UB-affiliated author found";
      }

      const summaryRow = ensureSummary(year, assignedFaculty.faculty_code, assignedFaculty.faculty_name);
      summaryRow.total_papers += 1;
      if (category === "first_author") summaryRow.first_author_papers += 1;
      if (category === "corresponding_author") summaryRow.corresponding_author_papers += 1;
      if (category === "co_author") summaryRow.co_author_papers += 1;

      assignmentRows.push({
        year,
        faculty_code: assignedFaculty.faculty_code,
        faculty_name: assignedFaculty.faculty_name,
        category,
        matched_lecturer_name: matchedLecturerName,
        matched_scopus_author_id: matchedAuthorId,
        assignment_reason: assignmentReason,
        title,
        eid,
        doi,
        first_author_ids: firstAuthorIds.join("; "),
        corresponding_author_ids: corrAuthorIds.join("; "),
        all_author_ids: allAuthorIds.join("; "),
      });
    }

    const summaryRows = [...summaryMap.values()].sort((a, b) => {
      if (a.year !== b.year) return Number(a.year) - Number(b.year);
      return a.faculty_code.localeCompare(b.faculty_code);
    });

    assignmentRows.sort((a, b) => {
      if (a.year !== b.year) return Number(a.year) - Number(b.year);
      if (a.faculty_code !== b.faculty_code) return a.faculty_code.localeCompare(b.faculty_code);
      return a.title.localeCompare(b.title);
    });

    buildWorkbook(outputXlsx, [
      {
        name: "Summary",
        headers: [
          "year",
          "faculty_code",
          "faculty_name",
          "first_author_papers",
          "corresponding_author_papers",
          "co_author_papers",
          "total_papers",
        ],
        rows: summaryRows,
      },
      {
        name: "Assignments",
        headers: [
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
        ],
        rows: assignmentRows,
      },
    ]);

    console.log(`Created ${outputXlsx}`);
  } finally {
    removeIfExists(tmpCopy);
    removeIfExists(tmpDir);
  }
}

main();
