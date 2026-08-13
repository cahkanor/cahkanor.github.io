#!/usr/bin/env node

const fs = require("fs");
const path = require("path");
const cp = require("child_process");

const rootDir = path.resolve(__dirname, "..");
const lecturersCsv = path.join(rootDir, "data", "master", "lecturers.csv");
const publicationsXlsx = path.join(
  rootDir,
  "data",
  "master",
  "Publications_at_Brawijaya_University_2025.xlsx"
);
const outputXlsx = path.join(
  rootDir,
  "data",
  "derived",
  "faculty_research_productivity_excellence_award_2025.xlsx"
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

function ensureDir(targetPath) {
  fs.mkdirSync(targetPath, { recursive: true });
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
        i += 1;
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
      if (ch === "\r" && next === "\n") i += 1;
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
  for (const match of xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)) {
    const texts = [...match[1].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((m) =>
      decodeXml(m[1])
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
    current -= 1;
    name = String.fromCharCode(65 + (current % 26)) + name;
    current = Math.floor(current / 26);
  }
  return name;
}

function parseCellValue(type, innerXml, sharedStrings) {
  if (type === "s") {
    const valueMatch = innerXml.match(/<v>([\s\S]*?)<\/v>/);
    const idx = valueMatch ? Number(valueMatch[1]) : -1;
    return idx >= 0 && idx < sharedStrings.length ? sharedStrings[idx] : "";
  }
  if (type === "inlineStr") {
    return [...innerXml.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((m) =>
      decodeXml(m[1])
    ).join("");
  }
  const valueMatch = innerXml.match(/<v>([\s\S]*?)<\/v>/);
  return valueMatch ? decodeXml(valueMatch[1]) : "";
}

function parseWorksheetRows(worksheetPath, sharedStrings) {
  const xml = fs.readFileSync(worksheetPath, "utf8");
  const rows = [...xml.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)];
  if (!rows.length) return [];

  const cellRegex = /<c\b([^>]*)>([\s\S]*?)<\/c>/g;
  const headers = new Map();
  for (const cell of rows[0][1].matchAll(cellRegex)) {
    const attrs = cell[1];
    const inner = cell[2];
    const ref = (attrs.match(/\br="([^"]+)"/) || [])[1];
    const type = (attrs.match(/\bt="([^"]+)"/) || [])[1] || "";
    if (!ref) continue;
    headers.set(parseCellValue(type, inner, sharedStrings), columnIndex(ref));
  }

  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const cellMap = new Map();
    for (const cell of rows[i][1].matchAll(cellRegex)) {
      const attrs = cell[1];
      const inner = cell[2];
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

function buildWorksheetXml(rows, merges = [], drawingRelId = "") {
  const rowXml = rows
    .map((cells, rowIdx) => {
      const xmlCells = cells
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
      return `<row r="${rowIdx + 1}">${xmlCells}</row>`;
    })
    .join("");

  const maxCols = Math.max(...rows.map((r) => r.length), 1);
  const dimension = `A1:${columnName(maxCols)}${rows.length}`;
  const mergeXml = merges.length
    ? `<mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${m}"/>`).join("")}</mergeCells>`
    : "";
  const drawingXml = drawingRelId
    ? `<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${drawingRelId}"/>`
    : "";

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="${dimension}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>${rowXml}</sheetData>
  ${mergeXml}
  ${drawingXml}
</worksheet>`;
}

function buildDrawingXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>15</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>25</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Faculty Productivity Chart"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;
}

function buildChartXml(lastDataRow) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:lang val="en-US"/>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>2025 Publication Counts by Faculty</a:t></a:r></a:p></c:rich></c:tx>
      <c:layout/>
    </c:title>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="bar"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Summary!$C$3</c:f></c:strRef></c:tx>
          <c:cat>
            <c:strRef>
              <c:f>Summary!$B$4:$B$${lastDataRow}</c:f>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Summary!$C$4:$C$${lastDataRow}</c:f>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:axId val="48650112"/>
        <c:axId val="48672768"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="48650112"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="48672768"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="48672768"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:majorGridlines/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="48650112"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="r"/><c:layout/></c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;
}

function writeWorkbook(outputPath, summaryXml, notesXml, chartXml, drawingXml) {
  const outputDir = path.dirname(outputPath);
  const buildDir = path.join(outputDir, "tmp_faculty_award_2025_xlsx");
  removeIfExists(buildDir);
  ensureDir(path.join(buildDir, "_rels"));
  ensureDir(path.join(buildDir, "xl", "_rels"));
  ensureDir(path.join(buildDir, "xl", "worksheets", "_rels"));
  ensureDir(path.join(buildDir, "xl", "worksheets"));
  ensureDir(path.join(buildDir, "xl", "drawings", "_rels"));
  ensureDir(path.join(buildDir, "xl", "drawings"));
  ensureDir(path.join(buildDir, "xl", "charts"));

  fs.writeFileSync(path.join(buildDir, "xl", "worksheets", "sheet1.xml"), summaryXml, "utf8");
  fs.writeFileSync(path.join(buildDir, "xl", "worksheets", "sheet2.xml"), notesXml, "utf8");
  fs.writeFileSync(path.join(buildDir, "xl", "drawings", "drawing1.xml"), drawingXml, "utf8");
  fs.writeFileSync(path.join(buildDir, "xl", "charts", "chart1.xml"), chartXml, "utf8");

  fs.writeFileSync(
    path.join(buildDir, "xl", "worksheets", "_rels", "sheet1.xml.rels"),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`,
    "utf8"
  );

  fs.writeFileSync(
    path.join(buildDir, "xl", "drawings", "_rels", "drawing1.xml.rels"),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`,
    "utf8"
  );

  fs.writeFileSync(
    path.join(buildDir, "[Content_Types].xml"),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
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
    <sheet name="Methodology" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`,
    "utf8"
  );

  fs.writeFileSync(
    path.join(buildDir, "xl", "_rels", "workbook.xml.rels"),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`,
    "utf8"
  );

  if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);
  const escapedBuildDir = buildDir.replace(/'/g, "''");
  const escapedOutput = String(outputPath).replace(/'/g, "''");
  cp.execFileSync("powershell", [
    "-Command",
    `Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::CreateFromDirectory('${escapedBuildDir}', '${escapedOutput}')`,
  ]);

  removeIfExists(buildDir);
}

function main() {
  ensureDir(path.dirname(outputXlsx));

  const lecturers = parseCsv(fs.readFileSync(lecturersCsv, "utf8"));
  const activeLecturerByScopus = new Map();
  for (const lecturer of lecturers) {
    const activeStatus = normalize(lecturer.active_status || lecturer.Keterangan || lecturer.Status).toLowerCase();
    if (activeStatus && !activeStatus.includes("aktif")) continue;
    const scopusId = normalize(lecturer.scopus_author_id || lecturer["SCOPUS ID"]);
    if (!scopusId || activeLecturerByScopus.has(scopusId)) continue;
    activeLecturerByScopus.set(scopusId, {
      faculty_code: normalize(lecturer.faculty_code),
      faculty_name: normalize(lecturer.faculty_name),
      lecturer_name: normalize(lecturer.lecturer_name || lecturer.Nama),
      scopus_author_id: scopusId,
    });
  }

  const tmpCopy = path.join(path.dirname(outputXlsx), "tmp_award_2025_copy.xlsx");
  const tmpDir = path.join(path.dirname(outputXlsx), "tmp_award_2025_dir");
  removeIfExists(tmpCopy);
  removeIfExists(tmpDir);
  fs.copyFileSync(publicationsXlsx, tmpCopy);
  cp.execFileSync("powershell", [
    "-Command",
    `Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::ExtractToDirectory('${tmpCopy.replace(/'/g, "''")}','${tmpDir.replace(/'/g, "''")}')`,
  ]);

  try {
    const sharedStrings = extractSharedStrings(path.join(tmpDir, "xl", "sharedStrings.xml"));
    const worksheetPath = getFirstWorksheetPath(tmpDir);
    const publicationRows = parseWorksheetRows(worksheetPath, sharedStrings);

    const uniquePublications = new Map();
    for (const row of publicationRows) {
      const year = normalize(row.Year || row.year);
      const fullDate = normalize(row["Full date"] || row["Full Date"]);
      const is2025 =
        year === "2025" ||
        fullDate.startsWith("2025-") ||
        fullDate.startsWith("2025/");
      if (!is2025) continue;

      const eid = normalize(row.EID);
      const doi = normalize(row.DOI);
      const title = normalize(row.Title);
      const key = eid || doi || title;
      if (!key) continue;
      if (!uniquePublications.has(key)) {
        uniquePublications.set(key, row);
      }
    }

    const assignments = [];
    const facultyCounts = new Map();

    for (const row of uniquePublications.values()) {
      const firstAuthorIds = getScopusIds(row["Scopus Author ID First Author"]);
      const corrAuthorIds = getScopusIds(row["Scopus Author ID Corresponding Author"]);
      const allAuthorIds = getScopusIds(row["Scopus Author Ids"]);
      const singleAuthorIds = getScopusIds(row["Scopus Author ID Single Author"]);

      let assignedFaculty = null;
      let category = "";
      let assignmentReason = "";
      let matchedLecturerName = "";
      let matchedScopusId = "";

      const effectiveFirstIds =
        singleAuthorIds.length === 1 && firstAuthorIds.length === 0 ? singleAuthorIds : firstAuthorIds;

      for (const scopusId of effectiveFirstIds) {
        const lecturer = activeLecturerByScopus.get(scopusId);
        if (lecturer) {
          assignedFaculty = lecturer;
          category = "First Author";
          assignmentReason = singleAuthorIds.includes(scopusId)
            ? "Single author paper counted as First Author"
            : "First author is an active UB lecturer";
          matchedLecturerName = lecturer.lecturer_name;
          matchedScopusId = scopusId;
          break;
        }
      }

      if (!assignedFaculty) {
        for (const scopusId of corrAuthorIds) {
          const lecturer = activeLecturerByScopus.get(scopusId);
          if (lecturer) {
            assignedFaculty = lecturer;
            category = "Corresponding Author";
            assignmentReason = "First author is not active UB; corresponding author is an active UB lecturer";
            matchedLecturerName = lecturer.lecturer_name;
            matchedScopusId = scopusId;
            break;
          }
        }
      }

      if (!assignedFaculty) {
        for (const scopusId of allAuthorIds) {
          const lecturer = activeLecturerByScopus.get(scopusId);
          if (lecturer) {
            assignedFaculty = lecturer;
            category = "Co-author";
            assignmentReason = "First UB-affiliated lecturer in author order";
            matchedLecturerName = lecturer.lecturer_name;
            matchedScopusId = scopusId;
            break;
          }
        }
      }

      if (!assignedFaculty) {
        assignedFaculty = {
          faculty_code: "OTHERS",
          faculty_name: "Others",
        };
        category = "Co-author";
        assignmentReason = "No active UB lecturer found in author list";
      }

      const facultyKey = assignedFaculty.faculty_code;
      if (!facultyCounts.has(facultyKey)) {
        facultyCounts.set(facultyKey, {
          faculty_code: assignedFaculty.faculty_code,
          faculty_name: assignedFaculty.faculty_name,
          publications: 0,
        });
      }
      facultyCounts.get(facultyKey).publications += 1;

      assignments.push({
        faculty_code: assignedFaculty.faculty_code,
        faculty_name: assignedFaculty.faculty_name,
        category,
        assignment_reason: assignmentReason,
        matched_lecturer_name: matchedLecturerName,
        matched_scopus_author_id: matchedScopusId,
        title: normalize(row.Title),
        eid: normalize(row.EID),
        doi: normalize(row.DOI),
      });
    }

    const totalUniquePublications = uniquePublications.size;
    const othersCount = facultyCounts.get("OTHERS")?.publications ?? 0;
    const rankingRows = [...facultyCounts.values()]
      .filter((row) => row.faculty_code !== "OTHERS")
      .sort((a, b) => {
        if (b.publications !== a.publications) return b.publications - a.publications;
        return a.faculty_name.localeCompare(b.faculty_name);
      });

    rankingRows.forEach((row, index) => {
      row.rank = index + 1;
      row.percentage = totalUniquePublications
        ? ((row.publications / totalUniquePublications) * 100).toFixed(2) + "%"
        : "0.00%";
      row.award =
        index === 0 ? "Faculty Research Productivity Excellence Award" : "";
    });

    const facultiesRepresented = rankingRows.filter((row) => row.publications > 0).length;
    const top3 = rankingRows.slice(0, 3);
    const sumAcrossAllBuckets = [...facultyCounts.values()].reduce((sum, row) => sum + row.publications, 0);
    const qualityPass = sumAcrossAllBuckets === totalUniquePublications ? "PASS" : "FAIL";
    const duplicatesRemoved = publicationRows.filter((row) => {
      const year = normalize(row.Year || row.year);
      const fullDate = normalize(row["Full date"] || row["Full Date"]);
      return year === "2025" || fullDate.startsWith("2025-") || fullDate.startsWith("2025/");
    }).length - totalUniquePublications;

    const summaryRows = [
      ["Faculty Research Productivity Excellence Award 2025"],
      [""],
      ["Rank", "Faculty Name", "Number of Publications", "Percentage of Total Publications", "Award"],
      ...rankingRows.map((row) => [
        row.rank,
        row.faculty_name,
        row.publications,
        row.percentage,
        row.award,
      ]),
      [""],
      ["Summary Statistics"],
      ["Total unique publications analyzed", totalUniquePublications],
      ["Total UB faculties represented", facultiesRepresented],
      ["Top 1 faculty", top3[0] ? `${top3[0].faculty_name} (${top3[0].publications})` : ""],
      ["Top 2 faculty", top3[1] ? `${top3[1].faculty_name} (${top3[1].publications})` : ""],
      ["Top 3 faculty", top3[2] ? `${top3[2].faculty_name} (${top3[2].publications})` : ""],
      ["Others bucket publications", othersCount],
      [""],
      ["Quality Check"],
      ["Publications counted exactly once", qualityPass],
      ["Sum across all faculties plus Others", sumAcrossAllBuckets],
      ["Unique publication total", totalUniquePublications],
      ["Duplicate source rows removed", duplicatesRemoved],
    ];

    const summaryMerges = ["A1:E1", "A6:B6", "A14:B14"];
    const summaryXml = buildWorksheetXml(summaryRows, summaryMerges, "rId1");
    const lastDataRow = 3 + rankingRows.length;
    const chartXml = buildChartXml(lastDataRow);
    const drawingXml = buildDrawingXml();

    const methodologyRows = [
      ["Methodology"],
      [""],
      [
        "The workbook was filtered to publications in January-December 2025 using the Year or Full date field."
      ],
      [
        "Each unique publication was identified by EID, or DOI, or Title when EID/DOI was blank."
      ],
      [
        "Attribution priority was applied once per publication: First Author, then Corresponding Author, then the first active UB lecturer appearing in the author list."
      ],
      [
        "A publication was assigned to Others only when no active UB lecturer from the master list appeared in the author metadata."
      ],
      [
        "Percentages are based on the total number of unique publications analyzed in the 2025 dataset."
      ],
      [""],
      ["Assumptions / Ambiguities"],
      [
        "UB affiliation was determined only from active lecturers in lecturers.csv matched by Scopus Author ID."
      ],
      [
        "If a publication listed multiple corresponding author IDs, the first active UB lecturer found in that list received the attribution."
      ],
      [
        "If duplicate publication rows existed in the workbook, they were de-duplicated before attribution."
      ],
      [""],
      ["Quality Check Result"],
      [
        `PASS status: ${qualityPass}. Sum across all faculties and Others = ${sumAcrossAllBuckets}; unique publications = ${totalUniquePublications}.`
      ],
    ];
    const notesXml = buildWorksheetXml(methodologyRows);

    writeWorkbook(outputXlsx, summaryXml, notesXml, chartXml, drawingXml);
    console.log(`Created ${outputXlsx}`);
  } finally {
    removeIfExists(tmpCopy);
    removeIfExists(tmpDir);
    removeIfExists(path.join(rootDir, "tmp_pub2025_inspect"));
  }
}

main();
