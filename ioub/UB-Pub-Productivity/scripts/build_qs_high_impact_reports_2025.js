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
const outputDir = path.join(rootDir, "data", "derived", "qs_high_impact_2025");

const SUBJECT_AREAS = [
  { code: "1", name: "Arts and Humanities", slug: "arts_and_humanities" },
  { code: "2", name: "Engineering and Technology", slug: "engineering_and_technology" },
  { code: "3", name: "Social Sciences and Management", slug: "social_sciences_and_management" },
  { code: "4", name: "Natural Sciences", slug: "natural_sciences" },
  { code: "5", name: "Life Sciences and Medicine", slug: "life_sciences_and_medicine" },
];

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
  const headers = rows[0].map((value) => normalize(value));
  return rows
    .slice(1)
    .filter((values) => values.some((value) => normalize(value)))
    .map((values) => Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""])));
}

function getScopusIds(value) {
  return [...new Set((normalize(value).match(/\d{6,}/g) || []))];
}

function extractSharedStrings(sharedStringsPath) {
  if (!fs.existsSync(sharedStringsPath)) return [];
  const xml = fs.readFileSync(sharedStringsPath, "utf8");
  const values = [];
  for (const match of xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)) {
    const texts = [...match[1].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((item) =>
      decodeXml(item[1])
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
  let out = "";
  let current = index;
  while (current > 0) {
    current -= 1;
    out = String.fromCharCode(65 + (current % 26)) + out;
    current = Math.floor(current / 26);
  }
  return out;
}

function parseCellValue(type, innerXml, sharedStrings) {
  if (type === "s") {
    const match = innerXml.match(/<v>([\s\S]*?)<\/v>/);
    const idx = match ? Number(match[1]) : -1;
    return idx >= 0 && idx < sharedStrings.length ? sharedStrings[idx] : "";
  }
  if (type === "inlineStr") {
    return [...innerXml.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((item) =>
      decodeXml(item[1])
    ).join("");
  }
  const match = innerXml.match(/<v>([\s\S]*?)<\/v>/);
  return match ? decodeXml(match[1]) : "";
}

function parseWorksheetRows(worksheetPath, sharedStrings) {
  const xml = fs.readFileSync(worksheetPath, "utf8");
  const rows = [...xml.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)];
  if (!rows.length) return { headers: [], rows: [] };

  const cellRegex = /<c\b([^>]*)>([\s\S]*?)<\/c>/g;
  const headerCols = new Map();
  for (const cell of rows[0][1].matchAll(cellRegex)) {
    const attrs = cell[1];
    const inner = cell[2];
    const ref = (attrs.match(/\br="([^"]+)"/) || [])[1];
    const type = (attrs.match(/\bt="([^"]+)"/) || [])[1] || "";
    if (!ref) continue;
    headerCols.set(columnIndex(ref), parseCellValue(type, inner, sharedStrings));
  }

  const headers = [...headerCols.entries()]
    .sort((a, b) => a[0] - b[0])
    .map((item) => item[1]);

  const out = [];
  for (let rowIndex = 1; rowIndex < rows.length; rowIndex++) {
    const valuesByCol = new Map();
    for (const cell of rows[rowIndex][1].matchAll(cellRegex)) {
      const attrs = cell[1];
      const inner = cell[2];
      const ref = (attrs.match(/\br="([^"]+)"/) || [])[1];
      const type = (attrs.match(/\bt="([^"]+)"/) || [])[1] || "";
      if (!ref) continue;
      valuesByCol.set(columnIndex(ref), parseCellValue(type, inner, sharedStrings));
    }

    const row = {};
    for (const [colIndex, header] of headerCols.entries()) {
      row[header] = valuesByCol.get(colIndex) ?? "";
    }
    if (Object.values(row).some((value) => normalize(value))) out.push(row);
  }

  return { headers, rows: out };
}

function parseNumber(value) {
  const text = normalize(value).replace(/,/g, "");
  if (!text || text === "-") return NaN;
  const num = Number(text);
  return Number.isFinite(num) ? num : NaN;
}

function publicationKey(row) {
  return normalize(row.EID) || normalize(row.DOI) || normalize(row.Title);
}

function formulaSheetName(sheetName) {
  return `'${sheetName.replace(/'/g, "''")}'`;
}

function buildWorksheetXml(rows, merges = [], options = {}) {
  const drawingRelId = options.drawingRelId || "";
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

  const maxCols = Math.max(1, ...rows.map((row) => row.length));
  const dimension = `A1:${columnName(maxCols)}${Math.max(rows.length, 1)}`;
  const mergeXml = merges.length
    ? `<mergeCells count="${merges.length}">${merges.map((value) => `<mergeCell ref="${value}"/>`).join("")}</mergeCells>`
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
    <xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>14</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Faculty Ranking Chart"/>
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

function buildChartXml(sheetName, lastDataRow, title) {
  const sheetRef = formulaSheetName(sheetName);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${escapeXml(
        title
      )}</a:t></a:r></a:p></c:rich></c:tx>
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
          <c:tx><c:v>Publications</c:v></c:tx>
          <c:cat>
            <c:strRef><c:f>${sheetRef}!$B$2:$B$${lastDataRow}</c:f></c:strRef>
          </c:cat>
          <c:val>
            <c:numRef><c:f>${sheetRef}!$C$2:$C$${lastDataRow}</c:f></c:numRef>
          </c:val>
        </c:ser>
        <c:axId val="64451712"/>
        <c:axId val="64453248"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="64451712"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="64453248"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="64453248"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:majorGridlines/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="64451712"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="r"/><c:layout/></c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;
}

function writeWorkbook(outputPath, sheets) {
  const buildDir = path.join(path.dirname(outputPath), `tmp_${path.basename(outputPath, ".xlsx")}`);
  removeIfExists(buildDir);
  ensureDir(buildDir);
  ensureDir(path.join(buildDir, "_rels"));
  ensureDir(path.join(buildDir, "docProps"));
  ensureDir(path.join(buildDir, "xl"));
  ensureDir(path.join(buildDir, "xl", "_rels"));
  ensureDir(path.join(buildDir, "xl", "worksheets"));

  const contentTypeOverrides = [
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
  ];

  const workbookSheetsXml = [];
  const workbookRelsXml = [];
  let relationshipId = 1;
  let chartCounter = 1;
  let drawingCounter = 1;

  sheets.forEach((sheet, index) => {
    const sheetId = index + 1;
    const worksheetPath = path.join(buildDir, "xl", "worksheets", `sheet${sheetId}.xml`);
    fs.writeFileSync(worksheetPath, sheet.xml, "utf8");
    contentTypeOverrides.push(
      `<Override PartName="/xl/worksheets/sheet${sheetId}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    );

    const relId = `rId${relationshipId++}`;
    workbookSheetsXml.push(
      `<sheet name="${escapeXml(sheet.name)}" sheetId="${sheetId}" r:id="${relId}"/>`
    );
    workbookRelsXml.push(
      `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${sheetId}.xml"/>`
    );

    if (sheet.drawingXml && sheet.chartXml) {
      ensureDir(path.join(buildDir, "xl", "drawings"));
      ensureDir(path.join(buildDir, "xl", "drawings", "_rels"));
      ensureDir(path.join(buildDir, "xl", "charts"));
      ensureDir(path.join(buildDir, "xl", "worksheets", "_rels"));

      const drawingName = `drawing${drawingCounter}.xml`;
      const chartName = `chart${chartCounter}.xml`;
      fs.writeFileSync(path.join(buildDir, "xl", "drawings", drawingName), sheet.drawingXml, "utf8");
      fs.writeFileSync(path.join(buildDir, "xl", "charts", chartName), sheet.chartXml, "utf8");
      fs.writeFileSync(
        path.join(buildDir, "xl", "drawings", "_rels", `drawing${drawingCounter}.xml.rels`),
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/${chartName}"/>
</Relationships>`,
        "utf8"
      );
      fs.writeFileSync(
        path.join(buildDir, "xl", "worksheets", "_rels", `sheet${sheetId}.xml.rels`),
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/${drawingName}"/>
</Relationships>`,
        "utf8"
      );
      contentTypeOverrides.push(
        `<Override PartName="/xl/drawings/${drawingName}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`
      );
      contentTypeOverrides.push(
        `<Override PartName="/xl/charts/${chartName}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`
      );
      drawingCounter += 1;
      chartCounter += 1;
    }
  });

  fs.writeFileSync(
    path.join(buildDir, "[Content_Types].xml"),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${contentTypeOverrides.join("\n  ")}
</Types>`,
    "utf8"
  );

  fs.writeFileSync(
    path.join(buildDir, "_rels", ".rels"),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`,
    "utf8"
  );

  fs.writeFileSync(
    path.join(buildDir, "docProps", "core.xml"),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>QS High Impact Reports 2025</dc:title>
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-07-20T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-20T00:00:00Z</dcterms:modified>
</cp:coreProperties>`,
    "utf8"
  );

  fs.writeFileSync(
    path.join(buildDir, "docProps", "app.xml"),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
</Properties>`,
    "utf8"
  );

  fs.writeFileSync(
    path.join(buildDir, "xl", "workbook.xml"),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${workbookSheetsXml.join("")}</sheets>
</workbook>`,
    "utf8"
  );

  fs.writeFileSync(
    path.join(buildDir, "xl", "_rels", "workbook.xml.rels"),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${workbookRelsXml.join("\n  ")}
</Relationships>`,
    "utf8"
  );

  const escapedBuildDir = buildDir.replace(/'/g, "''");
  const escapedOutput = outputPath.replace(/'/g, "''");
  removeIfExists(outputPath);
  cp.execFileSync("powershell", [
    "-Command",
    `Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::CreateFromDirectory('${escapedBuildDir}', '${escapedOutput}')`,
  ]);
  removeIfExists(buildDir);
}

function buildActiveLecturerMap(lecturers) {
  const map = new Map();
  for (const lecturer of lecturers) {
    const activeStatus = normalize(lecturer.active_status || lecturer.Keterangan || lecturer.Status).toLowerCase();
    if (activeStatus && !activeStatus.includes("aktif")) continue;
    const scopusId = normalize(lecturer.scopus_author_id || lecturer["SCOPUS ID"]);
    if (!scopusId || map.has(scopusId)) continue;
    map.set(scopusId, {
      lecturer_name: normalize(lecturer.lecturer_name || lecturer.Nama),
      faculty_code: normalize(lecturer.faculty_code),
      faculty_name: normalize(lecturer.faculty_name),
      scopus_author_id: scopusId,
    });
  }
  return map;
}

function assignPublication(row, activeLecturerByScopus) {
  const firstAuthorIds = getScopusIds(row["Scopus Author ID First Author"]);
  const corrAuthorIds = getScopusIds(row["Scopus Author ID Corresponding Author"]);
  const allAuthorIds = getScopusIds(row["Scopus Author Ids"]);
  const singleAuthorIds = getScopusIds(row["Scopus Author ID Single Author"]);

  let assignedLecturer = null;
  let rule = "";
  let reason = "";

  const effectiveFirstIds =
    singleAuthorIds.length === 1 && firstAuthorIds.length === 0 ? singleAuthorIds : firstAuthorIds;

  for (const scopusId of effectiveFirstIds) {
    const lecturer = activeLecturerByScopus.get(scopusId);
    if (lecturer) {
      assignedLecturer = lecturer;
      rule = "First Author";
      reason = singleAuthorIds.includes(scopusId)
        ? "Single-author paper counted as First Author"
        : "First author is an active UB lecturer";
      break;
    }
  }

  if (!assignedLecturer) {
    for (const scopusId of corrAuthorIds) {
      const lecturer = activeLecturerByScopus.get(scopusId);
      if (lecturer) {
        assignedLecturer = lecturer;
        rule = "Corresponding Author";
        reason = "First author is not active UB; corresponding author is an active UB lecturer";
        break;
      }
    }
  }

  if (!assignedLecturer) {
    for (const scopusId of allAuthorIds) {
      const lecturer = activeLecturerByScopus.get(scopusId);
      if (lecturer) {
        assignedLecturer = lecturer;
        rule = "First UB Co-author";
        reason = "First active UB lecturer in the author list";
        break;
      }
    }
  }

  if (!assignedLecturer) {
    assignedLecturer = {
      lecturer_name: "",
      faculty_code: "OTHERS",
      faculty_name: "Others",
      scopus_author_id: "",
    };
    rule = "Unmatched";
    reason = "No active UB lecturer found in the author metadata";
  }

  return {
    faculty_code: assignedLecturer.faculty_code,
    faculty_name: assignedLecturer.faculty_name,
    lecturer_name: assignedLecturer.lecturer_name,
    lecturer_scopus_id: assignedLecturer.scopus_author_id,
    attribution_rule: rule,
    attribution_reason: reason,
  };
}

function qsCodes(row) {
  return [...new Set(normalize(row["Quacquarelli Symonds (QS) Subject area code"])
    .split("|")
    .map((value) => normalize(value))
    .filter((value) => value))];
}

function filterUnique2025Publications(rows) {
  const unique = new Map();
  let raw2025Rows = 0;
  for (const row of rows) {
    const year = normalize(row.Year || row.year);
    const fullDate = normalize(row["Full date"] || row["Full Date"]);
    const is2025 = year === "2025" || fullDate.startsWith("2025-") || fullDate.startsWith("2025/");
    if (!is2025) continue;
    raw2025Rows += 1;

    const key = publicationKey(row);
    if (!key || unique.has(key)) continue;
    unique.set(key, row);
  }
  return { uniqueRows: [...unique.values()], raw2025Rows };
}

function createSubjectAreaReport(subjectArea, headers, sourceRows, activeLecturerByScopus) {
  const filteredRows = [];
  for (const row of sourceRows) {
    const fwci = parseNumber(row["Field-Weighted Citation Impact"]);
    if (!(fwci > 2)) continue;
    if (!qsCodes(row).includes(subjectArea.code)) continue;
    filteredRows.push(row);
  }

  const assignments = filteredRows.map((row) => {
    const assignment = assignPublication(row, activeLecturerByScopus);
    return {
      ...assignment,
      title: normalize(row.Title),
      eid: normalize(row.EID),
      doi: normalize(row.DOI),
      fwci: normalize(row["Field-Weighted Citation Impact"]),
      qs_subject_area_code: subjectArea.code,
      qs_subject_area_name: subjectArea.name,
    };
  });

  const facultyMap = new Map();
  for (const assignment of assignments) {
    const key = assignment.faculty_code || "OTHERS";
    if (!facultyMap.has(key)) {
      facultyMap.set(key, {
        faculty_code: assignment.faculty_code,
        faculty_name: assignment.faculty_name,
        publications: 0,
        lecturers: new Set(),
      });
    }
    const bucket = facultyMap.get(key);
    bucket.publications += 1;
    if (assignment.lecturer_scopus_id) bucket.lecturers.add(assignment.lecturer_scopus_id);
  }

  const rankingRows = [...facultyMap.values()].sort((a, b) => {
    if (b.publications !== a.publications) return b.publications - a.publications;
    if (b.lecturers.size !== a.lecturers.size) return b.lecturers.size - a.lecturers.size;
    return a.faculty_name.localeCompare(b.faculty_name);
  });

  rankingRows.forEach((row, index) => {
    row.rank = index + 1;
    row.highlight = index === 0 && row.faculty_code !== "OTHERS" ? "Top Faculty" : "";
  });

  const uniqueAttributedLecturers = new Set(
    assignments.map((item) => item.lecturer_scopus_id).filter((value) => value)
  );
  const totalAttributed = rankingRows.reduce((sum, row) => sum + row.publications, 0);
  const othersCount = facultyMap.get("OTHERS")?.publications ?? 0;
  const ubRankingRows = rankingRows.filter((row) => row.faculty_code !== "OTHERS");
  const topFaculty = ubRankingRows[0] || rankingRows[0] || null;
  const facultiesRepresented = ubRankingRows.filter((row) => row.publications > 0).length;
  const qualityPass = totalAttributed === filteredRows.length ? "PASS" : "FAIL";

  const publicationsSheetRows = [
    headers,
    ...filteredRows.map((row) => headers.map((header) => normalize(row[header]))),
  ];

  const rankingSheetRows = [
    ["Rank", "Faculty", "Number of Publications", "Number of Unique Lecturers", "Highlight"],
    ...rankingRows.map((row) => [
      row.rank,
      row.faculty_name,
      row.publications,
      row.lecturers.size,
      row.highlight,
    ]),
  ];

  const summarySheetRows = [
    [`High-Impact Publications Report 2025 - ${subjectArea.name}`],
    [""],
    ["Metric", "Value"],
    ["QS subject area code", subjectArea.code],
    ["QS subject area", subjectArea.name],
    ["FWCI threshold", "FWCI > 2"],
    ["Total publications in this subject-area file", filteredRows.length],
    ["Total UB faculties represented", facultiesRepresented],
    ["Total unique attributed lecturers", uniqueAttributedLecturers.size],
    ["Top-ranked UB faculty", topFaculty ? `${topFaculty.faculty_name} (${topFaculty.publications})` : ""],
    ["Others bucket publications", othersCount],
    [""],
    ["Quality Check", ""],
    ["Every publication assigned exactly once", qualityPass],
    ["Total attributed publications", totalAttributed],
    ["Subject-area publication total", filteredRows.length],
    ["Cross-area duplicates retained independently", "Yes"],
    ["Unique lecturer count based only on attribution lecturer", "Yes"],
  ];

  const methodologySheetRows = [
    ["Methodology"],
    [""],
    ["Source file filtered to January-December 2025 publications only."],
    ["Only publications with Field-Weighted Citation Impact (FWCI) greater than 2 were included."],
    [
      "QS subject-area membership was read from 'Quacquarelli Symonds (QS) Subject area code'; publications with multiple codes were copied into every matching subject-area workbook."
    ],
    [
      "Attribution priority was applied exactly once per publication inside each subject-area workbook: First Author, then Corresponding Author, then the first active UB lecturer in the author list."
    ],
    [
      "Active UB lecturers were matched from lecturers.csv using Scopus Author ID and active status containing 'aktif'."
    ],
    [
      "If no active UB lecturer could be matched, the publication was placed in the Others bucket to preserve one-and-only-one attribution."
    ],
    [""],
    ["Assumptions"],
    [
      "The lecturer chosen for faculty attribution is also the only lecturer counted for the unique-lecturer metric."
    ],
    [
      "If multiple corresponding-author IDs were present, the first active UB lecturer found in that list was used."
    ],
    [
      "Duplicate source rows were removed before subject-area filtering using EID, then DOI, then Title as the publication key."
    ],
  ];

  const attributionSheetRows = [
    [
      "Title",
      "EID",
      "DOI",
      "FWCI",
      "Attributed Faculty",
      "Attributed Lecturer",
      "Attributed Lecturer Scopus ID",
      "Attribution Rule",
      "Attribution Reason",
    ],
    ...assignments.map((item) => [
      item.title,
      item.eid,
      item.doi,
      item.fwci,
      item.faculty_name,
      item.lecturer_name,
      item.lecturer_scopus_id,
      item.attribution_rule,
      item.attribution_reason,
    ]),
  ];

  return {
    filteredRows,
    assignments,
    rankingRows,
    summarySheetRows,
    methodologySheetRows,
    publicationsSheetRows,
    rankingSheetRows,
    attributionSheetRows,
    qualityPass,
  };
}

function main() {
  ensureDir(outputDir);

  const lecturers = parseCsv(fs.readFileSync(lecturersCsv, "utf8"));
  const activeLecturerByScopus = buildActiveLecturerMap(lecturers);

  const tmpCopy = path.join(outputDir, "tmp_qs_high_impact_2025_copy.xlsx");
  const tmpDir = path.join(outputDir, "tmp_qs_high_impact_2025_dir");
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
    const parsed = parseWorksheetRows(worksheetPath, sharedStrings);
    const deduped = filterUnique2025Publications(parsed.rows);

    const results = [];

    for (const subjectArea of SUBJECT_AREAS) {
      const report = createSubjectAreaReport(
        subjectArea,
        parsed.headers,
        deduped.uniqueRows,
        activeLecturerByScopus
      );
      const filePath = path.join(
        outputDir,
        `high_impact_${subjectArea.slug}_2025.xlsx`
      );

      const rankingLastRow = Math.max(report.rankingRows.length + 1, 2);
      writeWorkbook(filePath, [
        {
          name: "Publications",
          xml: buildWorksheetXml(report.publicationsSheetRows),
        },
        {
          name: "Faculty Ranking",
          xml: buildWorksheetXml(report.rankingSheetRows, [], { drawingRelId: "rId1" }),
          drawingXml: buildDrawingXml(),
          chartXml: buildChartXml(
            "Faculty Ranking",
            rankingLastRow,
            `${subjectArea.name} - High-Impact Publications by Faculty`
          ),
        },
        {
          name: "Summary",
          xml: buildWorksheetXml(report.summarySheetRows, ["A1:B1", "A13:B13"]),
        },
        {
          name: "Methodology",
          xml: buildWorksheetXml(report.methodologySheetRows),
        },
        {
          name: "Attribution",
          xml: buildWorksheetXml(report.attributionSheetRows),
        },
      ]);

      results.push({
        subjectArea: subjectArea.name,
        publications: report.filteredRows.length,
        topFaculty: report.rankingRows.find((row) => row.faculty_code !== "OTHERS")?.faculty_name || "None",
        quality: report.qualityPass,
        filePath,
      });
    }

    console.table(results);
  } finally {
    removeIfExists(tmpCopy);
    removeIfExists(tmpDir);
  }
}

main();
