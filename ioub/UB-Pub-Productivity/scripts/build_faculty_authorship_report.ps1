param(
    [string]$LecturersCsv = (Join-Path $PSScriptRoot "..\data\master\lecturers.csv"),
    [string]$FacultiesCsv = (Join-Path $PSScriptRoot "..\data\master\faculties.csv"),
    [string]$PublicationsXlsx = (Join-Path $PSScriptRoot "..\data\master\Publications_at_Brawijaya_University.xlsx"),
    [string]$OutputXlsx = (Join-Path $PSScriptRoot "..\data\derived\faculty_authorship_report.xlsx")
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Normalize-Text {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    return ([string]$Value).Trim()
}

function Escape-XmlText {
    param([string]$Text)
    return [System.Security.SecurityElement]::Escape((Normalize-Text $Text))
}

function Get-ScopusIds {
    param([object]$Value)
    $text = Normalize-Text $Value
    if (-not $text) { return @() }

    $matches = [regex]::Matches($text, "\d{6,}")
    $ids = New-Object 'System.Collections.Generic.List[string]'
    foreach ($match in $matches) {
        $id = $match.Value.Trim()
        if ($id -and -not $ids.Contains($id)) {
            $ids.Add($id)
        }
    }
    return @($ids)
}

function Add-FacultyMapEntry {
    param(
        [hashtable]$Map,
        [string]$ScopusId,
        [string]$FacultyCode,
        [string]$FacultyName
    )

    if (-not $ScopusId -or -not $FacultyCode) { return }
    if (-not $Map.ContainsKey($ScopusId)) {
        $Map[$ScopusId] = New-Object 'System.Collections.Generic.Dictionary[string,string]'
    }
    $entry = $Map[$ScopusId]
    if (-not $entry.ContainsKey($FacultyCode)) {
        $entry[$FacultyCode] = $FacultyName
    }
}

function Get-OrAdd-Summary {
    param(
        [hashtable]$Summary,
        [string]$FacultyCode,
        [string]$FacultyName
    )

    if (-not $Summary.ContainsKey($FacultyCode)) {
        $Summary[$FacultyCode] = [ordered]@{
            faculty_code = $FacultyCode
            faculty_name = $FacultyName
            paper_count_first_author = 0
            paper_count_corresponding_author = 0
            paper_count_first_or_corresponding_author = 0
            paper_count_both_first_and_corresponding_author = 0
        }
    }
    return $Summary[$FacultyCode]
}

function Add-DetailRow {
    param(
        [System.Collections.Generic.List[object]]$Rows,
        [string]$FacultyCode,
        [string]$FacultyName,
        [string]$Role,
        [string]$Title,
        [string]$Year,
        [string]$Eid,
        [string]$Doi,
        [string]$FirstAuthorIds,
        [string]$CorrespondingAuthorIds
    )

    $Rows.Add([pscustomobject]@{
        faculty_code = $FacultyCode
        faculty_name = $FacultyName
        role = $Role
        title = $Title
        year = $Year
        eid = $Eid
        doi = $Doi
        scopus_author_id_first_author = $FirstAuthorIds
        scopus_author_id_corresponding_author = $CorrespondingAuthorIds
    }) | Out-Null
}

function Get-ColumnIndexFromReference {
    param([string]$Reference)
    $letters = ([regex]::Match($Reference, "^[A-Z]+")).Value
    $index = 0
    foreach ($char in $letters.ToCharArray()) {
        $index = ($index * 26) + ([int][char]$char - [int][char]'A' + 1)
    }
    return $index
}

function Get-ColumnName {
    param([int]$Index)
    $name = ""
    $current = $Index
    while ($current -gt 0) {
        $current--
        $name = [char](65 + ($current % 26)) + $name
        $current = [math]::Floor($current / 26)
    }
    return $name
}

function Get-NodeText {
    param($Node)
    if ($null -eq $Node) { return "" }
    if ($Node -is [string]) { return $Node }

    $parts = New-Object 'System.Collections.Generic.List[string]'
    foreach ($child in $Node.ChildNodes) {
        if ($child.LocalName -eq "t") {
            $parts.Add($child.InnerText) | Out-Null
        } else {
            $nested = Get-NodeText $child
            if ($nested) { $parts.Add($nested) | Out-Null }
        }
    }
    if ($parts.Count -eq 0 -and $Node.InnerText) {
        return $Node.InnerText
    }
    return ($parts -join "")
}

function Get-SharedStrings {
    param([string]$SharedStringsPath)
    $sharedStrings = @()
    if (-not (Test-Path $SharedStringsPath)) {
        return $sharedStrings
    }

    [xml]$xml = Get-Content -LiteralPath $SharedStringsPath
    foreach ($si in $xml.sst.si) {
        $sharedStrings += @(Get-NodeText $si)
    }
    return $sharedStrings
}

function Get-CellValue {
    param(
        $Cell,
        [string[]]$SharedStrings
    )

    $cellType = Normalize-Text $Cell.t
    if ($cellType -eq "s") {
        $index = [int](Normalize-Text $Cell.v)
        if ($index -ge 0 -and $index -lt $SharedStrings.Count) {
            return $SharedStrings[$index]
        }
        return ""
    }
    if ($cellType -eq "inlineStr") {
        return Get-NodeText $Cell.is
    }
    return Normalize-Text $Cell.v
}

function Get-FirstWorksheetPath {
    param([string]$ExtractedDir)

    [xml]$workbookXml = Get-Content -LiteralPath (Join-Path $ExtractedDir "xl\workbook.xml")
    [xml]$relsXml = Get-Content -LiteralPath (Join-Path $ExtractedDir "xl\_rels\workbook.xml.rels")

    $sheetNode = $workbookXml.workbook.sheets.sheet | Select-Object -First 1
    $relationshipId = $sheetNode.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    foreach ($rel in $relsXml.Relationships.Relationship) {
        if ($rel.Id -eq $relationshipId) {
            return Join-Path $ExtractedDir ("xl\" + $rel.Target.Replace("/", "\"))
        }
    }
    throw "Could not locate first worksheet in publication workbook."
}

function New-CellXml {
    param(
        [int]$RowIndex,
        [int]$ColumnIndex,
        [object]$Value
    )

    $cellRef = "$(Get-ColumnName $ColumnIndex)$RowIndex"
    $text = Normalize-Text $Value
    if ($text -match '^-?\d+(\.\d+)?$') {
        return "<c r=""$cellRef""><v>$text</v></c>"
    }
    $escaped = Escape-XmlText $text
    return "<c r=""$cellRef"" t=""inlineStr""><is><t xml:space=""preserve"">$escaped</t></is></c>"
}

function Write-WorksheetXml {
    param(
        [string]$Path,
        [string[]]$Headers,
        [System.Collections.IList]$Rows
    )

    $rowXml = New-Object 'System.Collections.Generic.List[string]'

    $headerCells = New-Object 'System.Collections.Generic.List[string]'
    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $headerCells.Add((New-CellXml -RowIndex 1 -ColumnIndex ($i + 1) -Value $Headers[$i])) | Out-Null
    }
    $rowXml.Add("<row r=""1"">$($headerCells -join '')</row>") | Out-Null

    $rowNumber = 2
    foreach ($row in $Rows) {
        $cells = New-Object 'System.Collections.Generic.List[string]'
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $header = $Headers[$i]
            $cells.Add((New-CellXml -RowIndex $rowNumber -ColumnIndex ($i + 1) -Value $row.$header)) | Out-Null
        }
        $rowXml.Add("<row r=""$rowNumber"">$($cells -join '')</row>") | Out-Null
        $rowNumber++
    }

    $dimensionEnd = "$(Get-ColumnName $Headers.Count)$([math]::Max(1, $Rows.Count + 1))"
    $content = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:$dimensionEnd"/>
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>$($rowXml -join '')</sheetData>
</worksheet>
"@
    Set-Content -LiteralPath $Path -Value $content -Encoding UTF8
}

function New-XlsxPackage {
    param(
        [string]$OutputPath,
        [hashtable[]]$Sheets
    )

    $tempDir = Join-Path ([System.IO.Path]::GetDirectoryName($OutputPath)) "tmp_xlsx_build"
    if (Test-Path $tempDir) {
        Remove-Item -LiteralPath $tempDir -Recurse -Force
    }

    New-Item -ItemType Directory -Force -Path $tempDir | Out-Null
    New-Item -ItemType Directory -Force -Path (Join-Path $tempDir "_rels") | Out-Null
    New-Item -ItemType Directory -Force -Path (Join-Path $tempDir "xl") | Out-Null
    New-Item -ItemType Directory -Force -Path (Join-Path $tempDir "xl\_rels") | Out-Null
    New-Item -ItemType Directory -Force -Path (Join-Path $tempDir "xl\worksheets") | Out-Null

    $worksheetOverrides = New-Object 'System.Collections.Generic.List[string]'
    $sheetEntries = New-Object 'System.Collections.Generic.List[string]'
    $sheetRelationships = New-Object 'System.Collections.Generic.List[string]'

    for ($i = 0; $i -lt $Sheets.Count; $i++) {
        $sheetIndex = $i + 1
        $sheet = $Sheets[$i]
        $sheetFile = Join-Path $tempDir "xl\worksheets\sheet$sheetIndex.xml"
        Write-WorksheetXml -Path $sheetFile -Headers $sheet.Headers -Rows $sheet.Rows
        $worksheetOverrides.Add("<Override PartName=""/xl/worksheets/sheet$sheetIndex.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>") | Out-Null
        $safeName = Escape-XmlText $sheet.Name
        $sheetEntries.Add("<sheet name=""$safeName"" sheetId=""$sheetIndex"" r:id=""rId$sheetIndex""/>") | Out-Null
        $sheetRelationships.Add("<Relationship Id=""rId$sheetIndex"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet$sheetIndex.xml""/>") | Out-Null
    }

    $contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  $($worksheetOverrides -join "`n  ")
</Types>
"@
    Set-Content -LiteralPath (Join-Path $tempDir "[Content_Types].xml") -Value $contentTypes -Encoding UTF8

    $rootRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"@
    Set-Content -LiteralPath (Join-Path $tempDir "_rels\.rels") -Value $rootRels -Encoding UTF8

    $workbookXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    $($sheetEntries -join "`n    ")
  </sheets>
</workbook>
"@
    Set-Content -LiteralPath (Join-Path $tempDir "xl\workbook.xml") -Value $workbookXml -Encoding UTF8

    $workbookRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  $($sheetRelationships -join "`n  ")
</Relationships>
"@
    Set-Content -LiteralPath (Join-Path $tempDir "xl\_rels\workbook.xml.rels") -Value $workbookRels -Encoding UTF8

    if (Test-Path $OutputPath) {
        Remove-Item -LiteralPath $OutputPath -Force
    }
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $OutputPath)
    Remove-Item -LiteralPath $tempDir -Recurse -Force
}

$lecturers = Import-Csv -LiteralPath $LecturersCsv
$faculties = Import-Csv -LiteralPath $FacultiesCsv

$facultyByScopusId = @{}
foreach ($lecturer in $lecturers) {
    $scopusId = Normalize-Text $lecturer.scopus_author_id
    $facultyCode = Normalize-Text $lecturer.faculty_code
    $facultyName = Normalize-Text $lecturer.faculty_name
    Add-FacultyMapEntry -Map $facultyByScopusId -ScopusId $scopusId -FacultyCode $facultyCode -FacultyName $facultyName
}

$summary = @{}
foreach ($faculty in $faculties) {
    $facultyCode = Normalize-Text $faculty.faculty_code
    $facultyName = Normalize-Text $faculty.faculty_name
    [void](Get-OrAdd-Summary -Summary $summary -FacultyCode $facultyCode -FacultyName $facultyName)
}

$outputDir = Split-Path -Parent $OutputXlsx
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
$tempCopy = Join-Path $outputDir "tmp_publications_working_copy.xlsx"
$tempExtracted = Join-Path $outputDir "tmp_publications_working_dir"

if (Test-Path $tempCopy) {
    Remove-Item -LiteralPath $tempCopy -Force
}
if (Test-Path $tempExtracted) {
    Remove-Item -LiteralPath $tempExtracted -Recurse -Force
}

Copy-Item -LiteralPath $PublicationsXlsx -Destination $tempCopy -Force
[System.IO.Compression.ZipFile]::ExtractToDirectory($tempCopy, $tempExtracted)

try {
    $sharedStrings = Get-SharedStrings -SharedStringsPath (Join-Path $tempExtracted "xl\sharedStrings.xml")
    $worksheetPath = Get-FirstWorksheetPath -ExtractedDir $tempExtracted
    [xml]$sheetXml = Get-Content -LiteralPath $worksheetPath

    $headers = @{}
    $rows = $sheetXml.worksheet.sheetData.row
    foreach ($cell in $rows[0].c) {
        $columnIndex = Get-ColumnIndexFromReference $cell.r
        $headers[(Get-CellValue -Cell $cell -SharedStrings $sharedStrings)] = $columnIndex
    }

    $requiredHeaders = @(
        "Title",
        "Year",
        "DOI",
        "EID",
        "Scopus Author ID First Author",
        "Scopus Author ID Corresponding Author"
    )
    foreach ($header in $requiredHeaders) {
        if (-not $headers.ContainsKey($header)) {
            throw "Missing required header in publication workbook: $header"
        }
    }

    for ($rowIndex = 1; $rowIndex -lt $rows.Count; $rowIndex++) {
        $rowNode = $rows[$rowIndex]
        $cellMap = @{}
        foreach ($cell in $rowNode.c) {
            $columnIndex = Get-ColumnIndexFromReference $cell.r
            $cellMap[$columnIndex] = Get-CellValue -Cell $cell -SharedStrings $sharedStrings
        }

        $title = Normalize-Text $cellMap[$headers["Title"]]
        $year = Normalize-Text $cellMap[$headers["Year"]]
        $doi = Normalize-Text $cellMap[$headers["DOI"]]
        $eid = Normalize-Text $cellMap[$headers["EID"]]
        $firstAuthorRaw = Normalize-Text $cellMap[$headers["Scopus Author ID First Author"]]
        $correspondingAuthorRaw = Normalize-Text $cellMap[$headers["Scopus Author ID Corresponding Author"]]

        $firstAuthorIds = Get-ScopusIds $firstAuthorRaw
        $correspondingAuthorIds = Get-ScopusIds $correspondingAuthorRaw

        $firstFaculties = New-Object 'System.Collections.Generic.Dictionary[string,string]'
        foreach ($id in $firstAuthorIds) {
            if ($facultyByScopusId.ContainsKey($id)) {
                foreach ($entry in $facultyByScopusId[$id].GetEnumerator()) {
                    $firstFaculties[$entry.Key] = $entry.Value
                }
            }
        }

        $correspondingFaculties = New-Object 'System.Collections.Generic.Dictionary[string,string]'
        foreach ($id in $correspondingAuthorIds) {
            if ($facultyByScopusId.ContainsKey($id)) {
                foreach ($entry in $facultyByScopusId[$id].GetEnumerator()) {
                    $correspondingFaculties[$entry.Key] = $entry.Value
                }
            }
        }

        foreach ($entry in $firstFaculties.GetEnumerator()) {
            $facultySummary = Get-OrAdd-Summary -Summary $summary -FacultyCode $entry.Key -FacultyName $entry.Value
            $facultySummary.paper_count_first_author++
        }

        foreach ($entry in $correspondingFaculties.GetEnumerator()) {
            $facultySummary = Get-OrAdd-Summary -Summary $summary -FacultyCode $entry.Key -FacultyName $entry.Value
            $facultySummary.paper_count_corresponding_author++
        }

        $eitherFaculties = New-Object 'System.Collections.Generic.Dictionary[string,string]'
        foreach ($entry in $firstFaculties.GetEnumerator()) {
            $eitherFaculties[$entry.Key] = $entry.Value
        }
        foreach ($entry in $correspondingFaculties.GetEnumerator()) {
            $eitherFaculties[$entry.Key] = $entry.Value
        }

        foreach ($entry in $eitherFaculties.GetEnumerator()) {
            $facultySummary = Get-OrAdd-Summary -Summary $summary -FacultyCode $entry.Key -FacultyName $entry.Value
            $facultySummary.paper_count_first_or_corresponding_author++
            if ($firstFaculties.ContainsKey($entry.Key) -and $correspondingFaculties.ContainsKey($entry.Key)) {
                $facultySummary.paper_count_both_first_and_corresponding_author++
            }
        }
    }

    $summaryRows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($facultyCode in ($summary.Keys | Sort-Object)) {
        $summaryRows.Add([pscustomobject]$summary[$facultyCode]) | Out-Null
    }

    $sheets = @(
        @{
            Name = "Summary"
            Headers = @(
                "faculty_code",
                "faculty_name",
                "paper_count_first_author",
                "paper_count_corresponding_author",
                "paper_count_first_or_corresponding_author",
                "paper_count_both_first_and_corresponding_author"
            )
            Rows = $summaryRows
        }
    )

    New-XlsxPackage -OutputPath $OutputXlsx -Sheets $sheets
    Write-Output "Created $OutputXlsx"
}
finally {
    if (Test-Path $tempCopy) {
        Remove-Item -LiteralPath $tempCopy -Force
    }
    if (Test-Path $tempExtracted) {
        Remove-Item -LiteralPath $tempExtracted -Recurse -Force
    }
}
