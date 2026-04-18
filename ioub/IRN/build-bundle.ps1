$files = @{
  lecturer    = "data\SCOPUS ID DATA - UB.xlsx"
  publication = "data\Publications_at_Brawijaya_University.xlsx"
  institution = "data\Institution Country File.xlsx"
}

$lines = @()
$lines += "window.BUNDLED_WORKBOOKS = {"

foreach ($key in $files.Keys) {
  $path = Join-Path $PSScriptRoot $files[$key]
  $bytes = [System.IO.File]::ReadAllBytes($path)
  $base64 = [Convert]::ToBase64String($bytes)
  $safePath = $files[$key].Replace("\", "/")
  $lines += ("  {0}: {{ path: '{1}', base64: '{2}' }}," -f $key, $safePath, $base64)
}

$lines += "};"

[System.IO.File]::WriteAllLines((Join-Path $PSScriptRoot "bundled-data.js"), $lines)
Write-Output "bundled-data.js rebuilt successfully."
