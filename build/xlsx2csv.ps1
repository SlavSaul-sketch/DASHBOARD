param([string]$Xlsx, [string]$OutDir, [string]$Prefix)
# ממיר כל טאב בקובץ xlsx ל-CSV בקידוד UTF-8, בלי צורך באקסל. תאריכים נשארים כמספר סידורי של אקסל.
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $OutDir | Out-Null
$zip = [IO.Compression.ZipFile]::OpenRead($Xlsx)
function ReadEntry($name) {
  $e = $zip.GetEntry($name); if (-not $e) { return $null }
  $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8); $t = $sr.ReadToEnd(); $sr.Close(); [xml]$t
}
function TextOf($node) { if ($node -eq $null) { '' } elseif ($node -is [string]) { $node } else { $node.'#text' } }

$ss = New-Object System.Collections.Generic.List[string]
$sx = ReadEntry 'xl/sharedStrings.xml'
if ($sx) {
  foreach ($si in $sx.sst.si) {
    if ($si.t -ne $null) { $ss.Add([string](TextOf $si.t)) }
    else { $ss.Add((($si.r | ForEach-Object { TextOf $_.t }) -join '')) }
  }
}
$wbx = ReadEntry 'xl/workbook.xml'
$rels = ReadEntry 'xl/_rels/workbook.xml.rels'
$relMap = @{}; foreach ($r in $rels.Relationships.Relationship) { $relMap[$r.Id] = $r.Target }
$i = 0
foreach ($sh in $wbx.workbook.sheets.sheet) {
  $i++
  $rid = $sh.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
  $target = $relMap[$rid] -replace '^/xl/', ''; if ($target -notlike 'xl/*') { $target = 'xl/' + $target }
  $x = ReadEntry $target
  $maxC = 0; $data = @{}
  foreach ($row in $x.worksheet.sheetData.row) {
    $rn = [int]$row.r; $cells = @{}
    foreach ($c in $row.c) {
      $col = 0; foreach ($ch in ($c.r -replace '\d', '').ToCharArray()) { $col = $col * 26 + ([int]$ch - 64) }
      if ($c.t -eq 's') { $v = $ss[[int]$c.v] } elseif ($c.t -eq 'inlineStr') { $v = TextOf $c.is.t } else { $v = $c.v }
      if ($v -ne $null -and "$v" -ne '') { $cells[$col] = "$v"; if ($col -gt $maxC) { $maxC = $col } }
    }
    if ($cells.Count) { $data[$rn] = $cells }
  }
  $sb = New-Object Text.StringBuilder
  foreach ($rn in ($data.Keys | Sort-Object)) {
    $cells = $data[$rn]
    $line = for ($c = 1; $c -le $maxC; $c++) { $v = $cells[$c]; if ($v -eq $null) { '' } else { '"' + ($v -replace '"', '""' -replace "`r?`n", ' ') + '"' } }
    $null = $sb.AppendLine(($line -join ','))
  }
  [IO.File]::WriteAllText((Join-Path $OutDir "$Prefix-$i.csv"), $sb.ToString(), [Text.Encoding]::UTF8)
  # שם הטאב נשמר בקובץ נפרד כדי שהבנייה תאתר טאבים לפי שם ולא לפי מיקום
  Add-Content -Path (Join-Path $OutDir 'tabs.txt') -Value "$Prefix-$i`t$($sh.name)" -Encoding UTF8
  "$Prefix-$i | $($sh.name) | rows $($data.Count)"
}
$zip.Dispose()
