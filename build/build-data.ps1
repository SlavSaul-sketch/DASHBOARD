param(
  [Parameter(Mandatory)] [string]$CsvDir,    # תיקייה עם ה-CSV של הטאבים (פלט של xlsx2csv.ps1) + tabs.txt
  [Parameter(Mandatory)] [string]$MetaCsv,   # דוח Meta ברמת מודעה
  [string]$From = '2026-06-01',              # תחילת הניתוח
  [string]$Template = (Join-Path $PSScriptRoot 'dashboard.template.html'),
  [string]$Out = (Join-Path (Split-Path $PSScriptRoot) 'dashboard.html')
)
# בונה את קובץ הדאשבורד מתוך הנתונים. לא כותב לשום גיליון. טלפונים ומיילים מוחלפים במפתח מוצפן חד כיווני.
$ErrorActionPreference = 'Stop'
$fromDate = [datetime]::ParseExact($From, 'yyyy-MM-dd', $null)
$attribFrom = $fromDate.AddDays(-90)   # הרשמות מ-90 יום לפני תחילת הניתוח, לצורך שיוך Last Touch
$warnings = New-Object System.Collections.Generic.List[string]

# ---------- טעינת טאבים לפי שם ----------
$tabs = @{}
foreach ($l in [IO.File]::ReadAllLines((Join-Path $CsvDir 'tabs.txt'), [Text.Encoding]::UTF8)) {
  $p = $l -split "`t"; if ($p.Count -eq 2) { $tabs[$p[1].Trim()] = Join-Path $CsvDir ($p[0] + '.csv') }
}
function TabFile([string]$name, [switch]$Prefix) {
  $k = if ($Prefix) { $tabs.Keys | Where-Object { $_.StartsWith($name) } | Select-Object -First 1 } else { $tabs.Keys | Where-Object { $_ -eq $name } | Select-Object -First 1 }
  if (-not $k) { throw "הטאב '$name' לא נמצא. ייתכן ששמו שונה בגיליון." }
  $tabs[$k]
}
function LoadCsv([string]$file) {
  $lines = [IO.File]::ReadAllLines($file, [Text.Encoding]::UTF8)
  if ($lines.Count -lt 2) { return @() }
  $hdr = ($lines[0] | ConvertFrom-Csv -Header (1..400 | ForEach-Object { "c$_" })).PSObject.Properties | Where-Object { $_.Value -ne $null } | ForEach-Object { $_.Value }
  $seen = @{}; $names = foreach ($h in $hdr) { $n = if ([string]::IsNullOrWhiteSpace($h)) { '_' } else { $h.Trim() }; if ($seen[$n]) { $seen[$n]++; "$n`_$($seen[$n])" } else { $seen[$n] = 1; $n } }
  $lines[1..($lines.Count - 1)] | ConvertFrom-Csv -Header $names
}

# ---------- נרמול ----------
function ToDate($v) {
  if ([string]::IsNullOrWhiteSpace($v)) { return $null }
  $v = "$v".Trim()
  if ($v -match '^\d{4,5}(\.\d+)?$') { return [DateTime]::FromOADate([double]$v) }
  if ($v -match '^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})(?:[ T]+(\d{1,2}):(\d{2}))?') {
    $a = [int]$matches[1]; $b = [int]$matches[2]; $y = [int]$matches[3]; if ($y -lt 100) { $y += 2000 }
    $hh = if ($matches[4]) { [int]$matches[4] } else { 0 }; $mm = if ($matches[5]) { [int]$matches[5] } else { 0 }
    if ($v -match '[AP]M') { $t = $a; $a = $b; $b = $t; if ($v -match 'PM' -and $hh -lt 12) { $hh += 12 } }  # פורמט אמריקאי ישן
    try { return New-Object DateTime($y, $b, $a, $hh, $mm, 0) } catch { return $null }
  }
  return $null
}
$sha = [Security.Cryptography.SHA256]::Create()
function Hash([string]$s) { if (-not $s) { return '' }; (-join ($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes('slavso|' + $s)) | Select-Object -First 6 | ForEach-Object { $_.ToString('x2') })) }
function PhoneKey($v) {
  if ([string]::IsNullOrWhiteSpace($v)) { return '' }
  $v = "$v".Trim()
  if ($v -match '^[\d.]+E\+?\d+$') { $v = ([decimal][double]::Parse($v, [Globalization.CultureInfo]::InvariantCulture)).ToString('0') }
  $d = $v -replace '\D', '' -replace '^972', '' -replace '^0+', ''
  if ($d.Length -lt 8 -or $d -match '^5?0{6,}$') { return '' }
  Hash ('p' + $d)
}
function MailKey($v) { $m = "$v".Trim().ToLower(); if ($m -notmatch '^[^@\s]+@[^@\s]+\.[a-z]{2,}$') { return '' }; Hash ('m' + $m) }
function IsTest($r, [string[]]$fields) { foreach ($f in $fields) { $x = "$($r.$f)"; if ($x -match 'טסט|test|yameliautomation') { return $true } }; $false }
function Num($v) { $x = "$v" -replace '[^\d.\-]', ''; if ($x -eq '' -or $x -eq '-' -or $x -eq '.') { 0 } else { [double]$x } }
function Iso($d) { $d.ToString('yyyy-MM-ddTHH:mm') }
function Called($v) { "$v".Trim() -match '^1(\.0)?$' }

# ---------- נרשמים לוובינר ----------
$regs = foreach ($r in (LoadCsv (TabFile 'נרשמים'))) {
  $d = ToDate $r.'תאריך הוספה'; if (-not $d -or $d -lt $attribFrom) { continue }
  if (IsTest $r 'שם פרטי', 'מייל', 'utm_source') { continue }
  [ordered]@{ d = Iso $d; kp = PhoneKey $r.'טלפון'; km = MailKey $r.'מייל'
    s = "$($r.utm_source)".Trim(); c = "$($r.utm_content)".Trim(); t = "$($r.utm_term)".Trim(); a = "$($r.utm_ad)".Trim() }
}

# ---------- משתתפי וובינר ----------
function Attendees($tab) {
  foreach ($r in (LoadCsv (TabFile $tab))) {
    $d = ToDate $r.'תאריך'; if (-not $d -or $d -lt $fromDate) { continue }
    [ordered]@{ d = Iso $d; kp = PhoneKey $r.'טלפון'; km = MailKey $r.'מייל'; call = [bool](Called $r.'שיחה') }
  }
}
$deep = @(Attendees 'הגיע 35')
$partial = @(Attendees 'הגיע פחות מ 35')

# ---------- לידים אורגניים ושיתופי פעולה ----------
$partnerRx = 'HTZONE|הייטקזון|היי טק זון|BUYME|ביימי|באיימי|בסלון|תוכנית עמית|תכנית עמית'
function PartnerName([string]$s) {
  if ($s -match 'HTZONE|הייטקזון|היי טק זון') { 'HTZONE' } elseif ($s -match 'BUYME|ביימי|באיימי') { 'BUYME' }
  elseif ($s -match 'בסלון') { 'בסלון' } elseif ($s -match 'תוכנית עמית|תכנית עמית') { 'תוכנית עמית' } else { '' }
}
$organic = foreach ($r in (LoadCsv (TabFile 'מסר 1'))) {
  $d = ToDate $r.'תאריך הוספה'; if (-not $d -or $d -lt $fromDate) { continue }
  if (IsTest $r 'שם מלא', 'מייל') { continue }
  $blob = "$($r.'מקור הגעה') $($r.'סטטוס') $($r.'Form Name') $($r.utm_source) $($r.utm_campaign)"
  $ch = if ($blob -match $partnerRx) { 'partner' } elseif ($blob -match 'gclid|gad_source|google') { 'google' } else { 'organic' }
  [ordered]@{ d = Iso $d; kp = PhoneKey $r.'טלפון'; km = MailKey $r.'מייל'; ch = $ch; p = (PartnerName $blob)
    src = "$($r.'מקור הגעה')".Trim(); call = [bool](Called $r.'שיחה') }
}

# ---------- רכישות ----------
$purch = foreach ($r in (LoadCsv (TabFile 'רכשו'))) {
  $d = ToDate $r.'תאריך'; if (-not $d -or $d -lt $fromDate) { continue }
  $amount = Num $r.'סכום'; if ($amount -le 0) { continue }
  [ordered]@{ d = Iso $d; kp = PhoneKey $r.'טלפון'; km = MailKey $r.'מייל'
    src = "$($r.'מקור הליד')".Trim(); s = "$($r.utm_source)".Trim(); c = "$($r.utm_content)".Trim(); t = "$($r.utm_term)".Trim()
    prod = "$($r.'מוצר שנרכש')".Trim(); amt = $amount; kit = Num $r.'עלות משלוח + ערכת חומרים'; pay = "$($r.'סליקה אוטומטית?')".Trim() }
}

# ---------- ספר הוצאות חודשי (טאב דוח חודשי) ----------
$monthNames = @{ 'ינואר' = 1; 'פברואר' = 2; 'מרץ' = 3; 'אפריל' = 4; 'מאי' = 5; 'יוני' = 6; 'יולי' = 7; 'אוגוסט' = 8; 'ספט' = 9; 'ספטמבר' = 9; 'אוק' = 10; 'אוקטובר' = 10; 'נוב' = 11; 'נובמבר' = 11; 'דצמ' = 12; 'דצמבר' = 12 }
$ledgerLines = [IO.File]::ReadAllLines((TabFile 'דוח חודשי' -Prefix), [Text.Encoding]::UTF8) | ForEach-Object { , @($_ | ConvertFrom-Csv -Header (1..60 | ForEach-Object { "c$_" }) | ForEach-Object { $o = $_; 1..60 | ForEach-Object { "$($o."c$_")".Trim() } }) }
$hdrRow = $ledgerLines | Where-Object { $_[0] -eq 'חודשי' } | Select-Object -First 1
$monthCols = @{}
for ($i = 1; $i -lt $hdrRow.Count; $i++) {
  if ($hdrRow[$i] -match '^(\S+)\s+(\d{2,4})$' -and $monthNames.ContainsKey($matches[1])) {
    $y = [int]$matches[2]; if ($y -lt 100) { $y += 2000 }
    $key = '{0:0000}-{1:00}' -f $y, $monthNames[$matches[1]]
    if ($key -ge $fromDate.ToString('yyyy-MM') -and $key -le (Get-Date).ToString('yyyy-MM')) { $monthCols[$key] = $i }
  }
}
$ledger = [ordered]@{}; $inCosts = $false
foreach ($row in $ledgerLines) {
  $label = $row[0]
  if ($label -eq 'עלות ליד') { $inCosts = $true; continue }
  if ($label -like 'יחס הכנסות*') { break }
  if ($label -in 'מכירות', 'עלות קמפיין פייסבוק', 'עלות קמפיין גוגל', 'סהכ עלויות' -or ($inCosts -and $label)) {
    $vals = [ordered]@{}; foreach ($m in ($monthCols.Keys | Sort-Object)) { $vals[$m] = [math]::Round((Num $row[$monthCols[$m]]), 2) }
    $ledger[$label] = $vals
  }
}

# ---------- דוח Meta ----------
$metaLines = [IO.File]::ReadAllLines($MetaCsv, [Text.Encoding]::UTF8)
$mh = ($metaLines[0] | ConvertFrom-Csv -Header (1..200 | ForEach-Object { "c$_" })).PSObject.Properties | Where-Object { $_.Value -ne $null } | ForEach-Object { $_.Value }
$seen = @{}; $mnames = foreach ($h in $mh) { if ($seen[$h]) { $seen[$h]++; "$h`_$($seen[$h])" } else { $seen[$h] = 1; $h } }
$metaRows = $metaLines[1..($metaLines.Count - 1)] | Where-Object { $_.Trim() } | ConvertFrom-Csv -Header $mnames
$dayCol = @('Day', 'Date') | Where-Object { $mnames -contains $_ } | Select-Object -First 1
$ads = foreach ($r in $metaRows) {
  [ordered]@{ name = ("$($r.'Ad name')" -replace '[‎‏]', '').Trim(); adset = ("$($r.'Ad set name')" -replace '[‎‏]', '').Trim()
    day = if ($dayCol) { "$($r.$dayCol)" } else { '' }; delivery = "$($r.'Ad delivery')"
    spend = Num $r.'Amount spent (USD)'; results = Num $r.Results; impr = Num $r.Impressions; reach = Num $r.Reach
    clicks = Num $r.'Clicks (all)'; ctr = Num $r.'CTR (all)'; cpm = Num $r.'CPM (cost per 1,000 impressions) (USD)'; freq = Num $r.Frequency
    quality = "$($r.'Quality ranking')"; engage = "$($r.'Engagement rate ranking')"; conv = "$($r.'Conversion rate ranking')" }
}
$metaStart = ($metaRows | Select-Object -First 1).'Reporting starts'; $metaEnd = ($metaRows | Select-Object -First 1).'Reporting ends'

# ---------- בדיקות שפיות ----------
$lastPurchase = ($purch | ForEach-Object { $_.d } | Sort-Object | Select-Object -Last 1)
if ($lastPurchase -and ([datetime]$metaEnd - [datetime]$lastPurchase.Substring(0, 10)).TotalDays -gt 30) { $warnings.Add("הרכישה האחרונה בטאב רכשו היא מ-$($lastPurchase.Substring(0,10)), יותר מ-30 יום לפני סוף דוח Meta. ייתכן שהנתונים חתוכים.") }
$regsInRange = @($regs | Where-Object { $_.d -ge $metaStart -and $_.d.Substring(0, 10) -le $metaEnd -and ($_.s -match 'facebook|^fb$|^ig$|instagram') }).Count
$metaResults = ($ads | ForEach-Object { $_.results } | Measure-Object -Sum).Sum
if ($metaResults -gt 0 -and ($regsInRange * 2 -lt $metaResults -or $metaResults * 2 -lt $regsInRange)) { $warnings.Add("פער חריג: $regsInRange הרשמות ממטא בגיליון מול $metaResults תוצאות בדוח Meta.") }
if (-not $monthCols.Count) { $warnings.Add('לא נמצאו עמודות חודשים בטאב דוח חודשי. ההוצאות לא יוצגו.') }
$noKey = @($purch | Where-Object { -not $_.kp -and -not $_.km }).Count
if ($noKey) { $warnings.Add("$noKey רכישות בלי טלפון או מייל תקינים. הן ישויכו לפי שדות הרכישה בלבד.") }

$data = [ordered]@{
  generated = (Get-Date).ToString('yyyy-MM-ddTHH:mm'); from = $From
  meta = [ordered]@{ start = $metaStart; end = $metaEnd; file = (Split-Path $MetaCsv -Leaf); byDay = [bool]$dayCol; ads = @($ads) }
  regs = @($regs); deep = $deep; partial = $partial; organic = @($organic); purchases = @($purch)
  ledger = $ledger; warnings = @($warnings)
}
$json = $data | ConvertTo-Json -Depth 6 -Compress
$json = $json -replace '</', '<\/'
$html = [IO.File]::ReadAllText($Template, [Text.Encoding]::UTF8).Replace('/*__DATA__*/null', $json)
[IO.File]::WriteAllText($Out, $html, (New-Object Text.UTF8Encoding($false)))
"OK: regs=$(@($regs).Count) deep=$($deep.Count) partial=$($partial.Count) organic=$(@($organic).Count) purchases=$(@($purch).Count) ads=$(@($ads).Count) months=$($monthCols.Count)"
$warnings | ForEach-Object { "WARN: $_" }
