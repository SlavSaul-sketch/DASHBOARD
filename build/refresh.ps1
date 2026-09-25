param(
  [Parameter(Mandatory)] [string]$WebinarExport,  # קובץ JSON של download_file_content לגיליון "וובינר אוטומטי - קובץ לידים" (שדה content ב-base64)
  [Parameter(Mandatory)] [string]$OrganicExport,  # קובץ JSON של download_file_content לגיליון "[אוטומציה] קורס דיגיטלי Slav&So - ליד חדש"
  [string]$MetaJson,                               # תשובת Airtable (list_records_for_table) שמורה כקובץ JSON
  [string]$MetaCsv,                                # דוח Meta מ-Ads Manager. ברירת מחדל: הקובץ החדש ביותר *Ads*.csv בתיקיית ההורדות
  [switch]$NoPush                                  # לבנות בלי commit ו-push
)
# רענון יומי מקצה לקצה, החלק המקומי: פענוח, המרה, בנייה, ניקוי, commit ו-push.
# לא כותב לגיליונות ולא ל-Airtable. הקבצים הגולמיים נשמרים רק בתיקייה זמנית ונמחקים בסוף.
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot
$tmp = Join-Path $env:TEMP 'slavso-dashboard'
Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force $tmp | Out-Null
try {
  # 1. פענוח הייצוא (base64) לקובצי xlsx
  foreach ($p in @(@{ f = $WebinarExport; n = 'webinar' }, @{ f = $OrganicExport; n = 'organic' })) {
    $j = [IO.File]::ReadAllText($p.f, [Text.Encoding]::UTF8) | ConvertFrom-Json
    if (-not $j.content) { throw "בקובץ $($p.f) אין שדה content. ייתכן שההורדה נכשלה." }
    [IO.File]::WriteAllBytes((Join-Path $tmp "$($p.n).xlsx"), [Convert]::FromBase64String($j.content))
  }
  # 2. המרה ל-CSV לכל טאב
  & (Join-Path $PSScriptRoot 'xlsx2csv.ps1') -Xlsx (Join-Path $tmp 'webinar.xlsx') -OutDir (Join-Path $tmp 'csv') -Prefix webinar | Out-Null
  & (Join-Path $PSScriptRoot 'xlsx2csv.ps1') -Xlsx (Join-Path $tmp 'organic.xlsx') -OutDir (Join-Path $tmp 'csv') -Prefix organic | Out-Null
  # 3. בחירת מקור מטא: העדכני מבין Airtable ודוח ה-CSV, לפי תאריך סוף הדוח
  if (-not $MetaCsv) { $MetaCsv = Get-ChildItem (Join-Path $env:USERPROFILE 'Downloads') -Filter '*Ads*.csv' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName }
  $airEnd = ''; if ($MetaJson -and (Test-Path $MetaJson)) { $airEnd = "$(([IO.File]::ReadAllText($MetaJson, [Text.Encoding]::UTF8) | ConvertFrom-Json).records | ForEach-Object { "$($_.cellValuesByFieldId.fldg3FFYMPtYeyWfP)" } | Sort-Object | Select-Object -Last 1)"; if ($airEnd.Length -gt 10) { $airEnd = $airEnd.Substring(0, 10) } }
  $csvEnd = ''; if ($MetaCsv -and (Test-Path $MetaCsv)) { $csvEnd = "$(([IO.File]::ReadAllLines($MetaCsv, [Text.Encoding]::UTF8) | Select-Object -Skip 1 -First 1 | ConvertFrom-Csv -Header 's','e').e)" }
  if (-not $airEnd -and -not $csvEnd) { throw 'אין נתוני מטא: לא התקבלה תשובת Airtable ולא נמצא דוח CSV בתיקיית ההורדות.' }
  $build = @{ CsvDir = (Join-Path $tmp 'csv') }
  if ($csvEnd -gt $airEnd) { $build.MetaCsv = $MetaCsv; "META: דוח CSV עד $csvEnd (Airtable עד $(if ($airEnd) { $airEnd } else { '-' })) - $(Split-Path $MetaCsv -Leaf)" }
  else { $build.MetaJson = $MetaJson; "META: Airtable עד $airEnd (CSV עד $(if ($csvEnd) { $csvEnd } else { '-' }))" }
  # 4. בנייה
  & (Join-Path $PSScriptRoot 'build-data.ps1') @build
}
finally {
  # 5. ניקוי: הקבצים הגולמיים מכילים פרטים אישיים
  Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
  foreach ($f in $WebinarExport, $OrganicExport, $MetaJson) { if ($f -and $f -notlike "$repo*") { Remove-Item $f -Force -ErrorAction SilentlyContinue } }
}

if ($NoPush) { return }

# 6. commit ו-push ל-main, רק של קובץ הדאשבורד
Push-Location $repo
try {
  # סנכרון הדרייב משאיר לפעמים קובץ נעילה של git. מוחקים אותו רק כשאין תהליך git פעיל.
  if (-not (Get-Process git -ErrorAction SilentlyContinue)) { Get-ChildItem .git -Filter '*.lock' -Recurse -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue }
  $branch = (git branch --show-current).Trim()
  if ($branch -ne 'main') { throw "הריפו נמצא על הענף '$branch' ולא על main. ה-push בוטל." }
  git pull --ff-only --quiet
  git add dashboard.html
  git diff --cached --quiet
  if ($LASTEXITCODE -eq 0) { 'PUSH: אין שינוי בדאשבורד מאז הרענון הקודם' ; return }
  $msgFile = Join-Path $env:TEMP 'slavso-commit-msg.txt'
  [IO.File]::WriteAllText($msgFile, "Daily refresh $((Get-Date).ToString('yyyy-MM-dd'))`n`nCo-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`n", (New-Object Text.UTF8Encoding($false)))
  git commit --quiet -F $msgFile
  Remove-Item $msgFile -Force -ErrorAction SilentlyContinue
  git push --quiet origin main
  if ($LASTEXITCODE -ne 0) { throw 'ה-push נכשל.' }
  "PUSH: $((git log --oneline -1))"
}
finally { Pop-Location }
