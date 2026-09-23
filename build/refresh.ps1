param(
  [Parameter(Mandatory)] [string]$WebinarExport,  # קובץ JSON של download_file_content לגיליון "וובינר אוטומטי - קובץ לידים" (שדה content ב-base64)
  [Parameter(Mandatory)] [string]$OrganicExport,  # קובץ JSON של download_file_content לגיליון "[אוטומציה] קורס דיגיטלי Slav&So - ליד חדש"
  [Parameter(Mandatory)] [string]$MetaJson,       # תשובת Airtable (list_records_for_table) שמורה כקובץ JSON
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
  # 3. בנייה
  & (Join-Path $PSScriptRoot 'build-data.ps1') -CsvDir (Join-Path $tmp 'csv') -MetaJson $MetaJson
}
finally {
  # 4. ניקוי: הקבצים הגולמיים מכילים פרטים אישיים
  Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
  foreach ($f in $WebinarExport, $OrganicExport, $MetaJson) { if ($f -and $f -notlike "$repo*") { Remove-Item $f -Force -ErrorAction SilentlyContinue } }
}

if ($NoPush) { return }

# 5. commit ו-push ל-main, רק של קובץ הדאשבורד
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
