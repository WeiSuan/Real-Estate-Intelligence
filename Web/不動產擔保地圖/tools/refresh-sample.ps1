param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$dataDir = Join-Path $ProjectRoot "data"
$samplePath = Join-Path $dataDir "sample-xlsx.js"
$manifestPath = Join-Path $dataDir "timeline-manifest.js"
$vacancyRefreshScript = Join-Path $PSScriptRoot "refresh-vacancy-data.py"
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)

$prefix = [string]([char]0x7E3D) + [string]([char]0x8868) + "_"
$yearChar = [string]([char]0x5E74)
$monthChar = [string]([char]0x6708)

if (!(Test-Path -LiteralPath $dataDir)) {
    throw "Missing data directory: $dataDir"
}

if (!(Test-Path -LiteralPath $vacancyRefreshScript)) {
    throw "Missing vacancy refresh script: $vacancyRefreshScript"
}

$periodFiles = @()
Get-ChildItem -LiteralPath $dataDir -File -Filter "*.xlsx" | ForEach-Object {
    if ($_.Name.StartsWith($prefix) -and $_.BaseName.Length -eq 9) {
        $period = $_.BaseName.Substring(3, 6)
        if ($period -match '^\d{6}$') {
            $periodFiles += [pscustomobject]@{
                Period = $period
                File = $_
            }
        }
    }
}
$periodFiles = @($periodFiles | Sort-Object Period)

if ($periodFiles.Count -eq 0) {
    throw "No Excel files found. Expected data\{prefix}YYYYMM.xlsx."
}

$latest = $periodFiles[-1]
$latestBytes = [System.IO.File]::ReadAllBytes($latest.File.FullName)
$latestBase64 = [Convert]::ToBase64String($latestBytes)

$datasetBlocks = @()
foreach ($item in $periodFiles) {
    $bytes = [System.IO.File]::ReadAllBytes($item.File.FullName)
    $base64 = [Convert]::ToBase64String($bytes)
    $datasetBlocks += @(
        "    {",
        "        periodId: `"$($item.Period)`",",
        "        fileName: `"$($item.File.Name)`",",
        "        base64: `"$base64`"",
        "    }"
    ) -join "`n"
}

$sampleLines = @(
    "window.DEFAULT_EXCEL_BASE64 = '$latestBase64';",
    "window.DEFAULT_EXCEL_FILE_NAME = '$($latest.File.Name)';",
    "window.DEFAULT_EXCEL_PERIOD_ID = '$($latest.Period)';",
    "window.TIMELINE_SAMPLE_DATASETS = [",
    ($datasetBlocks -join ",`n"),
    "];"
)
$sample = ($sampleLines -join "`n") + "`n"

$manifestItems = @()
foreach ($item in $periodFiles) {
    $year = $item.Period.Substring(0, 4)
    $month = $item.Period.Substring(4, 2)
    $label = $year + $yearChar + $month + $monthChar
    $manifestItems += @(
        "    {",
        "        periodId: `"$($item.Period)`",",
        "        label: `"$label`",",
        "        type: `"month`",",
        "        file: `"data/$($item.File.Name)`"",
        "    }"
    ) -join "`n"
}
$manifest = "window.TIMELINE_DATASETS = [`n" + ($manifestItems -join ",`n") + "`n];`n"

[System.IO.File]::WriteAllText($samplePath, $sample.Replace("`r", ""), $utf8NoBom)
[System.IO.File]::WriteAllText($manifestPath, $manifest.Replace("`r", ""), $utf8NoBom)

$python = Get-Command python -ErrorAction SilentlyContinue
if ($null -eq $python) {
    throw "Python is required to refresh vacancy data."
}
& $python.Source $vacancyRefreshScript
if ($LASTEXITCODE -ne 0) {
    throw "Vacancy data refresh failed."
}

Write-Host "Refreshed data/sample-xlsx.js from $($latest.File.Name)"
Write-Host "Refreshed data/timeline-manifest.js with $($periodFiles.Count) periods"
