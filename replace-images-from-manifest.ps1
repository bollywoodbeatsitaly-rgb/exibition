param(
  [string]$ManifestPath = ".\image-manifest.csv",
  [string]$ImagesDir = ".\imgs",
  [switch]$DryRun
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $ManifestPath)) {
  throw "Manifest not found: $ManifestPath"
}

if (-not (Test-Path -LiteralPath $ImagesDir)) {
  New-Item -ItemType Directory -Path $ImagesDir | Out-Null
}

$rows = Import-Csv -Path $ManifestPath
if (-not $rows -or $rows.Count -eq 0) {
  throw "Manifest is empty: $ManifestPath"
}

$downloaded = 0
$skipped = 0
$failed = 0

foreach ($row in $rows) {
  $filename = $row.filename
  $url = $row.source_url

  if ([string]::IsNullOrWhiteSpace($filename)) {
    Write-Host "SKIP row with empty filename" -ForegroundColor Yellow
    $skipped++
    continue
  }

  if ([string]::IsNullOrWhiteSpace($url)) {
    Write-Host "SKIP $filename (no source_url)" -ForegroundColor Yellow
    $skipped++
    continue
  }

  $dest = Join-Path $ImagesDir $filename
  Write-Host "FETCH $filename <- $url" -ForegroundColor Cyan

  if ($DryRun) {
    continue
  }

  try {
    Invoke-WebRequest -Uri $url -OutFile $dest -TimeoutSec 120
    $downloaded++
  } catch {
    Write-Host "FAIL $filename :: $($_.Exception.Message)" -ForegroundColor Red
    $failed++
  }
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host "Downloaded: $downloaded"
Write-Host "Skipped:    $skipped"
Write-Host "Failed:     $failed"

if (-not $DryRun) {
  Write-Host ""
  Write-Host "Tip: open http://127.0.0.1:5500/index.html after running your local server."
}
