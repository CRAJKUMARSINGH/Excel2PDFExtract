param(
    [switch]$DeleteMarkdown,
    [switch]$DeleteEmptyDirs
)

$root = 'e:\\Rajkumar\\Excel2PDFExtract'

# Gather files
$mdFiles = Get-ChildItem -Path $root -Recurse -Filter *.md -File
$txtFiles = Get-ChildItem -Path $root -Recurse -Filter *.txt -File

# Find empty directories
$emptyDirs = Get-ChildItem -Path $root -Directory -Recurse | Where-Object { -not (Get-ChildItem -Path $_.FullName -Force) }

# Find duplicate file names (same name in different folders)
$dupGroups = Get-ChildItem -Path $root -Recurse -File -Include *.md,*.txt | Group-Object Name | Where-Object { $_.Count -gt 1 }

# Build report
$report = @()
$report += "## Dry‑run report"
$report += ""
$report += "### Markdown files ($($mdFiles.Count))"
$mdFiles | ForEach-Object { $report += $_.FullName }
$report += ""
$report += "### Text files ($($txtFiles.Count))"
$txtFiles | ForEach-Object { $report += $_.FullName }
$report += ""
$report += "### Empty directories ($($emptyDirs.Count))"
$emptyDirs | ForEach-Object { $report += $_.FullName }
$report += ""
$report += "### Duplicate file names ($($dupGroups.Count))"
foreach($grp in $dupGroups){
    $report += "--- Duplicate name: $($grp.Name) (Count: $($grp.Count))"
    $grp.Group | ForEach-Object { $report += $_.FullName }
}

# Output report
$report | Out-File -FilePath 'dry_run_report.txt' -Encoding UTF8
Write-Output $report

if($DeleteMarkdown){
    Write-Host "Deleting markdown files..."
    $mdFiles | Remove-Item -Force
}
if($DeleteEmptyDirs){
    Write-Host "Deleting empty directories..."
    $emptyDirs | Remove-Item -Force -Recurse
}

