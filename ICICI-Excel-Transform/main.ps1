param(
    [string]$InputFolder  = "C:\TEST",
    [string]$OutputFolder = "C:\TEST\Output"
)

# Add module path (temporary for now)
$env:PSModulePath += ";C:\Projects\Automation\Modules"

Import-Module PCXLab.Excel -Force

# Ensure output folder exists
if (!(Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

$files = Get-ChildItem $InputFolder -Filter *.xls*

foreach ($file in $files) {

    Write-Host "Processing: $($file.Name)" -ForegroundColor Cyan

    try {
        $result = Convert-ICICIFormat -File $file

        $outFile = Join-Path $OutputFolder ($file.BaseName + "_Transformed.xlsx")

        $result | Export-Excel -Path $outFile -AutoSize -BoldTopRow

        Write-Host "Saved: $outFile" -ForegroundColor Green
    }
    catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}