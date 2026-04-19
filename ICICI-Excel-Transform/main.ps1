param(
    [Parameter(Mandatory)]
    [string]$Folder,

    [string]$OutputFolder
)

# 🔹 Resolve module path dynamically
$basePath = Split-Path $PSScriptRoot -Parent
$modulePath = Join-Path $basePath "Modules"

$env:PSModulePath = "$modulePath;$env:PSModulePath"

# 🔹 Import modules
Import-Module PCXLab.Core -Force
Import-Module PCXLab.Excel -Force

# 🔹 Basic check (only this one allowed before logging)
if (-not (Test-Path $Folder)) {
    throw "Input folder does not exist: $Folder"
}

# 🔹 Start logging
Start-LogSession -LogFolder (Join-Path $Folder "logs")

# 🔹 Pre-check framework (handles everything)
Test-PCXLabEnvironment -InputFolder $Folder -OutputFolder $OutputFolder

#  🔥 FIX
if (-not $OutputFolder) {
    $OutputFolder = $Folder
}

# 🔹 Process files
$files = Get-ChildItem $Folder -File

foreach ($file in $files) {

    if ($file.Name -match "_ConvertedFromXls" -or $file.Name -match "_Transformed") {
        Write-Log "Skipping: $($file.Name)"
        continue
    }

    Write-Log "Processing: $($file.Name)"

    try {
        $workingFile = Convert-XlsToXlsx -File $file
        $result = Convert-ICICIFormat -File $workingFile

        $outFileName = Get-OutputFileName `
            -File $file `
            -Converted:$($file.Extension -eq ".xls") `
            -Transformed

        $outFile = Join-Path $OutputFolder $outFileName

        $result | Export-Excel -Path $outFile -AutoSize -BoldTopRow

        Write-Log "Saved: $outFile" "SUCCESS"
    }
    catch {
        #Write-Host "Error processing $($file.Name): $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Error processing $($file.Name): $($_.Exception.Message)" "ERROR"
    }
}