param(
    [Parameter(Mandatory)]
    [string]$Folder,

    [string]$OutputFolder
)

# 🔹 Resolve module path
$basePath = Split-Path $PSScriptRoot -Parent
$modulePath = Join-Path $basePath "Modules"
$env:PSModulePath = "$modulePath;$env:PSModulePath"

# 🔹 Import modules
Import-Module PCXLab.Core -Force
Import-Module PCXLab.Excel -Force

# 🔹 Validate input folder
if (-not (Test-Path $Folder)) {
    throw "Input folder does not exist: $Folder"
}

# 🔹 Start logging
Start-LogSession -LogFolder (Join-Path $Folder "logs")

# 🔹 Environment check
Test-PCXLabEnvironment -InputFolder $Folder -OutputFolder $OutputFolder

# 🔹 Default output folder
if (-not $OutputFolder) {
    $OutputFolder = $Folder
}

# 🔹 Get all files
$files = Get-ChildItem $Folder -File

foreach ($file in $files) {

    # 🔹 Skip already processed files
    if ($file.Name -match "_ConvertedFromXls" -or $file.Name -match "_Transformed") {
        Write-Log "Skipping: $($file.Name)"
        continue
    }

    # 🔹 Process only supported files
    if ($file.Extension -notin ".xls", ".xlsx") {
        Write-Log "Skipping unsupported file: $($file.Name)" "WARNING"
        continue
    }

    Write-Log "Processing: $($file.Name)"

    try {

        # 🔥 IMPORTANT → reset result
        $result = $null

        # 🔹 Convert only if needed
        if ($file.Extension -eq ".xls") {
            $workingFile = Convert-XlsToXlsx -File $file
        }
        else {
            $workingFile = $file
        }

        # 🔹 Detect bank (use original filename)
        $bank = Get-BankFromFile -File $file

        switch ($bank) {

            "ICICI" {

                if ($file.Name -match "_DC_") {
                    $result = Convert-ICICIDCFormat -File $workingFile
                }
                else {
                    $result = Convert-ICICIFormat -File $workingFile
                }
            }

            "HDFC" {

                if ($file.Name -match "_DC_") {
                    $result = Convert-HDFCDCFormat -File $workingFile
                }
                else {
                    $result = Convert-HDFCFormat -File $workingFile
                }
            }

            default {
                Write-Log "Unknown bank format: $($file.Name)" "WARNING"
                continue
            }
        }

        # 🔥 Safety check
        if (-not $result) {
            Write-Log "No result generated for $($file.Name)" "ERROR"
            continue
        }

        # 🔹 Output filename
        $outFileName = Get-OutputFileName `
            -File $file `
            -Converted:$($file.Extension -eq ".xls") `
            -Transformed

        $outFile = Join-Path $OutputFolder $outFileName

        # 🔥 CRITICAL FIX → prevent 1E+22
        $result | Export-Excel `
            -Path $outFile `
            -AutoSize `
            -BoldTopRow `
            -NoNumberConversion "Chq./Ref.No."

        Write-Log "Saved: $outFile" "SUCCESS"
    }
    catch {
        Write-Log "Error processing $($file.Name): $($_.Exception.Message)" "ERROR"
    }
}