function Test-PCXLabEnvironment {

    param(
        [Parameter(Mandatory)]
        [string]$InputFolder,

        [string]$OutputFolder
    )

    Write-Log "Starting environment validation..."

    # 🔹 Check input folder
    if (-not (Test-Path $InputFolder)) {
        Write-Log "Input folder does not exist: $InputFolder" "ERROR"
        throw "Invalid input folder"
    }
    else {
        Write-Log "Input folder OK: $InputFolder" "SUCCESS"
    }

    # 🔹 Output folder
    if (-not $OutputFolder) {
        $OutputFolder = $InputFolder
    }

    if (!(Test-Path $OutputFolder)) {
        try {
            New-Item -ItemType Directory -Path $OutputFolder -ErrorAction Stop | Out-Null
            Write-Log "Created output folder: $OutputFolder" "SUCCESS"
        }
        catch {
            Write-Log "Cannot create output folder: $OutputFolder" "ERROR"
            throw
        }
    }
    else {
        Write-Log "Output folder OK: $OutputFolder" "SUCCESS"
    }

    # 🔹 NuGet check
    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Write-Log "NuGet not found. Installing..."

        try {
            Install-PackageProvider -Name NuGet `
                -MinimumVersion 2.8.5.201 `
                -Force `
                -Scope CurrentUser `
                -ErrorAction Stop

            Write-Log "NuGet installed successfully." "SUCCESS"
        }
        catch {
            Write-Log "NuGet installation failed: $($_.Exception.Message)" "ERROR"
            throw
        }
    }
    else {
        Write-Log "NuGet already available" "SUCCESS"
    }

    # 🔹 ImportExcel check
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Log "ImportExcel not found. Installing..."

        try {
            Install-Module ImportExcel `
                -Scope CurrentUser `
                -Force `
                -AllowClobber `
                -ErrorAction Stop

            Write-Log "ImportExcel installed successfully." "SUCCESS"
        }
        catch {
            Write-Log "ImportExcel install failed: $($_.Exception.Message)" "ERROR"
            throw
        }
    }
    else {
        Write-Log "ImportExcel available" "SUCCESS"
    }

    # 🔹 Excel availability (optional warning only)
    try {
        $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
        $excel.Quit()
        Write-Log "Excel COM available" "SUCCESS"
    }
    catch {
        Write-Log "Excel not installed or COM not available (XLS conversion may fail)" "WARNING"
    }

    Write-Log "Environment validation completed." "SUCCESS"
}