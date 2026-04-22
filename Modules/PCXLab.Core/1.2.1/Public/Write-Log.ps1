function Write-Log {

    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [string]$Level = "INFO"
    )

    if (-not $global:PCXLabLogFile) {
        throw "Log session not initialized. Use Start-LogSession first."
    }

    $time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $line = "$time [$Level] $Message"

    Add-Content -Path $global:PCXLabLogFile -Value $line

    Write-Host $line
}