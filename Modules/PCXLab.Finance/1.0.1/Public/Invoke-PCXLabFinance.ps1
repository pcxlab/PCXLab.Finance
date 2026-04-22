function Invoke-PCXLabFinance {

    param(
        [Parameter(Mandatory)]
        [string]$Folder,

        [string]$OutputFolder
    )

    # Go up to Automation root
    $automationRoot = Split-Path $PSScriptRoot -Parent      # 1.0.0
    $automationRoot = Split-Path $automationRoot -Parent    # PCXLab.Finance
    $automationRoot = Split-Path $automationRoot -Parent    # Modules
    $automationRoot = Split-Path $automationRoot -Parent    # Automation

    $mainScript = Join-Path $automationRoot "PCXLab.Finance\main.ps1"

    if (-not (Test-Path $mainScript)) {
        throw "main.ps1 not found at: $mainScript"
    }

    # Run in same scope
    . $mainScript -Folder $Folder -OutputFolder $OutputFolder
}