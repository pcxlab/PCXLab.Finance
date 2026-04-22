function Get-OutputFileName {
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File,

        [switch]$Converted,
        [switch]$Transformed
    )

    $baseName = $File.BaseName

    # Remove old suffixes
    $baseName = $baseName -replace "_ConvertedFromXls", ""
    $baseName = $baseName -replace "_Transformed", ""

    $suffix = ""

    if ($Converted) {
        $suffix += "_ConvertedFromXls"
    }

    if ($Transformed) {
        $suffix += "_Transformed"
    }

    return "$baseName$suffix.xlsx"
}