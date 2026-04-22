function Get-BankFromFile {

    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File
    )

    # 🔹 Fast detection using filename
    $name = $File.Name.ToUpper()

    if ($name -match "ICICI") { return "ICICI" }
    if ($name -match "HDFC")  { return "HDFC" }

    # 🔹 Fallback (content-based detection)
    try {
        $workingFile = Convert-XlsToXlsx -File $File
        $raw = Import-Excel $workingFile.FullName -NoHeader

        $scan = $raw[0..20]

        foreach ($row in $scan) {
            $text = ($row.PSObject.Properties.Value -join " ").ToUpper()

            if ($text -match "ICICI") { return "ICICI" }
            if ($text -match "HDFC")  { return "HDFC" }
        }
    }
    catch {
        # ignore
    }

    return "UNKNOWN"
}