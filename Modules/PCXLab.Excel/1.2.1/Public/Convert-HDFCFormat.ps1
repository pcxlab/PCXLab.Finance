function Convert-HDFCFormat {

    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File
    )

    # 🔹 Ensure XLS → XLSX
    $workingFile = Convert-XlsToXlsx -File $File

    # 🔹 Get MOP (UPDATED FORMAT with _)
    $mop = Get-MOPFromFileName -FileName $workingFile.Name

    # 🔹 Read Excel
    $raw = Import-Excel $workingFile.FullName -NoHeader

    # 🔹 Find header row
    $headerIndex = Get-HDFCHeader -RawData $raw

    if ($null -eq $headerIndex) {
        throw "HDFC Header not found in $($File.Name)"
    }

    # 🔹 Extract header row values
    $headerRow = $raw[$headerIndex].PSObject.Properties.Value

    # 🔹 Column mapping
    $colMap = @{
        DateTime = ($headerRow | ForEach-Object { $_ } | Where-Object { $_ -match "Date" } | ForEach-Object { [array]::IndexOf($headerRow, $_) })
        Details  = ($headerRow | Where-Object { $_ -match "Description" } | ForEach-Object { [array]::IndexOf($headerRow, $_) })
        Amount   = ($headerRow | Where-Object { $_ -match "AMT" } | ForEach-Object { [array]::IndexOf($headerRow, $_) })
        DrCr     = ($headerRow | Where-Object { $_ -match "Debit" } | ForEach-Object { [array]::IndexOf($headerRow, $_) })
    }

    # 🔹 Data starts after header
    $data = $raw[($headerIndex + 1)..($raw.Count - 1)]

    foreach ($row in $data) {

        $values = $row.PSObject.Properties.Value

        $dateTime = $values[$colMap.DateTime]
        $details  = $values[$colMap.Details]
        $amount   = $values[$colMap.Amount]
        $drcr     = $values[$colMap.DrCr]

        # 🔹 Skip empty rows
        if (-not $dateTime -or -not $details) { continue }

        # 🔹 Extract DATE (ignore time)
        if ($dateTime -match "(\d{2}/\d{2}/\d{4})") {
            $date = [datetime]::ParseExact($matches[1], "dd/MM/yyyy", $null).ToString("dd-MM-yyyy")
        }
        else {
            continue
        }

        # 🔹 Clean amount (remove commas)
        $amount = ($amount -replace ",", "")

        # 🔹 Convert to numeric safely
        $amtDr = 0
        $amtCr = 0

        if ($drcr -match "Cr") {
            $amtCr = [decimal]$amount
        }
        else {
            $amtDr = [decimal]$amount
        }

        # 🔹 Extract REF# from narration (IMPORTANT FIX)
        $ref = ""

        if ($details -match "Ref#\s*([A-Za-z0-9]+)") {
            $ref = "Ref# " + $matches[1]
        }

        # 🔹 Clean narration (remove ref part)
        $cleanDetails = ($details -replace "\s*\(Ref#.*?\)", "").Trim()

        # 🔹 Output object
        [PSCustomObject]@{
            Date           = $date
            Narration      = $cleanDetails
            Item           = ""
            Category       = ""
            Place          = ""
            Freq           = ""
            For            = ""
            MOP            = $mop
            "Amt (Dr)"     = $amtDr
            "Amt (Cr)"     = $amtCr
            "Value Dt"     = if ($amtDr -gt 0) { "Dr." } elseif ($amtCr -gt 0) { "Cr." };
            "Chq./Ref.No." = $ref
        }
    }
}