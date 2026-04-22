function Convert-HDFCDCFormat {

    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File
    )

    # 🔹 Convert if needed
    $workingFile = Convert-XlsToXlsx -File $File

    # 🔹 MOP
    $mop = Get-MOPFromFileName -FileName $workingFile.Name

    # 🔹 Read file
    $raw = Import-Excel $workingFile.FullName -NoHeader

    # 🔹 Find header row (Date | Narration | Chq...)
    $headerIndex = $null

    for ($i = 0; $i -lt $raw.Count; $i++) {
        $rowText = ($raw[$i].PSObject.Properties.Value -join " ")

        if ($rowText -match "Date" -and $rowText -match "Narration") {
            $headerIndex = $i
            break
        }
    }

    if ($null -eq $headerIndex) {
        throw "HDFC DC header not found in $($File.Name)"
    }

    $headerRow = $raw[$headerIndex].PSObject.Properties.Value

    # 🔹 Column mapping
    $colMap = @{
        Date       = ($headerRow | Where-Object { $_ -match "^Date" } | ForEach-Object { [array]::IndexOf($headerRow, $_) })
        Narration  = ($headerRow | Where-Object { $_ -match "Narration" } | ForEach-Object { [array]::IndexOf($headerRow, $_) })
        Ref        = ($headerRow | Where-Object { $_ -match "Chq" } | ForEach-Object { [array]::IndexOf($headerRow, $_) })
        Withdraw   = ($headerRow | Where-Object { $_ -match "Withdrawal" } | ForEach-Object { [array]::IndexOf($headerRow, $_) })
        Deposit    = ($headerRow | Where-Object { $_ -match "Deposit" } | ForEach-Object { [array]::IndexOf($headerRow, $_) })
    }

    $data = $raw[($headerIndex + 1)..($raw.Count - 1)]

    foreach ($row in $data) {

        $values = $row.PSObject.Properties.Value

        $date      = $values[$colMap.Date]
        $narration = $values[$colMap.Narration]
        $ref       = $values[$colMap.Ref]
        $withdraw  = $values[$colMap.Withdraw]
        $deposit   = $values[$colMap.Deposit]

        if (-not $date -or -not $narration) { continue }

        # 🔹 Convert date (01/03/26 → 01-03-2026)
        try {
            $date = [datetime]::ParseExact($date, "dd/MM/yy", $null).ToString("dd-MM-yyyy")
        }
        catch {
            continue
        }

        # 🔹 Clean amounts
        $withdraw = ($withdraw -replace ",", "")
        $deposit  = ($deposit -replace ",", "")

        $amtDr = if ($withdraw) { [decimal]$withdraw } else { 0 }
        $amtCr = if ($deposit)  { [decimal]$deposit } else { 0 }

        # 🔹 Keep ref as string (avoid 1E+ issue)
        if ($ref) {
            $ref = "'$ref"   # forces Excel text format
        }

        [PSCustomObject]@{
            Date           = $date
            Narration      = $narration
            Item           = ""
            Category       = ""
            Place          = ""
            Freq           = ""
            For            = ""
            MOP            = $mop
            "Amt (Dr)"     = $amtDr
            "Amt (Cr)"     = $amtCr
            "Value Dt"     = ""
            "Chq./Ref.No." = $ref
        }
    }
}