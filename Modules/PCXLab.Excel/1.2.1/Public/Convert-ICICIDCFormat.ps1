function Convert-ICICIDCFormat {

    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File
    )

    # 🔹 Convert if needed
    $workingFile = Convert-XlsToXlsx -File $File

    # 🔹 MOP
    $mop = Get-MOPFromFileName -FileName $workingFile.Name

    # 🔹 Read Excel
    $raw = Import-Excel $workingFile.FullName -NoHeader

    # 🔹 Find header row
    $headerIndex = $null

    for ($i = 0; $i -lt $raw.Count; $i++) {

        $rowText = ($raw[$i].PSObject.Properties.Value -join " ")

        if ($rowText -match "Value Date" -and $rowText -match "Transaction Date") {
            $headerIndex = $i
            break
        }
    }

    if ($null -eq $headerIndex) {
        throw "ICICI DC header not found in $($File.Name)"
    }

    $headerRow = $raw[$headerIndex].PSObject.Properties.Value

    # 🔹 Column mapping
    $colMap = @{
        ValueDate = ($headerRow | Where-Object { $_ -match "Value Date" } | ForEach-Object { [array]::IndexOf($headerRow, $_) })
        TransDate = ($headerRow | Where-Object { $_ -match "Transaction Date" } | ForEach-Object { [array]::IndexOf($headerRow, $_) })
        Ref       = ($headerRow | Where-Object { $_ -match "Cheque" } | ForEach-Object { [array]::IndexOf($headerRow, $_) })
        Details   = ($headerRow | Where-Object { $_ -match "Remarks" } | ForEach-Object { [array]::IndexOf($headerRow, $_) })
        Withdraw  = ($headerRow | Where-Object { $_ -match "Withdrawal" } | ForEach-Object { [array]::IndexOf($headerRow, $_) })
        Deposit   = ($headerRow | Where-Object { $_ -match "Deposit" } | ForEach-Object { [array]::IndexOf($headerRow, $_) })
    }

    $data = $raw[($headerIndex + 1)..($raw.Count - 1)]

    foreach ($row in $data) {

        $values = $row.PSObject.Properties.Value

        $valueDate = $values[$colMap.ValueDate]
        $details   = $values[$colMap.Details]
        $ref       = $values[$colMap.Ref]
        $withdraw  = $values[$colMap.Withdraw]
        $deposit   = $values[$colMap.Deposit]

        if (-not $valueDate -or -not $details) { continue }

        # 🔹 Format date
        try {
            $date = [datetime]::ParseExact($valueDate, "dd/MM/yyyy", $null).ToString("dd-MM-yyyy")
        }
        catch {
            continue
        }

        # 🔹 Clean amounts
        $withdraw = ($withdraw -replace ",", "")
        $deposit  = ($deposit -replace ",", "")

        $amtDr = if ($withdraw -and $withdraw -ne "0.00") { [decimal]$withdraw } else { 0 }
        $amtCr = if ($deposit  -and $deposit  -ne "0.00") { [decimal]$deposit } else { 0 }

        # 🔹 Force REF as text
        if ($ref) {
            $ref = "'$ref"
        }

        # 🔹 Clean narration
        $cleanDetails = $details.Trim()

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
            "Value Dt"     = $date   # ✅ actual value date
            "Chq./Ref.No." = $ref
        }
    }
}