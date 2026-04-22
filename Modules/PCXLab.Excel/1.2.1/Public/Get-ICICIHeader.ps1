function Get-ICICIHeader {
    param($RawData)

    for ($i = 0; $i -lt $RawData.Count; $i++) {

        $rowValues = $RawData[$i].PSObject.Properties.Value

        $hasDate   = $rowValues | Where-Object { $_ -match "Transaction Date" }
        $hasDetail = $rowValues | Where-Object { $_ -match "Details" }
        $hasAmount = $rowValues | Where-Object { $_ -match "Amount" }

        if ($hasDate -and $hasDetail -and $hasAmount) {
            return $i
        }
    }

    return $null
}