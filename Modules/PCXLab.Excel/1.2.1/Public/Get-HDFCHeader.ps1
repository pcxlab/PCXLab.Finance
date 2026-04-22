function Get-HDFCHeader {
    param($RawData)

    for ($i = 0; $i -lt $RawData.Count; $i++) {

        $row = $RawData[$i].PSObject.Properties.Value

        if (($row -join " ") -match "Transaction") {
            return $i
        }
    }

    return $null
}