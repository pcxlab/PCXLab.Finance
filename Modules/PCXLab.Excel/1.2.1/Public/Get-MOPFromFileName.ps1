function Get-MOPFromFileName {

    param([string]$FileName)

    $name = $FileName.ToUpper()

    # 🔹 Bank
    switch -Regex ($name) {
        "ICICI" { $bank = "ICICI"; break }
        "HDFC"  { $bank = "HDFC"; break }
        default { $bank = "UNK" }
    }

    # 🔹 Type (FIXED ORDER)
    if ($name -match "_CC_") {
        $type = "CC"
    }
    elseif ($name -match "_DC_") {
        $type = "DC"
    }
    elseif ($name -match "_SB_") {
        $type = "SB"
    }
    else {
        $type = "OT"
    }

    # 🔹 Initials
    if ($name -match "^[A-Z]+_[A-Z]+_([A-Z]{2})_") {
        $initials = $matches[1]
    }
    else {
        $initials = "NA"
    }

    return "${bank}_${type}_${initials}"
}