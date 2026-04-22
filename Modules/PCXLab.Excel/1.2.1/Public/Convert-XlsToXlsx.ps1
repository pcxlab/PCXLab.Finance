function Convert-XlsToXlsx {
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File
    )

    # If already xlsx → return as is
    if ($File.Extension -eq ".xlsx") {
        return $File
    }

    # Build output file path
    $newFileName = Get-OutputFileName -File $File -Converted
    $newFile = Join-Path $File.DirectoryName $newFileName

    # If already converted → reuse
    if (Test-Path $newFile) {
        return Get-Item $newFile
    }

    Write-Host "Converting XLS -> XLSX: $($File.Name)" -ForegroundColor Yellow

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {
        $workbook = $excel.Workbooks.Open($File.FullName)

        # 51 = xlOpenXMLWorkbook (.xlsx)
        $workbook.SaveAs($newFile, 51)

        $workbook.Close()
    }
    catch {
        throw "Failed to convert file: $($File.Name)"
    }
    finally {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    return Get-Item $newFile
}