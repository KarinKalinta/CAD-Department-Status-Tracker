$ErrorActionPreference = "Stop"

$path = Join-Path (Get-Location) "CAD_Status_Tracker_Batch_Aware_MANAGER_VIEW.xlsx"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open($path)
    $excel.Calculation = -4105
    $excel.CalculateFullRebuild()

    $sheet = $workbook.Worksheets.Item("Dashboard")
    $used = $sheet.UsedRange
    $errors = @()

    for ($row = 1; $row -le $used.Rows.Count; $row++) {
        for ($col = 1; $col -le $used.Columns.Count; $col++) {
            $cell = $used.Cells.Item($row, $col)
            $text = [string]$cell.Text
            if ($text -in @("#REF!", "#VALUE!", "#NAME?", "#DIV/0!")) {
                $errors += "$($cell.Address($false,$false))=$text"
            }
        }
    }

    if ($errors.Count -gt 0) {
        Write-Output "ERRORS"
        $errors | Select-Object -First 50
        $workbook.Close($false)
        exit 1
    }

    $workbook.Save()
    $workbook.Close($true)
    Write-Output "OK"
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
