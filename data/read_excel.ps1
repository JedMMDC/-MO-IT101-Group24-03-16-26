$excel = New-Object -ComObject Excel.Application
$excel.Visible=$false
$wb = $excel.Workbooks.Open('c:\Users\jedgb\OneDrive\MotorPH Payroll System\MotorPH_Employee Data _ Beltran, J..xlsx')
$ws = $wb.Sheets.Item(1)
$used = $ws.UsedRange
$rows = $used.Rows.Count
$cols = $used.Columns.Count
for($r=1; $r -le $rows; $r++){
    $line = @()
    for($c=1; $c -le $cols; $c++){
        $cell = $used.Item($r,$c).Text
        $line += $cell
    }
    Write-Output ($line -join ',')
}
$wb.Close($false)
$excel.Quit()