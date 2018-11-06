Remove-Item -Path multiplecharts.xlsx -ErrorAction Ignore

$data = @"
ID,Product,Quantity,Price,Total
12001,Nails,37,3.99,147.63
12002,Hammer,5,12.10,60.5
12003,Saw,12,15.37,184.44
12010,Drill,20,8,160
12011,Crowbar,7,23.48,164.36
"@

$c1 = New-ExcelChartDefinition -YRange "Price"   -XRange "Product" -Title "Item price"   -NoLegend -Height 225
$c2 = New-ExcelChartDefinition -YRange "Total   "-XRange "Product" -Title "Total sales"  -NoLegend -Height 225 -Row 9  -Column 15
$c3 = New-ExcelChartDefinition -YRange "Quantity"-XRange "Product" -Title "Sales volume" -NoLegend -Height 225 -Row 15

$data | ConvertFrom-Csv |
    Export-Excel -Path  "multiplecharts.xlsx" -AutoFilter -AutoNameRange -AutoSize -ExcelChartDefinition $c1, $c2, $c3  -Show



## Multiseries
Remove-Item MultiseriesChart.xlsx -ErrorAction Ignore

$data = Invoke-Sum -data (Get-Process) -dimension Company -measure Handles, PM, VirtualMemorySize

$c = New-ExcelChartDefinition -Title "ProcessStats" `
    -ChartType LineMarkersStacked `
    -XRange "Processes[Name]" `
    -YRange "Processes[PM]", "Processes[VirtualMemorySize]" `
    -SeriesHeader "PM", "VM"

$data |
    Export-Excel -Path MultiseriesChart.xlsx -AutoSize -TableName Processes -ExcelChartDefinition $c  -Show