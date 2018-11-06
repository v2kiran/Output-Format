$xlFile = "c:\temp\excel\testPivot.xlsx"
Remove-Item $xlFile -ErrorAction Ignore

$data = @"
Region,Area,Product,Units,Cost
North,A1,Apple,100,.5
South,A2,Pear,120,1.5
East,A3,Grape,140,2.5
West,A4,Banana,160,3.5
North,A1,Pear,120,1.5
North,A1,Grape,140,2.5
"@ | ConvertFrom-Csv

$data |
    Export-Excel $xlFile -Show `
    -AutoSize -AutoFilter `
    -IncludePivotTable `
    -PivotRows Product `
    -PivotData @{"Units" = "sum"} -PivotFilter Region, Area -Activate


#Example - 2

Remove-Item $xlFile -ErrorAction Ignore
gsv | select -First 10 | 
    Export-Excel $xlFile -Show `
    -AutoSize -AutoFilter `
    -IncludePivotTable `
    -PivotRows StartType, `
    -PivotFilter Status, Starttype -Activate



### Example - 3

Remove-Item $xlFile -ErrorAction Ignore
$data = Get-Service | Sort-Object StartType | Select Name, DisplayName, Status, StartType

# pivotdata is the values area - will give a count coulmn value for each pivot row value of the filed by default
$exportExcelSplat = @{
    Show              = $true
    AutoSize          = $true
    Path              = $xlFile
    WorksheetName     = 'Svc'
    PivotColumns      = 'Status'
    PivotRows         = 'StartType'
    PivotData         = 'StartType'
    IncludePivotTable = $true
    IncludePivotChart = $true
    ChartType         = 'PieExploded3D'
    AutoFilter        = $true
    Passthru          = $true
}
$wbook = $data | Export-Excel  @exportExcelSplat 

$wbook.Workbook.Worksheets['Svc'].View.ShowGridLines = $false
$wbook.Workbook.Worksheets['SvcPivotTable'].View.ShowGridLines = $false

Close-ExcelPackage -ExcelPackage $wbook -Show


