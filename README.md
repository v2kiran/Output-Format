# Excel

### Open an existing workbook

```powershell
 $excel = Open-ExcelPackage .\databars.xlsx
```

### Close excel package

```powershell
Close-ExcelPackage $excel -Show
```

### List worksheets

```powershell
 PS C:\temp\excel> Get-ExcelSheetInfo .\databars.xlsx

Name                Index  Hidden Path
----                -----  ------ ----
Processes               1 Visible C:\temp\excel\databars.xlsx
ProcessesPivotTable     2 Visible C:\temp\excel\databars.xlsx
```

### Get workbook info

```powershell
Get-ExcelWorkbookInfo .\databars.xlsx
```

### Highlight duplicates

```powershell
$data  | Export-Excel $f -Show -AutoSize -ConditionalText (New-ConditionalText -ConditionalType DuplicateValues)
```

### Selecting a worksheet

```powershell
$sheet = $excel.Workbook.Worksheets["Sheet1"]
Add-ConditionalFormatting -WorkSheet $sheet -Range "B1:D14" -DataBarColor CornflowerBlue
```

### hide gridlines

```powershell
$sheet1.View.ShowGridLines = $false

# hides column names A,B,C etc
$sheet1.View.ShowHeaders = $false
```

### Format

```powershell
# cell format
Set-Format -Address $sheet1.Cells["F:F"] -NumberFormat "#.#0%"  -WrapText -HorizontalAlignment Center -Width 12

# font settings
Set-Format -Address $sheet1.Cells["A2:C8"] -FontColor GrayText

# write to a cell
Set-Format -Address $sheet1.Cells["F1"] -HorizontalAlignment center -Bold -Value Revenue

# formula
Set-Format -Address $sheet1.Cells["E10"] -Formula "=Sum(E3:E8)" -Bold
```

### Join Worksheet

```powershell
#Create a summary page with a title of Summary, label the blocks with the name of the sheet they came from and hide the source sheets
Join-Worksheet -Path $path -HideSource -WorkSheetName Summary -NoHeader -LabelBlocks  -AutoSize -Title "Summary" -TitleBold -TitleSize 22 -show
```

### Charts

```powershell
# pass the below to export-excel
ExcelChartDefinition = New-ExcelChartDefinition -XRange Item -YRange UnitSold -Title 'Units Sold'
```

### Calculated properties

```powershell
$data | ConvertFrom-Csv |
    Add-Member -PassThru -MemberType NoteProperty -Name Total -Value "=units*cost" |
    Export-Excel -Path .\testFormula.xlsx -Show -AutoSize -AutoNameRange
```

### Hyperlinks

```powershell
$(
    New-PSItem '=Hyperlink("http://dougfinke.com/blog","Doug Finke")' @("Link")
    New-PSItem '=Hyperlink("http://blogs.technet.com/b/heyscriptingguy/","Hey, Scripting Guy")'

) | Export-Excel "c:\temp\excel\hyperlink.xlsx" -AutoSize -Show
```
