$file = "c:\temp\excel\conditionalTextFormatting.xlsx"
Remove-Item $file -ErrorAction Ignore

Get-Service |
    Select-Object Status, Name, DisplayName, ServiceName |
    Export-Excel $file -WorkSheetname "Services" -Show -AutoSize -AutoFilter -ConditionalText $(
    New-ConditionalText stop                                                  #Stop is the condition value, the rule is defaults to 'Contains text' and the default Colors are used
    New-ConditionalText runn darkblue cyan                                    #runn is the condition value, the rule is defaults to 'Contains text'; the foregroundColur is darkblue and the background is cyan
    New-ConditionalText -ConditionalType EndsWith svc wheat green             #the rule here is 'Ends with' and the value is 'svc' the forground is wheat and the background dark green
    New-ConditionalText -ConditionalType BeginsWith windows darkgreen wheat   #this is 'Begins with "Windows"' the forground is dark green and the background wheat
)


#region blanks
#Define a "Contains blanks" rule. No format is specified so it default to dark-red text on light-pink background.
$ContainsBlanks = New-ConditionalText -ConditionalType ContainsBlanks
$data | Export-Excel $file -show -ConditionalText $ContainsBlanks
#endregion


#region databars


$path = "c:\temp\excel\databars.xlsx"
Remove-Item -Path $path -ErrorAction Ignore

#Export processes, and get an ExcelPackage object representing the file.
$excel = Get-Process |
    Select-Object -Property Name, Company, Handles, CPU, PM, NPM, WS |
    Export-Excel -Path $path -ClearSheet -WorkSheetname "Processes"  -PassThru

$sheet = $excel.Workbook.Worksheets["Processes"]

#Apply fixed formatting to columns. Set-Format is an Alias for Set-Excel Range, -NFormat is an alias for numberformat
$sheet.Column(1) | Set-ExcelRange -Bold -AutoFit
$sheet.Column(2) | Set-Format -Width 29 -WrapText
$sheet.Column(3) | Set-Format -HorizontalAlignment Right -NFormat "#,###"

Set-ExcelRange -Range $sheet.Cells["E1:H1048576"]  -HorizontalAlignment Right -NFormat "#,###"
#Set-Format is an alias for Set-ExcelRange
Set-Format -Range   $sheet.Column(4)  -HorizontalAlignment Right -NFormat "#,##0.0" -Bold
#In Set-ExcelRange / Set-Format "-Address" is an alias for "-Range"
Set-Format -Address $sheet.Row(1) -Bold -HorizontalAlignment Center

#Create a Red Data-bar for the values in Column D
Add-ConditionalFormatting -WorkSheet $sheet -Address "D2:D1048576" -DataBarColor Red
# Conditional formatting applies to "Addreses" aliases allow either "Range" or "Address" to be used in Set-ExcelRange or Add-Conditional formatting.
Add-ConditionalFormatting -WorkSheet $sheet -Range  "G2:G1048576" -RuleType GreaterThan -ConditionValue "104857600"  -ForeGroundColor Red

foreach ($c in 5..9) {Set-Format -Address $sheet.Column($c)  -AutoFit }

#Create a pivot and save the file.
Export-Excel -ExcelPackage $excel -WorkSheetname "Processes" -IncludePivotChart -ChartType ColumnClustered -NoLegend -PivotRows company  -PivotData @{'Name' = 'Count'}  -Show


#endregion



#region existing workbook - 3 -icon set
$excel = Open-ExcelPackage .\databars.xlsx
Get-Process | Where-Object Company | Select-Object Company, Name, PM, Handles, *mem* |

#This example creates a 3 Icon set for the values in the "PM column, and Highlights company names (anywhere in the data) with different colors

Export-Excel -ExcelPackage $excel -WorkSheetname "IconSet" -Show -AutoSize -AutoNameRange `
    -ConditionalFormat $(
    New-ConditionalFormattingIconSet -Range "C:C" `
        -ConditionalFormat ThreeIconSet -IconType Arrows

) -ConditionalText $(
    New-ConditionalText Microsoft -ConditionalTextColor Black
    New-ConditionalText Google  -BackgroundColor Cyan -ConditionalTextColor Black
    New-ConditionalText authors -BackgroundColor LightBlue -ConditionalTextColor Black
    New-ConditionalText nvidia  -BackgroundColor LightGreen -ConditionalTextColor Black
)



#endregion