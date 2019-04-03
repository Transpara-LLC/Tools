$comments = @' 
'@ 
 
# ----------------------------------------------------- 
function Release-Ref ($ref) { 
([System.Runtime.InteropServices.Marshal]::ReleaseComObject( 
[System.__ComObject]$ref) -gt 0) 
[System.GC]::Collect() 
[System.GC]::WaitForPendingFinalizers() 
} 
# ----------------------------------------------------- date tag value
 
$arrExcelValues = @() 
 
$objExcel = new-object -comobject excel.application  
$objExcel.Visible = $True  
$objWorkbook = $objExcel.Workbooks.Open("C:\Scripts\Test.xls") 
$objWorksheet = $objWorkbook.Worksheets.Item(1) 
 
$i = 1 
 
Do { 
    $date =  $objWorksheet.Cells.Item($i, 1).Value() 
    $tag =  $objWorksheet.Cells.Item($i, 2).Value() 
    $value =  $objWorksheet.Cells.Item($i, 3).Value() 
    echo"$date $tag $value"
    $i++ 
} 
While ($objWorksheet.Cells.Item($i,1).Value() -ne $null) 
 
$a = $objExcel.Quit 

$a = Release-Ref($objWorksheet) 
$a = Release-Ref($objWorkbook) 
$a = Release-Ref($objExcel) 