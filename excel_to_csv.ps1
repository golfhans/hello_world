Function ExcelToCsv ($File) {
    $myDir = "D:\MyDocs\CurrentWorkDocuments\20200601_Summit_Files"
    $excelFile = "$myDir\" + "$File" + ".xlsx"
    $Excel = New-Object -ComObject Excel.Application
    $wb = $Excel.Workbooks.Open($excelFile)
	
    foreach ($ws in $wb.Worksheets) {
        $ws.SaveAs("$myDir\" + "$File" + "_" + $ws.name  + ".csv", 6)
    }
    $Excel.Quit()
}

$FileName = "MARS_HEFF_export_CF_04.30.2020"
ExcelToCsv -File $FileName