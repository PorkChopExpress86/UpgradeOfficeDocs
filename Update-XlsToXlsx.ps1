if (Get-Process -ProcessName "Excel" -ErrorAction SilentlyContinue) {
    Stop-Process -ProcessName "EXCEL"
}

$SaveFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault

$excel = new-object -ComObject excel.application
$excel.Visible = $False

#Get the files
$folderpath = ".\*"
$fileType = "*xls"

Get-ChildItem -path $folderpath -include $fileType -Recurse | foreach-object {
    
    # old document name
    $filePath = $PSitem.FullName

    $openworkbook = $excel.workbooks.open($_.FullName)

    # new document name
    $savename = ($_.fullname).substring(0,($_.FullName).lastindexOf("."))

    # convert to new format save and close

    # $openworkbook.Convert()
    $openworkbook.saveas("$savename", $SaveFormat);
    $openworkbook.close();
    Remove-Item -Path $filePath
}

#Clean up
$excel.quit()
