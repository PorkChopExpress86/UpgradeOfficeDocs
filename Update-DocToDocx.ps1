if (Get-Process -ProcessName "WINWORD" -ErrorAction SilentlyContinue) {
    Stop-Process -ProcessName "WINWORD"
}

$word = new-object -comobject word.application
$word.Visible = $True
[ref]$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type]
#Get the files
$folderpath = ".\*"
$fileType = "*doc"

Get-ChildItem -path $folderpath -include $fileType -recurse | foreach-object {
    

    $path = ($_.FullName).substring(0,($_.FullName).lastindexOf("."))

    $opendoc = $word.documents.open($_.FullName)

    # convert to new format save and close
    $opendoc.Convert()
    $opendoc.saveas([ref]$path, [ref]$SaveFormat::wdFormatDocumentDefault);
    $opendoc.close();

    #Remove-Item -Path $filePath
}

#Clean up
$word.quit()
$word = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()
