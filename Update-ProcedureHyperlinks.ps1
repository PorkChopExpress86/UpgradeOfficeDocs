if (Get-Process -ProcessName "WINWORD" -ErrorAction SilentlyContinue) {
    Stop-Process -ProcessName "WINWORD"
}

Copy-Item -Path ".\Man. 10 - Mfg Proc\Ethomeen T12_8000974.docx" -Destination .\Test

$word = new-object -comobject word.application
$word.Visible = $false
$folderpath = ".\Test\Man. 10 - Mfg Proc\*"
$fileType = "*.docx"

Get-ChildItem -path $folderpath -include $fileType  | ForEach-Object {
    
    # open word document
    $opendoc = $word.Documents.open($_.FullName)

    # loop over each hyperlink
    $opendoc.hyperlinks | ForEach-Object {
        if ($_.Address -like "*\bowdenb\Downloads\Calcs\*") {
            $NewAddress = $_.Address -Replace "C:\\users\\bowdenb\\downloads\\calcs\\", "Calcs\\"
            $NewAddress = $NewAddress -replace "xls","xlsx"
            write-host "$NewAddress"
            "Updating {0} to {1}" -f $_.Address, $NewAddress
            $_.Address = $NewAddress 
            #$._TextToDisplay =$NewAddress
        }
    }
    "Saving changes to {0}" -f $opendoc.fullname
    $opendoc.save()
}

#Clean up
$word.quit()