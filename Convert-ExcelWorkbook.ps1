<#
    .Synopsis
    PowerShell script to convert Excel workbooks

    .Description

    This script converts Excel compatible workbooks to a selected format utilizing the Excel SaveAs function. Each file is converted by a single dedicated Excel COM instance.

    The script converts either all workbook in a single folder of a matching include filter or a single file.

    Currently supported target document types:
    - Default --> Excel 2016
    - PDF
    - XPS
    - HTML

    Author: Thomas Stensitzki
    Converted to work with Excel by: Blake Bowden on 2021-7-14

    Version 1.1 2019-11-26

    .NOTES 

    Requirements 
    - Excel 2016+ installed locally

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0      Initial release

    .LINK
    http://scripts.granikos.eu

    .PARAMETER SourcePath
    Source path to a folder containing the workbooks to convert or full path to a single document

    .PARAMETER IncludeFilter
    File extension filter when converting all files  in a single folder. Default: *.xls

    .PARAMETER TargetFormat
    Excel Save AS target format. Currently supported: Default, CSV

    .PARAMETER DeleteExistingFiles
    Switch to delete an exiting target file

    .EXAMPLE
    Convert all .xls files in E:\temp to Default

    .\Convert-ExcelWorkbook.ps1 -SourcePath E:\Temp -IncludeFilter *.doc 

    .EXAMPLE
    Convert all .doc files in E:\temp to XPS

    .\Convert-ExcelWorkbook.ps1 -SourcePath E:\Temp -IncludeFilter *.doc -TargetFormat XPS

    .EXAMPLE
    Convert a single document to Word default format

    .\Convert-ExcelWorkbook.ps1 -SourcePath E:\Temp\MyDocument.doc
#>

[CmdletBinding()]
Param(
    [string]$SourcePath = '',
    [string]$IncludeFilter = '*.xls',
    [ValidateSet('Default', 'CSV')] # Only some of the supported file formats are currently tested
    [string]$TargetFormat = 'Default',
    [switch]$DeleteExistingFiles
)
  
$ERR_OK = 0
$ERR_COMOBJECT = 1001 
$ERR_SOURCEPATHMISSING = 1002
$ERR_WORDSAVEAS = 1003
  
# Define Word target document types
# Source: https://msdn.microsoft.com/en-us/vba/word-vba/articles/wdsaveformat-enumeration-word  

$EwFormat = @{
    'xlAddIn'                       = 18 # Microsoft Excel 97-2003 Add-In
    'xlAddIn8'                      = 18 # Microsoft Excel 97-2003 Add-In
    'CSV'                           = 6 # CSV
    'xlCSVMac'                      = 22 # Macintosh CSV
    'xlCSVMSDOS'                    = 24 # MSDOS CSV
    'xlCSVUTF8'                     = 62 # UTF8 CSV
    'xlCSVWindows'                  = 23 # Windows CSV
    'xlCurrentPlatformText'         = -4158 # Current Platform Text
    'xlDBF2'                        = 7 # Dbase 2 format
    'xlDBF3'                        = 8 # Dbase 3 format
    'xlDBF4'                        = 11 # Dbase 4 format
    'xlDIF'                         = 9 # Data Interchange format
    'xlExcel12'                     = 50 # Excel Binary Workbook
    'xlExcel2'                      = 16 # Excel version 2.0 (1987)
    'xlExcel2FarEast'               = 27 # Excel version 2.0 far east (1987)
    'xlExcel3'                      = 29 # Excel version 3.0 (1990)
    'xlExcel4'                      = 33 # Excel version 4.0 (1992)
    'xlExcel4Workbook'              = 35 # Excel version 4.0. Workbook format (1992)
    'xlExcel5'                      = 39 # Excel version 5.0 (1994)
    'xlExcel7'                      = 39 # Excel 95 (version 7.0)
    'xlExcel8'                      = 56 # Excel 97-2003 Workbook
    'xlExcel9795'                   = 43 # Excel version 95 and 97
    'xlHtml'                        = 44 # HTML format
    'xlIntlAddIn'                   = 26 # International Add-In
    'xlIntlMacro'                   = 25 # International Macro
    'xlOpenDocumentSpreadsheet'     = 60 # OpenDocument Spreadsheet
    'xlOpenXMLAddIn'                = 55 # Open XML Add-In
    'xlOpenXMLStrictWorkbook'       = 61 # Strict Open XML file
    'xlOpenXMLTemplate'             = 54 # Open XML Template
    'xlOpenXMLTemplateMacroEnabled' = 53 # Open XML Template Macro Enabled
    'xlOpenXMLWorkbook'             = 51 # Open XML Workbook
    'xlOpenXMLWorkbookMacroEnabled' = 52 # Open XML Workbook Macro Enabled
    'xlSYLK'                        = 2 # Symbolic Link format
    'Template'                    = 17 # Excel Template format
    'xlTemplate8'                   = 17 # Template 8
    'xlTextMac'                     = 19 # Macintosh Text
    'xlTextMSDOS'                   = 21 # MSDOS Text
    'xlTextPrinter'                 = 36 # Printer Text
    'xlTextWindows'                 = 20 # Windows Text
    'xlUnicodeText'                 = 42 # Unicode Text
    'xlWebArchive'                  = 45 # Web Archive
    'xlWJ2WD1'                      = 14 # Japanese 1-2-3
    'xlWJ3'                         = 40 # Japanese 1-2-3
    'xlWJ3FJ3'                      = 41 # Japanese 1-2-3 format
    'xlWK1'                         = 5 # Lotus 1-2-3 format
    'xlWK1ALL'                      = 31 # Lotus 1-2-3 format
    'xlWK1FMT'                      = 30 # Lotus 1-2-3 format
    'xlWK3'                         = 15 # Lotus 1-2-3 format
    'xlWK3FM3'                      = 32 # Lotus 1-2-3 format
    'xlWK4'                         = 38 # Lotus 1-2-3 format
    'xlWKS'                         = 4 # Lotus 1-2-3 format
    'Default'                       = 51 # Workbook default
    'xlWorkbookNormal'              = -4143 # Workbook normal
    'xlWorks2FarEast'               = 28 # Microsoft Works 2.0 far east format
    'xlWQ1'                         = 34 # Quattro Pro format
    'xlXMLSpreadsheet'              = 46 # XML Spreadsheet
}

$FileExtension = @{
    'Workbook' = '.xls'
    'Template' = '.xlt'
    'Default'  = '.xlsx'
    'CSV'      = '.csv'
}

function Invoke-Excel {
    [CmdletBinding()]
    param (
        [string]$FileSourcePath = '',
        [string]$SourceFileExtension = '',
        [string]$TargetFileExtension = '',
        [int]$EwSaveFormat = 51, # Default saveas xlsx
        [switch]$DeleteFile        
    )
    if ($FileSourcePath -ne '') {
        Write-Output ('Converting {0}' -f $FileSourcePath)

        $ExcelApplication = $null

        # create a new instance of word
        try {
            $ExcelApplication = New-Object -ComObject excel.application
        }
        catch {
            Write-Error -Message "Excel COM object caould not be loaded"
            exit $ERR_COMOBJECT
        }

        # try to open the word document and save in the new format
        try {
            $ExcelWorkbook = $ExcelApplication.Workbooks.Open($FileSourcePath)

            # replace the source file extension with the appropriate target file extension
            $NewFilePath = ($FileSourcePath).Replace($SourceFileExtension, $TargetFileExtension)

            if ((Test-Path -Path $NewFilePath) -and $DeleteFile) {
                # delete existing file
                $null = Remove-Item -Path $NewFilePath -Force -Confirm:$false
            }

            # save the new document
            $ExcelWorkbook.SaveAs([ref]$NewFilePath, [ref]$EwSaveFormat)
        }
        catch {
            # error
            Write-Error -Message "Error saving workbook$($FileSourcePath): Exception: $($_.Exception.Message)"
            exit $ERR_WORDSAVEAS
        }
        finally {
            #close open word applications
            $ExcelWorkbook.Close()
            $ExcelApplication.Quit()
            [Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelApplication) | Out-Null

            if (Test-Path variable:global:ExcelApplication) {
                Remove-Variable -Name ExcelApplication -Scope Global 4>$Null
            }

            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}

if ($SourcePath -ne '') {
  
    # Check whether SourcePath is a single file or directory
    $IsFolder = $false
    try {
        $IsFolder = ((Get-Item -Path $SourcePath ) -is [System.IO.DirectoryInfo])
    }
    catch {}
  
    if ($IsFolder) {
  
        # We need to iterate a source folder
        $SourceFiles = Get-ChildItem -Path $SourcePath -Include $IncludeFilter -Recurse
  
        Write-Verbose -Message ('{0} files found in {1}' -f ($SourceFiles | Measure-Object).Count, $SourcePath)
  
        # Let's work on all files
        foreach ($File in $SourceFiles) {
  
            Invoke-Excel -FileSourcePath $File.FullName -SourceFileExtension $File.Extension -TargetFileExtension $FileExtension.Item($TargetFormat) -EwSaveFormat $ewFormat.Item($TargetFormat)
          
        }
    }
    else {
        # It's just a single file
  
        $File = Get-Item -Path $SourcePath
  
        Invoke-Excel -FileSourcePath $File.FullName -SourceFileExtension $File.Extension -TargetFileExtension $FileExtension.Item($TargetFormat) -EwSaveFormat $ewFormat.Item($TargetFormat)
  
    }
}
else {
    Write-Warning -Message 'No document source path has been provided'
    exit $ERR_SOURCEPATHMISSING
}