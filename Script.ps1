function Get-FilesInFolder {
    # Parameter help description
    [CmdletBinding()]
    param (
        [Parameter(Position=1)]
        [String]
        $Path,
        [Parameter(Position=2)]
        [string]
        $OutPutFileName
    )

    Get-ChildItem -Path $Path -Recurse -File | Select-Object Name, BaseName, Extension | Export-Csv $OutPutFileName
}

Get-FilesInFolder -Path '.\Man. 10 - Mfg Proc' -OutPutFileName '.\man_10_files.csv'
Get-FilesInFolder -Path '.\Man. 11 - Lab Analytical Proc' -OutPutFileName '.\man_11_files.csv'
Get-FilesInFolder -Path '.\Man. 12 - Raw Material Specs' -OutPutFileName '.\man_12_files.csv'