$path = "C:\Users\841797\Downloads\Downloads_Path"
$environment = "AT"
Function Get-FileName {
   <# param (
        OptionalParameters
    )#>

    Get-ChildItem -Path $path | ForEach-Object {
        $name = $_.BaseName
        #$extension = $_.Extension
        If(($name -like "*_$environment_*") -and ($name -like "*_OUT")) {
            Write-Host ("$name" + ".csv")

        }
    }
    
}

Get-FileName