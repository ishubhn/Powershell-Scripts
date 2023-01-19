$snowusername = ""
$Snowpassword = ""
$RITMNumber = ""

$parentURL = ""

$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $Snowusername, $Snowpassword)))
$getheaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$getheaders.Add("Authorization", ("Basic {0}" -f $base64AuthInfo))
$getheaders.Add("Accept","Application/json")
$getheaders.Add("Content-Type","text/plain")
$getmethod = "get"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls, [Net.SecurityProtocolType]::Tls11, [Net.SecurityProtocolType]::Tls12, [Net.SecurityProtocolType]::Ssl3

If (!(Test-Path -Path "C:\Users\E0475735\Desktop\result.csv")) {
    New-Item -Path 'C:\Users\E0475735\Desktop\result.csv' -ItemType File
}

Import-CSV -Path "C:\Users\E0475735\Desktop\files.csv" | ForEach-Object {
    $RITMNumber = $_.number
    
    $url = "$parentURL=$RITMNumber&sysparm_display_value=false&sysparm_display_value=all&sysparm_fields=number,short_description,service_offering.name,description"

    try {
            $response  = Invoke-RestMethod -Headers $getheaders -Method $getmethod -Uri $url
    } catch {
            Write-Host "Error" $_.Exception.Message
    }

    $number = $response.result.number
    $so = $response.result."service_offering.name"
    $short_description = $response.result.short_description
    $description = $response.result.description

    #$numberA += $number
    #$soA += $so
    #$short_descriptionA += $short_description
    #$descriptionA += $description

    $newCSV = New-Object psobject -Property @{
        Number = $number
        SO = $so
        Short_Description = $short_description
        Description = $description
    }

    Export-Csv -InputObject $newCSV -Path "C:\Users\E0475735\Desktop\result.csv" -NoTypeInformation -Append -Force

}

Write-Host "Done"