<#
# Script FileName         : ExtractOTPFromMail.ps1
# developer Log fileName  : DeveloperLog.txt
# Description             : Extract OTP from mails
# Input                   : AA Environment details
# Output                  : OTP stored in OTPFile
# Developed Date          : 29/06/2022
# Developed By            : Shubham Nagre
# Reviewed By             : Sudhakar Govindasamy
# Prerequisite            : Microsoft Graph API, ObjectId, Access Token Secret Key, TenantId, ClientID
#----------------------------------------------------
#>

param (
    [parameter(mandatory = $true)]
    [string] $AAEnvironment
)

# Write log function
Function Write-IntoLog {
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, mandatory = $true)]  [String] $Type,
        [Parameter(Position = 1, mandatory = $true)]  [string] $Logstring
    )
   
    try {   
        $logtype = 'null'
            
        if ($Type.ToLower() -eq "i") {
            $logtype = "INFO"
        }
        elseif ($Type.ToLower() -eq "w") {
            $logtype = "WARNING"
        }
        elseif ($Type.ToLower() -eq "e") {
            $logtype = "ERROR"
        }
        else {
            $logtype = "INFO"
        }

        $LogFilePath = $DevLogFilePath
        $LogTime = Get-Date -Format '(dd/MMM/yy h:mm:ss tt)'
        $Type = $Type.ToUpper()
        $Text = "$LogTime $logtype | $logstring"

        if (!(Test-Path -Path $LogFilePath)) {
            New-Item -Path $LogFilePath -ItemType File | Out-Null
        }

        Add-Content -Path $LogFilePath -Value $Text
    }
    catch
    { }
}

# EXTRACT CONFIG FILE VALUES
Function Extract-ConfigFileData {  
    Try {
        $settingsFilePath = $env:USERPROFILE + "\Documents\Automation Anywhere Files\Automation Anywhere\My Docs\CTS\CTS\CTS EMEA MEDIDATA_USER_ACCESS\Medidata_Config.xml"
        $AAEnvironment = "DEV"  # Remove it later
        [xml]$configData = Get-Content -Path $settingsFilePath
        $AAEnvironment = $AAEnvironment.ToUpper()
        $settingsNode = $configData.SelectSingleNode("//config/$AAEnvironment")
        
        # Log File
        $DevLogFilePath = $settingsNode.file.DeveloperLog
        $Script:DevLogFilePath = $DevLogFilePath + "_" + $datetimeNew + ".txt"

        # Folder Path
        $Script:OTPFile = $settingsNode.file.OtpFolder

        # Mail Details
        $Script:TenantId = $settingsNode.Mail.TenantId           # Provide your Office 365 Tenant Id or Tenant Domain Name
        $Script:AppClientId = $settingsNode.Mail.AppClientId     # Provide Azure AD Application (client) Id of your app.
        $Script:Scope = $settingsNode.Mail.Scope                 # Scope
        $Script:SecretKey = $settingsNode.Mail.SecretKey         # Secret Key
        $Script:ObjectId = $settingsNode.Mail.ObjectId           # Object Id for user mailbox
        $Script:fromUser = $settingsNode.Mail.FromUser           # From User mail address

        Write-Host "Config Extraction Successfully"
    }
    Catch [System.Exception] {
        $sExceptionType = $_.Exception.GetType().FullName
        $sExpectionMessage = $_.Exception.Message
        
        Write-IntoLog "E" "__________________________Exception Message_______________________"
        Write-IntoLog "E" "Exception in the Extract-ConfigFileData Function" 
        Write-IntoLog "E" "Exception Type: $sExceptionType" 
        Write-IntoLog "E" "Exception Message: $sExpectionMessage" 
        Write-IntoLog "E" "__________________________Exception Message_______________________" 
    }
}

# Returns Token
Function Generate-Token ($TenantId, $AppClientId, $Scope, $SecretKey) {
    # Headers
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Content-Type", "application/x-www-form-urlencoded")
    $headers.Add("Cookie", "fpc=AmYpQ8qrEh5Hn8YtGP5-vDuTEjqNAgAAAFbETNoOAAAA; stsservicecookie=estsfd; x-ms-gateway-slice=estsfd")
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls, [Net.SecurityProtocolType]::Tls11, [Net.SecurityProtocolType]::Tls12, [Net.SecurityProtocolType]::Ssl3

    # URI To get Token
    $URI_Token = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    
    $body = "client_id=$AppClientId&scope=$Scope&client_secret=$SecretKey&grant_type=client_credentials"

    try {
        $response = Invoke-RestMethod $URI_Token -Method 'POST' -Headers $headers -Body $body
        Write-IntoLog "I" "Access Token is generated successfully"
    }
    catch {
        Write-IntoLog "E" "An error occured while generating token"
        $sExceptionType = $_.Exception.GetType().FullName
        $sExpectionMessage = $_.Exception.Message
        Write-IntoLog "E" "Exception Type - $sExceptionType"
        Write-IntoLog "E" "Exception Message - $sExpectionMessage"
    }

    return $response.access_token
}

# Read Mails for User
Function Read-Mail ($fromUser, $token) {

    $today_Date = Get-date -Format 'yyyy-MM-dd'

    # Create URI Mail 
    # Filters - From User, Mail Object Id (To read third party mail), recieved after today
    # Returns - Mail Subject, Mail Body
    $URI_Mail = 'https://graph.microsoft.com/v1.0/users/' + $ObjectId + '/messages?$filter=(from/emailAddress/address) eq ' + "'$fromUser'" + " and receivedDateTime gt $today_Date" + 'T00:00:00Z&isRead ne true$count=true&$select=subject,body'

    # Add authorization Bearer token in headers
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Prefer", "outlook.body-content-type=`"text`"")
    $headers.Add("Authorization", "Bearer $token")
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls, [Net.SecurityProtocolType]::Tls11, [Net.SecurityProtocolType]::Tls12, [Net.SecurityProtocolType]::Ssl3

    try {
        $response = Invoke-RestMethod $URI_Mail -Method 'GET' -Headers $headers
        Write-IntoLog "I" "Mail is fetched successfully on $today_Date"
    }
    catch {
        Write-IntoLog "E" "An error occured while reading mails"
        $sExceptionType = $_.Exception.GetType().FullName
        $sExpectionMessage = $_.Exception.Message
        Write-IntoLog "E" "Exception Type - $sExceptionType"
        Write-IntoLog "E" "Exception Message - $sExpectionMessage"
    }

    return $response
}

# Delete mail
Function Delete-Mail ($id) {
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization", "Bearer $token")

    $URI_Delete_Mail = "https://graph.microsoft.com/v1.0/users/$ObjectId/messages/$id"
    
    try {
        Invoke-RestMethod $URI_Delete_Mail -Method 'DELETE' -Headers $headers
    }
    catch {
        Write-IntoLog "E" "An error occured while deleting mail"
        $sExceptionType = $_.Exception.GetType().FullName
        $sExpectionMessage = $_.Exception.Message
        Write-IntoLog "E" "Exception Type - $sExceptionType"
        Write-IntoLog "E" "Exception Message - $sExpectionMessage"
    }
    
}

# Extract otp from mail body     
Function Fetch-OTP ($response) {
    # Capture total no of mails recieved from $FromUser, at $today_Date after 00:00:00 AM
    $count = ($response.value).Count

    # Validation if mail count is '0'
    if ($count -eq 0) {
        Write-Output "No Mails found"
        Write-IntoLog "E" "No mails recived from $fromUser for OTP"
        $value = -1
    }
    else {
        # Read latest mail recieved from $FromUser
        $mailBody = $response.value[$count - 1].body.content

        # Get Mail Id
        $mailId = $response.value[$count - 1].id

        $regexPattern = '\d{6}'     # Any 6 digit number in mail body content

        $otp = ($mailBody | Select-String $regexPattern -AllMatches).Matches
        Write-IntoLog "I" "OTP Token is extracted successfully"
        $value = $otp.value
        # Call delete mail function
        Delete-Mail $mailId
    }
    return $value.Trim()
}

try {
    Extract-ConfigFileData
    $token = Generate-Token -TenantId $TenantId -AppClientId $AppClientId -Scope $Scope -SecretKey $SecretKey
    $otp = Fetch-OTP(Read-Mail -fromUser $fromUser -token $token)
    
    # If otp value is -1; delete otp file; In AA code if file not found wait for 1 mins and run this script again
    if ($otp -eq -1) {
        Write-IntoLog "I" "OTP not found. stopping workflow."
        Write-IntoLog "I" "Waiting for 1 min."
        if (Test-Path -Path $OTPFile) { Remove-Item -Path $OTPFile }
        return "Mail not recieved"
    }
    else {
        if (Test-Path -Path $OTPFile) {
            # Check if File exist, override existing file; else create
            Set-Content -Path $OTPFile -Value $otp -NoNewline       # To not add new lines in content
            Write-IntoLog "I" "Otp file present. Overriding otp for mail"
            Write-IntoLog "I" "Otp fetched - $otp"
        }
        else {
            Set-Content -Path $OTPFile -Value $otp -NoNewline
            Write-IntoLog "I" "Otp fetched - $otp"
        }
    }
}
catch {
    $sExceptionType = $_.Exception.GetType().FullName
    $sExpectionMessage = $_.Exception.Message
        
    Write-IntoLog "E" "__________________________Exception Message_______________________"
    Write-IntoLog "E" "Exception in the Extract-MailPolling Function" 
    Write-IntoLog "E" "Exception Type: $sExceptionType" 
    Write-IntoLog "E" "Exception Message: $sExpectionMessage" 
    Write-IntoLog "E" "__________________________Exception Message_______________________" 
}
