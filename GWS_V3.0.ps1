#----------------------------------------------------
# Script FileName         : "Application Monitoring.ps1"
# developer Log fileName  : "Powershell_log.txt"
# Description             : Launches IE application and takes screenshot.
# Input                   : None
# Output                  : Job status updated successfully
# Developed Date          : 4/24/2020
# Developed By            : Haribaskar
# Reviewed By             : Ashok
#----------------------------------------------------
try
{
    if(-not (Test-Path $PSScriptRoot)){    # Validates the path
        throw "Not found - $ScriptRootPath."
    }
    
    $LogFilePath = Join-Path -Path $PSScriptRoot -ChildPath "GEWAESSERSCHUTZ_Log.txt"
    $ConfigFilePath = Join-Path -path $PSScriptRoot -ChildPath "GEWAESSERSCHUTZ_Config.xml"
    #$DevLogFilePath = "D:\AVM Automation\Tools\PowerShell\Log\Log_" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".txt"
    $DevLogFilePath = "$PSScriptRoot\Log_" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".txt"

    $DateStamp = get-date -Format "yyy-MM-dd hh:mm:ss"
    [xml]$XmlContent = Get-Content $ConfigFilePath
    $MailTo = $XmlContent.Settings.MailTo
    $MailCC = $XmlContent.Settings.MailCC
    $MailFrom = $XmlContent.Settings.MailFrom
    $SMTPServer = $XmlContent.Settings.SMTPServer
    $SMTPPort = $XmlContent.Settings.SMTPPort
    $Shortdesc = $XmlContent.Settings.Shortdescription
    $ITSolution = $XmlContent.Settings.ITSolution
    $SupportGroup = $XmlContent.Settings.AssignmentGroup
    $Priority = $XmlContent.Settings.Priority
    $serviceoffering = $XmlContent.Settings.Serviceoffering
    $Attachmentpath = $XmlContent.Settings.AttachmentPath
    $snowattachmentAPI = $XmlContent.Settings.snowattachmentAPI
    $URLValue = $XmlContent.Settings.URLtoMonitor
    $Successubject = $XmlContent.Settings.success.subject
    $Successbody = $XmlContent.Settings.success.body
    $SuccessTo = $XmlContent.Settings.success.To
    $Failubject = $XmlContent.Settings.Fail.subject
    $Failbody = $XmlContent.Settings.Fail.body
    $FailTo = $XmlContent.Settings.Fail.To
    $Snowapi = $XmlContent.Settings.SnowAPI
    $Shortdesc = $Shortdesc.Replace('@@Date@@',$DateStamp)
    Add-Type -Assembly "Microsoft.VisualBasic"

    if((Test-Path -Path $Attachmentpath)){
        Remove-Item -Force -Path $Attachmentpath
     }
       #Retrieve password
     [void][Windows.Security.Credentials.PasswordVault,Windows.Security.Credentials,ContentType=WindowsRuntime]
    $vault = New-Object Windows.Security.Credentials.PasswordVault
    $Creds = $vault.Retrieve('SNOWProdAPI','hive2snow')
    #$Creds = $vault.Retrieve('SNOWAPI','hive2snow')
    $Password = $creds.Password
    $UserName = $creds.UserName
    Function WriteToLog($Type, $Message)
    {
    $LogTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $Text = "$LogTime - $Type" + ":" + " $Message"
    if (!(Test-Path -Path $DevLogFilePath))
    {
        New-Item -Path $DevLogFilePath -ItemType File | Out-Null
    }    Add-Content -Path $DevLogFilePath -Value $Text    }
        #Launching IE
        $ie = new-object -com "InternetExplorer.Application"
        $ie.navigate("$URLValue")
        $ie.visible = $true
        $ie.fullscreen = $true
        While ($ie.Busy) { Sleep -m 10 }
        $ieProc = Get-Process | ? { $_.MainWindowHandle -eq $ie.HWND }
        $sessionid = [System.Diagnostics.Process]::GetCurrentProcess().SessionID
        WriteToLog "INFO" "Session ID $sessionid"
        Add-Content -Path $LogfilePath -Value "$DateStamp | $sessionid "
        #$ieProc = get-process | Where-Object { $_.MainWindowTitle -match 'Gewässerschutz - Internet Explorer'}
        

        Function Get-ScreenShot
        { 
  
        Param(
                [Parameter(ParameterSetName='File')]
                [string]$FullName
        )
              
        Add-type -AssemblyName System.Drawing
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
        $Screen = [System.Windows.Forms.SystemInformation]::WorkingArea #VirtualScreen
        <#$Width = $Screen.Width
        $Height = $Screen.Height#>
        $width = 1366
        $height = 728
        $Left = $Screen.Left
        $Top = $Screen.Top
  
        $Imageformat= [System.Drawing.Imaging.ImageFormat]::Jpeg;

        # Create bitmap using the top-left and bottom-right bounds
        $Bitmap = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $Width, $Height

        # Create Graphics object
        $Graphic = [System.Drawing.Graphics]::FromImage($Bitmap)

        # Capture screen
#        $Graphic.CopyFromScreen($Left, $Top, 0, 0, $Bitmap.Size)
    $Graphic.CopyFromScreen(0,0, 0,0, $Bitmap.Size)

        # Save to file
        $Bitmap.Save($FullName, $Imageformat) 
        Write-Verbose -Message "[$(get-date -Format T)] Screenshot saved to $FullName"
     
        }    
        
            Function Create-Incident
    { 
    $url = $SnowAPI
    WriteToLog "INFO" "Create Incident"
    #$body = '{"requested_by":"e0412362","service_offering":".Net Hosting non-Prod","contact_type":"email","assignment_group":"AM-CZ-OP3-ASSET_MGMT_DM","category":"Access","symptoms":"Account Locked","urgency":"2","short_description":"short description","number":"","correlation_display":"HIVE","action":"CREATE"}'
    $body = @{requested_by="e0410307";service_offering="$Serviceoffering";contact_type="email";assignment_group="$SupportGroup";category="Application";symptoms="Failure to Launch";urgency="1";prioirity="1";short_description="$Shortdesc";number="";correlation_display="HIVE";action="CREATE"}
    $body = (ConvertTo-Json $body)
    $UserID = $UserName
    $PlainPassword = $Password
    #$password = ConvertTo-SecureString $PlainPassword -AsPlainText -Force

    # Build auth header
    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $UserID, $PlainPassword)))

    # Set proper headers
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add('Authorization',('Basic {0}' -f $base64AuthInfo))
    $headers.Add('Accept','application/json')
    $headers.Add('Content-Type','application/json')
  
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $result = Invoke-RestMethod -Headers $headers -Method POST -Uri $url -Body $body
    if($result -ne $null)
    {
    $sys_ID = $result.result | Where-Object{$_.transform_map -eq 'Incident-HIVE center'} | select sys_id,display_value
    $Ticketvalue = $sys_ID.display_Value
    
    Remove-Variable -name headers
    Remove-Variable -name url
    if($offlineflag -eq $false)
    {
    UploadAttachment($sys_ID.sys_id)
    }
    Add-Content -Path $LogfilePath -Value "$DateStamp | URL is down | $Ticketvalue"  
    WriteToLog "INFO" "Ticket created sucessfully $Ticketvalue"
    Return $Ticketvalue
    }

    }
    Function UploadAttachment($SysID)
    { 
    #   $url = 'https://apidev-eu.sanofi.com:8443/gsuf/ams20/1.0/api/now/import/x_saag_sanofi_int_incident'
    WriteToLog "INFO" "Upload attachment"
    $url = "$SnowattachmentAPI" + "&table_sys_id=$SysID&file_name=$Attachmentpath"

    # Set proper headers
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add('Authorization',('Basic {0}' -f $base64AuthInfo))
    $headers.Add('Accept','application/json')
    $headers.Add('Content-Type','*/*')
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    #$result = Invoke-RestMethod -Headers $headers -Method POST -Uri $url 
    $result = Invoke-WebRequest -uri $url -Headers $headers -Method Post -InFile $Attachmentpath
    WriteToLog "INFO" "Upload success"
    }
    #Check if page is loaded correctly  
    Start-Sleep -s 5
        if($sessionid -ne 0)
        {
        $offlineflag = $false
        [Microsoft.VisualBasic.Interaction]::AppActivate($ieProc.Id)
        $LblUsername = $ie.Document.getElementById("ctl00_lblDate")
        Get-ScreenShot -FullName $Attachmentpath -Verbose
        if(-not ([string]::IsNullOrEmpty($LblUsername)))
        {
            $Successubject = $Successubject.replace('@@Datestamp@@',$DateStamp)
            $successbody = $Successbody.replace('@@Datestamp@@',$DateStamp)
            $successbody = $Successbody.replace('$br$','<br>')
           send-MailMessage -To $SuccessTo.split(";") -Cc $MailCC.split(";") -Subject $Successubject -Body $Successbody -SmtpServer $SMTPServer -BodyAsHtml -From $MailFrom -Port $SMTPPort -Attachments $Attachmentpath
            Add-Content -Path $LogfilePath -Value "$DateStamp | URL is up.  | "
            WriteToLog "INFO" "Success in backend"
        }
        else
        {
        $TicketID = Create-Incident
        Add-Content -Path $LogfilePath -Value "$DateStamp | Error frontend.  | "
        WriteToLog "INFO" "Frontend Error"
        ##sending Mail
        $Failubject = $Failubject.replace('@@Datestamp@@',$DateStamp)
            $Failbody = $Failbody.replace('@@Datestamp@@',$DateStamp)
            $Failbody = $Failbody.replace('$br$','<br>')
            $Failbody = $Failbody.replace('@@Incident@@',$TicketID)
           send-MailMessage -To $FailTo.split(";") -Cc $MailCC.split(";") -Subject $Failubject -Body $Failbody -SmtpServer $SMTPServer -BodyAsHtml -From $MailFrom -Port $SMTPPort -Attachments $Attachmentpath
           WriteToLog "INFO" "Incident created mail sent"
        }
        }
        else
        {
        $offlineflag = $true
        $webclient = New-Object System.Net.WebClient
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
        $webclient.UseDefaultCredentials = $true
        $Cookiecontainer = New-Object System.Net.CookieContainer
        $URLrequest = [system.Net.WebRequest]::Create($URLValue)
        Add-Content -Path $LogfilePath -Value "$DateStamp | Backend check.  | "
        $URLrequest.UseDefaultCredentials = $true
        try
        {
        $URLResponse = $URLrequest.GetResponse()
        $Statuscode = [int]$URLResponse.StatusCode
        
        if($Statuscode -eq 200)
        {
            WriteToLog "INFO" "URL up backend"
            $Successubject = $Successubject.replace('@@Datestamp@@',$DateStamp)
            $Successbody = 'Hi Nathalie,$br$$br$@GEWAESSERSCHUTZ application is up and available on @@Datestamp@@ $br$$br$Regards,$br$AMS Team'
            $successbody = $Successbody.replace('@@Datestamp@@',$DateStamp)
            $successbody = $Successbody.replace('$br$','<br>')
            
            send-MailMessage -To $SuccessTo.split(";") -Cc $MailCC.split(";") -Subject $Successubject -Body $Successbody -SmtpServer $SMTPServer -BodyAsHtml -From $MailFrom -Port $SMTPPort 
            Add-Content -Path $LogfilePath -Value "$DateStamp | URL is up.  | "
            WriteToLog "INFO" "Backend URL up mail sent"
        }
        else
    {
        WriteToLog "INFO" "URL down backend"
        throw 
    }
    }
    catch
    {
        $TicketID = Create-Incident
        Add-Content -Path $LogfilePath -Value "$DateStamp | Error.  | "
        ##sending Mail
        $Failubject = $Failubject.replace('@@Datestamp@@',$DateStamp)
            $Failbody = $Failbody.replace('@@Datestamp@@',$DateStamp)
            $Failbody = $Failbody.replace('$br$','<br>')
            $Failbody = $Failbody.replace('@@Incident@@',$TicketID)
            send-MailMessage -To $FailTo.split(";") -Cc $MailCC.split(";") -Subject $Failubject -Body $Failbody -SmtpServer $SMTPServer -BodyAsHtml -From $MailFrom -Port $SMTPPort
            WriteToLog "INFO" "Backend URL down mail sent"
    }

    }
}
catch
{
    $Output = $_.Exception.Message  
    $DateStamp = get-date -Format "yyy-MM-dd hh:mm:ss"
    $subject = 'Script Execution Failed'
    $mailbody =" Hi Team,<br>
    @GEWAESSERSCHUTZ  application cannot be monitored now dues to process fail. <br><br>
    Please Monitor the application manually
    <br><br> Regards,<br><br>Automation Team"
    #send-MailMessage -To $MailTo.split(";") -Cc $MailCC.split(";") -Subject $subject -Body $mailbody -SmtpServer $SMTPServer -BodyAsHtml -From $MailFrom -Port $SMTPPort
    Add-Content -Path $LogfilePath -Value "$DateStamp | Script execution failed. $Output"  
    WriteToLog "Error" "Error captured $Output"
    
}
finally
{
    (New-Object -COM 'Shell.Application').Windows() | Where-Object {
    $_.Name -like '*Internet Explorer*'
    } | ForEach-Object {
        $_.Quit()
    } 
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ie)| Out-Null
    Start-Sleep -seconds 3
    if($LblUsername -ne $null)
    {
    Remove-Variable -name LblUsername
    }
}