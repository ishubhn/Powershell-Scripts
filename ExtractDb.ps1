<#param(
    [parameter(mandatory=$true)]
    [string]$toPath,
    [parameter(mandatory=$true)]
    [string]$studyId,
    [parameter(mandatory=$true)]
    [string]$currentPath
)#>

$sISODateTime = Get-Date -Format yyyyMMMdd_HHmmss

# FUNCTION TO WRITE INTO A LOG FILE 
Function Write-IntoLog
{
   Param ([string]$logstring)
   
   try
        {
            $LogTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $Text = "$LogTime" + " | - " + " $Message"

            if (!(Test-Path -Path $LogFilePath))
            {
                New-Item -Path $LogFilePath -ItemType File | Out-Null
            }

            Add-Content -Path $LogFilePath -Value $logstring
        }
    catch
    { }
}

Function Extract-ConfigFileData
{  
    Try
    {
        $settingsFilePath = $env:USERPROFILE + "\Documents\Automation Anywhere Files\Automation Anywhere\My Docs\CTS\CTS\CTS EMEA USER_ACCESS_MEDIDATA\Medidata_PS_Config.xml"
        [xml]$configData = Get-Content -Path $settingsFilePath
        $settingsNode = $configData.SelectSingleNode("//config")
        $Script:LogFilePath = $settingsNode.DeveloperLog
        $Script:OracleDll = $settingsNode.OracleDll
        $Script:DBQuery = $settingsNode.DBQuery
        $Script:UserName = $settingsNode.UserName
        $Script:Password = $settingsNode.Password
        $Script:data_source = $settingsNode.DataSource
        $Script:sheetName = $settingsNode.WorksheetName
    }
    Catch [System.Exception] 
    {
        $sExceptionType = $_.Exception.GetType().FullName
        $sExpectionMessage = $_.Exception.Message
        
        Write-IntoLog ("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Exception Message~~~~~~~~~~~~~~~~~~~~~")
        Write-IntoLog ("Exception in the Extract-ConfigFileData Function Block")
        Write-IntoLog ("Exception Type: $sExceptionType")
        Write-IntoLog ("Exception Message: $sExpectionMessage")
        Write-IntoLog ("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Exception Message~~~~~~~~~~~~~~~~~~~~~")
    }
}

Write-IntoLog ("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
Write-IntoLog ("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~START OF A NEW SESSION @ $sISODateTime~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
Write-IntoLog ("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

#Extract query from database and store in temp csv file
Function Extract-OracleToCSV
{
    [CmdletBinding()]
    Param(
            [Parameter(Position=0,mandatory=$true)]  [String] $DSN,
            [Parameter(Position=0,mandatory=$true)]  [String] $UserID,
            [Parameter(Position=0,mandatory=$true)]  [String] $Password,
            [Parameter(Position=0,mandatory=$true)]  [String] $SQLQuery
         )

    Try 
    {
        $dllPath = $env:USERPROFILE + "\Documents\Automation Anywhere Files\Automation Anywhere\My Scripts\CTS\CTS\CTS EMEA USER_ACCESS_MEDIDATA\" + $OracleDll
        
        Add-PSSnapIn Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null

        #Oracle Database Connection - Load the assembly file (Oracle.DataAccess.dll) from the server
        Add-Type -Path $dllPath
        #Data Source
        #Prod server
        #e11sdm01-scan.pharma.aventis.com:1527/O1P42534.pharma.aventis.com
        $data_source = $data_source

        $connection_string = "User Id=$UserName;Password=$Password;Data Source=$data_source"

        #User Details Extraction Query
        #"Select * from tscadm.rave_$StudyId"
        $CommandText = $DBQuery + $studyId

        #Initiate the process
        $cn = New-Object Oracle.ManagedDataAccess.Client.OracleConnection -ArgumentList $connection_string

        $cn.Open()

        if($cn.State-eq'Open')
        {
            Write-IntoLog ("Oracle Database Connected")
        }

        #Load the Oracle Data in a Dataset
        $OracleCommand = New-Object -TypeName Oracle.ManagedDataAccess.Client.OracleCommand
        $OracleCommand.CommandText = $CommandText  
        $OracleCommand.Connection = $cn

        #Use the Oracle Data Adapter  
        $OracleDataAdapter = New-Object -TypeName Oracle.ManagedDataAccess.Client.OracleDataAdapter  
        $OracleDataAdapter.SelectCommand = $OracleCommand

        #Create a new Data set  
        $DataSet = New-Object -TypeName System.Data.DataSet  
        $OracleDataAdapter.Fill($DataSet) | Out-Null

        #Dispose the connection  
        $OracleDataAdapter.Dispose()  
        $OracleCommand.Dispose()  
        $cn.Close()
        
        Write-IntoLog ("All Oracle Connections Disposed and Closed")

        $iDataSetCount = $DataSet.Tables[0].Rows.Count

        If ($iDataSetCount -eq 0)
        {
            Write-IntoLog ("No Records Can Be retrieved. - Exception")
        }
        Else
        {
            $CSVFilePath = $currentPath
            
            Write-IntoLog ("The Path for placing the CSV File is: - $CSVFilePath")
            
            #COUNTING THE NUMBER OF RECORDS IN THE TABLE FOR THE MASTER EXTRACTION 41 Query
            Write-IntoLog ("The Records count extracted from Database is $iDataSetCount")
            
            <#$data= $DataSet.Tables[0]
            $data | Export-CSV -path $CSVFilePath -NoTypeInformation -Encoding UTF8
            Write-IntoLog ("The Records has been successfully extracted to the CSV file")
            #>

            For ($dataSetIndex = 0; $dataSetIndex -lt $DataSet.Tables.Count; $dataSetIndex++) 
            {
                $data = $DataSet.Tables[$Dsindex]
                $data | Export-CSV -path $CSVFilePath -NoTypeInformation -Encoding UTF8
                Write-IntoLog ("The Records has been successfully extracted to the CSV file")
                <#if($Dsindex -ne 0)
                {
                    $workbook.worksheets.Add() | Out-Null #Create new worksheet
                }#>
            }

            Return "Oracle Extraction Successfull"
        }
    }
    Catch [System.Exception]
    {
        $sExceptionType = $_.Exception.GetType().FullName
        $sExpectionMessage = $_.Exception.Message
        
        Write-IntoLog ("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Exception Message~~~~~~~~~~~~~~~~~~~~~")
        Write-IntoLog ("Exception in the Extract-OracleToCSV Function Block")
        Write-IntoLog ("Exception Type: $sExceptionType")
        Write-IntoLog ("Exception Message: $sExpectionMessage")
        Write-IntoLog ("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Exception Message~~~~~~~~~~~~~~~~~~~~~")
    }
}

Function Convert-CSVtoExcel
{
    [CmdletBinding()]
    Param(
            [Parameter(Position=0,mandatory=$true)]  [String]   $DefaultFilePath,
            [Parameter(Position=1,mandatory=$true)]  [String]   $CSVSourceFileName,
            [Parameter(Position=2,mandatory=$true)]  [String]   $InterMediateXLSFileName
        )
    
    Try
    {
        Write-IntoLog ("Default File Path is $DefaultFilePath")

        #Define Temp File Name Path 
        $sCSVSourceFileNamePath = $DefaultFilePath + $CSVSourceFileName
        Write-IntoLog ("CSV File is located at $sCSVSourceFileNamePath")

        #DEFINE INTERMEDIATE FILE NAME
        $sInterMediateXLSFileNamePath = $DefaultFilePath + $InterMediateXLSFileName
        Write-IntoLog ("The Intermediate Excel File Name is located at $sInterMediateXLSFileNamePath")

        #Define the Temp File Name Created and set it up to delete the file for the Next Run
        If (Test-Path $sInterMediateXLSFileNamePath) 
        {
            Remove-Item $sInterMediateXLSFileNamePath -Force -ErrorAction SilentlyContinue
            Write-IntoLog ("The Intermediate Excel File from the Previous Run Has Been Deleted") 
        }

        # Create a New Excel object to load into Excel
        $oCSVToExcel = New-Object -ComObject Excel.Application

        #Set The Excel Visibilty to FALSE
        $oCSVToExcel.Visible = $False

        #Worksheet
        $Workbook = $oCSVToExcel.Workbooks.Add(1)
        $Worksheet = $Workbook.worksheets.Item(1)
        $Worksheet.Name = $sheetName
        <#$worksheet = $workbook.worksheets.Item(1)
        $worksheet.Name = $Table.TableName#>
        
        # change thread culture
        [System.Threading.Thread]::CurrentThread.CurrentCulture = 'en-US'

        #Save the Converted CSV file into a Temp file for further Formatting as per the Required Formats
        $oCSVToExcel.Workbooks.Open($sCSVSourceFileNamePath).SaveAs($sInterMediateXLSFileNamePath,51)

        #Clear the oCSVToExcel Memory by Clearing the reference
        $oCSVToExcel.Quit()

        Return "Temporary Excel File Generated"
    }
    Catch [System.Exception]
    {
        $sExceptionType = $_.Exception.GetType().FullName
        $sExpectionMessage = $_.Exception.Message
        
        Write-IntoLog ("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Exception Message~~~~~~~~~~~~~~~~~~~~~")
        Write-IntoLog ("Exception in the Convert-CSVtoExcel Function Block")
        Write-IntoLog ("Exception Type: $sExceptionType")
        Write-IntoLog ("Exception Message: $sExpectionMessage")
        Write-IntoLog ("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Exception Message~~~~~~~~~~~~~~~~~~~~~")
    }
    Finally {
        Stop-Process -Name *excel* -force
    }
}

Function Move-File {
    Param (
        [parameter(Mandatory=$true)]
        [string]$From,
        [parameter(Mandatory=$true)]
        [string]$To
    )
    #Wait for 2 secs
    Start-Sleep -Seconds 2
    $input = $PSScriptRoot + "\Test.csv"
    #Move file to respective study folder
    #Move-Item "$PSScriptRoot\export.xls" "$toPath" -Force
    Move-Item $From $To
    Remove-Item $input
}

#Main Method
Function Run-MainMethod {
    Extract-ConfigFileData
    Extract-OracleToCSV
    Convert-CSVtoExcel
}