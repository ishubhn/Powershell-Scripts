Function Convert-CSVtoExcel
{​
    [CmdletBinding()]
    param(
            [Parameter(Position=0,mandatory=$true)]  [String]   $DefaultFilePath,
            [Parameter(Position=1,mandatory=$true)]  [String]   $CSVSourceFileName,
            [Parameter(Position=2,mandatory=$true)]  [String]   $InterMediateXLSFileName
           )
    try
    {​
            Write-IntoLog ($(Get-Date -Format 'u') + "| - Default File Path is $DefaultFilePath")

            #Define Temp File Name Path 
            $sCSVSourceFileNamePath = $DefaultFilePath + $CSVSourceFileName
            Write-IntoLog ($(Get-Date -Format 'u') + "| - CSV File is located at $sCSVSourceFileNamePath") 
            
            #DEFINE INTERMEDIATE FILE NAME
            $sInterMediateXLSFileNamePath = $DefaultFilePath + $InterMediateXLSFileName
            Write-IntoLog ($(Get-Date -Format 'u') + "| - The Intermediate Excel File Name is located at $sInterMediateXLSFileNamePath") 
            
            #Define the Temp File Name Created and set it up to delete the file for the Next Run
            if (Test-Path $sInterMediateXLSFileNamePath) 
            {​
              Remove-Item $sInterMediateXLSFileNamePath -Force -ErrorAction SilentlyContinue
              Write-IntoLog ($(Get-Date -Format 'u') + "| - The Intermediate Excel File from the Previous Run Has Been Deleted") 
            
            }​ 
            # Create a New Excel object to load into Excel
            $oCSVToExcel = New-Object -ComObject Excel.Application 
            
            #Set The Excel Visibilty to FALSE
            $oCSVToExcel.Visible = $False
            
            # change thread culture
            [System.Threading.Thread]::CurrentThread.CurrentCulture = 'en-US'
            
            #Save the Converted CSV file into a Temp file for further Formatting as per the Required Formats
            $oCSVToExcel.Workbooks.Open($sCSVSourceFileNamePath).SaveAs($sInterMediateXLSFileNamePath,51)
            
            #Clear the oCSVToExcel Memory by Clearing the reference
            $oCSVToExcel.Quit()
            Return "Temporary Excel File Generated"
    }​
    
    catch [System.Exception] 
       {​ 
            $sExceptionType = $_.Exception.GetType().FullName
            $sExpectionMessage = $_.Exception.Message
            Write-IntoLog ($(Get-Date -Format 'u') + "| - ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Exception Message~~~~~~~~~~~~~~~~~~~~~") 
            Write-IntoLog ($(Get-Date -Format 'u') + "| - Exception in the Convert-CSVtoExcel Function Block") 
            Write-IntoLog ($(Get-Date -Format 'u') + "| - Exception Type: $sExceptionType") 
            Write-IntoLog ($(Get-Date -Format 'u') + "| - Exception Message: $sExpectionMessage")
            Write-IntoLog ($(Get-Date -Format 'u') + "| - ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Exception Message~~~~~~~~~~~~~~~~~~~~~") 
       }​     
}​
    