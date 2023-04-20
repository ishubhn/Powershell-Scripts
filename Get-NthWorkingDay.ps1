Function Get-NthWorkingDay {
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory = $true)]
        [datetime]$StartDate,
        [Parameter(Mandatory = $true)]
        [int]$NthWorkingDay 
    )
    # Calculate the number of working days
    $Days = 0
    $currentDayOfWeek = (Get-Date -Date $StartDate).DayOfWeek 
    # If start date falls on non working day start iteration from 0
    if ($currentDayOfWeek -eq "Sunday" -or $currentDayOfWeek -eq "Saturday") {
        [int] $WorkingDays = 0 
    }
    else {
        [int] $WorkingDays = 1 
    } 
    
    while ($WorkingDays -le $NthWorkingDay) {
        $Days++ 
        $Date = $StartDate.AddDays($Days)
        if ($Date.DayOfWeek -ne "Saturday" -and $Date.DayOfWeek -ne "Sunday") { $WorkingDays++ } 
    } 
    # Return the nth working day 
    return $StartDate.AddDays($Days)
}