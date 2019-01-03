<#
.Description
Gets the users logged into the computer. Will display 'OFFLINE' for the user if the computer is not reachable via ping.

.Parameter ComputerName
The computer to get the logged in users from
#>
Function Get-RdpSessions 
{
    param(
        [string]$ComputerName 
    )

    If(Test-Connection -Computername $ComputerName -BufferSize 16 -Count 1 -Quiet){
        Try{
            $processInfo = Get-WmiObject -Query "select * from win32_process where name='explorer.exe'" -ComputerName $ComputerName -ErrorAction Stop

            $users = $processInfo | ForEach-Object { $_.GetOwner().User } | Sort-Object -Unique | ForEach-Object { New-Object psobject -Property @{LoggedOn=$_} } | Select-Object -ExpandProperty LoggedOn
            
            if($users -eq $null){
                $loggedIn = ''
                $creationDate = ''
            }
            else{
                $creationDate = $processinfo | ForEach-Object { Convert-WMIDateTime $_.creationdate}

                $loggedIn = $users
            }
        }
        Catch{
            $Exception = $_

            If($Exception.Exception.Message -like '*RPC server*'){
                $loggedIn = 'RPC ERROR'
            }
            else{$loggedIn = $Exception.Exception.Message}
        }  
    }else{$loggedIn = 'OFFLINE'}

    return $ComputerName, $loggedIn, $creationDate
}


<#
.Description
Gets an object array of computers and the users logged into them.

.Parameter ComputerList
List of computers to get logged in users from.
#>
Function Get-ComputerStatus 
{
    [cmdletbinding()]
    param([parameter(Mandatory)][string[]]$ComputerList,
        [switch] $Custom)

    $statusList = @()
    
    $total = $ComputerList.Count
    $current = 0

    foreach ($computer in $ComputerList){
        $current++
        Write-Progress -Activity 'Searching for systems' -Status "$current of $total : $computer" -PercentComplete (($current / $total) * 100)

        Write-Verbose $computer

        $computerStatus = New-Object PSObject
        $RdpSessions = Get-RdpSessions $computer

        $computerStatus | Add-Member NoteProperty -Name 'Computer' -Value $RdpSessions[0]
        $computerStatus | Add-Member NoteProperty -Name 'User' -Value $RdpSessions[1]
        $computerStatus | Add-Member NoteProperty -Name 'LogedInSince' -Value $RdpSessions[2]

        $statusList += $computerStatus        
    }
    
    return $statusList
}

<#
.Description
Converts from WMI time format to date time.

.Parameter WMIDateTime
WMI formatted time to convert.
#>
Function Convert-WMIDateTime($WMIDateTime){
    # WMI Event Log DateTime format is 20150625134302.000000+060 where the +060 is the time zone (BST in this case)
    $DateTimeArray = $WMIDateTime -split "\."
    $DateTimeOnly = $DateTimeArray[0] # This is the bit before the "."
    if($DateTimeArray[1] -like "*+*"){ # Is the time zone plus or minus
        $TimeZoneSymbol = "+"
    }else{
        $TimeZoneSymbol = "-"
    }
    # Split out the time zone from the zeros and convert from minutes to hours
    $WMISummerTime = ($DateTimeArray[1] -split "\$TimeZoneSymbol")[1] / 60
    # Build the formatted date/time string up ready for conversion to a datetime variable
    $FormattedDateTime = ($DateTimeOnly + $TimeZoneSymbol + $WMISummerTime)
    [datetime]::ParseExact($FormattedDateTime,"yyyyMMddHHmmssz",[System.Globalization.CultureInfo]::InvariantCulture)
}

<#
.Description
Enter a list of computers, split by returns (shift+enter), or a single computer to get an array of the entered computers.
#>
Function Get-LoggedInUsers
{
    [cmdletbinding()]
    param(
        [switch] $Custom,
        [array] $Computer
    )

    # Enter a list of computer split by new lines or a single computer to get an array of the entered computers
    If($Computer -eq $null){
        Do
        {
            $computerArray = (Read-Host "Please enter the computer name")
        }While($computerArray -eq '')
    }else{$computerArray = $Computer}

    $computerArray = $computerArray.Split([Environment]::NewLine)
    $computerArray = $computerArray | ? {$_ -ne ""}

    Get-ComputerStatus -ComputerList $computerArray -Custom:$Custom
}

Get-LoggedInUsers
