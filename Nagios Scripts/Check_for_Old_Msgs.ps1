<#
.SYNOPSIS
    Used to check a specified transport service's queues for messages in a retry state with a DateReceived older then specified
.DESCRIPTION
    Before attempting any work the script will verify that the Exchange PSSnapin is loaded by testing for the existence of Get-ExchangeServer. If the command is not present the script will then try to load the Exchange PSSnapin, if that fails the script will exit. Otherwise it will continue.
    The script will then check the Queue on the specified exchange transport service for any queue (That isn't of the DeliveryType ShadowRedundancy)
        in a retry state
        count more then 0 
        That has Messages
            In a retry state
            That have a Datereceived less than or equal to the current date at runtime minus the WarningAgeinMins
    If any messages are found the script will report as such
.PARAMETER Server 
    DNS name of the Hub transport server to check
.PARAMETER WarningAgeinMins 
    The warning Age in minutes, should be lower than the Critical Count
.PARAMETER CriticalAgeinMins 
    The critical Age in minutes 
.EXAMPLE
    PS C:\> .Check_TransportQueue -Server XCHHTBAL500 -WarningAgeinMins 60 -CriticalAgeinMins 1440
    
    Description
    -----------
    For transport service called XCHBAL500, check any non ShadowRedundancy queue in a retry state with a message count greater than 1 containing any messages in a retry state with a DateReceived older than 60 minutes from the run time of the script
.NOTES
    AUTHOR: John Mello
    CREATED : 08/23/2017
    CREATED BECAUSE: Needed a way to alert on old messages in the queues
    UPDATES :
        2017-10-16 : John Mello
            Switched from checking the queues on each server for a retry and then getting the messages from those queues to directly using the get-message command to get messages in a retry state.
            This was more efficient and better caught stuck messages
#>
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $True)] 
    [string]$Server,
    
    [Parameter(Mandatory = $True)] 
    [int]$WarningAgeinMins = 60,
    
    [Parameter(Mandatory = $True)]
    [int]$CriticalAgeinMins = 1440
)

IF ($WarningAgeinMins -gt $CriticalAgeinMins) {
    Write-Warning "WarningAgeinMins ($($WarningAgeinMins)) is greater than the CriticalAgeinMins ($($CriticalAgeinMins))"
    Write-Warning "Setting CriticalAgeinMins to the same Value as WarningAgeinMins" 
    $CriticalAgeinMins = $WarningAgeinMins
}#If

#region Dependencies
#Load Exchange Cmdlets via a PSSsession
If ((Get-Command Get-Mailbox -ErrorAction SilentlyContinue -Verbose:$false) -ne $Null) {Write-Verbose "Exchange PSSnapin is loaded, Proceeding with script"}
Else {
    Try {
        $CAS = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$Server/PowerShell/" -Authentication Kerberos -Name EXCH -ErrorAction Stop
        Import-PSSession $CAS -allowclobber -ErrorAction Stop -DisableNameChecking -WarningAction SilentlyContinue | Out-Null 3> $null
    }#Try
    Catch {
        Get-PSSession -Name EXCH -ErrorAction SilentlyContinue | Remove-PSSession
        Write-host "Can't load Exchange PSSession"
        $_.Exception.message
        exit 3
    }#Catch
}#else
#endregion
$CurrentDate = Get-date
$WarningDate = $CurrentDate.addminutes(-$WarningAgeinMins)
$CriticalDate = $CurrentDate.addminutes(-$CriticalAgeinMins)

Try {
    [array]$OldMSGCheck = Get-Message -Server $Server -Filter {Status -eq 'Retry'} -ResultSize Unlimited |
        Where-object DateReceived -le $WarningDate
}#Try
Catch {
    Write-host "Problem accessing queue"
    $_.Exception.message
    Get-PSSession -Name EXCH -ErrorAction SilentlyContinue | Remove-PSSession
    exit 3

}#Catch

#Remove Session now that we got the info we want
Get-PSSession -Name EXCH -ErrorAction SilentlyContinue | Remove-PSSession

If (-not $OldMSGCheck) {
    #"No old messages found, reporting an OK status"
    Write-Output "OK: All Queues have no messages older than $(get-date $WarningDate -format G)"
    Exit 0
}#If (-not $OldMSGCheck)
Elseif ($OldMSGCheck | Where DateReceived -gt $CriticalDate) {
    #Messages found older then the Warning Date but less then the Critical Date
    #Since we already searched for messages older then the warning date, we just need to see if they are gt then the critical date
    Write-Output "WARNING: $($OldMSGCheck.count) messages older than $(get-date $WarningDate -format G)"
    $OldMSGCheck |
        Select-Object OriginalFromAddress, Subject, DeferReason, MessageSourceName, Identity
    Exit 1
}#Elseif ($OldMSGCheck | Where DateReceived -lt $CriticalDate)
Elseif ($OldMSGCheck | Where DateReceived -le $CriticalDate) {
    #Messages found older than or equal to the Critical Date
    Write-Output "CRITICAL: $($OldMSGCheck.count) messages older than $(get-date $CriticalDate -format G)"
    $OldMSGCheck |
        Select-Object OriginalFromAddress, Subject, DeferReason, MessageSourceName, Identity
    Exit 2
}#Elseif ($OldMSGCheck | Where DateReceived -ge $CriticalDate) 
Else {
    #Shouldn't happen but catch it just to be safe
    Write-Output "UNKNOWN: Issue checking queues"
    Exit 3
}#Else
