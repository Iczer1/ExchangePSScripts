#requires -version 3
<#
.SYNOPSIS
    Checks managed availability states on a Exchange 2013 servers and report back a status that is consistent with nagios plug ins
.DESCRIPTION
    Check server Component state (Get-ServerComponentState)
    Check overall system health set (Get-HealthReport )
        if issues then check all health monitors (Get-ServerHealth)
    Reports the results of all checks and removes the PSSession
.PARAMETER server 
    Exchange server to run the check against, defualts to the $env:COMPUTERNAME
.EXAMPLE
    PS> .\check_Ex2013_ManagedAvailability.ps1

    Result
    -----------------------------------------
    CRITICAL: 3 Health sets are NOT HEALTHY
    --- Degraded HeathSets---

     ActiveSync.Protocol is Unhealthy 
        --- Associated Monitors---
        ActiveSyncDeepTestMonitor is Unhealthy

     IMAP.Protocol is Unhealthy 
        --- Associated Monitors---
        ImapDeepTestMonitor is Unhealthy

     POP.Protocol is Unhealthy 
        --- Associated Monitors---
        PopDeepTestMonitor is Unhealthy
.NOTES
    AUTHOR: John mello
    CREATED: 2015-04-09
    UPDATES :
        John Mello : 09/08/2015
            Change assoicated monitor logic so it ignores disabled monitors
        John Mello : 03/18/2016
            Get-ServerComponentState returns deserialized info when using a PSSession to Exchange in a standard shell, so you can't pull the request info since it's deserialized
            To counter act this I pull the info from the registry and AD and combine the results
            Script now only warns on bad health sets, and throws a critical on server component states
            When ServerWideOffline is InActive we only report that
        John Mello : 09/11/2017
            Updated Get-Ex2013DeserializedComponentStateHistory to get the proper info
#>

[CmdletBinding()]
param(
    #[string] $server = $env:COMPUTERNAME,
    [string] $server = ''
    #[switch] $BlackListOnly = $False,
    #[switch] $ComponentOnly = $TRUE
)

#region functions
function nagios_exit{
    #Clean up Exchange Cmdlets
    Get-PSSession -Name EXCH -ErrorAction SilentlyContinue | Remove-PSSession
    switch ($EC){
        $OK { "OK: $OUT" }
        $WARNING { "WARNING: $OUT" }
        $CRITICAL { "CRITICAL: $OUT" }
        default { "UNKNOWN: $OUT" }
    }
    exit $EC
}

Function Get-Ex2013DeserializedComponentStateHistory {
<#
.SYNOPSIS
Return Exchange 2013 component states, their most recent requester, and other relevant information.
.DESCRIPTION
Return Exchange 2013 component states, their most recent requester, and other relevant information.
.PARAMETER NonActiveOnly
Only return results for components not in an active state.
.PARAMETER ServerFilter
Limit results to specific servers.

.EXAMPLE
PS > Get-Ex2013DeserializedComponentStateHistory | ft -AutoSize

Description
-----------
Show all component states for all servers along with their most recent requester.

.EXAMPLE
PS > Get-Ex2013DeserializedComponentStateHistory | Where {($_.Requesters).Count -gt 1 } | ft -AutoSize

Description
-----------
List all components with more than 1 historical requester

.NOTES
Author: Zachary Loeber
Requires: Powershell 3.0, Exchange 2013
Version History
1.0.0 - 12/07/2014
- Initial release

The component state cliff notes:
    Component State Requesters
    * HealthAPI - Reserved by managed availability (probably shouldn't ever use this if setting component states)
    * Maintenance
    * Sidelined
    * Functional
    * Deployment

    Multiple requesters can set a component state. 
    Note: When there are multiple requesters 'Inactive' is prioritized over 'Active'

    Global Components
    * ServerwideOffline - Overrules the states of all other components except for Monitoring and RecoveryActionsEnabled

    Managed Availability Components (I think)
    * Monitoring
    * RecoveryActionsEnabled

    Transport Components (Only components which can be in 'draining' state)
    * FrontendTransport
    * HubTransport

    All other components
    * AutoDiscoverProxy
    * ActiveSyncProxy
    * EcpProxy
    * EwsProxy
    * ImapProxy
    * OabProxy
    * OwaProxy
    * PopProxy
    * PushNotificationsProxy
    * RpsProxy
    * RwsProxy
    * RpcProxy
    * UMCallRouter
    * XropProxy
    * HttpProxyAvailabilityGroup
    * ForwardSyncDaemon
    * ProvisioningRps
    * MapiProxy
    * EdgeTransport
    * HighAvailability
    * SharedCache

.LINK
https://github.com/zloeber/Powershell
.LINK
http://www.the-little-things.net
#>

    [CmdLetBinding()]
    param(
        [Parameter(Position=0, HelpMessage='Only return results for components not in an active state.')]
        [switch]$NonActiveOnly,
        [Parameter(Position=1, HelpMessage='Limit results to specific servers.')]
        [string]$ServerFilter = '*',
        [Parameter(Position=2, HelpMessage='ADSI path to Exchange Org')]
        [string]$ADSIPath = "CN=Servers,CN=Exchange Administrative Group (FYDIBOHF23SPDLT),CN=Administrative Groups,CN=SUSQ,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=dsroot,DC=susq,DC=com"

    )

    begin {
        try {
            $ExchangeServers = Get-ExchangeServer $ServerFilter | Where {$_.AdminDisplayVersion -like 'Version 15.*'}
        }
        catch {
            Write-Warning "Get-Exchange2013ComponentStateHistory: Unable to enumerate Exchange 2013 servers!"
            break
        }
    }
    process {
       Foreach ($Server in $ExchangeServers) {
            Write-Verbose "Get-Exchange2013ComponentStateHistory: Processing Server - $($Server.Name)"
            try {
                $ComponentStates = Get-ServerComponentState $Server.Name
                if ($NonActiveOnly) {
                    $ComponentStates = $ComponentStates | Where {$_.State -ne 'Active'}
                    Write-Verbose "Get-Exchange2013ComponentStateHistory: Non-active components - $($ComponentStates.Count)"
                }
                If ($ComponentStates) {
                    #Attemtping to load the AD Module, which is needed to run Get-Adobject
                    #TODO replace with ADSI searcher 
                    Try {import-module -Name ActiveDirectory -ErrorAction Stop | Out-Null}
                    Catch {
                        Write-host "AD module cannot be loaded"
                    }
                    
                    #Get component state change requests from the registry
                    #TODO how do we get this working remotley? 
                    #tried this but it didn't work https://support.microsoft.com/en-us/help/314837/how-to-manage-remote-access-to-the-registry
                    #Maybe this http://windowsitpro.com/security/granting-users-read-access-registry
                    [ScriptBlock]$GetRegEntries = {
                        $Keys = Get-ChildItem HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\ServerComponentStates
                        $Table = @()
                        foreach ($Key in $Keys) {
                            $States = @()
                            $Requesters = $Key.property
                            Foreach ($Requester in $Requesters) {  
                                $States += (Get-ItemProperty $Key.PSPATH).$Requester | 
                                            Select-Object @{N='Component';E={$Key.PSChildName}}, 
                                                    @{N='LastRequester';E={$Requester}}, 
                                                    @{N='LastChanged';E={Get-Date ([int64]($_ -split ":")[2])}},
                                                    @{N='LastSource';E={"Registry"}}

                            }
                            $Table += $States |
                                Sort-Object LastChanged -Descending |
                                Select-object -first 1
                        }
                        $Table
                    }
                    Try {
                        [Array]$RegEntries = Invoke-Command -ComputerName $Server.Name -ScriptBlock $GetRegEntries -ErrorAction stop |
                                                Select-Object Component,LastRequester,LastChanged,LastSource |
                                                Where-Object Component -in $ComponentStates.Component
                    }
                    Catch {
                        Write-Warning "Get-Exchange2013ComponentStateHistory: Can't pull registry info" 
                    }
                
                    #Get component state change requests from AD
                    Try {
                        $ADSIInfo = Get-ADObject ("CN=" + $Server.Name + "," + $ADSIPath)  -Properties msExchComponentStates -ErrorAction Stop
                        [Array]$ADSIEntries =
                            Foreach ($Entry in $ADSIInfo.msExchComponentStates) {
                                $Entry |
                                    Select-Object @{N='Component';E={($_ -split ":")[1]}}, 
                                            @{N='LastRequester';E={($_-split ":")[2]}}, 
                                            @{N='LastChanged';E={Get-Date ([int64]($_ -split ":")[4])}},
                                            @{N='LastSource';E={"AD"}} |
                                                Where-Object Component -in $ComponentStates.Component
                            }
                    }
                    Catch {
                        Write-Warning "Get-Exchange2013ComponentStateHistory: Can't access AD" 
                    }
                
                    #Combine the AD and Registry entries
                    $CombinedEntries = $RegEntries + $ADSIEntries
                    #Crate  list of Components
                    $UniqueComponents =  ($ComponentStates | select Component -Unique).Component
                    #Sort the combine entries and if multiple entries for one component then choose the newest one
                    $Finallist =@()                     
                    foreach ($Unique in $UniqueComponents) {
                        $Finallist += 
                            $CombinedEntries | 
                                where-object Component -eq $Unique |
                                Sort-object LastChanged -Descending |
                                Select-object -First 1
                    }
                    #Add the requester, cahnged, and source to the oringal info and return it
                    $ComponentStates | Select-object Component, State,
                                       @{N='LastRequester';E={($Finallist | where Component -eq $_.Component).LastRequester}}, 
                                        @{N='LastChanged';E={($Finallist | where Component -eq $_.Component).LastChanged}},
                                        @{N='LastSource';E={($Finallist | where Component -eq $_.Component).LastSource}}
                }

            }
            catch {
                Write-Warning "Get-Exchange2013ComponentStateHistory: Unable to get component state for $($Server.Name)!"
            }
        }
    }
    end {}
}
#endregion

#region Dependencies
#Load Exchange Cmdlets via a PSSsession
If ((Get-Command Get-Mailbox -ErrorAction SilentlyContinue -Verbose:$false) -ne $Null) {Write-Verbose "Exchange PSSnapin is loaded, Proceeding with script"}
Else {
    Try {
        $CAS = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$Server/PowerShell/" -Authentication Kerberos -Name EXCH -ErrorAction Stop
        Import-PSSession $CAS -allowclobber -ErrorAction Stop -DisableNameChecking -WarningAction SilentlyContinue | Out-Null
    }
    Catch {
        Get-PSSession -Name EXCH -ErrorAction SilentlyContinue | Remove-PSSession
        Write-host "Can't load Exchange PSSession"
        $_.Exception.message
        exit 3
    }
}

#endregion

#region Variables
#Exit variables for easier usage
$OK = 0
$WARNING = 1
$CRITICAL = 2
$UNKNOWN = 3
#Variables for check reports
$EC = $OK
$OUT = ""
$OUT_DETAIL = ""
$TESTED = ""
$TESTED_DETAIL = ""
#Whitelist for Health Report
#[Array]$HealthReportWhitelist = @("Transport.ServerCertExpireSoon.Monitor")
[Array]$HealthReportWhitelist = @()
#Pull Server info
Try {$ExServer = Get-ExchangeServer $server -ErrorAction Stop}
Catch {
    $OUT = "Get-ExchangeServer cmdlet failed`n" + $error[0]
    $EC = $UNKNOWN
    nagios_exit
}
#endregion

#region Managed availability checks

#get-servercomponentState 
Write-verbose ((get-date).toString() + " Checking Server component state")
#Casting to Array forces the count prpoerty even if there is only one entry
<#
if ($BlackListOnly) {
    $DegradedComponets = $NULL
} 
Else {
#>
[Array]$DegradedComponets = Get-Ex2013DeserializedComponentStateHistory -ServerFilter $Server -NonActiveOnly
#}

if($DegradedComponets) {
    $tmp = "--- Inactive Component ---`n"
    If ($DegradedComponets.component -contains "ServerWideOffline") {
        $tmp += "--- ServerWideOffline Found, reporting just that ---`n"
        $DegradedComponets = $DegradedComponets | Where component -eq "ServerWideOffline"
    }
    # Atleast one componets is inactive, list details
    foreach($Component in $DegradedComponets) {
        $tmp += "Component: $($Component.Component)`n"
        $tmp += "LastSource: $($Component.LastSource)`n"
        $tmp += "LastRequester: $($Component.LastRequester)`n"
        $tmp += "LastChanged: $($Component.LastChanged)`n`n"
        #$tmp += "Component: $($Component.Component)`nLastSource: $($Component.LastSource)`nLastRequester: $($Component.LastRequester)`nLastChangedt: $($Component.LastChanged)`n`n"
    }
    $OUT += "$($DegradedComponets.count) components in a NON ACTIVE state; "
    $OUT_DETAIL += $tmp
    $EC = $CRITICAL
}
Else {
    #$OUT_DETAIL += "$((Get-ServerComponentState -Identity $Server).count) Server Component States are ACTIVE`n"
    #$TESTED_DETAIL += "$((Get-ServerComponentState -Identity $Server).count) Server Component States are ACTIVE`n"
    $TESTED_DETAIL += "All Server Component States are ACTIVE`n"
    $TESTED += "Server Component State; "
}
Write-verbose ((get-date).toString() + " finishing Get-ServerComponentState check")

#Get-HealthReport
Write-verbose ((get-date).toString() + " Checking overall Health report")
# Get health checks for role
$HealthReport = Get-HealthReport -Identity $Server
#Casting to Array forces the count property even if there is only one entry
<#
if ($BlackListOnly) {
    [Array]$DegradedHealthReport = $HealthReport | 
                                    Where-object {$_.AlertValue -ne "Healthy" -and $_.AlertValue -ne "Disabled"} | 
                                    Where-Object {$_.Name -eq 'QueueMailboxRetryMonitor'}
}

Else {
#>
[Array]$DegradedHealthReport = $HealthReport | Where-object {$_.AlertValue -ne "Healthy" -and $_.AlertValue -ne "Disabled"} 

#}
if($DegradedHealthReport) {
    $tmp = ""
    Write-verbose ((get-date).toString() + "Unhealthy HealthSets found, gathering associated monitors")
    foreach($HealthCheck in $DegradedHealthReport) {

        [Array]$AssociatedMonitors = Get-ServerHealth -HealthSet $HealthCheck.HealthSet -Server $Server | 
                                        Where-object AlertValue -Notin @("Healthy",'Disabled') | 
                                        #WhiteList Entries
                                        Where-Object Name -notin $HealthReportWhitelist
        If ($AssociatedMonitors) {
            $tmp += "$($HealthCheck.HealthSet) is $($HealthCheck.AlertValue) `n"
            $tmp += "`* Associated Monitors *`n"
            $AssociatedMonitors | 
                Foreach-object {$tmp += "-> $($_.Name) is $($_.AlertValue)`n"}
        }#If ($AssociatedMonitors) {
        #$tmp += "`n"
    }#foreach($HealthCheck in $DegradedHealthReport)
    if ($tmp) {
        $OUT += "$($DegradedHealthReport.count) HealthSets are NOT HEALTHY; "
        $OUT_DETAIL += $tmp
        If ($EC -ne $CRITICAL) {$EC = $WARNING}
        Else {<#Keep Crtical State#>}
    }#if ($tmp)
}#if($DegradedHealthReport)
Else {
    $TESTED_DETAIL += "$($HealthReport.count) Health sets are HEALTHY`n"
    $TESTED += "Health Report; "
}#Else
Write-verbose ((get-date).toString() + " finishing Get-HealthReport check")

#endregion

#region display report
if ($EC -eq $OK){
    $TESTED = $TESTED.trimend("; ")
    $OUT = "$TESTED`n$TESTED_DETAIL"
}

$OUT = $OUT.trimend("; ")

if ($OUT_DETAIL -ne ""){
    $OUT += "`n$OUT_DETAIL"
}

nagios_exit
#endregion
