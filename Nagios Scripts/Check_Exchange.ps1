<#
.SYNOPSIS
    Used to check various server components on Exchange 2010 and 2013 servers and report back a status that is consistent with nagios plug ins
.DESCRIPTION
    Creates A PSSesson to the specifed Exchange server and performs the following
        Verifies all servies for all installed roles are running (Test-ServiceHealth)
        If this is a mailbox server
            verify that all mailbox databses and public folder databses (2010) are mounted
            Test mail flow (Test-Mailflow)
    Reports the results of all checks and removes the PSSession
.PARAMETER server 
    Exchange server to run the check agaiasnt, defualts to the $env:COMPUTERNAME
.EXAMPLE
    PS> .\check_ex2013.ps1

    Result
    -----------------------------------------
    OK: Required Services; Databases Mounted; Server Health; Server Component State; Health Report
    6 mounted databases
    No mailbox to test mail-flow
.NOTES
    AUTHOR: Jeff Chung
    CREATED: 2011-05-17
    UPDATES :
       2015-04-08 : John mello
        Added help and regions
        Changed $_ error catches to TRY/CATCH blocks
        Cleaned up to script to make it more "PowerShelly"
       2016-01-11 : John Mello
        Get-mailboxdatabase in 2013 only returns active databases, changed to get-mailboxdatabasecopystatus and edited code logic
       2016-07-15 : John Mello
        Added Content index state check
        Updated Mail flow test
        Added check for mounted databases matching the activation preference
       2016-09-28 : John Mello
        Fixed mailflow check for servers with lagged databases
       2017-10-11 : John Mello
        Ensured that the script is veiwing the entire AD forest
#>

[CmdletBinding()]
param(
    [string] $server
    #[string] $server = $env:COMPUTERNAME
)

#region functions
function nagios_exit {
    #Clean up Exchange Cmdlets
    $ExchSession = Get-PSSession -Name EXCH
    If ($ExchSession) {Remove-PSSession $ExchSession}
    switch ($EC) {
        $OK { "OK: $OUT" }
        $WARNING { "WARNING: $OUT" }
        $CRITICAL { "CRITICAL: $OUT" }
        default { "UNKNOWN: $OUT" }
    }#switch ($EC) 
    exit $EC
}#function nagios_exit

#endregion

#region Dependencies
#Load Exchange Cmdlets via a PSSsession
If ((Get-Command Get-Mailbox -ErrorAction SilentlyContinue -Verbose:$false) -ne $Null) {Write-Verbose "Exchange PSSnapin is loaded, Proceeding with script"}
Else {
    Try {
        $CAS = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$Server/PowerShell/" -Authentication Kerberos -Name EXCH -ErrorAction Stop
        Import-PSSession $CAS -allowclobber -ErrorAction Stop -DisableNameChecking -WarningAction SilentlyContinue | Out-Null
    }#Try
    Catch {
        $OUT = "Can't load Exchange PSSession"
        $EC = $WARNING
        nagios_exit
    }#Catch
}#Else
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

#Ensure we are viewing the entire AD forest
Try {
    Set-ADServerSettings -ViewEntireForest:$true -ErrorAction Stop
}#Try
Catch {
    $OUT = "Set-ADServerSettings cmdlet failed`n" + $error[0]
    $EC = $WARNING
    nagios_exit
}#Catch

#Pull Server info
Try {$ExServer = Get-ExchangeServer $server -Status -ErrorAction Stop}
Catch {
    $OUT = "Get-ExchangeServer cmdlet failed`n" + $error[0]
    $EC = $WARNING
    nagios_exit
}#Catch
#endregion

#region check if required services are running
Write-verbose ((get-date).toString() + " starting required services check")
Try {
    $roles = Test-ServiceHealth -ErrorAction STOP -Server $server
    $tmp = ""
    $roles | 
        foreach-object {
        if ($_.RequiredServicesRunning -eq $false) {
            $_.ServicesNotRunning |
                foreach-object {
                $tmp += "$_ not running`n";
            }#foreach-object
        }#if ($_.RequiredServicesRunning -eq $false)
    }#foreach-object
    if ($tmp -ne "") {
        $OUT += "required services are not running; "
        $OUT_DETAIL += $tmp
        $EC = $CRITICAL
    }#if ($tmp -ne "")
    else {
        $TESTED += "Required Services; "
    }#Else

    Write-verbose ((get-date).toString() + " finished required services check")
}#try
Catch {
    $OUT = "Test-ServiceHealth cmdlet failed`n" + $error[0]
    $EC = $CRITICAL
    nagios_exit
}#Catch
#endregion

#region check if mailbox databases are mounted and Healthy and on the right server
if ($ExServer.IsMailboxServer -eq $true) {
    $tmp = ""
    $numHealthyDatabase = 0

    Write-verbose ((get-date).toString() + " starting Get-MailboxDatabaseCopyStatus check")

    Try {$databases = Get-MailboxDatabaseCopyStatus -Server $server -ErrorAction STOP}
    Catch {
        $OUT += "Get-MailboxDatabaseCopyStatus cmdlet failed`n" + $error[0]
        $EC = $WARNING
        nagios_exit
    }#Catch
    if ($databases) {
        $ShouldBeMounted = $databases | 
            Where ActivationPreference -eq 1 |
            Where ActiveDatabaseCopy -ne $server |
            Select DataBaseName, ActiveDatabaseCopy
        If ($ShouldBeMounted) {
            $EC = $WARNING
            $OUT += "$($ShouldBeMounted.count) DB's activated on the wrong server;"
            $ShouldBeMounted | 
                Foreach-object {$OUT_DETAIL += "$($_.DataBaseName) is Active on $($_.ActiveDatabaseCopy)`n"}

        }#If ($ShouldBeMounted)

        $databases | 
            foreach-object {
            $DBState = ""
            if (($_.Status -eq 'Healthy' -or $_.Status -eq 'Mounted') -and $_.ContentIndexState -eq 'Healthy') {
                $numHealthyDatabase++
            }#if (($_.Status -eq 'Healthy' -or $_.Status -eq 'Mounted') -and $_.ContentIndexState -eq 'Healthy')
            Else {
                [String]$DBState
                If ($_.ContentIndexState -ne 'Healthy') {$DBState += 'Content Index not Healthy,'}
                if ($_.Status -ne 'Healthy' -and $_.Status -ne 'Mounted') {$DBState += 'DB not Healthy'}
                $DBState = $DBState.trim(",")
                $tmp += $_.databasename + " $DBState`n"
            }#Else
        }#foreach-object 
    }#if ($databases) 
    Write-verbose ((get-date).toString() + " finishing Get-MailboxDatabaseCopyStatus check")
    $OUT_DETAIL += "$numHealthyDatabase Healthy databases`n"
    if ($tmp -ne "") {
        $OUT += "databases not Healthy; "
        $OUT_DETAIL += $tmp
        $EC = $CRITICAL
    }#if ($tmp -ne "")
    else {
        $TESTED += "Databases Healthy; "
    }#Else
    Write-verbose ((get-date).toString() + " starting Test-Mailflow check")
    #Get active to use for mailflow testing
    $ActiveDBs = $databases |
        Where {$_.ActivationPreference -eq 1 -and $_.ActiveDatabaseCopy -eq $server}

    if ($numHealthyDatabase -gt 0 -and $ActiveDBs) {
        Try {
            #Due to errors with Exchange 2013 the Test-Mailflow needs to be run in a session
            #http://exchangeserverpro.com/exchange-2013-test-mailflow-error-for-remote-mailbox-servers/
            $MailFlow = Invoke-Command -Session $CAS {Test-Mailflow} -ErrorAction STOP
        }#Try
        Catch {
            $OUT += "Test-Mailflow cmdlet failed`n" + $error[0]
            $EC = $WARNING
            nagios_exit
        }#Catch
        if ($MailFlow.TestMailflowResult -ne "Success") {
            $OUT += "TestMailflowResult = " + $MailFlow.TestMailflowResult + ", latency = " + $MailFlow.MessageLatencyTime + "; "
            $EC = $CRITICAL
        }#if ($MailFlow.TestMailflowResult -ne "Success")
        else {
            $TESTED += "Mailflow; "
        }#Else
    }#if ($numHealthyDatabase -gt 0 -and $ActiveDBs)
    else {
        $OUT_DETAIL += "No active databases present to test mail-flow`n"
    }#Else

    Write-verbose ((get-date).toString() + " finishing Test-Mailflow check")

}#if ($ExServer.IsMailboxServer -eq $true)
#endregion

#region display report
if ($EC -eq $OK) {
    $OUT = $TESTED
}#if ($EC -eq $OK)

$OUT = $OUT.trimend("; ")

if ($OUT_DETAIL -ne "") {
    $OUT += "`n$OUT_DETAIL"
}#if ($OUT_DETAIL -ne "") 

nagios_exit
#endregion
