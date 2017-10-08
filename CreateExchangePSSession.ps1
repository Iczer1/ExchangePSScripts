Write-Verbose "INFO : Checking if Exchange cmdlets are loaded"
If ((Get-Command Get-Mailbox -ErrorAction SilentlyContinue -Verbose:$false) -ne $Null) {Write-Verbose "Exchange PSSnapin is loaded, Proceeding with script"}
Else {
       #Maybe in the future grab an Exchange server in the same AD site?
    #[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name
    #Create ADSI Object and set to root
       $root = [ADSI]'LDAP://RootDSE'
       $cfConfigRootpath = "LDAP://" + $root.ConfigurationNamingContext.tostring()
       $configRoot = [ADSI]$cfConfigRootpath
       #Create searcher and fliter
       $searcher = new-object System.DirectoryServices.DirectorySearcher($configRoot)
       #Mailbox=2, CAS=4, UM=16, HT=32, ET=64
       $searcher.Filter = '(&(&(objectCategory=msExchExchangeServer)(msExchCurrentServerRoles:1.2.840.113556.1.4.803:=4)))'
       #Perform the search
       [VOID]$searcher.PropertiesToLoad.Add("cn")
       $searchres = $searcher.FindAll()
       #$RandomCAS = ($searchres | Where-Object {$_.properties.cn -notlike "*CHI*"} | Get-Random).properties.cn
       $RandomCAS = ($searchres | Where-Object {$_.properties.cn -like "*XCHBAL50*" -or $_.properties.cn -like "*LABXCH*"} | Get-Random).properties.cn


       #Had to change to an an import session due to changes in Exchange SP3 that breaks certain cmdlets
       #http://support.microsoft.com/kb/2859999
       Try {
              $CAS = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$RandomCAS.ds.susq.com/PowerShell/" -Authentication Kerberos -Name EXCH -ErrorAction Stop
              Import-PSSession  $CAS -allowclobber -ErrorAction Stop -DisableNameChecking -WarningAction SilentlyContinue | Out-Null
       }
       Catch {
              Write-Warning "Exchange Cmdlets cannot be loaded, exiting script"
              Get-PSSession -Name EXCH -ErrorAction SilentlyContinue | Remove-PSSession
              Exit 1
       }
}
#endregion 
