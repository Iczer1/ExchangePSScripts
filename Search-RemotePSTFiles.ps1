<# 
.SYNOPSIS 
    Given one or more computer names the Search-RemotePSTFiles.ps will return a list of any PST founds on the remote computer and include the following information: Username of the PST owner, Outlook Profile associated with the PST, Outlook version, the PSTName, and the PST Path 
.DESCRIPTION 
    Will attempt to connect to each computer via a CIM Session by WSMAN, then DCOM. IF that fails the computer is added to the error list and the script moves onto the next computer. Once connected a list of user profiles will be pulled from the computer and then a search of each profiles registry will be performed in order to see if any Outlook profiles exist that contain a PST file (this portion can take up to 5 minutes for profiles with a lot of information) and if so it will pull the PST Name and file path. Once all PST information is collected it will be returned to the console or via CSV and/or email 
.PARAMETER ComputerName 
    One or more computer names to search for PST files 
.PARAMETER CSVExportDirectory 
    The path where a csv file for the PST list and error list should be dropped if needed 
.PARAMETER EmailTo 
    Will send an email report of all the changes to the email address specified. 
.PARAMETER EmailFrom 
    Used to specify the sending address for the email report. 
.PARAMETER SMTPServer 
    The IP address or DNS name of the email server you want to use in order to send an HTML email report of the results. 
.EXAMPLE 
    PS C:\> .\Search-RemotePSTFiles.ps1 -ComputerName "Spongebob100b","sideshowbob100b" 
 
    ComputerName   : Spongebob100b 
    UserName       : mello 
    ProfileName    : Default 
    OutlookVersion : 2013+ 
    PSTName        : SharePoint Lists 
    PSTPath        : C:\Users\mello\AppData\Local\Microsoft\Outlook\SharePoint Lists - Default(2).pst 
 
    ComputerName   : Spongebob100b 
    UserName       : mello 
    ProfileName    : Default 
    OutlookVersion : 2013+ 
    PSTName        : Archive 
    PSTPath        : C:\Users\mello\Documents\Archive.pst     
 
    Description 
    ----------- 
    Searches the computers "Spongebob100b","sideshowbob100b" for all PST files from all user profiles and reports back to the console 
.EXAMPLE 
    PS C:\> .\Search-RemotePSTFiles.ps1 –ComputerName "Spongebob100b","sideshowbob100b" -EmailTo joh.mello@home.com –EmailFrom "PSTSearch@Home.com" –SMTPServer "mail.home.com" -CSVExportDirectory "C:\Users\mello\Documents" 
     
    Description 
    ----------- 
    Searches the computers "Spongebob100b","sideshowbob100b" for all PST files from all user profiles and sends an email report to the specified user, from the specified address, using the specified SMTP server. Also drops a CSV version of the report along with an error report if one is generated 
.NOTES 
    Author: John Mello 
    Date: July 24, 2015 
    Reason for creation: To hunt down PSTs on the network 
.LINK 
    Mellositmusings.com 
    PhillyPosh.org 
#>  
 
#Requires -Version 3.0 
 
[CmdletBinding(DefaultParametersetName="Standard")] 
Param( 
    [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName,ParameterSetName = "Standard")] 
    [Parameter(Mandatory,ParameterSetName = "Email")] 
    [ValidateNotNullOrEmpty()] 
    [String[]]$ComputerName, 
 
    [Parameter(ParameterSetName = "Standard")] 
    [Parameter(ParameterSetName = "Email")] 
    [String]$CSVExportDirectory,  
 
    [Parameter(Mandatory,ParameterSetName = "Email")] 
       [string[]]$EmailTo, 
     
    [Parameter(Mandatory,ParameterSetName = "Email")]  
    [string]$EmailFrom,  
     
    [Parameter(Mandatory,ParameterSetName = "Email")]  
    [string]$SMTPServer 
) 
 
#BEGIN { 
    #region Functions 
 
    Function Search-Registry{ 
        <# 
        .SYNOPSIS 
        Searches the registry on one or more computers for a specified text pattern. 
 
        .DESCRIPTION 
        Searches the registry on one or more computers for a specified text pattern. Supports searching for any combination of key names, value names, and/or value data. The text pattern is a case-insensitive regular expression. 
 
        .PARAMETER StartKey 
        Starts searching at the specified key. The key name uses the following format: 
        subtree[:][\[keyname[\keyname...]]] 
        subtree can be any of the following: 
          HKCR or HKEY_CLASSES_ROOT 
          HKCU or HKEY_CURRENT_USER 
          HKLM or HKEY_LOCAL_MACHINE 
          HKU or HKEY_USERS 
        This parameter's format is compatible with PowerShell registry drive (e.g., HKLM:\SOFTWARE), reg.exe (e.g., HKLM\SOFTWARE), and regedit.exe (e.g., HKEY_LOCAL_MACHINE\SOFTWARE). 
 
        .PARAMETER Pattern 
        Searches for the specified regular expression pattern. The pattern is not case-sensitive. See help topic about_Regular_Expressions for more information. 
 
        .PARAMETER MatchKey 
        Matches registry key names. You must specify at least one of -MatchKey, -MatchValue, or -MatchData. 
 
        .PARAMETER MatchValue 
        Matches registry value names. You must specify at least one of -MatchKey, -MatchValue, or -MatchData. 
 
        .PARAMETER MatchData 
        Matches registry value data. You must specify at least one of -MatchKey, -MatchValue, or -MatchData. 
 
        .PARAMETER MaximumMatches 
        Specifies the maximum number of results per computer searched. 0 means "return the maximum number of possible matches." The default is 0. This parameter is useful when searching the registry on remote computers in order to minimize unnecessary network traffic. 
 
        .PARAMETER ComputerName 
        Searches the registry on the specified computer. This parameter supports piped input. 
 
        .OUTPUTS 
        PSObjects with the following properties: 
          ComputerName  The computer name on which the match occurred 
          Key           The key name (e.g., HKLM:\SOFTWARE) 
          Value         The registry value (empty for the default value) 
          Data          The registry value's data 
 
        .EXAMPLE 
        PS C:\> Search-Registry -StartKey HKLM -Pattern $ENV:USERNAME -MatchData 
        Searches HKEY_LOCAL_MACHINE (i.e., HKLM) on the current computer for registry values whose data contains the current user's name. 
 
        .EXAMPLE 
        PS C:\> Search-Registry -StartKey HKLM:\SOFTWARE\Classes\Installer -Pattern LastUsedSource -MatchValue | Select-Object Key,Value,Data | Format-List 
        Outputs the LastUsedSource registry entries on the current computer. 
 
        .EXAMPLE 
        PS C:\> Search-Registry -StartKey HKCR\.odt -Pattern .* -MatchKey -MaximumMatches 1 
        Outputs at least one match if the specified reistry key exists. This command returns a result if the current computer has a program registered to open files with the .odt extension. The pattern .* means 0 or more of any character (i.e., match everything). 
 
        .EXAMPLE 
        PS C:\> Get-Content Computers.txt | Search-Registry -StartKey "HKLM:\SOFTWARE\My Application\Installed" -Pattern "Installation Complete" -MatchValue -MaximumMatches 1 | Export-CSV C:\Reports\MyReport.csv -NoTypeInformation 
        Searches for the specified value name pattern in the registry on each computer listed in the file Computers.txt starting at the specified subkey. Output is sent to the specifed CSV file. 
         
        .NOTES 
        Written by Bill Stewart (bstewart@iname.com) 
        Pulled from : http://windowsitpro.com/scripting/searching-registry-powershell 
 
        #> 
 
        [CmdletBinding()] 
        param( 
          [parameter(Position=0,Mandatory=$TRUE)] 
            [String] $StartKey, 
          [parameter(Position=1,Mandatory=$TRUE)] 
            [String] $Pattern, 
            [Switch] $MatchKey, 
            [Switch] $MatchValue, 
            [Switch] $MatchData, 
            [UInt32] $MaximumMatches=0, 
          [parameter(ValueFromPipeline=$TRUE)] 
            [String[]] $ComputerName=$ENV:COMPUTERNAME 
        ) 
 
        begin { 
          $PIPELINEINPUT = (-not $PSBOUNDPARAMETERS.ContainsKey("ComputerName")) -and 
            (-not $ComputerName) 
 
          # Throw an error if -Pattern is not valid 
          try { 
            "" -match $Pattern | out-null 
          } 
          catch [System.Management.Automation.RuntimeException] { 
            throw "-Pattern parameter not valid - $($_.Exception.Message)" 
          } 
 
          # You must specify at least one matching criteria 
          if (-not ($MatchKey -or $MatchValue -or $MatchData)) { 
            throw "You must specify at least one of: -MatchKey -MatchValue -MatchData" 
          } 
 
          # Interpret zero as "maximum possible number of matches" 
          if ($MaximumMatches -eq 0) { $MaximumMatches = [UInt32]::MaxValue } 
 
          # These two hash tables speed up lookup of key names and hive types 
          $HiveNameToHive = @{ 
            "HKCR"               = [Microsoft.Win32.RegistryHive] "ClassesRoot"; 
            "HKEY_CLASSES_ROOT"  = [Microsoft.Win32.RegistryHive] "ClassesRoot"; 
            "HKCU"               = [Microsoft.Win32.RegistryHive] "CurrentUser"; 
            "HKEY_CURRENT_USER"  = [Microsoft.Win32.RegistryHive] "CurrentUser"; 
            "HKLM"               = [Microsoft.Win32.RegistryHive] "LocalMachine"; 
            "HKEY_LOCAL_MACHINE" = [Microsoft.Win32.RegistryHive] "LocalMachine"; 
            "HKU"                = [Microsoft.Win32.RegistryHive] "Users"; 
            "HKEY_USERS"         = [Microsoft.Win32.RegistryHive] "Users"; 
          } 
          $HiveToHiveName = @{ 
            [Microsoft.Win32.RegistryHive] "ClassesRoot"  = "HKCR"; 
            [Microsoft.Win32.RegistryHive] "CurrentUser"  = "HKCU"; 
            [Microsoft.Win32.RegistryHive] "LocalMachine" = "HKLM"; 
            [Microsoft.Win32.RegistryHive] "Users"        = "HKU"; 
          } 
 
          # Search for 'hive:\startkey'; ':' and starting key optional 
          $StartKey | select-string "([^:\\]+):?\\?(.+)?" | foreach-object { 
            $HiveName = $_.Matches[0].Groups[1].Value 
            $StartPath = $_.Matches[0].Groups[2].Value 
          } 
 
          if (-not $HiveNameToHive.ContainsKey($HiveName)) { 
            throw "Invalid registry path" 
          } else { 
            $Hive = $HiveNameToHive[$HiveName] 
            $HiveName = $HiveToHiveName[$Hive] 
          } 
 
          # Recursive function that searches the registry 
          function search-registrykey($computerName, $rootKey, $keyPath, [Ref] $matchCount) { 
            # Write error and return if unable to open the key path as read-only 
            try { 
              $subKey = $rootKey.OpenSubKey($keyPath, $FALSE) 
            } 
            catch [System.Management.Automation.MethodInvocationException] { 
              $message = $_.Exception.Message 
              write-error "$message - $HiveName\$keyPath" 
              return 
            } 
 
            # Write error and return if the key doesn't exist 
            if (-not $subKey) { 
              write-error "Key does not exist: $HiveName\$keyPath" -category ObjectNotFound 
              return 
            } 
 
            # Search for value and/or data; -MatchValue also returns the data 
            if ($MatchValue -or $MatchData) { 
              if ($matchCount.Value -lt $MaximumMatches) { 
                foreach ($valueName in $subKey.GetValueNames()) { 
                  $valueData = $subKey.GetValue($valueName) 
                  if (($MatchValue -and ($valueName -match $Pattern)) -or ($MatchData -and ($valueData -match $Pattern))) { 
                    "" | select-object ` 
                      @{N="ComputerName"; E={$computerName}}, 
                      @{N="Key"; E={"$HiveName\$keyPath"}}, 
                      @{N="Value"; E={$valueName}}, 
                      @{N="Data"; E={$valueData}} 
                    $matchCount.Value++ 
                  } 
                  if ($matchCount.Value -eq $MaximumMatches) { break } 
                } 
              } 
            } 
 
            # Iterate and recurse through subkeys; if -MatchKey requested, output 
            # objects only report computer and key (keys do not have values or data) 
            if ($matchCount.Value -lt $MaximumMatches) { 
              foreach ($keyName in $subKey.GetSubKeyNames()) { 
                if ($keyPath -eq "") { 
                  $subkeyPath = $keyName 
                } else { 
                  $subkeyPath = $keyPath + "\" + $keyName 
                } 
                if ($MatchKey -and ($keyName -match $Pattern)) { 
                  "" | select-object ` 
                    @{N="ComputerName"; E={$computerName}}, 
                    @{N="Key"; E={"$HiveName\$subkeyPath"}}, 
                    @{N="Value"; E={}}, 
                    @{N="Data"; E={}} 
                  $matchCount.Value++ 
                } 
                # $matchCount is a reference 
                search-registrykey $computerName $rootKey $subkeyPath $matchCount 
                if ($matchCount.Value -eq $MaximumMatches) { break } 
              } 
            } 
 
            # Close opened subkey 
            $subKey.Close() 
          } 
 
          # Core function opens the registry on a computer and initiates searching 
          function search-registry2($computerName) { 
          # Write error and return if unable to open the key on the computer 
          try { 
              $rootKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, 
                $computerName) 
            } 
            catch [System.Management.Automation.MethodInvocationException] { 
              $message = $_.Exception.Message 
              write-error "$message - $computerName" 
              return 
            } 
            # $matchCount is per computer; pass to recursive function as reference 
            $matchCount = 0 
            search-registrykey $computerName $rootKey $StartPath ([Ref] $matchCount) 
            $rootKey.Close() 
          } 
        } 
 
        process { 
          if ($PIPELINEINPUT) { 
            search-registry2 $_ 
          } 
          else { 
            $ComputerName | foreach-object { 
              search-registry2 $_ 
            } 
          } 
        } 
    } 
 
    function New-MrCimSession { 
    <# 
    .SYNOPSIS 
        Creates CimSessions to remote computer(s), automatically determining if the WSMAN 
        or Dcom protocol should be used. 
    .DESCRIPTION 
        New-MrCimSession is a function that is designed to create CimSessions to one or more 
        computers, automatically determining if the default WSMAN protocol or the backwards 
        compatible Dcom protocol should be used. PowerShell version 3 is required on the 
        computer that this function is being run on, but PowerShell does not need to be 
        installed at all on the remote computer. 
    .PARAMETER ComputerName 
        The name of the remote computer(s). This parameter accepts pipeline input. The local 
        computer is the default. 
    .PARAMETER Credential 
        Specifies a user account that has permission to perform this action. The default is 
        the current user. 
    .EXAMPLE 
         New-MrCimSession -ComputerName Server01, Server02 
    .EXAMPLE 
         New-MrCimSession -ComputerName Server01, Server02 -Credential (Get-Credential) 
    .EXAMPLE 
         Get-Content -Path C:\Servers.txt | New-MrCimSession 
    .INPUTS 
        String 
    .OUTPUTS 
        Microsoft.Management.Infrastructure.CimSession 
    .NOTES 
        Author:  Mike F Robbins 
        Website: http://mikefrobbins.com 
        Twitter: @mikefrobbins 
            Updates: 
            06/24/2015 : John Mello 
                Added a throw for the error conditions so that it works with a try/catch 
    #> 
        [CmdletBinding()] 
        param( 
            [Parameter(ValueFromPipeline)] 
            [ValidateNotNullorEmpty()] 
            [string[]]$ComputerName = $env:COMPUTERNAME, 
  
            [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty 
        ) 
 
        BEGIN { 
            $Opt = New-CimSessionOption -Protocol Dcom 
 
            $SessionParams = @{ 
                ErrorAction = 'Stop' 
            } 
 
            If ($PSBoundParameters['Credential']) { 
                $SessionParams.Credential = $Credential 
            } 
        } 
 
        PROCESS { 
            foreach ($Computer in $ComputerName) { 
                $SessionParams.ComputerName  = $Computer 
 
                if ((Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue).productversion -match 'Stack: ([3-9]|[1-9][0-9]+)\.[0-9]+') { 
                    try { 
                        Write-Verbose -Message "Attempting to connect to $Computer using the WSMAN protocol." 
                        New-CimSession @SessionParams 
                    } 
                    catch { 
                        Write-Warning -Message "Unable to connect to $Computer using the WSMAN protocol. Verify your credentials and try again." 
                        Throw 
                    } 
                } 
  
                else { 
                    $SessionParams.SessionOption = $Opt 
 
                    try { 
                        Write-Verbose -Message "Attempting to connect to $Computer using the DCOM protocol." 
                        New-CimSession @SessionParams 
                    } 
                    catch { 
                        Write-Warning -Message "Unable to connect to $Computer using the WSMAN or DCOM protocol. Verify $Computer is online and try again." 
                        Throw 
                    } 
 
                    $SessionParams.Remove('SessionOption') 
                }             
            } 
        } 
    } 
 
    function Set-AlternatingCSSClasses { 
        <# 
            .SYNOPSIS 
                Takes an HTML fragment made by Convertto-Html and add the CSS clas names to the <TR> tags 
            .DESCRIPTION 
            .PARAMETER HTMLFragment 
                HTML fragment containg a table from ConvertTo-Html 
            .PARAMETER CSSEvenClass 
                The CSS class name for your even row color 
            .PARAMETER CSSOddClass 
                The CSS class name for your odd row color     
            .EXAMPLE 
                PS C:\> Get-Process ConvertTo-HTML -Fragment |Out-String | Set-AlternatingCSSClasses -CSSEvenClass 'even' -CssOddClass 'odd' 
             
                Description 
                -----------     
                Returns a string with the rows in the alternating colors specifed 
            .NOTE 
                Taken from the PowerShell.org book "HTML Reports in PowerShell" 
        #> 
        [CmdletBinding()] 
         
        param( 
            [Parameter(Mandatory=$True,ValueFromPipeline=$True)] 
            [string]$HTMLFragment, 
             
            [Parameter(Mandatory=$True)] 
            [string]$CSSEvenClass, 
             
            [Parameter(Mandatory=$True)] 
            [string]$CSSOddClass 
        ) 
         
        [xml]$xml = $HTMLFragment 
        $table = $xml.SelectSingleNode('table') 
        $classname = $CSSOddClass 
        foreach ($tr in $table.tr) { 
            if ($classname -eq $CSSEvenClass) { 
                    $classname = $CssOddClass 
            }  
            else {$classname = $CSSEvenClass} 
            $class = $xml.CreateAttribute('class') 
            $class.value = $classname 
            $tr.attributes.append($class) | Out-null 
        } 
        $xml.innerxml | out-string 
    } 
 
    #endregion 
 
    #region Variables 
     
    $RemotePSTList = @() 
    $RemotePSTListProps = @{ 
        ComputerName = ''; 
        UserName = ''; 
        ProfileName = ''; 
        OutlookVersion = ''; 
        PSTName =''; 
        PSTPath ='' 
    } 
 
    $ErrorTable = @{} 
    #endregion 
 
    #region CSSStyleSheet 
 
#This helps create alterntating color rows 
$style = @" 
 
body { 
    background-color:#F2F2F2; 
    color:#000000; 
    font-family:Calibri,Tahoma; 
    font-size: 10pt; 
} 
h1 { 
    text-align:center; 
} 
h2 { 
    border-top:1px solid #000000; 
} 
 
Table,td { 
    border-collapse:collapse; 
} 
 
th { 
    font-weight:bold; 
    color:#FFFFFF; 
    background-color:#045FB4; 
    border:1px solid #000000; 
    padding:2px; 
} 
 
td { 
    border:1px solid black; 
    padding:2px; 
} 
     
.odd  {  
        background-color:#ffffff; 
} 
.even {  
        background-color:#BDBDBD;  
} 
 
"@ 
 
    #endregion 
 
#}#Begin 
 
#PROCESS{                                                                                                                                                                                                                                                                                                                                                                             Foreach ($Computer in $ComuterName) { 
    Foreach ($Computer in $ComputerName) { 
        #Try to connection via DCOM or WSMAN to the remote computer 
        Try {$CIMSession = New-MrCimSession -ComputerName $Computer -ErrorAction Stop} 
        Catch { 
            $ErrorTable.ADD($Computer,$_.Exception.Message) 
            Continue 
        }#Catch 
 
        #Grab All none system account profiles 
        #$Profiles = Get-WmiObject -Class win32_userprofile -ComputerName $Computer -ErrorAction STOP |  
        $Profiles = Get-CimInstance -ClassName win32_userprofile -CimSession $CIMSession -ErrorAction stop | 
                    Where-object localpath -notlike "C:\Windows\S*" |  
                    Select sid,localpath,@{L="UserName";E={($_.Localpath).split("\")[-1]}} 
 
        If (-not $Profiles ) { 
            Write-verbose "$Computer does not have users accounts, skipping" 
            contiune 
        }#IF 
 
        #Clean up CIM session since it was only needed to grab the profile list 
        Remove-CimSession $CIMSession 
 
        #Need to figure out a way to speed this up 
        #Maybe test for the key first then skip? 
        #http://psremoteregistry.codeplex.com/??? 
        Foreach ($Profile in $Profiles) { 
            #TODO 
            #Why am i searchinf this way again and not just doing this for each profile? 
            #Get-ChildItem -Path 'HKCU:\' -Recurse | Where-Object -Property property -Like '001f6700' 
            Write-verbose "Searching the registry for Outlook profiles entries that reference PST files for : $($Profile.username)" 
            $AllPSTs = @() 
            #2013 = "\Software\Microsoft\Office", 2010 and below = "\Software\Microsoft\Windows NT\CurrentVersion" 
            $ProfilePaths = @(("HKU:\" + $Profile.Sid + "\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"),  
                                ("HKU:\" + $Profile.Sid + "\Software\Microsoft\Office\15.0\Outlook\Profiles")) 
            foreach ($Path in  $ProfilePaths) { 
                Try { 
                    $AllPSTs += Search-Registry -StartKey $Path -Pattern '001f6700' -MatchValue -MaximumMatches 0 -ComputerName $Computer -ErrorAction Stop 
                } 
                Catch { 
                    If ($_.Exception.message -like "Key does not exist*") { 
                        If ($Path -like "*\Office\*") {Write-Verbose "No keys for 2013+ profiles found for $($Profile.username)"} 
                        Else {Write-Verbose "No keys for 2010- profiles found for $($Profile.username)"} 
                    }#IF 
                    ElseIf ($_.Exception.message -like "*The network path was not found*") { 
                        Write-Warning "Can't connect to registry on $Computer" 
                        $ErrorTable.ADD("$Computer : $($Profile.UserName))",$_.Exception.Message) 
                        break 
                    }#ElseIf 
                    Else { 
                        Write-Warning "Unknown error" 
                        $_ 
                    }#Else 
                }#Catch 
            }#foreach ($Path in  $ProfilePaths) 
 
            #Start populating the static values for this user 
            $RemotePSTListProps = @{ 
                ComputerName = $Computer; 
                UserName = $Profile.UserName; 
                ProfileName = ''; 
                OutlookVersion = ''; 
                PSTName =''; 
                PSTPath ='' 
            }#$RemotePSTListProps 
 
            If (-not $AllPSTs) { 
                Write-verbose "No Outlook profiles with PST's found for $($Profile.UserName), skipping" 
                $RemotePSTListProps.ProfileName = 'N/A' 
                $RemotePSTListProps.OutlookVersion = 'N/A' 
                $RemotePSTListProps.PSTName = 'N/A' 
                $RemotePSTListProps.PSTPath = 'N/A' 
                $RemotePSTList += New-Object PSObject -Property $RemotePSTListProps 
            }#If (-not $AllPSTs) 
            Else{ 
                Write-verbose "Outlook profiles with PST's found for $($Profile.username) found, pulling details" 
                Foreach ($PST in $AllPSTs) { 
                    #Not error checking since at this point we should have a connection to the remote PC 
                    #Grab PST Name, there should only be 1 value 
                    $PSTName = (Search-Registry -StartKey $PST.key -Pattern '001f3001' -Matchvalue -MaximumMatches 1 -ComputerName $Computer).data | 
                                Where-object {$_ -ne 0} 
                    #Grab PST Path, there should only be 1 value 
                    $PSTFilePath = (Search-Registry -StartKey $PST.key -Pattern '001f6700' -Matchvalue -MaximumMatches 1 -ComputerName $Computer).data | 
                        Where-object {$_ -ne 0} 
                    #Fill in the remaining Values 
                    $RemotePSTListProps.ProfileName = $PST.key.split("\")[-2] 
                    If ($PST.Key -like "*\Office\1*") {$RemotePSTListProps.OutlookVersion = "2013+"} 
                    Else {$RemotePSTListProps.OutlookVersion = "2010-"} 
                    #Sometimes PST Names can be blank 
                    If ($PSTNAME -eq $NULL) { 
                        $RemotePSTListProps.PSTName = "DEFAULT NAME" 
                    }#IF 
                    Else { 
                        $RemotePSTListProps.PSTName = [System.Text.Encoding]::Default.GetString($PSTName) 
                    }#ELSE 
                    $RemotePSTListProps.PSTPath = [System.Text.Encoding]::Default.GetString($PSTFilePath) 
                    #Add to the full list 
                    Write-Verbose "PST named `'$($RemotePSTListProps.PSTName)`' located at `'$($RemotePSTListProps.PSTPath)`', logging" 
                    $RemotePSTList += New-Object PSObject -Property $RemotePSTListProps 
                }#Foreach ($PST in $AllPSTs) 
            }#ELSE 
        }#Foreach ($Profile in $Profiles) 
    }#Foreach ($Computer in $ComputerName) 
 
#}#PROCESS  
 
#END{ 
    if ($RemotePSTList -or $ErrorTable) { 
        If (-not $CSVExportDirectory -and -not $SMTPServer) { 
            Write-verbose "No email or CSV options specifed, dumping info to console" 
            $RemotePSTList |  
                Where PSTName -ne "N/A" |  
                Select-Object ComputerName, UserName,ProfileName, OutlookVersion, PSTName, PSTPAth | 
                Sort-object ComputerName 
            $ErrorTable.GetEnumerator() |  
                select @{L="ComputerName";E={$_.Name}},@{L="Error";E={$_.Value}} | 
                Sort-Object ComputerName 
        }#If (-not $CSVExportDirectory -and -not $SMTPServer) 
         
        If ($CSVExportDirectory) { 
            Write-Verbose "CSV report specified, creating CSV export at $CSVExportDirectory" 
            Try { 
                $RemotePSTList |  
                    Where PSTName -ne "N/A" |  
                    Select-Object ComputerName, UserName,ProfileName, OutlookVersion, PSTName, PSTPAth | 
                    Export-Csv -NoTypeInformation -Path ("$CSVExportDirectory\" + (Get-Date -format "yyy-MM-dd  hh_mm") + " - PST Report.csv") -ErrorAction stop 
                If ($ErrorTable) { 
                    $ErrorTable.GetEnumerator() |  
                        select @{L="ComputerName";E={$_.Name}},@{L="Error";E={$_.Value}} | 
                        Sort-Object ComputerName | 
                        Export-Csv -NoTypeInformation -Path ("$CSVExportDirectory\" + (Get-Date -format "yyy-MM-dd  hh_mm") + " - ErrorReport.csv") -ErrorAction stop 
                } 
            } 
            Catch { 
                Write-Error "Error writing to $CSVExportDirectory, dumping to console" 
                $_ 
                $RemotePSTList |  
                    Where PSTName -ne "N/A" |  
                    Select-Object ComputerName, UserName,ProfileName, OutlookVersion, PSTName, PSTPAth | 
                    Sort-object ComputerName 
                $ErrorTable.GetEnumerator() |  
                    select @{L="ComputerName";E={$_.Name}},@{L="Error";E={$_.Value}} | 
                    Sort-Object ComputerName 
            } 
        }#If ($CSVExportDirectory) 
 
        if ($SMTPServer) { 
            $Date = Get-Date 
 
            Write-Verbose "Email report specified, creating and sending HTML report" 
            $AllHTMLTables = @() 
 
            $HTMLPSTs = $RemotePSTList |  
                            Where PSTName -ne "N/A" |  
                            Select-Object ComputerName, UserName,ProfileName, OutlookVersion, PSTName, PSTPAth | 
                            Sort-object ComputerName |  
                            ConvertTo-HTML -Fragment | 
                            Out-String | 
                            Set-AlternatingCSSClasses -CSSEvenClass 'even' -CssOddClass 'odd' 
            $HTMLPSTs = "<H3>Found PST files</H3>$HTMLPSTs" 
            $AllHTMLTables += $HTMLPSTs 
 
            If ($ErrorTable.GetEnumerator()) { 
                $HTMLErrors = $ErrorTable.GetEnumerator() |  
                                select @{L="ComputerName";E={$_.Name}},@{L="Error";E={$_.Value}} | 
                                Sort-Object ComputerName | 
                                ConvertTo-HTML -Fragment | 
                                Out-String | 
                                Set-AlternatingCSSClasses -CSSEvenClass 'even' -CssOddClass 'odd' 
                $HTMLErrors = "<H3>PC's with errors</H3>$HTMLErrors" 
                $AllHTMLTables += $HTMLErrors 
            } 
                 
 
            $params = @{'Head'="<title>PST report</title>$style"; 
                        'PreContent'="<h1>PST report for the run time of $Date</h1>"; 
                        'PostContent'= $AllHTMLTables 
            } 
 
            $Subject = "PST report for $Date" 
            send-mailmessage -from $EmailFrom -to $EmailTo -smtpserver $SmtpServer -subject $Subject -Body (ConvertTo-HTML @params | Out-String)  -BodyAsHtml  
        } #if ($SMTPServer) 
    }#if ($RemotePSTList -or $ErrorTable) 
    Else { 
        Write-Output "No PST's found on the specified computers" 
    }#ELSE 
#}#END
