<#
.PARAMETER EmailSubject
    The exact subject of the email that the script will search for, leave blank in order to leave the email subject out of the search
.PARAMETER SavePath
    The local or network path the attachments found will be save to
.PARAMETER PublicFolderPath
    Path to a public folder in the form of "\Rootlevelfolder\Subfolder\subfolder"
.PARAMETER From
    The from email addresse of the email you want to search for
.PARAMETER ExportEmailBody
    If specifed then then email body is also dumped to a text file with the naming format YYYY_MM_DD-hh_mm-SUBJECT.txt
.PARAMETER NoClobber
    Checks if a file with the same name exists in the directory, if so then it will prepend a random 10 string to the file name
.PARAMETER SkipWriteTest
    Will skip powershell's attempt to write and delete a test file to each directory it was presented
.PARAMETER SMTPAddress
    The SMTP address of the Exchange mailbox that you want to connect to.
    If left blank then the script will attempt to connect to the mailbox of the runtime account by searching LDAP for it's SMTP address
.PARAMETER DeleteEmail
    If specified any email found and it’s attachments successfully saved will be moved to the deleted folder. 
.PARAMETER TodayOnly
    If specified only emails found that match the Day, Month, and Year of the current script runtime will be returned by the inbox search
    If not specifed the script will leave the date of the email out of the search
.PARAMETER ReceivedInThePastNumberOfDays
    Will search for emails in the date range of the script runtime and X nuber of days prior
    If not specifed the script will leave the date of the email out of the search
.PARAMETER ExactSubjectSearch
    By default the script will search for any email matching any of the terms in the subject being searched for.
    when using this parameter the script will look for emails with a subject that exactly matches the subject being searched for
    For example if you are looking for "Expense Report" this will not return results for "RE: Expense Report" or "FW: Expense Report"
.PARAMETER UnZipFiles
    If any attachments found are zip files you can specify that the script extract the files within the zip file and store them at the specified location instead of the zip file itself
    NOTE, this will delete the zip file after it's contents have been unzipped
.PARAMETER EwsUrl
    Instead of using auto discover you can specify the EWS URL
.PARAMETER ExchangeVersion
    The EWS feature set to use when connecting to the Exchange instace, If not speficed then the API will default to the highest supported version.
    Note that you can't use a higher version then what is supported by the Exchange instance, but you can use lower versions.
.PARAMETER IgnoreSSLCertificate
    Use this option to ignore any certificate errors. This is helpful if the Exchange instance uses a self signed cert
.PARAMETER EWSManagedApiPath
    Used to specify the path for the EWS Managed API, if blank the script will search your system for a copy of the newest Microsoft.Exchange.WebServices.dll via a registry search
.PARAMETER EWSTracing
    Used to turn on EWS tracing for troubleshooting
.PARAMETER ImpersonationCredential
    Supplied credential (using Get-Credential) for impersonation
.PARAMETER EmailErrorsFrom
    The SMTP address error emails will come from
.PARAMETER EmailErrorsTo
    The SMTP address error emails will go to
.PARAMETER SMTPServer
    The SMTP address server used to send error emails
.PARAMETER ForceSearchFilter
    Forces the old searching methond instead of AQS, even on Exchange versions that support it. Sometimes AQS times out
.EXAMPLE
    PS C:\> .\Get-AttachmentFromEmail.ps1 -EmailSubject “Important Attachment” –SavePath “\\FileServer\Email Attachments” -SMTPAddress John.Doe@company.com
    
    Description
    -----------
    Will use the runtime account to access and search the mailbox associated with "John.Doe@company.com" for any emails in the inbox with a subject like “Important Attachment” and save the attachments to \\FileServer\Email Attachments
.EXAMPLE
    PS C:\> .\Get-AttachmentFromEmail.ps1 -EmailSubject “Important Attachment” –SavePath “\\FileServer\Email Attachments” -SMTPAddress John.Doe@company.com -NoClobber
    
    Description
    -----------
    Will use the runtime account to access and search the mailbox associated with "John.Doe@company.com" for any emails in the inbox with a subject like “Important Attachment” and save the attachments to \\FileServer\Email Attachments.
    If a file with the same name already exists in the directory then a random 10 digit number will be prepended to the file and a 2nd check of the directory will be made before saving the file
    This process will repeat until a unique file name is created
.EXAMPLE
    PS C:\> .\Get-AttachmentFromEmail.ps1 -EmailSubject “Important Attachment” –SavePath “\\FileServer\Email Attachments” -SMTPAddress John.Doe@company.com -ExactSubjectSearch –UnZipFiles 
    
    Description
    -----------
    Will use the runtime account to access and search the mailbox associated with "John.Doe@company.com" for any emails in the inbox with the exact subject of “Important Attachment” and save the attachments to \\FileServer\Email Attachments. If any of those attachments are zip files then the script will extract the contents of the zip files leaving only the contents and not the zip file
.EXAMPLE
    PS C:\> .\Get-AttachmentFromEmail.ps1 -EmailSubject “Sales Results” –SavePath “C:\Email Attachments” -SMTPAddress John.Doe@company.com -DeleteEmail -TodayOnly 
    
    Description
    -----------
    Will use the runtime account to access and search the mailbox associated with "John.Doe@company.com" for any emails in the inbox with a subject like “Sales Results” with the runtime date of today and save the attachments to \\FileServer\Email Attachments. Once the attachments have been saved then move the email to the delete items folder
.EXAMPLE
    PS C:\> .\Get-AttachmentFromEmail.ps1 -EmailSubject “Important Attachment” –SavePath \\FileServer\Email Attachments -SMTPAddress differentUser@company.com –ImpersonationUserName ServiceMailboxAutomation –ImpersonationPassword “$*fnh23587” –ImpersonationDomain CORP
    
    Description
    -----------
    Will search the mailbox associated with “differentUser@company.com” using the account “ServiceMailboxAutomation” which will need the impersonation RBAC role to access the mailbox. The search will be for any emails in the inbox with a subject like “Important Attachment” and save the attachments to \\FileServer\Email Attachments. 
.EXAMPLE
    PS C:\> .\Get-AttachmentFromEmail.ps1 -EmailSubject “Important Attachment” –SavePath \\FileServer\Email Attachments -SMTPAddress John.Doe@company.com –EmailErrorsFrom EmailTeam@company.com –EmailErrorsTo HelpDesk@company.com –SMTPServer mail.company.com
    
    Description
    -----------
    Will use the runtime account to access and search the mailbox associated with "John.Doe@company.com" for any emails in the inbox with a subject like “Important Attachment” and save the attachments to \\FileServer\Email Attachments. If any errors are encountered then send an email from “EmailTeam@company.com” to “HelpDesk@company.com” using the SMTP server of “mail.company.com”
.EXAMPLE
    PS C:\> .\Get-AttachmentFromEmail.ps1 -EmailSubject “Important Attachment” -From "Jane.Doe@Company.com","HR@Company.com" –SavePath \\FileServer\Email Attachments -SMTPAddress John.Doe@company.com
    
    Description
    -----------
    Will use the runtime account to access and search the mailbox associated with "John.Doe@company.com" for any emails in the inbox with a subject like “Important Attachment” and sent from either "Jane.Doe@Company.com","HR@Company.com" and save the attachments to \\FileServer\Email Attachments.
.EXAMPLE
    PS C:\> .\Get-AttachmentFromEmail.ps1 -EmailSubject “Important Attachment” –SavePath \\FileServer\Email Attachments -SMTPAddress John.Doe@company.com –EwsUrl https://webmail.company.com/EWS/Exchange.asmx –EWSManagedApiPath "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll" -$ExchangeVersion Exchange2010_SP1 –IgnoreSSLCertificate -EWSTracing
    
    Description
    -----------
    Will use the runtime account to access and search the mailbox associated with "John.Doe@company.com" any emails in the inbox that exactly match “Important Attachment” and save the attachments to \\FileServer\Email Attachments. Before connecting to the mailbox the EWS connection will use the following options
        1.  The EWS URL of https://webmail.company.com/EWS/Exchange.asmx
        2.  The EWS Managed API Path of "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"
        3.  The Exchange version of "Exchange2010_SP1"
        4.  Ignore all SSL Certificate errors
        5.  Turn on EWS tracing
.INPUTS
    System.String
        You cannot pipe to this script
.OUTPUTS
    NONE
.NOTES
    AUTHOR: John Mello 
    CREATED : 10/29/2013 
    CREATED BECAUSE: To automate the retrieval, download, and deletion of a daily delivered email attachment
    Updates
        12/17/2014 John Mello
            Added a parameter NoClobber to prepend a random 10 digit string to the file name before saving
            Added a parameter From that takes one or more email addresses to search for
            Added a function to do a test wrtie to the specifed destinations instead of just seeing if the path is accessible
            Added a parameter FindMyEWSManagedApi which will search the local system for the EWS Managed API DLL file
        09/15/2015 John Mello
            Noticed that when working with 2013 Public folders my pre search check that verifes that a folder has items doesn't work with 2013 public folders, they always return 0 items. So I skip the check for 2013 publics folders for the time being
            MovedToDeletedItems doesn't work with 2013 public folders, had to have it do a harddelete
            Change EWS default version to Exchange2013
        12/21/2015
            Added ReceivedInThePastNumberOfDays parameter at the request of users looking for a date range
            Corrected an issue that led to some issues with the search results not being properly counted in PS v2
        06/15/2016
            Simplified impersonation option with the use of get credential
            Simplified the function to create the EWS connection
            Added proxyaddresses search to LDAP email search as a fall back
            Added an option to force SearchFilter searches due to issues with AQS in my enviroment
        08/09/2016
            For some reason I wasn't getting the scoping right when passing the script parameters to the function that creates the EWS object, now i create a splat to pass
            Had to force the non impersonation parameter set. Similar to the scoping issue, I couldn't get the parameter sets to shake out
        08/10/2016
            Removed option for multiple FROM addresses to be included in the search query
            fixed an issue where the from was specifed as BLANK if left not unspecified, resulting in zero results since there was a search for a blank from address
        11/29/2017
            Added some extra vebose statements for the searching and file portion
.LINK
   DSMS Script Storage location : \\Bala01\World\Technology\EnterpriseServices\PlatformServices\Directory and Messaging Services\Scripts
   DSMS SharePoint Scripting Portal : http://techweb/ess/dci/dsms/scripting/default.aspx
   DSMS PowerShell KB : http://techweb/ess/dci/dsms/scripting/powershellscripting/Home.aspx
#>
[CmdletBinding(DefaultParameterSetName = "DEFAULT")]
Param(
    [Parameter(Mandatory = $False, Position = 0)]
    [string]$EmailSubject,
        
    [Parameter(Mandatory = $True, Position = 1)]
    [string]$SavePath = "P:\",
    
    [Parameter(Mandatory = $False)]
    [String]$PublicFolderPath,  

    [Parameter(Mandatory = $False)]
    [string]$From,

    [Parameter(Mandatory = $False)]
    [switch]$ExportEmailBody,
    
    [Parameter(Mandatory = $False)]
    [switch]$NoClobber,
    
    [Parameter(Mandatory = $False)]
    [switch]$SkipWriteTest,
    
    [Parameter(Mandatory = $False, ParameterSetName = "DEFAULT", Position = 1)]
    #[Parameter(Mandatory=$False,Position=1)]
    [Parameter(Mandatory = $True, ParameterSetName = "Impersonation", Position = 0)]
    [string]$SMTPAddress,
    
    [Parameter(Mandatory = $false)]
    [switch]$DeleteEmail,
    
    [Parameter(Mandatory = $false)]
    [switch]$TodayOnly,
    
    [Parameter(Mandatory = $false)]
    [int]$ReceivedInThePastNumberOfDays,

    [Parameter(Mandatory = $false)]
    [switch]$ExactSubjectSearch,
    
    [Parameter(Mandatory = $false)]
    [switch]$UnZipFiles,
    
    [Parameter(Mandatory = $false)]
    [string]$EwsUrl,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("Exchange2007", "Exchange2007_SP1", "Exchange2007_SP2", "Exchange2007_SP3", "Exchange2010", "Exchange2010_SP1", "Exchange2010_SP2", "Exchange2013", "Exchange2013_SP1")]
    [string]$ExchangeVersion = "Exchange2013_SP1",
    
    [Parameter(Mandatory = $false)]
    [switch]$IgnoreSSLCertificate,
    
    [Parameter(Mandatory = $false)]
    [string]$EWSManagedApiPath,

    [Parameter(Mandatory = $false)]
    [switch]$EWSTracing,
    
    [Parameter(Mandatory = $True, ParameterSetName = "Impersonation", Position = 1)]
    [System.Management.Automation.PSCredential]$ImpersonationCredential,
    
    [Parameter(Mandatory = $False)] 
    [string]$EmailErrorsFrom, 
    
    [Parameter(Mandatory = $False)] 
    [string[]]$EmailErrorsTo,

    [Parameter(Mandatory = $False)] 
    [string]$SMTPServer,

    [Parameter(Mandatory = $False)] 
    [switch]$ForceSearchFilter
)

#region Functions

Function New-EWSServiceObject {
    <#
    .SYNOPSIS
        Returns an EWS Service Object
    .DESCRIPTION
        Creates a EWS service object, with the option of using impersonation and/or An EWS URL or fall back to Autodiscover
    .PARAMETER SMTPAddress
        The SMTP address of the Exchange mailbox that you want to connect to.
        If left blank then the script will attempt to connect to the mailbox of the runtime account by searching LDAP for it's SMTP address 
    .PARAMETER EwsUrl
        The EWS URL for the Exchange instance you want to connect to, leave this blank to leverage AutoDiscover in Exchange to obtain the EWS URL
    .PARAMETER ExchangeVersion
        The EWS feature set to use when connecting to the Exchange instace, If not speficed then the API will default to the highest supported version.
        Note that you can't use a higher version then what is supported by the Exchange instance, but you can use lower versions.
    .PARAMETER IgnoreSSLCertificate
        Use this option to ignore any certificate errors. This is helpful if the Exchange instance uses a self signed cert
    .PARAMETER EWSManagedApiPath
        Used to specify the path for the EWS Managed API, if blank the script will search your system for a copy of the newest Microsoft.Exchange.WebServices.dll via a registry search
    .PARAMETER EWSTracing
        Used to turn on EWS tracing for troubleshooting
    .PARAMETER ImpersonationCredential
        Supplied credential (using Get-Credential) for impersonation
    .EXAMPLE
        PS C:\powershell> New-EWSServiceObject -SMTPAddress "John.Doe@Company.com"

        ------------------
        Description 
        ------------------
        
        This will return and EWS object that uses impersonation with a specifed URL
    #>
    [CmdletBinding(DefaultParameterSetName = "DEFAULT")]
    Param(
        [Parameter(Mandatory = $False, ParameterSetName = "DEFAULT", Position = 0)]
        #[Parameter(Mandatory=$False,Position=0)]
        [Parameter(Mandatory = $True, ParameterSetName = "Impersonation", Position = 0)]
        [string]$SMTPAddress,       

        [Parameter(Mandatory = $false)]
        [string]$EwsUrl,
    
        [Parameter(Mandatory = $false)]
        [ValidateSet("Exchange2007", "Exchange2007_SP1", "Exchange2007_SP2", "Exchange2007_SP3", "Exchange2010", "Exchange2010_SP1", "Exchange2010_SP2", "Exchange2013", "Exchange2013_SP1")]
        [string]$ExchangeVersion,
    
        [Parameter(Mandatory = $false)]
        [switch]$IgnoreSSLCertificate,
    
        [Parameter(Mandatory = $false)]
        [string]$EWSManagedApiPath,

        [Parameter(Mandatory = $false)]
        [switch]$EWSTracing,

        [Parameter(Mandatory = $True, ParameterSetName = "Impersonation", Position = 1)]
        [System.Management.Automation.PSCredential]$ImpersonationCredential
    )   

    # Check if EWS Managed API path was specifed, if not try to find it
    If (-not $EWSManagedApiPath) {
        Write-Verbose "No EWS path sepcifed, let me try to find it for you"
        $EWSPath = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
        If ($EWSPath) {
            Write-Verbose "Found the following EWS API Path : $EWSPath"
            $EWSManagedApiPath = $EWSPath
        } #If EWSPath
        Else {
            Write-Warning "Can't find the EWS DLL path"
            Write-Warning "Please verify that you have installed the Microsoft Exchange Web Services Managed API"
            Break
        }# Else
    } #If (-not $EWSManagedApiPath)

    #Verify Path for the EWS managed API is present 
    Try {
        Get-Item -Path $EWSManagedApiPath -ErrorAction Stop | Out-Null
    }#Try
    Catch {
        Write-Warning "EWS Managed API path cannot be accessed ($EWSManagedApiPath), please verify that the EWS managed API is installed"
        Break
    }#Catch

    Write-Verbose "Loading EWS Managed API"
    [void][Reflection.Assembly]::LoadFile($EWSManagedApiPath)

    Write-Verbose "Creating EWS Service connection object"
    If ($ExchangeVersion) {
        $EWSService = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ExchangeVersion)
    }#If ($ExchangeVersion)
    Else {
        #The API will default to the highest supported version by default
        $EWSService = new-object Microsoft.Exchange.WebServices.Data.ExchangeService
    }#Else

    If ($EWSTracing) {
        Write-Verbose "EWS Tracing enabled"
        $EWSService.traceenabled = $true
    }#If ($EWSTracing)
    
    if ($IgnoreSSLCertificate) {
        Write-Verbose "Ignoring any SSL certificate errors"
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true };
    }#if ($IgnoreSSLCertificate)    
    

    If (-not $SMTPAddress) {
        Write-Verbose "No SMTPAddress specifed, using LDAP to find runtime user's SMTPAddress"
        $strFilter = "(&(objectCategory=User)(name=$env:USERNAME))"

        $objDomain = New-Object System.DirectoryServices.DirectoryEntry

        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
        $objSearcher.SearchRoot = $objDomain
        $objSearcher.PageSize = 1
        $objSearcher.Filter = $strFilter
        $objSearcher.SearchScope = "Subtree"

        #Using Mail AD attribute
        $colProplist = "mail"
        foreach ($i in $colPropList) {$objSearcher.PropertiesToLoad.Add($i) | Out-Null} 

        $colResults = $objSearcher.FindAll()
        $SMTPAddress = $colResults[0].Properties.mail

        If (-not $SMTPAddress) {
            Write-verbose "Mail attrbute is blank, checking proxyaddresses attribute"
            #Using proxyAddresses AD attribute
            $colProplist = = "proxyAddresses"
            foreach ($i in $colPropList) {$objSearcher.PropertiesToLoad.Add($i) | Out-Null} 

            $colResults = $objSearcher.FindAll()
            $SMTPAddress = ($colResults[0].Properties.proxyaddresses | where {$_ -clike "SMTP*"}) -replace "SMTP:"
        }

        If ($SMTPAddress) {
            Write-Verbose "Found the following SMTP address : $SMTPAddress"
        }#If ($SMTPAddress) 
        Else {
            Write-Warning "No email address specifed, can't connect to EWS without one"
            Break
        }#Else
    }

    If ($SMTPAddress -as [Net.Mail.MailAddress]) {
        Write-Verbose "Pointing service object to the following address $SMTPAddress"
    }#If ($SMTPAddress -as [Net.Mail.MailAddress])
    Else {
        Write-Warning "$SMTPAddress is not a valid email address"
        Break
    }#Else

    Write-Verbose "Checking if the runtime or a seperate account will be used for impersonation"
    If ($ImpersonationCredential) {
        Write-Verbose "Secondary account for Impersonation specifed ($ImpersonationCredential.UserName)"

        $creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(), $Credentials.GetNetworkCredential().password.ToString())  
        $EWSService.Credentials = $creds     
    
        Write-Verbose "Saving impersonation credentials for EWS service"
        $EWSService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $SMTPAddress)
    }#If ($ImpersonationCredential)
    Else {
        Write-Verbose "Runtime account ($ENV:Username) will be used to authenticate"
        $EWSService.UseDefaultCredentials = $true   
    }#Else
    

    #Set the EWS URL, the EWS URL is needed if the the runtime user is specifed since they may not have a mailbox
    If ($EWSURL) {
        Write-Verbose "Using the specifed EWS URL of $EWSURL"
        $EWSService.URL = New-Object Uri($EWSURL)
    }#If If ($EWSURL)
    Else {
        Write-Verbose "Using the AutoDiscover to find the EWS URL for $SMTPAddress"
        $EWSService.AutodiscoverUrl($SMTPAddress, {$True})
    }#Else
    #Now Return the Service Object
    Return $EWSService
}

Function Send-ErrorReport {
    <#
    .SYNOPSIS
        Used to create a one line method to write an error warning to the console, send an error email, and exit a script when an terminal error is encountered
    .DESCRIPTION
        Send-ErrorReport Takes a Subject and message body field and uses the Subject to write a warning messages to the console and as the email subject. It uses the body as the body of the email that will be sent. You can also specify and SMTP Server, From, and To address or set the default options. Once complete the script sends the exit command to script calling the function 
    .PARAMETER Subject
        The subject of the email to be sent and the warning message that is sent to the console
    .PARAMETER body 
        The body of the email sent
    .PARAMETER HaltScript
        Sends the Exit command to the script caling the function, used to report on terminating errors
    .EXAMPLE
        PS C:\powershell> Send-ErrorReport -Subject "can't load EWS module" -Body "can't load EWS module, verify this path : $EWSManagedApiPath" -HaltScript

        WARNING: can't load EWS module

        ------------------
        Description 
        ------------------
        
        This will send an email through the specifed SMTP server, from and to the specifed addresses with the specifed subject and body and use the sbubject to write a warning message. Then function will stop the script that called it
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [string]$Subject,
        
        [Parameter(Position = 1, Mandatory = $true)]
        [string]$body,
        
        [Parameter()]
        [switch]$HaltScript
    )   
    
    Write-Warning $Subject
    #Appened Script name to email Subject
    If ($EmailErrorsFrom -and $EmailErrorsTo -and $SMTPServer) {
        $Subject = "Get-AttachmentFromEmail.ps1 : " + $Subject
        send-mailmessage -from $EmailErrorsFrom -to $EmailErrorsTo -smtpserver $SMTPServer -subject $Subject -Body ($body | Out-String) -BodyAsHtml
    }
    If ($HaltScript) {Exit 1}
}

function Test-Write {
    #Pulled from http://stackoverflow.com/questions/9735449/how-to-verify-whether-the-share-has-write-access
    [CmdletBinding()]
    param (
        [parameter()] [ValidateScript( {[IO.Directory]::Exists($_.FullName)})]
        [IO.DirectoryInfo] $Path
    )
    try {
        $testPath = Join-Path $Path ([IO.Path]::GetRandomFileName())
        [IO.File]::Create($testPath, 1, 'DeleteOnClose') > $null
        # Or...
        <# New-Item -Path $testPath -ItemType File -ErrorAction Stop > $null #>
        return $true
    }
    catch {
        return $false
    }
    finally {
        Remove-Item $testPath -ErrorAction SilentlyContinue
    }
}

Function Get-EWSFolderIdFromPath {  
    <#
    .SYNOPSIS
        When provided with a Public folder path this funciton will reutnr the folder ID of that public folder
    .DESCRIPTION
        Given a public folder path this function will split the path and loop through the public folder structure until it finds the folder specified in the path. Once found the folder ID will be returned as a string
    .PARAMETER FolderPath
        The full path of the public folder the function will provide an ID for (e.g. "\UK\HR\New Hires")
    .EXAMPLE
        PS C:\powershell> Get-EWSFolderIdFromPath -FolderPath "Managers\Reports\2015 -Service $EWSService

        ------------------
        Description 
        ------------------
        
        This will return the ID of the public folder specified
    .NOTES
        Lifted from Glen Scales : http://gsexdev.blogspot.com/2013/08/public-folder-ews-how-to-rollup-part-1.html
    #>
    
    
    param (  
        [Parameter(Position = 0, Mandatory = $True)]
        $FolderPath,
            
        [Parameter(Position = 1, Mandatory = $True)]
        [Object]$service
    )  
    process {  
        ## Find and Bind to Folder based on Path    
        #Define the path to search should be seperated with \    
        #Bind to the MSGFolder Root    
        Write-verbose "Public folder specifed, attempting to bind to the public folder root"
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)     
        $tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid)    
        #Split the Search path into an array    
        $fldArray = $FolderPath.Split("\")  
        Write-verbose "Looking for the public folder $FolderPath by splitting the path and search from the root down"
        #Loop through the Split Array and do a Search for each level of folder  
        for ($lint = 1; $lint -lt $fldArray.Length; $lint++) {  
            #Perform search based on the displayname of each folder level  
            $fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1)  
            $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $fldArray[$lint])  
            Write-verbose "Looking for $($fldArray[$lint])"
            $findFolderResults = $service.FindFolders($tfTargetFolder.Id, $SfSearchFilter, $fvFolderView)  
            if ($findFolderResults.TotalCount -gt 0) {  
                Write-Verbose "Found, moving to next folder level"
                foreach ($folder in $findFolderResults.Folders) {  
                    $tfTargetFolder = $folder                 
                }  
            }  
            else {  
                Write-Warning "Folder Not Found"   
                $tfTargetFolder = $null   
                break   
            }      
        }   
        if ($tfTargetFolder -ne $null) { 
            Write-Verbose "Folder $FolderPath found!"
            return $tfTargetFolder.id
        } 
    } 
} 

#endregion

#region PreWork

#Remove double qoutes from Subject, File Path, and SMTP address (just in case)
#Need to find a way around this
$EmailSubject = $EmailSubject -replace "`""
$SavePath = $SavePath -replace "`""
$SMTPAddress = $SMTPAddress -replace "`""

#Public folder should start with \
if ($PublicFolderPath) {
    if ($PublicFolderPath.StartsWith("\") -eq $FALSE) {$PublicFolderPath = "\" + $PublicFolderPath}
}

Write-verbose "Verifying that the specifed paths can be written to by the runtime account" 
If (-not $SkipWriteTest) {
    Try {
        If (-not (Test-Write $SavePath -ErrorAction Stop)) {Send-ErrorReport -Subject "Save path cannot be accessed" -body "Please the following user name <B>$($ENV:USERNAME)<?B>, has access to this path : <B>$SavePath</B>" -HaltScript}
    } #Try 
    Catch {Send-ErrorReport -Subject "Save path cannot be accessed" -body "Please the following user name <B>$($ENV:USERNAME)<?B>, has access to this path : <B>$SavePath</B>" -HaltScript} #Catch
}

#EWS requires double back whacks "\\" for each normal back whack "\" when
#it comes to the file path to save to
#E.g. C:\Windows\System32 needs to look like C:\\Windows\\System32
$EWSSavePath = $SavePath -replace "\\", "\\"

#Create EWS Service
#I kept running into scoping problems with this function for some reason, so Now we force it via a splat
$Splat = @{}
if ($SMTPAddress) {$Splat.add('SMTPAddress', $SMTPAddress)}
if ($EwsUrl) {$Splat.add('EwsUrl', $EwsUrl)}
if ($ExchangeVersion) {$Splat.add('ExchangeVersion', $ExchangeVersion)}
if ($IgnoreSSLCertificate) {$Splat.add('IgnoreSSLCertificate', $IgnoreSSLCertificate)}
if ($EWSManagedApiPath) {$Splat.add('EWSManagedApiPath', $EWSManagedApiPath)}
if ($EWSTracing) {$Splat.add('EWSTracing', $EWSTracing)}
if ($ImpersonationCredential) {$Splat.add('ImpersonationCredential', $ImpersonationCredential)}

$EWSservice = New-EWSServiceObject @Splat
#endregion

#region Main Program Work

If ($PublicFolderPath) {
    $SearchFolderID = Get-EWSFolderIdFromPath -FolderPath $PublicFolderPath -Service $EWSService
}
Else {
    $SearchFolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
}

Write-verbose "Checking if the specifed folder is empty"
#This will be used on in the search query so that we search the entire folder
Try {$FolderToSearch = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($EWSservice, $SearchFolderID)}
Catch {
    $_
    Send-ErrorReport -Subject "Account is not permissioned for Impersonation" -body "Please verify that the following account <B>$ImpersonationUserName</B> has the ability to impersonate <B>$SMTPAddress</B>" -HaltScript
}

If ($FolderToSearch.TotalCount -eq 0 -and -not $PublicFolderPath) {
    #TODO 2013 public folders return 0 items? 
    #Public folders return zero for some reason
    Write-Warning "Specifed folder is empty, exiting script or the Runtime Account of $($ENV:USERNAME) does not have impersonation rights for $SMTPAddress"
    Exit 1
}

if (-not $PublicFolderPath) {
    $NumEmailsToReturn = $FolderToSearch.TotalCount
}
Else {
    $NumEmailsToReturn = 500
}

#Create mailbox view
$view = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList $NumEmailsToReturn
#Define properties to use with our search filter and view
<# PROPERTY INFO
Microsoft.Exchange.WebServices.Data namespace
    http://msdn.microsoft.com/en-us/library/office/microsoft.exchange.webservices.data(v=exchg.80).aspx
EmailMessageSchema fields
    http://msdn.microsoft.com/en-us/library/office/microsoft.exchange.webservices.data.emailmessageschema_fields(v=exchg.80).aspx
ItemSchema fields
    http://msdn.microsoft.com/en-us/library/office/microsoft.exchange.webservices.data.itemschema_fields(v=exchg.80).aspx
NOTE
In general, the Sender is the actual originator of the email message.
The From Address, in contrast, is simply a header line in the email that may or may not be taken to mean anything.
Sadly we can only search for the from property
#>
$propertyset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet (
    [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly,
    [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject,
    [Microsoft.Exchange.WebServices.Data.ItemSchema]::HasAttachments,
    [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::From,
    [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived
)
#Assign view to porperty set
$view.PropertySet = $propertyset

Write-verbose "Checking to see which search type best fits the search query"
if (($PublicFolderPath -and $ExchangeVersion -notlike "Exchange2013*") -or $ForceSearchFilter) {
    Write-Verbose "Using Exchange Webservices SearchFilters"
    #Prior to Exchange 2013, you couldn't search public folders with an Advanced Query Syntax (AQS) query
    #instead have to use the SearchFilter and it's subclasses
    #https://msdn.microsoft.com/en-us/library/office/microsoft.exchange.webservices.data.searchfilter.containssubstring(v=exchg.80).aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1
    #Note that Powershell needs to use the "+" symbol to access enums on .net functions
    #Creat a search filter collection and use AND for each option
    $query = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)

    #Creat an option for Exact search
    if ($ExactSubjectSearch) {
        $SubjectSearch = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject, $EmailSubject)
    }
    Else {
        $SubjectSearch = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubString([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject, $EmailSubject)
    }
    $query.Add($SubjectSearch)

    #If today then add a search filter for any item equal or greater to midnight today
    If ($TodayOnly) {
        $Date = (get-date "00:00:00")
        $TodaySearch = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $Date)
        $query.Add($TodaySearch)
    }#If ($TodayOnly)
    Elseif ($ReceivedInThePastNumberOfDays) {
        $Date = (get-date).adddays($ReceivedInThePastNumberOfDays)
        $TodaySearch = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $Date)
        $query.Add($TodaySearch)
    }
    Else {
        #Default is no date range, meaning all items
    }

    #if we are looking for attachments and not the email body add that option
    If ($ExportEmailBody) {
        $SearchAttachments = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::HasAttachments, $True)
        $query.Add($SearchAttachments)
    }#If ($ExportEmailBody

    #if we are looking for specific From address add that options
    If ($From -and -not $PublicFolderPath) {
        #TODO, how do we add multiple from address to search from. Why did I bother with this in the first place?
        #If ($From) {
        #    Write-warning "Can only search 1 from address when targeting Public folders on a version of Exchange lower then 2013"
        #   $From = $From[0]
        #}
        $SearchFrom = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::From, (New-Object Microsoft.Exchange.WebServices.Data.EmailAddress($NULL, $From, "SMTP")))
        $query.Add($SearchFrom)
    }#If ($From -and -not $PublicFolderPath)
    Elseif ($From -and $PublicFolderPath) {
        Write-warning "From is currenly not supported on 2010 and lower servers OR when using search filters instead of AQS when searching for Public folders, ignoring"
    }#Elseif ($From -and $PublicFolderPath) 
    Else {}
} #if ($PublicFolderPath -and $ExchangeVersion -notlike "Exchange2013*")

Else {      
    Write-Verbose "Using Advanced Query Syntax (AQS)"
    #We are dealing with the Inbox or a 2013 Public folder
    #Define the AQS search quuery
    #An EWS exact search puts the search term in "" so if we need that then add the qoutes again
    #INFO : http://msdn.microsoft.com/en-us/library/office/ee693615(v=exchg.150).aspx
    #     https://msdn.microsoft.com/en-us/library/office/dn579420(v=exchg.150).aspx
    If ($ExactSubjectSearch) {
        $EmailSubject = "`"$EmailSubject`""
        #Orignally wanted to add an exact search for the from address but that would require an entry like "Mello, John <John.Mello@msx.bala.susq.com>"
        #Since this how the from address is stored
        #if ($From) {
        #   $From = $From | 
        #                Foreach-object {"`"$_`""}
        #} #If From
    } #If ExactSubjectSearch
    Else {
        $EmailSubject = "($($EmailSubject -replace " "," OR "))"
    }

    #Buidling search Query
    #INFO : http://msdn.microsoft.com/en-us/library/office/ee693615(v=exchg.150).aspx
    #If the email body is needed then don't limit search for just attachments, otherwise we want just attachments
    If ($ExportEmailBody) {$query = "Subject:$EmailSubject"}
    Else {$query = "Subject:$EmailSubject AND HasAttachments:True"}
    #Add Date Range
    If ($TodayOnly) {
        $query = "$query AND Received:today"
    }
    ELseIf ($ReceivedInThePastNumberOfDays) {
        $start = (get-date).adddays($ReceivedInThePastNumberOfDays)
        $query = "$query AND Received:$start..today"
    }
    Else {
        #Leave out the date criteria
    }
    #Add a search query for each from address
    If ($From) {
        If ($From.count -eq 1) {$query = "$query AND From:$From"}
        Else {
            $SenderSearch = " AND (From:"
            Foreach ($Addy in $From) {
                $SenderSearch = "$SenderSearch" + "$Addy OR " 
            }
            $SenderSearch = $SenderSearch.TrimEnd(" OR ") + ")"
            $query = $query + $SenderSearch
        }#Else
    }#If ($From)
}#Else ($PublicFolderPath -and $ExchangeVersion -notlike "Exchange2013*")

    
Write-Verbose "Performing search with the following query : $query"
Try {
    $FoundEmails = $EWSservice.FindItems($SearchFolderID, $query, $view)
}#Try
Catch {
    Write-Warning "Search Failed"
    if ($_.Exception -like "*The search operation could not be completed within the allotted time limit*") {
        Write-Warning "AQS search failed, try forcing a SearchFilter search with the -ForceSearchFilter parameter"
    }#if ($_.Exception -like "*The search operation could not be completed within the allotted time limit*") 
    $_
    Exit 1
}#Catch

#Note that a FindItems search only returns the bare minimum info wise on each email, not evening showing fields you searched agaisnt
#You need to bind to the invididual email item to get any further info
If ($FoundEmails.Totalcount) {
    Write-output "$($FoundEmails.TotalCount) emails found"
    If ($UnZipFiles) {$ZipFileNames = @()}
    Foreach ($Email in $FoundEmails) {
        $emailProps = New-Object -TypeName Microsoft.Exchange.WebServices.Data.PropertySet(
            [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly,
            [Microsoft.Exchange.WebServices.Data.ItemSchema]::Attachments,
            [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject,
            [Microsoft.Exchange.WebServices.Data.ItemSchema]::Body,
            [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived
        )
        $EmailwithAttachments = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($EWSservice, $Email.Id, $emailProps)

        If ($ExportEmailBody) {
            Write-verbose "Exporting email body from '$($Email.subject)'"
            $FileName = (get-date ($EmailwithAttachments.DateTimeReceived) -Format yyyy_MM_dd-hh_mm) + "-" + ($Email.subject -replace ("[{0}]" -f ([RegEx]::Escape([String][System.IO.Path]::GetInvalidFileNameChars())))) + ".HTML"
            If ($NoClobber) {
                If (Get-ChildItem $SavePath -Filter $FileName) {
                    do {
                        $FileName = "$(Get-Random)" + $FileName
                        $FilesInDir = Get-ChildItem $SavePath | 
                            Where-Object {$_.PSIsContainer -eq $FALSE} | 
                            ForEach-Object {$_.Name}
                    } while ($FilesInDir -contains $FileName)
                }
            }
            $EmailwithAttachments.Body.text | Out-File "$EWSSavePath\$FileName"
        }       

        Foreach ($File in $EmailwithAttachments.Attachments) {
            Try {
                Write-verbose "Saving Attachment from '$($Email.subject)' to $SavePath"
                #Just to be safe, remove any invaild characters from the attachment name
                $FileName = $File.name -replace ("[{0}]" -f ([RegEx]::Escape([String][System.IO.Path]::GetInvalidFileNameChars())))
                #Current can't handle Emails that are attacments, so we skip them
                If ($File.contenttype -eq $null -and ($FileName -like '*.msg') -or ($FileName -like '*.eml')) {
                    Write-Warning "Attachment is an email, skipping"
                    $EMLAttachment = $TRUE
                    continue
                }
                Else {$EMLAttachment = $FALSE}
                #If no clobber then generate a random 10 digit number and prepend to the file
                If ($NoClobber) {
                    If (Get-ChildItem $SavePath -Filter $FileName) {
                        do {
                            $FileName = "$(Get-Random)" + $FileName
                            $FilesInDir = Get-ChildItem $SavePath | 
                                Where-Object {$_.PSIsContainer -eq $FALSE} | 
                                ForEach-Object {$_.Name}
                        } while ($FilesInDir -contains $FileName)
                    }
                }
                $File.load($EWSSavePath + "\\" + $FileName) 
                If (($FileName -like "*.zip") -and $UnZipFiles) {$ZipFileNames += $SavePath + "\" + $FileName}
            }
            Catch {Send-ErrorReport -Subject "Attachment save path cannot be accessed" -body "Please verify the following path <B>$SavePath</B><BR>Full Error:<BR> $_" -HaltScript}
        }
        If ($DeleteEmail -and (-not $EMLAttachment)) {
            Write-verbose "Deleting the following email '$($Email.subject)'"
            If (-not $PublicFolderPath) {
                $Email.Delete("MoveToDeletedItems")
            }
            Else {
                #Public Folders require a hard delete
                $Email.Delete("HardDelete")
            }
        }
        ElseIf ($DeleteEmail) {Write-Warning "DeleteEmail specifed but embedded EML file in email could not be downloaded"}
    }
    If ($UnZipFiles) {
        Try {
            Write-Verbose "Unzipping attachment"
            #TODO Wrapp this in a try/catch, maybe by using ShouldProcess?
            $shell = new-object -com shell.application
            $Location = $shell.namespace($SavePath)
            foreach ($ZipFile in $ZipFileNames) {
                $ZipFolder = $shell.namespace($ZipFile)
                $Location.Copyhere($ZipFolder.items(), 8)
                Remove-item $ZipFile
                #TODO : Why can't I have duplicate files renamed when being unzipped?
                #http://msdn.microsoft.com/en-us/library/windows/desktop/ms723207(v=vs.85).aspx
                #http://msdn.microsoft.com/en-us/library/windows/desktop/bb787866(v=vs.85).aspx
                #http://technet.microsoft.com/en-us/library/ee176633.aspx
                #http://stackoverflow.com/questions/2359372/how-do-i-overwrite-existing-items-with-folder-copyhere-in-powershell
            }
        }
        Catch {Send-ErrorReport -Subject "Attachment save path cannot be accessed" -body "Please verify the following path <B>$SavePath</B><BR>Full Error:<BR> $_" -HaltScript}
    }
}#If ($FoundEmails) 
Else {Write-output "No emails found"}
#EndRegion