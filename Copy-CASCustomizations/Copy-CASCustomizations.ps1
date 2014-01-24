<#
.SYNOPSIS
       Copies the specified customized OWA and ECP files from a source Exchange server to a target Exchange server
.DESCRIPTION
    Using a predefined list in the script, the script will make copies of the original files and append "-Original" to those files while copying over the new versions.
    It will also make the predefined changes in the script to the logoff.apsx file
    This script assumes that exchange is installed on the C drive in the default directory
.PARAMETER TargetComputer 
    The Exchange client access server that will have it's OWA and ECP files customized
.PARAMETER SourceComputer
       The computer that will be the 
.EXAMPLE
       PS C:\> .\Copy-CASCustomizations –TargetComputer Labxchchbal500 –SourceComputer 2010caslab01
       Customizations successfully applied to Labxchchbal500 using 2010caslab01 as the source
 
       Description
       -----------
       On the target computer Labxchchbal500 rename the Original files and copy the customized files from 2010caslab01
 
.NOTES
       AUTHOR: John Mello
       CREATED : 01/23/2014 
       CREATED BECAUSE: Since Exchange removes the customizations during every SP and CU upgrade, and the process is tedious and error prone, this task was flaggged for automation
#>
 
 
[CmdletBinding()]
Param(
       [Parameter(Mandatory=$True)] 
    [string]$TargetComputer,
       
       [Parameter(Mandatory=$True)] 
    [string]$SourceComputer
)
 
Function ReanameOrginalCopyFromSource {
    <#
        .SYNOPSIS
            Helper function that takes a list of files and renames them on a target computer to "*-original*" and copies a version from the source computer
        .DESCRIPTION
        .PARAMETER FileCollection 
            Collection of files to move
        .PARAMETER TargetPath 
            The target path that will have these files renamed and then copied over from the source
        .PARAMETER SourcePath
               The source path the files will be copied from
        .EXAMPLE
               PS C:\> ReanameOrginalCopyFromSource -FileCollection $FilesToUpdate -TargetPath C:\Files -SourcePath D:\Files
 
               Description
               -----------
               Using the files in the list $FilesToUpdate rename the versions in C:\Files to "*-Original*" and copy the versions from D:\Files to C:\Files
    #>
   [CmdletBinding()]
    Param(
           [Parameter(Mandatory=$True)] 
        [Object]$FileCollection,
 
        [Parameter(Mandatory=$True)] 
        [String]$TargetPath,
 
        [Parameter(Mandatory=$True)] 
        [String]$SourcePath
       
    )
 
    Foreach ($File in $FileCollection) {
           Try {
                  $FileDetails = Get-item -Path "$TargetPath$File" -ErrorAction Stop
                  Write-Verbose "Renaming file $($FileDetails.Name) on Target server $TargetComputer"
            Rename-Item -Path $FileDetails.FullName -NewName "$($FileDetails.BaseName)-Original$($FileDetails.extension)" -ErrorAction Stop
                  Write-Verbose "Copying file $($FileDetails.Name) from Source server $SourceComputer to Target server $TargetComputer"
            Copy-Item -Path "$SourcePath$File" -Destination "$TargetPath$File" -ErrorAction Stop
           }
           Catch {
                  Write-Warning "Issue copying $File from $SourcePath to $TargetPath"
            Exit 1
           }
    }
}
 
Write-Verbose "Building path to Client Access directory on $TargetComputer and $SourceComputer"
$TargetPath = "\\$TargetComputer\C$\Program Files\Microsoft\Exchange Server\V14\ClientAccess"
$PullPath = "\\$SourceComputer\C$\Program Files\Microsoft\Exchange Server\V14\ClientAccess"
 
 
Write-Verbose "Testing access to Exchange folders on $TargetComputer and $SourceComputer"
If (($TargetPath,$PullPath | Test-path) -contains $False) {
    Write-Warning "Cannot access Exchange install directory on one of the servers, please check"
    Write-Output $_
    Exit 1
}
 
Write-Verbose "Building path to the various Client Access Directories"
#Get the latest OWA version folder path from the target and the path
$OWATargetPath =  (Get-ChildItem "$TargetPath\OWA" | 
       Where-object {$_.Name -match "^1[1-9]\."} | 
       Sort-Object -Descending | 
       Select-Object -First 1).FullName + "\Themes"
 
$OWAPullPath = (Get-ChildItem "$PullPath\OWA" | 
       Where-object {$_.Name -match "^1[1-9]\."} | 
       Sort-Object -Descending | 
       Select-Object -First 1).FullName + "\Themes"
 
#Get the latest ECP version folder path from the target and the path
$ECPTargetPath =  (Get-ChildItem "$TargetPath\ECP" | 
       Where-object {$_.Name -match "^1[1-9]\."} | 
       Sort-Object -Descending | 
       Select-Object -First 1).FullName + "\Themes\Default"
 
$ECPPullPath = (Get-ChildItem "$PullPath\ECP" | 
       Where-object {$_.Name -match "^1[1-9]\."} | 
       Sort-Object -Descending | 
       Select-Object -First 1).FullName + "\Themes\Default"
 
 
Write-Verbose "Testing access to Exchange folders on $TargetComputer and $SourceComputer"
If (($OWATargetPath,$OWAPullPath,$ECPTargetPath,$ECPPullPath | Test-path) -contains $False) {
    Write-Warning "Cannot access ClientAccess directory on one of the servers, please check"
    Write-Output $_
    Exit 1
}
 
 
#List of OWA theme files
$OWAThemeFiles = @("\base\Cnv-draft.png",
       "\base\Gradienth.png",
       "\base\Gradientv.png",
       "\base\Headerbgmain.png",
       "\base\Headerbgmainrtl.png",
       "\base\Headerbgright.png",
       "\base\Themepreview.png",
       "\base\Premium.css",
       "\base\Csssprites.png",
       "\base\Csssprites.css",
       "\resources\Lgnbotl.gif",
       "\resources\Lgnbotm.gif",
       "\resources\Lgnbotr.gif",
       "\resources\Lgnexlogo.gif",
       "\resources\Lgnleft.gif",
       "\resources\Lgnright.gif",
       "\resources\Lgntopl.gif",
       "\resources\Lgntopm.gif",
       "\resources\Lgntopr.gif",
       "\resources\Logon.css"
)
 
#List of ECP theme files
$ECPThemeFiles = @("\Headerbgmain.png",
       "\Headerbgmain-rtl.png",
       "\Headerbgright.png",
       "\Mainnavigationsprite.png",
       "\Mainnavigationsprite.css",
       "\Navigation.css",
       "\Editorstyles.css"
)
 
#OWA Themes : Rename old files and copy new files
ReanameOrginalCopyFromSource -FileCollection $OWAThemeFiles -TargetPath $OWATargetPath -SourcePath $OWAPullPath
 
#OWA Auth : Rename Logoff.aspx file and copy new file and set xml customizations
Try {
    $FileDetails = Get-Item "$TargetPath\OWA\Auth\Logoff.aspx" -ErrorAction Stop
    Write-Verbose "Renaming file $($FileDetails.Name) on Target server $TargetComputer"
    Rename-Item -Path $FileDetails.Fullname -NewName "$($FileDetails.BaseName)-Original$($FileDetails.extension)" -ErrorAction Stop
    Write-Verbose "Copying file $($FileDetails.Name) from Source server $SourceComputer to Target server $TargetComputer"
    Copy-Item -Path "$PullPath\OWA\Auth\Logoff.aspx"  -Destination "$TargetPath\OWA\Auth" -ErrorAction Stop
 
    Write-verbose "Making changes to $($FileDetails.Name) on Target server $TargetComputer"
    $FileToChange = Get-Content "$TargetPath\OWA\Auth\Logoff.aspx" -ErrorAction Stop
    $FileToChange -replace "document.execCommand\(`"ClearAuthenticationCache`"\);", "document.execCommand\(`"ClearAuthenticationCache`", false\);" |
        Out-Null
    $FileToChange | Out-File -FilePath "$TargetPath\OWA\Auth\Logoff.aspx" -ErrorAction Stop
}
Catch {
    Write-Warning "Cannot edit or access the Logoff.aspx file"
    Exit1
}
 
#ECP Themes : Rename old files and copy new files
ReanameOrginalCopyFromSource -FileCollection $ECPThemeFiles -TargetPath $ECPTargetPath -SourcePath $ECPPullPath
 
Write-Output "Customizations successfully applied to $TargetComputer using $SourceComputer as the source"

________________________________________

IMPORTANT: The information contained in this email and/or its attachments is confidential. If you are not the intended recipient, please notify the sender immediately by reply and immediately delete this message and all its attachments. Any review, use, reproduction, disclosure or dissemination of this message or any attachment by an unintended recipient is strictly prohibited. Neither this message nor any attachment is intended as or should be construed as an offer, solicitation or recommendation to buy or sell any security or other financial instrument. Neither the sender, his or her employer nor any of their respective affiliates makes any warranties as to the completeness or accuracy of any of the information contained herein or that this message or any of its attachments is free of viruses.
