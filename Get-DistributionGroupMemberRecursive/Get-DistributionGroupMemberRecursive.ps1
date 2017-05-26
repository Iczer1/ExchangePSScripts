<#
.SYNOPSIS
Returns a recursive list of distribution group members along with an two extra properties detailing the SourceGroup name and an array of the nested groups
.PARAMETER identity
Name of the group to return the recursive membership of
.PARAMETER MailObjectsOnly
Will return only group members that have a filled in PrimarySMTPAddress
.EXAMPLE
#>
Function Get-DistributionGroupMemberRecursive {

    [CmdletBinding()]
    Param(    
        [Parameter(Mandatory=$TRUE)] 
        [string]$identity,

        [Parameter()] 
        [Switch]$MailObjectsOnly
    )
       Try {
        $Members = Get-DistributionGroupMember -identity $identity -ResultSize Unlimited -ErrorAction stop
        if ($MailObjectsOnly) {
            $Members = $Members |
                Where-Object PrimarySmtpAddress
        }#if ($MailObjectsOnly) 
    }
    Catch [System.Management.Automation.CommandNotFoundException] {
        Write-warning "Command not found, are the Exchange cmdelts loaded?"
        Exit 1
    }#Catch
    Catch{
        Write-Warning "Issue getting group info, details"
        $_
        Exit 1
   }#Catch
    [Array]$NestedPath += $identity

    $Members | 
        ForEach-Object {
            $_ | 
                Add-Member -MemberType NoteProperty -Name SourceGroup -Value $identity -PassThru |
                Add-Member -MemberType NoteProperty -Name NestedGroupPath -Value $NestedPath
            if ($_.RecipientType -match "(MailUniversalSecurityGroup)|(MailUniversalDistributionGroup)") {
                Get-DistributionGroupMemberRecursive -identity $_.SamAccountName
            }#if ($_.RecipientType -match "(MailUniversalSecurityGroup)|(MailUniversalDistributionGroup)")
        }#ForEach-Object
    Return $Members
}#Get-DistributionGroupMemberRecursive 
