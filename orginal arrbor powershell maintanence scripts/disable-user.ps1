
<#PSScriptInfo

.VERSION 1.0

.AUTHOR William Higgs

.COMPANYNAME Providyn

.COPYRIGHT 2017, Providyn

.TAGS 

.LICENSEURI 

.PROJECTURI 

.ICONURI 

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES


#>

<# 
.SYNOPSIS
   	This script will perform the "day 1" offboarding procedure for the user whose first and last name is input into the input which appears when script is first run.

.DESCRIPTION 
 	This script will perform the "day 1" offboarding procedure for the user whose first and last name is input into the input which appears when script is first run.  Actions performed include removing the user from all groups except the domain user group and any group containing the text "Xenoport", adds the appropriate values to the "mailnickname" and "MSExchangeHidefromAddressbook" attributes, modify the display name of the user as required, moves the user to the disabled users organizational unit depending on the users department, and disables the account.
	
.NOTES
	
	
.EXAMPLE
	.\disable-user.ps1

#> 

function Use-RunAs
{
    # Check if script is running as Adminstrator and if not use RunAs 
    # Use Check Switch to check if admin 
     
    param([Switch]$Check)
     
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()`
        ).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
         
    if ($Check) { return $IsAdmin }
    if ($MyInvocation.ScriptName -ne "")
    {
        if (-not $IsAdmin)
        {
            try
            {
                $arg = "-file `"$($MyInvocation.ScriptName)`""
                Start-Process "$psHome\powershell.exe" -Verb Runas -ArgumentList $arg -ErrorAction 'stop'
            }
            catch
            {
                Write-Warning "Error - Failed to restart script with runas"
                break
            }
            exit # Quit this session of powershell
        }  
    }  
    else
    {
        Write-Warning "Error - Script must be saved as a .ps1 file first"
        break
    }
}
Use-RunAs
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$user = [Microsoft.VisualBasic.Interaction]::InputBox("Enter user's first and last name or employee id", "Disable AD user", "")
Add-Type -Assembly Microsoft.VisualBasic
If ([Microsoft.VisualBasic.Information]::IsNumeric($user))
{
    $sam = Get-ADUser -Filter {EmployeeID -eq $user} -Properties *
}
Else
{
    $sam = Get-ADUser -Filter {Name -like $user} -Properties *
}
Get-ADUser $sam -Properties SamAccountName | Get-ADPrincipalGroupMembership | ForEach-Object {
If (($_.SamAccountName -ne 'Domain Users') -and ($_.SamAccountName -notlike 'Xenoport'))
{
Remove-ADGroupMember -Identity $_ -Members $sam -Confirm:$false -Verbose
}
}
If (!($sam.DisplayName.StartsWith("zz")))
{
$newdi = "zz$($sam.DisplayName)"
}
Else
{
$newdi = $sam.DisplayName
}
$samacc = $sam.SamAccountName
$desti = $sam.DistinguishedName
$all = $sam.GivenName + '.' + $sam.SurName
$date = Get-Date -UFormat "%Y/%m/%d"
$type = $sam.office
Write-Host "$newdi`n$samacc"
Set-ADUser $sam -DisplayName $newdi -Description "Disabled $date" -Add @{msExchHideFromAddressLists="TRUE";mailNickname="$all"} -Verbose
Disable-ADAccount $sam
If ($type -eq 'Field'){
Move-ADObject -Identity $desti -TargetPath "OU=Disabled Users DO-NOT-DELETE,OU=Arbor FIELD,DC=ArborPharma,DC=Local" -Verbose}
Else {
Move-ADObject -Identity $desti -TargetPath "OU=Disabled Users DO-NOT-DELETE,OU=Arbor ATL,DC=ArborPharma,DC=Local" -Verbose}