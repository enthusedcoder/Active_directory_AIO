
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
   	This script will perform the "day 1" offboarding procedure for the users designated in a csv file.

.DESCRIPTION 
 	This script will perform the "day 1" offboarding procedure for the users designated in a csv file.  The csv file should only have the columns "First Name" and "Last Name" in the first row, and the first and last names of the users in the indicated columns.  Actions performed include removing the user from all groups except the domain user group and any group containing the text "Xenoport", adds the appropriate values to the "mailnickname" and "MSExchangeHidefromAddressbook" attributes, modify the display name of the user as required, moves the user to the disabled users organizational unit depending on the users department, and disables the account.
	
.NOTES
	
	
.EXAMPLE
	.\disable-usercsv.ps1 'C:\Path\to\csv\file.csv'

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
$v = Import-CSV -Delimiter ',' -Path $Args[0]
foreach ($user in $v)
{
    $dis = $user."First Name" + ' ' + $user."Last Name"
    $sam = Get-ADUser -Filter {Name -like $dis} -Properties *
    Write-Host $dis
    Get-ADUser $sam -Properties SamAccountName | Get-ADPrincipalGroupMembership | ForEach-Object {
        If (($_.SamAccountName -ne 'Domain Users') -and ($_.SamAccountName -notlike 'Xenoport'))
        {
            Remove-ADGroupMember -Identity $_ -Members $sam -Confirm $false -Verbose
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
    Write-Host $newdi
    $samacc = $sam.SamAccountName
    $desti = $sam.DistinguishedName
    $date = Get-Date -UFormat "%Y/%m/%d"
    $all = $sam.GivenName + '.' + $sam.SurName
    $type = $sam.office
    Write-Host "$newdi`n$samacc"
    Set-ADUser $sam -DisplayName $newdi -Description "Disabled $date" -Add @{msExchHideFromAddressLists="TRUE";mailNickname="$all"} -Verbose
    Disable-ADAccount $sam -Verbose
    If ($type -eq 'Field'){
    Move-ADObject -Identity $desti -TargetPath "OU=Disabled Users DO-NOT-DELETE,OU=Arbor FIELD,DC=ArborPharma,DC=Local" -Verbose}
    Else {
    Move-ADObject -Identity $desti -TargetPath "OU=Disabled Users DO-NOT-DELETE,OU=Arbor ATL,DC=ArborPharma,DC=Local" -Verbose}
}