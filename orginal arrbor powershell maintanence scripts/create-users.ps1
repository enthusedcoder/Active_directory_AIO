
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
   	This script will create users in active directory for arbor Pharmaceutical.

.DESCRIPTION 
 	This script will create users in active directory based on the "Pre hire to new hire" report that is called as this scripts only parameter.  The report must be converted to a csv and all flashy content removed so that the first row contains the labels which define the columns.
	
.NOTES
	To create the CSV file that is needed for this script to work, open the pre hire to new hire report, remove all of the rows above the row containing the column description information, then save as, then select "CVS" under "Save as Type" drop down.
	
.EXAMPLE
	.\create-users.ps1 'C:\Path\to\csv\file.csv'

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
ForEach ($ob in $v) {
$first = $ob."First Name"
$last = $ob."Last Name"
$dept = $ob."Default Department"
$title = $ob."Default Jobs (HR)"
$cell = $ob."Cell Phone"
$phone = $ob."Home Phone"
$email = $ob.Email
$finit = $first.ToUpper().substring(0,1)
$linit = $last.ToUpper().substring(0,1)
$comp = $ob."Employee EIN"
$id = $ob."Employee Id"


<#
Write-Host "Name: $first $last`nID: $id`nCell: $cell`nHome: $phone`nEmail: $email`nCompany: $comp`nDepartment: $dept`nLast initial: $linit`nFirst initial: $finit"

If ($dept.StartsWith('Sales')) {
Write-Host "Field`n`n`n`n"
}
Else
{
Write-host "Atlanta Corporate`n`n`n`n"
}
}
#>

If ($dept.StartsWith('Sales')) {
New-ADUser -Title $title -Company $comp -DisplayName "$first $last" -Department $dept -EmailAddress "$first.$last@arborpharma.com" -GivenName "$first" -AccountPassword (ConvertTo-SecureString -AsPlainText "P@ssw0rd$finit$linit" -Force) -ChangePasswordAtLogon $false -EmployeeID $id -SamAccountName "$first.$last" -MobilePhone $cell -HomePhone $phone -Name "$first $last" -Office 'Field' -UserPrincipalName "$first.$last@arborpharma.com" -Surname $last -Path "ou=Employees,ou=Arbor FIELD,dc=ArborPharma,dc=local" -OtherAttributes @{mailNickname="$first.$last";proxyAddresses="SMTP:$first.$last@arborpharma.com"} -Enabled $true
Add-ADGroupMember "PW SelfService" "$first.$last"
}
Else
{
New-ADUser -Title $title -Company $comp -DisplayName "$first $last" -Department $dept -EmailAddress "$first.$last@arborpharma.com" -GivenName "$first" -AccountPassword (ConvertTo-SecureString -AsPlainText "P@ssw0rd$finit$linit" -Force) -ChangePasswordAtLogon $false -EmployeeID $id -SamAccountName "$first.$last" -MobilePhone $cell -HomePhone $phone -Name "$first $last" -Office 'Atlanta Corporate' -UserPrincipalName "$first.$last@arborpharma.com" -Surname $last -Path "ou=Employees,ou=Arbor ATL,dc=ArborPharma,dc=local" -OtherAttributes @{mailNickname="$first.$last";proxyAddresses="SMTP:$first.$last@arborpharma.com"} -Enabled $true
}
}
