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
    This script will perform the "day 2" offboarding procedure for the user whose first and last name is input into the input which appears when script is first run.

    .DESCRIPTION 
    This script will perform the "day 2" offboarding procedure for the user whose first and last name is input into the input which appears when script is first run.  Actions performed include Removing the user from all distribution lists, adding them to the "departed users" security group, enabling archive (if notn already enabled), converting the mailbox to a shared mailbox, and removing all licenses associated with that user account.
	
    .NOTES
	
	
    .EXAMPLE
    .\disable-user-day-2-man.ps1

#> 

Function obtain-AdminPass {
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
  $inputbox = New-Object System.Windows.Forms.Form
  $inputbox.Text = 'Office 365 Admin Password'
  $inputbox.Size = New-Object System.Drawing.Size(300,150)

  $mLabel1 = New-Object System.Windows.Forms.Label
  $mLabel1.Text="Please input password for arbor office 365 account."
  $mLabel1.Top="18"
  $mLabel1.Left="5"
  $mLabel1.Anchor="Left,Top"
  $mLabel1.Size = New-Object System.Drawing.Size(295,23)
  $inputbox.Controls.Add($mLabel1)

  $password = New-Object Windows.Forms.MaskedTextBox
  $password.PasswordChar = '*'
  $password.Top = "50"
  $password.Left = "50"
  $password.Size = New-Object System.Drawing.Size(180,65)

  $inputbox.Controls.Add($password)

  $mButton1 = New-Object System.Windows.Forms.Button 
  $mButton1.Text="OK"
  $mButton1.Top="80"
  $mButton1.Left="90"
  $mButton1.Anchor="Left,Top"
  $mButton1.Size = New-Object System.Drawing.Size(100,23)
  $mButton1.Add_Click({$inputbox.Close()})
  $inputbox.Controls.Add($mButton1)

  $inputbox.ShowDialog() | Out-Null

  Return $password.Text
}
function Get-SBCredential
{
  [CmdletBinding(ConfirmImpact = 'Low')]
  Param (
    [Parameter(Mandatory = $false,
           ValueFromPipeLine = $true,
           ValueFromPipeLineByPropertyName = $true,
           Position = 0)]
    [String]$UserName = $env:USERNAME,
    [Parameter(Mandatory = $false,
           Position = 1)]
    [Switch]$Refresh = $false
  )
	
  $CredPath = "$env:Temp\$($UserName.Replace('\', '_')).txt"
  if ($Refresh)
  {
    try
    {
      Remove-Item -Path $CredPath -Force -Confirm:$false -ErrorAction Stop
    }
    catch
    {
    }
  }
  if (!(Test-Path -Path $CredPath))
  {
    $Temp = obtain-AdminPass | ConvertTo-SecureString -AsPlainText -Force
    try
    {
      ConvertFrom-SecureString $Temp | Out-File $CredPath -ErrorAction Stop
    }
    catch
    {
    }
  }
  $Pwd = Get-Content $CredPath | ConvertTo-SecureString
  try
  {
    New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName, $Pwd -ErrorAction Stop
  }
  catch
  {
  }
}
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
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
Use-RunAs
Try
{
  Get-Service -Name msoidsvc -ErrorAction Stop
}
Catch
{
    Get-PackageProvider -Name NuGet -ForceBootstrap
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    Try
    {
        Install-module Carbon -Force -ErrorAction Stop
    }
    Catch
    {
        Install-module Carbon -Force -AllowClobber
    }
    Install-Module -Name AzureAD
    Install-Module MSonline
    Invoke-WebRequest 'https://download.microsoft.com/download/7/1/E/71EF1D05-A42C-4A1F-8162-96494B5E615C/msoidcli_64bit.msi' -OutFile "$env:USERPROFILE\Downloads\signin.msi"
    Install-Msi -Path "$env:USERPROFILE\Downloads\signin.msi"
    Invoke-WebRequest 'https://bposast.vo.msecnd.net/MSOPMW/Current/amd64/AdministrationConfig-en.msi' -OutFile "$env:USERPROFILE\Downloads\azure.msi"
    Install-Msi -Path "$env:USERPROFILE\Downloads\azure.msi"
    Invoke-WebRequest 'https://download.microsoft.com/download/0/2/E/02E7E5BA-2190-44A8-B407-BC73CA0D6B87/sharepointonlinemanagementshell_6112-1200_x64_en-us.msi' -OutFile "$env:USERPROFILE\Downloads\share.msi"
    Install-Msi -Path "$env:USERPROFILE\Downloads\share.msi"
    Invoke-WebRequest 'https://download.microsoft.com/download/2/0/5/2050B39B-4DA5-48E0-B768-583533B42C3B/SkypeOnlinePowershell.exe' -OutFile "$env:USERPROFILE\Downloads\skype.exe"
    Start-Process "$env:USERPROFILE\Downloads\skype.exe" -ArgumentList '/passive /norestart'
}
    $cred2 = Get-SBCredential -UserName "mboxadmin@arborpharma.com" -Refresh
    $connect = Connect-MsolService -Credential $cred2
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred2 -Authentication Basic -AllowRedirection
    Import-PSSession $Session
    $name = [Microsoft.VisualBasic.Interaction]::InputBox("Enter user's first and last name", "Disable AD user", "")
    $priname = $name.Replace(' ', '.') + '@arborpharma.com'
    Write-Host $priname
    Get-DistributionGroup | Remove-DistributionGroupMember -Member $priname -Confirm:$false
    Add-DistributionGroupMember -Identity 'Departed Users' -Member $priname -Confirm:$false
    Get-Mailbox -Identity $priname | Enable-Mailbox -Archive
    Get-Mailbox -Identity $priname | Set-Mailbox -Type Shared
    Set-MsolUserLicense -UserPrincipalName $priname -RemoveLicenses "reseller-account:ENTERPRISEPACK" -Verbose
    Set-MsolUserLicense -UserPrincipalName $priname -RemoveLicenses "reseller-account:VISIOCLIENT" -verbose
    Set-MsolUserLicense -UserPrincipalName $priname -RemoveLicenses "reseller-account:RIGHTSMANAGEMENT" -verbose
    Set-MsolUserLicense -UserPrincipalName $priname -RemoveLicenses "reseller-account:POWER_BI_STANDARD" -verbose
    Set-MsolUserLicense -UserPrincipalName $priname -RemoveLicenses "reseller-account:EXCHANGEENTERPRISE" -verbose
    Set-MsolUserLicense -UserPrincipalName $priname -RemoveLicenses "reseller-account:STANDARDPACK" -verbose
    Get-PSSession | Remove-PSSession