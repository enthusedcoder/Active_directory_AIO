function Use-RunAs 
{    
     
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
                $arg = "-windowstyle hidden -file `"$($MyInvocation.ScriptName)`" -Noninteractive"
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
Function Get-FileName
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = "$env:USERPROFILE\Documents"
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv| Excel Documents (*.xlsx)| *.xlsx"
    $OpenFileDialog.ShowDialog() | Out-Null
    return $OpenFileDialog.filename
}
Function Search-User {
  Add-Type -Path (Join-Path -Path (Split-Path $script:MyInvocation.MyCommand.Path) -ChildPath 'bin\CubicOrange.Windows.Forms.ActiveDirectory.dll')

  $DialogPicker = New-Object CubicOrange.Windows.Forms.ActiveDirectory.DirectoryObjectPickerDialog

  $DialogPicker.AllowedLocations = [CubicOrange.Windows.Forms.ActiveDirectory.Locations]::All
  $DialogPicker.AllowedObjectTypes = [CubicOrange.Windows.Forms.ActiveDirectory.ObjectTypes]::Users
  $DialogPicker.DefaultLocations = [CubicOrange.Windows.Forms.ActiveDirectory.Locations]::JoinedDomain
  $DialogPicker.DefaultObjectTypes = [CubicOrange.Windows.Forms.ActiveDirectory.ObjectTypes]::Users
  $DialogPicker.ShowAdvancedView = $false
  $DialogPicker.MultiSelect = $true
  $DialogPicker.SkipDomainControllerCheck = $true
  $DialogPicker.Providers = [CubicOrange.Windows.Forms.ActiveDirectory.ADsPathsProviders]::Default

  $DialogPicker.AttributesToFetch.Add('samAccountName')
  #$DialogPicker.AttributesToFetch.Add('title')
  #$DialogPicker.AttributesToFetch.Add('department')
  #$DialogPicker.AttributesToFetch.Add('distinguishedName')


  $DialogPicker.ShowDialog()

  return $DialogPicker.Selectedobject
}
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
Function Show-MessageBox {
    Param([string]$Message="This is a default Message.",
        [string]$Title="Default Title",
        [ValidateSet(“Asterisk”,”Error”,”Exclamation”,“Hand”,”Information”,”None”,“Question”,”Stop”,”Warning”)]
        [string]$Type="Error",
        [ValidateSet(“AbortRetryIgnore”,”OK”,”OKCancel”,“RetryCancel”,”YesNo”,”YesNoCancel”)]
        [string]$Buttons="OK"
    )
    [void][System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”)
    $MsgBoxResult = [System.Windows.Forms.MessageBox]::Show($Message,$Title,[Windows.Forms.MessageBoxButtons]::$Buttons,[Windows.Forms.MessageBoxIcon]::$Type)
    Return $MsgBoxResult
}
Function Install-RSATTools
{
    $VerbosePreference = 'Continue'

    $x86 = 'https://download.microsoft.com/download/1/D/8/1D8B5022-5477-4B9A-8104-6A71FF9D98AB/WindowsTH-RSAT_WS2016-x86.msu'
    $x64 = 'https://download.microsoft.com/download/1/D/8/1D8B5022-5477-4B9A-8104-6A71FF9D98AB/WindowsTH-RSAT_WS2016-x64.msu'

    switch ($env:PROCESSOR_ARCHITECTURE)
    {
        'x86' {$version = $x86}
        'AMD64' {$version = $x64}
    }

    Write-Verbose -Message "OS Version is $env:PROCESSOR_ARCHITECTURE"
    Write-Verbose -Message "Now Downloading RSAT Tools installer"

    $Filename = $version.Split('/')[-1]
    Invoke-WebRequest -Uri $version -UseBasicParsing -OutFile "$env:TEMP\$Filename" 
    
    Write-Verbose -Message "Starting the Windows Update Service to install the RSAT Tools "
    
    Start-Process -FilePath wusa.exe -ArgumentList "$env:TEMP\$Filename /quiet" -Wait -Verbose
    
    Write-Verbose -Message "RSAT Tools are now be installed"
    
    Remove-Item "$env:TEMP\$Filename" -Verbose
    
    Write-Verbose -Message "Script Cleanup complete"
    
    Write-Verbose -Message "Remote Administration"
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
Function Disable-Dayone
{
    Param(
        [string]$user
  )
  $Cred = Get-SBCredential -UserName "Arborpharma01\providynadmin"
  Add-Type -Assembly Microsoft.VisualBasic
  If ([Microsoft.VisualBasic.Information]::IsNumeric($user))
  {
    $sam = Get-ADUser -Filter {EmployeeID -eq $user} -Properties *
  }
  Else
  {
    $sam = Get-ADUser -Filter {Name -like $user} -Properties *
  }
  $sam | Get-ADPrincipalGroupMembership | ForEach-Object {
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
  $mesname = $sam.Name
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
  Show-MessageBox -Message "The active directory account for $mesname has undergone day 1 termination proceedings." -Title "Day 1 complete" -Type Information -Buttons OK
}
Function Disable-Daytwo
{
    Param(
        [string]$usert
        )
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
    Install-module AzureRM
    Install-module MSOnline
    Invoke-WebRequest 'https://download.microsoft.com/download/7/1/E/71EF1D05-A42C-4A1F-8162-96494B5E615C/msoidcli_64bit.msi' -OutFile "$env:USERPROFILE\Downloads\signin.msi"
    Install-Msi -Path "$env:USERPROFILE\Downloads\signin.msi"
    Invoke-WebRequest 'https://bposast.vo.msecnd.net/MSOPMW/Current/amd64/AdministrationConfig-en.msi' -OutFile "$env:USERPROFILE\Downloads\azure.msi"
    Install-Msi -Path "$env:USERPROFILE\Downloads\azure.msi"
    Invoke-WebRequest 'https://download.microsoft.com/download/0/2/E/02E7E5BA-2190-44A8-B407-BC73CA0D6B87/sharepointonlinemanagementshell_6112-1200_x64_en-us.msi' -OutFile "$env:USERPROFILE\Downloads\share.msi"
    Install-Msi -Path "$env:USERPROFILE\Downloads\share.msi"
    Invoke-WebRequest 'https://download.microsoft.com/download/2/0/5/2050B39B-4DA5-48E0-B768-583533B42C3B/SkypeOnlinePowershell.exe' -OutFile "$env:USERPROFILE\Downloads\skype.exe"
    Start-Process "$env:USERPROFILE\Downloads\skype.exe" -ArgumentList '/passive /norestart'
  }
  $cred2 = Get-SBCredential -UserName "mboxadmin@arborpharma.com"
  $connect = Connect-MsolService -Credential $cred2
  $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred2 -Authentication Basic -AllowRedirection
  Import-PSSession $Session
  Get-MsolDomain
    $priname = $usert.Replace(' ', '.') + '@arborpharma.com'
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
  Show-MessageBox -Message "The office 365 account for $usert has undergone day 2 termination proceedings." -Title "Day 2 complete" -Type Information -Buttons OK
}
Function Create-userad
{
  $Cred = Get-SBCredential -UserName "Arborpharma01\providynadmin"
  Do
  {
    $csv = Get-FileName
  }
  Until ($csv.Contains(".xlsx") -or $csv.Contains(".csv"))
  #$hold = '"' + "$csv" + '"'
  #$path = Split-Path $hold
  Start-Process 'C:\PwdReset\excelmod.exe' -ArgumentList "`"$csv`"" -Wait
  $v = Import-CSV -Delimiter ',' -Path "$env:APPDATA\store.csv"
  ForEach ($ob in $v)
  {
    $first = $ob."First Name"
    $last = $ob."Last Name"
    $dept = $ob."Default Department"
    $title = $ob."Default Jobs (HR)"
    $cell = $ob."Cell Phone"
    $phone = $ob."Home Phone"
    $email = $ob.Email
    $finit = $first.ToUpper().substring(0, 1)
    $linit = $last.ToUpper().substring(0, 1)
    $comp = $ob."Employee EIN"
    $id = $ob."Employee Id"
		
    If ($dept.StartsWith('Sales'))
    {
      New-ADUser -Title $title -Company $comp -DisplayName "$first $last" -Department $dept -EmailAddress "$first.$last@arborpharma.com" -GivenName "$first" -AccountPassword (ConvertTo-SecureString -AsPlainText "s0m3thing$finit$linit" -Force) -ChangePasswordAtLogon $false -EmployeeID $id -SamAccountName "$first.$last" -MobilePhone $cell -HomePhone $phone -Name "$first $last" -Office 'Field' -UserPrincipalName "$first.$last@arborpharma.com" -Surname $last -Path "ou=Employees,ou=Arbor FIELD,dc=ArborPharma,dc=local" -OtherAttributes @{ mailNickname = "$first.$last"; proxyAddresses = "SMTP:$first.$last@arborpharma.com" } -Enabled $true
      Add-ADGroupMember "PW SelfService" "$first.$last"
    }
    Else
    {
      New-ADUser -Title $title -Company $comp -DisplayName "$first $last" -Department $dept -EmailAddress "$first.$last@arborpharma.com" -GivenName "$first" -AccountPassword (ConvertTo-SecureString -AsPlainText "s0m3thing$finit$linit" -Force) -ChangePasswordAtLogon $false -EmployeeID $id -SamAccountName "$first.$last" -MobilePhone $cell -HomePhone $phone -Name "$first $last" -Office 'Atlanta Corporate' -UserPrincipalName "$first.$last@arborpharma.com" -Surname $last -Path "ou=Employees,ou=Arbor ATL,dc=ArborPharma,dc=local" -OtherAttributes @{ mailNickname = "$first.$last"; proxyAddresses = "SMTP:$first.$last@arborpharma.com" } -Enabled $true
    }
  }
  Show-MessageBox -Message "The active directory accounts for the users in the pre hire report have been created.  Please wait for the users to sync with office 365 before creating the mailboxes." -Title "Users Created" -Type Information -Buttons OK
}
Function Create-userO365
{
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
        Install-module AzureRM
        Install-module MSOnline
    Invoke-WebRequest 'https://download.microsoft.com/download/7/1/E/71EF1D05-A42C-4A1F-8162-96494B5E615C/msoidcli_64bit.msi' -OutFile "$env:USERPROFILE\Downloads\signin.msi"
    Install-Msi -Path "$env:USERPROFILE\Downloads\signin.msi"
    Invoke-WebRequest 'https://bposast.vo.msecnd.net/MSOPMW/Current/amd64/AdministrationConfig-en.msi' -OutFile "$env:USERPROFILE\Downloads\azure.msi"
    Install-Msi -Path "$env:USERPROFILE\Downloads\azure.msi"
    Invoke-WebRequest 'https://download.microsoft.com/download/0/2/E/02E7E5BA-2190-44A8-B407-BC73CA0D6B87/sharepointonlinemanagementshell_6112-1200_x64_en-us.msi' -OutFile "$env:USERPROFILE\Downloads\share.msi"
    Install-Msi -Path "$env:USERPROFILE\Downloads\share.msi"
    Invoke-WebRequest 'https://download.microsoft.com/download/2/0/5/2050B39B-4DA5-48E0-B768-583533B42C3B/SkypeOnlinePowershell.exe' -OutFile "$env:USERPROFILE\Downloads\skype.exe"
    Start-Process "$env:USERPROFILE\Downloads\skype.exe" -ArgumentList '/passive /norestart'
  }
  $cred2 = Get-SBCredential -UserName "mboxadmin@arborpharma.com"
  Do
  {
    $csv2 = Get-FileName
  }
  Until ($csv2.Contains(".xlsx") -or $csv2.Contains(".csv"))
  #$hold = '"' + "$csv" + '"'
  #$path = Split-Path $hold
  Start-Process 'C:\PwdReset\excelmod.exe' -ArgumentList "`"$csv2`"" -Wait
  $ob = Import-CSV -Delimiter ',' -Path "$env:APPDATA\store.csv"
  $connect = Connect-MsolService -Credential $cred2
  $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred2 -Authentication Basic -AllowRedirection
  Import-PSSession $Session
  foreach ($obj in $ob)
  {
    $dept = $obj."Default Department"
    $namee = $obj."First Name" + ' ' + $obj."Last Name"
    $msolname = $obj."First Name" + '.' + $obj."Last Name" + '@arborpharma.com'
    Set-MsolUser -UserPrincipalName $msolname -UsageLocation US
    Set-MsolUserLicense -UserPrincipalName $msolname -AddLicenses 'reseller-account:ENTERPRISEPACK'
    Enable-Mailbox -Identity $namee
    Add-DistributionGroupMember -Identity 'allemployeesdl-arborpharma.com' -Member $msolname -Confirm:$false
    If ($dept.StartsWith('Sales'))
    {
      Add-DistributionGroupMember -Identity 'Arbor Field Employees' -Member $msolname -Confirm:$false
    }
    Else
    {
      Add-DistributionGroupMember -Identity 'arboratldl-arborpharma.com' -Member $msolname -Confirm:$false
    }
    Enable-Mailbox -Identity $namee -Archive
  }
  Get-PSSession | Remove-PSSession
  Show-MessageBox -Message "The office 365 accounts for the users in the pre hire report have been created." -Title "Users Created" -Type Information -Buttons OK
}
Function Test-ADAuthentication {
    param($username,$password)
    Add-Type -AssemblyName System.DirectoryServices.AccountManagement
    $ct = [System.DirectoryServices.AccountManagement.ContextType]::Domain
    $pc = New-Object System.DirectoryServices.AccountManagement.PrincipalContext $ct,$env:USERDOMAIN
    $pc.ValidateCredentials($username,$password)}
Use-RunAs
Import-module ActiveDirectory
<#
    New-PSDrive `
    -Name "test" `
    -PSProvider "ActiveDirectory" `
    -server 10.18.8.28 `
    -Credential ($Cred) `
    -Root "" `
    -Scope Global
    cd test:
    cd "dc=arborpharma,dc=local"
    Get-ADDomain
    Search-User
#>
$version = $PSVersionTable.PSVersion.Major
$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$text = & "C:\PwdReset\NET_Detector_cli_.exe"
  $dotnet = $False
ForEach ($item in $text)
{
    If (($item -match ".NET Framework 4.5.2*") -or ($item -match ".NET Framework 4.6*") -or ($item -match ".NET Framework 4.7*"))
    {
        $dotnet = $True
    }
}
If (!($dotnet))
{
    $powerPrompt = [System.Windows.Forms.MessageBox]::Show("You did not read the README.  Go take a look at it.  You do not have powershell 5 or the appropriate .NET framework needed to run powershell 5 installed on your computer.")
    Exit
}
If ($version -lt 5)
{
  $powerPrompt = [System.Windows.Forms.MessageBox]::Show("You did not read the README. You do not have powershell 5 installed.  I will go ahead and do it for you.")
  iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
  & "C:\ProgramData\chocolatey\bin\RefreshEnv.cmd"
  choco install powershell -y
  Restart-Computer -Force
}
If (!(Test-Path "$env:ProgramFiles\WindowsPowershell\modules\carbon"))
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
}
If (!(Test-Path "$env:ProgramFiles\WindowsPowershell\modules\PowerShellCookbook"))
{
  Get-PackageProvider -Name NuGet -ForceBootstrap
  Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
  Try
  {
    Install-module PowerShellCookbook -Force -ErrorAction Stop
  }
  Catch
  {
    Install-module PowerShellCookbook -Force -AllowClobber
  }
}
If (!(Test-Path "$env:ProgramFiles\WindowsPowershell\modules\SendEmail"))
{
  Get-PackageProvider -Name NuGet -ForceBootstrap
  Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
  Install-Module SendEmail
}
Try
{
  Import-Module ActiveDirectory -erroraction stop
}
Catch
{
  Use-RunAs
  Install-RSATTools
}


  [void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
  [void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
  [void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
  [void][reflection.assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
  [void][reflection.assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
  [void][reflection.assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][reflection.assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
Add-VpnConnection -Name "ArborVPN" -ServerAddress remote.arborpharma.com -AllUserConnection -AuthenticationMethod MSChapv2 -EncryptionLevel Maximum -TunnelType Sstp -RememberCredential -Verbose -ErrorAction SilentlyContinue

  [System.Windows.Forms.Application]::EnableVisualStyles()
  $MainForm = New-Object 'System.Windows.Forms.Form'
  $labelPasswordReset = New-Object 'System.Windows.Forms.Label'
  $textbox2 = New-Object 'System.Windows.Forms.TextBox'
  $textbox1 = New-Object 'System.Windows.Forms.TextBox'
  $buttonUnlockUserAccount = New-Object 'System.Windows.Forms.Button'
  $buttonResetUserPassword = New-Object 'System.Windows.Forms.Button'
    $buttonCreateAD = New-Object 'System.Windows.Forms.Button'
  $buttonSearch = New-Object 'System.Windows.Forms.Button'
  $buttonCreateO365 = New-Object 'System.Windows.Forms.Button'
  $buttonDisableone = New-Object 'System.Windows.Forms.Button'
  $buttonDisableO365 = New-Object 'System.Windows.Forms.Button'
  $buttonConnectO365 = New-Object 'System.Windows.Forms.Button'
  $buttonDisConnectO365 = New-Object 'System.Windows.Forms.Button'
  $tooltip1 = New-Object 'System.Windows.Forms.ToolTip'
  $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'

	
  $OnLoadFormEvent={
	
  }

$ShowHelp={
    Switch ($this.name) {
        "buttonSearch"  {$tip = "Lets you directly search active directory for user"}
        "buttonUnlockUserAccount" {$tip = "Unlock user account for selected user"}
    "buttonResetUserPassword" { $tip = "Reset password for selected user" }
    "buttonCreateAD" { $tip = "Create AD user accounts from CSV file" }
    "buttonCreateO365" { $tip = "Create Office 365 mailbox accounts from CSV file" }
    "buttonDisableone" { $tip = "Perform Day 1 termination proceedings for selected user" }
    "buttonDisableO365" { $tip = "Perform Day 2 termination proceedings for selected user"}
    "buttonConnectO365" { $tip = "Simply connect to Office 365 service" }
    "buttonDisConnectO365" { $tip = "Simply disconnect from Office 365 service" }
  }
    $tooltip1.SetToolTip($this,$tip)
}

    $buttonSearch_Click=
    {
        $textbox2.Text = $(Search-user).FetchedAttributes
    }
	
  $buttonResetUserPassword_Click=
    {
    $user = $textbox2.Text
      if ([string]::IsNullOrEmpty($user) -eq $false)            
            {
                Function Set-AdUserPwd
                { 
                Param( 
                [string]$user,
                [string]$pwd 
                ) #end param 
                $strFilter = "(&(objectCategory=User)(sAMAccountName=$user))"  
                $objDomain = New-Object System.DirectoryServices.DirectoryEntry 
                $objSearcher = New-Object System.DirectoryServices.DirectorySearcher 
                $objSearcher.SearchRoot = $objDomain 
                $objSearcher.PageSize = 1000 
                $objSearcher.Filter = $strFilter 
                $userLDAP = $objSearcher.FindOne() | select-object -ExpandProperty Path 
                if ($userLDAP.Length -gt 0)
                    {
                        $oUser = [adsi]"$userLDAP"
                        $setADUserPwdmsgbox = [System.Windows.Forms.MessageBox]::Show("You have selected $userLDAP. Is this correct?","",4)
                        if ($setADUserPwdmsgbox -eq "YES" ) 
                            {
          Get-ADUser -Filter { SamACcountName -like $user } -Credential $Cred -ErrorAction SilentlyContinue | Set-ADAccountPassword -NewPassword (ConvertTo-SecureString -AsPlainText $pwd -Force) -Reset -Credential $Cred -ErrorAction SilentlyContinue -Confirm:$false
        }
                        else
                            {
                            }
                    }
                    else 
                    {
                    [System.Windows.Forms.MessageBox]::Show("This username does not exist. Please try again.")
                    }
      #>
                }

                # CALL FUNCTION
                if ($NEWPWD.Length -gt 0)
                {
                $Reset_Error = $null
                Set-ADUserPwd -user $user -pwd $NEWPWD
                if ((Get-ADUser -Filter {SamACcountName -like $user} -Properties PasswordLastSet -ErrorVariable Reset_Error -ErrorAction SilentlyContinue -Credential $Cred | Select PasswordLastSet -ExpandProperty PasswordLastSet) -gt (Get-Date).AddMinutes(-1))
                    {
                    [System.Windows.Forms.MessageBox]::Show("$user's password has been reset to $NEWPWD.")
                    }
                else
                    {
                    if ($Reset_Error.Length -gt 0)
                        {
                            [System.Windows.Forms.MessageBox]::Show("There was an error using Active Directory. Are you using an account with proper privileges with RSAT installed?")
                        }
                    [System.Windows.Forms.MessageBox]::Show("Reset aborted.")
                    }
                }
                else
                {
                    [System.Windows.Forms.MessageBox]::Show("ERROR! The default_reset_pwd.txt file is missing.")
                }
            }
            
            else
            {
                [System.Windows.Forms.MessageBox]::Show("The username field is empty.")
            }
    }


  $buttonUnlockUserAccount_Click=
{
  $user = $textbox2.Text
  if ([string]::IsNullOrEmpty($user) -eq $false)
  {
    Function Unlock-ADUser
    {
      Param (
        [string]$user
      ) #end param 
			
      $strFilter = "(&(objectCategory=User)(sAMAccountName=$user))"
      $objDomain = New-Object System.DirectoryServices.DirectoryEntry
      $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
      $objSearcher.SearchRoot = $objDomain
      $objSearcher.PageSize = 1000
      $objSearcher.Filter = $strFilter
      $userLDAP = $objSearcher.FindOne() | select-object -ExpandProperty Path
      if ($userLDAP.Length -gt 0)
      {
        $oUser = [adsi]"$userLDAP"
        $setADUserPwdmsgbox = [System.Windows.Forms.MessageBox]::Show("You have selected $userLDAP. Is this correct?", "", 4)
        if ($setADUserPwdmsgbox -eq "YES")
        {
          Get-ADUser -Filter { SamACcountName -like $user } -Credential $Cred -ErrorAction SilentlyContinue | Unlock-ADAccount -Credential $Cred -ErrorAction SilentlyContinue
          #$ouser.psbase.invokeset("AccountDisabled","False") 
          #$ouser.psbase.CommitChanges()
        }
        else
        {
        }
      }
      else
      {
        [System.Windows.Forms.MessageBox]::Show("This username does not exist. Please try again.")
      }
    }
    # CALL FUNCTION
    $Unlock_Error = $null
    if ((Get-ADUser -Filter { SamACcountName -like $user } -Properties LockedOut -ErrorVariable Unlock_Error -ErrorAction SilentlyContinue | Select LockedOut -ExpandProperty LockedOut) -eq $False)
    {
      [System.Windows.Forms.MessageBox]::Show("$user is already unlocked.")
    }
    else
    {
      Unlock-ADUser -user $user
      if ((Get-ADUser -Filter { SamACcountName -like $user } -Properties LockedOut -ErrorVariable Unlock_Error -ErrorAction SilentlyContinue | Select LockedOut -ExpandProperty LockedOut) -eq $False)
      {
        [System.Windows.Forms.MessageBox]::Show("$user has been unlocked.")
      }
      else
      {
        if ($Unlock_Error.Length -gt 0)
        {
          [System.Windows.Forms.MessageBox]::Show("There was an error using Active Directory. Are you using an account with proper privileges with RSAT installed?")
        }
        [System.Windows.Forms.MessageBox]::Show("Unlock aborted.")
      }
			
    }
  }
  else
  {
    [System.Windows.Forms.MessageBox]::Show("The username field is empty.")
  }
}

  $buttonCreateAD_Click=
  {
    Create-userad
  }

  $buttonCreateO365_Click=
  {
    Create-userO365
  }

  $buttonDisableone_Click=
  {
    Disable-Dayone "$($textbox2.Text)"
  }

  $buttonDisableO365_Click=
  {
    Disable-Daytwo "$($textbox2.Text)"
  }
  
  $buttonConnectO365_Click=
  {
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
        Install-module AzureRM
        Install-module MSOnline
    Invoke-WebRequest 'https://download.microsoft.com/download/7/1/E/71EF1D05-A42C-4A1F-8162-96494B5E615C/msoidcli_64bit.msi' -OutFile "$env:USERPROFILE\Downloads\signin.msi"
    Install-Msi -Path "$env:USERPROFILE\Downloads\signin.msi"
    Invoke-WebRequest 'https://bposast.vo.msecnd.net/MSOPMW/Current/amd64/AdministrationConfig-en.msi' -OutFile "$env:USERPROFILE\Downloads\azure.msi"
    Install-Msi -Path "$env:USERPROFILE\Downloads\azure.msi"
    Invoke-WebRequest 'https://download.microsoft.com/download/0/2/E/02E7E5BA-2190-44A8-B407-BC73CA0D6B87/sharepointonlinemanagementshell_6112-1200_x64_en-us.msi' -OutFile "$env:USERPROFILE\Downloads\share.msi"
    Install-Msi -Path "$env:USERPROFILE\Downloads\share.msi"
    Invoke-WebRequest 'https://download.microsoft.com/download/2/0/5/2050B39B-4DA5-48E0-B768-583533B42C3B/SkypeOnlinePowershell.exe' -OutFile "$env:USERPROFILE\Downloads\skype.exe"
    Start-Process "$env:USERPROFILE\Downloads\skype.exe" -ArgumentList '/passive /norestart'
  }
  $cred2 = Get-SBCredential -UserName "mboxadmin@arborpharma.com"
  $connect = Connect-MsolService -Credential $cred2
  $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred2 -Authentication Basic -AllowRedirection
  Import-PSSession $Session
}

$buttonDisConnectO365_Click=
{
    Get-PSSession | Remove-PSSession
}
    $Form_StateCorrection_Load=
  {
    $MainForm.WindowState = $InitialFormWindowState
  }
	
  $Form_Cleanup_FormClosed=
  {

    try
    {
            $buttonSearch.remove_Click($buttonSearch_Click)
      $buttonUnlockUserAccount.remove_Click($buttonUnlockUserAccount_Click)
      $buttonResetUserPassword.remove_Click($buttonResetUserPassword_Click)
      $buttonCreateAD.remove_Click($buttonCreateAD_Click)
      $buttonCreateO365.remove_Click($buttonCreateO365_Click)
      $buttonDisableone.remove_Click($buttonDisableone_Click)
      $buttonDisableO365.remove_Click($buttonDisableO365_Click)
      $buttonConnectO365.remove_Click($buttonConnectO365_Click)
      $buttonDisConnectO365.remove_Click($buttonDisConnectO365_Click)
            $buttonSearch.remove_MouseHover($ShowHelp)
            $buttonUnlockUserAccount.remove_MouseHover($ShowHelp)
            $buttonResetUserPassword.remove_MouseHover($ShowHelp)
      $buttonCreateAD.remove_MouseHover($ShowHelp)
      $buttonCreateO365.remove_MouseHover($ShowHelp)
      $buttonDisableone.remove_MouseHover($ShowHelp)
      $buttonConnectO365.remove_MouseHover($ShowHelp)
      $buttonDisConnectO365.remove_MouseHover($ShowHelp)
      $MainForm.remove_Load($OnLoadFormEvent)
      $MainForm.remove_Load($Form_StateCorrection_Load)
      $MainForm.remove_FormClosed($Form_Cleanup_FormClosed)
    }
    catch [Exception]
    { }
  }

  $MainForm.Controls.Add($labelPasswordReset)
  $MainForm.Controls.Add($textbox2)
  $MainForm.Controls.Add($textbox1)
    $MainForm.Controls.Add($buttonSearch)
  $MainForm.Controls.Add($buttonUnlockUserAccount)
  $MainForm.Controls.Add($buttonResetUserPassword)
  $MainForm.Controls.Add($buttonCreateAD)
  $MainForm.Controls.Add($buttonCreateO365)
  $MainForm.Controls.Add($buttonDisableone)
  $MainForm.Controls.Add($buttonDisableO365)
  $MainForm.Controls.Add($buttonConnectO365)
  $MainForm.Controls.Add($buttonDisConnectO365)
    $MainForm.ClientSize = '380, 280'
  $MainForm.Name = "MainForm"
  $MainForm.StartPosition = 'CenterScreen'
  $MainForm.Text = "Arbor AD assistant"
  $MainForm.add_Load($OnLoadFormEvent)

  $labelPasswordReset.Font = "Tahoma, 9.75pt, style=Bold"
  $labelPasswordReset.Location = '3, 6'
  $labelPasswordReset.Name = "labelPasswordReset"
  $labelPasswordReset.Size = '330, 14'
  $labelPasswordReset.TabIndex = 6
  $labelPasswordReset.Text = "Arbor AD tool"
  $labelPasswordReset.TextAlign = 'TopCenter'

  $textbox2.Location = '61, 23'
  $textbox2.Name = "textbox2"
  $textbox2.Size = '275, 20'
  $textbox2.TabIndex = 8
  $textbox2.Enabled = $True
  #$textbox2.ReadOnly = $true

  $textbox1.BackColor = 'ControlLightLight'
  $textbox1.Enabled = $False
  $textbox1.Location = '4, 23'
  $textbox1.Name = "textbox1"
  $textbox1.ReadOnly = $True
  $textbox1.Size = '61, 20'
  $textbox1.TabIndex = 7
  $textbox1.Text = "Username: "

    $buttonSearch.Font = "Tahoma, 8pt"
    $buttonSearch.Location = '338, 23'
    $buttonSearch.Name = "buttonSearch"
    $buttonSearch.Size = '40, 20'
    $buttonSearch.TabIndex = 10
    $buttonSearch.Text = "...."
    $buttonSearch.UseVisualStyleBackColor = $True
    $buttonSearch.add_MouseHover($ShowHelp)
    $buttonSearch.add_Click($buttonSearch_Click)

  $buttonUnlockUserAccount.Font = "Tahoma, 8pt"
  $buttonUnlockUserAccount.Location = '170, 49'
  $buttonUnlockUserAccount.Name = "buttonUnlockUserAccount"
  $buttonUnlockUserAccount.Size = '165, 22'
  $buttonUnlockUserAccount.TabIndex = 11
  $buttonUnlockUserAccount.Text = "Unlock User Account"
  $buttonUnlockUserAccount.UseVisualStyleBackColor = $True
    $buttonUnlockUserAccount.add_MouseHover($ShowHelp)
  $buttonUnlockUserAccount.add_Click($buttonUnlockUserAccount_Click)

  $buttonResetUserPassword.Font = "Tahoma, 8pt"
  $buttonResetUserPassword.Location = '4, 49'
  $buttonResetUserPassword.Name = "buttonResetUserPassword"
  $buttonResetUserPassword.Size = '165, 22'
  $buttonResetUserPassword.TabIndex = 9
  $buttonResetUserPassword.Text = "Reset User Password"
  $buttonResetUserPassword.UseVisualStyleBackColor = $True
    $buttonResetUserPassword.add_MouseHover($ShowHelp)
  $buttonResetUserPassword.add_Click($buttonResetUserPassword_Click)

    $buttonCreateAD.Font = "Tahoma, 8pt"
    $buttonCreateAD.Location = '170, 75'
    $buttonCreateAD.Name = "buttonCreateAD"
    $buttonCreateAD.Size = "165, 22"
  $buttonCreateAD.TabIndex = 13
    $buttonCreateAD.Text = "Create users in AD with CSV"
    $buttonCreateAD.Enabled = $True
    $buttonCreateAD.UseVisualStyleBackColor = $True
    $buttonCreateAD.add_MouseHover($ShowHelp)
  $buttonCreateAD.add_Click($buttonCreateAD_Click)

  $buttonCreateO365.Font = "Tahoma, 8pt"
  $buttonCreateO365.Location = '4, 75'
  $buttonCreateO365.Name = "buttonCreateO365"
  $buttonCreateO365.Size = '165, 22'
  $buttonCreateO365.TabIndex = 14
  $buttonCreateO365.Text = "Create office 365 accounts"
  $buttonCreateO365.UseVisualStyleBackColor = $True
  $buttonCreateO365.add_MouseHover($ShowHelp)
  $buttonCreateO365.add_Click($buttonCreateO365_Click)

  $buttonDisableone.Font = "Tahoma, 8pt"
  $buttonDisableone.Location = '4, 100'
  $buttonDisableone.Name = "buttonDisableone"
  $buttonDisableone.Size = '165, 22'
  $buttonDisableone.TabIndex = 15
  $buttonDisableone.Text = "Perform day 1 termination"
  $buttonDisableone.UseVisualStyleBackColor = $True
  $buttonDisableone.add_MouseHover($ShowHelp)
  $buttonDisableone.add_Click($buttonDisableone_Click)


  $buttonDisableO365.Font = "Tahoma, 8pt"
  $buttonDisableO365.Location = '175, 100'
  $buttonDisableO365.Name = "buttonDisableO365"
  $buttonDisableO365.Size = '165, 22'
  $buttonDisableO365.TabIndex = 16
  $buttonDisableO365.Text = "Perform day 2 termination"
  $buttonDisableO365.UseVisualStyleBackColor = $True
  $buttonDisableO365.add_MouseHover($ShowHelp)
  $buttonDisableO365.add_Click($buttonDisableO365_Click)

  $buttonConnectO365.Font = "Tahoma, 8pt"
  $buttonConnectO365.Location = '4, 130'
  $buttonConnectO365.Name = "buttonConnectO365"
  $buttonConnectO365.Size = '165, 22'
  $buttonConnectO365.TabIndex = 17
  $buttonConnectO365.Text = "Connect to Office 365"
  $buttonConnectO365.UseVisualStyleBackColor = $True
  $buttonConnectO365.add_MouseHover($ShowHelp)
  $buttonConnectO365.add_Click($buttonConnectO365_Click)

  $buttonDisConnectO365.Font = "Tahoma, 8pt"
  $buttonDisConnectO365.Location = '175, 130'
  $buttonDisConnectO365.Name = "buttonDisConnectO365"
  $buttonDisConnectO365.Size = '165, 22'
  $buttonDisConnectO365.TabIndex = 18
  $buttonDisConnectO365.Text = "Disconnect from office 365"
  $buttonDisConnectO365.UseVisualStyleBackColor = $True
  $buttonDisConnectO365.add_MouseHover($ShowHelp)
  $buttonDisConnectO365.add_Click($buttonDisConnectO365_Click)


  $InitialFormWindowState = $MainForm.WindowState

  $MainForm.add_Load($Form_StateCorrection_Load)

  $MainForm.add_FormClosed($Form_Cleanup_FormClosed)
  return $MainForm.ShowDialog()
