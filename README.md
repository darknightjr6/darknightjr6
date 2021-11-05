# Notes for jcmontel
[Refer to this site for Markdown notes](https://www.markdownguide.org/basic-syntax)
[PowerShell and Configuration Manager 2012](https://devblogs.microsoft.com/scripting/powershell-and-configuration-manager-2012-r2part-1/)

## CMD notes
local dirctory variable
````
%~dp0
````

## MECM notes
Run powershell.exe command
````
powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -File .\%PSScript%.ps1
powershell.exe -noprofile -executionpolicy bypass -file ./Setup-KMS.ps1
````
Clear the ccmcache remotely
````powershell
$resman= New-Object -ComObject "UIResource.UIResourceMgr"
$cacheInfo=$resman.GetCacheInfo()
$cacheinfo.GetCacheElements()  | foreach {$cacheInfo.DeleteCacheElement($_.CacheElementID)
````
Increase the ccmcache size
````powershell
$CCM = New-Object -com UIResource.UIResourceMGR
$CC = $CCM.GetCacheInfo()
$CC.TotalSize = 8000
````
Add User to Remote Desktop Users ...This is not needed if machines are under TestOU3
````powershell
Add-LocalGroupMember -Group "Administrators" -Member "ASURITE\jcmontel"
Add-LocalGroupMember -Group "Remote Desktop Users" -Member "ASURITE\temp5100"
````
## Detection Method using gcim
Use either -like or -contains comparison operator
````powershell
if ((Get-CimInstance -Namespace "root\cimv2\sms" -Query "select * from sms_installedsoftware").ProductName `
    -contains 'Microsoft Office 365 ProPlus - en-us') { Write-Host "Installed" } else { }
# using get-childitem
if ((Get-ChildItem -Path "${env:ProgramFiles(x86)}\Zoom\bin\zoom.exe" -ErrorAction SilentlyContinue)) { "installed" } else { }
````

## Determine OS Info
Determine OS architecture
````powershell
$osArch = (Get-CimInstance -ClassName Win32_OperatingSystem).OSArchitecture
If ($osArch -eq '32-bit') {"32-bit"}
If ($osArch -eq '64-bit') {"64-bit"}
````
Determine OS Version
````powershell
$osVersion = (Get-CimInstance -ClassName Win32_OperatingSystem).Version
if ($osVersion -eq '10.0.18362') {$true} else {$false}
# other notes
[Environment]::OSVersion
([System.Environment]::OSVersion.Version).Build
````
[Determine PC Model](https://myserverissick.com/2013/09/find-a-computers-model-using-powershell/)
````powershell
(Get-WmiObject -Class:Win32_ComputerSystem).Model
````
## Uninstaller using gcim
````powershell
# GCIM using uninstallString
Get-CimInstance -Namespace "root\cimv2\sms" -Query "select * from sms_installedsoftware" |
    Where-Object {$_.ProductName -like "*deployment*"} | Select-Object -Property ProductName, UninstallString
# add/remove program name
if ((Get-CimInstance -ClassName Win32Reg_AddRemovePrograms).DisplayName `
    -like 'CrowdStrike*') {Write-Host "Installed"} else {}
if ((Get-CimInstance -ClassName Win32Reg_AddRemovePrograms).DisplayName `
            -notcontains 'System Center Endpoint Protection')
        {Write-Host "NotInstalled"} else {}
#using GWMI and ProductVerions
$gwmiProgram = Get-CimInstance -Namespace "root\cimv2\sms" -Query "select * from sms_installedsoftware"
if ($gwmiProgram | Where-Object { ($_.ProductName -contains 'SurfaceBook Update 18_050_00 (64 bit)') `
    -and ($_.ProductVersion -eq '18.050.18305.0')})
    {Write-Host "Installed" } else {  }
# shorthand
(Get-CimInstance -Namespace "root\cimv2\sms" -Query "select * from sms_installedsoftware").ProductName -like 'Microsoft Project Professional*'
# productCode
if (get-WmiObject Win32_Product | where -property IdentifyingNumber -EQ "{5E8556EC-62EA-49F0-919A-5C3C99E0C13A}"){write "installed"} else { }
# for Surface Firmware
if (get-wmiobject -query "SELECT * FROM Win32_ComputerSystem" -namespace "root\cimv2" | select model | where {$_.model -contains 'Surface Pro 4'})
{if (gwmi -query "select * from win32_pnpsigneddriver" | select deviceName, driverVersion | where {($_.deviceName -contains 'Surface Embedded Controller Firmware') -and ($_.driverVersion -eq '103.1744.256.0')} ) {Write-Host "Installed"} else { }}  else { }
````

Courtesy of jjourney ..this code is still a work in progress
````powershell
Get-CimInstance -query "select * from sms_installedsoftware" -namespace "root\cimv2\sms" |
	Where-Object {$_.ProductName -like "*365*"} |
Select-Object -Property ProductName, UninstallString

$uninstallstring = $app.UninstallString
$array1 = $uninstallstring.split('"', 3)
$array2 = $array1[2].split(" ")
$arrayFinal = $array2[1..($array2.Length - 1)]
````

## Install HyperV
Enable Hyper-V on Windows 10 using Powershell
````powershell
# Install only the PowerShell module
Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V-Management-PowerShell

# Install the Hyper-V management tool pack (Hyper-V Manager and the Hyper-V PowerShell module)
Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V-Tools-All

# Install the entire Hyper-V stack (hypervisor, services, and tools)
Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V-All
````

## Configure Server
[Configure Windows Server using Powershell](https://docs.microsoft.com/en-us/windows-server/administration/server-core/server-core-administer)
````powershell
# change password
$PWord = ConvertTo-SecureString -String 'P@sSwOrd1' -AsPlainText -Force
Set-LocalUser -name "administrator" -password $PWord

# rename the computer
rename-computer -NewName "RTS-CMG"

# enable PSRemoting
# Get-NetFirewallRule -Name 'WINRM*' | Select-Object Name #review the firewall rules
# Enable-PSRemoting -SkipNetworkProfileCheck -Force
# Set-NetFirewallRule -Name 'WINRM-HTTP-In-TCP' -RemoteAddress Any

# bind to domain
add-computer -DomainName 'asurite.ad.asu.edu' -Restart
````

## Get list of local users and create a folder for each user
````powershell
$epmFolder = '\appdata\local\EPMOfficeClient\'
        ForEach ($localUser In (GWmi Win32_UserProfile -F "Special != True" | Select -Expand LocalPath)) {
        $Null = New-Item -ItemType Directory -Force -Path "$($localUser)$epmFolder"
        Copy-File -path "$dirSupportFiles\*.*" -Destination "$($localUser)$epmFolder"
        }
````
