<#
 SCRIPT NAME: AppInstall_CiscoVMOPlugin.ps1
     VERSION: 1.2
        DATE: 6/7/2018
      AUTHOR: Justin Jahns @vdiguywi

 DISCRIPTION: Installs Cisco ViewMail for Outlook Plugin v11.5.1 and then configures user settings. Must run script as user.
DEPENDENCIES: none
  PARAMETERS: none
         USE: AppInstall_CiscoVMOPlugin.ps1
#>
$strLDAPSearchBase = "LDAP://domain.forest.com/DC=domain,DC=forest,DC=com"
$strUCServer = "serverfqdn.domain.forest.com"

#Install Cisco VMO Plugin
if ((Test-Path "C:\Program Files\Cisco Systems\ViewMail for Outlook\VMOAddIn.dll") -eq $false)
    {
    cmd /c "`"$PSScriptRoot\setup.exe`" /i /qb /LogFile `"$env:TEMP\CiscoViewMailforOutlook1151Install.log`""
    }

#Find current user in AD
$objUserSearchDomainRoot = New-Object System.DirectoryServices.DirectoryEntry("$strLDAPSearchBase")
$objUserSearch = New-Object System.DirectoryServices.DirectorySearcher
$objUserSearch.SearchRoot = $objUserSearchDomainRoot
$objUserSearch.PageSize = 200000
$objUserSearch.Filter = "(&(objectCategory=user)(Name=$env:USERNAME))"
$objUserSearch.SearchScope = "Subtree"
$objUserSearchResult = $objUserSearch.FindOne()
$objUser = $objUserSearchResult.GetDirectoryEntry()

#Get user's email and phone from AD
$strUserEmail = $objUser.EmailAddress
$strUserPhone = $objUser.ipPhone

#Create configuration REG Keys and Values
$strRegRoot = "HKCU:\SOFTWARE\Cisco Systems\ViewMail for Outlook"
if(!(Test-Path "$strRegRoot")) {New-Item -Path "$strRegRoot" -Force | Out-Null}
New-ItemProperty -Path "$strRegRoot" -Name "Dummy" -Value "DUMMY" -PropertyType String -Force | Out-Null

$strRegLanguage = "$strRegRoot\Language"
if(!(Test-Path "$strRegLanguage")) {New-Item -Path "$strRegLanguage" -Force | Out-Null}
New-ItemProperty -Path "$strRegLanguage" -Name "SelectedLanguage" -Value "en-US" -PropertyType String -Force | Out-Null

$strRegLogging = "$strRegRoot\Logging"
if(!(Test-Path "$strRegLogging")) {New-Item -Path "$strRegLogging" -Force | Out-Null}
New-ItemProperty -Path "$strRegLogging" -Name "LogLevel" -Value "Error" -PropertyType String -Force | Out-Null

$strRegServers = "$strRegRoot\Profiles\Outlook\Servers"
if(!(Test-Path "$strRegServers")) {New-Item -Path "$strRegServers" -Force | Out-Null}
New-ItemProperty -Path "$strRegServers" -Name "DefaultServer" -Value "$strUserEmail" -PropertyType String -Force | Out-Null

$strRegProfile = "$strRegRoot\Profiles\Outlook\Servers\$strUserEmail"
if(!(Test-Path "$strRegProfile")) {New-Item -Path "$strRegProfile" -Force | Out-Null}
New-ItemProperty -Path "$strRegProfile" -Name "Host" -Value "$strUCServer" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$strRegProfile" -Name "Phone" -Value "$strUserPhone" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$strRegProfile" -Name "PlaybackDevice" -Value "Telephone" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$strRegProfile" -Name "RecordDevice" -Value "Telephone" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$strRegProfile" -Name "RecordingCodec" -Value "7" -PropertyType DWORD -Force | Out-Null
New-ItemProperty -Path "$strRegProfile" -Name "SearchFolderVersion" -Value "3" -PropertyType DWORD -Force | Out-Null
New-ItemProperty -Path "$strRegProfile" -Name "Type" -Value "Connection8.5 SIB" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$strRegProfile" -Name "ViewMailCategory" -Value "ViewMail" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$strRegProfile" -Name "VoiceMailRequiresUserNameAndPassword" -Value "0" -PropertyType DWORD -Force | Out-Null
New-ItemProperty -Path "$strRegProfile" -Name "VoiceMailSSOStrictMode" -Value "0" -PropertyType DWORD -Force | Out-Null

$strRegWizard = "$strRegRoot\Profiles\Outlook\Wizard"
if(!(Test-Path "$strRegWizard")) {New-Item -Path "$strRegWizard" -Force | Out-Null}
New-ItemProperty -Path "$strRegWizard" -Name "HasRun" -Value "1" -PropertyType DWORD -Force | Out-Null

$strRegVMHotKey = "$strRegRoot\VoiceMailHotKey"
if(!(Test-Path "$strRegVMHotKey")) {New-Item -Path "$strRegVMHotKey" -Force | Out-Null}
New-ItemProperty -Path "$strRegVMHotKey" -Name "VoiceMailHotKey" -Value "" -PropertyType String -Force | Out-Null

$strRegVMNotify = "$strRegRoot\VoiceMailIcon_Notification"
if(!(Test-Path "$strRegVMNotify")) {New-Item -Path "$strRegVMNotify" -Force | Out-Null}
New-ItemProperty -Path "$strRegVMNotify" -Name "VoiceMailIcon_Notification" -Value "0" -PropertyType DWORD -Force | Out-Null

$strRegVMUploadPlay = "$strRegRoot\VoiceMailUpload_Play"
if(!(Test-Path "$strRegVMUploadPlay")) {New-Item -Path "$strRegVMUploadPlay" -Force | Out-Null}
New-ItemProperty -Path "$strRegVMUploadPlay" -Name "VoiceMailUpload_Play" -Value "0" -PropertyType DWORD -Force | Out-Null
