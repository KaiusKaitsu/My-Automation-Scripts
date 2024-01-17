#This script unistalls and reinstalls SCCM client on the computer is being run on with a restart. Merged from multiple cmd and batch files to streamline SCCM reinstallation process.
#Version 1.0 - 10.1.2024 by Kaius koivisto
#Requires -RunAsAdministrator

 

#Stop SCCM & WMI related services

Write-Host "Shutting down services"

Set-Service winmgmt -StartupType Disabled

Stop-Service winmgmt -Force

Stop-Service smstsmgr -Force

Stop-Service CcmExec -Force

#sccm reinstall registry key to be run once after the unistallation and restart.
reg add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce /f /v SCCM /d "C:\windows\ccmsetup\ccmetup.exe SMSSITECODE=[CODE] SMSMP=[YOUR DOMAIN] FSP=[YOUR DOMAIN] SMSCACHESIZE=[NUMBER]"


#unistall sccm

write-host "Unistalling sccm this may take a while"

Start-Process C:\windows\ccmsetup\ccmsetup.exe /unistall -NoNewWindow -wait

#check that the unistallation is succesful by checking logs 'C:\windows\ccmsetup\logs\ccmsetup.log' 

Write-Host "Removing leftover folders and registry keys"

 

#remove any leftover folders

Remove-Item -Recurse -Force 'C:\windows\CCM',  'C:\windows\ccmcache', 'C:\windows\SMSCFG.INI'

 
Write-host "Do not worry about all the errors, they are expected"
 

#Remove registry values

reg delete 'HKLM\Software\Microsoft\CCM' /f

reg delete 'HKLM\Software\Microsoft\CCMSetup' /f

reg delete 'HKLM\Software\Microsoft\SMS' /f 

Write-Host "Running WMI fixes"


#WMI Fixes: Translated from wmi.bat and WMI repository fix <Start>

#Stop-Process "winmget" -Force -ErrorAction SilentlyContinue

Set-Location "$env:windir\system32\wbem" 

#Make a backup folder of repository with the name rep_bak

if (Test-Path "Repository") {

    Remove-Item -Path "Rep_bak" -Recurse -Force -ErrorAction SilentlyContinue

    Rename-Item -Path "Repository" -NewName "Rep_bak"

}

#Refister DLL files

Get-ChildItem -Filter *.dll | ForEach-Object { regsvr32 /s $_.FullName }
 

#Register exe files

Get-ChildItem -Filter *.exe | ForEach-Object {

    if ($_ -notin @("wbemcntl.exe", "wbemtest.exe", "mofcomp.exe", "wmiprvse.exe", "winmgmt.exe")) {

        Start-Process -FilePath $_.FullName -ArgumentList "/RegServer" -Wait
    }
}
 

#mofcomp fix

Get-ChildItem -Filter *.mof | ForEach-Object { mofcomp $_.FullName }

Get-ChildItem -Filter *.mfl | ForEach-Object { mofcomp $_.FullName }
 

#Start WMI service

Set-Service winmgmt -StartupType Automatic

Start-Service winmgmt

#</stop>

 
#ccmsetup.exe /unistall should remove these wmi components, but this is a check

$ccm = Get-WmiObject -namespace root -class __Namespace -filter "name = 'ccm'"

$sms = Get-WmiObject -namespace root/cimv2 -class __Namespace -filter "name = 'sms'"

if ($ccm) {Remove-WmiObject $ccm}

if ($sms) {Remove-WmiObject $sms}


Write-Host "Running mofcomp fix"

#mofcomp fix

mofcomp 'C:\Program Files\Microsoft Policy Platform\ExtendedStatus.mof'

Read-Host "Requires a shut down, press enter to reboot"

shutdown /r /t 15
