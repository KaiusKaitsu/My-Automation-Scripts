#This script is to automate the deletion of computers from Active directory and SCCM. See the instructions in One note "Active directory group creation automatic scripts"
#Version 1.1 - 2.11.2023 
#Added: Logs to keep track of computers deleted by the script
#Dependencies configurationManager.psd1 and ps module activedirectory

$ready = $false 

    while (-not $ready) {

    $computername = Read-Host -Prompt "Enter computer to delete from AD and SCCM: "
    #Find the computer object from ad
    $computer = Get-ADComputer -Identity $computername
    #Search for the msFVE-RecoveryInformation (BitLocker) object of the computer. This is a childobject of the computer
    $bitlocker = get-ADObject -Filter {objectclass -eq "msFVE-RecoveryInformation"} -SearchBase $computer.DistinguishedName
    # Delete the childobject of the computer. Else the computer cannot be deleted recursively with certain rights
    remove-AdObject $bitlocker.DistinguishedName -Confirm:$false
    #Delete the computer object itself
    remove-AdObject $computer -Confirm:$false


    #Import powershell module for SCCM (you'll need the whole path)
    Import-Module X:\...\ConfigurationManager.psd1
    #Create instance of the SCCM server
    New-PSDrive -Name "SCCM-server" -PSProvider "[]" -Root "sccm.mycompany.com" -Description "Primary site"
    #CD to right folder
    Set-Location SCCM-server:
    #Remove the computer device
    Remove-CMDevice -DeviceName $computername -Confirm:$false
    #return to root
    Set-Location X:\
    #remove the created SCCM-server instance
    Remove-PSDrive -Name "SCCM-server" 

    #Write logs for the deletion action
    $logpath = 'X:\...\PC deletion logs.txt'
    $time = Get-Date
    $log = $computername + " deleted at: " + $time.tostring() + " by "+ $env:username +"`r`n"
    $log | Add-Content -path $logpath -Encoding UTF8
    Write-output "$computername removed from AD and SCCM`r`nLogs can be found from $logpath"   
    
    #Repeat with new computer?
    $yn = Read-Host "do you want to delete another computer?  y, n"
    if ($yn -eq "y") {$ready = $false} else {$ready = $true}
    }

Read-Host -Prompt "all done, press enter to exit"