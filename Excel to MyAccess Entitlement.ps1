#This script is to automate the entilement provisioning in MyAccess of AD groups. See the instructions in One note "Active Directory Atlassian provisioning automatic scripts"
#Version 1.2 - 4.10.2023 by Kaius Koivisto
#ChangeLog -- Made the endpoint creation optional and the script to work with only one owner
#Dependencies: Powershell module: ImportExcel  Install with command 'install-module importexcel' needs administrator session in powershell

#loop to create csv files in sequence
$ready = $false
while (-not $ready) {

    #import dependecy module
    Import-Module ImportExcel

    #file chooser window
    write-host "Please choose the excel file with the proper format"
    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "Select an Excel File"
    $openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $excelFilePath = $openFileDialog.FileName
        # Output the selected file path
        Write-Host "Selected File: $excelFilePath"
    } else {
        # The user canceled the dialog
        Write-Host -prompt "File selection canceled."
        exit
    }

    #The needed information should be on the Sheet1 of the excel
    $worksheetName = "Sheet1"

    # Read the Excel file and ready the variable with all the emails
    $data = Import-Excel -Path $excelFilePath -WorksheetName $worksheetName
    $commands = "Security System,Endpoint,Entitlement Type,Entitlement Value,Owner,Rank`r`n"
    $mails = @()

    # Iterate through each row of data and create AD groups
    foreach ($row in $data) {
        $groupName = $row.'Group name (max 64)'
        if ($row.'Owners (employee nbr + email)') {$owners = $row.'Owners (employee nbr + email)'} else {$owners = $row.'Owners (email)'} 
        $owners = if ($owners.contains(",")) {$owners.toLower().replace(" ", "").Split(",")} else {$owners.toLower().replace(" ", "")}
        $securitysystem = $row.'extensionAttribute1'
        $mails += $owners
        foreach ($owner in $owners) {
            $actdir = "Active Directory,Active Directory,memberOf,`"CN=$groupName,OU=Groups,OU=HEL,OU=Vaisala_Resources,DC=corp,DC=vaisala,DC=com`",$owner,1`r`n"
            $sysHEL =if ($securitysystem) {"Active Directory,$securitysystem,memberOf,`"CN=$groupName,OU=Groups,OU=HEL,OU=Vaisala_Resources,DC=corp,DC=vaisala,DC=com`",$owner,1`r`n"} else {""}
            $commands = $commands + $actdir + $sysHEL
        }
    }
    
    #User needs to manually give the MyAccess username of all the mails in the excel
    write-host "Now go to MyAccess and translate each mail to corresponding username eg `"12345`""
    $mails = $mails | Select-Object -Unique
    foreach ($mail in $mails) {
        Set-Clipboard -Value $mail
        $username = read-host "What is the username of "$mail" "
        $commands =$commands.Replace($mail, $username)
    }
    write-host "All mails replaced with usernames"

    #file can be found in the same place that the script is run from
    $filepath = "autoEntitlement.csv"

    if (test-path -path $filepath) {
        Remove-item -Path $filepath}
    $commands | Out-File -FilePath $filepath -Encoding UTF8 -noNewline
    
    read-host "CSV file can be found where this powershell is running $filepath, after using the file Press enter to continue"

    #remove the used csv file after being uploaded to myaccess
    Remove-item -Path $filepath 
    
    #want to loop?
    $answ = read-host "Do you want to create another entitlement .CSV  y, n"
    if ($answ -eq "n") {$ready = $true}
}

read-host "All finnished"