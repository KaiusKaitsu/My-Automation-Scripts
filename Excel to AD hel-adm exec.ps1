#This script is to automate the creation of groups to Active directory.
#Version 1.4 - 2.11.2023 by Kaius Koivisto
#Changes in 1.4: 
#       Added fields to specify groupScope, groupCategory and office site
#       Is backwards compatible for the older excel format without specifications for the mentioned new fields
#       Fixed the invisible newline/whitespace bug when there is newlines inside the members or owners cells
#Dependencies: Powershell module: ImportExcel  Install with command 'install-module importexcel' needs administrator session in powershell


# Connect to server hosting active directory Asks for your admin credentials to create a connection to server
$serverName = "server"
write-host "Connecting to " + $serverName
$credential = Get-Credential
$session = New-PSSession -ComputerName $serverName -Credential $credential
$ready = $false

#this is the loop that allows to create ad groups from multiple excel files in the same session
while (-not $ready) {

    #import the dependencie importexcel to the session
    Import-Module ImportExcel

    #file chooser dialog window
    write-host "Please choose the excel file with the proper format"
    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "Select an Excel File"
    $openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $excelFilePath = $openFileDialog.FileName
        Write-Host "Selected File: $excelFilePath"
    } else {
        Write-Host -prompt "File selection canceled."
        exit
    }

    #The needed information should be on the Sheet1 of the excel also the ticket number of the request
    $worksheetName = "Sheet1"
    $ticketNumber = read-host -Prompt "Give the ticket number of the request. This number is placed to the info of the group"


    # Read the Excel file and ready the data 
    $time =  Get-Date
    $data = Import-Excel -Path $excelFilePath -WorksheetName $worksheetName
    $commands = "`r`n`r`n`r`n`r`n`r`n`r`nimport-module activedirectory`r`n#This is temporary file for creating atlassian cloud ad groups`r`n" +'#'+ $time.tostring() +" : New group(s) file created by " + $env:username +" for ticket $ticketNumber`r`n#Here are the new commands that have been run:`r`n`r`n"


    # Iterate through each row of data and write the AD group creation commands
    foreach ($row in $data) {
        $groupName = $row.'Group name (max 64)'.replace(" ","")
        $displayName = $row.'Display name (max 256)' 
        $description = $row.'Description (max 1024)'
        $extensionAttribute1 = $row.extensionAttribute1
        $extensionAttribute5 = $row.extensionAttribute5
        $extensionAttribute2 = $row.extensionAttribute2
        if($row.'Owners (employee nbr + email)') {$owners = $row.'Owners (employee nbr + email)' -replace "`n", "" -replace "`r", ""} else {$owners = $row.'Owners (email)' -replace "`n", "" -replace "`r", ""}
        $member = $row.'Members (email)' -replace "`n", "" -replace "`r", ""
        $GroupCategory = if ($row.groupCategory) {$row.groupCategory} else {"Security"}
        $GroupScope = if ($row.groupScope) {$row.groupScope} else {"Global"}
        $path = if ($row.site) {"OU=Groups,OU="+$row.site+",OU=Resources,DC=corp,DC=company,DC=com"} else {"OU=Groups,OU=FIN,OU=Resources,DC=corp,DC=company,DC=com"}

        #cleans the owners from whitespace and takes the first owner as the AD group manager
        if ($owners.contains(",")) {
            $firstowner, $rest = $owners.replace(" ","") -split ","
        } else {$firstowner = $owners.replace(" ","")}
        
        #Cleans members and creates the creation string of members
        $members = ""
        if ($member) {
            if ($member.contains(",")) {
                $member = $member.replace(",","`",`"").replace(" ","")
            }
             $members = "`r`n"+ '$members = "' + $member + "`""
        }

        #Creates the extensionattribute string 1, 2 & 5 are supported
        $optionalAttributes = ""
        if ($extensionAttribute5) {
            $optionalAttributes = $optionalAttributes + '"extensionAttribute5" = "' + $extensionAttribute5 + '"; '
        }
        
        if ($extensionAttribute2) {
            $optionalAttributes = $optionalAttributes + '"extensionAttribute2" = "' + $extensionAttribute2 + '"; '
        }

        if ($extensionAttribute1) {
            $optionalAttributes = $optionalAttributes + '"extensionAttribute1" = "' + $extensionAttribute1 + '"; '
        }
        
        #These are the active directory creation commands that run on Server
        $findowner = "`r`n`$owner = Get-ADUser -Filter {EmailAddress -eq `"$firstowner`"}"
        $findMembers = if ($members) {"`r`n"+'$members2 = $members | ForEach-Object {$user = Get-ADUser -Filter {EmailAddress -eq $_ }'+"`r`n"+'if ($user) {$user} else {Write-Host "Email $_ not found"}}'} else {""}
        $newGroup = "`r`n"+'New-ADGroup -Name "' + $groupName + '" -SamAccountName ' + $groupName + ' -GroupCategory ' + $GroupCategory + ' -GroupScope ' + $GroupScope + ' -DisplayName "' + $displayName + '" -Path "' + $path + '" -Description "' + $description + '" -ManagedBy "' + '$owner' + '" -OtherAttributes @{'+ $optionalAttributes +'  "info" = "#' + $ticketNumber + '"}'
        $addMembers = if ($members) {"`r`nAdd-ADGroupMember -Identity $groupName -Members `$members2"} else {""}
        $commands = $commands + $members + $findMembers + $findowner + $newGroup + $addMembers
    }

    #file path of the created ps1 file that is to be executed on the Server remote session also the path for the log text file 
    $filepath = 'X:\...\ADgroupScript.ps1'
    $logpath = 'X:\...\provisioning logs.txt'

    #log just copies the creation script file on server and appends it to the dedicated log file
    $logCreator = "`r`n`r`n`r`nGet-Content -path '$filepath' | Add-Content -path '$logpath' -Encoding UTF8"
    
    $commands = $commands + $logCreator + "`r`nread-host `"ready, press enter to exit`"`r`n`r`n`r`n"

    #delete any previous file and write the new ps1 file
    Invoke-Command -Session $session -ScriptBlock {
        if (test-path -path $Using:filepath) {
            Remove-item -Path $Using:filepath}
        $Using:commands | Out-File -FilePath $Using:filepath -Encoding UTF8 -noNewline
    }
    read-host "Login to Server and execute $filepath powershell script 
    then press enter to continue"

    #delete the uneeded file
    Invoke-Command -Session $session -ScriptBlock {
        Remove-item -Path $Using:filepath
    }
    write-host "Logs can be found from $logpath"
    #do all again or exit loop? case sensitive y 
    $answ = read-host "Do you want to create another AD group?  y, n"
    if ($answ -eq "y") {$ready = $false} else {$ready = $true}
}

# Remove the remote session
Remove-PSSession -Session $sessionwrite-host "Connecting to $serverName"

read-host 'AD loop exited press enter to sync AD and azure'

#Automatically run AD to Azure AD sync after creating the groups. Connects to Azure with the previous credentials and runs the sync
write-host 'Running AD - Azure AD Sync'

$syncsession = new-pssession Azure-server -credential $credential
Invoke-Command -session $syncsession -ScriptBlock {
    start-adsyncsynccycle -policytype delta
}
#disconnects from Azure
remove-pssession -Session $syncsession

read-host "Sync done, press enter to exit"
