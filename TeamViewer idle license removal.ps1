$token = "YOUR-API-KEY-HERE" 
				#Needed rights (With these it works, might not need all of these experiment if you like!): 
				#Computers & Contacts:  View entries, add entries, edit entries, remove entries
				#User groups: Create user groups, read user groups, edit user groups, delete user groups

$bearer = "Bearer $token"

$header = @{
    "Authorization" = $bearer
}

$devices = (Invoke-RestMethod -Uri "https://webapi.teamviewer.com/api/v1/devices" -Method Get -Headers $header).devices

$Days = -60  #Put the wanted idle time for deletion, 60 = 2 months. Remember to mark it with a minus sign!

$date = (Get-Date).AddDays($Days)

$Counter = 0

Read-Host -Prompt "This script will delete computers that have been idle for $Days days. 

Press enter to continue"

foreach ($device in $devices) {
    if ($device.online_state -eq "Offline") {
        $lastSeen = $device.last_seen

        if ($lastSeen) {
            $lastSeenDate = $lastSeen.Split("T")[0]
            $dateLastSeen = [datetime]$lastSeenDate
	    $id = $device.device_id

            if ($dateLastSeen -le $date) {
                 Invoke-WebRequest -Uri "Https://webapi.teamviewer.com/api/v1/devices/$id" -Method Delete -Headers $header -UseBasicParsing
                 Write-Host "Deleted device:"$device.alias -ForegroundColor Yellow
		 $Counter ++
            }
        }
    }
}
"Deleted $Counter computers"
Read-Host -Prompt "Press Enter to exit"