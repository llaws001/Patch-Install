# Run this script / ISE as administrator!
$wshell = New-Object -ComObject Wscript.shell

#start the update service if it is stopped
$WinUpdate = Get-Service wuauserv
if($WinUpdate.Status -like "Stopped"){
    Get-Service wuauserv
    Start-Service wuauserv
    write-host "Starting Windows Update Service"
    start-sleep -s 8
    
}
Get-Service wuauserv


# Old hotfix list
Get-HotFix > "$PSScriptRoot\old_hotfix_list.txt"

# Get all updates
$Updates = Get-ChildItem -Path $PSScriptRoot -Recurse | Where-Object {$_.Name -like "*msu*"}


$CheckDirectory = Get-ChildItem -Path $PSScriptRoot -recurse -filter *.msu | Measure-Object
#Checks to see if directory has a patch to install
if($CheckDirectory.Count -eq 0){
 Write-Host "There are not any patches in the directory"
 $wshell.Popup("Remember to stop the windows update service!")
 } else {
# Iterate through each update
    ForEach ($update in $Updates) {

        
        $UpdateFilePath = $update.FullName
        write-host "Installing update $($update.BaseName)"

        Start-Process -wait wusa -ArgumentList "/update $UpdateFilePath","/quiet","/norestart" 
}
    # New hotfix list
    Get-HotFix > "$PSScriptRoot\new_hotfix_list.txt"
    write-host "The update process is complete. Please review the hotfix (a hotfix is a patch) logs to verify whether or not the patches successfully installed." 
    $wshell.Popup("Remember to stop the windows update service after reboot!")
    start-sleep -s 15

    #ATTENTION COMMENT THIS OUT IF YOU DO NOT WANT YOUR MACHINE TO REBOOT
    Restart-Computer
}




 
