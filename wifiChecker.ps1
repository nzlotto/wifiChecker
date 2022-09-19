#$computerName = "127.0.0.1"


$computerName = read-host "Enter the computer name or IP address"
$directoryPath = "C:\_WS\" + $computerName
$fileLocation = "\\" + $computerName + "\c$\ProgramData\Microsoft\Windows\WlanReport\*"
$saveLocation = $directoryPath
$removeLocation = "\\" + $computerName + "\c$\ProgramData\Microsoft\Windows\WlanReport\"
$launch = $saveLocation + "\wlan-report-latest.html"


if (Test-Connection -computerName $computerName -Quiet){
    Write-Host "$computerName is online" -ForegroundColor Green
    try {
        Invoke-Command -ComputerName $computerName -ScriptBlock {netsh wlan show wlanreport}
        New-Item -ItemType directory -Path $directoryPath
        Copy-Item -Path "$fileLocation" -Destination "$saveLocation"
        invoke-item "$launch"
        Remove-Item -Path  "$removeLocation" -Recurse
        }
        catch [System.Management.Automation.RuntimeException]
        {
         "$computerName is online but can't connect"
         }
         }
    else {
        Write-Host "$computerName is not available.  Please try again later." -ForegroundColor Red
       }