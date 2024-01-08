#Log name
$logName = 'Microsoft-Windows-DHCP-Client/Operational'

#New log config object
$log = New-Object System.Diagnostics.Eventing.Reader.EventLogConfiguration $logName

#If enabled return 0, if not enable and return 0, else return 1
if($log.IsEnabled -eq $true){return 0}
elseif($log.IsEnabled -eq $false){
    $log.IsEnabled = $true
    $log.SaveChanges()
    return 0;
}
else{return 1}
