#need to run powershell as local administrator, built in one
$Monitors = Get-WmiObject WmiMonitorID -Namespace root\wmi
ForEach ($Monitor in $Monitors)
{
	$Manufacturer = ($Monitor.ManufacturerName -notmatch 0 | ForEach{[char]$_}) -join ""
	$Name = ($Monitor.UserFriendlyName -notmatch 0 | ForEach{[char]$_}) -join ""
	$Serial = ($Monitor.SerialNumberID -notmatch 0 | ForEach{[char]$_}) -join ""}
$Manufacturer
$Name
$Serial