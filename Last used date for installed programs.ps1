Get-CimInstance -Class win32_softwareFeature |?{$_.lastuse -like "*"} | select productname,lastuse | Sort-Object lastuse -unique #excludes empty lastuse 

Get-CimInstance -Class win32_softwareFeature | Select-Object productname,lastuse -unique | Format-Table -Autosize # lists everything installed