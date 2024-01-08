$Begin = Get-Date -Date '23/01/2023 00:00:00'
$End = Get-Date -Date '16/03/2023 12:00:00'
Get-EventLog -LogName Security -InstanceId 4801,4778 -After $Begin -Before $End #| format-table -wrap 
#Get-EventLog -LogName System -InstanceId 7001 -After $Begin -Before $End | format-table -wrap