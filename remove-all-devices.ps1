$devices=Get-PnpDevice
ForEach ($device in $devices){
	&"pnputil" /remove-device $device.InstanceId
}
shutdown -r -t 1