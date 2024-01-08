function GetDiskInfo{
    param(
    #Disk
    [Parameter(Mandatory=$true)]
    $Disk
    )

    #Model of Disk
    $Model = $Disk.Model
    
    #if disk is a system drive
    if($Disk.BootFromDisk){$SystemDrive = "System Drive"}

    #size of Disk in GB rounded down
    $Size = [math]::Floor($Disk.Size /1gb)

    #GPT or MBR
    $PartitionStyle = $Disk.PartitionStyle

    return $DiskInfo = $Model + "|$Size GB" + "|$PartitionStyle" + "|$SystemDrive" 
}

#Get Disks
$Disks = Get-Disk

$Count = $Disks.Count

$Result = $null

#If more than one disk
if($Count -gt 1){
    for($i=0;$i -lt $Count;$i++){
        $Disk = GetDiskInfo -Disk $Disks[$i]
        $Result = $Result + $Disk + "`n"
    }
}


else{
    $Result = GetDiskInfo -Disk $Disks
}

$Result