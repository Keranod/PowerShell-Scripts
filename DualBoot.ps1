$BootMode = bcdedit | Select-String "path.*efi"
if ($null -eq $BootMode) {
    # I think non-uefi is \Windows\System32\winload.exe
    $BootMode = "Legacy"
}else {
    # UEFI is: 
    #path                    \EFI\MICROSOFT\BOOT\BOOTMGFW.EFI
    #path                    \Windows\system32\winload.efi
    $BootMode = "UEFI"
}
bcdedit /timeout 5 #reduces timeout in dualboot menu to 5 seconds
bcdedit /create "{ramdiskoptions}" /d '"RamdiskOptions"'
bcdedit /set "{ramdiskoptions}" ramdisksdidevice partition=\device\harddiskvolume3
bcdedit /set "{ramdiskoptions}" ramdisksdipath \WPEx64\media\Boot\boot.sdi
bcdedit /create "{11111111-2222-3333-4444-555555555555}" /d '"WinPE"' -application osloader
bcdedit /set "{11111111-2222-3333-4444-555555555555}" device ramdisk=[\device\harddiskvolume3]\WPEx64\media\sources\boot.wim','"{ramdiskoptions}"

If($BootMode -eq 'UEFI'){
bcdedit /set "{11111111-2222-3333-4444-555555555555}" path \windows\system32\boot\winload.efi
}
elseif($BootMode -eq 'Legacy'){
bcdedit /set "{11111111-2222-3333-4444-555555555555}" path \windows\system32\boot\winload.exe
}

bcdedit /set "{11111111-2222-3333-4444-555555555555}" locale en-US
bcdedit /set "{11111111-2222-3333-4444-555555555555}" inherit "{bootloadersettings}"
bcdedit /set "{11111111-2222-3333-4444-555555555555}" osdevice ramdisk=[\device\harddiskvolume3]\WPEx64\media\sources\boot.wim','"{ramdiskoptions}"
bcdedit /set "{11111111-2222-3333-4444-555555555555}" systemroot \windows
bcdedit /set "{11111111-2222-3333-4444-555555555555}" nx AlwaysOff
bcdedit /set "{11111111-2222-3333-4444-555555555555}" pae ForceDisable
bcdedit /set "{11111111-2222-3333-4444-555555555555}" detecthal Yes
bcdedit /set "{11111111-2222-3333-4444-555555555555}" winpe Yes
bcdedit /set "{11111111-2222-3333-4444-555555555555}" ems Yes
BCDEdit /displayorder "{11111111-2222-3333-4444-555555555555}" /addlast
