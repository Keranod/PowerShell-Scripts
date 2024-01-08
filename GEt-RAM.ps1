$RAM = Get-CimInstance Win32_PhysicalMemory

#Amount of physiacal RAM modules
$Count = $RAM.Count

#If there is one RAM module count is going to be null, but you still can call $RAM[0]
if([String]::IsNullOrEmpty($count)){$Count = 1}

if($Count -eq 0){
    Write-Host "No Physical RAM Modules"
    $RAM = "Error"
}

else{
    #Total Capacity of physical RAM
    $Capacity = ($RAM | Measure-Object -Property capacity -Sum).sum /1gb

    #DIMM or SODIMM
    $FormFactor = $RAM[0].FormFactor
    if($FormFactor -eq 8){$FormFactor = "DIMM"}
    elseif($FormFactor -eq 11){$FormFactor = "RIMM"}
    elseif($FormFactor -eq 12){$FormFactor = "SODIMM"}
    else{$FormFactor = "Unknown Form Factor"}

    <#
    0 = Unknown
1 = Other
2 = SIP
3 = DIP
4 = ZIP
5 = SOJ
6 = Proprietary
7 = SIMM
8 = DIMM
9 = TSOP
10 = PGA
11 = RIMM
12 = SODIMM
13 = SRIMM
14 = SMD
15 = SSMP
16 = QFP
17 = TQFP
18 = SOIC
19 = LCC
20 = PLCC
21 = DDR2
22 = FPBGA
23 = LGA
    #>

    #DDR Type
    $DDR = $RAM[0].SMBIOSMemoryType
    if($DDR -eq 26){$DDR = "DDR4"}
    elseif($DDR -eq 24){$DDR = "DDR3"}
    elseif($DDR -eq 22){$DDR = "DDR2 FB-DIMM"}
    elseif($DDR -eq 21){$DDR = "DDR2"}
    elseif($DDR -eq 20){$DRR = "DDR"}
    elseif($DDR -eq 19){$DRR = "RDRAM"}
    else{$DDR = "Unknown DDR Type"}


    for($i=0;$i -lt $Count;$i++){
        #capacity of ram module in GB
        [array]$ModuleCapacity += $RAM[$i].Capacity /1gb
        
        #speed of ram module in MHz
        [array]$Speed += $RAM[$i].Speed
    }

    $ModuleCapacityResult = $null
    #if all modules are not the same capacity
    if(@($ModuleCapacity -ne $ModuleCapacity[0]).Count -gt 0){
        for($i=0;$i -lt $Count;$i++){
            $ModuleCapacityResult = $ModuleCapacityResult + "$i=" + $ModuleCapacity[$i] + "GB;"
        }
    }
    else{$ModuleCapacityResult = [string]$ModuleCapacity[0] + "GB per Module"}

    $SpeedResult = $null
    #if all speeds of modules are not the same
    if(@($Speed -ne $Speed[0]).Count -gt 0){
        for($i=0;$i -lt $Count;$i++){
            $SpeedResult = $SpeedResult + "$i=" + $Speed[$i] + "MHz;" 
        }
    }
    else{$SpeedResult = [string]$Speed[0] + "MHz per Module" }

    $RAM = "$Capacity"+ "GB Total|$Count Modules|$FormFactor|$DDR|$ModuleCapacityResult|$SpeedResult"
}

$RAM

Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *