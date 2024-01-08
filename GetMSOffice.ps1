function GetMSOffice{
    param(
        #Passing $Search, used when $Search is an array and passing index one by one
        [Parameter(Mandatory=$true)]
        $Search
    )

    #Version
    #gets full path to ospp.vbs
    $File = $Search.DirectoryName + "\" + $Search.PSChildName
    
    #runs status check on MS Office suite license and writes is as an string
    $OSPP = cscript $File /dstatus | Out-String
    $OSPP

    #Type
    if($OSPP.contains("No installed product keys detected")){break}
    elseif($OSPP.Contains("Skype")){
        break
    }
    elseif($OSPP.Contains("Professional")){
        $Type = "Professional"
    }
    elseif($OSPP.Contains("Standard")){
        $Type = "Standard"
    }
    elseif($OSPP.Contains("Business")){
        $Type = "Business"
    }
    else{$Type = "Error unknown Type";Write-Host "Error when checking which Type of office is installed"}

    #If type could not be determined break the function, else continue
    if($Type.Contains("Error")){$Office = $type;break}
    else{
        #depending which Office *Number* is found on $OSPP do 
        if($OSPP.Contains("Office 11")){$Office = "MS Office 2003"}
        elseif($OSPP.contains("Office 12")){$Office = "MS Office 2007"}
        elseif($OSPP.contains("Office 14")){$Office = "MS Office 2010"}
        elseif($OSPP.contains("Office 15")){$Office = "MS Office 2013"}
        elseif($OSPP.contains("365")){$Office = "MS Office 365"}

        #There is no difference if Office 16 for 2016/2019 (Also 365 but it says 365 for that version so we can exclude that it)
        elseif($OSPP.contains("Office 16")){$Office = "MS Office 2016/2019"

            #gets installed packages on PC that contains Microsoft Office, writes them in wide format and as an string
            $Installed = Get-Package -Name "Microsoft Office*" | Format-Wide | Out-String

            #If either 2019 or 2016 is present do something
            if($Installed.Contains("$Type 2019")){
                $Office = "MS Office 2019"
            }
            elseif($Installed.Contains("$Type 2016")){
                $Office = "MS Office 2016"
            }
            else{$Office = "Error 2016/2019 Version";Write-Host "Error when checking if 2016/2019 Version of office is installed"}
        }
        else{$Office = "Error Version";Write-Host "Error when checking what Version of office is installed"}
        $Office = $Office + " $Type"
    }

    #Bitness
    if($Office -inotmatch "Error"){
        if($Search.DirectoryName.Contains("x86") -eq $true){$Bitness = 32}
        elseif($Search.DirectoryName.Contains("Program Files\") -eq $true){$Bitness = 64}
        else{Write-Host "Office Suite was not found"; $Bitness = 1}
        $Office = $Office + " x$Bitness"
    }

    $Office
}

$Error.Clear()

#Searches for licensing script for MS Office Suite
$Search = Get-ChildItem -Path 'C:\Program Files\Microsoft Office','C:\Program Files (x86)\Microsoft Office' -Recurse -Include ospp.vbs -ErrorAction SilentlyContinue

#If none scripts found MS Suite not installed
if($Error -ne 0 -and ($Error[0] | Out-String).Contains("because it does not exist") -eq $true){$Office = "N/A";$Office}

#If one found do
elseif($Search.Count -eq 1){
    $Office = GetMSOffice -Search $Search
}

#If more than one found do
elseif($Search.Count -ge 1){
    $Count = $Search.count - 1
    for($i = 0;$i -le $Count;$i++){
        $Office = GetMSOffice -Search $Search[$i]
        if($i -eq 0){$Offices = $Office}
        elseif($i -le $Count){$Offices = $Offices + " + " + $Office}
    }
}

#If something else throw error
else{$Office = "Error";throw "Error when checking MS Office"}

if($Offices){$Office = $Offices}