#Returns first SoftwareVersion from the list if installed, if not installed returns empty string
function Get-SoftwareVersion{
    param(
    #SoftwareName
    [Parameter(Mandatory=$true)]
    [String]$SoftwareName
    )

    #Registry uninstall paths in which check is performed
    $Paths = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*","HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"

    #For each path perform checks
    foreach($Paths in $Paths){

    #Gets displayversion from registry key of name simillar to $Softwarename and saves to $Version
    $Version = (Get-ItemProperty $Paths | select DisplayName,DisplayVersion | where {$_.DisplayName -like "*$SoftwareName*"}).DisplayVersion
    
    #If $Version string is empty write to host and return 1
    if([string]::IsNullOrEmpty($Version) -eq $true){Write-Host "Software with '$SoftwareName' in name is not installed";continue}

    #If $Version string is not empty continue
    elseif([string]::IsNullOrEmpty($Version) -eq $false){
        
        #If $Version is an array of strings return first string from $Version array
        if($Version.Count -ne 1){return $Version[0]}

        #Else return $Version
        else{return $Version}
    }

    #Else throw an error
    else{$Version = "Error";throw "Error when checking software version"}
    }
}