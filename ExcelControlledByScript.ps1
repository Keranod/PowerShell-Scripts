<# Create and use with a shortcut with target set to:
 C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -WindowStyle hidden -ExecutionPolicy Bypass "Path\ExcelControlledByScript.ps1" -spreadsheetPath "spreadsheet file path"
#>

param(
    #Path to the spreadsheet that is going to be controlled
    [Parameter(Mandatory=$true)]
    [String]$spreadsheetPath
)

$excel = New-Object -comobject Excel.Application
$excel.Visible = $true
$workbook = $excel.Workbooks.Open($spreadsheetPath)

while ($true) {
    Start-Sleep -Seconds 30
    try {
        $workbook.Save()
    }
    catch {break}
}

#Clears variables and modules
Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *