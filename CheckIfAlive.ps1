$PCs = Get-Content -Path "PCs.txt" #variable containing list of PCs

$Count = $PCs.Count

for($i=0;$i -lt $count;$i++){
    If(Test-Connection –BufferSize 32 –Count 1 –quiet –ComputerName $PCs[$i]){$PCs[$i]}
}
