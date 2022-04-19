$Computers = import-csv -Path "C:\Scripts\Inventory\csv\inventory.csv"
$Script:StatusList = New-Object System.Collections.ArrayList

Write-Host "`r`n`r`nPowershell is collecting the data. Please wait for script to finish; do not close this window.`r`n`r`n" -BackgroundColor DarkRed

foreach ($Computer in $Computers) {
    If (Test-Connection -BufferSize 32 -Count 1 -ComputerName $Computer.COMPUTER_NAME -Quiet) {
        Write-Host $Computer.COMPUTER_NAME ": ONLINE" -ForegroundColor Green
        $Script:StatusList += "ONLINE"
    } else {
        Write-Host $Computer.COMPUTER_NAME ": OFFLINE" -ForegroundColor Red
        $Script:StatusList += "OFFLINE"
    }
}

$Script:StatusList | Out-File -FilePath "C:\Scripts\Inventory\csv\status.csv"