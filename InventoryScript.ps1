$Computers = import-csv -Path "C:\Scripts\Inventory\csv\inventory.csv"
$Script:BoxSNList = New-Object System.Collections.ArrayList
$Script:HDDList = New-Object System.Collections.ArrayList

Write-Host "`r`n`r`nPowershell is collecting the inventory data. Please wait for script to finish; do not close this window.`r`n`r`n" -BackgroundColor DarkRed

foreach ($Computer in $Computers) {
    If (Test-Connection -BufferSize 32 -Count 1 -ComputerName $Computer.COMPUTER_NAME -Quiet) {
        Try {
            $HDDSerial = (Get-WmiObject Win32_PhysicalMedia -ComputerName $Computer.COMPUTER_NAME | ? {$_.Tag -match "PHYSICALDRIVE0"}).SerialNumber -replace " ", ""
            $ComputerSerial = (Get-WmiObject Win32_BIOS -ComputerName $Computer.COMPUTER_NAME).SerialNumber -replace " ", ""
            Write-Host $Computer.COMPUTER_NAME $ComputerSerial $HDDSerial
            $Script:BoxSNList += $ComputerSerial
            $Script:HDDList += $HDDSerial
        } Catch {
            Write-Host $Computer.COMPUTER_NAME "encountered an error." -ForegroundColor Red
            $Script:BoxSNList += "ERROR"
            $Script:HDDList += "ERROR"
        }
    } else {
        Write-Host $Computer.COMPUTER_NAME "is offline." -ForegroundColor Red
        $Script:BoxSNList += "OFFLINE"
        $Script:HDDList += "OFFLINE"
    }
}

$Script:BoxSNList | Out-File -FilePath "C:\Scripts\Inventory\csv\boxsnlist.csv"
$Script:HDDList | Out-File -FilePath "C:\Scripts\Inventory\csv\hddlist.csv"