if (Get-Process -processName "EXCEL" -ErrorAction SilentlyContinue) {
    Write-Host "Killing Excel"
    Stop-Process -ProcessName "EXCEL"
} else {
    Write-host "Excel is not open"
}

if (Get-Process -processName "WINWORD" -ErrorAction SilentlyContinue) {
    Write-Host "Killing Word"
    Stop-Process -ProcessName "WINWORD"
} else {
    Write-host "Word is not open"
}