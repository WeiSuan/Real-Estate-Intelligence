param(
    [string]$HostAddress = "127.0.0.1",
    [int]$Port = 8772
)

$ErrorActionPreference = "Stop"

try {
    $connections = Get-NetTCPConnection `
        -LocalAddress $HostAddress `
        -LocalPort $Port `
        -State Listen `
        -ErrorAction SilentlyContinue

    $serverPids = @($connections | Select-Object -ExpandProperty OwningProcess -Unique)

    foreach ($serverPid in $serverPids) {
        try {
            Stop-Process -Id $serverPid -Force -ErrorAction Stop
            Write-Host "已關閉舊服務 PID $serverPid"
        } catch {
            Write-Host "無法關閉舊服務 PID $serverPid：$($_.Exception.Message)"
            exit 1
        }
    }

    exit 0
} catch {
    Write-Host "檢查 $HostAddress`:$Port 失敗：$($_.Exception.Message)"
    exit 1
}
