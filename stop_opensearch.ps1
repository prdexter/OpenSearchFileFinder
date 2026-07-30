# PowerShell script to stop OpenSearch cluster
$OpenSearchDir = "D:\Active research\OpenSearch"

Write-Host "Stopping OpenSearch containers..." -ForegroundColor Yellow
Set-Location $OpenSearchDir
docker compose down
Write-Host "[+] OpenSearch stopped successfully." -ForegroundColor Green
