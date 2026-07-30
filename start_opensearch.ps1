# PowerShell script to start OpenSearch cluster and Dashboards
$OpenSearchDir = "D:\Active research\OpenSearch"

Write-Host "==================================================" -ForegroundColor Cipher
Write-Host "Starting OpenSearch Cluster & Dashboards..." -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cipher

Set-Location $OpenSearchDir

docker compose up -d

Write-Host "`nWaiting for OpenSearch service to initialize..." -ForegroundColor Yellow
Start-Sleep -Seconds 10

try {
    $response = Invoke-RestMethod -Uri "http://localhost:9200" -Method Get -ErrorAction Stop
    Write-Host "[+] OpenSearch is online! Cluster Name: $($response.cluster_name)" -ForegroundColor Green
    Write-Host "[+] OpenSearch Dashboards Web UI: http://localhost:5601" -ForegroundColor Green
} catch {
    Write-Host "[!] OpenSearch is still starting up. Please check http://localhost:9200 in a few seconds." -ForegroundColor Yellow
}
