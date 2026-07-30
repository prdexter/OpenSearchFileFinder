@echo off
title OpenSearch File Finder Launcher
echo ==================================================
echo Starting OpenSearch Engine and Smart Finder...
echo ==================================================

:: 1. Check and Start Docker Desktop if not already running
tasklist /FI "IMAGENAME eq Docker Desktop.exe" 2>NUL | find /I /N "Docker Desktop.exe" >NUL
if "%ERRORLEVEL%"=="1" (
    echo [*] Starting Docker Desktop...
    start "" "C:\Program Files\Docker\Docker\Docker Desktop.exe"
    echo [*] Waiting 12 seconds for Docker engine initialization...
    ping -n 12 127.0.0.1 >nul
) else (
    echo [+] Docker Desktop is already running.
)

:: 2. Spin up OpenSearch Docker containers
cd /d "D:\Active research\OpenSearch"
echo [*] Starting OpenSearch Docker containers...
docker compose up -d

:: 3. Start Smart Search App server in background if not running
start "SmartSearchApp" /min python search_app.py

:: 4. Wait for services to initialize
ping -n 3 127.0.0.1 >nul

:: 5. Open http://localhost:8080 in Microsoft Edge
echo [+] Opening OpenSearch File Finder in Microsoft Edge...
start msedge "http://localhost:8080"

exit
