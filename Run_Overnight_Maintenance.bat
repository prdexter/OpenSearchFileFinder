@echo off
title OpenSearch Overnight Maintenance & Optimization
color 0a

echo ======================================================================
echo          OPENSEARCH OVERNIGHT MAINTENANCE & OPTIMIZATION
echo ======================================================================
echo.
echo This script will:
echo   1. Prevent Windows from sleeping during the maintenance run
echo   2. Prune ~60 GB of orphaned thumbnails from .cache_thumbnails
echo   3. Run full document & CSV ingestion and backup
echo   4. Pre-generate all thumbnails multi-threaded (100%% coverage)
echo   5. Generate a morning report in overnight_maintenance_report.txt
echo.
echo ======================================================================

cd /d "%~dp0"

python run_overnight_maintenance.py

echo.
echo ======================================================================
echo Maintenance Complete! Press any key to exit.
echo ======================================================================
pause
