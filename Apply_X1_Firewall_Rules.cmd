@echo off
:: Batch file to elevate and run the hardened X1 Firewall script
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Running with Administrator privileges...
    powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Users\Paul Dexter\.gemini\antigravity-ide\brain\627ed139-4984-4ebd-866e-5071b8925705\scratch\setup_x1_firewall.ps1"
    pause
) else (
    echo Requesting Administrator privileges...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
)
