@echo off
:: ============================================================
:: Outlook MCP Server - Uninstall Auto-Start
:: ============================================================

echo ============================================================
echo  Outlook MCP Server - Uninstall
echo ============================================================
echo.

:: Stop the task if running
schtasks /end /tn "OutlookMCPServer" >nul 2>&1

:: Delete the task
schtasks /delete /tn "OutlookMCPServer" /f >nul 2>&1
if %errorlevel% equ 0 (
    echo [OK] Scheduled task 'OutlookMCPServer' removed.
) else (
    echo [INFO] No scheduled task found (already removed or never installed).
)

echo.
echo Done.
pause
