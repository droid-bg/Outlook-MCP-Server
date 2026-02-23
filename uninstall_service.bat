@echo off
:: ============================================================
:: Outlook MCP Server - Uninstall Auto-Start
:: ============================================================

echo ============================================================
echo  Outlook MCP Server - Uninstall
echo ============================================================
echo.

:: Stop and remove MCP server task
schtasks /end /tn "OutlookMCPServer" >nul 2>&1
schtasks /delete /tn "OutlookMCPServer" /f >nul 2>&1
if %errorlevel% equ 0 (
    echo [OK] OutlookMCPServer task removed.
) else (
    echo [INFO] OutlookMCPServer task not found.
)

:: Remove Classic Outlook auto-start task
schtasks /delete /tn "ClassicOutlookAutoStart" /f >nul 2>&1
if %errorlevel% equ 0 (
    echo [OK] ClassicOutlookAutoStart task removed.
) else (
    echo [INFO] ClassicOutlookAutoStart task not found.
)

echo.
echo Done.
pause
