@echo off
setlocal enabledelayedexpansion

:: ============================================================
:: Outlook MCP Server - Auto-Start Installer
::
:: Sets up TWO Windows Scheduled Tasks:
::   1. Launch Classic Outlook at logon (COM requires it)
::   2. Start the MCP server at logon (with auto-restart)
::
:: Run as Administrator for best results.
:: ============================================================

echo ============================================================
echo  Outlook MCP Server - Auto-Start Installer
echo ============================================================
echo.

:: Detect Python
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found in PATH.
    echo         Install Python and ensure it is on your PATH.
    pause
    exit /b 1
)

:: Get the directory this script lives in
set "SERVER_DIR=%~dp0"
set "SERVER_DIR=%SERVER_DIR:~0,-1%"

:: Get the full path to python.exe
for /f "tokens=*" %%i in ('where python') do (
    set "PYTHON_EXE=%%i"
    goto :found_python
)
:found_python

set "OUTLOOK_EXE=C:\Program Files\Microsoft Office\Root\Office16\OUTLOOK.EXE"

echo  Python:     %PYTHON_EXE%
echo  Outlook:    %OUTLOOK_EXE%
echo  Server dir: %SERVER_DIR%
echo.

:: Check Classic Outlook exists
if not exist "%OUTLOOK_EXE%" (
    echo [WARNING] Classic Outlook not found at expected path.
    echo          New Outlook does NOT support COM automation.
    echo          The MCP server requires Classic Outlook.
    echo.
)

:: ---- Task 1: Launch Classic Outlook at logon ----
echo [INFO] Creating scheduled task: ClassicOutlookAutoStart...
schtasks /query /tn "ClassicOutlookAutoStart" >nul 2>&1
if %errorlevel% equ 0 (
    schtasks /delete /tn "ClassicOutlookAutoStart" /f >nul 2>&1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$action = New-ScheduledTaskAction -Execute '\"%OUTLOOK_EXE%\"';" ^
    "$trigger = New-ScheduledTaskTrigger -AtLogon;" ^
    "$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Days 365);" ^
    "Register-ScheduledTask -TaskName 'ClassicOutlookAutoStart' -Action $action -Trigger $trigger -Settings $settings -Description 'Launches Classic Outlook at logon (required for MCP COM connection)' -Force"

if %errorlevel% neq 0 (
    echo [WARNING] Could not create ClassicOutlookAutoStart task.
    echo          You may need to run as Administrator.
) else (
    echo [OK] ClassicOutlookAutoStart task created.
)

:: ---- Task 2: Start MCP server at logon (with delay so Outlook starts first) ----
echo [INFO] Creating scheduled task: OutlookMCPServer...
schtasks /query /tn "OutlookMCPServer" >nul 2>&1
if %errorlevel% equ 0 (
    schtasks /delete /tn "OutlookMCPServer" /f >nul 2>&1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$action = New-ScheduledTaskAction -Execute '\"%PYTHON_EXE%\"' -Argument '\"%SERVER_DIR%\outlook_mcp.py\"' -WorkingDirectory '\"%SERVER_DIR%\"';" ^
    "$trigger = New-ScheduledTaskTrigger -AtLogon;" ^
    "$trigger.Delay = 'PT30S';" ^
    "$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit (New-TimeSpan -Days 365);" ^
    "Register-ScheduledTask -TaskName 'OutlookMCPServer' -Action $action -Trigger $trigger -Settings $settings -Description 'Outlook MCP Server - starts 30s after logon, restarts on crash' -Force"

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to create OutlookMCPServer task.
    echo         Try running this script as Administrator.
    pause
    exit /b 1
)

echo [OK] OutlookMCPServer task created.
echo.
echo ============================================================
echo  Setup Complete
echo ============================================================
echo.
echo  At logon:
echo    1. ClassicOutlookAutoStart - launches Classic Outlook
echo    2. OutlookMCPServer        - starts MCP server (30s delay)
echo.
echo  The MCP server auto-restarts up to 3 times on crash.
echo.
echo  Management commands:
echo    Start now:  schtasks /run /tn "OutlookMCPServer"
echo    Stop:       schtasks /end /tn "OutlookMCPServer"
echo    Remove all: run uninstall_service.bat
echo.

set /p START_NOW="Start the server now? (Y/N): "
if /i "%START_NOW%"=="Y" (
    schtasks /run /tn "OutlookMCPServer"
    echo [OK] Server started.
)

echo.
echo ============================================================
echo  Claude Desktop Configuration
echo ============================================================
echo.
echo  Add this to %%APPDATA%%\Claude\claude_desktop_config.json:
echo.
echo  {
echo    "mcpServers": {
echo      "outlook": {
echo        "command": "%PYTHON_EXE%",
echo        "args": ["%SERVER_DIR%\outlook_mcp.py"]
echo      }
echo    }
echo  }
echo.
echo  Claude Desktop manages the server process automatically.
echo  The scheduled tasks handle Classic Outlook + auto-start.
echo ============================================================
echo.
pause
