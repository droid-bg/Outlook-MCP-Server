@echo off
setlocal enabledelayedexpansion

:: ============================================================
:: Outlook MCP Server - Auto-Start Installer
::
:: Sets up a Windows Scheduled Task that:
::   - Starts the MCP server at user logon
::   - Restarts automatically on crash (up to 3 times)
::   - Runs in the background (no console window)
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

echo  Python:     %PYTHON_EXE%
echo  Server dir: %SERVER_DIR%
echo.

:: Remove existing task if present
schtasks /query /tn "OutlookMCPServer" >nul 2>&1
if %errorlevel% equ 0 (
    echo [INFO] Removing existing scheduled task...
    schtasks /delete /tn "OutlookMCPServer" /f >nul 2>&1
)

:: Create the scheduled task using PowerShell for advanced settings
:: (schtasks.exe cannot set RestartCount / RestartInterval)
echo [INFO] Creating scheduled task with auto-restart...

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$action = New-ScheduledTaskAction -Execute '\"%PYTHON_EXE%\"' -Argument '\"%SERVER_DIR%\outlook_mcp.py\"' -WorkingDirectory '\"%SERVER_DIR%\"';" ^
    "$trigger = New-ScheduledTaskTrigger -AtLogon;" ^
    "$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit (New-TimeSpan -Days 365);" ^
    "Register-ScheduledTask -TaskName 'OutlookMCPServer' -Action $action -Trigger $trigger -Settings $settings -Description 'Outlook MCP Server - auto-starts at logon, restarts on crash' -Force"

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to create scheduled task.
    echo         Try running this script as Administrator.
    pause
    exit /b 1
)

echo.
echo [OK] Scheduled task 'OutlookMCPServer' created successfully.
echo.
echo  - Starts automatically at logon
echo  - Restarts up to 3 times on crash (1 min interval)
echo  - To start now:    schtasks /run /tn "OutlookMCPServer"
echo  - To stop:         schtasks /end /tn "OutlookMCPServer"
echo  - To remove:       run uninstall_service.bat
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
echo  Add this to your Claude Desktop config
echo  (%%APPDATA%%\Claude\claude_desktop_config.json):
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
echo  The scheduled task is a fallback for standalone usage.
echo ============================================================
echo.
pause
