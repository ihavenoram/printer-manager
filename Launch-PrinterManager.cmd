@echo off
:: Double-click / Backstage launcher for PrinterManager-Gui.ps1.
::
:: ELEVATION: changing queues, drivers, firewall rules and per-computer
:: connections needs admin. Already elevated, or SYSTEM as in Backstage: it starts
:: straight away. Otherwise it relaunches itself through UAC; decline the prompt
:: and the console still opens, unelevated, for read-only work. Pass /noelevate
:: to skip the prompt. The paths and arguments travel to the elevated copy in
:: environment variables, so spaces and quotes in them need no escaping.
::
:: -ExecutionPolicy Bypass matters more once elevated: an elevated PowerShell
:: treats a script on a synced or network drive as remote and refuses it
:: under RemoteSigned. An elevated process may also not see a drive letter that
:: was mapped by the user, so stage the files on a local disk.
::
:: -STA is required for WinForms.
::
:: NOTE: deliberately NO "-WindowStyle Hidden" here. That switch hides the FORM,
:: not just the console host, and the GUI never appears. The script hides its own
:: console window by handle instead (GetConsoleWindow + ShowWindow), which leaves
:: the WinForms window alone. Pass -NoHideConsole if you want the console kept
:: visible for debugging.
setlocal
:: Spaces on BOTH sides keep the variable defined after the markers are removed:
:: an emptied variable makes the next %VAR:x=y% expand literally (measured - the
:: GUI was handed "/elevated=").
set "PTK_ARGS= %* "
set "PTK_MODE=%~1"
set "PTK_ARGS=%PTK_ARGS: /noelevate=%"
set "PTK_ARGS=%PTK_ARGS: /elevated=%"
if /i "%PTK_MODE%"=="/noelevate" goto run
if /i "%PTK_MODE%"=="/elevated" goto run

:: net session succeeds only with an administrator token (SYSTEM included).
net session >nul 2>&1
if not errorlevel 1 goto run

set "PTK_SELF=%~f0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "try { Start-Process -FilePath $env:PTK_SELF -ArgumentList ('/elevated' + $env:PTK_ARGS) -Verb RunAs -ErrorAction Stop; exit 0 } catch { exit 1 }"
if not errorlevel 1 exit /b 0
echo Elevation was declined or failed - opening the console unelevated (read-only actions only).

:run
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0PrinterManager-Gui.ps1"%PTK_ARGS%
endlocal
