@echo off
REM ============================================================
REM  Double-click launcher for the User Provisioning Wizard
REM ============================================================
REM  Keep this .bat in the SAME folder as UserProvisioningWizard.ps1.
REM  Double-clicking it launches the wizard directly - no need to
REM  open PowerShell, navigate to the folder, or change the system
REM  execution policy. The -ExecutionPolicy Bypass below applies to
REM  THIS launch only; it does not change any machine-wide setting.
REM ============================================================

REM %~dp0 is the folder this .bat lives in (with trailing backslash),
REM so the wizard is found no matter where the folder is moved to.
set "SCRIPT=%~dp0UserProvisioningWizard.ps1"

if not exist "%SCRIPT%" (
    echo.
    echo Could not find UserProvisioningWizard.ps1 next to this launcher.
    echo Make sure both files are in the same folder.
    echo.
    pause
    exit /b 1
)

REM Prefer Windows PowerShell 5.1 (powershell.exe) since the wizard's
REM WinForms UI wants STA, which is its default. -STA is passed
REM explicitly so it works even if a machine's default ever changes.
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%SCRIPT%"

REM If PowerShell exited with an error, keep the window open so the
REM message is readable instead of vanishing instantly.
if errorlevel 1 (
    echo.
    echo The wizard exited with an error ^(code %errorlevel%^).
    pause
)
