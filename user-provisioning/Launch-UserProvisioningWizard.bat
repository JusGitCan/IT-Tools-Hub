@echo off
REM ============================================================
REM  Double-click launcher for the User Provisioning Wizard
REM ============================================================
REM  Keep this .bat in the SAME folder as UserProvisioningWizard.ps1.
REM  Double-clicking it launches the wizard directly - no need to
REM  open PowerShell, navigate to the folder, or change the system
REM  execution policy. The -ExecutionPolicy Bypass below applies to
REM  THIS launch only; it does not change any machine-wide setting.
REM
REM  IMPORTANT: launch the tool with THIS .bat, not by right-clicking
REM  the .ps1. The .bat forces STA mode, which the Exchange sign-in
REM  requires; running the .ps1 directly can trigger a
REM  "single-threaded apartment" ActiveX error.
REM ============================================================

set "SCRIPT=%~dp0UserProvisioningWizard.ps1"

if not exist "%SCRIPT%" (
    echo.
    echo Could not find UserProvisioningWizard.ps1 next to this launcher.
    echo Make sure both files are in the same folder.
    echo.
    pause
    exit /b 1
)

REM Windows PowerShell 5.1 in STA mode - the combination the WinForms UI
REM and the Exchange Online sign-in control both require.
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%SCRIPT%"

if errorlevel 1 (
    echo.
    echo The wizard exited with an error ^(code %errorlevel%^).
    pause
)
