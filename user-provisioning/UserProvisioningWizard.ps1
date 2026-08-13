#Requires -Version 5.1
<#
    CorroHealth User Provisioning Wizard
    =====================================
    A WinForms wizard for creating Azure AD (Entra ID) users, assigning
    Microsoft 365 licenses, and adding group / distribution-list memberships,
    for both Domestic (US) and Global/India onboarding.

    Rebuilt from the original Weller/India onboarding scripts by
    Brent Waggoner and David Wilhite. Consolidated, de-bugged, GUI added.

    Key behavior
    ------------
    * Two modes via a toggle:
        - Domestic: License type (F3/F3+/E3) input, Country locked to US,
          India logic skipped.
        - Global/India: Mailbox size (2GB/50GB/E3) + EntApps input, Country
          US/IN, subcontractor group routing, CustomAttribute4 tag, and
          archiving/retention for 50GB mailboxes.
      Switching modes hides/shows columns but never deletes row data.
      Importing a file with any India row auto-switches to Global.
    * Existing users are ALWAYS skipped, never modified (no password reset,
      no re-enable). Status notes enabled vs disabled for manual handling.
    * Invalid rows (bad license/size/country, missing name) are skipped
      individually; the rest of the batch still runs.
    * A group-not-found or license failure does not fail the user; the account
      is still created and the problem is noted in Status.
    * Passwords auto-generate uniquely per user (16 chars); shared optional.
    * After creating users, polls until each mailbox exists (5-min timeout)
      before adding groups / archiving, instead of a fixed blind wait.
    * Results CSV (per-user password + status) + transcript written next to
      the script. Treat the results file as sensitive.

    Domestic License <-> Global Mailbox size (same underlying SKUs):
        F3   <-> 2GB   -> SPE_F1
        F3+  <-> 50GB  -> SPE_F1 + EXCHANGEARCHIVE_ADDON
        E3   <-> E3    -> SPE_E3
        EntApps = Y (Global) adds OFFICESUBSCRIPTION on top.

    "Groups" = Entra ID / Exchange distribution membership, NOT Azure
    Resource Groups.
#>

# ============================================================
#  STA self-guard
# ============================================================
# WinForms and Connect-ExchangeOnline's sign-in control both require the
# session to run in a Single-Threaded Apartment (STA). Windows PowerShell 5.1
# defaults to STA, but PowerShell 7 (pwsh) defaults to MTA, which produces:
#   "ActiveX control ... cannot be instantiated because the current thread
#    is not in a single-threaded apartment."
# If we detect we're in MTA, relaunch this same script in a fresh STA
# PowerShell process and exit the current one.
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    $scriptPath = $MyInvocation.MyCommand.Path
    if (-not $scriptPath) { $scriptPath = $PSCommandPath }
    if (-not $scriptPath) { $scriptPath = $MyInvocation.MyCommand.Definition }

    if ($scriptPath -and (Test-Path $scriptPath)) {
        Start-Process -FilePath "powershell.exe" `
            -ArgumentList @("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", "`"$scriptPath`"")
        exit
    } else {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show(
            "This tool must run in STA mode, but it was started in MTA and could not relaunch itself automatically.`n`nPlease launch it using Launch-UserProvisioningWizard.bat instead.",
            "Please use the .bat launcher", "OK", "Warning") | Out-Null
        exit
    }
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ============================================================
#  Corro brand palette + theme helpers
# ============================================================
# Corro brand accents are kept; the base palette is now LIGHT for readability.
# Variable names are unchanged so every downstream reference still works.
$corroPurple    = [System.Drawing.Color]::FromArgb(0x51, 0x2e, 0x6e)   # brand bar / selection
$corroSecondary = [System.Drawing.Color]::FromArgb(0x5e, 0x17, 0x4d)   # borders / inactive toggle
$corroAccent    = [System.Drawing.Color]::FromArgb(0xe7, 0xb5, 0x24)   # gold accent (Run, active toggle)
$themeBg        = [System.Drawing.Color]::White                         # form / grid background
$themePanel     = [System.Drawing.Color]::FromArgb(240, 240, 240)       # toolbars / panels
$themeControl   = [System.Drawing.Color]::White                         # inputs
$themeText      = [System.Drawing.Color]::FromArgb(30, 30, 30)          # main text
$themeTextMuted = [System.Drawing.Color]::FromArgb(130, 130, 130)       # placeholder text
$themeHeaderTxt = [System.Drawing.Color]::FromArgb(0x50, 0x2e, 0x6d)    # column headers in Corro purple
$themeRowAlt    = [System.Drawing.Color]::FromArgb(246, 243, 249)       # subtle purple-tinted zebra

# Names retained (Set-DarkButton / Set-DarkForm) so all call sites are unchanged,
# but they now apply the light theme.
function Set-DarkButton {
    param($Button, [switch]$Accent)
    $Button.FlatStyle = "Flat"
    if ($Accent) {
        $Button.ForeColor = $corroPurple
        $Button.BackColor = $corroAccent
    } else {
        $Button.ForeColor = $corroPurple
        $Button.BackColor = [System.Drawing.Color]::FromArgb(248, 246, 250)
    }
    $Button.FlatAppearance.BorderColor = $corroSecondary
    $Button.FlatAppearance.BorderSize = 1
    $Button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(230, 222, 240)
}
function Set-DarkForm { param($FormObj) $FormObj.BackColor = $themeBg; $FormObj.ForeColor = $themeText }

# ============================================================
#  Valid values + license/size mapping
# ============================================================
$validLicenses  = @("F3", "F3+", "E3")
$validSizes     = @("2GB", "50GB", "E3")
$validCountries = @("US", "IN")
$validYesNo     = @("N", "Y")

function Resolve-LicenseTier {
    param([string]$Value)
    switch (($Value + "").Trim().ToUpper()) {
        "F3"    { return "F3" }
        "2GB"   { return "F3" }
        "2 GB"  { return "F3" }
        "F3+"   { return "F3+" }
        "50GB"  { return "F3+" }
        "50 GB" { return "F3+" }
        "E3"    { return "E3" }
        default { return $null }
    }
}
function Tier-ToSize($tier) { switch ($tier) { "F3" {"2GB"} "F3+" {"50GB"} "E3" {"E3"} default {$null} } }

# ============================================================
#  Password generation
# ============================================================
function New-RandomPassword {
    param([int]$Length = 16)
    $lower="abcdefghijkmnopqrstuvwxyz"; $upper="ABCDEFGHJKLMNPQRSTUVWXYZ"
    $digits="23456789"; $symbols="!@#$%^&*-_="
    $all=$lower+$upper+$digits+$symbols
    $chars=@($lower[(Get-Random -Maximum $lower.Length)],$upper[(Get-Random -Maximum $upper.Length)],$digits[(Get-Random -Maximum $digits.Length)],$symbols[(Get-Random -Maximum $symbols.Length)])
    for ($i=$chars.Count; $i -lt $Length; $i++) { $chars += $all[(Get-Random -Maximum $all.Length)] }
    return -join ($chars | Sort-Object { Get-Random })
}

# ============================================================
#  Password mode dialog
# ============================================================
$pwdDialog = New-Object System.Windows.Forms.Form
$pwdDialog.Text = "Password mode"; $pwdDialog.Size = New-Object System.Drawing.Size(400, 230)
$pwdDialog.StartPosition = "CenterScreen"; $pwdDialog.FormBorderStyle = "FixedDialog"
$pwdDialog.MaximizeBox = $false; $pwdDialog.MinimizeBox = $false

$radioAuto = New-Object System.Windows.Forms.RadioButton
$radioAuto.Text = "Auto-generate a unique 16-character password per user (recommended)"
$radioAuto.Location = New-Object System.Drawing.Point(15, 15); $radioAuto.Size = New-Object System.Drawing.Size(360, 40); $radioAuto.Checked = $true
$radioShared = New-Object System.Windows.Forms.RadioButton
$radioShared.Text = "Use one shared password for all users in this run"
$radioShared.Location = New-Object System.Drawing.Point(15, 60); $radioShared.Size = New-Object System.Drawing.Size(360, 20)
$sharedPwdBox = New-Object System.Windows.Forms.MaskedTextBox
$sharedPwdBox.PasswordChar = "*"; $sharedPwdBox.Location = New-Object System.Drawing.Point(35, 85); $sharedPwdBox.Width = 320; $sharedPwdBox.Enabled = $false
$radioAuto.Add_CheckedChanged({ $sharedPwdBox.Enabled = $radioShared.Checked })
$radioShared.Add_CheckedChanged({ $sharedPwdBox.Enabled = $radioShared.Checked })
$btnPwdOk = New-Object System.Windows.Forms.Button
$btnPwdOk.Text = "Continue"; $btnPwdOk.Location = New-Object System.Drawing.Point(270, 140); $btnPwdOk.Width = 100; $btnPwdOk.DialogResult = "OK"
$pwdDialog.Controls.AddRange([System.Windows.Forms.Control[]]@($radioAuto, $radioShared, $sharedPwdBox, $btnPwdOk))
$pwdDialog.AcceptButton = $btnPwdOk
Set-DarkForm $pwdDialog
$radioAuto.ForeColor = $themeText; $radioShared.ForeColor = $themeText
$sharedPwdBox.BackColor = $themeControl; $sharedPwdBox.ForeColor = $themeText; $sharedPwdBox.BorderStyle = "FixedSingle"
Set-DarkButton $btnPwdOk -Accent

$passwordMode = "Auto"; $sharedPwd = $null
while ($true) {
    if ($pwdDialog.ShowDialog() -ne "OK") { exit }
    if ($radioShared.Checked) {
        if ([string]::IsNullOrWhiteSpace($sharedPwdBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Enter a shared password, or switch to auto-generate.", "Missing password", "OK", "Error") | Out-Null; continue
        }
        $passwordMode = "Shared"; $sharedPwd = $sharedPwdBox.Text
    } else { $passwordMode = "Auto"; $sharedPwd = $null }
    break
}

# ============================================================
#  Connect
# ============================================================
Write-Host "Connecting to Microsoft Graph and Exchange Online..." -ForegroundColor Cyan
try {
    Connect-MgGraph -NoWelcome -Scopes "User.ReadWrite.All", "Group.ReadWrite.All" -ErrorAction Stop
} catch {
    [System.Windows.Forms.MessageBox]::Show("Could not connect to Microsoft Graph:`n`n$($_.Exception.Message)", "Graph sign-in failed", "OK", "Error") | Out-Null
    exit
}
try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
} catch {
    [System.Windows.Forms.MessageBox]::Show("Could not connect to Exchange Online:`n`n$($_.Exception.Message)`n`nIf this is the STA / ActiveX error, make sure you launched via Launch-UserProvisioningWizard.bat.", "Exchange sign-in failed", "OK", "Error") | Out-Null
    exit
}

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$transcriptPath = Join-Path $scriptDir "UserProvisioning_$timestamp.txt"
Start-Transcript -Path $transcriptPath | Out-Null

function Get-LicenseSkus {
    return @{
        F3      = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq "SPE_F1" }
        E3      = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq "SPE_E3" }
        Archive = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq "EXCHANGEARCHIVE_ADDON" }
        Apps    = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq "OFFICESUBSCRIPTION" }
    }
}

# ============================================================
#  Main form + brand header
# ============================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "CorroHealth - User Provisioning Wizard"
$form.Size = New-Object System.Drawing.Size(1180, 660)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(1000, 560)
Set-DarkForm $form

$brandBar = New-Object System.Windows.Forms.Panel
$brandBar.Dock = "Top"; $brandBar.Height = 44; $brandBar.BackColor = $corroPurple

# CorroHealth logo mark (swoosh), white background removed, embedded as base64
# so the image travels inside the script - no separate file to keep alongside it.
$logoB64 = "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"
try {
    $logoBytes = [System.Convert]::FromBase64String($logoB64)
    $logoStream = New-Object System.IO.MemoryStream(,$logoBytes)
    $logoImage = [System.Drawing.Image]::FromStream($logoStream)
    $logoPic = New-Object System.Windows.Forms.PictureBox
    $logoPic.Image = $logoImage
    $logoPic.SizeMode = "AutoSize"
    $logoPic.BackColor = [System.Drawing.Color]::Transparent
    $logoPic.Location = New-Object System.Drawing.Point(12, 7)
    $brandBar.Controls.Add($logoPic)
    $labelX = 54
} catch {
    # If the logo fails to decode for any reason, fall back to text-only.
    $labelX = 12
}

$brandLabel = New-Object System.Windows.Forms.Label
$brandLabel.Text = "CorroHealth  |  User Provisioning"
$brandLabel.ForeColor = [System.Drawing.Color]::White
$brandLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$brandLabel.AutoSize = $true; $brandLabel.Location = New-Object System.Drawing.Point($labelX, 11)
$brandBar.Controls.Add($brandLabel)

$btnModeDomestic = New-Object System.Windows.Forms.Button
$btnModeDomestic.Text = "Domestic"; $btnModeDomestic.Size = New-Object System.Drawing.Size(110, 28); $btnModeDomestic.FlatStyle = "Flat"
$btnModeGlobal = New-Object System.Windows.Forms.Button
$btnModeGlobal.Text = "Global / India"; $btnModeGlobal.Size = New-Object System.Drawing.Size(110, 28); $btnModeGlobal.FlatStyle = "Flat"
$brandBar.Controls.Add($btnModeDomestic); $brandBar.Controls.Add($btnModeGlobal)
$brandBar.Add_Resize({
    $btnModeGlobal.Location = New-Object System.Drawing.Point(($brandBar.Width - 122), 8)
    $btnModeDomestic.Location = New-Object System.Drawing.Point(($brandBar.Width - 234), 8)
})

# ---- Toolbar ----
$toolbar = New-Object System.Windows.Forms.FlowLayoutPanel
$toolbar.Dock = "Top"; $toolbar.Height = 40; $toolbar.Padding = New-Object System.Windows.Forms.Padding(8, 6, 8, 6); $toolbar.BackColor = $themePanel
function New-ToolButton($text, $width) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $text; $b.Width = $width; $b.Height = 26; Set-DarkButton $b; return $b
}
$btnAddRow         = New-ToolButton "Add row" 80
$btnUserCreation   = New-ToolButton "User creation" 100
$btnUploadCsv      = New-ToolButton "Upload CSV" 95
$btnRemoveSelected = New-ToolButton "Remove selected" 115
$btnSelectAll      = New-ToolButton "Select all" 85
$btnLicenseRef     = New-ToolButton "License reference" 120
$toolbar.Controls.AddRange([System.Windows.Forms.Control[]]@($btnAddRow, $btnUserCreation, $btnUploadCsv, $btnRemoveSelected, $btnSelectAll, $btnLicenseRef))

# ---- Grid ----
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = "Fill"
$grid.AllowUserToAddRows = $false; $grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false; $grid.AllowUserToResizeColumns = $false
$grid.RowHeadersVisible = $false; $grid.ColumnHeadersVisible = $false
$grid.AutoSizeColumnsMode = "None"; $grid.RowTemplate.Height = 26; $grid.SelectionMode = "FullRowSelect"
$grid.BackgroundColor = $themeBg; $grid.GridColor = [System.Drawing.Color]::FromArgb(220, 214, 228); $grid.BorderStyle = "None"; $grid.EnableHeadersVisualStyles = $false
$grid.DefaultCellStyle.BackColor = $themeBg; $grid.DefaultCellStyle.ForeColor = $themeText
$grid.DefaultCellStyle.SelectionBackColor = $corroPurple; $grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
$grid.AlternatingRowsDefaultCellStyle.BackColor = $themeRowAlt; $grid.AlternatingRowsDefaultCellStyle.ForeColor = $themeText
$grid.AlternatingRowsDefaultCellStyle.SelectionBackColor = $corroPurple; $grid.AlternatingRowsDefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White

function New-TextColumn($name, $width) {
    $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $c.Name = $name; $c.Width = $width; return $c
}
function New-ComboColumn($name, $width, $items) {
    $c = New-Object System.Windows.Forms.DataGridViewComboBoxColumn; $c.Name = $name; $c.Width = $width; $c.FlatStyle = "Flat"; $c.Items.AddRange($items); return $c
}
$colInclude = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colInclude.Name = "Include"; $colInclude.Width = 34
$colFirst   = New-TextColumn "FirstName" 95
$colLast    = New-TextColumn "LastName" 95
$colUpn     = New-TextColumn "Upn" 160
$colTitle   = New-TextColumn "JobTitle" 105
$colLicense = New-ComboColumn "License" 95 $validLicenses
$colEntApps = New-ComboColumn "EntApps" 70 $validYesNo
$colCountry = New-ComboColumn "Country" 65 $validCountries
$colInternal= New-ComboColumn "InternalEmailOnly" 90 $validYesNo
$colSubcon  = New-ComboColumn "Subcontractor" 95 $validYesNo
$colCity    = New-TextColumn "City" 90
$colProvince= New-TextColumn "Province" 90
$colOffice  = New-TextColumn "Office" 70
$colGroups  = New-TextColumn "Groups" 160
$colStatus  = New-TextColumn "Status" 200
$colStatus.ReadOnly = $true
$colDelete = New-Object System.Windows.Forms.DataGridViewButtonColumn
$colDelete.Name = "Delete"; $colDelete.Text = "Remove"; $colDelete.UseColumnTextForButtonValue = $true; $colDelete.Width = 70
$grid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@(
    $colInclude, $colFirst, $colLast, $colUpn, $colTitle, $colLicense, $colEntApps,
    $colCountry, $colInternal, $colSubcon, $colCity, $colProvince, $colOffice, $colGroups, $colStatus, $colDelete
))

# ---- Header label bar ----
$gridHeaderBar = New-Object System.Windows.Forms.Panel
$gridHeaderBar.Dock = "Top"; $gridHeaderBar.Height = 30; $gridHeaderBar.BackColor = $themeBg
$headerFont = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$headerSeparator = New-Object System.Windows.Forms.Panel
$headerSeparator.Dock = "Top"; $headerSeparator.Height = 1; $headerSeparator.BackColor = $corroPurple

function Get-HeaderText($colName, $mode) {
    switch ($colName) {
        "Include"           { return "" }
        "FirstName"         { return "First name" }
        "LastName"          { return "Last name" }
        "Upn"               { return "Username" }
        "JobTitle"          { return "Job title" }
        "License"           { if ($mode -eq "Global") { return "Mailbox size" } else { return "License" } }
        "EntApps"           { return "EntApps" }
        "Country"           { return "Country" }
        "InternalEmailOnly" { return "Internal E-Mail" }
        "Subcontractor"     { return "Subcontractor" }
        "City"              { return "City" }
        "Province"          { return "Province" }
        "Office"            { return "Office" }
        "Groups"            { return "Groups" }
        "Status"            { return "Status" }
        "Delete"            { return "" }
        default             { return $colName }
    }
}
function Rebuild-HeaderBar($mode) {
    $gridHeaderBar.Controls.Clear()
    $x = 2 - $grid.HorizontalScrollingOffset
    foreach ($col in $grid.Columns) {
        if ($col.Visible) {
            $lbl = New-Object System.Windows.Forms.Label
            $lbl.Text = Get-HeaderText $col.Name $mode
            $lbl.Font = $headerFont; $lbl.ForeColor = $themeHeaderTxt
            $lbl.Location = New-Object System.Drawing.Point($x, 6)
            $lbl.Size = New-Object System.Drawing.Size($col.Width, 20); $lbl.TextAlign = "MiddleLeft"
            $gridHeaderBar.Controls.Add($lbl)
            $x += $col.Width
        }
    }
}
$grid.Add_Scroll({ Rebuild-HeaderBar $script:currentMode })

# ---- Bulk apply panel ----
$bulkPanel = New-Object System.Windows.Forms.Panel
$bulkPanel.Dock = "Bottom"; $bulkPanel.Height = 60; $bulkPanel.Padding = New-Object System.Windows.Forms.Padding(8); $bulkPanel.BackColor = $themePanel
$lblApply = New-Object System.Windows.Forms.Label
$lblApply.Text = "Apply to selected rows:"; $lblApply.AutoSize = $true; $lblApply.Location = New-Object System.Drawing.Point(8, 8); $lblApply.ForeColor = $themeText
$bulkLicense = New-Object System.Windows.Forms.ComboBox
$bulkLicense.DropDownStyle = "DropDownList"; $bulkLicense.Location = New-Object System.Drawing.Point(8, 28); $bulkLicense.Width = 110
$bulkLicense.BackColor = $themeControl; $bulkLicense.ForeColor = $themeText; $bulkLicense.FlatStyle = "Flat"
$bulkCountry = New-Object System.Windows.Forms.ComboBox
$bulkCountry.DropDownStyle = "DropDownList"; $bulkCountry.Location = New-Object System.Drawing.Point(126, 28); $bulkCountry.Width = 90
$bulkCountry.BackColor = $themeControl; $bulkCountry.ForeColor = $themeText; $bulkCountry.FlatStyle = "Flat"
$bulkCountry.Items.AddRange(@("(no change)") + $validCountries); $bulkCountry.SelectedIndex = 0
$bulkTitle = New-Object System.Windows.Forms.TextBox
$bulkTitle.Location = New-Object System.Drawing.Point(224, 28); $bulkTitle.Width = 150
$bulkTitle.BackColor = $themeControl; $bulkTitle.BorderStyle = "FixedSingle"; $bulkTitle.ForeColor = $themeTextMuted; $bulkTitle.Text = "Job title (blank = no change)"
$bulkGroups = New-Object System.Windows.Forms.TextBox
$bulkGroups.Location = New-Object System.Drawing.Point(382, 28); $bulkGroups.Width = 200
$bulkGroups.BackColor = $themeControl; $bulkGroups.BorderStyle = "FixedSingle"; $bulkGroups.ForeColor = $themeTextMuted; $bulkGroups.Text = "Groups (blank = no change)"
$btnApplySelected = New-Object System.Windows.Forms.Button
$btnApplySelected.Text = "Apply to selected"; $btnApplySelected.Location = New-Object System.Drawing.Point(590, 27); $btnApplySelected.Width = 130
Set-DarkButton $btnApplySelected
$bulkPanel.Controls.AddRange([System.Windows.Forms.Control[]]@($lblApply, $bulkLicense, $bulkCountry, $bulkTitle, $bulkGroups, $btnApplySelected))
foreach ($tb in @($bulkTitle, $bulkGroups)) {
    $tb.Add_Enter({ if ($this.ForeColor -eq $themeTextMuted) { $this.Text = ""; $this.ForeColor = $themeText } })
}

# ---- Action bar ----
$actionBar = New-Object System.Windows.Forms.Panel
$actionBar.Dock = "Bottom"; $actionBar.Height = 50; $actionBar.Padding = New-Object System.Windows.Forms.Padding(8); $actionBar.BackColor = $themePanel
$lblCount = New-Object System.Windows.Forms.Label
$lblCount.AutoSize = $true; $lblCount.Text = "0 users queued"; $lblCount.Location = New-Object System.Drawing.Point(8, 8); $lblCount.ForeColor = $themeText
$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.AutoSize = $true; $lblProgress.Text = ""; $lblProgress.Location = New-Object System.Drawing.Point(8, 28); $lblProgress.ForeColor = $corroPurple
$btnRunImport = New-Object System.Windows.Forms.Button
$btnRunImport.Text = "Run import"; $btnRunImport.Width = 130; $btnRunImport.Height = 32; $btnRunImport.Anchor = "Right"
Set-DarkButton $btnRunImport -Accent
$actionBar.Controls.Add($lblCount); $actionBar.Controls.Add($lblProgress); $actionBar.Controls.Add($btnRunImport)
$actionBar.Add_Resize({ $btnRunImport.Location = New-Object System.Drawing.Point(($actionBar.Width - 145), 9) })

$form.Controls.Add($grid)
$form.Controls.Add($headerSeparator)
$form.Controls.Add($gridHeaderBar)
$form.Controls.Add($bulkPanel)
$form.Controls.Add($actionBar)
$form.Controls.Add($toolbar)
$form.Controls.Add($brandBar)

# ============================================================
#  Mode handling
# ============================================================
$script:currentMode = "Domestic"
$domesticCols = @("Include","FirstName","LastName","Upn","JobTitle","License","Groups","Status","Delete")
$globalCols   = @("Include","FirstName","LastName","Upn","JobTitle","License","EntApps","Country","InternalEmailOnly","Subcontractor","City","Province","Office","Status","Delete")

function Set-Mode($mode) {
    $script:currentMode = $mode
    $visible = if ($mode -eq "Global") { $globalCols } else { $domesticCols }
    foreach ($col in $grid.Columns) { $col.Visible = ($visible -contains $col.Name) }

    if ($mode -eq "Global") {
        $colLicense.Items.Clear(); $colLicense.Items.AddRange($validSizes)
        $colCountry.ReadOnly = $false; $btnLicenseRef.Visible = $true
    } else {
        $colLicense.Items.Clear(); $colLicense.Items.AddRange($validLicenses)
        $colCountry.ReadOnly = $true; $btnLicenseRef.Visible = $false
        foreach ($row in $grid.Rows) { $row.Cells["Country"].Value = "US" }
    }

    $allowed = if ($mode -eq "Global") { $validSizes } else { $validLicenses }
    foreach ($row in $grid.Rows) {
        $v = "$($row.Cells['License'].Value)"
        if ($v -and ($allowed -notcontains $v)) {
            $tier = Resolve-LicenseTier $v
            $mapped = if ($mode -eq "Global") { Tier-ToSize $tier } else { $tier }
            $row.Cells["License"].Value = $mapped
        }
    }

    $inactiveBg  = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $inactiveTxt = [System.Drawing.Color]::FromArgb(60, 60, 60)
    if ($mode -eq "Global") {
        $btnModeGlobal.BackColor = $corroAccent; $btnModeGlobal.ForeColor = $corroPurple
        $btnModeDomestic.BackColor = $inactiveBg; $btnModeDomestic.ForeColor = $inactiveTxt
    } else {
        $btnModeDomestic.BackColor = $corroAccent; $btnModeDomestic.ForeColor = $corroPurple
        $btnModeGlobal.BackColor = $inactiveBg; $btnModeGlobal.ForeColor = $inactiveTxt
    }
    foreach ($b in @($btnModeDomestic, $btnModeGlobal)) { $b.FlatAppearance.BorderSize = 0 }
    Rebuild-HeaderBar $mode
}
$btnModeDomestic.Add_Click({ Set-Mode "Domestic"; Update-QueueCount })
$btnModeGlobal.Add_Click({ Set-Mode "Global"; Update-QueueCount })

# ============================================================
#  Row helpers
# ============================================================
function Add-BlankRow {
    $i = $grid.Rows.Add(); $r = $grid.Rows[$i]; $r.Height = 26
    $r.Cells["Include"].Value = $true
    $r.Cells["License"].Value = if ($script:currentMode -eq "Global") { "2GB" } else { "F3" }
    $r.Cells["EntApps"].Value = "N"; $r.Cells["Country"].Value = "US"
    $r.Cells["InternalEmailOnly"].Value = "N"; $r.Cells["Subcontractor"].Value = "N"
    return $i
}
function Update-QueueCount { $lblCount.Text = "$($grid.Rows.Count) users queued  |  Mode: $($script:currentMode)" }

# ============================================================
#  Toolbar behavior
# ============================================================
$btnAddRow.Add_Click({ Add-BlankRow | Out-Null; Update-QueueCount })
$btnSelectAll.Add_Click({ foreach ($row in $grid.Rows) { $row.Cells["Include"].Value = $true } })
$btnRemoveSelected.Add_Click({
    $toRemove = @(); foreach ($row in $grid.Rows) { if ($row.Cells["Include"].Value -eq $true) { $toRemove += $row } }
    foreach ($row in $toRemove) { $grid.Rows.Remove($row) }; Update-QueueCount
})
$grid.Add_CellClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    if ($grid.Columns[$e.ColumnIndex].Name -eq "Delete") { $grid.Rows.RemoveAt($e.RowIndex); Update-QueueCount }
})
$btnLicenseRef.Add_Click({
    $msg = "Mailbox size  ->  License granted`n`n2GB   ->  F3`n50GB  ->  F3 + Exchange Archive (F3+)`nE3    ->  E3 (full desktop Office)`n`nEntApps = Y adds Microsoft 365 Apps for Enterprise on top of any of the above."
    [System.Windows.Forms.MessageBox]::Show($msg, "License reference", "OK", "Information") | Out-Null
})

# ---- User creation (paste names) ----
$btnUserCreation.Add_Click({
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "User creation - paste names"; $dlg.Size = New-Object System.Drawing.Size(420, 340); $dlg.StartPosition = "CenterParent"
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Paste one name per line (First Last):"; $lbl.AutoSize = $true; $lbl.Location = New-Object System.Drawing.Point(10, 10)
    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Multiline = $true; $txt.ScrollBars = "Vertical"; $txt.Location = New-Object System.Drawing.Point(10, 32); $txt.Size = New-Object System.Drawing.Size(390, 220)
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Add to grid"; $btnOk.Location = New-Object System.Drawing.Point(230, 262); $btnOk.DialogResult = "OK"
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"; $btnCancel.Location = New-Object System.Drawing.Point(325, 262); $btnCancel.DialogResult = "Cancel"
    $dlg.Controls.AddRange([System.Windows.Forms.Control[]]@($lbl, $txt, $btnOk, $btnCancel))
    $dlg.AcceptButton = $btnOk; $dlg.CancelButton = $btnCancel
    Set-DarkForm $dlg; $lbl.ForeColor = $themeText
    $txt.BackColor = $themeControl; $txt.ForeColor = $themeText; $txt.BorderStyle = "FixedSingle"
    Set-DarkButton $btnOk; Set-DarkButton $btnCancel
    if ($dlg.ShowDialog() -eq "OK") {
        $lines = $txt.Text -split "`r?`n" | Where-Object { $_.Trim() -ne "" }
        foreach ($line in $lines) {
            $parts = $line.Trim() -split "\s+", 2
            $i = Add-BlankRow
            $grid.Rows[$i].Cells["FirstName"].Value = $parts[0]
            $grid.Rows[$i].Cells["LastName"].Value = if ($parts.Count -gt 1) { $parts[1] } else { "" }
        }
        Update-QueueCount
    }
})

# ---- Shared importer ----
function Import-Rows($rows) {
    $hasIndia = $false
    foreach ($r in $rows) { if (("$($r.UserCountry)$($r.Country)").Trim() -match "(?i)^(india|in)$") { $hasIndia = $true; break } }
    if ($hasIndia -and $script:currentMode -ne "Global") { Set-Mode "Global" }
    foreach ($r in $rows) {
        $i = Add-BlankRow; $row = $grid.Rows[$i]
        $row.Cells["FirstName"].Value = ("$($r.FirstName)").Trim()
        $row.Cells["LastName"].Value  = ("$($r.LastName)").Trim()
        $upn = ("$($r.UserPrincipalName)").Trim(); if ($upn) { $row.Cells["Upn"].Value = $upn }
        $row.Cells["JobTitle"].Value = ("$($r.Designation)").Trim()
        $rawLic = ("$($r.LicenseCode)$($r.RequiredMailboxSize)").Trim()
        $tier = Resolve-LicenseTier $rawLic
        if ($tier) {
            if ($script:currentMode -eq "Global") { $row.Cells["License"].Value = Tier-ToSize $tier }
            else { $row.Cells["License"].Value = $tier }
        }
        if (("$($r.EntApps)").Trim() -match "(?i)^y") { $row.Cells["EntApps"].Value = "Y" }
        $country = ("$($r.UserCountry)$($r.Country)").Trim()
        $row.Cells["Country"].Value = if ($country -match "(?i)^(india|in)$") { "IN" } else { "US" }
        if (("$($r.InternalEmailOnly)").Trim() -match "(?i)^y") { $row.Cells["InternalEmailOnly"].Value = "Y" }
        if (("$($r.SubContractor)$($r.Subcontractor)").Trim() -match "(?i)^y") { $row.Cells["Subcontractor"].Value = "Y" }
        $row.Cells["City"].Value = ("$($r.City)").Trim()
        $row.Cells["Province"].Value = ("$($r.Province)").Trim()
        $row.Cells["Office"].Value = ("$($r.Office)").Trim()
    }
    Update-QueueCount
}

# ---- Upload CSV ----
$btnUploadCsv.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"; $ofd.Title = "Select user import CSV"
    if ($ofd.ShowDialog() -ne "OK") { return }
    try { $rows = Import-Csv -Path $ofd.FileName }
    catch { [System.Windows.Forms.MessageBox]::Show("Could not read that CSV: $($_.Exception.Message)", "Import error", "OK", "Error") | Out-Null; return }
    Import-Rows $rows
})

# ---- Bulk apply ----
$btnApplySelected.Add_Click({
    $count = 0
    foreach ($row in $grid.Rows) {
        if ($row.Cells["Include"].Value -eq $true) {
            $count++
            if ($bulkLicense.SelectedItem -and $bulkLicense.SelectedItem -ne "(no change)") { $row.Cells["License"].Value = $bulkLicense.SelectedItem }
            if ($script:currentMode -eq "Global" -and $bulkCountry.SelectedItem -and $bulkCountry.SelectedItem -ne "(no change)") { $row.Cells["Country"].Value = $bulkCountry.SelectedItem }
            if ($bulkTitle.Text -and $bulkTitle.ForeColor -ne $themeTextMuted) { $row.Cells["JobTitle"].Value = $bulkTitle.Text }
            if ($bulkGroups.Text -and $bulkGroups.ForeColor -ne $themeTextMuted) { $row.Cells["Groups"].Value = $bulkGroups.Text }
        }
    }
    if ($count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No rows selected. Check the box on each row to update.", "Nothing selected", "OK", "Information") | Out-Null }
})
$bulkLicense.Add_DropDown({
    $bulkLicense.Items.Clear()
    $vals = if ($script:currentMode -eq "Global") { $validSizes } else { $validLicenses }
    $bulkLicense.Items.Add("(no change)") | Out-Null
    foreach ($v in $vals) { $bulkLicense.Items.Add($v) | Out-Null }
})

# ============================================================
#  Provisioning helpers
# ============================================================
function Get-UpnParts($firstName, $lastName) {
    $fn = ($firstName.Trim()) -replace '\s+', ''; $ln = ($lastName.Trim()) -replace '\s+', ''
    return @{ DisplayName = "$($firstName.Trim()) $($lastName.Trim())"; MailNickname = "$fn.$ln"; Upn = "$fn.$ln@corrohealth.com" }
}
function Add-UserToGroupOrList($upn, $userOid, $groupName) {
    $group = Get-MgGroup -Filter "DisplayName eq '$groupName'" -ErrorAction SilentlyContinue
    if (-not $group) { return "Group not found: $groupName" }
    if ($group.Mail -like "*corrohealth.com*") {
        try { Add-DistributionGroupMember -Identity $group.DisplayName -Member $upn -BypassSecurityGroupManagerCheck -ErrorAction Stop; return "Added to $groupName (Exchange)" }
        catch { return "Failed adding to $($groupName): $($_.Exception.Message)" }
    } else {
        try { New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $userOid -ErrorAction Stop; return "Added to $groupName (Entra ID)" }
        catch { return "Failed adding to $($groupName): $($_.Exception.Message)" }
    }
}
function Wait-ForMailbox($upn, $timeoutSeconds = 300, $intervalSeconds = 15) {
    $elapsed = 0
    while ($elapsed -lt $timeoutSeconds) {
        try { $mbx = Get-Mailbox -Identity $upn -ErrorAction SilentlyContinue; if ($mbx) { return $true } } catch { }
        Start-Sleep -Seconds $intervalSeconds; $elapsed += $intervalSeconds
    }
    return $false
}

# ============================================================
#  Run import
# ============================================================
$btnRunImport.Add_Click({
    $mode = $script:currentMode
    $rowsToProcess = @()
    foreach ($row in $grid.Rows) { if ($row.Cells["Include"].Value -eq $true) { $rowsToProcess += $row } }
    if ($rowsToProcess.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No rows selected to import.", "Nothing to run", "OK", "Information") | Out-Null; return }
    $confirm = [System.Windows.Forms.MessageBox]::Show("This will create $($rowsToProcess.Count) user(s) in $mode mode. Existing users will be skipped, not modified. Continue?", "Confirm import", "YesNo", "Question")
    if ($confirm -ne "Yes") { return }

    $skus = Get-LicenseSkus
    $btnRunImport.Enabled = $false; $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $results = New-Object System.Collections.ArrayList
    $created = @()

    foreach ($row in $rowsToProcess) {
        $fn = "$($row.Cells['FirstName'].Value)".Trim()
        $ln = "$($row.Cells['LastName'].Value)".Trim()
        $upnOverride = "$($row.Cells['Upn'].Value)".Trim()
        $title = "$($row.Cells['JobTitle'].Value)".Trim()
        $licRaw = "$($row.Cells['License'].Value)".Trim()
        $country = if ($mode -eq "Global") { "$($row.Cells['Country'].Value)".Trim() } else { "US" }
        $ent = "$($row.Cells['EntApps'].Value)".Trim()
        $internal = "$($row.Cells['InternalEmailOnly'].Value)".Trim()
        $subcon = "$($row.Cells['Subcontractor'].Value)".Trim()
        $city = "$($row.Cells['City'].Value)".Trim()
        $province = "$($row.Cells['Province'].Value)".Trim()
        $office = "$($row.Cells['Office'].Value)".Trim()
        $groupsRaw = "$($row.Cells['Groups'].Value)".Trim()

        $tier = Resolve-LicenseTier $licRaw
        $parts = Get-UpnParts $fn $ln
        $upn = if ($upnOverride) { $upnOverride } else { $parts.Upn }
        $userPwd = if ($passwordMode -eq "Shared") { $sharedPwd } else { New-RandomPassword -Length 16 }

        $recordStatus = {
            param($status)
            [void]$results.Add([PSCustomObject]@{ FirstName=$fn; LastName=$ln; UserPrincipalName=$upn; Password=$userPwd; LicenseTier=$tier; Country=$country; Groups=$groupsRaw; Status=$status })
            $row.Cells["Status"].Value = $status
        }

        if ([string]::IsNullOrWhiteSpace($fn) -or [string]::IsNullOrWhiteSpace($ln)) { & $recordStatus "Skipped - missing first or last name"; continue }
        if (-not $tier) { & $recordStatus "Skipped - invalid license/size '$licRaw'"; continue }
        if ($mode -eq "Global" -and $country -notin $validCountries) { & $recordStatus "Skipped - invalid country '$country'"; continue }

        $existing = $null
        try { $existing = Get-MgUser -UserId $upn -Property Id, AccountEnabled -ErrorAction SilentlyContinue } catch { $existing = $null }
        if ($existing) {
            $state = if ($existing.AccountEnabled) { "enabled" } else { "DISABLED" }
            & $recordStatus "Skipped - already exists ($state)"; continue
        }

        $pwProfile = @{ Password = $userPwd; ForceChangePasswordNextSignIn = $true }
        $mailNickname = ($upn -split "@")[0]
        $newParams = @{ GivenName=$fn; Surname=$ln; DisplayName=$parts.DisplayName; AccountEnabled=$true; UserPrincipalName=$upn; MailNickname=$mailNickname; PasswordProfile=$pwProfile; UsageLocation=$country; JobTitle=$title }
        if ($country) { $newParams["Country"] = $country }
        if ($city) { $newParams["City"] = $city }
        if ($province) { $newParams["State"] = $province }
        if ($office) { $newParams["OfficeLocation"] = $office }

        $userOid = $null
        try { $newUser = New-MgUser @newParams -ErrorAction Stop; $userOid = $newUser.Id }
        catch { & $recordStatus "FAILED creating user: $($_.Exception.Message)"; continue }

        $notes = @("Created")
        try {
            switch ($tier) {
                "F3"  { Set-MgUserLicense -UserId $upn -AddLicenses @{SkuId=$skus.F3.SkuId} -RemoveLicenses @() -ErrorAction Stop; $notes += "F3" }
                "F3+" {
                    Set-MgUserLicense -UserId $upn -AddLicenses @{SkuId=$skus.F3.SkuId} -RemoveLicenses @() -ErrorAction Stop
                    if ($mode -eq "Global") {
                        # Global "50GB" tier = F3 + Exchange Online Archiving (bigger mailbox).
                        Set-MgUserLicense -UserId $upn -AddLicenses @{SkuId=$skus.Archive.SkuId} -RemoveLicenses @() -ErrorAction Stop
                        $notes += "F3 + Exchange Archive (50GB)"
                    } else {
                        # Domestic "F3+" = F3 + Microsoft 365 Apps for enterprise (desktop Office).
                        Set-MgUserLicense -UserId $upn -AddLicenses @{SkuId=$skus.Apps.SkuId} -RemoveLicenses @() -ErrorAction Stop
                        $notes += "F3 + Apps for Enterprise"
                    }
                }
                "E3"  { Set-MgUserLicense -UserId $upn -AddLicenses @{SkuId=$skus.E3.SkuId} -RemoveLicenses @() -ErrorAction Stop; $notes += "E3" }
            }
            if ($ent -match "(?i)^y") { Set-MgUserLicense -UserId $upn -AddLicenses @{SkuId=$skus.Apps.SkuId} -RemoveLicenses @() -ErrorAction Stop; $notes += "Apps" }
        } catch { $notes += "License FAILED: $($_.Exception.Message)" }

        & $recordStatus ($notes -join " | ")
        $created += [PSCustomObject]@{ Row=$row; Upn=$upn; Oid=$userOid; Tier=$tier; Subcon=$subcon; Internal=$internal; GroupsRaw=$groupsRaw; Notes=$notes; Country=$country }
    }

    if ($created.Count -gt 0) {
        $ready = 0
        foreach ($c in $created) {
            $notes = $c.Notes

            # Build the full group list for this user first (before any waiting).
            $groupNames = @()
            if ($c.GroupsRaw) { $groupNames += ($c.GroupsRaw -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
            if ($mode -eq "Global") {
                $global = if ($c.Subcon -match "(?i)^y") { "P-SG-InTune-Global-SubContractor-User-Group" } else { "P-SG-InTune-Global-Team_Member-User-Group" }
                if ($c.Tier -eq "E3") { $groupNames += @("India O365 Login Access", $global) } else { $groupNames += @("internal email only", "Disable Outlook Access", $global) }
            } elseif ($c.Internal -match "(?i)^y") { $groupNames += "Internal Email Only" }
            $groupNames = @($groupNames | Select-Object -Unique)

            # Resolve each group now so we can tell Exchange (mail-enabled) from
            # Entra security groups. Exchange distribution membership, archiving,
            # and the subcontractor attribute all require the mailbox to exist;
            # Entra security-group adds and licensing do not.
            $resolved = @()
            $needsMailbox = $false
            foreach ($g in $groupNames) {
                $grp = Get-MgGroup -Filter "DisplayName eq '$g'" -ErrorAction SilentlyContinue
                $isExchange = ($grp -and ($grp.Mail -like "*corrohealth.com*"))
                if ($isExchange) { $needsMailbox = $true }
                $resolved += [PSCustomObject]@{ Name = $g; Group = $grp; IsExchange = $isExchange }
            }
            $wantsSubconAttr = ($mode -eq "Global" -and $c.Subcon -match "(?i)^y")
            $wantsArchive    = ($mode -eq "Global" -and $c.Tier -eq "F3+")
            if ($wantsSubconAttr -or $wantsArchive) { $needsMailbox = $true }

            # Add Entra (non-mailbox) group memberships immediately.
            foreach ($r in $resolved) {
                if (-not $r.IsExchange) {
                    if (-not $r.Group) { $notes += "Group not found: $($r.Name)" }
                    else {
                        try { New-MgGroupMember -GroupId $r.Group.Id -DirectoryObjectId $c.Oid -ErrorAction Stop; $notes += "Added to $($r.Name) (Entra ID)" }
                        catch { $notes += "Failed adding to $($r.Name): $($_.Exception.Message)" }
                    }
                }
            }

            # Only wait for the mailbox if something actually depends on it.
            $ok = $true
            if ($needsMailbox) {
                $lblProgress.Text = "Waiting for mailbox... ($ready of $($created.Count))"; [System.Windows.Forms.Application]::DoEvents()
                $ok = Wait-ForMailbox $c.Upn 300 15

                if ($wantsSubconAttr -and $ok) {
                    try { Set-Mailbox $c.Upn -CustomAttribute4 "SubContractor" -ErrorAction Stop; $notes += "CustomAttr4 set" }
                    catch { $notes += "CustomAttr4 FAILED: $($_.Exception.Message)" }
                }
                foreach ($r in $resolved) {
                    if ($r.IsExchange) {
                        if (-not $r.Group) { $notes += "Group not found: $($r.Name)" }
                        else {
                            try { Add-DistributionGroupMember -Identity $r.Group.DisplayName -Member $c.Upn -BypassSecurityGroupManagerCheck -ErrorAction Stop; $notes += "Added to $($r.Name) (Exchange)" }
                            catch { $notes += "Failed adding to $($r.Name): $($_.Exception.Message)" }
                        }
                    }
                }
                if ($wantsArchive -and $ok) {
                    try { Enable-Mailbox -Identity $c.Upn -Archive -ErrorAction Stop; $notes += "Archive enabled" } catch { $notes += "Archive FAILED: $($_.Exception.Message)" }
                    try { Set-Mailbox -Identity $c.Upn -RetentionPolicy "India F3 Users" -ErrorAction Stop; $notes += "Retention set" } catch { $notes += "Retention FAILED: $($_.Exception.Message)" }
                }
                if (-not $ok) { $notes += "Mailbox not ready within timeout - some Exchange steps may be incomplete" }
            }

            $c.Row.Cells["Status"].Value = ($notes -join " | ")
            foreach ($res in $results) { if ($res.UserPrincipalName -eq $c.Upn) { $res.Status = ($notes -join " | ") } }
            $ready++
        }
        $lblProgress.Text = "Done - $ready user(s) processed."
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::Default; $btnRunImport.Enabled = $true
    $resultsPath = Join-Path $scriptDir "UserProvisioning_Results_$timestamp.csv"
    try { $results | Export-Csv -Path $resultsPath -NoTypeInformation -ErrorAction Stop; $exportNote = "Results + credentials saved to:`n$resultsPath" }
    catch { $exportNote = "Could not save results file: $($_.Exception.Message)" }
    $pwdSummary = if ($passwordMode -eq "Shared") { "Shared password: $sharedPwd" } else { "Each user got a unique password - see the results file." }
    [System.Windows.Forms.MessageBox]::Show("Import finished. See the Status column for per-user results.`n`n$pwdSummary`n`n$exportNote", "Import complete", "OK", "Information") | Out-Null
})

# ============================================================
#  Startup
# ============================================================
$form.Add_Shown({
    Set-Mode "Domestic"
    foreach ($row in $grid.Rows) { $row.Height = 26 }
    $grid.PerformLayout(); $grid.Invalidate($true); $grid.Refresh()
    Rebuild-HeaderBar $script:currentMode
})
Set-Mode "Domestic"
Add-BlankRow | Out-Null
Update-QueueCount
[void]$form.ShowDialog()
Stop-Transcript | Out-Null
