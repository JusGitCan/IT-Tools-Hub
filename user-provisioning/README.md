# User Provisioning Wizard

A PowerShell WinForms wizard for creating/updating Azure AD (Entra ID) user
accounts, assigning Microsoft 365 licenses, and adding users to Entra ID
groups or Exchange distribution lists — built for onboarding subcontractor
and one-off batches (originally replacing `WellerUserImport_2025-11-21_v1.ps1`).

## What it does

- GUI grid for queuing up multiple users at once (add manually, paste a
  list of names, upload a CSV, or import directly from a Workday-style HR
  export)
- Bulk-apply License / Country / Job title / Groups to multiple selected
  rows in one action
- Validates License (`F3`, `F3+`, `E3`) and Country (`US`, `IN`) before
  anything is sent to Graph
- Optional per-row UPN override (auto-generates `First.Last@corrohealth.com`
  if left blank)
- Creates or re-enables users, assigns the correct license SKU(s), and adds
  them to the specified group(s)/distribution list(s)
- Writes a timestamped transcript and a results CSV (per-user status +
  generated password) next to the script after each run

## Prerequisites

- PowerShell 5.1 or later
- The following modules installed:
  - `Microsoft.Graph` (uses `Get-MgUser`, `New-MgUser`, `Update-MgUser`,
    `Set-MgUserLicense`, `Get-MgGroup`, `New-MgGroupMember`,
    `Get-MgSubscribedSku`)
  - `ExchangeOnlineManagement` (uses `Add-DistributionGroupMember`)
- An account with sufficient Graph scopes: `User.ReadWrite.All`,
  `Group.ReadWrite.All`, plus Exchange Online admin rights for
  distribution list membership

## Running it

```powershell
./UserProvisioningWizard.ps1
```

You'll be prompted to sign in to Microsoft Graph and Exchange Online, then
the wizard window opens.

1. Add users via **Add row**, **User creation** (paste names), **Upload
   CSV**, or **Import HR CSV** (header-matching, tolerant of Workday-style
   export layouts with metadata rows above the real header)
2. Fill in or bulk-apply License, Job title, Country, and Groups
3. Check the rows you want to process, then **Run import**
4. Review the Status column and the exported results CSV once it finishes

## CSV import format

`Upload CSV` expects: `UserPrincipalName, FirstName, LastName, Designation,
LicenseCode, UserCountry, InternalEmailOnly`. See `UserImportTemplate.csv`
for the expected layout and two example rows.

`Import HR CSV` doesn't require exact column positions — it searches the
first 15 rows of the file for a header row containing "First Name" and
"Last Name" labels and pulls from whichever columns match those names.
Job title, license, and groups aren't provided by HR exports and need to
be filled in manually after import.

## Output files (not committed to this repo)

Each run produces, next to the script:

- `UserProvisioning_<timestamp>.txt` — console transcript
- `UserProvisioning_Results_<timestamp>.csv` — per-user results, including
  the generated password for each new account

These contain personal data and plaintext temporary credentials. They are
excluded via `.gitignore` and should never be committed. Treat them as
sensitive — move or delete them once credentials have been handed off.

## Known limitations

- License codes are limited to `F3`, `F3+`, `E3` — anything else is
  rejected before the import runs
- Country is limited to `US` / `IN`; anything else defaults to `US` on
  import
- `Import HR CSV` depends on the header labels "First Name" / "Last Name"
  appearing somewhere in the first 15 rows — a differently labeled export
  (e.g. "Given Name") won't be auto-detected and needs manual conversion
