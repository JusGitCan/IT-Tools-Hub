# CorroHealth User Provisioning Wizard

A Windows desktop tool for creating Microsoft 365 / Entra ID (Azure AD) user
accounts, assigning licenses, and adding group and distribution-list
memberships. It supports both **Domestic (US)** and **Global / India**
onboarding from a single interface.

---

## Quick start

> ### ⚠ Always launch with the `.bat` — never the `.ps1`
> Double-click **`Launch-UserProvisioningWizard.bat`** to start the tool.
> **Do not** right-click the `.ps1` and "Run with PowerShell," and do not open
> the `.ps1` directly. Doing so can start PowerShell in the wrong threading mode
> and produce this error at the Exchange sign-in step:
>
> > *ActiveX control ... cannot be instantiated because the current thread is
> > not in a single-threaded apartment.*
>
> The `.bat` starts PowerShell in the required STA mode and avoids this
> entirely. (The script also tries to self-correct if started the wrong way,
> but the `.bat` is the reliable entry point — use it every time.)

1. Download all the files in this package and keep them **together in one
   folder**.
2. Double-click **`Launch-UserProvisioningWizard.bat`**.
3. On first launch, Windows may show a SmartScreen prompt because the files
   were downloaded from the internet. Click **More info → Run anyway**. This
   happens once per machine.
4. Sign in to Microsoft Graph and Exchange Online when prompted.
5. The wizard window opens.

---

## What's in this package

| File | Purpose |
|------|---------|
| `UserProvisioningWizard.ps1` | The wizard itself. |
| `Launch-UserProvisioningWizard.bat` | Double-click launcher. Keep it in the same folder as the `.ps1`. |
| `UserImportTemplate_Domestic.csv` | Blank template for bulk-importing domestic (US) users. |
| `README.md` | This file. |

---

## Prerequisites

- **Windows** with **Windows PowerShell 5.1** (built in) or PowerShell 7.
- The following PowerShell modules installed for the account that runs it:
  - `Microsoft.Graph`
  - `ExchangeOnlineManagement`
- An account with rights to create users and assign licenses:
  Graph scopes `User.ReadWrite.All` and `Group.ReadWrite.All`, plus Exchange
  Online administrator rights for distribution-list membership.

To install the modules (run PowerShell as administrator once):

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

---

## Using the wizard

### Modes

A toggle at the top right switches between two modes. **The mode changes both
the columns shown and the behavior.**

- **Domestic** — License type (`F3` / `F3+` / `E3`), Country locked to `US`.
  India-specific logic is skipped entirely.
- **Global / India** — Mailbox size (`2GB` / `50GB` / `E3`) plus EntApps,
  Country `US` or `IN`, subcontractor group routing, mailbox archiving and
  retention for 50 GB mailboxes, and the extra profile columns (City,
  Province, Office).

Switching modes hides or shows columns but does **not** delete data you've
already entered. Importing a file that contains any India row automatically
switches to Global mode.

### Adding users

- **Add row** — add one blank row.
- **User creation** — paste a list of names (one `First Last` per line).
- **Upload CSV** — load the standard template (see below).
- **Remove selected** / the per-row **Remove** button — delete rows.
- **Apply to selected** — stamp a license, country, job title, or groups
  onto every checked row at once.

### License reference (Global mode)

Mailbox size maps to licenses as follows:

| Mailbox size | License granted |
|--------------|-----------------|
| `2GB` | F3 |
| `50GB` | F3 + Exchange Archive (a.k.a. "F3+") |
| `E3` | E3 (full desktop Office) |
| `EntApps = Y` | Adds Microsoft 365 Apps for Enterprise on top of any of the above |

Domestic mode uses the license names (`F3` / `F3+` / `E3`) directly; they map
to the same underlying licenses.

### Groups

The **Groups** column accepts a comma-separated list, so a user can be added
to multiple groups at once, e.g. `Internal Email Only, VPN Users`. In Global
mode, the standard groups (based on license and subcontractor status) are added
automatically, and anything you type is added on top.

---

## Import CSV format (standard template)

Use **`UserImportTemplate_Domestic.csv`** as a starting point for domestic
imports. Columns:

- `UserPrincipalName` — leave blank to auto-generate `First.Last@corrohealth.com`,
  or fill in to override (useful for name collisions).
- `FirstName`, `LastName` — required.
- `Designation` — job title (optional).
- `LicenseCode` — `F3`, `F3+`, or `E3`.
- `UserCountry` — `US`.
- `InternalEmailOnly` — `Y` or `N`.

> HR / Workday exports (`.xlsx`, or multi-sheet workbooks) are **not** loaded
> directly. Convert them to this standard template first, then use **Upload
> CSV**.

---

## Safety behavior

The tool is deliberately conservative:

- **Existing users are never modified.** If a UPN already exists, it is skipped
  and marked in the Status column, noting whether the existing account is
  enabled or disabled — so it can be handled manually. No password is reset and
  no account is re-enabled.
- **Invalid rows are skipped individually** (bad license/size/country, missing
  name). The rest of the batch still runs.
- A **group that can't be found** or a **license that fails to assign** does not
  fail the user — the account is still created and the issue is noted in Status.
- **Passwords** are auto-generated uniquely per user (16 characters), or a
  single shared password can be chosen at launch. Users must change the
  password at first sign-in.
- After creating users, the tool **waits for each mailbox to provision**
  (polling, with a 5-minute safety timeout) before adding groups or enabling
  archiving, rather than a fixed blind wait.

---

## Output files

Each run writes two timestamped files next to the script:

- `UserProvisioning_<timestamp>.txt` — a transcript of the run.
- `UserProvisioning_Results_<timestamp>.csv` — one row per user with the
  generated password and the per-user status.

> **The results CSV contains plaintext temporary passwords.** Treat it as
> sensitive: move or delete it once credentials have been handed off, and do
> not commit it to source control.

---

## Notes and limitations

- "Groups" refers to Entra ID / Exchange distribution membership, **not** Azure
  Resource Groups.
- Group display names, the retention policy name (`India F3 Users`), and the
  subcontractor mailbox attribute are taken from the existing onboarding
  process. If any of those names change in the tenant, update them in the
  script.
- The mailbox-provisioning timeout is 5 minutes per user. For very large
  batches, run in smaller groups.

---

## Troubleshooting

**"ActiveX control ... cannot be instantiated because the current thread is not
in a single-threaded apartment."**
The tool was started in the wrong PowerShell threading mode. This almost always
means it was launched by running the `.ps1` directly instead of the `.bat`.
Fix: close the error and launch with **`Launch-UserProvisioningWizard.bat`**.

**"Running scripts is disabled on this system."**
An execution-policy restriction. The `.bat` already bypasses this for its own
launch, so use the `.bat`. (If you must run the `.ps1` some other way, start it
with `powershell.exe -STA -ExecutionPolicy Bypass -File "<path>"`.)

**A console window flashes and disappears, then the wizard opens.**
Expected. If the tool is ever started in the wrong mode, it relaunches itself
correctly — that brief flash is the relaunch.

**SmartScreen: "Windows protected your PC."**
Because the files were downloaded from the internet. Click **More info → Run
anyway**. This happens once per machine and is normal for any downloaded tool.

**Graph or Exchange sign-in fails.**
Confirm the account has the required permissions (see Prerequisites) and that
the `Microsoft.Graph` and `ExchangeOnlineManagement` modules are installed.
