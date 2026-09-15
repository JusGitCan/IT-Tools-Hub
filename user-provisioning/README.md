# CorroHealth User Provisioning Wizard

A WinForms wizard for creating Entra ID users, assigning Microsoft 365
licenses, and adding group / distribution-list memberships — for both
Domestic (US) and Global/India onboarding.

## Two modes
- **Domestic** — License (F3 / F3+ / E3 / E5), Country locked to US, no
  India-specific logic.
- **Global/India** — Mailbox size (2GB / 50GB / E3) + EntApps, Country
  US/IN, subcontractor group routing, and archiving/retention for 50GB
  mailboxes.

Switching modes hides/shows columns but never deletes row data. Importing
a file with any India row auto-switches to Global.

## License mapping
Same base SKU, mode-specific add-on for the "+" tier:

| Tier | Domestic | Global |
|---|---|---|
| F3 (2GB) | SPE_F1 | SPE_F1 |
| F3+ (50GB) | SPE_F1 + Apps for Enterprise | SPE_F1 + Exchange Archive |
| E3 | SPE_E3 | SPE_E3 |
| E5 | SPE_E5 | — (Domestic-only) |
| EntApps | — | Adds Apps for Enterprise on top of any tier above |

The "+" tier deliberately means different things in each mode: desktop
Office in Domestic, mailbox archiving in Global.

## Bringing in users
- **Manual entry** — type directly into the grid, including a Manager
  field (bulk-apply available: type a manager's UPN, apply to selected
  rows).
- **Upload CSV** — see `UserImportTemplate_Domestic.csv` in this bundle
  for the expected column headers.
- **Upload HR Export** — reads Workday exports directly (`.xlsx`), mapping
  by column header text rather than fixed position, so it tolerates
  Workday's export layout drifting between pulls. Handles `*`
  required-field markers, non-breaking spaces, and an optional "Example"
  sample row.
- Paste directly into a selected cell (Ctrl+V) without needing to
  double-click into edit mode first.

## Bulk-apply panels
Small captioned controls above the grid let you apply one License,
Country, Job title, or Groups value to every checked row at once — this
edits the grid before import, not existing Azure users.

## External tenant invite (Domestic mode)
After each Domestic-mode batch, every new user is also invited as a B2B
guest into the external tenant — same UPN, no invite email, name/title
copied over, no groups. A tenant-context check runs first to catch a
silent wrong-tenant sign-in before it causes a no-op run.

## What it never does
- Existing users are **always skipped, never modified** — no password
  reset, no re-enable. Status notes whether the existing account is
  enabled or disabled, for manual handling.
- Invalid rows (bad license/size/country, missing name) are skipped
  individually — the rest of the batch still runs.
- A group-not-found or license failure doesn't fail the user — the
  account is still created and the problem is noted in Status.
- M365/security groups route through Microsoft Graph; only true
  mail-enabled distribution lists use the Exchange cmdlet (fixed
  misclassification that used to reject M365 groups).

## Passwords
Three modes: auto-generate a unique 16-character password per user
(default), one shared password, or a fixed prefix + random 5-digit
suffix (e.g. `CorroHealth@@34961`). Forced password-change at first
sign-in is off.

## How to run it
Always launch via **Launch-UserProvisioningWizard.bat**, never by
double-clicking or right-clicking the `.ps1` directly, and never via
VS Code's F5 (that host runs MTA, not STA). The `.bat` forces Windows
PowerShell into STA mode, which the WinForms UI and the Exchange
Online sign-in control both require.

Keep the `.bat` and `.ps1` in the same folder.

## Sign-in
Prompts for Microsoft Graph sign-in, and Exchange Online sign-in (needed
for true distribution-list membership and mailbox polling), using your
own admin credentials.

## Practical limits
No hard user-count limit, but a practical ceiling around ~50 per run, due
to the synchronous UI and Microsoft Graph throttling on larger batches.

## Results
A per-run results CSV (per-user password + status — treat it as
sensitive) and PowerShell transcript are saved next to the script after
each run. After creating users, the wizard polls until each mailbox
exists (5-minute timeout) before adding groups or archiving, instead of
a fixed blind wait.
