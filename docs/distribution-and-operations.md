# Distribution and Operations

This guide explains how to distribute the app internally, how certificates work, and how to share customer configuration across colleagues.

## 1) Build a distributable package

From repo root:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\publish.ps1 -ZipOutput
```

Output:

- Published app folder: `dist/NewUserAutomation-win-x64`
- Optional zip: `dist/NewUserAutomation-win-x64.zip`

## 2) One-file installer script

Host these files on your internal web server:

- `NewUserAutomation-win-x64.zip`
- `scripts/Install-NewUserAutomation.ps1`

Colleagues install/update with one command:

```powershell
powershell -ExecutionPolicy Bypass -File .\Install-NewUserAutomation.ps1 -ZipUrl "https://<your-server>/NewUserAutomation-win-x64.zip"
```

This installs to `%LOCALAPPDATA%\NewUserAutomation\current` and creates Start Menu/Desktop shortcuts.

## 3) Certificate/auth model (current design)

The app uses app-only authentication with certificates:

- Graph app auth
- Exchange app auth
- SharePoint app auth (PnP)

Certificates are generated/used per machine/operator in `CurrentUser\My`.

Operational impact:

- Each operator machine needs one-time setup for each customer tenant.
- Admin consent and tenant role requirements still apply.

## 4) Share customer list across colleagues

Set environment variable `NEWUSERAUTOMATION_DATA_ROOT` to a shared path (for example, a network share).

Example:

```powershell
[Environment]::SetEnvironmentVariable("NEWUSERAUTOMATION_DATA_ROOT", "\\fileserver\IT\NewUserAutomationData", "User")
```

After restart, the app reads/writes:

- `customers\...` profile files
- `settings\customers.json`

under that shared root instead of local Documents.

## 5) Recommended pre-distribution checklist

- Confirm one-app customer setup works in a fresh user profile.
- Confirm personal SharePoint step creates/checks document library and applies group permissions.
- Validate install/update script on a colleague test machine.
- Keep hosted zip URL stable so install script doubles as updater.
