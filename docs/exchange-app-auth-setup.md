# One-App Setup (Graph + Exchange + SharePoint)

This app now uses a single app registration path.

## Setup flow in the app (recommended)

On the Sign In page, do this in order:

1. Select or create the customer.
2. Enter **Tenant domain** (for example `contoso.onmicrosoft.com`).
3. Click **Set Up Customer App (One-Time)**.
4. Complete admin consent in the browser.
5. Click **Save Customer**.
6. Click **Connect All (One App)**.

Expected result: Graph, Exchange, and SharePoint all show as connected.

## What the setup button does

- Creates/reuses one Entra app registration for this customer.
- Generates a certificate on this machine and uploads the public cert to the app.
- Sets the app client ID + thumbprint in the profile fields.
- Opens the admin consent URL.

Generated cert files are saved to:

- `artifacts/customers/<CustomerName>/exchange`

## Tenant admin requirements

The customer tenant still must allow and consent required permissions.

Minimum important items:

- Exchange application permission `Exchange.ManageAsApp` is present and admin-consented.
- Required Graph and SharePoint permissions are consented.
- Service principal has the needed role/RBAC to run Exchange cmdlets (for quick validation, Exchange Administrator role is common).

## Common failures

- **Unauthorized / OperationStopped**
  - Admin consent or Exchange role assignment is missing.

- **Certificate not found**
  - Cert thumbprint does not match, cert was removed, or private key is not in `CurrentUser\My` on this machine.

- **Consent completed but still failing**
  - Wait a minute for tenant propagation, then click **Connect All (One App)** again.

