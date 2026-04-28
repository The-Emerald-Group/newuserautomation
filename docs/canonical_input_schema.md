# Canonical Input Schema (v1)

This schema is the source contract for parser and validation logic.

## Field Definitions

- `FirstName`: string, required, scalar
- `LastName`: string, required, scalar
- `DisplayName`: string, required, scalar
- `PreferredUsername`: string, required, scalar, normalized lowercase
- `TemporaryPassword`: string, required, scalar
- `PrimaryEmail`: string, required, scalar, normalized lowercase, email format
- `SecondaryEmail`: string, optional, scalar, normalized lowercase, email format when provided
- `LicenseSkus`: array<string>, optional, multi-value
- `GroupAccess`: array<string>, optional, multi-value
- `SharedMailboxAccess`: array<string>, optional, multi-value
- `SharePointAccess`: array<string>, optional, multi-value
- `SpecialRequirements`: string, optional, scalar
- `RequestApprovedBy`: string, required, scalar

## Normalization Rules

- Trim leading and trailing whitespace.
- Convert smart quotes to ASCII equivalents.
- Convert em/en dashes to hyphen.
- Normalize list separators using split tokens: `,`, `;`, `|`, newline.
- Remove empty list entries and deduplicate case-insensitively.
- Lowercase identity fields where case-insensitive matching is expected (`PreferredUsername`, emails).

## Required Data For Execution

Execution is blocked when any of these are missing or invalid:

- `FirstName`
- `LastName`
- `PrimaryEmail`
- `RequestApprovedBy`
