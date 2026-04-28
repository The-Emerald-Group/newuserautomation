# Audit Format Decision (v1)

## Chosen Storage

Append-only `JSONL` file per environment.

Rationale:
- Immutable-by-convention append model.
- Human-readable and easy to diff.
- No additional runtime dependencies for Phase 1.

## Record Shape

```json
{
  "RunId": "string-guid",
  "TimestampUtc": "ISO-8601",
  "OperatorUpn": "string",
  "InputHash": "sha256-string",
  "PlannedActions": ["CreateUser", "AssignLicense"],
  "Outcomes": ["CreateUser:Succeeded", "AssignLicense:Succeeded"]
}
```

## Rules

- Never rewrite prior lines.
- One run emits at least one record.
- Record includes operator identity and command intent.
