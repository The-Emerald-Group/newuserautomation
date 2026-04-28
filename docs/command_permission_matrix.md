# Command-To-Permission Matrix (v1)

This matrix drives preflight checks before execution.

| Action | Graph Delegated Scope | Exchange Role | PnP Right |
|---|---|---|---|
| `CreateUser` | `User.ReadWrite.All` | n/a | n/a |
| `AssignLicense` | `User.ReadWrite.All` | n/a | n/a |
| `AddGroupMembership` | `Group.ReadWrite.All` | n/a | n/a |
| `GrantMailboxAccess` | n/a | `Mailbox.Permission.Assign` | n/a |
| `GrantSharePointAccess` | n/a | n/a | `Site.Member.Write` |

## Policy

- Preflight evaluates only selected actions.
- Missing required scope/role/right blocks execution.
- Failure output must include action name and exact missing permission string.
