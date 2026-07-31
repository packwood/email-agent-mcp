# Design

## Decisions

1. Maton is a provider transport, not a second business-logic implementation. Existing provider mapping and email-core actions remain authoritative.
2. Every request includes an explicit `Maton-Connection` header. The runtime never relies on Maton's default/oldest connection selection.
3. Connection records are resolved case-insensitively by configured mailbox and validated before use.
4. Maton pagination URLs must remain HTTPS on the configured Maton host and within the provider-specific path prefix.
5. Production wrappers select both Maton transports explicitly and do not configure local OAuth credentials.
6. Potentially consequential email writes are not automatically retried when the outcome is ambiguous.

## Rollback

The production wrapper can be rolled back to its previous pinned runtime and manifest. No provider credentials are migrated or deleted by this change.
