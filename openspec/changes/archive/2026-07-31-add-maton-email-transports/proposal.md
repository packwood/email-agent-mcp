# Change: Add Maton email transports

## Why

Production agents need Gmail and Microsoft email access without depending on locally persisted OAuth tokens or the macOS Keychain. Maton already manages the provider OAuth connections and exposes the native Gmail and Microsoft Graph APIs behind one authenticated gateway.

## What Changes

- Add explicit Maton transports for Gmail and Microsoft providers.
- Require an explicit active Maton connection for every configured mailbox.
- Preserve the existing native OAuth transports for other deployments, while allowing production wrappers to require Maton without fallback.
- Fail closed on missing, inactive, duplicate, mismatched, or untrusted Maton connection and pagination data.

## Impact

- Affected specs: `provider-gmail`, `provider-microsoft`.
- Affected code: provider clients and email MCP server startup/configuration.
- Production deployment will use Maton exclusively; native OAuth remains an opt-in upstream capability.
