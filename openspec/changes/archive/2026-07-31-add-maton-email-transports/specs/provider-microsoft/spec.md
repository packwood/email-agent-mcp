## ADDED Requirements

### Requirement: Maton Microsoft Transport

The system SHALL support Microsoft Graph through the Maton gateway when the Maton transport is explicitly selected, and SHALL require an explicit active connection for every configured Microsoft mailbox.

#### Scenario: Maton Microsoft mailbox starts
- **WHEN** the Maton Microsoft transport is selected with an API key and exactly one active connection matching the configured mailbox
- **THEN** Microsoft Graph requests are sent through the Maton gateway with that connection ID

#### Scenario: Maton Microsoft connection is invalid
- **WHEN** a configured Microsoft mailbox has no active connection, multiple matching connections, or a mismatched application connection
- **THEN** that mailbox fails closed without falling back to MSAL, the local Keychain, or another Maton connection

### Requirement: Trusted Maton Microsoft Pagination

The system SHALL follow Graph pagination only when the next URL remains HTTPS on the configured Maton gateway and within the Outlook Graph path prefix.

#### Scenario: Untrusted Graph pagination URL
- **WHEN** a Maton Graph response supplies a next URL outside the trusted gateway and Outlook path prefix
- **THEN** the request fails without transmitting credentials to that URL
