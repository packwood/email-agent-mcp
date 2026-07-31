## ADDED Requirements

### Requirement: Maton Gmail Transport

The system SHALL support Gmail through the Maton gateway when the Maton transport is explicitly selected, and SHALL require an explicit active connection for every configured Gmail mailbox.

#### Scenario: Maton Gmail mailbox starts
- **WHEN** the Maton Gmail transport is selected with an API key and exactly one active connection matching the configured mailbox
- **THEN** Gmail API requests are sent through the Maton gateway with that connection ID

#### Scenario: Maton Gmail connection is invalid
- **WHEN** a configured Gmail mailbox has no active connection, multiple matching connections, or a mismatched application connection
- **THEN** that mailbox fails closed without falling back to native OAuth or another Maton connection

### Requirement: Trusted Maton Gmail Pagination

The system SHALL follow Gmail pagination only when the next URL remains HTTPS on the configured Maton gateway and within the Gmail API path prefix.

#### Scenario: Untrusted Gmail pagination URL
- **WHEN** a Maton Gmail response supplies a next URL outside the trusted gateway and Gmail path prefix
- **THEN** the request fails without transmitting credentials to that URL
