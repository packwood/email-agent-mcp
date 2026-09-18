---
epic: Email Operations
feature: Attachment Handling
---

## Purpose

Defines attachment operations: downloading from inbound emails, attaching to outbound emails, embedded image handling, binary file detection, filename sanitization, and size/type validation. Outbound attachments are file-based only (no CID embedding in v1).
## Requirements
### Requirement: Download Attachments

The system SHALL download attachments from inbound emails using provider-specific retrieval and return metadata (filename, MIME type, size, contentId).

#### Scenario: List attachments
- **WHEN** `read_email` returns an email with attachments
- **THEN** each attachment includes `{id, filename, mimeType, size, contentId, isInline}`

### Requirement: Inline Image Handling

The system SHALL preserve CID references (`<img src="cid:...">`) in HTML email bodies as markdown image links during content transformation. The agent can correlate CID values with attachment metadata from `list_attachments`. Full resolution to attachment content is planned for a future phase.

#### Scenario: Embedded image in HTML body
- **WHEN** an email body contains `<img src="cid:image001">`
- **THEN** the content engine converts it to `![](cid:image001)` in the markdown output
- **AND** the agent can look up `contentId: "image001"` via `list_attachments`

### Requirement: Attach Files to Outbound

The system SHALL accept file attachments for outbound emails from a local file path or buffer. No CID embedding in v1 — email clients will show attachments inline if appropriate.

#### Scenario: Attach file to reply
- **WHEN** `reply_to_email` is called with `{attachments: [{path: "/tmp/report.pdf"}]}`
- **THEN** the system base64-encodes the file and includes it as a Graph/Gmail attachment

### Requirement: Size and Type Validation

The system SHALL validate attachment size (default max 25MB) and MIME type against a configurable allowlist. Reject oversized or disallowed types with clear errors.

#### Scenario: Oversized attachment rejected
- **WHEN** an attachment exceeds 25MB
- **THEN** the system returns an error: "Attachment exceeds maximum size of 25MB"

### Requirement: Binary File Detection

The system SHALL validate file content by checking actual bytes (null byte check + magic byte signatures), not declared content type alone. This prevents binary-as-text hallucination.

#### Scenario: MIME type detected from bytes
- **WHEN** an attachment declares `contentType: text/plain` but contains JPEG magic bytes
- **THEN** the system detects the true type as `image/jpeg` and handles accordingly

### Requirement: Filename Sanitization

The system SHALL sanitize filenames that it reports for inbound attachments (`list_attachments`, `download_attachment`) to a safe ASCII form — special characters such as spaces, parentheses, and non-ASCII replaced — preserving the file extension, because callers use that value to write the attachment to local storage. The provider's original name SHALL be returned alongside it as `original_filename`.

This storage-safe form SHALL NOT be applied to outbound attachment display names; see "Outbound Attachment Display Name".

#### Scenario: Special characters in filename
- **WHEN** an inbound attachment has filename "Term Sheet - Alexander Morgan (Draft).pdf"
- **THEN** the system reports a safe ASCII `filename` while preserving the `.pdf` extension
- **AND** reports the unmodified name as `original_filename`

### Requirement: Outbound Attachment Display Name

The system SHALL transmit an outbound attachment's filename to the provider as a display name, preserving ordinary document-name characters — spaces, parentheses, brackets, hyphens, periods, commas, apostrophes, ampersands, and non-ASCII letters. The display name is never used as a local filesystem path; the file itself is read through the sandboxed `path` reader, which is unaffected by this requirement.

The system SHALL still neutralize, by replacing with `_` or removing:
- directory components — only the final path segment is kept, and a name that is empty or consists only of dots becomes `attachment`;
- control characters (including CR, LF, NUL, and TAB), Unicode line/paragraph separators, bidirectional-override characters, and unpaired surrogates;
- `"` and `\`, which would break a MIME quoted-string, and the characters `/ < > : | ? *`, which a recipient on Windows could not save;
- trailing dots and spaces.

Display names SHALL be capped at 255 characters with the extension preserved.

Providers that serialize MIME themselves (Gmail) SHALL emit printable-ASCII names as a quoted-string, and SHALL encode names containing non-ASCII characters per RFC 2231 (`filename*`, using continuations so no header line exceeds the RFC 5322 line limit) with an RFC 2047 encoded-word `name` parameter.

#### Scenario: Business-document name is preserved
- **WHEN** a draft is created with an attachment named "Paxden NDA (Patty) (Silver Point) (Redline) (SP 2026-09-17 v01 vs PP 2026-09-18 v04).docx"
- **THEN** the Graph `fileAttachment.name` / MIME `filename` is exactly that string
- **AND** the MIME type is still detected from content and extension

#### Scenario: Hostile display name is neutralized
- **WHEN** an attachment `filename` override contains directory components, CR/LF, or a double quote
- **THEN** only the final path segment is used and the offending characters are replaced with `_`
- **AND** no additional MIME header or parameter is produced

