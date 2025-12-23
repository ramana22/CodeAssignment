# Word Redaction Add-in

React-based Word add-in that redacts sensitive information, inserts a confidentiality header, and enables Track Changes (when supported by Word API 1.5).

## Features
- Searches for emails, phone numbers, and SSNs across the document body.
- Replaces sensitive matches with redaction markers while preserving the rest of the document.
- Adds a `CONFIDENTIAL DOCUMENT` header to the primary header if it is missing.
- Enables Track Changes when Word JavaScript API requirement set 1.5 is available (shows a notice and continues without it otherwise).
- Clean React UI with no external CSS frameworks.

## Running locally
1. Install dependencies (requires internet access):
   ```bash
   npm install
   ```
2. Start the secured dev server on port 3000:
   ```bash
   npm start
   ```
3. Sideload `manifest.xml` into Word on the web or Word for Windows/Mac. The task pane loads from `https://localhost:3000`.

## How it works
- The add-in checks for `WordApi 1.5` support to decide whether Track Changes can be enabled.
- Document text is scanned with JavaScript regular expressions to collect sensitive matches. Each unique match is searched inside Word and replaced with a descriptive redaction marker.
- The primary header is populated with `CONFIDENTIAL DOCUMENT` when it is not already present.

## Testing
Use the provided `Document-To-Be-Redacted.docx` (or any document containing emails, phone numbers, and SSNs) and click **Redact & Track Changes**. Track Changes will log the header insertion and replacements whenever supported by the host.
