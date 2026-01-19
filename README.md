# GSMailer

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![License](https://img.shields.io/badge/license-MIT-green)

GSMailer is a production-ready Google Apps Script library that automates
event-based email workflows directly from Google Sheets using Google Docs templates.

It is designed for event organizers, NGOs, and teams who want to send personalized emails from Sheets easily and reliably.

---

## Features

- Send automated emails based on row status in Google Sheets
- Use Google Docs as HTML templates with dynamic variables
- Supports batch sending and dry run mode
- Optional registration number generation
- Optional Email attachements via Google drive URL
- Event context variables for common info (like event name, date)
- Logs email activity in a dedicated LOGS sheet
- Preview emails before sending
- Supports CC and BCC recipients

---

## Installation

1. Open your Google Sheet
2. Go to Extensions → Apps Script
3. Click + New Project or use an existing one
4. Add GSMailer as a library:
   - Library ID: 1PUNerdrrA-ZaUAegdlgkMvbmyJW_yLRrIHEg_wy-h69HE11Fd0mU_SeQ
   - Identifier: GSMailer
   - Version: 2.0.0  (Select latest)
5. Save and reload your Google Sheet

---

## Basic Usage

```javascript
function onOpen() {
  GSMailer.onOpen();
}

/* ===== Menu proxies ===== */
function initialSetup() {
  GSMailer.initialSetup();
}

function validateSetup() {
  GSMailer.validateSetup();
}

function sendEmailsLive() {
  GSMailer.sendEmailsLive();
}

function sendEmailsDryRun() {
  GSMailer.sendEmailsDryRun();
}

function previewSelectedRow() {
  GSMailer.previewSelectedRow();
}

function generateRegistrationNumbers() {
  GSMailer.generateRegistrationNumbers();
}

```

---

## Settings Sheet (SETTINGS)

GSMailer requires a sheet named SETTINGS with three columns: Key, Value, Help.

Key | Example Value | Description
--- | ------------- | -----------
DATA_SHEET | Attendees | Main sheet containing the rows of participants
STATUS_COLUMN | Status | Column that determines which email rule applies
EMAIL_COLUMN | Email | Column containing recipient email addresses
SENT_FLAG_COLUMN | Email Sent | Column marking emails as already sent
LAST_SENT_COLUMN | Last Sent At | Timestamp of last email sent
ERROR_COLUMN | Error | Column to store error messages
BATCH_LIMIT | 10 | Maximum number of emails per batch
REG_NUMBER_ENABLED | no | Enable registration number generation (yes/no)
REG_NUMBER_COLUMN | Registration No | Column to store registration numbers
REG_NUMBER_PATTERN | EVT-{{NUMBER}} | Pattern for registration numbers
REG_NUMBER_START | 1001 | Starting number if column is empty

---

## Email Rules Sheet (EMAIL_RULES)

Status | Subject | Template Doc URL | CC | BCC | Attachment URLs
------ | ------- | ---------------- | -- | --- | ---
Invited | Invitation to {{EventName}} | Google Doc URL | optional | optional | optional Google drive file url

Template variables use {{Variable}} syntax and are replaced using row data or EVENT_CONTEXT values.

---

## Event Context Sheet (EVENT_CONTEXT)

Key | Value
--- | -----
EventName | Annual Meetup 2026
EventDate | 2026-02-15

---

## Logs Sheet (LOGS)

Automatically tracks:
- Timestamp
- Row number
- Email
- Status
- Result (SENT / DRY RUN / ERROR)

---

## Changelog

v2.0.0
- Initial production-ready public release
- Batch email sending
- HTML template support
- Registration number generation
- Dry run and preview mode

---

## License

MIT License © 2026 Rohit Verma
