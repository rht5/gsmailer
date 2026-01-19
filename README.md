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
   - Click Libraries (+)
   - Library ID: 1PUNerdrrA-ZaUAegdlgkMvbmyJW_yLRrIHEg_wy-h69HE11Fd0mU_SeQ
   - Identifier: GSMailer
   - Version: 2.0.0  (or select latest)
   - Click Add
5. Add code from the Basic usage section in the editor
6. Save, run onOpen() and Grant Permissions.
7. Reload your Google Sheet.
8. In Google Sheet. You will see a new menu: 📨 GSMailer. Go to GSMailer → Admin / Setup → Initial Setup 
9. Configure your SETTINGS, EMAIL_RULES, EVENT_CONTEXT as needed.

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

## Email Templates (Google Docs)

GSMailer uses Google Docs as the source for email content.
This allows event organizers and NGOs to design rich, formatted emails without writing any code.

Emails are sent as HTML, generated automatically from the Google Doc.

**How It Works**
1. You create a Google Doc as your email template
2. Insert placeholders using {{VariableName}}
3. GSMailer replaces placeholders with values from: Your main data sheet (row values) and The EVENT_CONTEXT sheet (global event values)
4. The document is converted to HTML and sent via Gmail

**Example**

```
Dear {{Name}},

Thank you for registering for {{EventName}}.

📅 Date: {{EventDate}}  
📍 Venue: {{EventVenue}}

Your registration number is {{RegistrationNumber}}.

We look forward to seeing you!

Warm regards,  
{{OrganizerName}}
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
