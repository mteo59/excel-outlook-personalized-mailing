# Personalized Mailing with Excel & Outlook (VBA)

A lightweight VBA macro that sends a **personalized HTML email campaign** through **Microsoft Outlook**, using a contact list stored in an **Excel** sheet.

Each recipient gets an email addressed to them by title and last name, with a call-to-action button containing a personalized link (based on their contact ID), followed by your Outlook signature, including its images.

---

## Features

- Personalized greeting for each contact (title + last name)
- HTML email body with a styled call-to-action button
- Personalized button link per contact (`?id=<contact ID>`)
- Automatically appends your Outlook signature, including images
- Choose which Outlook account sends the emails
- **Review mode**: open each email before sending, or send directly
- Skips rows with no email address

---

## Requirements

- Windows
- Microsoft Excel (with macros enabled)
- Microsoft Outlook (desktop app), with the sender account configured

> Outlook must be the classic desktop version: the "new Outlook" and Outlook on the web do not support VBA automation.

---

## Project structure

```
.
├── src/
│   └── PersonalizedMailing.bas    # VBA module to import into Excel
├── example/
│   └── contacts_example.xlsm      # Sample workbook with the macro and fake contacts
└── README.md
```

---

## Contact sheet format

The macro reads the **active sheet**. Row 1 contains the headers and data starts on row 2.

| Column | Header     | Description                             |
|--------|------------|-----------------------------------------|
| A      | ID         | Unique contact ID (appended to the link)  |
| B      | CIVILITY   | Title (e.g. M., Mme)                    |
| C      | NAME       | Last name                               |
| D      | FIRST_NAME | First name                              |
| E      | EMAIL      | Recipient email address                 |

Example:

| ID | CIVILITY | NAME   | FIRST_NAME | EMAIL                     |
|----|----------|--------|------------|---------------------------|
| 1  | M.       | Dupont | Jean       | jean.dupont@example.com   |
| 2  | Mme      | Martin | Claire     | claire.martin@example.com |

---

## Installation

**Quick start:** download `example/contacts_example.xlsm`, which already contains the macro, then go to *Configuration*.

**Or use your own workbook:**

1. Open your Excel workbook and save it as **Excel Macro-Enabled Workbook (`.xlsm`)**.
2. Press **Alt + F11** to open the VBA editor.
3. Go to **File > Import File...** and select `src/PersonalizedMailing.bas`.
4. Close the VBA editor.

---

## Configuration

Open the `PersonalizedMailing` module and edit the constants in the **CONFIGURATION** section:

```vba
Private Const SENDER_ADDRESS As String = "your.address@example.com"
Private Const SIGNATURE_FILE As String = "MySignature.htm"
Private Const EMAIL_SUBJECT As String = "Your email subject"
Private Const BUTTON_BASE_URL As String = "https://www.example.com/form?id="
Private Const BUTTON_TEXT As String = "Click here"
Private Const SEND_DIRECTLY As Boolean = False
```

| Constant          | Description                                                                 |
|-------------------|-----------------------------------------------------------------------------|
| `SENDER_ADDRESS`  | Outlook account used to send the emails. If not found, the default account is used. |
| `SIGNATURE_FILE`  | Name of your signature file in `%APPDATA%\Microsoft\Signatures\`            |
| `EMAIL_SUBJECT`   | Subject of the email                                                        |
| `BUTTON_BASE_URL` | Link of the button; the contact ID is added at the end                      |
| `BUTTON_TEXT`     | Text displayed on the button                                                |
| `SEND_DIRECTLY`   | `False` = open each email for review, `True` = send immediately             |

Then write your own message in the `BuildEmailBody` function by replacing the `[...]` placeholders.

> **Finding your signature file:** press **Win + R**, type `%APPDATA%\Microsoft\Signatures` and press Enter. Use the name of the `.htm` file matching your signature.

---

## Usage

1. Open the sheet containing your contacts.
2. Press **Alt + F8**, select `SendPersonalizedEmails` and click **Run**.
3. A message box appears when the campaign is finished.

**Recommended:** run a first test with `SEND_DIRECTLY = False` and a contact list containing only your own email address, to check the content, the link and the signature before a real campaign.

---

## Limitations

- Outlook may show a security prompt when a macro sends emails, depending on your antivirus and Trust Center settings.
- Emails are sent one by one through your Outlook account, so your email provider's sending limits apply. This tool is not designed for very large campaigns.
- Signature images are referenced by local file paths; always check with a test email that they display correctly for the recipient.

---

## Responsible use

Only send emails to people who have agreed to be contacted, or with whom you have a legitimate business relationship, and always comply with applicable regulations (e.g. **GDPR** in the EU). Include a simple way for recipients to opt out of future emails.

**Never commit real contact data to this repository.** Use fake data such as `example.com` addresses for samples and tests.
