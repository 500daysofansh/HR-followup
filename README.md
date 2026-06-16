# 🤖 HR Candidate Follow-Up Automation

> Automated candidate communication system built for the **JCF Operations & Automation Internship Assignment**.  
> Eliminates manual follow-up emails by detecting overdue candidates and sending personalised, status-aware messages automatically.

---

## 📋 Table of Contents

- [Overview](#-overview)
- [Features](#-features)
- [Project Structure](#-project-structure)
- [Quick Start](#-quick-start)
- [Option A — Python Script](#option-a--python-script-local)
- [Option B — Google Apps Script](#option-b--google-apps-script-cloud)
- [Candidate Tracker Schema](#-candidate-tracker-schema)
- [Email Templates](#-email-templates)
- [Follow-Up Logic](#-follow-up-logic)
- [Environment Variables](#-environment-variables)
- [Tech Stack](#-tech-stack)

---

## 🧩 Overview

HR teams managing multiple candidates often miss follow-up emails due to manual tracking — leading to poor candidate experience and lost talent.

This automation:
1. Reads candidate data from an Excel / Google Sheet
2. Detects who is overdue for a follow-up based on their **status** and **days since last contact**
3. Sends a personalised, status-matched HTML email to each candidate
4. Updates the tracker with a timestamp and highlights the row green
5. Sends the HR team a daily digest of all follow-ups dispatched

---

## ✨ Features

| Feature | Python Script | Apps Script |
|---|---|---|
| Read candidate data | ✅ Excel (`.xlsx`) | ✅ Google Sheets |
| Status-aware email templates | ✅ | ✅ |
| Auto follow-up trigger (daily) | ✅ `--schedule` flag | ✅ Time-driven trigger |
| Tracker auto-update + timestamp | ✅ | ✅ |
| HR daily digest email | ❌ | ✅ |
| Dry-run / safe test mode | ✅ Default | ❌ |
| No server required | ✅ Local | ✅ Cloud |

---

## 📁 Project Structure

```
hr-followup-automation/
│
├── hr_followup_automation.py   # Python automation script
├── apps_script.js              # Google Apps Script (Sheets + Gmail)
├── candidate_tracker.xlsx      # Sample candidate tracker (15 entries)
├── generate_tracker.py         # Script that generated the sample tracker
├── JCF_HR_Automation_Workflow.docx  # Full assignment submission document
└── README.md                   # You are here
```

---

## 🚀 Quick Start

### Prerequisites

- Python 3.8+ **or** a Google account (for Apps Script)
- Gmail account with [App Password](https://myaccount.google.com/apppasswords) enabled (Python route)

---

## Option A — Python Script (Local)

### 1. Install dependencies

```bash
pip install pandas openpyxl schedule
```

### 2. Place files in the same folder

```
your-folder/
├── hr_followup_automation.py
└── candidate_tracker.xlsx
```

### 3. Run in dry-run mode (safe — no emails sent)

```bash
python hr_followup_automation.py
```

You'll see output like:
```
2026-05-02 09:00:01 [INFO] Loaded 15 candidates
2026-05-02 09:00:01 [INFO] 4 candidate(s) need a follow-up today
2026-05-02 09:00:01 [INFO] [DRY-RUN] Would email → aarav.mehta@gmail.com | Subject: Update on Your Application...
```

### 4. Send real emails

Set your Gmail credentials as environment variables:

**Windows:**
```cmd
set HR_EMAIL=your@gmail.com
set HR_EMAIL_PASS=your-16-char-app-password
python hr_followup_automation.py --send
```

**Mac / Linux:**
```bash
export HR_EMAIL="your@gmail.com"
export HR_EMAIL_PASS="your-16-char-app-password"
python hr_followup_automation.py --send
```

### 5. Schedule daily runs

```bash
python hr_followup_automation.py --send --schedule
```

Runs automatically every day at **09:00 AM**. Keep the terminal open (or run as a background service).

---

## Option B — Google Apps Script (Cloud)

> ✅ Recommended for submission — runs in the cloud 24/7, no laptop required.

### 1. Upload the tracker to Google Drive

- Go to [drive.google.com](https://drive.google.com)
- Upload `candidate_tracker.xlsx`
- Right-click → **Open with Google Sheets**

### 2. Open Apps Script

In your Google Sheet: **Extensions → Apps Script**

### 3. Paste the script

- Delete any existing code in the editor
- Copy the full contents of `apps_script.js` and paste it in
- Click **Save** (Ctrl+S)

### 4. Set the daily trigger

In the Apps Script editor, run:
```
createDailyTrigger()
```
(Click the function dropdown → select `createDailyTrigger` → click ▶ Run)

Grant permissions when prompted (Sheets + Gmail access).

### 5. Done ✅

The script now runs every day at **09:00 AM** automatically. After each run:
- Followed-up rows turn **green** in the sheet
- The `Follow-Up Sent` column is stamped with the date/time
- You receive a summary email with the count of follow-ups sent

---

## 📊 Candidate Tracker Schema

| Column | Field | Description |
|--------|-------|-------------|
| A | Name | Full name of candidate |
| B | Email | Contact email address |
| C | Phone | Phone number |
| D | Role Applied | Position they applied for |
| E | Status | Current stage (see below) |
| F | Last Contact Date | Date of last interaction (auto-updated) |
| G | Follow-Up Sent | Timestamp of last automated email |
| H | Notes | Manual HR notes |

### Valid Status Values

| Status | Follow-Up Triggered? | After (Days) |
|--------|---------------------|--------------|
| `Applied` | ✅ Yes | 2 days |
| `Shortlisted` | ✅ Yes | 3 days |
| `Interview Scheduled` | ✅ Yes | 1 day |
| `Under Review` | ✅ Yes | 4 days |
| `Hired` | ❌ No | — |
| `Rejected` | ❌ No | — |
| `Withdrew` | ❌ No | — |
| `On Hold` | ❌ No | — |

---

## 📧 Email Templates

The script automatically selects the right email based on the candidate's `Status`:

<details>
<summary><strong>Applied</strong> — Application received confirmation</summary>

> Subject: `Update on Your Application – [Role] at JCF`  
> Content: Confirms receipt, mentions review timeline of 3–5 business days.

</details>

<details>
<summary><strong>Shortlisted</strong> — Shortlisting congratulations</summary>

> Subject: `Great News! You've Been Shortlisted – [Role] at JCF`  
> Content: Congratulates candidate, previews next steps.

</details>

<details>
<summary><strong>Interview Scheduled</strong> — Interview reminder</summary>

> Subject: `Interview Reminder – [Role] at JCF`  
> Content: Friendly reminder, requests availability confirmation.

</details>

<details>
<summary><strong>Under Review</strong> — Status update</summary>

> Subject: `Application Status Update – [Role] at JCF`  
> Content: Acknowledges patience, confirms still in process.

</details>

---

## ⚙️ Follow-Up Logic

```
For each candidate row:
  IF status IN [Hired, Rejected, Withdrew, On Hold]  → SKIP
  IF email is missing                                → SKIP
  IF already followed up today                       → SKIP
  IF days_since_last_contact >= threshold[status]    → SEND EMAIL
                                                       UPDATE tracker
```

The script is **idempotent** — running it multiple times on the same day will not send duplicate emails.

---

## 🔐 Environment Variables

| Variable | Description | Example |
|----------|-------------|---------|
| `HR_EMAIL` | Your Gmail address | `hr@yourorg.com` |
| `HR_EMAIL_PASS` | Gmail App Password (not your login) | `abcd efgh ijkl mnop` |

> **How to generate a Gmail App Password:**  
> Google Account → Security → 2-Step Verification → App passwords → Create

---

## 🛠 Tech Stack

| Tool | Purpose |
|------|---------|
| `Python 3.8+` | Core scripting language |
| `pandas` | Reading and filtering Excel data |
| `openpyxl` | Excel formatting and timestamp updates |
| `smtplib` | Sending emails via Gmail SMTP |
| `schedule` | Daily trigger without a cron job |
| `Google Sheets` | Cloud-based candidate database |
| `Google Apps Script` | Automation + Gmail integration |
| `GmailApp` | Sending emails from Apps Script |

---

## 📝 Assignment Notes

**Problem solved:** HR teams lose track of candidate follow-ups when managing pipelines manually, causing delayed communication and poor candidate experience.

**How it improves efficiency:** Zero manual effort once set up — the script runs daily and handles all follow-ups automatically. HR only needs to update the `Status` field. Personalised templates ensure each candidate gets a relevant, professional message. A daily digest keeps HR informed without needing to monitor the sheet.

---

*Built for JCF Operations & Automation Internship Assignment — submitted May 2026*
