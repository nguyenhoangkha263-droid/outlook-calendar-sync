# Outlook Calendar Sync

## Overview

Outlook Calendar Sync exports events from Microsoft Outlook Desktop into an iCalendar (`.ics`) file and publishes the file through GitHub Pages.

The generated calendar can be subscribed from:

- iPhone
- Android
- Google Calendar
- Samsung Calendar
- Any calendar application supporting iCalendar subscriptions

The solution does not require Outlook Mobile, Microsoft Intune, Company Portal, Device Management, or Exchange configuration on personal devices.

---

## Architecture

```text
Outlook Desktop
        |
        v
calendar_export.py
        |
        v
calendar.ics
        |
        v
update.bat
        |
        v
GitHub Repository
        |
        v
GitHub Pages
        |
        v
Subscribed Calendar
        |
        +---- iPhone Calendar
        |
        +---- Google Calendar
        |
        +---- Samsung Calendar
```

---

## Features

- Export Outlook calendar events
- Export recurring meetings
- Preserve:
  - Subject
  - Start Time
  - End Time
  - Location
- Generate stable event UID per occurrence
- Publish automatically to GitHub Pages
- Auto refresh using Windows Task Scheduler
- Keep previous calendar history

---

## Export Window

Current export configuration:

```python
DAYS_BACK = 7
DAYS_AHEAD = 30
```

Meaning:

```text
7 days in the past
+
Today
+
30 days in the future
```

will be included in the generated calendar.

---

# Prerequisites

## Windows

Required:

- Windows 10 or newer
- Outlook Desktop installed
- Outlook configured and signed in
- Git installed
- Python 3.11 or newer

---

## Python Packages

Install:

```bash
pip install pywin32
```

---

# Setup Flow

A new user should complete the following steps:

```text
1. Fork repository
2. Clone repository
3. Install requirements
4. Configure Git
5. Verify Outlook export
6. Enable GitHub Pages
7. Configure automatic updates
8. Subscribe on mobile device
```

---

# 1. Fork Repository

Open the original repository.

Click:

```text
Fork
```

Create a copy under your own GitHub account.

Example:

```text
Original:
sybinh/outlook-calendar-sync

Fork:
johnsmith/outlook-calendar-sync
```

---

# 2. Clone Repository

Clone your fork:

```bash
git clone https://github.com/<your-account>/outlook-calendar-sync.git
```

Example:

```bash
git clone https://github.com/johnsmith/outlook-calendar-sync.git
```

---

# 3. Configure Git

Verify configuration:

```bash
git config --global user.name
git config --global user.email
```

Configure if necessary:

```bash
git config --global user.name "John Smith"
git config --global user.email "john.smith@example.com"
```

---

# 4. Configure Git Authentication

Verify that Git can push to your repository.

Recommended:

```text
Git Credential Manager
```

Verify:

```bash
git push
```

must complete without errors.

---

# 5. Verify Outlook Export

Run:

```bash
python calendar_export.py
```

Expected output:

```text
Calendar updated
ICS exported: calendar.ics
```

Verify:

```text
calendar.ics
```

contains your calendar meetings.

---

# Repository Structure

```text
outlook-calendar-sync/
│
├── calendar_export.py
├── calendar.ics
├── update.bat
├── setup.bat
└── README.md
```

---

# 6. Enable GitHub Pages

Open:

```text
Repository
    Settings
        Pages
```

Configuration:

```text
Source:
Deploy from a branch

Branch:
main

Folder:
/ (root)
```

Save configuration.

GitHub Pages will publish:

```text
https://<github-account>.github.io/outlook-calendar-sync/calendar.ics
```

Example:

```text
https://johnsmith.github.io/outlook-calendar-sync/calendar.ics
```

Verify that opening the URL downloads:

```text
calendar.ics
```

---

# 7. Manual Update

Run:

```bash
update.bat
```

The script will:

```text
1. Export Outlook calendar
2. Update calendar.ics
3. Stage calendar.ics
4. Commit changes
5. Push to GitHub
```

Expected output:

```text
Calendar updated and pushed successfully
```

or:

```text
No update to commit
```

---

# 8. Configure Automatic Updates

## Create Scheduled Task

Run:
```bash
setup.bat
```

```cmd
schtasks /create ^
 /tn "Outlook Calendar Sync" ^
 /sc minute ^
 /mo 15 ^
 /tr "cmd.exe /c \"C:\Path\To\update.bat\"" ^
 /f
```

Example:

```cmd
schtasks /create ^
 /tn "Outlook Calendar Sync" ^
 /sc minute ^
 /mo 15 ^
 /tr "cmd.exe /c \"D:\Projects\outlook-calendar-sync\update.bat\"" ^
 /f
```

---

## Verify Scheduled Task

List task:

```cmd
schtasks /query /tn "Outlook Calendar Sync"
```

Run immediately:

```cmd
schtasks /run /tn "Outlook Calendar Sync"
```

View detailed status:

```cmd
schtasks /query /tn "Outlook Calendar Sync" /v /fo list
```

Expected:

```text
Last Result: 0
```

---

## Delete Scheduled Task

```cmd
schtasks /delete /tn "Outlook Calendar Sync" /f
```

---

# iPhone Setup

Open:

```text
Settings
    Apps
        Calendar
            Calendar Accounts
                Add Account
                    Other
                        Add Subscribed Calendar
```

Enter:

```text
https://<github-account>.github.io/outlook-calendar-sync/calendar.ics
```

Tap:

```text
Next
```

Complete subscription.

---

## Verify iPhone Subscription

Open:

```text
Calendar
```

Tap:

```text
Calendars
```

Verify that the subscribed calendar is:

```text
Enabled
```

---

# Android Setup

Install:

```text
ICSx⁵
```

Open:

```text
ICSx⁵
    Add Subscription
```

Enter:

```text
https://<github-account>.github.io/outlook-calendar-sync/calendar.ics
```

Complete setup.

---

## Android Calendar Applications

After configuration the calendar can be viewed in:

```text
Google Calendar
Samsung Calendar
Business Calendar
aCalendar
```

depending on the configured provider.

---

# Updating Existing Installation

When repository changes are released:

```bash
git pull
```

Review changes.

Run:

```bash
python calendar_export.py
```

Verify successful export.

No additional configuration is normally required.

---

# Migration To Another PC

## Copy Repository

Clone your fork:

```bash
git clone <your-fork-url>
```

## Verify Outlook

Ensure Outlook Desktop is installed and signed in.

## Verify Export

Run:

```bash
python calendar_export.py
```

## Recreate Scheduled Task

Run:

```cmd
schtasks /create ...
```

using the local repository path.

Mobile devices do not need any modification as long as the GitHub Pages URL remains unchanged.

---

# Troubleshooting

## GitHub Pages returns 404

Verify:

- Repository is public
- GitHub Pages is enabled
- Branch configuration is correct
- Latest changes have been pushed

---

## Calendar Updates Are Not Visible On Mobile

Verify:

- Subscription is enabled
- GitHub Pages contains the latest file
- Mobile calendar refresh completed

If necessary:

```text
Remove subscription
Add subscription again
```

---

## Outlook Meetings Are Not Exported

Verify:

- Outlook Desktop is installed
- Outlook profile is signed in
- Outlook displays the expected meetings

Run:

```bash
python calendar_export.py
```

and inspect:

```text
calendar.ics
```

---

## Scheduled Task Does Not Run

Verify:

```cmd
schtasks /query /tn "Outlook Calendar Sync" /v /fo list
```

Check:

```text
Last Result: 0
```

If not:

```cmd
schtasks /run /tn "Outlook Calendar Sync"
```

and review console output.

---

## Git Push Fails

Verify:

```bash
git push
```

works manually.

Resolve Git authentication issues before using Task Scheduler.

---

# Security Considerations

The generated calendar may contain:

- Meeting subjects
- Meeting dates and times
- Meeting locations

Before using this solution, verify that publishing calendar information through GitHub Pages complies with your organization's security and data handling policies.

The exported calendar is read-only.

Updates made on mobile devices are not synchronized back to Outlook.
