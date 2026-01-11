# Backup Script

Simple PowerShell scripts to back up a notes folder (in my case notes taken in Obsidian) to a ZIP file and send an email with the backup status.

Designed to run automatically (monthly) using Windows Task Scheduler.

---

## Setup

### 1. Create SMTP credentials (required)

Run the credential script **once**.  
This stores your email credentials encrypted for your Windows user.

```powershell
powershell.exe -ExecutionPolicy Bypass -File create-smtp-credentials.ps1
```

This will prompt for:

- Email address
- Email password (or app password)

It creates an encrypted credential file

### 2. Create config file

Copy the example config:

config.example.json -> config.json


Edit config.json and set the variables.

### 3. Run backup manually
```powershell
powershell.exe `
  -ExecutionPolicy Bypass `
  -File backup-notes.ps1 `
  -ConfigPath "C:\path\to\config.json"
```

### Scheduling

Use Windows Task Scheduler to run the backup script on the 1st of every month.

Program:

powershell.exe


Arguments:

-ExecutionPolicy Bypass -File backup-notes.ps1 -ConfigPath "C:\path\to\config.json"


Run as the same Windows user that created the SMTP credentials.

# Notes

Backups are ZIP files with date in the filename

Email is sent on success and on failure

If you want to use it for **Gmail** and have 2 factor authentication enabled, you should generate an app password and use that password instead!

# Testing Manually
With an example folder `C:\Repositories\NotesBackup`
1. open powershell
2. cd "C:\Repositories\NotesBackup"
3. run
```
powershell.exe `
  -ExecutionPolicy Bypass `
  -File .\backup-notes.ps1 `
  -ConfigPath "C:\Repositories\NotesBackup\config.json"
  ```

# Windows Task Scheduler

- Click Create Task… (not “Basic Task”)

## General tab

- Name: NotesBackup

- ✅ Run whether user is logged on or not

- ✅ Run with highest privileges

- Configure for: Windows 10 / Windows 11


## Triggers tab

- Click New…

- Begin the task: On a schedule

- Settings: Monthly

- Months: All months

- Days: 1

- Set a time (e.g. 12:00)

- ✅ Enabled


## Actions tab

- Click New…

- Action: Start a program

- Program/script:

- search for `powershell.exe`


- Add arguments:
Use the real paths, example:

-NoProfile -ExecutionPolicy Bypass -File "C:\Repositories\NotesBackup\backup-notes.ps1" -ConfigPath "C:\Users\My\.notes-backup\config.json"


- Start in:

C:\Repositories\ObsidianNotesBackup


## Conditions tab

✅ Check “Wake the computer to run this task” (optional)

✅ Allow task to be run on demand

✅ Run task as soon as possible after a scheduled start is missed

If the task fails, restart every: 5 minutes, attempt: 2 (optional)
