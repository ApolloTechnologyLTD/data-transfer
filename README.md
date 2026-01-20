<div align="center">

# Apollo Technology Data Migration Utility

**Smart Backup & Restore tool with automated reporting and application data handling.**

</div>

---

## 📖 About

The **Apollo Technology Data Migration Utility (v2.7)** is a menu-driven PowerShell tool designed to streamline the transfer of user data between workstations. It simplifies the backup and restoration process for IT engineers by automating folder selection, preventing system sleep, and generating professional reports.

Key capabilities include **silent application data detection** (Chrome, Edge, Firefox, Outlook), **automated Google Chrome installation** during restoration, and **visual progress tracking**.

> ⚠️ **Disclaimer** – This tool operates in Administrative mode to copy system files and modify power settings during transfer.

---

## ✨ Features

### 🚀 Smart Data Transfer
- **Menu-Driven:** Simple selection for **Backup** (PC to External Drive) or **Restore** (External Drive to PC).
- **Robocopy Integration:** Uses robust file copying with retries and logging.
- **Visual Progress Bar:** Real-time feedback on current file and total transfer percentage.

### 🧠 Intelligent App Data Handling
- **Browser Data:** Automatically detects and migrates User Data for **Chrome, Edge, Firefox, Opera, and Brave**.
- **Outlook Data:** Captures PSTs, `Outlook Files` in Documents, and AppData configurations.
- **Silent Detection:** Only copies application folders if they exist on the source.

### 🛠️ System Management
- **Anti-Freeze:** Disables Console QuickEdit mode to prevent script freezing on clicks.
- **Anti-Sleep:** Prevents the computer from sleeping or turning off the screen during long transfers.
- **Chrome Installer:** Automatically downloads and installs Google Chrome Enterprise if missing during a restore.

### 📝 Reporting & Notifications
- **PDF Reporting:** Generates a branded HTML report and converts it to PDF using Microsoft Edge.
- **Email Integration:** Optional SMTP support to email the final report to a support mailbox.
- **Logs:** Creates detailed text logs of the Robocopy operation.

---

## 📋 Requirements

| Requirement | Details |
|------------|---------|
| **Operating System** | Windows 10 / Windows 11 |
| **PowerShell** | PowerShell 5.1 or later |
| **Permissions** | Administrator privileges required (Auto-elevates if not present) |
| **Dependencies** | Microsoft Edge (Required for PDF generation) |
| **Storage** | External Drive for Backup/Restore operations |

---

## 🚀 Quick Start

Run the following command to download and execute the utility:

```powershell
iwr https://short.apollotechnology.co.uk/backup_and_recovery -OutFile backupandrecovery.ps1; powershell -ExecutionPolicy Bypass .\backupandrecovery.ps1
```

> [!IMPORTANT]
> **ADMINISTRATOR PRIVILEGES REQUIRED**
>
> This script modifies system files/registries. You must launch your PowerShell with **"Run as Administrator"** rights.
> If you run this in a standard PowerShell, the script will fail or behave unexpectedly.


---
## ⚙️ Configuration

You can customize the script behavior by editing the variables at the top of the script file.

| Variable | Default | Description |
|--------|---------|-------------|
| `$DemoMode` | `false` | Simulates the transfer process without copying files. |
| `$EmailEnabled` | `false` | Set to `true` to enable email reporting. |
| `$SmtpServer` | `smtp.office365.com` | SMTP server used for sending reports. |
| `$ToAddress` | `support@...` | Email address that will receive the PDF report. |

---

## 📚 Usage Guide

### 1️⃣ Backup Mode

Run the script on the **Source Computer**.

1. Launch the script.
2. Select **Option 1 (BACKUP)**.
3. Enter the following details when prompted:
   - Engineer Name
   - Ticket Number
   - Customer Name
4. Select the external drive letter (for example: `D:` or `E:`).

📁 The script will create a folder on the external drive using the format:

```
[TicketNumber]-[CustomerName]
```

---

### 2️⃣ Restore Mode

Run the script on the **Destination Computer**.

1. Launch the script.
2. Select **Option 2 (RESTORE)**.
3. Enter the **Ticket Number**.

🔍 The script will automatically scan the external drive for a matching Ticket ID.

🌐 If **Google Chrome** is not installed, the script will automatically install it before starting the restore process.

---

## 📝 Report Output

After completion, the following files are generated inside the backup folder:

- `[BACKUP/RESTORE]_Log.txt`  
  Raw Robocopy transfer logs

- `[BACKUP/RESTORE]_Report.html`  
  Visual summary of the transfer

- `[BACKUP/RESTORE]_Report.pdf`  
  Professional PDF version of the summary

---

## 📊 Sample Report Data Includes

- Operation Mode (Backup or Restore)
- Engineer Name
- Customer Name
- Status of every folder transferred (Success / Failed)

---

<div align="center">
Created by: <strong>Lewis Wiltshire</strong> | Version: <strong>2.7</strong>
</div>
