<#
.SYNOPSIS
    Customer Data Backup Script for Field Engineers.
.DESCRIPTION
    Backs up user profiles, Chrome data, and Email data to an external drive using Robocopy.
    Generates a log and folder structure based on Ticket and Customer Name.
.NOTES
    Run as Administrator.
#>

# --- 1. SETUP & INPUTS ---
Clear-Host
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "   CUSTOMER DATA BACKUP UTILITY" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

# Check for Administrator privileges
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Please run this script as Administrator to ensure all files can be copied."
    Break
}

# Engineer and Ticket Details
$EngineerName = Read-Host "Enter Engineer Name"
$TicketNumber = Read-Host "Enter Ticket Number"
$CustomerName = Read-Host "Enter Customer Full Name"

# Drive Selection
Write-Host "`nAvailable Drives:" -ForegroundColor Yellow
Get-PSDrive -PSProvider FileSystem | Select-Object Name, Used, Free | Format-Table -AutoSize
$DriveLetterInput = Read-Host "Enter External Drive Letter (e.g. E or E:)"

# Sanitize Drive Letter (remove colon if typed)
$DriveLetter = $DriveLetterInput -replace ":", ""

# Validate Drive
if (!(Test-Path "$($DriveLetter):")) {
    Write-Error "Drive $($DriveLetter): not found. Please check connectivity and try again."
    Pause
    Exit
}

# Construct Destination Path
$DestRoot = "$($DriveLetter):\${TicketNumber}-$($CustomerName)"
$LogPath  = "$DestRoot\_Logs"

# Confirm with user
Write-Host "`n------------------------------------------"
Write-Host "Backing up current user: $env:USERNAME"
Write-Host "Destination: $DestRoot"
Write-Host "------------------------------------------"
$Confirm = Read-Host "Type 'Y' to proceed, or anything else to cancel"
if ($Confirm -ne 'Y') { Exit }

# --- 2. PREPARATION ---

# Create Destination Directory
if (!(Test-Path $DestRoot)) {
    New-Item -ItemType Directory -Path $DestRoot -Force | Out-Null
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}

# Stop Processes to release file locks (Chrome, Outlook)
Write-Host "`nStopping Chrome and Outlook to prevent file lock errors..." -ForegroundColor Yellow
Stop-Process -Name "chrome" -Force -ErrorAction SilentlyContinue
Stop-Process -Name "outlook" -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

# --- 3. THE BACKUP FUNCTION (ROBOCOPY) ---

function Run-Backup {
    param (
        [string]$Source,
        [string]$FolderName,
        [string]$LogFile
    )

    $FullDest = "$DestRoot\$FolderName"
    
    if (Test-Path $Source) {
        Write-Host "Backing up: $FolderName..." -ForegroundColor Green
        
        # Robocopy Switches:
        # /E   :: copy subdirectories, including Empty ones.
        # /XO  :: eXclude Older files.
        # /R:1 :: Retry 1 time on failed copies.
        # /W:1 :: Wait 1 second between retries.
        # /NP  :: No Progress - prevents log clutter.
        # /LOG+ :: Append output to log file.
        
        robocopy $Source $FullDest /E /XO /R:1 /W:1 /NP /LOG+:"$LogFile" /TEE
    } else {
        Write-Warning "Source not found: $Source"
        Add-Content -Path $LogFile -Value "SKIPPED: Source not found - $Source"
    }
}

# --- 4. EXECUTION ---

$MainLog = "$LogPath\TransferLog.txt"
Add-Content -Path $MainLog -Value "Backup Started: $(Get-Date)"
Add-Content -Path $MainLog -Value "Engineer: $EngineerName"
Add-Content -Path $MainLog -Value "Ticket: $TicketNumber"
Add-Content -Path $MainLog -Value "--------------------------------"

# Define User Profile Path
$UserProfile = $env:USERPROFILE

# A. Standard User Data
Run-Backup -Source "$UserProfile\Desktop"   -FolderName "Desktop"   -LogFile $MainLog
Run-Backup -Source "$UserProfile\Documents" -FolderName "Documents" -LogFile $MainLog
Run-Backup -Source "$UserProfile\Downloads" -FolderName "Downloads" -LogFile $MainLog
Run-Backup -Source "$UserProfile\Pictures"  -FolderName "Pictures"  -LogFile $MainLog
Run-Backup -Source "$UserProfile\Music"     -FolderName "Music"     -LogFile $MainLog
Run-Backup -Source "$UserProfile\Videos"    -FolderName "Videos"    -LogFile $MainLog

# B. Browser Data (Chrome)
# Note: Backs up "User Data". Passwords are encrypted by Windows DPAPI.
Run-Backup -Source "$UserProfile\AppData\Local\Google\Chrome\User Data" -FolderName "ChromeData" -LogFile $MainLog

# C. Email Data (Outlook)
# Outlook usually stores PST/OST in two possible locations
Run-Backup -Source "$UserProfile\AppData\Local\Microsoft\Outlook" -FolderName "Outlook_AppData" -LogFile $MainLog
Run-Backup -Source "$UserProfile\Documents\Outlook Files"         -FolderName "Outlook_Documents" -LogFile $MainLog

# --- 5. REPORT GENERATION ---

$ReportFile = "$DestRoot\Backup_Report.txt"

$ReportContent = @"
=================================================
          DATA MIGRATION REPORT
=================================================
Date:           $(Get-Date)
Engineer:       $EngineerName
Ticket Number:  $TicketNumber
Customer:       $CustomerName
-------------------------------------------------
Status:         COMPLETED
Destination:    $DestRoot

Items Attempted:
- Desktop, Documents, Downloads
- Pictures, Music, Video
- Google Chrome User Data (Bookmarks/History)
- Outlook Data Files (AppData & Documents)

Detailed logs can be found in the '_Logs' folder.
=================================================
"@

Set-Content -Path $ReportFile -Value $ReportContent

Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "   BACKUP COMPLETE" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "Report saved to: $ReportFile"
Write-Host "Please verify the folder size before wiping the old device."
Pause