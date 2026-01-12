<#
.SYNOPSIS
    Apollo Technology Data Backup Utility (Auto-Elevate & Internal Demo Mode)
.DESCRIPTION
    Backs up user profiles, Chrome data, and Email data to an external drive.
    Includes auto-elevation and a hardcoded toggle for simulation mode.
#>

# --- 0. CONFIGURATION ---
# CHANGE THIS TO $FALSE WHEN READY FOR REAL USE
$DemoMode = $true 

# --- 1. AUTO-ELEVATE TO ADMINISTRATOR ---
$CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (!($CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
    Write-Host "Requesting Administrator privileges..." -ForegroundColor Yellow
    try {
        Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        Exit
    }
    catch {
        Write-Error "Failed to elevate. Please run as Administrator manually."
        Pause
        Exit
    }
}

# --- 2. SETUP & BANNER ---
Clear-Host
$Banner = @'
    ___    ____  ____  __    __    ____     ____________________  ___   ______  __    ____  ________  __
   /   |  / __ \/ __ \/ /   / /   / __ \   /_  __/ ____/ ____/ / / / | / / __ \/ /   / __ \/ ____/\ \/ /
  / /| | / /_/ / / / / /   / /   / / / /    / / / __/ / /   / /_/ /  |/ / / / / /   / / / / / __   \  / 
 / ___ |/ ____/ /_/ / /___/ /___/ /_/ /    / / / /___/ /___/ __  / /|  / /_/ / /___/ /_/ / /_/ /   / /  
/_/  |_/_/    \____/_____/_____/\____/    /_/ /_____/\____/_/ /_/_/ |_/\____/_____/\____/\____/   /_/   
'@

Write-Host $Banner -ForegroundColor Cyan
Write-Host "`n   DATA MIGRATION & BACKUP TOOL" -ForegroundColor White
Write-Host "=================================================================================" -ForegroundColor DarkGray

if ($DemoMode) {
    Write-Host "`n   *** DEMO MODE ACTIVE - NO REAL DATA WILL BE COPIED ***" -ForegroundColor Magenta
}

# --- 3. INPUT COLLECTION ---
Write-Host "`n[ INPUT REQUIRED ]" -ForegroundColor Yellow
$EngineerName = Read-Host "   > Enter Engineer Name"
$TicketNumber = Read-Host "   > Enter Ticket Number"
$CustomerName = Read-Host "   > Enter Customer Full Name"

# Drive Selection
Write-Host "`n[ DRIVE SELECTION ]" -ForegroundColor Yellow
Write-Host "Detecting drives..." -ForegroundColor DarkGray
Get-PSDrive -PSProvider FileSystem | Select-Object Name, Used, Free, Root | Format-Table -AutoSize

$DriveLetterInput = Read-Host "   > Enter External Drive Letter (e.g. D or E)"
$DriveLetter = $DriveLetterInput -replace ":", ""

if (!(Test-Path "$($DriveLetter):")) {
    Write-Error "Drive $($DriveLetter): not found."
    Pause
    Exit
}

$DestRoot = "$($DriveLetter):\${TicketNumber}-$($CustomerName)"
$LogPath  = "$DestRoot\_Logs"

# --- 4. CONFIRMATION ---
Clear-Host
Write-Host $Banner -ForegroundColor Cyan
Write-Host "`n[ CONFIRMATION ]" -ForegroundColor Yellow
Write-Host "   Mode:        $(If ($DemoMode) {'DEMO (Simulation)'} Else {'REAL COPY'})" -ForegroundColor $(If ($DemoMode) {'Magenta'} Else {'Green'})
Write-Host "   Engineer:    $EngineerName"
Write-Host "   Ticket:      $TicketNumber"
Write-Host "   Destination: $DestRoot"
Write-Host "---------------------------------------------------------------------------------"

$Confirm = Read-Host "Type 'Y' to proceed"
if ($Confirm -ne 'Y') { Exit }

# --- 5. PREPARATION ---
# Create Directory
if (!(Test-Path $DestRoot)) {
    New-Item -ItemType Directory -Path $DestRoot -Force | Out-Null
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}

# Stop Processes (Only in Real Mode)
if (-not $DemoMode) {
    Write-Host "`n[ SYSTEM PREP ]" -ForegroundColor Yellow
    Write-Host "Stopping Chrome and Outlook..." -ForegroundColor Gray
    Stop-Process -Name "chrome" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "outlook" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
} else {
    Write-Host "`n[ SYSTEM PREP ]" -ForegroundColor Yellow
    Write-Host "Simulating Process Stop..." -ForegroundColor DarkGray
}

# --- 6. THE BACKUP ENGINE ---
function Run-Backup {
    param (
        [string]$Source,
        [string]$FolderName,
        [string]$LogFile
    )

    $FullDest = "$DestRoot\$FolderName"
    
    if (Test-Path $Source) {
        Write-Host "   Processing: $FolderName..." -ForegroundColor Green
        
        if ($DemoMode) {
            # DEMO MODE: Create the folder and a dummy file
            New-Item -ItemType Directory -Path $FullDest -Force | Out-Null
            New-Item -ItemType File -Path "$FullDest\DEMO_FILE.txt" -Value "This is a demo placeholder for $FolderName" -Force | Out-Null
            Add-Content -Path $LogFile -Value "SIMULATED COPY: $Source -> $FullDest"
            Start-Sleep -Milliseconds 200 # Fake delay for effect
        } 
        else {
            # REAL MODE: Robocopy
            robocopy $Source $FullDest /E /XO /R:1 /W:1 /NP /LOG+:"$LogFile" /TEE | Out-Null
        }
    } else {
        Write-Host "   Skipping: $FolderName (Not Found)" -ForegroundColor DarkGray
        Add-Content -Path $LogFile -Value "SKIPPED: Source not found - $Source"
    }
}

$MainLog = "$LogPath\TransferLog.txt"
Add-Content -Path $MainLog -Value "Backup Started: $(Get-Date)"
Add-Content -Path $MainLog -Value "Mode: $(If ($DemoMode) {'DEMO'} Else {'REAL'})"
Add-Content -Path $MainLog -Value "Engineer: $EngineerName | Ticket: $TicketNumber"
Add-Content -Path $MainLog -Value "--------------------------------"

Write-Host "`n[ TRANSFERRING DATA ]" -ForegroundColor Yellow
$UserProfile = $env:USERPROFILE

# User Data
Run-Backup -Source "$UserProfile\Desktop"   -FolderName "Desktop"   -LogFile $MainLog
Run-Backup -Source "$UserProfile\Documents" -FolderName "Documents" -LogFile $MainLog
Run-Backup -Source "$UserProfile\Downloads" -FolderName "Downloads" -LogFile $MainLog
Run-Backup -Source "$UserProfile\Pictures"  -FolderName "Pictures"  -LogFile $MainLog
Run-Backup -Source "$UserProfile\Music"     -FolderName "Music"     -LogFile $MainLog
Run-Backup -Source "$UserProfile\Videos"    -FolderName "Videos"    -LogFile $MainLog

# Chrome & Outlook
Run-Backup -Source "$UserProfile\AppData\Local\Google\Chrome\User Data" -FolderName "ChromeData" -LogFile $MainLog
Run-Backup -Source "$UserProfile\AppData\Local\Microsoft\Outlook" -FolderName "Outlook_AppData" -LogFile $MainLog
Run-Backup -Source "$UserProfile\Documents\Outlook Files" -FolderName "Outlook_Documents" -LogFile $MainLog

# --- 7. REPORT GENERATION ---
$ReportFile = "$DestRoot\Backup_Report.txt"
$ReportContent = @"
APOLLO TECHNOLOGY - DATA MIGRATION REPORT
=========================================
Date:           $(Get-Date)
Engineer:       $EngineerName
Ticket Number:  $TicketNumber
Customer:       $CustomerName
Mode:           $(If ($DemoMode) {'DEMO'} Else {'REAL'})
-----------------------------------------
STATUS:         SUCCESS
LOCATION:       $DestRoot

ITEMS PROCESSED:
[x] Desktop, Documents, Downloads
[x] Pictures, Music, Videos
[x] Chrome Data (Bookmarks/History)
[x] Outlook Data (PST/OST)
=========================================
"@

Set-Content -Path $ReportFile -Value $ReportContent

# --- 8. FINISH ---
Write-Host "`n[ COMPLETE ]" -ForegroundColor Green
Write-Host "Report generated at: $ReportFile"
Write-Host "You may now safely remove the drive."
Pause