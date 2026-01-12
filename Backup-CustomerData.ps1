<#
.SYNOPSIS
    Apollo Technology Data Migration Utility (Backup & Restore)
.DESCRIPTION
    Menu-driven utility to Backup data to external drive OR Restore data from external drive.
    Includes auto-elevation and Demo Mode.
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

# --- 2. HELPER FUNCTIONS ---
function Show-Header {
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
}

function Run-Robocopy {
    param (
        [string]$Source,
        [string]$Destination,
        [string]$LogFile
    )

    if ($DemoMode) {
        Write-Host "   [DEMO] Processing: $Source -> $Destination" -ForegroundColor Magenta
        # Create Dummy Dest
        if (!(Test-Path $Destination)) { New-Item -ItemType Directory -Path $Destination -Force | Out-Null }
        # Create Dummy File
        New-Item -ItemType File -Path "$Destination\DEMO_TRANSFER.txt" -Value "Simulated copy from $Source" -Force | Out-Null
        Add-Content -Path $LogFile -Value "DEMO COPY: $Source -> $Destination"
        Start-Sleep -Milliseconds 200
    }
    else {
        if (Test-Path $Source) {
            Write-Host "   Copying: $Source" -ForegroundColor Green
            # Robocopy Args: /E (Recursive) /XO (Exclude Older) /R:1 (Retry once) /W:1 (Wait 1 sec) /NP (No progress bar)
            robocopy $Source $Destination /E /XO /R:1 /W:1 /NP /LOG+:"$LogFile" /TEE | Out-Null
        } else {
            Write-Host "   Skipping: Source not found ($Source)" -ForegroundColor DarkGray
            Add-Content -Path $LogFile -Value "SKIPPED: Source Missing - $Source"
        }
    }
}

# --- 3. MAIN MENU ---
Show-Header
Write-Host "`n[ SELECT OPERATION MODE ]" -ForegroundColor Yellow
Write-Host "   1. BACKUP  (Computer -> External Drive)"
Write-Host "   2. RESTORE (External Drive -> New Computer)"
Write-Host "   3. EXIT"
Write-Host "---------------------------------------------------------------------------------"

$MenuSelection = Read-Host "   > Enter Option (1, 2, or 3)"

switch ($MenuSelection) {
    '1' { $Mode = "BACKUP" }
    '2' { $Mode = "RESTORE" }
    '3' { Exit }
    Default { Write-Host "Invalid selection."; Pause; Exit }
}

# --- 4. INPUT COLLECTION ---
Show-Header
Write-Host "`n[ $Mode CONFIGURATION ]" -ForegroundColor Yellow
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

# Define Paths
$ExternalStorePath = "$($DriveLetter):\${TicketNumber}-$($CustomerName)"
$UserProfile = $env:USERPROFILE
$LogPath = "$ExternalStorePath\_Logs"

# --- 5. PREPARATION & CONFIRMATION ---
Show-Header
Write-Host "`n[ CONFIRMATION ]" -ForegroundColor Yellow
Write-Host "   Operation:   $Mode" -ForegroundColor $(If ($Mode -eq 'BACKUP') {'Green'} Else {'Cyan'})
Write-Host "   Customer:    $CustomerName (Ticket: $TicketNumber)"
if ($Mode -eq "BACKUP") {
    Write-Host "   Source:      THIS COMPUTER ($UserProfile)"
    Write-Host "   Destination: EXTERNAL DRIVE ($ExternalStorePath)"
} else {
    Write-Host "   Source:      EXTERNAL DRIVE ($ExternalStorePath)"
    Write-Host "   Destination: THIS COMPUTER ($UserProfile)"
}

Write-Host "---------------------------------------------------------------------------------"
$Confirm = Read-Host "Type 'Y' to proceed"
if ($Confirm -ne 'Y') { Exit }

# Setup Logs
if (!(Test-Path $ExternalStorePath)) { New-Item -ItemType Directory -Path $ExternalStorePath -Force | Out-Null }
if (!(Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath -Force | Out-Null }
$MainLog = "$LogPath\${Mode}_Log.txt"
Add-Content -Path $MainLog -Value "Operation: $Mode | Date: $(Get-Date)"

# Stop Processes (Required for both backup and restore to avoid file locks)
Write-Host "`n[ SYSTEM PREP ]" -ForegroundColor Yellow
if (-not $DemoMode) {
    Write-Host "Stopping Chrome and Outlook..." -ForegroundColor Gray
    Stop-Process -Name "chrome" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "outlook" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
} else {
    Write-Host "Simulating Process Stop..." -ForegroundColor DarkGray
}

# --- 6. EXECUTION ENGINE ---
Write-Host "`n[ STARTING TRANSFER ]" -ForegroundColor Yellow

# Define the folders to process using a HashTable for clean mapping
# Key = Local Folder Name, Value = Folder Name on External Drive
$FoldersMap = @{
    "Desktop"   = "Desktop"
    "Documents" = "Documents"
    "Downloads" = "Downloads"
    "Pictures"  = "Pictures"
    "Music"     = "Music"
    "Videos"    = "Videos"
    "AppData\Local\Google\Chrome\User Data" = "ChromeData"
    "AppData\Local\Microsoft\Outlook"       = "Outlook_AppData"
    "Documents\Outlook Files"               = "Outlook_Documents"
}

foreach ($LocalSubPath in $FoldersMap.Keys) {
    $ExternalSubName = $FoldersMap[$LocalSubPath]
    
    # Calculate Full Paths based on Mode
    $LocalFull  = "$UserProfile\$LocalSubPath"
    $ExternalFull = "$ExternalStorePath\$ExternalSubName"

    if ($Mode -eq "BACKUP") {
        # Source = Local, Dest = External
        Run-Robocopy -Source $LocalFull -Destination $ExternalFull -LogFile $MainLog
    }
    elseif ($Mode -eq "RESTORE") {
        # Source = External, Dest = Local
        Run-Robocopy -Source $ExternalFull -Destination $LocalFull -LogFile $MainLog
    }
}

# --- 7. REPORT & FINISH ---
$ReportFile = "$ExternalStorePath\${Mode}_Report.txt"
$ReportContent = @"
APOLLO TECHNOLOGY - $Mode REPORT
=========================================
Date:           $(Get-Date)
Engineer:       $EngineerName
Ticket:         $TicketNumber
Operation:      $Mode
Status:         COMPLETED
=========================================
"@
Set-Content -Path $ReportFile -Value $ReportContent

Write-Host "`n[ COMPLETE ]" -ForegroundColor Green
Write-Host "Operation finished."
Write-Host "Log saved to: $MainLog"
Pause