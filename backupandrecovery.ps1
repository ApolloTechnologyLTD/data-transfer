<#
.SYNOPSIS
    Apollo Technology Data Migration Utility (Smart Restore & Backup) v4.8 Beta
.DESCRIPTION
    Menu-driven utility to Backup data or Restore data.
    - UPDATED v4.8: Added a mandatory user warning prompt when Verbose Logging is active.
    - NEW v4.7: Added Verbose Logging mode (captures all console/error output to C:\temp\backup).
    - UPDATED v4.6: Rebuilt Drive Scanner using .NET DriveInfo. Instant loading, cleanly formatted.
    - UPDATED v4.5: Added Out-Host to prevent prompt skipping.
    - UPDATED v4.3: Removed Disclaimer from Final Report.
    - NEW v4.2: Automated Permission Fixer for Slave Drives (Takeown/Icacls).
    - INCLUDES: Anti-Sleep, Anti-Freeze, Email Reports, Visual Progress Bar.
#>

# --- 0. CONFIGURATION ---
$DemoMode = $false
$VerboseMode = $true         # Set to $true to log all script output to C:\temp\backup\backuplogs.txt
$LogoUrl = "https://raw.githubusercontent.com/ApolloTechnologyLTD/computer-health-check/main/Apollo%20Cropped.png"
$Version = "4.8 Beta"

# --- EMAIL SETTINGS ---
$EmailEnabled = $false       # Set to $true to enable email
$SmtpServer   = "smtp.office365.com"
$SmtpPort     = 587
$FromAddress  = "reports@yourdomain.com"
$ToAddress    = "support@yourdomain.com"
$UseSSL       = $true

# --- 1. AUTO-ELEVATE TO ADMINISTRATOR ---
$CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (!($isAdmin)) {
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

# --- 1.5 VERBOSE LOGGING SETUP & WARNING ---
if ($VerboseMode) {
    $VerboseDir = "C:\temp\backup"
    if (!(Test-Path $VerboseDir)) {
        New-Item -ItemType Directory -Path $VerboseDir -Force | Out-Null
    }
    
    # Start capturing all console output, errors, and warnings
    Start-Transcript -Path "$VerboseDir\backuplogs.txt" -Append -Force | Out-Null
    
    Clear-Host
    Write-Host "`n=================================================================================" -ForegroundColor Magenta
    Write-Host " [ WARNING: VERBOSE LOGGING IS ENABLED ]" -ForegroundColor Red
    Write-Host "=================================================================================" -ForegroundColor Magenta
    Write-Host " All console output, background processes, and errors are currently being recorded."
    Write-Host " Log File Location: " -NoNewline; Write-Host "$VerboseDir\backuplogs.txt" -ForegroundColor Cyan
    Write-Host "`n Use this mode for debugging purposes only." -ForegroundColor Yellow
    Write-Host "---------------------------------------------------------------------------------" -ForegroundColor DarkGray
    $null = Read-Host " Press [ENTER] to acknowledge and continue"
}

# --- 2. PREVENT FREEZING & SLEEPING ---
# Disable Quick-Edit (Prevents freezing on click)
$consoleFuncs = @"
using System;
using System.Runtime.InteropServices;
public class ConsoleUtils {
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr GetStdHandle(int nStdHandle);
    [DllImport("kernel32.dll")]
    public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);
    [DllImport("kernel32.dll")]
    public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);
    public static void DisableQuickEdit() {
        IntPtr hConsole = GetStdHandle(-10); // STD_INPUT_HANDLE
        uint mode;
        GetConsoleMode(hConsole, out mode);
        mode &= ~0x0040u; // ENABLE_QUICK_EDIT_MODE = 0x0040
        SetConsoleMode(hConsole, mode);
    }
}
"@
try {
    Add-Type -TypeDefinition $consoleFuncs -Language CSharp
    [ConsoleUtils]::DisableQuickEdit()
} catch { }

# Prevent Sleep
$sleepBlocker = @"
using System;
using System.Runtime.InteropServices;
public class SleepUtils {
    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern uint SetThreadExecutionState(uint esFlags);
}
"@
try {
    Add-Type -TypeDefinition $sleepBlocker -Language CSharp
    $null = [SleepUtils]::SetThreadExecutionState(0x80000003) # ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
} catch { }

# --- 3. HELPER FUNCTIONS ---
function Show-Header {
    Clear-Host
    $Banner = @'
    ___    ____  ____  __    __    ____     ____________________  ___   ______  __    ____  ________  __
   /   |  / __ \/ __ \/ /   / /   / __ \   /_  __/ ____/ ____/ / / / | / / __ \/ /   / __ \/ ____/\ \/ /
  / /| | / /_/ / / / / /   / /   / / / /    / / / __/ / /   / /_/ /  |/ / / / / /   / / / / / __   \  / 
 / ___ |/ ____/ /_/ / /___/ /___/ /_/ /    / / / /___/ /___/ __  / /|  / /_/ / /___/ /_/ / /___/ /_/ /   / /  
/_/  |_/_/    \____/_____/_____/\____/    /_/ /_____/\____/_/ /_/_/ |_/\____/_____/\____/\____/   /_/   
'@
    Write-Host $Banner -ForegroundColor Cyan
    Write-Host "`n   DATA MIGRATION & BACKUP TOOL v$Version" -ForegroundColor White
    Write-Host "=================================================================================" -ForegroundColor DarkGray

    # --- NOTICE & DETAILS ---
    $Current = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $IsAdminHeader = $Current.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($IsAdminHeader) { 
        Write-Host "        [NOTICE] Running in Elevated Permissions" -ForegroundColor Red 
    } elseif ($DemoMode) {
        Write-Host "      [NOTICE] Running as Standard User" -ForegroundColor Yellow 
    }

    Write-Host "        Created by Lewis Wiltshire, Apollo Technology" -ForegroundColor Yellow
    Write-Host "      [POWER] Sleep Mode & Screen Timeout Blocked." -ForegroundColor DarkGray
    # -------------------------------

    if ($DemoMode) {
        Write-Host "`n   *** DEMO MODE ACTIVE - NO REAL DATA WILL BE COPIED ***" -ForegroundColor Magenta
    }
}

function Show-Disclaimer {
    Show-Header
    Write-Host "`n[ IMPORTANT DISCLAIMER & LIABILITY WAIVER ]" -ForegroundColor Red
    
    $DisclaimerText = @"
WARNING: You are about to use software currently in BETA (Version $Version).

1. NO WARRANTY: This script is provided "as-is" without warranty of any kind, express or implied.
2. DATA INTEGRITY: Do NOT rely on this tool completely. It is an automation utility designed to assist, 
   not replace, professional verification.
3. LIMITATION OF LIABILITY: Lewis Wiltshire accepts NO LIABILITY and NO BLAME 
   for any data loss, corruption, missing files, failed transfers, or damages resulting from the use 
   of this script.
4. RESPONSIBILITY: It is the sole responsibility of the Engineer using this tool to manually verify 
   that all critical data has been successfully transferred before formatting or disposing of the 
   source device.

By proceeding, you acknowledge these risks and agree to hold the creator harmless.
"@
    Write-Host $DisclaimerText -ForegroundColor Yellow
    Write-Host "`n---------------------------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "   PRESS [Y] or [ENTER] to Accept and Continue." -ForegroundColor Green
    Write-Host "   PRESS [N] or [ESC]   to Decline and Exit." -ForegroundColor Red
    
    while ($true) {
        $Key = [System.Console]::ReadKey($true)
        if ($Key.Key -eq 'Y' -or $Key.Key -eq 'Enter') {
            return # Continue script
        }
        elseif ($Key.Key -eq 'N' -or $Key.Key -eq 'Escape') {
            Write-Host "`n   User declined disclaimer. Exiting..." -ForegroundColor Red
            if ($VerboseMode) { Stop-Transcript | Out-Null }
            Start-Sleep -Seconds 1
            Exit
        }
    }
}

function Show-DriveList {
    # Using .NET DriveInfo. Instant loading, zero hangs, accurate names, and clean GB numbers.
    Write-Host "   Detecting drives..." -ForegroundColor DarkGray
    
    $drives = [System.IO.DriveInfo]::GetDrives() | Where-Object { $_.IsReady }
    $displayList = @()
    
    foreach ($d in $drives) {
        $Label = if ([string]::IsNullOrWhiteSpace($d.VolumeLabel)) { "[No Label]" } else { $d.VolumeLabel }
        $displayList += [PSCustomObject]@{
            Letter     = $d.Name.Substring(0,2) # Gets format like "C:"
            Name       = $Label
            "Size(GB)" = [math]::Round($d.TotalSize / 1GB, 2)
            "Free(GB)" = [math]::Round($d.TotalFreeSpace / 1GB, 2)
        }
    }
    
    # Out-Host ensures it renders synchronously and doesn't get skipped by the next prompt
    $displayList | Format-Table -AutoSize | Out-Host
}

function Install-GoogleChrome {
    Write-Host "`n[ CHROME INSTALLER ]" -ForegroundColor Yellow
    Write-Host "   Chrome not found. Downloading Enterprise Installer..." -ForegroundColor Cyan
    
    $InstallerUrl = "https://dl.google.com/chrome/install/googlechromestandaloneenterprise64.msi"
    $InstallerPath = "$env:TEMP\ChromeEnterprise.msi"

    try {
        Invoke-WebRequest -Uri $InstallerUrl -OutFile $InstallerPath -UseBasicParsing
        Write-Host "   Installing Chrome (Silent)..." -ForegroundColor Cyan
        
        $Process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$InstallerPath`" /quiet /norestart" -Wait -PassThru
        
        if ($Process.ExitCode -eq 0) {
            Write-Host "   SUCCESS: Google Chrome installed." -ForegroundColor Green
        } else {
            Write-Host "   WARNING: Install exited with code $($Process.ExitCode)." -ForegroundColor Red
        }
    }
    catch {
        Write-Error "   Failed to install Chrome. Please install manually."
    }
}

# --- NEW v4.2 PERMISSION FIXER ---
function Fix-SlaveDrivePermissions {
    param([string]$Path)
    
    Write-Host "`n[ PERMISSION CHECK ]" -ForegroundColor Yellow
    Write-Host "   You have selected a Slave Drive source."
    Write-Host "   Windows often denies access to User folders from other PCs."
    Write-Host "   We can attempt to Take Ownership and Grant Admin Access to:" -ForegroundColor Cyan
    Write-Host "   $Path" -ForegroundColor White
    
    $Choice = Read-Host "   > Attempt to fix permissions? (Recommended if Access Denied errors occur) [Y/N]"
    
    if ($Choice -match "Y") {
        Write-Host "   Applying permissions (This may take a moment)..." -ForegroundColor Yellow
        
        # 1. Take Ownership
        Write-Host "   > Step 1: Taking Ownership..." -ForegroundColor DarkGray
        try {
            # /F = File/Folder, /R = Recursive, /D Y = Answer Yes to confirmation
            cmd.exe /c "takeown /F `"$Path`" /R /D Y" | Out-Null
        } catch { Write-Warning "TakeOwn failed or partial success." }

        # 2. Grant Administrators Full Control
        Write-Host "   > Step 2: Granting Admin Access..." -ForegroundColor DarkGray
        try {
            # /grant Administrators:F = Full Control, /T = Recursive, /C = Continue on error, /Q = Quiet
            cmd.exe /c "icacls `"$Path`" /grant Administrators:F /T /C /Q" | Out-Null
            Write-Host "   > Permissions update complete." -ForegroundColor Green
        } catch { Write-Warning "Icacls failed or partial success." }
    } else {
        Write-Host "   Skipping permission fix." -ForegroundColor DarkGray
    }
}

function Get-SourceUserFromDrive {
    # Helper to select a user profile from an external drive
    Write-Host "`n[ SELECT SOURCE DRIVE (OFFLINE MODE) ]" -ForegroundColor Yellow
    
    # NEW READOUT
    Show-DriveList
    
    $SourceDriveLetter = Read-Host "   > Enter Drive Letter of the OLD computer (e.g. E)"
    $SourceDriveLetter = $SourceDriveLetter -replace ":", ""
    $UsersRoot = "$($SourceDriveLetter):\Users"

    if (!(Test-Path $UsersRoot)) {
        Write-Error "   Users folder not found at $UsersRoot. Please check the drive letter."
        Start-Sleep -Seconds 2
        return $null
    }

    Write-Host "`n[ DETECTED USER PROFILES ]" -ForegroundColor Yellow
    $UserFolders = Get-ChildItem -Path $UsersRoot -Directory | Where-Object { $_.Name -notin "Public", "Default", "All Users", "Default User" }
    
    $i = 1
    foreach ($u in $UserFolders) {
        Write-Host "   $i. $($u.Name)"
        $i++
    }

    $Selection = Read-Host "`n   > Select User Number to Backup"
    if ($Selection -match "^\d+$" -and $Selection -le $UserFolders.Count) {
        $SelectedUser = $UserFolders[$Selection - 1]
        Write-Host "   Selected: $($SelectedUser.FullName)" -ForegroundColor Green
        
        # CALL PERMISSION FIX HERE [v4.2]
        Fix-SlaveDrivePermissions -Path $SelectedUser.FullName
        
        return $SelectedUser.FullName
    } else {
        Write-Host "   Invalid selection." -ForegroundColor Red
        return $null
    }
}

function Run-Robocopy {
    param (
        [string]$Source,
        [string]$Destination,
        [string]$LogFile
    )

    # Exclude junction points that cause loops or access denied
    $Excludes = @("Temp", "Temporary Internet Files", "Application Data", "History", "Cookies")

    if ($DemoMode) {
        if (!(Test-Path $Destination)) { New-Item -ItemType Directory -Path $Destination -Force | Out-Null }
        Write-Host "   [DEMO] Processing: $Source" -ForegroundColor Magenta
        Add-Content -Path $LogFile -Value "DEMO COPY: $Source -> $Destination"
        Start-Sleep -Milliseconds 100
        return "Success (Demo)"
    }
    else {
        if (Test-Path $Source) {
            # /E = recursive, /XO = exclude older, /R:1 /W:1 = 1 retry, 1 sec wait
            # /COPY:DAT = Copy Data, Attributes, Time stamps
            # /ZB = Restartable Mode + Backup Mode (CRITICAL FOR SLAVE DRIVES TO BYPASS PERMISSIONS)
            robocopy $Source $Destination /E /XO /COPY:DAT /ZB /R:1 /W:1 /NP /LOG+:"$LogFile" /TEE /XD $Excludes | Out-Null
            return "Completed"
        } else {
            # Log skipped items
            Write-Host "   Skipping: Source not found ($Source)" -ForegroundColor DarkGray
            Add-Content -Path $LogFile -Value "SKIPPED: Source Missing - $Source"
            return "Skipped (Not Found)"
        }
    }
}

# --- 4. DISCLAIMER CHECK ---
Show-Disclaimer

# --- 5. MAIN MENU ---
Show-Header
Write-Host "`n[ SELECT OPERATION MODE ]" -ForegroundColor Yellow
Write-Host "   1. BACKUP  (Copy data FROM a source TO an external drive)"
Write-Host "   2. RESTORE (Copy data FROM an external drive TO this computer)"
Write-Host "   3. EXIT"
Write-Host "---------------------------------------------------------------------------------"

$MenuSelection = Read-Host "   > Enter Option (1, 2, or 3)"

switch ($MenuSelection) {
    '1' { $Mode = "BACKUP" }
    '2' { $Mode = "RESTORE" }
    '3' { 
        if ($VerboseMode) { Stop-Transcript | Out-Null }
        Exit 
    }
    Default { 
        Write-Host "Invalid selection."; Pause; 
        if ($VerboseMode) { Stop-Transcript | Out-Null }
        Exit 
    }
}

# --- 6. INPUT COLLECTION LOOP ---
do {
    Show-Header
    Write-Host "`n[ $Mode CONFIGURATION ]" -ForegroundColor Yellow

    $EngineerName = Read-Host "   > Enter Engineer Name"
    $TicketNumber = Read-Host "   > Enter Ticket Number"

    # --- NEW: SOURCE SELECTION FOR BACKUP ---
    $SourceProfilePath = $env:USERPROFILE # Default to current user
    
    if ($Mode -eq "BACKUP") {
        Write-Host "`n[ SOURCE LOCATION ]" -ForegroundColor Cyan
        Write-Host "   1. This Computer (Current User: $env:USERNAME)"
        Write-Host "   2. External/Slave Drive (Offline Windows)"
        $SourceType = Read-Host "   > Select Source (1 or 2)"
        
        if ($SourceType -eq "2") {
            $PickedPath = Get-SourceUserFromDrive
            if ($PickedPath) {
                $SourceProfilePath = $PickedPath
            } else {
                Write-Host "Selection failed. Restarting..."
                Start-Sleep -Seconds 2
                Continue
            }
        }
    }
    # ----------------------------------------

    # Target Drive Selection
    Write-Host "`n[ STORAGE DRIVE SELECTION ]" -ForegroundColor Yellow
    
    # NEW READOUT FOR DESTINATION TOO
    Show-DriveList

    $DriveLetterInput = Read-Host "   > Enter Storage/Backup Drive Letter (e.g. D or E)"
    $DriveLetter = $DriveLetterInput -replace ":", ""

    if (!(Test-Path "$($DriveLetter):")) {
        Write-Error "Drive $($DriveLetter): not found."
        Write-Host "Restarting input..." -ForegroundColor Red
        Start-Sleep -Seconds 2
        Continue 
    }

    # PATH LOGIC & SEARCH
    $ExternalStorePath = $null
    $CustomerName = $null

    if ($Mode -eq "BACKUP") {
        $CustomerName = Read-Host "   > Enter Customer Full Name"
        $ExternalStorePath = "$($DriveLetter):\${TicketNumber}-$($CustomerName)"
        
    } elseif ($Mode -eq "RESTORE") {
        Write-Host "`n[ SEARCHING FOR BACKUP ]" -ForegroundColor Yellow
        Write-Host "   Searching drive $DriveLetter for Ticket $TicketNumber..." -ForegroundColor DarkGray
        
        $SearchPath = "$($DriveLetter):\${TicketNumber}-*"
        $FoundFolders = Get-ChildItem -Path $SearchPath -Directory -ErrorAction SilentlyContinue
        
        if ($null -eq $FoundFolders -or $FoundFolders.Count -eq 0) {
            Write-Host "`n[ ERROR ] NO BACKUP FOUND" -ForegroundColor Red
            Write-Host "Press any key to restart inputs..."
            Pause
            Continue
        } elseif ($FoundFolders.Count -gt 1) {
            Write-Host "`n[ ERROR ] MULTIPLE BACKUPS FOUND" -ForegroundColor Red
            $FoundFolders | ForEach-Object { Write-Host "   - $($_.Name)" }
            Write-Host "Press any key to restart inputs..."
            Pause
            Continue
        }

        $ExternalStorePath = $FoundFolders.FullName
        $CustomerName = $FoundFolders.Name 
        Write-Host "   FOUND: $CustomerName" -ForegroundColor Green
    }

    # Get Email Creds Early if Enabled
    $EmailCreds = $null
    if ($EmailEnabled) {
        Write-Host "`n[ EMAIL CONFIGURATION ]" -ForegroundColor Cyan
        $EmailPass = Read-Host "   > Please enter the Password for $FromAddress" -AsSecureString
        $EmailCreds = New-Object System.Management.Automation.PSCredential ($FromAddress, $EmailPass)
    }

    # --- CONFIRMATION ---
    Show-Header
    Write-Host "`n[ CONFIRMATION ]" -ForegroundColor Yellow
    Write-Host "   Operation:   $Mode" -ForegroundColor $(If ($Mode -eq 'BACKUP') {'Green'} Else {'Cyan'})
    Write-Host "   Ticket:      $TicketNumber"
    Write-Host "   Customer:    $CustomerName"
    if ($Mode -eq "BACKUP") {
        Write-Host "   Source:      $SourceProfilePath"
        Write-Host "   Destination: $ExternalStorePath"
    } else {
        Write-Host "   Source:      $ExternalStorePath"
        Write-Host "   Destination: THIS COMPUTER ($env:USERPROFILE)"
    }
    
    if ($EmailEnabled) { Write-Host "   Email:       Enabled ($ToAddress)" }

    Write-Host "---------------------------------------------------------------------------------"
    
    $Confirm = Read-Host "Are these details correct? (Press [Enter] or [Y] for Yes, [N] for No)"
    if ($Confirm -eq "") { $Confirm = "Y" }

    if ($Confirm -match "N") {
        Clear-Host
        Write-Host "Restarting input..." -ForegroundColor Red
        Start-Sleep -Seconds 1
    }

} while ($Confirm -match "N")

# --- 7. EXECUTION PREP ---
if ($Mode -eq "RESTORE") {
    $ChromePath64 = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    $ChromePath32 = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    if ((-not (Test-Path $ChromePath64)) -and (-not (Test-Path $ChromePath32))) {
        Install-GoogleChrome
    }
}

$LogPath = "$ExternalStorePath\_Logs"

# Setup Logs
if (!(Test-Path $ExternalStorePath)) { New-Item -ItemType Directory -Path $ExternalStorePath -Force | Out-Null }
if (!(Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath -Force | Out-Null }
$MainLog = "$LogPath\${Mode}_Log.txt"
Add-Content -Path $MainLog -Value "Operation: $Mode | Date: $(Get-Date)"

# Stop Processes - CRITICAL FOR BROWSER DATA
Write-Host "`n[ SYSTEM PREP ]" -ForegroundColor Yellow
if (-not $DemoMode) {
    Write-Host "   Force closing Browsers and Outlook to unlock database files..." -ForegroundColor Red
    # Added braves, operas and generic names to ensure file release
    Stop-Process -Name "chrome", "msedge", "outlook", "firefox", "brave", "opera", "iexplore" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

# --- 8. MAPPING & EXECUTION ---
Write-Host "`n[ STARTING TRANSFER ]" -ForegroundColor Yellow

# Initialize Map - UPDATED FOR BETTER BROWSER COVERAGE
$FoldersMap = [ordered]@{
    "Desktop"   = "Desktop"
    "Documents" = "Documents"
    "Downloads" = "Downloads"
    "Pictures"  = "Pictures"
    "Music"     = "Music"
    "Videos"    = "Videos"
    
    # BROWSER PROFILES (Main)
    "AppData\Local\Google\Chrome\User Data" = "Chrome_UserData"
    "AppData\Local\Microsoft\Edge\User Data" = "Edge_UserData"
    
    # OUTLOOK
    "AppData\Local\Microsoft\Outlook"       = "Outlook_AppData"
    "Documents\Outlook Files"               = "Outlook_Documents"
}

# Extended Paths (Silent Detection)
$ExtendedPaths = @{
    "AppData\Roaming\Mozilla\Firefox"             = "Firefox_Data"
    "AppData\Roaming\Opera Software\Opera Stable" = "Opera_Data"
    "AppData\Local\BraveSoftware\Brave-Browser\User Data" = "Brave_Data"
    "AppData\Roaming"                             = "AppData_Roaming"
    # Capture Local AppData Root (sometimes contains other app configs)
    # "AppData\Local"                             = "AppData_Local" 
}

# Silent Detection (Populate Map with detected paths)
foreach ($RelPath in $ExtendedPaths.Keys) {
    $ExternalName = $ExtendedPaths[$RelPath]
    
    if ($Mode -eq "BACKUP") {
        # Check if source exists (using full path from selection)
        if (Test-Path "$SourceProfilePath\$RelPath") {
            if (-not $FoldersMap.Contains($RelPath)) { $FoldersMap[$RelPath] = $ExternalName }
        }
    }
    elseif ($Mode -eq "RESTORE") {
        if (Test-Path "$ExternalStorePath\$ExternalName") {
            if (-not $FoldersMap.Contains($RelPath)) { $FoldersMap[$RelPath] = $ExternalName }
        }
    }
}

# --- VALIDATE LIST FOR PROGRESS BAR ---
$ValidTransferItems = @()
Write-Host "   Verifying copy list..." -ForegroundColor DarkGray

foreach ($LocalSubPath in $FoldersMap.Keys) {
    $ExternalSubName = $FoldersMap[$LocalSubPath]
    
    # Logic uses $SourceProfilePath (Could be C:\Users\Admin OR E:\Users\OldUser)
    $LocalFull  = "$SourceProfilePath\$LocalSubPath"
    
    # On RESTORE, we write to Current User
    if ($Mode -eq "RESTORE") {
        $LocalFull = "$env:USERPROFILE\$LocalSubPath"
    }

    $ExternalFull = "$ExternalStorePath\$ExternalSubName"

    $ShouldAdd = $false
    if ($Mode -eq "BACKUP") {
        # Verify Source Exists
        if (Test-Path $LocalFull) { 
            $ShouldAdd = $true 
        } else {
            # DEBUG: Uncomment below to see why it skips
            # Write-Host "DEBUG: Skipping $LocalFull - Not Found" -ForegroundColor DarkGray
        }
    } elseif ($Mode -eq "RESTORE") {
        if (Test-Path $ExternalFull) { $ShouldAdd = $true }
    }

    if ($ShouldAdd) {
        $ValidTransferItems += [PSCustomObject]@{
            LocalPath = $LocalSubPath
            ExtName   = $ExternalSubName
            Src       = if ($Mode -eq "BACKUP") { $LocalFull } else { $ExternalFull }
            Dst       = if ($Mode -eq "BACKUP") { $ExternalFull } else { $LocalFull }
        }
    }
}

# Execution Loop with Progress Bar
$ReportItems = @()
$TotalItems = $ValidTransferItems.Count
$CurrentItemIndex = 0

if ($TotalItems -eq 0) {
    Write-Host "   WARNING: No folders found to copy!" -ForegroundColor Red
    Write-Host "   Check if the user profile ($SourceProfilePath) is correct." -ForegroundColor Yellow
}

foreach ($Task in $ValidTransferItems) {
    $CurrentItemIndex++
    
    # Calculate Percentage
    $PercentComplete = ($CurrentItemIndex / $TotalItems) * 100
    
    # Update Progress Bar
    Write-Progress -Activity "Migrating Data ($Mode)" -Status "Processing Folder $CurrentItemIndex of $TotalItems : $($Task.LocalPath)" -PercentComplete $PercentComplete
    
    # Visual Output to Console
    Write-Host "   [$CurrentItemIndex / $TotalItems] Processing: $($Task.LocalPath)" -ForegroundColor Green -NoNewline
    Write-Host " ..." -ForegroundColor Gray

    # Run Copy
    $Result = Run-Robocopy -Source $Task.Src -Destination $Task.Dst -LogFile $MainLog
    
    # Report
    $ReportItems += [PSCustomObject]@{ Item = if ($Mode -eq "BACKUP") {$Task.LocalPath} else {$Task.ExtName}; Status = $Result }
}

# Close Progress Bar
Write-Progress -Activity "Migrating Data ($Mode)" -Completed

# --- 9. REPORT GENERATION ---
Write-Host "`n[ REPORT GENERATION ]" -ForegroundColor Yellow
$CurrentDate = Get-Date -Format "yyyy-MM-dd HH:mm"
$ComputerInfo = Get-CimInstance Win32_ComputerSystem
$TransferTableRows = ""

foreach ($Row in $ReportItems) {
    $StatusColor = if ($Row.Status -match "Completed|Success") { "green" } else { "red" }
    $TransferTableRows += "<tr><td>$($Row.Item)</td><td><span style='color:$StatusColor'>$($Row.Status)</span></td></tr>"
}

$HtmlFile = "$ExternalStorePath\${Mode}_Report.html"
$PdfFile  = "$ExternalStorePath\${Mode}_Report.pdf"

$HtmlContent = @"
<!DOCTYPE html>
<html>
<head>
<style>
    body { font-family: 'Segoe UI', sans-serif; color: #333; padding: 20px; }
    .header { text-align: center; margin-bottom: 20px; }
    h1 { color: #0056b3; margin-bottom: 5px; }
    .meta { font-size: 0.9em; color: #666; text-align: center; margin-bottom: 30px; }
    .section { background: #f9f9f9; padding: 15px; border-left: 6px solid #0056b3; margin-bottom: 20px; }
    table { width: 100%; border-collapse: collapse; font-size: 0.9em; }
    th { text-align: left; background: #eee; padding: 8px; border-bottom: 1px solid #ddd; }
    td { padding: 8px; border-bottom: 1px solid #ddd; }
    .warning { color: #d9534f; font-size: 0.8em; margin-top: 10px; }
</style>
</head>
<body>
<div class="header">
    <img src="$LogoUrl" alt="Apollo Technology" style="max-height:100px;">
    <h1>Data $Mode Report</h1>
    <p>Report generated by <strong>$EngineerName</strong> for ticket (<strong>$TicketNumber</strong>)</p>
    <div class="meta">
        <strong>Date:</strong> $CurrentDate | <strong>Customer:</strong> $CustomerName
    </div>
</div>
<h2>Job Information</h2>
<div class="section">
    <strong>Operation Mode:</strong> $Mode <br>
    <strong>Workstation:</strong> $($ComputerInfo.Name) <br>
    <strong>Source Path:</strong> $SourceProfilePath <br>
    <strong>Engineer:</strong> $EngineerName
</div>
<h2>Transfer Details</h2>
<div class="section">
    <table><thead><tr><th>Data Folder</th><th>Status</th></tr></thead><tbody>$TransferTableRows</tbody></table>
    <p class="warning"><strong>Note on Browser Security:</strong> Browser profile files (Chrome/Edge) have been copied. However, due to Windows DPAPI encryption security, saved passwords are encrypted using the original user's specific account key. They may not auto-decrypt on a new computer unless a cloud sync account was active.</p>
</div>
<p style="text-align:center; font-size:0.8em; color:#888; margin-top:50px;">&copy; $(Get-Date -Format yyyy) by Apollo Technology. All rights reserved. Created by Apollo Technology (Lewis Wiltshire)</p>
</body>
</html>
"@

$HtmlContent | Out-File -FilePath $HtmlFile -Encoding UTF8

# Convert to PDF
$EdgeLoc1 = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
$EdgeLoc2 = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
$EdgeExe = if (Test-Path $EdgeLoc1) { $EdgeLoc1 } elseif (Test-Path $EdgeLoc2) { $EdgeLoc2 } else { $null }

if ($EdgeExe) {
    Write-Host "   Generating PDF Report..." -ForegroundColor Cyan
    $EdgeUserData = "$ExternalStorePath\EdgeTemp"
    if (-not (Test-Path $EdgeUserData)) { New-Item -Path $EdgeUserData -ItemType Directory -Force | Out-Null }
    try {
        Start-Process -FilePath $EdgeExe -ArgumentList "--headless", "--print-to-pdf=`"$PdfFile`"", "--no-pdf-header-footer", "--user-data-dir=`"$EdgeUserData`"", "`"$HtmlFile`"" -Wait
        if (Test-Path $PdfFile) {
            Write-Host "   Report saved to: $PdfFile" -ForegroundColor Green
            Remove-Item $HtmlFile -ErrorAction SilentlyContinue
            Remove-Item $EdgeUserData -Recurse -Force -ErrorAction SilentlyContinue
            Start-Process $PdfFile
        }
    } catch {
        Write-Warning "PDF Conversion failed. Report saved as HTML."
        $PdfFile = $HtmlFile
        Start-Process $HtmlFile
    }
} else {
    Write-Warning "Edge not found. Report saved as HTML."
    $PdfFile = $HtmlFile
    Start-Process $HtmlFile
}

# --- 10. EMAIL REPORT ---
if ($EmailEnabled -and $PdfFile -and (Test-Path $PdfFile)) {
    Write-Host "`n[ EMAIL REPORT ]" -ForegroundColor Yellow
    Write-Host "   Sending Email to $ToAddress..." -ForegroundColor Cyan
    try {
        Send-MailMessage -From $FromAddress -To $ToAddress -Subject "Backup & Recovery Report: $env:COMPUTERNAME ($Mode)" -Body "Attached is the $Mode report for Ticket $TicketNumber ($CustomerName)." -SmtpServer $SmtpServer -Port $SmtpPort -UseSsl $UseSSL -Credential $EmailCreds -Attachments $PdfFile -ErrorAction Stop
        Write-Host "   > Email Sent Successfully!" -ForegroundColor Green
    } catch {
        Write-Error "   > Failed to send email. Error: $_"
    }
}

# --- ALLOW SLEEP AGAIN ---
try { [SleepUtils]::SetThreadExecutionState(0x80000000) | Out-Null } catch { }

Write-Host "`n[ COMPLETE ]" -ForegroundColor Green
Write-Host "Operation finished."

# --- STOP TRANSCRIPT ---
if ($VerboseMode) {
    Write-Host "Stopping Verbose Logging..." -ForegroundColor DarkGray
    Stop-Transcript | Out-Null
}

Pause