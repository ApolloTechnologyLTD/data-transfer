<#
.SYNOPSIS
    Apollo Technology Data Migration Utility (Smart Restore & Backup) v2.6
.DESCRIPTION
    Menu-driven utility to Backup data or Restore data.
    - UPDATED: Header now includes Admin Notice, Credits, and Power Status.
    - INCLUDES: Anti-Sleep, Anti-Freeze (QuickEdit), Email Reports, Input Validation.
    - SILENT DETECTION for Edge/AppData.
    - AUTO-INSTALLS Google Chrome if missing.
#>

# --- 0. CONFIGURATION ---
$DemoMode = $false
$LogoUrl = "https://raw.githubusercontent.com/ApolloTechnologyLTD/computer-health-check/main/Apollo%20Cropped.png"
$Version = "2.6"

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
 / ___ |/ ____/ /_/ / /___/ /___/ /_/ /    / / / /___/ /___/ __  / /|  / /_/ / /___/ /_/ / /_/ /   / /  
/_/  |_/_/    \____/_____/_____/\____/    /_/ /_____/\____/_/ /_/_/ |_/\____/_____/\____/\____/   /_/   
'@
    Write-Host $Banner -ForegroundColor Cyan
    Write-Host "`n   DATA MIGRATION & BACKUP TOOL" -ForegroundColor White
    Write-Host "=================================================================================" -ForegroundColor DarkGray

    # --- ADDED: NOTICE & DETAILS ---
    # Re-check admin status for the header display
    $Current = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $IsAdminHeader = $Current.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($IsAdminHeader) { 
        Write-Host "        [NOTICE] Running in Elevated Permissions" -ForegroundColor Red 
    } elseif ($DemoMode) {
        Write-Host "      [NOTICE] Running as Standard User" -ForegroundColor Yellow 
    }

    Write-Host "        Created by Lewis Wiltshire, Version $Version" -ForegroundColor Yellow
    Write-Host "      [POWER] Sleep Mode & Screen Timeout Blocked." -ForegroundColor DarkGray
    # -------------------------------

    if ($DemoMode) {
        Write-Host "`n   *** DEMO MODE ACTIVE - NO REAL DATA WILL BE COPIED ***" -ForegroundColor Magenta
    }
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

function Run-Robocopy {
    param (
        [string]$Source,
        [string]$Destination,
        [string]$LogFile
    )

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
            Write-Host "   Copying: $Source" -ForegroundColor Green
            robocopy $Source $Destination /E /XO /R:1 /W:1 /NP /LOG+:"$LogFile" /TEE /XD $Excludes | Out-Null
            return "Completed"
        } else {
            Write-Host "   Skipping: Source not found ($Source)" -ForegroundColor DarkGray
            Add-Content -Path $LogFile -Value "SKIPPED: Source Missing - $Source"
            return "Skipped (Not Found)"
        }
    }
}

# --- 4. MAIN MENU ---
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

# --- 5. INPUT COLLECTION LOOP ---
do {
    Show-Header
    Write-Host "`n[ $Mode CONFIGURATION ]" -ForegroundColor Yellow

    $EngineerName = Read-Host "   > Enter Engineer Name"
    $TicketNumber = Read-Host "   > Enter Ticket Number"

    # Drive Selection
    Write-Host "`n[ DRIVE SELECTION ]" -ForegroundColor Yellow
    Write-Host "Detecting drives..." -ForegroundColor DarkGray
    Get-PSDrive -PSProvider FileSystem | Select-Object Name, Used, Free, Root | Format-Table -AutoSize

    $DriveLetterInput = Read-Host "   > Enter External Drive Letter (e.g. D or E)"
    $DriveLetter = $DriveLetterInput -replace ":", ""

    if (!(Test-Path "$($DriveLetter):")) {
        Write-Error "Drive $($DriveLetter): not found."
        Write-Host "Restarting input..." -ForegroundColor Red
        Start-Sleep -Seconds 2
        Continue 
    }

    # PATH LOGIC & SEARCH
    $UserProfile = $env:USERPROFILE
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
        Write-Host "   Source:      THIS COMPUTER"
        Write-Host "   Destination: $ExternalStorePath"
    } else {
        Write-Host "   Source:      $ExternalStorePath"
        Write-Host "   Destination: THIS COMPUTER"
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

# --- 6. EXECUTION PREP ---
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

# Stop Processes
Write-Host "`n[ SYSTEM PREP ]" -ForegroundColor Yellow
if (-not $DemoMode) {
    Write-Host "Stopping Browsers and Outlook..." -ForegroundColor Gray
    Stop-Process -Name "chrome", "msedge", "outlook", "firefox" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

# --- 7. MAPPING & EXECUTION ---
Write-Host "`n[ STARTING TRANSFER ]" -ForegroundColor Yellow

# Initialize Map
$FoldersMap = [ordered]@{
    "Desktop"   = "Desktop"
    "Documents" = "Documents"
    "Downloads" = "Downloads"
    "Pictures"  = "Pictures"
    "Music"     = "Music"
    "Videos"    = "Videos"
    "AppData\Local\Google\Chrome\User Data" = "Chrome_UserData"
    "AppData\Local\Microsoft\Outlook"       = "Outlook_AppData"
    "Documents\Outlook Files"               = "Outlook_Documents"
}

# Extended Paths
$ExtendedPaths = @{
    "AppData\Local\Microsoft\Edge\User Data"      = "Edge_UserData"
    "AppData\Roaming\Mozilla\Firefox"             = "Firefox_Data"
    "AppData\Roaming\Opera Software\Opera Stable" = "Opera_Data"
    "AppData\Local\BraveSoftware\Brave-Browser\User Data" = "Brave_Data"
    "AppData\Roaming"                             = "AppData_Roaming"
    "AppData\Local"                               = "AppData_Local"
}

# Silent Detection
foreach ($RelPath in $ExtendedPaths.Keys) {
    $ExternalName = $ExtendedPaths[$RelPath]
    
    if ($Mode -eq "BACKUP") {
        if (Test-Path "$UserProfile\$RelPath") {
            if (-not $FoldersMap.Contains($RelPath)) { $FoldersMap[$RelPath] = $ExternalName }
        }
    }
    elseif ($Mode -eq "RESTORE") {
        if (Test-Path "$ExternalStorePath\$ExternalName") {
            if (-not $FoldersMap.Contains($RelPath)) { $FoldersMap[$RelPath] = $ExternalName }
        }
    }
}

# Execution Loop
$ReportItems = @()

foreach ($LocalSubPath in $FoldersMap.Keys) {
    $ExternalSubName = $FoldersMap[$LocalSubPath]
    $LocalFull  = "$UserProfile\$LocalSubPath"
    $ExternalFull = "$ExternalStorePath\$ExternalSubName"

    if ($Mode -eq "BACKUP") {
        $Result = Run-Robocopy -Source $LocalFull -Destination $ExternalFull -LogFile $MainLog
        $ReportItems += [PSCustomObject]@{ Item = $LocalSubPath; Status = $Result }
    }
    elseif ($Mode -eq "RESTORE") {
        if (Test-Path $ExternalFull) {
            $Result = Run-Robocopy -Source $ExternalFull -Destination $LocalFull -LogFile $MainLog
            $ReportItems += [PSCustomObject]@{ Item = $ExternalSubName; Status = $Result }
        }
    }
}

# --- 8. REPORT GENERATION ---
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
    <strong>Engineer:</strong> $EngineerName <br>
    <strong>Storage Path:</strong> $ExternalStorePath
</div>
<h2>Transfer Details</h2>
<div class="section">
    <table><thead><tr><th>Data Folder</th><th>Status</th></tr></thead><tbody>$TransferTableRows</tbody></table>
</div>
<p style="text-align:center; font-size:0.8em; color:#888; margin-top:50px;">&copy; $(Get-Date -Format yyyy) by Apollo Technology.</p>
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

# --- 9. EMAIL REPORT ---
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
Pause