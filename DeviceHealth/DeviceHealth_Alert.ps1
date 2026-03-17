#Requires -Version 5.1
<#
.SYNOPSIS
    InfraEye - Device Health Alert Script
.DESCRIPTION
    Sends email alerts when device health thresholds are exceeded.
    Attach latest HTML report to the alert email.
.AUTHOR
    Tushar Gudde
.WEBSITE
    https://tushargudde.tech
.VERSION
    2.0
#>

param(
    [string]$ReportFile   = "",
    [double]$CPU          = 0,
    [double]$RAM          = 0,
    [int]   $DiskFreePct  = 100,
    [int]   $StartupCount = 0
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

# ─────────────────────────────────────────────────────────────────────────────
# DIRECTORY SETUP
# ─────────────────────────────────────────────────────────────────────────────
$LogDir  = Join-Path $PSScriptRoot "Logs"
if (!(Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile   = Join-Path $LogDir "DeviceHealthAlert_$Timestamp.log"

# ─────────────────────────────────────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────────────────────────────────────
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS")]
        [string]$Level = "INFO"
    )
    $Entry = "{0} | DeviceHealthAlert | {1} | {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -Path $LogFile -Value $Entry
    switch ($Level) {
        "INFO"    { Write-Host $Entry -ForegroundColor Cyan }
        "WARN"    { Write-Host $Entry -ForegroundColor Yellow }
        "ERROR"   { Write-Host $Entry -ForegroundColor Red }
        "SUCCESS" { Write-Host $Entry -ForegroundColor Green }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# LOAD CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────
function Get-Config {
    $configPath = Join-Path $PSScriptRoot "..\config.json"
    if (!(Test-Path $configPath)) {
        Write-Log "config.json not found at $configPath" "ERROR"
        throw "config.json not found."
    }
    return Get-Content -Path $configPath -Raw | ConvertFrom-Json
}

# ─────────────────────────────────────────────────────────────────────────────
# SEND EMAIL ALERT
# ─────────────────────────────────────────────────────────────────────────────
function Send-AlertEmail {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingConvertToSecureStringWithPlainText', '',
        Justification = 'Password is read from an external config file, not hardcoded in source.')]
    param(
        [string]$Subject,
        [string]$Body,
        [string]$AttachmentPath = ""
    )
    try {
        $config = Get-Config

        $mailParams = @{
            SmtpServer  = $config.smtp_server
            Port        = $config.smtp_port
            From        = $config.from_email
            To          = $config.to_email
            Subject     = $Subject
            Body        = $Body
            BodyAsHtml  = $true
            UseSsl      = $config.enable_ssl
        }

        if ($config.smtp_username -and $config.smtp_password) {
            $securePass = ConvertTo-SecureString $config.smtp_password -AsPlainText -Force
            $cred = New-Object System.Management.Automation.PSCredential($config.smtp_username, $securePass)
            $mailParams["Credential"] = $cred
        }

        if ($AttachmentPath -and (Test-Path $AttachmentPath)) {
            $mailParams["Attachments"] = $AttachmentPath
        }

        Send-MailMessage @mailParams
        Write-Log "Alert email sent to $($config.to_email)" "SUCCESS"
    } catch {
        Write-Log "Failed to send alert email: $_" "ERROR"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# BUILD ALERT MESSAGE
# ─────────────────────────────────────────────────────────────────────────────
function New-AlertBody {
    param(
        [array]$Alerts,
        [string]$ComputerName
    )
    $alertRows = ""
    foreach ($a in $Alerts) {
        $color = switch ($a.Severity) {
            "CRITICAL" { "#dc3545" }
            "WARNING"  { "#ffc107" }
            default    { "#6c757d" }
        }
        $alertRows += "<tr><td style='padding:8px;'><span style='background:$color;color:#fff;padding:2px 8px;border-radius:10px;font-size:0.85em;'>$($a.Severity)</span></td><td style='padding:8px;'>$($a.Metric)</td><td style='padding:8px;'>$($a.Value)</td><td style='padding:8px;'>$($a.Threshold)</td></tr>"
    }
    return @"
<html><body style='font-family:Segoe UI,sans-serif;background:#f4f6f8;padding:20px;'>
<div style='max-width:700px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1);'>
    <div style='background:#1a1a2e;color:#fff;padding:20px 30px;'>
        <h2 style='margin:0;'>&#x26A0; InfraEye Device Health Alert</h2>
        <p style='margin:5px 0 0;opacity:0.8;'>Computer: $ComputerName &nbsp;|&nbsp; $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    </div>
    <div style='padding:25px 30px;'>
        <p style='margin-bottom:15px;'>The following thresholds have been exceeded on <strong>$ComputerName</strong>:</p>
        <table style='width:100%;border-collapse:collapse;'>
            <thead><tr style='background:#0d6efd;color:#fff;'><th style='padding:8px;text-align:left;'>Severity</th><th style='padding:8px;text-align:left;'>Metric</th><th style='padding:8px;text-align:left;'>Current Value</th><th style='padding:8px;text-align:left;'>Threshold</th></tr></thead>
            <tbody>$alertRows</tbody>
        </table>
        <p style='margin-top:15px;font-size:0.9em;color:#6c757d;'>Please review the attached HTML report for detailed diagnostics.</p>
    </div>
    <div style='background:#f8f9fa;padding:15px 30px;text-align:center;font-size:0.8em;color:#6c757d;'>
        Report Version: 2.0 | Created by: Tushar Gudde | <a href='https://tushargudde.tech'>https://tushargudde.tech</a>
    </div>
</div>
</body></html>
"@
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN EXECUTION
# ─────────────────────────────────────────────────────────────────────────────
try {
    Write-Log "=== Device Health Alert Check Started ===" "INFO"

    $config      = Get-Config
    $thresholds  = $config.alert_thresholds
    $computerName = $env:COMPUTERNAME
    $alerts       = @()

    # If called directly, re-collect live data
    if ($CPU -eq 0) {
        try {
            $CPU = (Get-CimInstance -ClassName Win32_Processor | Measure-Object -Property LoadPercentage -Average).Average
        } catch { $CPU = 0 }
    }
    if ($RAM -eq 0) {
        try {
            $os  = Get-CimInstance -ClassName Win32_OperatingSystem
            $tot = $os.TotalVisibleMemorySize
            $free= $os.FreePhysicalMemory
            $RAM = if ($tot -gt 0) { [math]::Round((($tot - $free) / $tot) * 100, 1) } else { 0 }
        } catch { $RAM = 0 }
    }
    if ($DiskFreePct -eq 100) {
        try {
            $DiskFreePct = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" |
                ForEach-Object { if ($_.Size -gt 0) { [math]::Round($_.FreeSpace / $_.Size * 100, 1) } else { 100 } } |
                Measure-Object -Minimum).Minimum
        } catch { $DiskFreePct = 100 }
    }
    if ($StartupCount -eq 0) {
        try {
            $regPaths = @(
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
                "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
            )
            $StartupCount = ($regPaths | ForEach-Object {
                if (Test-Path $_) { (Get-ItemProperty -Path $_).PSObject.Properties | Where-Object { $_.Name -notmatch "^PS" } }
            } | Measure-Object).Count
        } catch { $StartupCount = 0 }
    }

    # Threshold checks
    if ($CPU -gt $thresholds.cpu_percent) {
        $alerts += [PSCustomObject]@{ Severity = "CRITICAL"; Metric = "CPU Usage"; Value = "$CPU%"; Threshold = ">$($thresholds.cpu_percent)%" }
        Write-Log "ALERT: CPU at $CPU% (threshold: $($thresholds.cpu_percent)%)" "WARN"
    }
    if ($RAM -gt $thresholds.ram_percent) {
        $alerts += [PSCustomObject]@{ Severity = "CRITICAL"; Metric = "RAM Usage"; Value = "$RAM%"; Threshold = ">$($thresholds.ram_percent)%" }
        Write-Log "ALERT: RAM at $RAM% (threshold: $($thresholds.ram_percent)%)" "WARN"
    }
    if ($DiskFreePct -lt $thresholds.disk_free_percent) {
        $alerts += [PSCustomObject]@{ Severity = "CRITICAL"; Metric = "Disk Free Space"; Value = "$DiskFreePct%"; Threshold = "<$($thresholds.disk_free_percent)%" }
        Write-Log "ALERT: Disk free at $DiskFreePct% (threshold: $($thresholds.disk_free_percent)%)" "WARN"
    }
    if ($StartupCount -gt $thresholds.startup_apps_max) {
        $alerts += [PSCustomObject]@{ Severity = "WARNING"; Metric = "Startup Programs"; Value = $StartupCount; Threshold = ">$($thresholds.startup_apps_max)" }
        Write-Log "ALERT: $StartupCount startup apps (threshold: $($thresholds.startup_apps_max))" "WARN"
    }

    if ($alerts.Count -gt 0) {
        Write-Log "$($alerts.Count) alert(s) triggered. Sending email..." "WARN"

        # Find latest report if not provided
        if (-not $ReportFile -or !(Test-Path $ReportFile)) {
            $ReportDir = Join-Path $PSScriptRoot "Reports"
            if (Test-Path $ReportDir) {
                $latestReport = Get-ChildItem -Path $ReportDir -Filter "DeviceHealth_*.html" |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -First 1
                if ($latestReport) { $ReportFile = $latestReport.FullName }
            }
        }

        $emailBody = New-AlertBody -Alerts $alerts -ComputerName $computerName
        Send-AlertEmail `
            -Subject  "[InfraEye ALERT] Device Health Issue on $computerName" `
            -Body     $emailBody `
            -AttachmentPath $ReportFile

        Write-Log "Alert processing completed." "SUCCESS"
    } else {
        Write-Log "No thresholds exceeded. No alert sent." "INFO"
    }

    Write-Log "=== Device Health Alert Check Completed ===" "SUCCESS"
} catch {
    Write-Log "FATAL ERROR in alert script: $_" "ERROR"
    exit 1
}
