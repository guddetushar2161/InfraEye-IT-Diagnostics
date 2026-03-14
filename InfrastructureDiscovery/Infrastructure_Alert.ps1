#Requires -Version 5.1
<#
.SYNOPSIS
    InfraEye - Infrastructure Discovery Alert Script
.DESCRIPTION
    Sends email alerts for infrastructure issues including:
    duplicate IPs, DHCP pool exhaustion, offline servers, and unknown devices.
.AUTHOR
    Tushar Gudde
.WEBSITE
    https://tushargudde.tech
.VERSION
    1.0
#>

param(
    [string]$ReportFile     = "",
    [int]   $DuplicateCount = 0,
    [int]   $UnknownDevices = 0,
    [int]   $DHCPCritical   = 0,
    [int]   $OfflineServers = 0
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

# ─────────────────────────────────────────────────────────────────────────────
# DIRECTORY SETUP
# ─────────────────────────────────────────────────────────────────────────────
$LogDir = Join-Path $PSScriptRoot "Logs"
if (!(Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile   = Join-Path $LogDir "InfrastructureAlert_$Timestamp.log"

# ─────────────────────────────────────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────────────────────────────────────
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS")]
        [string]$Level = "INFO"
    )
    $Entry = "{0} | InfrastructureAlert | {1} | {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
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
            SmtpServer = $config.smtp_server
            Port       = $config.smtp_port
            From       = $config.from_email
            To         = $config.to_email
            Subject    = $Subject
            Body       = $Body
            BodyAsHtml = $true
            UseSsl     = $config.enable_ssl
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
# BUILD ALERT HTML BODY
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
        $alertRows += "<tr><td style='padding:8px;'><span style='background:$color;color:#fff;padding:2px 8px;border-radius:10px;font-size:0.85em;'>$($a.Severity)</span></td><td style='padding:8px;'>$($a.Issue)</td><td style='padding:8px;'>$($a.Detail)</td></tr>"
    }
    return @"
<html><body style='font-family:Segoe UI,sans-serif;background:#f4f6f8;padding:20px;'>
<div style='max-width:700px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1);'>
    <div style='background:#1a1a2e;color:#fff;padding:20px 30px;'>
        <h2 style='margin:0;'>&#x26A0; InfraEye Infrastructure Alert</h2>
        <p style='margin:5px 0 0;opacity:0.8;'>Detected from: $ComputerName &nbsp;|&nbsp; $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    </div>
    <div style='padding:25px 30px;'>
        <p style='margin-bottom:15px;'>The following infrastructure issues have been detected:</p>
        <table style='width:100%;border-collapse:collapse;'>
            <thead><tr style='background:#0d6efd;color:#fff;'><th style='padding:8px;text-align:left;'>Severity</th><th style='padding:8px;text-align:left;'>Issue</th><th style='padding:8px;text-align:left;'>Details</th></tr></thead>
            <tbody>$alertRows</tbody>
        </table>
        <p style='margin-top:15px;font-size:0.9em;color:#6c757d;'>Please review the attached HTML report for full infrastructure details.</p>
    </div>
    <div style='background:#f8f9fa;padding:15px 30px;text-align:center;font-size:0.8em;color:#6c757d;'>
        Report Version: 1.0 | Created by: Tushar Gudde | <a href='https://tushargudde.tech'>https://tushargudde.tech</a>
    </div>
</div>
</body></html>
"@
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN EXECUTION
# ─────────────────────────────────────────────────────────────────────────────
try {
    Write-Log "=== Infrastructure Discovery Alert Check Started ===" "INFO"

    $config       = Get-Config
    $thresholds   = $config.alert_thresholds
    $computerName = $env:COMPUTERNAME
    $alerts       = @()

    # Threshold checks
    if ($DuplicateCount -gt 0) {
        $alerts += [PSCustomObject]@{
            Severity = "CRITICAL"
            Issue    = "Duplicate IP Addresses"
            Detail   = "$DuplicateCount duplicate IP(s) detected on the network. This may cause connectivity issues."
        }
        Write-Log "ALERT: $DuplicateCount duplicate IP addresses detected." "WARN"
    }

    if ($DHCPCritical -gt 0) {
        $alerts += [PSCustomObject]@{
            Severity = "CRITICAL"
            Issue    = "DHCP Pool Near Exhaustion"
            Detail   = "$DHCPCritical network(s) have DHCP pool usage above $($thresholds.dhcp_pool_percent)%. New devices may not get IP addresses."
        }
        Write-Log "ALERT: DHCP pool usage critical on $DHCPCritical network(s)." "WARN"
    }

    if ($OfflineServers -gt 0) {
        $alerts += [PSCustomObject]@{
            Severity = "CRITICAL"
            Issue    = "Offline Servers Detected"
            Detail   = "$OfflineServers server(s) are not responding to ping. Immediate investigation required."
        }
        Write-Log "ALERT: $OfflineServers offline server(s) detected." "WARN"
    }

    if ($UnknownDevices -gt 0) {
        $alerts += [PSCustomObject]@{
            Severity = "WARNING"
            Issue    = "Unknown Devices Detected"
            Detail   = "$UnknownDevices unclassified device(s) found on the network. Verify these are authorized devices."
        }
        Write-Log "ALERT: $UnknownDevices unknown devices detected." "WARN"
    }

    if ($alerts.Count -gt 0) {
        Write-Log "$($alerts.Count) infrastructure alert(s) triggered. Sending email..." "WARN"

        if (-not $ReportFile -or !(Test-Path $ReportFile)) {
            $ReportDir = Join-Path $PSScriptRoot "Reports"
            if (Test-Path $ReportDir) {
                $latestReport = Get-ChildItem -Path $ReportDir -Filter "InfrastructureDiscovery_*.html" |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -First 1
                if ($latestReport) { $ReportFile = $latestReport.FullName }
            }
        }

        $emailBody = New-AlertBody -Alerts $alerts -ComputerName $computerName
        Send-AlertEmail `
            -Subject        "[InfraEye ALERT] Infrastructure Issue Detected on $computerName" `
            -Body           $emailBody `
            -AttachmentPath $ReportFile

        Write-Log "Alert processing completed." "SUCCESS"
    } else {
        Write-Log "No infrastructure alert thresholds exceeded. No alert sent." "INFO"
    }

    Write-Log "=== Infrastructure Discovery Alert Check Completed ===" "SUCCESS"
} catch {
    Write-Log "FATAL ERROR in infrastructure alert script: $_" "ERROR"
    exit 1
}
