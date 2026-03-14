#Requires -Version 5.1
<#
.SYNOPSIS
    InfraEye - Network Diagnostics Alert Script
.DESCRIPTION
    Sends email alerts when network performance thresholds are exceeded.
    Attaches latest HTML network report.
.AUTHOR
    Tushar Gudde
.WEBSITE
    https://tushargudde.tech
.VERSION
    1.0
#>

param(
    [string]$ReportFile    = "",
    [double]$AvgLatency    = 0,
    [double]$PacketLoss    = 0,
    [double]$DownloadSpeed = 0
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

# ─────────────────────────────────────────────────────────────────────────────
# DIRECTORY SETUP
# ─────────────────────────────────────────────────────────────────────────────
$LogDir = Join-Path $PSScriptRoot "Logs"
if (!(Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile   = Join-Path $LogDir "NetworkDiagnosticsAlert_$Timestamp.log"

# ─────────────────────────────────────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────────────────────────────────────
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS")]
        [string]$Level = "INFO"
    )
    $Entry = "{0} | NetworkDiagnosticsAlert | {1} | {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
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
        $alertRows += "<tr><td style='padding:8px;'><span style='background:$color;color:#fff;padding:2px 8px;border-radius:10px;font-size:0.85em;'>$($a.Severity)</span></td><td style='padding:8px;'>$($a.Metric)</td><td style='padding:8px;'>$($a.Value)</td><td style='padding:8px;'>$($a.Threshold)</td></tr>"
    }
    return @"
<html><body style='font-family:Segoe UI,sans-serif;background:#f4f6f8;padding:20px;'>
<div style='max-width:700px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1);'>
    <div style='background:#1a1a2e;color:#fff;padding:20px 30px;'>
        <h2 style='margin:0;'>&#x26A0; InfraEye Network Alert</h2>
        <p style='margin:5px 0 0;opacity:0.8;'>Host: $ComputerName &nbsp;|&nbsp; $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    </div>
    <div style='padding:25px 30px;'>
        <p style='margin-bottom:15px;'>Network performance thresholds exceeded on <strong>$ComputerName</strong>:</p>
        <table style='width:100%;border-collapse:collapse;'>
            <thead><tr style='background:#0d6efd;color:#fff;'><th style='padding:8px;text-align:left;'>Severity</th><th style='padding:8px;text-align:left;'>Metric</th><th style='padding:8px;text-align:left;'>Current Value</th><th style='padding:8px;text-align:left;'>Threshold</th></tr></thead>
            <tbody>$alertRows</tbody>
        </table>
        <p style='margin-top:15px;font-size:0.9em;color:#6c757d;'>Please review the attached HTML report for full diagnostics details.</p>
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
    Write-Log "=== Network Diagnostics Alert Check Started ===" "INFO"

    $config       = Get-Config
    $thresholds   = $config.alert_thresholds
    $computerName = $env:COMPUTERNAME
    $alerts       = @()

    # If no values supplied, perform live checks
    if ($AvgLatency -eq 0 -and $PacketLoss -eq 0) {
        Write-Log "No values supplied. Running live ping tests..." "INFO"
        $targets = @("8.8.8.8","1.1.1.1")
        $latencies = @()
        $losses    = @()
        foreach ($t in $targets) {
            try {
                $ping = New-Object System.Net.NetworkInformation.Ping
                $replies = 1..5 | ForEach-Object {
                    try { $ping.Send($t, 2000) } catch { $null }
                }
                $ping.Dispose()
                $ok   = ($replies | Where-Object { $_ -and $_.Status -eq "Success" })
                $fail = 5 - $ok.Count
                if ($ok.Count -gt 0) {
                    $latencies += ($ok | Measure-Object -Property RoundtripTime -Average).Average
                }
                $losses += [math]::Round(($fail / 5) * 100, 0)
            } catch {
                $losses += 100
            }
        }
        $AvgLatency = if ($latencies.Count -gt 0) { [math]::Round(($latencies | Measure-Object -Average).Average, 1) } else { 9999 }
        $PacketLoss = if ($losses.Count -gt 0) { ($losses | Measure-Object -Maximum).Maximum } else { 100 }
    }

    # Threshold checks
    if ($AvgLatency -gt $thresholds.latency_ms) {
        $alerts += [PSCustomObject]@{ Severity = "CRITICAL"; Metric = "Avg Latency"; Value = "$AvgLatency ms"; Threshold = ">$($thresholds.latency_ms) ms" }
        Write-Log "ALERT: Latency $AvgLatency ms (threshold: $($thresholds.latency_ms) ms)" "WARN"
    }
    if ($PacketLoss -gt $thresholds.packet_loss_percent) {
        $alerts += [PSCustomObject]@{ Severity = "CRITICAL"; Metric = "Packet Loss"; Value = "$PacketLoss%"; Threshold = ">$($thresholds.packet_loss_percent)%" }
        Write-Log "ALERT: Packet loss $PacketLoss% (threshold: $($thresholds.packet_loss_percent)%)" "WARN"
    }
    if ($DownloadSpeed -gt 0 -and $DownloadSpeed -lt $thresholds.download_speed_mbps) {
        $alerts += [PSCustomObject]@{ Severity = "WARNING"; Metric = "Download Speed"; Value = "$DownloadSpeed Mbps"; Threshold = "<$($thresholds.download_speed_mbps) Mbps" }
        Write-Log "ALERT: Download speed $DownloadSpeed Mbps (threshold: $($thresholds.download_speed_mbps) Mbps)" "WARN"
    }

    if ($alerts.Count -gt 0) {
        Write-Log "$($alerts.Count) alert(s) triggered. Sending email..." "WARN"

        if (-not $ReportFile -or !(Test-Path $ReportFile)) {
            $ReportDir = Join-Path $PSScriptRoot "Reports"
            if (Test-Path $ReportDir) {
                $latestReport = Get-ChildItem -Path $ReportDir -Filter "NetworkDiagnostics_*.html" |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -First 1
                if ($latestReport) { $ReportFile = $latestReport.FullName }
            }
        }

        $emailBody = New-AlertBody -Alerts $alerts -ComputerName $computerName
        Send-AlertEmail `
            -Subject        "[InfraEye ALERT] Network Issue on $computerName" `
            -Body           $emailBody `
            -AttachmentPath $ReportFile

        Write-Log "Alert processing completed." "SUCCESS"
    } else {
        Write-Log "No network thresholds exceeded. No alert sent." "INFO"
    }

    Write-Log "=== Network Diagnostics Alert Check Completed ===" "SUCCESS"
} catch {
    Write-Log "FATAL ERROR in network alert script: $_" "ERROR"
    exit 1
}
