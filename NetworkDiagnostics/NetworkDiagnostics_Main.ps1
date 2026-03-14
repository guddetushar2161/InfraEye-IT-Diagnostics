#Requires -Version 5.1
<#
.SYNOPSIS
    InfraEye - Network Diagnostics Module
.DESCRIPTION
    Collects network adapter info, public IP, runs ping tests, measures bandwidth-heavy
    processes, and performs network root cause analysis. Generates HTML dashboard report.
.AUTHOR
    Tushar Gudde
.WEBSITE
    https://tushargudde.tech
.VERSION
    1.0
#>

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

# ─────────────────────────────────────────────────────────────────────────────
# DIRECTORY SETUP
# ─────────────────────────────────────────────────────────────────────────────
$ReportDir = Join-Path $PSScriptRoot "Reports"
$LogDir    = Join-Path $PSScriptRoot "Logs"

if (!(Test-Path $ReportDir)) { New-Item -ItemType Directory -Path $ReportDir | Out-Null }
if (!(Test-Path $LogDir))    { New-Item -ItemType Directory -Path $LogDir    | Out-Null }

$Timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile    = Join-Path $LogDir "NetworkDiagnostics_$Timestamp.log"
$ReportFile = Join-Path $ReportDir "NetworkDiagnostics_$Timestamp.html"

# ─────────────────────────────────────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────────────────────────────────────
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS")]
        [string]$Level = "INFO"
    )
    $Entry = "{0} | NetworkDiagnostics | {1} | {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -Path $LogFile -Value $Entry
    switch ($Level) {
        "INFO"    { Write-Host $Entry -ForegroundColor Cyan }
        "WARN"    { Write-Host $Entry -ForegroundColor Yellow }
        "ERROR"   { Write-Host $Entry -ForegroundColor Red }
        "SUCCESS" { Write-Host $Entry -ForegroundColor Green }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# DEPENDENCY CHECK
# ─────────────────────────────────────────────────────────────────────────────
function Install-RequiredModules {
    # Ensure NuGet provider is available for Install-Module
    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Write-Log "NuGet package provider not found. Installing..." "WARN"
        try {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
            Write-Log "NuGet provider installed successfully." "SUCCESS"
        } catch {
            Write-Log "Failed to install NuGet provider: $_" "WARN"
        }
    }

    $modules = @("ImportExcel", "PSWriteHTML")
    foreach ($mod in $modules) {
        if (-not (Get-Module -ListAvailable -Name $mod)) {
            Write-Log "Module '$mod' not found. Installing..." "WARN"
            try {
                Install-Module -Name $mod -Scope CurrentUser -Force -AllowClobber
                Write-Log "Module '$mod' installed successfully." "SUCCESS"
            } catch {
                Write-Log "Failed to install module '$mod': $_" "ERROR"
            }
        } else {
            Write-Log "Module '$mod' is available." "INFO"
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# DATA COLLECTION FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────
function Get-NetworkAdapterInfo {
    Write-Log "Collecting network adapter information..." "INFO"
    $adapters = @()
    try {
        $netAdapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True"
        foreach ($adapter in $netAdapters) {
            $adapterDetails = Get-CimInstance -ClassName Win32_NetworkAdapter -Filter "Index=$($adapter.Index)" -ErrorAction SilentlyContinue
            $speed = if ($adapterDetails -and $adapterDetails.Speed) {
                [math]::Round($adapterDetails.Speed / 1MB, 0)
            } else { 0 }

            $connectionType = if ($adapterDetails -and $adapterDetails.Name -match "Wi-Fi|Wireless|WLAN|802.11") {
                "WiFi"
            } elseif ($adapterDetails -and $adapterDetails.Name -match "Ethernet|LAN|Local Area") {
                "LAN"
            } else {
                "Other"
            }

            $adapters += [PSCustomObject]@{
                Name           = $adapter.Description
                ConnectionType = $connectionType
                MACAddress     = $adapter.MACAddress
                IPAddress      = ($adapter.IPAddress | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' }) -join ", "
                SubnetMask     = ($adapter.IPSubnet  | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' }) -join ", "
                DefaultGateway = ($adapter.DefaultIPGateway) -join ", "
                DNSServers     = ($adapter.DNSServerSearchOrder) -join ", "
                LinkSpeed_Mbps = $speed
                DHCPEnabled    = $adapter.DHCPEnabled
            }
        }
    } catch {
        Write-Log "Error collecting network adapter info: $_" "ERROR"
        throw
    }
    return $adapters
}

function Get-PublicIPInfo {
    Write-Log "Fetching public IP information from ipinfo.io..." "INFO"
    try {
        $response = Invoke-RestMethod -Uri "https://ipinfo.io/json" -TimeoutSec 10 -ErrorAction Stop
        return [PSCustomObject]@{
            PublicIP = $response.ip
            ISP      = $response.org
            City     = $response.city
            Region   = $response.region
            Country  = $response.country
            Timezone = $response.timezone
        }
    } catch {
        Write-Log "Could not fetch public IP info (API unavailable): $_" "WARN"
        return [PSCustomObject]@{
            PublicIP = "Unavailable"
            ISP      = "Unavailable"
            City     = "Unavailable"
            Region   = "Unavailable"
            Country  = "Unavailable"
            Timezone = "Unavailable"
        }
    }
}

function Invoke-PingTest {
    param([string[]]$Targets = @("8.8.8.8","1.1.1.1","8.8.4.4"))
    Write-Log "Running ping tests to: $($Targets -join ', ')..." "INFO"
    $results = @()
    foreach ($target in $Targets) {
        try {
            $pingResults = @()
            $failed = 0
            1..10 | ForEach-Object {
                $ping = New-Object System.Net.NetworkInformation.Ping
                try {
                    $reply = $ping.Send($target, 3000)
                    if ($reply.Status -eq "Success") {
                        $pingResults += $reply.RoundtripTime
                    } else {
                        $failed++
                    }
                } catch {
                    $failed++
                } finally {
                    $ping.Dispose()
                }
            }

            $avgLatency  = if ($pingResults.Count -gt 0) { [math]::Round(($pingResults | Measure-Object -Average).Average, 1) } else { 9999 }
            $minLatency  = if ($pingResults.Count -gt 0) { ($pingResults | Measure-Object -Minimum).Minimum } else { 9999 }
            $maxLatency  = if ($pingResults.Count -gt 0) { ($pingResults | Measure-Object -Maximum).Maximum } else { 9999 }
            $packetLoss  = [math]::Round(($failed / 10) * 100, 0)

            Write-Log "Ping $target`: avg=$avgLatency ms, loss=$packetLoss%" "INFO"
            $results += [PSCustomObject]@{
                Target      = $target
                AvgLatency  = $avgLatency
                MinLatency  = $minLatency
                MaxLatency  = $maxLatency
                PacketLoss  = $packetLoss
                Status      = if ($packetLoss -eq 100) { "Unreachable" } elseif ($packetLoss -gt 5) { "Degraded" } else { "OK" }
            }
        } catch {
            Write-Log "Ping test failed for $target`: $_" "WARN"
            $results += [PSCustomObject]@{
                Target     = $target; AvgLatency = 9999; MinLatency = 9999; MaxLatency = 9999
                PacketLoss = 100; Status = "Error"
            }
        }
    }
    return $results
}

function Invoke-SpeedTest {
    Write-Log "Estimating download speed..." "INFO"
    try {
        $url        = "http://speedtest.tele2.net/10MB.zip"
        $startTime  = Get-Date
        $wc         = New-Object System.Net.WebClient
        $tempFile   = [System.IO.Path]::GetTempFileName()
        try {
            $wc.DownloadFile($url, $tempFile)
            $elapsed    = ((Get-Date) - $startTime).TotalSeconds
            $fileSize   = (Get-Item $tempFile).Length
            $speedMbps  = [math]::Round(($fileSize * 8) / ($elapsed * 1MB), 2)
            Write-Log "Estimated download speed: $speedMbps Mbps" "INFO"
            return [PSCustomObject]@{
                DownloadMbps = $speedMbps
                TestFile     = "10 MB from speedtest.tele2.net"
                Duration_s   = [math]::Round($elapsed, 2)
                Status       = "Completed"
            }
        } finally {
            $wc.Dispose()
            if (Test-Path $tempFile) { Remove-Item $tempFile -Force -ErrorAction SilentlyContinue }
        }
    } catch {
        Write-Log "Speed test failed (network/timeout): $_" "WARN"
        return [PSCustomObject]@{
            DownloadMbps = 0
            TestFile     = "N/A"
            Duration_s   = 0
            Status       = "Failed"
        }
    }
}

function Get-NetworkUsage {
    Write-Log "Detecting bandwidth-heavy processes..." "INFO"
    $netProcesses = @()
    try {
        $netConnections = Get-NetTCPConnection -State Established -ErrorAction SilentlyContinue
        if ($netConnections) {
            $processGroups = $netConnections | Group-Object OwningProcess | Sort-Object Count -Descending | Select-Object -First 10
            foreach ($group in $processGroups) {
                try {
                    $proc = Get-Process -Id $group.Name -ErrorAction SilentlyContinue
                    if ($proc) {
                        $netProcesses += [PSCustomObject]@{
                            PID          = $group.Name
                            ProcessName  = $proc.Name
                            Connections  = $group.Count
                            CPU_s        = [math]::Round($proc.CPU, 1)
                            Memory_MB    = [math]::Round($proc.WorkingSet / 1MB, 1)
                        }
                    }
                } catch { }
            }
        }
    } catch {
        Write-Log "Could not collect network usage data: $_" "WARN"
    }
    return $netProcesses
}

function Get-NetworkIssueAnalysis {
    param(
        [array]$Adapters,
        [PSCustomObject]$PublicIP,
        [array]$PingResults,
        [PSCustomObject]$SpeedTest
    )
    Write-Log "Analyzing network issues..." "INFO"
    $issues = @()

    # Check WiFi signal
    foreach ($a in $Adapters) {
        if ($a.ConnectionType -eq "WiFi" -and $a.LinkSpeed_Mbps -lt 54) {
            $issues += [PSCustomObject]@{ Severity = "MEDIUM"; Issue = "Weak WiFi Signal"; Detail = "Adapter '$($a.Name)' link speed $($a.LinkSpeed_Mbps) Mbps suggests weak WiFi." }
        }
    }

    # Check ping results
    foreach ($p in $PingResults) {
        if ($p.Status -eq "Unreachable") {
            $issues += [PSCustomObject]@{ Severity = "HIGH"; Issue = "DNS/Internet Unreachable"; Detail = "Cannot reach $($p.Target). Check internet connectivity." }
        } elseif ($p.AvgLatency -gt 200) {
            $issues += [PSCustomObject]@{ Severity = "HIGH"; Issue = "High Latency"; Detail = "Latency to $($p.Target) is $($p.AvgLatency) ms. Expected < 200 ms." }
        }
        if ($p.PacketLoss -gt 5) {
            $issues += [PSCustomObject]@{ Severity = "HIGH"; Issue = "Packet Loss Detected"; Detail = "$($p.PacketLoss)% packet loss to $($p.Target). Network may be unstable." }
        }
    }

    # Check speed
    if ($SpeedTest.Status -eq "Completed" -and $SpeedTest.DownloadMbps -lt 10) {
        $issues += [PSCustomObject]@{ Severity = "MEDIUM"; Issue = "ISP Slow Speed"; Detail = "Download speed $($SpeedTest.DownloadMbps) Mbps is below 10 Mbps threshold." }
    }

    # Check DNS
    try {
        $dnsStart  = Get-Date
        $null = Resolve-DnsName -Name "google.com" -Type A -ErrorAction Stop -DnsOnly
        $dnsTime   = ((Get-Date) - $dnsStart).TotalMilliseconds
        if ($dnsTime -gt 500) {
            $issues += [PSCustomObject]@{ Severity = "MEDIUM"; Issue = "DNS Slow Response"; Detail = "DNS resolution took $([math]::Round($dnsTime,0)) ms. Consider using faster DNS (8.8.8.8 or 1.1.1.1)." }
        }
    } catch {
        $issues += [PSCustomObject]@{ Severity = "HIGH"; Issue = "DNS Resolution Failed"; Detail = "Could not resolve google.com. DNS may be misconfigured." }
    }

    if ($issues.Count -eq 0) {
        $issues += [PSCustomObject]@{ Severity = "OK"; Issue = "No Issues Detected"; Detail = "Network appears healthy." }
    }
    return $issues
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT GENERATION
# ─────────────────────────────────────────────────────────────────────────────
function New-HtmlReport {
    param(
        [array]$Adapters,
        [PSCustomObject]$PublicIP,
        [array]$PingResults,
        [PSCustomObject]$SpeedTest,
        [array]$NetProcesses,
        [array]$NetIssues
    )

    $reportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $adapterRows = ""
    foreach ($a in $Adapters) {
        $typeColor = if ($a.ConnectionType -eq "WiFi") { "#0d6efd" } else { "#28a745" }
        $adapterRows += "<tr><td>$($a.Name)</td><td><span style='background:$typeColor;color:#fff;padding:2px 8px;border-radius:10px;font-size:0.8em;'>$($a.ConnectionType)</span></td><td>$($a.MACAddress)</td><td>$($a.IPAddress)</td><td>$($a.DefaultGateway)</td><td>$($a.DNSServers)</td><td>$($a.LinkSpeed_Mbps) Mbps</td></tr>"
    }

    $pingRows = ""
    foreach ($p in $PingResults) {
        $statusColor = switch ($p.Status) { "OK" { "#28a745" } "Degraded" { "#ffc107" } "Unreachable" { "#dc3545" } default { "#6c757d" } }
        $lossColor   = if ($p.PacketLoss -gt 5) { "#dc3545" } elseif ($p.PacketLoss -gt 0) { "#ffc107" } else { "#28a745" }
        $pingRows += "<tr><td>$($p.Target)</td><td>$($p.AvgLatency) ms</td><td>$($p.MinLatency) ms</td><td>$($p.MaxLatency) ms</td><td style='color:$lossColor;font-weight:bold;'>$($p.PacketLoss)%</td><td><span style='background:$statusColor;color:#fff;padding:2px 8px;border-radius:10px;font-size:0.8em;'>$($p.Status)</span></td></tr>"
    }

    $procRows = ""
    foreach ($pr in $NetProcesses) {
        $procRows += "<tr><td>$($pr.PID)</td><td>$($pr.ProcessName)</td><td>$($pr.Connections)</td><td>$($pr.CPU_s) s</td><td>$($pr.Memory_MB) MB</td></tr>"
    }

    $issueRows = ""
    foreach ($i in $NetIssues) {
        $sevColor = switch ($i.Severity) { "HIGH" { "#dc3545" } "MEDIUM" { "#ffc107" } "LOW" { "#17a2b8" } "OK" { "#28a745" } default { "#6c757d" } }
        $issueRows += "<tr><td><span style='background:$sevColor;color:#fff;padding:2px 10px;border-radius:12px;font-size:0.8em;'>$($i.Severity)</span></td><td>$($i.Issue)</td><td>$($i.Detail)</td></tr>"
    }

    $avgLatency = if ($PingResults.Count -gt 0) {
        [math]::Round(($PingResults | Where-Object { $_.AvgLatency -lt 9999 } | Measure-Object -Property AvgLatency -Average).Average, 1)
    } else { 0 }
    $maxLoss = if ($PingResults.Count -gt 0) { ($PingResults | Measure-Object -Property PacketLoss -Maximum).Maximum } else { 0 }
    $latColor = if ($avgLatency -gt 200) { "#dc3545" } elseif ($avgLatency -gt 100) { "#ffc107" } else { "#28a745" }
    $lossColor2 = if ($maxLoss -gt 5) { "#dc3545" } elseif ($maxLoss -gt 0) { "#ffc107" } else { "#28a745" }
    $speedColor = if ($SpeedTest.DownloadMbps -lt 10) { "#ffc107" } else { "#28a745" }

    $pingLabels = ($PingResults | ForEach-Object { "'$($_.Target)'" }) -join ","
    $pingData   = ($PingResults | ForEach-Object { if ($_.AvgLatency -lt 9999) { $_.AvgLatency } else { 0 } }) -join ","
    $pingLoss   = ($PingResults | ForEach-Object { $_.PacketLoss }) -join ","

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>InfraEye - Network Diagnostics Report</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <style>
        :root { --bg:#f4f6f8;--card-bg:#ffffff;--text:#212529;--text-muted:#6c757d;--border:#dee2e6;--header-bg:#1a1a2e;--accent:#0d6efd; }
        body.dark-mode { --bg:#1e1e2f;--card-bg:#2b2b3c;--text:#e9ecef;--text-muted:#adb5bd;--border:#495057;--header-bg:#0d0d1a; }
        * { box-sizing:border-box;margin:0;padding:0; }
        body { font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:var(--bg);color:var(--text);transition:background 0.3s,color 0.3s; }
        header { background:var(--header-bg);color:#fff;padding:20px 40px;display:flex;align-items:center;justify-content:space-between; }
        header h1 { font-size:1.8em;font-weight:700; }
        header p { font-size:0.95em;opacity:0.8; }
        .toggle-btn { background:var(--accent);color:#fff;border:none;padding:8px 18px;border-radius:20px;cursor:pointer;font-size:0.9em; }
        .container { max-width:1400px;margin:0 auto;padding:30px 20px; }
        .summary-grid { display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:20px;margin-bottom:30px; }
        .card { background:var(--card-bg);border-radius:12px;padding:20px;box-shadow:0 2px 8px rgba(0,0,0,0.08);border:1px solid var(--border); }
        .card h3 { font-size:0.9em;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px; }
        .card .value { font-size:2em;font-weight:700; }
        .card .sub { font-size:0.85em;color:var(--text-muted);margin-top:4px; }
        .charts-grid { display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:20px;margin-bottom:30px; }
        .chart-card { background:var(--card-bg);border-radius:12px;padding:20px;box-shadow:0 2px 8px rgba(0,0,0,0.08);border:1px solid var(--border); }
        .chart-card h2 { font-size:1em;margin-bottom:15px; }
        .section { background:var(--card-bg);border-radius:12px;padding:20px;margin-bottom:20px;box-shadow:0 2px 8px rgba(0,0,0,0.08);border:1px solid var(--border); }
        .section h2 { font-size:1.1em;font-weight:600;margin-bottom:15px;padding-bottom:8px;border-bottom:2px solid var(--accent); }
        table { width:100%;border-collapse:collapse;font-size:0.9em; }
        th { background:var(--accent);color:#fff;padding:10px 14px;text-align:left;font-weight:600; }
        td { padding:9px 14px;border-bottom:1px solid var(--border); }
        tr:nth-child(even) td { background:rgba(0,0,0,0.03); }
        body.dark-mode tr:nth-child(even) td { background:rgba(255,255,255,0.04); }
        tr:hover td { background:rgba(13,110,253,0.07); }
        .pubip-grid { display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px; }
        .info-item { display:flex;flex-direction:column; }
        .info-label { font-size:0.8em;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.5px; }
        .info-value { font-size:0.95em;font-weight:600;margin-top:2px; }
        footer { text-align:center;padding:30px 20px;color:var(--text-muted);font-size:0.85em;border-top:1px solid var(--border);margin-top:20px; }
        footer a { color:var(--accent);text-decoration:none; }
    </style>
</head>
<body>
<header>
    <div>
        <h1>&#x1F310; InfraEye &mdash; Network Diagnostics Report</h1>
        <p>Generated: $reportDate &nbsp;|&nbsp; Host: $($env:COMPUTERNAME)</p>
    </div>
    <button class="toggle-btn" onclick="toggleDarkMode()">&#9790; Dark Mode</button>
</header>
<div class="container">
    <!-- Summary Cards -->
    <div class="summary-grid">
        <div class="card">
            <h3>Avg Latency</h3>
            <div class="value" style="color:$latColor;">$avgLatency ms</div>
            <div class="sub">To DNS servers (8.8.8.8, 1.1.1.1)</div>
        </div>
        <div class="card">
            <h3>Max Packet Loss</h3>
            <div class="value" style="color:$lossColor2;">$maxLoss%</div>
            <div class="sub">Across all ping targets</div>
        </div>
        <div class="card">
            <h3>Download Speed</h3>
            <div class="value" style="color:$speedColor;">$($SpeedTest.DownloadMbps) Mbps</div>
            <div class="sub">$($SpeedTest.Status)</div>
        </div>
        <div class="card">
            <h3>Public IP</h3>
            <div class="value" style="font-size:1.2em;">$($PublicIP.PublicIP)</div>
            <div class="sub">$($PublicIP.ISP)</div>
        </div>
        <div class="card">
            <h3>Active Adapters</h3>
            <div class="value">$($Adapters.Count)</div>
            <div class="sub">IP-enabled adapters</div>
        </div>
        <div class="card">
            <h3>Issues Found</h3>
            <div class="value" style="color:$(if(@($NetIssues | Where-Object {$_.Severity -eq 'HIGH'}).Count -gt 0){'#dc3545'}elseif(@($NetIssues | Where-Object {$_.Severity -eq 'OK'}).Count -eq $NetIssues.Count){'#28a745'}else{'#ffc107'});">$(@($NetIssues | Where-Object {$_.Severity -ne 'OK'}).Count)</div>
            <div class="sub">Network problems detected</div>
        </div>
    </div>

    <!-- Charts -->
    <div class="charts-grid">
        <div class="chart-card">
            <h2>Ping Latency (ms)</h2>
            <canvas id="latencyChart" height="200"></canvas>
        </div>
        <div class="chart-card">
            <h2>Packet Loss (%)</h2>
            <canvas id="lossChart" height="200"></canvas>
        </div>
        <div class="chart-card">
            <h2>Issues by Severity</h2>
            <canvas id="issueChart" height="200"></canvas>
        </div>
    </div>

    <!-- Public IP Info -->
    <div class="section">
        <h2>&#x1F30D; Public IP Information</h2>
        <div class="pubip-grid">
            <div class="info-item"><span class="info-label">Public IP</span><span class="info-value">$($PublicIP.PublicIP)</span></div>
            <div class="info-item"><span class="info-label">ISP / Organization</span><span class="info-value">$($PublicIP.ISP)</span></div>
            <div class="info-item"><span class="info-label">City</span><span class="info-value">$($PublicIP.City)</span></div>
            <div class="info-item"><span class="info-label">Region</span><span class="info-value">$($PublicIP.Region)</span></div>
            <div class="info-item"><span class="info-label">Country</span><span class="info-value">$($PublicIP.Country)</span></div>
            <div class="info-item"><span class="info-label">Timezone</span><span class="info-value">$($PublicIP.Timezone)</span></div>
        </div>
    </div>

    <!-- Network Adapters -->
    <div class="section">
        <h2>&#x1F4F6; Network Adapters</h2>
        <div style="overflow-x:auto;">
        <table>
            <thead><tr><th>Adapter Name</th><th>Type</th><th>MAC Address</th><th>IP Address</th><th>Gateway</th><th>DNS Servers</th><th>Link Speed</th></tr></thead>
            <tbody>$adapterRows</tbody>
        </table>
        </div>
    </div>

    <!-- Ping Results -->
    <div class="section">
        <h2>&#x1F4E1; Ping Test Results</h2>
        <table>
            <thead><tr><th>Target</th><th>Avg Latency</th><th>Min Latency</th><th>Max Latency</th><th>Packet Loss</th><th>Status</th></tr></thead>
            <tbody>$pingRows</tbody>
        </table>
    </div>

    <!-- Bandwidth-heavy Processes -->
    <div class="section">
        <h2>&#x1F4CA; Top Network-Active Processes</h2>
        <table>
            <thead><tr><th>PID</th><th>Process Name</th><th>TCP Connections</th><th>CPU (s)</th><th>Memory (MB)</th></tr></thead>
            <tbody>$procRows</tbody>
        </table>
    </div>

    <!-- Network Root Cause Analysis -->
    <div class="section">
        <h2>&#x1F50D; Network Root Cause Analysis</h2>
        <table>
            <thead><tr><th>Severity</th><th>Issue</th><th>Details</th></tr></thead>
            <tbody>$issueRows</tbody>
        </table>
    </div>
</div>
<footer>
    <p>Report Version: 1.0 &nbsp;|&nbsp; Created by: <strong>Tushar Gudde</strong> &nbsp;|&nbsp;
    Website: <a href="https://tushargudde.tech" target="_blank">https://tushargudde.tech</a></p>
</footer>
<script>
    function toggleDarkMode() {
        document.body.classList.toggle('dark-mode');
        const btn = document.querySelector('.toggle-btn');
        btn.textContent = document.body.classList.contains('dark-mode') ? '\u2600 Light Mode' : '\u263A Dark Mode';
    }

    const pingLabels = [$pingLabels];
    const pingData   = [$pingData];
    const pingLoss   = [$pingLoss];

    new Chart(document.getElementById('latencyChart'), {
        type: 'bar',
        data: {
            labels: pingLabels,
            datasets: [{ label: 'Avg Latency (ms)', data: pingData,
                backgroundColor: pingData.map(v => v > 200 ? '#dc3545' : v > 100 ? '#ffc107' : '#28a745'),
                borderRadius: 6 }]
        },
        options: { responsive: true, plugins: { legend: { display: false } } }
    });

    new Chart(document.getElementById('lossChart'), {
        type: 'bar',
        data: {
            labels: pingLabels,
            datasets: [{ label: 'Packet Loss %', data: pingLoss,
                backgroundColor: pingLoss.map(v => v > 5 ? '#dc3545' : v > 0 ? '#ffc107' : '#28a745'),
                borderRadius: 6 }]
        },
        options: { responsive: true, scales: { y: { min: 0, max: 100 } }, plugins: { legend: { display: false } } }
    });

    const issueCounts = { HIGH: 0, MEDIUM: 0, LOW: 0, OK: 0 };
    $( ($NetIssues | ForEach-Object { "issueCounts['$($_.Severity)'] = (issueCounts['$($_.Severity)'] || 0) + 1;" }) -join "`n    " )
    new Chart(document.getElementById('issueChart'), {
        type: 'doughnut',
        data: {
            labels: Object.keys(issueCounts).filter(k => issueCounts[k] > 0),
            datasets: [{ data: Object.values(issueCounts).filter(v => v > 0),
                backgroundColor: ['#dc3545','#ffc107','#17a2b8','#28a745'], borderWidth: 2 }]
        },
        options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
    });
</script>
</body>
</html>
"@
    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN EXECUTION
# ─────────────────────────────────────────────────────────────────────────────
try {
    Write-Log "=== InfraEye Network Diagnostics Started ===" "INFO"

    Install-RequiredModules

    $adapterInfo   = @(Get-NetworkAdapterInfo)
    $publicIPInfo  = Get-PublicIPInfo
    $pingResults   = @(Invoke-PingTest -Targets @("8.8.8.8","1.1.1.1","8.8.4.4"))
    $speedTest     = Invoke-SpeedTest
    $netProcesses  = @(Get-NetworkUsage)
    $netIssues     = @(Get-NetworkIssueAnalysis `
                        -Adapters    $adapterInfo `
                        -PublicIP    $publicIPInfo `
                        -PingResults $pingResults `
                        -SpeedTest   $speedTest)

    Write-Log "Generating HTML report..." "INFO"
    $htmlContent = New-HtmlReport `
        -Adapters     $adapterInfo `
        -PublicIP     $publicIPInfo `
        -PingResults  $pingResults `
        -SpeedTest    $speedTest `
        -NetProcesses $netProcesses `
        -NetIssues    $netIssues

    $htmlContent | Out-File -FilePath $ReportFile -Encoding UTF8 -Force
    Write-Log "Report saved: $ReportFile" "SUCCESS"

    # Trigger alert if needed
    $alertScript = Join-Path $PSScriptRoot "NetworkDiagnostics_Alert.ps1"
    if (Test-Path $alertScript) {
        $highIssues = @($netIssues | Where-Object { $_.Severity -eq "HIGH" })
        if ($highIssues.Count -gt 0) {
            Write-Log "Critical network issues found. Triggering alert..." "WARN"
            $avgLat  = if ($pingResults.Count -gt 0) { ($pingResults | Measure-Object -Property AvgLatency -Average).Average } else { 0 }
            $maxLoss = if ($pingResults.Count -gt 0) { ($pingResults | Measure-Object -Property PacketLoss -Maximum).Maximum } else { 0 }
            & $alertScript `
                -ReportFile    $ReportFile `
                -AvgLatency    $avgLat `
                -PacketLoss    $maxLoss `
                -DownloadSpeed $speedTest.DownloadMbps
        }
    }

    Write-Log "=== InfraEye Network Diagnostics Completed ===" "SUCCESS"
    Write-Host "`nReport generated: $ReportFile" -ForegroundColor Green
} catch {
    Write-Log "FATAL ERROR: $_" "ERROR"
    Write-Error $_
    exit 1
}
