#Requires -Version 5.1
<#
.SYNOPSIS
    InfraEye - Device Health Diagnostics Module
.DESCRIPTION
    Collects comprehensive device health data including CPU, RAM, Disk, GPU,
    Battery, and startup programs. Generates a professional HTML dashboard report.
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

$Timestamp    = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile      = Join-Path $LogDir "DeviceHealth_$Timestamp.log"
$ReportFile   = Join-Path $ReportDir "DeviceHealth_$Timestamp.html"
$ConfigFile   = Join-Path $PSScriptRoot "..\config.json"

# ─────────────────────────────────────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────────────────────────────────────
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS")]
        [string]$Level = "INFO"
    )
    $Entry = "{0} | DeviceHealth | {1} | {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
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
function Get-SystemType {
    Write-Log "Detecting system type..." "INFO"
    try {
        $enclosure = Get-CimInstance -ClassName Win32_SystemEnclosure
        $chassisTypes = $enclosure.ChassisTypes
        $type = switch ($chassisTypes[0]) {
            { $_ -in @(8,9,10,11,12,14,18,21) } { "Laptop" }
            { $_ -in @(3,4,5,6,7,15,16) }        { "Desktop" }
            { $_ -in @(1,2,13,17,19,20,22,23) }   { "Workstation" }
            default                                { "Unknown" }
        }
        Write-Log "System type detected: $type" "INFO"
        return $type
    } catch {
        Write-Log "Could not detect system type: $_" "WARN"
        return "Unknown"
    }
}

function Get-SystemInfo {
    Write-Log "Collecting system information..." "INFO"
    try {
        $cs  = Get-CimInstance -ClassName Win32_ComputerSystem
        $os  = Get-CimInstance -ClassName Win32_OperatingSystem
        $bio = Get-CimInstance -ClassName Win32_BIOS

        $uptimeSpan = (Get-Date) - $os.LastBootUpTime
        $uptime = "{0}d {1}h {2}m" -f $uptimeSpan.Days, $uptimeSpan.Hours, $uptimeSpan.Minutes

        return [PSCustomObject]@{
            ComputerName = $cs.Name
            Manufacturer = $cs.Manufacturer
            Model        = $cs.Model
            SerialNumber = $bio.SerialNumber
            OSVersion    = $os.Caption + " " + $os.Version
            Uptime       = $uptime
        }
    } catch {
        Write-Log "Error collecting system info: $_" "ERROR"
        throw
    }
}

function Get-CPUInfo {
    Write-Log "Collecting CPU data..." "INFO"
    try {
        $cpu  = Get-CimInstance -ClassName Win32_Processor
        $load = (Get-CimInstance -ClassName Win32_Processor | Measure-Object -Property LoadPercentage -Average).Average

        return [PSCustomObject]@{
            Model      = $cpu[0].Name.Trim()
            UsagePercent = [math]::Round($load, 1)
            TotalCores = ($cpu | Measure-Object -Property NumberOfCores -Sum).Sum
            TotalLogical = ($cpu | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
        }
    } catch {
        Write-Log "Error collecting CPU info: $_" "ERROR"
        throw
    }
}

function Get-RAMInfo {
    Write-Log "Collecting RAM data..." "INFO"
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        $totalGB = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
        $freeGB  = [math]::Round($os.FreePhysicalMemory   / 1MB, 2)
        $usedGB  = [math]::Round($totalGB - $freeGB, 2)
        $usedPct = if ($totalGB -gt 0) { [math]::Round(($usedGB / $totalGB) * 100, 1) } else { 0 }

        $recommendation = if ($totalGB -lt 8) { "UPGRADE RECOMMENDED: RAM < 8 GB detected. Upgrade to at least 16 GB for optimal performance." } else { "RAM is sufficient." }

        return [PSCustomObject]@{
            TotalGB        = $totalGB
            UsedGB         = $usedGB
            AvailableGB    = $freeGB
            UsedPercent    = $usedPct
            Recommendation = $recommendation
        }
    } catch {
        Write-Log "Error collecting RAM info: $_" "ERROR"
        throw
    }
}

function Get-DiskInfo {
    Write-Log "Collecting disk data..." "INFO"
    $disks = @()
    try {
        $logicalDisks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"
        foreach ($disk in $logicalDisks) {
            $totalGB = [math]::Round($disk.Size / 1GB, 2)
            $freeGB  = [math]::Round($disk.FreeSpace / 1GB, 2)
            $usedGB  = [math]::Round($totalGB - $freeGB, 2)
            $freePct = if ($totalGB -gt 0) { [math]::Round(($freeGB / $totalGB) * 100, 1) } else { 0 }

            # Detect SSD or HDD via Win32_DiskDrive
            $diskType = "Unknown"
            try {
                $partition = Get-CimInstance -Query "ASSOCIATORS OF {Win32_LogicalDisk.DeviceID='$($disk.DeviceID)'} WHERE AssocClass=Win32_LogicalDiskToPartition"
                if ($partition) {
                    $drive = Get-CimInstance -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='$($partition[0].DeviceID)'} WHERE AssocClass=Win32_DiskDriveToDiskPartition"
                    if ($drive) {
                        $mediaType = $drive[0].MediaType
                        if ($mediaType -match "SSD|Solid") { $diskType = "SSD" }
                        elseif ($mediaType -match "Fixed|HDD") { $diskType = "HDD" }
                        else {
                            # Try model name heuristic
                            $diskType = if ($drive[0].Model -match "SSD|NVMe|M\.2") { "SSD" } else { "HDD" }
                        }
                    }
                }
            } catch {
                $diskType = "Unknown"
            }

            $recommendation = if ($diskType -eq "HDD") {
                "UPGRADE RECOMMENDED: HDD detected on $($disk.DeviceID). Consider upgrading to SSD for better performance."
            } elseif ($freePct -lt 10) {
                "WARNING: Disk $($disk.DeviceID) is almost full ($freePct% free). Consider cleaning up or expanding storage."
            } else {
                "OK"
            }

            $disks += [PSCustomObject]@{
                Drive          = $disk.DeviceID
                TotalGB        = $totalGB
                UsedGB         = $usedGB
                FreeGB         = $freeGB
                FreePercent    = $freePct
                DiskType       = $diskType
                Recommendation = $recommendation
            }
        }
    } catch {
        Write-Log "Error collecting disk info: $_" "ERROR"
        throw
    }
    return $disks
}

function Get-GPUInfo {
    Write-Log "Collecting GPU data..." "INFO"
    $gpus = @()
    try {
        $videoControllers = Get-CimInstance -ClassName Win32_VideoController
        foreach ($gpu in $videoControllers) {
            $gpuType = if ($gpu.Name -match "Intel|AMD Radeon(.*?)(Graphics|HD|UHD|Vega)|Vega") {
                "Integrated"
            } elseif ($gpu.Name -match "NVIDIA|AMD Radeon RX|Radeon Pro|FirePro|Quadro") {
                "Dedicated"
            } else {
                "Unknown"
            }
            $gpus += [PSCustomObject]@{
                Model   = $gpu.Name
                GPUType = $gpuType
                VRAM_MB = [math]::Round($gpu.AdapterRAM / 1MB, 0)
            }
        }
    } catch {
        Write-Log "Error collecting GPU info: $_" "WARN"
    }
    return $gpus
}

function Get-BatteryInfo {
    Write-Log "Collecting battery data..." "INFO"
    try {
        $battery = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
        if ($battery) {
            $statusMap = @{
                1  = "Discharging"
                2  = "On AC Power"
                3  = "Fully Charged"
                4  = "Low"
                5  = "Critical"
                6  = "Charging"
                7  = "Charging and High"
                8  = "Charging and Low"
                9  = "Charging and Critical"
                10 = "Undefined"
                11 = "Partially Charged"
            }
            return [PSCustomObject]@{
                Status  = $statusMap[[int]$battery.BatteryStatus]
                ChargePercent = $battery.EstimatedChargeRemaining
                Present = $true
            }
        } else {
            return [PSCustomObject]@{
                Status        = "No Battery"
                ChargePercent = "N/A"
                Present       = $false
            }
        }
    } catch {
        Write-Log "Error collecting battery info: $_" "WARN"
        return [PSCustomObject]@{
            Status        = "Unknown"
            ChargePercent = "N/A"
            Present       = $false
        }
    }
}

function Get-TemperatureInfo {
    Write-Log "Collecting temperature data..." "INFO"
    $zones = @()
    try {
        $thermalZones = Get-CimInstance -Namespace "root/WMI" -ClassName MSAcpi_ThermalZoneTemperature -ErrorAction SilentlyContinue
        if ($thermalZones) {
            foreach ($zone in $thermalZones) {
                $tempCelsius = [math]::Round(($zone.CurrentTemperature / 10) - 273.15, 1)
                $zoneName = if ($zone.InstanceName -match "ThermalZone(\d+)") {
                    "Thermal Zone $($Matches[1])"
                } elseif ($zone.InstanceName -match "_TZ\.(\w+)") {
                    $Matches[1]
                } else {
                    # Extract the last non-numeric path component as a friendly name
                    ($zone.InstanceName -split '\\' | Where-Object { $_ -match '\D' } | Select-Object -Last 1) -replace '_\d+$', ''
                }
                $status = if ($tempCelsius -ge 90) { "Critical" }
                          elseif ($tempCelsius -ge 80) { "Hot" }
                          elseif ($tempCelsius -ge 65) { "Warm" }
                          else { "Normal" }
                $zones += [PSCustomObject]@{
                    ZoneName   = $zoneName
                    TempC      = $tempCelsius
                    Status     = $status
                }
            }
        }
    } catch {
        Write-Log "Could not read thermal zone temperatures: $_" "WARN"
    }

    # Try Win32_TemperatureProbe as a secondary source (rarely populated but worth trying)
    if ($zones.Count -eq 0) {
        try {
            $probes = Get-CimInstance -ClassName Win32_TemperatureProbe -ErrorAction SilentlyContinue
            if ($probes) {
                foreach ($p in $probes) {
                    # CurrentReading is in tenths of Kelvin; null or zero indicates no valid reading
                    if ($null -ne $p.CurrentReading -and $p.CurrentReading -ne 0) {
                        $tempCelsius = [math]::Round(($p.CurrentReading / 10) - 273.15, 1)
                        $status = if ($tempCelsius -ge 90) { "Critical" }
                                  elseif ($tempCelsius -ge 80) { "Hot" }
                                  elseif ($tempCelsius -ge 65) { "Warm" }
                                  else { "Normal" }
                        $zones += [PSCustomObject]@{
                            ZoneName = if ($p.Name) { $p.Name } else { "Temperature Probe" }
                            TempC    = $tempCelsius
                            Status   = $status
                        }
                    }
                }
            }
        } catch {
            Write-Log "Could not read temperature probes: $_" "WARN"
        }
    }

    if ($zones.Count -eq 0) {
        Write-Log "No temperature sensors accessible via WMI on this system." "WARN"
    }
    return $zones
}

function Get-StartupPrograms {
    Write-Log "Collecting startup programs..." "INFO"
    $startupItems = @()
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run"
    )
    foreach ($path in $regPaths) {
        try {
            if (Test-Path $path) {
                $items = Get-ItemProperty -Path $path
                $props = $items.PSObject.Properties | Where-Object { $_.Name -notmatch "^PS" }
                foreach ($prop in $props) {
                    $startupItems += [PSCustomObject]@{
                        Name    = $prop.Name
                        Command = $prop.Value
                        Source  = $path
                    }
                }
            }
        } catch {
            Write-Log "Could not read startup path $path`: $_" "WARN"
        }
    }
    # Startup folder
    try {
        $startupFolder = [System.Environment]::GetFolderPath("Startup")
        if (Test-Path $startupFolder) {
            $files = Get-ChildItem -Path $startupFolder -File
            foreach ($f in $files) {
                $startupItems += [PSCustomObject]@{
                    Name    = $f.Name
                    Command = $f.FullName
                    Source  = "Startup Folder"
                }
            }
        }
    } catch {
        Write-Log "Could not read startup folder: $_" "WARN"
    }
    return $startupItems
}

function Get-PerformanceAnalysis {
    param(
        [PSCustomObject]$CPU,
        [PSCustomObject]$RAM,
        [array]$Disks,
        [array]$StartupApps,
        [array]$Temperatures = @()
    )
    Write-Log "Performing performance analysis..." "INFO"
    $issues = @()

    if ($RAM.TotalGB -lt 8) {
        $issues += [PSCustomObject]@{ Issue = "Low RAM"; Severity = "HIGH"; Detail = "Only $($RAM.TotalGB) GB RAM installed. Minimum 8 GB recommended." }
    }
    if ($CPU.UsagePercent -gt 85) {
        $issues += [PSCustomObject]@{ Issue = "High CPU Usage"; Severity = "HIGH"; Detail = "CPU at $($CPU.UsagePercent)%. System may be overloaded." }
    }
    foreach ($d in $Disks) {
        if ($d.FreePercent -lt 10) {
            $issues += [PSCustomObject]@{ Issue = "Disk Bottleneck"; Severity = "HIGH"; Detail = "Drive $($d.Drive) has only $($d.FreePercent)% free space." }
        }
        if ($d.DiskType -eq "HDD") {
            $issues += [PSCustomObject]@{ Issue = "HDD Detected"; Severity = "MEDIUM"; Detail = "Drive $($d.Drive) is an HDD. Upgrading to SSD will significantly improve performance." }
        }
    }
    $processCount = (Get-Process).Count
    if ($processCount -gt 150) {
        $issues += [PSCustomObject]@{ Issue = "Too Many Processes"; Severity = "MEDIUM"; Detail = "$processCount processes running. Consider terminating unused processes." }
    }
    if ($StartupApps.Count -gt 10) {
        $severity = if ($StartupApps.Count -gt 20) { "HIGH" } else { "MEDIUM" }
        $issues += [PSCustomObject]@{ Issue = "Too Many Startup Programs"; Severity = $severity; Detail = "$($StartupApps.Count) startup apps detected. Reduce to improve boot time." }
    }
    if ($RAM.UsedPercent -gt 85) {
        $issues += [PSCustomObject]@{ Issue = "High RAM Usage"; Severity = "HIGH"; Detail = "RAM usage at $($RAM.UsedPercent)%. System may be memory-constrained." }
    }

    foreach ($zone in $Temperatures) {
        if ($zone.TempC -ge 90) {
            $issues += [PSCustomObject]@{ Issue = "Critical Temperature"; Severity = "HIGH"; Detail = "$($zone.ZoneName) is at $($zone.TempC)°C. Immediate cooling action required." }
        } elseif ($zone.TempC -ge 80) {
            $issues += [PSCustomObject]@{ Issue = "High Temperature"; Severity = "HIGH"; Detail = "$($zone.ZoneName) is at $($zone.TempC)°C. Check cooling system." }
        } elseif ($zone.TempC -ge 65) {
            $issues += [PSCustomObject]@{ Issue = "Elevated Temperature"; Severity = "MEDIUM"; Detail = "$($zone.ZoneName) is running warm at $($zone.TempC)°C. Monitor closely." }
        }
    }

    if ($issues.Count -eq 0) {
        $issues += [PSCustomObject]@{ Issue = "No Issues Detected"; Severity = "OK"; Detail = "System appears healthy." }
    }
    return $issues
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT GENERATION
# ─────────────────────────────────────────────────────────────────────────────
function New-HtmlReport {
    param(
        [string]$SystemType,
        [PSCustomObject]$SystemInfo,
        [PSCustomObject]$CPU,
        [PSCustomObject]$RAM,
        [array]$Disks,
        [array]$GPUs,
        [PSCustomObject]$Battery,
        [array]$StartupApps,
        [array]$PerfIssues,
        [array]$Temperatures = @()
    )

    $reportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $diskRowsHtml = ""
    foreach ($d in $Disks) {
        $badgeColor = if ($d.DiskType -eq "SSD") { "#28a745" } else { "#ffc107" }
        $freeColor  = if ($d.FreePercent -lt 10) { "#dc3545" } elseif ($d.FreePercent -lt 20) { "#ffc107" } else { "#28a745" }
        $diskRowsHtml += @"
            <tr>
                <td>$($d.Drive)</td>
                <td><span style='background:$badgeColor;color:#fff;padding:2px 8px;border-radius:12px;font-size:0.8em;'>$($d.DiskType)</span></td>
                <td>$($d.TotalGB) GB</td>
                <td>$($d.UsedGB) GB</td>
                <td><span style='color:$freeColor;font-weight:bold;'>$($d.FreeGB) GB ($($d.FreePercent)%)</span></td>
                <td>$($d.Recommendation)</td>
            </tr>
"@
    }

    $gpuRowsHtml = ""
    foreach ($g in $GPUs) {
        $gpuTypeColor = if ($g.GPUType -eq "Dedicated") { "#007bff" } else { "#6c757d" }
        $gpuRowsHtml += @"
            <tr>
                <td>$($g.Model)</td>
                <td><span style='background:$gpuTypeColor;color:#fff;padding:2px 8px;border-radius:12px;font-size:0.8em;'>$($g.GPUType)</span></td>
                <td>$($g.VRAM_MB) MB</td>
            </tr>
"@
    }

    $startupRowsHtml = ""
    foreach ($s in $StartupApps) {
        $startupRowsHtml += @"
            <tr>
                <td>$([System.Net.WebUtility]::HtmlEncode($s.Name))</td>
                <td style='font-size:0.8em;word-break:break-all;'>$([System.Net.WebUtility]::HtmlEncode($s.Command))</td>
                <td>$($s.Source)</td>
            </tr>
"@
    }

    $perfRowsHtml = ""
    foreach ($p in $PerfIssues) {
        $sevColor = switch ($p.Severity) {
            "HIGH"   { "#dc3545" }
            "MEDIUM" { "#ffc107" }
            "LOW"    { "#17a2b8" }
            "OK"     { "#28a745" }
            default  { "#6c757d" }
        }
        $perfRowsHtml += @"
            <tr>
                <td><span style='background:$sevColor;color:#fff;padding:2px 10px;border-radius:12px;font-size:0.8em;'>$($p.Severity)</span></td>
                <td>$($p.Issue)</td>
                <td>$($p.Detail)</td>
            </tr>
"@
    }

    $cpuColor  = if ($CPU.UsagePercent -gt 90) { "#dc3545" } elseif ($CPU.UsagePercent -gt 70) { "#ffc107" } else { "#28a745" }
    $ramColor  = if ($RAM.UsedPercent  -gt 90) { "#dc3545" } elseif ($RAM.UsedPercent  -gt 70) { "#ffc107" } else { "#28a745" }
    $diskFreeMin = if ($Disks.Count -gt 0) { ($Disks | Measure-Object -Property FreePercent -Minimum).Minimum } else { 100 }
    $diskColor = if ($diskFreeMin -lt 10) { "#dc3545" } elseif ($diskFreeMin -lt 20) { "#ffc107" } else { "#28a745" }

    # Temperature summary values
    $maxTempC       = if ($Temperatures.Count -gt 0) { ($Temperatures | Measure-Object -Property TempC -Maximum).Maximum } else { $null }
    $tempCardValue  = if ($null -ne $maxTempC) { "$maxTempC °C" } else { "N/A" }
    $tempCardColor  = if ($null -eq $maxTempC) { "#6c757d" }
                      elseif ($maxTempC -ge 90) { "#a71d2a" }
                      elseif ($maxTempC -ge 80) { "#dc3545" }
                      elseif ($maxTempC -ge 65) { "#ffc107" }
                      else { "#28a745" }
    $tempCardSub    = if ($null -eq $maxTempC) { "No sensor data available" }
                      elseif ($maxTempC -ge 90) { "CRITICAL — check cooling!" }
                      elseif ($maxTempC -ge 80) { "Hot — check cooling system" }
                      elseif ($maxTempC -ge 65) { "Warm — monitor closely" }
                      else { "Normal operating range" }

    # Temperature table rows
    # Bar width: scale so 120 °C maps to 100% (capped at 100% to handle any outliers)
    $tempBarMaxScale = 120
    $tempRowsHtml = ""
    if ($Temperatures.Count -gt 0) {
        foreach ($t in ($Temperatures | Sort-Object TempC -Descending)) {
            $tColor = if ($t.TempC -ge 90) { "#a71d2a" }
                      elseif ($t.TempC -ge 80) { "#dc3545" }
                      elseif ($t.TempC -ge 65) { "#ffc107" }
                      else { "#28a745" }
            $tBarWidth = [math]::Min([math]::Round(($t.TempC / $tempBarMaxScale) * 100, 0), 100)
            $tempRowsHtml += @"
            <tr>
                <td>$([System.Net.WebUtility]::HtmlEncode($t.ZoneName))</td>
                <td style='font-weight:700;color:$tColor;'>$($t.TempC) °C</td>
                <td>
                    <div style='background:#e9ecef;border-radius:6px;height:12px;width:100%;'>
                        <div style='background:$tColor;border-radius:6px;height:12px;width:$($tBarWidth)%;'></div>
                    </div>
                </td>
                <td><span style='background:$tColor;color:#fff;padding:2px 10px;border-radius:12px;font-size:0.8em;'>$($t.Status)</span></td>
            </tr>
"@
        }
    } else {
        $tempRowsHtml = "<tr><td colspan='4' style='text-align:center;color:#6c757d;'>Temperature sensor data not available on this system.</td></tr>"
    }

    # Temperature chart data
    $tempChartLabels = if ($Temperatures.Count -gt 0) {
        ($Temperatures | Sort-Object TempC -Descending | ForEach-Object { "'$([System.Net.WebUtility]::HtmlEncode($_.ZoneName))'" }) -join ','
    } else { "" }
    $tempChartValues = if ($Temperatures.Count -gt 0) {
        ($Temperatures | Sort-Object TempC -Descending | ForEach-Object { $_.TempC }) -join ','
    } else { "" }
    $tempChartColors = if ($Temperatures.Count -gt 0) {
        ($Temperatures | Sort-Object TempC -Descending | ForEach-Object {
            if ($_.TempC -ge 90) { "'#a71d2a'" }
            elseif ($_.TempC -ge 80) { "'#dc3545'" }
            elseif ($_.TempC -ge 65) { "'#ffc107'" }
            else { "'#28a745'" }
        }) -join ','
    } else { "" }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>InfraEye - Device Health Report</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <style>
        :root {
            --bg: #f4f6f8;
            --card-bg: #ffffff;
            --text: #212529;
            --text-muted: #6c757d;
            --border: #dee2e6;
            --header-bg: #1a1a2e;
            --header-text: #ffffff;
            --accent: #0d6efd;
        }
        body.dark-mode {
            --bg: #1e1e2f;
            --card-bg: #2b2b3c;
            --text: #e9ecef;
            --text-muted: #adb5bd;
            --border: #495057;
            --header-bg: #0d0d1a;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: var(--bg); color: var(--text); transition: background 0.3s, color 0.3s; }
        header { background: var(--header-bg); color: var(--header-text); padding: 20px 40px; display: flex; align-items: center; justify-content: space-between; }
        header h1 { font-size: 1.8em; font-weight: 700; }
        header p { font-size: 0.95em; opacity: 0.8; }
        .toggle-btn { background: var(--accent); color: #fff; border: none; padding: 8px 18px; border-radius: 20px; cursor: pointer; font-size: 0.9em; }
        .container { max-width: 1400px; margin: 0 auto; padding: 30px 20px; }
        .summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px; }
        .card { background: var(--card-bg); border-radius: 12px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); border: 1px solid var(--border); transition: background 0.3s; }
        .card h3 { font-size: 0.9em; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px; }
        .card .value { font-size: 2em; font-weight: 700; }
        .card .sub { font-size: 0.85em; color: var(--text-muted); margin-top: 4px; }
        .charts-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; margin-bottom: 30px; }
        .chart-card { background: var(--card-bg); border-radius: 12px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); border: 1px solid var(--border); }
        .chart-card h2 { font-size: 1em; margin-bottom: 15px; }
        .section { background: var(--card-bg); border-radius: 12px; padding: 20px; margin-bottom: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); border: 1px solid var(--border); }
        .section h2 { font-size: 1.1em; font-weight: 600; margin-bottom: 15px; padding-bottom: 8px; border-bottom: 2px solid var(--accent); }
        table { width: 100%; border-collapse: collapse; font-size: 0.9em; }
        th { background: var(--accent); color: #fff; padding: 10px 14px; text-align: left; font-weight: 600; }
        td { padding: 9px 14px; border-bottom: 1px solid var(--border); }
        tr:nth-child(even) td { background: rgba(0,0,0,0.03); }
        body.dark-mode tr:nth-child(even) td { background: rgba(255,255,255,0.04); }
        tr:hover td { background: rgba(13,110,253,0.07); }
        .sysinfo-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 12px; }
        .info-item { display: flex; flex-direction: column; }
        .info-label { font-size: 0.8em; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; }
        .info-value { font-size: 0.95em; font-weight: 600; margin-top: 2px; }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 0.8em; font-weight: 600; }
        footer { text-align: center; padding: 30px 20px; color: var(--text-muted); font-size: 0.85em; border-top: 1px solid var(--border); margin-top: 20px; }
        footer a { color: var(--accent); text-decoration: none; }
        @media (max-width: 600px) { header { flex-direction: column; gap: 12px; } .summary-grid { grid-template-columns: 1fr 1fr; } }
    </style>
</head>
<body>
<header>
    <div>
        <h1>&#x1F4BB; InfraEye &mdash; Device Health Report</h1>
        <p>Generated: $reportDate &nbsp;|&nbsp; Computer: $($SystemInfo.ComputerName) &nbsp;|&nbsp; Type: $SystemType</p>
    </div>
    <button class="toggle-btn" onclick="toggleDarkMode()">&#9790; Dark Mode</button>
</header>
<div class="container">
    <!-- Summary Cards -->
    <div class="summary-grid">
        <div class="card">
            <h3>CPU Usage</h3>
            <div class="value" style="color:$cpuColor;">$($CPU.UsagePercent)%</div>
            <div class="sub">$($CPU.Model.Substring(0, [Math]::Min(30, $CPU.Model.Length)))</div>
        </div>
        <div class="card">
            <h3>RAM Usage</h3>
            <div class="value" style="color:$ramColor;">$($RAM.UsedPercent)%</div>
            <div class="sub">$($RAM.UsedGB) GB / $($RAM.TotalGB) GB</div>
        </div>
        <div class="card">
            <h3>Disk Free (Min)</h3>
            <div class="value" style="color:$diskColor;">$diskFreeMin%</div>
            <div class="sub">Lowest free space across drives</div>
        </div>
        <div class="card">
            <h3>CPU Cores</h3>
            <div class="value">$($CPU.TotalCores)</div>
            <div class="sub">$($CPU.TotalLogical) Logical Processors</div>
        </div>
        <div class="card">
            <h3>Startup Apps</h3>
            <div class="value" style="color:$(if($StartupApps.Count -gt 20){'#dc3545'}elseif($StartupApps.Count -gt 10){'#ffc107'}else{'#28a745'});">$($StartupApps.Count)</div>
            <div class="sub">$(if($StartupApps.Count -gt 10){'Warning: Many startup items'}else{'OK'})</div>
        </div>
        <div class="card">
            <h3>Battery</h3>
            <div class="value">$(if($Battery.Present){$Battery.ChargePercent.ToString() + '%'}else{'N/A'})</div>
            <div class="sub">$($Battery.Status)</div>
        </div>
        <div class="card">
            <h3>Max Temperature</h3>
            <div class="value" style="color:$tempCardColor;">$tempCardValue</div>
            <div class="sub">$tempCardSub</div>
        </div>
    </div>

    <!-- Charts -->
    <div class="charts-grid">
        <div class="chart-card">
            <h2>CPU &amp; RAM Usage</h2>
            <canvas id="cpuRamChart" height="200"></canvas>
        </div>
        <div class="chart-card">
            <h2>Disk Usage</h2>
            <canvas id="diskChart" height="200"></canvas>
        </div>
        <div class="chart-card">
            <h2>Performance Issues</h2>
            <canvas id="issuesChart" height="200"></canvas>
        </div>
        <div class="chart-card">
            <h2>&#x1F321; Temperatures (°C)</h2>
            <canvas id="tempChart" height="200"></canvas>
        </div>
    </div>

    <!-- System Info -->
    <div class="section">
        <h2>&#x1F5A5; System Information</h2>
        <div class="sysinfo-grid">
            <div class="info-item"><span class="info-label">Computer Name</span><span class="info-value">$($SystemInfo.ComputerName)</span></div>
            <div class="info-item"><span class="info-label">Manufacturer</span><span class="info-value">$($SystemInfo.Manufacturer)</span></div>
            <div class="info-item"><span class="info-label">Model</span><span class="info-value">$($SystemInfo.Model)</span></div>
            <div class="info-item"><span class="info-label">Serial Number</span><span class="info-value">$($SystemInfo.SerialNumber)</span></div>
            <div class="info-item"><span class="info-label">OS Version</span><span class="info-value">$($SystemInfo.OSVersion)</span></div>
            <div class="info-item"><span class="info-label">System Uptime</span><span class="info-value">$($SystemInfo.Uptime)</span></div>
        </div>
    </div>

    <!-- Disk Table -->
    <div class="section">
        <h2>&#x1F4BE; Disk Storage</h2>
        <table>
            <thead><tr><th>Drive</th><th>Type</th><th>Total</th><th>Used</th><th>Free</th><th>Recommendation</th></tr></thead>
            <tbody>$diskRowsHtml</tbody>
        </table>
    </div>

    <!-- GPU Table -->
    <div class="section">
        <h2>&#x1F3AE; GPU Information</h2>
        <table>
            <thead><tr><th>Model</th><th>Type</th><th>VRAM</th></tr></thead>
            <tbody>$gpuRowsHtml</tbody>
        </table>
    </div>

    <!-- Startup Programs -->
    <div class="section">
        <h2>&#x1F680; Startup Programs ($($StartupApps.Count))</h2>
        $(if($StartupApps.Count -gt 10){'<p style="color:#ffc107;margin-bottom:12px;">&#x26A0; Warning: More than 10 startup programs detected. This may slow down boot time.</p>'}else{''})
        <table>
            <thead><tr><th>Name</th><th>Command</th><th>Source</th></tr></thead>
            <tbody>$startupRowsHtml</tbody>
        </table>
    </div>

    <!-- Performance Analysis -->
    <div class="section">
        <h2>&#x1F50D; Performance Root Cause Analysis</h2>
        <table>
            <thead><tr><th>Severity</th><th>Issue</th><th>Details</th></tr></thead>
            <tbody>$perfRowsHtml</tbody>
        </table>
    </div>

    <!-- Temperature -->
    <div class="section">
        <h2>&#x1F321; Device Temperature</h2>
        $(if($null -ne $maxTempC -and $maxTempC -ge 80){"<p style='color:#dc3545;margin-bottom:12px;'>&#x26A0; High temperature detected. Ensure adequate ventilation and check cooling system.</p>"}elseif($null -ne $maxTempC -and $maxTempC -ge 65){"<p style='color:#ffc107;margin-bottom:12px;'>&#x26A0; Elevated temperature detected. Monitor the system closely.</p>"}else{""})
        <table>
            <thead><tr><th>Zone / Sensor</th><th>Temperature</th><th>Heat Bar</th><th>Status</th></tr></thead>
            <tbody>$tempRowsHtml</tbody>
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

    const cpuVal  = $($CPU.UsagePercent);
    const ramVal  = $($RAM.UsedPercent);
    const diskLabels = [$(($Disks | ForEach-Object { "'$($_.Drive)'" }) -join ',')];
    const diskUsed   = [$(($Disks | ForEach-Object { $_.UsedGB }) -join ',')];
    const diskFree   = [$(($Disks | ForEach-Object { $_.FreeGB }) -join ',')];

    new Chart(document.getElementById('cpuRamChart'), {
        type: 'bar',
        data: {
            labels: ['CPU Usage', 'RAM Usage'],
            datasets: [{
                label: 'Usage %',
                data: [cpuVal, ramVal],
                backgroundColor: [
                    cpuVal > 90 ? '#dc3545' : cpuVal > 70 ? '#ffc107' : '#28a745',
                    ramVal > 90 ? '#dc3545' : ramVal > 70 ? '#ffc107' : '#28a745'
                ],
                borderRadius: 6
            }]
        },
        options: { responsive: true, scales: { y: { min: 0, max: 100 } }, plugins: { legend: { display: false } } }
    });

    new Chart(document.getElementById('diskChart'), {
        type: 'bar',
        data: {
            labels: diskLabels,
            datasets: [
                { label: 'Used (GB)', data: diskUsed, backgroundColor: '#0d6efd', borderRadius: 4 },
                { label: 'Free (GB)',  data: diskFree,  backgroundColor: '#28a745', borderRadius: 4 }
            ]
        },
        options: { responsive: true, plugins: { legend: { position: 'bottom' } }, scales: { x: { stacked: true }, y: { stacked: true } } }
    });

    const issuesBySeverity = { HIGH: 0, MEDIUM: 0, LOW: 0, OK: 0 };
    $( ($PerfIssues | ForEach-Object { "issuesBySeverity['$($_.Severity)'] = (issuesBySeverity['$($_.Severity)'] || 0) + 1;" }) -join "`n    " )
    new Chart(document.getElementById('issuesChart'), {
        type: 'doughnut',
        data: {
            labels: Object.keys(issuesBySeverity).filter(k => issuesBySeverity[k] > 0),
            datasets: [{
                data: Object.values(issuesBySeverity).filter(v => v > 0),
                backgroundColor: ['#dc3545','#ffc107','#17a2b8','#28a745'],
                borderWidth: 2
            }]
        },
        options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
    });

    const tempLabels = [$tempChartLabels];
    const tempValues = [$tempChartValues];
    const tempColors = [$tempChartColors];
    const tempChartMax = $tempBarMaxScale;
    if (tempLabels.length > 0) {
        new Chart(document.getElementById('tempChart'), {
            type: 'bar',
            data: {
                labels: tempLabels,
                datasets: [{
                    label: 'Temperature (°C)',
                    data: tempValues,
                    backgroundColor: tempColors,
                    borderRadius: 6
                }]
            },
            options: {
                responsive: true,
                plugins: { legend: { display: false } },
                scales: {
                    y: {
                        min: 0,
                        max: tempChartMax,
                        ticks: { callback: v => v + '°C' }
                    }
                }
            }
        });
    } else {
        const tempCanvas = document.getElementById('tempChart');
        const noDataMsg = document.createElement('p');
        noDataMsg.style.cssText = 'color:#6c757d;text-align:center;padding:20px;';
        noDataMsg.textContent = 'No temperature sensor data available.';
        tempCanvas.parentNode.replaceChild(noDataMsg, tempCanvas);
    }
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
    Write-Log "=== InfraEye Device Health Diagnostics Started ===" "INFO"

    Install-RequiredModules

    $systemType  = Get-SystemType
    $systemInfo  = Get-SystemInfo
    $cpuInfo     = Get-CPUInfo
    $ramInfo     = Get-RAMInfo
    $diskInfo    = Get-DiskInfo
    $gpuInfo     = Get-GPUInfo
    $batteryInfo = Get-BatteryInfo
    $startupApps = Get-StartupPrograms
    $tempInfo    = Get-TemperatureInfo
    $perfIssues  = Get-PerformanceAnalysis -CPU $cpuInfo -RAM $ramInfo -Disks $diskInfo -StartupApps $startupApps -Temperatures $tempInfo

    Write-Log "Generating HTML report..." "INFO"
    $htmlContent = New-HtmlReport `
        -SystemType    $systemType `
        -SystemInfo    $systemInfo `
        -CPU           $cpuInfo `
        -RAM           $ramInfo `
        -Disks         $diskInfo `
        -GPUs          $gpuInfo `
        -Battery       $batteryInfo `
        -StartupApps   $startupApps `
        -PerfIssues    $perfIssues `
        -Temperatures  $tempInfo

    $htmlContent | Out-File -FilePath $ReportFile -Encoding UTF8 -Force
    Write-Log "Report saved: $ReportFile" "SUCCESS"

    # Alert check
    $alertScript = Join-Path $PSScriptRoot "DeviceHealth_Alert.ps1"
    if (Test-Path $alertScript) {
        $criticalIssues = $perfIssues | Where-Object { $_.Severity -eq "HIGH" }
        if ($criticalIssues.Count -gt 0 -or $cpuInfo.UsagePercent -gt 95 -or $ramInfo.UsedPercent -gt 95) {
            Write-Log "Critical issues detected. Triggering alert script..." "WARN"
            & $alertScript -ReportFile $ReportFile -CPU $cpuInfo.UsagePercent -RAM $ramInfo.UsedPercent -StartupCount $startupApps.Count
        }
    }

    Write-Log "=== InfraEye Device Health Diagnostics Completed ===" "SUCCESS"
    Write-Host "`nReport generated: $ReportFile" -ForegroundColor Green
} catch {
    Write-Log "FATAL ERROR: $_" "ERROR"
    Write-Error $_
    exit 1
}
