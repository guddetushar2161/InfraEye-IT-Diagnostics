# InfraEye — IT Diagnostics & Infrastructure Intelligence

> **InfraEye** — Automated IT diagnostics and infrastructure intelligence powered by PowerShell.

[![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue)](https://docs.microsoft.com/powershell)
[![PowerShell 7+](https://img.shields.io/badge/PowerShell-7%2B-blue)](https://github.com/PowerShell/PowerShell)
[![Windows 10/11/Server](https://img.shields.io/badge/Windows-10%20%7C%2011%20%7C%20Server-lightgrey)](https://microsoft.com/windows)
[![Version](https://img.shields.io/badge/Version-2.1.6.1-green)](https://tushargudde.tech)
[![Python 3.6+](https://img.shields.io/badge/Python-3.6%2B-blue)](https://www.python.org)

Enterprise automation toolkit for **device health diagnostics**, **network troubleshooting**, and **infrastructure discovery**. Generates professional HTML dashboards, Excel inventories, PDF audit logs, and automated email alerts for critical IT issues.

**Author:** [Tushar Gudde](https://tushargudde.tech) | **License:** MIT

> **🎉 v2.1.6.1 NEW:** Historical trend analysis, PDF export, interactive dependency maps, and RAM calculation bug fix!

---

## 📋 Table of Contents

- [Project Overview](#-project-overview)
- [✨ New in v2.1.6.1](#-new-in-v2161)
- [Modules](#-modules)
- [Folder Structure](#-folder-structure)
- [Prerequisites](#-prerequisites)
- [How to Run Scripts](#-how-to-run-scripts)
- [Email Configuration](#-email-configuration)
- [Report Examples](#-report-examples)
- [Alert Thresholds](#-alert-thresholds)

---

## 🔭 Project Overview

InfraEye is a production-ready PowerShell automation suite designed for IT administrators and infrastructure engineers. It provides:

- **Real-time device health monitoring** — CPU, RAM, disk, GPU, battery, and startup programs
- **Network performance diagnostics** — latency, packet loss, bandwidth analysis, and public IP info
- **Infrastructure discovery** — ping sweep, ARP analysis, device classification, DHCP and subnet analysis
- **Professional HTML dashboards** — dark/light mode, Chart.js charts, responsive tables
- **Excel device inventory** — exported automatically with the `ImportExcel` module
- **Automated email alerts** — triggered when critical thresholds are exceeded

### Compatibility

| Platform         | Supported |
|------------------|-----------|
| Windows 10       | ✅         |
| Windows 11       | ✅         |
| Windows Server   | ✅         |
| PowerShell 5.1   | ✅         |
| PowerShell 7+    | ✅         |

---

## ✨ New in v2.1.6.1

### 🎯 Feature 1: Historical Comparison & Trend Analysis

**Automatically track system health over time with intelligent trend detection.**

Each time you run the Python diagnostics script, it creates a timestamped JSON snapshot in `history/` that captures all system metrics. On subsequent runs, the HTML report displays:

- **Trend Indicators:** Visual arrows showing whether metrics improved (▼ green), worsened (▲ red), or stayed the same (● grey)
- **Side-by-Side Comparison:** Current metrics displayed next to previous run values for easy comparison
- **New Issues (First Seen):** Services that failed in this run but were working before are marked 🆕 with "First Seen" label
- **Resolved Issues:** Services that were failing before but are now healthy are highlighted ✓ "Resolved Since Last Run"
- **Baseline Detection:** First run shows "Baseline Run — No Previous Data" so you know it's the initial snapshot
- **Automatic Cleanup:** Script keeps the last 30 snapshots (roughly one month of daily runs) and auto-deletes older ones

**Example Comparison Display:**
```
CPU Usage: ◀◀ Current: 45% | Previous: 52% ▼ (improved 7%)
RAM Usage: ◀◀ Current: 78% | Previous: 75% ▲ (worsened 3%)
Disk C:   ◀◀ Current: 62% | Previous: 62% ● (unchanged)

🆕 New Issues (First Seen)
  🔴 mongodb failed  
  
✓ Resolved (Since Last Run)
  🟢 elastic_search recovered
```

**CLI Usage:**
```bash
python infra_eye_diagnostics.py                # Run with history enabled (default)
python infra_eye_diagnostics.py --no-history   # Disable historical tracking
```

---

### 🖨️ Feature 2: Print-to-PDF / Weekly Audit Log Export

**Generate professional, print-ready PDF reports for audits and compliance documentation.**

##### Browser-Based PDF Export
A sticky "🖨️ Export PDF" button in the top-right corner of HTML reports allows immediate export via the browser's print dialog:
- One-click printing with `Ctrl+P` or button click
- Automated styling removes interactive elements from PDF
- Professional page headers with: Tool name, Hostname, Report Timestamp
- Confidential footer on every page: "Confidential — IT Internal Use | Page X of Y"
- A4 paper size with 20mm margins for professional printing
- Prevents tables and metrics from breaking across pages
- Black-and-white friendly colors while maintaining status badges

##### Python CLI PDF Generation (Optional)
For automated, headless PDF generation via command line, use `--export-pdf`:

```bash
pip install weasyprint  # One-time setup (optional)
python infra_eye_diagnostics.py --export-pdf
```

This generates a PDF file immediately without opening a browser:
- Filename: `InfraEye_Report_<hostname>_<YYYYMMDD>.pdf`
- Saved to the same directory as HTML report
- Graceful fallback if `weasyprint` not installed

**Example Use Case:**
```bash
# Generate HTML + PDF every Monday morning via cron/Task Scheduler
python infra_eye_diagnostics.py --export-pdf --output /var/log/infra_reports
# Result: InfraEye_Report_server01_20250317.html + .pdf in the same folder
```

**CLI Flags:**
```bash
python infra_eye_diagnostics.py --export-pdf        # Auto-generate PDF
python infra_eye_diagnostics.py --no-pdf-button     # Hide button in HTML
python infra_eye_diagnostics.py --export-pdf --no-pdf-button  # PDF only, no button
```

---

### 🔗 Feature 3: Visual Service Dependency Map

**Discover which service failures are root causes vs. cascading effects with an interactive network graph.**

An interactive HTML5 visualization powered by **vis.js** shows your service ecosystem as an interactive network diagram. Learn at a glance which service failures are probably root causes and which are downstream impacts.

#### How It Works:

1. **Service Dependency Configuration:** Define your service dependencies in Python:
   ```python
   DEPENDENCY_MAP = {
       "dns":      ["web browsing", "email", "database connections"],
       "database": ["web applications", "api services"],
       "networking": ["ssh", "http", "https", "dns"],
   }
   ```

2. **Intelligent Node Visualization:**
   - 🔴 **Red Node** = Service is currently **FAILED**
   - 🟡 **Yellow Node** = Service is **DEGRADED** (running but flagged)
   - 🟢 **Green Node** = Service is **HEALTHY**
   - ⚫ **Grey Node** = Service not detected / not applicable

3. **Root Cause Detection:**
   - Failing root services are labeled "⚠ ROOT CAUSE"
   - Dependency arrows turn **red** when a parent service is failed
   - Arrows thicker for critical dependencies
   - All downstream impacts are visually connected

4. **Interactive Features:**
   - Click any node to see full details (name, status, last check time)
   - Drag nodes to reorganize layout
   - Zoom and pan across the network
   - Automatically arranges using hierarchical physics (root services at top)

5. **Smart Display:**
   - If all services are healthy, shows ✅ green banner: "No Dependency Failures Detected"
   - Includes legend explaining node colors and edge meanings
   - No external dependencies — uses vis.js from CDN

**Example Output:**
```
┌────────────────────────────────────────────┐
│           Service Network                   │
│                                             │
│            🔴 DNS (FAILED)                  │
│             ↓ ↓ ↓ (red arrows)              │
│         🔴Web  🔴Email  🟢API               │
│                                             │
│  Legend:                                    │
│  🔴 Failed | 🟡 Degraded | 🟢 Healthy      │
└────────────────────────────────────────────┘
```

**CLI Usage:**
```bash
python infra_eye_diagnostics.py                 # Show dependency map (default)
python infra_eye_diagnostics.py --no-depmap     # Disable visualization
```

---

### 🐛 Bug Fix: RAM Upgrade Recommendation

**Fixed incorrect "upgrade to 16GB" recommendation that appeared regardless of current RAM.**

**Before Bug:**
```
RAM: 24GB installed, using 60%
⚠️ Upgrade to 16GB+ RAM for better performance  ❌ Incorrect & confusing!
```

**After Fix:**
```
RAM: 24GB installed, using 60%
✅ Healthy — No upgrade needed  ✓ Correct!

RAM: 8GB installed, using 85%
⚠️ Consider upgrading to 16GB for better performance  ✓ Helpful suggestion
```

**Logic:** Now only recommends upgrade if:
1. Current total RAM < 16GB, **AND**
2. Used percentage > 80% (threshold configurable)

---

### 🎛️ Feature Control Configuration

All three new features can be independently enabled/disabled:

```python
# At top of script
FEATURES = {
    "historical_comparison": True,   # Feature 1
    "pdf_export_button": True,       # Feature 2
    "dependency_map": True,          # Feature 3
}
```

**Via CLI (any combination):**
```bash
# Disable individual features
python infra_eye_diagnostics.py --no-history --no-pdf-button --no-depmap

# Mixed: enable trends and deps, disable PDF button
python infra_eye_diagnostics.py --no-pdf-button

# Show help with all options
python infra_eye_diagnostics.py --help
```

---

### 📊 Enhanced HTML Report Design

- **Professional Gradient Header:** Purple gradient with InfraEye branding and report metadata
- **Responsive Metrics Grid:** Adapts from 1 column (mobile) to 4 columns (desktop)
- **Color-Coded Badges:** Status indicators (✅ OK, ⚠️ WARNING, 🔴 CRITICAL)
- **Print Optimization:** Dedicated `@media print` CSS for clean PDF exports
- **Dark/Light Compatible:** Modern styles work in both browser modes
- **Zero Breaking Changes:** All updates are additive to existing report structure

---

## 🧩 Modules

### Module 1 — Device Health Diagnostics (`DeviceHealth/`)

Collects comprehensive system health data and generates an HTML dashboard with:

| Feature | Details |
|---|---|
| System Type Detection | Laptop / Desktop / Workstation via Win32_SystemEnclosure |
| System Info | Computer name, manufacturer, model, serial number, OS, uptime |
| CPU | Model, usage %, total cores & logical processors |
| RAM | Total / Used / Available GB, upgrade recommendation if < 8 GB |
| Disk | SSD/HDD detection, usage, free space per drive |
| GPU | Model, integrated vs. dedicated, VRAM |
| Battery | Status and charge % |
| Startup Programs | Count with warning if > 10, alert if > 20 |
| Performance Analysis | Root cause detection: low RAM, high CPU, disk bottleneck, too many processes/startups |

**Alert Script** (`DeviceHealth_Alert.ps1`) triggers when:
- CPU > 95%
- RAM > 95%
- Disk free < 10%
- Startup apps > 20

---

### Module 2 — Network Diagnostics (`NetworkDiagnostics/`)

Performs comprehensive network diagnostics and generates an HTML dashboard with:

| Feature | Details |
|---|---|
| Adapter Info | Name, type (WiFi/LAN), MAC, IP, gateway, DNS, link speed |
| Public IP | IP, ISP, city, country, timezone via ipinfo.io |
| Ping Tests | 8.8.8.8, 1.1.1.1, 8.8.4.4 — latency & packet loss |
| Speed Test | Download speed via public test endpoint |
| Bandwidth Processes | Top 10 processes by TCP connection count |
| Root Cause Analysis | Weak WiFi, high latency, packet loss, slow ISP, DNS delays |

**Alert Script** (`NetworkDiagnostics_Alert.ps1`) triggers when:
- Latency > 200 ms
- Packet loss > 5%
- Download speed < 10 Mbps

---

### Module 3 — Infrastructure Discovery (`InfrastructureDiscovery/`)

Scans all private network ranges and generates an HTML dashboard + Excel inventory:

| Feature | Details |
|---|---|
| Network Ranges | Auto-detects 192.168.x.x, 10.x.x.x, 172.16.x.x |
| Ping Sweep | Multi-threaded ping sweep with configurable timeout/threads |
| ARP Analysis | Reads ARP table for MAC addresses |
| Device Classification | Servers, workstations, printers, routers, switches, firewalls |
| Hostname Resolution | Reverse DNS for all discovered IPs |
| Vendor Lookup | MAC OUI-based vendor identification |
| Server Role Detection | Domain Controller, File Server, DB Server, DHCP, DNS, IIS |
| DHCP Analysis | Range, used/available IPs, pool usage % |
| Subnet Analysis | Subnet mask, total/usable hosts, broadcast address |
| Excel Export | All devices exported to `NetworkDevices_TIMESTAMP.xlsx` |
| Duplicate IP Detection | Flags and alerts on duplicate IP addresses |

**Alert Script** (`Infrastructure_Alert.ps1`) triggers when:
- Duplicate IP detected
- DHCP pool usage > 90%
- Offline servers detected
- Unknown devices found

---

## 📁 Folder Structure

```
InfraEye-IT-Diagnostics/
│
├── DeviceHealth/
│   ├── Reports/                    # HTML reports (auto-created)
│   ├── Logs/                       # Log files (auto-created)
│   ├── DeviceHealth_Main.ps1       # Main diagnostics script
│   └── DeviceHealth_Alert.ps1      # Alert/email script
│
├── NetworkDiagnostics/
│   ├── Reports/                    # HTML reports (auto-created)
│   ├── Logs/                       # Log files (auto-created)
│   ├── NetworkDiagnostics_Main.ps1 # Main diagnostics script
│   └── NetworkDiagnostics_Alert.ps1# Alert/email script
│
├── InfrastructureDiscovery/
│   ├── Reports/                    # HTML + Excel reports (auto-created)
│   ├── Logs/                       # Log files (auto-created)
│   ├── InfrastructureDiscovery_Main.ps1  # Main discovery script
│   └── Infrastructure_Alert.ps1   # Alert/email script
│
├── config.json                     # Global configuration
└── README.md                       # This file
```

> **Note:** `Reports/` and `Logs/` directories are created automatically when scripts run.

---

## ⚙️ Prerequisites

### Required PowerShell Modules

The scripts automatically install missing modules on first run:

```powershell
Install-Module -Name ImportExcel   -Scope CurrentUser -Force
Install-Module -Name PSWriteHTML   -Scope CurrentUser -Force
```

### Execution Policy

Enable script execution if needed:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

## 🚀 How to Run Scripts

> **⚠️ Administrator Privileges Recommended**
>
> For best results, **run all modules as Administrator**. Elevated privileges unlock:
> - Full battery report generation (`powercfg /batteryreport`) with hardware identification
> - Complete ARP table visibility and MAC address resolution (Infrastructure Discovery)
> - Network adapter diagnostics and TCP port scanning
> - System performance counters and cleanup analysis
>
> **Quick way to launch as Administrator:**
> ```powershell
> # Right-click PowerShell → "Run as administrator", then:
> Set-Location "C:\path\to\InfraEye-IT-Diagnostics"
> powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\DeviceHealth\DeviceHealth_Main.ps1"
> ```
> Or use `Start-Process` to elevate from an existing shell:
> ```powershell
> Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PWD\DeviceHealth\DeviceHealth_Main.ps1`""
> ```

### Module 1 — Device Health Diagnostics

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\DeviceHealth\DeviceHealth_Main.ps1"
```

Run alert check independently:
```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\DeviceHealth\DeviceHealth_Alert.ps1"
```

### Module 2 — Network Diagnostics

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\NetworkDiagnostics\NetworkDiagnostics_Main.ps1"
```

Run alert check independently:
```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\NetworkDiagnostics\NetworkDiagnostics_Alert.ps1"
```

### Module 3 — Infrastructure Discovery

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\InfrastructureDiscovery\InfrastructureDiscovery_Main.ps1"
```

> **Note:** Infrastructure discovery may require administrator privileges for ARP table and port scanning.

Run as Administrator:
```powershell
Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `".\InfrastructureDiscovery\InfrastructureDiscovery_Main.ps1`""
```

---

## 📧 Email Configuration

Edit `config.json` to configure email alerts:

```json
{
    "smtp_server": "smtp.office365.com",
    "smtp_port": 587,
    "from_email": "alerts@company.com",
    "to_email": "admin@company.com",
    "smtp_username": "alerts@company.com",
    "smtp_password": "YOUR_PASSWORD_HERE",
    "enable_ssl": true
}
```

### Common SMTP Configurations

| Provider | SMTP Server | Port | SSL |
|---|---|---|---|
| Microsoft 365 | smtp.office365.com | 587 | true |
| Gmail | smtp.gmail.com | 587 | true |
| Outlook.com | smtp-mail.outlook.com | 587 | true |
| On-premise Exchange | your-exchange-server | 25 or 587 | varies |

> **Security Note:** Store the SMTP password securely. Consider using an app-specific password or OAuth2 where supported.

---

## 📊 Report Examples

### Report Naming Format

```
ModuleName_YYYYMMDD_HHMMSS.html
```

**Examples:**
```
DeviceHealth_20260314_211200.html
NetworkDiagnostics_20260314_211305.html
InfrastructureDiscovery_20260314_211420.html
NetworkDevices_20260314_211420.xlsx
```

### Dashboard Features

All HTML reports include:
- 🌙 **Dark / Light mode toggle** (default: Light mode)
- 📊 **Chart.js charts** — bar, doughnut, and line charts
- 🎯 **Summary cards** — key metrics at a glance
- 📋 **Responsive data tables** — alternating row colors, hover effects
- 🔍 **Device search filter** (Infrastructure Discovery)
- 📌 **Footer** — Report Version, Author, Website link

---

## 🚨 Alert Thresholds

Configure thresholds in `config.json`:

```json
"alert_thresholds": {
    "cpu_percent": 95,
    "ram_percent": 95,
    "disk_free_percent": 10,
    "startup_apps_max": 20,
    "latency_ms": 200,
    "packet_loss_percent": 5,
    "download_speed_mbps": 10,
    "dhcp_pool_percent": 90
}
```

---

## 📝 Logging

All scripts generate timestamped log files in their respective `Logs/` directories.

**Log format:**
```
2026-03-14 21:10:11 | DeviceHealth | INFO | Collecting CPU data
2026-03-14 21:10:12 | DeviceHealth | SUCCESS | Report saved: C:\...\Reports\DeviceHealth_20260314_211011.html
```

**Log levels:** `INFO` | `WARN` | `ERROR` | `SUCCESS`

---

## 🔒 Security Notes

- Scripts use `$ErrorActionPreference = "Stop"` and `Set-StrictMode -Version Latest`
- SMTP password is stored in plain text in `config.json` — secure this file with appropriate NTFS permissions
- Infrastructure discovery performs port scanning — ensure you have authorization to scan the network
- The scripts use `Get-CimInstance` (preferred over deprecated `Get-WmiObject`)

---

## 📄 License

This project is provided as-is for IT automation and diagnostics purposes.

**Created by:** [Tushar Gudde](https://tushargudde.tech)  
**Website:** [https://tushargudde.tech](https://tushargudde.tech)  
**Version:** 2.0
