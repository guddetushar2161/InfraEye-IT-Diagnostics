# InfraEye v2.1.6.1 — Python Diagnostics Quick Start Guide

## 🚀 Installation

### Step 1: Install Python (if not already installed)
- **Minimum:** Python 3.6+
- **Recommended:** Python 3.10+
- Download from: https://www.python.org/downloads/

### Step 2: Install Dependencies

```bash
# Install required dependencies
pip install -r requirements.txt

# (Optional) Install weasyprint for PDF export feature
pip install weasyprint
```

**Windows Users (PowerShell):**
```powershell
python -m pip install -r requirements.txt
python -m pip install weasyprint  # optional
```

---

## 📊 Running the Diagnostics

### Basic Usage

```bash
# Run with all features enabled (default)
python infra_eye_diagnostics.py

# The script will:
# 1. Collect system metrics (CPU, RAM, Disk, Network, Services)
# 2. Create a JSON snapshot in history/
# 3. Generate HTML report in Reports/
# 4. Display feature status and report location
```

### Use Cases & Commands

#### Use Case 1: Quick Health Check
```bash
python infra_eye_diagnostics.py
# Shows current status + trend comparison if history exists
# Output: InfraEye_Report_<hostname>_<timestamp>.html
```

#### Use Case 2: Generate PDF for Weekly Audit
```bash
python infra_eye_diagnostics.py --export-pdf
# Generates both HTML and PDF versions
# PDFs suitable for printing and archiving
```

#### Use Case 3: Scheduled Daily Report (No Historical Tracking)
```bash
python infra_eye_diagnostics.py --no-history --output /var/log/infra
# Lightweight - ideal for cron jobs that just need current snapshot
# No historical comparison overhead
```

#### Use Case 4: Minimal Report (Basic Metrics Only)
```bash
python infra_eye_diagnostics.py --no-history --no-pdf-button --no-depmap
# Fastest execution, smallest HTML output
# Perfect for automated monitoring without extras
```

#### Use Case 5: Dependency Map Only
```bash
python infra_eye_diagnostics.py --no-history --no-pdf-button
# Focuses on service health and dependencies
# Shows which services are cascading failures
```

---

## 🎯 Feature Overview

### Feature 1: Historical Comparison & Trends
```
What it does:
  ✓ Saves snapshot of every run to history/snapshot_YYYYMMDD_HHMMSS.json
  ✓ Loads previous snapshot and compares metrics
  ✓ Shows trend arrows: ▼ (better), ▲ (worse), ● (unchanged)
  ✓ Lists new failures (since last run) and resolved issues
  ✓ Auto-deletes snapshots older than 30 runs

When to use:
  - Monitor system trends over time
  - Identify recurring issues
  - Prove system stability to auditors
  - Track growth/degradation patterns

Disable with:
  --no-history
```

### Feature 2: PDF Export
```
What it does:
  ✓ Adds "🖨️ Export PDF" button in HTML report
  ✓ Browser print dialog generates PDF (Ctrl+P)
  ✓ --export-pdf flag generates PDF automatically (requires weasyprint)
  ✓ Professional formatting: headers, footers, page numbers
  ✓ A4 paper, 20mm margins, confidential footer

When to use:
  - Create audit trail for compliance
  - Share reports via email
  - Print for management/board meetings
  - Archive historical reports

Disable button with:
  --no-pdf-button
```

### Feature 3: Service Dependency Map
```
What it does:
  ✓ Creates interactive network graph of services
  ✓ Shows which failures are root causes vs cascading
  ✓ Color codes: 🔴 Failed, 🟡 Degraded, 🟢 Healthy
  ✓ Click nodes to see detailed status
  ✓ Shows "ROOT CAUSE" label on critical failures

When to use:
  - Troubleshoot systemic failures
  - Understand service relationships
  - Identify redundancy gaps
  - Plan disaster recovery

Disable with:
  --no-depmap
```

---

## 📁 Output Directory Structure

```
Reports/
├── InfraEye_Report_SERVER01_20250317_094530.html    # HTML report
├── InfraEye_Report_SERVER01_20250317.pdf            # PDF (if --export-pdf)
└── ...

history/
├── snapshot_20250317_094530.json    # Full metrics snapshot
├── snapshot_20250317_090000.json
├── snapshot_20250316_090000.json
└── ...                              # Keeps last 30 snapshots

Logs/
├── diagnostic_run_20250317.log
└── ...
```

---

## 🔧 Configuration

### Edit Service Dependencies

In `infra_eye_diagnostics.py`, find the `DEPENDENCY_MAP` section and add your services:

```python
DEPENDENCY_MAP = {
    "dns": ["web services", "email", "api calls"],
    "database": ["web apps", "api services"],
    "nginx": ["web apps", "static content"],
    # Add your services here
}
```

### Adjust Alert Thresholds

Find the `THRESHOLDS` dict in script:

```python
THRESHOLDS = {
    "cpu_percent": 85,      # Change to your threshold
    "ram_percent": 80,      # RAM alert level
    "disk_percent": 85,     # Disk usage alert
}
```

### Enable/Disable Features

At the top of the script, modify:

```python
FEATURES = {
    "historical_comparison": True,   # Set to False to disable
    "pdf_export_button": True,
    "dependency_map": True,
}
```

---

## 📋 Command Reference

```bash
# Show help and all available options
python infra_eye_diagnostics.py --help

# Show version
python infra_eye_diagnostics.py --version

# Run with specific output directory
python infra_eye_diagnostics.py --output /custom/path
python infra_eye_diagnostics.py -o /custom/path

# Disable individual features
python infra_eye_diagnostics.py --no-history
python infra_eye_diagnostics.py --no-pdf-button
python infra_eye_diagnostics.py --no-depmap

# Combine flags
python infra_eye_diagnostics.py --no-history --no-depmap --export-pdf

# Minimal setup (basic metrics only)
python infra_eye_diagnostics.py --no-history --no-pdf-button --no-depmap --output /tmp

# Full-featured diagnostics
python infra_eye_diagnostics.py --export-pdf
```

---

## 🐛 Troubleshooting

### Issue: "psutil not found"
```bash
# Solution:
pip install psutil
```

### Issue: "weasyprint not found" (when using --export-pdf)
```bash
# Solution: Install weasyprint
pip install weasyprint

# Or disable PDF feature:
python infra_eye_diagnostics.py --no-pdf-button
```

### Issue: Reports directory is empty
```bash
# Check script output for errors
python infra_eye_diagnostics.py

# If no output, verify Python version:
python --version  # Should be 3.6+

# Check permissions:
# - Must have write access to current directory
# - Reports/ folder will be created automatically
```

### Issue: "Permission denied" on Linux/Mac
```bash
# Make script executable
chmod +x infra_eye_diagnostics.py

# Then run:
./infra_eye_diagnostics.py
```

### Issue: Network/DNS data missing
```bash
# This is normal if curl is not installed
# Script gracefully handles missing tools
# DNS and public IP features are optional

# To enable: install curl
# Windows: already available
# Linux: apt install curl / yum install curl
# macOS: should be pre-installed
```

---

## 📈 Performance Expectations

| Operation | Time | Notes |
|-----------|------|-------|
| Full diagnostics | 3-5s | Network/DNS lookups can add 1-2s |
| HTML report generation | 0.5-1s | |
| PDF export (weasyprint) | 2-4s | Only if --export-pdf flag used |
| Snapshot save/cleanup | 0.1-0.5s | Very fast |
| **Total (all features)** | **6-10s** | Acceptable for scheduled runs |

---

## 🔄 Scheduling (Optional)

### Windows Task Scheduler

```powershell
# Create scheduled task
$action = New-ScheduledTaskAction -Execute "python.exe" `
    -Argument "C:\InfraEye\infra_eye_diagnostics.py --export-pdf"

$trigger = New-ScheduledTaskTrigger -Daily -At 8:00AM

Register-ScheduledTask -Action $action -Trigger $trigger `
    -TaskName "InfraEye Daily Diagnostics" -Description "Daily system health check"
```

### Linux Cron

```bash
# Add to crontab
crontab -e

# Run every morning at 8 AM
0 8 * * * cd /home/user/InfraEye && python infra_eye_diagnostics.py --export-pdf >> /var/log/infra.log 2>&1
```

---

## 📚 More Information

- **Full Documentation:** See CHANGELOG.md for detailed feature descriptions
- **Updated Config:** Check config.json for all available settings
- **Main Script:** Comments in infra_eye_diagnostics.py explain how everything works
- **Author:** Tushar Gudde (https://tushargudde.tech)
- **GitHub:** https://github.com/guddetushar2161/InfraEye-IT-Diagnostics

---

## ✅ Testing Checklist

Before putting into production, test:

- [ ] **Installation:** Run `pip install -r requirements.txt` successfully
- [ ] **Basic Run:** `python infra_eye_diagnostics.py` completes without errors
- [ ] **Report Generation:** Check that HTML report was created in Reports/
- [ ] **History:** Run twice and verify history/snapshot_*.json files created
- [ ] **Trends:** Second run shows previous comparison, not "Baseline Run"
- [ ] **PDF Export:** If using --export-pdf, verify PDF generated successfully
- [ ] **Dependencies:** Click nodes in dependency map, verify no JavaScript errors
- [ ] **Permissions:** All files readable/writable by your user account
- [ ] **Schedule:** If scheduling, verify task runs at correct time

---

**Ready to monitor your infrastructure? Start with:**
```bash
python infra_eye_diagnostics.py
```
