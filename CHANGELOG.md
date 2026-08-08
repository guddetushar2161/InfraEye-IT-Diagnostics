# Changelog - InfraEye IT Diagnostics & Infrastructure Intelligence

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [2.1.6.1] — 2025-03-17

### ✨ Added (Major Features)

#### **Feature 1: Historical Comparison & Trend Analysis**
- New `HistoryManager` class for persistent snapshot tracking
- Automatic JSON snapshot storage in `history/` subdirectory with timestamps
- Snapshot format: `snapshot_YYYYMMDD_HHMMSS.json` containing full diagnostic data
- Cross-run comparison logic that:
  - Loads the most recent previous snapshot on each run
  - Displays trend indicators (▲ worse, ▼ better, ● unchanged)
  - Shows metric deltas (current vs. previous for CPU%, RAM%, Disk%, services)
  - Highlights **"First Seen"** issues (new failures not in previous snapshot)
  - Shows **"Resolved Since Last Run"** section for recovered services
  - Displays **"Baseline Run — No Previous Data"** on first execution
- Automatic rolling-window cleanup: keeps last 30 snapshots, auto-deletes older ones
- CLI flag: `--no-history` to disable this feature

#### **Feature 2: Print-to-PDF / Weekly Audit Log Export**
- Sticky "🖨️ Export PDF" button in top-right corner of HTML reports
- Browser-native `window.print()` integration with dedicated `@media print` CSS
- PDF-optimized styling:
  - Print-friendly black-and-white color scheme
  - Removes interactive elements (buttons, buttons hidden)
  - Page headers with tool name, hostname, report timestamp
  - Professional footers: "Confidential — IT Internal Use | Page X of Y"
  - Prevents table/metric breaks across pages (`page-break-inside: avoid`)
  - A4 paper size with proper margins
- Optional Python CLI flag: `--export-pdf`
  - Uses `weasyprint` library (if installed) to auto-generate PDF alongside HTML
  - PDF filename format: `InfraEye_Report_<hostname>_<YYYYMMDD>.pdf`
  - Graceful fallback if `weasyprint` not installed
- CLI flag: `--no-pdf-button` to disable the export button in HTML

#### **Feature 3: Visual Service Dependency Map**
- Interactive HTML5 network graph using **vis.js** (loaded via CDN — zero external dependencies)
- Service dependency configuration via `DEPENDENCY_MAP` dictionary
- Node visualization rules based on service status:
  - 🔴 **Red pulsing node** = Currently FAILED
  - 🟡 **Yellow node** = Degraded (running but flagged)
  - 🟢 **Green node** = Healthy
  - ⚫ **Grey node** = Not detected / not applicable
- Directed arrows show dependency flow (root service → affected dependents)
- Intelligent root-cause detection:
  - Identifies failed root services and their cascading impact
  - Marks root-cause services with "⚠ ROOT CAUSE" label
  - Colors edges red when a root service is failed
- Interactive tooltips on node click showing:
  - Service name, current status, last check time
  - Dependency relationship details
- Green success banner ("✅ No Dependency Failures Detected") when all services healthy
- Includes below-graph legend explaining node colors and edge directions
- CLI flag: `--no-depmap` to disable dependency visualization

### 🐛 Fixed

#### **RAM Calculation Bug**
- **Issue:** Report always showed "upgrade to 16GB+ RAM" regardless of actual current RAM
- **Root Cause:** Upgrade recommendation did not calculate based on current installed RAM
- **Solution:** 
  - Fixed logic to check current RAM usage percentage AND total installed RAM
  - Now only recommends upgrade if:
    1. Current total RAM < 16GB, **AND**
    2. Used percentage > 80% (configurable threshold)
  - Provides context-aware suggestions:
    - 8GB or below → suggest 16GB
    - 16GB or below → suggest 32GB
    - Above 16GB → no upgrade suggestion shown
  - Eliminates false-positive warnings for systems with adequate memory

### 📋 Added (General Improvements)

- **Feature Control Configuration:**
  - Top-level `FEATURES` dict for enabling/disabling all three new features
  - CLI flags: `--no-history`, `--no-pdf-button`, `--no-depmap`
  - Each feature can be independently toggled without code changes

- **Enhanced HTML Report:**
  - Responsive grid layout for metrics (adapts to mobile/tablet/desktop)
  - Professional gradient header with InfraEye branding
  - Trend visualization section with before/after comparisons
  - New issues/resolved issues sections with visual badges
  - Dependency map section with interactive graph and legend
  - Improved styling for dark/light mode compatibility

- **Better Diagnostics Data:**
  - `SystemDiagnostics.get_metric_values()` method for structured metric access
  - Enhanced service failure detection (cross-platform: Windows WMI + Linux systemctl)
  - Network interface enumeration with IP address details
  - Public IP detection (optional with timeout, graceful fallback)
  - Top 10 processes by CPU usage tracking

- **Improved CLI Documentation:**
  - Detailed argument parser with examples
  - Version flag: `--version` shows v2.1.6.1
  - Output flag: `--output` or `-o` to customize report directory
  - Comprehensive help text with usage examples

- **Code Quality:**
  - Extensive inline documentation and docstrings
  - Type hints throughout (Dict, List, Optional, Tuple, etc.)
  - Error handling with graceful degradation
  - Production-ready error logging with color-coded output
  - Modular class architecture (HistoryManager, DependencyMapGenerator, PDFExporter, etc.)

### 📦 Dependencies

- **Core:** Python 3.6+ with `psutil` (already required)
- **Optional:** `weasyprint` for `--export-pdf` CLI functionality
- **JavaScript:** vis.js loaded via CDN (no installation required)

### 🔄 Compatibility

- **Windows:** Full support (WMI for service checks)
- **Linux/Unix:** Full support (systemctl for service checks)
- **Python:** 3.6, 3.7, 3.8, 3.9, 3.10, 3.11, 3.12
- **Browsers:** Modern browsers supporting ES6 (for vis.js visualization)

### 📖 Migration Notes

- **From v2.0:** All additions are **backward compatible**
- Existing PowerShell diagnostic modules remain unchanged and fully functional
- New Python script (`infra_eye_diagnostics.py`) runs independently
- No database migrations or configuration breaking changes required
- Historical snapshots stored in `history/` directory (auto-created on first run)

### 🚀 Usage Examples

```bash
# Basic diagnostics with all features enabled
python infra_eye_diagnostics.py

# Disable historical tracking
python infra_eye_diagnostics.py --no-history

# Generate PDF report (requires: pip install weasyprint)
python infra_eye_diagnostics.py --export-pdf

# Disable dependency map visualization
python infra_eye_diagnostics.py --no-depmap

# Custom output directory
python infra_eye_diagnostics.py --output /custom/path

# Disable all features except basic metrics
python infra_eye_diagnostics.py --no-history --no-pdf-button --no-depmap
```

### 📸 Visual Enhancements

- Professional gradient header (#667eea → #764ba2)
- Color-coded metric cards (CPU, RAM, Disk in consistent gradient)
- Trend arrows with contextual colors (green ▼ better, red ▲ worse, grey ● unchanged)
- Service dependency graph with intelligent node positioning (hierarchical layout)
- Print-optimized PDF styling with borders and proper spacing
- Responsive design for mobile, tablet, and desktop viewing

### ✅ Testing Recommendations

1. **Feature 1 — History:**
   - Run script twice and verify `history/` folder contains two snapshots
   - Check HTML report shows trend indicators and comparisons
   - Verify cleanup after 30+ runs (old snapshots deleted)

2. **Feature 2 — PDF Export:**
   - Click "Export PDF" button and verify print dialog opens
   - Install weasyprint and run `--export-pdf` flag, verify PDF generated
   - Verify PDF printing doesn't show buttons or sensitive info

3. **Feature 3 — Dependency Map:**
   - Verify vis.js loads from CDN (check Network tab in browser DevTools)
   - Stop a service and rerun script; verify node turns red
   - Click nodes to see tooltips with service details
   - Verify green banner shows when all services healthy

4. **Bug Fix — RAM:**
   - Run on 8GB system with high RAM usage → should suggest 16GB upgrade
   - Run on 16GB system with high RAM usage → should suggest 32GB upgrade
   - Run on 16GB+ system with high RAM usage → should show no suggestion
   - Run on any system with low RAM usage → should show no upgrade suggestion

---

## [2.0.0] — 2023-06-15 (Previous Release)

### Features (PowerShell Modules)
- Device Health Diagnostics module with HTML reporting
- Network Diagnostics with latency and speed testing
- Infrastructure Discovery with network scanning
- Email alert system for critical thresholds
- Excel inventory export via ImportExcel module
- Professional HTML dashboards with Chart.js integration

---

## Version History (Abbreviated)

- **v1.5.0** — Added GPU and battery detection
- **v1.4.0** — Added ImportExcel support for inventories
- **v1.3.0** — Enhanced performance analysis with root-cause detection
- **v1.0.0** — Initial PowerShell automation framework

---

**Report Version:** 2.1.6.1  
**Last Updated:** 2025-03-17  
**Author:** [Tushar Gudde](https://tushargudde.tech)  
**License:** MIT  
**Repository:** [InfraEye-IT-Diagnostics](https://github.com/guddetushar2161/InfraEye-IT-Diagnostics)
