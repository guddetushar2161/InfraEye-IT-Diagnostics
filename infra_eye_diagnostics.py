#!/usr/bin/env python3
"""
InfraEye IT Diagnostics & Infrastructure Intelligence v2.1.6.1

A comprehensive, production-ready Python-based IT diagnostic tool that scans
system health (CPU, RAM, disk, services, network) and generates professional
HTML reports with advanced features:

    - Feature 1: Historical Comparison & Trend Analysis
    - Feature 2: Print-to-PDF / Weekly Audit Log Export
    - Feature 3: Visual Service Dependency Map

Author: Tushar Gudde (https://tushargudde.tech)
License: MIT
"""

import os
import sys
import json
import argparse
import socket
import platform
import psutil
import shutil
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any
import subprocess
import re

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CONFIGURATION & FEATURE FLAGS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

VERSION = "2.1.6.1"

# Global feature flags - can be toggled via CLI or config
FEATURES = {
    "historical_comparison": True,  # Feature 1: Track trends & compare with previous runs
    "pdf_export_button": True,      # Feature 2: Add PDF export functionality
    "dependency_map": True,         # Feature 3: Visual service dependency graph
}

# Service dependency map for Feature 3
# Maps root services to their downstream dependents
DEPENDENCY_MAP = {
    "networking": ["ssh", "http", "https", "dns", "smtp"],
    "dns": ["web services", "email", "api calls", "database connections"],
    "http": ["web applications", "api services"],
    "https": ["web applications", "api services", "authentication"],
    "database": ["web applications", "api services", "business logic"],
    "ssh": ["remote administration", "file transfer", "monitoring"],
    "smtp": ["email delivery", "notifications", "alerts"],
}

# Alert thresholds (percentage-based)
THRESHOLDS = {
    "cpu_percent": 85,          # Alert if CPU > 85%
    "ram_percent": 80,          # Alert if RAM > 80%
    "disk_percent": 85,         # Alert if Disk > 85%
    "temp_celsius": 80,         # Alert if CPU Temp > 80°C (if available)
}

# History settings
HISTORY_MAX_SNAPSHOTS = 30      # Keep last 30 snapshots (~1 month of daily runs)
SNAPSHOT_DIR = "history"        # Directory to store JSON snapshots

# Color scheme for HTML reports (theme-aware)
COLORS = {
    "success": "#4CAF50",        # Green
    "warning": "#FF9800",        # Orange
    "error": "#F44336",          # Red
    "info": "#2196F3",           # Blue
    "neutral": "#9E9E9E",        # Grey
    "bg_light": "#FFFFFF",       # White
    "bg_dark": "#1e1e1e",        # Dark grey
    "text_light": "#333333",     # Dark text
    "text_dark": "#FFFFFF",      # Light text
}

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# UTILITY FUNCTIONS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def log(message: str, level: str = "INFO"):
    """Log messages with timestamp and level."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    level_colors = {
        "INFO": "\033[36m",     # Cyan
        "WARN": "\033[33m",     # Yellow
        "ERROR": "\033[31m",    # Red
        "SUCCESS": "\033[32m",  # Green
    }
    color = level_colors.get(level, "\033[37m")  # Default white
    reset = "\033[0m"
    print(f"{color}[{timestamp}] {level:8} →{reset} {message}")


def ensure_directory(path: str) -> Path:
    """Create directory if it doesn't exist."""
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p


def get_hostname() -> str:
    """Get system hostname."""
    try:
        return socket.gethostname()
    except Exception:
        return platform.node()


def format_bytes(bytes_val: int) -> str:
    """Convert bytes to human-readable format."""
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if bytes_val < 1024.0:
            return f"{bytes_val:.2f} {unit}"
        bytes_val /= 1024.0
    return f"{bytes_val:.2f} PB"


def get_trend_indicator(current: float, previous: Optional[float]) -> Tuple[str, str]:
    """
    Determine trend direction and color.
    
    Args:
        current: Current metric value
        previous: Previous metric value
        
    Returns:
        Tuple of (indicator_symbol, color_code)
        - ▼ (green): Better trend (lower is better for these metrics)
        - ▲ (red): Worse trend
        - ● (grey): Unchanged or first run
    """
    if previous is None:
        return "●", COLORS["neutral"]  # First run - no baseline
    
    diff = current - previous
    if abs(diff) < 0.01:  # Essentially unchanged (tolerance for float precision)
        return "●", COLORS["neutral"]
    elif diff < 0:  # Improved (lower value = better for CPU/RAM/Disk)
        return "▼", COLORS["success"]
    else:  # Worsened (higher value = worse)
        return "▲", COLORS["error"]


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# FEATURE 1: HISTORICAL SNAPSHOT MANAGEMENT
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class HistoryManager:
    """Manages historical snapshots and trend analysis for Feature 1."""
    
    def __init__(self, snapshot_dir: str = SNAPSHOT_DIR):
        """Initialize history manager."""
        self.snapshot_dir = ensure_directory(snapshot_dir)
        self.current_snapshot = {}
        self.previous_snapshot = None
        self.snapshot_timestamp = None
    
    def save_snapshot(self, data: Dict[str, Any]) -> Path:
        """
        Save current diagnostic data as timestamped JSON snapshot.
        
        Args:
            data: System diagnostic data to save
            
        Returns:
            Path to saved snapshot file
        """
        self.snapshot_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.current_snapshot = data
        
        snapshot_file = self.snapshot_dir / f"snapshot_{self.snapshot_timestamp}.json"
        try:
            with open(snapshot_file, "w") as f:
                json.dump({
                    "timestamp": self.snapshot_timestamp,
                    "hostname": get_hostname(),
                    "data": data,
                }, f, indent=2)
            log(f"Snapshot saved: {snapshot_file}", "SUCCESS")
            return snapshot_file
        except Exception as e:
            log(f"Failed to save snapshot: {e}", "ERROR")
            return None
    
    def load_previous_snapshot(self) -> Optional[Dict[str, Any]]:
        """
        Load the most recent previous snapshot (excluding current run).
        
        Returns:
            Previous snapshot data or None if no history exists
        """
        try:
            snapshots = sorted(self.snapshot_dir.glob("snapshot_*.json"), reverse=True)
            
            # Filter out current snapshot if it exists
            previous_snapshots = [
                s for s in snapshots 
                if not s.name.startswith(f"snapshot_{self.snapshot_timestamp}")
            ] if self.snapshot_timestamp else snapshots
            
            if not previous_snapshots:
                return None
            
            with open(previous_snapshots[0], "r") as f:
                snapshot_data = json.load(f)
                self.previous_snapshot = snapshot_data.get("data", {})
                log(f"Loaded previous snapshot: {previous_snapshots[0].name}", "INFO")
                return self.previous_snapshot
        except Exception as e:
            log(f"Failed to load previous snapshot: {e}", "WARN")
            return None
    
    def cleanup_old_snapshots(self, max_snapshots: int = HISTORY_MAX_SNAPSHOTS):
        """
        Delete old snapshots beyond max limit (rolling window).
        
        Args:
            max_snapshots: Maximum number of snapshots to keep
        """
        try:
            snapshots = sorted(self.snapshot_dir.glob("snapshot_*.json"))
            if len(snapshots) > max_snapshots:
                to_delete = snapshots[:-max_snapshots]
                for old in to_delete:
                    old.unlink()
                log(f"Cleaned up {len(to_delete)} old snapshots", "INFO")
        except Exception as e:
            log(f"Cleanup failed: {e}", "WARN")
    
    def get_comparison(self) -> Dict[str, Any]:
        """
        Generate detailed comparison between current and previous data.
        
        Returns:
            Dictionary with trend information and diffs
        """
        if not self.previous_snapshot:
            return {"is_baseline": True, "message": "Baseline Run — No Previous Data"}
        
        comparison = {
            "is_baseline": False,
            "metrics": {},
            "new_issues": [],
            "resolved_issues": [],
        }
        
        # Compare key metrics (CPU, RAM, Disk usage percentages)
        metrics_to_compare = ["cpu_percent", "ram_percent", "disk_percent", "failed_services_count"]
        
        for metric in metrics_to_compare:
            current = self.current_snapshot.get(metric)
            previous = self.previous_snapshot.get(metric)
            
            if current is not None and previous is not None:
                trend, color = get_trend_indicator(float(current), float(previous))
                comparison["metrics"][metric] = {
                    "current": current,
                    "previous": previous,
                    "trend": trend,
                    "color": color,
                }
        
        # Identify new and resolved service failures
        current_failed = set(self.current_snapshot.get("failed_services", []))
        previous_failed = set(self.previous_snapshot.get("failed_services", []))
        
        comparison["new_issues"] = list(current_failed - previous_failed)
        comparison["resolved_issues"] = list(previous_failed - current_failed)
        
        return comparison


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SYSTEM DIAGNOSTICS DATA COLLECTION
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class SystemDiagnostics:
    """Collect comprehensive system health metrics."""
    
    def __init__(self):
        """Initialize diagnostics collector."""
        self.data = {}
        self.timestamp = datetime.now()
    
    def collect_all(self) -> Dict[str, Any]:
        """Collect all system metrics."""
        log("Starting system diagnostics...", "INFO")
        
        self.data = {
            "timestamp": self.timestamp.isoformat(),
            "hostname": get_hostname(),
            "platform": platform.system(),
            "version": VERSION,
            "cpu": self._collect_cpu(),
            "memory": self._collect_memory(),
            "disk": self._collect_disk(),
            "network": self._collect_network(),
            "processes": self._collect_processes(),
            "services": self._collect_services(),
        }
        
        log("Diagnostics collection complete", "SUCCESS")
        return self.data
    
    def _collect_cpu(self) -> Dict[str, Any]:
        """Collect CPU metrics."""
        try:
            cpu_percent = psutil.cpu_percent(interval=1)
            cpu_count = psutil.cpu_count(logical=False)
            logical_count = psutil.cpu_count(logical=True)
            cpu_freq = psutil.cpu_freq()
            
            return {
                "percent": cpu_percent,
                "cores": cpu_count,
                "logical_cores": logical_count,
                "frequency_mhz": cpu_freq.current if cpu_freq else 0,
                "status": "critical" if cpu_percent > THRESHOLDS["cpu_percent"] else "normal",
            }
        except Exception as e:
            log(f"CPU collection failed: {e}", "WARN")
            return {"percent": 0, "cores": 0, "error": str(e)}
    
    def _collect_memory(self) -> Dict[str, Any]:
        """
        Collect RAM metrics.
        
        BUG FIX: Correctly calculate upgrade recommendation based on CURRENT usage,
        not always showing "upgrade to 16GB". Only recommend upgrade if current
        usage is consistently high and less than 16GB installed.
        """
        try:
            vm = psutil.virtual_memory()
            total_gb = vm.total / (1024**3)  # Convert bytes to GB
            used_gb = vm.used / (1024**3)
            available_gb = vm.available / (1024**3)
            used_percent = vm.percent
            
            # Fixed logic: Only recommend upgrade if:
            # 1. Current RAM < 16GB, AND
            # 2. Used percentage is consistently high (>80%)
            upgrade_recommended = (
                total_gb < 16 and used_percent > THRESHOLDS["ram_percent"]
            )
            
            upgrade_suggestion = None
            if upgrade_recommended:
                # Suggest upgrade to next tier (8GB→16GB, 16GB→32GB, etc.)
                suggested_gb = 16 if total_gb < 8 else (32 if total_gb < 16 else 64)
                upgrade_suggestion = f"Consider upgrading to {suggested_gb}GB for better performance"
            
            return {
                "total_gb": round(total_gb, 2),
                "used_gb": round(used_gb, 2),
                "available_gb": round(available_gb, 2),
                "percent": used_percent,
                "upgrade_recommended": upgrade_recommended,
                "upgrade_suggestion": upgrade_suggestion,
                "status": "critical" if used_percent > THRESHOLDS["ram_percent"] else "normal",
            }
        except Exception as e:
            log(f"Memory collection failed: {e}", "WARN")
            return {"total_gb": 0, "used_gb": 0, "error": str(e)}
    
    def _collect_disk(self) -> Dict[str, Any]:
        """Collect disk metrics for all partitions."""
        try:
            disks = {}
            for partition in psutil.disk_partitions():
                try:
                    usage = psutil.disk_usage(partition.mountpoint)
                    disks[partition.device] = {
                        "mountpoint": partition.mountpoint,
                        "total_gb": round(usage.total / (1024**3), 2),
                        "used_gb": round(usage.used / (1024**3), 2),
                        "free_gb": round(usage.free / (1024**3), 2),
                        "percent": usage.percent,
                        "status": "critical" if usage.percent > THRESHOLDS["disk_percent"] else "normal",
                    }
                except (PermissionError, OSError):
                    pass
            return disks if disks else {"error": "No disk partitions found"}
        except Exception as e:
            log(f"Disk collection failed: {e}", "WARN")
            return {"error": str(e)}
    
    def _collect_network(self) -> Dict[str, Any]:
        """Collect network interface metrics."""
        try:
            interfaces = {}
            for name, addrs in psutil.net_if_addrs().items():
                interfaces[name] = [
                    {
                        "family": str(addr.family),
                        "address": addr.address,
                        "netmask": addr.netmask,
                    }
                    for addr in addrs
                ]
            
            # Get public IP (optional, graceful fallback)
            public_ip = "Unavailable"
            try:
                result = subprocess.run(
                    ["curl", "-s", "https://api.ipify.org"],
                    capture_output=True,
                    timeout=5,
                    text=True
                )
                if result.returncode == 0:
                    public_ip = result.stdout.strip()
            except Exception:
                pass
            
            return {
                "interfaces": interfaces,
                "public_ip": public_ip,
                "io_stats": {
                    "bytes_sent": psutil.net_io_counters().bytes_sent,
                    "bytes_recv": psutil.net_io_counters().bytes_recv,
                },
            }
        except Exception as e:
            log(f"Network collection failed: {e}", "WARN")
            return {"error": str(e)}
    
    def _collect_processes(self) -> Dict[str, Any]:
        """Collect top CPU and memory consuming processes."""
        try:
            processes = []
            for proc in psutil.process_iter(['pid', 'name', 'cpu_percent', 'memory_percent']):
                try:
                    processes.append({
                        "name": proc.info['name'],
                        "pid": proc.info['pid'],
                        "cpu_percent": proc.info['cpu_percent'],
                        "memory_percent": proc.info['memory_percent'],
                    })
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            
            # Sort and return top 10 by CPU usage
            return {
                "top_by_cpu": sorted(processes, key=lambda x: x["cpu_percent"], reverse=True)[:10],
                "total_count": len(processes),
            }
        except Exception as e:
            log(f"Process collection failed: {e}", "WARN")
            return {"error": str(e)}
    
    def _collect_services(self) -> Dict[str, Any]:
        """
        Collect service status (platform-specific).
        On Windows: checks critical Windows services.
        On Linux: checks systemd services.
        """
        failed_services = []
        
        try:
            if sys.platform == "win32":
                # Windows service check
                try:
                    import wmi
                    wmi_obj = wmi.WMI()
                    services = wmi_obj.Win32_Service(State="Stopped")
                    failed_services = [s.Name for s in services if "Disabled" not in s.StartMode]
                except (ImportError, Exception):
                    # Graceful fallback if WMI not available
                    pass
            else:
                # Linux/Unix service check using systemctl
                try:
                    result = subprocess.run(
                        ["systemctl", "list-units", "--failed", "--plain"],
                        capture_output=True,
                        text=True,
                        timeout=5
                    )
                    if result.returncode == 0:
                        for line in result.stdout.split('\n'):
                            if line.strip():
                                service_name = line.split()[0]
                                failed_services.append(service_name)
                except Exception:
                    pass
        except Exception as e:
            log(f"Service collection failed: {e}", "WARN")
        
        return {
            "failed_services": failed_services,
            "failed_services_count": len(failed_services),
            "status": "warning" if failed_services else "healthy",
        }
    
    # Store metric values for trending
    def get_metric_values(self) -> Dict[str, float]:
        """Extract key metrics for trend comparison."""
        return {
            "cpu_percent": self.data.get("cpu", {}).get("percent", 0),
            "ram_percent": self.data.get("memory", {}).get("percent", 0),
            "disk_percent": max(
                [d.get("percent", 0) for d in self.data.get("disk", {}).values() if isinstance(d, dict)]
                or [0]
            ),
            "failed_services_count": self.data.get("services", {}).get("failed_services_count", 0),
            "failed_services": self.data.get("services", {}).get("failed_services", []),
        }


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# FEATURE 3: SERVICE DEPENDENCY MAP GENERATOR
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class DependencyMapGenerator:
    """
    Generate interactive service dependency graph for Feature 3.
    Uses vis.js library for visualization.
    """
    
    def __init__(self, dependency_map: Dict[str, List[str]], failed_services: List[str]):
        """
        Initialize dependency map generator.
        
        Args:
            dependency_map: Root services mapped to dependents
            failed_services: List of currently failed services
        """
        self.dependency_map = dependency_map
        self.failed_services = set(failed_services)
        self.nodes = []
        self.edges = []
    
    def generate(self) -> Tuple[List[Dict], List[Dict], bool]:
        """
        Generate vis.js compatible nodes and edges.
        
        Returns:
            Tuple of (nodes_list, edges_list, all_healthy_bool)
        """
        # Check if all services are healthy
        if not self.failed_services:
            return [], [], True  # All healthy - no graph needed
        
        # Find root cause services (failed services with dependents)
        root_causes = set()
        all_services = set()
        
        for root, dependents in self.dependency_map.items():
            all_services.add(root)
            for dep in dependents:
                all_services.add(dep)
            
            if root in self.failed_services:
                root_causes.add(root)
        
        # Create nodes
        for service in all_services:
            if service in self.failed_services:
                if service in root_causes:
                    # Root cause - red pulsing
                    node = {
                        "id": service,
                        "label": f"⚠ {service.upper()}",
                        "color": COLORS["error"],
                        "title": f"{service}: FAILED (Root Cause)",
                        "font": {"color": "white", "weight": "bold"},
                        "borderWidth": 3,
                    }
                else:
                    # Dependent service failed - red
                    node = {
                        "id": service,
                        "label": f"🔴 {service.upper()}",
                        "color": COLORS["error"],
                        "title": f"{service}: FAILED (Cascading)",
                        "font": {"color": "white"},
                    }
            else:
                # Healthy service - green
                node = {
                    "id": service,
                    "label": f"🟢 {service.upper()}",
                    "color": COLORS["success"],
                    "title": f"{service}: HEALTHY",
                }
            
            self.nodes.append(node)
        
        # Create edges (dependencies)
        for root, dependents in self.dependency_map.items():
            for dep in dependents:
                if dep in all_services:
                    edge_color = COLORS["error"] if root in root_causes else COLORS["info"]
                    self.edges.append({
                        "from": root,
                        "to": dep,
                        "arrows": "to",
                        "color": edge_color,
                        "width": 2 if root in root_causes else 1,
                    })
        
        return self.nodes, self.edges, False


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# HTML REPORT GENERATOR WITH ALL THREE FEATURES
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class HTMLReportGenerator:
    """
    Generate professional HTML report with:
    - Feature 1: Historical trends & comparisons
    - Feature 2: PDF export button
    - Feature 3: Interactive dependency map
    """
    
    def __init__(self, diagnostics_data: Dict, comparison: Dict = None, 
                 dependency_nodes: List = None, dependency_edges: List = None,
                 all_services_healthy: bool = False):
        """Initialize report generator."""
        self.data = diagnostics_data
        self.comparison = comparison or {}
        self.dep_nodes = dependency_nodes or []
        self.dep_edges = dependency_edges or []
        self.all_healthy = all_services_healthy
        self.hostname = get_hostname()
        self.timestamp = datetime.now()
    
    def generate(self, enable_features: Dict = None) -> str:
        """
        Generate complete HTML report.
        
        Args:
            enable_features: Dict of feature flags to include
            
        Returns:
            Complete HTML string
        """
        enable_features = enable_features or FEATURES
        
        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>InfraEye Report - {self.hostname}</title>
    <style>
        {self._get_styles()}
    </style>
</head>
<body>
    <div class="container">
        {self._get_header(enable_features)}
        {self._get_metrics_section()}
        {self._get_trend_section() if enable_features.get("historical_comparison") else ""}
        {self._get_dependency_section() if enable_features.get("dependency_map") else ""}
        {self._get_footer()}
    </div>
    
    {self._get_vis_js_script() if enable_features.get("dependency_map") and self.dep_nodes else ""}
    {self._get_print_styles()}
</body>
</html>"""
        return html
    
    def _get_header(self, enable_features: Dict) -> str:
        """Generate report header with PDF export button."""
        pdf_button = ""
        if enable_features.get("pdf_export_button"):
            pdf_button = """
            <button class="pdf-button" onclick="window.print()" title="Export as PDF (Ctrl+P)">
                🖨️ Export PDF
            </button>"""
        
        return f"""
        <header class="header">
            <div class="header-content">
                <h1>📊 InfraEye Diagnostics Report</h1>
                <div class="header-info">
                    <p><strong>Hostname:</strong> {self.hostname}</p>
                    <p><strong>Timestamp:</strong> {self.timestamp.strftime('%Y-%m-%d %H:%M:%S')}</p>
                    <p><strong>Report Version:</strong> {VERSION}</p>
                </div>
            </div>
            {pdf_button}
        </header>"""
    
    def _get_metrics_section(self) -> str:
        """Generate metrics overview section."""
        cpu = self.data.get("cpu", {})
        mem = self.data.get("memory", {})
        
        cpu_status = self._get_status_badge(cpu.get("status", "unknown"))
        mem_status = self._get_status_badge(mem.get("status", "unknown"))
        
        disk_html = self._get_disk_section()
        services_html = self._get_services_section()
        
        return f"""
        <section class="metrics">
            <h2>📈 System Metrics</h2>
            
            <div class="metrics-grid">
                <div class="metric-card">
                    <h3>CPU Usage</h3>
                    <div class="metric-value">{cpu.get("percent", 0):.1f}%</div>
                    {cpu_status}
                    <p class="metric-detail">Cores: {cpu.get("cores", 0)} | Logical: {cpu.get("logical_cores", 0)}</p>
                </div>
                
                <div class="metric-card">
                    <h3>RAM Usage</h3>
                    <div class="metric-value">{mem.get("percent", 0):.1f}%</div>
                    {mem_status}
                    <p class="metric-detail">{mem.get("used_gb", 0):.1f}GB / {mem.get("total_gb", 0):.1f}GB</p>
                    {self._get_memory_alert(mem)}
                </div>
            </div>
            
            {disk_html}
            {services_html}
        </section>"""
    
    def _get_disk_section(self) -> str:
        """Generate disk usage section."""
        disks = self.data.get("disk", {})
        if not disks or "error" in disks:
            return "<p class='error'>Disk information unavailable</p>"
        
        disk_cards = ""
        for device, info in disks.items():
            if isinstance(info, dict) and "percent" in info:
                status = self._get_status_badge(info.get("status", "unknown"))
                disk_cards += f"""
                <div class="metric-card">
                    <h3>{device}</h3>
                    <div class="metric-value">{info.get("percent", 0):.1f}%</div>
                    {status}
                    <p class="metric-detail">{info.get("used_gb", 0):.1f}GB / {info.get("total_gb", 0):.1f}GB</p>
                </div>"""
        
        return f"<div class='metrics-grid'>{disk_cards}</div>" if disk_cards else ""
    
    def _get_services_section(self) -> str:
        """Generate services status section."""
        services = self.data.get("services", {})
        failed = services.get("failed_services", [])
        
        if not failed:
            return "<div class='success-banner'>✅ All Services Healthy</div>"
        
        failed_html = "".join([
            f"<li class='failed-service'>🔴 {service}</li>"
            for service in failed
        ])
        
        return f"""
        <div class='services-section'>
            <h3>⚠️ Failed Services ({len(failed)})</h3>
            <ul>{failed_html}</ul>
        </div>"""
    
    def _get_trend_section(self) -> str:
        """Generate historical trend comparison section (Feature 1)."""
        if self.comparison.get("is_baseline"):
            return f"""
            <section class="trends">
                <h2>📊 Historical Comparison</h2>
                <div class="baseline-notice">
                    ℹ️ {self.comparison.get("message", "Baseline Run — No Previous Data")}
                </div>
            </section>"""
        
        metrics_html = ""
        for metric_name, metric_data in self.comparison.get("metrics", {}).items():
            if metric_data:
                trend = metric_data.get("trend", "●")
                color = metric_data.get("color", COLORS["neutral"])
                current = metric_data.get("current", "N/A")
                previous = metric_data.get("previous", "N/A")
                
                pretty_name = metric_name.replace("_", " ").title()
                metrics_html += f"""
                <div class="trend-metric">
                    <h4>{pretty_name}</h4>
                    <span class="current">Current: <strong>{current}</strong></span>
                    <span class="previous">Previous: <strong style="color: #999;">{previous}</strong></span>
                    <span class="trend-arrow" style="color: {color}; font-size: 1.5em;">{trend}</span>
                </div>"""
        
        # New and resolved issues
        new_issues_html = ""
        if self.comparison.get("new_issues"):
            new_issues_html = """
            <div class="issues-section">
                <h4>🆕 New Issues (First Seen)</h4>
                <ul class="issues-list">"""
            for issue in self.comparison.get("new_issues", []):
                new_issues_html += f"<li class='new-issue'>🔴 {issue}</li>"
            new_issues_html += "</ul></div>"
        
        resolved_html = ""
        if self.comparison.get("resolved_issues"):
            resolved_html = """
            <div class="issues-section">
                <h4>✓ Resolved (Since Last Run)</h4>
                <ul class="issues-list">"""
            for issue in self.comparison.get("resolved_issues", []):
                resolved_html += f"<li class='resolved-issue'>🟢 {issue}</li>"
            resolved_html += "</ul></div>"
        
        return f"""
        <section class="trends">
            <h2>📊 Historical Comparison & Trends</h2>
            <div class="trends-grid">
                {metrics_html}
            </div>
            {new_issues_html}
            {resolved_html}
        </section>"""
    
    def _get_dependency_section(self) -> str:
        """Generate service dependency map section (Feature 3)."""
        if self.all_healthy:
            return f"""
            <section class="dependency-section">
                <h2>🔗 Service Dependencies</h2>
                <div class="success-banner">✅ No Dependency Failures Detected</div>
            </section>"""
        
        if not self.dep_nodes:
            return ""
        
        return f"""
        <section class="dependency-section">
            <h2>🔗 Service Dependency Map</h2>
            <p class="section-help">Click on nodes to see details. Red arrows indicate failed service roots.</p>
            <div id="dependency-network" style="width: 100%; height: 500px; border: 1px solid #ddd; border-radius: 8px;"></div>
            <div class="dependency-legend">
                <h4>Legend</h4>
                <ul>
                    <li><span style="color: {COLORS['error']};">🔴 Red Node</span> = Service FAILED</li>
                    <li><span style="color: {COLORS['success']};">🟢 Green Node</span> = Service HEALTHY</li>
                    <li><span style="color: {COLORS['warning']};">🟡 Yellow Node</span> = Service DEGRADED</li>
                    <li><span style="color: {COLORS['neutral']};">⚫ Grey Node</span> = Service NOT DETECTED</li>
                    <li>Arrows show dependency direction (root → dependents)</li>
                </ul>
            </div>
        </section>"""
    
    def _get_footer(self) -> str:
        """Generate report footer."""
        return f"""
        <footer class="footer">
            <p>InfraEye v{VERSION} | Generated {self.timestamp.strftime('%Y-%m-%d %H:%M:%S')} | 
            <a href="https://tushargudde.tech">Tushar Gudde</a></p>
            <p class="confidential" style="display: none;">Confidential — IT Internal Use</p>
        </footer>"""
    
    def _get_status_badge(self, status: str) -> str:
        """Get HTML badge for status."""
        badges = {
            "critical": "<span class='badge badge-critical'>⚠️ CRITICAL</span>",
            "warning": "<span class='badge badge-warning'>⚠️ WARNING</span>",
            "healthy": "<span class='badge badge-success'>✅ HEALTHY</span>",
            "normal": "<span class='badge badge-success'>✅ OK</span>",
            "unknown": "<span class='badge badge-neutral'>❓ UNKNOWN</span>",
        }
        return badges.get(status, badges["unknown"])
    
    def _get_memory_alert(self, mem: Dict) -> str:
        """Get memory upgrade alert if applicable."""
        if mem.get("upgrade_suggested"):
            return f"<p class='alert'>{mem['upgrade_suggestion']}</p>"
        return ""
    
    def _get_styles(self) -> str:
        """CSS styling for the report."""
        return """
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: #333;
            line-height: 1.6;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }
        
        .header {
            background: white;
            border-radius: 12px;
            padding: 30px;
            margin-bottom: 30px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            position: relative;
        }
        
        .header h1 {
            font-size: 2.5em;
            color: #667eea;
            margin-bottom: 15px;
        }
        
        .header-info {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 20px;
            margin-top: 15px;
        }
        
        .header-info p {
            color: #666;
            font-size: 0.95em;
        }
        
        .pdf-button {
            position: absolute;
            top: 30px;
            right: 30px;
            background: #667eea;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 0.95em;
            transition: all 0.3s;
        }
        
        .pdf-button:hover {
            background: #764ba2;
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        }
        
        section {
            background: white;
            border-radius: 12px;
            padding: 30px;
            margin-bottom: 30px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        
        section h2 {
            color: #667eea;
            font-size: 1.8em;
            margin-bottom: 25px;
            border-bottom: 2px solid #667eea;
            padding-bottom: 10px;
        }
        
        .metrics-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 25px;
            margin-bottom: 25px;
        }
        
        .metric-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 25px;
            border-radius: 10px;
            text-align: center;
            transition: transform 0.3s;
        }
        
        .metric-card:hover {
            transform: translateY(-5px);
        }
        
        .metric-card h3 {
            font-size: 0.95em;
            margin-bottom: 15px;
            opacity: 0.9;
        }
        
        .metric-value {
            font-size: 2.5em;
            font-weight: bold;
            margin: 10px 0;
        }
        
        .metric-detail {
            font-size: 0.85em;
            opacity: 0.85;
            margin-top: 10px;
        }
        
        .badge {
            display: inline-block;
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 0.85em;
            font-weight: bold;
            margin-top: 10px;
        }
        
        .badge-success {
            background: #4CAF50;
            color: white;
        }
        
        .badge-warning {
            background: #FF9800;
            color: white;
        }
        
        .badge-critical {
            background: #F44336;
            color: white;
        }
        
        .badge-neutral {
            background: #9E9E9E;
            color: white;
        }
        
        .success-banner {
            background: #E8F5E9;
            border-left: 4px solid #4CAF50;
            padding: 15px;
            border-radius: 6px;
            color: #2E7D32;
            margin: 20px 0;
        }
        
        .baseline-notice {
            background: #E3F2FD;
            border-left: 4px solid #2196F3;
            padding: 15px;
            border-radius: 6px;
            color: #1565C0;
            margin: 20px 0;
        }
        
        .services-section {
            margin-top: 25px;
            padding: 20px;
            background: #FFF3E0;
            border-radius: 8px;
            border-left: 4px solid #FF9800;
        }
        
        .services-section h3 {
            color: #E65100;
            margin-bottom: 15px;
        }
        
        .failed-service {
            color: #C62828;
            margin: 8px 0;
            padding: 8px 12px;
            background: white;
            border-radius: 4px;
        }
        
        .trends-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 25px;
        }
        
        .trend-metric {
            background: #F5F5F5;
            padding: 15px;
            border-radius: 8px;
            border-left: 4px solid #667eea;
        }
        
        .trend-metric h4 {
            color: #667eea;
            margin-bottom: 12px;
        }
        
        .current, .previous {
            display: block;
            font-size: 0.9em;
            margin: 5px 0;
        }
        
        .previous {
            color: #999;
        }
        
        .trend-arrow {
            display: block;
            margin-top: 10px;
            font-weight: bold;
        }
        
        .issues-section {
            margin-top: 20px;
            padding: 15px;
            background: #F5F5F5;
            border-radius: 8px;
        }
        
        .issues-section h4 {
            color: #667eea;
            margin-bottom: 10px;
        }
        
        .issues-list {
            list-style: none;
            padding-left: 20px;
        }
        
        .new-issue {
            color: #C62828;
            margin: 8px 0;
            font-weight: 500;
        }
        
        .resolved-issue {
            color: #2E7D32;
            margin: 8px 0;
            font-weight: 500;
        }
        
        .dependency-section {
            margin-top: 30px;
        }
        
        .dependency-legend {
            margin-top: 25px;
            padding: 15px;
            background: #F5F5F5;
            border-radius: 8px;
        }
        
        .dependency-legend h4 {
            color: #667eea;
            margin-bottom: 10px;
        }
        
        .dependency-legend ul {
            list-style: none;
            padding-left: 20px;
        }
        
        .dependency-legend li {
            margin: 8px 0;
            color: #666;
        }
        
        .section-help {
            color: #999;
            font-size: 0.9em;
            margin-bottom: 15px;
            font-style: italic;
        }
        
        .footer {
            text-align: center;
            padding: 20px;
            color: #999;
            border-top: 1px solid #eee;
            margin-top: 40px;
            font-size: 0.85em;
        }
        
        .footer a {
            color: #667eea;
            text-decoration: none;
        }
        
        .alert {
            background: #FFF3E0;
            color: #E65100;
            padding: 10px;
            border-radius: 4px;
            font-size: 0.9em;
            margin-top: 8px;
        }
        
        @media (max-width: 768px) {
            .header {
                padding: 20px;
            }
            
            .pdf-button {
                position: static;
                margin-top: 15px;
                width: 100%;
            }
            
            .metrics-grid {
                grid-template-columns: 1fr;
            }
            
            .metric-value {
                font-size: 2em;
            }
        }
        """
    
    def _get_vis_js_script(self) -> str:
        """
        Generate vis.js network graph initialization.
        vis.js is loaded from CDN for Feature 3.
        """
        if not self.dep_nodes or not self.dep_edges:
            return ""
        
        nodes_json = json.dumps(self.dep_nodes)
        edges_json = json.dumps(self.dep_edges)
        
        return f"""
        <script src="https://cdnjs.cloudflare.com/ajax/libs/vis/4.21.0/vis.min.js"></script>
        <link href="https://cdnjs.cloudflare.com/ajax/libs/vis/4.21.0/vis.min.css" rel="stylesheet">
        
        <script>
        document.addEventListener('DOMContentLoaded', function() {{
            let nodes = new vis.DataSet({nodes_json});
            let edges = new vis.DataSet({edges_json});
            
            let container = document.getElementById('dependency-network');
            let data = {{nodes: nodes, edges: edges}};
            
            let options = {{
                physics: {{
                    enabled: true,
                    solver: 'hierarchicalRepulsive',
                    hierarchicalRepulsive: {{
                        centralGravity: 0.0,
                        springLength: 200,
                        springConstant: 0.04
                    }}
                }},
                nodes: {{
                    shape: 'dot',
                    scaling: {{
                        min: 30,
                        max: 30
                    }}
                }},
                edges: {{
                    smooth: {{
                        type: 'continuous'
                    }},
                    arrows: {{
                        to: {{enabled: true, scaleFactor: 0.5}}
                    }}
                }},
                interaction: {{
                    navigationButtons: true,
                    keyboard: true,
                    zoomView: true
                }}
            }};
            
            let network = new vis.Network(container, data, options);
            
            // Click handler for node details
            network.on('click', function(params) {{
                if (params.nodes.length > 0) {{
                    let nodeId = params.nodes[0];
                    let nodeData = nodes.get(nodeId);
                    alert(nodeData.title);
                }}
            }});
        }});
        </script>"""
    
    def _get_print_styles(self) -> str:
        """CSS for @media print to optimize PDF export (Feature 2)."""
        return """
        <style media="print">
            @page {
                size: A4;
                margin: 20mm;
                @bottom-center {
                    content: "Confidential — IT Internal Use | Page " counter(page) " of " counter(pages);
                    font-size: 10pt;
                    color: #999;
                }
            }
            
            body {
                background: white !important;
            }
            
            .container {
                padding: 0;
                margin: 0;
            }
            
            .pdf-button {
                display: none !important;
            }
            
            section {
                page-break-inside: avoid;
                box-shadow: none !important;
                border: 1px solid #ddd;
                margin-bottom: 20px;
            }
            
            .header {
                border-bottom: 2px solid #667eea;
                margin-bottom: 20px;
            }
            
            .metric-card {
                background: white !important;
                color: black !important;
                border: 1px solid #ddd;
            }
            
            .footer {
                page-break-before: avoid;
                margin-top: 40px;
                border-top: 1px solid #ddd;
                padding-top: 20px;
            }
            
            .confidential {
                display: block !important;
            }
            
            #dependency-network {
                display: none !important;
            }
            
            table {
                page-break-inside: avoid;
                width: 100%;
            }
            
            tr {
                page-break-inside: avoid;
            }
        </style>"""


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# PDF EXPORT HANDLER (Feature 2 - Optional)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class PDFExporter:
    """Handle PDF export via weasyprint (optional dependency)."""
    
    @staticmethod
    def export(html_content: str, output_path: str) -> bool:
        """
        Export HTML to PDF using weasyprint.
        
        Args:
            html_content: HTML string to convert
            output_path: Path to save PDF
            
        Returns:
            True if successful, False otherwise
        """
        try:
            from weasyprint import HTML, CSS
            
            try:
                HTML(string=html_content).write_pdf(output_path)
                log(f"PDF exported successfully: {output_path}", "SUCCESS")
                return True
            except Exception as e:
                log(f"PDF generation failed: {e}", "ERROR")
                return False
        except ImportError:
            log("weasyprint not installed. Install with: pip install weasyprint", "WARN")
            return False


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# MAIN ORCHESTRATOR
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def main():
    """Main entry point for InfraEye diagnostics."""
    
    parser = argparse.ArgumentParser(
        description="InfraEye IT Diagnostics & Infrastructure Intelligence",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python infra_eye_diagnostics.py                 # Run full diagnostics
  python infra_eye_diagnostics.py --no-history    # Disable historical tracking
  python infra_eye_diagnostics.py --export-pdf    # Export to PDF (requires weasyprint)
  python infra_eye_diagnostics.py --no-depmap     # Disable dependency visualization
        """
    )
    
    parser.add_argument("--version", action="version", version=f"%(prog)s {VERSION}")
    parser.add_argument("--output", "-o", default="Reports", help="Output directory for reports")
    parser.add_argument("--no-history", action="store_true", help="Disable historical comparison feature")
    parser.add_argument("--no-pdf-button", action="store_true", help="Disable PDF export button")
    parser.add_argument("--no-depmap", action="store_true", help="Disable dependency map")
    parser.add_argument("--export-pdf", action="store_true", help="Auto-export to PDF (requires weasyprint)")
    
    args = parser.parse_args()
    
    # Override feature flags from CLI arguments
    features = FEATURES.copy()
    features["historical_comparison"] = not args.no_history
    features["pdf_export_button"] = not args.no_pdf_button
    features["dependency_map"] = not args.no_depmap
    
    log(f"InfraEye v{VERSION} - Starting diagnostics...", "INFO")
    
    # Ensure output directory
    output_dir = ensure_directory(args.output)
    
    # 1. COLLECT SYSTEM DATA
    diagnostics = SystemDiagnostics()
    system_data = diagnostics.collect_all()
    diag_metrics = diagnostics.get_metric_values()
    
    # 2. FEATURE 1: HISTORICAL SNAPSHOTS
    comparison_data = {}
    if features["historical_comparison"]:
        history = HistoryManager()
        history.current_snapshot = diag_metrics
        history.snapshot_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        history.save_snapshot(diag_metrics)
        history.load_previous_snapshot()
        comparison_data = history.get_comparison()
        history.cleanup_old_snapshots()
    
    # 3. FEATURE 3: DEPENDENCY MAP
    dep_nodes, dep_edges = [], []
    all_healthy = True
    if features["dependency_map"]:
        failed_services = diag_metrics.get("failed_services", [])
        dep_gen = DependencyMapGenerator(DEPENDENCY_MAP, failed_services)
        dep_nodes, dep_edges, all_healthy = dep_gen.generate()
    
    # 4. GENERATE HTML REPORT
    report_gen = HTMLReportGenerator(
        diagnostics_data=system_data,
        comparison=comparison_data,
        dependency_nodes=dep_nodes,
        dependency_edges=dep_edges,
        all_services_healthy=all_healthy
    )
    
    html_content = report_gen.generate(enable_features=features)
    
    # Save HTML report
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    hostname = get_hostname()
    html_file = output_dir / f"InfraEye_Report_{hostname}_{timestamp}.html"
    
    try:
        with open(html_file, "w") as f:
            f.write(html_content)
        log(f"HTML report saved: {html_file}", "SUCCESS")
    except Exception as e:
        log(f"Failed to save HTML report: {e}", "ERROR")
        return 1
    
    # 5. FEATURE 2: PDF EXPORT (if requested)
    if args.export_pdf:
        pdf_file = output_dir / f"InfraEye_Report_{hostname}_{timestamp.split('_')[0]}.pdf"
        if PDFExporter.export(html_content, str(pdf_file)):
            log(f"PDF report saved: {pdf_file}", "SUCCESS")
        else:
            log("PDF export failed. HTML report is still available.", "WARN")
    
    # Summary
    log("─" * 80, "INFO")
    log(f"✅ Diagnostics complete!", "SUCCESS")
    log(f"📄 Report: {html_file}", "INFO")
    log(f"🔄 Features enabled:", "INFO")
    for feature, enabled in features.items():
        status = "✓" if enabled else "✗"
        log(f"   {status} {feature.replace('_', ' ').title()}", "INFO")
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
