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
    2.0
#>

param(
    [switch]$RunCleanup,
    [switch]$ForceCleanup
)

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
$HistoryFile  = Join-Path $LogDir "DeviceHealth_History.jsonl"

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

function Convert-ToDouble {
    param(
        $Value,
        [double]$Default = 0
    )
    try {
        if ($null -eq $Value) { return $Default }
        if ($Value -is [string]) {
            $clean = ($Value -replace '[^0-9\.-]', '')
            if ([string]::IsNullOrWhiteSpace($clean)) { return $Default }
            return [double]$clean
        }
        return [double]$Value
    } catch {
        return $Default
    }
}

function Merge-ObjectArrays {
    param(
        $Primary,
        $Secondary
    )
    return @($Primary) + @($Secondary)
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

function Get-DriveTransferProfile {
    param(
        [string]$BusType,
        [string]$MediaType,
        [string]$Model
    )

    $bus = ([string]$BusType).ToUpper()
    $media = ([string]$MediaType).ToUpper()
    $modelText = ([string]$Model).ToUpper()

    if ($bus -eq "NVME" -or $modelText -match "NVME") {
        $gen = "NVMe"
        $maxMBps = 3500
        if ($modelText -match "GEN5|PCIE\s*5|PCIe\s*5") {
            $gen = "NVMe Gen5"
            $maxMBps = 14000
        } elseif ($modelText -match "GEN4|PCIE\s*4|PCIe\s*4") {
            $gen = "NVMe Gen4"
            $maxMBps = 7000
        } elseif ($modelText -match "GEN3|PCIE\s*3|PCIe\s*3") {
            $gen = "NVMe Gen3"
            $maxMBps = 3500
        }
        return [PSCustomObject]@{ Generation = $gen; MaxThroughput = "~$maxMBps MB/s (bus/theoretical)" }
    }

    if ($bus -eq "SATA" -or $modelText -match "SATA") {
        return [PSCustomObject]@{ Generation = "SATA III"; MaxThroughput = "~600 MB/s (bus/theoretical)" }
    }

    if ($bus -eq "SAS") {
        return [PSCustomObject]@{ Generation = "SAS"; MaxThroughput = "~1200 MB/s (bus/theoretical)" }
    }

    if ($bus -eq "USB") {
        return [PSCustomObject]@{ Generation = "USB Storage"; MaxThroughput = "Variable by USB generation" }
    }

    if ($media -eq "HDD") {
        return [PSCustomObject]@{ Generation = "HDD"; MaxThroughput = "~80-220 MB/s (typical sequential)" }
    }

    return [PSCustomObject]@{ Generation = if ($bus) { $bus } else { "Unknown" }; MaxThroughput = "Unavailable" }
}

function Get-DiskHardwareSummary {
    param(
        [string]$Vendor,
        [string]$Model,
        [string]$Generation,
        [string]$Serial,
        [string]$PartNumber,
        [string]$Fru,
        [string]$BusType,
        [string]$TransferSpeed
    )

    $v = if ($Vendor) { $Vendor } else { "Unknown" }
    $m = if ($Model) { $Model } else { "Unknown" }
    $g = if ($Generation) { $Generation } else { "Unknown" }
    $sn = if ($Serial) { $Serial } else { "N/A" }
    $pn = if ($PartNumber) { $PartNumber } else { "N/A" }
    $fruVal = if ($Fru) { $Fru } else { "N/A" }
    $bus = if ($BusType) { $BusType } else { "Unknown" }
    $speed = if ($TransferSpeed) { $TransferSpeed } else { "Unavailable" }

    $typeIdentity = if ($v -eq "Unknown" -and $m -ne "Unknown") { $m } elseif ($v -eq "Unknown" -and $m -eq "Unknown") { "Unknown Device" } else { "$v $m" }
    return "Type: $typeIdentity | Gen: $g | Bus: $bus | SN: $sn | PN: $pn | FRU: $fruVal | Data Rate: $speed"
}

function Get-DiskInfo {
    Write-Log "Collecting disk data..." "INFO"
    $disks = @()
    try {
        $physicalDiskTypeByNumber = @{}
        $hardwareByNumber = @{}
        try {
            $physicalDisks = @(Get-PhysicalDisk -ErrorAction Stop)
            foreach ($physicalDisk in $physicalDisks) {
                $diskMatch = Get-Disk -FriendlyName $physicalDisk.FriendlyName -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($diskMatch) {
                    $physicalDiskTypeByNumber[$diskMatch.Number] = if ($physicalDisk.MediaType) {
                        [string]$physicalDisk.MediaType
                    } elseif ($physicalDisk.SpindleSpeed -eq 0 -or $physicalDisk.BusType -match "NVMe") {
                        "SSD"
                    } else {
                        "Unknown"
                    }

                    $transferProfile = Get-DriveTransferProfile -BusType ([string]$diskMatch.BusType) -MediaType ([string]$physicalDisk.MediaType) -Model ([string]$physicalDisk.Model)
                    $hardwareByNumber[$diskMatch.Number] = [PSCustomObject]@{
                        Vendor       = if ($physicalDisk.Manufacturer) { [string]$physicalDisk.Manufacturer } else { "Unknown" }
                        Model        = if ($physicalDisk.Model) { [string]$physicalDisk.Model } else { [string]$physicalDisk.FriendlyName }
                        Serial       = if ($physicalDisk.SerialNumber) { [string]$physicalDisk.SerialNumber } else { $null }
                        PartNumber   = if ($physicalDisk.PartNumber) { [string]$physicalDisk.PartNumber } else { $null }
                        Fru          = if ($physicalDisk.FruId) { [string]$physicalDisk.FruId } else { $null }
                        BusType      = [string]$diskMatch.BusType
                        Generation   = $transferProfile.Generation
                        TransferRate = $transferProfile.MaxThroughput
                    }
                }
            }
        } catch {
            Write-Log "Modern physical disk metadata unavailable. Falling back to legacy disk detection." "WARN"
        }

        try {
            $legacyDiskInfo = @(Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction Stop)
            foreach ($legacyDisk in $legacyDiskInfo) {
                $legacyNumber = $null
                try { $legacyNumber = [int]$legacyDisk.Index } catch { $legacyNumber = $null }
                if ($null -eq $legacyNumber -or $hardwareByNumber.ContainsKey($legacyNumber)) { continue }

                $legacyTransfer = Get-DriveTransferProfile -BusType ([string]$legacyDisk.InterfaceType) -MediaType "" -Model ([string]$legacyDisk.Model)
                $hardwareByNumber[$legacyNumber] = [PSCustomObject]@{
                    Vendor       = if ($legacyDisk.Manufacturer) { [string]$legacyDisk.Manufacturer } else { "Unknown" }
                    Model        = [string]$legacyDisk.Model
                    Serial       = if ($legacyDisk.SerialNumber) { ([string]$legacyDisk.SerialNumber).Trim() } else { $null }
                    PartNumber   = $null
                    Fru          = $null
                    BusType      = [string]$legacyDisk.InterfaceType
                    Generation   = $legacyTransfer.Generation
                    TransferRate = $legacyTransfer.MaxThroughput
                }
            }
        } catch {
            Write-Log "Legacy disk hardware metadata unavailable: $_" "WARN"
        }

        $logicalDisks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"
        foreach ($disk in $logicalDisks) {
            $totalGB = [math]::Round($disk.Size / 1GB, 2)
            $freeGB  = [math]::Round($disk.FreeSpace / 1GB, 2)
            $usedGB  = [math]::Round($totalGB - $freeGB, 2)
            $freePct = if ($totalGB -gt 0) { [math]::Round(($freeGB / $totalGB) * 100, 1) } else { 0 }

            # Detect SSD/HDD using modern storage APIs first, then fall back to legacy WMI.
            $diskType = "Unknown"
            try {
                $driveLetter = $disk.DeviceID.TrimEnd(':')
                $partitionInfo = Get-Partition -DriveLetter $driveLetter -ErrorAction SilentlyContinue
                if ($partitionInfo) {
                    $diskInfo = $partitionInfo | Get-Disk -ErrorAction SilentlyContinue | Select-Object -First 1
                    if ($diskInfo) {
                        if ($physicalDiskTypeByNumber.ContainsKey($diskInfo.Number)) {
                            $diskType = $physicalDiskTypeByNumber[$diskInfo.Number]
                        } elseif ($diskInfo.BusType -match "NVMe") {
                            $diskType = "SSD"
                        } elseif ($diskInfo.MediaType) {
                            $diskType = [string]$diskInfo.MediaType
                        }
                    }
                }

                if ($diskType -eq "Unspecified") {
                    $diskType = "Unknown"
                }

                if ($diskType -eq "Unknown") {
                    $partition = Get-CimInstance -Query "ASSOCIATORS OF {Win32_LogicalDisk.DeviceID='$($disk.DeviceID)'} WHERE AssocClass=Win32_LogicalDiskToPartition"
                    if ($partition) {
                        $drive = Get-CimInstance -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='$($partition[0].DeviceID)'} WHERE AssocClass=Win32_DiskDriveToDiskPartition"
                        if ($drive) {
                            $model = [string]$drive[0].Model
                            if ($model -match "SSD|NVMe|M\.2") {
                                $diskType = "SSD"
                            } elseif ($drive[0].MediaType -match "Fixed|HDD|Hard") {
                                $diskType = "HDD"
                            }
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

            $diskNumber = $null
            $busTypeForDisk = ""
            try {
                $driveLetterForSummary = $disk.DeviceID.TrimEnd(':')
                $partitionInfoForSummary = Get-Partition -DriveLetter $driveLetterForSummary -ErrorAction SilentlyContinue
                if ($partitionInfoForSummary) {
                    $diskInfoForSummary = $partitionInfoForSummary | Get-Disk -ErrorAction SilentlyContinue | Select-Object -First 1
                    if ($diskInfoForSummary) {
                        $diskNumber = $diskInfoForSummary.Number
                        $busTypeForDisk = [string]$diskInfoForSummary.BusType
                    }
                }
            } catch { }

            $hw = if ($null -ne $diskNumber -and $hardwareByNumber.ContainsKey($diskNumber)) { $hardwareByNumber[$diskNumber] } else { $null }
            $transferProfileFallback = Get-DriveTransferProfile -BusType $busTypeForDisk -MediaType $diskType -Model ""
            $hardwareSummary = Get-DiskHardwareSummary `
                -Vendor $(if ($hw) { $hw.Vendor } else { "Unknown" }) `
                -Model $(if ($hw) { $hw.Model } else { "Unknown" }) `
                -Generation $(if ($hw) { $hw.Generation } else { $transferProfileFallback.Generation }) `
                -Serial $(if ($hw) { $hw.Serial } else { $null }) `
                -PartNumber $(if ($hw) { $hw.PartNumber } else { $null }) `
                -Fru $(if ($hw) { $hw.Fru } else { $null }) `
                -BusType $(if ($hw -and $hw.BusType) { $hw.BusType } else { $busTypeForDisk }) `
                -TransferSpeed $(if ($hw -and $hw.TransferRate) { $hw.TransferRate } else { $transferProfileFallback.MaxThroughput })

            $disks += [PSCustomObject]@{
                Drive          = $disk.DeviceID
                TotalGB        = $totalGB
                UsedGB         = $usedGB
                FreeGB         = $freeGB
                FreePercent    = $freePct
                DiskType       = $diskType
                HardwareSummary = $hardwareSummary
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

        # Battery report based health metrics (Design vs Full charge + Cycle count)
        $batteryHealth = [PSCustomObject]@{
            DesignCapacity_mWh     = $null
            FullChargeCapacity_mWh = $null
            CycleCount             = $null
            HealthPercent          = $null
            HealthStatus           = "Unknown"
            Recommendation         = "Battery data unavailable on this system."
            BatteryReportPath      = $null
            BatteryName            = $null
            BatteryManufacturer    = $null
            BatterySerial          = $null
            BatteryChemistry       = $null
        }

        try {
            function Get-BatteryTextField {
                param([string]$Raw, [string[]]$Patterns)
                foreach ($pattern in $Patterns) {
                    $m = [regex]::Match($Raw, $pattern)
                    if ($m.Success) {
                        $val = $m.Groups[1].Value.Trim()
                        if (-not [string]::IsNullOrWhiteSpace($val)) { return $val }
                    }
                }
                return $null
            }

            function Get-BatteryMetricFromReport {
                param(
                    [string]$Raw,
                    [string[]]$Patterns
                )

                foreach ($pattern in $Patterns) {
                    $m = [regex]::Match($Raw, $pattern)
                    if ($m.Success) {
                        $valueText = ($m.Groups[1].Value -replace '[^0-9]', '')
                        if (-not [string]::IsNullOrWhiteSpace($valueText)) {
                            return [int64]$valueText
                        }
                    }
                }
                return $null
            }

            $batteryReportPath = Join-Path $ReportDir ("BatteryReport_{0}.html" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
            $null = powercfg /batteryreport /output "$batteryReportPath" 2>$null
            if (Test-Path $batteryReportPath) {
                $batteryHealth.BatteryReportPath = $batteryReportPath
                $raw = Get-Content -Path $batteryReportPath -Raw -ErrorAction SilentlyContinue
                if (-not [string]::IsNullOrWhiteSpace($raw)) {
                    $normalizedRaw = ($raw -replace '&nbsp;', ' ' -replace "`r", ' ' -replace "`n", ' ')

                    $batteryHealth.DesignCapacity_mWh = Get-BatteryMetricFromReport -Raw $normalizedRaw -Patterns @(
                        '(?is)DESIGN\s*CAPACITY(?:\s*</span>)?\s*</td>\s*<td[^>]*>\s*([0-9][0-9,]*)\s*mWh',
                        '(?is)DESIGN\s*CAPACITY[^0-9]{0,160}([0-9][0-9,]*)\s*mWh'
                    )

                    $batteryHealth.FullChargeCapacity_mWh = Get-BatteryMetricFromReport -Raw $normalizedRaw -Patterns @(
                        '(?is)FULL\s*CHARGE\s*CAPACITY(?:\s*</span>)?\s*</td>\s*<td[^>]*>\s*([0-9][0-9,]*)\s*mWh',
                        '(?is)FULL\s*CHARGE\s*CAPACITY[^0-9]{0,160}([0-9][0-9,]*)\s*mWh'
                    )

                    $cycleValue = Get-BatteryMetricFromReport -Raw $normalizedRaw -Patterns @(
                        '(?is)CYCLE\s*COUNT(?:\s*</span>)?\s*</td>\s*<td[^>]*>\s*([0-9][0-9,]*)',
                        '(?is)CYCLE\s*COUNT[^0-9]{0,120}([0-9][0-9,]*)'
                    )
                    if ($null -ne $cycleValue) {
                        $batteryHealth.CycleCount = [int]$cycleValue
                    }

                    # Battery identification fields
                    $batteryHealth.BatteryName = Get-BatteryTextField -Raw $normalizedRaw -Patterns @(
                        '(?is)<span[^>]*class="label"[^>]*>NAME</span>\s*</td>\s*<td[^>]*>\s*([^<]+)',
                        '(?is)>NAME<[/\w ]*>\s*</td>\s*<td[^>]*>([^<]+)'
                    )
                    $batteryHealth.BatteryManufacturer = Get-BatteryTextField -Raw $normalizedRaw -Patterns @(
                        '(?is)<span[^>]*class="label"[^>]*>MANUFACTURER</span>\s*</td>\s*<td[^>]*>\s*([^<]+)',
                        '(?is)>MANUFACTURER<[/\w ]*>\s*</td>\s*<td[^>]*>([^<]+)'
                    )
                    $batteryHealth.BatterySerial = Get-BatteryTextField -Raw $normalizedRaw -Patterns @(
                        '(?is)<span[^>]*class="label"[^>]*>SERIAL\s*NUMBER</span>\s*</td>\s*<td[^>]*>\s*([^<]+)',
                        '(?is)>SERIAL\s*NUMBER<[/\w ]*>\s*</td>\s*<td[^>]*>([^<]+)'
                    )
                    $batteryHealth.BatteryChemistry = Get-BatteryTextField -Raw $normalizedRaw -Patterns @(
                        '(?is)<span[^>]*class="label"[^>]*>CHEMISTRY</span>\s*</td>\s*<td[^>]*>\s*([^<]+)',
                        '(?is)>CHEMISTRY<[/\w ]*>\s*</td>\s*<td[^>]*>([^<]+)'
                    )
                }
            }

            # Fallback: WMI battery classes when report parsing is unavailable or incomplete.
            if ($null -eq $batteryHealth.DesignCapacity_mWh -or $batteryHealth.DesignCapacity_mWh -le 0) {
                $staticData = Get-CimInstance -Namespace root\wmi -ClassName BatteryStaticData -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($staticData) {
                    $designedCapProp = $staticData.PSObject.Properties['DesignedCapacity']
                    if ($designedCapProp -and $null -ne $designedCapProp.Value) {
                        $batteryHealth.DesignCapacity_mWh = [int64](Convert-ToDouble $designedCapProp.Value 0)
                    }
                }
            }

            if ($null -eq $batteryHealth.FullChargeCapacity_mWh -or $batteryHealth.FullChargeCapacity_mWh -le 0) {
                $fullChargeData = Get-CimInstance -Namespace root\wmi -ClassName BatteryFullChargedCapacity -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($fullChargeData) {
                    $fullCapProp = $fullChargeData.PSObject.Properties['FullChargedCapacity']
                    if ($fullCapProp -and $null -ne $fullCapProp.Value) {
                        $batteryHealth.FullChargeCapacity_mWh = [int64](Convert-ToDouble $fullCapProp.Value 0)
                    }
                }
            }

            if ($null -eq $batteryHealth.CycleCount) {
                $cycleData = Get-CimInstance -Namespace root\wmi -ClassName BatteryCycleCount -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($cycleData) {
                    $cycleProp = $cycleData.PSObject.Properties['CycleCount']
                    if ($cycleProp -and $null -ne $cycleProp.Value) {
                        $batteryHealth.CycleCount = [int](Convert-ToDouble $cycleProp.Value 0)
                    }
                }
            }

            if ($null -ne $batteryHealth.DesignCapacity_mWh -and $batteryHealth.DesignCapacity_mWh -gt 0 -and
                $null -ne $batteryHealth.FullChargeCapacity_mWh -and $batteryHealth.FullChargeCapacity_mWh -gt 0) {
                $batteryHealth.HealthPercent = [math]::Round((($batteryHealth.FullChargeCapacity_mWh / $batteryHealth.DesignCapacity_mWh) * 100), 1)

                if ($batteryHealth.HealthPercent -ge 80) {
                    $batteryHealth.HealthStatus = "Healthy"
                    $batteryHealth.Recommendation = "Battery healthy"
                } elseif ($batteryHealth.HealthPercent -ge 60) {
                    $batteryHealth.HealthStatus = "Moderate degradation"
                    $batteryHealth.Recommendation = "Moderate degradation"
                } else {
                    $batteryHealth.HealthStatus = "Replacement recommended"
                    $batteryHealth.Recommendation = "Battery replacement recommended"
                }
            }
        } catch {
            Write-Log "Battery health report parsing failed: $_" "WARN"
        }

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
                Status                  = $statusMap[[int]$battery.BatteryStatus]
                ChargePercent           = $battery.EstimatedChargeRemaining
                Present                 = $true
                DesignCapacity_mWh      = $batteryHealth.DesignCapacity_mWh
                FullChargeCapacity_mWh  = $batteryHealth.FullChargeCapacity_mWh
                CycleCount              = $batteryHealth.CycleCount
                HealthPercent           = $batteryHealth.HealthPercent
                HealthStatus            = $batteryHealth.HealthStatus
                HealthRecommendation    = $batteryHealth.Recommendation
                BatteryReportPath       = $batteryHealth.BatteryReportPath
                BatteryName             = $batteryHealth.BatteryName
                BatteryManufacturer     = $batteryHealth.BatteryManufacturer
                BatterySerial           = $batteryHealth.BatterySerial
                BatteryChemistry        = $batteryHealth.BatteryChemistry
            }
        } else {
            return [PSCustomObject]@{
                Status                  = "No Battery"
                ChargePercent           = "N/A"
                Present                 = $false
                DesignCapacity_mWh      = $null
                FullChargeCapacity_mWh  = $null
                CycleCount              = $null
                HealthPercent           = $null
                HealthStatus            = "No Battery"
                HealthRecommendation    = "Battery data unavailable on this system."
                BatteryReportPath       = $batteryHealth.BatteryReportPath
                BatteryName             = $null
                BatteryManufacturer     = $null
                BatterySerial           = $null
                BatteryChemistry        = $null
            }
        }
    } catch {
        Write-Log "Error collecting battery info: $_" "WARN"
        return [PSCustomObject]@{
            Status                  = "Unknown"
            ChargePercent           = "N/A"
            Present                 = $false
            DesignCapacity_mWh      = $null
            FullChargeCapacity_mWh  = $null
            CycleCount              = $null
            HealthPercent           = $null
            HealthStatus            = "Unknown"
            HealthRecommendation    = "Battery data unavailable on this system."
            BatteryReportPath       = $null
            BatteryName             = $null
            BatteryManufacturer     = $null
            BatterySerial           = $null
            BatteryChemistry        = $null
        }
    }
}

function Get-DirectorySizeBytes {
    param([string]$Path)
    try {
        if (-not (Test-Path $Path)) { return [int64]0 }
        $sum = Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { -not $_.PSIsContainer } |
            Measure-Object -Property Length -Sum
        return [int64](Convert-ToDouble $sum.Sum 0)
    } catch {
        return [int64]0
    }
}

function Format-Bytes {
    param([double]$Bytes)
    if ($Bytes -lt 1KB) { return ("{0:N0} B" -f $Bytes) }
    if ($Bytes -lt 1MB) { return ("{0:N2} KB" -f ($Bytes / 1KB)) }
    if ($Bytes -lt 1GB) { return ("{0:N2} MB" -f ($Bytes / 1MB)) }
    return ("{0:N2} GB" -f ($Bytes / 1GB))
}

function Invoke-SystemCleanup {
    param([switch]$Force)

    $summary = [PSCustomObject]@{
        Requested           = $true
        Performed           = $false
        TotalRecoveredBytes = [int64]0
        TotalRecoveredText  = "0 B"
        Items               = @()
    }

    $confirmed = $Force.IsPresent
    if (-not $confirmed -and [Environment]::UserInteractive) {
        $response = Read-Host "Run optional cleanup now? (Y/N)"
        $confirmed = $response -match '^(y|yes)$'
    }
    if (-not $confirmed) {
        Write-Log "Cleanup skipped by user." "INFO"
        return $summary
    }

    $targets = @(
        [PSCustomObject]@{ Name = "%TEMP%"; Path = $env:TEMP },
        [PSCustomObject]@{ Name = "%LOCALAPPDATA%\Temp"; Path = (Join-Path $env:LOCALAPPDATA "Temp") },
        [PSCustomObject]@{ Name = "C:\Windows\Temp"; Path = "C:\Windows\Temp" },
        [PSCustomObject]@{ Name = "C:\Windows\Prefetch"; Path = "C:\Windows\Prefetch" },
        [PSCustomObject]@{ Name = "C:\Windows\SoftwareDistribution\Download"; Path = "C:\Windows\SoftwareDistribution\Download" },
        [PSCustomObject]@{ Name = "C:\Windows\SoftwareDistribution\DeliveryOptimization"; Path = "C:\Windows\SoftwareDistribution\DeliveryOptimization" },
        [PSCustomObject]@{ Name = "%LOCALAPPDATA%\D3DSCache"; Path = (Join-Path $env:LOCALAPPDATA "D3DSCache") },
        [PSCustomObject]@{ Name = "C:\ProgramData\Microsoft\Windows\WER"; Path = "C:\ProgramData\Microsoft\Windows\WER" },
        [PSCustomObject]@{ Name = "Thumbnail Cache"; Path = (Join-Path $env:LOCALAPPDATA "Microsoft\Windows\Explorer") }
    )

    foreach ($target in $targets) {
        try {
            $before = Get-DirectorySizeBytes -Path $target.Path
            if (Test-Path $target.Path) {
                if ($target.Name -eq "Thumbnail Cache") {
                    Get-ChildItem -Path $target.Path -Filter "thumbcache*" -File -Force -ErrorAction SilentlyContinue |
                        Remove-Item -Force -ErrorAction SilentlyContinue
                } else {
                    Get-ChildItem -Path $target.Path -Force -ErrorAction SilentlyContinue |
                        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
            $after = Get-DirectorySizeBytes -Path $target.Path
            $recovered = [math]::Max(0, ($before - $after))
            $summary.TotalRecoveredBytes += [int64]$recovered
            $summary.Items += [PSCustomObject]@{
                Target         = $target.Name
                RecoveredBytes = [int64]$recovered
                RecoveredText  = (Format-Bytes -Bytes $recovered)
                Status         = "Completed"
            }
        } catch {
            $summary.Items += [PSCustomObject]@{
                Target         = $target.Name
                RecoveredBytes = 0
                RecoveredText  = "0 B"
                Status         = "Skipped/Error: $_"
            }
            Write-Log "Cleanup target failed ($($target.Name)): $_" "WARN"
        }
    }

    try {
        Clear-RecycleBin -Force -ErrorAction SilentlyContinue | Out-Null
        $summary.Items += [PSCustomObject]@{
            Target         = "Recycle Bin"
            RecoveredBytes = 0
            RecoveredText  = "Cleared"
            Status         = "Completed"
        }
    } catch {
        $summary.Items += [PSCustomObject]@{
            Target         = "Recycle Bin"
            RecoveredBytes = 0
            RecoveredText  = "N/A"
            Status         = "Skipped/Error: $_"
        }
    }

    $summary.Performed = $true
    $summary.TotalRecoveredText = Format-Bytes -Bytes $summary.TotalRecoveredBytes
    Write-Log "Cleanup completed. Total recovered: $($summary.TotalRecoveredText)" "SUCCESS"
    return $summary
}

function Get-PerformanceBoostSuggestions {
    param(
        [PSCustomObject]$CPU,
        [PSCustomObject]$RAM,
        [array]$Disks,
        [array]$StartupApps,
        [PSCustomObject]$Battery,
        [array]$PerfIssues,
        [PSCustomObject]$CleanupSummary
    )

    $tips = @()

    if (Convert-ToDouble $CPU.UsagePercent 0 -ge 85) {
        $tips += "High CPU usage detected. Close non-essential CPU-heavy apps and review scheduled scans/tasks."
    }
    if (Convert-ToDouble $RAM.UsedPercent 0 -ge 80) {
        $tips += "Memory pressure is elevated. Keep fewer heavy apps open and consider upgrading to 16 GB+ RAM."
    }
    if (@($Disks | Where-Object { (Convert-ToDouble $_.FreePercent 100) -lt 15 }).Count -gt 0) {
        $tips += "One or more drives are low on free space. Keep at least 15-20% free for smoother performance."
    }
    if (@($StartupApps).Count -gt 10) {
        $tips += "Reduce startup applications to improve boot time and lower background load."
    }
    if ($Battery.Present -and $null -ne $Battery.HealthPercent -and (Convert-ToDouble $Battery.HealthPercent 100) -lt 60) {
        $tips += "Battery health is poor. Replacing the battery can improve stability and sustained performance."
    }
    if ($CleanupSummary -and $CleanupSummary.Performed) {
        $tips += "Cleanup recovered $($CleanupSummary.TotalRecoveredText). Run cleanup regularly to maintain responsiveness."
    } else {
        $tips += "Clean temporary locations periodically (%TEMP%, %LOCALAPPDATA%\\Temp, C:\\Windows\\Temp, Prefetch, SoftwareDistribution downloads, Recycle Bin)."
    }
    if (@($PerfIssues | Where-Object { $_.Severity -eq "HIGH" }).Count -eq 0) {
        $tips += "No critical bottlenecks detected. Keep drivers, BIOS, and Windows updates current for consistent performance."
    }

    return @($tips | Select-Object -Unique)
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

function Get-TopProcesses {
    param([int]$TopN = 15)
    Write-Log "Collecting top resource-consuming processes..." "INFO"
    $rows = @()
    try {
        $processes = @(Get-Process -ErrorAction SilentlyContinue)
        foreach ($proc in $processes) {
            $rows += [PSCustomObject]@{
                Name      = $proc.ProcessName
                PID       = $proc.Id
                MemoryMB  = [math]::Round($proc.WorkingSet64 / 1MB, 1)
                CPU_s     = [math]::Round($proc.CPU, 1)
                Handles   = $proc.Handles
                Threads   = $proc.Threads.Count
            }
        }
    } catch {
        Write-Log "Could not enumerate process metrics: $_" "WARN"
    }

    if ($rows.Count -eq 0) { return @() }
    return @($rows | Sort-Object -Property MemoryMB -Descending | Select-Object -First $TopN)
}

function Get-RecentHistoryEntries {
    param(
        [string]$Path,
        [int]$MaxEntries = 20
    )

    if (-not (Test-Path $Path)) { return @() }

    $entries = @()
    try {
        $lines = @(Get-Content -Path $Path -ErrorAction SilentlyContinue | Where-Object { $_ -and $_.Trim() -ne "" })
        $selected = @($lines | Select-Object -Last $MaxEntries)
        foreach ($line in $selected) {
            try {
                $entries += ($line | ConvertFrom-Json -ErrorAction Stop)
            } catch { }
        }
    } catch {
        Write-Log "Could not read history file $Path`: $_" "WARN"
    }
    return $entries
}

function Save-HistorySnapshot {
    param(
        [string]$Path,
        [PSCustomObject]$Snapshot
    )

    try {
        $Snapshot | ConvertTo-Json -Depth 6 -Compress | Add-Content -Path $Path
    } catch {
        Write-Log "Could not persist device history snapshot: $_" "WARN"
    }
}

function Get-RestartPowerEvents {
    param(
        [int]$Days = 7,
        [int]$MaxEvents = 15
    )

    $events = @()
    try {
        $startTime = (Get-Date).AddDays(-$Days)
        $eventIds = @(41, 1074, 6005, 6006, 6008)
        $rawEvents = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; StartTime = $startTime; Id = $eventIds } -ErrorAction Stop |
            Sort-Object TimeCreated -Descending |
            Select-Object -First $MaxEvents)

        foreach ($evt in $rawEvents) {
            $eventType = switch ($evt.Id) {
                41   { "Unexpected Power Loss / Crash" }
                1074 { "Planned Restart / Shutdown" }
                6005 { "Event Log Service Started (Boot)" }
                6006 { "Event Log Service Stopped (Clean Shutdown)" }
                6008 { "Unexpected Shutdown" }
                default { "System Event" }
            }
            $events += [PSCustomObject]@{
                Time    = $evt.TimeCreated
                EventId = $evt.Id
                Type    = $eventType
                Source  = $evt.ProviderName
                Detail  = if ($evt.Message) { ([string]$evt.Message).Split("`n")[0] } else { "No event message." }
            }
        }
    } catch {
        Write-Log "Could not query restart/power events: $_" "WARN"
    }
    return $events
}

function Get-DeviceHistoricalInsights {
    param(
        [string]$HistoryPath,
        [PSCustomObject]$CurrentSnapshot,
        [array]$TopProcesses = @(),
        [int]$TrendRuns = 15
    )

    $history = @(Get-RecentHistoryEntries -Path $HistoryPath -MaxEntries $TrendRuns)
    $trendRows = @()
    $derivedIssues = @()
    $appContributors = @()

    if ($history.Count -gt 0) {
        $avgCpu = [math]::Round((Convert-ToDouble ((@($history | Measure-Object -Property CPUUsagePercent -Average).Average)) 0), 1)
        $avgRam = [math]::Round((Convert-ToDouble ((@($history | Measure-Object -Property RAMUsedPercent -Average).Average)) 0), 1)
        $avgProc = [math]::Round((Convert-ToDouble ((@($history | Measure-Object -Property ProcessCount -Average).Average)) 0), 0)

        $trendRows += [PSCustomObject]@{
            Metric         = "CPU Usage"
            Current        = "$($CurrentSnapshot.CPUUsagePercent)%"
            Baseline       = "$avgCpu%"
            Delta          = "$([math]::Round($CurrentSnapshot.CPUUsagePercent - $avgCpu, 1))%"
            Interpretation = if ($CurrentSnapshot.CPUUsagePercent -gt ($avgCpu + 20)) { "Sudden CPU spike" } else { "Within expected variation" }
        }
        $trendRows += [PSCustomObject]@{
            Metric         = "RAM Usage"
            Current        = "$($CurrentSnapshot.RAMUsedPercent)%"
            Baseline       = "$avgRam%"
            Delta          = "$([math]::Round($CurrentSnapshot.RAMUsedPercent - $avgRam, 1))%"
            Interpretation = if ($CurrentSnapshot.RAMUsedPercent -gt ($avgRam + 12)) { "Sudden memory pressure" } else { "Within expected variation" }
        }
        $trendRows += [PSCustomObject]@{
            Metric         = "Process Count"
            Current        = "$($CurrentSnapshot.ProcessCount)"
            Baseline       = "$avgProc"
            Delta          = "$([math]::Round($CurrentSnapshot.ProcessCount - $avgProc, 0))"
            Interpretation = if ($CurrentSnapshot.ProcessCount -gt ($avgProc + 40)) { "Background app surge" } else { "Within expected variation" }
        }

        if ($CurrentSnapshot.RAMUsedPercent -gt ($avgRam + 12) -or $CurrentSnapshot.RAMUsedPercent -gt 90) {
            $derivedIssues += [PSCustomObject]@{
                Issue          = "Historical RAM Spike"
                Severity       = "HIGH"
                Detail         = "Current RAM ($($CurrentSnapshot.RAMUsedPercent)%) is significantly above recent baseline ($avgRam%)."
                RootCause      = "Memory consumption increased abruptly compared to previous diagnostic runs."
                Recommendation = "Inspect top memory consumers and disable/close persistent high-memory applications."
            }
        }

        if ($CurrentSnapshot.ProcessCount -gt ($avgProc + 40)) {
            $derivedIssues += [PSCustomObject]@{
                Issue          = "Historical Process Surge"
                Severity       = "MEDIUM"
                Detail         = "Process count increased from baseline $avgProc to $($CurrentSnapshot.ProcessCount)."
                RootCause      = "More background services/apps are active than in recent healthy runs."
                Recommendation = "Review startup entries and recurring background services introduced recently."
            }
        }
    } else {
        $trendRows += [PSCustomObject]@{
            Metric         = "History"
            Current        = "First tracked run"
            Baseline       = "N/A"
            Delta          = "N/A"
            Interpretation = "Trend analysis will improve after multiple runs."
        }
    }

    $restartEvents = @(Get-RestartPowerEvents -Days 7 -MaxEvents 15)
    $unexpectedEvents = @($restartEvents | Where-Object { $_.EventId -in @(41, 6008) })
    if ($unexpectedEvents.Count -gt 0) {
        $derivedIssues += [PSCustomObject]@{
            Issue          = "Unexpected Power/Restart Events"
            Severity       = "HIGH"
            Detail         = "$($unexpectedEvents.Count) unexpected shutdown/power-loss event(s) detected in the last 7 days."
            RootCause      = "Abrupt power loss, crash, thermal shutdown, or forced reset may be occurring."
            Recommendation = "Check power source/adapter, battery health, thermal conditions, and Windows reliability history."
        }
    }

    $recentProcessNames = @()
    foreach ($entry in $history) {
        foreach ($proc in @($entry.TopProcesses)) {
            if ($proc.Name) { $recentProcessNames += [string]$proc.Name }
        }
    }

    foreach ($proc in @($TopProcesses | Sort-Object MemoryMB -Descending | Select-Object -First 10)) {
        $repeatCount = @($recentProcessNames | Where-Object { $_ -eq $proc.Name }).Count
        if ($proc.MemoryMB -ge 400 -or $repeatCount -ge 4) {
            $appContributors += [PSCustomObject]@{
                ProcessName = $proc.Name
                MemoryMB    = $proc.MemoryMB
                CPU_s       = $proc.CPU_s
                Recurrence  = $repeatCount
                Reason      = if ($repeatCount -ge 4) { "Recurring heavy process across previous runs" } else { "High current memory usage" }
            }
        }
    }

    if ($appContributors.Count -eq 0) {
        $appContributors = @($TopProcesses | Sort-Object MemoryMB -Descending | Select-Object -First 5 | ForEach-Object {
            [PSCustomObject]@{
                ProcessName = $_.Name
                MemoryMB    = $_.MemoryMB
                CPU_s       = $_.CPU_s
                Recurrence  = 0
                Reason      = "Top active process in current snapshot"
            }
        })
    }

    return [PSCustomObject]@{
        TrendRows       = $trendRows
        RestartEvents   = $restartEvents
        AppContributors = $appContributors
        DerivedIssues   = $derivedIssues
        HistoryRuns     = $history.Count
    }
}

function Get-PerformanceAnalysis {
    param(
        [PSCustomObject]$CPU,
        [PSCustomObject]$RAM,
        [array]$Disks,
        [array]$StartupApps,
        [array]$Temperatures = @(),
        [array]$TopProcesses = @()
    )
    Write-Log "Performing performance analysis..." "INFO"
    $issues = @()
    $processCount = (Get-Process).Count
    $maxTemp = if ($Temperatures.Count -gt 0) { ($Temperatures | Measure-Object -Property TempC -Maximum).Maximum } else { $null }

    $slowdownScore = 0
    if ($CPU.UsagePercent -gt 85) { $slowdownScore += 35 }
    if ($RAM.UsedPercent -gt 85) { $slowdownScore += 35 }
    if ($processCount -gt 180) { $slowdownScore += 15 }
    if ($StartupApps.Count -gt 15) { $slowdownScore += 10 }
    if (@($Disks | Where-Object { $_.FreePercent -lt 10 }).Count -gt 0) { $slowdownScore += 20 }
    if ($null -ne $maxTemp -and $maxTemp -ge 80) { $slowdownScore += 25 }

    $slowIndicators = @()
    if ($CPU.UsagePercent -gt 85) { $slowIndicators += "CPU saturation ($($CPU.UsagePercent)%)" }
    if ($RAM.UsedPercent -gt 85) { $slowIndicators += "memory pressure ($($RAM.UsedPercent)%)" }
    if ($processCount -gt 180) { $slowIndicators += "high process count ($processCount)" }
    if ($StartupApps.Count -gt 15) { $slowIndicators += "heavy startup load ($($StartupApps.Count) items)" }
    $criticalDisks = @($Disks | Where-Object { $_.FreePercent -lt 10 } | ForEach-Object { "$($_.Drive) free $($_.FreePercent)%" })
    if ($criticalDisks.Count -gt 0) { $slowIndicators += "low disk free space ($($criticalDisks -join '; '))" }
    if ($null -ne $maxTemp -and $maxTemp -ge 80) { $slowIndicators += "thermal throttling risk ($maxTemp °C)" }

    if ($slowdownScore -ge 40) {
        $issues += [PSCustomObject]@{
            Issue          = "System Slowdown Risk"
            Severity       = if ($slowdownScore -ge 70) { "HIGH" } else { "MEDIUM" }
            Detail         = "Measured slowdown risk score: $slowdownScore/140. Indicators: $($slowIndicators -join ', ')."
            RootCause      = "Performance bottlenecks are cumulative. Multiple medium stressors (CPU, RAM, startup load, thermals, disk pressure) combine to create visible lag."
            Recommendation = "Prioritize top memory/CPU processes, reduce startup items, keep >15% free space on system drive, and address cooling if temperature is high."
        }
    }

    if ($RAM.TotalGB -lt 8) {
        $issues += [PSCustomObject]@{
            Issue          = "Low RAM Capacity"
            Severity       = "HIGH"
            Detail         = "Installed RAM is $($RAM.TotalGB) GB, below modern multitasking baseline."
            RootCause      = "Insufficient physical memory forces frequent paging to disk, causing app switching lag and stutter."
            Recommendation = "Upgrade to at least 16 GB RAM for smoother multitasking and lower disk paging pressure."
        }
    }
    if ($CPU.UsagePercent -gt 85) {
        $issues += [PSCustomObject]@{
            Issue          = "High CPU Usage"
            Severity       = "HIGH"
            Detail         = "CPU usage is $($CPU.UsagePercent)% during capture."
            RootCause      = "One or more workloads are saturating CPU cores, increasing response time for interactive tasks."
            Recommendation = "Identify top CPU processes from Task Manager/Process Explorer, close non-essential background tasks, and check for scheduled scans/updates."
        }
    }
    foreach ($d in $Disks) {
        if ($d.FreePercent -lt 10) {
            $issues += [PSCustomObject]@{
                Issue          = "Disk Space Bottleneck"
                Severity       = "HIGH"
                Detail         = "Drive $($d.Drive) has only $($d.FreePercent)% free space."
                RootCause      = "Low free space reduces filesystem efficiency, update/cache behavior, and can increase write amplification."
                Recommendation = "Free 15-20% space on $($d.Drive), clear temp files, move large archives/media, and uninstall unused software."
            }
        }
        if ($d.DiskType -eq "HDD") {
            $issues += [PSCustomObject]@{
                Issue          = "HDD Performance Constraint"
                Severity       = "MEDIUM"
                Detail         = "Drive $($d.Drive) is detected as HDD."
                RootCause      = "Mechanical seek latency is significantly higher than SSD, slowing boot, app launch, and random I/O-heavy tasks."
                Recommendation = "Migrate OS and frequently used apps to SSD/NVMe for major responsiveness gains."
            }
        }
    }
    if ($processCount -gt 150) {
        $issues += [PSCustomObject]@{
            Issue          = "Excessive Background Processes"
            Severity       = "MEDIUM"
            Detail         = "$processCount processes are running."
            RootCause      = "Excess startup and background services increase scheduler overhead, RAM footprint, and contention for CPU time."
            Recommendation = "Disable non-essential startup apps/services, uninstall unused software, and review scheduled tasks that spawn background agents."
        }
    }
    if ($StartupApps.Count -gt 10) {
        $severity = if ($StartupApps.Count -gt 20) { "HIGH" } else { "MEDIUM" }
        $issues += [PSCustomObject]@{
            Issue          = "Heavy Startup Load"
            Severity       = $severity
            Detail         = "$($StartupApps.Count) startup apps detected."
            RootCause      = "Too many autorun entries delay logon initialization and keep persistent background footprint after boot."
            Recommendation = "Keep startup entries below 10 where possible. Disable low-value autoruns from Task Manager > Startup or Autoruns utility."
        }
    }
    if ($RAM.UsedPercent -gt 85) {
        $topRam = @($TopProcesses | Sort-Object MemoryMB -Descending | Select-Object -First 5)
        $topRamSummary = if ($topRam.Count -gt 0) {
            ($topRam | ForEach-Object { "$($_.Name) ($($_.MemoryMB) MB)" }) -join ", "
        } else {
            "Top process data unavailable"
        }
        $issues += [PSCustomObject]@{
            Issue          = "High RAM Usage"
            Severity       = "HIGH"
            Detail         = "RAM usage is $($RAM.UsedPercent)% ($($RAM.UsedGB) GB used of $($RAM.TotalGB) GB). Top consumers: $topRamSummary."
            RootCause      = "One or more applications are keeping large working sets, increasing paging and reducing responsiveness."
            Recommendation = "Close/restart the highest-memory processes listed below, update memory-heavy apps, and consider a RAM upgrade if high usage is persistent."
        }
    }

    foreach ($zone in $Temperatures) {
        if ($zone.TempC -ge 90) {
            $issues += [PSCustomObject]@{
                Issue          = "Critical Temperature"
                Severity       = "HIGH"
                Detail         = "$($zone.ZoneName) is at $($zone.TempC)°C."
                RootCause      = "Thermal headroom is exhausted and CPU/GPU may throttle to prevent damage, which directly slows performance."
                Recommendation = "Clean vents/fans, improve airflow, reduce heavy workloads temporarily, and verify fan curve or thermal paste condition."
            }
        } elseif ($zone.TempC -ge 80) {
            $issues += [PSCustomObject]@{
                Issue          = "High Temperature"
                Severity       = "HIGH"
                Detail         = "$($zone.ZoneName) is at $($zone.TempC)°C."
                RootCause      = "Sustained high temperature reduces boost clocks and can trigger thermal throttling."
                Recommendation = "Check fan operation, remove dust buildup, ensure unobstructed airflow, and reduce simultaneous CPU+GPU load."
            }
        } elseif ($zone.TempC -ge 65) {
            $issues += [PSCustomObject]@{
                Issue          = "Elevated Temperature"
                Severity       = "MEDIUM"
                Detail         = "$($zone.ZoneName) is running warm at $($zone.TempC)°C."
                RootCause      = "Thermal load is above ideal baseline; prolonged operation can reduce sustained performance."
                Recommendation = "Monitor trend during heavy workloads and optimize cooling profile if temperature frequently rises further."
            }
        }
    }

    if ($issues.Count -eq 0) {
        $issues += [PSCustomObject]@{
            Issue          = "No Issues Detected"
            Severity       = "OK"
            Detail         = "No major performance bottlenecks were detected in this snapshot."
            RootCause      = "System metrics are currently within healthy operating ranges."
            Recommendation = "Keep drivers/OS updated and rerun diagnostics during peak usage if intermittent lag occurs."
        }
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
        [array]$Temperatures = @(),
        [array]$TopProcesses = @(),
        [PSCustomObject]$HistoricalInsights = $null,
        [PSCustomObject]$CleanupSummary = $null,
        [array]$PerformanceBoostSuggestions = @()
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
                <td style='font-size:0.82em;'>$([System.Net.WebUtility]::HtmlEncode($d.HardwareSummary))</td>
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
                <td>$($p.RootCause)</td>
                <td>$($p.Recommendation)</td>
            </tr>
"@
    }

    $processRowsHtml = ""
    foreach ($proc in ($TopProcesses | Sort-Object MemoryMB -Descending)) {
        $ramColorProc = if ($proc.MemoryMB -ge 1000) { "#dc3545" } elseif ($proc.MemoryMB -ge 500) { "#ffc107" } else { "#28a745" }
        $processRowsHtml += @"
            <tr>
                <td>$([System.Net.WebUtility]::HtmlEncode($proc.Name))</td>
                <td>$($proc.PID)</td>
                <td style='font-weight:700;color:$ramColorProc;'>$($proc.MemoryMB) MB</td>
                <td>$($proc.CPU_s)</td>
                <td>$($proc.Handles)</td>
                <td>$($proc.Threads)</td>
            </tr>
"@
    }
    if (-not $processRowsHtml) {
        $processRowsHtml = "<tr><td colspan='6' style='text-align:center;color:#6c757d;'>Process data unavailable.</td></tr>"
    }

    $trendRowsHtml = ""
    foreach ($t in @($HistoricalInsights.TrendRows)) {
        $trendRowsHtml += @"
            <tr>
                <td>$($t.Metric)</td>
                <td>$($t.Current)</td>
                <td>$($t.Baseline)</td>
                <td>$($t.Delta)</td>
                <td>$($t.Interpretation)</td>
            </tr>
"@
    }
    if (-not $trendRowsHtml) {
        $trendRowsHtml = "<tr><td colspan='5' style='text-align:center;color:#6c757d;'>No historical trend data yet.</td></tr>"
    }

    $restartRowsHtml = ""
    foreach ($e in @($HistoricalInsights.RestartEvents)) {
        $eventColor = if ($e.EventId -in @(41, 6008)) { "#dc3545" } elseif ($e.EventId -eq 1074) { "#ffc107" } else { "#6c757d" }
        $restartRowsHtml += @"
            <tr>
                <td>$($e.Time)</td>
                <td>$($e.EventId)</td>
                <td><span style='background:$eventColor;color:#fff;padding:2px 8px;border-radius:10px;font-size:0.8em;'>$($e.Type)</span></td>
                <td>$($e.Source)</td>
                <td>$([System.Net.WebUtility]::HtmlEncode($e.Detail))</td>
            </tr>
"@
    }
    if (-not $restartRowsHtml) {
        $restartRowsHtml = "<tr><td colspan='5' style='text-align:center;color:#6c757d;'>No restart/power events found in selected history window.</td></tr>"
    }

    $appRowsHtml = ""
    foreach ($a in @($HistoricalInsights.AppContributors)) {
        $appRowsHtml += @"
            <tr>
                <td>$([System.Net.WebUtility]::HtmlEncode($a.ProcessName))</td>
                <td>$($a.MemoryMB)</td>
                <td>$($a.CPU_s)</td>
                <td>$($a.Recurrence)</td>
                <td>$($a.Reason)</td>
            </tr>
"@
    }
    if (-not $appRowsHtml) {
        $appRowsHtml = "<tr><td colspan='5' style='text-align:center;color:#6c757d;'>No dominant application contributors identified.</td></tr>"
    }

    $cleanupRowsHtml = ""
    foreach ($row in @($CleanupSummary.Items)) {
        $cleanupRowsHtml += @"
            <tr>
                <td>$([System.Net.WebUtility]::HtmlEncode($row.Target))</td>
                <td>$([System.Net.WebUtility]::HtmlEncode($row.RecoveredText))</td>
                <td>$([System.Net.WebUtility]::HtmlEncode($row.Status))</td>
            </tr>
"@
    }
    if (-not $cleanupRowsHtml) {
        $cleanupRowsHtml = "<tr><td colspan='3' style='text-align:center;color:#6c757d;'>Cleanup was not executed in this run.</td></tr>"
    }

    $installedBatteryHtml = if ($Battery.Present -and ($Battery.BatteryName -or $Battery.BatteryManufacturer -or $Battery.BatterySerial)) {
        "<tr>" +
        "<td>$(if($Battery.BatteryName){[System.Net.WebUtility]::HtmlEncode($Battery.BatteryName)}else{'N/A'})</td>" +
        "<td>$(if($Battery.BatteryManufacturer){[System.Net.WebUtility]::HtmlEncode($Battery.BatteryManufacturer)}else{'N/A'})</td>" +
        "<td>$(if($Battery.BatterySerial){[System.Net.WebUtility]::HtmlEncode($Battery.BatterySerial)}else{'N/A'})</td>" +
        "<td>$(if($Battery.BatteryChemistry){[System.Net.WebUtility]::HtmlEncode($Battery.BatteryChemistry)}else{'N/A'})</td>" +
        "<td>N/A</td>" +
        "</tr>"
    } else {
        $noDataMsg = if ($Battery.Present) { "Battery identification data unavailable &mdash; re-run with Administrator privileges." } else { "No battery installed on this system." }
        "<tr><td colspan='5' style='text-align:center;color:#6c757d;'>$noDataMsg</td></tr>"
    }

    $cleanupSuggestions = @(
        [PSCustomObject]@{ Target = "%TEMP%"; Action = "Delete temporary app/installer files"; Benefit = "Frees disk space and reduces file-system clutter" },
        [PSCustomObject]@{ Target = "%LOCALAPPDATA%\\Temp"; Action = "Remove stale user temp data"; Benefit = "Improves responsiveness for daily workloads" },
        [PSCustomObject]@{ Target = "C:\\Windows\\Temp"; Action = "Clear system temporary files"; Benefit = "Reclaims space consumed by update/install leftovers" },
        [PSCustomObject]@{ Target = "C:\\Windows\\Prefetch"; Action = "Trim stale prefetch cache"; Benefit = "Helps reduce boot/runtime overhead from outdated entries" },
        [PSCustomObject]@{ Target = "C:\\Windows\\SoftwareDistribution\\Download"; Action = "Clean old Windows update downloads"; Benefit = "Recovers significant storage on long-running systems" },
        [PSCustomObject]@{ Target = "Recycle Bin"; Action = "Empty deleted files"; Benefit = "Immediately frees recoverable disk space" }
    )
    $cleanupSuggestionRowsHtml = ""
    foreach ($s in $cleanupSuggestions) {
        $cleanupSuggestionRowsHtml += @"
            <tr>
                <td>$([System.Net.WebUtility]::HtmlEncode($s.Target))</td>
                <td>$([System.Net.WebUtility]::HtmlEncode($s.Action))</td>
                <td>$([System.Net.WebUtility]::HtmlEncode($s.Benefit))</td>
            </tr>
"@
    }

    $cleanupStatusText = if ($CleanupSummary -and $CleanupSummary.Performed) {
        "Cleanup executed. Total recovered: $($CleanupSummary.TotalRecoveredText)"
    } else {
        "Automatic cleanup was not executed. Use these suggestions to safely free disk space and improve responsiveness."
    }

    $suggestionRowsHtml = ""
    foreach ($tip in @($PerformanceBoostSuggestions)) {
        $suggestionRowsHtml += "<li>$([System.Net.WebUtility]::HtmlEncode($tip))</li>"
    }
    if (-not $suggestionRowsHtml) {
        $suggestionRowsHtml = "<li>No additional optimization suggestions at this time.</li>"
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
                      elseif ($maxTempC -ge 90) { "CRITICAL - check cooling!" }
                      elseif ($maxTempC -ge 80) { "Hot - check cooling system" }
                      elseif ($maxTempC -ge 65) { "Warm - monitor closely" }
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
            <div class="sub">$($Battery.Status)$(if($null -ne $Battery.HealthPercent){" | Health: $($Battery.HealthPercent)% ($($Battery.HealthStatus))"}else{""})</div>
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
            <thead><tr><th>Drive</th><th>Type</th><th>Hardware Summary</th><th>Total</th><th>Used</th><th>Free</th><th>Recommendation</th></tr></thead>
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
            <thead><tr><th>Severity</th><th>Issue</th><th>Details</th><th>Why It Happens</th><th>What To Do</th></tr></thead>
            <tbody>$perfRowsHtml</tbody>
        </table>
    </div>

    <!-- Top RAM Processes -->
    <div class="section">
        <h2>&#x1F9E0; Top RAM Consumers (Slowdown/Heat Contributors)</h2>
        $(if($RAM.UsedPercent -ge 80){"<p style='color:#dc3545;margin-bottom:12px;'>&#x26A0; High memory pressure detected. Prioritize the highest RAM consumers below.</p>"}else{"<p style='color:#6c757d;margin-bottom:12px;'>Memory usage is currently manageable. This list still helps identify heavy background apps.</p>"})
        <table>
            <thead><tr><th>Process</th><th>PID</th><th>RAM (MB)</th><th>CPU Time (s)</th><th>Handles</th><th>Threads</th></tr></thead>
            <tbody>$processRowsHtml</tbody>
        </table>
    </div>

    <!-- Historical Trend & Incident Correlation -->
    <div class="section">
        <h2>&#x1F4CA; Historical Performance Trend (Previous Runs)</h2>
        <p style="margin-bottom:12px;color:#6c757d;">Trend window analyzed: $($HistoricalInsights.HistoryRuns) previous run(s).</p>
        <table>
            <thead><tr><th>Metric</th><th>Current</th><th>Baseline</th><th>Delta</th><th>Interpretation</th></tr></thead>
            <tbody>$trendRowsHtml</tbody>
        </table>
    </div>

    <div class="section">
        <h2>&#x1F50C; Restart &amp; Power Event Correlation (Last 7 Days)</h2>
        <table>
            <thead><tr><th>Time</th><th>Event ID</th><th>Type</th><th>Source</th><th>Detail</th></tr></thead>
            <tbody>$restartRowsHtml</tbody>
        </table>
    </div>

    <div class="section">
        <h2>&#x1F4BB; Likely Application Slowdown Contributors</h2>
        <table>
            <thead><tr><th>Process</th><th>RAM (MB)</th><th>CPU Time (s)</th><th>Seen In History</th><th>Why Flagged</th></tr></thead>
            <tbody>$appRowsHtml</tbody>
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

    <div class="section">
        <h2>&#x1F50B; Battery Health Analysis</h2>
        <h3 style='font-size:0.95em;margin:14px 0 10px;padding-bottom:6px;border-bottom:1px solid #dee2e6;'>&#x1F4B4; Installed Batteries</h3>
        <p style='font-size:0.85em;color:#6c757d;margin-bottom:10px;'>Physical battery identification data read from Windows battery report</p>
        <table>
            <thead><tr><th>Battery Name</th><th>Manufacturer</th><th>Serial Number</th><th>Chemistry</th><th>Battery Age</th></tr></thead>
            <tbody>$installedBatteryHtml</tbody>
        </table>
        <h3 style='font-size:0.95em;margin:20px 0 10px;padding-bottom:6px;border-bottom:1px solid #dee2e6;'>&#x2764;&#xFE0F; Battery Health Metrics</h3>
        <table>
            <thead><tr><th>Metric</th><th>Value</th><th>Recommendation</th></tr></thead>
            <tbody>
                <tr><td>Design Capacity</td><td>$(if($null -ne $Battery.DesignCapacity_mWh -and $Battery.DesignCapacity_mWh -gt 0){"$($Battery.DesignCapacity_mWh) mWh"}else{"N/A"})</td><td rowspan="4">$($Battery.HealthRecommendation)</td></tr>
                <tr><td>Full Charge Capacity</td><td>$(if($null -ne $Battery.FullChargeCapacity_mWh -and $Battery.FullChargeCapacity_mWh -gt 0){"$($Battery.FullChargeCapacity_mWh) mWh"}else{"N/A"})</td></tr>
                <tr><td>Cycle Count</td><td>$(if($null -ne $Battery.CycleCount){$Battery.CycleCount}else{"N/A"})</td></tr>
                <tr><td>Battery Health</td><td>$(if($null -ne $Battery.HealthPercent){"$($Battery.HealthPercent)% ($($Battery.HealthStatus))"}else{"N/A"})</td></tr>
            </tbody>
        </table>
    </div>

    <div class="section">
        <h2>&#x1F9F9; System Cleanup Suggestions</h2>
        <p style='margin-bottom:12px;color:#6c757d;'>Status: $cleanupStatusText</p>
        <table>
            <thead><tr><th>Location</th><th>Suggested Action</th><th>Expected Benefit</th></tr></thead>
            <tbody>$cleanupSuggestionRowsHtml</tbody>
        </table>
        $(if($CleanupSummary -and $CleanupSummary.Performed){"<h3 style='margin-top:16px;margin-bottom:10px;font-size:0.95em;'>Cleanup Results (This Run)</h3>"}else{""})
        $(if($CleanupSummary -and $CleanupSummary.Performed){"<table><thead><tr><th>Target</th><th>Recovered</th><th>Status</th></tr></thead><tbody>$cleanupRowsHtml</tbody></table>"}else{""})
    </div>

    <div class="section">
        <h2>&#x1F4CC; Battery Data Note</h2>
        <p style='color:#6c757d;'>$(if($null -eq $Battery.HealthPercent -and $null -eq $Battery.DesignCapacity_mWh -and $null -eq $Battery.FullChargeCapacity_mWh){"Battery data unavailable on this system."}else{"Battery health percentage is calculated as (Full Charge Capacity / Design Capacity) * 100."})</p>
        <table>
            <thead><tr><th>Formula</th><th>Current Value</th><th>Status Band</th></tr></thead>
            <tbody>
                <tr>
                    <td>(Full Charge Capacity / Design Capacity) * 100</td>
                    <td>$(if($null -ne $Battery.HealthPercent){"$($Battery.HealthPercent)%"}else{"N/A"})</td>
                    <td>$(if($null -eq $Battery.HealthPercent){"Unavailable"}elseif($Battery.HealthPercent -ge 80){">=80: Battery healthy"}elseif($Battery.HealthPercent -ge 60){"60-79: Moderate degradation"}else{"<60: Battery replacement recommended"})</td>
                </tr>
            </tbody>
        </table>
    </div>

    <div class="section">
        <h2>&#x1F680; Performance Boost Suggestions</h2>
        <ul style='padding-left:20px;line-height:1.7;'>$suggestionRowsHtml</ul>
    </div>
</div>
<footer>
    <p>Report Version: 2.0 &nbsp;|&nbsp; Created by: <strong>Tushar Gudde</strong> &nbsp;|&nbsp;
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
    $startupApps = @(Get-StartupPrograms)
    $topProcesses = @(Get-TopProcesses -TopN 15)
    $tempInfo    = @(Get-TemperatureInfo)
    $perfIssues  = Get-PerformanceAnalysis -CPU $cpuInfo -RAM $ramInfo -Disks $diskInfo -StartupApps $startupApps -Temperatures $tempInfo -TopProcesses $topProcesses

    $cleanupSummary = [PSCustomObject]@{
        Requested           = $false
        Performed           = $false
        TotalRecoveredBytes = 0
        TotalRecoveredText  = "0 B"
        Items               = @()
    }
    if ($RunCleanup) {
        $cleanupSummary = Invoke-SystemCleanup -Force:$ForceCleanup
    }

    $currentSnapshot = [PSCustomObject]@{
        Timestamp       = (Get-Date).ToString("o")
        CPUUsagePercent = $cpuInfo.UsagePercent
        RAMUsedPercent  = $ramInfo.UsedPercent
        ProcessCount    = (Get-Process).Count
        StartupCount    = $startupApps.Count
        MaxTempC        = if ($tempInfo.Count -gt 0) { ($tempInfo | Measure-Object -Property TempC -Maximum).Maximum } else { $null }
        TopProcesses    = @($topProcesses | Select-Object -First 8 -Property Name, MemoryMB, CPU_s)
    }

    $historicalInsights = Get-DeviceHistoricalInsights -HistoryPath $HistoryFile -CurrentSnapshot $currentSnapshot -TopProcesses $topProcesses
    if (@($historicalInsights.DerivedIssues).Count -gt 0) {
        $perfIssues = Merge-ObjectArrays -Primary $perfIssues -Secondary $historicalInsights.DerivedIssues
    }

    $perfBoostSuggestions = Get-PerformanceBoostSuggestions `
        -CPU $cpuInfo `
        -RAM $ramInfo `
        -Disks $diskInfo `
        -StartupApps $startupApps `
        -Battery $batteryInfo `
        -PerfIssues $perfIssues `
        -CleanupSummary $cleanupSummary

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
        -Temperatures  $tempInfo `
        -TopProcesses  $topProcesses `
        -HistoricalInsights $historicalInsights `
        -CleanupSummary $cleanupSummary `
        -PerformanceBoostSuggestions $perfBoostSuggestions

    $htmlContent | Out-File -FilePath $ReportFile -Encoding UTF8 -Force
    Write-Log "Report saved: $ReportFile" "SUCCESS"

    Save-HistorySnapshot -Path $HistoryFile -Snapshot $currentSnapshot

    # Alert check
    $alertScript = Join-Path $PSScriptRoot "DeviceHealth_Alert.ps1"
    if (Test-Path $alertScript) {
        $criticalIssues = @($perfIssues | Where-Object { $_.Severity -eq "HIGH" })
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
