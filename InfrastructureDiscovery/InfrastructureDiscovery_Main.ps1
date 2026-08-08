#Requires -Version 5.1
<#
.SYNOPSIS
    InfraEye - Infrastructure Discovery Module
.DESCRIPTION
    Scans local networks using ping sweep and ARP table analysis.
    Discovers, classifies, and exports all network devices to HTML report and Excel.
    Performs DHCP and subnet analysis.
.AUTHOR
    Tushar Gudde
.WEBSITE
    https://tushargudde.tech
.VERSION
    2.0
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

# Enforce TLS 1.2/1.3
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13

function Set-SecureDirectory {
    param([string]$Path)
    try {
        $acl = Get-Acl -Path $Path
        $acl.SetAccessRuleProtection($true, $false)
        $adminRule = New-Object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
        $systemRule = New-Object System.Security.AccessControl.FileSystemAccessRule("NT AUTHORITY\SYSTEM", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
        $acl.AddAccessRule($adminRule)
        $acl.AddAccessRule($systemRule)
        Set-Acl -Path $Path -AclObject $acl -ErrorAction SilentlyContinue
    } catch { }
}
Set-SecureDirectory -Path $ReportDir
Set-SecureDirectory -Path $LogDir

$Timestamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile     = Join-Path $LogDir "InfrastructureDiscovery_$Timestamp.log"
$ReportFile  = Join-Path $ReportDir "InfrastructureDiscovery_$Timestamp.html"
$ExcelFile   = Join-Path $ReportDir "NetworkDevices_$Timestamp.xlsx"
$ConfigFile  = Join-Path $PSScriptRoot "..\config.json"
$script:MacVendorCache = @{}

# ─────────────────────────────────────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────────────────────────────────────
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS")]
        [string]$Level = "INFO"
    )
    $Entry = "{0} | InfrastructureDiscovery | {1} | {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
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
# HELPER: GET LOCAL NETWORKS
# ─────────────────────────────────────────────────────────────────────────────
function Get-LocalNetworkRanges {
    Write-Log "Detecting local network ranges dynamically..." "INFO"
    $ranges = @()
    try {
        $adapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True"
        foreach ($a in $adapters) {
            if ($a.IPAddress -and $a.IPSubnet) {
                # Find the first IPv4 address
                for ($i = 0; $i -lt $a.IPAddress.Count; $i++) {
                    $ip = $a.IPAddress[$i]
                    $mask = $a.IPSubnet[$i]
                    if ($ip -match '^\d+\.\d+\.\d+\.\d+$') {
                        # Convert mask to CIDR prefix length
                        $maskOctets = $mask -split '\.'
                        $prefix = 0
                        foreach ($octet in $maskOctets) {
                            $bin = [Convert]::ToString([int]$octet, 2)
                            $prefix += ($bin -replace '0','').Length
                        }
                        
                        # Calculate Base Network Address
                        $ipOctets = $ip -split '\.'
                        $netOctets = @()
                        for ($j = 0; $j -lt 4; $j++) {
                            $netOctets += ([int]$ipOctets[$j] -band [int]$maskOctets[$j])
                        }
                        $networkAddress = $netOctets -join '.'
                        $cidr = "$networkAddress/$prefix"
                        
                        # Exclude 127.0.0.0/8 and 169.254.0.0/16 (APIPA)
                        if ($networkAddress -notmatch '^127\.' -and $networkAddress -notmatch '^169\.254\.') {
                            if ($ranges -notcontains $cidr) {
                                $ranges += $cidr
                            }
                        }
                    }
                }
            }
        }
    } catch {
        Write-Log "Error detecting dynamic network ranges: $_" "WARN"
    }

    if ($ranges.Count -eq 0) {
        Write-Log "No ranges detected. Attempting fallback from config.json" "WARN"
        try {
            $config = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json
            if ($config.network_scan.ranges -and $config.network_scan.ranges -ne "dynamic") {
                $ranges = $config.network_scan.ranges
            } else {
                $ranges = @("192.168.1.0/24")
            }
        } catch { $ranges = @("192.168.1.0/24") }
    }
    return $ranges
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER: PING SWEEP
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-PingSweep {
    param([string]$NetworkCIDR)
    Write-Log "Ping sweeping network: $NetworkCIDR" "INFO"

    $baseIP = ($NetworkCIDR -split '/')[0]
    $octets = $baseIP -split '\.'
    $prefix = "$($octets[0]).$($octets[1]).$($octets[2])"

    $config     = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json
    $timeout    = $config.network_scan.ping_timeout_ms
    $maxThreads = $config.network_scan.max_threads

    $tasks = @()
    1..254 | ForEach-Object {
        $ip = "$prefix.$_"
        $ping = New-Object System.Net.NetworkInformation.Ping
        $tasks += [PSCustomObject]@{
            IP = $ip
            Ping = $ping
            Task = $ping.SendPingAsync($ip, $timeout)
        }
    }

    try { [Threading.Tasks.Task]::WaitAll($tasks.Task) } catch { }

    $aliveHosts = [System.Collections.Generic.List[string]]::new()
    foreach ($t in $tasks) {
        try {
            if ($t.Task.Result.Status -eq 'Success') { $aliveHosts.Add($t.IP) }
        } catch { }
        finally { $t.Ping.Dispose() }
    }

    $sortedHosts = @($aliveHosts | Sort-Object { [System.Version]$_ })
    Write-Log "Ping sweep complete. Found $($sortedHosts.Count) live hosts in $NetworkCIDR" "INFO"
    return $sortedHosts
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER: GET ARP TABLE
# ─────────────────────────────────────────────────────────────────────────────
function Get-ARPTable {
    Write-Log "Reading ARP table..." "INFO"
    $arpEntries = @{}

    # Primary: Get-NetNeighbor (PS 3.0+, more complete; includes Reachable + Stale entries)
    try {
        $neighbors = Get-NetNeighbor -State Reachable,Stale,Permanent -ErrorAction SilentlyContinue
        foreach ($n in @($neighbors)) {
            if ($n.IPAddress -notmatch '^\d+\.\d+\.\d+\.\d+$') { continue }
            $rawMac = $n.LinkLayerAddress
            if (-not $rawMac -or $rawMac -match '^0{12}$' -or $rawMac -match '^F{12}$') { continue }
            # Normalize to AA:BB:CC:DD:EE:FF regardless of input format
            $hexOnly = $rawMac -replace '[^0-9A-Fa-f]', ''
            if ($hexOnly.Length -ne 12) { continue }
            $mac = [regex]::Replace($hexOnly.ToUpper(), '(..)(?!$)', '$1:')
            if (-not $arpEntries.ContainsKey($n.IPAddress)) {
                $arpEntries[$n.IPAddress] = $mac
            }
        }
    } catch {
        Write-Log "Get-NetNeighbor unavailable, falling back to arp.exe: $_" "DEBUG"
    }

    # Supplementary: arp.exe (fills any gaps on older systems / additional interfaces)
    try {
        $arpOutput = & arp -a 2>$null
        foreach ($line in $arpOutput) {
            if ($line -match '^\s*(\d+\.\d+\.\d+\.\d+)\s+([\w-]{17})\s+') {
                $ip  = $Matches[1]
                $mac = $Matches[2].ToUpper() -replace '-',':'
                if (-not $arpEntries.ContainsKey($ip)) {
                    $arpEntries[$ip] = $mac
                }
            }
        }
    } catch {
        Write-Log "Failed to read ARP table via arp.exe: $_" "WARN"
    }

    Write-Log "ARP table read. Found $($arpEntries.Count) entries." "INFO"
    return $arpEntries
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER: RESOLVE HOSTNAME
# ─────────────────────────────────────────────────────────────────────────────
function Resolve-HostnameForIP {
    param([string]$IPAddress)
    try {
        $hostEntry = [System.Net.Dns]::GetHostEntry($IPAddress)
        return $hostEntry.HostName
    } catch {
        return $IPAddress
    }
}

function Get-LocalDeviceProfile {
    Write-Log "Collecting local device profile..." "INFO"
    try {
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
        $operatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem
        $productType = [int]$operatingSystem.ProductType
        $adapters = @(Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True")

        $ipAddresses = @()
        $ipToMac = @{}
        foreach ($adapter in $adapters) {
            $normalizedMac = if ($adapter.MACAddress) { $adapter.MACAddress.ToUpper().Replace('-',':') } else { $null }
            $adapterIPs = @($adapter.IPAddress | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' })
            foreach ($ip in $adapterIPs) {
                $ipAddresses += $ip
                if ($normalizedMac) {
                    $ipToMac[$ip] = $normalizedMac
                }
            }
        }

        $isServerOS = ($productType -in @(2, 3))
        $deviceType = if ($isServerOS) {
            "Server"
        } else {
            switch ([int]$computerSystem.PCSystemTypeEx) {
                2 { "Laptop" }
                8 { "Tablet" }
                3 { "Desktop" }
                default {
                    if ($computerSystem.Model -match "Laptop|Notebook|Book|Ultrabook|Surface") { "Laptop" }
                    else { "Desktop" }
                }
            }
        }

        [PSCustomObject]@{
            Hostname      = $env:COMPUTERNAME
            Manufacturer  = $computerSystem.Manufacturer
            Model         = $computerSystem.Model
            OSName        = $operatingSystem.Caption
            OSProductType = $productType
            DeviceType    = $deviceType
            IsServer      = $isServerOS
            IPAddresses   = @($ipAddresses | Sort-Object -Unique)
            IPToMACMap    = $ipToMac
        }
    } catch {
        Write-Log "Could not build local device profile: $_" "WARN"
        [PSCustomObject]@{
            Hostname      = $env:COMPUTERNAME
            Manufacturer  = "Unknown"
            Model         = "Unknown"
            OSName        = "Unknown"
            OSProductType = 1
            DeviceType    = "Client Device"
            IsServer      = $false
            IPAddresses   = @()
            IPToMACMap    = @{}
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER: GET VENDOR FROM MAC
# ─────────────────────────────────────────────────────────────────────────────
function Get-VendorFromMAC {
    param([string]$MAC)
    if (-not $MAC -or $MAC -eq "N/A") { return "Unknown" }
    $normalizedMac = ($MAC -replace '[^A-F0-9]','').ToUpper()
    if ($normalizedMac.Length -lt 6) { return "Unknown" }

    if ($script:MacVendorCache.ContainsKey($normalizedMac)) {
        return $script:MacVendorCache[$normalizedMac]
    }

    $firstByte = [Convert]::ToInt32($normalizedMac.Substring(0, 2), 16)
    if (($firstByte -band 2) -eq 2) {
        $script:MacVendorCache[$normalizedMac] = "Randomized / Private MAC"
        return $script:MacVendorCache[$normalizedMac]
    }

    $oui = $normalizedMac.Substring(0, 6)

    $vendorMap = @{
        "000C29" = "VMware"
        "001C42" = "Parallels"
        "080027" = "Oracle VirtualBox"
        "001A11" = "Google"
        "002248" = "Microsoft"
        "3C5AB4" = "Microsoft Surface"
        "A0C589" = "Intel"
        "8CFE74" = "Apple"
        "3C15C2" = "Apple"
        "DC9B9C" = "Apple"
        "4865EE" = "HP"
        "38EAA7" = "HP"
        "001B78" = "Dell"
        "848506" = "Dell"
        "001560" = "Lenovo"
        "54E1AD" = "Lenovo"
        "00E04C" = "Realtek"
        "001320" = "Cisco"
        "00E0A3" = "Cisco"
        "94D469" = "Ubiquiti"
        "24A43C" = "Ubiquiti"
        "B4FBE4" = "Netgear"
        "C0A0BB" = "Netgear"
        "20AA4B" = "TP-Link"
        "F4F26D" = "TP-Link"
        "BCEE7B" = "Asus"
        "107B44" = "Asus"
    }

    if ($vendorMap.ContainsKey($oui)) {
        $script:MacVendorCache[$normalizedMac] = $vendorMap[$oui]
        return $script:MacVendorCache[$normalizedMac]
    }

    try {
        $lookupResult = Invoke-RestMethod -Uri "https://api.macvendors.com/$normalizedMac" -TimeoutSec 3 -ErrorAction Stop
        if ($lookupResult) {
            $script:MacVendorCache[$normalizedMac] = [string]$lookupResult
            return $script:MacVendorCache[$normalizedMac]
        }
    } catch { }

    $script:MacVendorCache[$normalizedMac] = "Unknown (OUI: $oui)"
    return $script:MacVendorCache[$normalizedMac]
}

function Get-HostnameInference {
    param([string]$Hostname)

    $normalized = ([string]$Hostname).ToUpper()
    if (-not $normalized) {
        return [PSCustomObject]@{ VendorHint = $null; DeviceTypeHint = $null; DeviceRoleHint = $null; Reason = "No hostname available." }
    }

    if ($normalized -match 'IPHONE|IOS|MACBOOK|IPAD') {
        return [PSCustomObject]@{ VendorHint = "Apple"; DeviceTypeHint = "Mobile Device"; DeviceRoleHint = "Client Endpoint"; Reason = "Hostname pattern indicates Apple mobile/client hardware." }
    }
    if ($normalized -match 'MOTO|MOTOROLA|ANDROID|GALAXY|PIXEL|ONEPLUS|REDMI|XIAOMI|OPPO|VIVO|POCO') {
        return [PSCustomObject]@{ VendorHint = "Mobile OEM"; DeviceTypeHint = "Mobile Device"; DeviceRoleHint = "Client Endpoint"; Reason = "Hostname pattern indicates Android mobile hardware." }
    }
    if ($normalized -match 'LAPTOP|NOTEBOOK|ULTRABOOK|SURFACE|THINKPAD') {
        return [PSCustomObject]@{ VendorHint = $null; DeviceTypeHint = "Laptop"; DeviceRoleHint = "Client Endpoint"; Reason = "Hostname pattern indicates laptop form factor." }
    }
    if ($normalized -match 'DESKTOP|WORKSTATION|^WS-|^PC-|-PC$') {
        return [PSCustomObject]@{ VendorHint = $null; DeviceTypeHint = "Desktop"; DeviceRoleHint = "Client Endpoint"; Reason = "Hostname pattern indicates desktop/workstation endpoint." }
    }
    if ($normalized -match 'NAS|QNAP|SYNOLOGY') {
        return [PSCustomObject]@{ VendorHint = $null; DeviceTypeHint = "NAS / File Storage"; DeviceRoleHint = "Storage Server"; Reason = "Hostname pattern indicates NAS device." }
    }

    return [PSCustomObject]@{ VendorHint = $null; DeviceTypeHint = $null; DeviceRoleHint = $null; Reason = "No specific hostname signature matched." }
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER: CLASSIFY DEVICE
# ─────────────────────────────────────────────────────────────────────────────
function Get-DeviceClassification {
    param(
        [string]$IPAddress,
        [string]$Hostname,
        [string]$Vendor,
        [string]$OpenPorts,
        [PSCustomObject]$LocalDeviceProfile
    )

    $normalizedHostname = ([string]$Hostname).ToUpper()
    $hostnameInference = Get-HostnameInference -Hostname $normalizedHostname
    $ports = @($OpenPorts -split ',' | Where-Object { $_ })
    $hasPort = {
        param([string[]]$Candidates)
        @($ports | Where-Object { $Candidates -contains $_ }).Count -gt 0
    }
    $isGatewayLike = ($IPAddress -match '\.(1|254)$')
    $isLocal = ($LocalDeviceProfile -and @($LocalDeviceProfile.IPAddresses | Where-Object { $_ -eq $IPAddress }).Count -gt 0)

    if ($isLocal) {
        return [PSCustomObject]@{
            DeviceType      = $LocalDeviceProfile.DeviceType
            DeviceRole      = if ($LocalDeviceProfile.IsServer) { "Server" } else { "Client Endpoint" }
            Confidence      = "High"
            DetectionReason = "Matched local IP address and local system profile ($($LocalDeviceProfile.Manufacturer) $($LocalDeviceProfile.Model), $($LocalDeviceProfile.OSName))."
        }
    }

    if ($hostnameInference.DeviceTypeHint -eq "Mobile Device") {
        return [PSCustomObject]@{
            DeviceType      = "Mobile Device"
            DeviceRole      = "Client Endpoint"
            Confidence      = "High"
            DetectionReason = $hostnameInference.Reason
        }
    }

    if (& $hasPort @('9100','631','515') -or $normalizedHostname -match 'PRINT|PRN|MFP|EPSON|RICOH|CANON|BROTHER') {
        return [PSCustomObject]@{
            DeviceType      = "Printer"
            DeviceRole      = "Peripheral"
            Confidence      = "High"
            DetectionReason = "Printing-related ports or printer hostname signature detected."
        }
    }

    if ($hostnameInference.DeviceTypeHint -eq "NAS / File Storage") {
        return [PSCustomObject]@{
            DeviceType      = "NAS / File Storage"
            DeviceRole      = "Storage Server"
            Confidence      = "High"
            DetectionReason = $hostnameInference.Reason
        }
    }

    if ($Vendor -match 'Cisco|Ubiquiti|Netgear|TP-Link|Asus|D-Link|Linksys|Juniper|Fortinet|MikroTik|Sophos|Palo Alto') {
        if ($normalizedHostname -match 'FW|FIREWALL|UTM') {
            return [PSCustomObject]@{
                DeviceType      = "Firewall"
                DeviceRole      = "Security Appliance"
                Confidence      = "High"
                DetectionReason = "Network-security vendor signature combined with firewall hostname pattern."
            }
        }
        if ($normalizedHostname -match 'SW|SWITCH') {
            return [PSCustomObject]@{
                DeviceType      = "Switch"
                DeviceRole      = "Network Infrastructure"
                Confidence      = "High"
                DetectionReason = "Network vendor signature combined with switch hostname pattern."
            }
        }
        if ($normalizedHostname -match 'AP|WIFI|WIRELESS') {
            return [PSCustomObject]@{
                DeviceType      = "Access Point"
                DeviceRole      = "Network Infrastructure"
                Confidence      = "High"
                DetectionReason = "Network vendor signature combined with wireless/access-point hostname pattern."
            }
        }

        if ($isGatewayLike -or (& $hasPort @('53','67','80','443','22'))) {
            return [PSCustomObject]@{
                DeviceType      = "Router / Network Device"
                DeviceRole      = "Network Infrastructure"
                Confidence      = "High"
                DetectionReason = "Networking vendor signature with gateway-like IP/ports strongly indicates routing infrastructure."
            }
        }

        return [PSCustomObject]@{
            DeviceType      = "Router / Network Device"
            DeviceRole      = "Network Infrastructure"
            Confidence      = "Medium"
            DetectionReason = "Vendor maps to a networking manufacturer."
        }
    }

    if ($normalizedHostname -match 'NAS|FILE|STOR|BACKUP|QNAP|SYNOLOGY') {
        return [PSCustomObject]@{
            DeviceType      = "NAS / File Storage"
            DeviceRole      = "Storage Server"
            Confidence      = "Medium"
            DetectionReason = "Storage-oriented hostname signature detected."
        }
    }

    $serverInfraPortCount = @($ports | Where-Object { @('53','67','88','389','636','3268','3269') -contains $_ }).Count
    if ($serverInfraPortCount -ge 2) {
        return [PSCustomObject]@{
            DeviceType      = "Server"
            DeviceRole      = "Infrastructure Server"
            Confidence      = "High"
            DetectionReason = "Multiple identity/infrastructure server ports are open."
        }
    }

    if (& $hasPort @('80','443','8080','8443') -and $normalizedHostname -match 'SRV|SERVER|WEB|APP') {
        return [PSCustomObject]@{
            DeviceType      = "Server"
            DeviceRole      = "Web / Application Server"
            Confidence      = "High"
            DetectionReason = "Web-service ports are open and hostname resembles a server naming convention."
        }
    }

    if ($normalizedHostname -match 'SRV|SERVER|DC|DB|SQL|APP|WEB') {
        return [PSCustomObject]@{
            DeviceType      = "Server"
            DeviceRole      = "General Server"
            Confidence      = "Medium"
            DetectionReason = "Hostname follows common server naming convention."
        }
    }

    if ($hostnameInference.DeviceTypeHint -eq "Laptop" -or $normalizedHostname -match 'LAPTOP|NOTEBOOK|NB|BOOK|SURFACE') {
        return [PSCustomObject]@{
            DeviceType      = "Laptop"
            DeviceRole      = "Client Endpoint"
            Confidence      = "High"
            DetectionReason = "Hostname matches common laptop naming patterns."
        }
    }

    if ($hostnameInference.DeviceTypeHint -eq "Desktop" -or $normalizedHostname -match 'DESKTOP|PC|DESK|WORKSTATION|WS') {
        return [PSCustomObject]@{
            DeviceType      = "Desktop"
            DeviceRole      = "Client Endpoint"
            Confidence      = "High"
            DetectionReason = "Hostname matches common desktop or workstation naming patterns."
        }
    }

    if (& $hasPort @('135','139','445','3389','5985','5986')) {
        return [PSCustomObject]@{
            DeviceType      = if ($hostnameInference.DeviceTypeHint) { $hostnameInference.DeviceTypeHint } else { "Desktop" }
            DeviceRole      = "Windows Endpoint"
            Confidence      = "Medium"
            DetectionReason = "Windows client-management / SMB ports are open, but no strong server indicators were found."
        }
    }

    if ($Vendor -eq 'Randomized / Private MAC') {
        return [PSCustomObject]@{
            DeviceType      = if ($hostnameInference.DeviceTypeHint) { $hostnameInference.DeviceTypeHint } else { "Mobile Device" }
            DeviceRole      = "Likely Mobile / Private Endpoint"
            Confidence      = "Medium"
            DetectionReason = "Locally administered MAC address detected; vendor cannot be reliably inferred."
        }
    }

    if ($isGatewayLike -and (& $hasPort @('53','67','80','443'))) {
        return [PSCustomObject]@{
            DeviceType      = "Router / Network Device"
            DeviceRole      = "Network Infrastructure"
            Confidence      = "Medium"
            DetectionReason = "Gateway-like IP with infrastructure ports open."
        }
    }

    return [PSCustomObject]@{
        DeviceType      = "Client Device"
        DeviceRole      = "Client Endpoint"
        Confidence      = "Low"
        DetectionReason = "Weak evidence profile; defaulting to client endpoint classification."
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER: QUICK PORT SCAN
# ─────────────────────────────────────────────────────────────────────────────
function Test-CommonPorts {
    param([string]$IP)
    $commonPorts = @(22, 53, 67, 80, 88, 135, 139, 389, 443, 445, 515, 631, 636, 3389, 5985, 5986, 8080, 8443, 9100)
    $openPorts = @()
    foreach ($port in $commonPorts) {
        try {
            $tcpClient = New-Object System.Net.Sockets.TcpClient
            $connect   = $tcpClient.BeginConnect($IP, $port, $null, $null)
            $wait      = $connect.AsyncWaitHandle.WaitOne(300, $false)
            if ($wait -and $tcpClient.Connected) {
                $openPorts += $port
            }
            $tcpClient.Close()
        } catch { }
    }
    return ($openPorts -join ", ")
}

function Get-InventoryDeviceSummary {
    param(
        [string]$DeviceType,
        [string]$DeviceRole,
        [string]$Vendor,
        [string]$Hostname,
        [string]$MACAddress,
        [string]$OpenPorts,
        [string]$Confidence
    )

    $typeText = if ($DeviceType) { $DeviceType } else { "Unknown" }
    $roleText = if ($DeviceRole) { $DeviceRole } else { "Unclassified" }
    $vendorText = if ($Vendor) { $Vendor } else { "Unknown" }
    $hostText = if ($Hostname) { $Hostname } else { "N/A" }
    $macText = if ($MACAddress) { $MACAddress } else { "N/A" }
    $portsText = if ($OpenPorts) { $OpenPorts } else { "None" }
    $confText = if ($Confidence) { $Confidence } else { "Low" }

    return "Type: $typeText | Role: $roleText | Vendor: $vendorText | Host: $hostText | MAC: $macText | Ports: $portsText | Confidence: $confText"
}

# ─────────────────────────────────────────────────────────────────────────────
# DETECT SERVER ROLES (LOCAL MACHINE)
# ─────────────────────────────────────────────────────────────────────────────
function Get-LocalServerRoles {
    param([PSCustomObject]$LocalDeviceProfile)
    Write-Log "Detecting local server roles..." "INFO"
    $roles = @()
    if (-not $LocalDeviceProfile.IsServer) {
        return @([PSCustomObject]@{
            System = $LocalDeviceProfile.Hostname
            Role   = "Not Applicable (Client OS)"
            Detail = "Local machine runs a client/workstation OS, so server-role detection is intentionally skipped."
        })
    }

    try {
        try {
            $dcService = Get-Service -Name "NTDS" -ErrorAction SilentlyContinue
            if ($dcService -and $dcService.Status -eq "Running") {
                $roles += [PSCustomObject]@{ System = $LocalDeviceProfile.Hostname; Role = "Domain Controller"; Detail = "NTDS service is running." }
            }
        } catch { }
        try {
            $fsService = Get-Service -Name "LanmanServer" -ErrorAction SilentlyContinue
            if ($fsService -and $fsService.Status -eq "Running") {
                $roles += [PSCustomObject]@{ System = $LocalDeviceProfile.Hostname; Role = "File Server"; Detail = "Server service is running on a server OS." }
            }
        } catch { }
        try {
            $sqlService = Get-Service -Name "MSSQLSERVER" -ErrorAction SilentlyContinue
            if ($sqlService -and $sqlService.Status -eq "Running") {
                $roles += [PSCustomObject]@{ System = $LocalDeviceProfile.Hostname; Role = "Database Server (SQL)"; Detail = "SQL Server service is running." }
            }
        } catch { }
        try {
            $dhcpService = Get-Service -Name "DHCPServer" -ErrorAction SilentlyContinue
            if ($dhcpService -and $dhcpService.Status -eq "Running") {
                $roles += [PSCustomObject]@{ System = $LocalDeviceProfile.Hostname; Role = "DHCP Server"; Detail = "DHCP Server service is running." }
            }
        } catch { }
        try {
            $dnsService = Get-Service -Name "DNS" -ErrorAction SilentlyContinue
            if ($dnsService -and $dnsService.Status -eq "Running") {
                $roles += [PSCustomObject]@{ System = $LocalDeviceProfile.Hostname; Role = "DNS Server"; Detail = "DNS service is running." }
            }
        } catch { }
        try {
            $iisService = Get-Service -Name "W3SVC" -ErrorAction SilentlyContinue
            if ($iisService -and $iisService.Status -eq "Running") {
                $roles += [PSCustomObject]@{ System = $LocalDeviceProfile.Hostname; Role = "Web Server (IIS)"; Detail = "World Wide Web Publishing Service is running." }
            }
        } catch { }
    } catch {
        Write-Log "Could not detect server roles: $_" "WARN"
    }
    if ($roles.Count -eq 0) {
        $roles += [PSCustomObject]@{ System = $LocalDeviceProfile.Hostname; Role = "No Server Roles Detected"; Detail = "Server OS detected, but no supported server-role services are running." }
    }
    return $roles
}

# ─────────────────────────────────────────────────────────────────────────────
# DHCP ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────
function Get-DHCPAnalysis {
    param([array]$AliveHosts, [string]$NetworkCIDR)
    Write-Log "Performing DHCP analysis for $NetworkCIDR..." "INFO"

    $baseIP = ($NetworkCIDR -split '/')[0]
    $octets = $baseIP -split '\.'
    $prefix = "$($octets[0]).$($octets[1]).$($octets[2])"
    $totalHosts = 254

    $usedIPs      = $AliveHosts.Count
    $availableIPs = $totalHosts - $usedIPs
    $usedPercent  = [math]::Round(($usedIPs / $totalHosts) * 100, 1)

    return [PSCustomObject]@{
        Network       = $NetworkCIDR
        DHCPRange     = "$prefix.1 - $prefix.254"
        TotalHosts    = $totalHosts
        UsedIPs       = $usedIPs
        AvailableIPs  = $availableIPs
        UsedPercent   = $usedPercent
        Status        = if ($usedPercent -gt 90) { "CRITICAL" } elseif ($usedPercent -gt 75) { "WARNING" } else { "OK" }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SUBNET ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────
function Get-SubnetAnalysis {
    param([string]$NetworkCIDR)
    Write-Log "Performing subnet analysis for $NetworkCIDR..." "INFO"

    $baseIP = ($NetworkCIDR -split '/')[0]
    $prefix = [int]($NetworkCIDR -split '/')[1]

    # Calculate subnet details
    $totalAddresses = [math]::Pow(2, (32 - $prefix))
    $usableHosts    = $totalAddresses - 2
    $subnetMask     = switch ($prefix) {
        24 { "255.255.255.0" }
        23 { "255.255.254.0" }
        22 { "255.255.252.0" }
        16 { "255.255.0.0" }
        8  { "255.0.0.0" }
        default {
            $maskBits = ("1" * $prefix).PadRight(32,"0")
            $octets   = @()
            for ($i = 0; $i -lt 4; $i++) {
                $octets += [Convert]::ToInt32($maskBits.Substring($i*8,8), 2)
            }
            $octets -join "."
        }
    }

    $parts  = $baseIP -split '\.'
    $bcastP = "$($parts[0]).$($parts[1]).$($parts[2]).255"

    return [PSCustomObject]@{
        Network        = $NetworkCIDR
        SubnetMask     = $subnetMask
        PrefixLength   = $prefix
        NetworkAddress = $baseIP
        BroadcastAddress = $bcastP
        TotalAddresses = [int]$totalAddresses
        UsableHosts    = [int]$usableHosts
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# DUPLICATE IP DETECTION
# ─────────────────────────────────────────────────────────────────────────────
function Find-DuplicateIPs {
    param([array]$Devices)
    Write-Log "Checking for duplicate IP addresses..." "INFO"
    $duplicates = $Devices | Group-Object -Property IPAddress | Where-Object { $_.Count -gt 1 }
    if ($duplicates) {
        Write-Log "WARNING: Duplicate IPs detected: $(($duplicates | ForEach-Object {$_.Name}) -join ', ')" "WARN"
    }
    return $duplicates
}

# ─────────────────────────────────────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────────────────────────────────────
function Export-DevicesToExcel {
    param(
        [array]$Devices,
        [string]$OutputPath
    )
    Write-Log "Exporting devices to Excel: $OutputPath" "INFO"
    try {
        if (Get-Module -ListAvailable -Name ImportExcel) {
            Import-Module ImportExcel -ErrorAction Stop
            $Devices | Select-Object IPAddress, Hostname, DeviceType, DeviceRole, Vendor, MACAddress, OpenPorts, DetectionConfidence, DetectionReason, DeviceSummary, Status, Network |
                Export-Excel -Path $OutputPath `
                             -WorksheetName "NetworkDevices" `
                             -TableName "DeviceInventory" `
                             -AutoSize `
                             -BoldTopRow `
                             -FreezeTopRow `
                             -Title "InfraEye Network Device Inventory" `
                             -TitleBold `
                             -TitleSize 14
            Write-Log "Excel export completed: $OutputPath" "SUCCESS"
            return $OutputPath
        } else {
            Write-Log "ImportExcel module not available. Exporting to CSV instead." "WARN"
            $csvPath = $OutputPath -replace '\.xlsx$', '.csv'
            $Devices | Select-Object IPAddress, Hostname, DeviceType, DeviceRole, Vendor, MACAddress, OpenPorts, DetectionConfidence, DetectionReason, DeviceSummary, Status, Network |
                Export-Csv -Path $csvPath -NoTypeInformation
            Write-Log "CSV export completed: $csvPath" "SUCCESS"
            return $csvPath
        }
    } catch {
        Write-Log "Excel export failed: $_" "WARN"
    }
    return $null
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT GENERATION
# ─────────────────────────────────────────────────────────────────────────────
function New-HtmlReport {
    param(
        [array]$Devices,
        [array]$DHCPData,
        [array]$SubnetData,
        [array]$ServerRoles,
        [array]$DuplicateIPs,
        [PSCustomObject]$LocalDeviceProfile
    )

    $reportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $deviceRows = ""
    foreach ($d in $Devices) {
        $statusColor = if ($d.Status -eq "Online") { "#28a745" } else { "#dc3545" }
        $typeColor   = switch -Wildcard ($d.DeviceType) {
            "*Server*"   { "#dc3545" }
            "Workstation"{ "#0d6efd" }
            "Printer"    { "#6f42c1" }
            "*Router*"   { "#fd7e14" }
            "*Switch*"   { "#fd7e14" }
            "Firewall"   { "#dc3545" }
            default      { "#6c757d" }
        }
        $deviceRows += @"
            <tr>
                <td>$($d.IPAddress)</td>
                <td>$($d.Hostname)</td>
                <td><span style='background:$typeColor;color:#fff;padding:2px 8px;border-radius:10px;font-size:0.8em;'>$($d.DeviceType)</span></td>
                <td>$($d.DeviceRole)</td>
                <td>$($d.Vendor)</td>
                <td style='font-family:monospace;'>$($d.MACAddress)</td>
                <td style='font-family:monospace;'>$($d.OpenPorts)</td>
                <td>$($d.DetectionConfidence)</td>
                <td>$($d.DetectionReason)</td>
                <td><span style='background:$statusColor;color:#fff;padding:2px 8px;border-radius:10px;font-size:0.8em;'>$($d.Status)</span></td>
            </tr>
"@
    }

    $dhcpRows = ""
    foreach ($dh in $DHCPData) {
        $dhcpStatusColor = switch ($dh.Status) { "CRITICAL" { "#dc3545" } "WARNING" { "#ffc107" } default { "#28a745" } }
        $dhcpRows += "<tr><td>$($dh.Network)</td><td>$($dh.DHCPRange)</td><td>$($dh.TotalHosts)</td><td>$($dh.UsedIPs)</td><td>$($dh.AvailableIPs)</td><td>$($dh.UsedPercent)%</td><td><span style='background:$dhcpStatusColor;color:#fff;padding:2px 8px;border-radius:10px;font-size:0.8em;'>$($dh.Status)</span></td></tr>"
    }

    $subnetRows = ""
    foreach ($s in $SubnetData) {
        $subnetRows += "<tr><td>$($s.Network)</td><td>$($s.SubnetMask)</td><td>/$($s.PrefixLength)</td><td>$($s.NetworkAddress)</td><td>$($s.BroadcastAddress)</td><td>$($s.TotalAddresses)</td><td>$($s.UsableHosts)</td></tr>"
    }

    $roleRows = ""
    foreach ($r in $ServerRoles) {
        $roleRows += "<tr><td>$($r.System)</td><td>$($r.Role)</td><td>$($r.Detail)</td></tr>"
    }

    $localDeviceInfoHtml = @"
        <div class='section'>
            <h2>&#x1F4BB; Local Device Classification</h2>
            <table>
                <thead><tr><th>Hostname</th><th>Device Type</th><th>Manufacturer</th><th>Model</th><th>OS</th><th>IP Addresses</th></tr></thead>
                <tbody>
                    <tr>
                        <td>$($LocalDeviceProfile.Hostname)</td>
                        <td>$($LocalDeviceProfile.DeviceType)</td>
                        <td>$($LocalDeviceProfile.Manufacturer)</td>
                        <td>$($LocalDeviceProfile.Model)</td>
                        <td>$($LocalDeviceProfile.OSName)</td>
                        <td>$((@($LocalDeviceProfile.IPAddresses) -join ', '))</td>
                    </tr>
                </tbody>
            </table>
        </div>
"@

    $dupWarning = if ($DuplicateIPs.Count -gt 0) {
        "<div style='background:#dc3545;color:#fff;padding:12px 20px;border-radius:8px;margin-bottom:15px;'><strong>&#x26A0; WARNING:</strong> Duplicate IP addresses detected: $(($DuplicateIPs | ForEach-Object {$_.Name}) -join ', ')</div>"
    } else { "" }

    # Device type breakdown
    $typeBreakdown = $Devices | Group-Object DeviceType | Sort-Object Count -Descending
    $typeLabels    = ($typeBreakdown | ForEach-Object { "'$($_.Name)'" }) -join ","
    $typeCounts    = ($typeBreakdown | ForEach-Object { $_.Count }) -join ","

    # Online vs offline
    $onlineCount  = @($Devices | Where-Object { $_.Status -eq "Online" }).Count
    $offlineCount = @($Devices | Where-Object { $_.Status -ne "Online" }).Count

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>InfraEye - Infrastructure Discovery Report</title>
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
        .summary-grid { display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:20px;margin-bottom:30px; }
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
        footer { text-align:center;padding:30px 20px;color:var(--text-muted);font-size:0.85em;border-top:1px solid var(--border);margin-top:20px; }
        footer a { color:var(--accent);text-decoration:none; }
        .search-box { padding:8px 14px;border:1px solid var(--border);border-radius:8px;background:var(--bg);color:var(--text);font-size:0.9em;width:100%;max-width:400px;margin-bottom:15px; }
    </style>
</head>
<body>
<header>
    <div>
        <h1>&#x1F5A7; InfraEye &mdash; Infrastructure Discovery Report</h1>
        <p>Generated: $reportDate &nbsp;|&nbsp; Total Devices: $($Devices.Count)</p>
    </div>
    <button class="toggle-btn" onclick="toggleDarkMode()">&#9790; Dark Mode</button>
</header>
<div class="container">
    $dupWarning

    <!-- Summary Cards -->
    <div class="summary-grid">
        <div class="card">
            <h3>Total Devices</h3>
            <div class="value">$($Devices.Count)</div>
            <div class="sub">Discovered on network</div>
        </div>
        <div class="card">
            <h3>Online</h3>
            <div class="value" style="color:#28a745;">$onlineCount</div>
            <div class="sub">Responding to ping</div>
        </div>
        <div class="card">
            <h3>Servers</h3>
            <div class="value" style="color:#dc3545;">$(@($Devices | Where-Object {$_.DeviceType -match 'Server'}).Count)</div>
            <div class="sub">Detected servers</div>
        </div>
        <div class="card">
            <h3>Client Devices</h3>
            <div class="value" style="color:#0d6efd;">$(@($Devices | Where-Object {$_.DeviceType -match 'Laptop|Desktop|Workstation|Client Device|Mobile Device|Tablet'}).Count)</div>
            <div class="sub">Detected user endpoints</div>
        </div>
        <div class="card">
            <h3>Network Devices</h3>
            <div class="value" style="color:#fd7e14;">$(@($Devices | Where-Object {$_.DeviceType -match 'Router|Switch|Firewall'}).Count)</div>
            <div class="sub">Routers, switches, firewalls</div>
        </div>
        <div class="card">
            <h3>Uncertain Classifications</h3>
            <div class="value" style="color:#6c757d;">$(@($Devices | Where-Object { $_.DetectionConfidence -eq 'Low' -or $_.Vendor -like 'Unknown*' }).Count)</div>
            <div class="sub">Need manual verification</div>
        </div>
    </div>

    <!-- Charts -->
    <div class="charts-grid">
        <div class="chart-card">
            <h2>Device Types</h2>
            <canvas id="typeChart" height="200"></canvas>
        </div>
        <div class="chart-card">
            <h2>Online vs Offline</h2>
            <canvas id="statusChart" height="200"></canvas>
        </div>
        <div class="chart-card">
            <h2>DHCP Pool Usage</h2>
            <canvas id="dhcpChart" height="200"></canvas>
        </div>
    </div>

    <!-- Device Inventory Table -->
    $localDeviceInfoHtml

    <div class="section">
        <h2>&#x1F4CB; Network Device Inventory</h2>
        <input type="text" id="deviceSearch" class="search-box" placeholder="Search devices..." onkeyup="filterTable()">
        <div style="overflow-x:auto;">
        <table id="deviceTable">
            <thead><tr><th>IP Address</th><th>Hostname</th><th>Device Type</th><th>Role</th><th>Vendor</th><th>MAC Address</th><th>Open Ports</th><th>Confidence</th><th>Detection Basis</th><th>Status</th></tr></thead>
            <tbody>$deviceRows</tbody>
        </table>
        </div>
    </div>

    <!-- DHCP Analysis -->
    <div class="section">
        <h2>&#x1F4E1; DHCP Pool Analysis</h2>
        <table>
            <thead><tr><th>Network</th><th>DHCP Range</th><th>Total Hosts</th><th>Used IPs</th><th>Available IPs</th><th>Used %</th><th>Status</th></tr></thead>
            <tbody>$dhcpRows</tbody>
        </table>
    </div>

    <!-- Subnet Analysis -->
    <div class="section">
        <h2>&#x1F310; Subnet Analysis</h2>
        <table>
            <thead><tr><th>Network</th><th>Subnet Mask</th><th>Prefix</th><th>Network Addr</th><th>Broadcast</th><th>Total Addr</th><th>Usable Hosts</th></tr></thead>
            <tbody>$subnetRows</tbody>
        </table>
    </div>

    <!-- Server Roles -->
    <div class="section">
        <h2>&#x1F5FC; Detected Server Roles (Local Machine - Server OS Only)</h2>
        <table>
            <thead><tr><th>System</th><th>Role</th><th>Details</th></tr></thead>
            <tbody>$roleRows</tbody>
        </table>
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

    function filterTable() {
        const input = document.getElementById('deviceSearch');
        const filter = input.value.toUpperCase();
        const table  = document.getElementById('deviceTable');
        const rows   = table.getElementsByTagName('tr');
        for (let i = 1; i < rows.length; i++) {
            const cells = rows[i].getElementsByTagName('td');
            let match = false;
            for (let j = 0; j < cells.length; j++) {
                if (cells[j].textContent.toUpperCase().includes(filter)) { match = true; break; }
            }
            rows[i].style.display = match ? '' : 'none';
        }
    }

    const typeLabels = [$typeLabels];
    const typeCounts = [$typeCounts];
    new Chart(document.getElementById('typeChart'), {
        type: 'doughnut',
        data: {
            labels: typeLabels,
            datasets: [{ data: typeCounts,
                backgroundColor: ['#dc3545','#0d6efd','#28a745','#fd7e14','#6f42c1','#6c757d','#17a2b8','#ffc107'],
                borderWidth: 2 }]
        },
        options: { responsive: true, plugins: { legend: { position: 'right', labels: { font: { size: 11 } } } } }
    });

    new Chart(document.getElementById('statusChart'), {
        type: 'doughnut',
        data: {
            labels: ['Online', 'Offline'],
            datasets: [{ data: [$onlineCount, $offlineCount],
                backgroundColor: ['#28a745','#dc3545'], borderWidth: 2 }]
        },
        options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
    });

    const dhcpUsed = [$(($DHCPData | ForEach-Object {$_.UsedPercent}) -join ',')];
    const dhcpLabels = [$(($DHCPData | ForEach-Object {"'$($_.Network)'"}) -join ',')];
    new Chart(document.getElementById('dhcpChart'), {
        type: 'bar',
        data: {
            labels: dhcpLabels,
            datasets: [{
                label: 'DHCP Used %',
                data: dhcpUsed,
                backgroundColor: dhcpUsed.map(v => v > 90 ? '#dc3545' : v > 75 ? '#ffc107' : '#28a745'),
                borderRadius: 6
            }]
        },
        options: { responsive: true, scales: { y: { min: 0, max: 100 } }, plugins: { legend: { display: false } } }
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
    Write-Log "=== InfraEye Infrastructure Discovery Started ===" "INFO"

    Install-RequiredModules

    $localDeviceProfile = Get-LocalDeviceProfile
    $networkRanges = Get-LocalNetworkRanges
    $allDevices    = @()
    $dhcpDataList  = @()
    $subnetDataList= @()
    $inventoryExportPath = $null

    foreach ($network in $networkRanges) {
        Write-Log "Scanning network: $network" "INFO"

        $aliveHosts = @()
        try {
            $aliveHosts = Invoke-PingSweep -NetworkCIDR $network
        } catch {
            Write-Log "Ping sweep failed for $network`: $_" "WARN"
            continue
        }

        # Get ARP table for MAC addresses
        $arpTable = Get-ARPTable

        foreach ($ip in $aliveHosts) {
            try {
                $isLocalDevice = @($localDeviceProfile.IPAddresses | Where-Object { $_ -eq $ip }).Count -gt 0
                $hostname  = Resolve-HostnameForIP -IPAddress $ip
                if ($isLocalDevice) {
                    $hostname = $localDeviceProfile.Hostname
                }

                $mac = if ($isLocalDevice -and $localDeviceProfile.IPToMACMap.ContainsKey($ip)) {
                    $localDeviceProfile.IPToMACMap[$ip]
                } elseif ($arpTable.ContainsKey($ip)) {
                    $arpTable[$ip]
                } else {
                    "N/A"
                }

                $vendor = if ($isLocalDevice) {
                    $localDeviceProfile.Manufacturer
                } else {
                    Get-VendorFromMAC -MAC $mac
                }

                if (-not $isLocalDevice -and ($vendor -eq "Unknown" -or $vendor -like "Unknown*" -or $vendor -eq "Randomized / Private MAC")) {
                    $hostnameInference = Get-HostnameInference -Hostname $hostname
                    if ($hostnameInference.VendorHint -eq "Mobile OEM" -and $hostname -match 'MOTO|MOTOROLA') {
                        $vendor = "Motorola (Hostname Inference)"
                    } elseif ($hostnameInference.VendorHint -eq "Mobile OEM") {
                        $vendor = "Android OEM (Hostname Inference)"
                    } elseif ($hostnameInference.VendorHint) {
                        $vendor = "$($hostnameInference.VendorHint) (Hostname Inference)"
                    }
                }

                $openPorts = ""
                try { $openPorts = Test-CommonPorts -IP $ip } catch { }
                $classification = Get-DeviceClassification -IPAddress $ip -Hostname $hostname -Vendor $vendor -OpenPorts $openPorts -LocalDeviceProfile $localDeviceProfile
                $summaryText = Get-InventoryDeviceSummary -DeviceType $classification.DeviceType -DeviceRole $classification.DeviceRole -Vendor $vendor -Hostname $hostname -MACAddress $mac -OpenPorts $openPorts -Confidence $classification.Confidence

                $allDevices += [PSCustomObject]@{
                    IPAddress  = $ip
                    Hostname   = $hostname
                    DeviceType = $classification.DeviceType
                    DeviceRole = $classification.DeviceRole
                    Vendor     = $vendor
                    MACAddress = $mac
                    Status     = "Online"
                    Network    = $network
                    OpenPorts  = $openPorts
                    DetectionConfidence = $classification.Confidence
                    DetectionReason = $classification.DetectionReason
                    DeviceSummary = $summaryText
                }
            } catch {
                Write-Log "Error processing host $ip`: $_" "WARN"
                $allDevices += [PSCustomObject]@{
                    IPAddress  = $ip
                    Hostname   = $ip
                    DeviceType = "Unknown Device"
                    DeviceRole = "Unclassified"
                    Vendor     = "Unknown"
                    MACAddress = "N/A"
                    Status     = "Online"
                    Network    = $network
                    OpenPorts  = ""
                    DetectionConfidence = "Low"
                    DetectionReason = "Unhandled discovery error during classification."
                    DeviceSummary = Get-InventoryDeviceSummary -DeviceType "Unknown Device" -DeviceRole "Unclassified" -Vendor "Unknown" -Hostname $ip -MACAddress "N/A" -OpenPorts "" -Confidence "Low"
                }
            }
        }

        # ARP-visible hosts: appeared in ARP table but did not respond to ICMP ping
        # These devices may have firewalls blocking ping but are still reachable on the LAN
        $netParts  = (($network -split '/')[0]).Split('.')
        $cidrBits  = [int](($network -split '/')[1])
        $netPrefix = if ($cidrBits -ge 24) { "$($netParts[0]).$($netParts[1]).$($netParts[2])." }
                     elseif ($cidrBits -ge 16) { "$($netParts[0]).$($netParts[1])." }
                     else { "$($netParts[0])." }

        $arpOnlyHosts = @($arpTable.Keys | Where-Object {
            $_.StartsWith($netPrefix) -and
            ($aliveHosts -notcontains $_) -and
            ($_ -notmatch '\.(0|255)$') -and
            $arpTable[$_] -notmatch '^FF:FF' -and
            $arpTable[$_] -ne '00:00:00:00:00:00'
        })

        if ($arpOnlyHosts.Count -gt 0) {
            Write-Log "Found $($arpOnlyHosts.Count) ARP-visible host(s) not responding to ICMP on $network" "INFO"
        }

        foreach ($ip in $arpOnlyHosts) {
            try {
                $isLocalDevice = @($localDeviceProfile.IPAddresses | Where-Object { $_ -eq $ip }).Count -gt 0
                $hostname = Resolve-HostnameForIP -IPAddress $ip
                if ($isLocalDevice) { $hostname = $localDeviceProfile.Hostname }

                $mac = if ($isLocalDevice -and $localDeviceProfile.IPToMACMap.ContainsKey($ip)) {
                    $localDeviceProfile.IPToMACMap[$ip]
                } else { $arpTable[$ip] }

                $vendor = if ($isLocalDevice) {
                    $localDeviceProfile.Manufacturer
                } else {
                    Get-VendorFromMAC -MAC $mac
                }

                $classification = Get-DeviceClassification -IPAddress $ip -Hostname $hostname -Vendor $vendor -OpenPorts "" -LocalDeviceProfile $localDeviceProfile
                $summaryText    = Get-InventoryDeviceSummary -DeviceType $classification.DeviceType -DeviceRole $classification.DeviceRole -Vendor $vendor -Hostname $hostname -MACAddress $mac -OpenPorts "" -Confidence "Low"

                $allDevices += [PSCustomObject]@{
                    IPAddress  = $ip
                    Hostname   = $hostname
                    DeviceType = $classification.DeviceType
                    DeviceRole = $classification.DeviceRole
                    Vendor     = $vendor
                    MACAddress = $mac
                    Status     = "ARP Visible (ICMP Blocked)"
                    Network    = $network
                    OpenPorts  = ""
                    DetectionConfidence = "Low"
                    DetectionReason = "Visible in ARP table but did not respond to ICMP. Firewall may block ping."
                    DeviceSummary = $summaryText
                }
            } catch {
                Write-Log "Error processing ARP-only host $ip`: $_" "WARN"
            }
        }

        # DHCP and Subnet analysis
        $dhcpDataList   += Get-DHCPAnalysis  -AliveHosts $aliveHosts -NetworkCIDR $network
        $subnetDataList += Get-SubnetAnalysis -NetworkCIDR $network
    }

    # Server roles (local)
    $serverRoles = @(Get-LocalServerRoles -LocalDeviceProfile $localDeviceProfile)

    # Duplicate IP check
    $duplicateIPs = @(Find-DuplicateIPs -Devices $allDevices)

    # Export to Excel
    if ($allDevices.Count -gt 0) {
        $inventoryExportPath = Export-DevicesToExcel -Devices $allDevices -OutputPath $ExcelFile
    }

    # Generate HTML report
    Write-Log "Generating HTML report..." "INFO"
    $htmlContent = New-HtmlReport `
        -Devices      $allDevices `
        -DHCPData     $dhcpDataList `
        -SubnetData   $subnetDataList `
        -ServerRoles  $serverRoles `
        -DuplicateIPs $duplicateIPs `
        -LocalDeviceProfile $localDeviceProfile

    $htmlContent | Out-File -FilePath $ReportFile -Encoding UTF8 -Force
    Write-Log "Report saved: $ReportFile" "SUCCESS"

    # Trigger alert if needed
    $alertScript = Join-Path $PSScriptRoot "Infrastructure_Alert.ps1"
    if (Test-Path $alertScript) {
        $dhcpCritical  = @($dhcpDataList | Where-Object { $_.Status -eq "CRITICAL" })
        $unknownDevices= @($allDevices   | Where-Object { $_.DetectionConfidence -eq "Low" -or $_.Vendor -like "Unknown*" })
        $dupCount      = $duplicateIPs.Count

        if ($dhcpCritical.Count -gt 0 -or $unknownDevices.Count -gt 0 -or $dupCount -gt 0) {
            Write-Log "Infrastructure alert conditions met. Triggering alert..." "WARN"
            & $alertScript `
                -ReportFile     $ReportFile `
                -InventoryFile  $inventoryExportPath `
                -DuplicateCount $dupCount `
                -UnknownDevices $unknownDevices.Count `
                -DHCPCritical   $dhcpCritical.Count
        }
    }

    Write-Log "Total devices discovered: $($allDevices.Count)" "SUCCESS"
    Write-Log "=== InfraEye Infrastructure Discovery Completed ===" "SUCCESS"
    Write-Host "`nReport generated: $ReportFile" -ForegroundColor Green
    if (Test-Path $ExcelFile) {
        Write-Host "Excel report: $ExcelFile" -ForegroundColor Green
    }
} catch {
    Write-Log "FATAL ERROR: $_" "ERROR"
    Write-Error $_
    exit 1
}
