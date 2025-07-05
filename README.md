# Network Status Reporter: PowerShell Script for IP & Website Checks with Excel Export

# Network Status Monitor 📡

This PowerShell script performs a comprehensive **network health check** by pinging devices, checking website availability, scanning CCTV IP ranges, and testing NVR connectivity. It exports the results into a **stylized Excel report**.

## Features ✨

- ✅ Ping test for critical infrastructure IPs
- 🌐 HTTP status check for websites
- 🎥 Range scan for CCTV cameras (192.168.10.x)
- 📺 NVR device ping test
- 📊 Excel report with:
  - Styled table (Light9 theme)
  - Conditional formatting (red font for offline)
  - Borders and no gridlines

## Requirements 🛠

- PowerShell module: `ImportExcel`

```
  Install-Module -Name ImportExcel -Scope CurrentUser
````
----

## ✅ Step-by-Step Fix

### 🔐 Step 1: Open PowerShell as Administrator

1. Press **Windows Key** → search for **PowerShell**.
2. Right-click on **Windows PowerShell** → click **"Run as administrator"**.

---

### ⚙️ Step 2: Check your current execution policy

Run:

```powershell
Get-ExecutionPolicy
```

It likely says: `Restricted` (which prevents all scripts from running).

---

### 🔓 Step 3: Temporarily allow scripts (safe for local use)

Run this command to **allow script execution just for the current user**:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force
```

✅ This enables you to run scripts **downloaded from the internet only if they are signed**, but allows local unsigned scripts (like `ImportExcel` module scripts).

---

### ✅ Step 4: Try again

Now run your PowerShell script again — starting with:

```powershell
Import-Module ImportExcel
```

Then proceed with the Excel export script.

---

## 🧠 Reminder

If you ever want to **revert** the setting later, use:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Restricted
```
---

### ✅ Step 5: PowerShell Script — `network_report.ps1`

```powershell

# Load the required module
Import-Module ImportExcel

# === IPs and Details ===

# List of critical infrastructure IPs
$ipAddresses = @(
    "192.168.5.10", "192.168.5.11", "192.168.5.205",
    "192.168.101.10", "192.168.101.1", "192.168.101.20",
    "192.168.5.250", "172.16.1.7"
)

# Metadata for each IP address
$ipDetails = @{
    "192.168.5.10"   = @{ Name = "Office Switch"; Location = "Server Room" }
    "192.168.5.11"   = @{ Name = "Core Router"; Location = "Rack 1" }
    "192.168.5.205"  = @{ Name = "Printer Room"; Location = "Floor 2" }
    "192.168.101.10" = @{ Name = "NAS Storage"; Location = "Backup Zone" }
    "192.168.101.1"  = @{ Name = "WiFi Controller"; Location = "Admin Office" }
    "192.168.101.20" = @{ Name = "Unused Device"; Location = "Unknown" }
    "192.168.5.250"  = @{ Name = "Firewall"; Location = "Server Rack" }
    "172.16.1.7"     = @{ Name = "Branch Link"; Location = "Branch Office" }
}

# List of websites to check
$websites = @(
    "https://chatgpt.com/",
    "https://www.facebook.com/",
    "https://www.instagram.com/",
    "https://www.xitizbasnet.com.np/",
    "https://golyan.com/our-businesses/shivam-plastic-industries-pvt-ltd/",
    "https://www.instagram.com/",
    "https://www.linkedin.com/in/xitizbasnet/",
    "https://golyan.com/"
)

# Camera IP range (192.168.10.25 to 192.168.10.35)
$cameraRange = 25..35
$cameraBase = "192.168.10."

# List of NVR devices
$nvrIPs = @("192.168.60.10", "192.168.60.20", "192.168.60.36")

# === FUNCTIONS ===

# Function to test ICMP ping
function Test-Ping {
    param($ip)
    $ping = New-Object System.Net.NetworkInformation.Ping
    try {
        ($ping.Send($ip, 1000)).Status -eq "Success"
    } catch { $false }
}

# Function to test website availability via HTTP HEAD request
function Test-Website {
    param($url)
    try {
        $res = Invoke-WebRequest -Uri $url -Method Head -TimeoutSec 5 -ErrorAction Stop
        return $res.StatusCode -ge 200 -and $res.StatusCode -lt 400
    } catch { $false }
}

# === MAIN LOGIC ===

$results = @()

# 1. Check all individual infrastructure IPs
foreach ($ip in $ipAddresses) {
    $status = if (Test-Ping $ip) { "Online" } else { "Offline" }
    $details = $ipDetails[$ip]
    $results += [PSCustomObject]@{
        Date     = (Get-Date -Format "yyyy-MM-dd")
        Time     = (Get-Date -Format "HH:mm:ss")
        Type     = "IP Ping"
        Name     = $details.Name
        Location = $details.Location
        Target   = $ip
        Status   = $status
    }
}

# 2. Check websites
foreach ($url in $websites) {
    $status = if (Test-Website $url) { "Online" } else { "Offline" }
    $results += [PSCustomObject]@{
        Date     = (Get-Date -Format "yyyy-MM-dd")
        Time     = (Get-Date -Format "HH:mm:ss")
        Type     = "Website"
        Name     = $url
        Location = "N/A"
        Target   = $url
        Status   = $status
    }
}

# 3. Check CCTV camera IPs
foreach ($i in $cameraRange) {
    $ip = "$cameraBase$i"
    $status = if (Test-Ping $ip) { "Online" } else { "Offline" }
    $results += [PSCustomObject]@{
        Date     = (Get-Date -Format "yyyy-MM-dd")
        Time     = (Get-Date -Format "HH:mm:ss")
        Type     = "Camera IP"
        Name     = "Camera $i"
        Location = "192.168.10 subnet"
        Target   = $ip
        Status   = $status
    }
}

# 4. Check NVRs
foreach ($ip in $nvrIPs) {
    $status = if (Test-Ping $ip) { "Online" } else { "Offline" }
    $results += [PSCustomObject]@{
        Date     = (Get-Date -Format "yyyy-MM-dd")
        Time     = (Get-Date -Format "HH:mm:ss")
        Type     = "NVR IP"
        Name     = "NVR Device"
        Location = "NVR subnet"
        Target   = $ip
        Status   = $status
    }
}

# === EXPORT TO EXCEL ===

# Output file path
$outputPath = [Environment]::GetFolderPath("Desktop") + "\full_network_report.xlsx"

# Export to Excel with basic formatting
$results | Export-Excel -Path $outputPath -WorksheetName "Network Status" `
    -TableName "NetworkReport" `
    -TableStyle Light9 `
    -AutoSize -AutoFilter

# Open Excel file to apply more formatting
$excel = Open-ExcelPackage -Path $outputPath
$ws = $excel.Workbook.Worksheets["Network Status"]

# Hide default Excel gridlines
$ws.View.ShowGridLines = $false

# Add borders around all cells
$range = $ws.Dimension.Address
$ws.Cells[$range].Style.Border.Top.Style = "Thin"
$ws.Cells[$range].Style.Border.Bottom.Style = "Thin"
$ws.Cells[$range].Style.Border.Left.Style = "Thin"
$ws.Cells[$range].Style.Border.Right.Style = "Thin"

# Apply red font color to "Offline" statuses
$rowCount = $ws.Dimension.End.Row
$statusCol = 7  # Column G = Status
for ($row = 2; $row -le $rowCount; $row++) {
    $cell = $ws.Cells[$row, $statusCol]
    if ($cell.Text -eq "Offline") {
        $cell.Style.Font.Color.SetColor([System.Drawing.Color]::Red)
    }
}

Close-ExcelPackage $excel

# Final Message
Write-Host "`n✅ Final Excel report created with:"
Write-Host "- Blue & white table (Light9)"
Write-Host "- Gridlines removed"
Write-Host "- Borders on all cells"
Write-Host "- Red text for Offline rows"
Write-Host "`n📄 Saved at: $outputPath";
```

---



## Output 📄

The script generates an Excel file on your desktop:

```
📁 full_network_report.xlsx
```
---
