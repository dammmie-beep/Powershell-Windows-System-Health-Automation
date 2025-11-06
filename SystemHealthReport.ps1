function Get-DiskSpaceReport {
    <#
    .SYNOPSIS
    Gathers disk space utilization data for all local fixed drives.
    
    .DESCRIPTION
    Uses Get-CimInstance to query Win32_LogicalDisk, calculates the free space
    percentage, and formats the output into GB for readability.
    #>
    
    Write-Host "--- Starting Disk Space Check ---" -ForegroundColor Cyan

    try {
        $Disks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType = 3" | 
            Select-Object @{Name='Drive'; Expression={$_.DeviceID}},
                          @{Name='TotalGB'; Expression={"{0:N2}" -f ($_.Size / 1GB)}},
                          @{Name='FreeGB'; Expression={"{0:N2}" -f ($_.FreeSpace / 1GB)}},
                          @{Name='UsedPercent'; Expression={"{0:N0}%" -f ((($_.Size - $_.FreeSpace) / $_.Size) * 100)}}

        # Output the data to the console in a clear table format
        # Change to force the formatted output to the console
        $Disks | Format-Table -AutoSize | Out-Host

        # Return the data as objects for later report compilation
        return $Disks

    }
    catch {
        Write-Error "Failed to retrieve disk information: $($_.Exception.Message)"
        return $null # Return null if an error occurs
    }
}

# --- Execution ---
# To run this part of the script, simply call the function:
# $DiskData = Get-DiskSpaceReport

function Get-ServiceStatusReport {
    <#
    .SYNOPSIS
    Checks the status of critical Windows services.
    
    .DESCRIPTION
    Checks essential services like the Event Log, Windows Update, and Print Spooler.
    #>

    Write-Host "`n--- Starting Service Status Check ---" -ForegroundColor Yellow

    # Define a list of critical services to check
    # You can customize this list based on the servers you administer
    $CriticalServices = @(
        "EventLog",      # Windows Event Log (Crucial for system health)
        "wuauserv",      # Windows Update
        "Spooler",       # Print Spooler
        "Dhcp",          # DHCP Client
        "LanmanServer"   # Server service (File/Print Sharing)
    )

    try {
        $ServiceData = Get-Service -Name $CriticalServices -ErrorAction SilentlyContinue |
            Select-Object Name, Status, DisplayName, StartType

        # Output the data to the console using Out-Host to force display
        $ServiceData | Format-Table -AutoSize | Out-Host

        # Return the data objects
        return $ServiceData
    }
    catch {
        Write-Error "Failed to retrieve service information: $($_.Exception.Message)"
        return $null
    }
}

function Get-EventLogReview {
    <#
    .SYNOPSIS
    Reviews the Application and System event logs for recent errors and warnings.
    
    .DESCRIPTION
    Filters logs for critical entries (Error and Warning) in the last 24 hours.
    #>

    Write-Host "`n--- Starting Event Log Review (Last 24 Hours) ---" -ForegroundColor Green

    # Define the time range (24 hours ago)
    $StartTime = (Get-Date).AddDays(-1)
    
    # Define the Event Log names and IDs to check (Error=2, Warning=3)
    $LogNames = @('System', 'Application')
    $SeverityIDs = @(2, 3) # 2=Error, 3=Warning

    try {
        # Using a structured query (hashtable) is more efficient than filtering objects later
        $FilterHashTable = @{
            LogName = $LogNames
            Level = $SeverityIDs
            StartTime = $StartTime
        }

        $EventData = Get-WinEvent -FilterHashTable $FilterHashTable -ErrorAction SilentlyContinue |
            Select-Object TimeCreated, ID, LevelDisplayName, ProviderName, Message -First 20

        # Output the data to the console using Out-Host to force display
        # We limit to the top 20 events to prevent massive output
        if ($EventData) {
            Write-Host "Found $($EventData.Count) relevant events in the last 24 hours (Displaying Top 20):" -ForegroundColor White
            $EventData | Format-Table -Wrap -AutoSize | Out-Host
        }
        else {
            Write-Host "SUCCESS: No Errors or Warnings found in the System or Application logs in the last 24 hours." -ForegroundColor Green
        }
        
        # Return the data objects
        return $EventData
    }
    catch {
        Write-Error "Failed to retrieve event log data: $($_.Exception.Message)"
        return $null
    }
}

function Publish-Report ($DiskData, $ServiceData, $EventLogData) {
    <#
    .SYNOPSIS
    Generates a single HTML report file from the collected system health data.
    #>

    Write-Host "`n--- Generating HTML Report ---" -ForegroundColor Magenta
    
    # --- Report Styling and Header ---
    $ReportPath = "$($env:USERPROFILE)\Desktop\SystemHealthReport_$((Get-Date).ToString('yyyyMMdd_HHmmss')).html"
    $ReportTitle = "Windows System Health Report for $($env:COMPUTERNAME)"
    
    $Style = @"
<style>
    body { font-family: Arial, sans-serif; margin: 20px; background-color: #f4f4f9; }
    h1 { color: #333; border-bottom: 2px solid #ccc; padding-bottom: 10px; }
    h2 { color: #555; margin-top: 30px; }
    table { width: 100%; border-collapse: collapse; margin-top: 10px; box-shadow: 0 2px 3px rgba(0,0,0,0.1); }
    th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
    th { background-color: #007bff; color: white; }
    tr:nth-child(even) { background-color: #e9e9e9; }
    .success { color: green; font-weight: bold; }
    .error { color: red; font-weight: bold; }
</style>
"@

    # --- Convert Data Sections to HTML ---

    # 1. Disk Space Table
    $DiskHTML = $DiskData | ConvertTo-Html -Fragment -PreContent "<h2>Disk Space Utilization</h2>"
    
    # 2. Service Status Table (Highlight Stopped Services)
    $ServiceHTML = $ServiceData | 
        Select-Object Name, DisplayName, Status, StartType, 
        @{Name='Status_Highlight'; Expression={
            if ($_.Status -ne "Running" -and $_.StartType -ne "Disabled") {
                "<span class='error'>$($_.Status)</span>"
            } else {
                "<span class='success'>$($_.Status)</span>"
            }
        }} | 
        ConvertTo-Html -Fragment -PreContent "<h2>Critical Service Status</h2>"

    # 3. Event Log Table
    $EventLogHTML = if ($EventLogData) {
        $EventLogData | ConvertTo-Html -Fragment -PreContent "<h2>Recent Critical Events (Last 24 Hours)</h2>"
    } else {
        "<p class='success'>No Errors or Warnings found in the last 24 hours.</p>"
    }

    # --- Combine and Export ---

    $HTMLBody = $DiskHTML + $ServiceHTML + $EventLogHTML

    ConvertTo-Html -Title $ReportTitle -Head $Style -Body $HTMLBody | Out-File $ReportPath

    Write-Host "SUCCESS: Report saved to $ReportPath" -ForegroundColor Green
    
    # Open the report in the default browser (Optional but helpful)
    Invoke-Item $ReportPath
}

function Publish-Report ($DiskData, $ServiceData, $EventLogData) {
    <#
    .SYNOPSIS
    Generates a single HTML report file from the collected system health data.
    #>

    Write-Host "`n--- Generating HTML Report ---" -ForegroundColor Magenta
    
    # --- Report Styling and Header ---
    $ReportPath = "$($env:USERPROFILE)\Desktop\SystemHealthReport_$((Get-Date).ToString('yyyyMMdd_HHmmss')).html"
    $ReportTitle = "Windows System Health Report for $($env:COMPUTERNAME)"
    
    $Style = @"
<style>
    body { font-family: Arial, sans-serif; margin: 20px; background-color: #f4f4f9; }
    h1 { color: #333; border-bottom: 2px solid #ccc; padding-bottom: 10px; }
    h2 { color: #555; margin-top: 30px; }
    table { width: 100%; border-collapse: collapse; margin-top: 10px; box-shadow: 0 2px 3px rgba(0,0,0,0.1); }
    th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
    th { background-color: #007bff; color: white; }
    tr:nth-child(even) { background-color: #e9e9e9; }
    .success { color: green; font-weight: bold; }
    .error { color: red; font-weight: bold; }
</style>
"@

    # --- Convert Data Sections to HTML ---

    # 1. Disk Space Table
    $DiskHTML = $DiskData | ConvertTo-Html -Fragment -PreContent "<h2>Disk Space Utilization</h2>"
    
    # 2. Service Status Table (Highlight Stopped Services)
    $ServiceHTML = $ServiceData | 
        Select-Object Name, DisplayName, Status, StartType, 
        @{Name='Status_Highlight'; Expression={
            if ($_.Status -ne "Running" -and $_.StartType -ne "Disabled") {
                "<span class='error'>$($_.Status)</span>"
            } else {
                "<span class='success'>$($_.Status)</span>"
            }
        }} | 
        ConvertTo-Html -Fragment -PreContent "<h2>Critical Service Status</h2>"

    # 3. Event Log Table
    $EventLogHTML = if ($EventLogData) {
        $EventLogData | ConvertTo-Html -Fragment -PreContent "<h2>Recent Critical Events (Last 24 Hours)</h2>"
    } else {
        "<p class='success'>No Errors or Warnings found in the last 24 hours.</p>"
    }

    # --- Combine and Export ---

    $HTMLBody = $DiskHTML + $ServiceHTML + $EventLogHTML

    ConvertTo-Html -Title $ReportTitle -Head $Style -Body $HTMLBody | Out-File $ReportPath

    Write-Host "SUCCESS: Report saved to $ReportPath" -ForegroundColor Green
    
    # Open the report in the default browser (Optional but helpful)
    Invoke-Item $ReportPath
}

# ----------------------------------------------------------------------
# --- FINAL SCRIPT EXECUTION BLOCK ---
# ----------------------------------------------------------------------

Write-Host "`n=================================================" -ForegroundColor White

# 1. Collect Data (These lines also print the console tables)
$DiskData = Get-DiskSpaceReport
$ServiceData = Get-ServiceStatusReport
$EventLogData = Get-EventLogReview

# 2. Generate and Publish the Report
# We pass the collected variables ($DiskData, etc.) to the publishing function.
Publish-Report -DiskData $DiskData -ServiceData $ServiceData -EventLogData $EventLogData

Write-Host "`n=================================================" -ForegroundColor White
Write-Host "Project Complete. Ready for submission!" -ForegroundColor Cyan

