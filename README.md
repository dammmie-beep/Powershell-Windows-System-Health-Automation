# Powershell-Windows-System-Health-Automation

## Windows System Health Reporter (PowerShell)
A robust PowerShell script designed to automate daily system administration checks on Windows environments and compile the results into a single, clean HTML report. This tool streamlines system health auditing, drastically reducing the time spent manually reviewing logs and metrics.

## Key Features
This script provides an immediate, aggregated status update across three critical system areas:

**Disk Space Monitoring:** Checks all local fixed drives and calculates Total GB, Free GB, and Used Percentage for quick capacity review.

**Critical Service Status:** Audits the status of essential services (EventLog, wuauserv, Spooler, etc.) and highlights any services that are Stopped but set to start automatically.

**Event Log Review:** Uses the modern Get-WinEvent cmdlet to query the System and Application logs for all Errors and Warnings generated in the last 24 hours.

**Professional HTML Output:** Generates a single, timestamped HTML file for easy sharing, archiving, and review.

### Getting Started

#### Prerequisites

A Windows machine running PowerShell 5.1 or newer (PowerShell 7+ is recommended).

Script execution permission (Set using: ```Set-ExecutionPolicy RemoteSigned -Scope CurrentUser```).

#### Installation and Execution

1. Clone the Repository:
   ```
   git clone https://github.com/dammmie-beep/Powershell-Windows-System-Health-Automation.git
    cd system-health-reporter
   ```
2.  Execute the script directly from a PowerShell console:
   ```
      .\SystemHealthReport.ps1 ```

3. The script will automatically open the generated report file, which is saved to your Desktop with a timestamped filename (e.g., ```SystemHealthReport_20251106_103000.html```).

#### Technical Implementation Highlights
1. Uses PowerShell functions with proper verb-noun naming conventions and Comment-Based Help (.SYNOPSIS, .DESCRIPTION).

2. Leverages Get-CimInstance over Get-WmiObject and Get-WinEvent (over Get-EventLog) for improved performance and reliability.

3. Implements try/catch blocks for error handling to ensure the script does not halt on minor failures (e.g., if a service name is misspelt).

4. Utilizes Calculated Properties (Select-Object @{...}) to format raw data (bytes) into human-readable metrics (GB and Percentages) before reporting.
