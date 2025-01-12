# Define log file path
$logFilePath = "$env:TEMP\install_excel_log.txt"

# Function to log messages
function Log-Message {
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $message"
    Add-Content -Path $logFilePath -Value $logMessage
}

# Function to check if Excel is installed
function Is-ExcelInstalled {
    $excelPath = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
    return Test-Path $excelPath
}

# Start logging
Log-Message "Starting Excel installation script."

try {
    if (-not (Is-ExcelInstalled)) {
        # Download the Office Deployment Tool
        $odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20158.exe"
        $odtPath = "$env:TEMP\officedeploymenttool.exe"
        Invoke-WebRequest -Uri $odtUrl -OutFile $odtPath
        Log-Message "Downloaded Office Deployment Tool."

        # Extract the ODT
        Start-Process -FilePath $odtPath -ArgumentList "/quiet /extract:$env:TEMP" -Wait
        Log-Message "Extracted Office Deployment Tool."

        # Create the configuration XML file for Excel (Volume Licensing)
        $configXml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="PerpetualVL2021">
    <Product ID="Excel2021Volume">
      <Language ID="en-us" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Property Name="SharedComputerLicensing" Value="1" />
  <Property Name="PinIconsToTaskbar" Value="FALSE" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Property Name="DisableFirstRunWizard" Value="TRUE" />
</Configuration>
"@
        $configXmlPath = "$env:TEMP\configuration.xml"
        $configXml | Out-File -FilePath $configXmlPath -Encoding UTF8
        Log-Message "Created configuration XML file."

        # Install Excel silently
        Start-Process -FilePath "$env:TEMP\setup.exe" -ArgumentList "/configure $configXmlPath" -Wait
        Log-Message "Started Excel installation."
    } else {
        Log-Message "Excel is already installed."
    }
} catch {
    Log-Message "Error: $_"
    throw
}

Log-Message "Excel installation script completed."