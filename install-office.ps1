# Download the Office Deployment Tool (replace with actual URL if needed)
$odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20158.exe"
$odtPath = "$env:TEMP\officedeploymenttool.exe" 
Invoke-WebRequest -Uri $odtUrl -OutFile $odtPath

# Extract the ODT
Start-Process -FilePath $odtPath -ArgumentList "/quiet /extract:$env:TEMP" -Wait

# Create the configuration XML file
$configXml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Monthly">
    <Product ID="ProPlusRetail">
      <Language ID="en-us" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@
$configXmlPath = "$env:TEMP\configuration.xml"
$configXml | Out-File -FilePath $configXmlPath -Encoding UTF8

# Install Office ProPlus
Start-Process -FilePath "$env:TEMP\setup.exe" -ArgumentList "/configure $env:TEMP\configuration.xml" -Wait
