# Download the Office Deployment Tool (replace with actual URL if needed)
$odtUrl = "https://download.microsoft.com/download/E/4/D/E4D9A805-C6D5-401B-8419-4D058AC3362A/officedeploymenttool_13822-20200.exe"
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