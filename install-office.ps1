# Download the Office Deployment Tool
$odtUrl = "https://download.microsoft.com/download/E/4/D/E4D9A805-C6D5-401B-8419-4D058AC3362A/officedeploymenttool_13822-20200.exe"
$odtPath = "C:\ODT\officedeploymenttool.exe"
Invoke-WebRequest -Uri $odtUrl -OutFile $odtPath

# Extract the ODT
Start-Process -FilePath $odtPath -ArgumentList "/quiet /extract:C:\ODT" -Wait

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
$configXmlPath = "C:\ODT\configuration.xml"
$configXml | Out-File -FilePath $configXmlPath -Encoding UTF8

# Install Office ProPlus
Start-Process -FilePath "C:\ODT\setup.exe" -ArgumentList "/configure C:\ODT\configuration.xml" -Wait
