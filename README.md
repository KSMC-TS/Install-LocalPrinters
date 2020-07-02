# Install-LocalPrinters
Install printers staged in Azure Blob Storage for deployment in Intune

## DESCRIPTION
- Use this script to deploy and redeploy printers with drivers and CSV located on Azure Blob Storage.
- Best used when packaged as an Intune Win32 application in tandem with Blob storage.
- When a printer needs to be added to a deployment, add the drivers to the container, update the CSV, and update the deployDate check on the Intune app deployment.
- CSV should have the following headers:
  - printerName - display name of the printer
  - printerIP - IP address of the printer
  - driverFilePath - path to driver (inf and other dependent files) - printernamefolder\printerdriver.inf
  - driverName - name found under [strings] within the INF file - HP OfficeJet Pro 8020 would be HP OfficeJet Pro 8020 series
  - colorSetting:
    - valid values: Grayscale, Color, Monochrome
    - refer to https://docs.microsoft.com/en-us/windows/win32/printdocs/pageoutputcolor
  - duplexSetting:
    - valid values: OneSided, TwoSidedShortEdge, TwoSidedLongEdge
    - refer to https://docs.microsoft.com/en-us/windows/win32/printdocs/jobduplexalldocumentscontiguously
  - stapleSetting:
    - valid values: None, StapleTopLeft, StapleTopRight, StapleBottomLeft, StapleBottomRight, StapleDualLeft, StapleDualRight, StapleDualTop, StapleDualBottom
    - refer to https://docs.microsoft.com/en-us/windows/win32/printdocs/jobstaplealldocuments
## PARAMETER blobSAS
- This should be the URL that the contents of the printer deployment are to be downloaded from.
- For Azure: container URL + SAS token
- Blob should contain the following
  - printerDeploy.csv with printers to be deployed/redeployed
  - folders for each printer containing driver files
## PARAMETER deployDate
- This will set a registry key at HKLM:\SOFTWARE\WOW6432Node\printerDeploy with the value specified here.
- Use this as a check that the most current deployment is installed.
## NOTES
    Version:         1.5
    Last updated:    03/04/2020
    Creation Date:   02/12/2020
    Author:          Zachary Choate
    URL:             https://raw.githubusercontent.com/zchoate/Install-LocalPrinters/main/Install-LocalPrinters.ps1
