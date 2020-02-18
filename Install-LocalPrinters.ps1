<#
.SYNOPSIS
    .
.DESCRIPTION
    Use this script to deploy and redeploy printers with drivers and CSV located on Azure Blob Storage.
    Best used when packaged as an Intune Win32 application in tandem with Blob storage.
    When a printer needs to be added to a deployment, add the drivers to the container, update the CSV,
        and update the deployDate check on the Intune app deployment.
    CSV should have the following headers:
        printerName - display name of the printer
        printerIP - IP address of the printer
        driverFilePath - path to driver (inf and other dependent files) - printernamefolder\printerdriver.inf
        driverName - name found under [strings] within the INF file - HP OfficeJet Pro 8020 would be HP OfficeJet Pro 8020 series
.PARAMETER blobSAS
    This should be the URL that the contents of the printer deployment are to be downloaded from.
    For Azure: container URL + SAS token
    Blob should contain the following
        printerDeploy.csv with printers to be deployed/redeployed
        folders for each printer containing driver files
.PARAMETER deployDate
    This will set a registry key at HKLM:\SOFTWARE\printerDeploy with the value specified here.
    Use this as a check that the most current deployment is installed.
.NOTES
    Version:         0.4
    Author:          Zachary Choate
    Creation Date:   02/12/2020
    URL:             https://raw.githubusercontent.com/zchoate/Install-LocalPrinters/master/Install-LocalPrinters.ps1
#>

param(
    [string] $blobSAS,
    [string] $deployDate
    )

function Install-LocalPrinter {
    
    Param($driverName,$driverFilePath,$printerIP,$printerName)

    #Install Printer Driver using PNP and Add-PrinterDriver
    $pnpOutput = Start-Process pnputil -ArgumentList "/add-driver", $driverFilePath, "/force" -NoNewWindow -PassThru -Wait | Select-String "Published name"
    $null = $pnpOutput -match "Published name :\s*(?<name>.*\.inf)"
    $driverINF = Get-ChildItem -Path C:\Windows\INF\$($matches.Name)
    Add-PrinterDriver -Name $driverName -InfPath $driverINF

    #Install Printer Port and Printer
    $printerPortStatus = Get-PrinterPort -Name "IP_$printerIP" -ErrorAction Ignore
    Add-PrinterPort -Name "IP_$printerIP" -PrinterHostAddress "IP_$printerIP" -ErrorAction Ignore
    Start-Sleep 10
    Add-Printer -Name $printerName -ShareName $printerName  -PortName "IP_$printerIP" -DriverName $driverName
}

function Invoke-BlobItems {  
    param (
        [Parameter(Mandatory)]
        [string]$URL,
        [string]$Path
    )

    $uri = $URL.split('?')[0]
    $sas = $URL.split('?')[1]

    $newurl = $uri + "?restype=container&comp=list&" + $sas 

    #Invoke REST API
    $body = Invoke-RestMethod -uri $newurl

    #cleanup answer and convert body to XML
    $xml = [xml]$body.Substring($body.IndexOf('<'))

    #use only the relative Path from the returned objects
    $files = $xml.ChildNodes.Blobs.Blob.Name

    #create folder structure and download files
    $files | ForEach-Object { $_; New-Item (Join-Path $Path (Split-Path $_)) -ItemType Directory -ea SilentlyContinue | Out-Null
        (New-Object System.Net.WebClient).DownloadFile($uri + "/" + $_ + "?" + $sas, (Join-Path $Path $_))
     }
}

function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
}

New-Item -ItemType Directory -Path "$env:TEMP\printerDeploy" -Force
Invoke-BlobItems -URL $blobSAS -Path "$env:TEMP\printerDeploy" | Out-Null
Start-Sleep -s 300

$logFile = "$env:TEMP\printerDeploy\printerDeploy.log"

$printers = Import-Csv -Path "$env:TEMP\printerDeploy\printerDeploy.csv" -ErrorVariable PrinterDeployError

ForEach($printer in $printers) {

    # Check to make sure printer isn't already installed - if it is, check other parameters in the event of a printer redeployment
    $printerstatus = Get-Printer -Name $printer.PrinterName -ErrorAction Ignore

    # Create path for printer driver
    $driverPath = "$env:TEMP\printerDeploy\" + $printer.DriverFilePath

    $printerIP = $printer.PrinterIP
    If(!($printerstatus)) {

        # Install printer per Install-LocalPrinter function defined.
        Install-LocalPrinter -driverName $printer.DriverName -driverFilePath $driverPath -printerIP $printer.PrinterIP -printerName $printer.PrinterName
        Start-Sleep -Seconds 30
        
        # Check to see that printer was successfully installed.
        If(!(Get-Printer -Name $printer.PrinterName -ErrorVariable PrinterDeployError -ErrorAction SilentlyContinue)) {
            
            $deployError = $printer.PrinterName + " failed to be redeployed."
            Write-Output "$(Get-TimeStamp) - $deployError" | Out-File $logFile -Append

        }

    # Look at currently installed printer and compare driver and printer port - if they don't match, let's redeploy.
    } elseif(($printerstatus.DriverName -notlike $printer.DriverName) -or ($printerstatus.PortName -notlike "*$printerIP*")) {

        # Remove printer and install printer with updated parameters. This can probably be written to update the existing printer but for time, this is quickest.
        ## Opportunity for rewrite.
        Remove-Printer -Name $printerstatus.Name
        Start-Sleep -Seconds 10
        Install-LocalPrinter -driverName $printer.DriverName -driverFilePath $driverPath -printerIP $printer.PrinterIP -printerName $printer.PrinterName
        Start-Sleep -Seconds 30
        If(!(Get-Printer -Name $printer.PrinterName -ErrorVariable PrinterDeployError -ErrorAction SilentlyContinue)) {

            $deployError = $printer.PrinterName + " failed to be redeployed."
            Write-Output "$(Get-TimeStamp) - $deployError" | Out-File $logFile -Append
        
        }

    } else { 
        
        $deployOutput = $printer.PrinterName + " is already installed. Moving on..."
        Write-Output "$(Get-TimeStamp) - $deployOutput" | Out-File $logFile -Append

    }

}

If($PrinterDeployError) {

    Exit 1618

}

Remove-Item -Recurse -Path "$env:TEMP\printerDeploy" -Force
Set-Location HKLM:
New-Item -Path .\SOFTWARE -Name "printerDeploy" -Value $deployDate -Force
Exit 0
