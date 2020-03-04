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
        colorSetting:
            - valid values: Grayscale, Color, Monochrome
            - refer to https://docs.microsoft.com/en-us/windows/win32/printdocs/pageoutputcolor
        duplexSetting:
            - valid values: OneSided, TwoSidedShortEdge, TwoSidedLongEdge
            - refer to https://docs.microsoft.com/en-us/windows/win32/printdocs/jobduplexalldocumentscontiguously
        stapleSetting:
            - valid values: None, StapleTopLeft, StapleTopRight, StapleBottomLeft, StapleBottomRight, StapleDualLeft,
StapleDualRight, StapleDualTop, StapleDualBottom
            - refer to https://docs.microsoft.com/en-us/windows/win32/printdocs/jobstaplealldocuments
.PARAMETER blobSAS
    This should be the URL that the contents of the printer deployment are to be downloaded from.
    For Azure: container URL + SAS token
    Blob should contain the following
        printerDeploy.csv with printers to be deployed/redeployed
        folders for each printer containing driver files
.PARAMETER deployDate
    This will set a registry key at HKLM:\SOFTWARE\printerDeploy with the value specified here.
    Use this as a check that the most current deployment is installed.
.PARAMETER defaultPrinter
    Specify the name of the printer that should be set as the default.
    This should be identical to the name of a printer deployed via this script or that is already installed.
.NOTES
    Version:         1.5
    Last updated:    03/03/2020
    Creation Date:   02/12/2020
    Author:          Zachary Choate
    URL:             https://raw.githubusercontent.com/zchoate/Install-LocalPrinters/master/Install-LocalPrinters.ps1
#>

param(
    [string] $blobSAS,
    [string] $deployDate,
    [string] $defaultPrinter
    )

function Install-LocalPrinter {
    
    Param($driverName,$driverFilePath,$printerIP,$printerName)

    #Install Printer Driver using PNP and Add-PrinterDriver
    Start-Process "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList "& c:\windows\sysnative\pnputil.exe /add-driver $driverFilePath /install /force" -RedirectStandardOutput "$env:TEMP\printerDeploy\pnpOutput.txt" -NoNewWindow -Wait 
    [string]$pnpOutput = (Get-Content "$env:TEMP\printerDeploy\pnpOutput.txt") -match "Published name:\s*(?<name>.*\.inf)"
    $driverInf = ($pnpOutput.Split(":") -match ".inf").Trim()
    $driverInf = "C:\Windows\INF\$driverInf"
    Add-PrinterDriver -Name $driverName -InfPath $driverINF

    #Install Printer Port and Printer
    $printerPortStatus = Get-PrinterPort -Name "IP_$printerIP" -ErrorAction Ignore
    Add-PrinterPort -Name "IP_$printerIP" -PrinterHostAddress "$printerIP" -ErrorAction Ignore
    Start-Sleep 10
    Add-Printer -Name $printerName -PortName "IP_$printerIP" -DriverName $driverName
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

function Set-DefaultPrinter {
    param(
        [string] $defaultPrinter
    )

    $Printers = Get-WmiObject -Class Win32_Printer
    $Printer = $Printers | Where-Object {$_.Name -eq "$defaultPrinter"}
    $Printer.SetDefaultPrinter() | Out-Null
    
}

function Set-PrintConfigurationSetting {
    param(
        [string] $printerName,
        [string] $configSetting,
        [string] $configValue
    )

        # Get current printer configuration defaults
        $printerConfig = Get-PrintConfiguration -PrinterName $printerName
        $printerConfigXML = [xml]$printerConfig.PrintTicketXML
        # Pull the specific setting that we're changing
        $settingConfig = ($printerConfigXML.PrintTicket.Feature).Where({$_.name -eq "$configSetting"}).option.name
        # Stage the updated setting in the Print Ticket XML.
        $updatedConfig = $printerConfig.PrintTicketXML -replace "$settingConfig","psk:$configValue"
        # Apply the updated print setting.
        Set-PrintConfiguration -PrinterName $printerName -PrintTicketXml $updatedConfig
}

New-Item -ItemType Directory -Path "$env:TEMP\printerDeploy" -Force
Invoke-BlobItems -URL $blobSAS -Path "$env:TEMP\printerDeploy" | Out-Null
Start-Sleep -s 300

$logFile = "$env:TEMP\printerDeploy\printerDeploy.log"

$printers = Import-Csv -Path "$env:TEMP\printerDeploy\printerDeploy.csv" -ErrorVariable PrinterDeployError

ForEach($printer in $printers) {

    $printerIP = $printer.PrinterIP

    # Check to make sure printer isn't already installed - if it is, check other parameters in the event of a printer redeployment
    $printerbyName = Get-Printer -Name $printer.PrinterName -ErrorAction Ignore
    $printerbyPort = Get-Printer | Where-Object {$_.PortName -like "*$printerIP"} -ErrorAction Ignore
    $printerbyIP = Get-PrinterPort | Where-Object {$_.PrinterHostAddress -eq $printer.PrinterIP} -ErrorAction Ignore

    # Create path for printer driver
    $driverPath = "$env:TEMP\printerDeploy\" + $printer.DriverFilePath

    If(!($printerbyName)) {

        # Install printer per Install-LocalPrinter function defined.
        Install-LocalPrinter -driverName $printer.DriverName -driverFilePath $driverPath -printerIP $printer.PrinterIP -printerName $printer.PrinterName
        Start-Sleep -Seconds 30
        
        # Check to see that printer was successfully installed.
        If(!(Get-Printer -Name $printer.PrinterName -ErrorVariable PrinterDeployError -ErrorAction SilentlyContinue)) {
            
            $deployError = $printer.PrinterName + " failed to be redeployed."
            Write-Output "$(Get-TimeStamp) - $deployError" | Out-File $logFile -Append

        }

    # Look at currently installed printer and compare driver, printer port, and IP - if they don't match, let's redeploy.
    } elseif($printerbyName.DriverName -notlike $printer.DriverName) {

        # Remove printer and install printer with updated parameters.
        $currentPrinter = Get-Printer -Name $printer.PrinterName
        Remove-Printer -Name $printerbyName.Name
        Start-Sleep -Seconds 10
        Try { Remove-PrinterPort -Name $currentPrinter.PortName } catch {
            Get-Printer | Where-Object {$_.PortName -like $currentPrinter.PortName} | Remove-Printer
            Start-Sleep -Seconds 10
            Restart-Service -Name Spooler
            Remove-PrinterPort -Name $currentPrinter.PortName
        }
        Start-Sleep -Seconds 10
        Install-LocalPrinter -driverName $printer.DriverName -driverFilePath $driverPath -printerIP $printer.PrinterIP -printerName $printer.PrinterName
        Start-Sleep -Seconds 30
        If(!(Get-Printer -Name $printer.PrinterName -ErrorVariable PrinterDeployError -ErrorAction SilentlyContinue)) {

            $deployError = $printer.PrinterName + " failed to be redeployed."
            Write-Output "$(Get-TimeStamp) - $deployError" | Out-File $logFile -Append
        
        }

    } elseif(!($printerbyPort)) {

        # Remove printer associated with port and associated port. Install with updated parameters.
        $currentPrinter = Get-Printer -Name $printer.PrinterName
        Remove-Printer -Name $printer.PrinterName
        Start-Sleep -Seconds 10
        Try { Remove-PrinterPort -Name $currentPrinter.PortName } catch {
            Get-Printer | Where-Object {$_.PortName -like $currentPrinter.PortName} | Remove-Printer
            Start-Sleep -Seconds 10
            Restart-Service -Name Spooler
            Remove-PrinterPort -Name $currentPrinter.PortName
        }
        Start-Sleep -Seconds 10
        Install-LocalPrinter -driverName $printer.DriverName -driverFilePath $driverPath -printerIP $printer.PrinterIP -printerName $printer.PrinterName
        Start-Sleep -Seconds 30
        If(!(Get-Printer -Name $printer.PrinterName -ErrorVariable PrinterDeployError -ErrorAction SilentlyContinue)) {

            $deployError = $printer.PrinterName + " failed to be redeployed."
            Write-Output "$(Get-TimeStamp) - $deployError" | Out-File $logFile -Append
        
        }

    } elseif(!($printerbyIP)) {

        # Remove printer associated with IP and associated port. Install with updated parameters.
        $currentPrinter = Get-Printer -Name $printer.PrinterName
        Remove-Printer -Name $printer.PrinterName
        Start-Sleep -Seconds 10
        Try { Remove-PrinterPort -Name $currentPrinter.PortName } catch {
            Get-Printer | Where-Object {$_.PortName -like $currentPrinter.PortName} | Remove-Printer
            Start-Sleep -Seconds 10
            Restart-Service -Name Spooler
            Remove-PrinterPort -Name $currentPrinter.PortName
        }
        Start-Sleep -Seconds 10
        Install-LocalPrinter -driverName $printer.DriverName -driverFilePath $driverPath -printerIP $printer.PrinterIP -printerName $printer.PrinterName
        Start-Sleep -Seconds 30
        If(!(Get-Printer -Name $printer.PrinterName -ErrorVariable PrinterDeployError -ErrorAction SilentlyContinue)) {

            $deployError = $printer.PrinterName + " failed to be redeployed."
            Write-Output "$(Get-TimeStamp) - $deployError" | Out-File $logFile -Append
        
        }

    } else { 
        
        # If we're here the printer was likely already deployed - updating log file and skipping.
        $deployOutput = $printer.PrinterName + " is already installed."
        Write-Output "$(Get-TimeStamp) - $deployOutput" | Out-File $logFile -Append

    }

    # Set the default color setting for the printer if specified.
    If($printer.ColorSetting) {

        Set-PrintConfigurationSetting -printerName $printer.PrinterName -configSetting "psk:PageOutputColor" -configValue $printer.ColorSetting
        $deployOutput = $printer.PrinterName + " has had the color default set."
        Write-Output "$(Get-TimeStamp) - $deployOutput" | Out-File $logFile -Append

    }

    # Set the default duplex setting for the printer if specified.
    If($printer.DuplexSetting) {

        Set-PrintConfigurationSetting -printerName $printer.PrinterName -configSetting "psk:JobDuplexAllDocumentsContiguously" -configValue $printer.DuplexSetting
        $deployOutput = $printer.PrinterName + " has had the duplex default set."
        Write-Output "$(Get-TimeStamp) - $deployOutput" | Out-File $logFile -Append

    }

    # Set the default staple setting for the printer if specified.
    If($printer.StapleSetting) {

        Set-PrintConfigurationSetting -printerName $printer.PrinterName -configSetting "psk:JobStapleAllDocuments" -configValue $printer.StapleSetting
        $deployOutput = $printer.PrinterName + " has had the staple default set."
        Write-Output "$(Get-TimeStamp) - $deployOutput" | Out-File $logFile -Append

    }

}

# Set default printer if specified
If(!($defaultPrinter)) {

    Write-Output "$(Get-TimeStamp) - No default printer set, skipping..." | Out-File $logFile -Append

} else {

    Set-DefaultPrinter -defaultPrinter $defaultPrinter
    Write-Output "$(Get-TimeStamp) - $defaultPrinter was set as the default printer." | Out-File $logFile -Append

}

If($PrinterDeployError) {

    Exit 1618

}

Remove-Item -Recurse -Path "$env:TEMP\printerDeploy" -Force
Set-Location -Path HKLM:
New-Item -Path .\SOFTWARE -Name "printerDeploy" -Value $deployDate -Force
Exit 0
