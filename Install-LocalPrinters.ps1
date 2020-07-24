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
    This will set a registry key at HKLM:\SOFTWARE\PrinterDeploy with the value specified here.
    Use this as a check that the most current deployment is installed.
.NOTES
    Version:         2.0
    Last updated:    07/20/2020
    Creation Date:   02/12/2020
    Author:          Zachary Choate
    URL:             https://raw.githubusercontent.com/zchoate/Install-LocalPrinters/main/Install-LocalPrinters.ps1
#>

param(
    [string] $blobSAS,
    [string] $deployDate
    )

function Install-LocalPrinter {
    
    Param($driverName,$driverFilePath,$printerIP,$printerName)

    #Install Printer Driver using PNP and Add-PrinterDriver
    Start-Process "powershell.exe" -ArgumentList "& c:\windows\system32\pnputil.exe /add-driver $driverFilePath /install /force" -RedirectStandardOutput "$env:TEMP\printerDeploy\pnpOutput.txt" -NoNewWindow -Wait 
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

Function Remove-LocalPrinters {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $printerIp,
        [Parameter()]
        [string]
        $printerName
    )

    If($printerIp) {
        $printerToRemove = Get-Printer | Where-Object {$_.PortName -like "*$printerIp*"}
    } elseif ($printerName) {
        $printerToRemove = Get-Printer | Where-Object {$_.Name -like "*$printerName*" -or $_.ShareName -like "*$printerName*"}
    } else {
        Write-Output "No printer IP or name was specified. Try passing either one as a parameter."
    }

    If($printerToRemove) {
        $printerPortName = $printerToRemove.PortName
        Remove-Printer -Name $printerToRemove.Name
        Remove-PrinterPort -Name $printerPortName
    } else {
        Write-Output "Printer $printerIp $printerName doesn't appear to be installed."
    }
}

Function Set-SharpDriverModules {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $moduleType,
        [Parameter(Mandatory=$true)]
        [string]
        $moduleConfiguration,
        [Parameter(Mandatory=$true)]
        [string]
        $printerName
    )

    switch ( $moduleType ) {
        "SharpUd3Punch" {
            $basePath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\$printerName\PrinterDriverData"
            $key = "dev_punchmodule_def"
            switch ( $moduleConfiguration) {
                "None" { $value = 0 }
                "2 Holes" { $value = 1 }
                "3 Holes" { $value = 6 }
                "4 Holes" { $value = 7 }
                "4 Holes Wide" { $value = 4 }
                "2/3 Holes" { $value = 2 }
                "2/4 Holes" { $value = 3 }
            }
        }
        "SharpUd3Staple" {
            $basePath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\$printerName\PrinterDriverData"
            $key = "dev_staplemodule_def"
            switch ( $moduleConfiguration ) {
                "None" { $value = 0 }
                "1 Staple" { $value = 1 }
                "1 Staple/2 Staples" { $value = 2 }
                "1/2/Saddle Staple" { $value = 4 }
                "1 Staple/2 Staples/Stapless" { $value = 10 }
                "1/2/Saddle/Staplesss" { $value = 12 }
            }
        }
    }

    New-ItemProperty -Path $basePath -Name $key -Value $value -PropertyType DWord -Force

}

New-Item -ItemType Directory -Path "$env:TEMP\printerDeploy" -Force
Invoke-BlobItems -URL $blobSAS -Path "$env:TEMP\printerDeploy" | Out-Null
Start-Sleep -s 300

$logFile = "$env:TEMP\printerDeploy\printerDeploy.log"

$uninstallPrinters = Import-Csv -Path "$env:TEMP\printerDeploy\printerDeploy.csv" | Where-Object {$_.install -eq 0} -ErrorVariable PrinterDeployError

ForEach ($uninstall in $uninstallPrinters) {

    Remove-LocalPrinters -printerIp $uninstall.printerIp -printerName $uninstall.printerName -ErrorVariable PrinterDeployError

}

$printers = Import-Csv -Path "$env:TEMP\printerDeploy\printerDeploy.csv" | Where-Object {$_.install -eq 1} -ErrorVariable PrinterDeployError

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

        try {
            Set-PrintConfigurationSetting -printerName $printer.PrinterName -configSetting "psk:PageOutputColor" -configValue $printer.ColorSetting -ErrorVariable PrinterDeployError
            $deployOutput = $printer.PrinterName + " has had the color default set."
        } catch {
            $deployOutput = $printer.PrinterName + " failed to have the color default set."
            $deployError = $deployOutput
        }
        Write-Output "$(Get-TimeStamp) - $deployOutput" | Out-File $logFile -Append

    }

    # Set the default duplex setting for the printer if specified.
    If($printer.DuplexSetting) {

        try {
            Set-PrintConfigurationSetting -printerName $printer.PrinterName -configSetting "psk:JobDuplexAllDocumentsContiguously" -configValue $printer.DuplexSetting -ErrorVariable PrinterDeployError
            $deployOutput = $printer.PrinterName + " has had the duplex default set."
        } catch {
            $deployOutput = $printer.PrinterName + " failed to have the duplex default set."
            $deployError = $deployOutput
        }
        Write-Output "$(Get-TimeStamp) - $deployOutput" | Out-File $logFile -Append

    }

    # Set the default staple setting for the printer if specified.
    If($printer.StapleSetting) {

        try {
            Set-PrintConfigurationSetting -printerName $printer.PrinterName -configSetting "psk:JobStapleAllDocuments" -configValue $printer.StapleSetting -ErrorVariable PrinterDeployError
            $deployOutput = $printer.PrinterName + " has had the staple default set."
        } catch {
            $deployOutput = $printer.PrinterName + " failed to have the staple default set."
            $deployError = $deployOutput
        }
        Write-Output "$(Get-TimeStamp) - $deployOutput" | Out-File $logFile -Append

    }

    # Enable Sharp staple module
    If($printer.custSharpUd3Staple) {

        try {
            Set-SharpDriverModules -moduleType "SharpUd3Staple" -moduleConfiguration $printer.custSharpUd3Staple -printerName $printer.printerName -ErrorVariable PrinterDeployError
            $deployOutput = $printer.PrinterName + " has had the staple module enabled."
        } catch {
            $deployOutput = $printer.PrinterName + " failed to have the staple module enabled."
            $deployError = $deployOutput
        }
        Write-Output "$(Get-TimeStamp) - $deployOutput" | Out-File $logFile -Append

    }

    # Enable Sharp hole punch module
    If($printer.custSharpUd3Punch) {
        try {
            Set-SharpDriverModules -moduleType "SharpUd3Punch" -moduleConfiguration $printer.custSharpUd3Punch -printerName $printer.printerName -ErrorVariable PrinterDeployError
            $deployOutput = $printer.PrinterName + " has had the hole punch module enabled."
        } catch {
            $deployOutput = $printer.PrinterName + " failed to have the hole punch module enabled."
            $deployError = $deployOutput
        }
        Write-Output "$(Get-TimeStamp) - $deployOutput" | Out-File $logFile -Append

    }
}

If($PrinterDeployError) {

    Exit 1618

}

Remove-Item -Recurse -Path "$env:TEMP\printerDeploy" -Force
New-Item -Path "HKLM:\SOFTWARE" -Name "PrinterDeploy" -Force
New-ItemProperty -Path "HKLM:\SOFTWARE\PrinterDeploy" -Name "version_applied" -Value $deployDate -Force
