using namespace System
using namespace System.Management.Automation
using namespace System.Collections.Generic

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [string[]]$PrinterNamePatterns,
    [string]$SelectedFaxFolder
)

$ErrorActionPreference = "Stop"

function Ensure-Parkview-Domain {
    param(
        [PSCredential]$Cred
    )

    if ($Cred.UserName -notlike "parkview\*") {
        $Cred = [PSCredential]::new("parkview\$($Cred.UserName)", $Cred.Password)
    }

    $Cred
}

function Get-Available-Drive-Identifer {
    $Drives = Get-PSDrive
    $Drives = $Drives | Where-Object { $_.Name -match '^[a-z]$:' }
    $IdsInUse = [HashSet[string]]::new(26)
    foreach ($Drive in $Drives) {
        $__ = $IdsInUse.Add($Drive.Name)
    }
    "Z".."D" | Where-Object { ("{0}:" -f $_ ) -notin $IdsInUse} | Select-Object -First 1
}

function Get-Parkview-Cred {
    $CredMsg = "parkview credentials"
    $Cred = Get-Credential -Message "$CredMsg"
    if ($null -eq $Cred) {
        throw "canceling with null credential"
    } 
    $Cred = Ensure-Parkview-Domain $Cred
    $Cred
}


function Connect-Fax-Network-Drive {
    $FaxServer = "\\xmfaxFS.parkviewmc.com\inbox"

    try { Remove-PSDrive -Name "tmp0" } 
    catch [DriveNotFoundException] { }

    $ParkviewCred = Get-Parkview-Cred
    New-PSDrive -Name "tmp0" -PSProvider "FileSystem" -Root "$FaxServer" -Credential $ParkviewCred

    try {
        <# \inbox has folders for each department--have user indicate their department folder. #>
        & "explorer" "$FaxServer"

        while ([string]::IsNullOrWhiteSpace($SelectedFaxFolder)) {
            $SelectedFaxFolder = Read-Host "enter full path to folder on fax server"
            if (Test-Path "$SelectedFaxFolder") {
                break
            }
        }
        
        [PSDriveInfo]$PersistentDrive = New-PSDrive -Name "$(Get-Available-Drive-Identifer)" -Root "$SelectedFolder" -Credential $ParkviewCred -Scope Global -Persist
        $PersistentDrive
    } finally {
        Remove-PSDrive -Name "tmp0"
    }
}

function Connect-Helpdesk-Drive {
    $DriveId = "Helpdesk"
    $KratosHelpdeskNetworkLocation = "\\kratos.parkviewmc.com\Helpdesk"
    
    try {
        $Drive = Get-PSDrive -Name "$DriveId"
    } catch [DriveNotFoundException] {
        $ParkviewCred = Get-Parkview-Cred
        Write-Host "Connecting via $($ParkviewCred.UserName)"    
        $Drive = New-PSDrive -Name "$DriveId" -PSProvider "FileSystem" -Root "$KratosHelpdeskNetworkLocation" -Credential $ParkviewCred -Scope Global
    }
    $Drive
}

function Copy-Parkview-Legacy-Apps {

    $ParkviewLegacyAppsSrc = "Helpdesk:\Reimaging Tools\Parkview Legacy Apps.url"
    $ParkviewLegacyAppsDest = "C:\Users\Public\Desktop"

    Copy-Item "$ParkviewLegacyAppsSrc" "$ParkviewLegacyAppsDest"
}

function Detect-Computer-Brand {
    $ComputerInfo = Get-ComputerInfo
    $Manufacturer = $ComputerInfo.CsManufacturer
    
    switch -Wildcard ($Manufacturer) {
        "*LENOVO*" {
            "Lenovo"
        }
        "*DELL*" {
            "Dell"
        }
        "*HP*" {
            "HP"
        }
        default {
            "Unknown"
        }
    }
}

function Driver-Setup-FileName {
    param(
        [string]$Manufacturer
    )

    switch ($Manufacturer) {
        "Lenovo" {
            "system_update_5.08.02.25.exe"
        }
        "Dell" {
            "SupportAssistInstaller.exe"
        }
        "HP" {
            "sp151723.exe"
        }
        default {
            throw "invalid manufacturer"
        }
    }
}

function Exec-Code {
    param(
        [string]$Command
    )
    Write-Host "Executing: & $Command"
    $Proc = Start-Process -FilePath "$Command" -PassThru -Wait
    $__ = $Proc.Handle
    $Proc.ExitCode
}

function Launch-Chrome-Setup-Await {
    $FullPathToChromeSetup = "Helpdesk:\Reimaging Tools\Ninite Chrome Installer.exe"
    Exec-Code "$FullPathToChromeSetup"
}


function Launch-Drivers-Setup-Await {
    param(
        [string]$DriverSetupFileName
    )

    $FullPathToDriverSetup = "Helpdesk:\Reimaging Tools\$DriverSetupFileName"
    Exec-Code "$FullPathToDriverSetup"
}

function Rundll32-Add-Printer {
    param(
        [string]$Name
    )
    Exec-Code "rundll32.dll printui.dll,PrintUIEntry /ga /n `"$Name`""
}

function Retrieve-Printers-Like {
    param(
        [string]$PrintServer,
        [string[]]$Likeness
    )
    
    $Printers = Get-Printer -Name "$PrintServer"
    $MatchedPrinters = @()

    foreach ($Like in $Likeness) {
        $Ms = $Printers | Where-Object { $_.Name -like "$Like" }
        foreach ($Match in $Ms) {
            $MatchedPrinters.Add($Match)
        }
    }
    $MatchedPrinters
}

<# https://wutils.com/wmi/root/standardcimv2/msft_printer/ #>
function Get-Printer-Network-Address {
    param(
        [CimInstance]$Printer
    )
    $PrinterAddress = "\\$($Printer.ComputerName)\$($Printer.Name)"
    $PrinterAddress
}

function Add-Printers {
    param(
        [CimInstance[]]$Printers
    )
    foreach ($Printer in $Printers) {
        $PrinterAddress = Get-Printer-Network-Address $Printer
        $AddPrinterExitCode = Rundll32-Add-Printer $PrinterAddress
        if ($AddPrinterExitCode -ne 0) {
            Write-Warning "Adding printer $PrinterAddress errored with exit code: $AddPrinterExitCode"
        }
    }
}

function Comm-Host-Connect-Helpdesk {
    Write-Host "Connecting to remote drive Helpdesk"
    $RemoteDrive = Connect-Helpdesk-Drive
    $RemoteDrive
}

function Comm-Host-Copy-Parkview-Legacy-Apps {
    Write-Host "Copying legacy apps"
    Copy-Parkview-Legacy-Apps 
}

function Comm-Host-Launch-Chrome-Setup {
    Write-Host "Launching chrome setup"
    $ChromeSetupExitCode = Launch-Chrome-Setup-Await

    if ($ChromeSetupExitCode -ne 0) {
        Write-Warning "chrome setup returned exit code: $ChromeSetupExitCode"
    }
}

function Comm-Host-Setup-Drivers {

    Write-Host "Detecting computer info"
    $Manufacturer = Detect-Computer-Brand

    if ($Manufacturer -ne "Unknown") {
        $DriverSetupFileName = Driver-Setup-FileName "$Manufacturer"

        Write-Host "Launching drivers setup"
        $DriversSetupExitCode = Launch-Drivers-Setup-Await "$DriverSetupFileName"
        
        if ($DriversSetupExitCode -eq 0) {        
            $DriversProgramExitCode = switch ($Manufacturer) {
                "Lenovo" {
                    Write-Host "Launching lenovo drivers program"
                    $PathToLenovoSysUpdate = "C:\Program Files (x86)\Lenovo\System Update\tvsu.exe"
                    Exec-Code "$PathToLenovoSysUpdate"
                }
                default {
                    Write-Warning "Drivers program not executed after setup."
                    0
                }
            }
            
            if ($DriversProgramExitCode -ne 0) {
                Write-Warning "Drivers program returned error exit code: $DriversProgramExitCode"
            }
        } else {
            Write-Warning "Driver setup program returned error exit code: $DriversSetupExitCode"    
        }
    } else {
        Write-Warning "Unknown computer manufacturer detected--won't launch driver setup."
    }
}

function Comm-Host-Rename-Computer {
    while ([string]::IsNullOrWhiteSpace($NewComputerName)) {
        $NewComputerName = Read-Host "Enter new computer name"
    }
    Rename-Computer -NewName "$NewComputerName"
}

function Comm-Host-Add-Printers {
    $PrinterNamePatterns = 
        $PrinterNamePatterns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if ($PrinterNamePatterns.Count -gt 0) {
        $PrintServer = "\\a-pmcprint1p"
        $Printers = Retrieve-Printers-Like "$PrintServer" $PrinterNamePatterns
        Write-Host "Adding printers $Printers"
        Add-Printers $Printers
    }
}

function Finishing-Touch {
    <# Connect to remote drive and mount it as Helpdesk:\ #>
    $__ = Comm-Host-Connect-Helpdesk

    <# Copy desktop shortcut to old apps #>
    Comm-Host-Copy-Parkview-Legacy-Apps

    <# Launch chrome setup and await user to complete setup #>
    Comm-Host-Launch-Chrome-Setup
    
    <# Launch drivers setup and await user to complete setup. 
       If drivers installed program is known, execute it. #>
    Comm-Host-Setup-Drivers

    Comm-Host-Rename-Computer
    Comm-Host-Add-Printers

    Write-Host "reboot to complete"
    Read-Host "Press enter to exit program"
}

Write-Host "call Finishing-Touch to execute"