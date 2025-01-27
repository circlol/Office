
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
[Console]::Title = 'OfficeScrubber'
[Console]::ForegroundColor = 'White'
[Console]::BackgroundColor = 'Black'
function Get-Question {
    $answer = Read-Host -Prompt "Would you like to run SaRA? (Y/N)"
    switch ($answer.ToUpper()) {
        'Y' { return $True }
        'N' { return $False}
        Default { Get-Question }
    }
}
function Remove-Office {
    # Checks
    If (!$InstalledOffice) {
        $x32 = ${env:ProgramFiles(x86)} + "\Microsoft Office"
        $x64 = $env:ProgramFiles + "\Microsoft Office"
        if (Test-Path -Path $x32) {$Excel32 = Get-ChildItem -Recurse -Path $x32 -Filter "EXCEL.EXE"}
        if (Test-Path -Path $x64) {$Excel64 = Get-ChildItem -Recurse -Path $x64 -Filter "EXCEL.EXE"}
        if ($Excel32) {$Excel = $Excel32}
        if ($Excel64) {$Excel = $Excel64}
        $RegPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
        $DisplayVersion = Get-ItemProperty -Path $RegPath -Name "DisplayVersion" | Where-Object {$_.DisplayVersion -eq $Excel.VersionInfo.ProductVersion -and $_.PSChildName -notlike "{*}"}
        $Office = Get-ItemProperty -Path $DisplayVersion.PSPath
        $InstalledOffice = $Office | ForEach-Object {$_.DisplayName + $(if ($_.InstallLocation -eq $x32) {" 32 Bit "} else {" 64 Bit "})  + $_.DisplayVersion}
    }
    
    If ($null -eq $InstalledOffice) { $InstalledOffice = "No Office Products Detected." }

    # Warns
    Write-Output "
    
   Office Removal Tool

Office Products Installed:
$InstalledOffice

What will happen?`nThe latest version of the Microsoft Support and Recovery Assistant (SaRA) tool will be downloaded and run using the commands below:

SaRACMD.exe -OfficeScrubScenario -AcceptEula -OfficeVersion All 

URL: https://aka.ms/SaRA_EnterpriseVersionFiles
THIS TOOL WILL REMOVE EVERY VERSION OF OFFICE.`n     THIS IS THE ONLY CONFIRMATION`n`n"

    $answer = Get-Question



    If ($answer -eq $True) {
        # Downloads
        try { Write-Output "Downloading SaRA."
        Invoke-WebRequest https://aka.ms/SaRA_EnterpriseVersionFiles -OutFile $env:temp 
    }
    catch { return "Error downloading SaRA:`n$($Error[0])" }
    
    # Extracts
    try { Write-Output "Extracting files."
        Expand-Archive -Path "$env:temp\SaRACmd_*.zip" -DestinationPath "$env:temp\SaRA" -Force
    }
    catch { return "Error expanding SaRA:`n$($Error[0])" }

    # Starts
    try { Write-Output "Starting SaRA."
        Start-Process (Get-ChildItem "$env:temp\SaRA" -Include SaRAcmd.exe -Recurse).FullName -ArgumentList "-S OfficeScrubScenario -AcceptEula -OfficeVersion All" -Wait -NoNewWindow
    }
    catch { return "Error Starting SaRA:`n$($Error[0])" }

    # Cleans
    try { Write-Output "Cleaning up."
        Remove-Item "$env:temp\SaRACmd_*.zip"
        Remove-Item "$env:temp\SaRA" -Force
    }
    catch { return "Error Cleaning up:`n$($Error[0])" }
    } else {
        return "Closing."
    }

}



Remove-Office