
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
[Console]::Title = 'OfficeScrubber'
[Console]::ForegroundColor = 'White'
[Console]::BackgroundColor = 'Black'

#region Functions
function Remove-Office {
    If (!$InstalledOffice) {

        $x32 = ${env:ProgramFiles(x86)} + "\Microsoft Office"
        $x64 = $env:ProgramFiles + "\Microsoft Office"
        $OK = $true
                
        if (Test-Path -Path $x32) {$Excel32 = Get-ChildItem -Recurse -Path $x32 -Filter "EXCEL.EXE"}
        if (Test-Path -Path $x64) {$Excel64 = Get-ChildItem -Recurse -Path $x64 -Filter "EXCEL.EXE"}
        if ($Excel32) {$Excel = $Excel32}
        if ($Excel64) {$Excel = $Excel64}
        if ($Excel32 -and $Excel64) {"Error: x32 and x64 installation found." ; $Excel32.Fullname ; $Excel64.Fullname ; $OK = $false}
        if ($Excel.Count -gt 1) {"Error: More than one Excel.exe found." ; $Excel.Fullname ; $OK = $false}
        if ($Excel.Count -eq 0) {"Error: Excel.exe not found." ; $OK = $false}
        
        
        if ($OK) {
            $RegPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
            $DisplayVersion = Get-ItemProperty -Path $RegPath -Name "DisplayVersion" | Where-Object {$_.DisplayVersion -eq $Excel.VersionInfo.ProductVersion -and $_.PSChildName -notlike "{*}"}
            $Office = Get-ItemProperty -Path $DisplayVersion.PSPath
            $InstalledOffice = $Office | ForEach-Object {$_.DisplayName + $(if ($_.InstallLocation -eq $x32) {" 32 Bit "} else {" 64 Bit "})  + $_.DisplayVersion}
        }
    }
    
    If ($null -eq $InstalledOffice) {
        $InstalledOffice = "No Office Products Detected."
    }

    Write-Output "
    
   Office Removal Tool

Office Products Installed:
$InstalledOffice
    
What will happen?
The latest version of the Microsoft Support and Recovery Assistant (SaRA) tool will be downloaded and run using the commands below:

Downloaded from https://aka.ms/SaRA_EnterpriseVersionFiles
SaRACMD.exe -OfficeScrubScenario -AcceptEula -OfficeVersion All 
    
THIS TOOL WILL REMOVE EVERY VERSION OF OFFICE.
     THIS IS THE ONLY CONFIRMATION"

    try {
        Write-Output "Downloading SaRA."
        Invoke-WebRequest https://aka.ms/SaRA_EnterpriseVersionFiles -OutFile "$env:temp\"
    }
    catch {
        return "Error downloading SaRA: `n$($Error[0]) "
    }


    try {
        Write-Output "Extracting files."
        Expand-Archive -Path "$env:temp\SaRACmd_*.zip" -DestinationPath "$env:temp\SaRA" -Force
    }
    catch {
        return "Error expanding SaRA: `n$($Error[0]) "
    }
    
    try {
        $SaRAcmdexe = (Get-ChildItem "$env:temp\SaRA" -Include SaRAcmd.exe -Recurse).FullName
        Write-Output "Starting SaRA..."
        Start-Process $SaRAcmdexe -ArgumentList "-S OfficeScrubScenario -AcceptEula -OfficeVersion All"
    }
    catch {
        return "Error Starting SaRA: `n$($Error[0])"
    }
}

function Get-Question {
    $answer = Read-Host -Prompt "Would you like to run SaRA? (Y/N)"
    switch ($answer.ToUpper()) {
        'Y' { Remove-Office }
        'N' { return }
        Default { Get-Question }
    }
}

#endregion

Remove-Office