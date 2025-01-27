$Border = "*" * [Console]::WindowWidth
$Message1 = "Invalid selection. Please try again."
$Message2 = "Invalid key selected. Returning to Main Menu."
$Message3 = "Offline Installer not available for this version."
$OfficeVersion =  @{
    "365Home"        = @{
        Name         = "Microsoft Office 365 Home"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365HomePremRetail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365HomePremRetail&platform=x86&language=en-us&version=O16GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/O365HomePremRetail.img"
    }
    "365Business"    = @{
        Name         = "Microsoft Office 365 Business"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365BusinessRetail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365BusinessRetail&platform=x86&language=en-us&version=O16GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/O365BusinessRetail.img"
    }
    "365ProPlus"     = @{
        Name         = "Microsoft Office 365 Professional Plus"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365ProPlusRetail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365ProPlusRetail&platform=x86&language=en-us&version=O16GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/O365ProPlusRetail.img"
    }

    "2024Home"       = @{
        Name         = "Microsoft Office 2024 Home"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Home2024Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Home2024Retail&platform=x86&language=en-us&version=O16GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/Home2024Retail.img"
    }
    "2024HomeBusiness"= @{
        Name         = "Microsoft Office 2024 Home and Business"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusiness2024Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusiness2024Retail&platform=x86&language=en-us&version=O16GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/HomeBusiness2024Retail.img"
    }
    "2024ProPlus"    = @{
        Name         = "Microsoft Office 2024 Professional Plus"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2024Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2024Retail&platform=x86&language=en-us&version=O16GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2024Retail.img"
    }

    "2021HomeBusiness"= @{
        Name         = "Microsoft Office 2021 Home and Business"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusiness2021Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusiness2021Retail&platform=x86&language=en-us&version=O16GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/HomeBusiness2021Retail.img"
    }
    "2021HomeStudent"= @{
        Name         = "Microsoft Office 2021 Home and Student"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudent2021Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudent2021Retail&platform=x86&language=en-us&version=O16GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/HomeStudent2021Retail.img"
    }
    "2021Pro"        = @{
        Name         = "Microsoft Office 2021 Professional"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Professional2021Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Professional2021Retail&platform=x86&language=en-us&version=O16GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/Professional2021Retail.img"
    }
    "2021ProPlus"    = @{
        Name         = "Microsoft Office 2021 Professional Plus"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2021Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2021Retail&platform=x86&language=en-us&version=O16GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2021Retail.img"
    }
    "2019HomeBusiness"= @{
        Name         = "Microsoft Office 2019 Home and Business"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusiness2019Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusiness2019Retail&platform=x86&language=en-us&version=O16GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/HomeBusiness2019Retail.img"
    }
    "2019HomeStudent"= @{
        Name         = "Microsoft Office 2019 Home and Student"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudent2019Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudent2019Retail&platform=x86&language=en-us&version=O16GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/HomeStudent2019Retail.img"
    }
    "2019Pro"        = @{
        Name         = "Microsoft Office 2019 Professional"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Professional2019Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Professional2019Retail&platform=x86&language=en-us&version=O16GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/Professional2019Retail.img"
    }
    "2019ProPlus"    = @{
        Name         = "Microsoft Office 2019 Professional Plus"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2019Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2019Retail&platform=x86&language=en-us&version=O16GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2019Retail.img"
    }

    "2016HomeBusiness" =    @{
        Name         = "Microsoft Office 2016 Home and Business"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusinessRetail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusinessRetail&platform=x86&language=en-us&version=O16GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/HomeBusinessRetail.img"
    }
    "2016HomeStudent"  = @{
        Name         = "Microsoft Office 2016 Home and Student"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudentRetail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudentRetail&platform=x86&language=en-us&version=O16GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/HomeStudentRetail.img"
    }
    "2016Pro"        = @{
        Name         = "Microsoft Office 2016 Professional"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProfessionalRetail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProfessionalRetail&platform=x64&language=en-us&version=O16GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProfessionalRetail.img"
    }
    "2016ProPlus"       = @{
        Name         = "Microsoft Office 2016 Professional Plus"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x86&language=en-us&version=O16GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlusRetail.img"
    }

    "2013HomeBusiness"  = @{
        Name = "Microsoft Office 2013 Home and Business"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusinessRetail&platform=x64&language=en-us&version=O15GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusinessRetail&platform=x86&language=en-us&version=O15GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/39168d7e-077b-48e7-872c-b232c3e72675/media/en-us/HomeBusinessRetail.img"
    }
    "2013HomeStudent"   = @{
        Name = "Microsoft Office 2013 Home and Student"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudentRetail&platform=x64&language=en-us&version=O15GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudentRetail&platform=x86&language=en-us&version=O15GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/39168d7e-077b-48e7-872c-b232c3e72675/media/en-us/HomeStudentRetail.img"
    }
    "2013Pro"        = @{
        Name         = "Microsoft Office 2013 Professional"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProfessionalRetail&platform=x64&language=en-us&version=O15GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProfessionalRetail&platform=x86&language=en-us&version=O15GA"
        OfflineLink  = "https://officecdn.microsoft.com/db/39168d7e-077b-48e7-872c-b232c3e72675/media/en-us/ProfessionalRetail.img"
    }
    "2013ProPlus"       = @{
        Name = "Microsoft Office 2013 Professional Plus"
        OnlineLink64 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x64&language=en-us&version=O15GA"
        OnlineLink86 = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x86&language=en-us&version=O15GA"
    }
}

function Show-MainMenu {
    Write-Output "Going to main menu"
    Write-Output "$Border`n"
    Write-Host "      Select the Office version:" -ForegroundColor Cyan
    $i = 1
    
    $sortedKeys = $OfficeVersion.Keys | Sort-Object
    foreach ($key in $SortedKeys) {
        Write-Host "         $i) $($OfficeVersion[$key].Name)"
        $i++
    }
    Write-Host "      $i) Exit"
    Write-Output "`n$Border"

    # Prompt user for choice
    $choice = Read-Host "Enter your choice (1-$i)"
    if ($choice -eq $i) { Exit } elseif ($choice -ge 1 -and $choice -lt $i) {
        # Map choice to key
        $selectedKey = $sortedKeys[$choice - 1]
        Show-SubMenu $selectedKey
    } else {
        # Returns to Main Menu with Invalid Selection
        Write-Host $Message1 -ForegroundColor Red
        Show-MainMenu
    }
}

function Show-SubMenu {
    param (
        [string]$selectedKey
    )

    # Debugging output to verify key resolution
    if (-not $OfficeVersion.ContainsKey($selectedKey)) {
        Write-Host $Message2 -ForegroundColor Red
        Show-MainMenu
        return
    }

    Write-Host "`nSelected Option - $($OfficeVersion[$selectedKey].Name)" -ForegroundColor Yellow
    Write-Host "Select the download type for $($OfficeVersion[$selectedKey].Name):" -ForegroundColor Cyan
    Write-Host "1) Online Installer (x64)"
    Write-Host "2) Online Installer (x86)"
    Write-Host "3) Offline Installer"
    Write-Host "4) Back to Main Menu"

    $choice = Read-Host "Enter your choice (1-4)"
    switch ($choice) {
        1 {
            Start-Download $OfficeVersion[$selectedKey].OnlineLink64 $selectedKey "x64"
        }
        2 {
            Start-Download $OfficeVersion[$selectedKey].OnlineLink86 $selectedKey "x86"
        }
        3 {
            if ($OfficeVersion[$selectedKey].ContainsKey("OfflineLink")) {
                Start-Download $OfficeVersion[$selectedKey].OfflineLink $selectedKey "Offline"
            } else {
                Write-Host $Message3 -ForegroundColor Red
                Show-SubMenu $selectedKey
            }
        }
        4 {
            Show-MainMenu
        }
        default {
            Write-Host $Message1 -ForegroundColor Red
            Show-SubMenu $selectedKey
        }
    }
}

function Start-Activation {
    # Prompt user for choice
    $choice = Read-Host "Would you like to activate office with MAS(Microsoft Activation Scripts)? (Y/N)"
    switch ($choice.ToUpper()) {
        'Y' {
            Write-Output "Starting activation with MAS(Microsoft Activation Scripts)"
            Invoke-RestMethod https://get.activated.win | Invoke-Expression
        }
        'N' {
            Write-Output "Skipping MAS(Microsoft Activation Scripts)"
            return
        }
        default {
            Write-Output $Message1
        }
    }
}

function Start-Download {
    param (
        [string]$url,
        [string]$selectedKey,
        [string]$type
    )

    $fileName = "$($OfficeVersion[$selectedKey].Name) - $type.exe"
    $outputPath = Join-Path -Path $PWD -ChildPath $fileName

    # Attempts to download setup
    Write-Host "`nDownloading $($OfficeVersion[$selectedKey].Name) ($type)..." -ForegroundColor Yellow
    try {
        Invoke-WebRequest -Uri $url -OutFile $outputPath
        Write-Host "Download completed: $outputPath" -ForegroundColor Green
    } catch {
        Write-Host "Download failed: $($_.Exception.Message)" -ForegroundColor Red
        Show-MainMenu
    }

    # Attempts to run setup
    try {
        $fileDownloaded = Test-Path $OutputPath -ErrorAction SilentlyContinue
        If ($fileDownloaded -eq $True) {
            Write-Host "Starting: $outputPath" -ForegroundColor Yellow
            Start-Process $outputPath
        }
    }
    catch {
        Write-Host "Starting setup failed: $($_.Exception.Message)" -ForegroundColor Red
        Show-MainMenu
    }

    Start-Activation

    Show-MainMenu
}

Clear-Host

# Start the script
Show-MainMenu