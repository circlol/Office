$Border = "*" * [Console]::WindowWidth
$OS = [Environment]::OSVersion.Platform
$DownloadDirectory = Join-Path -Path $PWD -ChildPath "Office Downloads"
if (-not (Test-Path -Path $DownloadDirectory)) { New-Item -Path $DownloadDirectory -ItemType Directory | Out-Null }
$EnableVerbose = $True
$Logo = @"
 ██████╗ ███████╗███████╗██╗ ██████╗███████╗    ██████╗  ██████╗ ██╗    ██╗███╗   ██╗██╗      ██████╗  █████╗ ██████╗ 
██╔═══██╗██╔════╝██╔════╝██║██╔════╝██╔════╝    ██╔══██╗██╔═══██╗██║    ██║████╗  ██║██║     ██╔═══██╗██╔══██╗██╔══██╗
██║   ██║█████╗  █████╗  ██║██║     █████╗      ██║  ██║██║   ██║██║ █╗ ██║██╔██╗ ██║██║     ██║   ██║███████║██║  ██║
██║   ██║██╔══╝  ██╔══╝  ██║██║     ██╔══╝      ██║  ██║██║   ██║██║███╗██║██║╚██╗██║██║     ██║   ██║██╔══██║██║  ██║
╚██████╔╝██║     ██║     ██║╚██████╗███████╗    ██████╔╝╚██████╔╝╚███╔███╔╝██║ ╚████║███████╗╚██████╔╝██║  ██║██████╔╝
 ╚═════╝ ╚═╝     ╚═╝     ╚═╝ ╚═════╝╚══════╝    ╚═════╝  ╚═════╝  ╚══╝╚══╝ ╚═╝  ╚═══╝╚══════╝ ╚═════╝ ╚═╝  ╚═╝╚═════╝ 
"@
$ErrorMessage = @{
    1 = "Invalid selection. Please try again."
    2 = "Invalid key selected. Returning to Main Menu."
    3 = "Offline installer not available for this version."
    4 = "Online installer not available for this version."
}

<#####################################################################################

                                Office Downloader


            Office versions are automatically added into program if following example.

            Example:
    "YearVersion"           = @{
        Name            = "Microsoft Office Year Version"
        OffFileName     = "YearVersion.img/.iso/.exe/.msi/.zip"
        OnFileName64    = "YearVersionArch.exe"
        OnFileName86    = "YearVersionArch.exe"
        OfflineLink     = "URL"
        OnlineLink64    = "URL"
        OnlineLink86    = "URL"
    }

#####################################################################################>

$Office =  @{
    "365Home"           = @{
        Name            = "Microsoft Office 365 Home"
        OffFileName     = "O365HomePremRetail.img"
        OnFileName64    = "O365HomePremRetail_x64.exe"
        OnFileName86    = "O365HomePremRetail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/O365HomePremRetail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365HomePremRetail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365HomePremRetail&platform=x86&language=en-us&version=O16GA"
    }
    "365Business"       = @{
        Name            = "Microsoft Office 365 Business"
        OffFileName     = "O365BusinessRetailRetail.img"
        OnFileName64    = "O365BusinessRetailRetail_x64.exe"
        OnFileName86    = "O365BusinessRetailRetail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/O365BusinessRetail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365BusinessRetail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365BusinessRetail&platform=x86&language=en-us&version=O16GA"
    }
    "365ProPlus"        = @{
        Name            = "Microsoft Office 365 Professional Plus"
        OffFileName     = "O365ProPlusRetail.img"
        OnFileName64    = "O365ProPlusRetail_x64.exe"
        OnFileName86    = "O365ProPlusRetail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/O365ProPlusRetail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365ProPlusRetail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365ProPlusRetail&platform=x86&language=en-us&version=O16GA"
    }
    "2024Home"          = @{
        Name            = "Microsoft Office 2024 Home"
        OffFileName     = "Home2024Retail.img"
        OnFileName64    = "Home2024Retail_x64.exe"
        OnFileName86    = "Home2024Retail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/Home2024Retail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Home2024Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Home2024Retail&platform=x86&language=en-us&version=O16GA"
    }
    "2024HomeBusiness"  = @{
        Name            = "Microsoft Office 2024 Home and Business"
        OffFileName     = "HomeBusiness2024Retail.img"
        OnFileName64    = "HomeBusiness2024Retail_x64.exe"
        OnFileName86    = "HomeBusiness2024Retail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/HomeBusiness2024Retail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusiness2024Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusiness2024Retail&platform=x86&language=en-us&version=O16GA"
    }
    "2024ProPlus"       = @{
        Name            = "Microsoft Office 2024 Professional Plus"
        OffFileName     = "ProPlus2024Retail.img"
        OnFileName64    = "ProPlus2024Retail_x64.exe"
        OnFileName86    = "ProPlus2024Retail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2024Retail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2024Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2024Retail&platform=x86&language=en-us&version=O16GA"
    }
    "2021HomeBusiness"  = @{
        Name            = "Microsoft Office 2021 Home and Business"
        OffFileName     = "HomeBusiness2021Retail.img"
        OnFileName64    = "HomeBusiness2021Retail_x64.exe"
        OnFileName86    = "HomeBusiness2021Retail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/HomeBusiness2021Retail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusiness2021Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusiness2021Retail&platform=x86&language=en-us&version=O16GA"
    }
    "2021HomeStudent"   = @{
        Name            = "Microsoft Office 2021 Home and Student"
        OffFileName     = "HomeStudent2021Retail.img"
        OnFileName64    = "HomeStudent2021Retail_x64.exe"
        OnFileName86    = "HomeStudent2021Retail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/HomeStudent2021Retail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudent2021Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudent2021Retail&platform=x86&language=en-us&version=O16GA"
    }
    "2021Pro"           = @{
        Name            = "Microsoft Office 2021 Professional"
        OffFileName     = "Professional2021Retail.img"
        OnFileName64    = "Professional2021Retail_x64.exe"
        OnFileName86    = "Professional2021Retail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/Professional2021Retail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Professional2021Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Professional2021Retail&platform=x86&language=en-us&version=O16GA"
    }
    "2021ProPlus"       = @{
        Name            = "Microsoft Office 2021 Professional Plus"
        OffFileName     = "ProPlus2021Retail.img"
        OnFileName64    = "ProPlus2021Retail_x64.exe"
        OnFileName86    = "ProPlus2021Retail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2021Retail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2021Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2021Retail&platform=x86&language=en-us&version=O16GA"
    }
    "2019HomeBusiness"  = @{
        Name            = "Microsoft Office 2019 Home and Business"
        OffFileName     = "HomeBusiness2019Retail.img"
        OnFileName64    = "HomeBusiness2019Retail_x64.exe"
        OnFileName86    = "HomeBusiness2019Retail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/HomeBusiness2019Retail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusiness2019Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusiness2019Retail&platform=x86&language=en-us&version=O16GA"
    }
    "2019HomeStudent"   = @{
        Name            = "Microsoft Office 2019 Home and Student"
        OffFileName     = "HomeStudent2019Retail.img"
        OnFileName64    = "HomeStudent2019Retail_x64.exe"
        OnFileName86    = "HomeStudent2019Retail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/HomeStudent2019Retail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudent2019Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudent2019Retail&platform=x86&language=en-us&version=O16GA"
    }
    "2019Pro"           = @{
        Name            = "Microsoft Office 2019 Professional"
        OffFileName     = "Professional2019Retail.img"
        OnFileName64    = "Professional2019Retail_x64.exe"
        OnFileName86    = "Professional2019Retail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/Professional2019Retail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Professional2019Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Professional2019Retail&platform=x86&language=en-us&version=O16GA"
    }
    "2019ProPlus"       = @{
        Name            = "Microsoft Office 2019 Professional Plus"
        OffFileName     = "ProPlus2019Retail.img"
        OnFileName64    = "ProPlus2019Retail_x64.exe"
        OnFileName86    = "ProPlus2019Retail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2019Retail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2019Retail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2019Retail&platform=x86&language=en-us&version=O16GA"
    }
    "2016HomeBusiness"  = @{
        Name            = "Microsoft Office 2016 Home and Business"
        OffFileName     = "HomeBusiness2016Retail.img"
        OnFileName64    = "HomeBusiness2016Retail_x64.exe"
        OnFileName86    = "HomeBusiness2016Retail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/HomeBusinessRetail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusinessRetail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusinessRetail&platform=x86&language=en-us&version=O16GA"
    }
    "2016HomeStudent"   = @{
        Name            = "Microsoft Office 2016 Home and Student"
        OffFileName     = "HomeStudent2016Retail.img"
        OnFileName64    = "HomeStudent2016Retail_x64.exe"
        OnFileName86    = "HomeStudent2016Retail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/HomeStudentRetail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudentRetail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudentRetail&platform=x86&language=en-us&version=O16GA"
    }
    "2016Pro"           = @{
        Name            = "Microsoft Office 2016 Professional"
        OffFileName     = "Professional2016Retail.img"
        OnFileName64    = "Professional2016Retail_x64.exe"
        OnFileName86    = "Professional2016Retail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProfessionalRetail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProfessionalRetail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProfessionalRetail&platform=x64&language=en-us&version=O16GA"
    }
    "2016ProPlus"       = @{
        Name            = "Microsoft Office 2016 Professional Plus"
        OffFileName     = "ProPlus2016Retail.img"
        OnFileName64    = "ProPlus2016Retail_x64.exe"
        OnFileName86    = "ProPlus2016Retail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlusRetail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x64&language=en-us&version=O16GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x86&language=en-us&version=O16GA"
    }
    "2013HomeBusiness"  = @{
        Name            = "Microsoft Office 2013 Home and Business"
        OffFileName     = "HomeBusiness2013Retail.img"
        OnFileName64    = "HomeBusiness2013Retail_x64.exe"
        OnFileName86    = "HomeBusiness2013Retail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/39168d7e-077b-48e7-872c-b232c3e72675/media/en-us/HomeBusinessRetail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusinessRetail&platform=x64&language=en-us&version=O15GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusinessRetail&platform=x86&language=en-us&version=O15GA"
    }
    "2013HomeStudent"   = @{
        Name            = "Microsoft Office 2013 Home and Student"
        OffFileName     = "HomeStudent2013Retail.img"
        OnFileName64    = "HomeStudent2013Retail_x64.exe"
        OnFileName86    = "HomeStudent2013Retail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/39168d7e-077b-48e7-872c-b232c3e72675/media/en-us/HomeStudentRetail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudentRetail&platform=x64&language=en-us&version=O15GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudentRetail&platform=x86&language=en-us&version=O15GA"
    }
    "2013Pro"           = @{
        Name            = "Microsoft Office 2013 Professional"
        OffFileName     = "Professional2013Retail.img"
        OnFileName64    = "Professional2013Retail_x64.exe"
        OnFileName86    = "Professional2013Retail_x86.exe"
        OfflineLink     = "https://officecdn.microsoft.com/db/39168d7e-077b-48e7-872c-b232c3e72675/media/en-us/ProfessionalRetail.img"
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProfessionalRetail&platform=x64&language=en-us&version=O15GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProfessionalRetail&platform=x86&language=en-us&version=O15GA"
    }
    "2013ProPlus"       = @{
        Name            = "Microsoft Office 2013 Professional Plus"
        OffFileName     = ""
        OnFileName64    = "ProPlus2013Retail_x64.exe"
        OnFileName86    = "ProPlus2013Retail_x86.exe"
        OfflineLink     = $False
        OnlineLink64    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x64&language=en-us&version=O15GA"
        OnlineLink86    = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x86&language=en-us&version=O15GA"
    }
    "2010ProPlus"       = @{
        Name            = "Microsoft Office 2010 Professional Plus"
        OffFileName     = "ProPlus2010Retail.iso"
        OnFileName64    = ""
        OnFileName86    = ""
        OfflineLink     = "https://drive.massgrave.dev/SW_DVD5_Office_Professional_Plus_2010w_SP1_64Bit_English_CORE_MLF_X17-76756.ISO"
        OnlineLink64    = $False
        OnlineLink86    = $False
    }
    "2007Enterprise"    = @{
        Name            = "Microsoft Office 2007 Enterprise"
        OnFileName64    = ""
        OnFileName86    = ""
        OffFileName     = "Enterprise2007.iso"
        OnlineLink64    = $False
        OnlineLink86    = $False
        OfflineLink     = "https://drive.massgrave.dev/en_office_enterprise_2007_united_states_x86_cd_481472.iso"
    }
}

function Select-Folder {
    If ($OS -eq "Win32NT") {

        Add-Type -AssemblyName System.Windows.Forms
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "Select a folder"
        $folderDialog.RootFolder = [System.Environment+SpecialFolder]::Desktop
        
        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedPath = $folderDialog.SelectedPath
            return $selectedPath
        } else {
            return "No folder selected."
        }
    } elseif ($OS -eq "Unix") {
        $selectedPath = zenity --file-selection --directory --title="Select a folder"
        return $selectedPath
    }
}
function Select-IndividualApps {

    $Check = "✔"
    $Cross = "❌"

    $Apps = @{
        "Access" = @{
            Name = "Microsoft Access"
        }
        "Word" = @{
            Name = "Microsoft Word"
        }
        "Excel" = @{
            Name = "Microsoft Excel"
        }
        "Publisher" = @{
            Name = "Microsoft Publisher"
        }
        "Powerpoint" = @{
            Name = "Microsoft Powerpoint"
        }
        "Visio" = @{
            Name = "Microsoft Visio"
        }
        "Outlook" = @{
            Name = "Microsoft Outlook"
        }
        "OneDrive" = @{
            Name = "Microsoft OneDrive"
        }
        "Defender" = @{
            Name = "Microsoft Defender"
        }

    }

}
Function Show-Logo {
    Write-Output "`n$border`n`n`n$Logo`n"
}
function Show-MainMenu {
    Write-Output "`n$Border`n"
    Write-Host "      Select the Office version:" -ForegroundColor Cyan
    
    $i = 1
    $sortedKeys = $Office.Keys | Sort-Object -descending
    foreach ($key in $SortedKeys) {
        Write-Output "         $i) $($Office[$key].Name)"
        $i++
    }
    
    Write-Host "`n`n      Current download directory : $DownloadDirectory" -ForegroundColor Cyan
    Write-Output "      $i) Change download directory"
    $i++
    Write-Output "      $i) Exit`n`n$Border"

    $choice = [int](Read-Host "Enter your choice (1-$i)")
    #Write-Output "User selected: $choice, accessing index: $($choice - 1)"
    
    
    if ($choice -eq $i) { 
        return
    } elseif ($choice -eq ($i -1)) {
        $OldFolder = $DownloadDirectory
        $NewFolder = Select-Folder
        If (Test-Path $NewFolder) {
            If ($NewFolder -ne $OldFolder){
                $DownloadDirectory = $NewFolder
            }
        }
        Show-MainMenu
    } elseif ($choice -ge 1 -and $choice -lt $i) {
        $selectedKey = $sortedKeys[$choice - 1]
        Show-SubMenu $selectedKey
    } else {
        Write-Host $ErrorMessage.1 -ForegroundColor Red
        Show-MainMenu
    }
}
function Show-SubMenu( [string]$selectedKey )  {
    if (-not $Office.ContainsKey($selectedKey)) {
        Write-Host $ErrorMessage.2 -ForegroundColor Red
        Show-MainMenu
    }

    Write-Host "`nSelect the download type for $($Office[$selectedKey].Name):" -ForegroundColor Cyan
    Write-Host "`nSelected Option - $($Office[$selectedKey].Name)" -ForegroundColor Yellow

    $options = @{}
    $i = 1

    if ($Office[$selectedKey].OnlineLink64) {
        Write-Output "$i) Online Installer (x64)"
        $options[$i] = @{ Link = $Office[$selectedKey].OnlineLink64; Arch = "x64" ; Filename = $Office[$selectedKey].OnFileName64}
        $i++
    }
    if ($Office[$selectedKey].OnlineLink86) {
        Write-Output "$i) Online Installer (x86)"
        $options[$i] = @{ Link = $Office[$selectedKey].OnlineLink86; Arch = "x86" ; Filename = $Office[$selectedKey].OnFileName86}
        $i++
    }
    if ($Office[$selectedKey].OfflineLink) {
        Write-Output "$i) Offline Installer"
        $options[$i] = @{ Link = $Office[$selectedKey].OfflineLink; Arch = "Offline" ; Filename = $Office[$selectedKey].OffFileName}
        $i++
    }

    Write-Output "$i) Back to Main Menu"

    $choice = [int](Read-Host "Enter your choice (1-$i)")
    if ($choice -eq $i) {
        Show-MainMenu
    } elseif ($options.ContainsKey($choice)) {
        Start-Download $options[$choice].Link $selectedKey $options[$choice].Arch $options[$choice].Filename 
    } else {
        Write-Host $ErrorMessage.1 -ForegroundColor Red
        Show-SubMenu $selectedKey
    }
}
function Start-Activation {
    $Name = "MAS (Microsoft Activation Scripts)"
    $choice = Read-Host "Would you like to activate office with $Name ? (Y/N)"
    switch ($choice.ToUpper()) {
        'Y' {
            Write-Output "Starting activation with $Name"
            Invoke-RestMethod https://get.activated.win -Verbose:$EnableVerbose | Invoke-Expression -Verbose:$EnableVerbose
            return
        }
        'N' { return }
        default { return $ErrorMessage.1 }
    }
}
function Start-Download( [string]$url, [string]$selectedKey, [string]$type, [string]$filename ) {
    $outputPath = Join-Path -Path $DownloadDirectory -ChildPath $filename

    Write-Host "`nDownloading $($Office[$selectedKey].Name) ($type)..." -ForegroundColor Yellow
    try {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

        Invoke-WebRequest -Uri $url -OutFile $outputPath -Verbose:$EnableVerbose -ErrorAction Stop
        Write-Host "Download completed to path: $outputPath" -ForegroundColor Green
    } catch {
        Write-Host "Download failed: $($_.Exception.Message)" -ForegroundColor Red
        Show-MainMenu
    }

    #Start-Installer $OutputPath
    Open-Path $outputPath.Parent
    Start-Activation
    
    
    $choice = Read-Host "Go back to the main menu? (Y/N)"
    switch ($choice.ToUpper()) {
        'Y' {  Show-MainMenu  ;  return  }
        'N' {  Write-Output "Exiting"  ;  return  }
        'Exit' { Write-Output "Exiting"  ;  return }
        default {  return $ErrorMessage.1  }
    }
}

function Start-Installer( [string]$Installer ) {
    $choice = Read-Host "Would you like to start office installer named $Installer ? (Y/N)"
    switch ($choice.ToUpper()) {
        'Y' {
            try { 
                $fileDownloaded = Test-Path $Installer -ErrorAction SilentlyContinue
                If ($fileDownloaded -eq $True) {
                    Write-Host "Starting: $Installer" -ForegroundColor Yellow
                    Start-Process $Installer -Verbose:$EnableVerbose
                }
            } catch {
                Write-Host "Starting setup failed: $($_.Exception.Message)" -ForegroundColor Red
                Show-MainMenu
            }

            return
        }
        'N' {  return  }
        default {  Write-Output $ErrorMessage.1  }
    }
}


Clear-Host
Show-Logo
Show-MainMenu