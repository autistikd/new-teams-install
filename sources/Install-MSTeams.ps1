[CmdletBinding()]
param (
    $EXE = "Teamsbootstrapper.exe",
    $MSIX = "MSTeams-x64.msix",
    [switch]$Offline,
    [switch]$Uninstall,
    [Alias("TryFix")]
    [switch]$ForceInstall,
    [switch]$Teamslnk
)

function Install-MSTeams {
    param (
        [switch]$Offline
    )
    if ($Offline) {
        $Result = & "$EXEFolder\$EXE" -p --offline-install "$MSIXFolder\$MSIX"
    } else {
        $Result = & "$EXEFolder\$EXE" -p
    }
    $ResultPSO = try { $Result | ConvertFrom-Json } catch {$null}
    if ($null -ne $ResultPSO) {
        return $ResultPSO
    } else {
        return $Result
    }
}

function Uninstall-MSTeams {
    $Appx = Get-AppxPackage -AllUsers | Where-Object {$_.Name -eq "MSTeams"}
    if ($Appx) {
        $Appx | Remove-AppxPackage -AllUsers
    }
    $Result = & "$EXEFolder\$EXE" -x
    $ResultPSO = try { $Result | ConvertFrom-Json } catch {$null}
    if ($null -ne $ResultPSO) {
        return $ResultPSO
    } else {
        return $Result
    }
}

function IsAppInstalled {
    param (
        $AppName = "MSTeams"
    )
    $Appx = Get-AppxPackage -AllUsers | Where-Object {$_.Name -eq $AppName}
    $ProvApp = Get-ProvisionedAppPackage -Online | Where-Object {$_.DisplayName -eq $AppName}
    if ($Appx) {
        Write-Output "$AppName AppxPackage is currently installed"
    } else {
        Write-Output "$AppName AppxPackage is NOT installed"
    }
    if ($ProvApp) {
        Write-Output "$AppName ProvisionedAppPackage is currently installed"
    } else {
        Write-Output "$AppName ProvisionedAppPackage is NOT installed"
    }
}

$EXEFolder = $PSScriptRoot
$MSIXFolder = $PSScriptRoot


if ($Teamslnk) {

	$publicDesktopPath = [Environment]::GetFolderPath("CommonDesktopDirectory")
    $shortcutPath = Join-Path $publicDesktopPath "Microsoft Teams.lnk"

    $iconPath = Join-Path $PSScriptRoot "teams-icon-256x256.ico"

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)

    $shortcut.TargetPath = "explorer.exe"
    $shortcut.Arguments = "shell:AppsFolder\MSTeams_8wekyb3d8bbwe!MSTeams"

    $shortcut.IconLocation = $iconPath

    $shortcut.Save()

}

if (-not(Test-Path -Path $EXEFolder\$EXE)) {
    Write-Output "Failed to find $EXE"
    exit 2
}

$EXEinfo = Get-ChildItem -Path "$EXEFolder\$EXE"

if ($Uninstall) {
    Write-Output "Attempting to uninstall MSTeams"

	$shortcutPath = Join-Path $publicDesktopPath "Microsoft Teams.lnk"
	if (Test-Path $shortcutPath) {
        Remove-Item $shortcutPath -Force
    }

    IsAppInstalled "MSTeams"

    $result = Uninstall-MSTeams
    $Appx = Get-AppxPackage -AllUsers | Where-Object {$_.Name -eq "MSTeams"}
    $ProvApp = Get-ProvisionedAppPackage -Online | Where-Object {$_.DisplayName -eq "MSTeams"}

    if (!$Appx -and !$ProvApp) {
        Remove-Item HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\Teams -Force -ErrorAction SilentlyContinue
        Write-Output "MSTeams successfully uninstalled"
        exit 0
    } else {
        Write-Output "Error uninstalling MSTeams"
        IsAppInstalled "MSTeams"
        exit 1
    }
}

if ($Offline) {
    if (-not(Test-Path -Path "$MSIXFolder\$MSIX")) {
        Write-Output "Offline parameter specified but failed to find $MSIX"
        exit 2
    }
    Write-Output "Attempting to install MSTeams offline"

    if ($ForceInstall) {
        Write-Output "ForceInstall: uninstalling before install"
        IsAppInstalled "MSTeams"
        $result = Uninstall-MSTeams
        Remove-Item HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\Teams -Force -ErrorAction SilentlyContinue
    }

    $result = Install-MSTeams -Offline
    if ($result.Success) {
        Write-Output "MSTeams successfully installed offline"
        exit 0
    } else {
        Write-Output "Error installing MSTeams offline"
        IsAppInstalled "MSTeams"
        exit 1
    }
} else {
    Write-Output "Attempting to install MSTeams online"
    if ($ForceInstall) {
        Write-Output "ForceInstall: uninstalling before install"
        $result = Uninstall-MSTeams
        Remove-Item HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\Teams -Force -ErrorAction SilentlyContinue
    }

    $result = Install-MSTeams
    if ($result.Success) {
        Write-Output "MSTeams successfully installed online"
        exit 0
    } else {
        Write-Output "Error installing MSTeams online"
        IsAppInstalled "MSTeams"
        exit 1
    }
}
