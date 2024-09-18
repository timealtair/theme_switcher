$themeKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"

function Get-CurrentThemeSettings {
    $appsTheme = Get-ItemPropertyValue -Path $themeKey -Name AppsUseLightTheme
    $systemTheme = Get-ItemPropertyValue -Path $themeKey -Name SystemUsesLightTheme
    return @{ AppsTheme = $appsTheme; SystemTheme = $systemTheme }
}

function Save-ExplorerWindows {
    $shellWindows = New-Object -ComObject Shell.Application
    $windows = $shellWindows.Windows() | Where-Object { $_.Name -ne "" }
    $windowStates = @()

    foreach ($window in $windows) {
        $path = $window.Document.Folder.Self.Path
        $windowStates += [PSCustomObject]@{
            Path = $path
        }
    }

    return $windowStates
}

function Reopen-ExplorerWindows {
    param (
        [array]$windowStates
    )

    if ($windowStates.Count -eq 0) {
	$process = Get-Process | Where-Object { $_.ProcessName -like "*explorer*" }

	if ($process -ne $null) {
	} else {
	    start explorer.exe
	}
    } else {
        foreach ($state in $windowStates) {
            Start-Process explorer.exe $state.Path
        }
    }
}

function Toggle-Theme {
    $currentSettings = Get-CurrentThemeSettings

    if ($currentSettings.AppsTheme -eq 1 -and $currentSettings.SystemTheme -eq 1) {
        Set-ItemProperty -Path $themeKey -Name AppsUseLightTheme -Value 0
        Set-ItemProperty -Path $themeKey -Name SystemUsesLightTheme -Value 0
    } else {
        Set-ItemProperty -Path $themeKey -Name AppsUseLightTheme -Value 1
        Set-ItemProperty -Path $themeKey -Name SystemUsesLightTheme -Value 1
    }

    $explorerStates = Save-ExplorerWindows

$process = Get-Process | Where-Object { $_.ProcessName -like "*explorer*" }

if ($process -ne $null) {
    $processId = $process.Id
    Stop-Process -Id $processId -Force
} else {
}

    Reopen-ExplorerWindows -windowStates $explorerStates
}

try {
    Toggle-Theme
} catch {
    Write-Error "An error occurred: $_"
}
