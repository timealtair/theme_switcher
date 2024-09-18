# Define the registry path for theme settings
$themeKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"

# Function to get the current theme settings
function Get-CurrentThemeSettings {
    $appsTheme = Get-ItemPropertyValue -Path $themeKey -Name AppsUseLightTheme
    $systemTheme = Get-ItemPropertyValue -Path $themeKey -Name SystemUsesLightTheme
    return @{ AppsTheme = $appsTheme; SystemTheme = $systemTheme }
}

# Function to save the state of open Explorer windows
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

# Function to reopen Explorer windows with their saved state
function Reopen-ExplorerWindows {
    param (
        [array]$windowStates
    )

    foreach ($state in $windowStates) {
        Start-Process explorer.exe $state.Path
        # Start-Sleep -Seconds 2 # Wait to ensure the process has time to start
    }
}

# Function to toggle the theme
function Toggle-Theme {
    $currentSettings = Get-CurrentThemeSettings

    if ($currentSettings.AppsTheme -eq 1 -and $currentSettings.SystemTheme -eq 1) {
        # Switch to dark mode
        Set-ItemProperty -Path $themeKey -Name AppsUseLightTheme -Value 0
        Set-ItemProperty -Path $themeKey -Name SystemUsesLightTheme -Value 0
        # Write-Output "Switched to dark theme"
    } else {
        # Switch to light mode
        Set-ItemProperty -Path $themeKey -Name AppsUseLightTheme -Value 1
        Set-ItemProperty -Path $themeKey -Name SystemUsesLightTheme -Value 1
        # Write-Output "Switched to light theme"
    }

    # Save the state of open Explorer windows
    $explorerStates = Save-ExplorerWindows

    # Restart Explorer to apply theme changes
    Stop-Process -Name explorer -Force
    # Start-Sleep -Seconds 5 # Wait for Explorer to fully terminate

    # Reopen the saved Explorer windows
    Reopen-ExplorerWindows -windowStates $explorerStates
}

# Main script
try {
    Toggle-Theme
} catch {
    Write-Error "An error occurred: $_"
}
