Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser

Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted

Install-Module -Name PS2EXE -Scope CurrentUser -Force

Invoke-PS2EXE -InputFile "source/theme_switcher.ps1" -OutputFile "ThemeSwitcher.exe" -NoConsole -IconFile "source/theme_switcher.ico"

