# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## 2025-10-07
### Added
- Added HTML file association when other browsers (instead of chrome or edge) are default
### Changed
- Improved Orchestrator script
- Improved notepad startup script
### Fixed
- Added missing startup-scripts folder

## 2025-10-06
### Added
- Added startup functionality, making it possible to run multiple startup Scripts in the Sandbox. Improves (and kinda Fixes) [#11]https://github.com/Joly0/Run-in-Sandbox/issues/11
- Added some useful startup script (Notepad, Context Menu customization, fix for slow MSI installation); Thanks to ThioJoe https://github.com/ThioJoe/Windows-Sandbox-Tools
- Added code to unblock files in the host folder that is mapped to the sandbox. Fixes [#10]https://github.com/Joly0/Run-in-Sandbox/issues/10
### Fixed
- Fixed running multiple Apps through SDBApp
### Changed
- Changed some startup behaviour and adjusted showing some cmd/powershell windows and hiding some


## 2025-08-20
### Security
- **Enhanced 7-Zip Integration:** Removed bundled 7-Zip executables for improved security
- **On-Demand 7-Zip:** Now uses host-installed 7-Zip when available, or downloads latest version from GitHub releases
- **Smart Caching:** Automatic download and caching of 7-Zip installers with 7-day refresh cycle
- **Offline Support:** Works offline using cached installers when network is unavailable
- **Always Current:** Ensures latest 7-Zip version is used, eliminating security vulnerabilities from outdated bundled files
### Changed
- Modified `RunInSandbox.ps1` to detect and mount host 7-Zip installation
- Updated installation process to cache latest 7-Zip installer during setup
- Enhanced `New-WSB` function to support additional mapped folders
- Reduced project size by ~2MB by removing bundled 7-Zip files
### Added
- `Find-Host7Zip` function for detecting system 7-Zip installations
- `Get-Latest7ZipDownloadUrl` function for GitHub API integration
- `Update-7ZipCache` function for smart installer management
- `Ensure-7ZipCache` function for offline-capable cache validation
- Version tracking and age-based cache refresh mechanism


## 2025-05-07
### Fixed
- Fixed running batch files in Sandbox
- Fixed label for "Run BAT in Sandbox"


## 2025-01-06
### Fixed
- Fixed [#7]https://github.com/Joly0/Run-in-Sandbox/issues/7 and [#58]https://github.com/damienvanrobaeys/Run-in-Sandbox/issues/58
- Fixed indendation for wsb file
- Fixed [#8]https://github.com/Joly0/Run-in-Sandbox/issues/8


## 2024-11-8
### Changed
- Slightly adjusted RunInSandbox installer script
- Improved readme for RunInSandbox installation


## 2024-10-15
### Fixed
- Fixed [#56]https://github.com/damienvanrobaeys/Run-in-Sandbox/issues/56
### Changed
- Formatting improvements for RunInSandbox.ps1 script


## 2024-08-26
### Fixed
- Fixed Subcommands Entries for Powershell and VBS


## 2024-08-20
### Added
- Added common functions script
- Added deep-clean option for uninstalling Run-in-Sandbox
### Changed
- Rewritten install and uninstall script and exported common functions to separate file
### Fixed
- Slightly adjusted intunewin and sdbapp scripts


## 2024-05-22
### Added
- Added easy install Script
### Changed
- Improved console output and make it easier readable
- Improved readme and install steps


## 2024-05-14
### Added
-  Added some better error handling and checking for needed features
### Fixed
- Probably fixed [#4]https://github.com/Joly0/Run-in-Sandbox/issues/4
### Changed
- Improved the way, exe files are handled inside the sandbox


## 2023-07-14
### Fixed
- Finally fixed running intunewin with serviceUI and psexec
- Fixed [#40]https://github.com/damienvanrobaeys/Run-in-Sandbox/issues/40
- Fixed [#41]https://github.com/damienvanrobaeys/Run-in-Sandbox/issues/41
### Changed
- Changed formatting to OTBS using "Invoke-Formatter" cmdlet in "Script-Analyzer" module (On-going discussion [#44]https://github.com/damienvanrobaeys/Run-in-Sandbox/discussions/44) and applied some powershell best-pratices


## 2023-05-01
### Added
- Reimplemented running Intunewin as System using psexec (serviceui will stay)
### Fixed
- Fixed [#18]https://github.com/damienvanrobaeys/Run-in-Sandbox/issues/18


## 2023-05-03
### Added
- Added option to run .intunewin via sdbapp
### Changed
- Changed Intunewin_Content_File and Intunewin_Command_File to be parameters for IntuneWin_Install.ps1


## 2023-03-29
### Added
- Added context menu entry for opening PDF in Sandbox
### Changed
- Completely rewrote alot of code in Add_Structure.ps1


## 2023-03-22
### Added
- Added context menu entry for running CMD/BAT in Sandbox


## 2023-03-21
### Changed
- Readded 7z part and adjusted 7z reg key path


## 2023-03-20
### Changed
- Completly refactored RunInSandbox.ps1 to use switch instead of ifelse and rearranged alot of code
### Fixed
- Fixed some issues with loading iso´s, exe´s and zip´s
### Removed
- Removed 7z part of RunInSandbox.ps1 because non-functional


## 2023-03-07
### Added
- Added ServiceUI
### Changed
- Replaced PSexec with ServiceUI for intunewin sandbox
### Removed
- Removed PSexec in favor of ServiceUI


## 2023-03-06
### Added
- Added option to Sandbox_Config.xml to cleanup leftover .wsb file afterwards (default is true)
### Changed
- .wsb is not executed by the "Start-Process"-cmdlet with -wait parameter


## 2023-03-03
### Added
- Added -noprofile to powershell commands to improve performance
### Changed
- Applied formatting of scripts and applied best practices
### Fixed
- Fixed .ps1 conext menu


## 2021-11-16
### Added
- Add a context menu for running PS1 as system in Sandbox
- Add a context menu for running MSIX in Sandbox
- Add a context menu for running PPKG in Sandbox
- Add a context menu for opening URL in Sandbox
- Add a context menu for extracting ISO in Sandbox
- Add a context menu for extracting 7z file in Sandbox
### Fixed
- Fix a bug where context menu for PS1 does not appear on Windows 11 


## 2021-09-21
### Added
- Add a context menu for reg file, to run them in Sandbox
- Add ability to run multiple apps in the same Sandbox session


## 2021-08-03
### Added
- Add a context menu for intunewin file, to run them in Sandbox
- Add ability to choose which content menu to add


## 2021-07-27
### Changed
- Change the default path where WSB are saved after running Sandbox: now in %temp%


## 2021-07-21
### Changed
- Updated the GUI when running EXE or MSI for more understanding
- Updated the GUI when running PS1 for more understanding


## 2021-07-16
### Added
- The Add_Structure.ps1 will now create a restore point
- It will then check if Sources folder exists


## 2020-06-24
### Removed
- Temporarily removed the main file [#9](https://github.com/damienvanrobaeys/Run-in-Sandbox/issues/9)
### Changed
- Fixed detail language setting being French


## 2020-06-02 
### Added
 - Add new WSB config options for Windows 10 2004. These new settings can be managed in the **Sources\Run_in_Sandbox\Sandbox_Config.xml**
 - New options: AudioInput, VideoInput, ProtectedClient, PrinterRedirection, ClipboardRedirection, MemoryInMB


## 2020-05-19
### Added
- Added French, Italian, Spanish, English, and German languages for context menus. To configure language, edit **Main_Language** in **Sources\Run_in_Sandbox\Sandbox_Config.xml**