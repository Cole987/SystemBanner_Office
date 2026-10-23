@echo off
rem removes the SystemBanner Office add-in, run as an administrator; close Word, Excel and PowerPoint first
setlocal
cd /d "%~dp0"
set TARGET=%ProgramFiles%\SystemBanner\Office
if exist "%WINDIR%\SysWOW64\regedit.exe" (
    "%WINDIR%\SysWOW64\regedit.exe" /s SystemBannerOffice_remove.reg
) else (
    regedit /s SystemBannerOffice_remove.reg
)
if exist "%WINDIR%\Microsoft.NET\Framework64\v2.0.50727\RegAsm.exe" (
    "%WINDIR%\regedit.exe" /s SystemBannerOffice_remove.reg
    "%WINDIR%\Microsoft.NET\Framework64\v2.0.50727\RegAsm.exe" /nologo /unregister "%TARGET%\SystemBannerOffice.dll"
)
"%WINDIR%\Microsoft.NET\Framework\v2.0.50727\RegAsm.exe" /nologo /unregister "%TARGET%\SystemBannerOffice.dll"
del /q "%TARGET%\SystemBannerOffice.dll" "%TARGET%\SystemBannerOffice.Markings.dll"
rmdir "%TARGET%" 2>nul
echo SystemBanner Office add-in removed.
