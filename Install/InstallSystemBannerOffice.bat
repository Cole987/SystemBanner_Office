@echo off
rem installs the SystemBanner Office add-in for every user of the computer, run as an administrator
rem Windows 2000 has no reg.exe: the Office keys are written with regedit /s and a .reg file, like SystemBanner for Windows 2000
setlocal
cd /d "%~dp0"
set TARGET=%ProgramFiles%\SystemBanner\Office
if not exist "%TARGET%" mkdir "%TARGET%"
copy /y SystemBannerOffice.dll "%TARGET%\" >nul || goto failed
copy /y SystemBannerOffice.Markings.dll "%TARGET%\" >nul || goto failed

rem 32 bit Office (every Office from 2000 to 2007, and 32 bit Office 2010)
"%WINDIR%\Microsoft.NET\Framework\v2.0.50727\RegAsm.exe" /nologo /codebase "%TARGET%\SystemBannerOffice.dll" || goto failed
rem on 64 bit Windows, 32 bit Office reads the 32 bit registry view, written by the 32 bit regedit
if exist "%WINDIR%\SysWOW64\regedit.exe" (
    "%WINDIR%\SysWOW64\regedit.exe" /s SystemBannerOffice.reg
) else (
    regedit /s SystemBannerOffice.reg
)
rem 64 bit Office 2010 on 64 bit Windows: the 64 bit registry view needs its own registration
if exist "%WINDIR%\Microsoft.NET\Framework64\v2.0.50727\RegAsm.exe" (
    "%WINDIR%\Microsoft.NET\Framework64\v2.0.50727\RegAsm.exe" /nologo /codebase "%TARGET%\SystemBannerOffice.dll" || goto failed
    "%WINDIR%\regedit.exe" /s SystemBannerOffice.reg
)
echo SystemBanner Office add-in installed: restart Word, Excel and PowerPoint.
exit /b 0
:failed
echo Installation failed. Run this script as an administrator.
exit /b 1
