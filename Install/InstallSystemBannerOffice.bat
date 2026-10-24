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
rem Office 2000 to 2003 start the .NET runtime of their first managed add-in, which can be .NET 1.1 (Word and Excel 2000:
rem "File or assembly name SystemBannerOffice ... was not found"): a config file next to each program asks for .NET 2.0
for %%D in ("%ProgramFiles%\Microsoft Office\Office" "%ProgramFiles%\Microsoft Office\Office10" "%ProgramFiles%\Microsoft Office\Office11") do (
    for %%E in (WINWORD EXCEL POWERPNT) do (
        if exist "%%~D\%%E.EXE" if not exist "%%~D\%%E.EXE.config" copy /y SystemBannerOffice.exe.config "%%~D\%%E.EXE.config" >nul
        if exist "%%~D\%%E.EXE.config" find "SystemBanner" "%%~D\%%E.EXE.config" >nul || echo Warning: %%~D\%%E.EXE.config was not written by SystemBanner: add supportedRuntime v2.0.50727 to it, see SystemBannerOffice.exe.config.
    )
)
echo SystemBanner Office add-in installed: restart Word, Excel and PowerPoint.
exit /b 0
:failed
echo Installation failed. Run this script as an administrator.
exit /b 1
