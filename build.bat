@echo off
rem builds the marking engine, its tests and the Office add-in with the C# 3.0 compiler of the .NET Framework 3.5,
rem against the .NET Framework 2.0 assemblies only, so they run on Windows 2000 with Office 2000 like SystemBanner
rem usage: build.bat [output folder], build by default; the old compiler fails on paths longer than about 200 characters
setlocal
set CSC=%WINDIR%\Microsoft.NET\Framework\v3.5\csc.exe
set NET2=%WINDIR%\Microsoft.NET\Framework\v2.0.50727
if not exist "%CSC%" (
    echo The .NET Framework 3.5 compiler was not found: %CSC%
    exit /b 1
)
cd /d "%~dp0"
set OUT=build
if not "%~1"=="" set OUT=%~1
if not exist "%OUT%" mkdir "%OUT%"
set REFS=/nostdlib+ /r:"%NET2%\mscorlib.dll" /r:"%NET2%\System.dll"

"%CSC%" /nologo /target:library /out:"%OUT%\SystemBannerOffice.Markings.dll" %REFS% Code\MarkingEngine\*.cs
if errorlevel 1 exit /b 1

"%CSC%" /nologo /target:exe /out:"%OUT%\SystemBannerOffice.Tests.exe" %REFS% /r:"%OUT%\SystemBannerOffice.Markings.dll" Code\Tests\*.cs
if errorlevel 1 exit /b 1

if exist Code\OfficeAddin\*.cs (
    "%CSC%" /nologo /target:library /out:"%OUT%\SystemBannerOffice.dll" %REFS% /r:"%NET2%\System.Windows.Forms.dll" /r:"%NET2%\System.Drawing.dll" /r:"%OUT%\SystemBannerOffice.Markings.dll" Code\OfficeAddin\*.cs
    if errorlevel 1 exit /b 1
    rem drives the installed Word, Excel and PowerPoint with the add-in's document code, without installing it
    "%CSC%" /nologo /target:exe /out:"%OUT%\OfficeCheck.exe" %REFS% /r:"%OUT%\SystemBannerOffice.Markings.dll" /r:"%OUT%\SystemBannerOffice.dll" Code\OfficeCheck\*.cs
    if errorlevel 1 exit /b 1
)
copy /y "%OUT%\SystemBannerOffice.dll" Install\ >nul
copy /y "%OUT%\SystemBannerOffice.Markings.dll" Install\ >nul
echo Build succeeded.
