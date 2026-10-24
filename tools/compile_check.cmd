@echo off
REM Compile-check only: builds to obj\scratch\ so it does NOT touch bin\Release
REM (which is locked by a running Outlook). Use this to catch compile errors
REM without asking the user to close Outlook. For the real build use build_publish.cmd.
setlocal

set "VSWHERE=C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe"
set "MSBUILD="

if exist "%VSWHERE%" (
  for /f "usebackq tokens=*" %%i in (`"%VSWHERE%" -latest -products * -requires Microsoft.VisualStudio.Component.VSTO -property installationPath`) do (
    if exist "%%i\MSBuild\Current\Bin\MSBuild.exe" set "MSBUILD=%%i\MSBuild\Current\Bin\MSBuild.exe"
  )
)

if not defined MSBUILD (
  if exist "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe" set "MSBUILD=C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe"
  if exist "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" set "MSBUILD=C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
  if exist "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe" set "MSBUILD=C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
)

echo Using MSBUILD: %MSBUILD%
cd /d "%~dp0..\OutlookDateJump"
"%MSBUILD%" OutlookDateJump.csproj /p:Configuration=Release /p:OutputPath=obj\scratch\ /t:Build /v:n /nologo
if errorlevel 1 (
  echo.
  echo COMPILE CHECK FAILED.
  exit /b 1
)
echo.
echo COMPILE CHECK OK ^(output: obj\scratch\, bin\Release untouched^)
