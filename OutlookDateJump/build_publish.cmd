@echo off
REM Build + ClickOnce publish for OutlookDateJump VSTO add-in.
REM Self-contained: locates MSBuild via vswhere (VS 2022 with VSTO workload only).
REM Can be run by double-click - no Developer Command Prompt needed.
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

if not defined MSBUILD (
  echo ERROR: MSBuild.exe of VS 2022 not found. Install VS 2022 with the
  echo "Office/SharePoint development" workload first.
  exit /b 1
)

echo Using MSBUILD: %MSBUILD%
cd /d "%~dp0"

echo === Build (Release) ===
"%MSBUILD%" OutlookDateJump.csproj /p:Configuration=Release /t:Build /v:n /nologo
if errorlevel 1 goto fail

echo === Bump ClickOnce revision (ASCII state file) so Outlook actually redeploys ===
for /f "usebackq tokens=*" %%n in (`powershell -NoProfile -Command "$f='..\tools\revision.txt'; $n=[int](Get-Content $f -Raw) + 1; Set-Content $f -Value $n -NoNewline; Write-Output $n"`) do set NEWREV=%%n
echo New revision: %NEWREV%

echo === Publish (ClickOnce) ===
"%MSBUILD%" OutlookDateJump.csproj /p:Configuration=Release /t:Publish /p:ApplicationRevision=%NEWREV% /v:m /nologo
if errorlevel 1 goto fail

echo.
echo SUCCESS. Install package: bin\Release\app.publish\
exit /b 0

:fail
echo.
echo BUILD or PUBLISH FAILED.
exit /b 1
