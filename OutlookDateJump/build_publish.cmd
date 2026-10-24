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

REM --- Semantic version (single source of truth: repo-root VERSION file) ---
set "VERFILE=%~dp0..\VERSION"
if not exist "%VERFILE%" (
  echo ERROR: VERSION file not found: %VERFILE%
  echo Create it with a SemVer string, e.g. 1.4.0 or 1.4.0-rc.1
  exit /b 1
)
REM Read via PowerShell so a UTF-8 BOM / newline does not leak into the value.
for /f "usebackq tokens=*" %%v in (`powershell -NoProfile -Command "[System.IO.File]::ReadAllText('%VERFILE%').Trim()"`) do set SEMVER=%%v
if not defined SEMVER (
  echo ERROR: could not read VERSION file: %VERFILE%
  exit /b 1
)
echo Semantic version: %SEMVER%

echo === Build (Release) ===
"%MSBUILD%" OutlookDateJump.csproj /p:Configuration=Release /t:Build /v:n /nologo
if errorlevel 1 goto fail

echo === Bump ClickOnce deployment revision (monotonic counter, NOT the version) ===
REM revision.txt only tells ClickOnce "this is a newer build, redeploy".
REM The user-visible version comes from the SemVer in the VERSION file above.
for /f "usebackq tokens=*" %%n in (`powershell -NoProfile -Command "$f='..\tools\revision.txt'; $n=[int](Get-Content $f -Raw) + 1; Set-Content $f -Value $n -NoNewline; Write-Output $n"`) do set NEWREV=%%n
echo New ClickOnce revision: %NEWREV%

echo === Publish (ClickOnce) ===
"%MSBUILD%" OutlookDateJump.csproj /p:Configuration=Release /t:Publish /p:ApplicationRevision=%NEWREV% /v:m /nologo
if errorlevel 1 goto fail

echo.
echo SUCCESS. Version %SEMVER% (ClickOnce revision %NEWREV%)
echo Install package: bin\Release\app.publish\
exit /b 0

:fail
echo.
echo BUILD or PUBLISH FAILED.
exit /b 1
