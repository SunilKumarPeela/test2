@echo off
setlocal
call build.bat
if errorlevel 1 exit /b 1

if defined POINT_SIGNING_CERT_SHA1 (
  where signtool >nul 2>&1
  if errorlevel 1 (
    echo POINT_SIGNING_CERT_SHA1 is set, but SignTool was not found.
    exit /b 1
  )
  signtool sign /sha1 "%POINT_SIGNING_CERT_SHA1%" /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 build\Point.exe
  if errorlevel 1 exit /b 1
  signtool sign /sha1 "%POINT_SIGNING_CERT_SHA1%" /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 build\PointFetcher.exe
  if errorlevel 1 exit /b 1
  signtool sign /sha1 "%POINT_SIGNING_CERT_SHA1%" /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 build\PointBrowserFetcher.exe
  if errorlevel 1 exit /b 1
)

set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" (
  echo Inno Setup 6 is required to create Point-Setup.exe.
  echo Download it from https://jrsoftware.org/isdl.php and run this file again.
  exit /b 1
)

"%ISCC%" installer\Point.iss
if errorlevel 1 exit /b 1
if defined POINT_SIGNING_CERT_SHA1 (
  signtool sign /sha1 "%POINT_SIGNING_CERT_SHA1%" /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 build-installer\Point-v8.20.6-Setup.exe
  if errorlevel 1 exit /b 1
)
echo.
echo Installer created: build-installer\Point-v8.20.6-Setup.exe
endlocal
