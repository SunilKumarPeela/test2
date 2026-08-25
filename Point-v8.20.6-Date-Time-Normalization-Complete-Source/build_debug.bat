@echo off
setlocal
if not exist point.ico (
  echo ERROR: point.ico is required. Point will not build without its logo.
  exit /b 1
)
if not exist build-debug mkdir build-debug
if not defined WEBVIEW2_SDK_DIR set "WEBVIEW2_SDK_DIR=packages\Microsoft.Web.WebView2"
if not exist "%WEBVIEW2_SDK_DIR%\build\native\include\WebView2.h" (
  echo WebView2 SDK is missing. Installing it with NuGet...
  where nuget >nul 2>&1
  if errorlevel 1 (
    echo ERROR: nuget.exe is required to download the WebView2 SDK.
    exit /b 1
  )
  nuget install Microsoft.Web.WebView2 -OutputDirectory packages -ExcludeVersion -NonInteractive
  if errorlevel 1 exit /b 1
)
if not exist "%WEBVIEW2_SDK_DIR%\build\native\include\WebView2.h" (
  for /d %%D in (packages\Microsoft.Web.WebView2*) do (
    if exist "%%D\build\native\include\WebView2.h" set "WEBVIEW2_SDK_DIR=%%D"
  )
)
if not exist "%WEBVIEW2_SDK_DIR%\build\native\include\WebView2.h" (
  echo ERROR: WebView2 SDK installation completed, but WebView2.h was not found.
  exit /b 1
)
rc /nologo /fo build-debug\point.res src\point.rc
if errorlevel 1 exit /b 1
cl /nologo /std:c++20 /utf-8 /EHsc /W4 /WX /sdl /guard:cf ^
  /DUNICODE /D_UNICODE /Od /Zi /RTC1 /I"%WEBVIEW2_SDK_DIR%\build\native\include" ^
  src\point_browser_fetcher.cpp build-debug\point.res /Fe:build-debug\PointBrowserFetcher-Debug.exe ^
  /link /DEBUG /DYNAMICBASE /NXCOMPAT /GUARD:CF ^
  user32.lib gdi32.lib comctl32.lib shell32.lib ole32.lib oleaut32.lib ^
  shlwapi.lib advapi32.lib version.lib runtimeobject.lib ^
  "%WEBVIEW2_SDK_DIR%\build\native\x64\WebView2LoaderStatic.lib"
if errorlevel 1 exit /b 1
cl /nologo /std:c++20 /utf-8 /EHsc /W4 /WX /sdl /guard:cf ^
  /DUNICODE /D_UNICODE /Od /Zi /RTC1 ^
  src\point_core.cpp src\point_compliance.cpp src\point_excel_import.cpp src\point_win32.cpp build-debug\point.res ^
  /Fe:build-debug\Point-Debug.exe ^
  /link /DEBUG /DYNAMICBASE /NXCOMPAT /GUARD:CF ^
  user32.lib gdi32.lib comctl32.lib comdlg32.lib shell32.lib ole32.lib oleaut32.lib ^
  advapi32.lib crypt32.lib bcrypt.lib
if errorlevel 1 exit /b 1
cl /nologo /std:c++20 /utf-8 /EHsc /W4 /WX /sdl /guard:cf ^
  /DUNICODE /D_UNICODE /Od /Zi /RTC1 ^
  src\point_fetcher.cpp build-debug\point.res /Fe:build-debug\PointFetcher-Debug.exe ^
  /link /DEBUG /DYNAMICBASE /NXCOMPAT /GUARD:CF ^
  user32.lib gdi32.lib comctl32.lib shell32.lib advapi32.lib winhttp.lib
if errorlevel 1 exit /b 1
echo Built Point, Point Fetcher, and Point Browser Fetcher debug executables with runtime checks.
