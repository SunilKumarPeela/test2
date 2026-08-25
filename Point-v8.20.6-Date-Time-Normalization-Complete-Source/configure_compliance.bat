@echo off
setlocal
net session >nul 2>&1
if errorlevel 1 (
  echo Run this file once from an Administrator Command Prompt.
  exit /b 1
)
net localgroup "Point Users" >nul 2>&1
if errorlevel 1 net localgroup "Point Users" /add
if errorlevel 1 exit /b 1
net localgroup "Point Administrators" >nul 2>&1
if errorlevel 1 net localgroup "Point Administrators" /add
if errorlevel 1 exit /b 1
net localgroup "Point Exporters" >nul 2>&1
if errorlevel 1 net localgroup "Point Exporters" /add
if errorlevel 1 exit /b 1
net localgroup "Point Users" "%USERDOMAIN%\%USERNAME%" /add
if errorlevel 1 exit /b 1
net localgroup "Point Exporters" "%USERDOMAIN%\%USERNAME%" /add
if errorlevel 1 exit /b 1
echo Point Windows groups are configured.
echo Sign out and sign in again before starting Point.
