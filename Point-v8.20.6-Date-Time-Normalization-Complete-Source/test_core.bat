@echo off
setlocal
if not exist build-tests mkdir build-tests
cl /nologo /std:c++20 /utf-8 /EHsc /W4 /WX /sdl ^
  /DUNICODE /D_UNICODE /Od /Zi ^
  /Isrc src\point_core.cpp tests\schema_mapping_test.cpp ^
  /Fe:build-tests\schema_mapping_test.exe ^
  /link bcrypt.lib
if errorlevel 1 exit /b 1
build-tests\schema_mapping_test.exe
if errorlevel 1 exit /b 1
echo All core schema mapping tests passed.
