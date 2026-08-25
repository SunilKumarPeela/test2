# Creating the Point installer

Only the release owner needs the build tools. Install Visual Studio Build Tools
2022 with **Desktop development with C++**, then install Inno Setup 6. Double-
click `build_installer.bat`. The distributable file will appear as:

`build-installer\Point-v8.20.6-Setup.exe`

Send only that Setup file to end users. They double-click it and follow the
wizard; they do not need a compiler, source code, or Command Prompt.

For a trusted publisher name instead of Windows' “Unknown publisher” warning,
obtain an Authenticode code-signing certificate and define
`POINT_SIGNING_CERT_SHA1` with its certificate thumbprint before running the
builder. The builder signs both Point.exe and the final installer when the
certificate and Windows SDK SignTool are available.
