@echo off
pushd "%~dp0"
powershell.exe -NoProfile -File .\Export-XrayTestausfuehrungen.ps1 %*
if errorlevel 1 pause
popd
