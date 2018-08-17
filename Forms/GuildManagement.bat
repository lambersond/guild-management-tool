@ECHO OFF
Powershell.exe -NoProfile -Command "& {Start Powershell.exe -ArgumentList '-WindowStyle Hidden -NoProfile -NoExit -ExecutionPolicy Unrestricted -File %~dpn0.ps1'}"