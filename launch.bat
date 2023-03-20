@echo off
if exist "C:\PwdReset" goto LAUNCH

:COPY
xcopy "%CD%\*" "C:\PwdReset\*" /Y /E
xcopy "%CD%\Password Reset Tool.lnk" "C:\Users\%username%\Desktop\*"

:LAUNCH
powershell.exe -setexecutionpolicy bypass -command "Get-ChildItem C:\PwdReset\* -Recurse | Unblock-File"
start powershell.exe -windowstyle hidden -File "C:\PwdReset\MainForm.ps1" -SetExecutionPolicy Bypass -Noninteractive -nologo