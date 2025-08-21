@echo off
echo Stopping npm...
taskkill /F /IM "node.exe" /T > nul 2>&1

echo Closing Excel instance...
taskkill /F /IM "EXCEL.EXE" /T > nul 2>&1

echo Clearing Office Add-in cache...
rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\"
rd /s /q "%USERPROFILE%\AppData\Local\Microsoft\Office\16.0\Wef\"

echo Clearing dist folder...
if exist "dist" (
    rd /s /q "dist"
)

echo Waiting for processes to fully terminate...
timeout /t 2 /nobreak > nul

echo Starting the add-in...
call npm run start:dev

echo Starting Node server...
start cmd /c "node server.mjs"

echo Done! Please wait for Excel to open.