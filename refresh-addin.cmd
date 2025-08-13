@echo off
echo Stopping npm...
taskkill /F /IM "node.exe" /T

echo Closing Excel instance with add-in ID: 28477825-5554-4111-9edf-b1366c316ff9...
taskkill /F /IM "EXCEL.EXE" /T

echo Clearing Office Add-in cache...
rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\"
rd /s /q "%USERPROFILE%\AppData\Local\Microsoft\Office\16.0\Wef\"

echo Clearing dist folder...
rd /s /q "dist"

echo Waiting for processes to fully terminate...
timeout /t 2 /nobreak

echo Starting npm...
start cmd /c "npm start"

echo Done! Please reopen Excel and your workbook.