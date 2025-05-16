@echo off

REM 1. 複製 Bookmarks 和 Bookmarks.bak
xcopy /Y /Q "%~dp0Bookmarks\Bookmarks" "%LOCALAPPDATA%\Google\Chrome\User Data\Default\"
xcopy /Y /Q "%~dp0Bookmarks\Bookmarks.bak" "%LOCALAPPDATA%\Google\Chrome\User Data\Default\"

REM 2. 複製 Signatures 資料夾內所有檔案
xcopy /Y /Q "%~dp0Signatures\*" "%AppData%\Microsoft\Signatures\"

REM 3. 複製 quick_access 資料夾內所有檔案
xcopy /Y /Q "%~dp0quick_access\*" "%AppData%\Microsoft\Windows\Recent\AutomaticDestinations\"

REM 4. 複製 Network Shortcuts 資料夾內所有檔案
xcopy /Y /Q "%~dp0network_shortcuts\*" "%AppData%\Microsoft\Windows\Network Shortcuts\"

REM 等待5秒後自動關閉
timeout /t 5 >nul
exit