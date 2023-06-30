ipconfig /flushdns
taskkill /IM Lexi.Desktop.exe
taskkill /IM Lexi.Desktop.exe
timeout /t 2
start C:\Lexi\Console\Lexi.Desktop.exe
timeout /t 5
call sendkeys.bat "lexi" ""
exit