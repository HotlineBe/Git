ipconfig /flushdns
taskkill /IM Lexi.Desktop.exe
taskkill /IM Lexi.Desktop.exe
timeout /t 2
start A:\App\IDSAV\Lexi.Desktop.exe
timeout /t 5
call sendkeys.bat "lexi" ""
exit