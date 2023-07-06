mkdir c:\Lexi
robocopy "\\172.16.1.14\be\publications\ID-SAV\IDT\Console" "A:\App\IDSAV"  /MIR /LOG+:C:\Lexi\updateConsoleSAV.log
A:\App\IDSAV\ngencustom.exe
A:\App\IDSAV\Lexi.Desktop.exe