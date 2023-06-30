$myService = "LexiClientIds"

$etatService = (get-service | Where-Object {$_.name -eq $myService } |select Status | ft -HideTableHeader) | Out-String
if ($etatService.trim() -eq "Stopped")
{
    start-service -name $myService
}