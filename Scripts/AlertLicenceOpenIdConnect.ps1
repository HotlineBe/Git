$path = "A:\OneDrive\Innovative Digital Technologies\Support BE - General\Registres\Registre de certificat OpenIdConnect.xlsx"
$worksheetname = "Form1"
$datas = Import-Excel -Path $path  -WorksheetName $worksheetname
$today = [System.DateTime]::Now.AddDays(0).ToString('dd/MM/yyyy')
$today2 = [System.DateTime]::Now.AddDays(0)

$delaitPopAlerte = 35
$dateLimite = $today2.AddDays($delaitPopAlerte).ToString('dd/MM/yyyy')
$dataSend = "<table><tr><th>Client</th><th>Date expiration</th></tr>"


# Traitement des dates excel
$firstDayExcelDate = [DateTime]"1900-01-01"
$dateRefecenceExcel = $firstDayExcelDate.AddDays(-2)

foreach ($data in $datas){
    
    $dateExpiration = $dateRefecenceExcel.AddDays($data.'Date d''expiration').ToString('dd/MM/yyyy')
    
    if($data.'Date d''expiration' -lt $today2){
        Write-Host $data.Client " - " $dateExpiration
        #$dataSend += "<tr><td>" + $data.Client + "</td><td>" + $dateExpiration + "</td></tr>"
    }

}

$dataSend += "</table>"
 #Write-Output $dataSend

#if($dataSend -ne "<table><tr><th>Client</><th>Date expiration</th></tr></table>"){
#
#    Write-Output $dataSend
#
#    $SecurePassword = ConvertTo-SecureString '1597e612e984c0453d60502e7425d755' -AsPlainText -Force
#    $Credential = New-Object System.Management.Automation.PSCredential ('986328c41d9044ce7fe166c76e0bdb08', $SecurePassword)
#    $SmtpServer = 'in-v3.mailjet.com'
#    $encodingMail = [System.Text.Encoding]::UTF8
#    $To = 'gheitaa-be@idt.pf'
#    $From = 'support-be@idt.pf'
#    $Cc = 'gheitaa@idt.pf' 
#    $Subject = "[Registre certificat OpenIdConnect] Liste des certificats a revenouveller"
#    $Body = 'Bonjour, <br><br> ' + $dataSend + '<br><i>Equipe support</i>'
#    Send-MailMessage -To $To -From $From -Cc $Cc -SmtpServer $SmtpServer -Credential $Credential -Port "587" -UseSsl -Subject $Subject -BodyAsHtml $Body -Encoding $encodingMail
#
#}