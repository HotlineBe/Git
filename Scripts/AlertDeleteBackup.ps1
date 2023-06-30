
$data = Import-Excel -Path 'A:\OneDrive\Innovative Digital Technologies\Support BE - General\Registres\Registre de backup.xlsx'  -WorksheetName 'Form1'
$deathLine = @{}
$today = [System.DateTime]::Now.AddDays(0).ToString('dd/MM/yyyy')
$infos = ""


for($i = 0 ; $i -lt $data.Length; $i++){
    $dateLimite = $data[$i].Date + $data[$i]."Durée accordée"
    $dateLimite = $dateLimite.AddDays(0).ToString('dd/MM/yyyy')

    if($dateLimite -eq $today){
        $deathLine.add($data[$i].Client,$dateLimite)
        $infos += '<li>Le backup de <strong>' + $data[$i].Client + '</strong> récupere le <strong>' + $data[$i].Date + '</strong> pour une durée de <strong>' + $data[$i]."Durée souhaitée" + ' jours</strong> doit être supprimé. Il est stocké sur <strong>' + $data[$i]."Emplacement de stockage" + '</strong> et se nomme <strong>' + $data[$i]."Nom de la base de données" +'</strong></li>'
        
    }

}

Write-Output $infos

if($deathLine.Count -gt 0) {

    $SecurePassword = ConvertTo-SecureString '1597e612e984c0453d60502e7425d755' -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential ('986328c41d9044ce7fe166c76e0bdb08', $SecurePassword)
    $SmtpServer = 'in-v3.mailjet.com'  
    $encodingMail = [System.Text.Encoding]::UTF8
    $To = 'support-be@idt.pf'
    $From = 'support-be@idt.pf'
    $Cc = 'dmelzani@idt.pf','sreverdy@idt.pf' 
    $Subject = "[Registre backup] Liste des backups à supprimer"
    #$Body = "Clients : <br> " + $deathLine.Keys + "<br><br>Date de suppression : <br>" + $deathLine.Values
    $Body = "Bonjour,<br> <ul>" + $infos + '</ul><br><i>Equipe support</i>'
    Send-MailMessage -To $To -From $From -Cc $Cc -SmtpServer $SmtpServer -Credential $Credential -Port "587" -UseSsl -Subject $Subject -BodyAsHtml $Body -Encoding $encodingMail

}