# GET des données venant du fichier Excel
$path = "A:\OneDrive\Innovative Digital Technologies\Support BE - General\Registres\SuiviDesTachesSAV.xlsm"
$worksheetname = "Data"
$datas = Import-Excel -Path $path  -WorksheetName $worksheetname

# Variables
$today = [System.DateTime]::Now.AddDays(0).ToString('MM-dd')
$dataSend = "";

foreach ($data in $datas){

    $dataSend += '<tr><td>' + $data.Intervenant + '</td><td>' + $data.Description + '</td><td>' + $data.Notes + '</td><td>' +  $data.NomClient + '</td><td>' + $data.PlanningDebut + '</td></tr>'
}


# Autres tâches
$memos = "";
$worksheetname2 = "Mémos"
$autres = Import-Excel -Path $path  -WorksheetName $worksheetname2

foreach ($data in $autres){

    $memos += '<tr><td>' + $data.NomClient + '</td><td>' + $data.Description + '</td><td>' +  $data.Notes + '</td><td>' + $data.DateLimite + '</td></tr>'
}



if($memos -ne ''){

    $SecurePassword = ConvertTo-SecureString '1597e612e984c0453d60502e7425d755' -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential ('986328c41d9044ce7fe166c76e0bdb08', $SecurePassword)
    $SmtpServer = 'in-v3.mailjet.com'  
    $encodingMail = [System.Text.Encoding]::UTF8
    $To = 'gheitaa@idt.pf'
    $From = 'support-be@idt.pf'
    $Cc = 'hello@idt.pf'
    $Subject = "Rappels de tâches du " + $today
    $Body = '<style>table{border-collapse: collapse; width:80%; margin:auto;}th, td{ border: 1px solid black; padding: 10px;} th{background-color : #557CBA; color : #F5EFF4;}</style><table><tr><th>Client</th><th>Description</th><th>Détail</th><th>Date limite</th></tr>'+ $memos + '</table>'
    Send-MailMessage -To $To -From $From -Cc $Cc -SmtpServer $SmtpServer -Credential $Credential -Port "587" -UseSsl -Subject $Subject -BodyAsHtml $Body -Encoding $encodingMail

}

if($dataSend -ne ''){

    $SecurePassword = ConvertTo-SecureString '1597e612e984c0453d60502e7425d755' -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential ('986328c41d9044ce7fe166c76e0bdb08', $SecurePassword)
    $SmtpServer = 'in-v3.mailjet.com'  
    $encodingMail = [System.Text.Encoding]::UTF8
    $To = 'support-be@idt.pf'
    $From = 'support-be@idt.pf'
    $Cc = 'support-be@idt.pf'
    $Subject = "Rappel de tâches"
    $Body = '<style>table{border-collapse: collapse; width:80%; margin:auto;}th, td{ border: 1px solid black; padding: 10px;} th{background-color : #557CBA; color : #F5EFF4;}</style><table><tr><th>Intervenant</th><th>Description</th><th>Détail</th><th>Client</th><th>Date</th></tr>'+ $dataSend + '</table>'
    Send-MailMessage -To $To -From $From -Cc $Cc -SmtpServer $SmtpServer -Credential $Credential -Port "587" -UseSsl -Subject $Subject -BodyAsHtml $Body -Encoding $encodingMail

}