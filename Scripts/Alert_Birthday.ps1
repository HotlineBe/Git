# GET des données venant du fichier Excel
$path = "A:\OneDrive\OneDrive - Innovative Digital Technologies\IDT\Divers\Date de naissance IDT.xlsx"
$worksheetname = "Feuil1"
$datas = Import-Excel -Path $path  -WorksheetName $worksheetname

# Variables
$today = [System.DateTime]::Now.AddDays(0).ToString('MM-dd')
$dataSend = "";

foreach ($data in $datas){

    if($today -eq $data.Date.ToString('MM-dd')){
        
        $dataSend += 'C''est l''Anniversaire de ' + $data.Nom + '<br>'
    }
}

if($dataSend -ne ''){

    $SecurePassword = ConvertTo-SecureString '1597e612e984c0453d60502e7425d755' -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential ('986328c41d9044ce7fe166c76e0bdb08', $SecurePassword)
    $SmtpServer = 'in-v3.mailjet.com'  
    $encodingMail = [System.Text.Encoding]::UTF8
    $From = 'support-be@idt.pf'
    $To = 'gheitaa@idt.pf'
    $Cc = 'ltuahiva@idt.pf','ssurdacki@idt.pf'
    $Subject = "HAPPY BIRTHDAY"
    $Body = $dataSend
    Send-MailMessage -To $To -From $From -Cc $Cc -SmtpServer $SmtpServer -Credential $Credential -Port "587" -UseSsl -Subject $Subject -BodyAsHtml $Body -Encoding $encodingMail

}