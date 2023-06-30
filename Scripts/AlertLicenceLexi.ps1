$path = "A:\OneDrive\Innovative Digital Technologies\Support BE - General\Registres\Registre des licences Lexi.xlsx"
$worksheetname = "Form1"
$datas = Import-Excel -Path $path  -WorksheetName $worksheetname
$today = [System.DateTime]::Now.AddDays(0)
$todayMMddyyy = [System.DateTime]::Now.AddDays(0).ToString('MM-dd-yyyy')

$delaitPopAlerte = 35
$elements = @{}
$infos = ""


# Traitement des dates excel
$firstDayExcelDate = [DateTime]"1900-01-01"
$dateRefecenceExcel = $firstDayExcelDate.AddDays(-2)

foreach ($data in $datas){
    
    # Toutes les dates > à celle d'aujourd'hui
    if(($dateRefecenceExcel.AddDays($data.Date_expiration)) -gt $today){
        
        $datePopAlert = $dateRefecenceExcel.AddDays($data.Date_expiration - $delaitPopAlerte)

        

        if( (($datePopAlert - $today).Days) -le $delaitPopAlerte){

            if($data.etat -ne 'Supprimé'){
                
                #if($data.Maling_off -ne "Oui"){

                    $dateExpiration = $dateRefecenceExcel.AddDays($data.Date_expiration).ToString('dd/MM/yyyy')
                    $elements.Add($data.Client,$dateExpiration)
                    $infos += '  <li> La licence <strong>' + $data.Module + '</strong> du client <strong>' + $data.Client + '</strong> expire le <strong>' + $dateExpiration + '</strong>. Il a <strong>' + $data.Delais + ' jours</strong> de grace</li>'

                #}
            }
        }

    }

}


Write-Output $infos

if($elements.Count -gt 0) {

    $SecurePassword = ConvertTo-SecureString '1597e612e984c0453d60502e7425d755' -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential ('986328c41d9044ce7fe166c76e0bdb08', $SecurePassword)
    $SmtpServer = 'in-v3.mailjet.com'  
    $encodingMail = [System.Text.Encoding]::UTF8
    $To = 'support-be@idt.pf'
    $From = 'support-be@idt.pf'
    $Cc = 'dmelzani@idt.pf','sreverdy@idt.pf' 
    $Subject = "[Registre Licence Lexi] Liste des licences à revenouveller"
    #$Body = "Clients : <br> " + $elements.Keys + "<br><br>Date d'expiration : <br>" + $elements.Values
    $Body = 'Bonjour, <br><ul>' + $infos + '</ul><br><i>Equipe support</i>'
    Send-MailMessage -To $To -From $From -Cc $Cc -SmtpServer $SmtpServer -Credential $Credential -Port "587" -UseSsl -Subject $Subject -BodyAsHtml $Body -Encoding $encodingMail

}