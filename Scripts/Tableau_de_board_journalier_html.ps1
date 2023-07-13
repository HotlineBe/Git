# Variables générales
$today = [System.DateTime]::Now.AddDays(0).ToString('MM-dd')
$todayJJMMYYY = [System.DateTime]::Now.AddDays(0).ToString('dd/MM/yyyy')
$hierJJMMYYY = [System.DateTime]::Now.AddDays(-1).ToString('dd/MM/yyyy')

# GET des données venant du fichier Excel
$pathTachesIdSav = "A:\OneDrive\Innovative Digital Technologies\Support BE - General\Registres\TB.xlsm"
$worksheetnameTachesIdSav = "Memo"
$datasTachesIdSav = Import-Excel -Path $pathTachesIdSav  -WorksheetName $worksheetnameTachesIdSav

$pathBackups = "A:\OneDrive\Innovative Digital Technologies\Support BE - General\Registres\Registre de backup.xlsx"
$worksheetnameBackups = "Form1"
$datasBackups = Import-Excel -Path $pathBackups  -WorksheetName $worksheetnameBackups

$pathConnexionTse = "A:\OneDrive\Innovative Digital Technologies\Support BE - General\Registres\Registre de connexion.xlsx"
$worksheetnameConnexionTse = "Form1"
$datasConnexionTse = Import-Excel -Path $pathConnexionTse  -WorksheetName $worksheetnameConnexionTse

$pathLicenceLexi = "A:\OneDrive\Innovative Digital Technologies\Support BE - General\Registres\Registre des licences Lexi.xlsx"
$worksheetnameLicenceLexi = "Form1"
$datasLicenceLexi = Import-Excel -Path $pathLicenceLexi  -WorksheetName $worksheetnameLicenceLexi

$pathMAJ = "A:\OneDrive\Innovative Digital Technologies\Support BE - General\Registres\Registre de livraison.xlsx"
$worksheetnameMAJ = "Form1"
$datasMAJ = Import-Excel -Path $pathMAJ  -WorksheetName $worksheetnameMAJ

$pathTacheHier = "A:\OneDrive\Innovative Digital Technologies\Support BE - General\Registres\TB.xlsm"
$worksheetnameTacheHier = "Tâches N-1"
$datasTacheHier = Import-Excel -Path $pathTacheHier  -WorksheetName $worksheetnameTacheHier

$pathTacheToday = "A:\OneDrive\Innovative Digital Technologies\Support BE - General\Registres\TB.xlsm"
$worksheetnameTacheToday = "Tâches N"
$datasTacheToday = Import-Excel -Path $pathTacheToday  -WorksheetName $worksheetnameTacheToday

# Résultat de traitement
$resultat = "";
$tachesIdSav = "";
$backups = "";
$connexionTse = "";
$licenceLexi = "";
$MAJ = "";
$gilles = "";
$leilanie = "";
$stephanie = "";

$compteurTaches = 0;
$compteurBackups = 0;
$compteurConnexionTse = 0;
$compteurLicenceLexi = 0;
$compteurMAJ = 0;
$compteurGilles = 0;
$compteurLeilanie = 0;
$compteurStephanie = 0;
$compteurHier = 0;

# Entête
$css = '<style>h1{text-align: center; padding: 2%; color: #557CBA;} #recap{background: #2ca747;} h2{text-decoration: underline; color: #572B50;} h3{font-style: italic; color: #B271A8;} h4{margin-left: 20px;} table {border-collapse: collapse; width:80%; margin:auto;}th, td{ border: 1px solid black; padding: 10px;} th{background-color : #557CBA; color : #F5EFF4;}</style>'
$enteteTacheIdsav = '<h3>B) Liste de tâches à ne pas oublier :</h3><table><tr><th>Intervenant</th><th>Description</th><th>Détail</th><th>Client</th><th>Date</th></tr>'
$enteteBackups = "<h3>C) Liste des backups à supprimer :</h3><table><tr><th>Client</th><th>Date du backup</th><th>Durée de conservation</th><th>Emplacement</th><th>Nom de la base de données</th></tr>";
$enteteConnexionTse = "<h3>D) Les connexions TSE :</h3><table><tr><th>Intervenant</th><th>Client</th><th>Description</th><th>Type de connexion</th></tr>";
$enteteLicenceLexi = "<h3>E) Liste des licences Lexi arrivant à expiration :</h3><table><tr><th>Client</th><th>Module</th><th>Date d'expiration</th><th>Jour de grace</th></tr>";
$enteteMAJ = "<h3>F) Les mises à jour client :</h3><table><tr><th>Opérateur</th><th>Client</th><th>Application</th><th>Type</th><th>Description</th></tr>";
$enteteanniversaire = "";


## Tâches IDSAV
foreach ($data in $datasTachesIdSav){

    $tachesIdSav += '<tr><td>' + $data.Intervenant + '</td><td>' + $data.Description + '</td><td>' + $data.Notes + '</td><td>' +  $data.NomClient + '</td><td>' + $data.PlanningDebut + '</td></tr>'
    $compteurTaches = $compteurTaches + 1
}

### Tâches faite la veille
$gilles += '<tr><td colspan="4" style="text-align: center; font-weight:bold;background:#C9B9F3;">Tâches du '+ $hierJJMMYYY + '</td></tr>'
$leilanie += '<tr><td colspan="4" style="text-align: center; font-weight:bold;background:#C9B9F3;">Tâches du '+ $hierJJMMYYY + '</td></tr>'
$stephanie += '<tr><td colspan="4" style="text-align: center; font-weight:bold;background:#C9B9F3;">Tâches du '+ $hierJJMMYYY + '</td></tr>'

foreach ($data in $datasTacheHier){
    
    if($data.Intervenant -eq "Gilles"){
        $gilles += '<tr><td>' + $data.NomClient + '</td><td>' + $data.Description + '</td><td>' + $data.Notes + '</td><td>' + $data.PlanningDebut + '</td></tr>'
        $compteurGilles = $compteurGilles + 1
        $compteurHier = $compteurHier + 1
    }
    if($data.Intervenant -eq "Leilanie"){
        $leilanie += '<tr><td>' + $data.NomClient + '</td><td>' + $data.Description + '</td><td>' + $data.Notes + '</td><td>' + $data.PlanningDebut + '</td></tr>'
        $compteurLeilanie = $compteurLeilanie + 1
        $compteurHier = $compteurHier + 1
    }
    if($data.Intervenant -eq "Stéphanie"){
        $stephanie += '<tr><td>' + $data.NomClient + '</td><td>' + $data.Description + '</td><td>' + $data.Notes + '</td><td>' + $data.PlanningDebut + '</td></tr>'
        $compteurStephanie = $compteurStephanie + 1
        $compteurHier = $compteurHier + 1
    }
}

### Tâches faite aujourd'hui
$gilles += '<tr><td colspan="4" style="text-align: center; font-weight:bold;background:#C9B9F3;">Tâches du '+ $todayJJMMYYY + '</td></tr>'
$leilanie += '<tr><td colspan="4" style="text-align: center; font-weight:bold;background:#C9B9F3;">Tâches du '+ $todayJJMMYYY + '</td></tr>'
$stephanie += '<tr><td colspan="4" style="text-align: center; font-weight:bold;background:#C9B9F3;">Tâches du '+ $todayJJMMYYY + '</td></tr>'

foreach ($data in $datasTacheToday){
    
    if($data.Intervenant -eq "Gilles"){
        $gilles += '<tr><td>' + $data.NomClient + '</td><td>' + $data.Description + '</td><td>' + $data.Notes + '</td><td>' + $data.PlanningDebut + '</td></tr>'
        $compteurGilles = $compteurGilles + 1
        $compteurHier = $compteurHier + 1
    }
    if($data.Intervenant -eq "Leilanie"){
        $leilanie += '<tr><td>' + $data.NomClient + '</td><td>' + $data.Description + '</td><td>' + $data.Notes + '</td><td>' + $data.PlanningDebut + '</td></tr>'
        $compteurLeilanie = $compteurLeilanie + 1
        $compteurHier = $compteurHier + 1
    }
    if($data.Intervenant -eq "Stéphanie"){
        $stephanie += '<tr><td>' + $data.NomClient + '</td><td>' + $data.Description + '</td><td>' + $data.Notes + '</td><td>' + $data.PlanningDebut + '</td></tr>'
        $compteurStephanie = $compteurStephanie + 1
        $compteurHier = $compteurHier + 1
    }
}


## Backups
for($i = 0 ; $i -lt $datasBackups.Length; $i++){
    $dateLimite = $datasBackups[$i].Date + $datasBackups[$i]."Durée accordée"
    $dateLimite = $dateLimite.AddDays(0).ToString('dd/MM/yyyy')

    if($dateLimite -eq $todayJJMMYYY){
        $backups += '<tr><td>' + $datasBackups[$i].Client + '</td><td>' + $datasBackups[$i].Date + '</td><td>' + $datasBackups[$i]."Durée souhaitée" + '</td><td>' + $datasBackups[$i]."Emplacement de stockage" + '</td><td>' + $datasBackups[$i]."Nom de la base de données" +'</td></tr>'
        $compteurBackups = $compteurBackups + 1
    }

}


## Connexion TSE
$connexionTse += '<tr><td colspan="4" style="text-align: center; font-weight:bold;background:#C9B9F3;">Connexion du '+ $todayJJMMYYY + '</td></tr>'
for($i = 0 ; $i -lt $datasConnexionTse.Length; $i++){
    if($datasConnexionTse[$i].Date.AddDays(0).ToString('dd/MM/yyyy') -eq $todayJJMMYYY){
        $connexionTse += '<tr><td>' + $datasConnexionTse[$i].Name + '</td><td>' + $datasConnexionTse[$i].Client + '</td><td>' + $datasConnexionTse[$i]."Description de l'intervention" + '</td><td>' + $datasConnexionTse[$i]."Serveur ou type de connexion (Anydesk/Teamviewer)" + '</td></tr>'
        $compteurConnexionTse = $compteurConnexionTse + 1
    }

}

$connexionTse += '<tr><td colspan="4" style="text-align: center; font-weight:bold;background:#C9B9F3;">Connexion du '+ $hierJJMMYYY + '</td></tr>'
for($i = 0 ; $i -lt $datasConnexionTse.Length; $i++){
    if($datasConnexionTse[$i].Date.AddDays(0).ToString('dd/MM/yyyy') -eq $hierJJMMYYY){
        $connexionTse += '<tr><td>' + $datasConnexionTse[$i].Name + '</td><td>' + $datasConnexionTse[$i].Client + '</td><td>' + $datasConnexionTse[$i]."Description de l'intervention" + '</td><td>' + $datasConnexionTse[$i]."Serveur ou type de connexion (Anydesk/Teamviewer)" + '</td></tr>'
        $compteurConnexionTse = $compteurConnexionTse + 1
    }

}



## Licence Lexi
$firstDayExcelDate = [DateTime]"1900-01-01"
$dateRefecenceExcel = $firstDayExcelDate.AddDays(-2)
$todayLicenceLexi = [System.DateTime]::Now.AddDays(0)
$delaitPopAlerte = 35

foreach ($data in $datasLicenceLexi){
    
    # Toutes les dates > à celle d'aujourd'hui
    if(($dateRefecenceExcel.AddDays($data.Date_expiration)) -gt $todayLicenceLexi){
        
        $datePopAlert = $dateRefecenceExcel.AddDays($data.Date_expiration - $delaitPopAlerte)

        

        if( (($datePopAlert - $todayLicenceLexi).Days) -le $delaitPopAlerte){

            if($data.etat -ne 'Supprimé'){
                
                #if($data.Maling_off -ne "Oui"){

                    $dateExpiration = $dateRefecenceExcel.AddDays($data.Date_expiration).ToString('dd/MM/yyyy')
                    $licenceLexi += '<tr><td>' + $data.Client + '</td><td>' + $data.Module + '</td><td>' + $dateExpiration + '</td><td>' + $data.Delais + '</td></tr>'
                    $compteurLicenceLexi = $compteurLicenceLexi + 1
                #}
            }
        }

    }

}


## Livraison

for($i = 0 ; $i -lt $datasMAJ.Length; $i++){

    if($datasMAJ[$i].Date.AddDays(0).ToString('dd/MM/yyyy') -eq $hierJJMMYYY){
        $MAJ += '<tr><td>' + $datasMAJ[$i].Name + '</td><td>' + $datasMAJ[$i].Client + '</td><td>' + $datasMAJ[$i].Application + '</td><td>' + $datasMAJ[$i].Type  + '</td><td>' + $datasMAJ[$i].Description + '</td></tr>'
        $compteurMAJ = $compteurConnexionTse + 1
    }

}


### Début
$resultat += $css
$resultat += "<h1>Rapport d'activité du " + $todayJJMMYYY + "</h1>"

### RECAP
$resultat += "<h2>I - Récapitulatif : </h2><table><tr><th id='recap'>Tâches faite hier</th><th id='recap'>Tâches en mémo</th><th id='recap'>Backup</th><th id='recap'>Nombre de connexion TSE</th><th id='recap'>Licence Lexi</th><th id='recap'>Mise à jour client</th></tr><td>" + $compteurHier +"</td><td>" + $compteurTaches + "</td><td>" + $compteurBackups + "</td><td>" + $compteurConnexionTse + "</td><td>" + $compteurLicenceLexi + "</td><td>" + $compteurMAJ + "</td></tr></table>"

#### Alimenter la variable resultat

$resultat += "<h2>II - Détails</h2>"

$resultat += "<h3>A) Récapitulatif des tâches faites du " + $hierJJMMYYY + " au " + $todayJJMMYYY + "(" + $compteurHier + " tâches)</h3>"


$enteteGilles = "<h4>Gilles (" + $compteurGilles +" tâches)</h4><table><tr><th>Client</th><th>Description</th><th>Détail</th><th>Date</th></tr>";
$enteteLeilanie = "<h4>Leilnaie (" + $compteurLeilanie +" tâches)</h4><table><tr><th>Client</th><th>Description</th><th>Détail</th><th>Date</th></tr>";
$enteteStephanie = "<h4>Stéphanie (" + $compteurStephanie +" tâches)</h4><table><tr><th>Client</th><th>Description</th><th>Détail</th><th>Date</th></tr>";

$resultat += $enteteGilles;
$resultat += $gilles;
$resultat += "</table>";

$resultat += $enteteStephanie;
$resultat += $stephanie;
$resultat += "</table>";

$resultat += $enteteLeilanie;
$resultat += $leilanie;
$resultat += "</table>";


if($tachesIdSav -ne ''){
    
    $resultat += $enteteTacheIdsav;
    $resultat += $tachesIdSav;
    $resultat += "</table>";
    
}


if($tachesIdSav -ne ''){
    
    $resultat += $enteteBackups;
    $resultat += $backups;
    $resultat += "</table>";
    
}

if($tachesIdSav -ne ''){
    
    $resultat += $enteteConnexionTse;
    $resultat += $connexionTse;
    $resultat += "</table>";
    
}



if($tachesIdSav -ne ''){
    
    $resultat += $enteteLicenceLexi;
    $resultat += $licenceLexi;
    $resultat += "</table>";
    
}


if($tachesIdSav -ne ''){
    
    $resultat += $enteteMAJ;
    $resultat += $MAJ;
    $resultat += "</table>";
    
}

Clear-Content -path "A:\OneDrive\Innovative Digital Technologies\Support BE - General\Registres\Tableau de bord\Tableau de bord journalier.html"
ADD-content -path "A:\OneDrive\Innovative Digital Technologies\Support BE - General\Registres\Tableau de bord\Tableau de bord journalier.html" -value $resultat