
Install-Module -Name ImportExcel -Scope CurrentUser -Force


# Erstellen Sie einen relativen Pfad, falls das Skript von einem anderen Ort aus ausgeführt wird ---- 
$dataPath = Join-Path $PSScriptRoot "Modul3.xlsx"
$errorDataPath = Join-Path $PSScriptRoot "ErrorLog.csv"

try {
    $data = Import-Excel -Path $dataPath -WorksheetName "testdaten_Modul3"
}
catch {
    Write-Host "Die Datei $dataPath existiert nicht oder kann nicht gelesen werden. Das Programm wird unterbrochen." -ForegroundColor Red
    Write-Host "Stell bitte sicher, dass der Name der Datei mit den Eingabedaten Users.csv heißt. Falls nicht, bitte benenne sie um!" -ForegroundColor Red
    $lastError = "Fehler beim Vorgang. Die Datei $dataPath existiert nicht oder kann nicht gelesen werden."
    $errorOutput = New-Object -TypeName PSObject -Property @{Date = (Get-Date); LastError = $lastError}
    $errorOutput | Export-Csv -Path $errorDataPath -NoTypeInformation -Append -Force -Encoding "UTF8"
    exit
}

# Ordner e:\user_shares\ auf dem Server anlegen------------------------------------------------------------------
$ordnerPath = "E:\user_shares"

if (Test-Path $ordnerPath) {
    Write-Host "Folderr $ordnerPath exist on E."
} else {
    Write-Host "Folderr $ordnerPath NOT exist on E and will be created."
    New-Item -ItemType Directory -Path $ordnerPath
}

#disable OU ProtectedFromAccidentalDeletion
$protection = $false


$standort = "Stuttgart"
$standortOUPath = "OU=$standort"
$domainDCPath = "DC=s,DC=zukunftsmotor,DC=org"
$domainOUPath = "$standortOUPath,$domainDCPath"

# Ordner e:\user_shares\ auf dem Server anlegen------------------------------------------------------------------
if (Get-ADOrganizationalUnit -Filter {Name -eq $standort}) {
    Write-Host "OU '$($standort)' schon exist."
}
else {
    New-ADOrganizationalUnit -Name $standort -Path $domainDCPath -ProtectedFromAccidentalDeletion $protection
    New-ADGroup -Name $standort -SamAccountName $standort -GroupCategory Security -GroupScope Global -DisplayName $standort -Path $domainOUPath -Description "Members of $standort"
}
#----------------------------------------------------------------------------------------------------------------

$good = 0
$doppelter_eintrag = 0

foreach ($line in $data) {
    if ($line.Standort -eq $standort) {
        $vorname = $line.Vorname
        $nachname = $line.Name
        $abteilung = $line.Abteilung
        $raum = [int]($line.Raum)

        # Für jedes Benutzerkonto einen Unterordner von e:\user_shares anlegen------------------------------------- 
        $userFileName = $nachname+"."+$vorname
        $userOrdnerPath = "$ordnerPath\$userFileName"
        if (Test-Path $userOrdnerPath) {
            Write-Host "Folderr $userOrdnerPath  exist on E."
        } else {
            Write-Host "Folderr $userOrdnerPath NOT exist on E and will be created."
            New-Item -ItemType Directory -Path $userOrdnerPath
        }
        #----------------------------------------------------------------------------------------------------------


        # OU Hiearchy fur Spalten „Abteilung“ und „Raum“ anlegen---------------------------------------------------
        $abteilungOUPath = "OU=$abteilung,$domainOUPath"

        if (Get-ADOrganizationalUnit -Filter {Name -eq $abteilung}) {
            Write-Host "OU '$($line.Abteilung)' schon exist."

            if (Get-ADOrganizationalUnit -Filter {Name -eq $raum}) {
                Write-Host "OU '$($line.Raum)' schon exist."
            }

            else {
                New-ADOrganizationalUnit -Name $raum -Path $abteilungOUPath -ProtectedFromAccidentalDeletion $protection
                Write-Host "OU '$($line.Raum)' wurde erstellt."
            }
        }

        else {
            New-ADOrganizationalUnit -Name $abteilung -Path $domainOUPath -ProtectedFromAccidentalDeletion $protection
            New-ADGroup -Name $abteilung -SamAccountName $abteilung -GroupCategory Security -GroupScope Global -DisplayName $abteilung -Path $abteilungOUPath -Description "Members of $abteilung"
            Write-Host "OU '$($line.Abteilung)' wurde erstellt."

            if (Get-ADOrganizationalUnit -Filter {Name -eq $raum}) {
                Write-Host "OU '$($line.Raum)' schon exist."
            }

            else {
                New-ADOrganizationalUnit -Name $raum -Path $abteilungOUPath -ProtectedFromAccidentalDeletion $protection
                Write-Host "OU '$($line.Raum)' wurde erstellt."
            }
          
        }
        #------------------------------------------------------------------------------------------------------------------

        #New Benutzer anlegen-----------------------------------------------------------------------------------------------
        $mitabeiterNummer = $line.Mitarbeiternummer
        $jobBezeichnung = $line.Jobbezeichnung
        $email = $vorname+"."+$nachname+"@zukunftsmotor.org"
        $name = $vorname+" "+$nachname
        $account_name = $vorname+"."+$nachname
        $account_name_fd = $account_name.Normalize([System.Text.NormalizationForm]::FormD)
        $sam_account_name = $account_name_fd -replace '\p{M}', '' 
        $company = "Zuukunftsmotor GmbH"
        $office = "$standort Abteilung:$abteilung Raum:$raum"
        #Write-Host $userOrdnerPath

        if ($sam_account_name.Length -gt 20) {$sam_account_name = $sam_account_name.Substring(0,20)} 
        
            
            
            
        try {
            
            # New Benutzer anlegen -----------------------------------------
            New-ADUser -Name $name `
              -SamAccountName $sam_account_name `
              -UserPrincipalName $sam_account_name `
              -EmployeeNumber $mitabeiterNummer `
              -GivenName $vorname `
              -Surname $nachname `
              -AccountPassword (ConvertTo-SecureString "Password*123" -AsPlainText -Force) `
              -Enabled $true `
              -EmailAddress $email `
              -ProfilePath $userOrdnerPath `
              -Company $company `
              -Title $jobBezeichnung `
              -Department $abteilung `
              -City $standort `
              -Office $office
            

            # Benutzer in OU verschieben ------------------------------------
            $identityOU = "OU=$raum,OU=$abteilung,OU=$standort,$domainDCPath"
            $identityUser = "CN=$name,CN=Users,DC=s,DC=zukunftsmotor,DC=org"
       
            Move-ADObject -Identity $identityUser -TargetPath $identityOU
            #-----------------------------------------------------------------


            # Benutzer in Gruppe verschieben ------------------------------------
            
            Add-ADGroupMember -Identity $standort -Members $sam_account_name       
            Add-ADGroupMember -Identity $abteilung -Members $sam_account_name

            #-----------------------------------------------------------------

            $good++
        
        }
        catch {

            # falls Benuzer schon exist Error message in ErrorLog.csv speichern ----------------------------------------------------------------------------------
            $doppelter_eintrag++
            $lastError = "Der Index lag außerhalb des Bereichs. Der Benutzer existiert schon."
            $errorOutput = New-Object -TypeName PSObject -Property @{Date = (Get-Date); LastError = $lastError}
            $errorOutput | Export-Csv -Path $errorDataPath -NoTypeInformation -Append -Force -Encoding "UTF8"
            Write-Host "Der Benutzer existiert schon" -ForegroundColor Yellow
            Write-Host $line -ForegroundColor Yellow
            #------------------------------------------------------------------------------------------------------------------------------------------------------
        }

        if (Test-Path $userOrdnerPath) {
            # Berechtigungen für ein Verzeichnis festlegen -------------------------------------------------------------------------------------------------
            $acl = Get-Acl $userOrdnerPath
            $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($sam_account_name, "Modify", "ContainerInherit,ObjectInherit", "None", "Allow")
            $acl.SetAccessRule($accessRule)
            $acl | Set-Acl $userOrdnerPath
            
            # Berechtigungsvererbung entfernen  ----------------------------------------------------------------------------------------------------------
            $acl.SetAccessRuleProtection($True, $False)
            $acl | Set-Acl $userOrdnerPath
       

            # Share File -------------------------------------------------------------------------------------------------------------------------------------------
                 
      
            $folderPath = $userOrdnerPath
            $shareName = "$userFileName"

            $shareParams = @{
                Name        = $shareName
                Path        = $folderPath
                FullAccess  = $sam_account_name
                Description = "Sdílená složka pro uživatele $name"
                }

            try {
                New-SmbShare @shareParams -ErrorAction Stop
            }

            catch {
                $lastError = "Fehler beim New-SmbShare. Der Name wurde bereits freigegeben oder kann nicht angelegt werden."
                $errorOutput = New-Object -TypeName PSObject -Property @{Date = (Get-Date); LastError = $lastError}
                $errorOutput | Export-Csv -Path $errorDataPath -NoTypeInformation -Append -Force -Encoding "UTF8"
                Write-Host $lastError
            }

        } 
        
        else {
            Write-Host "Path do NOT EXIST"
        }
    }
}







gpupdate /force


Write-Host "Von der Gesamtzahl der Einträge: $($data.Count) wurden: $good Benutzerkonten erstellt."
Write-Host "Es war nicht möglich, $doppelter_eintrag Benutzerkonten zu erstellen. Der Name und die Fehlermeldungen wurden in dein Path : ErrorLog.csv gespeichert." -ForegroundColor Yellow 
#Read-Host "Drücke Enter Taste, um fortzufahren."

#Get-SmbShare -Name *
