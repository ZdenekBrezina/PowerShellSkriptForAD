Import-Module ActiveDirectory

#####################################################################
##############   BEKANNTES PROBLEM: '#NAME?' in Excel-Formeln   #####
#####################################################################
# Wenn Zellen in der Excel-Datei (z.B. ConstructData!B7-B10, B12) ploetzlich
# '#NAME?' anzeigen, obwohl Vorname/Nachname korrekt sind, liegt das NICHT
# an diesem Skript, sondern daran, dass das Windows-Konto, unter dem
# PowerShell/Excel laeuft (049ad4...), keine Microsoft 365 Lizenz hat.
# Funktionen wie XLOOKUP brauchen eine gueltige, angemeldete M365-Lizenz -
# ohne sie liefert Excel im Hintergrund '#NAME?' statt des berechneten Werts.
#
# LOESUNG (einmalig, z.B. auf dem Server WVC-0090):
#   1. Als 049ad4... auf dem Server anmelden.
#   2. Excel manuell oeffnen und sich dort NUR fuer diese Anwendung als
#      049ext4... (lizenziertes Konto) anmelden (Datei > Konto > Anmelden).
#   3. Excel komplett schliessen und wieder neu oeffnen.
#   4. Man kann sich danach wieder als 049ext4... auf dem Server anmelden -
#      die Office-Anmeldung bleibt fuer 049ad4... im Hintergrund bestehen.
#   5. Dieses Skript muss weiterhin unter 049ad4... in der PS ISE laufen
#      (fuer die AD-Schreibrechte: Set-ADUser, Passwort-Reset, Gruppen).
#
# Falls '#NAME?' wieder auftaucht: Schritt 2+3 (Excel-Anmeldung als
# 049ext4...) unter dem 049ad4-Konto wiederholen.
#####################################################################

#####################################################################
##############     BENUTZER IDENTIFIZIEREN (Anfang)     #############
#####################################################################

# Der Benutzer wurde bereits (z.B. durch das HR-System) angelegt.
# Deshalb wird zuerst nach dem SamAccountName gefragt, damit sichergestellt ist,
# dass wir den richtigen Benutzer bearbeiten. Aus dem SamAccountName ergeben sich
# gleichzeitig die ersten 3 Ziffern als Prefix (z.B. "049Mustermann" -> Prefix "049").

$ExistingADUser = $null
do {
    $InputUsername = Read-Host "Bitte den SamAccountName des bereits angelegten Benutzers eingeben (z.B. 049Mustermann)"
    $ExistingADUser = Get-ADUser -Identity $InputUsername -Properties StreetAddress, City, State, PostalCode, Country, Office -ErrorAction SilentlyContinue
    if (-not $ExistingADUser) {
        Write-Host "Benutzer '$InputUsername' wurde in AD nicht gefunden. Bitte erneut versuchen." -ForegroundColor Yellow
    }
} while (-not $ExistingADUser)

$saccount = $ExistingADUser.SamAccountName
$prefix = $saccount.Substring(0,3)

Write-Host "Gefundener Benutzer: $saccount (Prefix: $prefix)" -ForegroundColor Cyan

#####################################################################################################

Add-Type -AssemblyName System.Windows.Forms

[System.Windows.Forms.MessageBox]::Show("Bitte aktualisieren Sie die Datei: C:\IT\NewITUser\testtxt.txt. Sobald Sie fertig sind, klicken Sie auf OK.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

# Exit Excel if open (auch versteckte Hintergrund-Instanzen, nicht nur sichtbare Fenster)
Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.CloseMainWindow() | Out-Null }
Start-Sleep -Seconds 2
Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

# Open Excel file
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false # Prevents unnecessary warning messages
$Workbook = $Excel.Workbooks.Open("C:\IT\NewITUser\NewITUser.xlsm")
$WsSource = $Workbook.Sheets("NewStarters")

# Update data in Excel
# Jedes Connection/Query einzeln und synchron aktualisieren
foreach ($conn in $Workbook.Connections) {
    try { $conn.OLEDBConnection.BackgroundQuery = $false } catch {}
    try {
        Write-Host "Aktualisiere Connection: $($conn.Name)..." -ForegroundColor DarkCyan
        $conn.Refresh()
        Write-Host "  OK" -ForegroundColor Green
    }
    catch {
        Write-Host "  FEHLER beim Refresh von '$($conn.Name)': $($_.Exception.Message)" -ForegroundColor Red
    }
}
$Excel.CalculateFullRebuild()
Start-Sleep -Seconds 2

# Speichern, Excel komplett schliessen und beenden
Write-Host "Speichere und schliesse Excel komplett, um sauber neu zu laden..." -ForegroundColor DarkCyan
$Workbook.Save()
$Workbook.Close($false)
$Excel.Quit()

try {
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
} catch {}
Remove-Variable Workbook, Excel -ErrorAction SilentlyContinue
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
[GC]::Collect()

Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

# Frisch neu oeffnen - Formeln werden jetzt aus dem sauber gespeicherten Zustand geladen
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false
$Workbook = $Excel.Workbooks.Open("C:\IT\NewITUser\NewITUser.xlsm")
$WsSource = $Workbook.Sheets("NewStarters")

Write-Host "Excel wurde frisch neu geladen." -ForegroundColor Cyan

#####################################################################
##############          MANAGER                      ################
#####################################################################

# Read "Manager" value from sheet "ConstructData", cell "B12"
$ListConstructData = $Workbook.Worksheets.Item("ConstructData")
$ManagerName = $ListConstructData.Range("B12").Text
$ManagerNameSplit = $ManagerName -split " "

if ([string]::IsNullOrWhiteSpace($ManagerName) -or $ManagerName.StartsWith("#") -or $ManagerNameSplit.Count -lt 2) {
    Write-Host "Manager-Name aus Excel ist ungueltig ('$ManagerName')." -ForegroundColor Yellow
    if ($ManagerName -like "#*") {
        Write-Host "--> Das ist das bekannte Lizenzproblem (siehe Hinweis am Skriptanfang)." -ForegroundColor Yellow
        Write-Host "--> Kurzfassung: In Excel (unter 049ad4 angemeldet) einmalig als 049ext4... anmelden, Excel neu starten." -ForegroundColor Yellow
        Write-Host "--> Danach in PS ISE weiterhin als 049ad4... arbeiten (AD-Schreibrechte)." -ForegroundColor Yellow
    }
    $ADUsers = @()
}
else {
    # Search for "Manager" value in Active Directory
    $ManagerName0 = $ManagerNameSplit[0]
    $ManagerName1 = $ManagerNameSplit[1]

    $Filter = "DisplayName -like '*$ManagerName0*' -and DisplayName -like '*$ManagerName1*'"

    Write-Host $Filter

    # Search "Manager" SamAccountname in Active Directory
    $ADUsers = Get-ADUser -Filter $Filter | Select-Object -ExpandProperty SamAccountName

    Start-Sleep -Seconds 3
}

# Check if "Manager" SamAccountname were found
if ($ADUsers.Count -eq 0) {
    Write-Host "Es wurde kein Benutzer gefunden."
    $SamAccountName = Read-Host "Bitte geben Sie den Benutzernamen manuell ein:"
} elseif ($ADUsers.Count -eq 1) {
    # If only one "Manager" SamAccountname is found
    [string]$SamAccountName = $ADUsers
    Write-Host "Gefundener Benutzer: $SamAccountName"
} else {
    # If multiple "Manager" SamAccountname are found
    Write-Host "Es wurden mehrere Benutzer gefunden. Bitte wählen Sie einen aus:"

    # Display all "Manager" SamAccountname with numbering
    for ($i = 0; $i -lt $ADUsers.Count; $i++) {
        Write-Host "$i : $($ADUsers[$i])"
    }

    # User selects a number
    $Selection = Read-Host "Geben Sie die Nummer des gewünschten Benutzers ein"

    # Input validation
    if ($Selection -match '^\d+$' -and [int]$Selection -lt $ADUsers.Count) {
        $SamAccountName = $ADUsers[$Selection]
        Write-Host "Ausgewählter Benutzer: $SamAccountName"
    } else {
        Write-Host "Ungültige Auswahl. Es wird kein Benutzer verwendet."
        $SamAccountName = "N/A"
    }
}

$UpnnameManager = (Get-ADUser -Identity $SamAccountName -Properties UserPrincipalName).UserPrincipalName

# Write "Manager" SamAccountname and E-mail to Excel
Write-Host $SamAccountName
$BlattPowerShell = $Workbook.Worksheets.Item("PowerShell")
$BlattPowerShell.Range("J2").Value = [string]$SamAccountName

Write-Host $UpnnameManager
$BlattPowerShell = $Workbook.Worksheets.Item("E-mail")
$BlattPowerShell.Range("C2").Value = [string]$UpnnameManager

#####################################################################################################
######################                     SENDER                               #####################
#####################################################################################################

$Absender = Get-ADUser $env:USERNAME -Properties GivenName, Surname
$Absender = $Absender.GivenName + ' ' + $Absender.Surname
Write-Host $Absender

$BlattPowerShell = $Workbook.Worksheets.Item("E-mail")
$BlattPowerShell.Range("C1").Value = [string]$Absender

#####################################################################################################
######################                     GENERATE PASSWORD                    #####################
#####################################################################################################

# Der Benutzer wurde vom HR-System angelegt, daher kennt niemand das aktuelle Passwort.
# Es wird deshalb ein neues Passwort generiert, in Excel gespeichert (zum Ausdrucken)
# und weiter unten in AD zurueckgesetzt.

function New-MemorablePassword {
    param (
        [int]$Length = 12
    )

    # English-like syllables
    $syllables = @("ba", "be", "bi", "bo", "bu", "ca", "co", "da", "de", "fa", "fi", "go", "ha", "la", "ma", "na", "ra", "si", "ta", "za")
    $specialChars = "!#=?%§&*"
    $digits = "0123456789"

    # Generate a base string using syllables
    $base = ""
    while ($base.Length -lt ($Length - 4)) {
        $base += ($syllables | Get-Random)
    }

    $base = $base.Substring(0, $Length - 4)

    # Add three random digits and one special character
    $extras = @(
        ($digits.ToCharArray() | Get-Random -Count 3)
        ($specialChars.ToCharArray() | Get-Random -Count 1)
    )

    # Combine everything and capitalize the first letter
    $password = ($base + ($extras -join ''))
    $password = $password.Substring(0,1).ToUpper() + $password.Substring(1)

    return $password
}

$password = New-MemorablePassword
$BlattPowerShell = $Workbook.Worksheets.Item("Drucken")
$BlattPowerShell.Range("L15").Value = [string]$password

#####################################################################################################
######################                   EXCEL DATA UPDATE                      #####################
#####################################################################################################

$ListPowerShell = $Workbook.Worksheets.Item("PowerShell")
$ExcelPrefix = $ListPowerShell.Range("A2").Text
$uname = $ListPowerShell.Range("B2").Text
$name = $ListPowerShell.Range("C2").Text
$givenname = $ListPowerShell.Range("D2").Text
$surname = $ListPowerShell.Range("E2").Text
$dname = $ListPowerShell.Range("F2").Text
$ExcelSaccount = $ListPowerShell.Range("G2").Text
$office = $ListPowerShell.Range("H2").Text
$department = $ListPowerShell.Range("I2").Text
$manager = $SamAccountName
$job = $ListPowerShell.Range("K2").Text
$company = $ListPowerShell.Range("L2").Text
$address = $ListPowerShell.Range("M2").Text
$city = $ListPowerShell.Range("N2").Text
$state = $ListPowerShell.Range("O2").Text
$postal = $ListPowerShell.Range("P2").Text
$country = $ListPowerShell.Range("Q2").Text
$ou = $ListPowerShell.Range("R2").Text
$employeeID = $ListPowerShell.Range("S2").Text

$Workbook.Save()

# Hinweis, falls der in Excel berechnete SamAccountName vom eingegebenen Benutzer abweicht
# (z.B. wenn ein falscher Benutzer eingegeben wurde). Der eingegebene Benutzer ($saccount)
# bleibt in jedem Fall die massgebliche Quelle fuer alle AD-Operationen weiter unten.
if ($ExcelSaccount -and $ExcelSaccount -ne $saccount) {
    Write-Host "ACHTUNG: Der in Excel berechnete SamAccountName ('$ExcelSaccount') unterscheidet sich vom eingegebenen Benutzer ('$saccount')." -ForegroundColor Yellow
}

Write-Host "==================================================================================="

    Write-Host "Prefix: $prefix"
    Write-Host "UserPrincipalName: $uname"
    Write-Host "Name: $name"
    Write-Host "GivenName: $givenname"
    Write-Host "Surname: $surname"
    Write-Host "DisplayName: $dname"
    Write-Host "SamAccountName: $saccount"
    Write-Host "Office: $office"
    Write-Host "Department: $department"
    Write-Host "Manager: $manager"
    Write-Host "Title: $job"
    Write-Host "Company: $company"
    Write-Host "StreetAddress: $address"
    Write-Host "City: $city"
    Write-Host "State: $state"
    Write-Host "PostalCode: $postal"
    Write-Host "Country: $country"
    Write-Host "OrganizationalUnit:$ou"
    Write-Host "EmployeeID: $employeeID"
    Write-Host "Password: $password"

Write-Host "==================================================================================="

$Workbook.Save()

#####################################################################################################
######################                   LAST CHECK                             #####################
#####################################################################################################

$choice = Read-Host "Möchten Sie mit dem Skript fortfahren? (J/N)"

if ($choice -match '^[Jj]$') {
    Write-Host "Skript wird fortgesetzt..."
} else {
    Write-Host "Skript wird beendet."
    $Workbook.Close($true)
    $Excel.Quit()
    Stop-Process -Name "EXCEL" -Force -ErrorAction SilentlyContinue
    exit
}

#####################################################################################################
######################          BESTEHENDEN AD-BENUTZER AKTUALISIEREN           #####################
#####################################################################################################

# Adresse und Office werden nur auf Wunsch aktualisiert (kein Manager-Feld - das wird nur in Excel gepflegt)
# Die Werte kommen primaer aus dem bestehenden AD-Konto; nur wenn das jeweilige
# AD-Attribut leer ist ($null/""), wird der Wert aus Excel verwendet.
$finalOffice = if ($ExistingADUser.Office) { $ExistingADUser.Office } else { $office }
$finalAddress = if ($ExistingADUser.StreetAddress) { $ExistingADUser.StreetAddress } else { $address }
$finalCity = if ($ExistingADUser.City) { $ExistingADUser.City } else { $city }
$finalState = if ($ExistingADUser.State) { $ExistingADUser.State } else { $state }
$finalPostal = if ($ExistingADUser.PostalCode) { $ExistingADUser.PostalCode } else { $postal }
$finalCountry = if ($ExistingADUser.Country) { $ExistingADUser.Country } else { $country }

# Raum/Office = "-" bedeutet "kein Raum vorhanden" und wird nie als Wert nach AD geschrieben
if ($finalOffice -eq "-") {
    $finalOffice = $null
}

Write-Host "Office/Raum (final): $finalOffice"
Write-Host "StreetAddress (final): $finalAddress / $finalPostal $finalCity, $finalState, $finalCountry"

$updateChoice = Read-Host "Sollen Adresse und Office bei $saccount aktualisiert werden? (J/N)"

if ($updateChoice -match '^[Jj]$') {
    $SetADUserParams = @{
        Identity      = $saccount
        StreetAddress = $finalAddress
        City          = $finalCity
        State         = $finalState
        PostalCode    = $finalPostal
        Country       = $finalCountry
    }

    if ($finalOffice) {
        $SetADUserParams["Office"] = $finalOffice
    } else {
        Write-Host "Raum/Office ist '-' oder leer - Office wird NICHT gesetzt." -ForegroundColor Yellow
    }

    Set-ADUser @SetADUserParams

    Write-Host "Adresse und Office wurden fuer $saccount aktualisiert." -ForegroundColor Green
} else {
    Write-Host "Adresse und Office wurden NICHT aktualisiert." -ForegroundColor Yellow
}

# Passwort zuruecksetzen: Der Benutzer wurde vom HR-System angelegt, das aktuelle
# Passwort ist niemandem bekannt. Es wird deshalb zurueckgesetzt und muss beim
# naechsten Login geaendert werden.
Set-ADAccountPassword -Identity $saccount -Reset -NewPassword (ConvertTo-SecureString $password -AsPlainText -Force)
Set-ADUser -Identity $saccount -ChangePasswordAtLogon $true

Write-Output "Passwort fuer $saccount wurde zurueckgesetzt: $password"

# Manual group assignment (choice 1 or 2)
# Default = 1) Standard User (press Enter)

Write-Host ""
Write-Host "Choose group assignment for user: $saccount"
Write-Host "  1) Standard User"
Write-Host "  2) Production - First line worker"
Write-Host ""

$choice = Read-Host "Enter your choice (1 or 2) or just Enter for 1"

# Enter = default (Standard User)
if ([string]::IsNullOrWhiteSpace($choice)) {
    $choice = "1"
}

while ($choice -notin @("1", "2")) {
    $choice = Read-Host "Invalid input. Enter 1 or 2 (or press Enter for Standard User)"
    if ([string]::IsNullOrWhiteSpace($choice)) { $choice = "1" }
}

if ($choice -eq "1") {
    # Standard User (default)
    Add-ADGroupMember -Identity "049-OLD-Sennheiser" -Members $saccount
    Add-ADGroupMember -Identity "$prefix-rma-home" -Members $saccount
    if ($prefix -eq "149") {
        Add-ADGroupMember -Identity "149-LY-Enable-EnterpriseVoice_Barleben" -Members $saccount
        Add-ADGroupMember -Identity "149-EX-Enable-Mail_Barleben" -Members $saccount
    }
    else {
        Add-ADGroupMember -Identity "$prefix-LY-Enable-EnterpriseVoice" -Members $saccount
        try {
            Add-ADGroupMember -Identity "$prefix-EX-Enable-Mail" -Members $saccount
        }
        catch {
            Add-ADGroupMember -Identity "$prefix-EX-Enabled_Mail" -Members $saccount
        }
    }
    Write-Host "Added $saccount to Standard User groups."
}
else {
    # Production - First line worker
    Add-ADGroupMember -Identity "QA-FLW" -Members $saccount
    Add-ADGroupMember -Identity "049-EX-Enable_Mail_PlanF3" -Members $saccount
    Write-Host "Added $saccount to Production - First line worker groups."
}

# Add Mitarbeiterhandbuch group if prefix is 249 (applies to both choice 1 and 2)
if ($prefix -eq "249") {
    Add-ADGroupMember -Identity "249-NTFS-Mitarbeiterhandbuch-R" -Members $saccount
    Add-ADGroupMember -Identity "249-NTFS-Neumann_ALL-R" -Members $saccount
    Write-Host "Added $saccount to 249-NTFS-Mitarbeiterhandbuch-R (Neumann staff)." -ForegroundColor Cyan
}

# Save and close file
$Workbook.Save()
$Workbook.Close($true)
$Excel.Quit()

# Release COM objects to free memory
try {
    if ($wsDest) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsDest) | Out-Null }
    if ($wbDest) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbDest) | Out-Null }
    if ($WsSource) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WsSource) | Out-Null }
    if ($Workbook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null }
    if ($Excel) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null }
}
catch {
    Write-Host "Error releasing COM objects: $_"
}

[GC]::Collect()
[GC]::WaitForPendingFinalizers()
[GC]::Collect()

# Exit Excel
Stop-Process -Name "EXCEL" -Force -ErrorAction SilentlyContinue

Write-Host "Daten wurden aktualisiert und das Skript erfolgreich abgeschlossen."
