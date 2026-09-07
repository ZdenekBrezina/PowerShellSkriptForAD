Import-Module ActiveDirectory

#####################################################################
##############     BENUTZER PER E-MAIL IDENTIFIZIEREN    ############
#####################################################################

# Der Benutzer wurde bereits angelegt. Wir identifizieren ihn anhand
# seiner E-Mail-Adresse (UserPrincipalName) und lesen daraus alle
# weiteren benoetigten Werte direkt aus dem AD-Konto.

$ExistingADUser = $null
do {
    $InputEmail = Read-Host "Bitte die E-Mail-Adresse (UserPrincipalName) des Benutzers eingeben"
    $ExistingADUser = Get-ADUser -Filter "UserPrincipalName -eq '$InputEmail'" -Properties GivenName, Surname, UserPrincipalName, Manager -ErrorAction SilentlyContinue
    if (-not $ExistingADUser) {
        Write-Host "Kein Benutzer mit der E-Mail-Adresse '$InputEmail' gefunden. Bitte erneut versuchen." -ForegroundColor Yellow
    }
} while (-not $ExistingADUser)

$saccount = $ExistingADUser.SamAccountName
$givenname = $ExistingADUser.GivenName
$surname = $ExistingADUser.Surname
$upn = $ExistingADUser.UserPrincipalName
$prefix = $saccount.Substring(0,3)
$domain = ($upn -split "@")[1]
$fullname = "$givenname $surname"

Write-Host "Gefundener Benutzer: $saccount ($upn)" -ForegroundColor Cyan

# Manager aus dem AD-Attribut "Manager" des Benutzers ermitteln
$managerName = ""
if ($ExistingADUser.Manager) {
    $managerName = (Get-ADUser -Identity $ExistingADUser.Manager -Properties DisplayName).DisplayName
    Write-Host "Manager: $managerName" -ForegroundColor Cyan
} else {
    Write-Host "Kein Manager-Attribut fuer $saccount in AD gesetzt." -ForegroundColor Yellow
}

#####################################################################
##############          PASSWORT GENERIEREN               ###########
#####################################################################

function New-MemorablePassword {
    param (
        [int]$Length = 12
    )

    $syllables = @("ba", "be", "bi", "bo", "bu", "ca", "co", "da", "de", "fa", "fi", "go", "ha", "la", "ma", "na", "ra", "si", "ta", "za")
    $specialChars = "!#=?%§&*"
    $digits = "0123456789"

    $base = ""
    while ($base.Length -lt ($Length - 4)) {
        $base += ($syllables | Get-Random)
    }
    $base = $base.Substring(0, $Length - 4)

    $extras = @(
        ($digits.ToCharArray() | Get-Random -Count 3)
        ($specialChars.ToCharArray() | Get-Random -Count 1)
    )

    $password = ($base + ($extras -join ''))
    $password = $password.Substring(0,1).ToUpper() + $password.Substring(1)

    return $password
}

$password = New-MemorablePassword

#####################################################################
##############          EXCEL AKTUALISIEREN                ##########
#####################################################################

Add-Type -AssemblyName System.Windows.Forms

# Falls Excel noch offen ist, schliessen
Get-Process -Name Excel -ErrorAction SilentlyContinue | ForEach-Object { if ($_.MainWindowHandle -ne 0) { $_.CloseMainWindow() } }
Start-Sleep -Seconds 2
Get-Process -Name Excel -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { $_.Kill() } }

$Excel = New-Object -ComObject Excel.Application
$Excel.DisplayAlerts = $false
$Excel.Visible = $true
$Workbook = $Excel.Workbooks.Open("C:\IT\NewITUser\ExternalUser.xlsm")
$SheetDrucken = $Workbook.Worksheets.Item("Drucken")
$SheetConstructData = $Workbook.Worksheets.Item("ConstructData")

$SheetDrucken.Range("L10").Value = [string]$givenname
$SheetDrucken.Range("L11").Value = [string]$surname
$SheetDrucken.Range("L13").Value = [string]$upn
$SheetDrucken.Range("L14").Value = [string]$saccount
$SheetDrucken.Range("L15").Value = [string]$password

$SheetConstructData.Range("B5").Value = [string]$givenname
$SheetConstructData.Range("B6").Value = [string]$surname
$SheetConstructData.Range("B7").Value = [string]$prefix
$SheetConstructData.Range("B8").Value = [string]$domain
$SheetConstructData.Range("B9").Value = [string]$upn
$SheetConstructData.Range("B10").Value = [string]$saccount
$SheetConstructData.Range("B12").Value = [string]$managerName
$SheetConstructData.Range("B13").Value = [string]$fullname

$Workbook.Save()

Write-Host "==================================================================================="
Write-Host "SamAccountName: $saccount"
Write-Host "GivenName: $givenname"
Write-Host "Surname: $surname"
Write-Host "UserPrincipalName: $upn"
Write-Host "Password: $password"
Write-Host "==================================================================================="
Write-Host "ExternalUser.xlsm wurde aktualisiert und bleibt geoeffnet (Sheet 'Drucken')." -ForegroundColor Green

#####################################################################
##############          PASSWORT ZURUECKSETZEN              ##########
#####################################################################

Set-ADAccountPassword -Identity $saccount -Reset -NewPassword (ConvertTo-SecureString $password -AsPlainText -Force)
Set-ADUser -Identity $saccount -ChangePasswordAtLogon $true

Write-Host "Passwort fuer $saccount wurde zurueckgesetzt." -ForegroundColor Green

#####################################################################
##############          GRUPPEN HINZUFUEGEN                 ##########
#####################################################################

Add-ADGroupMember -Identity "O365-Licensing-Plan_F3" -Members $saccount
Add-ADGroupMember -Identity "O365-Licensing-Plan_F3_withEXO" -Members $saccount

Write-Host "Benutzer $saccount wurde zu O365-Licensing-Plan_F3 und O365-Licensing-Plan_F3_withEXO hinzugefuegt." -ForegroundColor Green

Write-Host "Skript erfolgreich abgeschlossen."
