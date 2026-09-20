<#
.SYNOPSIS
    Installe le complement « Envoi programme » dans Outlook classique.

.DESCRIPTION
    Deux modes d'hebergement du volet :

    -Mode Local   (defaut) : le volet est servi par un petit serveur HTTPS qui
                  tourne sur le poste, sur 127.0.0.1:3000. Rien ne sort du poste.
                  Exige d'installer un certificat auto-signe pour localhost :
                  Windows affiche alors un avertissement a accepter, UNE fois.

    -Mode Distant : le volet est servi par une URL HTTPS publique (GitHub Pages,
                  Cloudflare Pages...). Aucun certificat, aucun serveur, aucune
                  tache planifiee : seul le manifeste est installe.

    Dans les deux cas, le complement est charge par la cle de registre de
    sideload d'Office (HKCU), sans passer par le centre d'administration.

.EXAMPLE
    .\installer.ps1
    .\installer.ps1 -Mode Distant -UrlBase https://gcfiscalite.github.io/outlook-envoi-differe
    .\installer.ps1 -Desinstaller
#>

[CmdletBinding()]
param(
    [ValidateSet('Local', 'Distant')]
    [string] $Mode = 'Local',

    [string] $UrlBase,

    [int] $Port = 3000,

    [switch] $Desinstaller
)

$ErrorActionPreference = 'Stop'

$Racine   = Split-Path -Parent $PSScriptRoot
$Outils   = Join-Path $Racine 'tools'
$Manifest = Join-Path $Racine 'manifest.xml'
$CleWef   = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
$NomValeur = 'GCFiscalite.EnvoiProgramme'
$NomTache = 'GC - Complement Outlook envoi programme'
$SujetCert = 'GC Fiscalite - complement Outlook'

function Info    { param($m) Write-Host "  $m" }
function Etape   { param($m) Write-Host "`n$m" -ForegroundColor Cyan }
function Succes  { param($m) Write-Host "  $m" -ForegroundColor Green }
function Alerte  { param($m) Write-Host "  $m" -ForegroundColor Yellow }

# Lanceur Python unique pour tout le script : « py -3.12 » si le lanceur Windows
# est la, sinon le python du PATH.
$script:PyExe  = $null
$script:PyArgs = @()

function Initialize-Python {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        $script:PyExe  = 'py'
        $script:PyArgs = @('-3.12')
    } elseif (Get-Command python -ErrorAction SilentlyContinue) {
        $script:PyExe  = 'python'
        $script:PyArgs = @()
    } else {
        throw "Python introuvable. Installer Python 3.12 ou l'ajouter au PATH."
    }
    return "$script:PyExe $($script:PyArgs -join ' ')".Trim()
}

function Invoke-Py {
    param([Parameter(ValueFromRemainingArguments = $true)] [string[]] $Arguments)
    & $script:PyExe @($script:PyArgs + $Arguments)
}

# ---------------------------------------------------------------- desinstallation

if ($Desinstaller) {
    Etape 'Desinstallation'

    if (Test-Path $CleWef) {
        Remove-ItemProperty -Path $CleWef -Name $NomValeur -ErrorAction SilentlyContinue
        Succes 'Complement retire du registre Office.'
    }

    $tache = Get-ScheduledTask -TaskName $NomTache -ErrorAction SilentlyContinue
    if ($tache) {
        Unregister-ScheduledTask -TaskName $NomTache -Confirm:$false
        Succes 'Tache planifiee supprimee.'
    }

    # Get-Process n'expose pas CommandLine en PowerShell 5.1 : on passe par CIM.
    Get-CimInstance Win32_Process -Filter "Name='pythonw.exe' OR Name='python.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -like '*serveur.py*' } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }

    $certs = Get-ChildItem Cert:\CurrentUser\Root |
             Where-Object { $_.Subject -like "*$SujetCert*" }
    foreach ($c in $certs) {
        Remove-Item $c.PSPath -Force
        Succes "Certificat retire : $($c.Thumbprint)"
    }

    Write-Host "`nDesinstalle. Fermer et rouvrir Outlook pour que le bouton disparaisse.`n"
    return
}

# ---------------------------------------------------------------- verifications

Etape 'Verification de l''environnement'

Info "Python : $(Initialize-Python)"

$versionOutlook = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' `
                   -ErrorAction SilentlyContinue).VersionToReport
if ($versionOutlook) {
    Info "Office : $versionOutlook"
    $v = [version]$versionOutlook
    # Mailbox 1.13 exige la version 2304, build 16327.20248.
    if ($v -lt [version]'16.0.16327.20248') {
        throw "Office $versionOutlook est trop ancien : l'envoi programme cote serveur exige 2304 (16.0.16327.20248) ou plus recent."
    }
    Succes 'Version compatible avec Mailbox 1.13.'
} else {
    Alerte "Version d'Office indeterminee, on continue."
}

# ---------------------------------------------------------------- hebergement

if ($Mode -eq 'Distant') {
    if (-not $UrlBase) { throw "-Mode Distant exige -UrlBase https://..." }
    $base = $UrlBase.TrimEnd('/')
} else {
    $base = "https://localhost:$Port"

    Etape 'Certificat du serveur local'

    $existant = Get-ChildItem Cert:\CurrentUser\Root |
                Where-Object { $_.Subject -like "*$SujetCert*" -and $_.NotAfter -gt (Get-Date) }

    if ($existant) {
        Succes "Deja installe (expire le $($existant[0].NotAfter.ToString('yyyy-MM-dd')))."
    } else {
        Invoke-Py (Join-Path $Outils 'generer_certificat.py')

        Alerte "Windows va demander confirmation pour installer le certificat."
        Alerte "C'est normal : il ne vaut que pour localhost, sur ce poste."

        $magasin = New-Object System.Security.Cryptography.X509Certificates.X509Store('Root', 'CurrentUser')
        $magasin.Open('ReadWrite')
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(
                    (Join-Path $Outils 'cert\localhost.cer'))
        $magasin.Add($cert)
        $magasin.Close()
        Succes "Certificat installe : $($cert.Thumbprint)"
    }

    Etape 'Serveur local'

    $scriptServeur = Join-Path $Outils 'serveur.py'
    $pythonw = (Get-Command pythonw -ErrorAction SilentlyContinue).Source
    if (-not $pythonw) {
        $pythonw = Join-Path (Split-Path (Invoke-Py -c "import sys;print(sys.executable)")) 'pythonw.exe'
    }
    Info "pythonw : $pythonw"

    $action    = New-ScheduledTaskAction -Execute $pythonw -Argument "`"$scriptServeur`" --port $Port"
    $declench  = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
    $reglages  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries `
                    -DontStopIfGoingOnBatteries -StartWhenAvailable `
                    -ExecutionTimeLimit ([TimeSpan]::Zero)

    Register-ScheduledTask -TaskName $NomTache -Action $action -Trigger $declench `
        -Settings $reglages -Description 'Sert le volet du complement Outlook sur 127.0.0.1 uniquement.' `
        -Force | Out-Null
    Succes "Tache planifiee creee : demarre a l'ouverture de session."

    $dejaLa = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
    if (-not $dejaLa) {
        Start-Process -FilePath $pythonw -ArgumentList "`"$scriptServeur`" --port $Port" -WindowStyle Hidden
        Start-Sleep -Seconds 2
        Succes 'Serveur demarre.'
    } else {
        Succes 'Serveur deja en ecoute.'
    }
}

# ---------------------------------------------------------------- manifeste

Etape 'Manifeste'

Invoke-Py (Join-Path $Outils 'generer_manifeste.py') --base $base
Invoke-Py (Join-Path $Outils 'verifier.py')
if ($LASTEXITCODE -ne 0) { throw "Le manifeste n'a pas passe la verification." }

# ---------------------------------------------------------------- sideload

Etape 'Chargement dans Outlook'

if (-not (Test-Path $CleWef)) { New-Item -Path $CleWef -Force | Out-Null }
New-ItemProperty -Path $CleWef -Name $NomValeur -Value $Manifest `
    -PropertyType String -Force | Out-Null
Succes "Inscrit dans $CleWef"

Write-Host @"

Termine.

  1. Fermer Outlook completement, puis le rouvrir.
  2. Nouveau message : le bouton « Programmer l'envoi » apparait dans le ruban.
  3. Si le bouton manque, patienter : Outlook met parfois jusqu'a une heure
     a relire un manifeste sideloade, et jusqu'a 24 h dans le pire des cas.

  Mode : $Mode
  Volet : $base/src/taskpane.html

"@
