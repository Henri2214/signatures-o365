<#
.SYNOPSIS
    Importe des signatures Outlook (Outlook classique / Microsoft 365 Apps) depuis un dossier.

.DESCRIPTION
    - Copie les fichiers .htm/.rtf/.txt et les dossiers "<nom>_files" (images) dans le dossier
      de signatures de l'utilisateur (%APPDATA%\Microsoft\Signatures ou nom localise).
    - Optionnel : definit la signature par defaut (nouveaux messages / reponses-transferts)
      pour tous les comptes de tous les profils Outlook.
    - Optionnel : desactive les "roaming signatures" (signatures cloud M365) qui peuvent
      ecraser les signatures locales.
    A executer dans le contexte de l'utilisateur (pas en SYSTEM), ex. script de logon GPO / Intune.

.EXAMPLE
    .\Import-OutlookSignatures.ps1 -SourcePath "\\srv\partage\Signatures" -DefaultNew "Societe" -DefaultReply "Societe-court" -DisableRoaming -CloseOutlook
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SourcePath,

    # Nom de la signature (sans extension) pour les nouveaux messages
    [string]$DefaultNew,

    # Nom de la signature (sans extension) pour les reponses/transferts
    [string]$DefaultReply,

    # Desactive les signatures itinerantes M365 (sinon Outlook peut resynchroniser depuis le cloud)
    [switch]$DisableRoaming,

    # Ferme Outlook avant l'import (recommande pour que les valeurs registre soient prises en compte)
    [switch]$CloseOutlook,

    # Remplace les signatures existantes portant le meme nom
    [switch]$Force = $true
)

$ErrorActionPreference = 'Stop'
$OfficeVer = '16.0'   # Office 2016/2019/2021/M365
$OfficeKey = "HKCU:\Software\Microsoft\Office\$OfficeVer"

function Write-Log($msg, $color = 'Gray') { Write-Host "[$(Get-Date -Format HH:mm:ss)] $msg" -ForegroundColor $color }

# --- 0. Verifications ---------------------------------------------------------
if (-not (Test-Path $SourcePath)) { throw "Dossier source introuvable : $SourcePath" }

$sigFiles = Get-ChildItem -Path $SourcePath -File | Where-Object Extension -in '.htm', '.html', '.rtf', '.txt'
if (-not $sigFiles) { throw "Aucun fichier de signature (.htm/.rtf/.txt) dans $SourcePath" }

# --- 1. Fermeture d'Outlook ----------------------------------------------------
if ($CloseOutlook) {
    $ol = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
    if ($ol) {
        Write-Log "Fermeture d'Outlook..." Yellow
        $ol | ForEach-Object { $_.CloseMainWindow() | Out-Null }
        Start-Sleep -Seconds 5
        Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Stop-Process -Force
    }
}

# --- 2. Dossier de destination (nom localise possible) -------------------------
$folderName = (Get-ItemProperty "$OfficeKey\Common\General" -Name 'Signatures' -ErrorAction SilentlyContinue).Signatures
if ([string]::IsNullOrWhiteSpace($folderName)) { $folderName = 'Signatures' }
$dest = Join-Path $env:APPDATA "Microsoft\$folderName"
if (-not (Test-Path $dest)) { New-Item -ItemType Directory -Path $dest | Out-Null }
Write-Log "Destination : $dest" Cyan

# --- 3. Copie des signatures -----------------------------------------------------
$imported = @()
foreach ($f in $sigFiles) {
    $target = Join-Path $dest $f.Name
    if ((Test-Path $target) -and -not $Force) { Write-Log "Ignore (existe) : $($f.Name)" DarkYellow; continue }
    Copy-Item $f.FullName $target -Force
    $imported += $f.BaseName
    Write-Log "Copie : $($f.Name)" Green
}

# Dossiers d'images associes (<nom>_files ou <nom>-fichiers selon la langue)
Get-ChildItem -Path $SourcePath -Directory | Where-Object { $_.Name -match '(_files|_fichiers|-fichiers)$' } | ForEach-Object {
    Copy-Item $_.FullName (Join-Path $dest $_.Name) -Recurse -Force
    Write-Log "Dossier images copie : $($_.Name)" Green
}

$imported = $imported | Sort-Object -Unique
Write-Log "Signatures importees : $($imported -join ', ')" Cyan

# --- 4. Desactivation des signatures itinerantes ---------------------------------
if ($DisableRoaming) {
    $setup = "$OfficeKey\Outlook\Setup"
    if (-not (Test-Path $setup)) { New-Item -Path $setup -Force | Out-Null }
    New-ItemProperty -Path $setup -Name 'DisableRoamingSignaturesTemporaryToggle' -Value 1 -PropertyType DWord -Force | Out-Null
    Write-Log "Roaming signatures desactivees (DisableRoamingSignaturesTemporaryToggle=1)" Yellow
}

# --- 5. Signatures par defaut ------------------------------------------------------
function Set-DefaultSignature([string]$ValueName, [string]$SigName) {
    if (-not (Test-Path (Join-Path $dest "$SigName.htm"))) {
        Write-Log "Signature '$SigName' introuvable dans $dest - non definie par defaut" Red
        return
    }
    $bytes = [System.Text.Encoding]::Unicode.GetBytes("$SigName`0")

    # Par compte, dans chaque profil Outlook
    $accounts = Get-ChildItem "$OfficeKey\Outlook\Profiles\*\9375CFF0413111d3B88A00104B2A6676" -ErrorAction SilentlyContinue |
                Get-ChildItem | Where-Object { $_.GetValue('Account Name') }
    foreach ($acc in $accounts) {
        Set-ItemProperty -Path $acc.PSPath -Name $ValueName -Value $bytes -Type Binary
        Write-Log "  $ValueName = '$SigName' -> $([System.Text.Encoding]::Unicode.GetString($acc.GetValue('Account Name')) -replace "`0")" Green
    }
    if (-not $accounts) { Write-Log "Aucun compte Outlook trouve (profil pas encore cree ?)" DarkYellow }
}

if ($DefaultNew -or $DefaultReply) {
    # Valeurs globales (equivalent GPO) : appliquees aussi aux comptes crees plus tard
    $mail = "$OfficeKey\Common\MailSettings"
    if (-not (Test-Path $mail)) { New-Item -Path $mail -Force | Out-Null }

    if ($DefaultNew) {
        New-ItemProperty -Path $mail -Name 'NewSignature' -Value $DefaultNew -PropertyType String -Force | Out-Null
        Set-DefaultSignature 'New Signature' $DefaultNew
    }
    if ($DefaultReply) {
        New-ItemProperty -Path $mail -Name 'ReplySignature' -Value $DefaultReply -PropertyType String -Force | Out-Null
        Set-DefaultSignature 'Reply-Forward Signature' $DefaultReply
    }
}

Write-Log "Termine. Relancer Outlook pour prendre en compte les changements." Cyan
