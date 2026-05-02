#Requires -Version 5.1
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$RepoUrl    = 'https://raw.githubusercontent.com/KevinDSM/kevin-lust-musiques/main'
$ReleaseUrl = 'https://github.com/KevinDSM/kevin-lust-musiques/releases/latest/download/KevinLust.zip'

Write-Host ''
Write-Host '=============================================' -ForegroundColor Cyan
Write-Host '  Kevin Lust - Installateur / Mise a jour'   -ForegroundColor Cyan
Write-Host '=============================================' -ForegroundColor Cyan
Write-Host ''

# -------- 1. Detecter le dossier AddOns de WoW --------
Write-Host '[1/3] Recherche de World of Warcraft...' -ForegroundColor White

$AddOnsDir = $null

# Tentative via le registre Windows
try {
    $reg = Get-ItemProperty `
        -Path 'HKLM:\SOFTWARE\WOW6432Node\Blizzard Entertainment\World of Warcraft' `
        -ErrorAction SilentlyContinue
    if ($reg -and $reg.InstallPath) {
        $candidate = Join-Path $reg.InstallPath '_retail_\Interface\AddOns'
        if (Test-Path $candidate) { $AddOnsDir = $candidate }
    }
} catch {}

# Chemins classiques en fallback
if (-not $AddOnsDir) {
    $fallbacks = @(
        'C:\Program Files (x86)\World of Warcraft\_retail_\Interface\AddOns',
        'C:\Program Files\World of Warcraft\_retail_\Interface\AddOns',
        'D:\World of Warcraft\_retail_\Interface\AddOns',
        'D:\Games\World of Warcraft\_retail_\Interface\AddOns',
        'E:\World of Warcraft\_retail_\Interface\AddOns'
    )
    foreach ($c in $fallbacks) {
        if (Test-Path $c) { $AddOnsDir = $c; break }
    }
}

# Introuvable -> l'utilisateur choisit manuellement
if (-not $AddOnsDir) {
    Write-Host '  WoW introuvable automatiquement.' -ForegroundColor Yellow
    Write-Host '  Selectionne le dossier AddOns manuellement...' -ForegroundColor Yellow
    $browser = New-Object System.Windows.Forms.FolderBrowserDialog
    $browser.Description = "Impossible de trouver WoW. Selectionne le dossier AddOns de WoW : _retail_\Interface\AddOns"
    $browser.ShowNewFolderButton = $false
    if ($browser.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host '  Installation annulee.' -ForegroundColor Yellow
        exit 0
    }
    $AddOnsDir = $browser.SelectedPath
}

$AddonDir = Join-Path $AddOnsDir 'KevinLust'
$SongsDir = Join-Path $AddonDir  'Songs'
$SongsLua = Join-Path $AddonDir  'Songs.lua'

Write-Host ("  -> Dossier AddOns : {0}" -f $AddOnsDir) -ForegroundColor Gray

# -------- 2. Installer ou mettre a jour l'addon --------
$addonInstalled = Test-Path $AddonDir

Write-Host ''
if ($addonInstalled) {
    Write-Host '[2/3] Mise a jour de l addon Kevin Lust...' -ForegroundColor White
} else {
    Write-Host '[2/3] Addon Kevin Lust non installe. Telechargement...' -ForegroundColor White
}

$tmpZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), 'KevinLust.zip')
try {
    Write-Host '  Connexion a GitHub...' -NoNewline
    Invoke-WebRequest -Uri $ReleaseUrl -OutFile $tmpZip -UseBasicParsing -TimeoutSec 120
    Write-Host '  [OK]' -ForegroundColor Green
} catch {
    Write-Host '  [ECHEC]' -ForegroundColor Red
    Write-Host "  Impossible de telecharger l'addon : $($_.Exception.Message)" -ForegroundColor DarkRed
    Write-Host ''
    Write-Host '  Que faire :' -ForegroundColor Yellow
    Write-Host '  - Verifie ta connexion internet' -ForegroundColor Yellow
    Write-Host '  - Ou demande le dossier KevinLust directement a Kevin' -ForegroundColor Yellow
    Write-Host ("    et place-le dans : {0}" -f $AddOnsDir) -ForegroundColor Yellow
    if (Test-Path $tmpZip) { Remove-Item $tmpZip -Force }
    Write-Host ''
    pause
    exit 1
}

try {
    Write-Host '  Extraction dans AddOns...' -NoNewline

    # Sauvegarder Songs/ et Songs.lua avant d'ecraser (ils ne sont pas dans le zip
    # mais Expand-Archive -Force ne les supprime pas non plus — precaution par securite)
    Expand-Archive -Path $tmpZip -DestinationPath $AddOnsDir -Force

    Write-Host '  [OK]' -ForegroundColor Green
} catch {
    Write-Host '  [ECHEC]' -ForegroundColor Red
    Write-Host "  Erreur extraction : $($_.Exception.Message)" -ForegroundColor DarkRed
    if (Test-Path $tmpZip) { Remove-Item $tmpZip -Force }
    Write-Host ''
    pause
    exit 1
} finally {
    if (Test-Path $tmpZip) { Remove-Item $tmpZip -Force -ErrorAction SilentlyContinue }
}

if ($addonInstalled) {
    Write-Host ("  -> Addon mis a jour : {0}" -f $AddonDir) -ForegroundColor Green
} else {
    Write-Host ("  -> Addon installe dans : {0}" -f $AddonDir) -ForegroundColor Green
}

# Creer le dossier Songs s'il manque (premiere installation ou nettoyage)
if (-not (Test-Path $SongsDir)) {
    New-Item -ItemType Directory -Path $SongsDir | Out-Null
}

# -------- 3. Telecharger le manifest --------
Write-Host ''
Write-Host '[3/3] Telechargement du manifest des musiques...' -ForegroundColor White
try {
    $apiResp = (Invoke-WebRequest -Uri 'https://api.github.com/repos/KevinDSM/kevin-lust-musiques/contents/manifest.txt' -UseBasicParsing -TimeoutSec 30).Content | ConvertFrom-Json
    $manifestText = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String(($apiResp.content -replace "`n","")))
} catch {
    Write-Host '  ECHEC : impossible de joindre GitHub.' -ForegroundColor Red
    Write-Host "  Detail : $($_.Exception.Message)" -ForegroundColor DarkRed
    Write-Host ''
    Write-Host "  L'addon est installe mais les musiques n'ont pas pu etre chargees." -ForegroundColor Yellow
    Write-Host '  Relance le script plus tard pour les ajouter.' -ForegroundColor Yellow
    Write-Host ''
    pause
    exit 1
}

$entries = @()
foreach ($line in ($manifestText -split "`r?`n")) {
    $line = $line.Trim()
    if ($line -eq '' -or $line.StartsWith('#')) { continue }
    $parts = $line.Split('|', 2)
    if ($parts.Count -ne 2) { continue }
    $entries += [pscustomobject]@{
        File     = $parts[0].Trim()
        SafeFile = ($parts[0].Trim() -replace ' ', '_')
        Label    = $parts[1].Trim()
    }
}

Write-Host ("  -> {0} morceau(x) disponible(s) sur GitHub." -f $entries.Count) -ForegroundColor Gray

# Pas encore de musiques dans le manifest
if ($entries.Count -eq 0) {
    Write-Host ''
    if (-not $addonInstalled) {
        Write-Host '  Addon installe avec succes !' -ForegroundColor Green
    }
    Write-Host '  Aucune musique dans le manifest pour l instant.' -ForegroundColor Yellow
    Write-Host '  Lance WoW et tape /djlust pour ouvrir les reglages.' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '=============================================' -ForegroundColor Cyan
    Write-Host '  Termine !' -ForegroundColor Cyan
    Write-Host '=============================================' -ForegroundColor Cyan
    Write-Host ''
    pause
    exit 0
}

# -------- 4. Fenetre de selection des musiques --------
$form = New-Object System.Windows.Forms.Form
$form.Text            = 'Kevin Lust - Selection des musiques'
$form.Size            = New-Object System.Drawing.Size(560, 660)
$form.StartPosition   = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false

# Titre dynamique selon contexte
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location  = New-Object System.Drawing.Point(15, 12)
$titleLabel.Size      = New-Object System.Drawing.Size(515, 18)
$titleLabel.Font      = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
if (-not $addonInstalled) {
    $titleLabel.Text     = 'Addon installe !  Coche les musiques a telecharger :'
    $titleLabel.ForeColor = [System.Drawing.Color]::DarkGreen
} else {
    $titleLabel.Text = 'Coche les musiques a installer dans ton dossier Kevin Lust.'
}
$form.Controls.Add($titleLabel)

$subLabel = New-Object System.Windows.Forms.Label
$subLabel.Text      = 'Decoche pour SUPPRIMER une musique deja installee.'
$subLabel.Location  = New-Object System.Drawing.Point(15, 32)
$subLabel.Size      = New-Object System.Drawing.Size(515, 18)
$subLabel.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($subLabel)

$btnAll = New-Object System.Windows.Forms.Button
$btnAll.Text     = 'Tout cocher'
$btnAll.Location = New-Object System.Drawing.Point(15, 58)
$btnAll.Size     = New-Object System.Drawing.Size(110, 26)
$form.Controls.Add($btnAll)

$btnNone = New-Object System.Windows.Forms.Button
$btnNone.Text     = 'Tout decocher'
$btnNone.Location = New-Object System.Drawing.Point(132, 58)
$btnNone.Size     = New-Object System.Drawing.Size(110, 26)
$form.Controls.Add($btnNone)

$panel = New-Object System.Windows.Forms.Panel
$panel.Location    = New-Object System.Drawing.Point(15, 93)
$panel.Size        = New-Object System.Drawing.Size(515, 490)
$panel.AutoScroll  = $true
$panel.BorderStyle = 'FixedSingle'
$form.Controls.Add($panel)

$checkboxes = @{}
$y = 8
foreach ($e in $entries) {
    $isInstalled = Test-Path (Join-Path $SongsDir $e.SafeFile)
    $cb = New-Object System.Windows.Forms.CheckBox
    $suffix = if ($isInstalled) { '   [deja installe]' } else { '' }
    $cb.Text     = "{0}    ({1}){2}" -f $e.Label, $e.File, $suffix
    $cb.Location = New-Object System.Drawing.Point(10, $y)
    $cb.Size     = New-Object System.Drawing.Size(488, 24)
    $cb.Checked  = $isInstalled
    if ($isInstalled) { $cb.ForeColor = [System.Drawing.Color]::DarkGreen }
    $panel.Controls.Add($cb)
    $checkboxes[$e.File] = $cb
    $y += 26
}

$btnAll.Add_Click({  foreach ($cb in $checkboxes.Values) { $cb.Checked = $true  } })
$btnNone.Add_Click({ foreach ($cb in $checkboxes.Values) { $cb.Checked = $false } })

$btnOk = New-Object System.Windows.Forms.Button
$btnOk.Text         = 'Appliquer'
$btnOk.Location     = New-Object System.Drawing.Point(355, 592)
$btnOk.Size         = New-Object System.Drawing.Size(85, 30)
$btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.Controls.Add($btnOk)
$form.AcceptButton  = $btnOk

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text         = 'Annuler'
$btnCancel.Location     = New-Object System.Drawing.Point(445, 592)
$btnCancel.Size         = New-Object System.Drawing.Size(85, 30)
$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.Controls.Add($btnCancel)
$form.CancelButton      = $btnCancel

$formResult = $form.ShowDialog()
if ($formResult -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host '  Annule par l utilisateur.' -ForegroundColor Yellow
    Write-Host ''
    pause
    exit 0
}

$selection = @{}
foreach ($k in $checkboxes.Keys) { $selection[$k] = $checkboxes[$k].Checked }

# -------- 5. Appliquer : telechargements / suppressions --------
Write-Host ''
Write-Host 'Application des changements...' -ForegroundColor White

$downloaded = 0
$skipped    = 0
$removed    = 0
$failed     = @()

foreach ($e in $entries) {
    $dest    = Join-Path $SongsDir $e.SafeFile
    $wanted  = $selection[$e.File]
    $present = Test-Path $dest

    if ($wanted -and -not $present) {
        Write-Host ("  + {0}" -f $e.Label) -NoNewline
        try {
            Invoke-WebRequest -Uri "$RepoUrl/$([System.Uri]::EscapeDataString($e.File))" -OutFile $dest -UseBasicParsing -TimeoutSec 60
            Write-Host '   [OK]' -ForegroundColor Green
            $downloaded++
        } catch {
            Write-Host '   [ECHEC]' -ForegroundColor Yellow
            $failed += $e.File
            if (Test-Path $dest) { Remove-Item $dest -Force }
        }
    } elseif ($wanted -and $present) {
        $skipped++
    } elseif (-not $wanted -and $present) {
        Write-Host ("  - {0} (suppression)" -f $e.Label) -ForegroundColor DarkYellow
        Remove-Item $dest -Force
        $removed++
    }
}

Write-Host ''
Write-Host ("  Telecharges  : {0}" -f $downloaded) -ForegroundColor Green
Write-Host ("  Conserves    : {0}" -f $skipped)    -ForegroundColor Gray
Write-Host ("  Supprimes    : {0}" -f $removed)    -ForegroundColor DarkYellow
if ($failed.Count -gt 0) {
    Write-Host ("  Echecs       : {0}" -f ($failed -join ', ')) -ForegroundColor Yellow
}

# -------- 6. Regenerer Songs.lua --------
Write-Host ''
Write-Host 'Mise a jour de Songs.lua...' -ForegroundColor White

$lines = @()
$lines += '-- Songs.lua : AUTO-GENERE par Install-KevinLust.bat. Ne pas editer a la main.'
$lines += '-- Pour ajouter un morceau, utilise plutot le manifest GitHub.'
$lines += ''
$lines += 'local addonName, addon = ...'
$lines += ''
$lines += 'addon.SONG_MANIFEST = {'
$lines += '    { label = "Chipi Chipi (built-in)", path = "Interface\\AddOns\\KevinLust\\chipilust.mp3" },'
$lines += '    { label = "Pedro (built-in)",       path = "Interface\\AddOns\\KevinLust\\pedrolust.mp3" },'

foreach ($e in $entries) {
    $dest = Join-Path $SongsDir $e.SafeFile
    if (-not (Test-Path $dest)) { continue }
    $luaLabel = $e.Label -replace '\\', '\\\\' -replace '"', '\"'
    $luaPath  = 'Interface\\AddOns\\KevinLust\\Songs\\' + $e.SafeFile
    $lines += ('    {{ label = "{0}", path = "{1}" }},' -f $luaLabel, $luaPath)
}
$lines += '}'
$lines += ''

Set-Content -Path $SongsLua -Value $lines -Encoding UTF8
Write-Host '  -> Songs.lua mis a jour.' -ForegroundColor Green
Write-Host ''

Write-Host '=============================================' -ForegroundColor Cyan
Write-Host '  Termine !' -ForegroundColor Cyan
Write-Host '  Lance WoW (ou tape /reload si deja en jeu).' -ForegroundColor Cyan
Write-Host '=============================================' -ForegroundColor Cyan
Write-Host ''
pause
