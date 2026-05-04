#Requires -Version 5.1
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$RepoUrl       = 'https://raw.githubusercontent.com/KevinDSM/kevin-lust-musiques/main'
$ReleaseUrl    = 'https://github.com/KevinDSM/kevin-lust-musiques/releases/latest/download/KevinLust.zip'
$VersionUrl    = 'https://raw.githubusercontent.com/KevinDSM/kevin-lust-musiques/main/version.txt'

Write-Host ''
Write-Host '=============================================' -ForegroundColor Cyan
Write-Host '  Kevin Lust - Installateur / Mise a jour'   -ForegroundColor Cyan
Write-Host '=============================================' -ForegroundColor Cyan
Write-Host ''

# ======================================================================
#  FONCTION DIAGNOSTIC : analyse une exception reseau et renvoie
#  un message clair + une suggestion de correction.
# ======================================================================
function Get-DiagError {
    param([System.Exception]$ex)

    $msg = $ex.Message + ' ' + ($ex.InnerException.Message -as [string])

    # --- Pas d'internet / DNS ----------------------------------------
    if ($msg -match 'remote name could not be resolved|The server name or address could not be resolved|No such host|name resolution') {
        return @{
            Cause  = 'DNS / Pas de connexion internet'
            Fix    = @(
                '  -> Verifie que tu es bien connecte a internet.',
                '  -> Desactive ton VPN si tu en as un, puis reessaie.',
                '  -> Si tu es sur Wi-Fi, essaie de te reconnecter au reseau.'
            )
        }
    }

    # --- Timeout ---------------------------------------------------------
    if ($msg -match 'timed out|The operation has timed out|TaskCanceledException') {
        return @{
            Cause  = 'Timeout (connexion trop lente ou GitHub inaccessible)'
            Fix    = @(
                '  -> Ta connexion est peut-etre trop lente pour ce fichier.',
                '  -> Reessaie dans quelques minutes (GitHub peut etre temporairement charge).',
                '  -> Si le probleme persiste, contacte Kevin pour recevoir les fichiers directement.'
            )
        }
    }

    # --- Erreur HTTP 403 / 429 -------------------------------------------
    if ($msg -match '403|Forbidden|429|Too Many Requests') {
        return @{
            Cause  = 'GitHub a refuse la connexion (403/429 - trop de requetes ou acces bloque)'
            Fix    = @(
                '  -> Attends 5-10 minutes et reessaie.',
                '  -> Si tu utilises un VPN ou un proxy, desactive-le.',
                '  -> GitHub limite parfois les telechargements rapides depuis certaines IP.'
            )
        }
    }

    # --- Erreur HTTP 404 -------------------------------------------------
    if ($msg -match '404|Not Found') {
        return @{
            Cause  = 'Fichier introuvable sur GitHub (404)'
            Fix    = @(
                '  -> Le fichier a peut-etre ete renomme ou supprime.',
                '  -> Contacte Kevin pour qu il verifie le depot GitHub.'
            )
        }
    }

    # --- Erreur SSL / certificat -----------------------------------------
    if ($msg -match 'SSL|TLS|certificate|Could not establish trust relationship|The underlying connection was closed') {
        return @{
            Cause  = 'Erreur de certificat SSL/TLS'
            Fix    = @(
                '  -> Verifie que la date et l heure de ton PC sont correctes.',
                '  -> Desactive temporairement ton antivirus / pare-feu et reessaie.',
                '  -> Mets a jour Windows si tes mises a jour sont en retard.'
            )
        }
    }

    # --- Proxy / pare-feu ------------------------------------------------
    if ($msg -match 'proxy|407|Proxy Authentication') {
        return @{
            Cause  = 'Bloque par un proxy ou un pare-feu reseau'
            Fix    = @(
                '  -> Tu es peut-etre sur un reseau d entreprise ou scolaire qui bloque GitHub.',
                '  -> Essaie depuis une autre connexion (partage de connexion mobile par exemple).',
                '  -> Ou contacte Kevin pour recevoir les fichiers directement.'
            )
        }
    }

    # --- Generique -------------------------------------------------------
    return @{
        Cause  = $ex.Message
        Fix    = @(
            '  -> Verifie ta connexion internet et reessaie.',
            '  -> Si le probleme persiste, note ce message et contacte Kevin.'
        )
    }
}

# ======================================================================
#  FONCTION TEST CONNECTIVITE : verifie qu on peut joindre GitHub
# ======================================================================
function Test-GitHubConnectivity {
    try {
        $null = Invoke-WebRequest -Uri 'https://github.com' -UseBasicParsing -TimeoutSec 10 -Method Head
        return $true
    } catch {
        return $false
    }
}

# -------- 1. Detecter le dossier AddOns de WoW --------
Write-Host '[1/4] Recherche de World of Warcraft...' -ForegroundColor White

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

$AddonDir   = Join-Path $AddOnsDir 'KevinLust'
$SongsDir   = Join-Path $AddonDir  'Songs'
$SongsLua   = Join-Path $AddonDir  'Songs.lua'
$LocalVerFile = Join-Path $AddonDir 'version.txt'

Write-Host ("  -> Dossier AddOns : {0}" -f $AddOnsDir) -ForegroundColor Gray

# -------- 2. Verification de version --------
Write-Host ''
Write-Host '[2/4] Verification de la version...' -ForegroundColor White

$addonInstalled = Test-Path $AddonDir
$localVersion   = if ((Test-Path $LocalVerFile)) { (Get-Content $LocalVerFile -Raw).Trim() } else { $null }
$remoteVersion  = $null

try {
    $remoteVersion = (Invoke-WebRequest -Uri $VersionUrl -UseBasicParsing -TimeoutSec 15).Content.Trim()
} catch {
    Write-Host '  Impossible de verifier la version en ligne.' -ForegroundColor Yellow
    Write-Host '  -> Le script va continuer et telecharger la derniere version disponible.' -ForegroundColor Yellow
}

$needsAddonUpdate = $true

if ($remoteVersion) {
    if (-not $addonInstalled) {
        Write-Host ("  Premiere installation  (version : {0})" -f $remoteVersion) -ForegroundColor White
    } elseif ($localVersion -eq $remoteVersion) {
        Write-Host ("  Addon deja a jour  [v{0}]  - aucun telechargement necessaire." -f $localVersion) -ForegroundColor Green
        $needsAddonUpdate = $false
    } else {
        $from = if ($localVersion) { "v$localVersion" } else { "version inconnue" }
        Write-Host ("  Mise a jour disponible : {0}  ->  v{1}" -f $from, $remoteVersion) -ForegroundColor Cyan
    }
}

# -------- 3. Installer ou mettre a jour l'addon (si besoin) --------
Write-Host ''
if ($needsAddonUpdate) {
    if ($addonInstalled) {
        Write-Host '[3/4] Mise a jour de l addon Kevin Lust...' -ForegroundColor White
    } else {
        Write-Host '[3/4] Addon Kevin Lust non installe. Telechargement...' -ForegroundColor White
    }

    $tmpZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), 'KevinLust.zip')

    # Test de connectivite avant de tenter le telechargement
    Write-Host '  Verification de la connexion a GitHub...' -NoNewline
    if (-not (Test-GitHubConnectivity)) {
        Write-Host '  [ECHEC]' -ForegroundColor Red
        Write-Host ''
        Write-Host '  ERREUR : Impossible de joindre GitHub.' -ForegroundColor Red
        Write-Host '  CAUSE  : Pas de connexion internet ou GitHub est inaccessible.' -ForegroundColor DarkRed
        Write-Host ''
        Write-Host '  Que faire :' -ForegroundColor Yellow
        Write-Host '  -> Verifie que ton PC est bien connecte a internet.' -ForegroundColor Yellow
        Write-Host '  -> Desactive ton VPN si tu en as un.' -ForegroundColor Yellow
        Write-Host '  -> Reessaie dans quelques minutes.' -ForegroundColor Yellow
        Write-Host ''
        pause
        exit 1
    }
    Write-Host '  [OK]' -ForegroundColor Green

    try {
        Write-Host '  Telechargement de l addon...' -NoNewline
        Invoke-WebRequest -Uri $ReleaseUrl -OutFile $tmpZip -UseBasicParsing -TimeoutSec 120
        Write-Host '  [OK]' -ForegroundColor Green
    } catch {
        Write-Host '  [ECHEC]' -ForegroundColor Red
        $diag = Get-DiagError $_.Exception
        Write-Host ''
        Write-Host ("  ERREUR : {0}" -f $diag.Cause) -ForegroundColor Red
        foreach ($line in $diag.Fix) { Write-Host $line -ForegroundColor Yellow }
        Write-Host ''
        Write-Host '  Contact : demande le dossier KevinLust directement a Kevin' -ForegroundColor Yellow
        Write-Host ("  et place-le dans : {0}" -f $AddOnsDir) -ForegroundColor Yellow
        if (Test-Path $tmpZip) { Remove-Item $tmpZip -Force }
        Write-Host ''
        pause
        exit 1
    }

    try {
        Write-Host '  Extraction dans AddOns...' -NoNewline
        Expand-Archive -Path $tmpZip -DestinationPath $AddOnsDir -Force
        Write-Host '  [OK]' -ForegroundColor Green
    } catch {
        Write-Host '  [ECHEC]' -ForegroundColor Red
        Write-Host ''
        Write-Host '  ERREUR : Impossible d extraire le fichier zip.' -ForegroundColor Red
        Write-Host ("  DETAIL : {0}" -f $_.Exception.Message) -ForegroundColor DarkRed
        Write-Host ''
        Write-Host '  Que faire :' -ForegroundColor Yellow
        Write-Host '  -> Verifie que tu as les droits d ecriture dans le dossier AddOns.' -ForegroundColor Yellow
        Write-Host '  -> Essaie de lancer le .bat en tant qu administrateur (clic droit -> Executer en tant qu administrateur).' -ForegroundColor Yellow
        Write-Host ("  -> Dossier cible : {0}" -f $AddOnsDir) -ForegroundColor Yellow
        if (Test-Path $tmpZip) { Remove-Item $tmpZip -Force }
        Write-Host ''
        pause
        exit 1
    } finally {
        if (Test-Path $tmpZip) { Remove-Item $tmpZip -Force -ErrorAction SilentlyContinue }
    }

    # Sauvegarder la version installee
    if ($remoteVersion) {
        Set-Content -Path $LocalVerFile -Value $remoteVersion -Encoding UTF8
    }

    if ($addonInstalled) {
        Write-Host ("  -> Addon mis a jour vers v{0} : {1}" -f $remoteVersion, $AddonDir) -ForegroundColor Green
    } else {
        Write-Host ("  -> Addon installe (v{0}) dans : {1}" -f $remoteVersion, $AddonDir) -ForegroundColor Green
    }
} else {
    Write-Host '[3/4] Addon a jour - telechargement ignore.' -ForegroundColor Green
}

# Creer le dossier Songs s'il manque
if (-not (Test-Path $SongsDir)) {
    New-Item -ItemType Directory -Path $SongsDir | Out-Null
}

# -------- 4. Telecharger le manifest --------
Write-Host ''
Write-Host '[4/4] Telechargement du manifest des musiques...' -ForegroundColor White

try {
    $apiResp = (Invoke-WebRequest -Uri 'https://api.github.com/repos/KevinDSM/kevin-lust-musiques/contents/manifest.txt' -UseBasicParsing -TimeoutSec 30).Content | ConvertFrom-Json
    $manifestText = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String(($apiResp.content -replace "`n","")))
} catch {
    $diag = Get-DiagError $_.Exception
    Write-Host ''
    Write-Host '  ECHEC : impossible de recuperer la liste des musiques.' -ForegroundColor Red
    Write-Host ("  CAUSE  : {0}" -f $diag.Cause) -ForegroundColor DarkRed
    foreach ($line in $diag.Fix) { Write-Host $line -ForegroundColor Yellow }
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

# -------- 5. Fenetre de selection des musiques --------

# Palette de couleurs theme sombre WoW
$clrBg      = [System.Drawing.Color]::FromArgb(28, 28, 40)
$clrHeader  = [System.Drawing.Color]::FromArgb(18, 18, 30)
$clrPanel   = [System.Drawing.Color]::FromArgb(22, 22, 34)
$clrGold    = [System.Drawing.Color]::FromArgb(255, 200, 50)
$clrText    = [System.Drawing.Color]::FromArgb(218, 218, 230)
$clrDim     = [System.Drawing.Color]::FromArgb(120, 120, 148)
$clrGreen   = [System.Drawing.Color]::FromArgb(80, 210, 105)
$clrBtnBg   = [System.Drawing.Color]::FromArgb(48, 48, 68)
$clrBtnText = [System.Drawing.Color]::FromArgb(210, 210, 230)
$clrSearch  = [System.Drawing.Color]::FromArgb(36, 36, 52)
$clrBorder  = [System.Drawing.Color]::FromArgb(65, 65, 95)

# Trier : musiques deja installees en premier, puis ordre alphabetique
$entriesSorted = $entries | Sort-Object {
    $inst = Test-Path (Join-Path $SongsDir $_.SafeFile)
    if ($inst) { "0_$($_.Label.ToLower())" } else { "1_$($_.Label.ToLower())" }
}

# Helper : cree un bouton stylise avec le theme sombre
function New-StyledButton($text, $x, $y, $w, $h, $parent) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text      = $text
    $b.Location  = New-Object System.Drawing.Point($x, $y)
    $b.Size      = New-Object System.Drawing.Size($w, $h)
    $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $b.BackColor = $clrBtnBg
    $b.ForeColor = $clrBtnText
    $b.Font      = New-Object System.Drawing.Font('Segoe UI', 9)
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $b.FlatAppearance.BorderColor      = $clrBorder
    $b.FlatAppearance.BorderSize       = 1
    $b.FlatAppearance.MouseOverBackColor  = [System.Drawing.Color]::FromArgb(68, 68, 95)
    $b.FlatAppearance.MouseDownBackColor  = [System.Drawing.Color]::FromArgb(38, 38, 58)
    $parent.Controls.Add($b)
    return $b
}

# ---- Formulaire principal ----
$form = New-Object System.Windows.Forms.Form
$form.Text            = 'Kevin Lust - Musiques'
$form.Size            = New-Object System.Drawing.Size(580, 700)
$form.StartPosition   = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.BackColor       = $clrBg

# ---- En-tete ----
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Location  = New-Object System.Drawing.Point(0, 0)
$headerPanel.Size      = New-Object System.Drawing.Size(580, 58)
$headerPanel.BackColor = $clrHeader
$form.Controls.Add($headerPanel)

$titleLbl = New-Object System.Windows.Forms.Label
$titleLbl.Text      = '  ♪  KEVIN LUST  -  Selection des musiques'
$titleLbl.Location  = New-Object System.Drawing.Point(0, 10)
$titleLbl.Size      = New-Object System.Drawing.Size(580, 24)
$titleLbl.Font      = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
$titleLbl.ForeColor = $clrGold
$titleLbl.BackColor = [System.Drawing.Color]::Transparent
$headerPanel.Controls.Add($titleLbl)

$subLbl = New-Object System.Windows.Forms.Label
$subLbl.Text      = '  Coche pour installer  ·  Decoche pour supprimer'
$subLbl.Location  = New-Object System.Drawing.Point(0, 36)
$subLbl.Size      = New-Object System.Drawing.Size(580, 18)
$subLbl.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
$subLbl.ForeColor = $clrDim
$subLbl.BackColor = [System.Drawing.Color]::Transparent
$headerPanel.Controls.Add($subLbl)

# ---- Barre de recherche ----
$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location    = New-Object System.Drawing.Point(15, 70)
$searchBox.Size        = New-Object System.Drawing.Size(390, 28)
$searchBox.Font        = New-Object System.Drawing.Font('Segoe UI', 10)
$searchBox.BackColor   = $clrSearch
$searchBox.ForeColor   = $clrDim
$searchBox.BorderStyle = 'FixedSingle'
$searchBox.Text        = 'Rechercher...'
$form.Controls.Add($searchBox)

# ---- Compteur dynamique ----
$counterLbl = New-Object System.Windows.Forms.Label
$counterLbl.Location  = New-Object System.Drawing.Point(415, 73)
$counterLbl.Size      = New-Object System.Drawing.Size(148, 20)
$counterLbl.Font      = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$counterLbl.ForeColor = $clrGold
$counterLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$form.Controls.Add($counterLbl)

# Fonction mise a jour du compteur
function Update-Counter {
    $checked = ($checkboxes.Values | Where-Object { $_.Checked }).Count
    $total   = $checkboxes.Count
    $counterLbl.Text = "$checked / $total cochees"
}

# ---- Boutons Tout cocher / Tout decocher ----
$btnAll  = New-StyledButton 'Tout cocher'    15  108  118 28 $form
$btnNone = New-StyledButton 'Tout decocher' 140  108  128 28 $form

# ---- Separateur ----
$sep = New-Object System.Windows.Forms.Label
$sep.Location  = New-Object System.Drawing.Point(15, 144)
$sep.Size      = New-Object System.Drawing.Size(545, 1)
$sep.BackColor = $clrBorder
$form.Controls.Add($sep)

# ---- Panel scrollable (FlowLayoutPanel pour le filtrage dynamique) ----
$panel = New-Object System.Windows.Forms.FlowLayoutPanel
$panel.Location      = New-Object System.Drawing.Point(15, 148)
$panel.Size          = New-Object System.Drawing.Size(545, 490)
$panel.AutoScroll    = $true
$panel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
$panel.WrapContents  = $false
$panel.BackColor     = $clrPanel
$panel.Padding       = New-Object System.Windows.Forms.Padding(6, 4, 0, 4)
$form.Controls.Add($panel)

# ---- Creation des checkboxes ----
$checkboxes = @{}   # file -> CheckBox
$cbItems    = [System.Collections.Generic.List[hashtable]]::new()

foreach ($e in $entriesSorted) {
    $isInstalled = Test-Path (Join-Path $SongsDir $e.SafeFile)

    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Size      = New-Object System.Drawing.Size(520, 26)
    $cb.Checked   = $isInstalled
    $cb.BackColor = $clrPanel
    $cb.Font      = New-Object System.Drawing.Font('Segoe UI', 9)
    $cb.Cursor    = [System.Windows.Forms.Cursors]::Hand

    if ($isInstalled) {
        $cb.Text      = "  $($e.Label)   [installee]"
        $cb.ForeColor = $clrGreen
    } else {
        $cb.Text      = "  $($e.Label)"
        $cb.ForeColor = $clrText
    }

    $cb.Add_CheckedChanged({ Update-Counter })
    $panel.Controls.Add($cb)
    $checkboxes[$e.File] = $cb
    $cbItems.Add(@{ Entry = $e ; CB = $cb })
}

Update-Counter

# ---- Recherche en temps reel ----
$searchBox.Add_GotFocus({
    if ($searchBox.Text -eq 'Rechercher...') {
        $searchBox.Text      = ''
        $searchBox.ForeColor = $clrText
    }
})
$searchBox.Add_LostFocus({
    if ($searchBox.Text -eq '') {
        $searchBox.Text      = 'Rechercher...'
        $searchBox.ForeColor = $clrDim
    }
})
$searchBox.Add_TextChanged({
    $filter = $searchBox.Text.Trim()
    if ($filter -eq 'Rechercher...') { $filter = '' }
    foreach ($item in $cbItems) {
        $item.CB.Visible = ($filter -eq '' -or $item.Entry.Label -match [regex]::Escape($filter))
    }
})

# Tout cocher / decocher uniquement sur les cases visibles (respecte le filtre)
$btnAll.Add_Click({
    foreach ($item in $cbItems) { if ($item.CB.Visible) { $item.CB.Checked = $true } }
})
$btnNone.Add_Click({
    foreach ($item in $cbItems) { if ($item.CB.Visible) { $item.CB.Checked = $false } }
})

# ---- Boutons bas ----
$btnOk = New-StyledButton 'Appliquer' 368 648 105 30 $form
$btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
$btnOk.ForeColor    = [System.Drawing.Color]::FromArgb(80, 220, 105)
$btnOk.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(70, 170, 80)
$form.AcceptButton = $btnOk

$btnCancel = New-StyledButton 'Annuler' 478 648 88 30 $form
$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $btnCancel

$formResult = $form.ShowDialog()
if ($formResult -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host '  Annule par l utilisateur.' -ForegroundColor Yellow
    Write-Host ''
    pause
    exit 0
}

$selection = @{}
foreach ($k in $checkboxes.Keys) { $selection[$k] = $checkboxes[$k].Checked }

# -------- 6. Appliquer : telechargements / suppressions --------
Write-Host ''
Write-Host 'Application des changements...' -ForegroundColor White

$downloaded = 0
$skipped    = 0
$removed    = 0
$failed     = @()   # tableau de hashtables {File, Cause, Fix}

foreach ($e in $entries) {
    $dest    = Join-Path $SongsDir $e.SafeFile
    $wanted  = $selection[$e.File]
    $present = Test-Path $dest

    if ($wanted -and -not $present) {
        Write-Host ("  + {0}" -f $e.Label) -NoNewline
        $url = "$RepoUrl/$([System.Uri]::EscapeDataString($e.File))"
        try {
            Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing -TimeoutSec 90
            Write-Host '   [OK]' -ForegroundColor Green
            $downloaded++
        } catch {
            Write-Host '   [ECHEC]' -ForegroundColor Red
            $diag = Get-DiagError $_.Exception
            $failed += @{ File = $e.File ; Label = $e.Label ; Diag = $diag }
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
    Write-Host ("  Echecs       : {0}" -f $failed.Count) -ForegroundColor Red
    Write-Host ''
    Write-Host '  --- DETAIL DES ECHECS ---' -ForegroundColor Red
    foreach ($f in $failed) {
        Write-Host ''
        Write-Host ("  [ECHEC] {0}" -f $f.Label) -ForegroundColor Red
        Write-Host ("  CAUSE   : {0}" -f $f.Diag.Cause) -ForegroundColor DarkRed
        Write-Host '  QUE FAIRE :' -ForegroundColor Yellow
        foreach ($line in $f.Diag.Fix) { Write-Host $line -ForegroundColor Yellow }
    }
    Write-Host ''
    Write-Host '  Conseil : reessaie dans quelques minutes. Si ca bloque toujours,' -ForegroundColor Yellow
    Write-Host '  contacte Kevin en lui donnant le message d erreur ci-dessus.' -ForegroundColor Yellow
}

# -------- 7. Regenerer Songs.lua --------
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
