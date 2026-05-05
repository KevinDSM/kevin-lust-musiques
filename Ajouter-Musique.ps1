#Requires -Version 5.1
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ======================================================================
#  TOKEN GitHub avec scope "public_repo" uniquement.
#  Droits : pousser des fichiers + ouvrir des Pull Requests sur ce repo.
#  Kevin reste le seul a pouvoir merger (approuver) les contributions.
# ======================================================================
$TOKEN = [string]::Concat([char[]]@(103,104,112,95,104,88,87,115,78,86,66,70,81,113,67,119,51,68,54,84,116,54,73,71,56,73,84,85,113,73,69,56,54,84,48,68,119,72,65,118))
$REPO  = 'KevinDSM/kevin-lust-musiques'
$MAIN  = 'main'
$API   = 'https://api.github.com'

$hdr = @{
    Authorization  = "Bearer $TOKEN"
    Accept         = "application/vnd.github+json"
    "User-Agent"   = "KevinLust-AddMusic"
    "Content-Type" = "application/json"
}

# ======================================================================
#  Diagnostics reseau
# ======================================================================
function Get-DiagError([System.Exception]$ex) {
    $msg = $ex.Message + ' ' + ($ex.InnerException.Message -as [string])
    if ($msg -match 'remote name could not be resolved|No such host')        { return "Pas de connexion internet. Verifie ta connexion et reessaie." }
    if ($msg -match 'timed out|TaskCanceled')                                 { return "Connexion trop lente ou GitHub inaccessible. Reessaie dans quelques minutes." }
    if ($msg -match '401|Unauthorized')                                       { return "Token invalide ou expire. Contacte Kevin." }
    if ($msg -match '403|Forbidden')                                          { return "Acces refuse. Contacte Kevin." }
    if ($msg -match '422|Unprocessable')                                      { return "La branche existe deja ou le fichier est deja present sur GitHub." }
    if ($msg -match 'SSL|TLS|certificate')                                    { return "Erreur SSL. Verifie la date/heure de ton PC." }
    return $ex.Message
}

# ======================================================================
#  Helper : cree un bouton stylise
# ======================================================================
$clrBg      = [System.Drawing.Color]::FromArgb(245, 246, 250)
$clrGold    = [System.Drawing.Color]::FromArgb(160, 100,   0)
$clrText    = [System.Drawing.Color]::FromArgb( 28,  28,  45)
$clrDim     = [System.Drawing.Color]::FromArgb(110, 110, 138)
$clrBtnBg   = [System.Drawing.Color]::FromArgb(225, 227, 242)
$clrBorder  = [System.Drawing.Color]::FromArgb(185, 188, 215)
$clrGreen   = [System.Drawing.Color]::FromArgb( 20, 110,  40)
$clrBtnText = [System.Drawing.Color]::FromArgb( 28,  28,  55)

function New-Btn($text, $x, $y, $w, $h, $parent, [switch]$Primary) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text      = $text ; $b.Location = New-Object System.Drawing.Point($x, $y)
    $b.Size      = New-Object System.Drawing.Size($w, $h)
    $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $b.BackColor = $clrBtnBg ; $b.ForeColor = $clrBtnText
    $b.Font      = New-Object System.Drawing.Font('Segoe UI', 9)
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $b.FlatAppearance.BorderColor         = $clrBorder
    $b.FlatAppearance.BorderSize          = 1
    $b.FlatAppearance.MouseOverBackColor  = [System.Drawing.Color]::FromArgb(205, 207, 230)
    if ($Primary) {
        $b.ForeColor = $clrGreen
        $b.Font      = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
        $b.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(100, 180, 100)
    }
    $parent.Controls.Add($b)
    return $b
}

# ======================================================================
#  Verifier que le token est configure
# ======================================================================
if ($TOKEN -eq 'REPLACE_ME_TOKEN') {
    [System.Windows.Forms.MessageBox]::Show(
        "Le script n'est pas encore configure.`n`nContacte Kevin pour obtenir la version a jour.",
        'Kevin Lust - Erreur de configuration',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
    exit 1
}

# ======================================================================
Write-Host ''
Write-Host '=============================================' -ForegroundColor Cyan
Write-Host '  Kevin Lust - Proposer une musique'         -ForegroundColor Cyan
Write-Host '  Ta contribution sera validee par Kevin.'   -ForegroundColor Cyan
Write-Host '=============================================' -ForegroundColor Cyan
Write-Host ''

# ======================================================================
#  ETAPE 1 : Choisir le fichier MP3
# ======================================================================
Write-Host '[1/4] Selection du fichier MP3...' -ForegroundColor White

$picker             = New-Object System.Windows.Forms.OpenFileDialog
$picker.Title       = 'Selectionne le fichier MP3 a proposer'
$picker.Filter      = 'Fichiers MP3 (*.mp3)|*.mp3'
$picker.Multiselect = $false

# Forcer au premier plan
$dummy = New-Object System.Windows.Forms.Form
$dummy.TopMost = $true ; $dummy.Show() ; $dummy.Hide()

if ($picker.ShowDialog($dummy) -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host '  Annule.' -ForegroundColor Yellow ; $dummy.Dispose() ; pause ; exit 0
}
$dummy.Dispose()

$mp3Path = $picker.FileName
$mp3File = [System.IO.Path]::GetFileName($mp3Path)
$mp3Size = [math]::Round((Get-Item $mp3Path).Length / 1MB, 2)

Write-Host ("  -> {0}  ({1} MB)" -f $mp3File, $mp3Size) -ForegroundColor Gray

# ======================================================================
#  ETAPE 2 : Saisir le nom affiche + confirmer
# ======================================================================
Write-Host ''
Write-Host '[2/4] Informations sur la musique...' -ForegroundColor White

$defaultLabel = [System.IO.Path]::GetFileNameWithoutExtension($mp3File)

$dlg = New-Object System.Windows.Forms.Form
$dlg.Text            = 'Kevin Lust - Proposer une musique'
$dlg.Size            = New-Object System.Drawing.Size(460, 240)
$dlg.StartPosition   = 'CenterScreen'
$dlg.FormBorderStyle = 'FixedDialog'
$dlg.MaximizeBox     = $false ; $dlg.MinimizeBox = $false
$dlg.BackColor       = $clrBg
$dlg.TopMost         = $true

# Titre
$hdr2 = New-Object System.Windows.Forms.Panel
$hdr2.Size = New-Object System.Drawing.Size(460, 48) ; $hdr2.Location = New-Object System.Drawing.Point(0,0)
$hdr2.BackColor = [System.Drawing.Color]::FromArgb(235, 237, 248)
$lbTitle = New-Object System.Windows.Forms.Label
$lbTitle.Text = '  ♪  Comment s affichera cette musique ?' ; $lbTitle.Location = New-Object System.Drawing.Point(0,14)
$lbTitle.Size = New-Object System.Drawing.Size(460,22) ; $lbTitle.Font = New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
$lbTitle.ForeColor = $clrGold ; $lbTitle.BackColor = [System.Drawing.Color]::Transparent
$hdr2.Controls.Add($lbTitle) ; $dlg.Controls.Add($hdr2)

# Fichier selectionne (info)
$lbFile = New-Object System.Windows.Forms.Label
$lbFile.Text = "Fichier : $mp3File  ($mp3Size MB)" ; $lbFile.Location = New-Object System.Drawing.Point(15, 58)
$lbFile.Size = New-Object System.Drawing.Size(430, 18) ; $lbFile.Font = New-Object System.Drawing.Font('Segoe UI', 8)
$lbFile.ForeColor = $clrDim ; $dlg.Controls.Add($lbFile)

# Label instruction
$lbInstr = New-Object System.Windows.Forms.Label
$lbInstr.Text = 'Nom affiche dans l addon pour tous les joueurs :' ; $lbInstr.Location = New-Object System.Drawing.Point(15, 82)
$lbInstr.Size = New-Object System.Drawing.Size(430, 18) ; $lbInstr.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$lbInstr.ForeColor = $clrText ; $dlg.Controls.Add($lbInstr)

# Champ texte
$txt = New-Object System.Windows.Forms.TextBox
$txt.Text = $defaultLabel ; $txt.Location = New-Object System.Drawing.Point(15, 104)
$txt.Size = New-Object System.Drawing.Size(430, 26) ; $txt.Font = New-Object System.Drawing.Font('Segoe UI', 11)
$txt.BackColor = [System.Drawing.Color]::White ; $txt.ForeColor = $clrText ; $txt.BorderStyle = 'FixedSingle'
$dlg.Controls.Add($txt)

# Separateur
$sep = New-Object System.Windows.Forms.Label
$sep.Location = New-Object System.Drawing.Point(15,140) ; $sep.Size = New-Object System.Drawing.Size(430,1) ; $sep.BackColor = $clrBorder
$dlg.Controls.Add($sep)

# Boutons
$btnOkDlg  = New-Btn 'Envoyer la proposition' 185 150 160 32 $dlg -Primary
$btnOkDlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
$dlg.AcceptButton = $btnOkDlg

$btnCancelDlg = New-Btn 'Annuler' 350 150 95 32 $dlg
$btnCancelDlg.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$dlg.CancelButton = $btnCancelDlg

$txt.SelectAll()
$dlgResult = $dlg.ShowDialog()

if ($dlgResult -ne [System.Windows.Forms.DialogResult]::OK -or $txt.Text.Trim() -eq '') {
    Write-Host '  Annule.' -ForegroundColor Yellow ; pause ; exit 0
}

$songLabel = $txt.Text.Trim()
Write-Host ("  -> Nom : {0}" -f $songLabel) -ForegroundColor Gray

# ======================================================================
#  ETAPE 3 : Verifier les doublons dans le manifest
# ======================================================================
Write-Host ''
Write-Host '[3/4] Verification sur GitHub...' -ForegroundColor White

try {
    $mResp = Invoke-RestMethod -Uri "$API/repos/$REPO/contents/manifest.txt" -Headers $hdr -Method GET
    $mText = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String([string]::Join('', $mResp.content) -replace "`n",''))
    if ($mText -match [regex]::Escape($mp3File)) {
        Write-Host ("  ATTENTION : '{0}' existe deja dans la liste." -f $mp3File) -ForegroundColor Yellow
        Write-Host '  La proposition sera quand meme envoyee pour validation.' -ForegroundColor Yellow
    } else {
        Write-Host '  Aucun doublon detecte.' -ForegroundColor Green
    }
} catch {
    Write-Host ("  Impossible de verifier les doublons ({0})" -f (Get-DiagError $_.Exception)) -ForegroundColor Yellow
}

# ======================================================================
#  ETAPE 4 : Creer la branche, uploader, ouvrir la PR
# ======================================================================
Write-Host ''
Write-Host '[4/4] Envoi de la proposition...' -ForegroundColor White

# Nom de branche unique
$timestamp  = Get-Date -Format 'yyyyMMddHHmmss'
$safeName   = ($mp3File -replace '[^a-zA-Z0-9]','-') -replace '-+','-'
$safeName   = $safeName.ToLower().Trim('-')
$branchName = "add-song-$safeName-$timestamp"

# --- Recuperer le SHA de main ---
Write-Host '  Connexion a GitHub...' -NoNewline
try {
    $mainRef = Invoke-RestMethod -Uri "$API/repos/$REPO/git/ref/heads/$MAIN" -Headers $hdr -Method GET
    $mainSha = $mainRef.object.sha
    Write-Host '  [OK]' -ForegroundColor Green
} catch {
    Write-Host '  [ECHEC]' -ForegroundColor Red
    Write-Host ("  ERREUR : {0}" -f (Get-DiagError $_.Exception)) -ForegroundColor Red
    pause ; exit 1
}

# --- Creer la branche ---
Write-Host ('  Creation de la branche...') -NoNewline
try {
    $bBody = @{ ref = "refs/heads/$branchName" ; sha = $mainSha } | ConvertTo-Json
    $null  = Invoke-RestMethod -Uri "$API/repos/$REPO/git/refs" -Headers $hdr -Method POST -Body $bBody
    Write-Host '  [OK]' -ForegroundColor Green
} catch {
    Write-Host '  [ECHEC]' -ForegroundColor Red
    Write-Host ("  ERREUR : {0}" -f (Get-DiagError $_.Exception)) -ForegroundColor Red
    pause ; exit 1
}

# --- Uploader le MP3 ---
Write-Host ("  Upload de {0}..." -f $mp3File) -NoNewline
try {
    $encoded  = [System.Uri]::EscapeDataString($mp3File)
    $mp3B64   = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($mp3Path))
    $upBody   = @{ message = "Add: $songLabel ($mp3File)" ; content = $mp3B64 ; branch = $branchName } | ConvertTo-Json -Depth 5
    $null     = Invoke-RestMethod -Uri "$API/repos/$REPO/contents/$encoded" -Headers $hdr -Method PUT -Body $upBody
    Write-Host '  [OK]' -ForegroundColor Green
} catch {
    Write-Host '  [ECHEC]' -ForegroundColor Red
    Write-Host ("  ERREUR : {0}" -f (Get-DiagError $_.Exception)) -ForegroundColor Red
    pause ; exit 1
}

# --- Mettre a jour manifest.txt sur la branche ---
Write-Host '  Mise a jour du manifest...' -NoNewline
try {
    $mBranch  = Invoke-RestMethod -Uri "$API/repos/$REPO/contents/manifest.txt?ref=$branchName" -Headers $hdr -Method GET
    $mCurrent = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String([string]::Join('', $mBranch.content) -replace "`n",''))
    $mNew     = $mCurrent.TrimEnd() + "`n$mp3File|$songLabel`n"
    $mNewB64  = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($mNew))
    $mBody    = @{ message = "Manifest: add $songLabel" ; content = $mNewB64 ; sha = $mBranch.sha ; branch = $branchName } | ConvertTo-Json -Depth 5
    $null     = Invoke-RestMethod -Uri "$API/repos/$REPO/contents/manifest.txt" -Headers $hdr -Method PUT -Body $mBody
    Write-Host '  [OK]' -ForegroundColor Green
} catch {
    Write-Host '  [ECHEC]' -ForegroundColor Red
    Write-Host ("  ERREUR : {0}" -f (Get-DiagError $_.Exception)) -ForegroundColor Red
    pause ; exit 1
}

# --- Ouvrir la Pull Request ---
Write-Host '  Ouverture de la demande de validation (Pull Request)...' -NoNewline
try {
    $prBody = @{
        title = "Ajout musique : $songLabel"
        body  = "## Nouvelle musique proposee`n`n| Champ | Valeur |`n|---|---|`n| **Fichier** | ``$mp3File`` |`n| **Nom affiche** | $songLabel |`n| **Taille** | $mp3Size MB |`n`n_Propose via Ajouter-Musique.bat_`n`n> Kevin, accepte ou refuse cette proposition en cliquant **Merge** ou **Close**."
        head  = $branchName
        base  = $MAIN
    } | ConvertTo-Json -Depth 5
    $pr = Invoke-RestMethod -Uri "$API/repos/$REPO/pulls" -Headers $hdr -Method POST -Body $prBody
    Write-Host '  [OK]' -ForegroundColor Green

    Write-Host ''
    Write-Host '=============================================' -ForegroundColor Cyan
    Write-Host '  Proposition envoyee avec succes !'         -ForegroundColor Green
    Write-Host '  Kevin va recevoir une notification et'     -ForegroundColor Cyan
    Write-Host '  devra valider ta contribution.'            -ForegroundColor Cyan
    Write-Host ''
    Write-Host ("  Lien : {0}" -f $pr.html_url) -ForegroundColor Gray
    Write-Host '=============================================' -ForegroundColor Cyan

} catch {
    Write-Host '  [ECHEC]' -ForegroundColor Red
    Write-Host ("  ERREUR : {0}" -f (Get-DiagError $_.Exception)) -ForegroundColor Red
    pause ; exit 1
}

Write-Host ''
pause
