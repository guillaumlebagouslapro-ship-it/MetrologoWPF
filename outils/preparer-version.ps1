param(
  [Parameter(Mandatory = $true)][string]$AppId,
  [Parameter(Mandatory = $true)][string]$Feed,
  [Parameter(Mandatory = $true)][string]$OutDir,
  # Fichier "MAJOR.MINOR" utilise seulement si le dossier reseau est vide.
  [string]$Repli = ''
)

# Choix du numero de version + description, appele par publier-*.bat.
#
# Numerotation MAJOR.MINOR (paquet Velopack : MAJOR.MINOR.0) :
#   Entree = MINOR + 1 (mise a jour courante), M = MAJOR + 1 (mise a jour importante),
#   ou un numero tape a la main.
# Garde-fous : pas de zero en tete (Update.exe refuse d'appliquer 1.261007.0919) et
# version strictement superieure a la derniere publiee (sinon les postes ne la voient
# jamais comme une mise a jour : 1.3 < 1.260706.1451).
#
# Ecrit <OutDir>\version.txt (MAJOR.MINOR) et, si une description est saisie,
# <OutDir>\notes.md (affichee dans Accueil > Mises a jour). Code retour 1 = abandon.

$ErrorActionPreference = 'Stop'

# --- Derniere version publiee (paquets reellement presents sur le reseau) ---
$max = $null
Get-ChildItem -LiteralPath $Feed -Filter '*.nupkg' -ErrorAction SilentlyContinue | ForEach-Object {
  if ($_.Name -match ('^' + [regex]::Escape($AppId) + '-(\d+\.\d+\.\d+)-')) {
    $v = [version]$Matches[1]
    if (-not $max -or $v -gt $max) { $max = $v }
  }
}

if ($max) {
  if ($max.Major -eq 1 -and $max.Minor -gt 1000) {
    # Ancienne numerotation datee (1.AAMMJJ.HHMM) : on passe a la numerotation simple.
    Write-Host "Derniere version publiee : $max (ancienne numerotation datee)"
    $mineure = '2.0'
    $majeure = '2.0'
    Write-Host "Passage a la numerotation simple : premiere version = 2.0"
  } else {
    Write-Host ("Derniere version publiee : {0}.{1}" -f $max.Major, $max.Minor)
    $mineure = '{0}.{1}' -f $max.Major, ($max.Minor + 1)
    $majeure = '{0}.0' -f ($max.Major + 1)
  }
} else {
  $base = '1.0'
  if ($Repli -and (Test-Path -LiteralPath $Repli)) { $base = (Get-Content -LiteralPath $Repli -Raw).Trim() }
  $p = $base -split '\.'
  $mineure = '{0}.{1}' -f [int]$p[0], ([int]$p[1] + 1)
  $majeure = '{0}.0' -f ([int]$p[0] + 1)
  Write-Host "ATTENTION : aucun paquet dans $Feed"
  Write-Host "Le numero propose vient de '$base' et peut etre trop BAS : il doit etre"
  Write-Host "SUPERIEUR a la version installee sur les postes, sinon pas de mise a jour."
}

Write-Host ''
Write-Host "  Entree = $mineure   (mise a jour courante)"
Write-Host "  M      = $majeure   (mise a jour importante)"
Write-Host "  ou tape un numero, ex $mineure"
Write-Host ''

while ($true) {
  $saisie = (Read-Host 'Version a publier').Trim()
  if ($saisie -eq '') { $ver = $mineure }
  elseif ($saisie -ieq 'M') { $ver = $majeure }
  else { $ver = $saisie -replace '\.0$', '' -replace '^(\d+)$', '$1.0' }

  if ($ver -notmatch '^(0|[1-9]\d*)\.(0|[1-9]\d*)$') {
    Write-Host "  Numero invalide : '$saisie'. Format MAJEUR.MINEUR, sans zero en tete (ex 2.4)."
    continue
  }
  if ($max -and ([version]"$ver.0") -le $max) {
    Write-Host "  $ver n'est pas superieur a la derniere version publiee ($max) :"
    Write-Host "  les postes ne la verraient jamais comme une mise a jour."
    continue
  }
  break
}

Write-Host ''
$notes = (Read-Host 'Ce qui change dans cette version (une ligne, Entree = rien)').Trim()

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$utf8 = New-Object System.Text.UTF8Encoding $false
[IO.File]::WriteAllText((Join-Path $OutDir 'version.txt'), $ver, $utf8)
if ($notes -ne '') { [IO.File]::WriteAllText((Join-Path $OutDir 'notes.md'), $notes, $utf8) }

Write-Host ''
Write-Host "Version retenue : $ver   (paquet $ver.0)"
if ($notes -ne '') { Write-Host "Description     : $notes" }
exit 0
