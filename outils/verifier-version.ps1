param(
  [string]$Version,
  [string]$Feed
)

# Verifie qu'une version peut etre publiee sur le canal Velopack :
#  - format X.Y.Z (3 nombres) ;
#  - strictement superieure a la derniere version deja publiee dans $Feed,
#    sinon les postes ne la verront jamais comme une mise a jour
#    (ex : 1.3.0 < 1.260706.1451, la comparaison se fait nombre par nombre).
#  - aucun zero en tete (SemVer strict, sinon Update.exe refuse d'appliquer).
# Code retour : 0 = OK, 2 = format invalide, 3 = version trop basse.

if ($Version -notmatch '^\d+\.\d+\.\d+$') {
  Write-Host "ERREUR : version '$Version' invalide. Il faut 3 nombres, ex 2.0.0"
  exit 2
}

# Zero en tete (ex 1.261007.0919) : publie sans erreur, mais Update.exe refuse de
# l'appliquer (SemVer strict) -> la MAJ se telecharge et echoue en boucle sur les postes.
if ($Version -notmatch '^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$') {
  Write-Host "ERREUR : version '$Version' invalide : pas de zero en tete d'un nombre (ex 0919 -> 919)."
  exit 2
}

$max = $null
Get-ChildItem -LiteralPath $Feed -Filter '*-full.nupkg' -ErrorAction SilentlyContinue | ForEach-Object {
  if ($_.Name -match '-(\d+\.\d+\.\d+)-full\.nupkg$') {
    $v = [version]$Matches[1]
    if (-not $max -or $v -gt $max) { $max = $v }
  }
}

if ($max -and [version]$Version -le $max) {
  Write-Host "ERREUR : version $Version trop basse. Derniere version publiee : $max"
  Write-Host "Les postes ne la verraient jamais comme une mise a jour."
  Write-Host "Accepte la version automatique (Entree) ou choisis un numero plus grand."
  exit 3
}

if ($max) { Write-Host "OK : $Version > derniere publiee ($max)" }
exit 0
