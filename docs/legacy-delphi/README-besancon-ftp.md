# Tâche quotidienne Besançon / FTP / moyenne hebdo (analyse du code Delphi legacy)

Ce dossier contient le **code source Delphi de l'ancien Metrologo** (référence, non compilé).
Source principale : `F_Main.pas`. Ce document synthétise le mécanisme de la tâche
planifiée quotidienne (récupération FTP des valeurs corrigées de Besançon + calcul
de la moyenne hebdomadaire du rubidium), à réimplémenter dans la version WPF
(= TODO 5.3 « Marche hebdo / Correction Besançon »).

## Tableau de synthèse

| Étape | Quand | Action | Code / SQL |
|---|---|---|---|
| 1. Programmation | Au démarrage | Timer programmé à **10h00** (variante commentée = **9h50**), puis toutes les **24 h** | `tmrBesancon` (F_Main.pas ~349) |
| 2. Déclenchement | Chaque jour | Attend la ressource Mesure (ne perturbe pas une mesure en cours) | `tmrBesanconTimer` (~2939) |
| 3. Téléchargement FTP | À chaque déclenchement | Connexion **FTP E2M**, télécharge **`ef_utcop`** (mesures corrigées Observatoire de Besançon), puis le supprime du FTP | `ftpe2m`, `S_FIC_BESANCON` |
| 4. Sauvegarde | Après téléchargement | Copie datée dans `SavBesancon\yyyymmdd_hhnnss.txt` | `S_PATH_SAVEBESANCON` |
| 5. Intégration BDD | Pour chaque ligne | Parse `<date julienne> <valeur>` → insère la mesure journalière | proc `Metrologo_Add_Mesure_Journaliere_GPS` |
| 6. Moyenne hebdo | **Le mardi** | Exige **7 mesures** sur les 7 jours précédents, calcule `AVG(DAT_VALEUR)` → `SUV_ECARTF` + `SUV_DELTATPS` (= moy × 86400) | `GetMoyenneHebdo` (~3810), proc `Metrologo_Calcul_Ecart_Hebdo` |
| 7. Maj référence | Après calcul | **GPS** → maj fréq réf rubidium + rappel rubidiums non actifs · **Allouis** → rappel marche hebdo | `Metrologo_EnregFRef`, `tmrReminder` |

## Détails clés

- **Heure** : `dlHeureDeclench := IncHour(DateOf(Now), 10)` (10h00). Variante 9h50 :
  `IncMinute(DateOf(Now), 590)`. Intervalle initial = ms jusqu'à cette heure, puis 24 h.
- **Fichier FTP** : `ef_utcop` (constante `S_FIC_BESANCON` dans `U_DeclarationsMETROLOGO.pas`).
  Format : lignes `<dateJulienne> <valeur> …` séparées par des espaces. Date julienne =
  **Modified Julian Date** (`DateTimeToModifiedJulianDate`).
- **Moyenne hebdo** : sur les 7 jours précédant un mardi ; refuse le calcul si la base
  ne contient pas exactement 7 mesures. `SUV_ECARTF` = moyenne ; `SUV_DELTATPS` = moyenne × 86400.

## Tables SQL
- `T_METROLOGO_DATESRUBIS` (DAT_ID = date julienne, DAT_VALEUR, RUB_ACTIF)
- `TJ_METROLOGO_SUIVIRUBI` (SUV_ECARTF, SUV_DELTATPS, DAT_ID, RUB_ID)
- `TR_METROLOGO_RUBIDIUMS`

## Manque pour réimplémenter
- **Hôte / identifiants du FTP E2M** : dans le `.dfm` du formulaire (non fourni — seuls
  les `.pas` + `.ini` sont ici).
- **Schéma exact** des tables + procédures stockées SQL Server (à recréer côté WPF).
