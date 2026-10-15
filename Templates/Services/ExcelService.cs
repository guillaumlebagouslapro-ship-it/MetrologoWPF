using ClosedXML.Excel;
using Metrologo.Models;
using Metrologo.Services.Catalogue;
using Metrologo.Services.Incertitude;
using Metrologo.Services.Journal;
using JournalLog = Metrologo.Services.Journal.Journal;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;

namespace Metrologo.Services
{
    public interface IExcelService
    {
        /// <summary>Crée la feuille de mesure dans le classeur du FI. gateIndexOverride remplace config.GateIndex
        /// dans les zones nommées (balayage stab) sans modifier la configuration.</summary>
        Task InitialiserRapportAsync(string numeroFI, Mesure configuration, Rubidium rubidium, int? gateIndexOverride = null, bool nouvelleSession = false);

        /// <summary>Ajoute les lignes de mesure manquantes (au-delà des 2 du template) avec leurs formules.
        /// HEURE et mesure restent vides : elles sont remplies en direct.</summary>
        Task PreparerLignesMesureAsync(int nbMesures);

        /// <summary>Enregistre le classeur ClosedXML sur disque, sans Excel, pour qu'ExcelInteropHost le reprenne en écriture live.</summary>
        Task<string> SauvegarderSurDisqueAsync();

        /// <summary>Nom de la feuille créée pour cette mesure (ex: Freq1, Stab1).</summary>
        string NomFeuilleMesure { get; }

        /// <summary>Chemin attendu du fichier de mesure pour une FI et un type donnés, sans effet de bord.
        /// Sert à vérifier que le classeur Excel déjà ouvert correspond à la même FI.</summary>
        string CalculerCheminFichierAttendu(Mesure mesure);

        /// <summary>Supprime une feuille de mesure (et ses lignes Récap.) dans le .xlsx sur disque, via ClosedXML.
        /// Utilisé après un STOP, quand Excel a été tué. Best-effort : réessaie et ne lève jamais.</summary>
        void SupprimerFeuilleSurDisque(string cheminFichier, string nomFeuille);

        /// <summary>Inscrit la moyenne et la variance dans les zones nommées en fin de boucle (le Récap. en dépend).</summary>
        Task EcrireStatsAsync(List<double> resultats);

        /// <summary>Écrit les N mesures (HEURE + valeur) via ClosedXML, sans Interop. Mode invisible de la Stabilité :
        /// Excel n'est jamais ouvert, le fichier n'apparaît qu'à la fin du balayage.</summary>
        Task EcrireValeursBatchClosedXMLAsync(int ligneDebut, IList<(DateTime ts, double valeur)> mesures);

        Task MettreAJourRecapFreqAsync(Mesure mesure);
        Task MettreAJourRecapStabAsync(Mesure mesure);

        /// <summary>Recopie le fichier de mesure vers le chemin réseau configuré (CheminsMetrologo.MesuresLocal),
        /// sans workbook ClosedXML ouvert. Appelé après la finalisation COM quand SauvegarderFinalAsync ne l'est pas.</summary>
        Task DupliquerSurReseauAsync();

        /// <summary>Idem avec le chemin source explicite : indispensable pour la voie COM pure, où _cheminFichier n'est jamais initialisé.</summary>
        Task DupliquerSurReseauAsync(string cheminSourceExplicite);

        /// <summary>Rouvre depuis le disque le classeur rempli en live par l'Interop. Ne crée aucune feuille :
        /// récupère celle nommée NomFeuilleMesure.</summary>
        Task RouvrirClasseurAsync();

        /// <summary>Sauvegarde finale une fois le Récap et les stats posés. Le fichier reste ouvert côté ExcelInteropHost.</summary>
        Task SauvegarderFinalAsync();

        void FermerExcel();
    }

    public class ExcelService : IExcelService
    {
        // Feuille modèle, jamais modifiée (copiée pour chaque mesure).
        private const string NOM_MODELE = "ModFeuille";
        private const string NOM_RECAP = "Récap.";

        // Ligne où commencent les mesures dans le template.
        private const int LIGNE_DEBUT_MESURES = 9;

        private XLWorkbook? _workbook;
        private IXLWorksheet? _feuilleMesure;   // la feuille créée pour cette mesure
        private string _cheminFichier = string.Empty;
        /// <summary>Fichier de la stab en cours (stab, stab1…), fixé à sa 1re gate et réutilisé par les suivantes.</summary>
        private string _cheminFichierStabSession = string.Empty;
        private string _nomFeuilleMesure = string.Empty;
        private TypeMesure _typeMesureCourant;  // pour PreparerLignesMesureAsync (col K Tachy)

        // Module d'incertitude retenu à l'init : EcrireStatsAsync y cherche CoeffA/CoeffB
        // (ModulesIncertitudeService) selon la moyenne. Vide : valeurs codées en dur de l'init.
        private string _numModuleIncertitudeCourant = string.Empty;

        // Module Fréquence auxiliaire (tachymètre) : remplit ZNCoeffA/B (côté Hz), le module
        // principal remplit ZNCoeffC/D (côté RPM). Vide : A/B gardent les valeurs du template.
        private string _numModuleIncertitudeFreqCourant = string.Empty;
        private double _tempsGateSecondesCourant;

        // Permettent de refaire en C# la conversion Indirect de la formule de la colonne C (Fréq. Réelle),
        // pour que la moyenne passée à ObtenirCoefficients colle à l'AVERAGE(C) d'Excel.
        private int _indexMultiplicateurCourant;
        private double _fNominaleCourant;
        private ModeMesure _modeMesureCourant;

        // Sigmas relatifs par gate (clé = nom de feuille en int). Remplis par EcrireStatsAsync, relus par
        // SauvegarderSurDisqueAsync pour calibrer l'axe Y log du graphe Stab (Excel auto-scale sur 1E-9..1E0
        // alors que les données sont vers 1E-8..1E-6).
        private readonly Dictionary<int, double> _sigmasRelatifsParGate = new();

        /// <summary>Colonne de conversion Hz → tr/min (tachy/strobo), à droite du bloc A-D pour ne pas gêner
        /// la Récap, qui lit via les zones nommées.</summary>
        private const string COL_CONVERSION_TR_MIN = "K";

        public string NomFeuilleMesure => _nomFeuilleMesure;

        /// <summary>Chemin du fichier Excel finalement utilisé (stab1.xlsx, stab2.xlsx… en stabilité itérée), affiché dans l'UI.</summary>
        public string CheminFichierGenere => _cheminFichier;

        /// <summary>Refait le calcul de chemin d'InitialiserRapportAsync (dossier de la FI — réseau si joignable, sinon
        /// Bureau\Metrologo\FI — puis freq.xlsx ou stab.xlsx) pour savoir si le classeur ouvert est réutilisable.
        /// Peut tester le réseau : hors thread UI de préférence. Les noms freq/stab sont imposés par le logiciel tiers d'extraction.</summary>
        public string CalculerCheminFichierAttendu(Mesure mesure)
        {
            string dossier = CheminsMetrologo.DossierFI(mesure.NumFI ?? string.Empty);
            bool estStab = mesure.TypeMesure == TypeMesure.Stabilite;
            string baseNom = estStab ? "stab" : "freq";
            return Path.Combine(dossier, $"{baseNom}.xlsx");
        }

        /// <summary>Vrai quand on a dû écrire le fichier sous un nom de repli plutôt que sous son nom principal.</summary>
        public bool FallbackTimestampUtilise { get; private set; }

        public async Task InitialiserRapportAsync(string numeroFI, Mesure config, Rubidium rubidium, int? gateIndexOverride = null, bool nouvelleSession = false)
        {
            int gateInscrite = gateIndexOverride ?? config.GateIndex;
            bool estStab = config.TypeMesure == TypeMesure.Stabilite;
            bool estTachy = EnTetesMesureHelper.EstTachymetre(config.TypeMesure);
            bool estInterval = config.TypeMesure == TypeMesure.Interval;

            await Task.Run(() =>
            {
                // --- 1. Dossier et fichier ---
                // Le rapport vit sur le partage réseau (Admin > Chemins d'accès > Mesures) s'il est joignable,
                // sinon dans Bureau\Metrologo\FI, rapatrié plus tard (CheminsMetrologo.DossierFI).
                string dossier = CheminsMetrologo.DossierFI(numeroFI);
                Directory.CreateDirectory(dossier);

                // Noms imposés par le logiciel tiers d'extraction : stab.xlsx pour la Stabilité, freq.xlsx
                // pour tout le reste (fréquence ET tachy, jamais mélangées sur une même FI). Le FI est porté par le dossier.
                string baseNom = estStab ? "stab" : "freq";
                _cheminFichier = Path.Combine(dossier, $"{baseNom}.xlsx");
                FallbackTimestampUtilise = false;

                // Stabilité, nouvelle session et fichier existant : on passe à stab1, stab2… pour que le
                // graphe Stab ne mélange pas deux sessions. La fréquence réutilise toujours freq.xlsx.
                if (estStab && nouvelleSession && File.Exists(_cheminFichier))
                {
                    string baseSansExt = Path.Combine(dossier, "stab");
                    int v = 1;
                    string candidat;
                    do
                    {
                        candidat = $"{baseSansExt}{v}.xlsx";
                        v++;
                    } while (File.Exists(candidat) && v < 1000);
                    _cheminFichier = candidat;
                }

                // Le fichier choisi à la 1re gate est mémorisé pour toute la stab, sinon les gates 2+
                // (nouvelleSession = false) retomberaient sur « stab.xlsx » au lieu de stab1/stab2…
                if (estStab)
                {
                    if (nouvelleSession)
                        _cheminFichierStabSession = _cheminFichier;
                    else if (!string.IsNullOrEmpty(_cheminFichierStabSession))
                        _cheminFichier = _cheminFichierStabSession;
                }

                // --- 2. Ouverture : fichier existant ou copie du template ---
                // Un template par famille : Stab (Récap. 8 cols + zones ZNRecapS_*), Tachy pour Contact/Optique
                // (col I = stats en tr/min, zones ZNCoeffC/D et ZNIncertResolRpm), Interval, et Fréquence pour le reste.
                string nomTemplate =
                    estStab ? "METROLOGO_Stab.xlsx" :
                    estTachy ? "METROLOGO_Tachy.xlsx" :
                    estInterval ? "METROLOGO_Interval.xlsx" :
                    "METROLOGO.xlsx";
                string templatePath = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory, "Templates", nomTemplate);

                bool partirDuTemplate = false;
                if (File.Exists(_cheminFichier))
                {
                    if (FichierEstVerrouille(_cheminFichier))
                    {
                        // Verrou probablement dû à notre propre Excel COM qui n'a pas encore rendu son handle :
                        // on attend jusqu'à 5 s, puis erreur explicite (pas de sous-fichier horodaté, tout doit rester dans le même fichier).
                        var sw = System.Diagnostics.Stopwatch.StartNew();
                        while (FichierEstVerrouille(_cheminFichier) && sw.ElapsedMilliseconds < 5000)
                        {
                            System.Threading.Thread.Sleep(200);
                        }

                        if (FichierEstVerrouille(_cheminFichier))
                        {
                            throw new InvalidOperationException(
                                $"Le fichier « {Path.GetFileName(_cheminFichier)} » est verrouillé "
                                + "(probablement ouvert dans une autre fenêtre Excel). "
                                + "Ferme-le manuellement et relance la mesure.");
                        }
                    }

                    _workbook = new XLWorkbook(_cheminFichier);

                    // Garde-fou « famille de template » : la ModFeuille de freq.xlsx hérite du template qui a CRÉÉ
                    // le fichier. Mélanger tachy et fréquence donnerait un rapport bancal (zones ZNCoeffC/D absentes,
                    // colonne tr/min manquante, nom du module décalé). La famille se repère à la zone ZNCoeffC,
                    // présente seulement dans le template tachy ; on refuse le mélange.
                    if (!estStab && _workbook.Worksheets.Any(w => w.Name == NOM_MODELE))
                    {
                        var modExistant = _workbook.Worksheet(NOM_MODELE);
                        bool fichierFamilleTachy =
                            modExistant.DefinedNames.Contains("ZNCoeffC")
                            || _workbook.DefinedNames.Contains("ZNCoeffC");
                        if (fichierFamilleTachy != estTachy)
                        {
                            string familleFichier = fichierFamilleTachy
                                ? "tachymètre/stroboscope" : "fréquence";
                            string familleMesure = estTachy
                                ? "tachymètre/stroboscope" : "fréquence";
                            throw new InvalidOperationException(
                                $"Le fichier « {Path.GetFileName(_cheminFichier)} » de cette FI a été créé "
                              + $"pour des mesures de type {familleFichier} : impossible d'y ajouter "
                              + $"une mesure de type {familleMesure}.\n\n"
                              + "Une même FI ne peut pas mélanger ces deux familles de mesures "
                              + "(mise en page Excel différente). Utilise une autre FI, ou "
                              + "supprime/déplace ce fichier depuis le dossier de la FI puis relance.");
                        }
                    }
                }
                else
                {
                    _workbook = new XLWorkbook(templatePath);
                    partirDuTemplate = true;
                }

                // --- 2b. Nettoyage des feuilles "1".."10" héritées du Stab1.xls historique (Delphi) ---
                // Le template Stab contient 10 feuilles vides ; sans nettoyage, les nouvelles gates seraient
                // nommées 11, 12… et la Récap. aurait 10 lignes parasites. Uniquement sur un fichier neuf
                // (issu du template), jamais sur un fichier existant (mesures déjà faites).
                if (partirDuTemplate && estStab)
                {
                    NettoyerFeuillesNumeriquesResiduellesEtRecap();
                    // Nouvelle session Stab : on repart de zéro sur les sigmas (calibration de l'axe Y).
                    _sigmasRelatifsParGate.Clear();
                }

                var modFeuille = _workbook.Worksheet(NOM_MODELE);

                // --- 3. Création d'une nouvelle feuille (copie de ModFeuille) ---
                _nomFeuilleMesure = TrouverNomFeuilleUnique(config.TypeMesure);

                // Feuilles à nom fixe (F_avant_interv / F_finale) : en cas de relance sur la même FI, on ÉCRASE
                // l'ancienne feuille et sa ligne Récap (sinon doublon / erreur au CopyTo). No-op pour les types numérotés.
                SupprimerFeuilleEtLigneRecapSiExiste(_nomFeuilleMesure);

                _feuilleMesure = modFeuille.CopyTo(_nomFeuilleMesure);

                // ModFeuille est un modèle interne : on la masque (elle reste dans le classeur pour les copies suivantes).
                // La copie peut hériter de Hidden : on rend la nouvelle feuille visible et active
                // (Excel n'accepte pas un onglet actif masqué).
                _feuilleMesure.Visibility = XLWorksheetVisibility.Visible;
                foreach (var ws in _workbook.Worksheets) ws.TabSelected = false;
                _feuilleMesure.SetTabActive();
                _feuilleMesure.TabSelected = true;
                modFeuille.Visibility = XLWorksheetVisibility.Hidden;

                DeprotegerFeuille(_feuilleMesure);

                // Le tableau des mesures commence en colonne A (HEURE=A, Mesurée=B, Réelle=C, Delta=D).
                _feuilleMesure.Column("B").Width = 30;   // Mesurée
                _feuilleMesure.Column("C").Width = 28;   // Réelle
                _feuilleMesure.Column("E").Width = 30;   // bloc labels droite

                // --- 4. Clonage des zones nommées en sheet-scope sur la nouvelle feuille ---
                ClonerZonesNommeesPourNouvelleFeuille(modFeuille, _feuilleMesure);

                // --- 5. En-têtes adaptés au type de mesure (labels de lignes en col B, valeurs en col C via formules/zones) ---
                var entetes = EnTetesMesureHelper.Pour(config.TypeMesure);
                _feuilleMesure.Cell("A7").SetValue(entetes.EnteteHeure);
                _feuilleMesure.Cell("B7").SetValue(entetes.EnteteMesuree);
                _feuilleMesure.Cell("C7").SetValue(entetes.EnteteReelle);
                _feuilleMesure.Cell("D7").SetValue(entetes.EnteteDelta);
                _feuilleMesure.Cell("B13").SetValue(entetes.LabelMoyenne);
                _feuilleMesure.Cell("B21").SetValue(entetes.LabelFreqRef);
                _feuilleMesure.Cell("B23").SetValue(entetes.LabelFreqCorr);
                _feuilleMesure.Cell("B25").SetValue(entetes.LabelIncertResol);
                _feuilleMesure.Cell("B31").SetValue(entetes.LabelIncertGlob);

                // --- 5b. Module d'incertitude affiché une seule fois ---
                // Module choisi (celui qui alimente les coefficients) à côté du label « Module = » :
                // F10 en standard, G10 en tachy. Le Récap. lit cette cellule.
                string moduleSelectionne = config.NumModuleIncertitude ?? string.Empty;
                string celluleModule = estTachy ? "G10" : "F10";
                if (!string.IsNullOrEmpty(moduleSelectionne))
                    _feuilleMesure.Cell(celluleModule).SetValue(moduleSelectionne);

                // Mémorisé pour PreparerLignesMesureAsync (col K conversion tr/min).
                _typeMesureCourant = config.TypeMesure;

                // Mémorisé pour EcrireStatsAsync, qui appelle ObtenirCoefficients (fonction, temps de gate,
                // fréquence moyenne) une fois la moyenne connue en fin de boucle.
                _numModuleIncertitudeCourant = config.NumModuleIncertitude ?? string.Empty;
                _numModuleIncertitudeFreqCourant = config.NumModuleIncertitudeFreq ?? string.Empty;
                _tempsGateSecondesCourant = EnTetesMesureHelper.SecondesGate(gateInscrite);
                _indexMultiplicateurCourant = config.IndexMultiplicateur;
                _fNominaleCourant = config.FNominale;
                _modeMesureCourant = config.ModeMesure;

                // --- 5c. Tachymétrie : colonne K dédiée à la conversion Hz → tr/min ---
                // Le compteur GPIB donne des Hz (mesure brute en col B). tr/min = Hz × 60 (1 imp/tour)
                // est calculé par Excel pour rester visible de l'utilisateur.
                if (EnTetesMesureHelper.EstTachymetre(config.TypeMesure))
                {
                    _feuilleMesure.Cell($"{COL_CONVERSION_TR_MIN}7").SetValue("Vitesse (tr/min)");
                    _feuilleMesure.Column(COL_CONVERSION_TR_MIN).Width = 18;
                    // Lignes 9 et 10 du template ; les suivantes sont remplies par PreparerLignesMesureAsync.
                    _feuilleMesure.Cell($"{COL_CONVERSION_TR_MIN}9").FormulaA1 = "=B9*60";
                    _feuilleMesure.Cell($"{COL_CONVERSION_TR_MIN}10").FormulaA1 = "=B10*60";
                }

                // --- 6. Métadonnées via zones nommées (sheet-scope) ---
                SetNamed("ZNNoFiche", numeroFI);
                SetNamed("ZNDate", DateTime.Now.ToString("dd/MM/yyyy"));
                SetNamed("ZNTypeMesure", EnTetesMesureHelper.LibelleType(config.TypeMesure));
                SetNamed("ZNFreqUtilise", NomAppareilDepuisCatalogue(config.IdModeleCatalogue));
                SetNamed("ZNRubidium", rubidium.DesignationAvecRaccord);
                SetNamed("ZNGate", EnTetesMesureHelper.LibelleGate(gateInscrite));
                SetNamed("ZNLibGate", EnTetesMesureHelper.LibelleGate(gateInscrite));
                SetNamed("ZNValGateSecondes", EnTetesMesureHelper.SecondesGate(gateInscrite));
                SetNamed("ZNModeMesure", config.ModeMesure == ModeMesure.Direct ? "Direct" : "Indirect");
                SetNamed("ZNCoeffMult", config.IndexMultiplicateur);
                SetNamed("ZNValFNominale", config.FNominale);
                SetNamed("ZNNbMesures", config.NbMesures);
                // Résolution et incertitude relative = 0 à l'init. Les vraies valeurs viennent des popups post-mesure
                // (AcceuilViewModel.SaisiePostMesureFrequenceAsync / SaisiePostMesureTachyAsync) via EcrireZoneNommeeAsync ;
                // les types sans popup (FreqAvantInterv, FreqFinale, Stabilité, Generateur…) restent à 0.
                // Ne pas seeder depuis config.IncertSupp : config est réutilisé entre relances sur une même FI
                // et la valeur saisie réapparaîtrait sur la mesure suivante.
                SetNamed("ZNIncertResol", 0.0);
                SetNamed("ZNIncertSup", 0.0);

                // Fréquence de référence (rubidium actif) écrite directement en C21 (zone ZNFreqRef), sans SetNamed :
                // sur une feuille fraîchement copiée, la zone nommée pouvait se résoudre sur le scope workbook
                // (ModFeuille cachée) et laisser C21 vide.
                var celluleFreqRef = _feuilleMesure.Cell("C21");
                celluleFreqRef.SetValue(rubidium.FrequenceMoyenne);
                // 7 décimales = maximum fidèle pour ~10 MHz : Excel (comme un double) ne garde que 15 chiffres
                // significatifs (8 entiers + 7 décimales) ; au-delà il afficherait des « 00 » parasites.
                celluleFreqRef.Style.NumberFormat.Format = "#,##0.0000000";
                SetNamed("ZNFreqRef", rubidium.FrequenceMoyenne);

                // Valeurs par défaut, conservées si aucun module d'incertitude n'est sélectionné ou si la moyenne
                // sort de ses plages. Sinon EcrireStatsAsync les surcharge en fin de boucle.
                SetNamed("ZNCoeffA", 1e-10);
                SetNamed("ZNCoeffB", 5e-13);
                SetNamed("ZNNbMesAccredite", 30);
                SetNamed("ZNTempsMesureAccredite", 10);

                // --- 7. Nettoyage anticipé des lignes « fantômes » du Récap. ---
                // Le template contient des lignes pré-remplies pointant vers ModFeuille (=[0]!ZNxxx, évaluées à 0).
                // On les purge dès l'init (et non à l'écriture du récap final, cf. EcrireLigneRecap) pour ne pas
                // montrer ~10 lignes à zéro à qui ouvre le fichier pendant la mesure.
                if (_workbook.Worksheets.Any(w => w.Name == NOM_RECAP))
                {
                    var recap = _workbook.Worksheet(NOM_RECAP);
                    DeprotegerFeuille(recap);
                    NettoyerLignesGhost(recap);
                }
            });
        }

        public async Task PreparerLignesMesureAsync(int nbMesures)
        {
            if (_feuilleMesure == null || _workbook == null || nbMesures <= 0) return;

            await Task.Run(() =>
            {
                // Le template a 2 lignes de mesures (9 et 10), les labels commencent en 13 :
                // les mesures supplémentaires sont insérées AVANT ZNPointInsertion.
                if (nbMesures > 2)
                {
                    int pointInsertionRow = TrouverLigneZone("ZNPointInsertion") ?? 11;
                    _feuilleMesure.Row(pointInsertionRow).InsertRowsAbove(nbMesures - 2);
                }

                // Formules Fréq. Réelle (col C) et delta (col D) pour les lignes insérées. HEURE et mesure restent
                // vides : ExcelInteropHost les écrit en direct pendant la boucle de mesures.
                bool conversionTrMin = EnTetesMesureHelper.EstTachymetre(_typeMesureCourant);
                for (int i = 2; i < nbMesures; i++)
                {
                    int row = LIGNE_DEBUT_MESURES + i;
                    // B = mesurée (GPIB), C = réelle (formule de correction), D = delta entre 2 mesures consécutives.
                    _feuilleMesure.Cell($"C{row}").FormulaA1 =
                        $"IF(ISBLANK(ZNCoeffMult),B{row},"
                        + $"(((B{row}-10000000)/(POWER(10,ZNCoeffMult)*10000000))+1)*ZNValFNominale)";
                    _feuilleMesure.Cell($"D{row}").FormulaA1 = $"C{row - 1}-C{row}";
                    if (conversionTrMin)
                        _feuilleMesure.Cell($"{COL_CONVERSION_TR_MIN}{row}").FormulaA1 = $"=B{row}*60";
                }

                // Formules métier historiques (héritage Stab1.xls) des statistiques. Nécessaires car la ModFeuille
                // du template Stab n'en a aucune ; pour Freq elles existent déjà, la réécriture est idempotente.
                EcrireFormulesStatsMetier(nbMesures);
            });
        }

        /// <summary>Pose les formules métier historiques : ZNFreqMoyReel (AVERAGE), ZNVariance (macro VBA
        /// Cal_variance, pas la variance n-1 classique) et les stats dérivées absentes du template Stab.
        /// Cellules cibles résolues par zones nommées sheet-scope.</summary>
        private void EcrireFormulesStatsMetier(int nbMesures)
        {
            if (_feuilleMesure == null) return;
            int ligneDeb = LIGNE_DEBUT_MESURES;
            int ligneFin = LIGNE_DEBUT_MESURES + nbMesures - 1;

            void EcrireSurZN(string nomZone, string formule)
            {
                IXLCell? cell = null;
                try
                {
                    // Zone sheet-scope (clonée depuis ModFeuille).
                    if (_feuilleMesure!.DefinedNames.TryGetValue(nomZone, out var defName))
                    {
                        cell = defName?.Ranges.FirstOrDefault()?.FirstCell();
                    }
                }
                catch { /* zone absente sur la feuille, silencieux */ }
                if (cell != null) cell.FormulaA1 = formule;
            }

            // Moyenne sur la colonne C (Fréq. Réelle).
            EcrireSurZN("ZNFreqMoyReel",
                $"IF(ISBLANK(ZNNbMesures),,AVERAGE(C{ligneDeb}:C{ligneFin}))");

            // Variance via la macro VBA Metrologo.xla. Les deltas commencent en D{deb+1} (pas de mesure précédente pour la 1re).
            EcrireSurZN("ZNVariance",
                $"[1]!Cal_variance(SUMSQ(D{ligneDeb + 1}:D{ligneFin}),ZNNbMesures,ZNFreqMoyReel)");

            // Stats dérivées : idempotent pour Freq (déjà dans ModFeuille), créées pour Stab.
            EcrireSurZN("ZNEcartType",
                "IF(ISBLANK(ZNVariance),,[1]!Cal_ecart_type(ZNVariance))");
            EcrireSurZN("ZNIncertEcartType",
                "IF(ISBLANK(ZNNbMesures),,[1]!Cal_incert_ecart_type(ZNVariance,ZNNbMesures))");
            EcrireSurZN("ZNFreqCorr",
                "IF(ISBLANK(ZNFreqMoyReel),,[1]!Cal_freq_corrigee(ZNFreqMoyReel,ZNFreqRef))");
        }

        public Task EcrireValeursBatchClosedXMLAsync(int ligneDebut, IList<(DateTime ts, double valeur)> mesures)
        {
            if (_feuilleMesure == null || mesures.Count == 0) return Task.CompletedTask;
            return Task.Run(() =>
            {
                for (int i = 0; i < mesures.Count; i++)
                {
                    int row = ligneDebut + i;
                    // A = HEURE, B = mesure brute.
                    _feuilleMesure.Cell(row, 1).SetValue(mesures[i].ts.ToString("HH:mm:ss"));
                    _feuilleMesure.Cell(row, 2).SetValue(mesures[i].valeur);
                }
            });
        }

        public async Task EcrireStatsAsync(List<double> resultats)
        {
            if (_feuilleMesure == null || _workbook == null || resultats.Count == 0) return;

            await Task.Run(() =>
            {
                if (resultats.Count == 0) return;

                // ZNFreqMoyReel et ZNVariance ne sont pas écrits depuis C# : leurs formules (AVERAGE / Cal_variance VBA)
                // sont posées par EcrireFormulesStatsMetier et calculées par Excel à l'ouverture
                // (fullCalcOnLoad="1", cf. ForcerRecalculAuOuverture).

                // Sigma relatif (n-1 classique, pas la formule métier Cal_variance) : seul l'ordre de grandeur compte,
                // pour fixer min/max de l'axe Y log du graphe Stab à ±0.5 décade autour des données.
                if (_typeMesureCourant == TypeMesure.Stabilite
                    && resultats.Count >= 2
                    && TryExtraireNumeroGate(_nomFeuilleMesure, out int numGate))
                {
                    double moy = resultats.Average();
                    double sumSq = 0;
                    foreach (var v in resultats) sumSq += (v - moy) * (v - moy);
                    double sigma = Math.Sqrt(sumSq / (resultats.Count - 1));
                    if (moy > 0 && sigma > 0)
                    {
                        _sigmasRelatifsParGate[numGate] = sigma / moy;
                    }
                }

                // --- Coefficients d'incertitude depuis le module sélectionné ---
                // Le choix de la ligne du module CSV exige la fréquence moyenne RÉELLE. ClosedXML ne calcule pas
                // les formules : on la recalcule en C# (ConvertirEnFreqReelle), la conversion affine commutant avec la moyenne.
                // Sans module sélectionné, les valeurs par défaut posées à l'init restent en place.
                if (!string.IsNullOrEmpty(_numModuleIncertitudeCourant))
                {
                    double moyenneBrute = resultats.Average();
                    double moyenneReelle = ConvertirEnFreqReelle(moyenneBrute);

                    // Types en tr/min (tachy contact/optique + strobo) : bornes du CSV saisies en tr/min côté admin,
                    // donc moyenne Hz × 60 avant le lookup. Les autres types restent en Hz.
                    bool uniteRpm = EnTetesMesureHelper.EstUniteRpm(_typeMesureCourant);
                    double valeurLookup = uniteRpm ? moyenneReelle * 60.0 : moyenneReelle;

                    string fonction = IncertitudeFonctionHelper.NomFonction(_typeMesureCourant);
                    var (coeffA, coeffB) = ModulesIncertitudeService.ObtenirCoefficients(
                        _numModuleIncertitudeCourant, _typeMesureCourant, fonction,
                        _tempsGateSecondesCourant, valeurLookup);

                    if (coeffA > 0 || coeffB > 0)
                    {
                        // Tachymètres : les coeffs A/B du CSV sont les coeffs C/D du template tachy
                        // (formule I29 = Vitesse_RPM × C + D), injectés dans ZNCoeffC/ZNCoeffD.
                        if (EnTetesMesureHelper.EstTachymetre(_typeMesureCourant))
                        {
                            SetNamed("ZNCoeffC", coeffA);
                            SetNamed("ZNCoeffD", coeffB);
                            JournalLog.Info(CategorieLog.Mesure, "INCERT_COEFFS_RESOLUS_TACHY",
                                $"Module {_numModuleIncertitudeCourant} : CoeffC={coeffA:G6} CoeffD={coeffB:G6} "
                              + $"(fonction={fonction}, gate={_tempsGateSecondesCourant}s, valeurLookup={valeurLookup:G6} tr/min) "
                              + "[A/B du CSV → C/D du template tachy, lookup en tr/min]");
                        }
                        else
                        {
                            SetNamed("ZNCoeffA", coeffA);
                            SetNamed("ZNCoeffB", coeffB);
                            string uniteLookup = uniteRpm ? "tr/min" : "Hz";
                            JournalLog.Info(CategorieLog.Mesure, "INCERT_COEFFS_RESOLUS",
                                $"Module {_numModuleIncertitudeCourant} : CoeffA={coeffA:G6} CoeffB={coeffB:G6} "
                              + $"(fonction={fonction}, gate={_tempsGateSecondesCourant}s, valeurLookup={valeurLookup:G6} {uniteLookup})");
                        }
                    }
                    else
                    {
                        string uniteLookup = uniteRpm ? "tr/min" : "Hz";
                        JournalLog.Warn(CategorieLog.Mesure, "INCERT_NO_MATCH",
                            $"Module {_numModuleIncertitudeCourant} : aucune ligne ne couvre "
                          + $"(fonction={fonction}, gate={_tempsGateSecondesCourant}s, valeurLookup={valeurLookup:G6} {uniteLookup}) "
                          + "— coefficients par défaut utilisés.");
                    }
                }

                // Tachymétrie : ZNCoeffA/B viennent aussi du module Fréquence auxiliaire (fréquencemètre support) ;
                // le module tachy a déjà rempli ZNCoeffC/D ci-dessus.
                if (EnTetesMesureHelper.EstTachymetre(_typeMesureCourant)
                    && !string.IsNullOrEmpty(_numModuleIncertitudeFreqCourant))
                {
                    double moyenneBrute = resultats.Average();
                    double moyenneReelle = ConvertirEnFreqReelle(moyenneBrute);
                    string fonctionFreq = IncertitudeFonctionHelper.NomFonction(TypeMesure.Frequence);
                    var (coeffAFreq, coeffBFreq) = ModulesIncertitudeService.ObtenirCoefficients(
                        _numModuleIncertitudeFreqCourant, TypeMesure.Frequence, fonctionFreq,
                        _tempsGateSecondesCourant, moyenneReelle);

                    if (coeffAFreq > 0 || coeffBFreq > 0)
                    {
                        SetNamed("ZNCoeffA", coeffAFreq);
                        SetNamed("ZNCoeffB", coeffBFreq);
                        JournalLog.Info(CategorieLog.Mesure, "INCERT_COEFFS_AB_FREQ",
                            $"Module Freq {_numModuleIncertitudeFreqCourant} : CoeffA={coeffAFreq:G6} CoeffB={coeffBFreq:G6} "
                          + $"(gate={_tempsGateSecondesCourant}s, freq={moyenneReelle:G6} Hz) "
                          + "[module Fréquence auxiliaire pour tachymètre]");
                    }
                    else
                    {
                        JournalLog.Warn(CategorieLog.Mesure, "INCERT_NO_MATCH_FREQ",
                            $"Module Freq {_numModuleIncertitudeFreqCourant} : aucune ligne ne couvre "
                          + $"(gate={_tempsGateSecondesCourant}s, freq={moyenneReelle:G6} Hz) "
                          + "— ZNCoeffA/B restent aux valeurs par défaut.");
                    }
                }
            });
        }

        /// <summary>Reproduit en C# la conversion de la formule Fréq. Réelle (col C). Affine, donc commutative avec la moyenne :
        /// convertir la moyenne brute donne la moyenne d'Excel.</summary>
        private double ConvertirEnFreqReelle(double mesureBrute)
        {
            // Mode Direct = pas de conversion (Excel : ZNCoeffMult vide). Côté C# ZNCoeffMult est toujours
            // posé, d'où ce court-circuit.
            if (_modeMesureCourant == ModeMesure.Direct) return mesureBrute;

            return (((mesureBrute - 10_000_000.0)
                    / (Math.Pow(10, _indexMultiplicateurCourant) * 10_000_000.0)) + 1.0)
                   * _fNominaleCourant;
        }

        public async Task<string> SauvegarderSurDisqueAsync()
        {
            if (_workbook == null) return string.Empty;
            await Task.Run(() =>
            {
                // Compté AVANT SaveAs : ClosedXML libère ses ressources après sauvegarde.
                int nbGatesStab = CompterGatesStabPourGraphe();

                try { _workbook.SaveAs(_cheminFichier); }
                catch (IOException)
                {
                    throw new InvalidOperationException(
                        $"Impossible de sauvegarder « {Path.GetFileName(_cheminFichier)} » : "
                        + "le fichier est verrouillé. Fermez-le dans Excel et relancez la mesure.");
                }

                try { PatcherLienMacroXLA(_cheminFichier, Preferences.CheminMacroXLA); }
                catch { /* best-effort */ }

                try { ForcerRecalculAuOuverture(_cheminFichier); }
                catch { /* best-effort */ }

                // Graphe Stab : plages des séries adaptées au nombre de gates balayées, axe Y auto (retrait des
                // min/max fixes 1E-12/1E-08 hérités de Stab1.xls). Ne touche pas aux cellules ni aux zones nommées
                // de la Récap : CEAO/SDAO consomme les données, pas les graphes.
                if (nbGatesStab > 0)
                {
                    try { RendreGrapheStabDynamique(_cheminFichier, nbGatesStab); }
                    catch { /* best-effort */ }
                }

                // Vidage des caches du graphe hérités de Stab1.xls (10 valeurs Y figées à 0) : sinon Excel s'en sert
                // pour calibrer l'axe Y log (échelle énorme et vide). Avec fullCalcOnLoad, force un vrai recalcul.
                try { ViderCachesGraphe(_cheminFichier); }
                catch { /* best-effort */ }

                // Axe Y log : sans cache ni bornes Excel auto-scale vers 1E-9..1E0 ; on force min/max
                // à ~0.5 décade autour des sigmas relatifs estimés.
                if (nbGatesStab > 0 && _sigmasRelatifsParGate.Count > 0)
                {
                    try { CalibrerAxeYGrapheStab(_cheminFichier, _sigmasRelatifsParGate.Values); }
                    catch { /* best-effort */ }
                }
            });
            return _cheminFichier;
        }

        public async Task RouvrirClasseurAsync()
        {
            if (string.IsNullOrEmpty(_cheminFichier) || string.IsNullOrEmpty(_nomFeuilleMesure))
                throw new InvalidOperationException(
                    "RouvrirClasseurAsync appelée avant InitialiserRapportAsync — pas de classeur à rouvrir.");

            await Task.Run(() =>
            {
                _workbook?.Dispose();
                _workbook = new XLWorkbook(_cheminFichier);

                _feuilleMesure = _workbook.Worksheets
                    .FirstOrDefault(w => string.Equals(w.Name, _nomFeuilleMesure, StringComparison.OrdinalIgnoreCase));

                if (_feuilleMesure == null)
                    throw new InvalidOperationException(
                        $"Feuille « {_nomFeuilleMesure} » introuvable après réouverture de « {Path.GetFileName(_cheminFichier)} ».");
            });
        }

        public async Task SauvegarderFinalAsync()
        {
            if (_workbook == null) return;
            await Task.Run(() =>
            {
                int nbGatesStab = CompterGatesStabPourGraphe();

                try { _workbook.SaveAs(_cheminFichier); }
                catch (IOException)
                {
                    throw new InvalidOperationException(
                        $"Impossible de sauvegarder « {Path.GetFileName(_cheminFichier)} » : "
                        + "le fichier est verrouillé. Fermez-le dans Excel et relancez la mesure.");
                }

                try { PatcherLienMacroXLA(_cheminFichier, Preferences.CheminMacroXLA); }
                catch { /* best-effort */ }

                try { ForcerRecalculAuOuverture(_cheminFichier); }
                catch { /* best-effort */ }

                if (nbGatesStab > 0)
                {
                    try { RendreGrapheStabDynamique(_cheminFichier, nbGatesStab); }
                    catch { /* best-effort */ }
                }

                // Duplication best-effort vers le chemin local de sauvegarde si configuré.
                DupliquerVersCheminLocal(_cheminFichier);
            });
        }

        /// <summary>Appelle <see cref="DupliquerVersCheminLocal"/> depuis MesureOrchestrator après la voie COM
        /// (où SauvegarderFinalAsync n'est plus invoquée). Best-effort : ne lève jamais.</summary>
        public Task DupliquerSurReseauAsync() => Task.Run(() => DupliquerVersCheminLocal(_cheminFichier));

        public Task DupliquerSurReseauAsync(string cheminSourceExplicite)
            => Task.Run(() => DupliquerVersCheminLocal(cheminSourceExplicite));

        /// <summary>Copie le fichier vers <c>CheminsMetrologo.MesuresLocal</c> si configuré, sous <c>&lt;dest&gt;\&lt;FI&gt;\&lt;nomFichier&gt;</c>.
        /// Best-effort : un échec (réseau coupé, permissions) est journalisé sans lever, le fichier principal reste valide.</summary>
        private void DupliquerVersCheminLocal(string cheminSource)
        {
            if (!CheminsMetrologo.MesuresLocalConfigure) return;
            if (string.IsNullOrWhiteSpace(cheminSource) || !File.Exists(cheminSource)) return;

            try
            {
                string racineLocale = CheminsMetrologo.MesuresLocal;
                string nomFichier = Path.GetFileName(cheminSource);
                string nomFI = Path.GetFileName(Path.GetDirectoryName(cheminSource) ?? "");
                if (string.IsNullOrWhiteSpace(nomFI)) nomFI = "Mesures";

                string dossierCible = Path.Combine(racineLocale, nomFI);
                Directory.CreateDirectory(dossierCible);
                string cheminCible = Path.Combine(dossierCible, nomFichier);

                // Écrasement : le fichier source contient tout l'historique (feuilles freq1, freq2…),
                // la copie réseau doit en être le reflet exact.
                File.Copy(cheminSource, cheminCible, overwrite: true);
                JournalLog.Info(CategorieLog.Excel, "EXCEL_DUPLIQUE_LOCAL",
                    $"Rapport dupliqué (écrasement) : {cheminCible}");
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "EXCEL_DUPLIQUE_LOCAL_KO",
                    $"Duplication vers {CheminsMetrologo.MesuresLocal} échouée : {ex.Message}");
            }
        }

        /// <summary>Ajoute une ligne de récapitulatif Fréquence dans la feuille <c>Récap.</c>
        /// (portage de <c>TfrmMain.MajRecapFreq</c>, F_Main.pas:2421).</summary>
        public async Task MettreAJourRecapFreqAsync(Mesure mesure)
        {
            if (_workbook == null || _feuilleMesure == null) return;
            string nomFeuille = _feuilleMesure.Name;
            string celluleModule = EnTetesMesureHelper.EstTachymetre(mesure.TypeMesure) ? "G10" : "F10";

            await Task.Run(() => EcrireLigneRecap(
                new[]
                {
                    $"='{nomFeuille}'!ZNFreqMoyReel",       // Col 1 : fréquence moyenne
                    $"='{nomFeuille}'!ZNLibGate",           // Col 2 : temps de mesure (libellé gate)
                    $"='{nomFeuille}'!ZNFreqCorr",          // Col 3 : fréquence corrigée
                    $"='{nomFeuille}'!ZNEcartType",         // Col 4 : écart-type
                    // Col 5 : fréquence indiquée, vide ici. Générateur : « Géné. » via valeurDirecte. Fréquencemètre :
                    // la valeur de la pop-up post-mesure est écrite directement par ExcelInteropHost.EcrireFreqIndiqueeRecapAsync.
                    null,
                    $"='{nomFeuille}'!ZNIncertResol",       // Col 6 : incertitude de résolution
                    $"='{nomFeuille}'!ZNIncertSup",         // Col 7 : incertitude supplémentaire
                    $"='{nomFeuille}'!ZNIncertAccreditee",  // Col 8 : incertitude accréditée
                    $"='{nomFeuille}'!ZNIncertGlobale",     // Col 9 : incertitude globale
                    // Col 10 : fréquence finale (formule historique du Delphi) : si la fréquence indiquée (col E) est
                    // saisie, la corrigée (col C) est projetée selon E/C*A, sinon C est gardée. {N} = n° de ligne (EcrireLigneRecap).
                    "=IF(ISNUMBER(E{N}),(E{N}/C{N})*A{N},C{N})",
                    $"='{nomFeuille}'!{celluleModule}"      // Col 11 : n°Module (cellule « Module = » de la feuille)
                },
                colValeurDirecte: 5,
                valeurDirecte: mesure.SourceMesure == SourceMesure.Generateur
                    ? (object)"Géné."
                    : null,   // Fréquencemètre : col 5 remplie par la pop-up
                entetesColonne: new[]
                {
                    null, null, null, null, null, null, null, null, null, null,
                    "n°Module"    // Col 11 : en-tête ajouté s'il manque
                }));
        }

        /// <summary>Ajoute une ligne dans la <c>Récap.</c> du fichier Stabilité, selon la structure historique (Stab1.xls) :
        /// 8 colonnes (Temps, Fréq Moyenne, Écart type, Incertitude, Incert accréditée, Incert globale, Maxi, Mini),
        /// une ligne par gate (gate 1 → L6, gate 2 → L7…), cols 7-8 = moyenne ± incert. globale, NumFI + date en L1.</summary>
        public async Task MettreAJourRecapStabAsync(Mesure mesure)
        {
            if (_workbook == null || _feuilleMesure == null) return;
            string nomFeuille = _feuilleMesure.Name;

            await Task.Run(() =>
            {
                if (!_workbook.Worksheets.Any(w => w.Name == NOM_RECAP)) return;
                var recap = _workbook.Worksheet(NOM_RECAP);
                DeprotegerFeuille(recap);

                // Colonne D (date) élargie pour éviter « #### ».
                recap.Column(4).Width = 13;

                // En-tête fiche : NumFI + date (idempotent à chaque gate).
                recap.Cell("B1").SetValue(mesure.NumFI);
                recap.Cell("D1").SetValue(DateTime.Now.ToString("dd/MM/yyyy"));

                // N° de la gate : les feuilles Stab sont numériques, on parse le nom ; sinon on compte les lignes remplies.
                int numero;
                if (!int.TryParse(nomFeuille, out numero))
                {
                    numero = CompterLignesStabRemplies(recap) + 1;
                }

                int ligne = 5 + numero;  // L6 = 1ère gate balayée, L7 = 2ème, etc.

                recap.Cell(ligne, 1).FormulaA1 = $"='{nomFeuille}'!ZNValGateSecondes";
                recap.Cell(ligne, 2).FormulaA1 = $"='{nomFeuille}'!ZNFreqMoyReel";
                recap.Cell(ligne, 3).FormulaA1 = $"='{nomFeuille}'!ZNEcartType";
                recap.Cell(ligne, 4).FormulaA1 = $"='{nomFeuille}'!ZNIncertEcartType";
                recap.Cell(ligne, 5).FormulaA1 = $"='{nomFeuille}'!ZNIncertAccreditee";
                recap.Cell(ligne, 6).FormulaA1 = $"='{nomFeuille}'!ZNIncertGlobale";
                recap.Cell(ligne, 7).FormulaA1 = $"=C{ligne}+F{ligne}";
                // Valeurs Mini : formule d'origine de Stab1.xls, plancher 1E-12 quand sigma ≤ incertitude
                // (« * » en col I) : un 0 rendrait le point invisible sur l'axe log.
                recap.Cell(ligne, 8).FormulaA1 =
                    $"=IF(ISBLANK(C{ligne}),,IF((C{ligne}-F{ligne})<=0,0.000000000001,MAX(0.000000000001,C{ligne}-F{ligne})))";

                // Col K : n°Module, lu dans la cellule « Module = » de la feuille de mesure (F10 en stabilité).
                recap.Cell(ligne, 11).FormulaA1 = $"='{nomFeuille}'!F10";
                // En-tête en ligne 5, posé seulement s'il manque.
                if (recap.Cell(5, 11).IsEmpty()) recap.Cell(5, 11).SetValue("n°Module");
            });
        }

        /// <summary>Nombre de lignes de données remplies dans la zone Stab (L6+), colonne A (Temps) non vide.</summary>
        private static int CompterLignesStabRemplies(IXLWorksheet recap)
        {
            int compteur = 0;
            for (int row = 6; row <= 18; row++)
            {
                var cell = recap.Cell(row, 1);
                if (!cell.IsEmpty()) compteur++;
            }
            return compteur;
        }

        // Ligne d'en-tête des colonnes dans le template Récap. (METROLOGO.xlsx).
        private const int LIGNE_ENTETE_RECAP = 5;

        /// <summary>Ajoute à la suite des lignes existantes de la Récap. une ligne de formules cross-sheet
        /// vers la feuille de mesure.</summary>
        private void EcrireLigneRecap(
            string?[] formulesParColonne,
            int colValeurDirecte = -1,
            object? valeurDirecte = null,
            string?[]? entetesColonne = null)
        {
            if (_workbook == null) return;
            if (!_workbook.Worksheets.Any(w => w.Name == NOM_RECAP)) return;

            var recap = _workbook.Worksheet(NOM_RECAP);

            DeprotegerFeuille(recap);

            // Colonne D (date) élargie pour éviter « #### ».
            recap.Column(4).Width = 13;

            // 1. Retrait des lignes « fantômes » du template (=[0]!ZNxxx vers ModFeuille, plage de zéros sous l'en-tête).
            NettoyerLignesGhost(recap);

            // 2. Nouvelle ligne sous la dernière non vide (la plus récente en bas).
            int nouvelleLigne = LIGNE_ENTETE_RECAP + 1;
            while (!recap.Row(nouvelleLigne).IsEmpty()) nouvelleLigne++;

            // 3. Remplissage des colonnes
            for (int i = 0; i < formulesParColonne.Length; i++)
            {
                int col = i + 1;
                if (col == colValeurDirecte && valeurDirecte != null)
                {
                    recap.Cell(nouvelleLigne, col).SetValue(XLCellValue.FromObject(valeurDirecte));
                    continue;
                }
                var formule = formulesParColonne[i];
                if (!string.IsNullOrEmpty(formule))
                {
                    // {N} = n° de la ligne courante, pour les formules qui référencent la même ligne.
                    formule = formule.Replace("{N}", nouvelleLigne.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    recap.Cell(nouvelleLigne, col).FormulaA1 = formule;
                }
            }

            // 4. En-têtes des colonnes ajoutées au-delà du template (ex: n°Module), posés seulement s'ils manquent :
            //    évite de modifier le template Excel.
            if (entetesColonne != null)
            {
                for (int i = 0; i < entetesColonne.Length; i++)
                {
                    string? entete = entetesColonne[i];
                    if (string.IsNullOrEmpty(entete)) continue;
                    int col = i + 1;
                    var cellEntete = recap.Cell(LIGNE_ENTETE_RECAP, col);
                    if (cellEntete.IsEmpty()) cellEntete.SetValue(entete);
                }
            }
        }

        /// <summary>Supprime les lignes « fantômes » de la Récap. : formules <c>=[0]!ZNxxx</c> du template pointant vers ModFeuille
        /// (le <c>[0]!</c> les identifie sans ambiguïté). Parcours descendant pour ne pas décaler les indices.</summary>
        private static void NettoyerLignesGhost(IXLWorksheet recap)
        {
            int derniereLigne = recap.LastRowUsed()?.RowNumber() ?? LIGNE_ENTETE_RECAP;
            for (int row = derniereLigne; row > LIGNE_ENTETE_RECAP; row--)
            {
                if (LigneEstGhost(recap, row))
                    recap.Row(row).Delete();
            }
        }

        private static bool LigneEstGhost(IXLWorksheet recap, int row)
        {
            foreach (var cell in recap.Row(row).CellsUsed())
            {
                if (!cell.HasFormula) continue;
                var f = cell.FormulaA1 ?? string.Empty;
                if (f.Contains("[0]!")) return true;
            }
            return false;
        }


        /// <summary>Nombre de gates balayées dans le fichier Stab, compté en feuilles à nom numérique (une par gate).
        /// Plus fiable que la Récap : ClosedXML ne calcule pas les formules, le <c>CachedValue</c> de la col A est toujours 0.
        /// Renvoie 0 hors Stabilité.</summary>
        private int CompterGatesStabPourGraphe()
        {
            if (_workbook == null || _typeMesureCourant != TypeMesure.Stabilite) return 0;
            int compteur = 0;
            foreach (var ws in _workbook.Worksheets)
            {
                if (int.TryParse(ws.Name, out _)) compteur++;
            }
            return compteur;
        }

        /// <summary>Adapte le graphe Stabilité du .xlsx aux mesures : plages des séries ramenées au nombre réel de gates
        /// (le <c>$A$6:$A$15</c> figé de Stab1.xls lisait les lignes vides comme 0 et l'axe Y log rendait un graphe vide),
        /// bornes min/max de l'axe Y retirées (<c>[1E-12, 1E-08]</c> historique), axe X recalé en bas.
        /// Ne modifie que <c>xl/charts/chart*.xml</c> : CEAO/SDAO lit les zones nommées <c>ZNRecapS_*</c>, pas le graphe.</summary>
        private static void RendreGrapheStabDynamique(string xlsmPath, int nbGates)
        {
            if (nbGates <= 0) return;
            int derniereLigne = 5 + nbGates;

            using var zip = ZipFile.Open(xlsmPath, ZipArchiveMode.Update);
            var entries = zip.Entries
                .Where(e => e.FullName.StartsWith("xl/charts/chart") && e.FullName.EndsWith(".xml"))
                .ToList();

            foreach (var entry in entries)
            {
                string contenu;
                using (var s = entry.Open())
                using (var r = new StreamReader(s))
                {
                    contenu = r.ReadToEnd();
                }

                string nouveau = contenu;

                // 1. Plages des séries : $X$6:$X$<N> → $X$6:$X$<5+nbGates>. `\d+` permet de re-patcher à chaque gate du balayage.
                nouveau = System.Text.RegularExpressions.Regex.Replace(
                    nouveau,
                    @"\$([A-Z]+)\$6:\$([A-Z]+)\$\d+",
                    m => $"${m.Groups[1].Value}$6:${m.Groups[2].Value}${derniereLigne}");

                // 2. Retire <c:min/> et <c:max/> de l'axe Y (<c:valAx>) pour qu'Excel les calcule ; l'axe X (catAx) n'en a pas.
                nouveau = System.Text.RegularExpressions.Regex.Replace(
                    nouveau,
                    "(<c:valAx>.*?<c:scaling>)(.*?)(</c:scaling>.*?</c:valAx>)",
                    m =>
                    {
                        string scaling = m.Groups[2].Value;
                        scaling = System.Text.RegularExpressions.Regex.Replace(
                            scaling, @"<c:min\s+val=""[^""]*""\s*/>", "");
                        scaling = System.Text.RegularExpressions.Regex.Replace(
                            scaling, @"<c:max\s+val=""[^""]*""\s*/>", "");
                        return m.Groups[1].Value + scaling + m.Groups[3].Value;
                    },
                    System.Text.RegularExpressions.RegexOptions.Singleline);

                // 3. Axe X en bas : le template croise l'axe Y à 1E-12 (Stab1.xls), ce qui le met au milieu du graphe
                //    si l'échelle descend plus bas. <c:crossesAt> du catAx → <c:crosses val="min"/>.
                nouveau = System.Text.RegularExpressions.Regex.Replace(
                    nouveau,
                    "(<c:catAx>.*?)<c:crossesAt val=\"[^\"]*\"/>(.*?</c:catAx>)",
                    "$1<c:crosses val=\"min\"/>$2",
                    System.Text.RegularExpressions.RegexOptions.Singleline);

                if (nouveau == contenu) continue;

                string fullName = entry.FullName;
                entry.Delete();
                var nouvelle = zip.CreateEntry(fullName);
                using (var s = nouvelle.Open())
                using (var w = new StreamWriter(s))
                {
                    w.Write(nouveau);
                }
            }
        }

        /// <summary>Force Excel à tout recalculer (formules ET caches de graphes) à l'ouverture, via
        /// <c>fullCalcOnLoad="1"</c> dans <c>xl/workbook.xml</c>. ClosedXML ne touche pas aux <c>numCache</c> des
        /// graphes : figés à 0, ils déclenchent sur l'axe Y log de la Stabilité le popup « Impossible de représenter
        /// les valeurs nulles ou négatives » et le graphe reste vide.</summary>
        private static void ForcerRecalculAuOuverture(string xlsmPath)
        {
            const string relPath = "xl/workbook.xml";
            using var zip = ZipFile.Open(xlsmPath, ZipArchiveMode.Update);
            var entry = zip.GetEntry(relPath);
            if (entry == null) return;

            string contenu;
            using (var s = entry.Open())
            using (var r = new StreamReader(s))
            {
                contenu = r.ReadToEnd();
            }

            // Le workbook peut avoir un préfixe de namespace XML (<x:workbook>, <x:calcPr>) : d'où (\w+:)?.
            // 3 cas : calcPr avec fullCalcOnLoad → mis à 1 ; calcPr sans l'attribut → ajouté ;
            // pas de calcPr → injecté avant </workbook> avec le même préfixe.
            if (System.Text.RegularExpressions.Regex.IsMatch(contenu, @"<(\w+:)?calcPr\b[^>]*\bfullCalcOnLoad="))
            {
                contenu = System.Text.RegularExpressions.Regex.Replace(
                    contenu, @"\bfullCalcOnLoad=""[^""]*""", "fullCalcOnLoad=\"1\"");
            }
            else if (System.Text.RegularExpressions.Regex.IsMatch(contenu, @"<(\w+:)?calcPr\b"))
            {
                contenu = System.Text.RegularExpressions.Regex.Replace(
                    contenu, @"<((\w+:)?calcPr)\b", "<$1 fullCalcOnLoad=\"1\"");
            }
            else
            {
                var matchPrefix = System.Text.RegularExpressions.Regex.Match(
                    contenu, @"</(?<p>(\w+:)?)workbook>");
                if (matchPrefix.Success)
                {
                    string p = matchPrefix.Groups["p"].Value;
                    contenu = contenu.Replace(
                        $"</{p}workbook>",
                        $"<{p}calcPr fullCalcOnLoad=\"1\"/></{p}workbook>");
                }
            }

            entry.Delete();
            var nouvelle = zip.CreateEntry(relPath);
            using (var s = nouvelle.Open())
            using (var w = new StreamWriter(s))
            {
                w.Write(contenu);
            }
        }

        /// <summary>Force les bornes min/max de l'axe Y log du graphe Stab d'après les sigmas relatifs estimés par gate :
        /// min = 10^floor(log10(min_sigma) - 0.5), max = 10^ceil(log10(max_sigma) + 0.5) (~3× de marge, arrondi à la décade).</summary>
        private static void CalibrerAxeYGrapheStab(string xlsmPath, IEnumerable<double> sigmas)
        {
            var positifs = sigmas.Where(s => s > 0).ToList();
            if (positifs.Count == 0) return;

            double minS = positifs.Min();
            double maxS = positifs.Max();
            double minBorne = Math.Pow(10, Math.Floor(Math.Log10(minS) - 0.5));
            double maxBorne = Math.Pow(10, Math.Ceiling(Math.Log10(maxS) + 0.5));

            // Au moins 1 décade d'écart pour un graphe lisible.
            if (maxBorne / minBorne < 10) maxBorne = minBorne * 10;

            string minXml = $"<c:min val=\"{minBorne.ToString("R", System.Globalization.CultureInfo.InvariantCulture)}\"/>";
            string maxXml = $"<c:max val=\"{maxBorne.ToString("R", System.Globalization.CultureInfo.InvariantCulture)}\"/>";

            using var zip = ZipFile.Open(xlsmPath, ZipArchiveMode.Update);
            var entries = zip.Entries
                .Where(e => e.FullName.StartsWith("xl/charts/chart") && e.FullName.EndsWith(".xml"))
                .ToList();

            foreach (var entry in entries)
            {
                string contenu;
                using (var s = entry.Open())
                using (var r = new StreamReader(s))
                {
                    contenu = r.ReadToEnd();
                }

                // Réinjecte min/max dans le <c:scaling> de <c:valAx> (les éventuels existants sont remplacés).
                string nouveau = System.Text.RegularExpressions.Regex.Replace(
                    contenu,
                    "(<c:valAx>.*?<c:scaling>)(.*?)(</c:scaling>.*?</c:valAx>)",
                    m =>
                    {
                        string scaling = m.Groups[2].Value;
                        scaling = System.Text.RegularExpressions.Regex.Replace(
                            scaling, @"<c:min\s+val=""[^""]*""\s*/>", "");
                        scaling = System.Text.RegularExpressions.Regex.Replace(
                            scaling, @"<c:max\s+val=""[^""]*""\s*/>", "");
                        return m.Groups[1].Value + scaling + minXml + maxXml + m.Groups[3].Value;
                    },
                    System.Text.RegularExpressions.RegexOptions.Singleline);

                if (nouveau == contenu) continue;

                string fullName = entry.FullName;
                entry.Delete();
                var nouvelle = zip.CreateEntry(fullName);
                using (var s = nouvelle.Open())
                using (var w = new StreamWriter(s))
                {
                    w.Write(nouveau);
                }
            }
        }

        /// <summary>Supprime les <c>numCache</c>/<c>strCache</c> de tous les <c>xl/charts/chartN.xml</c>. Ceux du graphe Stab
        /// viennent de Stab1.xls (10 valeurs Y à 0) et feraient auto-scaler l'axe Y log jusqu'à 1E-308 ; sans cache,
        /// Excel relit les séries depuis les cellules.</summary>
        private static void ViderCachesGraphe(string xlsmPath)
        {
            using var zip = ZipFile.Open(xlsmPath, ZipArchiveMode.Update);
            var entries = zip.Entries
                .Where(e => e.FullName.StartsWith("xl/charts/chart") && e.FullName.EndsWith(".xml"))
                .ToList();

            foreach (var entry in entries)
            {
                string contenu;
                using (var s = entry.Open())
                using (var r = new StreamReader(s))
                {
                    contenu = r.ReadToEnd();
                }

                string nouveau = System.Text.RegularExpressions.Regex.Replace(
                    contenu,
                    @"<c:numCache>.*?</c:numCache>",
                    "",
                    System.Text.RegularExpressions.RegexOptions.Singleline);

                nouveau = System.Text.RegularExpressions.Regex.Replace(
                    nouveau,
                    @"<c:strCache>.*?</c:strCache>",
                    "",
                    System.Text.RegularExpressions.RegexOptions.Singleline);

                if (nouveau == contenu) continue;

                string fullName = entry.FullName;
                entry.Delete();
                var nouvelle = zip.CreateEntry(fullName);
                using (var s = nouvelle.Open())
                using (var w = new StreamWriter(s))
                {
                    w.Write(nouveau);
                }
            }
        }

        /// <summary>Réécrit le Target de la relation externe du .xlsx pour qu'il pointe vers le Metrologo.xla configuré.</summary>
        private static void PatcherLienMacroXLA(string xlsmPath, string xlaPath)
        {
            const string relPath = "xl/externalLinks/_rels/externalLink1.xml.rels";

            // URI file:/// avec slashes (compatibilité XML)
            string uri = "file:///" + xlaPath.Replace('\\', '/');

            using var zip = ZipFile.Open(xlsmPath, ZipArchiveMode.Update);
            var entry = zip.GetEntry(relPath);
            if (entry == null) return;

            string contenu;
            using (var reader = new StreamReader(entry.Open()))
            {
                contenu = reader.ReadToEnd();
            }

            var rewritten = Regex.Replace(contenu,
                @"Target=""[^""]*Metrologo\.xla""",
                $@"Target=""{uri}""",
                RegexOptions.IgnoreCase);

            if (rewritten == contenu) return; // rien à changer

            entry.Delete();
            var newEntry = zip.CreateEntry(relPath);
            using var writer = new StreamWriter(newEntry.Open());
            writer.Write(rewritten);
        }

        public void FermerExcel()
        {
            _workbook?.Dispose();
            _workbook = null;
            _feuilleMesure = null;
        }

        // ---------- Utilitaires ----------

        // Mots de passe connus de ModFeuille dans les templates métier.
        private static readonly string[] _motsDePasseFeuille = { "METROL", "metrol" };

        private static void DeprotegerFeuille(IXLWorksheet feuille)
        {
            try { feuille.Unprotect(); return; }
            catch { /* protégée par mot de passe : essais ci-dessous */ }

            foreach (var mdp in _motsDePasseFeuille)
            {
                try { feuille.Unprotect(mdp); return; }
                catch { /* mauvais mot de passe — on essaie le suivant */ }
            }

            throw new InvalidOperationException(
                $"Impossible de déprotéger la feuille « {feuille.Name} » : "
                + "aucun mot de passe connu ne correspond. Vérifiez le template.");
        }

        private void SetNamed(string name, object value)
        {
            if (_feuilleMesure == null) return;
            try
            {
                if (_feuilleMesure.DefinedNames.TryGetValue(name, out var defName))
                {
                    defName.Ranges.First().FirstCell().SetValue(XLCellValue.FromObject(value));
                }
            }
            catch { /* zone nommée absente → silencieux */ }
        }

        private static string NomAppareilDepuisCatalogue(string idModele)
        {
            if (string.IsNullOrEmpty(idModele)) return string.Empty;
            var modele = CatalogueAppareilsService.Instance.Modeles
                .FirstOrDefault(m => m.Id == idModele);
            return modele?.Nom ?? idModele;
        }

        private int? TrouverLigneZone(string name)
        {
            if (_feuilleMesure == null) return null;
            try
            {
                if (_feuilleMesure.DefinedNames.TryGetValue(name, out var defName))
                {
                    return defName.Ranges.First().FirstCell().Address.RowNumber;
                }
            }
            catch { }
            return null;
        }

        private static bool FichierEstVerrouille(string chemin)
        {
            try
            {
                using var fs = File.Open(chemin, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                return false;
            }
            catch (IOException)
            {
                return true;
            }
        }

        /// <summary>Supprime les feuilles à nom numérique du template Stab (héritées de <c>Stab1.xls</c>, 10 procédures
        /// auto figées du Delphi) et les formules de la Récap. qui les référencent. À n'appeler que sur un fichier
        /// neuf : sur un fichier existant ces feuilles contiennent de vraies mesures.</summary>
        private void NettoyerFeuillesNumeriquesResiduellesEtRecap()
        {
            if (_workbook == null) return;

            // 1. Feuilles dont le nom est un entier ("1".."10"), liste copiée avant suppression.
            var aSupprimer = _workbook.Worksheets
                .Where(w => int.TryParse(w.Name, out _))
                .Select(w => w.Name)
                .ToList();

            foreach (var nom in aSupprimer)
            {
                try { _workbook.Worksheet(nom).Delete(); }
                catch { /* feuille déjà partie ou nom invalide — non bloquant */ }
            }

            // 2. Formules pré-remplies de la Récap. (A6:H15) pointant vers ces feuilles : effacées, sinon #REF! ou 0 fantômes.
            //    Repeuplée à chaque gate par MettreAJourRecapStabAsync.
            if (_workbook.Worksheets.Any(w => w.Name == NOM_RECAP))
            {
                var recap = _workbook.Worksheet(NOM_RECAP);
                DeprotegerFeuille(recap);
                recap.Range("A6:H15").Clear(XLClearOptions.Contents);

                // G19/H19 = GRANDE.VALEUR/PETITE.VALEUR héritées de Stab1.xls : plus lues et fausses au-delà de 10 temps de porte.
                recap.Range("G19:H19").Clear(XLClearOptions.Contents);
            }

            // 3. Les zones nommées workbook-scope vers ces feuilles (ex. _1ZNDeltaF) deviennent orphelines mais le
            //    graphe ne les utilise pas : laissées en place (ClosedXML les tolère).
        }

        /// <summary>Supprime la feuille <paramref name="nomFeuille"/> du classeur en cours, avec les lignes de la Récap.
        /// qui la référencent, pour écraser une feuille à nom fixe (F_avant_interv / F_finale) lors d'une relance. No-op si elle n'existe pas.</summary>
        private void SupprimerFeuilleEtLigneRecapSiExiste(string nomFeuille)
        {
            if (_workbook == null || string.IsNullOrWhiteSpace(nomFeuille)) return;

            var feuille = _workbook.Worksheets
                .FirstOrDefault(w => string.Equals(w.Name, nomFeuille, StringComparison.OrdinalIgnoreCase));
            if (feuille == null) return; // nom libre → rien à écraser

            // 1. Lignes de la Récap. pointant vers cette feuille (sinon doublon après recréation).
            try
            {
                if (_workbook.Worksheets.Any(w => w.Name == NOM_RECAP))
                {
                    var recap = _workbook.Worksheet(NOM_RECAP);
                    int derniere = recap.LastRowUsed()?.RowNumber() ?? LIGNE_ENTETE_RECAP;
                    for (int row = derniere; row > LIGNE_ENTETE_RECAP; row--)
                    {
                        if (LigneRefereFeuille(recap, row, nomFeuille))
                            recap.Row(row).Delete();
                    }
                }
            }
            catch { /* best-effort : la suppression de la feuille reste prioritaire */ }

            // 2. Supprime la feuille elle-même.
            try { feuille.Delete(); } catch { /* best-effort */ }
        }

        /// <inheritdoc/>
        public void SupprimerFeuilleSurDisque(string cheminFichier, string nomFeuille)
        {
            if (string.IsNullOrWhiteSpace(cheminFichier) || string.IsNullOrWhiteSpace(nomFeuille))
                return;

            // Le STOP tue Excel de façon asynchrone : le verrou saute après un court délai, d'où les réessais.
            const int maxTentatives = 6;
            for (int tentative = 1; tentative <= maxTentatives; tentative++)
            {
                if (!File.Exists(cheminFichier)) return;
                try
                {
                    using var wb = new XLWorkbook(cheminFichier);

                    var feuille = wb.Worksheets
                        .FirstOrDefault(w => string.Equals(w.Name, nomFeuille, StringComparison.OrdinalIgnoreCase));
                    if (feuille == null) return; // déjà absente

                    // 1. Lignes de la Récap. qui référencent cette feuille.
                    try
                    {
                        if (wb.Worksheets.Any(w => w.Name == NOM_RECAP))
                        {
                            var recap = wb.Worksheet(NOM_RECAP);
                            int derniere = recap.LastRowUsed()?.RowNumber() ?? LIGNE_ENTETE_RECAP;
                            for (int row = derniere; row > LIGNE_ENTETE_RECAP; row--)
                            {
                                if (LigneRefereFeuille(recap, row, nomFeuille))
                                    recap.Row(row).Delete();
                            }
                        }
                    }
                    catch { /* best-effort : la suppression de la feuille reste prioritaire */ }

                    // 2. Supprime la feuille et sauvegarde le fichier.
                    feuille.Delete();
                    wb.Save();

                    JournalLog.Info(CategorieLog.Excel, "EXCEL_FEUILLE_STOP_SUPPRIMEE_DISQUE",
                        $"Feuille « {nomFeuille} » supprimée sur disque suite à l'arrêt de la mesure.");
                    return;
                }
                catch (IOException)
                {
                    // Fichier encore verrouillé (Excel pas tout à fait mort) : on réessaie.
                    if (tentative < maxTentatives)
                        System.Threading.Thread.Sleep(400);
                }
                catch (Exception ex)
                {
                    JournalLog.Warn(CategorieLog.Excel, "EXCEL_SUPPR_FEUILLE_DISQUE_KO",
                        $"Suppression sur disque de « {nomFeuille} » échouée : {ex.Message}");
                    return;
                }
            }

            JournalLog.Warn(CategorieLog.Excel, "EXCEL_SUPPR_FEUILLE_DISQUE_VERROU",
                $"Suppression sur disque de « {nomFeuille} » abandonnée : fichier resté verrouillé.");
        }

        /// <summary>Vrai si une cellule de la ligne contient une formule référençant la feuille donnée.</summary>
        private static bool LigneRefereFeuille(IXLWorksheet recap, int row, string nomFeuille)
        {
            foreach (var cell in recap.Row(row).CellsUsed())
            {
                if (!cell.HasFormula) continue;
                var f = cell.FormulaA1 ?? string.Empty;
                if (f.IndexOf($"{nomFeuille}!", StringComparison.OrdinalIgnoreCase) >= 0
                    || f.IndexOf($"'{nomFeuille}'!", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }

        private string TrouverNomFeuilleUnique(TypeMesure type)
        {
            // Feuilles uniques (1 par classeur) : pas de numérotation.
            if (type == TypeMesure.FreqAvantInterv) return "F_avant_interv";
            if (type == TypeMesure.FreqFinale) return "F_finale";

            // Préfixes courts (inter1, topti1…). Fréquence et Stabilité : onglets numérotés seuls (1, 2, 3…) ;
            // en stab c'est le FICHIER qui s'incrémente (stab, stab1…), avec une feuille par temps de porte.
            string prefixe = type switch
            {
                TypeMesure.Frequence    => "",
                TypeMesure.Stabilite    => "",
                TypeMesure.Interval     => "inter",
                TypeMesure.TachyOptique => "topti",
                TypeMesure.TachyContact => "tcont",
                TypeMesure.Stroboscope  => "strob",
                _ => "Mesure"
            };

            int idx = 1;
            var existants = _workbook!.Worksheets.Select(w => w.Name)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            while (existants.Contains($"{prefixe}{idx}")) idx++;
            return $"{prefixe}{idx}";
        }

        /// <summary>Extrait le numéro de gate d'un nom de feuille Stab (« 1 » → 1). Tolère l'ancien format préfixé
        /// (« stab1 » → 1) pour les classeurs générés avant le renommage.</summary>
        private static bool TryExtraireNumeroGate(string nomFeuille, out int numGate)
        {
            numGate = 0;
            if (string.IsNullOrEmpty(nomFeuille)) return false;
            // Chiffres en fin de nom, quel que soit le préfixe.
            int i = nomFeuille.Length;
            while (i > 0 && char.IsDigit(nomFeuille[i - 1])) i--;
            string chiffres = nomFeuille.Substring(i);
            return int.TryParse(chiffres, out numGate);
        }

        private void ClonerZonesNommeesPourNouvelleFeuille(IXLWorksheet source, IXLWorksheet dest)
        {
            if (_workbook == null) return;

            foreach (var defName in _workbook.DefinedNames.ToList())
            {
                try
                {
                    if (!defName.RefersTo.Contains(source.Name + "!", StringComparison.OrdinalIgnoreCase))
                        continue;

                    string newRefersTo = defName.RefersTo
                        .Replace(source.Name + "!", dest.Name + "!",
                                 StringComparison.OrdinalIgnoreCase);

                    if (!dest.DefinedNames.Contains(defName.Name))
                    {
                        dest.DefinedNames.Add(defName.Name, newRefersTo);
                    }
                }
                catch { /* ignore #REF! */ }
            }
        }
    }
}
