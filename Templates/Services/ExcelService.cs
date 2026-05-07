using ClosedXML.Excel;
using Metrologo.Models;
using Metrologo.Services.Catalogue;
using Metrologo.Services.Incertitude;
using Metrologo.Services.Journal;
using JournalLog = Metrologo.Services.Journal.Journal;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Metrologo.Services
{
    public interface IExcelService
    {
        /// <summary>
        /// Initialise une feuille de mesure dans le classeur du FI. Le paramètre optionnel
        /// <paramref name="gateIndexOverride"/> permet de spécifier la gate à inscrire dans
        /// les zones nommées (<c>ZNGate</c>/<c>ZNLibGate</c>/<c>ZNValGateSecondes</c>) — utile
        /// pour les balayages de stabilité où chaque feuille est associée à une gate différente
        /// sans qu'on souhaite muter <c>config.GateIndex</c> partout.
        /// </summary>
        Task InitialiserRapportAsync(string numeroFI, Mesure configuration, Rubidium rubidium, int? gateIndexOverride = null, bool nouvelleSession = false);

        /// <summary>
        /// Pré-insère les lignes vides nécessaires pour les mesures à venir (au-delà des 2 lignes
        /// par défaut du template). Les formules <c>Fréq. Réelle</c> et <c>F(i)-F(i+1)</c> sont
        /// ajoutées. Les cellules HEURE et mesure restent vides — elles seront remplies en direct.
        /// </summary>
        Task PreparerLignesMesureAsync(int nbMesures);

        /// <summary>
        /// Sauvegarde le classeur ClosedXML sur le disque (sans ouvrir Excel) pour qu'il puisse
        /// être repris ensuite par <see cref="ExcelInteropHost"/> en écriture live.
        /// </summary>
        Task<string> SauvegarderSurDisqueAsync();

        /// <summary>Nom de la feuille créée pour cette mesure (ex: Freq1, Stab1).</summary>
        string NomFeuilleMesure { get; }

        /// <summary>
        /// Écrit la moyenne et la variance dans les zones nommées — à appeler après la boucle
        /// de mesures, pour que le Récap. cross-sheet fonctionne.
        /// </summary>
        Task EcrireStatsAsync(List<double> resultats);

        /// <summary>
        /// Écrit les N mesures (HEURE + valeur) directement dans la feuille de mesure courante
        /// via ClosedXML, **sans passer par Excel/Interop**. Utilisé en mode « invisible » pour
        /// la Stabilité : Excel n'est jamais ouvert pendant la mesure, tout se fait en mémoire,
        /// et l'utilisateur voit le fichier final apparaître à la toute fin du balayage.
        /// </summary>
        Task EcrireValeursBatchClosedXMLAsync(int ligneDebut, IList<(DateTime ts, double valeur)> mesures);

        Task MettreAJourRecapFreqAsync(Mesure mesure);
        Task MettreAJourRecapStabAsync(Mesure mesure);

        /// <summary>
        /// Re-ouvre depuis le disque un classeur déjà initialisé, après qu'Interop l'ait rempli
        /// avec les valeurs live. Ne crée pas de nouvelle feuille : on récupère la feuille dont
        /// le nom est <see cref="NomFeuilleMesure"/>.
        /// </summary>
        Task RouvrirClasseurAsync();

        /// <summary>
        /// Sauvegarde finale après Recap + stats. Le fichier reste ouvert par <see cref="ExcelInteropHost"/>
        /// (l'utilisateur peut continuer à l'inspecter).
        /// </summary>
        Task SauvegarderFinalAsync();

        void FermerExcel();
    }

    public class ExcelService : IExcelService
    {
        // Nom de la feuille modèle — jamais modifiée.
        private const string NOM_MODELE = "ModFeuille";
        private const string NOM_RECAP = "Récap.";

        // Première ligne de mesures dans le template
        private const int LIGNE_DEBUT_MESURES = 9;

        // Zones nommées indiquant l'emplacement d'insertion des lignes Recap (cf. template xltm).
        private const string ZN_RECAPF_DEBZONE = "ZNRecapF_DebZone";
        private const string ZN_RECAPS_DEBZONE = "ZNRecapS_DebZone";

        // Lignes de fallback si les zones ZNRecapF/S_DebZone sont absentes (valeurs par défaut du template).
        private const int LIGNE_FALLBACK_RECAPF = 19;
        private const int LIGNE_FALLBACK_RECAPS = 19;

        private XLWorkbook? _workbook;
        private IXLWorksheet? _feuilleMesure;   // feuille NOUVELLEMENT créée pour cette mesure
        private string _cheminFichier = string.Empty;
        private string _nomFeuilleMesure = string.Empty;
        private TypeMesure _typeMesureCourant;  // mémorisé pour PreparerLignesMesureAsync (col K Tachy)

        // Mémorisé à l'init pour qu'EcrireStatsAsync puisse résoudre les coefficients
        // CoeffA/CoeffB depuis le module CSV (cf. ModulesIncertitudeService) en fonction
        // de la moyenne calculée. Vide → on garde les hardcoded de l'init.
        private string _numModuleIncertitudeCourant = string.Empty;
        private double _tempsGateSecondesCourant;

        // Mémorisés pour reproduire en C# la conversion Indirect appliquée par la formule
        // de la colonne F (Fréq. Réelle) — ainsi la moyenne C# passée à ObtenirCoefficients
        // est mathématiquement identique à AVERAGE(F) calculée par Excel à la fin.
        private int _indexMultiplicateurCourant;
        private double _fNominaleCourant;
        private ModeMesure _modeMesureCourant;

        // Sigmas relatifs estimés par numéro de gate (clé = nom de la feuille parsé en int).
        // Alimenté par EcrireStatsAsync à chaque fin de gate Stab. Lu par
        // SauvegarderSurDisqueAsync pour calibrer l'axe Y log du graphe Stab — sans cette
        // calibration explicite, Excel auto-scale parfois sur une plage 1E-9..1E0 même
        // quand les vraies données sont 1E-8..1E-6.
        private readonly Dictionary<int, double> _sigmasRelatifsParGate = new();

        /// <summary>
        /// Colonne dédiée à la conversion Hz → tr/min pour les mesures de tachymétrie /
        /// stroboscope. Choisie volontairement éloignée des colonnes D-G (qui alimentent
        /// la Récap. Fréquence via les zones nommées scope-feuille) pour ne PAS interférer
        /// avec les formules cross-sheet de la Récap. CEAO/SDAO ne lit pas cette colonne.
        /// Position N après le décalage de 3 colonnes pour les nouvelles colonnes
        /// n°Module/Fonction/Condition 1 ajoutées en A/B/C.
        /// </summary>
        private const string COL_CONVERSION_TR_MIN = "N";

        public string NomFeuilleMesure => _nomFeuilleMesure;

        /// <summary>
        /// Chemin du fichier Excel réellement utilisé pour cette mesure (peut différer de
        /// <c>Mesures_{FI}.xlsm</c> si un fallback timestampé a été appliqué — cf. Excel verrouillé).
        /// Exposé pour que la UI puisse en informer l'utilisateur.
        /// </summary>
        public string CheminFichierGenere => _cheminFichier;

        /// <summary>Vrai si le fichier a dû être écrit sous un nom de fallback au lieu du nom principal.</summary>
        public bool FallbackTimestampUtilise { get; private set; }

        public async Task InitialiserRapportAsync(string numeroFI, Mesure config, Rubidium rubidium, int? gateIndexOverride = null, bool nouvelleSession = false)
        {
            int gateInscrite = gateIndexOverride ?? config.GateIndex;
            bool estStab = config.TypeMesure == TypeMesure.Stabilite;

            await Task.Run(() =>
            {
                // --- 1. Détermination du dossier et du fichier ---
                //    La Stabilité utilise un fichier séparé (Récap. à 8 colonnes spécifique)
                //    pour ne pas polluer la Récap. Fréquence du fichier principal du FI.
                string numFISafe = SanitizerNomFichier(numeroFI);

                string dossier = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "Metrologo", numFISafe);
                Directory.CreateDirectory(dossier);

                string suffixe = estStab ? "_Stab" : string.Empty;
                _cheminFichier = Path.Combine(dossier, $"Mesures{suffixe}_{numFISafe}.xlsm");
                FallbackTimestampUtilise = false;

                // Stabilité + nouvelle session + fichier existant → suffixe _v2, _v3, etc.
                // Évite que le graphe Stab affiche les valeurs de la mesure précédente
                // mélangées avec celles de la nouvelle session sur le même FI.
                if (estStab && nouvelleSession && File.Exists(_cheminFichier))
                {
                    string baseSansExt = Path.Combine(dossier, $"Mesures{suffixe}_{numFISafe}");
                    int v = 2;
                    string candidat;
                    do
                    {
                        candidat = $"{baseSansExt}_v{v}.xlsm";
                        v++;
                    } while (File.Exists(candidat) && v < 1000);
                    _cheminFichier = candidat;
                }

                // --- 2. Ouverture : fichier existant ou copie du template ---
                //    Template Stab dédié pour la Stabilité (Récap. 8 cols + zones nommées
                //    ZNRecapS_*), template Fréquence pour le reste.
                string nomTemplate = estStab ? "METROLOGO_Stab.xltm" : "METROLOGO.xltm";
                string templatePath = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory, "Templates", nomTemplate);

                bool partirDuTemplate = false;
                if (File.Exists(_cheminFichier))
                {
                    if (FichierEstVerrouille(_cheminFichier))
                    {
                        // Tentative 1 : fermer Excel poliment via WM_CLOSE (marche si pas de dialogue bloquant).
                        WindowHelper.FermerFenetresExcel(Path.GetFileName(_cheminFichier));
                        var sw = System.Diagnostics.Stopwatch.StartNew();
                        while (FichierEstVerrouille(_cheminFichier) && sw.ElapsedMilliseconds < 3000)
                        {
                            System.Threading.Thread.Sleep(150);
                        }

                        if (FichierEstVerrouille(_cheminFichier))
                        {
                            // Tentative 2 : fallback sur un nom de fichier timestampé — on n'interrompt
                            //                ni la mesure ni le travail de l'utilisateur dans Excel.
                            string cheminAlternatif = Path.Combine(
                                dossier,
                                $"Mesures_{numeroFI}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsm");

                            // On tente de copier le fichier existant pour garder Récap. + historique.
                            // Si Excel bloque aussi la lecture, on repart du template vierge.
                            try
                            {
                                File.Copy(_cheminFichier, cheminAlternatif, overwrite: false);
                            }
                            catch (IOException)
                            {
                                partirDuTemplate = true;
                            }

                            _cheminFichier = cheminAlternatif;
                            FallbackTimestampUtilise = true;
                        }
                    }

                    _workbook = partirDuTemplate
                        ? new XLWorkbook(templatePath)
                        : new XLWorkbook(_cheminFichier);
                }
                else
                {
                    _workbook = new XLWorkbook(templatePath);
                    partirDuTemplate = true;
                }

                // --- 2b. Nettoyage des feuilles « 1 » à « N » héritées du Stab1.xls historique ---
                // Le template METROLOGO_Stab.xltm contient 10 feuilles vides nommées "1".."10"
                // (slots des 10 procédures auto figées du Delphi historique). Sans nettoyage,
                // TrouverNomFeuilleUnique attribue 11, 12, 13… aux nouvelles gates et la
                // Récap. affiche 10 lignes parasites pointant vers des feuilles vides.
                // On ne fait ce nettoyage qu'à l'OUVERTURE du template (= nouveau fichier),
                // jamais sur un fichier existant — sinon on détruirait des mesures précédentes.
                if (partirDuTemplate && estStab)
                {
                    NettoyerFeuillesNumeriquesResiduellesEtRecap();
                    // Nouvelle session Stab → on oublie les sigmas de la session précédente
                    // (sinon le dico accumule et la calibration de l'axe Y serait faussée).
                    _sigmasRelatifsParGate.Clear();
                }

                var modFeuille = _workbook.Worksheet(NOM_MODELE);

                // --- 3. Création d'une nouvelle feuille (copie de ModFeuille) ---
                _nomFeuilleMesure = TrouverNomFeuilleUnique(config.TypeMesure);
                _feuilleMesure = modFeuille.CopyTo(_nomFeuilleMesure);

                // ModFeuille est cachée dans le template (convention métier — c'est juste un
                // modèle interne qui ne doit pas apparaître à l'utilisateur). Mais la copie
                // hérite de cet attribut Hidden — il faut donc forcer la visibilité de chaque
                // nouvelle feuille de mesure (1, 2, 3, … ou Freq1, Stab1, …).
                _feuilleMesure.Visibility = XLWorksheetVisibility.Visible;

                DeprotegerFeuille(_feuilleMesure);

                // Largeurs : le template a été décalé de 3 colonnes pour accueillir
                // n°Module / Fonction / Condition 1. Les anciennes col B/C/E deviennent
                // E/F/H (le bloc des mesures HEURE/Mesurée/Réelle/Delta est maintenant
                // en col D-G).
                _feuilleMesure.Column("E").Width = 30;
                _feuilleMesure.Column("F").Width = 28;
                _feuilleMesure.Column("H").Width = 30;

                // --- 4. Clonage des zones nommées en sheet-scope sur la nouvelle feuille ---
                ClonerZonesNommeesPourNouvelleFeuille(modFeuille, _feuilleMesure);

                // --- 5. En-têtes adaptatifs (colonnes D/E/F/G + labels de lignes col E) ---
                // Décalage post-ajout des colonnes A/B/C (n°Module / Fonction / Condition 1) :
                //   Ancien A → Nouveau D (HEURE)
                //   Ancien B → Nouveau E (Mesurée)
                //   Ancien C → Nouveau F (Réelle)
                //   Ancien D → Nouveau G (Delta)
                //   Labels en B?? → E?? (les valeurs associées en col F via formules)
                var entetes = EnTetesMesureHelper.Pour(config.TypeMesure);
                _feuilleMesure.Cell("D7").SetValue(entetes.EnteteHeure);
                _feuilleMesure.Cell("E7").SetValue(entetes.EnteteMesuree);
                _feuilleMesure.Cell("F7").SetValue(entetes.EnteteReelle);
                _feuilleMesure.Cell("G7").SetValue(entetes.EnteteDelta);
                _feuilleMesure.Cell("E13").SetValue(entetes.LabelMoyenne);
                _feuilleMesure.Cell("E21").SetValue(entetes.LabelFreqRef);
                _feuilleMesure.Cell("E23").SetValue(entetes.LabelFreqCorr);
                _feuilleMesure.Cell("E25").SetValue(entetes.LabelIncertResol);
                _feuilleMesure.Cell("E31").SetValue(entetes.LabelIncertGlob);

                // --- 5b. Nouvelles colonnes n°Module / Fonction / Condition 1 (ligne 9) ---
                // Ces 3 valeurs sont écrites une seule fois sur la 1ère ligne de mesure
                // (ligne 9 = 1ère mesure dans le template). Elles décrivent la session :
                // quel module/fonction de l'instrument est utilisé + temps de gate sélectionné.
                var (module, fonction) = MesureConfigService.ObtenirPourType(config.TypeMesure);
                if (!string.IsNullOrEmpty(module)) _feuilleMesure.Cell("A9").SetValue(module);
                if (!string.IsNullOrEmpty(fonction)) _feuilleMesure.Cell("B9").SetValue(fonction);
                _feuilleMesure.Cell("C9").SetValue(EnTetesMesureHelper.SecondesGate(gateInscrite));

                // Mémorisation pour PreparerLignesMesureAsync (col N conversion tr/min)
                _typeMesureCourant = config.TypeMesure;

                // Mémorisation pour EcrireStatsAsync — la moyenne sera calculée à la fin
                // de la boucle de mesures, à ce moment on appellera ObtenirCoefficients
                // pour ce module avec (fonction, temps de gate, fréquence moyenne).
                _numModuleIncertitudeCourant = config.NumModuleIncertitude ?? string.Empty;
                _tempsGateSecondesCourant = EnTetesMesureHelper.SecondesGate(gateInscrite);
                _indexMultiplicateurCourant = config.IndexMultiplicateur;
                _fNominaleCourant = config.FNominale;
                _modeMesureCourant = config.ModeMesure;

                // --- 5c. Tachymétrie : colonne N dédiée à la conversion Hz → tr/min ---
                // Le compteur GPIB mesure une fréquence d'impulsions (Hz). La conversion
                // tr/min = Hz × 60 (1 imp/tour) est faite côté Excel pour rester visible
                // à l'utilisateur. Colonne N (= ancienne K + 3) choisie loin de D-G pour
                // ne pas interférer avec la Récap qui lit E/F/G via les zones nommées.
                // Note : la mesure brute est désormais en col E (post-décalage de 3).
                if (config.TypeMesure == TypeMesure.TachyContact)
                {
                    _feuilleMesure.Cell($"{COL_CONVERSION_TR_MIN}7").SetValue("Vitesse (tr/min)");
                    _feuilleMesure.Column(COL_CONVERSION_TR_MIN).Width = 18;
                    // Formules pour les 2 lignes de mesure pré-existantes du template (9 et 10).
                    // Les lignes additionnelles (11+) seront alimentées par PreparerLignesMesureAsync.
                    _feuilleMesure.Cell($"{COL_CONVERSION_TR_MIN}9").FormulaA1 = "=E9*60";
                    _feuilleMesure.Cell($"{COL_CONVERSION_TR_MIN}10").FormulaA1 = "=E10*60";
                }

                // --- 6. Métadonnées via zones nommées (sheet-scope) ---
                SetNamed("ZNNoFiche", numeroFI);
                SetNamed("ZNDate", DateTime.Now.ToString("dd/MM/yyyy"));
                SetNamed("ZNTypeMesure", EnTetesMesureHelper.LibelleType(config.TypeMesure));
                SetNamed("ZNFreqUtilise", NomAppareilDepuisCatalogue(config.IdModeleCatalogue));
                SetNamed("ZNRubidium",
                    rubidium.Designation + (rubidium.AvecGPS ? " (raccord GPS)" : " (raccord Allouis)"));
                SetNamed("ZNGate", EnTetesMesureHelper.LibelleGate(gateInscrite));
                SetNamed("ZNLibGate", EnTetesMesureHelper.LibelleGate(gateInscrite));
                SetNamed("ZNValGateSecondes", EnTetesMesureHelper.SecondesGate(gateInscrite));
                SetNamed("ZNModeMesure", config.ModeMesure == ModeMesure.Direct ? "Direct" : "Indirect");
                SetNamed("ZNCoeffMult", config.IndexMultiplicateur);
                SetNamed("ZNValFNominale", config.FNominale);
                SetNamed("ZNNbMesures", config.NbMesures);
                SetNamed("ZNIncertResol", config.Resolution);
                SetNamed("ZNIncertSup", config.IncertSupp);
                SetNamed("ZNFreqRef", rubidium.FrequenceMoyenne);

                // Fallback hardcoded : utilisé tel quel si aucun module d'incertitude n'est
                // sélectionné, ou si la moyenne tombe hors des plages couvertes par le module.
                // Si un module est associé, EcrireStatsAsync surcharge ZNCoeffA/ZNCoeffB en
                // fin de boucle avec la valeur correspondant à la fréquence moyenne mesurée.
                SetNamed("ZNCoeffA", 1e-10);
                SetNamed("ZNCoeffB", 5e-13);
                SetNamed("ZNNbMesAccredite", 30);
                SetNamed("ZNTempsMesureAccredite", 10);
            });
        }

        public async Task PreparerLignesMesureAsync(int nbMesures)
        {
            if (_feuilleMesure == null || _workbook == null || nbMesures <= 0) return;

            await Task.Run(() =>
            {
                // *** KEY FIX (identique à l'ancien AjouterResultatsAsync) ***
                // Le template a 2 lignes de mesures (9 et 10) et les labels commencent en 13.
                // Pour chaque mesure au-delà de la 2e, on insère une rangée AVANT ZNPointInsertion.
                if (nbMesures > 2)
                {
                    int pointInsertionRow = TrouverLigneZone("ZNPointInsertion") ?? 11;
                    _feuilleMesure.Row(pointInsertionRow).InsertRowsAbove(nbMesures - 2);
                }

                // Ajoute les formules Fréq. Réelle (col C) et delta (col D) pour les lignes
                // créées par InsertRowsAbove. Les cellules HEURE/mesure restent vides — elles
                // seront écrites en direct par ExcelInteropHost pendant la boucle de mesures.
                bool conversionTrMin = _typeMesureCourant == TypeMesure.TachyContact;
                for (int i = 2; i < nbMesures; i++)
                {
                    int row = LIGNE_DEBUT_MESURES + i;
                    // Décalage de 3 colonnes : ancien B/C/D → nouveau E/F/G.
                    //   E = mesurée (lue par GPIB) ; F = réelle (formule de correction) ;
                    //   G = delta entre 2 mesures consécutives.
                    _feuilleMesure.Cell($"F{row}").FormulaA1 =
                        $"IF(ISBLANK(ZNCoeffMult),E{row},"
                        + $"(((E{row}-10000000)/(POWER(10,ZNCoeffMult)*10000000))+1)*ZNValFNominale)";
                    _feuilleMesure.Cell($"G{row}").FormulaA1 = $"F{row - 1}-F{row}";
                    if (conversionTrMin)
                        _feuilleMesure.Cell($"{COL_CONVERSION_TR_MIN}{row}").FormulaA1 = $"=E{row}*60";
                }

                // Restaure les formules métier historiques (héritage Stab1.xls) qui calculent
                // les statistiques à partir des plages de mesures. Faites ici car ModFeuille
                // du template Stab n'a aucune de ces formules (elles n'existaient que sur
                // les feuilles 1..10 vestiges) ; pour Freq, ModFeuille a déjà ZNEcartType /
                // ZNIncertEcartType / ZNFreqCorr — la réécriture est idempotente.
                EcrireFormulesStatsMetier(nbMesures);
            });
        }

        /// <summary>
        /// Écrit sur la feuille de mesure courante les formules métier historiques :
        /// <list type="bullet">
        ///   <item><c>ZNFreqMoyReel</c> = AVERAGE de la colonne F (Fréq. Réelle post-conversion Indirect).</item>
        ///   <item><c>ZNVariance</c> = appel macro VBA <c>Cal_variance(SUMSQ(deltas), nbMesures, moyenne)</c>
        ///         — formule métier différente de la variance n-1 classique.</item>
        ///   <item>Statistiques dérivées (<c>ZNEcartType</c>, <c>ZNIncertEcartType</c>, <c>ZNFreqCorr</c>)
        ///         pour le cas où ModFeuille ne les contiendrait pas (template Stab).</item>
        /// </list>
        /// Les zones nommées sont sheet-scope (clonées par <c>ClonerZonesNommeesPourNouvelleFeuille</c>) ;
        /// on récupère la cellule cible via <c>NamedRange(...)</c> pour rester abstrait du layout exact.
        /// </summary>
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
                    // Sheet-scope d'abord (cloné depuis ModFeuille pour cette feuille de mesure).
                    if (_feuilleMesure!.DefinedNames.TryGetValue(nomZone, out var defName))
                    {
                        cell = defName?.Ranges.FirstOrDefault()?.FirstCell();
                    }
                }
                catch { /* zone absente sur la feuille — silencieux */ }
                if (cell != null) cell.FormulaA1 = formule;
            }

            // Moyenne sur la colonne F (= Fréq. Réelle, après conversion Indirect ligne par ligne).
            EcrireSurZN("ZNFreqMoyReel",
                $"IF(ISBLANK(ZNNbMesures),,AVERAGE(F{ligneDeb}:F{ligneFin}))");

            // Variance via macro VBA Metrologo.xla. La 1ère ligne de delta est G{deb+1}
            // (G{deb} reste vide car il n'y a pas de mesure précédente pour la 1ère).
            EcrireSurZN("ZNVariance",
                $"[1]!Cal_variance(SUMSQ(G{ligneDeb + 1}:G{ligneFin}),ZNNbMesures,ZNFreqMoyReel)");

            // Statistiques dérivées — idempotentes pour Freq (déjà dans ModFeuille),
            // créées pour Stab (ModFeuille du template Stab n'avait rien).
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
                    // Décalage de 3 colonnes : col 4 (D) = HEURE ; col 5 (E) = mesure brute.
                    _feuilleMesure.Cell(row, 4).SetValue(mesures[i].ts.ToString("HH:mm:ss"));
                    _feuilleMesure.Cell(row, 5).SetValue(mesures[i].valeur);
                }
            });
        }

        public async Task EcrireStatsAsync(List<double> resultats)
        {
            if (_feuilleMesure == null || _workbook == null || resultats.Count == 0) return;

            await Task.Run(() =>
            {
                if (resultats.Count == 0) return;

                // ZNFreqMoyReel et ZNVariance ne sont PLUS écrits depuis C# : la formule
                // métier (AVERAGE / Cal_variance VBA) est posée par EcrireFormulesStatsMetier
                // dans PreparerLignesMesureAsync. Excel calculera ces zones en ouvrant le
                // fichier (fullCalcOnLoad="1" garanti par ForcerRecalculAuOuverture).

                // Estimation sigma relatif (n-1 classique) pour calibrer l'axe Y du graphe
                // Stab à la sauvegarde. Pas la formule métier exacte (Cal_variance VBA), mais
                // le bon ordre de grandeur — ce qui suffit pour fixer min/max log à
                // ±0.5 décade autour des données réelles.
                if (_typeMesureCourant == TypeMesure.Stabilite
                    && resultats.Count >= 2
                    && int.TryParse(_nomFeuilleMesure, out int numGate))
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
                // Pour choisir la bonne ligne du module CSV, on a besoin de la fréquence
                // moyenne RÉELLE (post-conversion Indirect). On ne peut pas la lire d'Excel
                // ici (ClosedXML ne calcule pas les formules). On la recalcule en C# en
                // appliquant la même conversion linéaire que la formule de la colonne F :
                //   AVERAGE(F_i) = AVERAGE( conv(E_i) ) = conv( AVERAGE(E_i) )
                // (la conversion étant affine, elle commute avec la moyenne).
                // Si pas de module sélectionné, les valeurs hardcoded posées à l'init dans
                // ZNCoeffA / ZNCoeffB restent en place (fallback).
                if (!string.IsNullOrEmpty(_numModuleIncertitudeCourant))
                {
                    double moyenneBrute = resultats.Average();
                    double moyenneReelle = ConvertirEnFreqReelle(moyenneBrute);

                    string fonction = IncertitudeFonctionHelper.NomFonction(_typeMesureCourant);
                    var (coeffA, coeffB) = ModulesIncertitudeService.ObtenirCoefficients(
                        _numModuleIncertitudeCourant, _typeMesureCourant, fonction,
                        _tempsGateSecondesCourant, moyenneReelle);

                    if (coeffA > 0 || coeffB > 0)
                    {
                        SetNamed("ZNCoeffA", coeffA);
                        SetNamed("ZNCoeffB", coeffB);
                        JournalLog.Info(CategorieLog.Mesure, "INCERT_COEFFS_RESOLUS",
                            $"Module {_numModuleIncertitudeCourant} : CoeffA={coeffA:G6} CoeffB={coeffB:G6} "
                          + $"(fonction={fonction}, gate={_tempsGateSecondesCourant}s, freq={moyenneReelle:G6} Hz)");
                    }
                    else
                    {
                        JournalLog.Warn(CategorieLog.Mesure, "INCERT_NO_MATCH",
                            $"Module {_numModuleIncertitudeCourant} : aucune ligne ne couvre "
                          + $"(fonction={fonction}, gate={_tempsGateSecondesCourant}s, freq={moyenneReelle:G6} Hz) "
                          + "— coefficients par défaut utilisés.");
                    }
                }
            });
        }

        /// <summary>
        /// Reproduit en C# la conversion appliquée par la formule de la colonne F (Fréq. Réelle)
        /// du template :
        /// <code>
        /// IF(ISBLANK(ZNCoeffMult), E, (((E-1e7) / (10^Mult * 1e7)) + 1) * FNominale)
        /// </code>
        /// La moyenne arithmétique des résultats convertis (= AVERAGE(F) côté Excel) est
        /// mathématiquement égale à <c>ConvertirEnFreqReelle(AVERAGE(E))</c> car la
        /// conversion est affine (commute avec la moyenne).
        /// </summary>
        private double ConvertirEnFreqReelle(double mesureBrute)
        {
            // Mode Direct = pas de conversion (Excel : ZNCoeffMult ISBLANK → renvoie E).
            // Côté C#, ZNCoeffMult est toujours posé (int) → on reproduit le comportement
            // attendu en court-circuitant pour le mode Direct.
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

                // Patche le graphe Stab : plages des séries adaptées au nombre de gates
                // effectivement balayées + axe Y rendu auto (retrait des min/max fixes
                // 1E-12/1E-08 hérités de Stab1.xls qui figeaient l'échelle). Ne touche
                // PAS aux cellules ni aux zones nommées de la Récap → CEAO/SDAO non
                // impacté (il consomme les données, pas les graphes).
                if (nbGatesStab > 0)
                {
                    try { RendreGrapheStabDynamique(_cheminFichier, nbGatesStab); }
                    catch { /* best-effort */ }
                }

                // Vidage des caches du graphe (numCache/strCache hérités du Stab1.xls
                // = 10 valeurs Y figées à 0). Sans ce nettoyage, Excel à l'ouverture
                // utilise les zéros pour calibrer l'axe Y log → échelle énorme et vide.
                // Combiné à fullCalcOnLoad, force un vrai recalcul depuis les cellules.
                try { ViderCachesGraphe(_cheminFichier); }
                catch { /* best-effort */ }

                // Calibration explicite de l'axe Y log du graphe Stab : Excel auto-scale
                // souvent vers 1E-9..1E0 sans cache et sans bornes (~5 décades en trop).
                // On force min/max ~0.5 décade autour des sigmas relatifs estimés.
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

                // Récupère la feuille de mesure par son nom (créée à l'initialisation).
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
            });
        }

        /// <summary>
        /// Ajoute une ligne de récapitulatif Fréquence dans la feuille <c>Récap.</c>.
        /// Portage de <c>TfrmMain.MajRecapFreq</c> (F_Main.pas:2421) — la structure de Recap
        /// n'est pas modifiée, seules de nouvelles lignes de données sont insérées au point
        /// défini par la zone nommée <c>ZNRecapF_DebZone</c>.
        /// </summary>
        public async Task MettreAJourRecapFreqAsync(Mesure mesure)
        {
            if (_workbook == null || _feuilleMesure == null) return;
            string nomFeuille = _feuilleMesure.Name;

            await Task.Run(() => EcrireLigneRecap(
                nomFeuille,
                ZN_RECAPF_DEBZONE, LIGNE_FALLBACK_RECAPF,
                new[]
                {
                    $"='{nomFeuille}'!ZNFreqMoyReel",       // Col 1 : fréquence moyenne
                    $"='{nomFeuille}'!ZNLibGate",           // Col 2 : temps de mesure (libellé gate)
                    $"='{nomFeuille}'!ZNFreqCorr",          // Col 3 : fréquence corrigée
                    $"='{nomFeuille}'!ZNEcartType",         // Col 4 : écart-type
                    null,                                    // Col 5 : fréquence indiquée (valeur directe)
                    $"='{nomFeuille}'!ZNIncertResol",       // Col 6 : incertitude de résolution
                    $"='{nomFeuille}'!ZNIncertSup",         // Col 7 : incertitude supplémentaire
                    $"='{nomFeuille}'!ZNIncertAccreditee",  // Col 8 : incertitude accréditée
                    $"='{nomFeuille}'!ZNIncertGlobale",     // Col 9 : incertitude globale
                    null,                                    // Col 10 : Fréquence finale (valeur historique, laissée vide)
                    $"='{nomFeuille}'!A9",                  // Col 11 : n°Module (depuis A9 de la feuille de mesure)
                    $"='{nomFeuille}'!B9"                   // Col 12 : Fonction (depuis B9 de la feuille de mesure)
                },
                colValeurDirecte: 5,
                valeurDirecte: mesure.SourceMesure == SourceMesure.Generateur
                    ? (object)"Géné."
                    : mesure.FNominale,
                entetesColonne: new[]
                {
                    null, null, null, null, null, null, null, null, null, null,
                    "n°Module",   // Col 11 : header en ligne 5 si pas déjà présent
                    "Fonction"    // Col 12
                }));
        }

        /// <summary>
        /// Ajoute une ligne dans la feuille <c>Récap.</c> du fichier Stabilité, en s'alignant
        /// sur la structure historique (cf. Stab1.xls) :
        /// <list type="bullet">
        ///   <item>8 colonnes : Temps (s), Fréq Moyenne, Écart type, Incertitude, Incert accréditée, Incert globale, Valeurs Maxi, Valeurs Mini</item>
        ///   <item>Insertion en ordre chronologique (gate 1 → L6, gate 2 → L7, …)</item>
        ///   <item>Cols 7-8 = formules locales <c>=C{n}+F{n}</c> et <c>=C{n}-F{n}</c> (moyenne ± incert. globale)</item>
        ///   <item>L1 mise à jour avec NumFI + date</item>
        /// </list>
        /// Le template Stab a déjà la ligne L5 comme template + L19 globaux (Max/Min).
        /// </summary>
        public async Task MettreAJourRecapStabAsync(Mesure mesure)
        {
            if (_workbook == null || _feuilleMesure == null) return;
            string nomFeuille = _feuilleMesure.Name;

            await Task.Run(() =>
            {
                if (!_workbook.Worksheets.Any(w => w.Name == NOM_RECAP)) return;
                var recap = _workbook.Worksheet(NOM_RECAP);
                DeprotegerFeuille(recap);

                // En-tête fiche : NumFI + date — idempotent à chaque itération de gate.
                recap.Cell("B1").SetValue(mesure.NumFI);
                recap.Cell("D1").SetValue(DateTime.Now.ToString("dd/MM/yyyy"));

                // Numéro de la gate dans la séquence (1, 2, 3…). Le nommage des feuilles de
                // mesure étant numérique pur en Stab, on parse directement le nom — sinon
                // fallback : on compte les lignes déjà remplies dans la zone de données.
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
                recap.Cell(ligne, 8).FormulaA1 =
                    $"=IF(ISBLANK(C{ligne}),,IF((C{ligne}-F{ligne})<=0,0,C{ligne}-F{ligne}))";

                // Cols K et L : n°Module et Fonction (lus depuis A9/B9 de la feuille de
                // mesure stab). Pour la Stabilité, la 1ère gate écrit ces valeurs, les
                // gates suivantes les répètent (même session = même module/fonction).
                recap.Cell(ligne, 11).FormulaA1 = $"='{nomFeuille}'!A9";
                recap.Cell(ligne, 12).FormulaA1 = $"='{nomFeuille}'!B9";
                // En-têtes ligne 5 (idempotent : on n'écrase que si la cellule est vide).
                if (recap.Cell(5, 11).IsEmpty()) recap.Cell(5, 11).SetValue("n°Module");
                if (recap.Cell(5, 12).IsEmpty()) recap.Cell(5, 12).SetValue("Fonction");
            });
        }

        /// <summary>
        /// Compte le nombre de lignes de données déjà remplies dans la zone Stab (L6+).
        /// Une ligne est considérée remplie si sa colonne A (Temps) contient quelque chose.
        /// </summary>
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

        // Ligne d'entête des colonnes dans le template Récap. (observé dans METROLOGO.xltm).
        // Toute nouvelle mesure est insérée juste en dessous (row 6) pour avoir "newest on top".
        private const int LIGNE_ENTETE_RECAP = 5;

        /// <summary>
        /// Insère une nouvelle ligne en tête de la zone de données Récap. et la remplit avec
        /// des formules cross-sheet vers la feuille de mesure. Les anciennes mesures descendent
        /// d'une unité ; la plus récente reste ainsi toujours au sommet.
        /// </summary>
        private void EcrireLigneRecap(
            string nomFeuilleMesure,
            string zoneDebZone,
            int ligneFallback,
            string?[] formulesParColonne,
            int colValeurDirecte = -1,
            object? valeurDirecte = null,
            string?[]? entetesColonne = null)
        {
            if (_workbook == null) return;
            if (!_workbook.Worksheets.Any(w => w.Name == NOM_RECAP)) return;

            var recap = _workbook.Worksheet(NOM_RECAP);

            DeprotegerFeuille(recap);

            // 1. Nettoyage des lignes "fantômes" héritées du template : ces lignes contiennent
            //    des formules =[0]!ZNxxx qui pointent vers ModFeuille (toujours vide) et
            //    produisent une plage de zéros sans intérêt juste sous l'entête.
            NettoyerLignesGhost(recap);

            // 2. Insertion de la nouvelle ligne directement sous l'entête — la plus récente en haut.
            int nouvelleLigne = LIGNE_ENTETE_RECAP + 1;
            recap.Row(nouvelleLigne).InsertRowsAbove(1);

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
                    recap.Cell(nouvelleLigne, col).FormulaA1 = formule;
            }

            // 4. En-têtes pour les colonnes au-delà de la structure historique du template
            //    (idempotent : on n'écrase que si la cellule de l'entête est vide). Permet
            //    d'ajouter de nouvelles colonnes — ex: n°Module, Fonction — sans avoir à
            //    modifier le template Excel.
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

        /// <summary>
        /// Supprime les lignes de la zone de données qui ne contiennent que des formules
        /// <c>=[0]!ZNxxx</c> (placeholders du template pointant vers ModFeuille, toujours vide).
        /// Le <c>[0]!</c> est un indicateur sans ambiguïté de ces lignes issues du template xltm.
        /// Itération descendante pour ne pas décaler les indices pendant la suppression.
        /// </summary>
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


        /// <summary>
        /// Remplace les caractères interdits par Windows dans les noms de fichier/dossier
        /// (<c>&lt; &gt; : " / \ | ? *</c> + caractères de contrôle) par un underscore. Utilisé
        /// pour transformer un numéro FI saisi par l'utilisateur en nom de dossier sûr.
        /// </summary>
        private static string SanitizerNomFichier(string nom)
        {
            if (string.IsNullOrWhiteSpace(nom)) return "sans-nom";

            var invalides = new HashSet<char>(Path.GetInvalidFileNameChars());
            var sb = new System.Text.StringBuilder(nom.Length);
            foreach (var c in nom)
            {
                sb.Append(invalides.Contains(c) ? '_' : c);
            }
            string resultat = sb.ToString().Trim(' ', '.');
            return string.IsNullOrEmpty(resultat) ? "sans-nom" : resultat;
        }

        /// <summary>
        /// Réécrit le Target de la relation externe dans le fichier .xlsm pour qu'il
        /// pointe vers le chemin configuré du fichier Metrologo.xla.
        /// </summary>
        /// <summary>
        /// Compte les gates effectivement balayées dans le fichier Stab — basé sur le
        /// nombre de feuilles à nom numérique (<c>1</c>, <c>2</c>, …) car chaque gate
        /// crée sa propre feuille de mesure. Approche plus fiable que de compter les
        /// lignes de la Récap : ClosedXML ne calcule pas les formules, donc le
        /// <c>CachedValue</c> de col A (formule <c>='N'!ZNValGateSecondes</c>) est
        /// systématiquement 0 → tableau systématiquement vide selon ce critère.
        ///
        /// Retourne 0 si pas dans un fichier Stab (autres types n'ont pas de feuilles
        /// numériques, donc protection naturelle).
        /// </summary>
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

        /// <summary>
        /// Patche le graphe Stabilité du fichier .xlsm pour qu'il s'adapte aux mesures :
        ///   • Plages des séries ajustées au nombre réel de gates balayées (au lieu du
        ///     <c>$A$6:$A$15</c> figé du Stab1.xls historique qui forçait 10 lignes — les
        ///     lignes vides étaient lues comme 0, l'axe Y log refusait → graphe vide).
        ///   • Bornes <c>min</c>/<c>max</c> de l'axe Y retirées : Excel les calcule auto
        ///     à chaque ouverture en fonction des valeurs réelles. L'historique fixait
        ///     <c>[1E-12, 1E-08]</c> qui marchait pour des écart-types Allan ~1E-9, mais
        ///     toute mesure hors plage (signal mal calibré, gamme différente, etc.)
        ///     sortait du cadre.
        ///
        /// Modifie UNIQUEMENT <c>xl/charts/chart1.xml</c> — pas les cellules ni les zones
        /// nommées de la Récap. CEAO/SDAO consomme les données via les zones nommées
        /// (<c>ZNRecapS_*</c>), pas le graphe → impact zéro.
        /// </summary>
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

                // 1. Adapte les plages des séries au nb de gates : $X$6:$X$<N> → $X$6:$X$<5+nbGates>.
                // Le `\d+` (au lieu de `15`) permet de re-patcher après un patch précédent
                // (à chaque gate du balayage stab, le code repasse pour ajuster la plage).
                nouveau = System.Text.RegularExpressions.Regex.Replace(
                    nouveau,
                    @"\$([A-Z]+)\$6:\$([A-Z]+)\$\d+",
                    m => $"${m.Groups[1].Value}$6:${m.Groups[2].Value}${derniereLigne}");

                // 2. Retire <c:min .../> et <c:max .../> à l'intérieur des <c:valAx>...</c:valAx>
                //    (axe Y) pour qu'Excel les recalcule auto. L'axe X (catAx) n'a pas de
                //    bornes fixes dans le template historique — non concerné.
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

        /// <summary>
        /// Force Excel à recalculer toutes les formules ET à régénérer les caches des graphes
        /// à l'ouverture, en injectant <c>fullCalcOnLoad="1"</c> dans <c>xl/workbook.xml</c>.
        ///
        /// Pourquoi : ClosedXML écrit correctement les cellules de la Récap. mais ne touche
        /// pas aux <c>numCache</c> stockés dans les <c>xl/charts/chartN.xml</c> — Excel les
        /// utilise alors tels quels (figés à 0) au lieu de relire les cellules. Sur le graphe
        /// Stabilité (axe Y log), des zéros déclenchent le popup « Impossible de représenter
        /// les valeurs nulles ou négatives sur des graphiques logarithmiques » et la zone
        /// de traçage reste vide.
        ///
        /// Solution : <c>fullCalcOnLoad="1"</c> dit à Excel de tout recalculer à la prochaine
        /// ouverture du fichier (formules ET caches de graphes), ~50 ms d'overhead à
        /// l'ouverture, négligeable.
        /// </summary>
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

            // Le workbook peut utiliser un préfixe XML namespace (ex: <x:workbook>, <x:calcPr>).
            // Les regex acceptent donc un éventuel préfixe via (\w+:)?.
            //
            // 3 cas :
            //   1. <[ns:]calcPr ... fullCalcOnLoad="X" /> existant → remplacer X par 1.
            //   2. <[ns:]calcPr ... /> sans fullCalcOnLoad → ajouter l'attribut.
            //   3. pas de <[ns:]calcPr> → en injecter un avant </[ns:]workbook>, en
            //      reprenant le préfixe utilisé pour <[ns:]workbook>.
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
                // Détection du préfixe utilisé sur la balise <workbook> pour l'appliquer à <calcPr>.
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

        /// <summary>
        /// Supprime les blocs <c>&lt;c:numCache&gt;...&lt;/c:numCache&gt;</c> et
        /// <c>&lt;c:strCache&gt;...&lt;/c:strCache&gt;</c> de tous les <c>xl/charts/chartN.xml</c>.
        ///
        /// Pourquoi : les caches du graphe Stab héritent du Stab1.xls (10 valeurs Y figées
        /// à 0). Sans nettoyage, Excel à l'ouverture utilise ces zéros pour auto-scaler
        /// l'axe Y log → l'axe descend à 1E-308 (auto-fit sur valeurs nulles) et l'échelle
        /// est démesurément grande par rapport aux mesures réelles.
        ///
        /// Sans cache, Excel est obligé de recalculer les séries depuis les cellules — qui
        /// ont elles-mêmes été recalculées via <c>fullCalcOnLoad="1"</c>. Le ptCount va aussi
        /// être recalculé automatiquement par Excel.
        /// </summary>
        /// <summary>
        /// Force les bornes <c>min</c>/<c>max</c> de l'axe Y log du graphe Stab à partir
        /// d'une estimation des sigmas relatifs par gate. Évite l'auto-scale d'Excel qui,
        /// sans cache et sans bornes, choisit parfois 1E-9..1E0 (5 décades en trop).
        ///
        /// Borne min = 10^floor(log10(min_sigma) - 0.5)  (= ~3× en dessous, arrondi décade).
        /// Borne max = 10^ceil(log10(max_sigma) + 0.5)   (= ~3× au-dessus, arrondi décade).
        /// </summary>
        private static void CalibrerAxeYGrapheStab(string xlsmPath, IEnumerable<double> sigmas)
        {
            var positifs = sigmas.Where(s => s > 0).ToList();
            if (positifs.Count == 0) return;

            double minS = positifs.Min();
            double maxS = positifs.Max();
            double minBorne = Math.Pow(10, Math.Floor(Math.Log10(minS) - 0.5));
            double maxBorne = Math.Pow(10, Math.Ceiling(Math.Log10(maxS) + 0.5));

            // Si toutes les sigmas sont sur la même décade, on garantit au moins 1 décade
            // d'écart pour un graphe lisible.
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

                // Réinjecte min/max dans <c:scaling> à l'intérieur de <c:valAx>.
                // Si le scaling contient déjà min/max (peu probable après ViderCaches +
                // RendreGrapheStabDynamique qui les ont retirés), on les remplace.
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

        private static void PatcherLienMacroXLA(string xlsmPath, string xlaPath)
        {
            const string relPath = "xl/externalLinks/_rels/externalLink1.xml.rels";

            // URI file:/// avec backslashes → slashes pour compatibilité XML
            string uri = "file:///" + xlaPath.Replace('\\', '/');

            using var zip = ZipFile.Open(xlsmPath, ZipArchiveMode.Update);
            var entry = zip.GetEntry(relPath);
            if (entry == null) return;

            string contenu;
            using (var reader = new StreamReader(entry.Open()))
            {
                contenu = reader.ReadToEnd();
            }

            // Remplace chaque Target="...Metrologo.xla" par le chemin configuré
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

        // Mots de passe connus appliqués à ModFeuille dans les templates métier.
        private static readonly string[] _motsDePasseFeuille = { "METROL", "metrol" };

        private static void DeprotegerFeuille(IXLWorksheet feuille)
        {
            try { feuille.Unprotect(); return; }
            catch { /* feuille protégée par mot de passe — on essaie ci-dessous */ }

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

        /// <summary>
        /// Supprime les feuilles à nom numérique (« 1 », « 2 », …) du template Stab — qui
        /// proviennent du <c>Stab1.xls</c> historique (10 procédures auto figées du Delphi)
        /// — ainsi que les formules de la Récap. qui les référençaient.
        ///
        /// À n'appeler QUE lorsqu'on part du template (= nouveau fichier). Sur un fichier
        /// existant, ces feuilles contiennent les mesures réelles d'une session précédente
        /// — il ne faut jamais les supprimer.
        /// </summary>
        private void NettoyerFeuillesNumeriquesResiduellesEtRecap()
        {
            if (_workbook == null) return;

            // 1. Suppression des feuilles dont le nom est un entier (cas template Stab :
            //    "1".."10" en général). On clone la liste avant pour itérer sans collision
            //    avec la modification de la collection Worksheets.
            var aSupprimer = _workbook.Worksheets
                .Where(w => int.TryParse(w.Name, out _))
                .Select(w => w.Name)
                .ToList();

            foreach (var nom in aSupprimer)
            {
                try { _workbook.Worksheet(nom).Delete(); }
                catch { /* feuille déjà partie ou nom invalide — non bloquant */ }
            }

            // 2. Effacement des formules pré-remplies de la Récap. qui pointaient vers
            //    ces feuilles supprimées (lignes 6 à 15, colonnes A à H = ValGate / Moy /
            //    Sigma / IncSigma / IncAccr / IncGlob / Total / Stab). Sans cet effacement,
            //    les cellules afficheraient #REF! ou des 0 fantômes après suppression des
            //    feuilles cibles. Récap. sera repeuplée à chaque gate via MettreAJourRecapStabAsync.
            if (_workbook.Worksheets.Any(w => w.Name == NOM_RECAP))
            {
                var recap = _workbook.Worksheet(NOM_RECAP);
                DeprotegerFeuille(recap);
                recap.Range("A6:H15").Clear(XLClearOptions.Contents);
            }

            // 3. Les zones nommées workbook-scope qui référençaient ces feuilles (ex.
            //    _1ZNDeltaF, _10ZNPlageFReelle) deviennent orphelines mais ne sont pas
            //    consommées par le graphe (qui pointe vers Récap. directement) — on les
            //    laisse en place ; ClosedXML ne plante pas dessus.
        }

        private string TrouverNomFeuilleUnique(TypeMesure type)
        {
            if (type == TypeMesure.FreqAvantInterv) return "F_Avant_Interv";
            if (type == TypeMesure.FreqFinale) return "F_Finale";

            // Stabilité : nommage numérique pur (1, 2, 3…) — aligné sur la convention
            // historique du Delphi/Stab1.xls. La feuille Récap. attend ce format pour
            // ses formules cross-sheet ='1'!ZN…, ='2'!ZN…, etc.
            if (type == TypeMesure.Stabilite)
            {
                int n = 1;
                var existantsStab = _workbook!.Worksheets.Select(w => w.Name)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);
                while (existantsStab.Contains(n.ToString())) n++;
                return n.ToString();
            }

            string prefixe = type switch
            {
                TypeMesure.Frequence => "Freq",
                TypeMesure.Interval => "Interv",
                TypeMesure.TachyContact => "TachyC",
                TypeMesure.Stroboscope => "Strobo",
                _ => "Mesure"
            };

            int idx = 1;
            var existants = _workbook!.Worksheets.Select(w => w.Name)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            while (existants.Contains($"{prefixe}{idx}")) idx++;
            return $"{prefixe}{idx}";
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
