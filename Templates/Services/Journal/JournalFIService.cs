using System;
using System.IO;
using System.Text;

namespace Metrologo.Services.Journal
{
    /// <summary>
    /// Journal utilisateur par FI : un fichier <c>Journal_&lt;FI&gt;.txt</c> est créé dans
    /// le dossier de la FI (à côté des <c>Mesures_*.xlsx</c>) dès que l'utilisateur valide
    /// le numéro FI. Toutes les actions métier (configuration, mesures, relances, arrêts,
    /// échecs) y sont consignées en format texte lisible, séparé du journal système
    /// centralisé sur le réseau (<see cref="Journal"/>) qui reste réservé aux logs techniques.
    ///
    /// Singleton thread-safe — une seule session active à la fois (changement de FI = bascule
    /// vers un nouveau fichier après écriture FIN_SESSION sur l'ancien).
    /// </summary>
    public static class JournalFIService
    {
        private static readonly object _sync = new();
        private static string _cheminFichier = string.Empty;
        private static string _numFICourant = string.Empty;
        private static string _utilisateurCourant = string.Empty;
        private static string _posteCourant = string.Empty;
        private static DateTime _debutSession = DateTime.MinValue;

        // Compteurs cumulés sur la session (pour le récap final FIN_SESSION).
        private static int _nbMesuresEffectuees;
        private static int _nbMesuresEchouees;

        /// <summary>Numéro FI de la session en cours (vide si aucune session active).</summary>
        public static string NumFICourant
        {
            get { lock (_sync) { return _numFICourant; } }
        }

        /// <summary>Vrai si une session journal FI est actuellement ouverte.</summary>
        public static bool EstActif
        {
            get { lock (_sync) { return !string.IsNullOrEmpty(_cheminFichier); } }
        }

        /// <summary>Chemin du fichier journal de la session en cours (vide si aucune).</summary>
        public static string CheminFichier
        {
            get { lock (_sync) { return _cheminFichier; } }
        }

        /// <summary>
        /// Démarre une session journal pour la FI donnée. Crée le dossier de la FI si
        /// absent et le fichier <c>Journal_&lt;FI&gt;.txt</c>. Si une session était déjà
        /// ouverte sur une autre FI, elle est terminée (FIN_SESSION écrit) avant la bascule.
        /// Idempotent si la FI demandée est déjà la session courante.
        /// </summary>
        public static void DemarrerSession(string numFI, string utilisateur, string poste)
        {
            if (string.IsNullOrWhiteSpace(numFI)) return;

            lock (_sync)
            {
                // Bascule depuis une autre FI ? On ferme proprement la précédente.
                if (!string.IsNullOrEmpty(_cheminFichier)
                    && !string.Equals(_numFICourant, numFI, StringComparison.OrdinalIgnoreCase))
                {
                    TerminerSessionInterne("Changement de FI");
                }
                // Même FI mais l'opérateur (ou le poste) a changé — ex. un collègue reprend la
                // FI en se reconnectant sans relancer l'app. On clôt le bloc courant et on en
                // ouvre un nouveau, pour que le journal reflète bien qui a travaillé sur la FI.
                else if (!string.IsNullOrEmpty(_cheminFichier)
                    && string.Equals(_numFICourant, numFI, StringComparison.OrdinalIgnoreCase)
                    && (!string.Equals(_utilisateurCourant, utilisateur, StringComparison.Ordinal)
                        || !string.Equals(_posteCourant, poste, StringComparison.Ordinal)))
                {
                    TerminerSessionInterne("Changement d'utilisateur");
                }
                // Déjà sur cette FI avec le même opérateur ? Rien à faire.
                if (string.Equals(_numFICourant, numFI, StringComparison.OrdinalIgnoreCase)
                    && !string.IsNullOrEmpty(_cheminFichier))
                {
                    return;
                }

                try
                {
                    string numFISafe = SanitizerNomFichier(numFI);
                    string dossier = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                        "Metrologo", numFISafe);
                    Directory.CreateDirectory(dossier);
                    _cheminFichier = Path.Combine(dossier, $"Journal_{numFISafe}.txt");
                    _numFICourant = numFI;
                    _utilisateurCourant = utilisateur ?? string.Empty;
                    _posteCourant = poste ?? string.Empty;
                    _debutSession = DateTime.Now;
                    _nbMesuresEffectuees = 0;
                    _nbMesuresEchouees = 0;

                    // Si le fichier existe déjà (ouvertures FI précédentes), on append un
                    // séparateur de session — on perd pas l'historique des sessions précédentes.
                    bool fichierExistant = File.Exists(_cheminFichier);
                    var sb = new StringBuilder();
                    if (fichierExistant)
                    {
                        sb.AppendLine();
                        sb.AppendLine();
                    }
                    sb.AppendLine("=================================================");
                    sb.AppendLine($"Journal utilisateur — FI {numFI}");
                    sb.AppendLine("=================================================");
                    sb.AppendLine($"Utilisateur     : {utilisateur}");
                    sb.AppendLine($"Poste           : {poste}");
                    sb.AppendLine($"Début session   : {_debutSession:dd/MM/yyyy HH:mm:ss}");
                    sb.AppendLine();

                    File.AppendAllText(_cheminFichier, sb.ToString(), Encoding.UTF8);
                    Ecrire("FI_OUVERTE", numFI);
                }
                catch (Exception ex)
                {
                    // Best-effort : si l'écriture échoue (disque plein, droits, etc.), on
                    // logue dans le journal central et on désactive cette session FI.
                    Journal.Warn(CategorieLog.Systeme, "JOURNAL_FI_DEMARRAGE_KO",
                        $"Impossible de créer le journal FI pour {numFI} : {ex.Message}");
                    _cheminFichier = string.Empty;
                    _numFICourant = string.Empty;
                }
            }
        }

        /// <summary>
        /// Append un événement au journal de la session courante. No-op si aucune session
        /// n'est active (ex: utilisateur n'a pas encore validé de FI). Thread-safe — peut
        /// être appelé depuis n'importe quel thread (UI, mesure GPIB, etc.).
        /// </summary>
        public static void Ecrire(string typeAction, string detail = "")
        {
            lock (_sync)
            {
                if (string.IsNullOrEmpty(_cheminFichier)) return;

                // Compteurs métier
                if (typeAction == "MESURE_FIN") _nbMesuresEffectuees++;
                else if (typeAction == "MESURE_ECHEC") _nbMesuresEchouees++;

                try
                {
                    string timestamp = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                    // typeAction sur 22 chars max (alignement colonne), détail à la suite.
                    string typePadded = typeAction.Length > 22
                        ? typeAction.Substring(0, 22)
                        : typeAction.PadRight(22);
                    string ligne = string.IsNullOrEmpty(detail)
                        ? $"[{timestamp}]  {typePadded}\n"
                        : $"[{timestamp}]  {typePadded}{detail}\n";
                    File.AppendAllText(_cheminFichier, ligne, Encoding.UTF8);
                    // PLUS de duplication temps-réel : le transfert vers le réseau est fait en
                    // BLOC à la fin de chaque mesure via TransfertReseauService.TransfererDossierFIAsync.
                }
                catch (Exception ex)
                {
                    // Silencieux : ne pas faire planter une mesure si le journal échoue.
                    // On logue dans le journal central pour traçabilité.
                    Journal.Warn(CategorieLog.Systeme, "JOURNAL_FI_ECRITURE_KO",
                        $"Écriture journal FI {_numFICourant} échouée : {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Termine la session courante en écrivant une ligne FIN_SESSION avec récap
        /// (nombre de mesures effectuées / échouées, durée totale). Idempotent — appelable
        /// plusieurs fois sans erreur (no-op à partir du 2e appel).
        /// </summary>
        public static void TerminerSession(string motif = "")
        {
            lock (_sync) { TerminerSessionInterne(motif); }
        }

        // ---- internes (déjà sous lock) ----

        private static void TerminerSessionInterne(string motif)
        {
            if (string.IsNullOrEmpty(_cheminFichier)) return;
            try
            {
                TimeSpan duree = DateTime.Now - _debutSession;
                string dureeFmt = duree.TotalHours >= 1
                    ? $"{(int)duree.TotalHours}h{duree.Minutes:D2}min"
                    : $"{(int)duree.TotalMinutes} min {duree.Seconds:D2} s";

                string detail = $"{_nbMesuresEffectuees} mesure(s) effectuée(s)";
                if (_nbMesuresEchouees > 0) detail += $", {_nbMesuresEchouees} échec(s)";
                detail += $" · durée totale {dureeFmt}";
                if (!string.IsNullOrEmpty(motif)) detail += $" · {motif}";

                string timestamp = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                string ligne = $"[{timestamp}]  {"FIN_SESSION".PadRight(22)}{detail}\n";
                File.AppendAllText(_cheminFichier, ligne, Encoding.UTF8);
                // PLUS de duplication temps-réel : transfert différé via TransfertReseauService.
            }
            catch (Exception ex)
            {
                Journal.Warn(CategorieLog.Systeme, "JOURNAL_FI_FIN_KO",
                    $"Écriture FIN_SESSION journal FI {_numFICourant} échouée : {ex.Message}");
            }
            finally
            {
                _cheminFichier = string.Empty;
                _numFICourant = string.Empty;
                _utilisateurCourant = string.Empty;
                _posteCourant = string.Empty;
                _debutSession = DateTime.MinValue;
                _nbMesuresEffectuees = 0;
                _nbMesuresEchouees = 0;
            }
        }

        // Méthode DupliquerSurReseauAsync RETIRÉE — la duplication réseau temps-réel par
        // ligne est remplacée par un transfert en bloc du dossier FI à la fin de chaque
        // mesure (cf. TransfertReseauService.TransfererDossierFIAsync).

        private static string SanitizerNomFichier(string nom)
        {
            if (string.IsNullOrWhiteSpace(nom)) return "sans-nom";
            var invalides = new System.Collections.Generic.HashSet<char>(Path.GetInvalidFileNameChars());
            var sb = new StringBuilder(nom.Length);
            foreach (var c in nom)
            {
                sb.Append(invalides.Contains(c) ? '_' : c);
            }
            string resultat = sb.ToString().Trim(' ', '.');
            return string.IsNullOrEmpty(resultat) ? "sans-nom" : resultat;
        }
    }
}
