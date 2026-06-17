using System;
using System.IO;
using System.Text;

namespace Metrologo.Services.Journal
{
    /// <summary>
    /// Journal utilisateur, un par FI : dès que l'opérateur valide le numéro de FI, on crée un
    /// fichier <c>Journal_&lt;FI&gt;.txt</c> dans le dossier de la FI (à côté des
    /// <c>Mesures_*.xlsx</c>). Toutes les actions métier — configuration, mesures, relances,
    /// arrêts, échecs — y sont notées en texte lisible. C'est volontairement distinct du journal
    /// système centralisé sur le réseau (<see cref="Journal"/>), qui lui reste pour les logs techniques.
    ///
    /// Singleton thread-safe : une seule session ouverte à la fois. Changer de FI bascule vers un
    /// nouveau fichier, après avoir écrit FIN_SESSION dans l'ancien.
    /// </summary>
    public static class JournalFIService
    {
        private static readonly object _sync = new();
        private static string _cheminFichier = string.Empty;
        private static string _numFICourant = string.Empty;
        private static string _utilisateurCourant = string.Empty;
        private static string _posteCourant = string.Empty;
        private static DateTime _debutSession = DateTime.MinValue;

        // Compteurs cumulés sur la durée de la session (servent au récap final FIN_SESSION).
        private static int _nbMesuresEffectuees;
        private static int _nbMesuresEchouees;

        /// <summary>Numéro de FI de la session en cours (vide si aucune session n'est ouverte).</summary>
        public static string NumFICourant
        {
            get { lock (_sync) { return _numFICourant; } }
        }

        /// <summary>Vrai s'il y a une session de journal FI ouverte en ce moment.</summary>
        public static bool EstActif
        {
            get { lock (_sync) { return !string.IsNullOrEmpty(_cheminFichier); } }
        }

        /// <summary>Chemin du fichier journal de la session courante (vide s'il n'y en a pas).</summary>
        public static string CheminFichier
        {
            get { lock (_sync) { return _cheminFichier; } }
        }

        /// <summary>
        /// Ouvre une session de journal pour la FI indiquée. Crée au besoin le dossier de la FI
        /// et le fichier <c>Journal_&lt;FI&gt;.txt</c>. Si une session était déjà ouverte sur une
        /// autre FI, on la clôt proprement (FIN_SESSION) avant de basculer. Si la FI demandée est
        /// déjà la session en cours, l'appel ne fait rien.
        /// </summary>
        public static void DemarrerSession(string numFI, string utilisateur, string poste)
        {
            if (string.IsNullOrWhiteSpace(numFI)) return;

            lock (_sync)
            {
                // On vient d'une autre FI ? On clôt proprement la précédente avant de continuer.
                if (!string.IsNullOrEmpty(_cheminFichier)
                    && !string.Equals(_numFICourant, numFI, StringComparison.OrdinalIgnoreCase))
                {
                    TerminerSessionInterne("Changement de FI");
                }
                // Même FI, mais l'opérateur (ou le poste) a changé — typiquement un collègue qui
                // reprend la FI en se reconnectant sans relancer l'app. On ferme le bloc en cours et
                // on en rouvre un, pour que le journal montre bien qui a travaillé sur la FI.
                else if (!string.IsNullOrEmpty(_cheminFichier)
                    && string.Equals(_numFICourant, numFI, StringComparison.OrdinalIgnoreCase)
                    && (!string.Equals(_utilisateurCourant, utilisateur, StringComparison.Ordinal)
                        || !string.Equals(_posteCourant, poste, StringComparison.Ordinal)))
                {
                    TerminerSessionInterne("Changement d'utilisateur");
                }
                // Déjà sur cette FI avec le même opérateur ? Alors il n'y a rien à faire.
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

                    // Si le fichier existe déjà (la FI a été ouverte par le passé), on ajoute juste
                    // un séparateur de session — comme ça on garde l'historique des sessions d'avant.
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
                    // Au mieux : si l'écriture échoue (disque plein, droits manquants...), on
                    // le signale dans le journal central et on désactive cette session FI.
                    Journal.Warn(CategorieLog.Systeme, "JOURNAL_FI_DEMARRAGE_KO",
                        $"Impossible de créer le journal FI pour {numFI} : {ex.Message}");
                    _cheminFichier = string.Empty;
                    _numFICourant = string.Empty;
                }
            }
        }

        /// <summary>
        /// Ajoute un événement au journal de la session courante. Ne fait rien s'il n'y a pas de
        /// session ouverte (par ex. l'utilisateur n'a pas encore validé de FI). Thread-safe : on
        /// peut l'appeler depuis n'importe quel thread (UI, mesure GPIB, etc.).
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
                    // typeAction limité à 22 caractères (pour aligner la colonne), puis le détail.
                    string typePadded = typeAction.Length > 22
                        ? typeAction.Substring(0, 22)
                        : typeAction.PadRight(22);
                    string ligne = string.IsNullOrEmpty(detail)
                        ? $"[{timestamp}]  {typePadded}\n"
                        : $"[{timestamp}]  {typePadded}{detail}\n";
                    File.AppendAllText(_cheminFichier, ligne, Encoding.UTF8);
                    // Fini la duplication en temps réel : le transfert vers le réseau se fait
                    // désormais en BLOC à la fin de chaque mesure, via
                    // TransfertReseauService.TransfererDossierFIAsync.
                }
                catch (Exception ex)
                {
                    // En silence : pas question de faire planter une mesure parce que le journal
                    // a échoué. On garde une trace dans le journal central, c'est tout.
                    Journal.Warn(CategorieLog.Systeme, "JOURNAL_FI_ECRITURE_KO",
                        $"Écriture journal FI {_numFICourant} échouée : {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Clôt la session courante en écrivant une ligne FIN_SESSION avec un petit récap
        /// (mesures réussies / échouées, durée totale). On peut l'appeler plusieurs fois sans
        /// problème : à partir du deuxième appel, ça ne fait rien.
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
                // Plus de duplication en temps réel : le transfert est différé via TransfertReseauService.
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

        // Méthode DupliquerSurReseauAsync supprimée : on ne duplique plus ligne par ligne en
        // temps réel sur le réseau ; on transfère tout le dossier FI en une fois à la fin de
        // chaque mesure (voir TransfertReseauService.TransfererDossierFIAsync).

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
