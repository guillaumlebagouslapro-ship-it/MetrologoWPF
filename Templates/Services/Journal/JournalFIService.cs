using System;
using System.IO;
using System.Text;

namespace Metrologo.Services.Journal
{
    /// <summary>
    /// Journal utilisateur, un par FI : <c>Journal_&lt;FI&gt;.txt</c> dans le dossier de la FI (à côté
    /// des <c>Mesures_*.xlsx</c>). Un bandeau par session, puis des lignes CONCISES :
    /// <code>
    /// =================================================
    /// Journal utilisateur — FI 12345
    /// =================================================
    /// Utilisateur     : Jean Dupont
    /// Poste           : Baie
    /// Début session   : 14/10/2026 09:00:12
    ///
    /// --- Fréquence -----------------------------------
    /// 09:00  Config    30 mesures · 53131A · gate 1 s
    /// 09:01  ✔ Mesure  30 mesures
    /// 09:05  ✖ Relance (fichier neuf)  arrêtée par l'opérateur
    ///
    /// --- Stabilité -----------------------------------
    /// 10:02  Config    30 mesures par gate · 53131A · toutes les gates
    /// 10:14  ✔ Stab 1/3  480 mesures
    ///
    /// 11:00  Fin       2 mesures OK, 1 échec · 2 h 00 · Fermeture application
    /// </code>
    /// Un séparateur au nom du type de mesure s'insère à chaque changement de type.
    /// Une ligne par mesure avec son résultat (qui / quoi / quand ; les valeurs sont dans le
    /// rapport Excel), la config seulement quand elle change. Le détail technique reste dans le
    /// journal système centralisé (<see cref="Journal"/>).
    ///
    /// Singleton thread-safe : une seule session ouverte à la fois. Changer de FI (ou d'opérateur)
    /// clôt la session en cours avant d'en ouvrir une nouvelle.
    /// </summary>
    public static class JournalFIService
    {
        /// <summary>Début des lignes d'en-tête de session (et des changements de jour).</summary>
        public const string DebutEntete = "── ";
        /// <summary>Marque d'une mesure réussie (sert aussi au comptage dans la liste des FI).</summary>
        public const string MarqueOk = "✔";
        public const string MarqueEchec = "✖";

        private static readonly object _sync = new();
        private static string _cheminFichier = string.Empty;
        private static string _numFICourant = string.Empty;
        private static string _utilisateurCourant = string.Empty;
        private static string _posteCourant = string.Empty;
        private static DateTime _debutSession = DateTime.MinValue;
        private static DateTime _dernierJourEcrit = DateTime.MinValue;
        private static string _derniereConfig = string.Empty;
        private static string _dernierType = string.Empty;
        /// <summary>Largeur des lignes de séparation (celle du bandeau de session).</summary>
        private const int LargeurSeparateur = 49;
        private static string _libelleMesure = "Mesure";
        // Mesure démarrée mais pas encore notée, et arrêt demandé par l'opérateur pendant celle-ci.
        private static bool _mesureEnAttente;
        private static bool _arretDemande;

        // Compteurs de la session (récap de la ligne « Fin »).
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
        /// Ouvre une session de journal pour la FI indiquée (crée au besoin le dossier et le fichier).
        /// Si une session était ouverte sur une autre FI, ou par un autre opérateur / poste, on la clôt
        /// d'abord. Même FI, même opérateur : ne fait rien.
        /// </summary>
        public static void DemarrerSession(string numFI, string utilisateur, string poste)
        {
            if (string.IsNullOrWhiteSpace(numFI)) return;

            lock (_sync)
            {
                if (!string.IsNullOrEmpty(_cheminFichier))
                {
                    bool memeFI = string.Equals(_numFICourant, numFI, StringComparison.OrdinalIgnoreCase);
                    bool memeOperateur = string.Equals(_utilisateurCourant, utilisateur, StringComparison.Ordinal)
                                      && string.Equals(_posteCourant, poste, StringComparison.Ordinal);
                    if (memeFI && memeOperateur) return;
                    TerminerSessionInterne(memeFI ? "Changement d'utilisateur" : "Changement de FI");
                }

                try
                {
                    // Même dossier que le rapport : réseau si joignable, sinon Bureau\Metrologo\FI.
                    string dossier = CheminsMetrologo.DossierFI(numFI);
                    Directory.CreateDirectory(dossier);
                    _cheminFichier = Path.Combine(dossier, $"Journal_{CheminsMetrologo.NomDossierFI(numFI)}.txt");
                    _numFICourant = numFI;
                    _utilisateurCourant = utilisateur ?? string.Empty;
                    _posteCourant = poste ?? string.Empty;
                    _debutSession = DateTime.Now;
                    _dernierJourEcrit = _debutSession.Date;
                    _derniereConfig = string.Empty;
                    _dernierType = string.Empty;
                    _libelleMesure = "Mesure";
                    _mesureEnAttente = false;
                    _arretDemande = false;
                    _nbMesuresEffectuees = 0;
                    _nbMesuresEchouees = 0;

                    // Bandeau de session (format historique). Si le fichier existe déjà (FI ouverte par
                    // le passé), deux lignes vides le séparent de la session précédente, qu'on garde.
                    var sb = new StringBuilder();
                    if (File.Exists(_cheminFichier))
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
        /// Note la configuration validée, seulement si elle diffère de la précédente de la session.
        /// Changement de type de mesure (Fréquence → Stabilité…) : un séparateur au nom du type la
        /// précède, et la ligne « Config » ne répète donc pas le type.
        /// </summary>
        public static void NoterConfiguration(string typeMesure, string resume)
        {
            lock (_sync)
            {
                if (string.IsNullOrEmpty(_cheminFichier)) return;
                bool nouveauType = typeMesure != _dernierType;
                if (!nouveauType && resume == _derniereConfig) return;

                if (nouveauType)
                {
                    // Le bandeau se termine déjà par une ligne vide : pas de seconde pour le 1er type.
                    string titre = $"--- {typeMesure} ";
                    EcrireBrut((_dernierType.Length > 0 ? Environment.NewLine : string.Empty)
                        + titre + new string('-', Math.Max(3, LargeurSeparateur - titre.Length))
                        + Environment.NewLine);
                    _dernierType = typeMesure;
                }
                _derniereConfig = resume;
                AjouterLigne("Config   ", resume);
            }
        }

        /// <summary>
        /// Mémorise le libellé de la mesure qui démarre (« Lancement », « Relance (fichier neuf) »,
        /// « Stab 2/5 »…). Rien n'est écrit : la ligne part avec le résultat, dans <see cref="NoterResultat"/>.
        /// </summary>
        public static void DebutMesure(string preambule)
        {
            lock (_sync)
            {
                // Le préambule commence par un pictogramme (▶, 🔁) : on part de la première lettre.
                string p = preambule ?? string.Empty;
                int i = 0;
                while (i < p.Length && !char.IsLetterOrDigit(p[i])) i++;
                string libelle = p[i..].Trim();

                // La mesure précédente s'est terminée sans résultat noté (chemin d'abandon) : on la
                // note maintenant pour ne pas perdre la trace.
                if (_mesureEnAttente) NoterResultatInterne(false, "interrompue");

                _libelleMesure = libelle.Length == 0 || libelle == "Lancement" ? "Mesure" : libelle;
                _mesureEnAttente = true;
                _arretDemande = false;
            }
        }

        /// <summary>Clic sur STOP : la ligne de la mesure dira « arrêtée par l'opérateur » au lieu
        /// du message d'erreur technique qui découle de l'annulation.</summary>
        public static void SignalerArret()
        {
            lock (_sync) { _arretDemande = true; }
        }

        /// <summary>Une ligne par mesure terminée : ✔ si elle a abouti, ✖ sinon (avec la raison).</summary>
        public static void NoterResultat(bool ok, string detail)
        {
            lock (_sync) { NoterResultatInterne(ok, detail); }
        }

        private static void NoterResultatInterne(bool ok, string detail)
        {
            if (string.IsNullOrEmpty(_cheminFichier) || !_mesureEnAttente) return;
            if (ok) _nbMesuresEffectuees++; else _nbMesuresEchouees++;
            if (!ok && _arretDemande) detail = "arrêtée par l'opérateur";
            AjouterLigne($"{(ok ? MarqueOk : MarqueEchec)} {_libelleMesure} ", detail);
            _libelleMesure = "Mesure";
            _mesureEnAttente = false;
            _arretDemande = false;
        }

        /// <summary>Début / fin d'une série de stabs répétées.</summary>
        public static void NoterSerie(string texte)
        {
            lock (_sync)
            {
                if (string.IsNullOrEmpty(_cheminFichier)) return;
                AjouterLigne("Série    ", texte);
            }
        }

        /// <summary>
        /// Clôt la session courante par une ligne « Fin » (mesures réussies / échouées, durée, motif).
        /// Appelable plusieurs fois : à partir du deuxième appel, ne fait rien.
        /// </summary>
        public static void TerminerSession(string motif = "")
        {
            lock (_sync) { TerminerSessionInterne(motif); }
        }

        // ---- internes (déjà sous lock) ----

        /// <summary>Le journal suit le dossier ACTIF de la FI (passage local → réseau après un
        /// rapatriement) : sans ça, il recréerait un dossier sur le Bureau.</summary>
        private static void SuivreDossierFI()
        {
            if (string.IsNullOrEmpty(_cheminFichier) || string.IsNullOrEmpty(_numFICourant)) return;
            try
            {
                _cheminFichier = Path.Combine(CheminsMetrologo.DossierFI(_numFICourant),
                    Path.GetFileName(_cheminFichier));
            }
            catch { /* on garde le chemin connu */ }
        }

        /// <summary>Texte écrit tel quel (séparateurs).</summary>
        private static void EcrireBrut(string texte)
        {
            SuivreDossierFI();
            try { File.AppendAllText(_cheminFichier, texte, Encoding.UTF8); }
            catch (Exception ex)
            {
                Journal.Warn(CategorieLog.Systeme, "JOURNAL_FI_ECRITURE_KO",
                    $"Écriture journal FI {_numFICourant} échouée : {ex.Message}");
            }
        }

        /// <summary>« HH:mm  libellé  détail ». Si la session déborde sur un autre jour (stab de
        /// nuit), une ligne de date est insérée pour que les heures restent lisibles.</summary>
        private static void AjouterLigne(string libelle, string detail, bool ligneVideAvant = false)
        {
            SuivreDossierFI();
            try
            {
                var maintenant = DateTime.Now;
                var sb = new StringBuilder();
                if (ligneVideAvant) sb.AppendLine();
                if (maintenant.Date != _dernierJourEcrit)
                {
                    sb.Append(DebutEntete).Append(maintenant.ToString("dd/MM/yyyy")).AppendLine(" ──");
                    _dernierJourEcrit = maintenant.Date;
                }
                string d = (detail ?? string.Empty).Replace("\r", " ").Replace("\n", " ").Trim();
                sb.Append(maintenant.ToString("HH:mm")).Append("  ").Append(libelle);
                if (d.Length > 0) sb.Append(' ').Append(d);
                sb.AppendLine();
                File.AppendAllText(_cheminFichier, sb.ToString(), Encoding.UTF8);
                // Pas de duplication en temps réel : le dossier FI part sur le réseau en bloc à la
                // fin de chaque mesure (TransfertReseauService.TransfererDossierFIAsync).
            }
            catch (Exception ex)
            {
                // En silence : pas question de faire planter une mesure parce que le journal
                // a échoué. On garde une trace dans le journal central, c'est tout.
                Journal.Warn(CategorieLog.Systeme, "JOURNAL_FI_ECRITURE_KO",
                    $"Écriture journal FI {_numFICourant} échouée : {ex.Message}");
            }
        }

        private static void TerminerSessionInterne(string motif)
        {
            if (string.IsNullOrEmpty(_cheminFichier)) return;
            try
            {
                if (_mesureEnAttente) NoterResultatInterne(false, "interrompue");

                TimeSpan duree = DateTime.Now - _debutSession;
                string dureeFmt = duree.TotalHours >= 1
                    ? $"{(int)duree.TotalHours} h {duree.Minutes:D2}"
                    : $"{Math.Max(1, (int)Math.Round(duree.TotalMinutes))} min";

                string detail = $"{_nbMesuresEffectuees} mesure{(_nbMesuresEffectuees > 1 ? "s" : "")} OK";
                if (_nbMesuresEchouees > 0)
                    detail += $", {_nbMesuresEchouees} échec{(_nbMesuresEchouees > 1 ? "s" : "")}";
                detail += $" · {dureeFmt}";
                if (!string.IsNullOrEmpty(motif)) detail += $" · {motif}";

                AjouterLigne("Fin      ", detail, ligneVideAvant: true);
            }
            finally
            {
                _cheminFichier = string.Empty;
                _numFICourant = string.Empty;
                _utilisateurCourant = string.Empty;
                _posteCourant = string.Empty;
                _debutSession = DateTime.MinValue;
                _derniereConfig = string.Empty;
                _dernierType = string.Empty;
                _libelleMesure = "Mesure";
                _mesureEnAttente = false;
                _arretDemande = false;
                _nbMesuresEffectuees = 0;
                _nbMesuresEchouees = 0;
            }
        }
    }
}
