namespace Metrologo.Services.Journal
{
    public enum CategorieLog
    {
        Authentification,
        Session,
        Configuration,
        Mesure,
        Rubidium,
        Administration,
        Excel,
        Systeme,
        Erreur
    }

    public enum SeveriteLog
    {
        Info,
        Avertissement,
        Erreur
    }

    public class LogEntry
    {
        public long EntryId { get; set; }
        public string SessionId { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; }
        public CategorieLog Categorie { get; set; }
        public string Action { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
        public string? Details { get; set; }
        public SeveriteLog Severite { get; set; }

        /// <summary>Login de l'opérateur au moment de l'entrée. Absent des anciennes entrées : la
        /// lecture le complète alors depuis sessions.json (voir FichierJournalService.ChargerEntreesAsync).</summary>
        public string? Utilisateur { get; set; }

        /// <summary>Poste (nom de machine) qui a écrit l'entrée, même remarque que <see cref="Utilisateur"/>.</summary>
        public string? Machine { get; set; }

        // Petits helpers d'affichage
        public string HorodatageAffiche => Timestamp.ToString("dd/MM/yyyy HH:mm:ss");
        public string UtilisateurAffiche => string.IsNullOrWhiteSpace(Utilisateur) ? "—" : Utilisateur;
        public string SeveriteLibelle => Severite switch
        {
            SeveriteLog.Erreur        => "Erreur",
            SeveriteLog.Avertissement => "Avert.",
            _                         => "Info",
        };

        /// <summary>
        /// Version lisible des Details JSON : décode les échappements Unicode du JsonSerializer
        /// (guillemets, +, &gt;...) pour que les valeurs SCPI s'affichent proprement.
        /// </summary>
        public string? DetailsLisibles
        {
            get
            {
                if (string.IsNullOrEmpty(Details)) return Details;
                return System.Text.RegularExpressions.Regex.Replace(
                    Details,
                    @"\\u([0-9a-fA-F]{4})",
                    m => ((char)System.Convert.ToInt32(m.Groups[1].Value, 16)).ToString());
            }
        }
        public string CategorieLibelle => Categorie switch
        {
            CategorieLog.Authentification => "Authentification",
            CategorieLog.Session => "Session",
            CategorieLog.Configuration => "Configuration",
            CategorieLog.Mesure => "Mesure",
            CategorieLog.Rubidium => "Rubidium",
            CategorieLog.Administration => "Administration",
            CategorieLog.Excel => "Excel",
            CategorieLog.Systeme => "Système",
            CategorieLog.Erreur => "Erreur",
            _ => "Autre"
        };
    }

    public class SessionJournal
    {
        public string SessionId { get; set; } = string.Empty;
        public string Utilisateur { get; set; } = string.Empty;
        public string Machine { get; set; } = string.Empty;
        /// <summary>Mode de mesure retenu au début de la session : « Baie » ou « Paillasse ». Null tant qu'on n'a pas choisi.</summary>
        public string? Poste { get; set; }
        public DateTime Debut { get; set; }
        public DateTime? Fin { get; set; }
        public List<LogEntry> Entrees { get; set; } = new();

        public bool Active => Fin == null;

        public int NbErreurs
        {
            get
            {
                int n = 0;
                foreach (var e in Entrees) if (e.Severite == SeveriteLog.Erreur) n++;
                return n;
            }
        }

        /// <summary>Nombre de mesures faites pendant la session (on compte les MESURE_DEBUT).</summary>
        public int NbMesures
        {
            get
            {
                int n = 0;
                foreach (var e in Entrees)
                    if (string.Equals(e.Action, "MESURE_DEBUT", StringComparison.OrdinalIgnoreCase)) n++;
                return n;
            }
        }

    }

    public class FiltreJournal
    {
        public DateTime? Depuis { get; set; }
        public DateTime? Jusqu_a { get; set; }
        public string? Utilisateur { get; set; }
        public CategorieLog? Categorie { get; set; }
        public SeveriteLog? SeveriteMin { get; set; }
        public string? Recherche { get; set; }

        /// <summary>
        /// Si renseigné, ne charge que les entrées dont l'<see cref="LogEntry.Action"/> figure dans
        /// cette liste, ou de sévérité Avertissement/Erreur (toujours visibles) : évite de remonter
        /// des milliers d'entrées techniques en mode normal. Null = aucune restriction (mode debug).
        /// </summary>
        public IReadOnlyCollection<string>? ActionsMetier { get; set; }
    }
}
