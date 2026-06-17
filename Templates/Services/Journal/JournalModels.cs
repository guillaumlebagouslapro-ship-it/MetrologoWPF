using System;
using System.Collections.Generic;

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

        // Petits helpers d'affichage
        public string TimestampAffiche => Timestamp.ToString("HH:mm:ss");

        /// <summary>
        /// Version lisible des Details JSON : on décode les échappements Unicode les plus
        /// courants (<c>"</c> → <c>"</c>, <c>+</c> → <c>+</c>, <c>></c> → <c>&gt;</c>…)
        /// pour que les valeurs SCPI s'affichent proprement dans l'UI, au lieu d'être noyées
        /// sous les échappements du JsonSerializer.
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
        public string CategorieEmoji => Categorie switch
        {
            CategorieLog.Authentification => "🔐",
            CategorieLog.Session => "🚪",
            CategorieLog.Configuration => "⚙",
            CategorieLog.Mesure => "📊",
            CategorieLog.Rubidium => "🎯",
            CategorieLog.Administration => "🛡",
            CategorieLog.Excel => "📈",
            CategorieLog.Systeme => "💻",
            CategorieLog.Erreur => "⚠",
            _ => "•"
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
        public TimeSpan Duree => (Fin ?? DateTime.Now) - Debut;
        public int NbEntrees => Entrees.Count;

        public string DebutAffiche => Debut.ToString("dd/MM/yyyy HH:mm:ss");
        public string FinAffiche => Fin.HasValue ? Fin.Value.ToString("HH:mm:ss") : "— en cours —";
        public string DureeAffiche => Duree.TotalHours >= 1
            ? $"{(int)Duree.TotalHours} h {Duree.Minutes} min"
            : Duree.TotalMinutes >= 1
                ? $"{(int)Duree.TotalMinutes} min {Duree.Seconds} s"
                : $"{(int)Duree.TotalSeconds} s";

        /// <summary>Le bandeau d'en-tête affiché dans la liste des sessions. Il reprend le poste
        /// (Baie/Paillasse) s'il a été choisi ; la machine, elle, est masquée à dessein — c'est une
        /// info technique qui n'apporte rien à la vue métier.</summary>
        public string EnteteAffiche
        {
            get
            {
                string posteTag = string.IsNullOrEmpty(Poste) ? "" : $"{Poste} · ";
                return $"{posteTag}{DebutAffiche} → {FinAffiche} ({DureeAffiche})";
            }
        }
        public int NbErreurs
        {
            get
            {
                int n = 0;
                foreach (var e in Entrees) if (e.Severite == SeveriteLog.Erreur) n++;
                return n;
            }
        }
        public int NbAvertissements
        {
            get
            {
                int n = 0;
                foreach (var e in Entrees) if (e.Severite == SeveriteLog.Avertissement) n++;
                return n;
            }
        }

        public bool HasErreurs => NbErreurs > 0;
        public bool HasAvertissements => NbAvertissements > 0;

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

        /// <summary>Résumé court pour la carte compacte, du genre : 47 actions · 2 mesures · 0 erreur.</summary>
        public string ResumeAffiche
        {
            get
            {
                var parts = new List<string> { $"{NbEntrees} action{(NbEntrees > 1 ? "s" : "")}" };
                if (NbMesures > 0) parts.Add($"{NbMesures} mesure{(NbMesures > 1 ? "s" : "")}");
                if (NbErreurs > 0) parts.Add($"{NbErreurs} erreur{(NbErreurs > 1 ? "s" : "")}");
                else if (NbAvertissements > 0) parts.Add($"{NbAvertissements} avert.");
                return string.Join(" · ", parts);
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
        /// Si renseigné, on ne charge que les entrées dont l'<see cref="LogEntry.Action"/> figure
        /// dans cette liste, ou dont la sévérité est Avertissement/Erreur (qui restent toujours
        /// visibles). Ça permet de pré-filtrer en amont et d'éviter de remonter des milliers
        /// d'entrées techniques en mode normal. Null = aucune restriction (mode debug).
        /// </summary>
        public IReadOnlyCollection<string>? ActionsMetier { get; set; }
    }
}
