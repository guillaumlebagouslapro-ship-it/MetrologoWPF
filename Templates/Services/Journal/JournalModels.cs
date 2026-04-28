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

        // Helpers de présentation
        public string TimestampAffiche => Timestamp.ToString("HH:mm:ss");

        /// <summary>
        /// Version lisible des Details JSON : décode les échappements Unicode courants
        /// (<c>"</c> → <c>"</c>, <c>+</c> → <c>+</c>, <c>></c> → <c>&gt;</c>…)
        /// pour que les valeurs SCPI soient lisibles dans la UI au lieu d'être noyées
        /// d'échappements JsonSerializer.
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
        /// <summary>Mode de mesure choisi en début de session : « Baie » ou « Paillasse ». Null avant choix.</summary>
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

        /// <summary>Bandeau d'en-tête montré dans la liste des sessions. Inclut le poste
        /// (Baie/Paillasse) s'il a été choisi par l'utilisateur ; la machine est volontairement
        /// masquée — info technique non pertinente pour la vue métier.</summary>
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
        /// Si non null, ne charger que les entrées dont <see cref="LogEntry.Action"/> est
        /// dans cette liste OU dont la sévérité est Avertissement/Erreur (toujours visible).
        /// Permet de pré-filtrer côté SQL et éviter de charger des milliers d'entrées techniques
        /// pour les afficher en mode normal. Null = aucune restriction (mode debug).
        /// </summary>
        public IReadOnlyCollection<string>? ActionsMetier { get; set; }
    }
}
