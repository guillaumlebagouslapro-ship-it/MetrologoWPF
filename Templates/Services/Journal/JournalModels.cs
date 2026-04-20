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

        public string EnteteAffiche =>
            $"{Utilisateur} · {Machine} · {DebutAffiche} → {FinAffiche} ({DureeAffiche})";
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
    }
}
