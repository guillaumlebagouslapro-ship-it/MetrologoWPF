using System;
using System.IO;
using System.Text.Json;

namespace Metrologo.Services.Besancon
{
    /// <summary>
    /// Config de la tâche quotidienne Besançon. Stockée sur le partage (besancon.ftp.json)
    /// pour que tous les postes partagent les mêmes identifiants FTP. Mot de passe en clair :
    /// compromis assumé. Le marqueur derniere_recuperation.json garantit qu'un seul poste
    /// télécharge par jour ; Active=false sur un poste est respecté définitivement.
    /// </summary>
    public sealed class BesanconConfig
    {
        public bool Active { get; set; } = true;

        /// <summary>Marque interne : la bascule one-shot vers Active sur tous les postes a déjà été
        /// appliquée. Évite de réactiver une tâche que l'admin a remise à Active=false depuis.</summary>
        public bool MigrationMultiPosteAppliquee { get; set; } = false;

        public string FtpHote { get; set; } = "ftp.aserti-group.com";
        public int FtpPort { get; set; } = 21;
        public string FtpUtilisateur { get; set; } = string.Empty;
        public string FtpMotDePasse { get; set; } = string.Empty;
        public bool FtpSsl { get; set; } = false;

        /// <summary>Nom du fichier à récupérer sur le FTP (legacy : ef_utcop).</summary>
        public string FichierDistant { get; set; } = "ef_utcop";

        /// <summary>Heure de déclenchement quotidien (HH:mm). Défaut 14:00 : valeur de la veille
        /// dispo vers 12h, transfert baie→FTP vers 13h30 ; 14h laisse une marge confortable.</summary>
        public string HeureDeclenchement { get; set; } = "14:00";

        /// <summary>Si vrai, supprime le fichier sur le FTP après téléchargement (comportement
        /// legacy). Déconseillé en multi-poste, false par défaut.</summary>
        public bool SupprimerApresTelechargement { get; set; } = false;

        /// <summary>Emplacement PARTAGÉ (tous les postes lisent/écrivent la même config).</summary>
        public static string Chemin =>
            Path.Combine(CheminsMetrologo.Besancon, "besancon.ftp.json");

        /// <summary>Ancien emplacement local (avant mutualisation), sert à la migration one-shot.</summary>
        private static string CheminLocalLegacy =>
            Path.Combine(CheminsMetrologo.Configuration, "besancon.ftp.json");

        public static BesanconConfig Charger()
        {
            var cfg = ChargerBrut();
            bool modifie = false;

            // Anciens défauts (09:50, 09:38, 13:38) remontés à 14:00. Idempotent ;
            // une heure réglée manuellement (autre valeur) n'est pas touchée.
            if (cfg.HeureDeclenchement == "09:50" || cfg.HeureDeclenchement == "09:38"
                || cfg.HeureDeclenchement == "13:38")
            {
                cfg.HeureDeclenchement = "14:00";
                modifie = true;
            }

            // Bascule one-shot Active=true (multi-poste). Flag MigrationMultiPosteAppliquee
            // empêche la réactivation si l'admin remet Active=false.
            if (!cfg.MigrationMultiPosteAppliquee)
            {
                cfg.Active = true;
                cfg.MigrationMultiPosteAppliquee = true;
                modifie = true;
            }

            if (modifie) cfg.Sauvegarder();
            return cfg;
        }

        private static BesanconConfig ChargerBrut()
        {
            // 1. Config partagée valide : on l'utilise.
            var partagee = LireFichier(Chemin);
            if (partagee != null && !string.IsNullOrWhiteSpace(partagee.FtpUtilisateur))
                return partagee;

            // 2. Config partagée absente/vide : promeut la config locale legacy vers le partage.
            var locale = LireFichier(CheminLocalLegacy);
            if (locale != null && !string.IsNullOrWhiteSpace(locale.FtpUtilisateur))
            {
                locale.Sauvegarder();   // écrit sur le partage (Chemin)
                return locale;
            }

            // 3. Partagée vide, rien à migrer : on garde le gabarit.
            if (partagee != null)
                return partagee;

            // 4. Rien nulle part : gabarit par défaut sur le partage (à remplir).
            var defaut = new BesanconConfig();
            defaut.Sauvegarder();
            return defaut;
        }

        private static BesanconConfig? LireFichier(string chemin)
        {
            try
            {
                if (File.Exists(chemin))
                    return JsonSerializer.Deserialize<BesanconConfig>(File.ReadAllText(chemin));
            }
            catch { /* fichier corrompu ou injoignable */ }
            return null;
        }

        public void Sauvegarder()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(Chemin)!);
                File.WriteAllText(Chemin,
                    JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }));
            }
            catch { /* best-effort */ }
        }

        /// <summary>Parse HeureDeclenchement ; retombe sur 14:00 si invalide.</summary>
        public TimeSpan HeureParsee()
        {
            if (TimeSpan.TryParseExact(HeureDeclenchement, "hh\\:mm",
                    System.Globalization.CultureInfo.InvariantCulture, out var ts))
                return ts;
            return new TimeSpan(14, 0, 0);
        }
    }
}
