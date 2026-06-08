using System;
using System.IO;
using System.Text.Json;

namespace Metrologo.Services.Besancon
{
    /// <summary>
    /// Configuration de la tâche quotidienne Besançon (récupération FTP + moyenne hebdo).
    /// Stockée sur le PARTAGE réseau (<c>M:\…\Besancon\besancon.ftp.json</c>) pour que tous les
    /// postes partagent les mêmes identifiants FTP — plus besoin de reconfigurer chaque poste.
    /// ⚠ Le mot de passe FTP est donc en clair sur le partage : toute personne ayant accès au
    /// dossier Besançon peut le lire (compromis assumé pour la mutualisation).
    ///
    /// <para><see cref="Active"/> est <c>true</c> par défaut : la tâche tourne sur TOUS les postes,
    /// mais un marqueur partagé (<c>derniere_recuperation.json</c>) fait que seul le premier poste
    /// télécharge réellement le fichier chaque jour — les autres voient « déjà fait » et passent.
    /// L'admin peut désactiver la tâche sur un poste en mettant <c>Active=false</c> (ce choix tient,
    /// la bascule one-shot ne se réapplique pas).</para>
    /// </summary>
    public sealed class BesanconConfig
    {
        public bool Active { get; set; } = true;

        /// <summary>Marque interne : la bascule one-shot « Active sur tous les postes » a déjà été
        /// appliquée à ce fichier de config. Empêche de ré-activer une tâche que l'admin a
        /// volontairement remise à <c>Active=false</c> par la suite.</summary>
        public bool MigrationMultiPosteAppliquee { get; set; } = false;

        public string FtpHote { get; set; } = "ftp.aserti-group.com";
        public int FtpPort { get; set; } = 21;
        public string FtpUtilisateur { get; set; } = string.Empty;
        public string FtpMotDePasse { get; set; } = string.Empty;
        public bool FtpSsl { get; set; } = false;

        /// <summary>Nom du fichier à récupérer sur le FTP (legacy : <c>ef_utcop</c>).</summary>
        public string FichierDistant { get; set; } = "ef_utcop";

        /// <summary>Heure de déclenchement quotidien, format <c>HH:mm</c>. Défaut 09h38 : décalé
        /// après 09h30 pour éviter le bouchon réseau de cet horaire.</summary>
        public string HeureDeclenchement { get; set; } = "09:38";

        /// <summary>Si vrai, supprime le fichier sur le FTP après téléchargement (comportement
        /// legacy). DÉCONSEILLÉ en multi-poste — laissé à false par défaut.</summary>
        public bool SupprimerApresTelechargement { get; set; } = false;

        /// <summary>Emplacement PARTAGÉ (tous les postes lisent/écrivent la même config).</summary>
        public static string Chemin =>
            Path.Combine(CheminsMetrologo.Besancon, "besancon.ftp.json");

        /// <summary>Ancien emplacement LOCAL (avant mutualisation) — sert à la migration one-shot.</summary>
        private static string CheminLocalLegacy =>
            Path.Combine(CheminsMetrologo.Configuration, "besancon.ftp.json");

        public static BesanconConfig Charger()
        {
            var cfg = ChargerBrut();
            bool modifie = false;

            // Migration heure : l'ancien défaut « 09:50 » (jamais réglable via une UI) est remonté
            // à « 09:38 » pour éviter le bouchon de 09h30. Idempotent (ne s'applique qu'une fois).
            if (cfg.HeureDeclenchement == "09:50")
            {
                cfg.HeureDeclenchement = "09:38";
                modifie = true;
            }

            // Migration multi-poste : bascule ONE-SHOT vers Active=true sur tous les postes (le
            // marqueur partagé empêche les doublons). Tracée par un flag → l'admin peut ensuite
            // remettre Active=false sans que ça se réactive tout seul.
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
            // 1. Config partagée déjà valide (identifiants renseignés) → on l'utilise.
            var partagee = LireFichier(Chemin);
            if (partagee != null && !string.IsNullOrWhiteSpace(partagee.FtpUtilisateur))
                return partagee;

            // 2. Sinon (partagée absente OU vide), on promeut une config LOCALE déjà remplie
            //    (ancien emplacement, identifiants saisis avant la mutualisation) vers le partage.
            //    Robuste même si un autre poste a déjà créé une config partagée vide.
            var locale = LireFichier(CheminLocalLegacy);
            if (locale != null && !string.IsNullOrWhiteSpace(locale.FtpUtilisateur))
            {
                locale.Sauvegarder();   // écrit sur le partage (Chemin)
                return locale;
            }

            // 3. Partagée existante mais vide, et rien à migrer → on garde la partagée (gabarit).
            if (partagee != null)
                return partagee;

            // 4. Rien nulle part → crée un gabarit par défaut sur le partage (à remplir une fois).
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

        /// <summary>Parse <see cref="HeureDeclenchement"/> ; retourne 09:38 par défaut si invalide.</summary>
        public TimeSpan HeureParsee()
        {
            if (TimeSpan.TryParseExact(HeureDeclenchement, "hh\\:mm",
                    System.Globalization.CultureInfo.InvariantCulture, out var ts))
                return ts;
            return new TimeSpan(9, 38, 0);
        }
    }
}
