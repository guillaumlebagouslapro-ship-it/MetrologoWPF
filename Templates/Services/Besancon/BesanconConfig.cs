using System;
using System.IO;
using System.Text.Json;

namespace Metrologo.Services.Besancon
{
    /// <summary>
    /// Configuration de la tâche quotidienne Besançon (récupération FTP + moyenne hebdo).
    /// Stockée LOCALEMENT par poste (contient des identifiants FTP) dans
    /// <c>%LocalAppData%\Metrologo\Configuration\besancon.ftp.json</c> — donc PAS sur le partage
    /// réseau et PAS dans le dépôt git.
    ///
    /// <para><see cref="Active"/> est <c>false</c> par défaut : la tâche ne doit tourner que sur
    /// UN SEUL poste (sinon plusieurs postes iraient chercher le même fichier en même temps).
    /// L'admin l'active sur le poste « maître ».</para>
    /// </summary>
    public sealed class BesanconConfig
    {
        public bool Active { get; set; } = false;

        public string FtpHote { get; set; } = "ftp.aserti-group.com";
        public int FtpPort { get; set; } = 21;
        public string FtpUtilisateur { get; set; } = string.Empty;
        public string FtpMotDePasse { get; set; } = string.Empty;
        public bool FtpSsl { get; set; } = false;

        /// <summary>Nom du fichier à récupérer sur le FTP (legacy : <c>ef_utcop</c>).</summary>
        public string FichierDistant { get; set; } = "ef_utcop";

        /// <summary>Heure de déclenchement quotidien, format <c>HH:mm</c> (legacy : 09h50 / 10h00).</summary>
        public string HeureDeclenchement { get; set; } = "09:50";

        /// <summary>Si vrai, supprime le fichier sur le FTP après téléchargement (comportement
        /// legacy). DÉCONSEILLÉ en multi-poste — laissé à false par défaut.</summary>
        public bool SupprimerApresTelechargement { get; set; } = false;

        public static string Chemin =>
            Path.Combine(CheminsMetrologo.Configuration, "besancon.ftp.json");

        public static BesanconConfig Charger()
        {
            try
            {
                if (File.Exists(Chemin))
                {
                    var cfg = JsonSerializer.Deserialize<BesanconConfig>(File.ReadAllText(Chemin));
                    if (cfg != null) return cfg;
                }
            }
            catch { /* fichier corrompu → on régénère un défaut */ }

            var defaut = new BesanconConfig();
            defaut.Sauvegarder();   // crée le gabarit à remplir (hôte pré-rempli, identifiants vides)
            return defaut;
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

        /// <summary>Parse <see cref="HeureDeclenchement"/> ; retourne 09:50 par défaut si invalide.</summary>
        public TimeSpan HeureParsee()
        {
            if (TimeSpan.TryParseExact(HeureDeclenchement, "hh\\:mm",
                    System.Globalization.CultureInfo.InvariantCulture, out var ts))
                return ts;
            return new TimeSpan(9, 50, 0);
        }
    }
}
