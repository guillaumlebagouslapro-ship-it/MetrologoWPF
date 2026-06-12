using System;
using System.Collections.Generic;
using System.IO;
using Metrologo.Models;

namespace Metrologo.Services
{
    /// <summary>
    /// Lit Mesures_Config.txt (LocalAppData\Metrologo) : n° de module et fonction par type
    /// de mesure, écrits en colonnes A/B de la feuille de mesure (la col C = temps de gate
    /// vient de la sélection utilisateur). Format INI simple éditable par l'admin,
    /// créé avec les valeurs livrées au premier démarrage s'il n'existe pas.
    /// </summary>
    public static class MesureConfigService
    {
        private const string NomFichier = "Mesures_Config.txt";
        private static Dictionary<string, (string Module, string Fonction)>? _cache;

        /// <summary>
        /// (Module, Fonction) pour un type de mesure ; les sections du fichier portent le nom
        /// du TypeMesure. Si la section manque on retourne ("", ""), cellules vides côté Excel.
        /// </summary>
        public static (string Module, string Fonction) ObtenirPourType(TypeMesure type)
        {
            EnsureCache();
            return _cache!.TryGetValue(type.ToString(), out var v) ? v : ("", "");
        }

        /// <summary>Force la relecture du fichier au prochain appel (si l'admin l'a modifié en cours de session).</summary>
        public static void Recharger() => _cache = null;

        public static string CheminFichier =>
            CheminsMetrologo.ResoudreCheminAvecFallback(
                CheminsMetrologo.FichierMesuresConfig, NomFichier);

        // -------------------------------------------------------------------------

        private static void EnsureCache()
        {
            if (_cache != null) return;
            _cache = new Dictionary<string, (string, string)>(StringComparer.OrdinalIgnoreCase);

            string chemin = CheminFichier;
            if (!File.Exists(chemin)) CreerFichierParDefaut(chemin);

            try
            {
                string? sectionCourante = null;
                foreach (var ligne in File.ReadAllLines(chemin))
                {
                    string l = ligne.Trim();
                    if (string.IsNullOrEmpty(l) || l.StartsWith("#") || l.StartsWith(";")) continue;

                    if (l.StartsWith("[") && l.EndsWith("]"))
                    {
                        sectionCourante = l[1..^1].Trim();
                        if (!_cache.ContainsKey(sectionCourante)) _cache[sectionCourante] = ("", "");
                        continue;
                    }

                    if (sectionCourante == null) continue;
                    int eq = l.IndexOf('=');
                    if (eq <= 0) continue;
                    string cle = l[..eq].Trim();
                    string val = l[(eq + 1)..].Trim();

                    var actuel = _cache[sectionCourante];
                    if (cle.Equals("Module", StringComparison.OrdinalIgnoreCase))
                        _cache[sectionCourante] = (val, actuel.Item2);
                    else if (cle.Equals("Fonction", StringComparison.OrdinalIgnoreCase))
                        _cache[sectionCourante] = (actuel.Item1, val);
                }
            }
            catch (Exception)
            {
                // fichier mal formé : on garde le cache vide, ça donnera des cellules vides côté Excel
            }
        }

        private static void CreerFichierParDefaut(string chemin)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(chemin)!);
            string contenu =
                "# Configuration des étiquettes Module + Fonction écrites dans les feuilles" + Environment.NewLine +
                "# de mesure Excel (cols A et B). Une section par TypeMesure de l'app." + Environment.NewLine +
                "# Modifiable à chaud — l'app relit le fichier à chaque démarrage." + Environment.NewLine +
                Environment.NewLine +
                "[Frequence]" + Environment.NewLine +
                "Module=78015" + Environment.NewLine +
                "Fonction=freq" + Environment.NewLine +
                Environment.NewLine +
                "[Stabilite]" + Environment.NewLine +
                "Module=78020" + Environment.NewLine +
                "Fonction=stab" + Environment.NewLine +
                Environment.NewLine +
                "[FreqAvantInterv]" + Environment.NewLine +
                "Module=78010" + Environment.NewLine +
                "Fonction=freqAv" + Environment.NewLine +
                Environment.NewLine +
                "[FreqFinale]" + Environment.NewLine +
                "Module=78011" + Environment.NewLine +
                "Fonction=freqFin" + Environment.NewLine +
                Environment.NewLine +
                "[Interval]" + Environment.NewLine +
                "Module=78030" + Environment.NewLine +
                "Fonction=interv" + Environment.NewLine +
                Environment.NewLine +
                "[TachyContact]" + Environment.NewLine +
                "Module=78050" + Environment.NewLine +
                "Fonction=tachyC" + Environment.NewLine +
                Environment.NewLine +
                "[Stroboscope]" + Environment.NewLine +
                "Module=78060" + Environment.NewLine +
                "Fonction=strobo" + Environment.NewLine;

            File.WriteAllText(chemin, contenu);
        }

        private static string DossierConfig() => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Metrologo");
    }
}
