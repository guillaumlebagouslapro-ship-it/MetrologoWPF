using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Metrologo.Services.Config
{
    /// <summary>
    /// Lecteur INI minimaliste sans dépendance. Gère les sections [xxx], les paires clé=valeur
    /// et les commentaires en début de ligne (;). Les clés sont insensibles à la casse.
    /// </summary>
    public class IniFile
    {
        private readonly Dictionary<string, Dictionary<string, string>> _sections
            = new(StringComparer.OrdinalIgnoreCase);

        public IReadOnlyCollection<string> Sections => _sections.Keys;

        public static IniFile Charger(string cheminFichier)
        {
            if (!File.Exists(cheminFichier))
                throw new FileNotFoundException(
                    $"Fichier de configuration introuvable : {cheminFichier}", cheminFichier);

            var ini = new IniFile();
            string? sectionCourante = null;

            foreach (var brute in File.ReadAllLines(cheminFichier, Encoding.UTF8))
            {
                var ligne = brute.Trim();
                if (ligne.Length == 0 || ligne.StartsWith(";") || ligne.StartsWith("#"))
                    continue;

                if (ligne.StartsWith("[") && ligne.EndsWith("]"))
                {
                    sectionCourante = ligne[1..^1].Trim();
                    if (!ini._sections.ContainsKey(sectionCourante))
                        ini._sections[sectionCourante] = new(StringComparer.OrdinalIgnoreCase);
                    continue;
                }

                if (sectionCourante == null) continue;

                int egal = ligne.IndexOf('=');
                if (egal < 0) continue;

                string cle = ligne[..egal].Trim();
                string val = ligne[(egal + 1)..].Trim();
                if (cle.Length == 0) continue;

                ini._sections[sectionCourante][cle] = val;
            }

            return ini;
        }

        public bool ContientSection(string section) => _sections.ContainsKey(section);

        public string? Valeur(string section, string cle)
            => _sections.TryGetValue(section, out var map) && map.TryGetValue(cle, out var v) ? v : null;

        public string ValeurObligatoire(string section, string cle)
            => Valeur(section, cle)
               ?? throw new InvalidDataException(
                   $"Clé « {cle} » manquante dans la section [{section}].");

        public int EntierObligatoire(string section, string cle)
        {
            var brut = ValeurObligatoire(section, cle);
            if (int.TryParse(brut, out var n)) return n;
            throw new InvalidDataException(
                $"Clé « {cle} » dans [{section}] n'est pas un entier valide : « {brut} ».");
        }

        public bool BooleenObligatoire(string section, string cle)
            => EntierObligatoire(section, cle) != 0;
    }
}
