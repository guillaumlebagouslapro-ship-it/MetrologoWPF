using System;
using System.IO;
using System.Text.Json;
using Metrologo.Services;

namespace Metrologo.Models
{
    /// <summary>
    /// Préférences persistées entre sessions.
    /// Fichier JSON : <c>%LocalAppData%\Metrologo\Configuration\settings.json</c>
    /// (chemin centralisé dans <see cref="CheminsMetrologo"/>).
    /// </summary>
    public static class Preferences
    {
        private static readonly string _chemin = CheminsMetrologo.FichierSettings;

        private static Settings _settings = new();

        public static Rubidium? RubidiumActif => _settings.RubidiumActif;

        public static string CheminMacroXLA
        {
            get => _settings.CheminMacroXLA ?? @"C:\Exe_Spe\Fct_VBA\Metrologo.xla";
            set { _settings.CheminMacroXLA = value; Sauvegarder(); }
        }

        public static void Charger()
        {
            try
            {
                if (!File.Exists(_chemin)) return;
                var json = File.ReadAllText(_chemin);
                var parsed = JsonSerializer.Deserialize<Settings>(json);
                if (parsed != null) _settings = parsed;
            }
            catch
            {
                // Fichier corrompu / inaccessible : on repart sur des préférences vides.
                _settings = new Settings();
            }
        }

        public static void SauvegarderRubidium(Rubidium? rubi)
        {
            _settings.RubidiumActif = rubi;
            Sauvegarder();
        }

        private static void Sauvegarder()
        {
            try
            {
                var dir = Path.GetDirectoryName(_chemin);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                var json = JsonSerializer.Serialize(_settings,
                    new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_chemin, json);
            }
            catch
            {
                // Silencieux — la perte de préférence n'est pas bloquante.
            }
        }

        private class Settings
        {
            public Rubidium? RubidiumActif { get; set; }
            public string? CheminMacroXLA { get; set; }
        }
    }
}
