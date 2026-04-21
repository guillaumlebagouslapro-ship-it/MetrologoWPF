using System;

namespace Metrologo.Models
{
    /// <summary>
    /// État global de l'application partagé entre les ViewModels.
    /// Persiste automatiquement le rubidium actif dans les préférences locales.
    /// </summary>
    public static class EtatApplication
    {
        private static Rubidium? _rubidiumActif;
        private static bool _chargeDepuisPreferences;

        public static Rubidium? RubidiumActif
        {
            get
            {
                AssurerChargement();
                return _rubidiumActif;
            }
            set
            {
                AssurerChargement();
                if (ReferenceEquals(_rubidiumActif, value)) return;
                _rubidiumActif = value;
                Preferences.SauvegarderRubidium(value);
                RubidiumActifChange?.Invoke(null, EventArgs.Empty);
            }
        }

        public static string RubidiumActifTexte
        {
            get
            {
                AssurerChargement();
                return _rubidiumActif == null
                    ? "Rubidium : non défini"
                    : $"Rubidium : {_rubidiumActif.Designation} — "
                        + (_rubidiumActif.AvecGPS ? "raccord GPS" : "raccord Allouis");
            }
        }

        public static event EventHandler? RubidiumActifChange;

        private static void AssurerChargement()
        {
            if (_chargeDepuisPreferences) return;
            _chargeDepuisPreferences = true;
            Preferences.Charger();
            _rubidiumActif = Preferences.RubidiumActif;
        }
    }
}
