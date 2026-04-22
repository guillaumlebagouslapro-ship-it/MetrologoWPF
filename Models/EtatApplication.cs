using System;
using System.Collections.ObjectModel;
using Metrologo.Services.Config;

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

        /// <summary>Configuration des appareils IEEE chargée depuis Metrologo.ini au démarrage.</summary>
        public static ConfigAppareils? ConfigAppareils { get; set; }

        /// <summary>
        /// Appareils actuellement détectés sur le bus GPIB (rafraîchi après chaque scan).
        /// Utilisé par la fenêtre de Configuration pour peupler la liste des fréquencemètres.
        /// </summary>
        public static ObservableCollection<AppareilDetecte> AppareilsDetectes { get; } = new();

        /// <summary>Levé après chaque mise à jour de <see cref="AppareilsDetectes"/>.</summary>
        public static event EventHandler? AppareilsDetectesChange;

        public static void NotifierAppareilsDetectesChange()
            => AppareilsDetectesChange?.Invoke(null, EventArgs.Empty);

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
