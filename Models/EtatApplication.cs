using System;
using System.Collections.ObjectModel;
using System.Globalization;

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

        /// <summary>
        /// Utilisateur déclaré au démarrage via le dropdown (identité affichage + journal).
        /// Aucune authentification — sert juste à savoir qui est en face de l'app.
        /// </summary>
        public static Utilisateur? UtilisateurConnecte { get; set; }

        /// <summary>
        /// Admin réellement authentifié via la modale login + mot de passe. <c>null</c>
        /// tant qu'on n'est pas dans la zone admin. Set au déverrouillage admin, clear
        /// au retour Accueil. Sert à déterminer les actions sensibles autorisées
        /// (ex. seul un SuperAdmin peut gérer les rôles).
        /// </summary>
        public static Utilisateur? AdminConnecte { get; set; }

        /// <summary>
        /// Vrai si l'admin authentifié dans la session courante est SuperAdministrateur.
        /// Reflète le compte qui a validé la modale Administration, PAS le dropdown
        /// d'identité au démarrage.
        /// </summary>
        public static bool EstSuperAdmin =>
            AdminConnecte?.Role == RoleUtilisateur.SuperAdministrateur;

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
                if (_rubidiumActif == null) return "Rubidium : non défini";

                // Format français avec espace milliers : "10 000 000,00 Hz".
                string freq = _rubidiumActif.FrequenceMoyenne.ToString(
                    "N2", CultureInfo.GetCultureInfo("fr-FR"));
                string libelle = _rubidiumActif.EstReglageManuel
                    ? "Réglage manuel"
                    : _rubidiumActif.Designation;
                return $"Rubidium : {libelle} · {freq} Hz";
            }
        }

        public static event EventHandler? RubidiumActifChange;

        /// <summary>
        /// Force la notification + la sauvegarde du rubidium courant. À appeler
        /// quand on a modifié en place les champs (Designation / FrequenceMoyenne)
        /// du rubidium actif via la gestion du catalogue : le setter standard
        /// short-circuit sur <see cref="ReferenceEquals"/> et ne se déclencherait pas.
        /// </summary>
        public static void NotifierRubidiumActifChange()
        {
            Preferences.SauvegarderRubidium(_rubidiumActif);
            RubidiumActifChange?.Invoke(null, EventArgs.Empty);
        }

        private static void AssurerChargement()
        {
            if (_chargeDepuisPreferences) return;
            _chargeDepuisPreferences = true;
            Preferences.Charger();
            _rubidiumActif = Preferences.RubidiumActif;
        }
    }
}
