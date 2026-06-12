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

        /// <summary>Utilisateur déclaré au démarrage via le dropdown (affichage + journal).
        /// Pas d'authentification, sert juste à savoir qui est en face de l'app.</summary>
        public static Utilisateur? UtilisateurConnecte { get; set; }

        /// <summary>
        /// Admin réellement authentifié via la modale login + mot de passe. null hors zone admin
        /// (set au déverrouillage, clear au retour Accueil). Sert aux actions sensibles,
        /// ex. seul un SuperAdmin gère les rôles.
        /// </summary>
        public static Utilisateur? AdminConnecte { get; set; }

        /// <summary>Vrai si l'admin authentifié est SuperAdministrateur. Reflète le compte de la
        /// modale Administration, pas le dropdown d'identité du démarrage.</summary>
        public static bool EstSuperAdmin =>
            AdminConnecte?.Role == RoleUtilisateur.SuperAdministrateur;

        /// <summary>Appareils détectés sur le bus GPIB (rafraîchi à chaque scan).
        /// Alimente la liste des fréquencemètres de la fenêtre Configuration.</summary>
        public static ObservableCollection<AppareilDetecte> AppareilsDetectes { get; } = new();

        /// <summary>Levé après chaque mise à jour de AppareilsDetectes.</summary>
        public static event EventHandler? AppareilsDetectesChange;

        /// <summary>
        /// Mode adresses fixes (poste Baie). Si true, la fenêtre Configuration propose les appareils
        /// legacy du catalogue à une adresse GPIB éditable au lieu des appareils détectés par scan.
        /// Toujours false en Paillasse.
        /// </summary>
        public static bool ModeAdressesFixes { get; set; }

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

                // fréquence de référence exacte (toute la précision du double), groupée par
                // milliers pour la lisibilité ("10 000 000"). C'est la valeur reprise par ZNFreqRef dans Excel.
                string freq = Metrologo.Services.SaisieHelper.FormaterFrequence(
                    _rubidiumActif.FrequenceMoyenne);
                string libelle = _rubidiumActif.EstReglageManuel
                    ? "Réglage manuel"
                    : _rubidiumActif.Designation;
                return $"Rubidium : {libelle} · {freq} Hz";
            }
        }

        public static event EventHandler? RubidiumActifChange;

        /// <summary>
        /// Force notification + sauvegarde du rubidium courant. À appeler quand on a modifié ses
        /// champs en place via la gestion du catalogue : le setter short-circuit sur ReferenceEquals
        /// et ne se déclencherait pas.
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

        /// <summary>
        /// Relit le rubidium actif depuis le fichier réseau et notifie l'UI. Sans ça, le chargement
        /// ne se fait qu'au démarrage et le rubidium resterait figé après un changement admin fait
        /// depuis un autre poste. Appelé par l'actualisation à chaud de la configuration.
        /// </summary>
        public static void RechargerRubidiumActif()
        {
            Preferences.Charger();
            _rubidiumActif = Preferences.RubidiumActif;
            _chargeDepuisPreferences = true;
            RubidiumActifChange?.Invoke(null, EventArgs.Empty);
        }
    }
}
