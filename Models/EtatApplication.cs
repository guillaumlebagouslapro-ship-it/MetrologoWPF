using System;
using System.Collections.ObjectModel;
using System.Globalization;

namespace Metrologo.Models
{
    /// <summary>
    /// L'état global de l'application, partagé par tous les ViewModels.
    /// Au passage, il sauvegarde automatiquement le rubidium actif dans les préférences locales.
    /// </summary>
    public static class EtatApplication
    {
        private static Rubidium? _rubidiumActif;
        private static bool _chargeDepuisPreferences;

        /// <summary>L'utilisateur déclaré au démarrage via le menu déroulant (pour l'affichage et le journal).
        /// Aucune authentification ici : ça sert juste à savoir qui est devant l'app.</summary>
        public static Utilisateur? UtilisateurConnecte { get; set; }

        /// <summary>
        /// L'admin vraiment authentifié via la modale login + mot de passe. Vaut null en dehors de la
        /// zone admin (rempli au déverrouillage, remis à null au retour à l'Accueil). On s'en sert pour
        /// les actions sensibles, par ex. seul un SuperAdmin a le droit de gérer les rôles.
        /// </summary>
        public static Utilisateur? AdminConnecte { get; set; }

        /// <summary>Vrai quand l'admin authentifié est un SuperAdministrateur. Attention : on parle ici du
        /// compte saisi dans la modale Administration, pas du menu déroulant d'identité du démarrage.</summary>
        public static bool EstSuperAdmin =>
            AdminConnecte?.Role == RoleUtilisateur.SuperAdministrateur;

        /// <summary>Les appareils détectés sur le bus GPIB (remis à jour à chaque scan).
        /// C'est ce qui alimente la liste des fréquencemètres dans la fenêtre Configuration.</summary>
        public static ObservableCollection<AppareilDetecte> AppareilsDetectes { get; } = new();

        /// <summary>Déclenché à chaque fois que AppareilsDetectes change.</summary>
        public static event EventHandler? AppareilsDetectesChange;

        /// <summary>
        /// Mode adresses fixes (poste Baie). Quand c'est true, la fenêtre Configuration propose les
        /// appareils legacy du catalogue à une adresse GPIB qu'on peut éditer, plutôt que les appareils
        /// trouvés par scan. Toujours false en Paillasse.
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

                // la fréquence de référence exacte (toute la précision du double), groupée par
                // milliers pour qu'elle reste lisible ("10 000 000"). C'est cette valeur que ZNFreqRef reprend dans Excel.
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
        /// Force la notification et la sauvegarde du rubidium courant. À appeler quand on a modifié ses
        /// champs en place depuis la gestion du catalogue : comme le setter court-circuite sur
        /// ReferenceEquals, il ne se déclencherait pas tout seul dans ce cas.
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
        /// Relit le rubidium actif depuis le fichier réseau et prévient l'UI. Sans ça, le chargement
        /// n'aurait lieu qu'au démarrage et le rubidium resterait figé après un changement fait par un
        /// admin depuis un autre poste. C'est l'actualisation à chaud de la configuration qui l'appelle.
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
