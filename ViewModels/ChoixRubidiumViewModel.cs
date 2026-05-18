using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using System;
using System.Collections.ObjectModel;
using System.Globalization;

namespace Metrologo.ViewModels
{
    public partial class ChoixRubidiumViewModel : ObservableObject
    {
        [ObservableProperty] private ObservableCollection<Rubidium> _rubidiums = new();
        [ObservableProperty] private Rubidium? _rubidiumSelectionne;
        [ObservableProperty] private bool _avecGPS;
        [ObservableProperty] private bool _sansGPS = true;

        // ---------- Mode Réglage manuel ----------

        /// <summary>
        /// Si vrai, l'admin saisit directement une fréquence moyenne plutôt que de
        /// choisir un rubidium du catalogue. Pour les cas où aucun rubidium n'est
        /// raccordé / pour des tests / pour des fréquences de référence custom.
        /// </summary>
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(EstModeListe))]
        private bool _modeManuel;

        /// <summary>Inverse de <see cref="ModeManuel"/> — utile pour bindings XAML.</summary>
        public bool EstModeListe => !ModeManuel;

        /// <summary>Fréquence saisie en mode manuel (texte pour parsing tolérant).</summary>
        [ObservableProperty] private string _frequenceManuelleTexte = "10000000";

        // ---------- Erreurs / résultat ----------

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _messageErreur = string.Empty;

        public bool HasError => !string.IsNullOrEmpty(MessageErreur);

        public Rubidium? Resultat { get; private set; }
        public bool AvecGpsResultat { get; private set; }
        public Action<bool>? CloseAction { get; set; }

        public bool ListeVide => Rubidiums.Count == 0;

        public ChoixRubidiumViewModel()
        {
            ChargerRubidiums();
        }

        private void ChargerRubidiums()
        {
            // TODO : remplacer par une requête SQL : Select RUB_ID, RUB_ACTIF, RUB_DESIGNATION from TR_METROLOGO_RUBIDIUMS
            Rubidiums.Add(new Rubidium { Id = 1, Designation = "Syref",       FrequenceMoyenne = 10000000.0 });
            Rubidiums.Add(new Rubidium { Id = 2, Designation = "Redondances", FrequenceMoyenne = 10000000.0 });
            RubidiumSelectionne = Rubidiums[0];
            OnPropertyChanged(nameof(ListeVide));
        }

        partial void OnAvecGPSChanged(bool value)
        {
            if (value) SansGPS = false;
        }

        partial void OnSansGPSChanged(bool value)
        {
            if (value) AvecGPS = false;
        }

        [RelayCommand]
        private void Valider()
        {
            MessageErreur = string.Empty;

            if (ModeManuel)
            {
                // Saisie manuelle : on parse la fréquence et on crée un Rubidium "factice"
                // avec Designation="Réglage manuel" et EstReglageManuel=true (utilisé par
                // NomAffichage et le badge en bas de l'écran d'accueil).
                string txt = (FrequenceManuelleTexte ?? "").Trim().Replace(',', '.');
                if (!double.TryParse(txt, NumberStyles.Float, CultureInfo.InvariantCulture, out double freq) || freq <= 0)
                {
                    MessageErreur = $"Fréquence invalide : « {FrequenceManuelleTexte} ». Saisis une valeur positive en Hz (ex. 10000000).";
                    return;
                }

                // Le mode de raccordement reste optionnel — l'admin peut rester sur Allouis
                // (par défaut) sans incidence si la fréquence est saisie à la main.
                Resultat = new Rubidium
                {
                    Id = 0,
                    Designation = "Réglage manuel",
                    FrequenceMoyenne = freq,
                    EstReglageManuel = true,
                    AvecGPS = AvecGPS
                };
                AvecGpsResultat = AvecGPS;
                CloseAction?.Invoke(true);
                return;
            }

            if (RubidiumSelectionne == null)
            {
                MessageErreur = "Veuillez sélectionner un rubidium.";
                return;
            }

            if (!AvecGPS && !SansGPS)
            {
                MessageErreur = "Veuillez spécifier le mode de raccordement.";
                return;
            }

            Resultat = RubidiumSelectionne;
            AvecGpsResultat = AvecGPS;
            Resultat.AvecGPS = AvecGPS;
            Resultat.EstReglageManuel = false;
            CloseAction?.Invoke(true);
        }

        [RelayCommand]
        private void Annuler() => CloseAction?.Invoke(false);
    }
}
