using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Globalization;

namespace Metrologo.ViewModels
{
    public partial class SaisieValFreqViewModel : ObservableObject
    {
        [ObservableProperty] private string _titre = "Saisie de la fréquence";
        [ObservableProperty] private string _sousTitre = "Entrez la valeur lue sur l'appareil";
        [ObservableProperty] private string _libelleChamp = "Fréquence lue (Hz)";

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _valeurTexte = "10000000";

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _messageErreur = string.Empty;

        public bool HasError => !string.IsNullOrEmpty(MessageErreur);
        public double ValeurLue { get; private set; }
        public Action<bool>? CloseAction { get; set; }

        public SaisieValFreqViewModel() { }

        public SaisieValFreqViewModel(double valeurInitiale, string? titre = null, string? sousTitre = null, string? libelle = null)
        {
            ValeurTexte = valeurInitiale.ToString("R", CultureInfo.InvariantCulture);
            if (titre != null) Titre = titre;
            if (sousTitre != null) SousTitre = sousTitre;
            if (libelle != null) LibelleChamp = libelle;
        }

        [RelayCommand]
        private void Valider()
        {
            var txt = (ValeurTexte ?? string.Empty).Trim().Replace(',', '.');

            if (!double.TryParse(txt, NumberStyles.Float, CultureInfo.InvariantCulture, out var val))
            {
                MessageErreur = $"Valeur incorrecte : '{ValeurTexte}'.";
                return;
            }

            if (val <= 0)
            {
                MessageErreur = "La fréquence doit être strictement positive.";
                return;
            }

            ValeurLue = val;
            CloseAction?.Invoke(true);
        }

        [RelayCommand]
        private void Annuler() => CloseAction?.Invoke(false);
    }
}
