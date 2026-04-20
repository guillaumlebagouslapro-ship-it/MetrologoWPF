using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;

namespace Metrologo.ViewModels
{
    public partial class MdpValidationViewModel : ObservableObject
    {
        // Mots de passe hérités de l'ancien code
        public const string MdpValidation = "METROL";
        public const string MdpIncertitudes = "1135";

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _messageErreur = string.Empty;

        [ObservableProperty] private string _titre = "Authentification requise";
        [ObservableProperty] private string _sousTitre = "Entrez le mot de passe pour continuer";
        [ObservableProperty] private string _mdpAttendu = MdpValidation;

        public bool HasError => !string.IsNullOrEmpty(MessageErreur);
        public bool PasswordOk { get; private set; }
        public Action<bool>? CloseAction { get; set; }

        public MdpValidationViewModel() { }

        public MdpValidationViewModel(string mdpAttendu, string? titre = null, string? sousTitre = null)
        {
            MdpAttendu = mdpAttendu;
            if (titre != null) Titre = titre;
            if (sousTitre != null) SousTitre = sousTitre;
        }

        [RelayCommand]
        public void Verifier(string mdpSaisi)
        {
            if (string.IsNullOrEmpty(mdpSaisi))
            {
                MessageErreur = "Saisissez le mot de passe.";
                return;
            }

            if (string.Equals(mdpSaisi, MdpAttendu, StringComparison.Ordinal)
                || string.Equals(mdpSaisi.ToUpperInvariant(), MdpAttendu, StringComparison.Ordinal))
            {
                PasswordOk = true;
                CloseAction?.Invoke(true);
                return;
            }

            MessageErreur = "Mot de passe incorrect.";
        }

        [RelayCommand]
        private void Annuler() => CloseAction?.Invoke(false);
    }
}
