using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Globalization;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// Saisie post-mesure pour TachyContact : un seul champ, « Résolution (tr/min) ».
    /// Ici pas d'incertitude supplémentaire à saisir — la dégradation relative est déjà
    /// portée par les coefficients C/D du module RPM. La valeur part dans la zone nommée
    /// ZNIncertResol du classeur Excel, convertie en Hz (saisie / 60) pour rester en phase
    /// avec les formules Hz existantes ; côté template tachy, la cellule I25 réaffiche le
    /// résultat en tr/min grâce à la formule =ZNIncertResol*60.
    /// </summary>
    public partial class SaisiePostMesureTachyViewModel : ObservableObject
    {
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _resolutionTexte = "0.1";

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _messageErreur = string.Empty;

        public bool HasError => !string.IsNullOrEmpty(MessageErreur);

        /// <summary>La résolution telle que saisie par l'utilisateur, en tr/min.</summary>
        public double ResolutionRpm { get; private set; }

        /// <summary>
        /// La même résolution, mais convertie en Hz et prête à écrire dans ZNIncertResol
        /// (c'est simplement ResolutionRpm / 60).
        /// </summary>
        public double ResolutionHz => ResolutionRpm / 60.0;

        public Action<bool>? CloseAction { get; set; }

        public SaisiePostMesureTachyViewModel() { }

        public SaisiePostMesureTachyViewModel(double resolutionRpmInitiale)
        {
            ResolutionTexte = resolutionRpmInitiale.ToString("R", CultureInfo.InvariantCulture);
        }

        [RelayCommand]
        private void Valider()
        {
            string txt = (ResolutionTexte ?? "").Trim().Replace(',', '.');
            if (!double.TryParse(txt, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) || v < 0)
            {
                MessageErreur = $"Résolution invalide : '{ResolutionTexte}' (doit être ≥ 0).";
                return;
            }

            ResolutionRpm = v;
            CloseAction?.Invoke(true);
        }

        [RelayCommand]
        private void Annuler() => CloseAction?.Invoke(false);
    }
}
