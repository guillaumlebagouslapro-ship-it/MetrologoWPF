using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Globalization;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// Saisie post-mesure pour TachyContact : un unique champ « Résolution (tr/min) ».
    /// Pas d'incertitude supplémentaire (la dégradation relative est portée par les coeffs
    /// C/D du module RPM, pas par une saisie manuelle). Valeur injectée dans la zone
    /// nommée ZNIncertResol du classeur Excel (convertie en Hz = saisie/60 pour rester
    /// cohérente avec les formules Hz existantes ; la cellule I25 du template tachy
    /// affiche le retour en tr/min via la formule =ZNIncertResol*60).
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

        /// <summary>Résolution saisie par l'utilisateur, en tr/min.</summary>
        public double ResolutionRpm { get; private set; }

        /// <summary>
        /// Résolution convertie en Hz, prête à être écrite dans ZNIncertResol
        /// (équivaut à ResolutionRpm / 60).
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
