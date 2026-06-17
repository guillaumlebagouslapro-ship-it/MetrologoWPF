using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Services;
using System;
using System.Globalization;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// Saisie post-mesure regroupée pour le cas Fréquence avec source Fréquencemètre :
    /// la fréquence relevée sur l'afficheur, la résolution du compteur et l'incertitude
    /// supplémentaire relative (la dégradation). Tout tient sur une seule page, injectée
    /// dans Excel via les zones nommées ZNFreqRef / ZNIncertResol / ZNIncertSup.
    /// </summary>
    public partial class SaisiePostMesureFreqViewModel : ObservableObject
    {
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _frequenceLueTexte = "10000000";

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _resolutionTexte = "0.01";

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _incertSuppTexte = "0.0";

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _messageErreur = string.Empty;

        public bool HasError => !string.IsNullOrEmpty(MessageErreur);

        public double FrequenceLue { get; private set; }
        public double Resolution { get; private set; }
        public double IncertSupp { get; private set; }

        public Action<bool>? CloseAction { get; set; }

        public SaisiePostMesureFreqViewModel() { }

        public SaisiePostMesureFreqViewModel(double frequenceInitiale, double resolution, double incertSupp)
        {
            FrequenceLueTexte = frequenceInitiale.ToString("R", CultureInfo.InvariantCulture);
            ResolutionTexte = resolution.ToString("R", CultureInfo.InvariantCulture);
            IncertSuppTexte = incertSupp.ToString("R", CultureInfo.InvariantCulture);
        }

        [RelayCommand]
        private void Valider()
        {
            if (!SaisieHelper.TryParsePositiveDouble(FrequenceLueTexte, out var f))
            {
                MessageErreur = $"Fréquence lue invalide : '{FrequenceLueTexte}' (doit être > 0).";
                return;
            }
            if (!SaisieHelper.TryParseNonNegativeDouble(ResolutionTexte, out var res))
            {
                MessageErreur = $"Résolution invalide : '{ResolutionTexte}'.";
                return;
            }
            if (!SaisieHelper.TryParseNonNegativeDouble(IncertSuppTexte, out var supp))
            {
                MessageErreur = $"Incertitude sup. relative invalide : '{IncertSuppTexte}'.";
                return;
            }

            FrequenceLue = f;
            Resolution = res;
            IncertSupp = supp;
            CloseAction?.Invoke(true);
        }

        [RelayCommand]
        private void Annuler() => CloseAction?.Invoke(false);
    }
}
