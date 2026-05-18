using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Globalization;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// Saisie post-mesure consolidée pour Fréquence + source Fréquencemètre :
    /// fréquence lue sur l'afficheur, résolution du compteur et incertitude
    /// supplémentaire relative (dégradation). Page unique injectée dans Excel
    /// via les zones nommées ZNFreqRef / ZNIncertResol / ZNIncertSup.
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
            if (!TryParse(FrequenceLueTexte, out var f) || f <= 0)
            {
                MessageErreur = $"Fréquence lue invalide : '{FrequenceLueTexte}' (doit être > 0).";
                return;
            }
            if (!TryParse(ResolutionTexte, out var res) || res < 0)
            {
                MessageErreur = $"Résolution invalide : '{ResolutionTexte}'.";
                return;
            }
            if (!TryParse(IncertSuppTexte, out var supp) || supp < 0)
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

        private static bool TryParse(string s, out double v)
        {
            return double.TryParse((s ?? "").Trim().Replace(',', '.'),
                NumberStyles.Float, CultureInfo.InvariantCulture, out v);
        }
    }
}
