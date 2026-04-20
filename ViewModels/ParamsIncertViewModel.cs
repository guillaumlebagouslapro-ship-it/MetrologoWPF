using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Globalization;

namespace Metrologo.ViewModels
{
    public partial class ParamsIncertViewModel : ObservableObject
    {
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

        public double Resolution { get; private set; }
        public double IncertSupp { get; private set; }

        public Action<bool>? CloseAction { get; set; }

        public ParamsIncertViewModel() { }

        public ParamsIncertViewModel(double resolution, double incertSupp)
        {
            ResolutionTexte = resolution.ToString("R", CultureInfo.InvariantCulture);
            IncertSuppTexte = incertSupp.ToString("R", CultureInfo.InvariantCulture);
        }

        [RelayCommand]
        private void Valider()
        {
            if (!TryParse(ResolutionTexte, out var res) || res < 0)
            {
                MessageErreur = $"Résolution invalide : '{ResolutionTexte}'";
                return;
            }
            if (!TryParse(IncertSuppTexte, out var supp) || supp < 0)
            {
                MessageErreur = $"Incertitude supplémentaire invalide : '{IncertSuppTexte}'";
                return;
            }

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
