using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Globalization;

namespace Metrologo.ViewModels
{
    public partial class ModifValBesanconViewModel : ObservableObject
    {
        private static readonly DateTime MjdEpoch = new(1858, 11, 17);

        [ObservableProperty] private DateTime _dateSelectionnee = DateTime.Today;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private int _dateJulienne;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _valeurJourTexte = "";

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _messageErreur = string.Empty;

        public bool HasError => !string.IsNullOrEmpty(MessageErreur);

        public double ValeurJourResultat { get; private set; }
        public int DateJulienneResultat { get; private set; }
        public Action<bool>? CloseAction { get; set; }

        public ModifValBesanconViewModel() { Recalculer(); }

        partial void OnDateSelectionneeChanged(DateTime value)
        {
            Recalculer();
            ChargerValeurExistante();
        }

        private void Recalculer()
        {
            DateJulienne = (int)(DateSelectionnee.Date - MjdEpoch).TotalDays;
        }

        private void ChargerValeurExistante()
        {
            // TODO : requête SQL : select DAT_VALEUR from T_METROLOGO_DATESRUBIS where DAT_ID=@jd and RUB_ACTIF=1
            // Pour l'instant, simulation : rien
            ValeurJourTexte = "";
        }

        [RelayCommand]
        private void Valider()
        {
            var txt = (ValeurJourTexte ?? "").Trim().Replace(',', '.');
            if (!double.TryParse(txt, NumberStyles.Float, CultureInfo.InvariantCulture, out var v))
            {
                MessageErreur = "Valeur incorrecte.";
                return;
            }
            if (Math.Abs(v) > 1e-9)
            {
                MessageErreur = $"Valeur trop élevée ({v:E3}). Limite : ±1e-9.";
                return;
            }
            if (DateJulienne < 50000 || DateJulienne > 80000)
            {
                MessageErreur = "Date julienne hors plage autorisée (50000 – 80000).";
                return;
            }

            ValeurJourResultat = v;
            DateJulienneResultat = DateJulienne;
            CloseAction?.Invoke(true);
        }

        [RelayCommand]
        private void Annuler() => CloseAction?.Invoke(false);
    }
}
