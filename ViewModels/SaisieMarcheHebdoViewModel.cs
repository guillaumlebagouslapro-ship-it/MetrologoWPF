using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Globalization;
using System.Windows;

namespace Metrologo.ViewModels
{
    public partial class SaisieMarcheHebdoViewModel : ObservableObject
    {
        private static readonly DateTime MjdEpoch = new(1858, 11, 17);

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _marcheTexte = "0.0";

        [ObservableProperty] private DateTime _dateSelectionnee = DateTime.Today;
        [ObservableProperty] private int _dateJulienne;
        [ObservableProperty] private int _numeroSemaine;
        [ObservableProperty] private string _rubidiumActifTexte = "Rubidium : non défini";

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _messageErreur = string.Empty;

        public bool HasError => !string.IsNullOrEmpty(MessageErreur);

        public double ValeurMarcheSecondes { get; private set; }
        public int DateJulienneResultat { get; private set; }
        public Action<bool>? CloseAction { get; set; }

        public SaisieMarcheHebdoViewModel() : this("non défini") { }

        public SaisieMarcheHebdoViewModel(string nomRubidiumActif)
        {
            RubidiumActifTexte = $"Rubidium actif : {nomRubidiumActif}";
            RecalculerDates();
        }

        partial void OnDateSelectionneeChanged(DateTime value) => RecalculerDates();

        private void RecalculerDates()
        {
            DateJulienne = (int)(DateSelectionnee.Date - MjdEpoch).TotalDays;
            NumeroSemaine = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(
                DateSelectionnee, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }

        [RelayCommand]
        private void Valider()
        {
            var txt = (MarcheTexte ?? "").Trim().Replace(',', '.');
            if (!double.TryParse(txt, NumberStyles.Float, CultureInfo.InvariantCulture, out var valUs))
            {
                MessageErreur = $"Valeur incorrecte : '{MarcheTexte}'.";
                return;
            }

            var r = MessageBox.Show(
                $"Confirmez-vous que la valeur de la marche est {valUs} µs ?",
                "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (r != MessageBoxResult.Yes) return;

            ValeurMarcheSecondes = valUs * 1e-6;
            DateJulienneResultat = DateJulienne;
            CloseAction?.Invoke(true);
        }

        [RelayCommand]
        private void Annuler() => CloseAction?.Invoke(false);
    }
}
