using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace Metrologo.ViewModels
{
    public partial class SaisieFreqAutresViewModel : ObservableObject
    {
        [ObservableProperty] private string _labelRubi1 = "Rubidium 1";
        [ObservableProperty] private string _labelRubi2 = "Rubidium 2";

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _valeurRubi1Texte = "10000000";

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _valeurRubi2Texte = "10000000";

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _messageErreur = string.Empty;

        public bool HasError => !string.IsNullOrEmpty(MessageErreur);

        public int IdRubi1 { get; private set; }
        public int IdRubi2 { get; private set; }
        public double ValRubi1 { get; private set; }
        public double ValRubi2 { get; private set; }
        public Action<bool>? CloseAction { get; set; }

        public SaisieFreqAutresViewModel() : this(ChargerRubidiumsInactifsSimules()) { }

        public SaisieFreqAutresViewModel(IList<Rubidium> rubidiumsInactifs)
        {
            if (rubidiumsInactifs.Count < 2) return;

            IdRubi1 = rubidiumsInactifs[0].Id;
            LabelRubi1 = rubidiumsInactifs[0].Designation;

            IdRubi2 = rubidiumsInactifs[1].Id;
            LabelRubi2 = rubidiumsInactifs[1].Designation;
        }

        private static IList<Rubidium> ChargerRubidiumsInactifsSimules()
        {
            // TODO : requête SQL : Select RUB_ID, RUB_DESIGNATION from TR_METROLOGO_RUBIDIUMS where RUB_ACTIF=0 order by RUB_ID
            return new List<Rubidium>
            {
                new() { Id = 2, Designation = "FS725 - SN 67890" },
                new() { Id = 3, Designation = "LPRO-101 - SN 54321" }
            };
        }

        [RelayCommand]
        private void Valider()
        {
            if (!TryParse(ValeurRubi1Texte, out var v1) || v1 <= 0)
            {
                MessageErreur = $"Valeur incorrecte pour {LabelRubi1}.";
                return;
            }
            if (!TryParse(ValeurRubi2Texte, out var v2) || v2 <= 0)
            {
                MessageErreur = $"Valeur incorrecte pour {LabelRubi2}.";
                return;
            }

            ValRubi1 = v1;
            ValRubi2 = v2;
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
