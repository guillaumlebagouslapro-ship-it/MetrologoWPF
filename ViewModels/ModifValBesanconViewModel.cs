using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services.Besancon;
using System;
using System.Globalization;
using System.Threading.Tasks;

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
            _ = ChargerValeurExistanteAsync();
        }

        private void Recalculer()
        {
            DateJulienne = (int)(DateSelectionnee.Date - MjdEpoch).TotalDays;
        }

        private async Task ChargerValeurExistanteAsync()
        {
            // Lit la valeur journalière déjà stockée en base (T_METROLOGO_DATESRUBIS) pour le
            // rubidium actif + la date julienne sélectionnée. Vide si aucune valeur enregistrée.
            var rub = EtatApplication.RubidiumActif;
            if (rub == null) { ValeurJourTexte = ""; return; }

            try
            {
                var v = await BesanconStore.LireValeurJournaliereAsync(rub.Id, DateJulienne);
                ValeurJourTexte = v.HasValue ? v.Value.ToString(CultureInfo.InvariantCulture) : "";
            }
            catch { ValeurJourTexte = ""; }
        }

        [RelayCommand]
        private async Task Valider()
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

            // Persiste la correction en base (T_METROLOGO_DATESRUBIS) pour le rubidium actif —
            // ainsi la valeur saisie/corrigée entre dans les moyennes hebdo.
            var rub = EtatApplication.RubidiumActif;
            if (rub != null)
            {
                try { await BesanconStore.UpsertValeurJournaliereAsync(rub.Id, DateJulienne, v); }
                catch (Exception ex)
                {
                    MessageErreur = $"Enregistrement en base échoué : {ex.Message}";
                    return;
                }
            }

            CloseAction?.Invoke(true);
        }

        [RelayCommand]
        private void Annuler() => CloseAction?.Invoke(false);
    }
}
