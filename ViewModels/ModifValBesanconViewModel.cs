using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services.Besancon;
using System;
using System.Globalization;
using System.Linq;

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
            // Lit la valeur journalière déjà stockée (suivi partagé besancon-suivi.json) pour le
            // rubidium actif + la date julienne sélectionnée. Vide si aucune valeur enregistrée.
            var rub = EtatApplication.RubidiumActif;
            if (rub == null) { ValeurJourTexte = ""; return; }

            var donnees = BesanconStore.Charger();
            var existante = donnees.Journalieres
                .FirstOrDefault(x => x.RubidiumId == rub.Id && x.Mjd == DateJulienne);
            ValeurJourTexte = existante != null
                ? existante.Valeur.ToString(CultureInfo.InvariantCulture)
                : "";
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

            // Persiste la correction dans le suivi partagé (pour le rubidium actif) — ainsi la
            // valeur saisie/corrigée est consultable et entre dans les moyennes hebdo.
            var rub = EtatApplication.RubidiumActif;
            if (rub != null)
            {
                var donnees = BesanconStore.Charger();
                BesanconStore.UpsertValeurJournaliere(donnees, rub.Id, DateJulienne, v);
                BesanconStore.Sauvegarder(donnees);
            }

            CloseAction?.Invoke(true);
        }

        [RelayCommand]
        private void Annuler() => CloseAction?.Invoke(false);
    }
}
