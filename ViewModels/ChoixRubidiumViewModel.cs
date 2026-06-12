using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Views;
using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Windows;

namespace Metrologo.ViewModels
{
    public partial class ChoixRubidiumViewModel : ObservableObject
    {
        [ObservableProperty] private ObservableCollection<Rubidium> _rubidiums = new();
        [ObservableProperty] private Rubidium? _rubidiumSelectionne;

        // ---------- Mode Réglage manuel ----------

        // si vrai, l'admin tape directement une fréquence au lieu de choisir un rubidium
        // du catalogue (pas de rubidium raccordé, tests, fréquence de référence custom)
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(EstModeListe))]
        private bool _modeManuel;

        // inverse de ModeManuel, pour les bindings XAML
        public bool EstModeListe => !ModeManuel;

        // fréquence saisie en mode manuel (texte, pour un parsing tolérant)
        [ObservableProperty] private string _frequenceManuelleTexte = "10000000";

        // ---------- Erreurs / résultat ----------

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _messageErreur = string.Empty;

        public bool HasError => !string.IsNullOrEmpty(MessageErreur);

        public Rubidium? Resultat { get; private set; }
        public Action<bool>? CloseAction { get; set; }

        public bool ListeVide => Rubidiums.Count == 0;

        public ChoixRubidiumViewModel()
        {
            ChargerRubidiums();
            RestaurerEtatCourant();
        }

        /// <summary>Charge le catalogue depuis Preferences (seed par défaut si settings.json n'a encore rien).</summary>
        private void ChargerRubidiums()
        {
            Rubidiums.Clear();
            foreach (var r in Preferences.CatalogueRubidiums)
            {
                Rubidiums.Add(new Rubidium
                {
                    Id = r.Id,
                    Designation = r.Designation,
                    FrequenceMoyenne = r.FrequenceMoyenne,
                });
            }
            OnPropertyChanged(nameof(ListeVide));
        }

        // À l'ouverture, coche le bon mode et pré-sélectionne le rubidium en cours
        // (sinon le 1er du catalogue). Sans ça la fenêtre s'ouvrirait toujours en mode
        // Catalogue même quand la mesure tourne en Réglage manuel.
        private void RestaurerEtatCourant()
        {
            var actif = EtatApplication.RubidiumActif;
            if (actif != null && actif.EstReglageManuel)
            {
                ModeManuel = true;
                // valeur exacte, pas tronquée à 2 décimales : on doit retrouver
                // la fréquence de référence telle qu'elle a été saisie
                FrequenceManuelleTexte = actif.FrequenceMoyenne.ToString(
                    CultureInfo.InvariantCulture);
                RubidiumSelectionne = Rubidiums.FirstOrDefault();
                return;
            }

            ModeManuel = false;
            if (actif != null)
            {
                RubidiumSelectionne = Rubidiums.FirstOrDefault(r => r.Id == actif.Id)
                                      ?? Rubidiums.FirstOrDefault();
            }
            else
            {
                RubidiumSelectionne = Rubidiums.FirstOrDefault();
            }
        }

        // ouvre la gestion du catalogue (CRUD) puis recharge la liste en gardant
        // la sélection courante si possible
        [RelayCommand]
        private void GererCatalogue()
        {
            var win = new GestionCatalogueRubidiumsWindow
            {
                Owner = Application.Current.Windows
                    .OfType<Window>().FirstOrDefault(w => w.IsActive)
            };
            if (win.ShowDialog() == true)
            {
                int? idCourant = RubidiumSelectionne?.Id;
                ChargerRubidiums();
                RubidiumSelectionne = idCourant.HasValue
                    ? Rubidiums.FirstOrDefault(r => r.Id == idCourant.Value) ?? Rubidiums.FirstOrDefault()
                    : Rubidiums.FirstOrDefault();
            }
        }

        [RelayCommand]
        private void Valider()
        {
            MessageErreur = string.Empty;

            if (ModeManuel)
            {
                // on parse la fréquence et on fabrique un Rubidium factice avec
                // EstReglageManuel=true (repris par NomAffichage et le badge de l'accueil)
                string txt = (FrequenceManuelleTexte ?? "").Trim().Replace(',', '.');
                if (!double.TryParse(txt, NumberStyles.Float, CultureInfo.InvariantCulture, out double freq) || freq <= 0)
                {
                    MessageErreur = $"Fréquence invalide : « {FrequenceManuelleTexte} ». Saisis une valeur positive en Hz (ex. 10000000).";
                    return;
                }

                Resultat = new Rubidium
                {
                    Id = 0,
                    Designation = "Réglage manuel",
                    FrequenceMoyenne = freq,
                    EstReglageManuel = true,
                };
                CloseAction?.Invoke(true);
                return;
            }

            if (RubidiumSelectionne == null)
            {
                MessageErreur = "Veuillez sélectionner un rubidium.";
                return;
            }

            Resultat = RubidiumSelectionne;
            Resultat.EstReglageManuel = false;
            CloseAction?.Invoke(true);
        }

        [RelayCommand]
        private void Annuler() => CloseAction?.Invoke(false);
    }
}
