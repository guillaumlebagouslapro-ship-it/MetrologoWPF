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

        /// <summary>
        /// Si vrai, l'admin saisit directement une fréquence moyenne plutôt que de
        /// choisir un rubidium du catalogue. Pour les cas où aucun rubidium n'est
        /// raccordé / pour des tests / pour des fréquences de référence custom.
        /// </summary>
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(EstModeListe))]
        private bool _modeManuel;

        /// <summary>Inverse de <see cref="ModeManuel"/> — utile pour bindings XAML.</summary>
        public bool EstModeListe => !ModeManuel;

        /// <summary>Fréquence saisie en mode manuel (texte pour parsing tolérant).</summary>
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

        /// <summary>
        /// Charge le catalogue depuis <see cref="Preferences"/> (avec seed par défaut
        /// si le fichier settings.json n'a pas encore de catalogue persisté).
        /// </summary>
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

        /// <summary>
        /// Au moment d'ouvrir la fenêtre, on coche le bon mode et on pré-sélectionne
        /// le rubidium actuellement utilisé (sinon le 1er du catalogue). Sans ça,
        /// la fenêtre s'ouvre toujours en mode Catalogue alors que la mesure tourne
        /// peut-être en Réglage manuel.
        /// </summary>
        private void RestaurerEtatCourant()
        {
            var actif = EtatApplication.RubidiumActif;
            if (actif != null && actif.EstReglageManuel)
            {
                ModeManuel = true;
                // Valeur exacte (toute la précision saisie), pas tronquée à 2 décimales :
                // au retour dans la fenêtre on retrouve la fréquence de référence telle quelle.
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

        /// <summary>
        /// Ouvre la fenêtre de gestion du catalogue (CRUD) puis recharge la liste
        /// en préservant si possible la sélection courante.
        /// </summary>
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
                // Saisie manuelle : on parse la fréquence et on crée un Rubidium "factice"
                // avec Designation="Réglage manuel" et EstReglageManuel=true (utilisé par
                // NomAffichage et le badge en bas de l'écran d'accueil).
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
