using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services.Catalogue;
using Metrologo.Services.Journal;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// Ligne éditable : un appareil legacy + son adresse GPIB.
    /// </summary>
    public partial class LigneAdresseLegacy : ObservableObject
    {
        public ModeleAppareil Modele { get; }
        public string Nom => Modele.Nom;

        [ObservableProperty] private int _adresse;

        public LigneAdresseLegacy(ModeleAppareil modele)
        {
            Modele = modele;
            _adresse = modele.Parametres.AdresseFixeParDefaut;
        }

        /// <summary>Recopie l'adresse saisie dans le modèle (avant sauvegarde).</summary>
        public void Appliquer() => Modele.Parametres.AdresseFixeParDefaut = Adresse;
    }

    /// <summary>
    /// VM admin : édition + enregistrement des adresses GPIB des appareils legacy
    /// (EIP / Racal / Stanford), pour pouvoir les brancher simultanément à des adresses
    /// distinctes. Réservé à l'Administration — les utilisateurs ne modifient pas ces valeurs.
    /// Persiste dans le fichier réseau via <see cref="SeedLegacyAppareils.Sauvegarder"/>.
    /// </summary>
    public partial class GestionAdressesLegacyViewModel : ObservableObject
    {
        public ObservableCollection<LigneAdresseLegacy> Lignes { get; } = new();

        public Action? CloseAction { get; set; }

        public GestionAdressesLegacyViewModel()
        {
            foreach (var m in CatalogueAppareilsService.Instance.Modeles
                         .Where(x => x.Parametres.Legacy)
                         .OrderBy(x => x.Nom))
            {
                Lignes.Add(new LigneAdresseLegacy(m));
            }
        }

        [RelayCommand]
        private void Enregistrer()
        {
            // 1) Bornes GPIB valides (1..30).
            foreach (var l in Lignes)
            {
                if (l.Adresse < 1 || l.Adresse > 30)
                {
                    MessageBox.Show(
                        $"L'adresse GPIB de « {l.Nom} » doit être comprise entre 1 et 30 (valeur : {l.Adresse}).",
                        "Adresse invalide", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            // 2) Doublons : on prévient (deux appareils sur la même adresse ne peuvent pas être
            //    branchés en même temps), mais on laisse le choix de continuer.
            var doublon = Lignes.GroupBy(l => l.Adresse).FirstOrDefault(g => g.Count() > 1);
            if (doublon != null)
            {
                var noms = string.Join(" et ", doublon.Select(l => l.Nom));
                var res = MessageBox.Show(
                    $"{noms} partagent l'adresse GPIB {doublon.Key}.\n\n"
                  + "Ils ne pourront pas être branchés sur le bus en même temps.\n\n"
                  + "Enregistrer quand même ?",
                    "Adresse en double", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (res != MessageBoxResult.Yes) return;
            }

            // 3) Applique + persiste sur le réseau.
            foreach (var l in Lignes) l.Appliquer();

            try
            {
                string fichier = SeedLegacyAppareils.Sauvegarder();
                MessageBox.Show(
                    $"Adresses GPIB enregistrées.\n\nFichier : {fichier}",
                    "Enregistré", MessageBoxButton.OK, MessageBoxImage.Information);
                CloseAction?.Invoke();
            }
            catch (Exception ex)
            {
                JournalLog.Erreur(CategorieLog.Administration, "APPAREILS_LEGACY_SAVE_ERR",
                    $"Échec sauvegarde adresses legacy : {ex.Message}");
                MessageBox.Show(
                    $"Impossible d'enregistrer les adresses :\n\n{ex.Message}\n\n"
                  + "Vérifie l'accès au partage réseau M:.",
                    "Erreur d'enregistrement", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        [RelayCommand]
        private void Annuler() => CloseAction?.Invoke();
    }
}
