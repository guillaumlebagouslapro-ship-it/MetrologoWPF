using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services;
using Metrologo.Services.Incertitude;
using Metrologo.Services.Journal;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// VM de la fenêtre admin "Gérer les modules d'incertitude".
    /// CRUD sur les fichiers CSV de <c>%LocalAppData%\Metrologo\Incertitudes\</c>.
    /// </summary>
    public partial class GestionModulesIncertitudeViewModel : ObservableObject
    {
        public ObservableCollection<ModuleIncertitude> Modules { get; } = new();

        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(SupprimerCommand))]
        [NotifyCanExecuteChangedFor(nameof(EnregistrerCommand))]
        [NotifyCanExecuteChangedFor(nameof(AjouterTrioCommand))]
        [NotifyCanExecuteChangedFor(nameof(SupprimerLignesCommand))]
        [NotifyCanExecuteChangedFor(nameof(CopierVersCommand))]
        [NotifyPropertyChangedFor(nameof(InfosFonctions))]
        private ModuleIncertitude? _moduleSelectionne;

        [ObservableProperty] private string _statut = "Prêt.";

        // ------- Catégorie (sous-dossier) actuellement consultée -------

        /// <summary>
        /// Type de mesure actuellement sélectionné dans le ComboBox « Catégorie ».
        /// Définit le sous-dossier consulté (ex. <c>Incertitudes\Frequence\</c>) — chaque
        /// type a ses propres modules. Recharge la liste à chaque changement.
        /// </summary>
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(CheminDossier))]
        [NotifyPropertyChangedFor(nameof(LibelleCategorie))]
        private TypeMesure _typeMesureSelectionne = TypeMesure.Frequence;

        partial void OnTypeMesureSelectionneChanged(TypeMesure value)
        {
            // Reset de la sélection — le module actif n'a pas de sens dans une autre catégorie.
            ModuleSelectionne = null;
            Recharger();
        }

        /// <summary>Options affichées dans le ComboBox « Catégorie » avec libellés conviviaux.</summary>
        public OptionTypeMesure[] CategoriesDisponibles { get; } = new[]
        {
            new OptionTypeMesure(TypeMesure.Frequence,       EnTetesMesureHelper.LibelleType(TypeMesure.Frequence)),
            new OptionTypeMesure(TypeMesure.FreqAvantInterv, EnTetesMesureHelper.LibelleType(TypeMesure.FreqAvantInterv)),
            new OptionTypeMesure(TypeMesure.FreqFinale,      EnTetesMesureHelper.LibelleType(TypeMesure.FreqFinale)),
            new OptionTypeMesure(TypeMesure.Stabilite,       EnTetesMesureHelper.LibelleType(TypeMesure.Stabilite)),
            new OptionTypeMesure(TypeMesure.Interval,        EnTetesMesureHelper.LibelleType(TypeMesure.Interval)),
            new OptionTypeMesure(TypeMesure.TachyContact,    EnTetesMesureHelper.LibelleType(TypeMesure.TachyContact)),
            new OptionTypeMesure(TypeMesure.Stroboscope,     EnTetesMesureHelper.LibelleType(TypeMesure.Stroboscope)),
        };

        public string LibelleCategorie => EnTetesMesureHelper.LibelleType(TypeMesureSelectionne);

        public string InfosFonctions
        {
            get
            {
                if (ModuleSelectionne == null) return "";
                var fns = ModuleSelectionne.FonctionsSupportees.ToList();
                return fns.Count == 0 ? "Aucune ligne — clique « Ajouter un temps de mesure » pour démarrer."
                                      : "Fonctions : " + string.Join(", ", fns);
            }
        }

        public string CheminDossier => ModulesIncertitudeService.DossierComplet(TypeMesureSelectionne);

        public int NbTempsDistincts => ModuleSelectionne?.Lignes
            .Select(l => new { l.Fonction, l.TempsDeMesure }).Distinct().Count() ?? 0;
        public int NbLignes => ModuleSelectionne?.Lignes.Count ?? 0;

        /// <summary>Liste fermée des fonctions affichées dans le ComboBox de la colonne Fonction.</summary>
        public string[] FonctionsDisponibles { get; } = new[]
        {
            "Freq", "FreqAv", "FreqFin", "Stab", "Interv", "TachyC", "Strobo"
        };

        /// <summary>Liste éditable des temps de mesure standards (en secondes) — l'admin
        /// peut sélectionner ou taper sa propre valeur pour des cas atypiques.</summary>
        public double[] TempsDeMesureSuggeres { get; } = new[]
        {
            0.01, 0.02, 0.05, 0.1, 0.2, 0.5, 1d, 2d, 5d, 10d, 20d, 50d, 100d
        };

        public GestionModulesIncertitudeViewModel()
        {
            Recharger();
        }

        [RelayCommand]
        private void Recharger()
        {
            Modules.Clear();
            foreach (var m in ModulesIncertitudeService.Lister(TypeMesureSelectionne))
                Modules.Add(m);
            Statut = $"{Modules.Count} module(s) chargé(s) — catégorie « {LibelleCategorie} ».";
        }

        [RelayCommand]
        private void Ajouter()
        {
            var dlg = new Views.AjoutModuleIncertitudeDialog { Owner = Application.Current.MainWindow };
            if (dlg.ShowDialog() != true) return;

            string num = dlg.NumModuleSaisi;
            string nom = dlg.NomAffichageSaisi;
            if (string.IsNullOrWhiteSpace(num)) return;

            if (Modules.Any(m => string.Equals(m.NumModule, num, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show(
                    $"Un module « {num} » existe déjà dans la catégorie « {LibelleCategorie} ».",
                    "Doublon", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var nouveau = new ModuleIncertitude { NumModule = num, NomAffichage = nom };
            ModulesIncertitudeService.Sauvegarder(nouveau, TypeMesureSelectionne);
            Modules.Add(nouveau);
            ModuleSelectionne = nouveau;
            Statut = $"Module {num} créé dans « {LibelleCategorie} ». Ajoute des lignes puis Enregistre.";
        }

        [RelayCommand(CanExecute = nameof(PeutEnregistrerSupprimer))]
        private void Enregistrer()
        {
            if (ModuleSelectionne == null) return;
            try
            {
                ModulesIncertitudeService.Sauvegarder(ModuleSelectionne, TypeMesureSelectionne);
                Statut = $"Module {ModuleSelectionne.NumModule} enregistré ({ModuleSelectionne.Lignes.Count} lignes).";
                OnPropertyChanged(nameof(InfosFonctions));
                OnPropertyChanged(nameof(NbTempsDistincts));
                OnPropertyChanged(nameof(NbLignes));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Enregistrement échoué : {ex.Message}",
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        [RelayCommand(CanExecute = nameof(PeutEnregistrerSupprimer))]
        private void Supprimer()
        {
            if (ModuleSelectionne == null) return;
            var conf = MessageBox.Show(
                $"Supprimer le module « {ModuleSelectionne.NumModule} » de la catégorie « {LibelleCategorie} » ?\nLe fichier CSV sera effacé.",
                "Confirmer", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (conf != MessageBoxResult.Yes) return;

            try
            {
                ModulesIncertitudeService.Supprimer(ModuleSelectionne.NumModule, TypeMesureSelectionne);
                Modules.Remove(ModuleSelectionne);
                ModuleSelectionne = null;
                Statut = "Module supprimé.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Suppression échouée : {ex.Message}",
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // AjouterLigne (1 ligne unique) retiré — l'admin ajoute désormais des temps de
        // mesure via AjouterTrio (3 plages auto). Évite des lignes orphelines hors du
        // schéma "1 temps de mesure = 3 plages" du format C.E.A.O.

        /// <summary>
        /// Ajoute en une seule fois 3 lignes correspondant aux 3 plages de fréquence
        /// typiques d'un temps de mesure donné (basse / moyenne / haute) — accélère la
        /// saisie quand on suit le format du tableau papier C.E.A.O.
        /// </summary>
        [RelayCommand(CanExecute = nameof(PeutEnregistrerSupprimer))]
        private void AjouterTrio()
        {
            if (ModuleSelectionne == null) return;

            var dlg = new Views.AjoutTrioDialog(ModuleSelectionne) { Owner = Application.Current.MainWindow };
            if (dlg.ShowDialog() != true) return;

            string fn = dlg.Fonction;
            double t = dlg.TempsDeMesure;

            // 3 plages typiques C.E.A.O. : basse / moyenne / haute. L'admin ajuste
            // ensuite les bornes et incertitudes ligne par ligne.
            ModuleSelectionne.Lignes.Add(new LigneIncertitude
            {
                Fonction = fn, TempsDeMesure = t,
                BorneBasse = 0, BorneHaute = 10000.01
            });
            ModuleSelectionne.Lignes.Add(new LigneIncertitude
            {
                Fonction = fn, TempsDeMesure = t,
                BorneBasse = 10000.01, BorneHaute = 1000001.0
            });
            ModuleSelectionne.Lignes.Add(new LigneIncertitude
            {
                Fonction = fn, TempsDeMesure = t,
                BorneBasse = 1000001.0, BorneHaute = 1000000100.0
            });
            OnPropertyChanged(nameof(InfosFonctions));
            OnPropertyChanged(nameof(NbTempsDistincts));
            OnPropertyChanged(nameof(NbLignes));
            Statut = $"3 lignes ajoutées pour {fn} @ {t} s — ajuste les bornes et coefficients.";
        }

        /// <summary>Supprime les lignes sélectionnées dans le DataGrid.</summary>
        [RelayCommand(CanExecute = nameof(PeutEnregistrerSupprimer))]
        private void SupprimerLignes(System.Collections.IList? selection)
        {
            if (ModuleSelectionne == null || selection == null) return;
            var aSupprimer = selection.Cast<LigneIncertitude>().ToList();
            foreach (var l in aSupprimer) ModuleSelectionne.Lignes.Remove(l);
            OnPropertyChanged(nameof(InfosFonctions));
            OnPropertyChanged(nameof(NbTempsDistincts));
            OnPropertyChanged(nameof(NbLignes));
            Statut = $"{aSupprimer.Count} ligne(s) supprimée(s).";
        }

        /// <summary>
        /// Copie le module sélectionné vers une autre catégorie. Utile quand un module
        /// physique (ex. compteur de fréquence) est valide pour plusieurs types de mesure
        /// (Fréquence + FreqAvantInterv + FreqFinale par exemple). L'admin choisit la
        /// catégorie cible dans un mini-dialog.
        /// </summary>
        [RelayCommand(CanExecute = nameof(PeutEnregistrerSupprimer))]
        private void CopierVers()
        {
            if (ModuleSelectionne == null) return;

            var dlg = new Views.CopierVersDialog(TypeMesureSelectionne, ModuleSelectionne.NumModule)
            {
                Owner = Application.Current.MainWindow
            };
            if (dlg.ShowDialog() != true) return;

            var cible = dlg.CategorieChoisie;

            // Doublon dans la cible → confirmation avant écrasement.
            string ciblePath = System.IO.Path.Combine(
                ModulesIncertitudeService.DossierComplet(cible),
                ModuleSelectionne.NumModule + ".csv");
            if (System.IO.File.Exists(ciblePath))
            {
                var conf = MessageBox.Show(
                    $"Un module « {ModuleSelectionne.NumModule} » existe déjà dans "
                  + $"« {EnTetesMesureHelper.LibelleType(cible)} ».\nÉcraser ?",
                    "Doublon", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (conf != MessageBoxResult.Yes) return;
            }

            try
            {
                // Important : enregistrer d'abord le module courant pour que les éventuelles
                // modifs en cours soient incluses dans la copie (sinon la copie part du
                // disque, ignore le travail non sauvé).
                ModulesIncertitudeService.Sauvegarder(ModuleSelectionne, TypeMesureSelectionne);
                ModulesIncertitudeService.Copier(ModuleSelectionne.NumModule,
                    TypeMesureSelectionne, cible);

                Statut = $"Module {ModuleSelectionne.NumModule} copié vers "
                       + $"« {EnTetesMesureHelper.LibelleType(cible)} ».";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Copie échouée : {ex.Message}",
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        [RelayCommand]
        private void OuvrirDossier()
        {
            try
            {
                Directory.CreateDirectory(CheminDossier);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = CheminDossier,
                    UseShellExecute = true
                });
            }
            catch { /* silencieux */ }
        }

        private bool PeutEnregistrerSupprimer() => ModuleSelectionne != null;
    }

    /// <summary>
    /// Option du ComboBox « Catégorie » : couple (TypeMesure, libellé affichable).
    /// Utilise <c>Type</c> comme clé pour le binding à <c>TypeMesureSelectionne</c>.
    /// </summary>
    public record OptionTypeMesure(TypeMesure Type, string Libelle);
}
