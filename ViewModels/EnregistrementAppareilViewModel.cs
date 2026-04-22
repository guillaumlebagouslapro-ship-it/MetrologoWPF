using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services.Catalogue;
using Metrologo.Services.Journal;
using System;
using System.Threading.Tasks;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// VM de la fenêtre d'enregistrement d'un nouvel appareil au catalogue local.
    /// Pré-rempli à partir de l'IDN détecté, l'utilisateur ajuste les commandes SCPI puis sauvegarde.
    /// </summary>
    public partial class EnregistrementAppareilViewModel : ObservableObject
    {
        private readonly AppareilDetecte? _detecte;
        private readonly ModeleAppareil? _modeleExistant;
        private readonly string _utilisateurActuel;

        public string Titre => _modeleExistant != null ? "Modifier un appareil" : "Enregistrer un appareil";
        public bool EstModification => _modeleExistant != null;

        [ObservableProperty] private string _nom = string.Empty;
        [ObservableProperty] private string _fabricantIdn = string.Empty;
        [ObservableProperty] private string _modeleIdn = string.Empty;

        [ObservableProperty] private string _chaineInit = "*RST;*CLS";
        [ObservableProperty] private string _confEntree = string.Empty;
        [ObservableProperty] private string _exeMesure = ":READ?";
        [ObservableProperty] private string _commandeGate = ":FREQ:ARM:STOP:TIM {0}";

        [ObservableProperty] private int _termWrite = 1;
        [ObservableProperty] private int _termRead = 10;
        [ObservableProperty] private int _tailleHeader = 1;

        [ObservableProperty] private bool _gereSrq;
        [ObservableProperty] private string _srqOn = string.Empty;
        [ObservableProperty] private string _srqOff = string.Empty;

        [ObservableProperty] private string _gatesTexte = "10 ms, 100 ms, 1 s, 10 s";
        [ObservableProperty] private string _entreesTexte = string.Empty;
        [ObservableProperty] private string _couplagesTexte = "AC, DC";

        [ObservableProperty] private string _idnDetecte = string.Empty;

        public Action<bool>? CloseAction { get; set; }
        public ModeleAppareil? Resultat { get; private set; }

        /// <summary>Création depuis un appareil fraîchement détecté (flux Diagnostic GPIB).</summary>
        public EnregistrementAppareilViewModel(AppareilDetecte detecte, string utilisateur)
        {
            _detecte = detecte;
            _utilisateurActuel = utilisateur;

            Nom = detecte.Modele ?? "Nouvel appareil";
            FabricantIdn = detecte.Fabricant ?? string.Empty;
            ModeleIdn = detecte.Modele ?? string.Empty;
            IdnDetecte = detecte.IdnBrut ?? string.Empty;

            // Pré-remplissage intelligent selon le modèle détecté
            if ((detecte.Modele ?? "").ToUpperInvariant().Contains("53131A"))
            {
                ChaineInit = "*RST;*CLS";
                ExeMesure = ":READ?";
                CommandeGate = ":FREQ:ARM:STOP:TIM {0}";
                EntreesTexte = "Canal 1, Canal 2, Canal 3";
                GatesTexte = "10 ms, 100 ms, 1 s, 10 s, 100 s";
            }
        }

        /// <summary>Édition d'un modèle existant (flux Admin « Gérer les appareils »).</summary>
        public EnregistrementAppareilViewModel(ModeleAppareil modeleExistant, string utilisateur)
        {
            _modeleExistant = modeleExistant;
            _utilisateurActuel = utilisateur;

            Nom = modeleExistant.Nom;
            FabricantIdn = modeleExistant.FabricantIdn;
            ModeleIdn = modeleExistant.ModeleIdn;
            IdnDetecte = $"(enregistré le {modeleExistant.DateCreation:dd/MM/yyyy} par {modeleExistant.CreePar})";

            ChaineInit = modeleExistant.Parametres.ChaineInit;
            ConfEntree = modeleExistant.Parametres.ConfEntree;
            ExeMesure = modeleExistant.Parametres.ExeMesure;
            CommandeGate = modeleExistant.Parametres.CommandeGate;
            TermWrite = modeleExistant.Parametres.TermWrite;
            TermRead = modeleExistant.Parametres.TermRead;
            TailleHeader = modeleExistant.Parametres.TailleHeader;
            GereSrq = modeleExistant.Parametres.GereSrq;
            SrqOn = modeleExistant.Parametres.SrqOn;
            SrqOff = modeleExistant.Parametres.SrqOff;

            GatesTexte = string.Join(", ", modeleExistant.Gates);
            EntreesTexte = string.Join(", ", modeleExistant.Entrees);
            CouplagesTexte = string.Join(", ", modeleExistant.Couplages);
        }

        [RelayCommand]
        private async Task EnregistrerAsync()
        {
            if (string.IsNullOrWhiteSpace(Nom))
                Nom = _detecte?.Modele ?? _modeleExistant?.Nom ?? "Modèle sans nom";

            if (_modeleExistant != null)
            {
                // Modification
                await CatalogueAppareilsService.Instance.ModifierAsync(_modeleExistant.Id, m =>
                {
                    m.Nom = Nom.Trim();
                    m.FabricantIdn = FabricantIdn.Trim();
                    m.ModeleIdn = ModeleIdn.Trim();
                    m.Parametres = ConstruireParametres();
                    m.Gates = SplitCSV(GatesTexte);
                    m.Entrees = SplitCSV(EntreesTexte);
                    m.Couplages = SplitCSV(CouplagesTexte);
                });

                JournalLog.Info(CategorieLog.Administration, "CATALOGUE_MODIF",
                    $"Modèle « {Nom} » modifié par {_utilisateurActuel}.",
                    new { Id = _modeleExistant.Id, Nom });

                Resultat = _modeleExistant;
            }
            else
            {
                // Création
                var modele = new ModeleAppareil
                {
                    Nom = Nom.Trim(),
                    FabricantIdn = FabricantIdn.Trim(),
                    ModeleIdn = ModeleIdn.Trim(),
                    Parametres = ConstruireParametres(),
                    Gates = SplitCSV(GatesTexte),
                    Entrees = SplitCSV(EntreesTexte),
                    Couplages = SplitCSV(CouplagesTexte),
                    DateCreation = DateTime.Now,
                    CreePar = _utilisateurActuel
                };

                await CatalogueAppareilsService.Instance.AjouterAsync(modele);

                JournalLog.Info(CategorieLog.Configuration, "CATALOGUE_AJOUT",
                    $"Modèle « {modele.Nom } » ajouté au catalogue par {_utilisateurActuel}.",
                    new { modele.Nom, modele.FabricantIdn, modele.ModeleIdn });

                Resultat = modele;
            }

            CloseAction?.Invoke(true);
        }

        private ParametresIeee ConstruireParametres() => new()
        {
            ChaineInit = ChaineInit ?? string.Empty,
            ConfEntree = ConfEntree ?? string.Empty,
            ExeMesure = ExeMesure ?? string.Empty,
            CommandeGate = CommandeGate ?? string.Empty,
            TermWrite = TermWrite,
            TermRead = TermRead,
            TailleHeader = TailleHeader,
            GereSrq = GereSrq,
            SrqOn = SrqOn ?? string.Empty,
            SrqOff = SrqOff ?? string.Empty
        };

        [RelayCommand]
        private void Annuler() => CloseAction?.Invoke(false);

        private static System.Collections.Generic.List<string> SplitCSV(string texte)
        {
            var liste = new System.Collections.Generic.List<string>();
            if (string.IsNullOrWhiteSpace(texte)) return liste;
            foreach (var part in texte.Split(','))
            {
                var t = part.Trim();
                if (!string.IsNullOrEmpty(t)) liste.Add(t);
            }
            return liste;
        }
    }
}
