using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace Metrologo.Services.Incertitude
{
    /// <summary>
    /// Module métrologique (compteur, tachymètre…) avec son tableau d'incertitudes.
    /// Un module = un CSV dans <c>%LocalAppData%\Metrologo\Incertitudes\</c> ; le nom
    /// du fichier (sans extension) est l'identifiant. Les valeurs distinctes de
    /// <c>Fonction</c> définissent les types de mesure supportés (filtre du ComboBox Configuration).
    /// </summary>
    public partial class ModuleIncertitude : ObservableObject
    {
        /// <summary>Identifiant du module (= nom du fichier CSV sans extension, ex: "MF51901A").
        /// Renommer ici = renommer le fichier CSV.</summary>
        [ObservableProperty] private string _numModule = string.Empty;

        /// <summary>Libellé affiché dans les listes UI. Optionnel — fallback sur NumModule.</summary>
        [ObservableProperty] private string _nomAffichage = string.Empty;

        partial void OnNumModuleChanged(string value) => OnPropertyChanged(nameof(LibelleAffiche));
        partial void OnNomAffichageChanged(string value) => OnPropertyChanged(nameof(LibelleAffiche));

        /// <summary>
        /// Faux pour les modules sans porte temporelle (tachy contact, stroboscope) —
        /// la colonne Temps est masquée dans l'UI et ignorée dans <see cref="Trouver"/>.
        /// </summary>
        public bool UtiliseTempsDeMesure { get; set; } = true;

        /// <summary>Lignes du tableau d'incertitudes. ObservableCollection pour le DataGrid admin.</summary>
        public ObservableCollection<LigneIncertitude> Lignes { get; set; } = new();

        /// <summary>Fonctions distinctes du CSV — filtre le ComboBox "Module" dans Configuration.</summary>
        public IEnumerable<string> FonctionsSupportees =>
            Lignes.Select(l => l.Fonction).Distinct(System.StringComparer.OrdinalIgnoreCase);

        /// <summary>Vrai si ce module supporte la fonction donnée.</summary>
        public bool SupportFonction(string fonction) =>
            Lignes.Any(l => string.Equals(l.Fonction, fonction, System.StringComparison.OrdinalIgnoreCase));

        /// <summary>
        /// Cherche la ligne (Fonction, Temps, plage Hz). Retourne null si aucune ne couvre
        /// la combinaison — l'appelant laisse CoeffA/CoeffB à 0 et logue un avertissement.
        /// </summary>
        public LigneIncertitude? Trouver(string fonction, double tempsSec, double freqHz)
        {
            return Lignes.FirstOrDefault(l =>
                string.Equals(l.Fonction, fonction, System.StringComparison.OrdinalIgnoreCase) &&
                (!UtiliseTempsDeMesure || System.Math.Abs(l.TempsDeMesure - tempsSec) < 1e-9) &&
                l.Couvre(freqHz));
        }

        public string LibelleAffiche =>
            string.IsNullOrWhiteSpace(NomAffichage) ? NumModule : $"{NumModule} — {NomAffichage}";
    }
}
