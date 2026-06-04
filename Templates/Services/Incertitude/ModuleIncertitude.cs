using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace Metrologo.Services.Incertitude
{
    /// <summary>
    /// Module métrologique (compteur de fréquence, tachymètre, etc.) avec son tableau
    /// d'incertitudes. Chaque module = 1 fichier CSV dans le dossier
    /// <c>%LocalAppData%\Metrologo\Incertitudes\</c>. Le nom du fichier (sans extension)
    /// donne le numéro/identifiant du module.
    ///
    /// Un module supporte une ou plusieurs fonctions : si son CSV contient des lignes
    /// avec différentes valeurs de <c>Fonction</c>, le module est utilisable pour ces
    /// types de mesure. C'est le filtre du ComboBox dans la fenêtre Configuration.
    /// </summary>
    public partial class ModuleIncertitude : ObservableObject
    {
        /// <summary>Identifiant du module (= nom du fichier CSV sans extension). Ex: "MF51901A".
        /// Éditable dans la gestion des modules (un renommage = renommage du fichier CSV).</summary>
        [ObservableProperty] private string _numModule = string.Empty;

        /// <summary>Nom/commentaire lisible affiché dans les listes UI (ex. « Compteur de fréquence
        /// ou Tachymètre »). Optionnel — fallback sur NumModule. Éditable.</summary>
        [ObservableProperty] private string _nomAffichage = string.Empty;

        partial void OnNumModuleChanged(string value) => OnPropertyChanged(nameof(LibelleAffiche));
        partial void OnNomAffichageChanged(string value) => OnPropertyChanged(nameof(LibelleAffiche));

        /// <summary>
        /// Faux pour les modules dont l'incertitude ne dépend pas d'une porte temporelle —
        /// typiquement les tachymètres à contacts et les stroboscopes, qui comptent par
        /// événement plutôt que par fenêtre de comptage. Quand faux, la colonne Temps est
        /// masquée dans l'UI et <see cref="Trouver"/> ignore cette dimension.
        /// </summary>
        public bool UtiliseTempsDeMesure { get; set; } = true;

        /// <summary>
        /// Toutes les lignes du tableau d'incertitudes de ce module.
        /// <see cref="ObservableCollection{T}"/> pour que le DataGrid de l'UI admin
        /// se rafraîchisse automatiquement à chaque ajout/suppression.
        /// </summary>
        public ObservableCollection<LigneIncertitude> Lignes { get; set; } = new();

        /// <summary>
        /// Liste des fonctions distinctes supportées par ce module — déduite des lignes
        /// du CSV. Permet de filtrer le ComboBox "Module" dans Configuration selon le
        /// type de mesure choisi.
        /// </summary>
        public IEnumerable<string> FonctionsSupportees =>
            Lignes.Select(l => l.Fonction).Distinct(System.StringComparer.OrdinalIgnoreCase);

        /// <summary>Vrai si ce module supporte la fonction donnée.</summary>
        public bool SupportFonction(string fonction) =>
            Lignes.Any(l => string.Equals(l.Fonction, fonction, System.StringComparison.OrdinalIgnoreCase));

        /// <summary>
        /// Cherche la ligne qui matche (Fonction, Temps de mesure, Fréquence dans la plage).
        /// Retourne null si aucune ligne ne couvre cette combinaison — l'appelant doit gérer
        /// (typiquement : laisser CoeffA/CoeffB à 0 et logger un avertissement).
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
