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
    public class ModuleIncertitude
    {
        /// <summary>Identifiant du module (= nom du fichier CSV sans extension). Ex: "MF51901A".</summary>
        public string NumModule { get; set; } = string.Empty;

        /// <summary>Nom lisible affiché dans les listes UI. Optionnel — fallback sur NumModule.</summary>
        public string NomAffichage { get; set; } = string.Empty;

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
                System.Math.Abs(l.TempsDeMesure - tempsSec) < 1e-9 &&
                l.Couvre(freqHz));
        }

        public string LibelleAffiche =>
            string.IsNullOrWhiteSpace(NomAffichage) ? NumModule : $"{NumModule} — {NomAffichage}";
    }
}
