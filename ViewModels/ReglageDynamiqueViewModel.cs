using CommunityToolkit.Mvvm.ComponentModel;
using Metrologo.Models;
using System.Globalization;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// VM d'un réglage dynamique, tel qu'affiché dans la fenêtre Configuration.
    /// L'UI le rend de trois façons possibles :
    ///   - <see cref="EstBinaire"/> : deux RadioButtons côte à côte (par ex. 50 Ω / 1 MΩ, ou AC / DC).
    ///   - <see cref="EstChoixMultiple"/> : une ComboBox classique (3 options ou plus, par ex. le mode de mesure).
    ///   - <see cref="EstNumerique"/> : une TextBox avec son unité (par ex. le Trigger en volts).
    /// </summary>
    public partial class ReglageDynamiqueViewModel : ObservableObject
    {
        private readonly ReglageAppareil _source;

        public string Nom { get; }
        public System.Collections.Generic.IReadOnlyList<OptionReglage> Options { get; }
        public string Unite { get; }

        public bool EstChoix => _source.Type == TypeReglage.Choix;
        public bool EstNumerique => _source.Type == TypeReglage.Numerique;

        /// <summary>Choix à exactement 2 options : on l'affiche en RadioButtons plutôt qu'en ComboBox.</summary>
        public bool EstBinaire => EstChoix && Options.Count == 2;

        /// <summary>Choix à 3 options ou plus : on l'affiche en ComboBox.</summary>
        public bool EstChoixMultiple => EstChoix && Options.Count > 2;

        [ObservableProperty]
        private OptionReglage? _optionSelectionnee;

        [ObservableProperty]
        private string _valeur = string.Empty;

        partial void OnOptionSelectionneeChanged(OptionReglage? value)
        {
            OnPropertyChanged(nameof(EstPremiereOption));
            OnPropertyChanged(nameof(EstSecondeOption));
        }

        public ReglageDynamiqueViewModel(ReglageAppareil source)
        {
            _source = source;
            Nom = source.Nom;
            Options = source.Options;
            Unite = source.Unite;

            if (source.Type == TypeReglage.Choix && source.Options.Count > 0)
            {
                OptionSelectionnee = source.Options[0];
            }
            else if (source.Type == TypeReglage.Numerique)
            {
                // Sans valeur par défaut, le champ reste vide (pas de "0" prérempli, qui serait une
                // valeur valable ambiguë). Voir CommandeSelectionnee pour le traitement d'une saisie vide.
                Valeur = source.ValeurDefaut ?? string.Empty;
            }
        }

        /// <summary>Binding TwoWay du premier RadioButton, pour un réglage binaire.</summary>
        public bool EstPremiereOption
        {
            get => Options.Count > 0 && ReferenceEquals(OptionSelectionnee, Options[0]);
            set { if (value && Options.Count > 0) OptionSelectionnee = Options[0]; }
        }

        /// <summary>Binding TwoWay du second RadioButton, pour un réglage binaire.</summary>
        public bool EstSecondeOption
        {
            get => Options.Count > 1 && ReferenceEquals(OptionSelectionnee, Options[1]);
            set { if (value && Options.Count > 1) OptionSelectionnee = Options[1]; }
        }

        public string LibellePremiereOption => Options.Count > 0 ? Options[0].Libelle : string.Empty;
        public string LibelleSecondeOption  => Options.Count > 1 ? Options[1].Libelle : string.Empty;

        /// <summary>GroupName unique pour les RadioButtons du binding binaire, histoire d'éviter que deux réglages se marchent dessus.</summary>
        public string GroupeRadio => "Reglage_" + Nom.Replace(' ', '_');

        /// <summary>
        /// La commande SCPI finale, celle qu'on va vraiment envoyer :
        /// - Choix : la <c>CommandeScpi</c> de l'option retenue.
        /// - Numerique : le template, avec <c>{0}</c> remplacé par <see cref="Valeur"/>.
        /// Renvoie null s'il n'y a rien à envoyer.
        /// </summary>
        public string? CommandeSelectionnee
        {
            get
            {
                if (_source.Type == TypeReglage.Numerique)
                {
                    string? template = _source.Options.Count > 0 ? _source.Options[0].CommandeScpi : null;
                    if (string.IsNullOrWhiteSpace(template)) return null;

                    // Pas de saisie = on force 0 plutôt que d'omettre le réglage : sur certains compteurs
                    // (53131A), un trigger non configuré bascule en mode AUTO, à éviter en métrologie.
                    string valeurNum = string.IsNullOrWhiteSpace(Valeur) ? "0" : Valeur.Replace(',', '.');
                    if (!double.TryParse(valeurNum, NumberStyles.Float, CultureInfo.InvariantCulture, out var v)) return null;

                    return string.Format(CultureInfo.InvariantCulture, template, v);
                }

                return OptionSelectionnee?.CommandeScpi;
            }
        }

        /// <summary>
        /// Restaure la sélection depuis les commandes SCPI mémorisées (<c>MesureConfig.CommandesScpiReglages</c>) :
        /// Choix = re-coche l'option dont la commande est présente ; Numérique = retrouve la commande
        /// correspondant au motif <c>"préfixe{0}suffixe"</c> et en réextrait la valeur.
        /// </summary>
        public void RestaurerDepuis(System.Collections.Generic.IEnumerable<string>? commandesPersistees)
        {
            if (commandesPersistees == null) return;
            var liste = commandesPersistees.Where(c => !string.IsNullOrEmpty(c)).ToList();
            if (liste.Count == 0) return;

            if (_source.Type == TypeReglage.Choix)
            {
                var opt = Options.FirstOrDefault(o => liste.Any(c => c == o.CommandeScpi));
                if (opt != null) OptionSelectionnee = opt;
                return;
            }

            // Cas numérique : motif "préfixe{0}suffixe" pour extraire la valeur mémorisée.
            string? template = _source.Options.Count > 0 ? _source.Options[0].CommandeScpi : null;
            if (string.IsNullOrWhiteSpace(template)) return;

            // Un template peut être CHAÎNÉ (53230A : ":INP1:LEV:AUTO OFF;:INP1:LEV {0}") et a été éclaté en
            // entrées séparées par ConfigurationViewModel.SplitCommandesScpi : préfixe/suffixe doivent
            // venir du seul sous-template portant {0}, sinon aucun match et la valeur est perdue.
            string sousTemplate = SousCommandeAvecValeur(template);
            int pos = sousTemplate.IndexOf("{0}", System.StringComparison.Ordinal);
            if (pos < 0) return;

            string prefixe = sousTemplate.Substring(0, pos);
            string suffixe = sousTemplate.Substring(pos + 3);

            foreach (var c in liste)
            {
                if (!c.StartsWith(prefixe, System.StringComparison.Ordinal)) continue;
                if (suffixe.Length > 0 && !c.EndsWith(suffixe, System.StringComparison.Ordinal)) continue;
                int longueur = c.Length - prefixe.Length - suffixe.Length;
                if (longueur <= 0) continue;
                Valeur = c.Substring(prefixe.Length, longueur).Trim();
                return;
            }
        }

        /// <summary>
        /// Extrait d'un template SCPI possiblement chaîné ("a;:b{0};:c") le seul sous-template contenant
        /// <c>{0}</c>. Reproduit exactement le découpage de <c>ConfigurationViewModel.SplitCommandesScpi</c>
        /// (séparateur ";:", ":" initial remis sur les parties suivantes). Template non chaîné : renvoyé tel quel.
        /// </summary>
        private static string SousCommandeAvecValeur(string template)
        {
            if (!template.Contains(";:")) return template;

            var parts = template.Split(";:", System.StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < parts.Length; i++)
            {
                string p = parts[i].Trim();
                if (p.Length == 0) continue;
                if (p.Contains("{0}"))
                    return i == 0 ? p : ":" + p;
            }
            return template;
        }
    }
}
