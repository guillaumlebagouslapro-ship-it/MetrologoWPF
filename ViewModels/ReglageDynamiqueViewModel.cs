using CommunityToolkit.Mvvm.ComponentModel;
using Metrologo.Models;
using System.Globalization;
using System.Linq;

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
                // Quand aucune valeur n'est prévue, on laisse le champ vide : une TextBox vide dit
                // bien à l'utilisateur "rien ne sera appliqué" (et CommandeSelectionnee renverra
                // alors null). Préremplir avec "0" prêterait à confusion — "0 V" est une valeur tout
                // à fait valable, et on ne veut pas l'envoyer par mégarde si l'utilisateur n'a rien tapé.
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

                    // Si l'utilisateur n'a rien saisi, on envoie quand même un 0 explicite plutôt
                    // que de zapper le réglage : sur certains compteurs (le 53131A), un trigger
                    // laissé non configuré bascule en mode AUTO, ce qu'on ne veut surtout pas en
                    // métrologie. La convention est donc : pas de saisie = on force la valeur 0.
                    string valeurNum = string.IsNullOrWhiteSpace(Valeur) ? "0" : Valeur.Replace(',', '.');
                    if (!double.TryParse(valeurNum, NumberStyles.Float, CultureInfo.InvariantCulture, out var v)) return null;

                    return string.Format(CultureInfo.InvariantCulture, template, v);
                }

                return OptionSelectionnee?.CommandeScpi;
            }
        }

        /// <summary>
        /// Restaure la sélection de ce réglage depuis les commandes SCPI mémorisées
        /// (<c>MesureConfig.CommandesScpiReglages</c>), pour que la fenêtre Configuration
        /// retrouve les choix de l'utilisateur à sa réouverture. Les deux types sont gérés :
        ///  - <b>Choix</b> : on re-coche l'option dont la commande est présente dans la liste.
        ///  - <b>Numérique</b> (le Trigger, par ex.) : on repère la commande qui colle au template
        ///    <c>"préfixe{0}suffixe"</c>, puis on en réextrait la valeur pour réalimenter le champ.
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

            // Cas numérique : on reconstitue le motif "préfixe{0}suffixe" afin d'extraire la
            // valeur de la commande mémorisée correspondante.
            string? template = _source.Options.Count > 0 ? _source.Options[0].CommandeScpi : null;
            if (string.IsNullOrWhiteSpace(template)) return;

            // Attention : un template peut être CHAÎNÉ (ex. 53230A : ":INP1:LEV:AUTO OFF;:INP1:LEV {0}").
            // Au moment d'enregistrer, ConfigurationViewModel.SplitCommandesScpi l'a éclaté en
            // entrées séparées (":INP1:LEV:AUTO OFF", ":INP1:LEV 1"). Pour retomber sur la bonne
            // entrée, le préfixe/suffixe doit être dérivé du SEUL sous-template qui porte le {0}
            // (":INP1:LEV {0}"), et non du template chaîné complet — sinon aucun match et la valeur est perdue.
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
        /// Dans un template SCPI possiblement chaîné ("a;:b{0};:c"), extrait le seul sous-template
        /// qui contient le marqueur <c>{0}</c>. On reproduit ici exactement le découpage de
        /// <c>ConfigurationViewModel.SplitCommandesScpi</c> (séparateur ";:", et le ":" initial
        /// remis sur les parties suivantes), pour que le préfixe/suffixe obtenu colle au caractère
        /// près à la commande mémorisée après split. Si le template n'est pas chaîné, on le renvoie tel quel.
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
