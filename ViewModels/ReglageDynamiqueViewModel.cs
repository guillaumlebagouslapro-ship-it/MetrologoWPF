using CommunityToolkit.Mvvm.ComponentModel;
using Metrologo.Models;
using System.Globalization;
using System.Linq;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// VM d'un réglage dynamique affiché dans la fenêtre Configuration.
    /// Trois modes de rendu dans l'UI :
    ///   - <see cref="EstBinaire"/> : deux RadioButtons côte à côte (ex: 50 Ω / 1 MΩ, AC / DC).
    ///   - <see cref="EstChoixMultiple"/> : ComboBox classique (3+ options, ex: Mode de mesure).
    ///   - <see cref="EstNumerique"/> : TextBox avec unité (ex: Trigger en volts).
    /// </summary>
    public partial class ReglageDynamiqueViewModel : ObservableObject
    {
        private readonly ReglageAppareil _source;

        public string Nom { get; }
        public System.Collections.Generic.IReadOnlyList<OptionReglage> Options { get; }
        public string Unite { get; }

        public bool EstChoix => _source.Type == TypeReglage.Choix;
        public bool EstNumerique => _source.Type == TypeReglage.Numerique;

        /// <summary>Choix à exactement 2 options — rendu en RadioButtons plutôt qu'en ComboBox.</summary>
        public bool EstBinaire => EstChoix && Options.Count == 2;

        /// <summary>Choix à 3+ options — rendu en ComboBox.</summary>
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
                // Laisse vide par défaut si aucune valeur n'est prévue : une TextBox vide signale
                // clairement à l'utilisateur "je n'applique rien" (et CommandeSelectionnee renverra
                // null en conséquence). Mettre "0" induirait en erreur — "0 V" est une valeur valide
                // et on ne veut pas l'envoyer par accident si l'utilisateur n'a rien saisi.
                Valeur = source.ValeurDefaut ?? string.Empty;
            }
        }

        /// <summary>Binding TwoWay pour le premier RadioButton d'un réglage binaire.</summary>
        public bool EstPremiereOption
        {
            get => Options.Count > 0 && ReferenceEquals(OptionSelectionnee, Options[0]);
            set { if (value && Options.Count > 0) OptionSelectionnee = Options[0]; }
        }

        /// <summary>Binding TwoWay pour le second RadioButton d'un réglage binaire.</summary>
        public bool EstSecondeOption
        {
            get => Options.Count > 1 && ReferenceEquals(OptionSelectionnee, Options[1]);
            set { if (value && Options.Count > 1) OptionSelectionnee = Options[1]; }
        }

        public string LibellePremiereOption => Options.Count > 0 ? Options[0].Libelle : string.Empty;
        public string LibelleSecondeOption  => Options.Count > 1 ? Options[1].Libelle : string.Empty;

        /// <summary>GroupName unique pour les RadioButtons du binding binaire (évite les conflits entre réglages).</summary>
        public string GroupeRadio => "Reglage_" + Nom.Replace(' ', '_');

        /// <summary>
        /// Commande SCPI finale à envoyer :
        /// - Choix : la <c>CommandeScpi</c> de l'option sélectionnée.
        /// - Numerique : le template avec <c>{0}</c> remplacé par <see cref="Valeur"/>.
        /// Retourne null si rien à envoyer.
        /// </summary>
        public string? CommandeSelectionnee
        {
            get
            {
                if (_source.Type == TypeReglage.Numerique)
                {
                    string? template = _source.Options.Count > 0 ? _source.Options[0].CommandeScpi : null;
                    if (string.IsNullOrWhiteSpace(template)) return null;

                    // Si l'utilisateur n'a rien saisi, on envoie 0 explicite plutôt que de
                    // laisser le réglage non envoyé : sur certains compteurs (53131A) un
                    // trigger non configuré active le mode AUTO qui n'est pas voulu en
                    // métrologie. Convention : pas de saisie = valeur 0 forcée.
                    string valeurNum = string.IsNullOrWhiteSpace(Valeur) ? "0" : Valeur.Replace(',', '.');
                    if (!double.TryParse(valeurNum, NumberStyles.Float, CultureInfo.InvariantCulture, out var v)) return null;

                    return string.Format(CultureInfo.InvariantCulture, template, v);
                }

                return OptionSelectionnee?.CommandeScpi;
            }
        }

        /// <summary>
        /// Restaure la sélection de ce réglage à partir des commandes SCPI persistées
        /// (<c>MesureConfig.CommandesScpiReglages</c>), pour que la fenêtre Configuration
        /// retrouve les choix de l'utilisateur quand il la rouvre. Gère les deux types :
        ///  - <b>Choix</b> : re-sélectionne l'option dont la commande figure dans la liste.
        ///  - <b>Numérique</b> (ex. Trigger) : retrouve la commande correspondant au template
        ///    <c>"préfixe{0}suffixe"</c> et en ré-extrait la valeur pour repeupler le champ.
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

            // Numérique : on reconstruit le motif "préfixe{0}suffixe" pour isoler la valeur
            // dans la commande persistée correspondante.
            string? template = _source.Options.Count > 0 ? _source.Options[0].CommandeScpi : null;
            if (string.IsNullOrWhiteSpace(template)) return;

            // ⚠ Un template peut être CHAÎNÉ (ex. 53230A : ":INP1:LEV:AUTO OFF;:INP1:LEV {0}").
            // À la persistance, ConfigurationViewModel.SplitCommandesScpi l'a découpé en
            // entrées distinctes (":INP1:LEV:AUTO OFF", ":INP1:LEV 1"). Pour matcher la bonne
            // entrée, on doit dériver le préfixe/suffixe du SEUL sous-template porteur de {0}
            // (":INP1:LEV {0}"), pas du template chaîné entier — sinon aucun match → valeur perdue.
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
        /// Isole, dans un template SCPI éventuellement chaîné ("a;:b{0};:c"), le sous-template
        /// qui porte le marqueur <c>{0}</c>. Réplique le découpage de
        /// <c>ConfigurationViewModel.SplitCommandesScpi</c> (séparateur ";:", réajout du ":"
        /// initial sur les parties suivantes) pour que le préfixe/suffixe dérivé corresponde
        /// EXACTEMENT à la commande persistée après split. Si le template n'est pas chaîné,
        /// il est renvoyé tel quel.
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
