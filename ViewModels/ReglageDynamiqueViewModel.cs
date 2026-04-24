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
                    if (string.IsNullOrWhiteSpace(template) || string.IsNullOrWhiteSpace(Valeur)) return null;

                    // Normalise la virgule décimale en point (SCPI attend du point).
                    string valeurNum = Valeur.Replace(',', '.');
                    if (!double.TryParse(valeurNum, NumberStyles.Float, CultureInfo.InvariantCulture, out var v)) return null;

                    return string.Format(CultureInfo.InvariantCulture, template, v);
                }

                return OptionSelectionnee?.CommandeScpi;
            }
        }
    }
}
