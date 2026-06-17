using CommunityToolkit.Mvvm.ComponentModel;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// Case à cocher pour un temps de porte standard (ex: "1 s", "100 ms"). L'utilisateur
    /// coche ceux que son fréquencemètre supporte. Remplace la saisie libre pour éviter
    /// les coquilles ("100ms", "1sec"...) non reconnues par le parseur.
    /// </summary>
    public partial class GateCochable : ObservableObject
    {
        public string Libelle { get; }

        [ObservableProperty] private bool _estCoche;

        /// <summary>
        /// Index canonique du libellé dans l'échelle 0..15 (10 ms → 1000 s). -1 si non
        /// renseigné — utilisé par <c>SelectionGateViewModel</c> pour mapper directement
        /// les cases cochées vers <c>Mesure.GateIndices</c> sans reparser le libellé.
        /// </summary>
        public int SlotCanonique { get; set; } = -1;

        public GateCochable(string libelle)
        {
            Libelle = libelle;
        }
    }
}
