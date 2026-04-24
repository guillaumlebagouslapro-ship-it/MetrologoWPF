using CommunityToolkit.Mvvm.ComponentModel;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// Entrée "cochable" pour la sélection des temps de porte dans la fenêtre d'enregistrement
    /// d'un appareil. Chaque instance représente un libellé standard (ex: "1 s", "100 ms") que
    /// l'utilisateur active ou désactive selon ce que son fréquencemètre supporte.
    ///
    /// Remplace la saisie libre texte qui pouvait laisser passer des coquilles typo ("100ms"
    /// sans espace, "1sec", etc.) non reconnues par le parseur.
    /// </summary>
    public partial class GateCochable : ObservableObject
    {
        public string Libelle { get; }

        [ObservableProperty] private bool _estCoche;

        public GateCochable(string libelle)
        {
            Libelle = libelle;
        }
    }
}
