using System.Threading.Tasks;

namespace Metrologo.Services
{
    /// <summary>
    /// Recharge à chaud la configuration partagée modifiée par un admin depuis un autre
    /// poste (chemins, étiquettes module/fonction, catalogue d'appareils…), sans
    /// redémarrer l'application ni interrompre une mesure en cours.
    /// <para/>
    /// Stratégie « invalidation de caches » : on vide / relit uniquement les caches
    /// persistants en mémoire. Les nouveaux réglages sont donc pris en compte à la
    /// prochaine lecture (nouvelle mesure, ouverture d'un écran). Les services
    /// sans cache long-vécu (modules d'incertitude, catalogue rubidiums, utilisateurs)
    /// relisent déjà leurs fichiers à l'ouverture de leur écran : rien à invalider.
    /// <para/>
    /// Appelé quand l'utilisateur clique « Actualiser maintenant » sur le pop-up
    /// signalant des changements administrateur (cf. <c>MainViewModel.AcquitterChangementsAdmin</c>).
    /// </summary>
    public static class RafraichirConfigurationService
    {
        public static async Task RafraichirAsync()
        {
            // 1. Chemins EN PREMIER : les autres services en dérivent (dossier
            //    Incertitudes, fichier catalogue, étiquettes…). Relit le cache local
            //    puis le fichier maître serveur si renseigné. Idempotent.
            CheminsMetrologo.ChargerConfigChemins();

            // 2. Étiquettes Module/Fonction écrites dans les feuilles Excel : vide le
            //    cache, relu au prochain ObtenirPourType().
            MesureConfigService.Recharger();

            // 3. Catalogue d'appareils : singleton en mémoire. ChargerAsync() relit le
            //    JSON réseau et lève CatalogueChange → les écrans bindés se rafraîchissent.
            await Catalogue.CatalogueAppareilsService.Instance.ChargerAsync();
        }
    }
}
