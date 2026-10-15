using Metrologo.Models;

namespace Metrologo.Services
{
    /// <summary>
    /// Recharge à chaud la configuration partagée (chemins, étiquettes, catalogue d'appareils)
    /// sans redémarrer ni interrompre une mesure.
    /// Les caches en mémoire sont invalidés et relus à la prochaine lecture ; les services sans
    /// cache long-vécu (modules d'incertitude, rubidiums, utilisateurs) relisent déjà leurs
    /// fichiers à l'ouverture de leur écran.
    /// Appelé via « Actualiser maintenant » (cf. <c>MainViewModel.AcquitterChangementsAdmin</c>).
    /// </summary>
    public static class RafraichirConfigurationService
    {
        public static async Task RafraichirAsync()
        {
            // 1. Chemins en premier : tous les accès réseau ci-dessous en dépendent.
            CheminsMetrologo.ChargerConfigChemins();

            // 2. Étiquettes Module/Fonction des feuilles Excel.
            MesureConfigService.Recharger();

            // 3. Caches partagés, relus à la prochaine ouverture d'écran.
            Preferences.InvaliderCacheUtilisateurs();
            Preferences.InvaliderCacheCatalogueRubidiums();

            // 4. Rubidium actif : lève RubidiumActifChange. Synchrone (avant tout await) pour que
            //    la notification parte sur le thread UI.
            EtatApplication.RechargerRubidiumActif();

            // 5. Catalogue d'appareils : relit le JSON réseau et lève CatalogueChange.
            await Catalogue.CatalogueAppareilsService.Instance.ChargerAsync();
        }
    }
}
