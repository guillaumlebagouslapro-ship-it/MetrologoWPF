using System.Threading.Tasks;
using Metrologo.Models;

namespace Metrologo.Services
{
    /// <summary>
    /// Recharge à chaud la configuration partagée (chemins, étiquettes, catalogue d'appareils)
    /// sans redémarrer ni interrompre une mesure.
    /// <para/>
    /// Stratégie : invalidation des caches persistants en mémoire ; les nouveaux réglages
    /// sont pris en compte à la prochaine lecture. Les services sans cache long-vécu
    /// (modules d'incertitude, rubidiums, utilisateurs) relisent déjà leurs fichiers à
    /// l'ouverture de leur écran.
    /// <para/>
    /// Appelé via « Actualiser maintenant » (cf. <c>MainViewModel.AcquitterChangementsAdmin</c>).
    /// </summary>
    public static class RafraichirConfigurationService
    {
        public static async Task RafraichirAsync()
        {
            // 1. Chemins en premier : tous les accès réseau ci-dessous en dépendent. Idempotent.
            CheminsMetrologo.ChargerConfigChemins();

            // 2. Étiquettes Module/Fonction (feuilles Excel) : cache vidé, relu au prochain ObtenirPourType().
            MesureConfigService.Recharger();

            // 3. Caches partagés invalidés → relus à la prochaine ouverture d'écran.
            Preferences.InvaliderCacheUtilisateurs();
            Preferences.InvaliderCacheCatalogueRubidiums();

            // 4. Rubidium actif : relit le fichier partagé et lève RubidiumActifChange.
            //    Synchrone (avant tout await) → notification sur le thread UI, bandeau mis à jour.
            EtatApplication.RechargerRubidiumActif();

            // 5. Catalogue d'appareils : ChargerAsync() relit le JSON réseau et lève CatalogueChange.
            await Catalogue.CatalogueAppareilsService.Instance.ChargerAsync();
        }
    }
}
