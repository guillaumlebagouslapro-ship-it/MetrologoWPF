using System.Threading.Tasks;
using Metrologo.Models;

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
            // 1. Chemins EN PREMIER : tous les fichiers réseau ci-dessous en dérivent
            //    (rubidium actif, utilisateurs, catalogues, incertitudes…). Relit le cache
            //    local puis le fichier maître serveur si renseigné. Idempotent.
            CheminsMetrologo.ChargerConfigChemins();

            // 2. Étiquettes Module/Fonction écrites dans les feuilles Excel : vide le
            //    cache, relu au prochain ObtenirPourType().
            MesureConfigService.Recharger();

            // 3. Caches mémoire des données partagées : invalidés → relus à la prochaine
            //    ouverture de leur écran (sélection utilisateur, choix rubidium).
            Preferences.InvaliderCacheUtilisateurs();
            Preferences.InvaliderCacheCatalogueRubidiums();

            // 4. Rubidium actif : relit le fichier partagé ET lève RubidiumActifChange.
            //    L'AccueilViewModel y est abonné → le bandeau se met à jour immédiatement.
            //    Synchrone et exécuté sur le thread appelant (UI) avant tout await, donc
            //    la notification de binding part bien du thread UI.
            EtatApplication.RechargerRubidiumActif();

            // 5. Catalogue d'appareils : singleton en mémoire. ChargerAsync() relit le
            //    JSON réseau et lève CatalogueChange → les écrans bindés se rafraîchissent.
            await Catalogue.CatalogueAppareilsService.Instance.ChargerAsync();
        }
    }
}
