using System.Threading.Tasks;
using Metrologo.Views;
using Velopack;

namespace Metrologo;

/// <summary>
/// Verifie au lancement s'il existe une version plus recente de Metrologo dans le
/// dossier reseau de deploiement. Si oui, affiche une fenetre de progression,
/// telecharge la mise a jour (barre de progression reelle) puis redemarre l'appli
/// pour l'appliquer.
///
/// Robustesse : si le lecteur reseau n'est pas connecte ou si l'application n'a
/// pas ete installee via Velopack (ex. lancement depuis Visual Studio), la
/// verification est ignoree en silence et l'application demarre normalement.
/// </summary>
public static class UpdateService
{
    /// <summary>
    /// Dossier reseau ou <c>publier-metrologo.bat</c> depose les versions.
    /// DOIT etre identique au chemin de sortie du script de publication.
    /// </summary>
    private const string FeedPath = @"M:\exe_spe\Data_Metrologo\Metrologo";

    /// <summary>
    /// Recherche et applique une eventuelle mise a jour.
    /// Retourne apres avoir termine ; si une mise a jour est appliquee, le
    /// processus est relance par Velopack et ne revient pas de cet appel.
    /// </summary>
    public static async Task CheckAndApplyAsync()
    {
        try
        {
            var mgr = new UpdateManager(FeedPath);

            // Non installe via Velopack (dev / Visual Studio) : rien a faire.
            if (!mgr.IsInstalled)
                return;

            var updateInfo = await mgr.CheckForUpdatesAsync();
            if (updateInfo is null)
                return; // deja a jour

            var win = new UpdateWindow();
            win.Show();
            win.SetStatus("Telechargement de la mise a jour...");
            win.SetProgress(0);

            await mgr.DownloadUpdatesAsync(
                updateInfo,
                percent => win.Dispatcher.Invoke(() => win.SetProgress(percent)));

            win.SetStatus("Installation et redemarrage...");
            win.SetIndeterminate();

            // Laisse le message visible un court instant.
            await Task.Delay(700);

            // Relance l'application sur la nouvelle version (ne revient pas).
            mgr.ApplyUpdatesAndRestart(updateInfo);
        }
        catch
        {
            // Reseau indisponible (M: non connecte), dossier absent, etc.
            // On ignore : l'application doit demarrer normalement.
        }
    }
}
