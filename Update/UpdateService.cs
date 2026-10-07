using System;
using System.IO;
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
    private const string FeedPath = @"M:\exe_spe\Data_Metrologo\SUITE ASERTI Guillaume\Metrologo";

    /// <summary>
    /// Trace de chaque verification (version installee, version trouvee, erreur).
    /// Les echecs restent silencieux pour l'utilisateur mais doivent etre
    /// diagnosticables : sans ce fichier, une MAJ qui ne se fait pas ne laisse
    /// aucune trace.
    /// </summary>
    private static readonly string LogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Metrologo", "Logs", "maj.log");

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
            {
                Log("Non installe via Velopack (dev / copie manuelle) : verification ignoree.");
                return;
            }

            var updateInfo = await mgr.CheckForUpdatesAsync();
            if (updateInfo is null)
            {
                Log($"Version {mgr.CurrentVersion} : deja a jour (canal {FeedPath}).");
                return;
            }

            Log($"Version {mgr.CurrentVersion} -> {updateInfo.TargetFullRelease.Version} : mise a jour.");

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
        catch (Exception ex)
        {
            // Reseau indisponible (M: non connecte), dossier absent, etc.
            // On ignore : l'application doit demarrer normalement.
            Log($"ECHEC verification (canal {FeedPath}) : {ex.GetType().Name} : {ex.Message}");
        }
    }

    private static void Log(string message)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
            File.AppendAllText(LogPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}  {message}{Environment.NewLine}");
        }
        catch
        {
            // Journal facultatif : ne doit jamais bloquer le demarrage.
        }
    }
}
