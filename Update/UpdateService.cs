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
///
/// Anti-boucle (meme principe qu'Asertools) : on memorise la derniere version qu'on
/// a TENTE d'appliquer. Si au redemarrage la meme mise a jour est re-proposee
/// (l'installation n'a pas pris : fichier verrouille, antivirus, droits...), on
/// n'insiste pas et on ouvre l'application au lieu de reboucler indefiniment.
/// </summary>
public static class UpdateService
{
    /// <summary>
    /// Dossier reseau ou <c>publier-metrologo.bat</c> depose les versions.
    /// DOIT etre identique au chemin de sortie du script de publication.
    /// </summary>
    private const string FeedPath = @"M:\exe_spe\Data_Metrologo\SUITE ASERTI Guillaume\Metrologo";

    private static readonly string RacineLocale = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Metrologo");

    /// <summary>
    /// Trace de chaque verification (version installee, version trouvee, erreur).
    /// Les echecs restent silencieux pour l'utilisateur mais doivent etre
    /// diagnosticables : sans ce fichier, une MAJ qui ne se fait pas ne laisse
    /// aucune trace.
    /// </summary>
    private static readonly string LogPath = Path.Combine(RacineLocale, "Logs", "maj.log");

    /// <summary>
    /// Marqueur : numero de la derniere version dont on a lance l'installation.
    /// Supprimer ce fichier pour forcer une nouvelle tentative.
    /// </summary>
    private static readonly string MarqueurTentative = Path.Combine(RacineLocale, ".update-attempt");

    /// <summary>
    /// Recherche et applique une eventuelle mise a jour.
    /// Retourne apres avoir termine ; si une mise a jour est appliquee, le
    /// processus est relance par Velopack et ne revient pas de cet appel.
    /// </summary>
    public static async Task CheckAndApplyAsync()
    {
        try
        {
            Log("---- Demarrage, exe : " + Environment.ProcessPath);
            var mgr = new UpdateManager(FeedPath);

            // Non installe via Velopack (dev / Visual Studio) : rien a faire.
            if (!mgr.IsInstalled)
            {
                Log("Non installe via Velopack (dev / copie manuelle) : verification ignoree.");
                return;
            }
            Log($"Version installee : {mgr.CurrentVersion}");

            if (!Directory.Exists(FeedPath))
            {
                Log("Dossier reseau introuvable (lecteur M: non connecte ?) : " + FeedPath);
                return;
            }

            var updateInfo = await mgr.CheckForUpdatesAsync();
            if (updateInfo is null)
            {
                Log("Deja a jour.");
                // L'install a fini par prendre (ou plus rien a faire) : on oublie la tentative.
                EffacerMarqueur();
                return;
            }

            string cible = updateInfo.TargetFullRelease.Version.ToString();
            Log($"Version disponible : {cible}");

            // ANTI-BOUCLE : cette version a deja ete tentee au lancement precedent et on
            // nous la re-propose -> l'installation ne prend pas sur ce poste. On ouvre
            // l'appli au lieu d'enfermer l'utilisateur dans une boucle sans fin.
            if (LireMarqueur() == cible)
            {
                Log($"MAJ {cible} deja tentee sans succes (anti-boucle) : ignoree. " +
                    $"Voir {Path.Combine(RacineLocale, "velopack.log")} ; supprimer {MarqueurTentative} pour reessayer.");
                return;
            }

            // Note AVANT de redemarrer la version qu'on s'apprete a installer.
            EcrireMarqueur(cible);

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
            Log($"Telechargement termine : installation de {cible} et redemarrage.");
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
            if (File.Exists(LogPath) && new FileInfo(LogPath).Length > 200_000)
                File.Delete(LogPath);
            File.AppendAllText(LogPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}  {message}{Environment.NewLine}");
        }
        catch
        {
            // Journal facultatif : ne doit jamais bloquer le demarrage.
        }
    }

    private static string? LireMarqueur()
    {
        try { return File.Exists(MarqueurTentative) ? File.ReadAllText(MarqueurTentative).Trim() : null; }
        catch { return null; }
    }

    private static void EcrireMarqueur(string version)
    {
        try
        {
            Directory.CreateDirectory(RacineLocale);
            File.WriteAllText(MarqueurTentative, version);
        }
        catch { /* sans marqueur : comportement d'origine */ }
    }

    private static void EffacerMarqueur()
    {
        try { if (File.Exists(MarqueurTentative)) File.Delete(MarqueurTentative); }
        catch { /* sans importance */ }
    }
}
