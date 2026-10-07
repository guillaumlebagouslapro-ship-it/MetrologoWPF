using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Metrologo.Views;
using Velopack;
using Velopack.Locators;

namespace Metrologo;

/// <summary>
/// Verifie au lancement s'il existe une version plus recente de Metrologo dans le
/// dossier reseau de deploiement. Si oui, affiche une fenetre de progression,
/// sauvegarde la version actuelle (pour un retour arriere), telecharge la mise a
/// jour puis redemarre l'appli pour l'appliquer.
///
/// Robustesse : si le lecteur reseau n'est pas connecte ou si l'application n'a
/// pas ete installee via Velopack (ex. lancement depuis Visual Studio), la
/// verification est ignoree en silence et l'application demarre normalement.
///
/// Anti-boucle : si une mise a jour lancee n'a pas pris (meme version au
/// redemarrage), on ne la retente pas automatiquement (voir <see cref="HistoriqueMaj"/>).
///
/// Retour arriere : une seule version en arriere, depuis la copie locale faite
/// avant chaque mise a jour. Ensuite les MAJ sont suspendues jusqu'a la
/// publication d'une version plus recente que celle abandonnee.
/// </summary>
public static class UpdateService
{
    /// <summary>
    /// Dossier reseau ou <c>publier-metrologo.bat</c> depose les versions.
    /// DOIT etre identique au chemin de sortie du script de publication.
    /// </summary>
    private const string FeedPath = @"M:\exe_spe\Data_Metrologo\SUITE ASERTI Guillaume\Metrologo";

    private const string AppId = "Metrologo";

    private static readonly string RacineLocale = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), AppId);

    /// <summary>
    /// Trace de chaque verification (version installee, version trouvee, erreur).
    /// Les echecs restent silencieux pour l'utilisateur mais doivent etre
    /// diagnosticables : sans ce fichier, une MAJ qui ne se fait pas ne laisse
    /// aucune trace.
    /// </summary>
    private static readonly string LogPath = Path.Combine(RacineLocale, "Logs", "maj.log");

    /// <summary>Journal de Velopack lui-meme (cause precise d'une installation ratee).</summary>
    public static readonly string JournalVelopack = Path.Combine(
        Path.GetDirectoryName(RacineLocale)!, "velopack", $"velopack_{AppId}.log");

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
            if (!mgr.IsInstalled || mgr.CurrentVersion is null)
            {
                Log("Non installe via Velopack (dev / copie manuelle) : verification ignoree.");
                return;
            }
            string actuelle = mgr.CurrentVersion.ToString();
            Log($"Version installee : {actuelle}");

            // Historique : note la MAJ / le retour arriere qui vient d'avoir lieu (ou son echec).
            var etat = HistoriqueMaj.Synchroniser(actuelle);
            SupprimerAncienMarqueur();

            // Raccourcis bureau / menu Démarrer : icône de CETTE version (nouveau logo
            // visible dès la mise à jour, sans réinstallation). En arrière-plan : le
            // démarrage ne l'attend pas.
            _ = Task.Run(() => Log(IconeRaccourcis.Appliquer(actuelle)));

            if (!Directory.Exists(FeedPath))
            {
                Log("Dossier reseau introuvable (lecteur M: non connecte ?) : " + FeedPath);
                return;
            }

            var updateInfo = await mgr.CheckForUpdatesAsync();
            if (updateInfo is null)
            {
                Log("Deja a jour.");
                return;
            }

            string cible = updateInfo.TargetFullRelease.Version.ToString();
            Log($"Version disponible : {cible}");

            // Retour arriere fait par l'utilisateur : on n'impose pas a nouveau la version
            // abandonnee. Une version plus recente leve le blocage.
            if (etat.VersionBloquee != null)
            {
                if (HistoriqueMaj.Comparer(cible, etat.VersionBloquee) <= 0)
                {
                    Log($"MAJ suspendues apres retour arriere (version {etat.VersionBloquee} abandonnee) : {cible} ignoree.");
                    return;
                }
                Log($"Version {cible} plus recente que la version abandonnee {etat.VersionBloquee} : MAJ reprises.");
                etat.VersionBloquee = null;
            }

            // ANTI-BOUCLE : cette version a deja ete tentee et n'a pas pris sur ce poste.
            if (etat.TentativeEchouee == cible)
            {
                Log($"MAJ {cible} deja tentee sans succes (anti-boucle) : ignoree. Cause : {JournalVelopack}. " +
                    "Bouton 'Reactiver les mises a jour' (Accueil > Mises a jour) pour reessayer.");
                HistoriqueMaj.Sauver(etat);
                return;
            }

            var win = new UpdateWindow();
            win.Show();
            win.SetStatus("Sauvegarde de la version actuelle...");
            win.SetIndeterminate();

            await SauvegarderVersionActuelleAsync(mgr, etat, actuelle);

            win.SetStatus("Telechargement de la mise a jour...");
            win.SetProgress(0);

            await mgr.DownloadUpdatesAsync(
                updateInfo,
                percent => win.Dispatcher.Invoke(() => win.SetProgress(percent)));

            win.SetStatus("Installation et redemarrage...");
            win.SetIndeterminate();

            // Note AVANT de redemarrer ce qu'on s'apprete a installer (historique + anti-boucle).
            etat.EnAttente = new OperationMajEnAttente
            {
                Version = cible,
                Type = HistoriqueMaj.TypeMiseAJour,
                Description = Nettoyer(updateInfo.TargetFullRelease.NotesMarkdown),
            };
            HistoriqueMaj.Sauver(etat);

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

    /// <summary>Version installee (null si lancee hors installation : Visual Studio...).</summary>
    public static string? VersionInstallee()
    {
        try
        {
            var mgr = new UpdateManager(FeedPath);
            return mgr.IsInstalled ? mgr.CurrentVersion?.ToString() : null;
        }
        catch { return null; }
    }

    /// <summary>
    /// Reinstalle la version precedente sauvegardee puis redemarre (ne revient pas en
    /// cas de succes). <paramref name="nettoyerAvantArret"/> fait le menage de fermeture
    /// (journal, Excel) avant que Velopack ne coupe le processus.
    /// Leve <see cref="InvalidOperationException"/> avec un message affichable sinon.
    /// </summary>
    public static void RetournerVersionPrecedente(Action nettoyerAvantArret)
    {
        var mgr = new UpdateManager(FeedPath);
        if (!mgr.IsInstalled || mgr.CurrentVersion is null)
            throw new InvalidOperationException("Metrologo n'est pas lance depuis son installation : retour arriere impossible.");

        string actuelle = mgr.CurrentVersion.ToString();
        var etat = HistoriqueMaj.Charger();
        string? chemin = HistoriqueMaj.CheminPrecedente(etat);
        if (etat.Precedente is null || chemin is null)
            throw new InvalidOperationException("Aucune version precedente n'est sauvegardee sur ce poste.");
        if (HistoriqueMaj.Comparer(etat.Precedente.Version, actuelle) >= 0)
            throw new InvalidOperationException("La version sauvegardee n'est pas plus ancienne que la version actuelle.");

        // Velopack applique un paquet depose dans son dossier packages.
        string dossierPackages = DossierPackages();
        Directory.CreateDirectory(dossierPackages);
        string dest = Path.Combine(dossierPackages, Path.GetFileName(chemin));
        File.Copy(chemin, dest, overwrite: true);
        var asset = VelopackAsset.FromNupkgNoChecksum(dest);

        etat.EnAttente = new OperationMajEnAttente
        {
            Version = etat.Precedente.Version,
            Type = HistoriqueMaj.TypeRetourArriere,
            Description = $"Retour à la version {HistoriqueMaj.Affichage(etat.Precedente.Version)} " +
                          $"(version {HistoriqueMaj.Affichage(actuelle)} abandonnée).",
        };
        etat.VersionBloquee = actuelle;
        HistoriqueMaj.Sauver(etat);

        Log($"Retour arriere demande : {actuelle} -> {etat.Precedente.Version}.");
        nettoyerAvantArret();
        mgr.ApplyUpdatesAndRestart(asset);
    }

    /// <summary>Leve la suspension (apres retour arriere) et l'anti-boucle : la MAJ sera
    /// retentee au prochain demarrage.</summary>
    public static void ReactiverMisesAJour()
    {
        var etat = HistoriqueMaj.Charger();
        etat.VersionBloquee = null;
        etat.TentativeEchouee = null;
        HistoriqueMaj.Sauver(etat);
        Log("Mises a jour reactivees par l'utilisateur.");
    }

    /// <summary>
    /// Copie le paquet complet de la version installee dans MAJ\precedente\ (remplace
    /// l'ancienne copie : un seul retour arriere possible). Source : dossier packages
    /// local si present, sinon le dossier reseau. Sans copie, pas de retour arriere,
    /// mais la mise a jour se fait quand meme.
    /// </summary>
    private static async Task SauvegarderVersionActuelleAsync(UpdateManager mgr, EtatMaj etat, string actuelle)
    {
        try
        {
            string fichier = $"{AppId}-{actuelle}-full.nupkg";
            string? notes = null;
            var asset = LireFlux()?.Assets?.FirstOrDefault(a =>
                a.Type == VelopackAssetType.Full && a.Version?.ToString() == actuelle);
            if (asset != null)
            {
                fichier = asset.FileName;
                notes = Nettoyer(asset.NotesMarkdown);
            }

            string? source = new[]
                {
                    Path.Combine(DossierPackages(), fichier),
                    Path.Combine(FeedPath, fichier),
                }
                .FirstOrDefault(p => p != null && File.Exists(p));

            if (source is null)
            {
                Log($"Paquet de la version {actuelle} introuvable : pas de retour arriere possible apres cette MAJ.");
                HistoriqueMaj.SupprimerPrecedente(etat);
                return;
            }

            string tmp = HistoriqueMaj.DossierPrecedente + ".tmp";
            await Task.Run(() =>
            {
                if (Directory.Exists(tmp)) Directory.Delete(tmp, recursive: true);
                Directory.CreateDirectory(tmp);
                File.Copy(source, Path.Combine(tmp, fichier));
                if (Directory.Exists(HistoriqueMaj.DossierPrecedente))
                    Directory.Delete(HistoriqueMaj.DossierPrecedente, recursive: true);
                Directory.Move(tmp, HistoriqueMaj.DossierPrecedente);
            });

            etat.Precedente = new VersionSauvegardee { Version = actuelle, Fichier = fichier, Description = notes };
            HistoriqueMaj.Sauver(etat);
            Log($"Version {actuelle} sauvegardee pour retour arriere (depuis {source}).");
        }
        catch (Exception ex)
        {
            Log($"Sauvegarde de la version {actuelle} impossible : {ex.Message}");
        }
    }

    /// <summary>Dossier "packages" de Velopack (la ou il cherche le paquet a appliquer).</summary>
    private static string DossierPackages()
    {
        try
        {
            if (VelopackLocator.IsCurrentSet && VelopackLocator.Current.PackagesDir is { } dossier)
                return dossier;
        }
        catch { /* repli ci-dessous */ }
        return Path.Combine(RacineLocale, "packages");
    }

    private static VelopackAssetFeed? LireFlux()
    {
        try
        {
            string json = Path.Combine(FeedPath, "releases.win.json");
            return File.Exists(json) ? VelopackAssetFeed.FromJson(File.ReadAllText(json)) : null;
        }
        catch { return null; }
    }

    /// <summary>Notes de version saisies a la publication : texte simple sur une ligne.</summary>
    private static string? Nettoyer(string? notes)
    {
        if (string.IsNullOrWhiteSpace(notes)) return null;
        return string.Join(" ", notes.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                                     .Select(l => l.Trim().TrimStart('#', '-', '*', ' ')))
                     .Trim();
    }

    /// <summary>Ancien marqueur anti-boucle (remplace par etat-maj.json).</summary>
    private static void SupprimerAncienMarqueur()
    {
        try
        {
            string ancien = Path.Combine(RacineLocale, ".update-attempt");
            if (File.Exists(ancien)) File.Delete(ancien);
        }
        catch { /* sans importance */ }
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
}
