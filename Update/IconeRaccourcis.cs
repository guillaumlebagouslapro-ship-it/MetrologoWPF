using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace Metrologo;

/// <summary>
/// Garde l'icône des raccourcis (bureau, menu Démarrer, barre des tâches épinglée) et de
/// l'entrée « Applications installées » alignée sur la version installée.
/// <para/>
/// Pourquoi : les raccourcis Velopack pointent vers un petit lanceur créé à l'installation,
/// qui garde l'icône de cette époque, et Windows met les icônes en cache. Ici, chaque
/// version écrit SA propre icône (embarquée dans l'exe) sous un nom qui contient le numéro
/// de version, et les raccourcis pointent dessus : un nouveau logo publié apparaît donc au
/// premier démarrage qui suit la mise à jour, sans réinstaller ni vider de cache.
/// </summary>
public static class IconeRaccourcis
{
    private const string AppId = "Metrologo";
    private const string RessourceIcone = "pack://application:,,,/Resources/logo.ico";

    private static readonly string Racine = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), AppId);

    /// <summary>
    /// A appeler au démarrage (application installée). Ne lève jamais d'exception.
    /// Démarrage ordinaire : l'icône de la version existe déjà -> rien à faire (simple test
    /// d'existence). Le vrai travail n'a lieu qu'au 1er démarrage d'une nouvelle version.
    /// </summary>
    public static string Appliquer(string version)
    {
        try
        {
            if (File.Exists(Path.Combine(Racine, $"icone-{version}.ico")))
                return "icone des raccourcis deja a jour";

            return Appliquer(version, Racine, DossiersRaccourcis(), () =>
                System.Windows.Application.GetResourceStream(new Uri(RessourceIcone))?.Stream);
        }
        catch (Exception ex)
        {
            return "icone non mise a jour : " + ex.Message;
        }
    }

    /// <summary>Coeur testable : <paramref name="racine"/> = dossier d'installation.</summary>
    internal static string Appliquer(string version, string racine, IEnumerable<string> dossiers,
                                     Func<Stream?> ouvrirIcone, bool registre = true)
    {
        // 1) L'icône de CETTE version, sous un nom propre à la version (cache Windows contourné).
        string icone = Path.Combine(racine, $"icone-{version}.ico");
        if (!File.Exists(icone))
        {
            using var source = ouvrirIcone() ?? throw new InvalidOperationException("ressource icone introuvable");
            Directory.CreateDirectory(racine);
            string tmp = icone + ".tmp";
            using (var f = File.Create(tmp)) source.CopyTo(f);
            File.Move(tmp, icone, overwrite: true);
        }
        foreach (string ancienne in Directory.GetFiles(racine, "icone-*.ico"))
        {
            if (!string.Equals(ancienne, icone, StringComparison.OrdinalIgnoreCase))
                try { File.Delete(ancienne); } catch { /* encore utilisee : au prochain coup */ }
        }

        // 2) Raccourcis qui lancent l'application installée.
        int modifies = 0;
        string voulu = icone + ",0";
        var typeShell = Type.GetTypeFromProgID("WScript.Shell")
                        ?? throw new InvalidOperationException("WScript.Shell indisponible");
        dynamic shell = Activator.CreateInstance(typeShell)!;
        try
        {
            foreach (string dossier in dossiers)
            {
                if (!Directory.Exists(dossier)) continue;
                foreach (string lnk in Directory.EnumerateFiles(dossier, "*.lnk"))
                {
                    dynamic raccourci = shell.CreateShortcut(lnk);
                    try
                    {
                        string cible = raccourci.TargetPath ?? "";
                        if (!cible.StartsWith(racine + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                            continue;
                        if (string.Equals((string)raccourci.IconLocation, voulu, StringComparison.OrdinalIgnoreCase))
                            continue;
                        raccourci.IconLocation = voulu;
                        raccourci.Save();
                        modifies++;
                    }
                    finally { Marshal.FinalReleaseComObject(raccourci); }
                }
            }
        }
        finally { Marshal.FinalReleaseComObject(shell); }

        // 3) Entrée « Applications installées » (écrite par Velopack).
        if (registre)
        {
            using var cle = Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Uninstall\" + AppId, writable: true);
            if (cle != null && !string.Equals(cle.GetValue("DisplayIcon") as string, icone, StringComparison.OrdinalIgnoreCase))
            {
                cle.SetValue("DisplayIcon", icone);
                modifies++;
            }
        }

        // 4) Demande à l'Explorateur de redessiner les icônes.
        if (modifies > 0)
            SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, IntPtr.Zero, IntPtr.Zero);

        return modifies == 0 ? "icone des raccourcis deja a jour" : $"icone {version} appliquee ({modifies} element(s))";
    }

    /// <summary>Emplacements où Velopack et l'utilisateur mettent les raccourcis.</summary>
    private static IEnumerable<string> DossiersRaccourcis()
    {
        string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        yield return Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        yield return Environment.GetFolderPath(Environment.SpecialFolder.Programs);
        yield return Path.Combine(appData, @"Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar");
        yield return Path.Combine(appData, @"Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu");
    }

    private const int SHCNE_ASSOCCHANGED = 0x08000000;
    private const uint SHCNF_IDLIST = 0x0000;

    [DllImport("shell32.dll")]
    private static extern void SHChangeNotify(int wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
}
