using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using Velopack;

namespace Metrologo;

/// <summary>
/// Point d'entree de l'application.
///
/// <see cref="VelopackApp"/> DOIT etre execute en tout premier : lors d'une
/// installation, d'une mise a jour ou d'une desinstallation, Velopack relance
/// brievement l'exe avec des arguments speciaux ; cet appel intercepte ces cas,
/// fait le necessaire, puis quitte. En fonctionnement normal il ne fait rien et
/// l'application demarre comme d'habitude (App.OnStartup).
/// </summary>
public static class Program
{
    /// <summary>
    /// Verrou d'instance unique. "Global\" : un seul Metrologo sur le poste, toutes
    /// sessions Windows confondues (le bus GPIB et Excel ne se partagent pas).
    /// Libere automatiquement par Windows a la fin du processus, ce qui laisse la
    /// relance apres mise a jour Velopack fonctionner normalement.
    /// </summary>
    private const string NomMutex = @"Global\Metrologo-InstanceUnique-7C1E4A52";

    // Champ statique : garde le mutex vivant (pas collecte par le GC) jusqu'a la sortie.
    private static Mutex? _mutexInstance;

    [STAThread]
    public static void Main(string[] args)
    {
        VelopackApp.Build().Run();

        bool premiereInstance;
        try
        {
            _mutexInstance = new Mutex(true, NomMutex, out premiereInstance);
        }
        catch (UnauthorizedAccessException)
        {
            // Verrou cree par un autre utilisateur Windows : Metrologo tourne deja.
            premiereInstance = false;
        }

        if (!premiereInstance || _mutexInstance is null)
        {
            if (!ActiverInstanceExistante())
            {
                MessageBox.Show(
                    "Metrologo est déjà ouvert (ou en cours de démarrage) sur ce poste.",
                    "Metrologo", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            return;
        }

        try
        {
            var app = new App();
            app.InitializeComponent();
            app.Run();
        }
        finally
        {
            _mutexInstance.ReleaseMutex();
            _mutexInstance.Dispose();
            _mutexInstance = null;
        }
    }

    /// <summary>
    /// Ramene au premier plan la fenetre du Metrologo deja lance (restaure si
    /// reduite). Retourne false si aucune fenetre n'est trouvee (instance encore
    /// en demarrage, ou ouverte dans une autre session Windows).
    /// </summary>
    private static bool ActiverInstanceExistante()
    {
        using var courant = Process.GetCurrentProcess();
        foreach (var p in Process.GetProcessesByName(courant.ProcessName))
        {
            using (p)
            {
                if (p.Id == courant.Id || p.SessionId != courant.SessionId)
                    continue;

                IntPtr hwnd = p.MainWindowHandle;
                if (hwnd == IntPtr.Zero)
                    continue;

                if (IsIconic(hwnd))
                    ShowWindow(hwnd, SW_RESTORE);
                SetForegroundWindow(hwnd);
                return true;
            }
        }
        return false;
    }

    private const int SW_RESTORE = 9;

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool IsIconic(IntPtr hWnd);
}
