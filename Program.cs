using System;
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
    [STAThread]
    public static void Main(string[] args)
    {
        VelopackApp.Build().Run();

        var app = new App();
        app.InitializeComponent();
        app.Run();
    }
}
