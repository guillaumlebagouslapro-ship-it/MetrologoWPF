using System;
using System.Linq;
using System.Windows;

namespace Metrologo.Views;

/// <summary>
/// Accueil > Mises à jour : version installée, historique des mises à jour du poste
/// et retour à la version précédente (une seule version en arrière).
/// </summary>
public partial class MisesAJourWindow : Window
{
    private sealed record LigneHistorique(string Date, string Operation, string Versions, string Description);

    public MisesAJourWindow()
    {
        InitializeComponent();
        Charger();
    }

    private void Charger()
    {
        string? installee = UpdateService.VersionInstallee();
        var etat = HistoriqueMaj.Charger();

        VersionTexte.Text = installee is null
            ? "Version de développement (lancée hors installation)"
            : "Version installée : " + HistoriqueMaj.Affichage(installee);

        // --- Bandeau d'alerte ---
        string? alerte = null;
        bool reactivable = false;
        if (installee is null)
            alerte = "Metrologo n'est pas lancé depuis son installation : les mises à jour et le retour arrière ne s'appliquent pas ici.";
        else if (etat.VersionBloquee != null)
        {
            alerte = $"Mises à jour suspendues : vous êtes revenu en arrière. La version {HistoriqueMaj.Affichage(etat.VersionBloquee)} " +
                     "ne sera plus installée automatiquement ; une version plus récente le sera.";
            reactivable = true;
        }
        else if (etat.TentativeEchouee != null)
        {
            alerte = $"La mise à jour {HistoriqueMaj.Affichage(etat.TentativeEchouee)} n'a pas pu s'installer : elle n'est plus retentée " +
                     "automatiquement (pour éviter une boucle de redémarrages). Cause détaillée : " + UpdateService.JournalVelopack;
            reactivable = true;
        }
        BandeauAlerte.Visibility = alerte is null ? Visibility.Collapsed : Visibility.Visible;
        AlerteTexte.Text = alerte ?? "";
        BoutonReactiver.Visibility = reactivable ? Visibility.Visible : Visibility.Collapsed;

        // --- Version précédente ---
        bool retourPossible = installee != null
            && etat.Precedente != null
            && HistoriqueMaj.CheminPrecedente(etat) != null
            && HistoriqueMaj.Comparer(etat.Precedente.Version, installee) < 0;

        if (retourPossible)
        {
            string v = HistoriqueMaj.Affichage(etat.Precedente!.Version);
            PrecedenteTexte.Text = $"Version {v} sauvegardée sur ce poste"
                + (string.IsNullOrEmpty(etat.Precedente.Description) ? "." : $" : {etat.Precedente.Description}");
            BoutonRetour.Content = $"Revenir à la version {v}";
        }
        else
        {
            PrecedenteTexte.Text = "Aucune version précédente disponible. La version en place est sauvegardée "
                + "automatiquement à chaque mise à jour (un seul retour en arrière possible).";
            BoutonRetour.Content = "Revenir à la version précédente";
        }
        BoutonRetour.IsEnabled = retourPossible;

        // --- Historique (le plus récent en haut) ---
        var lignes = etat.Historique
            .OrderByDescending(h => h.Date)
            .Select(h => new LigneHistorique(
                h.Date.ToString("dd/MM/yyyy HH:mm"),
                Libelle(h.Type),
                h.De is null || h.Type == HistoriqueMaj.TypeInstallation
                    ? HistoriqueMaj.Affichage(h.Vers)
                    : $"{HistoriqueMaj.Affichage(h.De)} → {HistoriqueMaj.Affichage(h.Vers)}",
                h.Description ?? ""))
            .ToList();
        ListeHistorique.ItemsSource = lignes;
        HistoriqueVide.Visibility = lignes.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        ListeHistorique.Visibility = lignes.Count == 0 ? Visibility.Collapsed : Visibility.Visible;
    }

    private static string Libelle(string type) => type switch
    {
        HistoriqueMaj.TypeMiseAJour => "Mise à jour",
        HistoriqueMaj.TypeRetourArriere => "Retour arrière",
        HistoriqueMaj.TypeEchec => "Échec",
        HistoriqueMaj.TypeInstallation => "Installation",
        _ => type,
    };

    private void Retour_Click(object sender, RoutedEventArgs e)
    {
        var etat = HistoriqueMaj.Charger();
        string? installee = UpdateService.VersionInstallee();
        if (etat.Precedente is null || installee is null) return;

        string ancienne = HistoriqueMaj.Affichage(etat.Precedente.Version);
        string actuelle = HistoriqueMaj.Affichage(installee);
        var rep = MessageBox.Show(this,
            $"Revenir à la version {ancienne} ?\n\n" +
            $"Metrologo va se fermer puis redémarrer en version {ancienne}. Une mesure en cours sera interrompue.\n\n" +
            $"La version {actuelle} ne sera plus installée automatiquement : les mises à jour reprendront " +
            "à la prochaine version publiée (ou avec « Réactiver les mises à jour »).\n\n" +
            "Un seul retour en arrière est possible.",
            "Revenir à la version précédente", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
        if (rep != MessageBoxResult.Yes) return;

        try
        {
            UpdateService.RetournerVersionPrecedente(App.NettoyerAvantArret);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Retour arrière impossible :\n\n" + ex.Message,
                "Revenir à la version précédente", MessageBoxButton.OK, MessageBoxImage.Warning);
            Charger();
        }
    }

    private void Reactiver_Click(object sender, RoutedEventArgs e)
    {
        UpdateService.ReactiverMisesAJour();
        MessageBox.Show(this, "Les mises à jour seront recherchées au prochain démarrage de Metrologo.",
            "Mises à jour", MessageBoxButton.OK, MessageBoxImage.Information);
        Charger();
    }

    private void Fermer_Click(object sender, RoutedEventArgs e) => Close();
}
