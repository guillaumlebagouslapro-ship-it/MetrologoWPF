using System.Windows;

namespace Metrologo.Views;

/// <summary>
/// Petite fenetre affichee pendant le telechargement / l'installation d'une mise
/// a jour, avec une barre de progression reelle (pourcentage du telechargement).
/// </summary>
public partial class UpdateWindow : Window
{
    public UpdateWindow() => InitializeComponent();

    /// <summary>Met a jour le texte d'etat sous le titre.</summary>
    public void SetStatus(string message) => StatusText.Text = message;

    /// <summary>Affiche une progression precise (0 a 100 %).</summary>
    public void SetProgress(int percent)
    {
        Bar.IsIndeterminate = false;
        Bar.Value = percent;
        PercentText.Text = percent + " %";
    }

    /// <summary>Affiche une progression indeterminee (etape sans pourcentage).</summary>
    public void SetIndeterminate()
    {
        Bar.IsIndeterminate = true;
        PercentText.Text = string.Empty;
    }
}
