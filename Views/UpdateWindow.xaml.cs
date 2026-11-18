using System.Windows;

namespace Metrologo.Views;

/// <summary>
/// Petite fenetre affichee pendant une mise a jour : etape en cours (1/3, 2/3, 3/3),
/// barre de progression reelle et detail (Mo copies).
/// </summary>
public partial class UpdateWindow : Window
{
    public UpdateWindow() => InitializeComponent();

    /// <summary>Titre en gras (ex. « Mise à jour de Metrologo 2.3 → 2.5 »).</summary>
    public void SetTitre(string titre) => TitleText.Text = titre;

    /// <summary>Met a jour le texte d'etat sous le titre.</summary>
    public void SetStatus(string message) => StatusText.Text = message;

    /// <summary>Annonce une nouvelle etape et remet la barre a zero.</summary>
    public void SetEtape(int numero, int total, string texte)
    {
        StatusText.Text = $"Étape {numero}/{total} : {texte}";
        DetailText.Text = string.Empty;
        SetProgress(0);
    }

    /// <summary>Affiche une progression precise (0 a 100 %), avec un detail facultatif.</summary>
    public void SetProgress(int percent, string? detail = null)
    {
        Bar.IsIndeterminate = false;
        Bar.Value = percent;
        PercentText.Text = percent + " %";
        if (detail != null) DetailText.Text = detail;
    }

    /// <summary>Affiche une progression indeterminee (etape sans pourcentage).</summary>
    public void SetIndeterminate(string? detail = null)
    {
        Bar.IsIndeterminate = true;
        PercentText.Text = string.Empty;
        DetailText.Text = detail ?? string.Empty;
    }
}
