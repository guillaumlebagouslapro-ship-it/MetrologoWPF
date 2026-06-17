using CommunityToolkit.Mvvm.ComponentModel;
using Metrologo.Models;
using Metrologo.Services;
using Metrologo.Services.Journal;
using System.Windows;
using System.Windows.Input;
using Wpf.Ui.Controls;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Views
{
    /// <summary>
    /// Fenêtre modale qui demande un login et un mot de passe avant d'ouvrir la zone admin.
    /// Les identifiants sont contrôlés par <see cref="ComptesLocauxService.AuthentifierAdmin"/> ;
    /// si tout est bon, le compte connecté est mis à disposition dans <see cref="AdminAuthentifie"/>.
    /// </summary>
    public partial class SaisieMdpAdminWindow : FluentWindow
    {
        public SaisieMdpAdminVM ViewModel { get; } = new();

        /// <summary>Le compte connecté ; reste à null tant qu'aucune saisie valide n'a abouti.</summary>
        public Utilisateur? AdminAuthentifie { get; private set; }

        public SaisieMdpAdminWindow()
        {
            InitializeComponent();
            DataContext = ViewModel;
            Loaded += (_, _) => LoginBox.Focus();
        }

        private void OnValider(object sender, RoutedEventArgs e) => Verifier();

        private void OnAnnuler(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void OnPasswordKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter) Verifier();
        }

        private void Verifier()
        {
            string login = LoginBox.Text ?? string.Empty;
            string mdp = MdpBox.Password ?? string.Empty;

            var admin = ComptesLocauxService.AuthentifierAdmin(login, mdp);
            if (admin != null)
            {
                AdminAuthentifie = admin;
                ViewModel.MessageErreur = string.Empty;
                JournalLog.Info(CategorieLog.Administration, "ACCES_ADMIN_OK",
                    $"Connexion admin : {admin.Login} ({admin.Role}).");
                DialogResult = true;
                Close();
                return;
            }

            ViewModel.MessageErreur = "Identifiant ou mot de passe incorrect";

            MdpBox.Clear();
            MdpBox.Focus();
            JournalLog.Warn(CategorieLog.Administration, "ACCES_ADMIN_KO",
                $"Tentative d'accès admin refusée (login = « {login} »).");
        }
    }

    public partial class SaisieMdpAdminVM : ObservableObject
    {
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _messageErreur = string.Empty;

        public bool HasError => !string.IsNullOrEmpty(MessageErreur);
    }
}
