using System.Windows;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class InfoCompteCreeDialog : FluentWindow
    {
        public InfoCompteCreeDialog(string login, string motDePasse, string? titre = null)
        {
            InitializeComponent();
            TbLogin.Text = login;
            TbMdp.Text = motDePasse;
            if (!string.IsNullOrWhiteSpace(titre)) Title = titre;
        }

        private void OnCopierLogin(object sender, RoutedEventArgs e)
            => Clipboard.SetText(TbLogin.Text);

        private void OnCopierMdp(object sender, RoutedEventArgs e)
            => Clipboard.SetText(TbMdp.Text);

        private void OnFermer(object sender, RoutedEventArgs e)
            => Close();
    }
}
