using System.Windows;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class InfoMdpGenereDialog : FluentWindow
    {
        public InfoMdpGenereDialog(string login, string motDePasseClair, string? titre = null)
        {
            InitializeComponent();
            if (!string.IsNullOrEmpty(titre)) TbTitre.Text = titre;
            TbLogin.Text = login;
            TbMdp.Text = motDePasseClair;
        }

        private void OnCopierLogin(object sender, RoutedEventArgs e) =>
            Clipboard.SetText(TbLogin.Text);

        private void OnCopierMdp(object sender, RoutedEventArgs e) =>
            Clipboard.SetText(TbMdp.Text);

        private void OnFermer(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
    }
}
