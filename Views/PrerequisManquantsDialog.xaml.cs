using Metrologo.Services;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;
using System.Windows.Navigation;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class PrerequisManquantsDialog : FluentWindow
    {
        public PrerequisManquantsDialog(IEnumerable<VerificationPrerequis.Prerequis> manquants)
        {
            InitializeComponent();
            LstPrerequis.ItemsSource = manquants;
        }

        private void OnContinuer(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void OnQuitter(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        /// <summary>Ouvre le lien de téléchargement dans le navigateur par défaut.</summary>
        private void OnHyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = e.Uri.AbsoluteUri,
                    UseShellExecute = true,
                });
                e.Handled = true;
            }
            catch
            {
                // Best-effort : si le browser n'est pas configuré, on laisse tomber.
            }
        }
    }
}
