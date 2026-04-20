using System.Windows;
using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class ConfigurationWindow : FluentWindow
    {
        public ConfigurationWindow(ConfigurationViewModel vm)
        {
            InitializeComponent();
            DataContext = vm;

            vm.CloseAction = result =>
            {
                this.DialogResult = result;
            };
        }

        private void ComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            (DataContext as ConfigurationViewModel)?.RefreshAll();
        }

        private void TypeMesure_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            // On appelle la logique du ViewModel
            (DataContext as ConfigurationViewModel)?.OnTypeMesureChanged();
        }
    }
}