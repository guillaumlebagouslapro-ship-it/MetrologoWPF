using System.Windows;
using Metrologo.ViewModels;

namespace Metrologo.Views
{
    public partial class ConfigurationWindow : Window
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