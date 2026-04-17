using Metrologo.ViewModels;
using System.Windows;

namespace Metrologo
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            // CETTE LIGNE EST LE BRANCHEMENT :
            this.DataContext = new MainViewModel();
        }
    }
}