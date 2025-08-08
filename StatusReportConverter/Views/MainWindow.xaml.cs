using System.Windows;
using StatusReportConverter.ViewModels;

namespace StatusReportConverter.Views
{
    public partial class MainWindow : Window
    {
        public MainWindow(MainViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }
    }
}