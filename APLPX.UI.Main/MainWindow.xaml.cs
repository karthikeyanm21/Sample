using APLPX.Modules.StagingDBConfig.ViewModels;
using APLPX.UI.Main.ViewModels;
using MahApps.Metro.Controls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace APLPX.UI.Main
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {

        public MainWindow(MainViewModel viewModel)
        {
            DataContext = viewModel;

            var themes = MahApps.Metro.ThemeManager.AppThemes;

            var accent = MahApps.Metro.ThemeManager.Accents.First(x => x.Name == "Blue");
            var theme = MahApps.Metro.ThemeManager.AppThemes.First(x => x.Name == "BaseDark");

            MahApps.Metro.ThemeManager.ChangeAppStyle(Application.Current, accent, theme);

            InitializeComponent();
        }      
    }
}
