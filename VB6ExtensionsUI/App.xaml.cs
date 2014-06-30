using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using VB6Extensions.Parser;

namespace VB6ExtensionsUI
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            var parser = new CodeFileParser();
            var tree = parser.Parse(@"VB6\SqlCommand.cls");
            var viewModel = new TreeNodeViewModel(tree);

            var view = new MainWindow();
            view.DataContext = viewModel;

            view.ShowDialog();
        }
    }
}
