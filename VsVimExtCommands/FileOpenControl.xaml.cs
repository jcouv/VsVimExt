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

namespace VsVimExtCommands
{
    /// <summary>
    /// Interaction logic for FileOpenControl.xaml
    /// </summary>
    public partial class FileOpenControl : UserControl
    {
        public FileOpenControl()
        {
            InitializeComponent();
        }


        // Allow the tool window to create the toolbar tray.  Set its style and
        // add it to the grid.
        public void SetTray(ToolBarTray tray)
        {
            //tray.Style = FindResource("ToolBarTrayStyle") as Style;
            //grid.Children.Add(tray);
        }
    }
}
