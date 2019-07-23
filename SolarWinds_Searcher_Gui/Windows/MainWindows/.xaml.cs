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
using System.Windows.Shapes;

namespace SolarWinds_Searcher_Gui
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        private CustomParams par;
        public Window1(CustomParams parent)
        {
            SizeChanged += Revert_Size;
            InitializeComponent();
            par = parent;
            par.IsEnabled = false;
            Closing += Window_Closing;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            par.IsEnabled = true;
        }

        private void Revert_Size(object sender, SizeChangedEventArgs e)
        {
            Width = 267.606;
            Height = 192.254; 
        }
    }
}
