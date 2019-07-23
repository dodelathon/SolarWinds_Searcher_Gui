using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
    /// Interaction logic for CustomParams.xaml
    /// </summary>
    /// 
    public partial class CustomParams : Window
    {
        private string[] Attributes;
        private int col;
        private int row;
        private bool submit_btn;
        private Window1 error;
        private MainWindow main;

        public CustomParams(MainWindow parent, string[] temp)
        {
            Attributes = temp;
            main = parent;
            main.IsEnabled = false;
            row = 10;
            col = 10;
            MinHeight = 390;
            MinWidth = 400;
            submit_btn = false;
            InitializeComponent();
            Closing += WindowClosing;
        }

        private void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            main.callback();
            submit_btn = true;
            Close();

        }

        private void ColInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            Regex rx = new Regex("[0-9]{1,3}");
            Match res = rx.Match(ColInput.Text);
            if (res != null && res.Success && ColInput.Text.Length <= 3)
            {
                col = int.Parse(ColInput.Text);
            }
            else if(ColInput.Equals(""))
            {
            }
            else
            {
                error = new Window1(this);
                error.Activate();
                error.Visibility = Visibility.Visible;
                ColInput.Text = "";
            }
        }

        private void RowInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            Regex rx = new Regex("[0-9]{1,3}");
            Match res = rx.Match(RowInput.Text);
            if (res != null && res.Success && RowInput.Text.Length <= 3)
            {
                row = int.Parse(RowInput.Text);
            }
            else if (RowInput.Equals(""))
            { 
            }
            else
            {
                error = new Window1(this);
                error.Activate();
                error.Visibility = Visibility.Visible;
                RowInput.Text = "";
            }
        }
           
        public object[] GetAll()
        {
            return new object[] {row, col};
        }

        private void WindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (submit_btn == false)
            {
                main.CustomBox.IsChecked = false;
            }
            main.IsEnabled = true;
        }


    }
}
