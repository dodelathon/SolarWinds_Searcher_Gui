using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SolarWinds_Searcher_Gui
{
    public partial class SearcherThreadErrorPopUp : Form
    {
        MainWindow main;
        public SearcherThreadErrorPopUp(MainWindow parent)
        {
            main = parent;
            InitializeComponent();
            main.IsEnabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            main.IsEnabled = true;
            Close();
        }
    }
}
