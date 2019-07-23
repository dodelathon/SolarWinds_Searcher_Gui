using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SolarWinds_Searcher_Gui.Windows.ErrorPopUps
{
    public partial class UnableToOpenPopup : Form
    {

        MainWindow main;
        public UnableToOpenPopup(MainWindow parent)
        {
            main = parent;
            main.IsEnabled = false;
            InitializeComponent();
            FormClosed += Closed;
        }

        private new void Closed(object sender, EventArgs e)
        {
            main.IsEnabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            main.IsEnabled = true;
            Close();
        }
    }
}
