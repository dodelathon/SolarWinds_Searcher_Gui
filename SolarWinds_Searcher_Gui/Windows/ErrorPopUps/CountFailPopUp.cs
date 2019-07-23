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
    public partial class CountFailPopUp : Form
    {
        MainWindow parent;
        public CountFailPopUp(MainWindow parent)
        {
            this.parent = parent;
            InitializeComponent();
            parent.IsEnabled = false;
            FormClosed += Closed;

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private new void Closed(object sender, EventArgs e)
        {
            parent.IsEnabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            parent.IsEnabled = true;
            Close();
        }
    }
}
