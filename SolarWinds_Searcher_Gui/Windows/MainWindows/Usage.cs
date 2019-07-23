using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SolarWinds_Searcher_Gui.Windows
{
    public partial class Usage : Form
    {
        private MainWindow parent;
        private Credits credits;
        public Usage(MainWindow main)
        {
            parent = main;
            parent.IsEnabled = false;
            InitializeComponent();
            FormClosed += Closed;
        }

        private new void Closed(object sender, EventArgs e)
        {
            parent.IsEnabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            credits = new Credits(this);
            credits.Visible = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            parent.IsEnabled = true;
            Close();
        }
    }
}
