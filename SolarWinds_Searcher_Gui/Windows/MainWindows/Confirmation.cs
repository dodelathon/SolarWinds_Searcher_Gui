using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SolarWinds_Searcher_Gui.Windows.MainWindows
{
    public partial class Confirmation : Form
    {
        private MainWindow parent;
        public Confirmation(MainWindow main)
        {
            parent = main;
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            parent.Start();
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {

            parent.Close();
        }
    }
}
