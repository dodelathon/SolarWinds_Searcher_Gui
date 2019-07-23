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
    public partial class Credits : Form
    {
        private Usage parent;
        public Credits(Usage parent)
        {
            this.parent = parent;
            parent.Enabled = false;
            InitializeComponent();
            FormClosed += Closed;
        }

        private new void Closed(object sender, EventArgs e)
        {
            parent.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            parent.Enabled = true;
            Close();
        }
    }
}
