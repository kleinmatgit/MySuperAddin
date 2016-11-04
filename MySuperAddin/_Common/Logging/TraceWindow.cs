using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MySuperAddin._Common.Logging
{
    public partial class TraceWindow : Form
    {
        public TraceWindow()
        {
            InitializeComponent();
            listView1.HeaderStyle = ColumnHeaderStyle.None;
            Show();
        }

        private void TraceWindow_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Hide();
            e.Cancel = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //TO DO: implements
        }
    }
}
