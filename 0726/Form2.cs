using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace _0726
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void pGridView1_Scroll(object sender, ScrollEventArgs e)
        {
            pGridView2.FirstDisplayedScrollingRowIndex = pGridView1.FirstDisplayedScrollingRowIndex;
        }

        private void pGridView2_Scroll(object sender, ScrollEventArgs e)
        {
            pGridView1.FirstDisplayedScrollingRowIndex = pGridView2.FirstDisplayedScrollingRowIndex;
        }
    }
}
