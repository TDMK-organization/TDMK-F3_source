using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.Views
{
    public partial class NormalForm : Form
    {
        public NormalForm(UserControl uc)
        {
            InitializeComponent();
            uc.Dock = DockStyle.Fill;
            tableLayoutPanel.Controls.Add(uc);
        }
    }
}
