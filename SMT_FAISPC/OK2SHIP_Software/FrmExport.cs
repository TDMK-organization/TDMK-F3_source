using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_Software
{
    public partial class FrmExport : Form
    {
        myVar MyVar = new myVar();
        public FrmExport()
        {
            InitializeComponent();
        }

        private void btnLoadData_Click(object sender, EventArgs e)
        {
            string process = cbProcess.SelectedItem.ToString();
            FAI_Data frmFAI = new FAI_Data(txtItemCode.Text,txtLotNo.Text) { Dock = DockStyle.Fill, TopLevel = false, TopMost = true };
            this.splitContainer_main.Panel2.Controls.Add(frmFAI);
            frmFAI.Show();
        }

        private void FrmExport_Load(object sender, EventArgs e)
        {
            myVar.sqlcon_SMT = MyVar.initial_data("OK2SHIP_SMT", true);
        }
    }
}
