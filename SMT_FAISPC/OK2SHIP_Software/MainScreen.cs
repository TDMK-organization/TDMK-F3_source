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
    public partial class MainScreen : Form
    {
        public MainScreen()
        {
            InitializeComponent();
        }
        private void lblExport_Click(object sender, EventArgs e)
        {
            FrmMain _frmMain = new FrmMain();
            _frmMain.Show();
            //myVar.frm_export = new FrmExport();
            //myVar.frm_export.Show();
            this.Hide();
        }

        private void lblExit_Click(object sender, EventArgs e)
        {
            this. Close();
            Application.Exit();
        }

        private void MainScreen_Load(object sender, EventArgs e)
        {

        }
        private void lblSPC_Click(object sender, EventArgs e)
        {
            DE_Data_Info Frm_DE = new DE_Data_Info();
            Frm_DE.Show();
            this.Hide();
        }

        private void lblUser_Setup_Click(object sender, EventArgs e)
        {

        }

        private void lblUserRegister_Click(object sender, EventArgs e)
        {
            User_Register frmUser = new User_Register();
            frmUser.Show();
        }

        private void lblFormatSetup_Click(object sender, EventArgs e)
        {
            FormatSetup frmFormat = new FormatSetup();
            frmFormat.Show();
        }

        private void label1_Click(object sender, EventArgs e)
        {
            Edit_ItemsText frmItem_Edit = new Edit_ItemsText();
            frmItem_Edit.Show();
            this.Hide();
        }

        private void lblSearchData_Click(object sender, EventArgs e)
        {
            FrmAnalysis frmAnalys = new FrmAnalysis();
            frmAnalys.Show();
        }
    }
}
