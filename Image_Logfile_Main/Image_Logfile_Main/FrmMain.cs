using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using VHX;
using Echeck_LogFile_Process;
using Manual_Input;

namespace OK2SHIP
{
    public partial class FrmMain : Form
    {
        public FrmMain()
        {
            InitializeComponent();
        }

        private void lblType2_Click(object sender, EventArgs e)
        {
            myVar._frmCAM = new FrmCamera();
            myVar._frmCAM.Show();
            this.Hide();
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
            
        }

        private void lblType3_Click(object sender, EventArgs e)
        {
            myVar._frmVHX = new FrmMainVHX();
            myVar._frmVHX.Show();
            this.Hide();
        }

        private void lblType1_Click(object sender, EventArgs e)
        {
            myVar._frmECheck = new ECheck_Main();
            myVar._frmECheck.Show();
            this.Hide();
        }

        private void lblType4_Click(object sender, EventArgs e)
        {
            myVar._frmManual = new Frm_Manual_input();
            myVar._frmManual.Show();
            this.Hide();
        }

        private void lblTest_Click(object sender, EventArgs e)
        {
            myVar._frmCheckLogfile = new Check_Impedance_logfile();
            myVar._frmCheckLogfile.Show();
            this.Hide();
        }
    }
}
