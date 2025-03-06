using SEEV_SMT_F3.TDMK_Tool;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace SEEV_SMT_F3
{
    public partial class MainScreen : Form
    {
        public MainScreen()
        {
            InitializeComponent();
            AutoFit();
        }


        private void AutoFit()
        {
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
        }

        private void btn_InputData_Click(object sender, EventArgs e)
        {
            TDMK_UI_Winform.OpenAndCloseForm(this, new FormHandleObject());
        }

        private void MainScreen_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }
    }
}
