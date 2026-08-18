using Funtion_F3_SMT;
using OK2SHIP_SMT.UserControls;
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
    public partial class CommonForm : Form
    {
        public string PROCESSNAME { get; set; }
        public UserControl UserControl { get; set; }
        private bool? notbackToMenu = false;
        public CommonForm(string process = null, UserControl uc = null, bool? notbackToMenu = false)
        {
            if (process != null)
            {
                this.PROCESSNAME = process;
            }
            else
            {
                this.PROCESSNAME = "TDMK Screen";
            }
            UserControl = uc;
            InitializeComponent();
            this.notbackToMenu = notbackToMenu;
            this.label.Text = this.PROCESSNAME;
     
        }
        private void CommonForm_Load(object sender, EventArgs e)
        {  
            SetUpForm(UserControl);
        }
        #region Event
        private void CommonForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (notbackToMenu == null)
            {
                return;
            }


            if (notbackToMenu == false)
            {
                BackToMain();
                return;
            }
            ExitProgram();

        }

        private void backToMainMenuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BackToMain();
        }

        private void resetWindowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Reset(this);
        }
        #endregion
        #region Action
        private void ExitProgram()
        {
            Application.Exit();
        }
        private void Reset(CommonForm from)
        {
            this.Hide();
            new CommonForm(from.PROCESSNAME, from.UserControl).Show();
        }
        public void BackToMain()
        {
            this.Hide();
            new MainScreen().Show();
        }
        private void SetUpForm(UserControl userControl)
        {
            this.WindowState = FormWindowState.Maximized;
            userControl.Dock = DockStyle.Fill;
            spc_Main.Panel2.Controls.Add(userControl);

        }



        #endregion

        private void exitProgramToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExitProgram();
        }

       
    }
}
