using OK2SHIP_SMT.UserControls;
using OK2SHIP_SMT.Views;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Funtion_F3_SMT
{
    public partial class MainScreen : Form
    {

        public MainScreen()
        {
            InitializeComponent();
        }


        private void lblExit_Click(object sender, EventArgs e)
        {
            this.Close();
            Application.Exit();
        }

        private void MainScreen_Load(object sender, EventArgs e)
        {

        }




        private void lblFormatSetup_Click(object sender, EventArgs e)
        {
            //Setup_Spec_SMT frmFormat = new Setup_Spec_SMT();
            //frmFormat.Show();
            //this.Hide();
        }



        private void lblSearchData_Click(object sender, EventArgs e)
        {

        }

        private void lblExport_Click(object sender, EventArgs e)
        {

        }

        private void lblInputdata_Click(object sender, EventArgs e)
        {
            FrmMain frmInput = new FrmMain();
            frmInput.Show();
            this.Hide();
        }

        private void MainScreen_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SwitchScreen("Table Of Content");
        }

        private void SwitchScreen(string process)
        {
            try
            {
                switch (process)
                {
                    case "Table Of Content":
                        TableOfContent toc = new TableOfContent();
                        CommonForm frm = new CommonForm("Table Of Content", toc);
                        frm.Show();
                        break;
                    default:
                        throw new Exception("Process not found");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

            this.Hide();

        }
    }
}
