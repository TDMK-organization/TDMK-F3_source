using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using TDMK_SQL;
using myExcel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace OK2SHIP_Software
{
    public partial class FrmLogin : Form
    {
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public static string app_path;
        public static SqlConnection sqlcon;
        myVar myCode = new myVar();
        public string depart = "QA";
        public FrmLogin()
        {
            InitializeComponent();
        }

        private void OK_Click(object sender, EventArgs e)
        {
            bool login_en = false;
            if ((txtUsername.Text == "Admin") && (txtPassword .Text =="TDMK"))
            {
                login_en = true;
            }
            else
            {
                //DataTable info =  TDMK_Code.Datatable_Filter(myVar.sqlcon_OK2SHIP, "USerInfo", TDMK_Code.filter_str(new string[] { "User_Name", "User_Password", "Department" }, new string[] { txtUsername.Text, txtPassword.Text, depart }));
                //if(info.Rows.Count==0)
                //{
                //    login_en = false;
                //}
                //else
                //{
                //    login_en = true;
                //}
            }
            if(login_en)
            {
                myVar.UserID = txtUsername.Text;
                myVar.USerDepart = depart;
                if(!myVar.confirm_mode)
                {
                    myVar.frm_mainscreen = new MainScreen();
                    myVar.frm_mainscreen.Show();
                    this.Hide();
                }
                else
                {
                    FindControl fc_DGV = new FindControl();
                    DataGridView DGV_Data = (DataGridView)fc_DGV.Ctrl(myVar.frm_export, "DGV_Data");
                    DataGridViewCell curr_cell = DGV_Data.CurrentCell;
                    curr_cell.Value = myVar.edit_val;
                    string cur_item = DGV_Data.Columns[curr_cell.ColumnIndex].Name;
                    TDMK_Code.update_item_val("Approve_Data", myVar.sqlcon_OK2SHIP, "ID", DGV_Data.Rows[curr_cell.RowIndex].Cells["ID"].Value.ToString(), cur_item, myVar. edit_val);
                    this.Hide();
                }
            }
            else
            {
                MessageBox.Show("Sai mật khẩu!");
            }
        }

        private void FrmLogin_Load(object sender, EventArgs e)
        {
            if(!myVar.confirm_mode)
            {
                app_path = Application.StartupPath;
                rbQA.Checked = true;
                myVar.app_path = Application.StartupPath;
                //listBox1.DataSource = myCode.GetAllTables(myVar.sqlcon_Recycle).ToList();
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            string src_file = Path.Combine(myVar.app_path, "Temp_Folder", "NPI OK2Ship Report for Bare FPC_22B050000003.xlsm");//"NPI OK2Ship Report for Bare FPC_22B050000003.xlsm"
            string tar_folder = @"D:\Customer Projects\SEEV\TestAreas\OK2SHIP automation system";
             

            /***********************  Export FAI Data **********************************************/
            //myCode.Export_FAI_Data(src_file, "22B0500", "00003");
            /*********************** Finish export FAI Data **********************************************/

            /***********************  Export IPQC Data **********************************************/
            //myCode.Export_IPQC_Data(src_file, "22B0500", "00003");
            /*********************** Finish export IPQC Data **********************************************/

            /***********************  Export Materials Data **********************************************/
            //myCode.Export_Materials_Data(src_file, "22B0500", "00003");
            /*********************** Finish export Materials Data **********************************************/

            /***********************  Export Recycle_PGC_Copper Data **********************************************/
            //myCode. Export_Recycle_Data(src_file, "22B0500", "00003");
            /*********************** Finish export Recycle_PGC_Copper Data **********************************************/

            /***********************  Export Others from Excel Data **********************************************/
            // myCode.Export_fromExcel(src_file, "22B0358", "00003", tar_folder);
            /*********************** Finish Export Others from Excel Data **********************************************/

        }

        private void rbQA_CheckedChanged(object sender, EventArgs e)
        {
            if (rbQA.Checked)
            {
                depart = rbQA.Text;
            }
        }

        private void rbDE_CheckedChanged(object sender, EventArgs e)
        {
            if (rbDE.Checked)
            {
                depart = rbDE.Text;
            }
        }

        private void rbNPI_CheckedChanged(object sender, EventArgs e)
        {
            if (rbNPI.Checked)
            {
                depart = rbNPI.Text;
            }
        }

        private void rbPro_CheckedChanged(object sender, EventArgs e)
        {
            if (rbPro.Checked)
            {
                depart = rbPro.Text;
            }
        }

        private void rbPE_CheckedChanged(object sender, EventArgs e)
        {
            if (rbPE.Checked)
            {
                depart = rbPE.Text;
            }
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
            //Application.Exit();
        }
    }
}
