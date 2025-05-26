using OK2SHIP_Lib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SQL;
using OK2SHIP_Software;

namespace OK2SHIP_Measurements
{
    public partial class FrmLogin : Form
    {
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public FrmLogin(string src_str)
        {
            InitializeComponent();
            this.Value = src_str;
        }

        public FrmLogin()
        {
            InitializeComponent();
            ///txtTest.Text = src_str;
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            txtPass.Text = Value;
            //string connstr_OK2SHIP = TDMK_Code.data_connection(myVar.server_name, "OK2SHIP_Items", myVar.server_acc, myVar.server_pass).ConnectionString;
        }
        public string Value { get; set; }

        private void btnConfirm_Click(object sender, EventArgs e)
        {
            bool login_en = false;
            if ((txtPass.Text == "TDMK") && (txtUserID.Text =="Admin"))
            {
                //TDMK_OK2SHIP.admin_mode = true;
                //FindControl fc = new FindControl();
                //FindControl fc_numQty = new FindControl();
                //FindControl fc_GBControl = new FindControl();
                //if (TDMK_OK2SHIP.FAI_form !=null)
                //{
                //    Button btnLogin = (Button)fc.Ctrl(TDMK_OK2SHIP.FAI_form, "btnLogin");
                //    btnLogin.Text = "Logout";
                //    NumericUpDown numQty = (NumericUpDown)fc_numQty.Ctrl(TDMK_OK2SHIP.FAI_form, "numQty");
                //    numQty.Enabled = true;
                //    GroupBox GBControl = (GroupBox)fc_GBControl.Ctrl(TDMK_OK2SHIP.FAI_form, "GBControl");
                //    GBControl.Enabled = true;
                //    btnLogin.BackColor = Color.Green;
                //}
                //if(TDMK_OK2SHIP.IPQC_Man_form!=null)
                //{
                //    NumericUpDown numQty = (NumericUpDown)fc.Ctrl(TDMK_OK2SHIP.IPQC_Man_form, "numQty");
                //    numQty.Enabled = true;
                //}
                //Dispose();
                login_en = true;
            }
            else
            {
                //TDMK_OK2SHIP.admin_mode = false;
                //MessageBox.Show("Mat khau khong hop le","Thong bao");
                //DataTable info = TDMK_Code.Datatable_Filter(myVar.sqlcon_OK2SHIP, "USerInfo", TDMK_Code.filter_str(new string[] { "User_Name", "User_Password", "Department" }, new string[] { txtUsername.Text, txtPassword.Text, depart }));
                //if (info.Rows.Count == 0)
                //{
                //    login_en = false;
                //}
                //else
                //{
                //    login_en = true;
                //}
                login_en = false;
            }
            if (login_en)
            {
                TDMK_OK2SHIP.admin_mode = true;
                FindControl fc = new FindControl();
                FindControl fc_numQty = new FindControl();
                FindControl fc_GBControl = new FindControl();
                if (myVar.FAI_form != null)
                {
                    Button btnLogin = (Button)fc.Ctrl(myVar.FAI_form, "btnLogin");
                    btnLogin.Text = "Logout";
                    NumericUpDown numQty = (NumericUpDown)fc_numQty.Ctrl(myVar.FAI_form, "numQty");
                    numQty.Enabled = true;
                    GroupBox GBControl = (GroupBox)fc_GBControl.Ctrl(myVar.FAI_form, "GBControl");
                    GBControl.Enabled = true;
                    btnLogin.BackColor = Color.Green;
                }
                if (myVar.IPQC_Man_form != null)
                {
                    NumericUpDown numQty = (NumericUpDown)fc.Ctrl(myVar.IPQC_Man_form, "numQty");
                    numQty.Enabled = true;
                }
                Dispose();
            }
            else
            {
                TDMK_OK2SHIP.admin_mode = false;
                MessageBox.Show("Mat khau khong hop le", "Thong bao");
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Dispose();
        }
    }
}
