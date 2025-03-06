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
using System.Data.SqlClient;
using System.Diagnostics.Eventing.Reader;

namespace Funtion_F3_SMT
{
    public partial class Login : Form
    {
        public SqlConnection sql_login = new SqlConnection();
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        private bool en_ = false;
        public bool mode_ = false;


        //public bool admin_mode = false;

        public bool mode
        {
            get { return mode_; }
            set { mode_ = value; }
        }

        public Login(Log_en sender)
        {

            InitializeComponent();
            this.admin_mode = sender;
        }
        public delegate void Log_en(string mode);

        public Log_en admin_mode;


        //Main frm1 = new Main();
        public string depart;

        //public Login()
        //{
        //    InitializeComponent();
        //}

        public bool en
        {
            get { return en_; }
            set { en_ = value; }
        }

        private void OK_Click(object sender, EventArgs e)
        {
            
            if (btnLogin.Text == "Login")
            { 
                bool login_en = false;
                if ((txtUsername.Text == "Admin") && (txtPassword.Text == "TDMK"))
                {
                    login_en = true;
                }
                else
                {
                    //DataTable info = TDMK_Code.Datatable_Filter(sql_login, "USerInfo", TDMK_Code.filter_str(new string[] { "User_Name", "User_Password", "Department" }, new string[] { txtUsername.Text, txtPassword.Text, depart }));
                    //if (info.Rows.Count == 0)
                    //{
                    //    login_en = false;
                    //}
                    //else
                    //{
                    //    login_en = true;
                    //}
                }
                if (login_en)
                {
                    //FindControl fc_btnLogin = new FindControl();
                    //Button btnLogin = (Button)fc_btnLogin.Ctrl(myVar_ECheck.frmMain, "btnLogin");
                    //myVar_ECheck.confirm_mode = true; 
                    //btnLogin.BackColor = Color.GreenYellow; 
                    
                    btnLogin.Text = "Logout";
                    this.admin_mode("Admin mode"); 
                    this.Hide();
                }
                else
                {
                    MessageBox.Show("Sai mật khẩu!");
                }
            }
            else if(btnLogin.Text == "Logout")
            {
                this.admin_mode("LOG IN");
                this.Hide();
            }
        }

        private void Login_Load(object sender, EventArgs e)
        {
            //sql_login = myVar.mysqlcon; //myVar_ECheck.frmMain.initial_data("OK2SHIP_Items");
            rbQA.Checked = true;

            if (mode_)
            {
                btnLogin.Text = "Logout";
                txtPassword.Text = "TDMK";
               
            }
            else
            {
                btnLogin.Text = "Login";
            }
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
            this.Hide();
        }

        private void txtPassword_TextChanged(object sender, EventArgs e)
        {
             
        }
             

        private void txtUsername_TextChanged(object sender, EventArgs e)
        {
            btnLogin.Text = "Login";
        }
    }
}
