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
using OK2SHIP_Measurements.Services;

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
        public Login()
        {
            InitializeComponent();
            this.admin_mode = notThing;
        }
        private void notThing(string str)
        {
            // Do nothing
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
            try
            {


                if (btnLogin.Text == "Login")
                {

                    bool login_en = false;
                    string userName = txtUsername.Text.Trim();
                    string password = txtPassword.Text.Trim();
                    UserSession.Instance.Login(userName, password);
                    login_en = UserSession.Instance.IsLoggedIn;
                    //if ((txtUsername.Text == "Admin") && (txtPassword.Text == "TDMK"))
                    //{
                    //    login_en = true;
                    //}
                    //else
                    //{
                    //    //DataTable info = TDMK_Code.Datatable_Filter(sql_login, "ACCOUNT_USER", TDMK_Code.filter_str(new string[] { "UserName", "Password" }, new string[] { txtUsername.Text, txtPassword.Text }));
                    //    if (info.Rows.Count == 0)
                    //    {
                    //        login_en = false;
                    //    }
                    //    else
                    //    {
                    //        login_en = true;
                    //    }
                    //}
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
                else if (btnLogin.Text == "Logout")
                {
                    this.admin_mode("LOG IN");
                    this.Hide();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Login_Load(object sender, EventArgs e)
        {

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
