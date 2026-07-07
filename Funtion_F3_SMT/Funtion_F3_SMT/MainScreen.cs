using OK2SHIP_SMT.Repositories;
using OK2SHIP_SMT.Services;
using OK2SHIP_SMT.UserControls;
using OK2SHIP_SMT.UserControls.Logins;
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
            if (!checkLogin())
            {
                Application.Exit();
            }

            LoadStatusAsync();

        }

        private async void LoadStatusAsync()
        {
            // Cập nhật trạng thái đang kiểm tra (tùy chọn)
            lb_database.Text = "Database: Checking...";
            lb_nas.Text = "Nas: Checking...";

            // Chạy việc kiểm tra trong một luồng nền (Background Thread)
            // Task.Run giúp UI không bị treo
            await Task.Run(() => checkingStatus());

        }

        private void checkingStatus()
        {
            // Kiểm tra Database
            try
            {
                bool dbOk = checkingExport();
                // Cập nhật UI từ luồng nền cần dùng Invoke
                this.Invoke(new Action(() =>
                {
                    lb_Export.Text = dbOk ? "Export: Connected" : "Export: Disconnected";
                    lb_Export.BackColor = dbOk ? Color.Green : Color.Red;
                }));
            }
            catch (Exception)
            {
                try
                {

                    this.Invoke(new Action(() =>
                    {
                        lb_Export.Text = "Export: Error";
                        lb_Export.BackColor = Color.Red;
                    }));
                }
                catch
                {

                }
            }
            // Kiểm tra Database
            try
            {
                bool dbOk = checkingDB();
                // Cập nhật UI từ luồng nền cần dùng Invoke
                this.Invoke(new Action(() =>
                {
                    lb_database.Text = dbOk ? "Database: Connected" : "Database: Disconnected";
                    lb_database.BackColor = dbOk ? Color.Green : Color.Red;
                }));
            }
            catch (Exception)
            {
                try
                {

                    this.Invoke(new Action(() =>
                    {
                        lb_database.Text = "Database: Error";
                        lb_database.BackColor = Color.Red;
                    }));
                }
                catch { }
            }

            // Kiểm tra NAS
            try
            {
                bool nasOk = checkingNAS();
                this.Invoke(new Action(() =>
                {
                    lb_nas.Text = nasOk ? "Nas: Connected" : "Nas: Disconnected";
                    lb_nas.BackColor = nasOk ? Color.Green : Color.Red;
                }));
            }
            catch (Exception)
            {
                try
                {

                    this.Invoke(new Action(() =>
                    {
                        lb_nas.Text = "Nas: Error";
                        lb_nas.BackColor = Color.Red;
                    }));
                }
                catch { }
            }
        }
        private bool checkingNAS()
        {
            NasRepository nas = new NasRepository();
            return true;
        }
        private bool checkingExport()
        {
            ExportProcess export = new ExportProcess();
            return true;
        }
        private bool checkingDB()
        {
            DBContext _db = new DBContext();
            return true;
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
            UC_NewFeature uC = new UC_NewFeature();
            CommonForm frm = new CommonForm("Format Setup", uC, true);
            frm.Show();
            this.Hide();
        }



        private void lblSearchData_Click(object sender, EventArgs e)
        {

        }

        private void lblExport_Click(object sender, EventArgs e)
        {

        }
        private bool checkLogin()
        {
            if (!UserSession.Instance.IsLoggedIn)
            {
                Login loginForm = new Login();
                loginForm.ShowDialog();
            }
            return UserSession.Instance.IsLoggedIn;
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

        private void button1_Click_1(object sender, EventArgs e)
        {
            CommonForm frm = new CommonForm("ACCOUNT MANAGER", new UserDashboard());
            frm.Show();
            this.Hide();
        }

        private void lb_Export_Click(object sender, EventArgs e)
        {
            string msg = "";
            try
            {

                ExportProcess export = new ExportProcess();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

        }
    }
}
