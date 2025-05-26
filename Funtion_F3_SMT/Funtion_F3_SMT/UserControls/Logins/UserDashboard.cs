using Funtion_F3_SMT;
using OK2SHIP_SMT.Services;
using OK2SHIP_SMT.Views;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Security.Authentication;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.UserControls.Logins
{
    public partial class UserDashboard : UserControl
    {

        public UserDashboard()
        {
            InitializeComponent();
            CheckState();
        }
        private void CheckState()
        {


        }

        private UC_Information infor = new UC_Information(new Dictionary<string, string>(), new List<Button>());
        private void btn_Create_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("User ID", "");
            dic.Add("User Name", "");
            dic.Add("Password", "");
            dic.Add("Role", "QA");
            dic.Add("Active", "ACTIVE");
            List<Button> buttons = new List<Button>();
            Button btn = new Button();
            btn.Text = "Save";
            btn.Click += btn_Create_ClickZ;
            buttons.Add(btn);
            infor = new UC_Information(dic, buttons);
            NormalForm c = new NormalForm(infor);
            c.ShowDialog();
        }
        private void btn_Create_ClickZ(object sender, EventArgs e)
        {
            Dictionary<string, string> dicz = infor.ListTextBox;
            while (true)
            {
                if (UserSession.Instance.IsLoggedIn == true && UserSession.Instance.Role.Equals("admin"))
                {
                    try
                    {

                        int res = UserSession.Instance.CreateUser(dicz["User Name"], dicz["Password"], dicz["User ID"], dicz["Role"], dicz["Active"].ToLower().Equals("active") ? 1 : 0);
                        if (res > 0)
                        {
                            break;
                        }
                        throw new Exception("lỗi không lưu được dữ liệu");

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        return;
                    }
                }
                else
                {
                    if (MessageBox.Show("Bạn cần đăng nhập", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        new Login().ShowDialog();
                    }
                    else
                    {
                        return;
                    }
                }
            }

            MessageBox.Show("Lưu thành công!");
        }

        private void btn_Search_Click(object sender, EventArgs e)
        {
            try
            {
                listAccount.DataSource = UserSession.Instance.ListOfUser();
                /// <exception cref="AuthenticationException"></exception>
            }
            catch (AuthenticationException z)
            {
                Console.WriteLine(z.Message);
                new Login().ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
