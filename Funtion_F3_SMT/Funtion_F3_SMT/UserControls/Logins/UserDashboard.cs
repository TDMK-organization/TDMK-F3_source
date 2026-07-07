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
        private AccountService _service = new AccountService();


        public UserDashboard()
        {
            InitializeComponent();


            CheckState();
        }
        private void CheckState()
        {
            if (UserSession.Instance.User_ID != "ADMIN")
            {
                tabPage1.Controls.Add(new DatabaseTool());
                tabControl.TabPages.Remove(tp_AccountManager);
            }
            if (UserSession.Instance.User_ID == "ADMIN")
            {
                tabPage1.Controls.Add(new DatabaseTool());
                tabControl.TabPages.Remove(tp_EditProfile);
            }
            if (UserSession.Instance.IsLoggedIn)
            {
                yourUserName.Text = UserSession.Instance.Username;

            }

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

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {

                string result = _service.UpdateAccount((DataTable)listAccount.DataSource);
                MessageBox.Show(result);
                btn_Search_Click(sender, e); // Refresh the list after deletion
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            // Delete Account
            if (MessageBox.Show("Bạn có chắc chắn muốn xóa tài khoản này không?", "Xác nhận", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                try
                {
                    string[] userIds = listAccount.SelectedCells
                    .Cast<DataGridViewCell>()
                    .Select(cell => cell.OwningRow)
                    .Distinct()
                    .Select(row => row.Cells["User_ID"].Value?.ToString())
                    .Where(id => !string.IsNullOrEmpty(id)) // Loại bỏ giá trị null hoặc rỗng
                    .ToArray();


                    string result = _service.DeleteAccount(userIds);
                    MessageBox.Show(result);
                    btn_Search_Click(sender, e); // Refresh the list after deletion
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // Deactive Account 
        }

        private void tdmK_Button1_Click(object sender, EventArgs e)
        {
            string pass = null;
            if(!yourPass.Text.Equals("") )
            {
                pass = yourPass.Text;  
            }
            AccountService accountService = new AccountService();
            string res = accountService.UpdateAccount(yourUserName.Text, pass);
            MessageBox.Show(res);
        }
    }
}
