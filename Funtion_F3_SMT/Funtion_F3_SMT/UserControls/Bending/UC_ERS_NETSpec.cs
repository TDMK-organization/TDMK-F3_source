using System;
using System.Data;
using System.Windows.Forms;
using AntdUI;
using Funtion_F3_SMT;
using OK2SHIP_SMT.Services;

namespace OK2SHIP_SMT.UserControls.Bending
{
    public partial class UC_ERS_NETSpec : UserControl
    {
        private AntdUI.Window window;

        public UC_ERS_NETSpec(AntdUI.Window _window)
        {
            InitializeComponent();
            DisplayData();
            window = _window;
            updateDisplay();
        }

        private string __iconLogin =
            "<svg viewBox=\"0 0 24 24\" focusable=\"false\" width=\"1em\" height=\"1em\" fill=\"currentColor\" aria-hidden=\"true\"><path d=\"M13 2C10.2386 2 8 4.23858 8 7C8 7.55228 8.44772 8 9 8C9.55228 8 10 7.55228 10 7C10 5.34315 11.3431 4 13 4H17C18.6569 4 20 5.34315 20 7V17C20 18.6569 18.6569 20 17 20H13C11.3431 20 10 18.6569 10 17C10 16.4477 9.55228 16 9 16C8.44772 16 8 16.4477 8 17C8 19.7614 10.2386 22 13 22H17C19.7614 22 22 19.7614 22 17V7C22 4.23858 19.7614 2 17 2H13Z M3 11C2.44772 11 2 11.4477 2 12C2 12.5523 2.44772 13 3 13H11.2821C11.1931 13.1098 11.1078 13.2163 11.0271 13.318C10.7816 13.6277 10.5738 13.8996 10.427 14.0945C10.3536 14.1921 10.2952 14.2705 10.255 14.3251L10.2084 14.3884L10.1959 14.4055L10.1915 14.4115C10.1914 14.4116 10.191 14.4122 11 15L10.1915 14.4115C9.86687 14.8583 9.96541 15.4844 10.4122 15.809C10.859 16.1336 11.4843 16.0346 11.809 15.5879L11.8118 15.584L11.822 15.57L11.8638 15.5132C11.9007 15.4632 11.9553 15.3897 12.0247 15.2975C12.1637 15.113 12.3612 14.8546 12.5942 14.5606C13.0655 13.9663 13.6623 13.2519 14.2071 12.7071L14.9142 12L14.2071 11.2929C13.6623 10.7481 13.0655 10.0337 12.5942 9.43937C12.3612 9.14542 12.1637 8.88702 12.0247 8.7025C11.9553 8.61033 11.9007 8.53682 11.8638 8.48679L11.822 8.43002L11.8118 8.41602L11.8095 8.41281C11.4848 7.96606 10.859 7.86637 10.4122 8.19098C9.96541 8.51561 9.86636 9.14098 10.191 9.58778L11 9C10.191 9.58778 10.1909 9.58773 10.191 9.58778L10.1925 9.58985L10.1959 9.59454L10.2084 9.61162L10.255 9.67492C10.2952 9.72946 10.3536 9.80795 10.427 9.90549C10.5738 10.1004 10.7816 10.3723 11.0271 10.682C11.1078 10.7837 11.1931 10.8902 11.2821 11H3Z\"></path></svg>\n";

        private string __iconLogout =
            "<svg viewBox=\"0 0 24 24\" focusable=\"false\" width=\"1em\" height=\"1em\" fill=\"currentColor\" aria-hidden=\"true\"><path d=\"M12.9999 2C10.2385 2 7.99991 4.23858 7.99991 7C7.99991 7.55228 8.44762 8 8.99991 8C9.55219 8 9.99991 7.55228 9.99991 7C9.99991 5.34315 11.3431 4 12.9999 4H16.9999C18.6568 4 19.9999 5.34315 19.9999 7V17C19.9999 18.6569 18.6568 20 16.9999 20H12.9999C11.3431 20 9.99991 18.6569 9.99991 17C9.99991 16.4477 9.55219 16 8.99991 16C8.44762 16 7.99991 16.4477 7.99991 17C7.99991 19.7614 10.2385 22 12.9999 22H16.9999C19.7613 22 21.9999 19.7614 21.9999 17V7C21.9999 4.23858 19.7613 2 16.9999 2H12.9999Z M13.9999 11C14.5522 11 14.9999 11.4477 14.9999 12C14.9999 12.5523 14.5522 13 13.9999 13V11Z M5.71783 11C5.80685 10.8902 5.89214 10.7837 5.97282 10.682C6.21831 10.3723 6.42615 10.1004 6.57291 9.90549C6.64636 9.80795 6.70468 9.72946 6.74495 9.67492L6.79152 9.61162L6.804 9.59454L6.80842 9.58848C6.80846 9.58842 6.80892 9.58778 5.99991 9L6.80842 9.58848C7.13304 9.14167 7.0345 8.51561 6.58769 8.19098C6.14091 7.86637 5.51558 7.9654 5.19094 8.41215L5.18812 8.41602L5.17788 8.43002L5.13612 8.48679C5.09918 8.53682 5.04456 8.61033 4.97516 8.7025C4.83623 8.88702 4.63874 9.14542 4.40567 9.43937C3.93443 10.0337 3.33759 10.7481 2.7928 11.2929L2.08569 12L2.7928 12.7071C3.33759 13.2519 3.93443 13.9663 4.40567 14.5606C4.63874 14.8546 4.83623 15.113 4.97516 15.2975C5.04456 15.3897 5.09918 15.4632 5.13612 15.5132L5.17788 15.57L5.18812 15.584L5.19045 15.5872C5.51509 16.0339 6.14091 16.1336 6.58769 15.809C7.0345 15.4844 7.13355 14.859 6.80892 14.4122L5.99991 15C6.80892 14.4122 6.80897 14.4123 6.80892 14.4122L6.804 14.4055L6.79152 14.3884L6.74495 14.3251C6.70468 14.2705 6.64636 14.1921 6.57291 14.0945C6.42615 13.8996 6.21831 13.6277 5.97282 13.318C5.89214 13.2163 5.80685 13.1098 5.71783 13H13.9999V11H5.71783Z\"></path></svg>\n";

        private void btn_login_Click(object sender, EventArgs e)
        {
            if (UserSession.Instance.IsLoggedIn)
            {
                UserSession.Instance.Logout();
            }
            else
            {
                Login fr1 = new Login();
                fr1.ShowDialog();
            }

            updateSession();
        }

        private void updateSession()
        {
            tb_login.Enabled = false;
            if (UserSession.Instance.IsLoggedIn)
            {
                btn_login.IconSvg = __iconLogout;
                btn_login.Type = TTypeMini.Error;
                tb_login.Text = UserSession.Instance.Username;
            }
            else
            {
                btn_login.IconSvg = __iconLogin;
                btn_login.Type = TTypeMini.Primary;
                tb_login.Text = "";
            }
        }

        private void updateDisplay()
        {
            updateSession();
        }

        ERS_NetSpecService _service = new ERS_NetSpecService();
        TAlignFrom align = TAlignFrom.Top;

        private void button_loading_Click(object sender, EventArgs e)
        {
            try
            {
                string location = tb_logfile.Text;
                string itemCode = tb_itemCode.Text;
                string maker = tb_maker.Text;
                _service.getData(location.Trim(), itemCode.Trim(), maker.Trim());
                DisplayData();
            }
            catch (Exception ex)
            {
                AntdUI.Notification.error(window, $"Error", ex.Message, autoClose: 10, align: align);
            }
        }

        private void DisplayData()
        {
            dgv.DataSource = _service._DATA;
        }

        private void tb_logfile_TextChanged(object sender, EventArgs e)
        {
            try
            {
                string location = tb_logfile.Text.Trim();
                _service.getInfo(location, out string itemCode, out string maker);
                tb_itemCode.Text = itemCode;
                tb_maker.Text = maker;
            }
            catch (Exception ex)
            {
                AntdUI.Notification.error(window, $"Error", ex.Message, autoClose: 10, align: align);
            }
        }

        private void btn_Save_Click(object sender, EventArgs e)
        {
            string itemCode = tb_itemCode.Text.Trim();
            string maker = tb_maker.Text.Trim();
            try
            {
                _service.Save(itemCode, maker);
                AntdUI.Notification.success(window, $"Info", "Đẩy dữ liệu hoàn tất", autoClose: 10, align: align);
            }
            catch (DataException de)
            {
                try
                {
                    if (MessageBox.Show(de.Message, "Thông báo!", MessageBoxButtons.YesNo, MessageBoxIcon.Error) ==
                        DialogResult.Yes)
                    {
                        _service.Save(itemCode, maker, true);
                        AntdUI.Notification.success(window, $"Info", "Đẩy dữ liệu hoàn tất", autoClose: 10,
                            align: align);
                    }
                }
                catch (Exception ex)
                {
                    AntdUI.Notification.error(window, $"Error", ex.Message, autoClose: 10, align: align);
                }
            }
            catch (Exception ex)
            {
                AntdUI.Notification.error(window, $"Error", ex.Message, autoClose: 10, align: align);
            }
        }

        private void btn_loadData_Click(object sender, EventArgs e)
        {
            string itemCode = tb_itemCode.Text.Trim();
            string maker = tb_maker.Text.Trim();
            try
            {
                _service.LoadData(itemCode, maker);
                DisplayData();
                AntdUI.Notification.success(window, $"Notice", "Lấy dữ liệu thành công!", autoClose: 10, align: align);
                
            }
            catch (Exception ex)
            {
                AntdUI.Notification.error(window, $"Error", ex.Message, autoClose: 10, align: align);
            }
        }
    }
}