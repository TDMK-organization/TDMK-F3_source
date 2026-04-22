using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TDMK_Helper
{
    public partial class FrmLogin : Form
    {
        TextBox txtUser;
        TextBox txtPass;

        public FrmLogin()
        {
            // ===== Form =====
            this.Text = "Leader Login";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Font = new Font("Segoe UI", 9F);
            this.ClientSize = new Size(320, 170);

            // ===== Labels =====
            Label lblUser = new Label()
            {
                Text = "Tài khoản",
                AutoSize = true,
                Left = 20,
                Top = 28
            };

            Label lblPass = new Label()
            {
                Text = "Mật khẩu",
                AutoSize = true,
                Left = 20,
                Top = 63
            };

            // ===== TextBoxes =====
            txtUser = new TextBox()
            {
                Left = 100,
                Top = 25,
                Width = 180
            };

            txtPass = new TextBox()
            {
                Left = 100,
                Top = 60,
                Width = 180,
                UseSystemPasswordChar = true
            };

            // ===== Buttons =====
            Button btnLogin = new Button()
            {
                Text = "Đăng nhập",
                Width = 90,
                Height = 30,
                Left = 55,
                Top = 105
            };

            Button btnCancel = new Button()
            {
                Text = "Hủy",
                Width = 90,
                Height = 30,
                Left = 160,
                Top = 105
            };

            btnLogin.Click += BtnLogin_Click;
            btnCancel.Click += (s, e) => this.Close();

            this.Controls.AddRange(new Control[]
            {
            lblUser, txtUser,
            lblPass, txtPass,
            btnLogin, btnCancel
            });
        }

        private void BtnLogin_Click(object sender, EventArgs e)
        {
            if (txtUser.Text == "canhnd00850" && txtPass.Text == "@dulieung!xacnhan!")
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                MessageBox.Show(
                    this,
                    "Sai tài khoản hoặc mật khẩu",
                    "Login",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }
    }
}
