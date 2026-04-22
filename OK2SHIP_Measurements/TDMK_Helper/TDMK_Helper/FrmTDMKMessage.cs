using System;
using System.Drawing;
using System.Windows.Forms;
using TDMK_Helper;

namespace TDMK_Helper
{
    public class FrmTDMKMessage : Form
    {
        private Label lblMessage;
        private PictureBox picIcon;
        private Button btnOK;

        public bool AdminConfirmed { get; private set; }

        public FrmTDMKMessage(
            string message,
            string title,
            bool allowAdminLogin)
        {
            // ===== Form =====
            this.Text = title;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Font = new Font("Segoe UI", 9F);
            this.ClientSize = new Size(280, 160);
            this.KeyPreview = true;

            // ===== Icon =====
            picIcon = new PictureBox()
            {
                Image = SystemIcons.Warning.ToBitmap(),
                SizeMode = PictureBoxSizeMode.AutoSize,
                Left = 20,
                Top = 30
            };

            // ===== Message =====
            lblMessage = new Label()
            {
                Text = message,
                Left = 70,
                Top = 25,
                Width = 270,
                Height = 60,
                AutoSize = false
            };

            // ===== OK Button =====
            btnOK = new Button()
            {
                Text = "OK",
                Width = 90,
                Height = 30,
                Left = (this.ClientSize.Width - 90) / 2,
                Top = 105,
                DialogResult = DialogResult.OK
            };

            btnOK.Click += (s, e) => this.Close();

            this.Controls.AddRange(new Control[]
            {
            picIcon,
            lblMessage,
            btnOK
            });

            // ===== Hotkey =====
            this.KeyDown += (s, e) =>
            {
                if (allowAdminLogin && e.Alt && e.KeyCode == Keys.O)
                {
                    using (FrmLogin login = new FrmLogin())
                    {
                        if (login.ShowDialog(this) == DialogResult.OK)
                        {
                            AdminConfirmed = true;
                            this.Close();
                        }
                    }
                }
            };
        }
    }

}