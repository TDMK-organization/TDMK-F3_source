using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TDMK_Helper
{
    internal class Frm_Loading : Form
    {
        public Frm_Loading()
        {
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Name = "Frm_Loading";
            this.ResumeLayout(false);
            this.ControlBox = false; // Ẩn nút X
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Vui lòng chờ...";
            this.Width = 300;
            this.Height = 100;
            Label lbl = new Label();
            lbl.Text = "Vui lòng chờ hệ thống xử lí....";
            lbl.AutoSize = false;
            lbl.TextAlign = ContentAlignment.MiddleCenter;
            lbl.Dock = DockStyle.Fill;
            this.Controls.Add(lbl);
        }
    }
}
