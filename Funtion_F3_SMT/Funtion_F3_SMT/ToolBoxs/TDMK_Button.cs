
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.ToolBoxs
{
    public partial class TDMK_Button : Button
    {
        public TDMK_Button()
        {
            InitializeComponent();
            Init();
        }
        private void Init()
        {
            this.Font = new Font(new FontFamily("Microsoft Sans Serif"), (float)14.25, FontStyle.Bold);
            this.Text = "TDMK BUTTON";
            this.SizeChanged += TDMK_Button_SizeChanged;
            this.Dock = DockStyle.Fill;
        }

        private void TDMK_Button_SizeChanged(object sender, EventArgs e)
        {

        }
    }
}
