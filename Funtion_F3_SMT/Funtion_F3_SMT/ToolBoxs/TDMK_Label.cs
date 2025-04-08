using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.ToolBoxs
{
    public class TDMK_Label : Label
    {
        public TDMK_Label()
        {
            this.AutoSize = false;
            this.TextAlign = ContentAlignment.MiddleCenter;
            this.Text = "Logfile";
            this.Dock = DockStyle.Fill;
            this.Font = new Font("Arial", 12, FontStyle.Bold);
        }
    }
}
