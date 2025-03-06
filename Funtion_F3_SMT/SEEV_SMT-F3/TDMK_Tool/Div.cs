using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace SEEV_SMT_F3.TDMK_Tool
{
    public class Div : UserControl
    {
        #region Constructor
        private void Constructor()
        {
            this.Dock = DockStyle.Fill;
        }

        private void InitializeComponent()
        {
            SuspendLayout();
            // 
            // Div
            // 
            BackColor = Color.Gray;
            Name = "Div";
            Size = new Size(1245, 662);
            ResumeLayout(false);

        }

        public Div()
        {
            Constructor();
        }
        #endregion

    }
}
