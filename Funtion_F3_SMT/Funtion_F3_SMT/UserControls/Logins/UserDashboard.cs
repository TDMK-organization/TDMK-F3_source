using OK2SHIP_SMT.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
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
            foreach(TabControl item in tabControl.TabPages)
            {
                item.Visible = true;
            }
            string role = UserSession.Instance.Role;
            switch (role) {
                case "ADMIN":
                    break;
                
            }

        }
    }
}
