using SEEV_SMT_F3.TDMK_Tool;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SEEV_SMT_F3
{

    public partial class FormHandleObject : Form
    {
        public FormHandleObject()
        {
            InitializeComponent();
            changeName("TDMK");
        }
        private void changeName(string name)
        {
            this.Text = name + "Process";
        }
        private void FormHandleObject_Load(object sender, EventArgs e)
        {
            TDMK_UI_Winform.OpenFormLarge(this);
        }

        private void FormHandleObject_FormClosed(object sender, FormClosedEventArgs e)
        {
            TDMK_UI_Winform.OpenAndCloseForm(this, new MainScreen());
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string str = toolStripComboBox1.SelectedItem.ToString();
                changeName(str);
            }
            catch (Exception ex)
            {
                changeName("TDMK");
            }
        }
    }
}
