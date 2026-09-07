using System;
using System.Windows.Forms;

namespace OK2SHIP_SMT.Views
{
    public partial class BendingForm : Form
    {
        public bool? PRIME = null;
        public BendingForm()
        {
            InitializeComponent();
        }

    

        private void btn_new_Click_1(object sender, EventArgs e)
        {
            PRIME = true;
            this.Close();
        }

        private void btn_oldversion_Click(object sender, EventArgs e)
        {
            PRIME = false;
            this.Close();
        }
    }
}