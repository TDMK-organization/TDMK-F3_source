using System;
using System.Windows.Forms;

namespace OK2SHIP_SMT.Views
{
    public partial class BendingForm : Form
    {
        public BendingForm()
        {
            InitializeComponent();
        }

    

        private void btn_new_Click_1(object sender, EventArgs e)
        {
            this.Close();
            new BendingWindow().Show();
        }
    }
}