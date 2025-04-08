using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.UserControls
{
    public partial class PickData : UserControl
    {
        private Action closeAction;
        public PickData(Action close)
        {
            InitializeComponent();
            closeAction = close;
        }
        public string ShouldClose()
        {
            return tb_Value.Text;
        }


        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            DateTime selectedDate = monthCalendar1.SelectionStart;
            tb_Value.Text = selectedDate.ToString("dd-MMM");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            closeAction();
        }
    }
}
