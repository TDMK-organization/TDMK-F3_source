using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT
{
    public partial class Message_Error : Form
    {
        public string message = "";

        public string message_
        {
            get { return message_; }
            set { message_ = value; }
        }

        public Message_Error()
        {
            InitializeComponent();
        }
 
        private void Message_Error_Load(object sender, EventArgs e)
        {
            lbl_msg.Text = message_;
        }
    }
}
