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

namespace OK2SHIP_SMT.Views
{
    public partial class CommentItemCodeAirBubble : Form
    {
        AirBubbleService _service = null;
        public CommentItemCodeAirBubble(AirBubbleService service = null)
        {
            if (service == null)
            {
                _service = new AirBubbleService();
            }
            InitializeComponent();
        }

        private void btn_load_Click(object sender, EventArgs e)
        {
            try
            {
                KeyValuePair<string, string> pair = _service.LoadComment(tb_itemCode.Text, tb_Tape.Text);
                textBox3.Text = pair.Key;
                textBox4.Text = pair.Value;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Comment airbubble!");
            }
        }

        private void btn_save_Click(object sender, EventArgs e)
        {
            try
            {
                _service.SaveComment(tb_itemCode.Text, tb_Tape.Text, textBox3.Text, textBox4.Text);
            }
            catch (Exception ex)
            {
                try
                {
                    if (ex.Message.Split('-').Count() > 1 && ex.Message.Split('-')[0].Contains("1234") && MessageBox.Show($"{ex.Message.Split('-')[1]}", "Comment Airbbule!", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        _service.SaveComment(tb_itemCode.Text, tb_Tape.Text, textBox3.Text, textBox4.Text, true);
                    }
                    else
                    {
                        throw ex;
                    }
                }
                catch (Exception ex2)
                {

                    MessageBox.Show(ex2.Message, "Comment airbubble");
                }
            }

        }

        private void btn_delete_Click(object sender, EventArgs e)
        {
            string itemCode = tb_itemCode.Text;
            string tape = tb_Tape.Text;
            _service.Remove(itemCode, tape);
            MessageBox.Show($"Remove compplete");
        }
    }
}
