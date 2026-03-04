using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Export_FPCA_OK2ship_Auto_System.Views
{
    public partial class SettingManual : Form
    {
        public string _LOCATION = "";
        public string _CATEGORY = "";
        public bool _TAKEALL  = false;
        public SettingManual(string[] list)
        {
            InitializeComponent();
            checkStatusBtn();
            comboBox1.Items.AddRange(list);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            checkStatusBtn();
        }

        private void comboBox1_StyleChanged(object sender, EventArgs e)
        {
            checkStatusBtn();
        }
        private void checkStatusBtn()
        {
            string text = comboBox1.Text;
            string combo = textBox1.Text;
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(combo))
            {
                button1.Enabled = false;

                return;
            }
            button1.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (button1.Text.Contains("Check") || button1.Text.Contains("Error"))
            {
                string location = textBox1.Text.Trim();
                if (File.Exists(location))
                {
                    button1.Text = "Save";
                    button1.BackColor = Color.Green;
                }
                else
                {
                    button1.Text = "Error";
                    button1.BackColor = Color.Red;

                }
            }
            if (button1.Text.Contains("Save"))
            {
                _LOCATION = textBox1.Text.Trim();
                _CATEGORY = comboBox1.Text.Trim();
                _TAKEALL = checkBox1.Checked;
                this.Hide();
            }

        }

        private void textBox1_DoubleClick(object sender, EventArgs e)
        {
            OpenFileDialog f_open = new OpenFileDialog();
            f_open.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (f_open.ShowDialog() == DialogResult.OK)
            { 
                textBox1.Text = f_open.FileName;
            }
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            checkStatusBtn();
        }
    }
}
