using OK2SHIP_SMT.ToolBoxs;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.UserControls.Logins
{
    public partial class UC_Information : UserControl
    {
        public Dictionary<string, string> ListTextBox { get; set; } = new Dictionary<string, string>();
        public UC_Information(Dictionary<string, string> dic, List<Button> listButton)
        {
            InitializeComponent();
            ListTextBox = dic;
            settingProfile(listButton);
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string name = ((TextBox)sender).Name;
            string value = ((TextBox)sender).Text;
            ListTextBox[name] = value;
        }
        private void settingProfile(List<Button> listButton)
        {
            foreach (var button in listButton)
            {
                Button_Area.Controls.Add(button);
            }
            int i = 0;
            TableLayout.RowCount = ListTextBox.Keys.Count;
            TableLayout.RowStyles.Remove(TableLayout.RowStyles[0]);
            foreach (string item in ListTextBox.Keys)
            {
                TDMK_Label tDMK_Label = new TDMK_Label();
                tDMK_Label.Text = item;
                TextBox textBox = new TextBox();
                textBox.Name = item;
                textBox.Text = ListTextBox[item];
                textBox.Dock = DockStyle.Fill;
                textBox.Multiline = true;
                textBox.Leave += textBox1_TextChanged;
                TableLayout.Controls.Add(textBox, 1, i);
                TableLayout.Controls.Add(tDMK_Label, 0, i++);
                TableLayout.RowStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            }

        }
    }
}
