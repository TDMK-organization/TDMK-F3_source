using OK2SHIP_SMT.Services;
using OK2SHIP_SMT.Views;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.UserControls
{
    public partial class UC_EnvironmentTable : UserControl
    {
        public string memory = "";
        public Dictionary<string, Dictionary<string, DataTable>> dictionary_Data = new Dictionary<string, Dictionary<string, DataTable>>();
        public CustomDataGridView psaList = new CustomDataGridView(new DataTable(), new Dictionary<string, string[]>());
        public CustomDataGridView linearList = new CustomDataGridView(new DataTable(), new Dictionary<string, string[]>());
        public UC_EnvironmentTable()
        {
            InitializeComponent();
            setupInit();
        }

        private void dataGridView_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dataGridView = (DataGridView)sender;
            object dataSource = dataGridView.DataSource;
            displayImage((DataTable)dataSource, e);
        }

        private void displayImage(DataTable sender, DataGridViewCellEventArgs e)
        {
            int col = e.ColumnIndex;
            int row = e.RowIndex;
            if (sender.Columns[col].DataType == typeof(Image))
            {
                Image imageData = (Image)sender.Rows[row][col];
                pictureBox.Image = imageData;
            }
            else if (sender.Columns[col].DataType == typeof(byte[]))
            {
                Image imageData = TDMK_ImageConverter.ByteArrayToImage((byte[])sender.Rows[row][col]);
                pictureBox.Image = imageData;
            }
            if (sender.Columns.Contains("Name") && listBox.SelectedItem.ToString().Contains("SAMPLE"))
            {
                DataRow rowz = sender.Rows[row];
                memory = $"{rowz["Name"]}-{rowz["Tape"]}-{row}";
            }
            else
            {
                memory = "";
            }
            btn_editImage.Text = $"Edit Image {memory}";


        }
        private void setupInit()
        {
            linearList.CellContentDoubleClick += dataGridView_CellContentDoubleClick;
            psaList.CellContentDoubleClick += dataGridView_CellContentDoubleClick;
            splitLinear.Panel2.Controls.Add(linearList);
            splitPSA.Panel2.Controls.Add(psaList);
        }

        private void listBox_SelectedIndexChanged(object sender, EventArgs e)
        {

            ListBox listBox = (ListBox)sender;
            if (listBox.SelectedIndex != -1) // Check if an item is selected
            {
                object selectedItem = listBox.SelectedItem;

                // If the items are strings:
                string selectedText = selectedItem.ToString();
                btn_editImage.Enabled = selectedText.Equals("SAMPLE");
                if (dictionary_Data.TryGetValue(selectedText, out Dictionary<string, DataTable> dic))
                {
                    string PSA_str = "PSA";
                    string Linear = "LINER";
                    if (dic.TryGetValue(PSA_str, out DataTable dt))
                    {
                        psaList.DataSource = dt;
                    }
                    if (dic.TryGetValue(Linear, out DataTable dtz))
                    {
                        linearList.DataSource = dtz;
                    }
                }

            }
        }

        private void btn_editImage_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(memory) && memory.Contains('-'))
            {
                string[] str = memory.Split('-');
                string name = str[0], tape = str[1], row = str[2];
                Image image = TDMK_ImageConverter.ByteArrayToImage((byte[])((DataTable)dictionary_Data["SAMPLE"][name]).Rows[int.Parse(row)]["Image"]);
                EditImageView edit = new EditImageView(image);
                Form dialog = new CommonForm("", edit, null);
                dialog.ShowDialog();
                if(!edit.save_status)
                {
                    return;
                }
                ((DataTable)dictionary_Data["SAMPLE"][name]).Rows[int.Parse(row)]["Image"] = TDMK_ImageConverter.ImageToByteArray(edit.image, ImageFormat.Jpeg);
                pictureBox.Image = edit.image;
            }
        }
    }
}
