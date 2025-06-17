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
        public void ClearData()
        {
            DataTable dt = new DataTable();
            psaList.DataSource = dt;
            linearList.DataSource = dt;
            dictionary_Data = new Dictionary<string, Dictionary<string, DataTable>>();
            dataGridView.DataSource = dt;
            pictureBox.Image = null;
            listBox.Items.Clear();


        }
        public void FillData(DataTable dataTable = null, Dictionary<string, Dictionary<string, DataTable>> dictionary = null)
        {
            ClearData();
            dataGridView.RowTemplate.Height = 150;

            if (dataTable != null)
            {
                dataGridView.DataSource = dataTable;
            }
            if (dictionary != null)
            {
                dictionary_Data = dictionary;
                listBox.Items.AddRange(dictionary.Keys.ToArray());
            }
            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                if (dataGridView.Columns[column.Name] is DataGridViewImageColumn)
                {
                    ((DataGridViewImageColumn)dataGridView.Columns[column.Name]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)dataGridView.Columns[column.Name]).Width = 250; // Đặt chiều rộng cột
                }
            }
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
            if (sender.Columns.Contains("Image Sample"))
            {
                DataRow rowz = sender.Rows[row];
                memory = $"{rowz["Name"]}-{rowz["Tape"]}-{row}";
                btn_editImage.Enabled = true;
            }
            else
            {
                btn_editImage.Enabled = false;
                memory = "";
            }
            btn_editImage.Text = $"Edit Image {memory}";


        }
        private void setupInit()
        {
            psaList.CellPainting += dataGridView_CellPainting;
            linearList.CellPainting += dataGridView_CellPainting;
            //dataGridView.CellContentDoubleClick += dataGridView_CellContentDoubleClick;
            linearList.CellContentDoubleClick += dataGridView_CellContentDoubleClick;
            psaList.CellContentDoubleClick += dataGridView_CellContentDoubleClick;
            splitLinear.Panel2.Controls.Add(linearList);
            splitPSA.Panel2.Controls.Add(psaList);
        }

        private void listBox_SelectedIndexChanged(object sender, EventArgs e)
        {

            Dictionary<string, string[]> keyValuePairs = new Dictionary<string, string[]>();
            keyValuePairs.Add("Judgement failure mode", new[] { "No adhesive stick on liner" });
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
                        psaList.columnDropdowns = keyValuePairs;
                        psaList.DataSource = dt;
                    }
                    if (dic.TryGetValue(Linear, out DataTable dtz))
                    {
                        linearList.columnDropdowns = keyValuePairs;
                        linearList.DataSource = dtz;
                    }
                }

            }
        }
        Form dialog = new Form();
        private void btn_close_Click(object sender, EventArgs e)
        {
            dialog.Hide();
        }

        private void btn_editImage_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(memory) && memory.Contains('-'))
            {
                string[] str = memory.Split('-');
                string name = str[0], tape = str[1], row = str[2];
                var image = ((DataTable)dataGridView.DataSource).Rows[int.Parse(row)]["Image Sample"];
                bool prime = false;
                if (image is byte[])
                {
                    prime = true;
                    image = TDMK_ImageConverter.ByteArrayToImage((byte[])image);
                }
                EditImageView edit = new EditImageView((Image)image, btn_close_Click);
                dialog = new CommonForm("", edit, null);
                dialog.ShowDialog();
                if (!edit.save_status)
                {
                    return;
                }
                if (prime)
                {
                    image = TDMK_ImageConverter.ImageToByteArray(edit.image, ImageFormat.Jpeg);
                }
                else
                {
                    image = edit.image;
                }
                ((DataTable)dataGridView.DataSource).Rows[int.Parse(row)]["Image Sample"] = image;
                pictureBox.Image = edit.image;
            }
        }

        private void dataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void splitPSA_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dataGridView_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.Value != null)
                if (e.Value.ToString().Contains("Fail"))
                {
                    e.CellStyle.BackColor = Color.Red;
                }
        }

        private void dataGridView_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            //Debugger.Break();
        }
    }
}
