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

namespace OK2SHIP_SMT.UserControls
{
    public partial class UC_XrayPicture : UserControl
    {
        public CustomDataGridView dataGridView = new CustomDataGridView(new DataTable(), new Dictionary<string, string[]>()) { Dock = DockStyle.Fill };
        public UC_XrayPicture()
        {
            InitializeComponent();
            tbl_dta.Controls.Add(dataGridView, 0, 1);
        }
        public void setData(DataTable dataTable, Dictionary<string, string[]> dic = null)
        {
            if (dic == null)
            {
                dic = new Dictionary<string, string[]>();
            }
            dic.Add("Result", new[] { "OK" });
            dataGridView.columnDropdowns = dic;
            dataGridView.DataSource = dataTable;
            dataGridView.CellContentDoubleClick += handleClickImage;
            tbl_dta.Controls.Add(dataGridView, 0, 1);
        }
        public void handleClickImage(object sender, DataGridViewCellEventArgs e)
        {

            var data = ((DataTable)dataGridView.DataSource).Rows[e.RowIndex][e.ColumnIndex];
            if (data is Image)
            {
                pictureBox.Image = (Image)data;
            }
            else if (data is byte[])
            {
                pictureBox.Image = TDMK_ImageConverter.ByteArrayToImage((byte[])data);

            }
        }

        public void ClearData()
        {
            pictureBox.Image = null;
            dataGridView.DataSource = null;
        }
    }
}
