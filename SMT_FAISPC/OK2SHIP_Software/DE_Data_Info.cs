using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using TDMK_SQL;
using OK2SHIP_Measurements;

namespace OK2SHIP_Software
{
    public partial class DE_Data_Info : Form
    {
        //TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        myVar myCode = new myVar();
        public SqlConnection sel_sqlcon;
        public string sel_tbl;
        public DE_Data_Info()
        {
            InitializeComponent();
        }

        private void DE_Data_Info_Load(object sender, EventArgs e)
        {
            RBDeclare.Checked = true;
            Disable_Sort_DGV(DGV_Data);
        }
        private void RBDeclare_CheckedChanged(object sender, EventArgs e)
        {
            if (RBDeclare.Checked)
            {
                sel_sqlcon = myVar.sqlcon_Declare;
                sel_tbl = "DE_Data";
                lstTable.DataSource = null;
                lstTable.DataSource = new List<string>() { sel_tbl };
            }
        }
        private void RBMaterials_CheckedChanged(object sender, EventArgs e)
        {
            if (RBMaterials.Checked)
            {
                sel_sqlcon = myVar.sqlcon_Materials;
                lstTable.DataSource = null;
                lstTable.DataSource = myCode.GetAllTables(sel_sqlcon).ToList();
            }
        }
        private void RBRecycle_CheckedChanged(object sender, EventArgs e)
        {
            if (RBRecycle.Checked)
            {
                sel_sqlcon = myVar.sqlcon_Recycle;
                lstTable.DataSource = null;
                lstTable.DataSource = myCode.GetAllTables(sel_sqlcon).ToList();
                btnUpdate.Enabled = true;
                //txtMasterList.Enabled = false;
            }
            else
            {
                btnUpdate.Enabled = false;
                //txtMasterList.Enabled = false;
            }
        }
        private void btnLoad_Click(object sender, EventArgs e)
        {
            string cur_tbl = lstTable.SelectedItem.ToString();
            if (txtItemCode.Text == "")
            {
                myVar.TDMK_Code.fill_dataset_DGV(cur_tbl, myVar.TDMK_Code.SQL_CMD(cur_tbl), DGV_Data, sel_sqlcon);
            }
            else
            {
                myVar.TDMK_Code.fill_dataset_DGV(cur_tbl, myVar.TDMK_Code.filter_cmd(cur_tbl, "ItemCode = '" + txtItemCode.Text + "'", "ID"), DGV_Data, sel_sqlcon);
            }
            if (DGV_Data.Rows.Count == 0)
            {
                MessageBox.Show("No Data");
            }
        }

        private void DGV_Data_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            mnuAction.Show(DGV_Data, new Point(e.X, e.Y));
        }
        public void Disable_Sort_DGV(DataGridView sel_DGV)
        {
            foreach (DataGridViewColumn column in sel_DGV.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }
        private void DGV_Data_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                DataGridViewCell curr_dgv_cell = DGV_Data.CurrentCell;
                int curr_col_inx = curr_dgv_cell.ColumnIndex;
                if (DGV_Data.Columns[curr_col_inx].Name == "ID")
                {
                    MessageBox.Show("Cannot edit this field");                    
                }
                else
                {
                    DGV_Data.BeginEdit(true);
                }
            }
        }
        private void mnuAddNew_Click(object sender, EventArgs e)
        {
            if (txtItemCode .Text !="")
            {
                string ID = (myVar.TDMK_Code.SQL_MAX(sel_tbl, "ID", sel_sqlcon) + 1).ToString();
                myVar.TDMK_Code.insert_val_arr(sel_tbl, sel_sqlcon, new string[] { "ID", "ItemCode" }, new string[] { ID, txtItemCode.Text });
                btnLoad.PerformClick();
            }
            else
            {
                MessageBox.Show("Please, enter ItemCode number");
            }
        }
        private void mnuDelete_Click(object sender, EventArgs e)
        {
            if (DGV_Data .Rows .Count >0)
            {
                int sel_row_inx = DGV_Data.SelectedCells[0].RowIndex;
                string ID = DGV_Data.Rows[sel_row_inx].Cells["ID"].Value.ToString(); 
                myVar.TDMK_Code.Delelte_FilteredItem_arr(sel_tbl, sel_sqlcon, "ID = " + ID);
                btnLoad.PerformClick();
            }
        }
        private void mnuUpdate_Click(object sender, EventArgs e)
        {
            string[] items = new string[DGV_Data.Columns.Count-1];
            string[] item_vals = new string[DGV_Data.Columns.Count-1];
            int sel_row_inx = DGV_Data.SelectedCells[0].RowIndex;
            int ID = Convert.ToInt32(DGV_Data.Rows[sel_row_inx].Cells[0].Value);
            for (int i = 1; i< DGV_Data .ColumnCount; i++)
            {
                items[i-1] = DGV_Data.Columns[i].Name;
                item_vals[i-1] = DGV_Data.Rows[sel_row_inx].Cells[i].Value.ToString();
            }          
            myVar.TDMK_Code.updatebyID_val_arr(sel_tbl, sel_sqlcon, ID, items, item_vals);
            btnLoad.PerformClick();
        }

        private void lstTable_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstTable .SelectedIndex >-1)
            {
                sel_tbl = lstTable.SelectedItem.ToString();
                btnLoad.PerformClick();
            }
        }
        private void btnExit_Click(object sender, EventArgs e)
        {
            myVar.frm_mainscreen.Show();
            this.Close();
        }

        private void lstTable_RightToLeftChanged(object sender, EventArgs e)
        {

        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            OpenFileDialog f_dialog = new OpenFileDialog();
            if (f_dialog.ShowDialog()==DialogResult.OK)
            {
                string f_name = f_dialog.FileName;
                txtMasterList.Text = f_name;
                if (myCode.Recycle_data_process(f_name))
                {
                    MessageBox.Show("Completed Update from Master List", "Warning");
                }
                else
                {
                    MessageBox.Show("Error", "Warning");
                }
            }
        }
    }
}
