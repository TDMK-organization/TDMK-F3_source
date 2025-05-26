using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SQL;
using OK2SHIP_Measurements;
using OK2SHIP_Lib;

namespace OK2SHIP_Software
{
    public partial class Edit_ItemsText : Form
    {
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        TDMK_OK2SHIP myCode = new TDMK_OK2SHIP();
        public string sel_tbl;
        public string Sel_Depart;
        string filtr_str;
        public AutoCompleteStringCollection LAB_Printing_IgnoredList = new AutoCompleteStringCollection() { "LPI_Open_B1", "LPI_Open_B2" };
        public AutoCompleteStringCollection Etching_DESData_list = new AutoCompleteStringCollection() { "DES_DATA_4", "DES_DATA_7", "DES_DATA_11", "DES_DATA_14", "DES_DATA_18", "DES_DATA_21", "DES_DATA_25", "DES_DATA_28" };
        public Edit_ItemsText()
        {
            InitializeComponent();
        }
        public void Set_target_Table(RadioButton src_RB)
        {
            if (src_RB.Checked)
            {
                sel_tbl = src_RB.Text;
                DGV_Data.DataSource = null;
            }
        }

        private void RB_Lab_CheckedChanged(object sender, EventArgs e)
        {
            if (RB_Lab.Checked)
            {
                Sel_Depart = RB_Lab.Text;
                GBIPQC_Items.Enabled = true;
                foreach (RadioButton c in GBIPQC_Items.Controls)
                {
                    if (c.Checked)
                    {
                        Set_target_Table(c);
                        break;
                    }
                }
            }
        }
        private void RB_Etching_CheckedChanged(object sender, EventArgs e)
        {
            if (RB_Product.Checked)
            {
                Sel_Depart = RB_Product.Text;
                GBIPQC_Items.Enabled = true;
                foreach (RadioButton c in GBIPQC_Items.Controls)
                {
                    if (c.Checked)
                    {
                        Set_target_Table(c);
                        break;
                    }
                }
            }
        }
        private void RB_UV_Process_CheckedChanged(object sender, EventArgs e)
        {
            if (RB_UV_Process.Checked)
            {
                Sel_Depart = RB_UV_Process.Text;
                GBIPQC_Items.Enabled = false;
            }
            Set_target_Table(RB_UV_Process);
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            filtr_str = "";
            switch (Sel_Depart)
            {
                case "LAB":
                    switch (sel_tbl)
                    {
                        case "Etching_Process":
                            foreach (string t in Etching_DESData_list)
                            {
                                filtr_str = filtr_str + "'" + t + "'" + ",";
                            }
                            filtr_str = "Col_Name not in (" + filtr_str.TrimEnd(',') + ")" + " and Process_Name ='" + sel_tbl + "'";
                            break;
                        case "Printing_Process":
                            foreach (string t in LAB_Printing_IgnoredList)
                            {
                                filtr_str = filtr_str + "'" + t + "'" + ",";
                            }
                            filtr_str = "Col_Name not in (" + filtr_str.TrimEnd(',') + ")" + " and Process_Name ='" + sel_tbl + "'";
                            break;
                        default:
                            filtr_str = "Process_Name = '" + sel_tbl + "'";
                            break;
                    }

                    break;
                case "Production":
                    switch (sel_tbl)
                    {
                        case "Etching_Process":
                            foreach (string t in Etching_DESData_list)
                            {
                                filtr_str = filtr_str + "'" + t + "'" + ",";
                            }
                            filtr_str = "Col_Name in (" + filtr_str.TrimEnd(',') + ")" + " and Process_Name ='" + sel_tbl + "'";
                            break;
                        case "Printing_Process":
                            foreach (string t in LAB_Printing_IgnoredList)
                            {
                                filtr_str = filtr_str + "'" + t + "'" + ",";
                            }
                            filtr_str = "Col_Name in (" + filtr_str.TrimEnd(',') + ")" + " and Process_Name ='" + sel_tbl + "'";
                            break;
                        default:
                            filtr_str = ""; 
                            break;
                    }
                    break;
                case "UV_Process":
                    filtr_str = "Process_Name = '" + sel_tbl + "'";
                    break;
            }
            if(filtr_str!="")
            {
                DataTable temp = TDMK_Code.Datatable_Filter(myVar.sqlcon_IPQC, "All_Items", filtr_str);
                DGV_Data.DataSource = temp;
                List<string> edit_col = new List<string>() { "Item_Name", "Internal_Point", "LogFile_Point" };
                foreach(DataGridViewColumn dc in DGV_Data.Columns)
                {
                    if(edit_col.IndexOf(dc.Name)==-1)
                    {
                        dc.ReadOnly = true;
                    }
                }
                DGV_Data.AutoResizeColumns();
                myCode.Disable_Sort_DGV(DGV_Data);
            }
        }

        private void RB_Etching_Process_CheckedChanged(object sender, EventArgs e)
        {
            Set_target_Table(RB_Etching_Process);
        }

        private void RB_CopperPlating_CheckedChanged(object sender, EventArgs e)
        {
            Set_target_Table(RB_CopperPlating);
        }

        private void RB_Printing_Process_CheckedChanged(object sender, EventArgs e)
        {
            Set_target_Table(RB_Printing_Process);
        }

        private void RB_CoverLay_Process_CheckedChanged(object sender, EventArgs e)
        {
            Set_target_Table(RB_CoverLay_Process);
        }

        private void Edit_ItemsText_Load(object sender, EventArgs e)
        {
            RB_Lab.Checked = true;
            RB_Etching_Process.Checked = true;
        }
        private void DGV_Data_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                int r_num = 0;
                r_num = DGV_Data.SelectedRows.Count;
                if (r_num > 0)
                {
                    if (MessageBox.Show("Do you want to update Item_Name ?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        for (int i = 0; i < r_num; i++)
                        {
                            int r_index = DGV_Data.SelectedRows[i].Index;
                            string ID = DGV_Data.Rows[r_index].Cells["ID"].Value.ToString();
                            string item_val = DGV_Data.Rows[r_index].Cells["Item_Name"].Value.ToString();
                            TDMK_Code.update_item_val_filter("All_Items", myVar. sqlcon_IPQC, "ID = " + ID, "Item_Name", item_val);
                        }
                        DGV_Data.DataSource = TDMK_Code.Datatable_Filter(myVar. sqlcon_IPQC, "All_Items", "Process_Name ='" + sel_tbl + "'");
                    }
                }
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
            myVar.frm_mainscreen.Show();
        }

        private void DGV_Data_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if(MessageBox.Show("Do you want to update?","Warning",MessageBoxButtons.YesNo)==DialogResult.Yes)
            {
                try
                {
                    TDMK_Code.Delelte_FilteredItem_arr("All_Items", myVar.sqlcon_IPQC, filtr_str);
                    myCode.BatchBulkCopy(myVar.sqlcon_IPQC, (DataTable)DGV_Data.DataSource, "All_Items");
                }
                catch
                {
                    MessageBox.Show("Update false");
                }
                btnLoad.PerformClick();
            }
        }
    }
}
