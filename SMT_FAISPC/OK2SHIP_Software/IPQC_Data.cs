using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
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
    public partial class IPQC_Data : Form
    {

        myVar myCode2 = new myVar();
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public TDMK_OK2SHIP myCode = new TDMK_OK2SHIP();
        public string app_path;
        public string origin_cell_val;
        public bool edit_en = false;
        public bool pro_en = false;
        public delegate void Data_process(string _filename, DataGridView tar_DGV, DataGridView DGV_Spec, DataGridView tar_DGV_plus, bool edit_mode);
        public delegate void SetText(Button tar_btn, string btn_str);
        public delegate void checkTextbox(TextBox src_txt);
        public Color curr_color;
        public SqlConnection sqlcon { get; set; }
        public string sel_tbl;
        public DataTable sel_dt = new DataTable();
        public DataTable cur_sel_dt = new DataTable();
        public bool update_en = false;
        public string ItemCode { get; set; }
        public string LotNo { get; set; }

        public IPQC_Data()
        {
            InitializeComponent();
        }
        public IPQC_Data(string _itemcode, string _lotno, SqlConnection _sqlcon)
        {
            InitializeComponent();
            ItemCode = _itemcode;
            LotNo = _lotno;
            sqlcon = _sqlcon;
        }

        public void Set_target_Table(RadioButton src_RB)
        {
            if (src_RB.Checked)
            {
                sel_tbl = src_RB.Text;
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

        private void RB_UV_Process_CheckedChanged(object sender, EventArgs e)
        {
            Set_target_Table(RB_UV_Process);
        }

        private void RB_CoverLay_Process_CheckedChanged(object sender, EventArgs e)
        {
            Set_target_Table(RB_CoverLay_Process);
        }

        private void btnLoadSpec_Click(object sender, EventArgs e)
        {
            string flt_str = "";
            DataTable tbl_all = new DataTable();
            switch(sel_tbl)
            {
                case "Roughness":
                    flt_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Col_Name" }, new string[] { ItemCode, "%Roughness%" });
                    tbl_all = TDMK_Code.Datatable_Filter(sqlcon, "All_Items", "Col_Name like '%Roughness%'");
                    break;
                case "AU_NI":
                    flt_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Col_Name" }, new string[] { ItemCode, "%THICKNESS%" });
                    tbl_all = TDMK_Code.Datatable_Filter(sqlcon, "All_Items", "Col_Name like '%THICKNESS%'");
                    break;
                default:
                    flt_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, sel_tbl });
                    tbl_all = TDMK_Code.Datatable_Filter(sqlcon, "All_Items", "Process_Name = '" + sel_tbl + "'");
                    break;
            }
            sel_dt = TDMK_Code.Datatable_Filter(sqlcon, "SpecList", flt_str);                       
            DGV_SpecView.Columns.Clear();
            DGV_SpecView.DataSource = null;
            if (sel_dt.Rows.Count > 0)
            {         
                string[] Col_name = sel_dt.AsEnumerable().Select(r => r.Field<string>("Col_name")).ToArray();
                string[] col_header = new string[Col_name.Length];
                int inx = 0;
                foreach (string c_name in Col_name)
                {
                    string[] item_name = tbl_all.AsEnumerable().Where(r => r.Field<string>("Col_name") == c_name).Select(r => r.Field<string>("Item_Name")).ToArray();
                    col_header[inx] = item_name[0];
                    if (!myCode.check_columns_existed(myCode.DGV_To_Table(DGV_SpecView), c_name))
                    {
                        DGV_SpecView.Columns.Add(c_name, item_name[0]);
                    }
                    inx++;
                }
                myCode.IPQC_Spec_Process_Man3(sel_dt, DGV_SpecView);
                DataTable sel_tbl_data = TDMK_Code.Datatable_Filter(sqlcon, sel_tbl, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                DGV_DataView.DataSource = sel_tbl_data.AsDataView().ToTable(false, Col_name);
                inx = 0;
                foreach(DataGridViewColumn c in DGV_DataView.Columns)
                {
                    c.HeaderText = col_header[inx];
                    inx++;
                }
                inx = 0;
                foreach(DataGridViewRow r in DGV_DataView.Rows)
                {
                    myCode.Check_IPQC_Data_inSpec(DGV_DataView, DGV_SpecView);
                    r.HeaderCell.Value = (inx + 1).ToString();
                    inx++;
                }
                DGV_DataView.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                DGV_DataView.AutoResizeColumns();
                myCode.Disable_Sort_DGV(DGV_DataView);
            }
        }

        private void IPQC_Data_Load(object sender, EventArgs e)
        {
            lblSpec.Text = "Specification: " + ItemCode + "-" + LotNo;
            lblIPQC.Text = "IPQC Data: " + ItemCode + "-" + LotNo;
        }

        private void DGV_DataView_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                List<string> col_list = new List<string>();
                List<string> header_list = new List<string>();
                foreach (DataGridViewCell c in DGV_DataView.SelectedCells)
                {
                    int k = c.ColumnIndex;
                    if (!col_list.Contains(DGV_DataView.Columns[k].Name))
                    {
                        col_list.Add(DGV_DataView.Columns[k].Name);
                        header_list.Add(DGV_DataView.Columns[k].HeaderText);
                    }
                }
                int arr_num = col_list.Count;
                TDMK_OK2SHIP.IPQC_Spec[] IPQC_spec_lst = new TDMK_OK2SHIP.IPQC_Spec[arr_num];
                int col_list_inx = 0;
                foreach (string t in col_list)
                {
                    TDMK_OK2SHIP.IPQC_Spec sel_IPQC_test = new TDMK_OK2SHIP.IPQC_Spec();
                    sel_IPQC_test.IPQC_Name = header_list[col_list_inx];
                    sel_IPQC_test.UL = myCode.checkDBNull(DGV_SpecView.Rows[1].Cells[t].Value);
                    sel_IPQC_test.LL = myCode.checkDBNull(DGV_SpecView.Rows[2].Cells[t].Value);
                    IPQC_spec_lst[col_list_inx] = sel_IPQC_test;
                    double[] IPQC_Data = new double[DGV_DataView.Rows.Count];
                    int inx = 0;
                    foreach(DataGridViewRow r in DGV_DataView.Rows)
                    {
                        string cur_val = myCode.checkDBNull(r.Cells[t].Value);
                        if (myCode.IsNumeric(cur_val))
                        {
                            IPQC_Data[inx] = Convert.ToDouble(cur_val);
                            inx++;
                        }
                    }
                    Array.Resize(ref IPQC_Data, inx);
                    if (inx > 0)
                    {
                        IPQC_Chart frmIPQC_chart = new IPQC_Chart(IPQC_Data, sel_IPQC_test);
                        frmIPQC_chart.Show();
                    }
                    col_list_inx++;
                }
            }
            catch
            {

            }
        }

        private void RBRoughness_CheckedChanged(object sender, EventArgs e)
        {
            if (RBRoughness.Checked)
            {
                sel_tbl = "Roughness";
            }
        }

        private void RBAU_NI_CheckedChanged(object sender, EventArgs e)
        {
            if (RBAU_NI.Checked)
            {
                sel_tbl = "AU_NI";
            }
        }
    }
}
