using Echeck_LogFile_Process;
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
using System.Data.SqlClient;
using TDMK_SEEV_DLL;
using System.IO;
using VHX;

namespace OK2SHIP
{
    public partial class ReCheck_Data : Form
    {
        public SqlConnection sqlcon { get; set; }
        public string ItemCode { get; set; }
        public string LotNo { get; set; }
        public string Cycle { get; set; }
        public string _Process { get; set; }
        public DataTable Cycle_tbl { get; set; }
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        ECheck_Process proc_data = new ECheck_Process();
        SEI_Lib myCode = new SEI_Lib();
        myVar myVar_code = new myVar();
        DataTable spec_dt = new DataTable();
        Dictionary<string, DataTable> result_table_lst = new Dictionary<string, DataTable>();
        Dictionary<string, List<int>> update_lst = new Dictionary<string, List<int>>();
        public ReCheck_Data()
        {
            InitializeComponent();
        }
        public ReCheck_Data(SqlConnection _sqlcon, string itemcode, string lotno, string cycle, string process, DataTable cycle_dt)
        {
            InitializeComponent();
            sqlcon = _sqlcon;
            ItemCode = itemcode;
            LotNo = lotno;
            Cycle = cycle;
            _Process = process;
            Cycle_tbl = cycle_dt;
            cbMachine.SelectedIndex = 0;
        }

        private void ReCheck_Data_Load(object sender, EventArgs e)
        {
            DataTable temp_spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, _Process }));
            spec_dt = temp_spec_dt.AsDataView().ToTable(false, new string[] { "Point+V", "Point-V", "LSL", "USL" });
            Dictionary<string, List<int>> dic_NG = check_in_spec_tbl(spec_dt, Cycle_tbl);
            DataTable NG_tbl = Cycle_tbl.AsDataView().ToTable(false, dic_NG.Keys.ToArray());
            DGV_Details.DataSource = NG_tbl;
            foreach (var NG_lst in dic_NG)
            {
                string col_name = NG_lst.Key;
                foreach(var r_inx in NG_lst.Value)
                {
                    DGV_Details.Rows[r_inx].Cells[col_name].Style.BackColor = Color.Pink;
                }
            }
            myCode.DGV_Auto_Resize(DGV_Details);    
        }
        public Dictionary<string, List<int>> check_in_spec_tbl(DataTable dgv_spec, DataGridView dgv_data)
        {
            Dictionary<string, List<int>> result = new Dictionary<string, List<int>>();
            int r_inx = 0;
            foreach (DataGridViewRow dr in dgv_data.Rows)
            {
                string USL = dgv_spec.Rows[r_inx][3].ToString();
                string LSL = dgv_spec.Rows[r_inx][2].ToString();
                foreach (DataGridViewColumn dc in dgv_data.Columns)
                {
                    DataGridViewCell sel_cell = dr.Cells[dc.Index];
                    sel_cell.Style.BackColor = myCode.check_in_limit2(USL, LSL, sel_cell.Value.ToString());
                    if (!result.ContainsKey(dc.Name))
                    {
                        result.Add(dc.Name, new List<int>() { r_inx });
                    }
                    else
                    {
                        result[dc.Name].Add(r_inx);
                    }
                }
                r_inx++;
            }
            return result;
        }
        public Dictionary<string, List<int>> check_in_spec_tbl(DataTable dgv_spec, DataTable dgv_data)
        {
            Dictionary<string, List<int>> result = new Dictionary<string, List<int>>();
            int r_inx = 0;
            foreach (DataRow dr in dgv_data.Rows)
            {
                string USL = dgv_spec.Rows[r_inx][3].ToString();
                string LSL = dgv_spec.Rows[r_inx][2].ToString();
                foreach (DataColumn dc in dgv_data.Columns)
                {
                    string act_val = myCode.checkDBNull(dr[dc]);
                    if(myCode.check_in_limit2(USL, LSL, act_val) !=Color.White)
                    {
                        if (!result.ContainsKey(dc.ColumnName))
                        {
                            result.Add(dc.ColumnName, new List<int>() { r_inx });
                        }
                        else
                        {
                            result[dc.ColumnName].Add(r_inx);
                        }
                    }
                }
                r_inx++;
            }
            return result;
        }

        private void txtLogfile_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                string tar_loc = txtLogfile.Text.Replace(Environment.NewLine, "");
                DirectoryInfo di = new DirectoryInfo(tar_loc);
                FileInfo[] files= di.GetFiles("*.csv");
                result_table_lst = new Dictionary<string, DataTable>();
                Dictionary<string, List<DataTable>> dic_tbl_data_lst = new Dictionary<string, List<DataTable>>();
                List<string> lst_log = new List<string>();
                foreach (var t in files)
                {
                    lst_log.Add(t.Name);
                }
                Dictionary<string, List<string>> block_logfile = proc_data.Get_BlockofLogFile_Taiyo(lst_log);
                foreach (var block in block_logfile)
                {
                    foreach (var logfile in block.Value)
                    {
                        DataTable dt = new DataTable();
                        string f_name = Path.Combine(txtLogfile.Text, logfile);
                        DataTable netSpec_tbl = new DataTable();
                        DataTable result_data = new DataTable();
                        int PCS_num = (int)numPCS.Value;
                        string sel_dev = cbMachine.Text;
                        try
                        {
                            switch (sel_dev)
                            {
                                case "YAMAHA":
                                    dt = proc_data.Yamaha_3GSPD_Process(f_name, ref netSpec_tbl);
                                    break;
                                case "TAIYO":
                                    dt = proc_data.Tayo_3GMRD_Process(f_name, PCS_num, ref netSpec_tbl);
                                    break;
                            }

                        }
                        catch
                        {
                            MessageBox.Show("Hãy chọn đúng loại thiết bị!","Thông báo");
                            txtLogfile.Text = txtLogfile.Text.Replace(Environment.NewLine, "");
                            return;
                        }
                        if (dt.Rows.Count > 0)
                        {
                            if (TDMK_Code.check_exist_list_index2(block.Key, dic_tbl_data_lst.Keys.ToList()) == -1)
                            {
                                dic_tbl_data_lst.Add(block.Key, new List<DataTable>() { dt });
                            }
                            else
                            {
                                dic_tbl_data_lst[block.Key].Add(dt);
                            }
                        }
                    }
                }
                foreach (var tbl in dic_tbl_data_lst)
                {
                    DataTable block_tbl = proc_data.Summary_Data_Table_from_list(tbl.Value, spec_dt);
                    result_table_lst.Add(tbl.Key, block_tbl);
                }
                DataTable Recheck_tbl_all = Summary_DataTable_FromList(result_table_lst, spec_dt);
                List<string> src_NG_Cols = Get_Columns_list((DataTable)(DGV_Details.DataSource));
                List<string> recheck_cols = Get_Columns_list(Recheck_tbl_all);
                List<string> sel_cols_lst = proc_data.Get_intersec(new List<List<string>> { src_NG_Cols, recheck_cols });
                DGV_Recheck.DataSource = Recheck_tbl_all.AsDataView().ToTable(false, sel_cols_lst.ToArray());
                check_in_spec_tbl(spec_dt, DGV_Recheck);
                myCode.DGV_Auto_Resize(DGV_Recheck);
            }
        }
        public DataTable Summary_DataTable_FromList(Dictionary<string, DataTable> src_result_table_lst, DataTable src_spec_dt, int numQty=0)
        {
            int div_factor = src_result_table_lst.Count;
            int fact = numQty / div_factor;
            List<int> item_qty = new List<int>();
            List<DataTable> summ_tbl_lst = new List<DataTable>();
            for (int i = 0; i < div_factor - 1; i++)
            {
                item_qty.Add(fact);
            }
            item_qty.Add(numQty - (div_factor - 1) * fact);
            int inx = 0;
            foreach (var tbl in src_result_table_lst)
            {
                List<ECheck_Process.NG_list2> cur_NGList = proc_data.Get_NG_point_tbl(src_spec_dt, tbl.Value);
                DataTable sel_dt = proc_data.Summary_Selected_FromExisted(tbl.Value, cur_NGList, item_qty[inx]);
                summ_tbl_lst.Add(sel_dt);
                inx++;
            }
            DataTable disp_result = new DataTable();
            string bl_name;
            for (int i = 0; i < summ_tbl_lst.Count; i++)
            {
                string[] bl_name_arr = result_table_lst.Keys.ToList()[i].Split('-');
                bl_name = bl_name_arr[bl_name_arr.Length - 1];
                for (int j = 0; j < summ_tbl_lst[i].Columns.Count; j++)
                {

                    string col_name = summ_tbl_lst[i].Columns[j].ColumnName + "_" + "BL" + bl_name;
                    disp_result.Columns.Add(col_name);
                }
            }
            for (int i = 0; i < summ_tbl_lst.Count; i++)
            {

                string[] bl_name_arr = result_table_lst.Keys.ToList()[i].Split('-');
                bl_name = bl_name_arr[bl_name_arr.Length - 1];
                for (int r_inx = 0; r_inx < summ_tbl_lst[i].Rows.Count; r_inx++)
                {
                    if (disp_result.Rows.Count < r_inx + 1)
                    {
                        disp_result.Rows.Add();
                    }
                    for (int c_inx = 0; c_inx < summ_tbl_lst[i].Columns.Count; c_inx++)
                    {
                        string col_name = summ_tbl_lst[i].Columns[c_inx].ColumnName + "_" + "BL" + bl_name;
                        disp_result.Rows[r_inx][col_name] = summ_tbl_lst[i].Rows[r_inx][c_inx];
                    }
                }
            }
            return disp_result;
        }
        public List<string> Get_Columns_list(DataTable src_dt)
        {
            List<string> result = new List<string>();
            foreach(DataColumn dc in src_dt.Columns)
            {
                result.Add(dc.ColumnName);
            }
            return result;
        }

        private void btnReplace_Click(object sender, EventArgs e)
        {
            update_lst = Get_update_lst(spec_dt, (DataTable)DGV_Recheck.DataSource, DGV_Details);
        }
        public Dictionary<string, List<int>> Get_update_lst(DataTable src_spec_dt, DataTable recheck_data, DataGridView tar_DGV)
        {
            Dictionary<string, List<int>> NG_lst = check_in_spec_tbl(src_spec_dt, (DataTable)tar_DGV.DataSource);
            Dictionary<string, List<int>> update_lst = new Dictionary<string, List<int>>();
            foreach (var NG_item in NG_lst)
            {
                string col_name = NG_item.Key;
                if (myCode.check_columns_existed(recheck_data, col_name))
                {
                    foreach (var net_no in NG_item.Value)
                    {
                        string replace_val = DGV_Recheck.Rows[net_no].Cells[col_name].Value.ToString();
                        if (myCode.IsNumeric(replace_val))
                        {
                            string USL = spec_dt.Rows[net_no][3].ToString();
                            string LSL = spec_dt.Rows[net_no][2].ToString();
                            if (myCode.check_in_limit2(USL, LSL, replace_val) == Color.White)
                            {
                                tar_DGV.Rows[net_no].Cells[col_name].Value = replace_val;
                                tar_DGV.Rows[net_no].Cells[col_name].Style.BackColor = Color.White;
                                if (update_lst.ContainsKey(col_name))
                                {
                                    update_lst[col_name].Add(net_no);
                                }
                                else
                                {
                                    update_lst.Add(col_name, new List<int> { net_no });
                                }
                            }
                        }
                    }
                }
            }
            return update_lst;
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, _Process, filter_str);
            DataView dv = src_dt.AsDataView();
            foreach(var t in update_lst)
            {
                string pcs_no = t.Key;
                foreach(var net_no in t.Value)
                {
                    string sel_net = (net_no + 1).ToString();
                    string sel_filter = TDMK_Code.filter_str(new string[] { "Pcs_No", "Net_No" }, new string[] { pcs_no, sel_net });
                    dv.RowFilter = sel_filter;
                    DataRow dr = dv[0].Row;
                    int r_x = src_dt.Rows.IndexOf(dr);
                    if (r_x != -1)
                    {
                        src_dt.Rows[r_x][Cycle] = DGV_Details.Rows[net_no].Cells[pcs_no].Value;
                    }
                }
            }
            TDMK_Code.Delelte_FilteredItem_arr(_Process, sqlcon, filter_str);
            proc_data.BatchBulkCopy(sqlcon, src_dt, _Process);
            MessageBox.Show("Hoàn thành cập nhật dữ liệu", "Thông báo");
        }
        private void btnExport_Click(object sender, EventArgs e)
        {
            //TDMK_Code.Export_DGV_Excel3(DGV_Details, TDMK_Code.Create_workbook(), true);
            //TDMK_Code.Export_DGV_Excel3(DGV_Recheck, TDMK_Code.Create_workbook(), true);
            string tbl_name = _Process + "_" + ItemCode + "-" + LotNo + "_NG_" + DateTime.Now.ToString("yyyyMMMdd_HHmmss");
            string recheck_name = _Process + "_" + ItemCode + "-" + LotNo + "_Recheck_" + DateTime.Now.ToString("yyyyMMMdd_HHmmss");
            myVar_code. Table_To_CSV(tbl_name, (DataTable)(DGV_Details.DataSource));
            myVar_code.Table_To_CSV(recheck_name, (DataTable)(DGV_Recheck.DataSource));
        }
    }
}
