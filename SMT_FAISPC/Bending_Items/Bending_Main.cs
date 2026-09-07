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
using TDMK_SEEV_DLL;
using TDMK_SQL;
using myExcel = Microsoft.Office.Interop.Excel;
using OK2SHIP;
using VHX;
using Resize_Lib;
using System.Data.SqlClient;
using IniLibs;
using Bending_Export;
using OfficeOpenXml;
using System.Diagnostics;

namespace Bending_Items
{
    public partial class Bending_Main : Form
    {
        SEI_Lib myCode = new SEI_Lib();
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        ECheck_Process proc_data = new ECheck_Process();
        myVar exp_proc = new myVar();
        //IniFile TDMK_init = new IniFile("Config.ini");
        SqlConnection sqlcon_OK2SHIP;
        DataTable mainTbl = new DataTable();
        List<ECheck_Process.NG_list2> myNGlst2;
        Dictionary<string, List<int>> Delta_R_NGlist;
        string lblItem_title = "NG NET List of Item: ";
        string lblNGDetail = "Details NG info: ";
        bool Logfile_mode = true;
        myVar_ECheck.item_info sel_item;
        myVar_ECheck.edit_cell sel_cell;
        //List<string> item_NG_lst = new List<string>();
        List<List<string>> item_lst_sum = new List<List<string>>();
        Dictionary<string, DataTable> result_table_lst = new Dictionary<string, DataTable>();
        SortedDictionary<int, string> dic_index_lst;
        bool manual_en = false;
        Resize rs = new Resize();
        public Bending_SMT_Lib bending_proc = new Bending_SMT_Lib();
        string Data_Location;
        string format_folder;
        string log_folder;
        List<string> Net_lst = new List<string>();
        Dictionary<string, string> sel_Log_info = new Dictionary<string, string>();
        public Bending_SMT_Lib.Cell_data_details sel_before_Cell;
        DataTable sel_spec_dt = new DataTable();
        DataTable sel_log_dt = new DataTable();
        public Dictionary<string, List<Bending_SMT_Lib.Cell_data_details>> dic_updated_bef = new Dictionary<string, List<Bending_SMT_Lib.Cell_data_details>>();
        public Bending_Export_EPPLUS_Lib bending_proc_Epplus = new Bending_Export_EPPLUS_Lib();
        public Bending_Main()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
        }

        private void Bending_Main_Load(object sender, EventArgs e)
        {
            //sqlcon_OK2SHIP = bending_proc.initial_data("OK2SHIP_SMT", true);
            sqlcon_OK2SHIP = bending_proc.initial_data("OK2SHIP_SMT_Q3_2025", true);
            //cbMachine.SelectedIndex = 0;
            cbShift.SelectedIndex = 0;
            cbShift_Sel.SelectedIndex = 0;
            myVar_ECheck.frmMain = this;
            rs.FindAllControls(this);
            blNG_Details.form_load();
        }
        private void cbProcess_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> cycles_lst = new List<string>();
            string app_path = Application.StartupPath;
            string tar_file = Path.Combine(app_path, "Config", cbProcess.Text + ".txt");
            cycles_lst = myCode.read_config_arr(tar_file).ToList();
            cbCycles.Text = "";
            cbCycles.DataSource = cycles_lst;
            reset_all();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder_diag = new FolderBrowserDialog();
            folder_diag.Description = "Select log file location...";
            folder_diag.SelectedPath = Application.StartupPath;
            if (folder_diag.ShowDialog() == DialogResult.OK)
            {
                txtLocation.Text = folder_diag.SelectedPath;
                DirectoryInfo di = new DirectoryInfo(folder_diag.SelectedPath);
                FileInfo[] files;
                switch (cbProcess.Text)
                {
                    case "CQRA_HOT_OIL":
                        files = di.GetFiles("*.xlsx");
                        break;
                    case "CQRA_BENDING":
                        files = di.GetFiles("*.xlsx");
                        break;
                    case "CQRA_THERMAL_CYCLING_AND_BEND":
                        //files = di.GetFiles("*.xlsx");
                        files = di.GetFiles("*.csv");
                        break;
                    case "CQRA_HEAT_SOAK_AND_BEND":
                        //files = di.GetFiles("*.xlsx");
                        files = di.GetFiles("*.csv");
                        break;
                    case "CQRA_SURVIVAL_HOT_OIL":
                        files = di.GetFiles("*.xlsx");
                        break;
                    case "CQRA_THERMAL_CYCLING_ON_COUPON":
                        files = di.GetFiles("*.xlsx");
                        break;
                    default:
                        files = di.GetFiles("*.DAT");
                        break;
                }
                List<string> _lstLogfile = new List<string>();
                foreach (FileInfo f in files)
                {
                    if (!f.Name.Contains("$"))
                    {
                        _lstLogfile.Add(f.Name);
                    }
                }
                lstLogFile.DataSource = _lstLogfile;
            }

        }
        private void lstLogFile_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstLogFile.SelectedIndex > -1)
            {
                txtItemCode.Enabled = false;
                txtLotNo.Enabled = false;
                DGV_NG_detail.DataSource = null;
                DGV_LogFile.DataSource = null;
                Net_lst.Clear();
                sel_Log_info.Clear();
                if (Logfile_mode)
                {
                    string log_locate = txtLocation.Text.Replace(Environment.NewLine, "");
                    string category = switchCategory(cbProcess.Text);
                    string cycle = switchCycle(cbCycles.Text);
                    log_locate = $"{log_locate}\\{category}\\{cycle}";
                    string f_name = Path.Combine(log_locate, lstLogFile.SelectedItem.ToString());
                    string itemname = "";
                    string userid = "";
                    DataTable netSpec_tbl = new DataTable();
                    DataTable result_data = new DataTable();
                    //result_data = bending_proc.DAT_To_DataTable_details_time(f_name, ref Net_lst, ref netSpec_tbl,ref sel_Log_info);
                    result_data = bending_proc.DAT_To_DataTable_details_time(f_name, ref Net_lst, ref netSpec_tbl, ref sel_Log_info, ref itemname, ref userid);
                    DGV_LogFile.DataSource = result_data;
                    DGV_NET_Spec.DataSource = netSpec_tbl;
                    txtItemName.Text = itemname;
                    txtOperator.Text = userid;
                }
                else
                {
                    string userid = "";
                    string itemname = "";
                    DGV_LogFile.DataSource = bending_proc.Load_Log_Data(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, cbCycles.Text, cbShift.Text, ref userid, ref itemname, lstLogFile.SelectedItem.ToString(), Net_lst);
                    DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { txtItemCode.Text, cbProcess.Text }));
                    DGV_NET_Spec.DataSource = spec_dt.AsDataView().ToTable(false, new string[] { "Net_Name", "Point+V", "Point-V", "LSL", "USL" });
                    txtOperator.Text = userid;
                    txtItemName.Text = itemname;
                }
                bending_proc.Marking_DGV_Rows_Net(DGV_LogFile, Net_lst);
                myCode.DGV_Auto_Resize(DGV_NET_Spec);
                if (DGV_NET_Spec.Rows.Count == DGV_LogFile.Rows.Count)
                {
                    myNGlst2 = proc_data.Get_NG_point3(DGV_NET_Spec, DGV_LogFile);
                    List<string> item_NG_lst = new List<string>();
                    DataTable tbl_NG_NET = new DataTable();
                    List<int> NET_index_lst = new List<int>();
                    proc_data.Summary_NG_list(myNGlst2, ref item_NG_lst, ref tbl_NG_NET, (DataTable)DGV_NET_Spec.DataSource, ref NET_index_lst);
                    lstNG.DataSource = item_NG_lst;
                    DGV_NET_NG.DataSource = tbl_NG_NET;
                    int r_inx = 0;
                    foreach (int inx in NET_index_lst)
                    {
                        DGV_NET_NG.Rows[r_inx].HeaderCell.Value = (inx + 1).ToString();
                        r_inx++;
                    }
                    DGV_NET_NG.AutoResizeColumns();
                    DGV_NET_NG.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                    txtTotal_NET.Text = NET_index_lst.Count.ToString();
                    txtTotalItem.Text = item_NG_lst.Count.ToString();
                }
                else
                {
                    MessageBox.Show("NET Spec data is not enough!\r\nPlease, check again NET Spec", "Warning");
                }
                lblNG_Detail.Text = lblNGDetail + " Log file No " + (lstLogFile.SelectedIndex + 1).ToString();
            }
            else
            {
                DGV_LogFile.DataSource = null;
                lblNG_Detail.Text = lblNGDetail.Replace(": ", "");
            }
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "")
            {
                if (DGV_LogFile.DataSource != null)
                {
                    //DataTable dt = bending_proc.Save_LogFile_detail_time(sqlcon_OK2SHIP, (DataTable)DGV_LogFile.DataSource, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, cbCycles.Text, lstLogFile.SelectedItem.ToString(), Net_lst, sel_Log_info);
                    DataTable dt = bending_proc.Save_LogFile_detail_time(sqlcon_OK2SHIP, (DataTable)DGV_LogFile.DataSource, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, cbCycles.Text, cbShift.Text, txtOperator.Text, txtItemName.Text, lstLogFile.SelectedItem.ToString(), Net_lst, sel_Log_info);
                    DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { txtItemCode.Text, cbProcess.Text }));
                    if (spec_dt.Rows.Count == 0)
                    {
                        List<string> items = new List<string>() { "ID", "ItemCode" };
                        DataTable temp_spec = (DataTable)DGV_NET_Spec.DataSource;
                        foreach (DataColumn dc in temp_spec.Columns)
                        {
                            items.Add("[" + dc.ColumnName + "]");
                        }
                        items.Add("Remark");
                        int id = TDMK_Code.SQL_MAX("NET_SPEC", "ID", sqlcon_OK2SHIP);
                        foreach (DataRow dr in temp_spec.Rows)
                        {
                            id++;
                            List<string> items_val = new List<string>() { id.ToString(), txtItemCode.Text };
                            foreach (DataColumn dc in temp_spec.Columns)
                            {
                                items_val.Add(dr[dc].ToString());
                            }
                            items_val.Add(cbProcess.Text);
                            TDMK_Code.insert_val_arr2("NET_SPEC", sqlcon_OK2SHIP, items.ToArray(), items_val.ToArray());
                        }
                    }
                    MessageBox.Show("Completed!", "Warning");
                }
                else
                {
                    MessageBox.Show("No data", "Warning");
                }
            }
            else
            {
                MessageBox.Show("Please, enter ItemCode!", "Warning");
            }
        }
        public void BatchBulkCopy(DataTable dataTable, string DestinationTbl)
        {
            // Get the DataTable 
            DataTable dtInsertRows = dataTable;
            //SqlBulkCopy tg = new SqlBulkCopy(sqlcon_OK2SHIP);
            using (SqlBulkCopy sbc = new SqlBulkCopy(sqlcon_OK2SHIP))
            {
                if (sqlcon_OK2SHIP.State != ConnectionState.Open)
                {
                    sqlcon_OK2SHIP.Open();
                }
                sbc.DestinationTableName = DestinationTbl;

                // Number of records to be processed in one go
                //sbc.BatchSize = batchSize;

                // Add your column mappings here
                foreach (DataColumn dc in dataTable.Columns)
                {
                    sbc.ColumnMappings.Add(dc.ColumnName, dc.ColumnName);
                }
                // Finally write to server
                sbc.WriteToServer(dtInsertRows);
                sqlcon_OK2SHIP.Close();
            }
        }

        private void btnLoadData_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "")
            {
                DataTable dt = new DataTable();
                DataTable spec_dt = new DataTable();
                if (!Logfile_mode)
                {
                    List<string> logfile_lst = new List<string>();
                    dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess.Text, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Cycles_name", "Shift" }, new string[] { txtItemCode.Text, txtLotNo.Text, cbCycles.Text, cbShift.Text }));
                    proc_data.Get_List_data(-1, dt, new string[] { "ItemCode", "LotNo", "Cycles_name" }, ref logfile_lst, "Logfile", true);
                    lstLogFile.DataSource = logfile_lst;
                    if (logfile_lst.Count == 0)
                    {
                        MessageBox.Show("No data", "Warning");
                    }
                }
            }
            else
            {
                MessageBox.Show("Please, enter ItemCode!", "Warning");
            }

        }

        public SqlConnection initial_data(string DB_name)
        {
            SqlConnection _sqlcon_OK2SHIP;
            string app_path = Application.StartupPath;
            string config_file = Path.Combine(app_path, "Config", "config.txt");
            string[] my_config = myCode.read_config_arr(config_file);
            string server_name = "";
            string server_acc = "";
            string server_pass = "";
            string Data_Location = "";
            string format_folder = "";
            string log_folder = "";
            foreach (string c in my_config)
            {
                if (c.Contains("Server"))
                {
                    server_name = c.Split(':')[1].Trim();
                }
                if (c.Contains("Account"))
                {
                    server_acc = c.Split(':')[1].Trim();
                }
                if (c.Contains("Password"))
                {
                    server_pass = c.Split(':')[1].Trim();
                }
                if (c.Contains("Data_Location"))
                {
                    string[] temp = c.Split(':');
                    string tg = "";
                    if (temp.Length > 2)
                    {

                        for (int t = 1; t < temp.Length; t++)
                        {
                            tg = tg + temp[t] + ":";
                        }
                    }
                    Data_Location = tg.TrimEnd(':').Trim();
                }
                if (c.Contains("Format_Folder"))
                {
                    format_folder = c.Split('#')[1].Trim();
                }
                if (c.Contains("Log_folder"))
                {
                    log_folder = c.Split('#')[1].Trim();
                }
            }
            string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, DB_name, server_acc, server_pass).ConnectionString;
            _sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);
            return _sqlcon_OK2SHIP;
        }
        //public SqlConnection initial_data(string DB_name, bool sa_en)
        //{
        //    SqlConnection _sqlcon_OK2SHIP;

        //    if (!File.Exists("Config.ini"))
        //    {
        //        TDMK_init.Write("Server", "10.212.1.243", "SMT_Config");
        //        TDMK_init.Write("Account", "sa", "SMT_Config");
        //        TDMK_init.Write("Password", "seev@123;", "SMT_Config");
        //        TDMK_init.Write("Report_Location", @"\\10.212.1.200\rgo\OK2SHIP automation system", "SMT_Config");
        //        TDMK_init.Write("Data_Location", @"\\10.212.1.200\rgo\OK2SHIP automation system\TDMK_Data\TestAreas\OK2SHIP_Measurements\OMM", "SMT_Config");
        //        TDMK_init.Write("Format_Folder", @"\\10.212.1.200\rgo\OK2SHIP automation system\0.OK2SHIP report format", "SMT_Config");
        //        TDMK_init.Write("Log_folder", @"\\10.212.1.200\rgo\OK2SHIP automation system\TDMK_Data\TestAreas\OK2SHIP_Measurements", "SMT_Config");
        //    }
        //    string server_name = TDMK_init.Read("Server", "SMT_Config");
        //    string server_acc = TDMK_init.Read("Account", "SMT_Config");
        //    string server_pass = TDMK_init.Read("Password", "SMT_Config");
        //    Data_Location = TDMK_init.Read("Data_Location", "SMT_Config");
        //    format_folder = TDMK_init.Read("Format_Folder", "SMT_Config");
        //    log_folder = TDMK_init.Read("Log_folder", "SMT_Config");
        //    if (sa_en)
        //    {
        //        string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, DB_name, server_acc, server_pass).ConnectionString;
        //        _sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);
        //    }
        //    else
        //    {
        //        string _strcon = TDMK_Code.data_connection2(server_name, DB_name).ConnectionString;
        //        _sqlcon_OK2SHIP = new SqlConnection(_strcon);
        //    }
        //    return _sqlcon_OK2SHIP;
        //}

        private void btnSummary_Click(object sender, EventArgs e)
        {
            /***************************** Main program ******************************************/
            //List<string> pcs_lst = new List<string>();
            //DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess.Text, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
            //bending_proc.Get_List_data(-1, src_tbl, new string[] { "ItemCode", "LotNo" }, ref pcs_lst, "Pcs_No", true);
            //List<string> sum_col_lst = bending_proc.Sum_OK_Item(sqlcon_OK2SHIP, (DataTable)DGV_LogFile.DataSource, txtItemCode.Text, cbProcess.Text, 1.67);
            //List<string> sel_sum_col_lst = sum_col_lst;
            //if (pcs_lst.Count>0)
            //{
            //    MessageBox.Show("Đã dữ liệu các pcs", "Thông báo");
            //    sel_sum_col_lst = bending_proc.Get_intersec(new List<List<string>> { sum_col_lst, pcs_lst });
            //}          
            //int qty = Math.Min(sel_sum_col_lst.Count, (int)numQty.Value);
            //if (sel_sum_col_lst.Count < (int)numQty.Value)
            //{
            //    MessageBox.Show("Không đủ số lượng dữ liệu OK", "Cảnh báo");
            //}
            //List<string> sel_col_lst = new List<string>();
            //for (int i = 0; i < qty; i++)
            //{
            //    sel_col_lst.Add(sel_sum_col_lst[i]);
            //}
            //DataTable src_dt = (DataTable)DGV_LogFile.DataSource;
            //DGV_Data.DataSource = src_dt.AsDataView().ToTable(false, sel_col_lst.ToArray());
            //myCode.DGV_Auto_Resize(DGV_Data);

            /************************************************************************************/
            /***************************** Test program ******************************************/
            Summary_data_Block();
            //Dictionary<string, DataTable> dic_sum = bending_proc. Summary_Cycles_Logfile(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, cbShift.Text, Convert.ToInt32(numQty.Value));
            //if(dic_sum.Keys.ToList().IndexOf(cbCycles.Text)!=-1)
            //{
            //    DataTable tar_dt = dic_sum[cbCycles.Text];
            //    DGV_Data.DataSource = tar_dt;
            //}
            //else
            //{
            //    MessageBox.Show("Không có dữ liệu", "Thông báo");
            //}

            /************************************************************************************/

        }
        public void Summary_data_Block()
        {
            DGV_Data.DataSource = null;
            if (lstLogFile.Items.Count > 0)
            {
                result_table_lst = new Dictionary<string, DataTable>();
                Dictionary<string, List<DataTable>> dic_tbl_data_lst = new Dictionary<string, List<DataTable>>();
                List<string> lst_log = new List<string>();
                foreach (string t in lstLogFile.Items)
                {
                    lst_log.Add(t);
                }
                Dictionary<string, List<string>> block_logfile = bending_proc.Get_BlockofLogFile(lst_log);
                foreach (var block in block_logfile)
                {
                    foreach (var logfile in block.Value)
                    {
                        DataTable dt = new DataTable();
                        if (Logfile_mode)
                        {
                            string location = txtLocation.Text;
                            string category = switchCategory(cbProcess.Text);
                            string cycle = switchCycle(cbCycles.Text);
                            location = $"{location}\\{category}\\{cycle}";
                            string f_name = Path.Combine(location, logfile);
                            DataTable netSpec_tbl = new DataTable();
                            dt = bending_proc.DAT_To_DataTable_details(f_name, ref Net_lst, ref netSpec_tbl);
                        }
                        else
                        {
                            List<string> temp_net_lst = new List<string>();
                            string userid = "";
                            string itemname = "";
                            dt = bending_proc.Load_Log_Data(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, cbCycles.Text, cbShift.Text, ref userid, ref itemname, logfile, temp_net_lst);
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
                //foreach (var tbl in dic_tbl_data_lst)
                //{
                //    DataTable block_tbl = bending_proc.Summary_Data_Table_from_list(tbl.Value, (DataTable)DGV_NET_Spec.DataSource);
                //    result_table_lst.Add(tbl.Key, block_tbl);
                //}

                string spec_filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { txtItemCode.Text, cbProcess.Text });
                DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", spec_filter_str);
                foreach (var tbl in dic_tbl_data_lst)
                {
                    DataTable block_tbl = bending_proc.Summary_Data_Table_from_list(tbl.Value, spec_dt);
                    List<string> col_lst = bending_proc.Get_column_name(block_tbl);
                    if (col_lst.Count > 1)
                    {
                        List<string> main_col_lst = col_lst.Where(x => !x.Contains("_")).ToList();
                        DataTable final_dt = new DataTable();
                        foreach (string sel_col in main_col_lst)
                        {
                            string[] cur_col = col_lst.Where(x => x.Contains(sel_col)).ToArray();
                            DataTable cur_dt = block_tbl.AsDataView().ToTable(false, cur_col);
                            final_dt.Columns.Add(sel_col);
                            int r_inx = 0;
                            foreach (DataRow dr in cur_dt.Rows)
                            {
                                string USL = spec_dt.Rows[r_inx]["USL"].ToString();
                                string LSL = spec_dt.Rows[r_inx]["LSL"].ToString();
                                List<double> col_data_lst = new List<double>();
                                foreach (DataColumn dc in cur_dt.Columns)
                                {
                                    string cur_val = myCode.checkDBNull(dr[dc]);
                                    if (bending_proc.check_in_limit(USL, LSL, cur_val))
                                    {
                                        col_data_lst.Add(Convert.ToDouble(cur_val));
                                    }
                                }
                                if (final_dt.Rows.Count <= r_inx)
                                {
                                    final_dt.Rows.Add();
                                }
                                if (col_data_lst.Count > 0)
                                {
                                    final_dt.Rows[r_inx][sel_col] = col_data_lst.Min();
                                }
                                r_inx++;
                            }
                        }
                        result_table_lst.Add(tbl.Key, final_dt);
                    }
                    else
                    {
                        result_table_lst.Add(tbl.Key, block_tbl);
                    }
                }
                btnDataSelect.Enabled = true;
                MessageBox.Show("Summary completed!\r\nPlease, choose data Qty then press Data Selected button", "Warning");
            }
        }
        public void check_in_spec(DataGridView dgv_spec, DataGridView dgv_data)
        {
            foreach (DataGridViewRow dr in dgv_data.Rows)
            {
                string USL = dgv_spec.Rows[dr.Index].Cells[3].Value.ToString();
                string LSL = dgv_spec.Rows[dr.Index].Cells[2].Value.ToString();
                foreach (DataGridViewColumn dc in dgv_data.Columns)
                {
                    DataGridViewCell sel_cell = dr.Cells[dc.Index];
                    sel_cell.Style.BackColor = myCode.check_in_limit2(USL, LSL, sel_cell.Value.ToString());
                }
            }
        }
        public void check_in_spec_tbl(DataTable dgv_spec, DataGridView dgv_data)
        {
            int r_inx = 0;
            foreach (DataGridViewRow dr in dgv_data.Rows)
            {
                string USL = dgv_spec.Rows[r_inx][3].ToString();
                string LSL = dgv_spec.Rows[r_inx][2].ToString();
                foreach (DataGridViewColumn dc in dgv_data.Columns)
                {
                    DataGridViewCell sel_cell = dr.Cells[dc.Index];
                    sel_cell.Style.BackColor = myCode.check_in_limit2(USL, LSL, sel_cell.Value.ToString());
                }
                r_inx++;
            }
        }
        public Color check_in_limit(string src_UL, string src_LL, string src_act_val)
        {
            Color tar_color = Color.LightPink;
            if (myCode.IsNumeric(src_act_val))
            {
                if (src_UL != "")
                {
                    double UL = Convert.ToDouble(src_UL);
                    double act_val = Convert.ToDouble(src_act_val);
                    if (src_LL != "")
                    {
                        double LL = Convert.ToDouble(src_LL);
                        if ((act_val <= UL) && (act_val >= LL))
                        {

                            tar_color = Color.White;

                        }
                        else
                        {
                            if (act_val > UL)
                            {
                                tar_color = Color.Yellow;
                            }
                            else
                            {
                                tar_color = Color.LightBlue;
                            }
                        }
                    }
                    else
                    {
                        if (act_val <= UL)
                        {
                            tar_color = Color.White;

                        }
                        else
                        {
                            tar_color = Color.Yellow;
                        }
                    }
                }
                else
                {
                    if (src_LL != "")
                    {
                        double LL = Convert.ToDouble(src_LL);
                        double act_val = Convert.ToDouble(src_act_val);
                        if (act_val >= LL)
                        {
                            tar_color = Color.White;

                        }
                        else
                        {
                            tar_color = Color.LightBlue;
                        }
                    }
                    else
                    {
                        tar_color = Color.Gray;
                    }
                }
            }
            return tar_color;
        }
        private void btnDataSelect_Click(object sender, EventArgs e)
        {
            List<string> pcs_lst = new List<string>();
            DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess.Text, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Shift" }, new string[] { txtItemCode.Text, txtLotNo.Text, cbShift.Text }));
            proc_data.Get_List_data(-1, src_tbl, new string[] { "ItemCode", "LotNo" }, ref pcs_lst, "Pcs_No", true);
            if (pcs_lst.Count > 0)
            {
                //MessageBox.Show("Existed!");
                DataTable cur_dt = new DataTable();
                for (int i = 0; i < pcs_lst.Count; i++)
                {
                    cur_dt.Columns.Add(pcs_lst[i]);
                }
                foreach (var tbl in result_table_lst)
                {
                    //string bl_name = tbl.Key.Split('-')[2];
                    string[] bl_name_arr = tbl.Key.Split('-');
                    string bl_name = bl_name_arr[bl_name_arr.Length - 1];
                    for (int r_inx = 0; r_inx < tbl.Value.Rows.Count; r_inx++)
                    {
                        if (cur_dt.Rows.Count < r_inx + 1)
                        {
                            cur_dt.Rows.Add();
                        }
                        for (int c_inx = 0; c_inx < tbl.Value.Columns.Count; c_inx++)
                        {
                            string col_name = tbl.Value.Columns[c_inx].ColumnName;// + "_" + "BL" + bl_name;
                            if (TDMK_Code.check_exist_list_index2(col_name, pcs_lst) != -1)
                            {
                                cur_dt.Rows[r_inx][col_name] = tbl.Value.Rows[r_inx][c_inx];
                            }
                        }
                    }
                }
                DGV_Data.DataSource = cur_dt;
            }
            else
            {
                int div_factor = result_table_lst.Count;
                int fact = (int)numQty.Value / div_factor;
                List<int> item_qty = new List<int>();
                List<DataTable> summ_tbl_lst = new List<DataTable>();
                for (int i = 0; i < div_factor - 1; i++)
                {
                    item_qty.Add(fact);
                }
                item_qty.Add((int)numQty.Value - (div_factor - 1) * fact);
                int inx = 0;
                foreach (var tbl in result_table_lst)
                {
                    List<ECheck_Process.NG_list2> cur_NGList = proc_data.Get_NG_point_tbl((DataTable)DGV_NET_Spec.DataSource, tbl.Value);
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

                        string col_name = summ_tbl_lst[i].Columns[j].ColumnName;// + "_" + "BL" + bl_name;
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
                            string col_name = summ_tbl_lst[i].Columns[c_inx].ColumnName;// + "_" + "BL" + bl_name;
                            disp_result.Rows[r_inx][col_name] = summ_tbl_lst[i].Rows[r_inx][c_inx];
                        }
                    }
                }
                DGV_Data.DataSource = disp_result;
            }
            myCode.DGV_Auto_Resize(DGV_Data);
            if (DGV_NET_Spec.Rows.Count == DGV_Data.Rows.Count)
            {
                DataTable tbl_NG_NET = new DataTable();
                List<int> NET_index_lst = new List<int>();
                List<string> item_NG_lst = new List<string>();
                proc_data.Submit_Data(DGV_NET_Spec, DGV_Data, ref tbl_NG_NET, ref NET_index_lst, DGV_NET_NG, ref item_NG_lst, ref myNGlst2);
                lstNG.DataSource = item_NG_lst;
                DGV_NET_NG.DataSource = tbl_NG_NET;
                int r_inx = 0;
                foreach (int inx in NET_index_lst)
                {
                    DGV_NET_NG.Rows[r_inx].HeaderCell.Value = (inx + 1).ToString();
                    r_inx++;
                }
                DGV_NET_NG.AutoResizeColumns();
                DGV_NET_NG.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                txtTotal_NET.Text = NET_index_lst.Count.ToString();
                txtTotalItem.Text = item_NG_lst.Count.ToString();
            }
            else
            {
                MessageBox.Show("NET Spec Data is not enough!\r\nPlease, check NET Spec", "Warning");
            }
        }

        private void DGV_Data_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //List<string> col_lst = new List<string>() { "ItemCode", "LotNo", "Net_No", "Pcs_No" };
            //DataTable cur_tbl = (DataTable)DGV_Data.DataSource;
            //string sel_col_val = cur_tbl.Columns[e.ColumnIndex].ColumnName;
            //lblLogfile.Text = "Log file data: " + sel_col_val;
            //col_lst.Add(sel_col_val);
            //DataTable filter_tbl = cur_tbl.AsDataView().ToTable(false, col_lst.ToArray());
            //DataTable tbl_detail = proc_data.Detail_NET_PCS(filter_tbl);
            //DGV_LogFile.DataSource = tbl_detail;
            //myCode.DGV_Auto_Resize(DGV_LogFile);
            //List<ECheck_Process.NG_list> myNGlst = proc_data.Get_NG_point2(DGV_NET_Spec,DGV_LogFile);
            //Marking_NG_Point2(myNGlst);
        }
        public void Marking_NG_Point(DataGridView tar_DGV, List<ECheck_Process.NG_list> templst)
        {
            List<string> item_NG_lst = new List<string>();
            foreach (ECheck_Process.NG_list item in templst)
            {
                int r_inx = item.row_inx;
                string col_name = item.col_val;
                item_NG_lst.Add(col_name);
                DGV_LogFile.Rows[r_inx].Cells[col_name].Style.BackColor = Color.Red;
            }
            lstNG.DataSource = item_NG_lst.Distinct().ToList();
            if (item_NG_lst.Count > 0)
            {
                lblDATA.BackColor = Color.OrangeRed;
            }
            else
            {
                lblDATA.BackColor = Color.LightGreen;
            }
        }
        public void Marking_NG_Point2(List<ECheck_Process.NG_list> templst)
        {
            List<string> item_NG_lst = new List<string>();
            foreach (ECheck_Process.NG_list item in templst)
            {
                int r_inx = item.row_inx;
                string col_name = item.col_val;
                item_NG_lst.Add(col_name);
            }
            lstNG.DataSource = item_NG_lst.Distinct().ToList();
            if (item_NG_lst.Count > 0)
            {
                lblDATA.BackColor = Color.OrangeRed;
            }
            else
            {
                lblDATA.BackColor = Color.LightGreen;
            }
        }
        public void Marking_NG_Point3(List<ECheck_Process.NG_list2> templst)
        {
            List<string> item_NG_lst = new List<string>();
            foreach (ECheck_Process.NG_list2 item in templst)
            {
                //int r_inx = item.row_inx;
                string col_name = item.col_val;
                item_NG_lst.Add(col_name);
            }
            lstNG.DataSource = item_NG_lst.Distinct().ToList();
            if (item_NG_lst.Count > 0)
            {
                lblDATA.BackColor = Color.OrangeRed;
            }
            else
            {
                lblDATA.BackColor = Color.LightGreen;
            }
        }

        public List<string> Get_NG_List(DataTable src_dt)
        {
            List<string> item_NG_lst = new List<string>();

            foreach (DataColumn dc in src_dt.Columns)
            {
                bool ignored_en = false;
                foreach (DataRow dr in src_dt.Rows)
                {
                    if (!myCode.IsNumeric(dr[dc].ToString()))
                    {
                        ignored_en = true;
                    }
                }
                if (ignored_en)
                {
                    item_NG_lst.Add(dc.ColumnName);
                }
            }
            return item_NG_lst;
        }
        public SortedDictionary<int, string> Get_NG_point(DataTable src_dt)
        {
            SortedDictionary<int, string> _result = new SortedDictionary<int, string>();
            foreach (DataColumn dc in src_dt.Columns)
            {
                int r_inx = 0;
                foreach (DataRow dr in src_dt.Rows)
                {
                    if (!myCode.IsNumeric(dr[dc].ToString()))
                    {
                        _result.Add(r_inx, dc.ColumnName);
                    }
                    r_inx++;
                }
            }
            return _result;
        }
        //public List<NG_list> Get_NG_point2(DataTable src_dt)
        //{
        //    List<NG_list> _result = new List<NG_list>();
        //    foreach (DataColumn dc in src_dt.Columns)
        //    {
        //        int r_inx = 0;
        //        foreach (DataRow dr in src_dt.Rows)
        //        {
        //            if (!myCode.IsNumeric(dr[dc].ToString()))
        //            {
        //                NG_list sel_NG = new NG_list(r_inx, dc.ColumnName);
        //                _result.Add(sel_NG);
        //            }
        //            r_inx++;
        //        }
        //    }
        //    return _result;
        //}
        private void cbCycles_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbCycles.SelectedIndex >= 0)
            {
                lblLogfile.Text = "Log file data: " + cbCycles.Text;
                reset_all();
            }
            else
            {
                lblLogfile.Text = "Log file data";
            }

        }

        private void lstNG_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstNG.SelectedIndex > -1)
            {
                int inx = lstNG.SelectedIndex;
                lblItemNG.Text = lblItem_title + lstNG.Items[inx];
                DGV_NG_detail.DataSource = proc_data.NG_NET_detail((DataTable)DGV_NET_Spec.DataSource, myNGlst2[inx]);
                int r_inx = 0;
                foreach (int index in myNGlst2[inx].lst_row_inx)
                {
                    DGV_NG_detail.Rows[r_inx].HeaderCell.Value = (index + 1).ToString();
                    r_inx++;
                }
                DGV_NG_detail.AutoResizeColumns();
                DGV_NG_detail.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            }
            else
            {
                lblItemNG.Text = lblItem_title;
            }
        }

        private void rbLogfile_CheckedChanged(object sender, EventArgs e)
        {
            //if(rbLogfile.Checked)
            //{
            //    Logfile_mode = true;
            //    //btnLoadData.Enabled = false;
            //}
            //else
            //{
            //    Logfile_mode = false;
            //    //btnLoadData.Enabled = true;
            //}
        }
        //        public DataTable save_data_type3(string ItemCode, string LotNo)
        //        {
        //            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
        //            DataTable temp_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess.Text, filter_str);
        //            if (temp_dt.Rows.Count > 0)
        //            {
        //                int r_inx = 0;
        //                foreach (DataGridViewColumn dc in DGV_LogFile.Columns)
        //                {
        //                    foreach (DataGridViewRow dr in DGV_LogFile.Rows)
        //                    {
        //                        if (r_inx < temp_dt.Rows.Count)
        //                        {
        //                            temp_dt.Rows[r_inx][cbCycles.Text] = dr.Cells[dc.Index].Value.ToString();
        //                        }
        //                        r_inx++;
        //                    }
        //                }
        //                TDMK_Code.Delelte_FilteredItem_arr(cbProcess.Text, sqlcon_OK2SHIP, filter_str);
        //                BatchBulkCopy(temp_dt, cbProcess.Text);
        //                //MessageBox.Show("update data completed!");
        //            }
        //            else
        //            {
        //                int id = TDMK_Code.SQL_MAX(cbProcess.Text, "ID", sqlcon_OK2SHIP);
        //                foreach (DataGridViewColumn dc in DGV_LogFile.Columns)
        //                {
        //                    foreach (DataGridViewRow dr in DGV_LogFile.Rows)
        //                    {
        //                        id++;
        //                        TDMK_Code.insert_val_arr2(cbProcess.Text, sqlcon_OK2SHIP, new string[] { "ID", "ItemCode", "LotNo", "Net_No", "Pcs_No", cbCycles.Text }, new string[] { id.ToString(), txtItemCode.Text, txtLotNo.Text, (dr.Index + 1).ToString(), (dc.Index + 1).ToString(), dr.Cells[dc.Index].Value.ToString() });
        //                    }
        //                }
        //                //MessageBox.Show("Insert data completed!");
        //            }
        //            return TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess.Text, filter_str);
        //        }
        //        public DataTable Load_Cycles_Data(string ItemCode, string LotNo, string process_name, string cycle_name)
        //        {
        //            DataTable dgv_dt = new DataTable();
        //            List<DataTable> myLst_tbl = new List<DataTable>();
        //            DataSet DS = new DataSet();
        //            if(cycle_name =="")
        //            {
        //                cycle_name = "Data";
        //            }    
        //            TDMK_Code.fill_dataset_Filter_arr(DS, sqlcon_OK2SHIP, process_name, new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
        //            DataTable dt = DS.Tables[0];
        //            Get_ListTable(-1, dt, new string[] { "ItemCode", "LotNo", "Net_No" }, ref myLst_tbl, cycle_name);
        //            foreach (DataTable temp_dt in myLst_tbl)
        //            {
        //                List<string> temp_data = new List<string>();
        //                Get_List_data(-1, temp_dt, new string[] { "ItemCode", "LotNo" }, ref temp_data,cycle_name, false);
        //                if (dgv_dt.Columns.Count < temp_data.Count)
        //                {
        //                    for (int i = 0; i < temp_data.Count; i++)
        //                    {
        //                        string col_name = "ItemNo_" + (i + 1).ToString();
        //                        if (!myCode.check_columns_existed(dgv_dt, col_name))
        //                        {
        //                            dgv_dt.Columns.Add(col_name);
        //                        }
        //                    }
        //                }
        //                DataRow dr = dgv_dt.NewRow();
        //                int col_inx = 0;
        //                foreach (string col in temp_data)
        //                {
        //                    dr[col_inx] = col;
        //                    col_inx++;
        //                }
        //                dgv_dt.Rows.Add(dr);
        //            }
        //            return dgv_dt;
        //        }
        //        public void save_data_type4()
        //        {
        //            start_lbl: string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
        //            DataTable temp_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess.Text, filter_str);
        //            DataTable src_dt =  (DataTable) DGV_LogFile.DataSource;
        //            if (temp_dt.Rows.Count > 0)
        //            {
        //                if (MessageBox.Show("Data existed!\r\n" + "Do you want to overwrite it?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
        //                {
        //                    if(temp_dt.Rows.Count==src_dt.Rows.Count * src_dt.Columns.Count)
        //                    {
        //                        int id = 0;
        //                        foreach (DataColumn dc in src_dt.Columns)
        //                        {
        //                            foreach (DataRow dr in src_dt.Rows)
        //                            {
        //                                temp_dt.Rows[id]["Data"] = dr[dc].ToString();
        //                                id++;
        //                            }
        //                        }
        //                        TDMK_Code.Delelte_FilteredItem_arr(cbProcess.Text, sqlcon_OK2SHIP, filter_str);
        //                        BatchBulkCopy(temp_dt, cbProcess.Text);
        //                        MessageBox.Show("Update Completed!");
        //                    }
        //                    else
        //                    {
        //                        TDMK_Code.Delelte_FilteredItem_arr(cbProcess.Text, sqlcon_OK2SHIP, filter_str);
        //                        goto start_lbl;
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                int id = TDMK_Code.SQL_MAX(cbProcess.Text, "ID", sqlcon_OK2SHIP);
        //                foreach (DataGridViewColumn dc in DGV_LogFile.Columns)
        //                {
        //                    foreach (DataGridViewRow dr in DGV_LogFile.Rows)
        //                    {
        //                        id++;
        //                        TDMK_Code.insert_val_arr2(cbProcess.Text, sqlcon_OK2SHIP, new string[] { "ID", "ItemCode", "LotNo", "Net_No", "Pcs_No", "Data", "Process" }, new string[] { id.ToString(), txtItemCode.Text, txtLotNo.Text, (dr.Index + 1).ToString(), (dc.Index + 1).ToString(), dr.Cells[dc.Index].Value.ToString(), cbCycles.Text });
        //                    }
        //                }
        //                MessageBox.Show("Insert Completed!");
        //            }
        //        }
        //        public DataTable Save_LogFile(DataTable logfile_tbl, string ItemCode, string LotNo, string process_name, string cycle_name)
        //        {
        //start_lbl:  string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
        //            DataTable temp_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
        //            if(cycle_name =="")
        //            {
        //                cycle_name = "Data";
        //            }
        //            if (temp_dt.Rows.Count > 0)
        //            {
        //                if(temp_dt.Rows.Count == logfile_tbl.Rows.Count*logfile_tbl.Columns.Count)
        //                {
        //                    int r_inx = 0;
        //                    foreach (DataColumn dc in logfile_tbl.Columns)
        //                    {
        //                        foreach (DataRow dr in logfile_tbl.Rows)
        //                        {
        //                            if (r_inx < temp_dt.Rows.Count)
        //                            {
        //                                temp_dt.Rows[r_inx][cycle_name] = dr[dc].ToString();
        //                            }
        //                            r_inx++;
        //                        }
        //                    }
        //                    TDMK_Code.Delelte_FilteredItem_arr(process_name, sqlcon_OK2SHIP, filter_str);
        //                    BatchBulkCopy(temp_dt, process_name);
        //                }
        //                else
        //                {
        //                    TDMK_Code.Delelte_FilteredItem_arr(process_name, sqlcon_OK2SHIP, filter_str);
        //                    goto start_lbl;
        //                }

        //            }
        //            else
        //            {
        //                int id = TDMK_Code.SQL_MAX(process_name, "ID", sqlcon_OK2SHIP);
        //                int col_inx = 0;
        //                foreach (DataColumn dc in logfile_tbl.Columns)
        //                {
        //                    int r_inx = 0;
        //                    foreach (DataRow dr in logfile_tbl.Rows)
        //                    {
        //                        id++;
        //                        TDMK_Code.insert_val_arr2(process_name, sqlcon_OK2SHIP, new string[] { "ID", "ItemCode", "LotNo", "Net_No", "Pcs_No", cycle_name }, new string[] { id.ToString(), txtItemCode.Text, txtLotNo.Text, (r_inx + 1).ToString(), (col_inx + 1).ToString(), dr[dc].ToString()});
        //                        r_inx++;
        //                    }
        //                    col_inx++;
        //                }
        //            }
        //            return TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
        //        }
        public void load_spec()
        {
        start_lbl: string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { txtItemCode.Text, cbProcess.Text });
            DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", filter_str);
            if (spec_dt.Rows.Count > 0)
            {
                if (MessageBox.Show("Do you want to update NET Spec table?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    TDMK_Code.Delelte_FilteredItem_arr("NET_SPEC", sqlcon_OK2SHIP, filter_str);
                    goto start_lbl;
                }
            }
            else
            {
                int inx = TDMK_Code.SQL_MAX("NET_SPEC", "ID", sqlcon_OK2SHIP) + 1;
                foreach (DataGridViewRow dr in DGV_NET_Spec.Rows)
                {
                    List<string> item_val = new List<string>();
                    string ID = inx.ToString();
                    item_val.Add(ID);
                    item_val.Add(txtItemCode.Text);
                    foreach (DataGridViewColumn dc in DGV_NET_Spec.Columns)
                    {
                        item_val.Add(dr.Cells[dc.Index].Value.ToString());
                    }
                    item_val.Add(cbProcess.Text);
                    TDMK_Code.insert_val_arr2("NET_SPEC", sqlcon_OK2SHIP, new string[] { "ID", "ItemCode", "[Point+V]", "[Point-V]", "LSL", "USL", "Remark" }, item_val.ToArray());
                    inx++;
                }
                spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", filter_str);
            }
            DGV_NET_Spec.DataSource = spec_dt.AsDataView().ToTable(false, new string[] { "Point+V", "Point-V", "LSL", "USL" });
        }

        private void rbDatabase_CheckedChanged(object sender, EventArgs e)
        {
            if (rbDatabase.Checked)
            {
                btnBrowse.Enabled = false;
                Logfile_mode = false;
                reset_all();
            }
            else
            {
                Logfile_mode = true;
                btnBrowse.Enabled = true;
                reset_all();
            }
        }
        public void reset_all()
        {
            DGV_LogFile.DataSource = null;
            DGV_Data.DataSource = null;
            DGV_NET_NG.DataSource = null;
            DGV_NET_Spec.DataSource = null;
            DGV_NG_detail.DataSource = null;
            lstLogFile.DataSource = null;
            lstNG.DataSource = null;
            txtTotalItem.Clear();
            txtTotal_NET.Clear();
            txtLocation.Clear();
            lblDATA.BackColor = this.BackColor;
            btnSaveSubmit.Enabled = false;
            GB_Submit.Enabled = false;
            btnSave.Enabled = false;
            btnSaveAll.Enabled = false;
            btnDataSelect.Enabled = false;
            if (Logfile_mode)
            {
                btnLoadData.Enabled = false;
                //btnSave.Enabled = true;
                txtLocation.Enabled = true;
                btnBrowse.Enabled = true;
            }
            else
            {
                btnLoadData.Enabled = true;
                //btnSave.Enabled = false;
                txtLocation.Enabled = false;
                btnBrowse.Enabled = false;
            }
            txtItemCode.Enabled = true;
            txtLotNo.Enabled = true;
            txtItemName.Clear();
            txtOperator.Clear();
        }
        public void reset_all_sel()
        {
            //DGV_LogFile.DataSource = null;
            //DGV_Data.DataSource = null;
            //DGV_NET_NG.DataSource = null;
            //DGV_NET_Spec.DataSource = null;
            //DGV_NG_detail.DataSource = null;
            //lstLogFile.DataSource = null;
            //lstNG.DataSource = null;
            //txtTotalItem.Clear();
            //txtTotal_NET.Clear();
            //txtLocation.Clear();
            //lblDATA.BackColor = this.BackColor;
            blNG_Details.reset_all();
            DGV_Cycles_Data_Sel.DataSource = null;
            DGV_Data_Sel.DataSource = null;
        }
        private void btnSaveSubmit_Click(object sender, EventArgs e)
        {
            DGV_Data.DataSource = proc_data.Save_Submit_data(sqlcon_OK2SHIP, (DataTable)DGV_Data.DataSource, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, cbCycles.Text, cbShift.Text);
            myCode.DGV_Auto_Resize(DGV_Data);
            btnSaveSubmit.Enabled = false;
        }

        private void DGV_Data_DataSourceChanged(object sender, EventArgs e)
        {
            if (DGV_Data.DataSource != null)
            {
                btnSaveSubmit.Enabled = true;
            }
            else
            {
                btnSaveSubmit.Enabled = false;
            }
        }

        private void DGV_LogFile_DataSourceChanged(object sender, EventArgs e)
        {
            if (DGV_LogFile.DataSource == null)
            {
                btnSave.Enabled = false;
                GB_Submit.Enabled = false;
                btnSaveAll.Enabled = false;
            }
            else
            {
                GB_Submit.Enabled = true;
                if (Logfile_mode)
                {
                    btnSave.Enabled = true;
                    btnSaveAll.Enabled = true;
                }
                //else
                //{
                //    Data_Details frmData_details = new Data_Details((DataTable)DGV_LogFile.DataSource, txtItemCode.Text, cbProcess.Text, sqlcon_OK2SHIP);
                //    frmData_details.Show();
                //}

                //if (myVar_ECheck.CSV_en)
                //{
                //    GB_Submit.Enabled = true;
                //    if (Logfile_mode)
                //    {
                //        btnSave.Enabled = true;
                //    }
                //}
                //else
                //{
                //    GB_Submit.Enabled = false;
                //    if (cbProcess.SelectedItem.ToString() == "CQRA_HOT_OIL")
                //    {
                //        if (Logfile_mode)
                //        {
                //            btnSave.Enabled = true;
                //        }
                //        else
                //        {
                //            btnSaveSubmit.Enabled = true;
                //        }
                //    }
                //}

            }
        }

        private void numQty_ValueChanged(object sender, EventArgs e)
        {
            //btnDataSelect.Enabled = false;
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            int ng_item_inx = 0;
            foreach (ECheck_Process.NG_list2 ng in blNG_Details.myNGlst2)
            {
                string cur_pcs = ng.col_val.Split('_')[1];
                foreach (int r_inx in ng.lst_row_inx)
                {
                    int tar_row = ng_item_inx * DGV_Cycles_Data_Sel.Rows.Count + r_inx;
                    string val_before = DGV_Data_Sel.Rows[tar_row].Cells[sel_item.cur_Cycles].Value.ToString();
                    string val_after = DGV_Cycles_Data_Sel.Rows[r_inx].Cells[ng_item_inx].Value.ToString();
                    if (val_before != val_after)
                    {
                        DGV_Data_Sel.Rows[tar_row].Cells[sel_item.cur_Cycles].Value = val_after;
                        string id = (TDMK_Code.SQL_MAX("EDIT_HISTORY", "ID", sqlcon_OK2SHIP) + 1).ToString();
                        string[] item = new string[] { "ID", "ItemCode", "LotNo", "Net_No", "Pcs_No", "Data_Before", "Data_After", "UserID", "Date_Modify", "Process", "Remark" };
                        string[] item_val = new string[] { id, sel_item.cur_ItemCode, sel_item.cur_LotNo, (r_inx + 1).ToString(), cur_pcs, val_before, val_after, myVar_ECheck.cur_User, DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), sel_item.cur_Process, sel_item.cur_Cycles };
                        TDMK_Code.insert_val_arr2("EDIT_HISTORY", sqlcon_OK2SHIP, item, item_val);
                    }
                }
                ng_item_inx++;
            }
        }

        private void cbProcess_Sel_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbProcess_Sel.SelectedIndex != -1)
            {
                List<string> cycles_lst = new List<string>();
                string app_path = Application.StartupPath;
                string tar_file = Path.Combine(app_path, "Config", cbProcess_Sel.Text + ".txt");
                cycles_lst = myCode.read_config_arr(tar_file).ToList();
                cbCycles_Sel.Text = "";
                cbCycles_Sel.DataSource = cycles_lst;
                reset_all_sel();
            }
        }
        private void btnLoad_Data_Sel_Click(object sender, EventArgs e)
        {
            try
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Shift" }, new string[] { txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbShift_Sel.Text });
                DGV_Data_Sel.DataSource = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess_Sel.Text , filter_str);
                sel_spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { txtItemCode_Sel.Text, cbProcess_Sel.Text }));
                sel_log_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess_Sel.Text , filter_str);
                //myCode.Disable_Sort_DGV(DGV_Data_Sel);
                myCode.DGV_Auto_Resize(DGV_Data_Sel);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void DGV_Data_Sel_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DGV_CellColumnDoubleClick((DataTable)DGV_Data_Sel.DataSource, DGV_Cycles_Data_Sel, e.ColumnIndex);
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            if (myVar_ECheck.confirm_mode)
            {
                btnLogin.Text = "Login";
                btnLogin.BackColor = this.BackColor;
                myVar_ECheck.confirm_mode = false;
            }
            else
            {
                Login frm_login = new Login();
                frm_login.ShowDialog();
            }
        }
        private void DGV_Cycles_Data_Sel_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (myVar_ECheck.confirm_mode)
                {
                    DGV_Cycles_Data_Sel.ContextMenuStrip = cmsCopyPaste;
                }
                else
                {
                    MessageBox.Show("Please, login at first!");
                    DGV_Cycles_Data_Sel.ReadOnly = true;
                }
            }
        }
        private void DGV_Cycles_Data_Sel_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewCell cur_cell = DGV_Cycles_Data_Sel.CurrentCell;
            sel_cell.after_val = cur_cell.Value.ToString();
            string USL = blNG_Details.Net_Spec_tbl.Rows[cur_cell.RowIndex][3].ToString();
            string LSL = blNG_Details.Net_Spec_tbl.Rows[cur_cell.RowIndex][2].ToString();
            double ref_rate = 0.1;
            if (sel_cell.after_val != sel_cell.before_val)
            {
                if (!myCode.IsNumeric(sel_cell.after_val))
                {
                    cur_cell.Value = sel_cell.before_val;
                }
                Color cell_color = myCode.check_in_limit2(USL, LSL, cur_cell.Value.ToString());
                if (cell_color == Color.White)
                {
                    if (cbProcess_Sel.Text == "CQRA_SURVIVAL_HOT_OIL")
                    {
                        ref_rate = 0.05;
                    }
                    string col_name = DGV_Cycles_Data_Sel.Columns[cur_cell.ColumnIndex].Name;
                    //cur_cell.Style.BackColor = proc_data.CheckCell_Variability_Resistance(ref_rate, (DataTable)DGV_Data_Sel.DataSource, cur_cell, col_name);
                    cur_cell.Style.BackColor = proc_data.CheckCell_Variability_Resistance_bending(ref_rate, (DataTable)DGV_Data_Sel.DataSource, cur_cell, col_name, sel_item.cur_Cycles, dic_index_lst);

                }
                else
                {
                    cur_cell.Style.BackColor = cell_color;
                }
            }
            if (!manual_en)
            {
                DGV_Cycles_Data_Sel.ReadOnly = true;
            }
        }
        private void DGV_Cycles_Data_Sel_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (myVar_ECheck.confirm_mode)
            {
                if (MessageBox.Show("Do you want to update?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    DataTable src_dt = (DataTable)DGV_Data_Sel.DataSource;
                    if (!manual_en)
                    {
                        List<string> Net_lst = new List<string>();

                        Net_lst = src_dt.AsEnumerable().Select(x => x.Field<string>("Net_No")).Distinct().ToList();
                        foreach (ECheck_Process.NG_list2 ng in blNG_Details.myNGlst2)
                        {
                            string cur_pcs = ng.col_val;// ng.col_val.Split('_')[1];
                            foreach (int r_inx in ng.lst_row_inx)
                            {
                                int ng_item_inx = DGV_Cycles_Data_Sel.Columns[ng.col_val].Index;
                                int src_row = r_inx;
                                int total_row = DGV_Cycles_Data_Sel.Rows.Count;
                                if (sel_item.cur_Cycles.Contains("Bending"))
                                {
                                    src_row = dic_index_lst.Keys.ToArray()[r_inx];
                                    total_row = Net_lst.Count;
                                }
                                int tar_row = ng_item_inx * total_row + src_row;
                                string val_before = DGV_Data_Sel.Rows[tar_row].Cells[sel_item.cur_Cycles].Value.ToString();
                                string val_after = DGV_Cycles_Data_Sel.Rows[r_inx].Cells[ng_item_inx].Value.ToString();
                                if (val_before != val_after)
                                {
                                    DGV_Data_Sel.Rows[tar_row].Cells[sel_item.cur_Cycles].Value = val_after;
                                    string id = (TDMK_Code.SQL_MAX("EDIT_HISTORY", "ID", sqlcon_OK2SHIP) + 1).ToString();
                                    string[] item = new string[] { "ID", "ItemCode", "LotNo", "Net_No", "Pcs_No", "Data_Before", "Data_After", "UserID", "Date_Modify", "Process", "Remark" };
                                    string[] item_val = new string[] { id, sel_item.cur_ItemCode, sel_item.cur_LotNo, (src_row + 1).ToString(), cur_pcs, val_before, val_after, myVar_ECheck.cur_User, DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), sel_item.cur_Process, sel_item.cur_Cycles };
                                    TDMK_Code.insert_val_arr2("EDIT_HISTORY", sqlcon_OK2SHIP, item, item_val);
                                }
                            }
                        }
                        foreach (var ng in Delta_R_NGlist)
                        {
                            string cur_pcs = ng.Key;// ng.col_val.Split('_')[1];
                            foreach (int r_inx in ng.Value)
                            {
                                int ng_item_inx = DGV_Cycles_Data_Sel.Columns[ng.Key].Index;
                                int src_row = r_inx;
                                int total_row = DGV_Cycles_Data_Sel.Rows.Count;
                                if (sel_item.cur_Cycles.Contains("Bending"))
                                {
                                    src_row = dic_index_lst.Keys.ToArray()[r_inx];
                                    total_row = Net_lst.Count;
                                }
                                int tar_row = ng_item_inx * total_row + src_row;
                                string val_before = DGV_Data_Sel.Rows[tar_row].Cells[sel_item.cur_Cycles].Value.ToString();
                                string val_after = DGV_Cycles_Data_Sel.Rows[r_inx].Cells[ng_item_inx].Value.ToString();
                                if (val_before != val_after)
                                {
                                    DGV_Data_Sel.Rows[tar_row].Cells[sel_item.cur_Cycles].Value = val_after;
                                    string id = (TDMK_Code.SQL_MAX("EDIT_HISTORY", "ID", sqlcon_OK2SHIP) + 1).ToString();
                                    string[] item = new string[] { "ID", "ItemCode", "LotNo", "Net_No", "Pcs_No", "Data_Before", "Data_After", "UserID", "Date_Modify", "Process", "Remark" };
                                    string[] item_val = new string[] { id, sel_item.cur_ItemCode, sel_item.cur_LotNo, (src_row + 1).ToString(), cur_pcs, val_before, val_after, myVar_ECheck.cur_User, DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), sel_item.cur_Process, sel_item.cur_Cycles };
                                    TDMK_Code.insert_val_arr2("EDIT_HISTORY", sqlcon_OK2SHIP, item, item_val);
                                }
                            }
                        }
                    }
                    else
                    {
                        proc_data.update_bending_table(sqlcon_OK2SHIP, src_dt, (DataTable)DGV_Cycles_Data_Sel.DataSource, sel_item.cur_ItemCode, sel_item.cur_Process, sel_item.cur_Cycles);
                    }
                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Shift" }, new string[] { sel_item.cur_ItemCode, sel_item.cur_LotNo, cbShift_Sel.Text });
                    DataTable temp_dt = (DataTable)DGV_Data_Sel.DataSource;
                    TDMK_Code.Delelte_FilteredItem_arr(sel_item.cur_Process, sqlcon_OK2SHIP, filter_str);
                    proc_data.BatchBulkCopy(sqlcon_OK2SHIP, temp_dt, sel_item.cur_Process);
                    btnLoad_Data_Sel.PerformClick();
                    blNG_Details.reset_all();
                    DGV_Cycles_Data_Sel.DataSource = null;
                }
            }
        }
        public void Summary_data()
        {
            DGV_Data.DataSource = null;
            if (lstLogFile.Items.Count > 0)
            {
                List<DataTable> tbl_data_lst = new List<DataTable>();
                if (Logfile_mode)
                {
                    foreach (string logfile in lstLogFile.Items)
                    {

                        string f_name = Path.Combine(txtLocation.Text, logfile);
                        DataTable netSpec_tbl = new DataTable();
                        DataTable result_data = new DataTable();
                        //int Net_no = (int)numNET.Value;
                        //string sel_dev = cbMachine.Text;
                        //try
                        //{
                        //    switch (sel_dev)
                        //    {
                        //        case "YAMAHA":
                        //            result_data = proc_data.Yamaha_3GSPD_Process(f_name, ref netSpec_tbl);
                        //            break;
                        //        case "TAIYO":
                        //            result_data = proc_data.Tayo_3GMRD_Process(f_name, Net_no, ref netSpec_tbl);
                        //            break;
                        //    }
                        //    tbl_data_lst.Add(result_data);
                        //}
                        //catch
                        //{
                        //    MessageBox.Show("Wrong device type!");
                        //}
                    }
                }
                else
                {
                    foreach (string logfile in lstLogFile.Items)
                    {
                        DataTable dt = new DataTable();
                        dt = proc_data.Load_Log_Data(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, cbCycles.Text, logfile);
                        if (dt.Rows.Count > 0)
                        {
                            tbl_data_lst.Add(dt);
                        }
                    }
                }
                DataTable final_tbl = new DataTable();
                if (tbl_data_lst.Count > 1)
                {
                    if (cbProcess.Text != "ELECTRICAL")
                    {
                        final_tbl = tbl_data_lst[0];
                        int c_inx = 0;
                        foreach (DataColumn dc in final_tbl.Columns)
                        {
                            int r_inx = 0;
                            foreach (DataRow dr in final_tbl.Rows)
                            {
                                string USL = DGV_NET_Spec.Rows[r_inx].Cells[3].Value.ToString();
                                string LSL = DGV_NET_Spec.Rows[r_inx].Cells[2].Value.ToString();
                                string cur_val = dr[dc].ToString();
                                if (!proc_data.check_in_limit(USL, LSL, cur_val))
                                {
                                    for (int i = 1; i < tbl_data_lst.Count; i++)
                                    {
                                        DataTable cur_dt = tbl_data_lst[i];
                                        if (proc_data.check_in_limit(USL, LSL, cur_dt.Rows[r_inx][c_inx].ToString()))
                                        {
                                            dr[dc] = cur_dt.Rows[r_inx][c_inx];
                                        }
                                    }
                                }
                                r_inx++;
                            }
                            c_inx++;
                        }
                    }
                    else
                    {
                        final_tbl = proc_data.Electric_Summary(sqlcon_OK2SHIP, tbl_data_lst, txtItemCode.Text, cbProcess.Text, 100);
                    }
                }
                else
                {
                    final_tbl = tbl_data_lst[0];
                }
                DGV_Data.DataSource = final_tbl;
                myCode.DGV_Auto_Resize(DGV_Data);
                if (DGV_NET_Spec.Rows.Count == DGV_Data.Rows.Count)
                {
                    DataTable tbl_NG_NET = new DataTable();
                    List<int> NET_index_lst = new List<int>();
                    List<string> item_NG_lst = new List<string>();
                    proc_data.Submit_Data(DGV_NET_Spec, DGV_Data, ref tbl_NG_NET, ref NET_index_lst, DGV_NET_NG, ref item_NG_lst, ref myNGlst2);
                    lstNG.DataSource = item_NG_lst;
                    txtTotal_NET.Text = NET_index_lst.Count.ToString();
                    txtTotalItem.Text = item_NG_lst.Count.ToString();
                }
                else
                {
                    MessageBox.Show("NET Spec Data is not enough!\r\nPlease, check NET Spec", "Warning");
                }
                lblNG_Detail.Text = lblNGDetail + "Summary Data";
                //btnDataSelect.Enabled = true;
            }
        }

        private void btnTest_Click_1(object sender, EventArgs e)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
            DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess_Sel.Text + "_LOGFILE", filter_str);
            List<DataTable> tbl_lst = new List<DataTable>();
            List<string> logfile_lst = new List<string>();
            proc_data.Get_List_data2(-1, src_tbl, new string[] { "ItemCode", "LotNo" }, ref logfile_lst, "Logfile", true);
            foreach (var log_f in logfile_lst)
            {
                DataTable temp = proc_data.Load_Log_Data(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, cbCycles.Text, log_f);
                tbl_lst.Add(temp);
            }
            DataTable dest_tbl = new DataTable();
            List<List<string>> src_NET_lst = new List<List<string>>();
            foreach (var tbl in tbl_lst)
            {
                int dr_inx = 0;
                foreach (DataRow dr in tbl.Rows)
                {
                    if (src_NET_lst.Count < dr_inx + 1)
                    {
                        src_NET_lst.Add(new List<string>());
                    }
                    foreach (DataColumn dc in tbl.Columns)
                    {
                        src_NET_lst[dr_inx].Add(dr[dc].ToString());
                    }
                    dr_inx++;
                }
            }
            for (int i = 0; i < src_NET_lst[0].Count; i++)
            {
                dest_tbl.Columns.Add("ItemNo_" + (i + 1).ToString());
            }
            foreach (var n_lst in src_NET_lst)
            {
                DataRow dr = dest_tbl.NewRow();
                int c_inx = 0;
                foreach (string t in n_lst)
                {
                    dr[c_inx] = t;
                    c_inx++;
                }
                dest_tbl.Rows.Add(dr);
            }
            MessageBox.Show("Fill datatable");
            DGV_LogFile.DataSource = dest_tbl;
        }

        private void DGV_LogFile_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //DataTable src_tbl = (DataTable)DGV_LogFile.DataSource;
            //Electrical_Details frmElectric = new Electrical_Details(src_tbl, txtItemCode.Text, cbProcess.Text, sqlcon_OK2SHIP);
            //frmElectric.Show();
        }

        private void DGV_LogFile_ColumnAdded(object sender, DataGridViewColumnEventArgs e)
        {

        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            string sheet_name = "";
            //string process_name = "";
            switch (cbProcess_Sel.Text)
            {
                case "FLEX_BENDING":
                    sheet_name = "Flex bending";
                    break;
                case "THERMAL_CYCLING_AND_BEND":
                    sheet_name = "Thermal Cycling & bending";
                    break;
                case "HEAT_SOAK_AND_BEND":
                    sheet_name = "Heat Soak & bending";
                    break;
            }
            if (RB_Mass.Checked)
            {
                //bending_proc.Export_Thermal_HeatSoak_Bend_MultiType(sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, sheet_name, 1,"MASS", cbShift_Sel.Text);
                bending_proc_Epplus.Export_Thermal_HeatSoak_Bend_MultiType(sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, sheet_name, 1, "MASS", cbShift_Sel.Text);
            }
            else
            {
                //bending_proc.Export_Thermal_HeatSoak_Bend_All(sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, sheet_name, 20, "NPI");
                bending_proc_Epplus.Export_Thermal_HeatSoak_Bend_All(sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, sheet_name, 20, "NPI");
            }
        }
        private void DGV_Data_Sel_DataSourceChanged(object sender, EventArgs e)
        {
            double ref_rate = 0.1;
            if (cbProcess_Sel.SelectedItem.ToString() == "CQRA_SURVIVAL_HOT_OIL")
            {
                ref_rate = 0.05;
                //proc_data.Check_Cycles_Data(DGV_Data_Sel, 0.05);

            }
            else
            {
                //proc_data.Check_Cycles_Data(DGV_Data_Sel, 0.1);
            }
            lblCycleDetail.BackColor = Color.FromArgb(255, 255, 192);
            proc_data.Check_Cycles_Data(DGV_Data_Sel, ref_rate);
        }

        private void txtLocation_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
                e.SuppressKeyPress = true;
                //if(txtLocation.Text != "")
                //{
                //    string tar_loc = txtLocation.Text.Replace(Environment.NewLine, "");
                //    DirectoryInfo di = new DirectoryInfo(tar_loc);
                //    if (di.Exists)
                //    {
                //        FileInfo[] files;
                //        switch (cbProcess.Text)
                //        {
                //            case "CQRA_HOT_OIL":
                //                files = di.GetFiles("*.xlsx");
                //                break;
                //            case "CQRA_BENDING":
                //                files = di.GetFiles("*.xlsx");
                //                break;
                //            case "CQRA_THERMAL_CYCLING_AND_BEND":
                //                files = di.GetFiles("*.xlsx");
                //                break;
                //            case "CQRA_HEAT_SOAK_AND_BEND":
                //                files = di.GetFiles("*.xlsx");
                //                break;
                //            case "CQRA_SURVIVAL_HOT_OIL":
                //                files = di.GetFiles("*.xlsx");
                //                break;
                //            default:
                //                files = di.GetFiles("*.csv");
                //                break;
                //        }
                //        List<string> _lstLogfile = new List<string>();
                //        foreach (FileInfo f in files)
                //        {
                //            if (!f.Name.Contains("$"))
                //            {
                //                _lstLogfile.Add(f.Name);
                //            }
                //        }
                //        lstLogFile.DataSource = _lstLogfile;
                //    }
                //    else
                //    {
                //        MessageBox.Show("Folder not existed!", "Warning");
                //    }
                //}
            }
        }

        private void ECheck_Main_old_FormClosed(object sender, FormClosedEventArgs e)
        {
            //myVar._frmMain.Show();
        }

        private void tsmCopy_Click(object sender, EventArgs e)
        {
            myCode.CopyToClipboard(DGV_Cycles_Data_Sel);
        }

        private void selectedColumnsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PasteClipboardValue(true, DGV_Cycles_Data_Sel);
        }
        private void PasteClipboardValue(bool _transpose, DataGridView tar_DGV)
        {
            //Show Error if no cell is selected
            if (tar_DGV.SelectedCells.Count == 0)
            {
                MessageBox.Show("Please select a cell", "Paste", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            //Get the satring Cell
            DataGridViewCell startCell = GetStartCell(tar_DGV);
            //Get the clipboard value in a dictionary
            Dictionary<int, Dictionary<int, string>> cbValue = ClipBoardValues(Clipboard.GetText());
            if (cbValue.Count > 0)
            {
                double ref_rate = 0.1;
                if (cbProcess_Sel.Text == "CQRA_SURVIVAL_HOT_OIL")
                {
                    ref_rate = 0.05;
                }
                if (!_transpose)
                {
                    int iRowIndex = startCell.RowIndex;
                    foreach (int rowKey in cbValue.Keys)
                    {
                        int iColIndex = startCell.ColumnIndex;
                        foreach (int cellKey in cbValue[rowKey].Keys)
                        {
                            //Check if the index is with in the limit
                            if (iColIndex <= tar_DGV.Columns.Count - 1 && iRowIndex <= tar_DGV.Rows.Count - 1)
                            {
                                DataGridViewCell curr_cell = tar_DGV[iColIndex, iRowIndex];
                                //Copy to selected cells if 'chkPasteToSelectedCells' is checked
                                if (curr_cell.Selected)
                                {
                                    string cur_col_1 = DGV_Cycles_Data_Sel.Columns[curr_cell.ColumnIndex].Name;
                                    string cur_col = cur_col_1.Split('_').Last();
                                    if (!manual_en)
                                    {
                                        sel_cell = new myVar_ECheck.edit_cell(cur_col_1, (curr_cell.RowIndex + 1).ToString(), curr_cell.Value.ToString());
                                        bool R_val_en = false;
                                        bool R_vary_en = false;
                                        int item_temp = blNG_Details.item_NG_lst.FindIndex(x => x == cur_col_1);
                                        if (item_temp != -1)
                                        {
                                            ECheck_Process.NG_list2 cur_NG = blNG_Details.myNGlst2[item_temp];
                                            int r_temp = cur_NG.lst_row_inx.FindIndex(x => x == curr_cell.RowIndex);
                                            if (r_temp != -1)
                                            {
                                                R_val_en = true;
                                            }
                                        }
                                        if (TDMK_Code.check_exist_list_index2(cur_col_1, Delta_R_NGlist.Keys.ToList()) != -1)
                                        {
                                            if (Delta_R_NGlist[cur_col_1].FindIndex(x => x == curr_cell.RowIndex) != -1)
                                            {
                                                R_vary_en = true;
                                            }
                                        }
                                        if (R_val_en || R_vary_en)
                                        {
                                            DGV_Cycles_Data_Sel.BeginEdit(true);
                                            curr_cell.Value = cbValue[rowKey][cellKey].Trim();
                                            curr_cell.Style.BackColor = proc_data.CheckCell_Variability_Resistance_bending(ref_rate, (DataTable)DGV_Data_Sel.DataSource, curr_cell, cur_col_1, sel_item.cur_Cycles, dic_index_lst);
                                        }
                                    }
                                    else
                                    {
                                        DGV_Cycles_Data_Sel.BeginEdit(true);
                                        curr_cell.Value = cbValue[rowKey][cellKey].Trim();
                                        curr_cell.Style.BackColor = proc_data.CheckCell_Variability_Resistance_bending(ref_rate, (DataTable)DGV_Data_Sel.DataSource, curr_cell, cur_col_1, sel_item.cur_Cycles, dic_index_lst);
                                    }
                                }
                            }
                            iColIndex++;
                        }
                        iRowIndex++;
                    }
                }
                else
                {
                    int iColIndex = startCell.ColumnIndex;
                    foreach (int rowKey in cbValue.Keys)
                    {
                        int iRowIndex = startCell.RowIndex;
                        foreach (int cellKey in cbValue[rowKey].Keys)
                        {
                            //Check if the index is with in the limit
                            if (iColIndex <= tar_DGV.Columns.Count - 1 && iRowIndex <= tar_DGV.Rows.Count - 1)
                            {
                                DataGridViewCell curr_cell = tar_DGV[iColIndex, iRowIndex];
                                //Copy to selected cells if 'chkPasteToSelectedCells' is checked
                                string cur_col_1 = DGV_Cycles_Data_Sel.Columns[curr_cell.ColumnIndex].Name;
                                string cur_col = cur_col_1.Split('_').Last();
                                if (curr_cell.Selected)
                                {
                                    if (!manual_en)
                                    {

                                        sel_cell = new myVar_ECheck.edit_cell(cur_col_1, (curr_cell.RowIndex + 1).ToString(), curr_cell.Value.ToString());
                                        bool R_val_en = false;
                                        bool R_vary_en = false;
                                        int item_temp = blNG_Details.item_NG_lst.FindIndex(x => x == cur_col_1);
                                        if (item_temp != -1)
                                        {
                                            ECheck_Process.NG_list2 cur_NG = blNG_Details.myNGlst2[item_temp];
                                            int r_temp = cur_NG.lst_row_inx.FindIndex(x => x == curr_cell.RowIndex);
                                            if (r_temp != -1)
                                            {
                                                R_val_en = true;
                                            }
                                        }
                                        if (TDMK_Code.check_exist_list_index2(cur_col_1, Delta_R_NGlist.Keys.ToList()) != -1)
                                        {
                                            if (Delta_R_NGlist[cur_col_1].FindIndex(x => x == curr_cell.RowIndex) != -1)
                                            {
                                                R_vary_en = true;
                                            }
                                        }
                                        if (R_val_en || R_vary_en)
                                        {
                                            DGV_Cycles_Data_Sel.BeginEdit(true);
                                            curr_cell.Value = cbValue[rowKey][cellKey].Trim();
                                            curr_cell.Style.BackColor = proc_data.CheckCell_Variability_Resistance_bending(ref_rate, (DataTable)DGV_Data_Sel.DataSource, curr_cell, cur_col_1, sel_item.cur_Cycles, dic_index_lst);
                                        }
                                    }
                                    else
                                    {
                                        DGV_Cycles_Data_Sel.BeginEdit(true);
                                        curr_cell.Value = cbValue[rowKey][cellKey].Trim();
                                        curr_cell.Style.BackColor = proc_data.CheckCell_Variability_Resistance_bending(ref_rate, (DataTable)DGV_Data_Sel.DataSource, curr_cell, cur_col_1, sel_item.cur_Cycles, dic_index_lst);
                                    }
                                }
                            }
                            iRowIndex++;
                        }
                        iColIndex++;
                    }
                }
            }
            else
            {
                MessageBox.Show("PLease, select data!", "Warning");
            }
        }
        private DataGridViewCell GetStartCell(DataGridView dgView)
        {
            //get the smallest row,column index
            if (dgView.SelectedCells.Count == 0)
                return null;

            int rowIndex = dgView.Rows.Count - 1;
            int colIndex = dgView.Columns.Count - 1;

            foreach (DataGridViewCell dgvCell in dgView.SelectedCells)
            {
                if (dgvCell.RowIndex < rowIndex)
                    rowIndex = dgvCell.RowIndex;
                if (dgvCell.ColumnIndex < colIndex)
                    colIndex = dgvCell.ColumnIndex;
            }

            return dgView[colIndex, rowIndex];
        }
        private Dictionary<int, Dictionary<int, string>> ClipBoardValues(string clipboardValue)
        {
            Dictionary<int, Dictionary<int, string>> copyValues = new Dictionary<int, Dictionary<int, string>>();

            String[] lines = clipboardValue.Split('\n');

            for (int i = 0; i <= lines.Length - 1; i++)
            {
                copyValues[i] = new Dictionary<int, string>();
                String[] lineContent = lines[i].Split('\t');

                //if an empty cell value copied, then set the dictionay with an empty string
                //else Set value to dictionary
                if (lineContent.Length == 0)
                    copyValues[i][0] = string.Empty;
                else
                {
                    for (int j = 0; j <= lineContent.Length - 1; j++)
                        copyValues[i][j] = lineContent[j];
                }
            }
            return copyValues;
        }

        private void tsmPasteRows_Click(object sender, EventArgs e)
        {
            PasteClipboardValue(false, DGV_Cycles_Data_Sel);
        }

        private void DGV_Cycles_Data_Sel_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if ((e.RowIndex > -1) && (e.ColumnIndex > -1))
            {
                if (!manual_en)
                {
                    if (myVar_ECheck.confirm_mode)
                    {
                        DGV_Cycles_Data_Sel.ReadOnly = false;
                        DataGridViewCell cur_cell = DGV_Cycles_Data_Sel.CurrentCell;
                        string cur_col_1 = DGV_Cycles_Data_Sel.Columns[cur_cell.ColumnIndex].Name;
                        string cur_col = cur_col_1.Split('_').Last();
                        sel_cell = new myVar_ECheck.edit_cell(cur_col_1, (cur_cell.RowIndex + 1).ToString(), cur_cell.Value.ToString());
                        bool R_val_en = false;
                        bool R_vary_en = false;
                        int item_temp = blNG_Details.item_NG_lst.FindIndex(x => x == cur_col_1);
                        if (item_temp != -1)
                        {
                            ECheck_Process.NG_list2 cur_NG = blNG_Details.myNGlst2[item_temp];
                            int r_temp = cur_NG.lst_row_inx.FindIndex(x => x == cur_cell.RowIndex);
                            if (r_temp != -1)
                            {
                                R_val_en = true;
                            }
                        }
                        if (TDMK_Code.check_exist_list_index2(cur_col_1, Delta_R_NGlist.Keys.ToList()) != -1)
                        {
                            if (Delta_R_NGlist[cur_col_1].FindIndex(x => x == cur_cell.RowIndex) != -1)
                            {
                                R_vary_en = true;
                            }
                        }
                        if (R_vary_en || R_val_en)
                        {
                            DGV_Cycles_Data_Sel.BeginEdit(true);
                        }
                        else
                        {
                            MessageBox.Show("Cannot Edit", "Warning");
                            DGV_Cycles_Data_Sel.ReadOnly = true;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please, login at first!", "Warning");
                        DGV_Cycles_Data_Sel.ReadOnly = true;
                    }
                }
            }
        }

        private void DGV_Cycles_Data_Sel_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            int r_inx = e.RowIndex;
            int c_inx = e.ColumnIndex;
            if ((r_inx > -1) && (c_inx > -1))
            {
                e.ToolTipText = proc_data.Details_Variability_Resistance_bending(r_inx, c_inx, (DataTable)DGV_Data_Sel.DataSource, (DataTable)DGV_Cycles_Data_Sel.DataSource, sel_item.cur_Cycles, dic_index_lst);
            }
        }
        public void DGV_CellColumnDoubleClick(DataTable cur_tbl, DataGridView tar_DGV, int ColumnIndex)
        {
            try
            {
                double ref_rate = 0.1;
                manual_en = false;
                if (cbProcess_Sel.Text == "CQRA_SURVIVAL_HOT_OIL")
                {
                    ref_rate = 0.05;
                }
                List<string> col_lst = new List<string>() { "ItemCode", "LotNo", "Net_No", "Pcs_No" };
                List<string> mycol_lst = new List<string>();
                foreach (DataColumn dc in cur_tbl.Columns)
                {
                    if (dc.ColumnName.Contains("Before") || dc.ColumnName.Contains("After") || dc.ColumnName.Contains("Data"))
                    {
                        mycol_lst.Add(dc.ColumnName);
                    }
                }
                string sel_col_val = cur_tbl.Columns[ColumnIndex].ColumnName;
                sel_item = new myVar_ECheck.item_info(txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, sel_col_val);
                lblCycleDetail.Text = "Details Cycles Data: " + sel_col_val;
                blNG_Details.reset_all();
                if (TDMK_Code.check_exist_list_index(sel_col_val, mycol_lst) != -1)
                {
                    col_lst.Add(sel_col_val);
                    DataTable filter_tbl = cur_tbl.AsDataView().ToTable(false, col_lst.ToArray());
                    DataTable temp_spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { txtItemCode_Sel.Text, cbProcess_Sel.Text }));
                    DataTable spec_dt = temp_spec_dt.AsDataView().ToTable(false, new string[] { "Point+V", "Point-V", "LSL", "USL" });
                    dic_index_lst = new SortedDictionary<int, string>();
                    int inx = 0;
                    foreach (DataRow dr in temp_spec_dt.Rows)
                    {
                        if (myCode.checkDBNull(dr["Sel_Report"]) != "")
                        {
                            string cur_net = dr["Sel_Report"].ToString() + "(" + dr["Point+V"].ToString() + "_" + dr["Point-V"].ToString() + ")";
                            dic_index_lst.Add(inx, cur_net);
                        }
                        inx++;
                    }
                    List<int> index_lst = dic_index_lst.Keys.ToList();
                    if (sel_col_val.Contains("Bending"))
                    {
                        if (index_lst.Count > 0)
                        {
                            DataTable _filter_tbl = cur_tbl.AsEnumerable().Where(x => index_lst.Contains(Convert.ToInt32(x.Field<string>("Net_no")) - 1)).CopyToDataTable();
                            filter_tbl = _filter_tbl.AsDataView().ToTable(false, col_lst.ToArray());
                            DataTable _spec_dt = proc_data.Filter_table_by_index(temp_spec_dt, index_lst);
                            spec_dt = _spec_dt.AsDataView().ToTable(false, new string[] { "Point+V", "Point-V", "LSL", "USL" });
                            List<string> cur_col_data = cur_tbl.AsEnumerable().Where(x => index_lst.Contains(Convert.ToInt32(x.Field<string>("Net_No")) - 1) && x.Field<string>(sel_col_val) != null).Select(x => x.Field<string>(sel_col_val)).ToList();
                            if (cur_col_data.Count == 0)
                            {
                                DataTable bending_tbl = new DataTable();
                                List<string> col_name_lst = cur_tbl.AsEnumerable().Select(x => x.Field<string>("Pcs_No")).Distinct().ToList();
                                foreach (var col in col_name_lst)
                                {
                                    bending_tbl.Columns.Add(col);
                                }
                                for (int i = 0; i < index_lst.Count; i++)
                                {
                                    bending_tbl.Rows.Add();
                                }
                                tar_DGV.DataSource = bending_tbl;
                                for (int i = 0; i < index_lst.Count; i++)
                                {
                                    tar_DGV.Rows[i].HeaderCell.Value = dic_index_lst[index_lst[i]];
                                }
                                tar_DGV.ReadOnly = false;
                                tar_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                                blNG_Details.fill_data(new DataTable(), spec_dt);
                                manual_en = true;
                                return;
                            }
                        }
                        else
                        {
                            MessageBox.Show("Please, setup Net for process: " + cbProcess_Sel.Text, "Warning");
                            return;
                        }

                    }
                    DataTable tbl_detail = proc_data.Detail_NET_PCS(filter_tbl);
                    tar_DGV.DataSource = tbl_detail;
                    blNG_Details.fill_data(tbl_detail, spec_dt);
                    proc_data.check_in_spec_tbl(spec_dt, tar_DGV);
                    if (!sel_col_val.Contains("Bending"))
                    {
                        myCode.DGV_Auto_Resize(tar_DGV);
                    }
                    else
                    {
                        for (int i = 0; i < index_lst.Count; i++)
                        {
                            tar_DGV.Rows[i].HeaderCell.Value = dic_index_lst[index_lst[i]];
                        }
                        myCode.Disable_Sort_DGV(tar_DGV);
                        tar_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                    }
                    sel_item = new myVar_ECheck.item_info(txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, sel_col_val);
                    if (myCode.check_columns_existed(cur_tbl, "Before"))
                    {
                        List<string> bef_col_lst = new List<string>() { "ItemCode", "LotNo", "Net_No", "Pcs_No", "Before" };
                        DataTable before_tbl = cur_tbl.AsDataView().ToTable(false, bef_col_lst.ToArray());
                        if (sel_col_val.Contains("Bending"))
                        {
                            DataTable _filter_tbl = cur_tbl.AsEnumerable().Where(x => index_lst.Contains(Convert.ToInt32(x.Field<string>("Net_no")) - 1)).CopyToDataTable();
                            before_tbl = _filter_tbl.AsDataView().ToTable(false, bef_col_lst.ToArray());
                        }
                        DataTable bef_tbl_detail = proc_data.Detail_NET_PCS(before_tbl);
                        Delta_R_NGlist = new Dictionary<string, List<int>>();
                        int c_inx = 0;
                        foreach (DataColumn dc in bef_tbl_detail.Columns)
                        {
                            int r_inx = 0;
                            bool col_NG = false;
                            foreach (DataRow dr in bef_tbl_detail.Rows)
                            {
                                DataGridViewCell cur_cell = DGV_Cycles_Data_Sel.Rows[r_inx].Cells[c_inx];
                                bool add_en = false;
                                if (myCode.IsNumeric(cur_cell.Value.ToString()))
                                {
                                    double cur_val = Convert.ToDouble(dr[dc]);
                                    double cyc_val = Convert.ToDouble(cur_cell.Value);
                                    double rate = Math.Abs(cur_val - cyc_val) / cur_val;
                                    if (rate > ref_rate)
                                    {
                                        cur_cell.Style.BackColor = Color.Red;
                                        add_en = true;
                                        col_NG = true;
                                    }
                                }
                                else
                                {
                                    cur_cell.Style.BackColor = Color.LightPink;
                                    add_en = true;
                                    col_NG = true;
                                }
                                if (add_en)
                                {
                                    string col_name = dc.ColumnName;
                                    int r_no = r_inx;// + 1;
                                    if (TDMK_Code.check_exist_list_index2(col_name, Delta_R_NGlist.Keys.ToList()) == -1)
                                    {
                                        Delta_R_NGlist.Add(col_name, new List<int>());
                                    }
                                    Delta_R_NGlist[col_name].Add(r_no);
                                }
                                r_inx++;
                            }
                            if (col_NG)
                            {
                                DGV_Cycles_Data_Sel.EnableHeadersVisualStyles = false;
                                DGV_Cycles_Data_Sel.Columns[c_inx].HeaderCell.Style.BackColor = Color.HotPink;
                            }
                            else
                            {
                                DGV_Cycles_Data_Sel.EnableHeadersVisualStyles = false;
                                DGV_Cycles_Data_Sel.Columns[c_inx].HeaderCell.Style.BackColor = Color.White;
                            }

                            c_inx++;
                        }
                    }
                    bool data_ok_status = true;
                    foreach (DataGridViewColumn dgv_c in DGV_Cycles_Data_Sel.Columns)
                    {
                        foreach (DataGridViewRow dgv_r in DGV_Cycles_Data_Sel.Rows)
                        {
                            if (dgv_r.Cells[dgv_c.Index].Style.BackColor != Color.White)
                            {

                                data_ok_status = false;
                                break;
                            }
                        }
                    }
                    if (data_ok_status)
                    {
                        lblCycleDetail.BackColor = Color.Green;
                    }
                    else
                    {
                        lblCycleDetail.BackColor = Color.Red;
                    }
                }
            }
            catch
            {

            }
        }
        private void txtLotNo_Validated(object sender, EventArgs e)
        {
            txtLotNo.Text = proc_data.Lotno_Formated(txtLotNo.Text);
        }

        private void txtLotNo_Sel_Validated(object sender, EventArgs e)
        {
            txtLotNo_Sel.Text = proc_data.Lotno_Formated(txtLotNo_Sel.Text);
        }

        private void txtLocation_Validated(object sender, EventArgs e)
        {

        }

        private void txtLocation_TextChanged(object sender, EventArgs e)
        {

        }
        private string switchCategory(string category)
        {
            switch (category)
            {
                case "THERMAL_CYCLING":
                    return "TC";
                case "FLEX_BENDING":
                    return "F";
                case "HEAT_SOAK_AND_BEND":
                    return "HF";
                case "THERMAL_CYCLING_AND_BEND":
                    return "TF";
                case "THERMAL_SHOCK":
                    return "TS";
                case "HEAT_SOAK":
                    return "HS";
                default:
                    return "EXCEPTION";
            }
        }
        private string switchCycle(string cycle)
        {
            if(cycle == "Before")
            {
                return "BF";
            }
            if (cycle.Contains("After_"))
            {
                return cycle.Replace("After_", "L");
            }
            Debugger.Break();
            return "ERROR";
        }
        private void _txtLocation_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                if (txtLocation.Text != "")
                {
                    string tar_loc = txtLocation.Text.Replace(Environment.NewLine, "");
                    string category = switchCategory(cbProcess.Text);
                    string cycle = switchCycle(cbCycles.Text);
                    tar_loc = $"{tar_loc}\\{category}\\{cycle}";
             
                    DirectoryInfo di = new DirectoryInfo(tar_loc);
                    if (di.Exists)
                    {
                        FileInfo[] files;
                        files = di.GetFiles("*.DAT");
                        List<string> _lstLogfile = files.Select(x => x.Name).ToList();
                        lstLogFile.DataSource = _lstLogfile;
                    }
                    else
                    {
                        MessageBox.Show("Folder not existed!", "Warning");
                    }
                }
            }
        }
        private void Bending_Main_Resize(object sender, EventArgs e)
        {
            rs.ResizeAllControls(this);
            blNG_Details.resize();
        }

        private void btnSaveAll_Click(object sender, EventArgs e)
        {
            //bending_proc.Save_Log(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, (List<string>)lstLogFile.DataSource, txtLocation.Text);
            //bending_proc.Summary_Logfile_All(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.Text);
            bending_proc.Save_Log(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, cbShift.Text, (List<string>)lstLogFile.DataSource, txtLocation.Text);
            bending_proc.Summary_Logfile_All(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, cbShift.Text, Convert.ToInt32(numQty.Value));
        }

        private void DGV_Data_Sel_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int r_inx = e.RowIndex;
            int c_inx = e.ColumnIndex;
            if (r_inx != -1 && c_inx != -1)
            {
                string col_name = DGV_Data_Sel.Columns[c_inx].Name;
                if (col_name == "Before")
                {
                    DataGridViewCell cur_cell = DGV_Data_Sel.CurrentCell;
                    string NetNo = DGV_Data_Sel.Rows[r_inx].Cells["Net_No"].Value.ToString();
                    int spec_r_inx = Convert.ToInt32(NetNo) - 1;
                    string usl = sel_spec_dt.Rows[spec_r_inx]["USL"].ToString();
                    string lsl = sel_spec_dt.Rows[spec_r_inx]["LSL"].ToString();
                    sel_before_Cell = new Bending_SMT_Lib.Cell_data_details(cur_cell.Value.ToString(), "", NetNo, usl, lsl);
                    DGV_Data_Sel.BeginEdit(true);
                }
            }
            else
            {
                if (r_inx == -1 && c_inx == -1)
                {
                    DataTable src_dt = (DataTable)DGV_Data_Sel.DataSource;
                    dic_updated_bef = new Dictionary<string, List<Bending_SMT_Lib.Cell_data_details>>();
                    foreach (DataGridViewRow dr in DGV_Data_Sel.Rows)
                    {
                        int r = dr.Index;
                        string NetNo = DGV_Data_Sel.Rows[r].Cells["Net_No"].Value.ToString();
                        string pcs = DGV_Data_Sel.Rows[r].Cells["Pcs_No"].Value.ToString();
                        int spec_r_inx = Convert.ToInt32(NetNo) - 1;
                        string usl = sel_spec_dt.Rows[spec_r_inx]["USL"].ToString();
                        string lsl = sel_spec_dt.Rows[spec_r_inx]["LSL"].ToString();
                        string Cell_bef_val = dr.Cells["Before"].Value.ToString();
                        List<string> after_lst = new List<string>();
                        List<string> before_lst = new List<string>();
                        foreach (DataGridViewColumn dc in DGV_Data_Sel.Columns)
                        {
                            string col = dc.Name;
                            if (col == "Before")
                            {
                                before_lst = logfile_data_lst(r, dc.Index, src_dt);
                            }
                            if (col.Contains("After"))
                            {
                                if (myCode.checkDBNull(dr.Cells[col].Value) != "")
                                {
                                    after_lst.Add(dr.Cells[col].Value.ToString());
                                }
                            }
                        }
                        Dictionary<string, int> R_vary_NG_lst = new Dictionary<string, int>();
                        foreach (string bef in before_lst)
                        {
                            int sum = 0;
                            foreach (string aft in after_lst)
                            {
                                double bef_val = Convert.ToDouble(bef);
                                double aft_val = Convert.ToDouble(aft);
                                double vary = Math.Abs(bef_val - aft_val) * 100 / bef_val;
                                if (vary > 10)
                                {
                                    sum++;
                                }
                            }
                            R_vary_NG_lst.Add(bef, sum);
                        }
                        string tar_bef = R_vary_NG_lst.Where(x => x.Value == R_vary_NG_lst.Values.Min()).First().Key;
                        if (Cell_bef_val != tar_bef)
                        {
                            //dr.Cells["Before"].Value = tar_bef;
                            src_dt.Rows[r]["Before"] = tar_bef;
                            Bending_SMT_Lib.Cell_data_details sel_bef_update = new Bending_SMT_Lib.Cell_data_details(Cell_bef_val, tar_bef, NetNo, usl, lsl);
                            if (dic_updated_bef.Keys.ToList().IndexOf(pcs) == -1)
                            {
                                dic_updated_bef.Add(pcs, new List<Bending_SMT_Lib.Cell_data_details> { sel_bef_update });
                            }
                            else
                            {
                                dic_updated_bef[pcs].Add(sel_bef_update);
                            }
                        }
                    }
                    proc_data.Check_Cycles_Data(DGV_Data_Sel, 0.1);
                    MessageBox.Show(new Form { TopMost = true }, "Hoàn thành chọn dữ liệu Before tốt nhất", "Thông báo");
                }
            }
        }

        private void DGV_Data_Sel_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewCell cur_cell = DGV_Data_Sel.CurrentCell;
            //List<string> log_data_lst = logfile_data_lst(cur_cell.RowIndex, cur_cell.ColumnIndex, (DataTable)DGV_Data_Sel.DataSource);
            //string cell_val = cur_cell.Value.ToString();
            //if(log_data_lst.IndexOf(cell_val)==-1)
            //{
            //    MessageBox.Show("Chọn đúng dữ liệu từ log file đã đo", "Thông báo");
            //    cur_cell.Value = sel_before_Cell.before_edit_data;
            //}
            //else
            //{

            //}
            sel_before_Cell.after_edit_data = cur_cell.Value.ToString();
            double ref_rate = 0.1;
            if (sel_before_Cell.after_edit_data != sel_before_Cell.before_edit_data)
            {
                if (!myCode.IsNumeric(sel_before_Cell.after_edit_data))
                {
                    cur_cell.Value = sel_before_Cell.before_edit_data;
                }
                else
                {
                    int r_inx = cur_cell.RowIndex;
                    string pcs = DGV_Data_Sel.Rows[r_inx].Cells["Pcs_No"].Value.ToString();
                    int inx = dic_updated_bef.Keys.ToList().IndexOf(pcs);
                    if (inx == -1)
                    {
                        dic_updated_bef.Add(pcs, new List<Bending_SMT_Lib.Cell_data_details> { sel_before_Cell });
                    }
                    else
                    {
                        dic_updated_bef[pcs].Add(sel_before_Cell);
                    }
                }
                cur_cell.Style.BackColor = myCode.check_in_limit2(sel_before_Cell.USL, sel_before_Cell.LSL, cur_cell.Value.ToString());
                for (int c_inx = cur_cell.ColumnIndex + 1; c_inx < DGV_Data_Sel.Columns.Count; c_inx++)
                {
                    DataGridViewCell cell_checked = DGV_Data_Sel.Rows[cur_cell.RowIndex].Cells[c_inx];
                    if (myCode.checkDBNull(cell_checked.Value) != "")
                    {
                        double _cur_val = Convert.ToDouble(cur_cell.Value);
                        double _sel_val = Convert.ToDouble(DGV_Data_Sel.Rows[cur_cell.RowIndex].Cells[c_inx].Value);
                        double rate = Math.Abs(_cur_val - _sel_val) / _sel_val;
                        if (rate < ref_rate)
                        {
                            cell_checked.Style.BackColor = Color.White;
                        }
                        else
                        {
                            cell_checked.Style.BackColor = Color.Red;
                        }
                    }
                }
            }
            //if (!manual_en)
            //{
            //    DGV_Data_Sel.ReadOnly = true;
            //}
        }

        private void DGV_Data_Sel_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            int r_inx = e.RowIndex;
            int c_inx = e.ColumnIndex;
            if ((r_inx > -1) && (c_inx > -1))
            {
                string col_name = DGV_Data_Sel.Columns[c_inx].Name;
                if (col_name.Contains("Before") || col_name.Contains("After"))
                {
                    string itemcode = DGV_Data_Sel.Rows[r_inx].Cells["ItemCode"].Value.ToString();
                    string lotno = DGV_Data_Sel.Rows[r_inx].Cells["LotNo"].Value.ToString();
                    string pcs = DGV_Data_Sel.Rows[r_inx].Cells["Pcs_No"].Value.ToString();
                    string Net_no = DGV_Data_Sel.Rows[r_inx].Cells["Net_No"].Value.ToString();
                    int net_inx = Convert.ToInt32(Net_no) - 1;
                    string Net_name = sel_spec_dt.Rows[net_inx]["Net_Name"].ToString();
                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Net_No", "Pcs_No", "Cycles_name" }, new string[] { itemcode, lotno, Net_name, pcs, col_name });
                    DataView dv = sel_log_dt.AsDataView();
                    dv.RowFilter = filter_str;
                    DataTable log_dt = dv.ToTable();
                    string log_before = "";
                    foreach (DataRow dr in log_dt.Rows)
                    {
                        log_before += dr["Data"].ToString() + " @" + dr["Logfile"].ToString() + "\r\n";
                    }
                    e.ToolTipText = log_before.Trim();
                }
            }
        }
        public List<string> logfile_data_lst(int r_inx, int c_inx, DataTable src_data)
        {
            List<string> result = new List<string>();
            string col_name = src_data.Columns[c_inx].ColumnName;
            string itemcode = src_data.Rows[r_inx]["ItemCode"].ToString();
            string lotno = src_data.Rows[r_inx]["LotNo"].ToString();
            string pcs = src_data.Rows[r_inx]["Pcs_No"].ToString();
            string Net_no = src_data.Rows[r_inx]["Net_No"].ToString();
            int net_inx = Convert.ToInt32(Net_no) - 1;
            string Net_name = sel_spec_dt.Rows[net_inx]["Net_Name"].ToString();
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Net_No", "Pcs_No", "Cycles_name" }, new string[] { itemcode, lotno, Net_name, pcs, col_name });
            DataView dv = sel_log_dt.AsDataView();
            dv.RowFilter = filter_str;
            DataTable log_dt = dv.ToTable();
            result = log_dt.AsEnumerable().Select(x => x.Field<string>("Data")).Distinct().ToList();
            return result;
        }

        private void DGV_Data_Sel_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (myVar_ECheck.confirm_mode)
                {
                    DGV_Data_Sel.ContextMenuStrip = cmsUpdate;
                }
                else
                {
                    MessageBox.Show("Please, login at first!");
                }
            }
        }

        private void tsmUpdateData_Click(object sender, EventArgs e)
        {
            if (myVar_ECheck.confirm_mode)
            {
                if (MessageBox.Show("Do you want to update?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    if (!manual_en)
                    {
                        foreach (var t in dic_updated_bef)
                        {
                            string cur_pcs = t.Key;
                            List<Bending_SMT_Lib.Cell_data_details> sel_updated_lst = t.Value;
                            int id = TDMK_Code.SQL_MAX("EDIT_HISTORY", "ID", sqlcon_OK2SHIP);
                            foreach (var cur_data in sel_updated_lst)
                            {
                                id++;
                                string net = cur_data.Net_No;
                                string val_after = cur_data.after_edit_data;
                                string val_before = cur_data.before_edit_data;
                                string[] item = new string[] { "ID", "ItemCode", "LotNo", "Net_No", "Pcs_No", "Data_Before", "Data_After", "UserID", "Date_Modify", "Process", "Remark" };
                                string[] item_val = new string[] { id.ToString(), txtItemCode_Sel.Text, txtLotNo_Sel.Text, net, cur_pcs, val_before, val_after, myVar_ECheck.cur_User, DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), sel_item.cur_Process, "Before" };
                                TDMK_Code.insert_val_arr2("EDIT_HISTORY", sqlcon_OK2SHIP, item, item_val);
                            }
                        }
                    }
                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode_Sel.Text, txtLotNo_Sel.Text });
                    DataTable temp_dt = (DataTable)DGV_Data_Sel.DataSource;
                    TDMK_Code.Delelte_FilteredItem_arr(sel_item.cur_Process, sqlcon_OK2SHIP, filter_str);
                    proc_data.BatchBulkCopy(sqlcon_OK2SHIP, temp_dt, sel_item.cur_Process);
                    btnLoad_Data_Sel.PerformClick();
                    blNG_Details.reset_all();
                    DGV_Cycles_Data_Sel.DataSource = null;
                }
            }
        }

        private void cbShift_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbShift.SelectedIndex != -1)
            {
                reset_all();
            }
        }

        private void tapMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            string tar_tab = tapMain.SelectedTab.Text;
            if (tar_tab == "Selection")
            {
                txtItemCode_Sel.Text = txtItemCode.Text;
                txtLotNo_Sel.Text = txtLotNo.Text;
                cbShift_Sel.SelectedIndex = cbShift.SelectedIndex;
                cbProcess_Sel.SelectedIndex = cbProcess.SelectedIndex;
            }
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void txtTotal_NET_TextChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void txtTotalItem_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
