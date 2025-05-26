using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
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
using OK2SHIP_Software;
using OfficeOpenXml;
using TDMK_EPPLUS_7;
using System.Diagnostics;
namespace Echeck_LogFile_Process
{
    public partial class ECheck_Main : Form
    {
        SEI_Lib myCode = new SEI_Lib();
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        ECheck_Process proc_data = new ECheck_Process();
        myVar exp_proc = new myVar();
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
        List<string> export_pcs_lst = new List<string>();
        DataTable temp_pcs_sel_dt = new DataTable();
        Echeck_Export Echeck_export_EPPlus = new Echeck_Export();
        TDMK_EPPLUS7_lib Excel_lib = new TDMK_EPPLUS7_lib();
        ExportEPPlus F_exportEPPlus = new ExportEPPlus();
        //public struct NG_list
        //{
        //    public int row_inx { get; set; }
        //    public string col_val { get; set; }
        //    public NG_list(int _rinx, string _col)
        //    {
        //        row_inx = _rinx;
        //        col_val = _col;
        //    }               
        //}
        public ECheck_Main()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            List<string> tar_process = new List<string>() { "ELECTRICAL", "CQRA_REFLOW", "CQRA_HOT_OIL", "CQRA_THERMAL_CYCLING", "CQRA_HEAT_SOAK", "CQRA_THERMAL_SHOCK", "CQRA_BENDING", "CQRA_THERMAL_CYCLING_AND_BEND", "CQRA_HEAT_SOAK_AND_BEND", "CQRA_SURVIVAL_REFLOW", "CQRA_SURVIVAL_HOT_OIL","CQRA_THERMAL_CYCLING_ON_COUPON" };
            cbProcess.DataSource = tar_process;
            cbProcess_Sel.DataSource = tar_process;
            cbProcess_pcs.DataSource = tar_process;
            cbProcess_NetSel.DataSource = tar_process;
            cbMachine.SelectedIndex = 0;
            sqlcon_OK2SHIP = exp_proc.initial_data(myVar.sel_DB,true);
            myVar_ECheck.frmMain = myVar._frmECheck;
            rs.FindAllControls(this);
            blNG_Details.form_load();
        }

        private void cbProcess_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> cycles_lst = new List<string>();
            string app_path = Application.StartupPath.Replace("\\FPCA OK2SHIP Auto System\\Bending_Items", "");
            string tar_file = Path.Combine(app_path, "Config", cbProcess.Text + ".txt");
            cycles_lst = myCode.read_config_arr(tar_file).ToList();
            cbCycles.Text = "";
            cbCycles.DataSource = cycles_lst;
            reset_all();
            switch (cbProcess.Text)
            {
                case "CQRA_HOT_OIL":
                    myVar_ECheck.CSV_en = false;
                    cbCycles.Enabled = false;
                    break;
                case "CQRA_BENDING":
                    myVar_ECheck.CSV_en = false;
                    cbCycles.Enabled = false;
                    break;
                case "CQRA_THERMAL_CYCLING_AND_BEND":
                    myVar_ECheck.CSV_en = false;
                    break;
                case "CQRA_HEAT_SOAK_AND_BEND":
                    myVar_ECheck.CSV_en = false;
                    break;
                case "CQRA_SURVIVAL_HOT_OIL":
                    myVar_ECheck.CSV_en = false;
                    cbCycles.Enabled = false;
                    break;
                case "CQRA_THERMAL_CYCLING_ON_COUPON":
                    myVar_ECheck.CSV_en = false;
                    cbCycles.Enabled = false;
                    //btnSave.Enabled = false;
                    break;
                default:
                    myVar_ECheck.CSV_en = true;
                    break;
            }

            //if (cbProcess.Text != "CQRA_HOT_OIL")
            //{
            //    myVar_ECheck.CSV_en = true;
            //}
            //else
            //{
            //    myVar_ECheck.CSV_en = false;
            //}
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
                        files = di.GetFiles("*.xlsx");
                        // = di.GetFiles("*.csv");
                        break;
                    case "CQRA_HEAT_SOAK_AND_BEND":
                        files = di.GetFiles("*.xlsx");
                        //files = di.GetFiles("*.csv");
                        break;
                    case "CQRA_SURVIVAL_HOT_OIL":
                        files = di.GetFiles("*.xlsx");
                        break;
                    case "CQRA_THERMAL_CYCLING_ON_COUPON":
                        files = di.GetFiles("*.xlsx");
                        break;
                    default:
                        files = di.GetFiles("*.csv");
                        break;
                }
                //if (cbProcess.Text != "CQRA_HOT_OIL")
                //{
                //    files = di.GetFiles("*.csv");
                //}
                //else
                //{
                //    files = di.GetFiles("*.xlsx");
                //}
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
                if (Logfile_mode)
                {
                    if(proc_data.Logfile_ItemCode_LotNo_Matching(lstLogFile.SelectedItem.ToString(),txtItemCode.Text,txtLotNo.Text))
                    {
                        string log_locate = txtLocation.Text.Replace(Environment.NewLine, "");
                        string f_name = Path.Combine(log_locate, lstLogFile.SelectedItem.ToString());
                        DataTable netSpec_tbl = new DataTable();
                        DataTable result_data = new DataTable();
                        if (myVar_ECheck.CSV_en)
                        {
                            /*if(cbProcess.Text.ToUpper()!="ELECTRICAL")
                            {
                                btnSaveAll.Enabled = true;
                            }*/

                            int PCS_num = (int)numPCS.Value;
                            string sel_dev = cbMachine.Text;
                            try
                            {
                                switch (sel_dev)
                                {
                                    case "YAMAHA":
                                        result_data = proc_data.Yamaha_3GSPD_Process(f_name, ref netSpec_tbl);
                                        break;
                                    case "TAIYO":
                                        result_data = proc_data.Tayo_3GMRD_Process(f_name, PCS_num, ref netSpec_tbl);
                                        break;
                                }
                            }
                            catch
                            {
                                MessageBox.Show("Wrong device type!");
                            }
                            DGV_NET_Spec.DataSource = proc_data.load_spec_from_Logfile(netSpec_tbl, txtItemCode.Text, cbProcess.Text, sqlcon_OK2SHIP, true);
                            btnSaveAll.Enabled = true;
                        }
                        else
                        {
                            switch (cbProcess.Text)
                            {
                                case "CQRA_HOT_OIL":
                                    result_data = proc_data.HotOil_process(f_name);
                                    break;
                                case "CQRA_BENDING":
                                    result_data = proc_data.Bending_process(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, f_name);
                                    break;
                                case "CQRA_THERMAL_CYCLING_AND_BEND":
                                    result_data = proc_data.Bending_process(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, f_name);
                                    break;
                                case "CQRA_HEAT_SOAK_AND_BEND":
                                    result_data = proc_data.Bending_process(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, f_name);
                                    break;
                                case "CQRA_SURVIVAL_HOT_OIL":
                                    result_data = proc_data.Survival_HotOil_process(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, f_name);
                                    break;
                                case "CQRA_THERMAL_CYCLING_ON_COUPON":
                                    result_data = proc_data.ThermalCycling_On_Coupon(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, f_name);
                                    break;
                            }

                        }
                        DGV_LogFile.DataSource = result_data;
                    }
                    else
                    {
                        MessageBox.Show("Logfile is not matching with ItemCode and LotNo", "Warning");
                        return;
                    }
                }
                else
                {
                    if (myVar_ECheck.CSV_en)
                    {
                        DGV_LogFile.DataSource = proc_data.Load_Log_Data(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, cbCycles.Text, lstLogFile.SelectedItem.ToString());
                        DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { txtItemCode.Text, cbProcess.Text }));
                        DGV_NET_Spec.DataSource = spec_dt.AsDataView().ToTable(false, new string[] { "Point+V", "Point-V", "LSL", "USL" });
                    }
                    else
                    {
                        switch (cbProcess.Text)
                        {
                            case "CQRA_HOT_OIL":
                                string Logfile_name = lstLogFile.SelectedItem.ToString();
                                string region = "";
                                string f_name = Path.GetFileNameWithoutExtension(Logfile_name);
                                string[] temp = f_name.Split('-');
                                region = temp[temp.Length - 1];
                                DGV_LogFile.DataSource = proc_data.Load_Log_Data_HotOil(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, region, Logfile_name);
                                break;
                            case "CQRA_BENDING":
                                break;
                            case "CQRA_THERMAL_CYCLING_AND_BEND":
                                break;
                            case "CQRA_HEAT_SOAK_AND_BEND":
                                break;
                            case "CQRA_THERMAL_CYCLING_ON_COUPON":
                                //string Logfile_name = lstLogFile.SelectedItem.ToString();
                                DGV_LogFile.DataSource = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess.Text, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Logfile" }, new string[] { txtItemCode.Text, txtLotNo.Text, lstLogFile.SelectedItem.ToString() }));
                                btnSave.Enabled = false;
                                btnSaveSubmit.Enabled = false;
                                break;
                        }
                    }
                }
                myCode.DGV_Auto_Resize(DGV_LogFile);
                myCode.DGV_Auto_Resize(DGV_NET_Spec);
                if (myVar_ECheck.CSV_en)
                {
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
                    switch (cbProcess.Text)
                    {
                        case "CQRA_HOT_OIL":
                            if (proc_data.Check_HotOil_data(DGV_LogFile))
                            {
                                MessageBox.Show("OK");
                            }
                            else
                            {
                                MessageBox.Show("NG");
                            }
                            break;
                        case "CQRA_BENDING":
                            break;
                        case "CQRA_THERMAL_CYCLING_AND_BEND":
                            break;
                        case "CQRA_HEAT_SOAK_AND_BEND":
                            break;
                    }
                }
            }
            else
            {
                DGV_LogFile.DataSource = null;
                lblNG_Detail.Text = lblNGDetail.Replace(": ", "");
            }
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            /********************************************************* Main Program **************************************************************/
            if (txtItemCode.Text != "")
            {
                if (DGV_LogFile.DataSource != null)
                {
                    if (myVar_ECheck.CSV_en)
                    {
                        List<string> db_block_lst = proc_data.Get_block_from_Report_table(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.SelectedItem.ToString());
                        List<string> block_from_logfile = proc_data.Get_block_from_logfile_table(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.SelectedItem.ToString());
                        string cur_block = proc_data.Get_blockname_from_Logfile(lstLogFile.SelectedItem.ToString()).Split('-').LastOrDefault();
                        bool enable_save_log = false;
                        if(cbProcess.Text.ToUpper()=="ELECTRICAL")
                        {
                            enable_save_log = true;
                            goto next_act;
                        }
                        if(cbCycles.SelectedItem.ToString()=="Before")
                        {
                            if(db_block_lst.Count==0)
                            {
                                enable_save_log = true;
                            }
                            else
                            {
                                if(db_block_lst.IndexOf(cur_block)!=-1)
                                {
                                    enable_save_log = true;
                                }
                            }
                        }
                        else
                        {
                            if (db_block_lst.Count == 0)
                            {
                                if(block_from_logfile.IndexOf(cur_block)!=-1)
                                {
                                    enable_save_log = true;
                                }
                            }
                            else
                            {
                                if (db_block_lst.IndexOf(cur_block) != -1)
                                {
                                    enable_save_log = true;
                                }
                            }
                        }
              next_act: if(enable_save_log)
                        {
                            DGV_Data.DataSource = proc_data.Save_LogFile_detail(sqlcon_OK2SHIP, (DataTable)DGV_LogFile.DataSource, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, cbCycles.Text, lstLogFile.SelectedItem.ToString());
                        }
                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Block file không phù hợp với dữ liệu đã lưu", "Cảnh báo");
                            return;
                        }
                    }
                    else
                    {
                        DataTable src_tbl = (DataTable)DGV_LogFile.DataSource;
                        DGV_Data.DataSource = proc_data.Save_LogFile_HotOil(sqlcon_OK2SHIP, src_tbl, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, lstLogFile.SelectedItem.ToString());
                    }
                    MessageBox.Show("Lưu dữ liệu logfile thành công!", "Thông báo");
                }
                else
                {
                    MessageBox.Show("Không có dữ liệu", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show("Nhập thông tin ItemCode!", "Thông báo");
            }
            /********************************************************* End Main Program **************************************************************/
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
            /*if (txtItemCode.Text != "")
            {
                DataTable dt = new DataTable();
                DataTable spec_dt = new DataTable();
                if (!Logfile_mode)
                {
                    List<string> logfile_lst = new List<string>();
                    if (cbCycles.Text!="")
                    {
                        dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess.Text + "_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Cycles_name" }, new string[] { txtItemCode.Text, txtLotNo.Text, cbCycles.Text }));
                        proc_data.Get_List_data(-1, dt, new string[] { "ItemCode", "LotNo", "Cycles_name" }, ref logfile_lst, "Logfile", true);
                    }
                    else
                    {
                        dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess.Text + "_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo"}, new string[] { txtItemCode.Text, txtLotNo.Text }));
                        proc_data.Get_List_data(-1, dt, new string[] { "ItemCode", "LotNo"}, ref logfile_lst, "Logfile", true);
                    }                   
                    lstLogFile.DataSource = logfile_lst;
                    if(logfile_lst.Count==0)
                    {
                        MessageBox.Show("No data", "Warning");
                    }  
                }
                //else
                //{
                //    dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess.Text, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                //}
                
                //if (dt.Rows.Count == 0)
                //{
                //    MessageBox.Show("No data");
                //}
                //else
                //{
                //    DGV_Data.DataSource = dt;
                //}
                //myCode.DGV_Auto_Resize(DGV_NET_Spec);
            }
            else
            {
                MessageBox.Show("Please, enter ItemCode!", "Warning");
            }*/

            if (txtItemCode.Text != "")
            {
                DataTable dt = new DataTable();
                DataTable spec_dt = new DataTable();
                if (!Logfile_mode)
                {
                    List<string> logfile_lst = new List<string>();
                    if (cbCycles.Text != "")
                    {
                        if (myVar_ECheck.CSV_en)
                        {
                            dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess.Text + "_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Cycles_name" }, new string[] { txtItemCode.Text, txtLotNo.Text, cbCycles.Text }));
                            proc_data.Get_List_data(-1, dt, new string[] { "ItemCode", "LotNo", "Cycles_name" }, ref logfile_lst, "Logfile", true);
                        }
                        else
                        {

                            dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess.Text + "_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                            proc_data.Get_List_data(-1, dt, new string[] { "ItemCode", "LotNo" }, ref logfile_lst, "Logfile", true);

                        }

                    }
                    else
                    {
                        if (cbProcess.Text != "CQRA_THERMAL_CYCLING_ON_COUPON")
                        {
                            dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess.Text + "_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                        }
                        else
                        {
                            dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess.Text, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                        }
                        proc_data.Get_List_data(-1, dt, new string[] { "ItemCode", "LotNo" }, ref logfile_lst, "Logfile", true);

                    }
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
            string app_path =  Application.StartupPath;
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

        private void btnSummary_Click(object sender, EventArgs e)
        {
            /***************************** Main program ******************************************/
            Summary_data_Block();
            /************************************************************************************/
            /***************************** Test program ******************************************/
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
                Dictionary<string, List<string>> block_logfile = proc_data.Get_BlockofLogFile_Taiyo(lst_log);
                foreach (var block in block_logfile)
                {
                    foreach (var logfile in block.Value)
                    {
                        DataTable dt = new DataTable();
                        if (Logfile_mode)
                        {
                            string f_name = Path.Combine(txtLocation.Text, logfile);
                            DataTable netSpec_tbl = new DataTable();
                            DataTable result_data = new DataTable();
                            int Pcs_num = (int)numPCS.Value;
                            string sel_dev = cbMachine.Text;
                            try
                            {
                                switch (sel_dev)
                                {
                                    case "YAMAHA":
                                        dt = proc_data.Yamaha_3GSPD_Process(f_name, ref netSpec_tbl);
                                        break;
                                    case "TAIYO":
                                        dt = proc_data.Tayo_3GMRD_Process(f_name, Pcs_num, ref netSpec_tbl);
                                        break;
                                }

                            }
                            catch
                            {
                                MessageBox.Show("Wrong device type!");
                            }
                        }
                        else
                        {
                            dt = proc_data.Load_Log_Data(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, cbCycles.Text, logfile);
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
                    DataTable block_tbl = proc_data.Summary_Data_Table_from_list(tbl.Value, (DataTable)DGV_NET_Spec.DataSource);
                    result_table_lst.Add(tbl.Key, block_tbl);
                }
                btnDataSelect.Enabled = true;
                MessageBox.Show("Summary completed!\r\nPlease, choose data Qty then press Data Selected button","Warning");
                if(cbProcess.Text=="ELECTRICAL")
                {
                    DataTable electric_data = proc_data.Electric_Data_Summary(result_table_lst);
                    Electrical_Details frmElectric = new Electrical_Details(electric_data, txtItemCode.Text, cbProcess.Text,sqlcon_OK2SHIP);
                    frmElectric.Show();
                }
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
            DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess.Text, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
            proc_data.Get_List_data(-1, src_tbl, new string[] { "ItemCode", "LotNo" }, ref pcs_lst, "Pcs_No", true);
start_lbl:  if(pcs_lst.Count > 0)
            {
                if( MessageBox.Show("Dữ liệu các PCS đã có!\r\nBạn có lấy dữ liệu theo các pcs này không?","Thông báo",MessageBoxButtons.YesNo)==DialogResult.Yes)
                {
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
                                string col_name = tbl.Value.Columns[c_inx].ColumnName + "_" + "BL" + bl_name;
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
                    pcs_lst.Clear();
                    goto start_lbl;
                }
            }
            else
            {
                if (cbProcess.Text!="ELECTRICAL")
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
                    DGV_Data.DataSource = disp_result;
                }
                else
                {
                    int sel_qty = (int)numQty.Value;
                    List<DataTable> summ_tbl_lst = new List<DataTable>();
                    foreach (var tbl in result_table_lst)
                    {
                        List<ECheck_Process.NG_list2> cur_NGList = proc_data.Get_NG_point_tbl((DataTable)DGV_NET_Spec.DataSource, tbl.Value);
                        DataTable sel_dt = proc_data.Summary_Selected_FromExisted(tbl.Value, cur_NGList, tbl.Value.Columns.Count);
                        summ_tbl_lst.Add(sel_dt);
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
                    List<string> sel_Col = proc_data.Sum_OK_Item(sqlcon_OK2SHIP, disp_result, txtItemCode.Text, cbProcess.Text, 1.67);
                    List<string> col_name_sel = new List<string>();
                    int qty = Math.Min(sel_Col.Count, sel_qty);
                    if(sel_qty==0)
                    {
                        qty = sel_Col.Count;
                    }
                    for (int i = 0; i < qty; i++)
                    {
                        string t = sel_Col[i];
                        col_name_sel.Add(disp_result.Columns[int.Parse(t)].ColumnName);
                    }
                    DGV_Data.DataSource =disp_result.AsDataView().ToTable(false, col_name_sel.ToArray());
                    //DGV_Data.DataSource = proc_data.Electric_Summary_Block(sqlcon_OK2SHIP, result_table_lst, txtItemCode.Text, cbProcess.Text, (int)numQty.Value);
                    try
                    {
                        DataTable sel_data_tbl = (DataTable)DGV_Data.DataSource;
                        Electrical_Details frmElectric = new Electrical_Details(sel_data_tbl, txtItemCode.Text, cbProcess.Text, sqlcon_OK2SHIP);
                        frmElectric.Show();
                    }
                    catch
                    {

                    }
                }
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
            if(cbCycles.SelectedIndex >=0)
            {
                lblLogfile.Text = "Log file data: " + cbCycles.Text;
                reset_all();
            }
            else
            {
                lblLogfile.Text = "Log file data" ;
            }
            
        }

        private void lstNG_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstNG.SelectedIndex>-1)
            {
                int inx = lstNG.SelectedIndex;
                lblItemNG.Text = lblItem_title + lstNG.Items[inx];
                DGV_NG_detail.DataSource = proc_data.NG_NET_detail((DataTable) DGV_NET_Spec.DataSource, myNGlst2[inx]);
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
            if(rbDatabase.Checked)
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
                //txtLocation.Enabled = true;
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
            cbCycles.Enabled = true;
            if (Logfile_mode)
            {
                btnLoadData.Enabled = false;
                txtLocation.Enabled = true;
                btnBrowse.Enabled = true;
            }
            else
            {
                btnLoadData.Enabled = true;
                txtLocation.Enabled = false;
                btnBrowse.Enabled = false;
            }
            txtItemCode.Enabled = true;
            txtLotNo.Enabled = true;
        }
        public void reset_all_sel()
        {
            blNG_Details.reset_all();
            DGV_Cycles_Data_Sel.DataSource = null;
            DGV_Data_Sel.DataSource = null;
        }
        private void btnSaveSubmit_Click(object sender, EventArgs e)
        {
            if(myVar_ECheck.CSV_en)
            {
                DGV_Data.DataSource = proc_data.Save_Submit_data(sqlcon_OK2SHIP, (DataTable)DGV_Data.DataSource, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, cbCycles.Text);
                btnSaveSubmit.Enabled = false;
            }
            else
            {
                proc_data.Save_HotOil(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.Text);
            }
        }

        private void DGV_Data_DataSourceChanged(object sender, EventArgs e)
        {
            if(DGV_Data.DataSource==null)
            {
                btnSaveSubmit.Enabled = false;
            }
            else
            {
                if((btnDataSelect.Enabled)&&(!Logfile_mode))
                {
                    btnSaveSubmit.Enabled = true;
                }
            }
        }

        private void DGV_LogFile_DataSourceChanged(object sender, EventArgs e)
        {
            if(DGV_LogFile.DataSource ==null)
            {
                btnSave.Enabled = false;
                GB_Submit.Enabled = false;
                btnDataSelect.Enabled = false;
            }
            else
            {
                if (myVar_ECheck.CSV_en)
                {
                    GB_Submit.Enabled = true;
                    if (Logfile_mode)
                    {
                        btnSave.Enabled = true;
                    }
                }
                else
                {
                    GB_Submit.Enabled = false;
                    if(cbProcess.SelectedItem.ToString()=="CQRA_HOT_OIL")
                    {
                        if (Logfile_mode)
                        {
                            btnSave.Enabled = true;
                        }
                        else
                        {
                            btnSaveSubmit.Enabled = true;
                        }
                    }
                    //if (!Logfile_mode)
                    //{
                    //    btnSaveSubmit.Enabled = true;
                    //}
                }

            }
        }

        private void numQty_ValueChanged(object sender, EventArgs e)
        {
            //btnDataSelect.Enabled = false;
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            int ng_item_inx = 0;
            foreach(ECheck_Process.NG_list2 ng in blNG_Details.myNGlst2)
            {
                string cur_pcs = ng.col_val.Split('_')[1];
                foreach(int r_inx in ng.lst_row_inx)
                {
                    int tar_row = ng_item_inx * DGV_Cycles_Data_Sel.Rows.Count + r_inx;
                    string val_before = DGV_Data_Sel.Rows[tar_row].Cells[sel_item.cur_Cycles].Value.ToString();
                    string val_after = DGV_Cycles_Data_Sel.Rows[r_inx].Cells[ng_item_inx].Value.ToString();
                    if (val_before != val_after)
                    {
                        DGV_Data_Sel.Rows[tar_row].Cells[sel_item.cur_Cycles].Value = val_after ;
                        string id = (TDMK_Code.SQL_MAX("EDIT_HISTORY", "ID", sqlcon_OK2SHIP) + 1).ToString();
                        string[] item = new string[] { "ID", "ItemCode", "LotNo", "Net_No", "Pcs_No", "Data_Before", "Data_After", "UserID", "Date_Modify", "Process", "Remark" };
                        string[] item_val = new string[] { id, sel_item.cur_ItemCode, sel_item.cur_LotNo, (r_inx+1).ToString(), cur_pcs, val_before, val_after, myVar_ECheck.cur_User, DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), sel_item.cur_Process, sel_item.cur_Cycles };
                        TDMK_Code.insert_val_arr2("EDIT_HISTORY", sqlcon_OK2SHIP, item, item_val);
                    }
                }
                ng_item_inx++;
            }
        }

        private void cbProcess_Sel_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> cycles_lst = new List<string>();
            string app_path = Application.StartupPath;
            string tar_file = Path.Combine(app_path, "Config", cbProcess_Sel.Text + ".txt");
            cycles_lst = myCode.read_config_arr(tar_file).ToList();
            cbCycles_Sel.Text = "";
            cbCycles_Sel.DataSource = cycles_lst;
            reset_all_sel();
            switch(cbProcess_Sel.Text)
            {
                case "CQRA_HOT_OIL":
                    myVar_ECheck.CSV_en = false;
                    break;
                case "CQRA_BENDING":
                    myVar_ECheck.CSV_en = false;
                    break;
                case "CQRA_THERMAL_CYCLING_AND_BEND":
                    myVar_ECheck.CSV_en = false;
                    break;
                case "CQRA_HEAT_SOAK_AND_BEND":
                    myVar_ECheck.CSV_en = false;
                    break;
                case "CQRA_SURVIVAL_HOT_OIL":
                    myVar_ECheck.CSV_en = false;
                    break;
                case "CQRA_THERMAL_CYCLING_ON_COUPON":
                    myVar_ECheck.CSV_en = false;
                    break;
                default:
                    myVar_ECheck.CSV_en = true;
                    break;
            }
        }
        private void btnLoad_Data_Sel_Click(object sender, EventArgs e)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode_Sel.Text, txtLotNo_Sel.Text });
            DGV_Data_Sel.DataSource= TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess_Sel.Text, filter_str);
            myCode.Disable_Sort_DGV(DGV_Data_Sel);
        }

        private void DGV_Data_Sel_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DGV_CellColumnDoubleClick((DataTable)DGV_Data_Sel.DataSource, DGV_Cycles_Data_Sel, e.ColumnIndex);
            /*try
            {
                List<string> col_lst = new List<string>() { "ItemCode", "LotNo", "Net_No", "Pcs_No" };
                List<string> mycol_lst = new List<string>();
                DataTable cur_tbl = (DataTable)DGV_Data_Sel.DataSource;
                foreach (DataColumn dc in cur_tbl.Columns)
                {
                    if (dc.ColumnName.Contains("Before") || dc.ColumnName.Contains("After") || dc.ColumnName.Contains("Data"))
                    {
                        mycol_lst.Add(dc.ColumnName);
                    }
                }
                string sel_col_val = cur_tbl.Columns[e.ColumnIndex].ColumnName;
                lblCycleDetail.Text = "Details Cycles Data: " + sel_col_val;
                blNG_Details.reset_all();
                if (TDMK_Code.check_exist_list_index(sel_col_val, mycol_lst) != -1)
                {
                    col_lst.Add(sel_col_val);
                    DataTable filter_tbl = cur_tbl.AsDataView().ToTable(false, col_lst.ToArray());
                    DataTable tbl_detail = proc_data.Detail_NET_PCS(filter_tbl);
                    DGV_Cycles_Data_Sel.DataSource = tbl_detail;
                    DataTable temp_spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { txtItemCode_Sel.Text, cbProcess_Sel.Text }));
                    DataTable spec_dt = temp_spec_dt.AsDataView().ToTable(false, new string[] { "Point+V", "Point-V", "LSL", "USL" });
                    blNG_Details.fill_data(tbl_detail, spec_dt);
                    proc_data.check_in_spec_tbl(spec_dt, DGV_Cycles_Data_Sel);
                    myCode.DGV_Auto_Resize(DGV_Cycles_Data_Sel);
                    sel_item = new myVar_ECheck.item_info(txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, sel_col_val);
                    if (myCode.check_columns_existed(cur_tbl, "Before"))
                    {
                        List<string> bef_col_lst = new List<string>() { "ItemCode", "LotNo", "Net_No", "Pcs_No", "Before" };
                        DataTable before_tbl = cur_tbl.AsDataView().ToTable(false, bef_col_lst.ToArray());
                        DataTable bef_tbl_detail = proc_data.Detail_NET_PCS(before_tbl);
                        Delta_R_NGlist = new Dictionary<string, List<int>>();
                        int c_inx = 0;
                        foreach (DataColumn dc in bef_tbl_detail.Columns)
                        {
                            int r_inx = 0;
                            foreach (DataRow dr in bef_tbl_detail.Rows)
                            {
                                DataGridViewCell cur_cell = DGV_Cycles_Data_Sel.Rows[r_inx].Cells[c_inx];
                                bool add_en = false;
                                if (myCode.IsNumeric(cur_cell.Value.ToString()))
                                {
                                    double cur_val = Convert.ToDouble(dr[dc]);
                                    double cyc_val = Convert.ToDouble(cur_cell.Value);
                                    double rate = Math.Abs(cur_val - cyc_val) / cur_val;
                                    if (rate > 0.1)
                                    {
                                        cur_cell.Style.BackColor = Color.Red;
                                        add_en = true;
                                    }
                                }
                                else
                                {
                                    cur_cell.Style.BackColor = Color.LightPink;
                                    add_en = true;
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
                            c_inx++;
                        }
                    }
                }
            }
            catch
            {

            }*/

        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            if(myVar_ECheck.confirm_mode)
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
            if(e.Button == MouseButtons.Right)
            {
                if (myVar_ECheck.confirm_mode)
                {
                    DGV_Cycles_Data_Sel.ContextMenuStrip = cmsCopyPaste;
                    /*DGV_Cycles_Data_Sel.ReadOnly = false;
                    DataGridViewCell cur_cell = DGV_Cycles_Data_Sel.CurrentCell;
                    string cur_col_1 = DGV_Cycles_Data_Sel.Columns[cur_cell.ColumnIndex].Name;
                    string cur_col = cur_col_1.Split('_')[1];
                    sel_cell = new myVar_ECheck.edit_cell(cur_col, (cur_cell.RowIndex+1).ToString(), cur_cell.Value.ToString());
                    int item_temp = blNG_Details.item_NG_lst.FindIndex(x => x == cur_col_1);
                    if (item_temp!=-1)
                    {
                        ECheck_Process.NG_list2 cur_NG = blNG_Details.myNGlst2[item_temp];
                        int r_temp = cur_NG.lst_row_inx.FindIndex(x => x == cur_cell.RowIndex);
                        if(r_temp!=-1)
                        {
                            DGV_Cycles_Data_Sel.BeginEdit(true);
                        }
                        else
                        {
                            MessageBox.Show("Cannot Edit");
                            DGV_Cycles_Data_Sel.ReadOnly = true;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Cannot Edit");
                        DGV_Cycles_Data_Sel.ReadOnly = true;
                    }*/
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
            if (sel_cell.after_val!=sel_cell.before_val)
            {
                if (!myCode.IsNumeric(sel_cell.after_val))
                {
                    cur_cell.Value = sel_cell.before_val;
                }
                Color cell_color = myCode.check_in_limit2(USL, LSL, cur_cell.Value.ToString());
                if(cell_color==Color.White)
                {
                    if(cbProcess_Sel.Text == "CQRA_SURVIVAL_HOT_OIL")
                    {
                        ref_rate = 0.05;
                    }
                    string col_name = DGV_Cycles_Data_Sel.Columns[cur_cell.ColumnIndex].Name;
                    //cur_cell.Style.BackColor = proc_data.CheckCell_Variability_Resistance(ref_rate, (DataTable)DGV_Data_Sel.DataSource, cur_cell, col_name);
                    cur_cell.Style.BackColor = proc_data.CheckCell_Variability_Resistance_bending(ref_rate, (DataTable)DGV_Data_Sel.DataSource, cur_cell, col_name, sel_item.cur_Cycles,dic_index_lst);

                }
                else
                {
                    cur_cell.Style.BackColor = cell_color;
                }
            }
            if(!manual_en)
            {
                DGV_Cycles_Data_Sel.ReadOnly = true;
            }
        }
        private void DGV_Cycles_Data_Sel_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if(myVar_ECheck.confirm_mode)
            {
                if (MessageBox.Show("Do you want to update?","Warning",MessageBoxButtons.YesNo)==DialogResult.Yes)
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
                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { sel_item.cur_ItemCode, sel_item.cur_LotNo });
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
                        int PCS_Num = (int)numPCS.Value;
                        string sel_dev = cbMachine.Text;
                        try
                        {
                            switch (sel_dev)
                            {
                                case "YAMAHA":
                                    result_data = proc_data.Yamaha_3GSPD_Process(f_name, ref netSpec_tbl);
                                    break;
                                case "TAIYO":
                                    result_data = proc_data.Tayo_3GMRD_Process(f_name, PCS_Num, ref netSpec_tbl);
                                    break;
                            }
                            tbl_data_lst.Add(result_data);
                        }
                        catch
                        {
                            MessageBox.Show("Wrong device type!");
                        }
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
                btnDataSelect.Enabled = true;
            }                                                                                                                                                                                                                                                                                                                                                                  
        }

        private void btnTest_Click_1(object sender, EventArgs e)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
            DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess_Sel.Text + "_LOGFILE", filter_str);
            List<DataTable> tbl_lst = new List<DataTable>();
            List<string> logfile_lst = new List<string>();
            proc_data.Get_List_data2(-1, src_tbl, new string[] { "ItemCode", "LotNo" }, ref logfile_lst, "Logfile", true);
            foreach(var log_f in logfile_lst)
            {
                DataTable temp = proc_data.Load_Log_Data(sqlcon_OK2SHIP, txtItemCode.Text, txtLotNo.Text, cbProcess.Text, cbCycles.Text, log_f);
                tbl_lst.Add(temp);
            }
            DataTable dest_tbl = new DataTable();
            List<List<string>> src_NET_lst = new List<List<string>>();
            foreach(var tbl in tbl_lst)
            {
                int dr_inx = 0;
                foreach(DataRow dr in tbl.Rows)
                {
                    if(src_NET_lst.Count < dr_inx+1)
                    {
                        src_NET_lst.Add(new List<string>());
                    }
                    foreach(DataColumn dc in tbl.Columns)
                    {
                        src_NET_lst[dr_inx].Add(dr[dc].ToString());
                    }
                    dr_inx++;
                }
            }
            for(int i=0;i< src_NET_lst[0].Count;i++)
            {
                dest_tbl.Columns.Add("ItemNo_" + (i + 1).ToString());
            }
            foreach(var n_lst in src_NET_lst)
            {
                DataRow dr = dest_tbl.NewRow();
                int c_inx = 0;
                foreach(string t in n_lst)
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
            DataTable src_tbl = (DataTable)DGV_LogFile.DataSource;
            Electrical_Details frmElectric = new Electrical_Details(src_tbl, txtItemCode.Text, cbProcess.Text, sqlcon_OK2SHIP);
            frmElectric.Show();
        }

        private void DGV_LogFile_ColumnAdded(object sender, DataGridViewColumnEventArgs e)
        {
            
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode_Sel.Text);
            if (format_file != "")
            {
                string export_path = Path.Combine(Application.StartupPath, "Export", cbProcess_Sel.Text);
                if(!Directory.Exists(export_path))
                {
                    Directory.CreateDirectory(export_path);
                }
                ExcelPackage format_pack = Excel_lib.open_excel(format_file);
                ExcelWorkbook format_wrkbk = format_pack.Workbook;
                string export_success = "";
                switch (cbProcess_Sel.Text)
                {
                    case "CQRA_HOT_OIL":
                        //proc_data.Export_HotOil(sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Hot oil");
                        export_success = Echeck_export_EPPlus.Export_HotOil(format_wrkbk, sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Hot oil");
                        break;
                    case "CQRA_BENDING":
                        //proc_data.Export_Bending(sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Bending");
                        export_success = Echeck_export_EPPlus.Export_Bending(format_wrkbk, sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Bending");
                        break;
                    case "CQRA_THERMAL_CYCLING_AND_BEND":
                        //proc_data.Export_Thermal_HeatSoak_Bend(sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Thermal cycling and bend", 20);
                        export_success = Echeck_export_EPPlus.Export_Bending(format_wrkbk, sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Thermal cycling and bend");
                        break;
                    case "CQRA_HEAT_SOAK_AND_BEND":
                        //proc_data.Export_Thermal_HeatSoak_Bend(sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Heat soak and bend", 20);
                        export_success = Echeck_export_EPPlus.Export_Bending(format_wrkbk, sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Heat soak and bend");
                        break;
                    case "CQRA_REFLOW":
                        //proc_data.Export_Reflow_new(sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Reflow");
                        export_success = Echeck_export_EPPlus.Export_Reflow_new(format_wrkbk, sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Reflow");
                        break;
                    case "CQRA_THERMAL_CYCLING":
                        //proc_data.Export_ThermalCycling(sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Thermal cycling");
                        export_success = Echeck_export_EPPlus.Export_ThermalCycling(format_wrkbk, sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Thermal cycling");
                        break;
                    case "CQRA_HEAT_SOAK":
                        //proc_data.Export_ThermalCycling(sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Heat Soak");
                        export_success = Echeck_export_EPPlus.Export_ThermalCycling(format_wrkbk, sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Heat Soak");
                        break;
                    case "CQRA_THERMAL_SHOCK":
                        //proc_data.Export_ThermalCycling(sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Thermal shock");
                        export_success = Echeck_export_EPPlus.Export_ThermalCycling(format_wrkbk, sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Thermal shock");
                        break;
                    case "CQRA_SURVIVAL_REFLOW":
                        //proc_data.Export_Survival(sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Survival reflow");
                        export_success = Echeck_export_EPPlus.Export_Survival(format_wrkbk, sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Survival reflow");
                        break;
                    case "CQRA_SURVIVAL_HOT_OIL":
                        //proc_data.Export_Surv_HotOil_Multi(sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Survival hot oil");
                        export_success = Echeck_export_EPPlus.Export_Surv_HotOil_Multi(format_wrkbk, sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Survival hot oil");
                        break;
                    case "ELECTRICAL":
                        //proc_data.Export_Electrical(sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "Electrical");
                        export_success = Echeck_export_EPPlus.Export_Electrical(format_wrkbk, sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "Electrical");
                        break;
                    case "CQRA_THERMAL_CYCLING_ON_COUPON":
                        //proc_data.Export_ThermalCycling_On_Coupon(sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Thermal cycling");
                        export_success = Echeck_export_EPPlus.Export_ThermalCycling_On_Coupon(format_wrkbk, sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbProcess_Sel.Text, "CQRA - Thermal cycling");
                        break;
                }
                if(export_success!="")
                {
                    string export_file = Path.Combine(export_path, cbProcess_Sel.Text + "_" + txtItemCode_Sel.Text + "-" + txtLotNo_Sel.Text + "_" + txtUser_Sel.Text+"_" + DateTime.Now.ToString("ddMMMyyyy HHmmss") + ".xlsx");
                    ExcelPackage report_pack = new ExcelPackage( export_file);
                    ExcelWorkbook report_wrkbk = report_pack.Workbook;
                    report_wrkbk.Worksheets.Add(export_success, format_wrkbk.Worksheets[export_success]);
                    report_pack.Save();
                    report_pack = null;
                    report_wrkbk = null;
                    try
                    {
                        ProcessStartInfo pi = new ProcessStartInfo(export_file);
                        Process.Start(pi);
                    }
                    catch { };

                }
                format_pack = null;
                format_wrkbk = null;
                MessageBox.Show(new Form { TopMost = true }, "Hoàn thành xuất dữ liệu", "Thông báo");
            }
            else
            {
                MessageBox.Show( new Form { TopMost = true}, "Không tìm thấy Format", "Thông báo");
            }
            
        }
        private void DGV_Data_Sel_DataSourceChanged(object sender, EventArgs e)
        {
            double ref_rate = 0.1;
            if(cbProcess_Sel.SelectedItem.ToString()== "CQRA_SURVIVAL_HOT_OIL")
            {
                ref_rate = 0.05;
                //proc_data.Check_Cycles_Data(DGV_Data_Sel, 0.05);
                
            }
            else
            {
                //proc_data.Check_Cycles_Data(DGV_Data_Sel, 0.1);
            }
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
            myVar._frmMain.Show();
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
                                    string cur_col = cur_col_1.Split('_')[1];
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
                                string cur_col = cur_col_1.Split('_')[1];
                                if (curr_cell.Selected)
                                {
                                    if(!manual_en)
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
            if((e.RowIndex>-1)&&(e.ColumnIndex >-1))
            {
                if(!manual_en)
                {
                    if (myVar_ECheck.confirm_mode)
                    {
                        DGV_Cycles_Data_Sel.ReadOnly = false;
                        DataGridViewCell cur_cell = DGV_Cycles_Data_Sel.CurrentCell;
                        string cur_col_1 = DGV_Cycles_Data_Sel.Columns[cur_cell.ColumnIndex].Name;
                        string cur_col = cur_col_1.Split('_')[1];
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
            if(cbProcess_Sel.Text.ToUpper()!="ELECTRICAL")
            {
                int r_inx = e.RowIndex;
                int c_inx = e.ColumnIndex;
                if ((r_inx > -1) && (c_inx > -1))
                {
                    e.ToolTipText = proc_data.Details_Variability_Resistance_bending(r_inx, c_inx, (DataTable)DGV_Data_Sel.DataSource, (DataTable)DGV_Cycles_Data_Sel.DataSource, sel_item.cur_Cycles, dic_index_lst);
                    //if (sel_item.cur_Cycles.Contains("Bending"))
                    //{
                    //    e.ToolTipText = proc_data.Details_Variability_Resistance_bending(r_inx, c_inx, (DataTable)DGV_Data_Sel.DataSource, (DataTable)DGV_Cycles_Data_Sel.DataSource, sel_item.cur_Cycles,dic_index_lst);
                    //}
                    //else
                    //{
                    //    e.ToolTipText = proc_data.Details_Variability_Resistance(r_inx, c_inx, (DataTable)DGV_Data_Sel.DataSource, (DataTable)DGV_Cycles_Data_Sel.DataSource);
                    //}

                }
            }
        }

        private void lstNG_DataSourceChanged(object sender, EventArgs e)
        {

        }
        public void DGV_CellColumnDoubleClick(DataTable cur_tbl, DataGridView tar_DGV, int ColumnIndex)
        {
            try
            {
                double ref_rate = 0.1;
                manual_en = false;
                if(cbProcess_Sel.Text=="CQRA_SURVIVAL_HOT_OIL")
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
                    List<string>sel_pcs_report= proc_data.Get_SelectedPcs_List(cur_tbl, "Sel_report", "Yes");
                    if(sel_pcs_report.Count>0)
                    {
                        tbl_detail = tbl_detail.AsDataView().ToTable(false, sel_pcs_report.ToArray());
                    }    
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
                                    }
                                }
                                else
                                {
                                    cur_cell.Style.BackColor = Color.LightPink;
                                    add_en = true;
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
                            c_inx++;
                        }
                    }
                }
            }
            catch (Exception E)
            {
                MessageBox.Show(E.Message.ToString());
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
            //if (txtLocation.Text != "")
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

        private void txtLocation_TextChanged(object sender, EventArgs e)
        {

        }

        private void _txtLocation_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode==Keys.Enter)
            {
                e.SuppressKeyPress = true;
                if (txtLocation.Text != "")
                {
                    string tar_loc = txtLocation.Text.Replace(Environment.NewLine, "");
                    DirectoryInfo di = new DirectoryInfo(tar_loc);
                    if (di.Exists)
                    {
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
                                files = di.GetFiles("*.xlsx");
                                break;
                            case "CQRA_HEAT_SOAK_AND_BEND":
                                files = di.GetFiles("*.xlsx");
                                break;
                            case "CQRA_SURVIVAL_HOT_OIL":
                                files = di.GetFiles("*.xlsx");
                                break;
                            case "CQRA_THERMAL_CYCLING_ON_COUPON":
                                files = di.GetFiles("*.xlsx");
                                break ;
                            default:
                                files = di.GetFiles("*.csv");
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
                    else
                    {
                        MessageBox.Show("Folder not existed!", "Warning");
                    }
                }
            }
            
        }

        private void ECheck_Main_Resize(object sender, EventArgs e)
        {
            rs.ResizeAllControls(this);
            blNG_Details.resize();
        }

        private void btnReCheck_Click(object sender, EventArgs e)
        {
            if (DGV_Cycles_Data_Sel.DataSource != null)
            {
                ReCheck_Data frmRecheck = new ReCheck_Data(sqlcon_OK2SHIP, txtItemCode_Sel.Text, txtLotNo_Sel.Text, lblCycleDetail.Text.Split(':').LastOrDefault().Trim(), cbProcess_Sel.SelectedItem.ToString(), (DataTable)DGV_Cycles_Data_Sel.DataSource);
                frmRecheck.Show();
            }
        }

        private void btnToExcel_Click(object sender, EventArgs e)
        {
            string tbl_name = cbProcess_Sel.SelectedItem.ToString() + "_" + txtItemCode_Sel.Text + "-" + txtLotNo_Sel.Text +"_"+ DateTime.Now.ToString("yyyyMMMdd_HHmmss");
            exp_proc.Table_To_CSV(tbl_name, (DataTable)(DGV_Data_Sel.DataSource));
            //TDMK_Code.Export_DGV_Excel3(DGV_Data_Sel, TDMK_Code.Create_workbook(), true);
        }

        private void btnExportDetail_Click(object sender, EventArgs e)
        {
            TDMK_Code.Export_DGV_Excel3(DGV_Cycles_Data_Sel, TDMK_Code.Create_workbook(), true);
        }

        private void btnSaveAll_Click(object sender, EventArgs e)
        {
            string log_locate = txtLocation.Text.Replace(Environment.NewLine, "");
            List<string> logfile_lst = new List<string>();
            logfile_lst = (List<string>)lstLogFile.DataSource;
            if (!save_data_worker.IsBusy)
            {
                object arg = new List<object>() { txtItemCode.Text, txtLotNo.Text, cbProcess.Text, cbCycles.Text, numPCS.Value, cbMachine.Text, logfile_lst, log_locate, numQty.Value };
                save_data_worker.RunWorkerAsync(arg);
                timer1.Enabled = true;
                GBInfo.Enabled = false;
                GB_Data.Enabled = false;
                GBAction.Enabled = false;
                GB_Submit.Enabled = false;
            }
            else
            {
                MessageBox.Show("Hệ thống đang xử lý", "Thông báo");
            }
            
        }

        private void lstLogFile_DataSourceChanged(object sender, EventArgs e)
        {

        }

        private void tapMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(tapMain.SelectedIndex==1)
            {
                txtItemCode_Sel.Text = txtItemCode.Text;
                txtLotNo_Sel.Text = txtLotNo.Text;
                cbProcess_Sel.SelectedIndex = cbProcess.SelectedIndex;
                cbCycles_Sel.SelectedIndex = cbCycles.SelectedIndex;
                txtUser_Sel.Text = txtOperator.Text;
            }    
        }

        private void btnLoadData_pcs_Click(object sender, EventArgs e)
        {

            /*************************************************************** Main Program *************************************************************************/
            reset_selected_pcs();
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode_pcs.Text, txtLotNo_pcs.Text });
            temp_pcs_sel_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess_pcs.Text, filter_str);
            export_pcs_lst = temp_pcs_sel_dt.AsEnumerable().Select(x => x.Field<string>("Pcs_No")).Distinct().ToList();
            List<string> block_lst = export_pcs_lst.Select(x => x.Split('_').LastOrDefault()).Distinct().ToList();
            List<string> pcs_selected_lst = proc_data.Get_SelectedPcs_List(temp_pcs_sel_dt, "Sel_report", "Yes");
            lstPcs_selected.Items.AddRange(pcs_selected_lst.ToArray());
            block_lst.Add("All");
            lstBlock.DataSource = block_lst;
            if (pcs_selected_lst.Count != 0)
            {
                lblSelectedPcs.Text = "Selected PCS = " + pcs_selected_lst.Count.ToString();
                DGV_Details_Pcs.DataSource = temp_pcs_sel_dt.AsEnumerable().Where(x => pcs_selected_lst.IndexOf(x.Field<string>("Pcs_No")) != -1).CopyToDataTable();
            }
            else
            {
                lblSelectedPcs.Text = "Selected PCS";
                DGV_Details_Pcs.DataSource = temp_pcs_sel_dt;
            }
            /*************************************************************** End Main Program *************************************************************************/
        }

        private void lstBlock_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(lstBlock.SelectedIndex!=-1)
            {
                List<string> pcs_lst = new List<string>();
                string lst_bl = lstBlock.SelectedItem.ToString();
                if (lst_bl == "All")
                {
                    pcs_lst= export_pcs_lst;
                }
                else
                {
                    pcs_lst = export_pcs_lst.Where(x => x.Contains(lst_bl)).Distinct().ToList();
                }
                chkPcs.Items.Clear();
                chkPcs.Items.AddRange(pcs_lst.ToArray());
                foreach(var chk in lstPcs_selected.Items)
                {
                    int inx = chkPcs.Items.IndexOf(chk);
                    if (inx != -1)
                    {
                        chkPcs.SetItemChecked(inx, true);
                    }
                }
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            foreach(var chk in chkPcs.CheckedItems)
            {
                if(lstPcs_selected.Items.IndexOf(chk)==-1)
                {
                    lstPcs_selected.Items.Add(chk);
                }
            }
            
        }

        private void lstPcs_selected_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            int inx = lstPcs_selected.SelectedIndex;
            if (inx != -1)
            {
                if(MessageBox.Show("Loại bỏ khỏi danh sách lựa chọn?","Thông báo",MessageBoxButtons.YesNo)==DialogResult.Yes)
                {
                    object del_item = lstPcs_selected.Items[inx];
                    lstPcs_selected.Items.RemoveAt(inx);
                    int chk_inx = chkPcs.Items.IndexOf(del_item);
                    if (chk_inx != -1)
                    {
                        chkPcs.SetItemChecked(chk_inx, false);
                    }    
                }    

            }
        }

        private void chkPcs_MouseDown(object sender, MouseEventArgs e)
        {
            if(e.Button==MouseButtons.Right)
            {
                if(chkPcs.CheckedItems.Count!=0)
                {
                    chkPcs.ContextMenuStrip = cmsPcsSelected;
                }
                else
                {
                    chkPcs.ContextMenuStrip = null;
                }
            }
        }

        private void tsmAdd_Click(object sender, EventArgs e)
        {
            foreach (var chk in chkPcs.CheckedItems)
            {
                if (lstPcs_selected.Items.IndexOf(chk) == -1)
                {
                    lstPcs_selected.Items.Add(chk);
                }
            }
            List<string> sel_lst = new List<string>();
            foreach(var pcs in lstPcs_selected.Items)
            {
                sel_lst.Add(pcs.ToString());
            }
            DGV_Details_Pcs.DataSource= temp_pcs_sel_dt.AsEnumerable().Where(x => sel_lst.IndexOf(x.Field<string>("Pcs_No")) != -1).CopyToDataTable();
            lblSelectedPcs.Text = "Selected PCS = " + sel_lst.Count.ToString();
        }

        private void cbProcess_pcs_SelectedIndexChanged(object sender, EventArgs e)
        {
            reset_selected_pcs();
        }

        private void btnSaveDB_Click(object sender, EventArgs e)
        {

            foreach (DataRow dr in temp_pcs_sel_dt.Rows)
            {
                string pcs_val = dr["Pcs_No"].ToString();
                if(lstPcs_selected.Items.IndexOf(pcs_val)!=-1)
                {
                    dr["Sel_report"] = "Yes";
                }
                else
                {
                    dr["Sel_report"] = "";
                }
            }
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode_pcs.Text, txtLotNo_pcs.Text });
            TDMK_Code.Delelte_FilteredItem_arr(cbProcess_pcs.SelectedItem.ToString(), sqlcon_OK2SHIP, filter_str);
            proc_data.BatchBulkCopy(sqlcon_OK2SHIP, temp_pcs_sel_dt, cbProcess_pcs.SelectedItem.ToString());
            MessageBox.Show("Lưu dữ liệu thành công", "Thông báo");

        }
        public void reset_selected_pcs()
        {
            temp_pcs_sel_dt = null;
            DGV_Details_Pcs.DataSource = null;
            lstBlock.DataSource = null;
            lstPcs_selected.Items.Clear();
            chkPcs.Items.Clear();
            lblSelectedPcs.Text = "Selected PCS";
        }

        private void save_data_worker_DoWork(object sender, DoWorkEventArgs e)
        {

            List<object> src_ojb_lst = (List<object>)e.Argument;
            string ItemCode = (string)src_ojb_lst[0];
            string LotNo = (string)src_ojb_lst[1];
            string process = (string)src_ojb_lst[2];
            string cycle = (string)src_ojb_lst[3];
            int numNET = Convert.ToInt32(src_ojb_lst[4]);
            string Machine = (string)src_ojb_lst[5];
            List<string> logfile_lst = (List<string>)src_ojb_lst[6];
            string log_locate = (string)src_ojb_lst[7];
            int numQty = Convert.ToInt32( src_ojb_lst[8]);
            save_data_worker.ReportProgress(0);
            if (proc_data.Save_logfile_all_en(sqlcon_OK2SHIP, ItemCode, LotNo, process, cycle, numNET, Machine, logfile_lst, log_locate))
            {
                if(process!="ELECTRICAL")
                {
                    save_data_worker.ReportProgress(50);
                    proc_data.Summary_Logfile_All(sqlcon_OK2SHIP, ItemCode, LotNo, process, cycle, numQty);
                }
                else
                {
                    MessageBox.Show("Lưu dữ liệu logfile hoàn thành", "Thông báo");
                }
            }            
        }

        private void save_data_worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int process_val = e.ProgressPercentage;
            switch(process_val)
            {
                case 0:
                    lblProcess_Data.Text = "Saving logfile data.....";
                    break;
                case 50:
                    lblProcess_Data.Text = "Saving summary data.....";
                    break;
            }

        }

        private void save_data_worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            
            if (e.Cancelled)
            {
                lblProcess_Data.Text = "Hủy tác vụ";
            }

            else if (e.Error != null)
            {
                lblProcess_Data.Text = "Xảy ra lỗi";
            }
            else
            {
                lblProcess_Data.Text = "Summary Process Data";
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, cbProcess.SelectedItem.ToString(), filter_str);
                DGV_Data.DataSource = src_dt;
            }
            timer1.Enabled = false;
            lblProcess_Data.BackColor = Color.Aqua;
            GBInfo.Enabled = true;
            GB_Data.Enabled = true;
            GBAction.Enabled = true;
            GB_Submit.Enabled = true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if(lblProcess_Data.BackColor == Color.Aqua)
            {
                lblProcess_Data.BackColor = Color.White;
            }
            else
            {
                lblProcess_Data.BackColor = Color.Aqua;
            }
        }

        private void btnLoad_Net_Click(object sender, EventArgs e)
        {
            if (txtItemCode_NetSel.Text != "")
            {
                //DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "NET_SPEC", "ItemCode = '" + txtItemCode.Text + "'");
                DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { txtItemCode_NetSel.Text, cbProcess_NetSel.Text }));
                DGV_NetSel.DataSource = spec_dt;
                if (spec_dt.Rows.Count == 0)
                {
                    if (MessageBox.Show("No data! \r\n Do you want to setup ?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        myExcel.Workbook spec_wrkbk = TDMK_Code.Create_workbook();
                        myExcel.Worksheet spec_wrksht = spec_wrkbk.Sheets[1];
                        int col_inx = 0;
                        foreach (DataColumn dc in spec_dt.Columns)
                        {
                            spec_wrksht.Range["A1"].Offset[0, col_inx].Value = dc.ColumnName;
                            col_inx++;
                        }
                        spec_wrkbk.Activate();
                    }
                }
            }
            else
            {
                MessageBox.Show("Please, enter ItemCode!", "Warning");
            }
        }

        private void btnExport_NetSel_Click(object sender, EventArgs e)
        {
            if (DGV_NetSel.Rows.Count > 0)
            {
                myExcel.Workbook exp_wrkbk = TDMK_Code.Create_workbook();
                TDMK_Code.Export_DGV_Excel3(DGV_NetSel, exp_wrkbk, true);
                myExcel.Worksheet exp_sht = exp_wrkbk.ActiveSheet;
                exp_sht.Range["A:A"].Delete();
            }
            else
            {
                MessageBox.Show("No data!", "Warning");
            }
        }

        private void btnImport_NETSel_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            TDMK_Code.fill_dataset(ds, "NET_SPEC", sqlcon_OK2SHIP);
            DataTable spec_tbl = ds.Tables[0];
            DataTable cur_tbl = spec_tbl.Copy();
            cur_tbl.Rows.Clear();
            myExcel.Application xlapp = TDMK_Code.StartExcel();
            myExcel.Workbook cur_wrkbk = xlapp.ActiveWorkbook;
            if (cur_wrkbk != null)
            {
                myExcel.Worksheet cur_wrksht = cur_wrkbk.Sheets[1];
                myExcel.Range cur_rgn = cur_wrksht.Range["E2"];
                int r_inx = 0;

                while (myCode.checkDBNull(cur_rgn.Offset[r_inx, 0].Value) != "")
                {
                    DataRow dr = cur_tbl.NewRow();
                    for (int i = 0; i < cur_tbl.Columns.Count; i++)
                    {
                        dr[i] = cur_rgn.Offset[r_inx, i - 4].Value;
                    }
                    cur_tbl.Rows.Add(dr);
                    r_inx++;
                }
                DGV_NetSel.DataSource = cur_tbl;
            }
            else
            {
                if (MessageBox.Show("Excel format file is closed!\r\nDo you want to re-open it?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    myExcel.Workbook spec_wrkbk = TDMK_Code.Create_workbook();
                    myExcel.Worksheet spec_wrksht = spec_wrkbk.Sheets[1];
                    int col_inx = 0;
                    foreach (DataColumn dc in cur_tbl.Columns)
                    {
                        spec_wrksht.Range["A1"].Offset[0, col_inx].Value = dc.ColumnName;
                        col_inx++;
                    }
                    spec_wrkbk.Activate();
                }
            }
        }

        private void btnSave_NETSel_Click(object sender, EventArgs e)
        {
            DataTable myDt = (DataTable)DGV_NetSel.DataSource;
            List<DataTable> mylst_dt = new List<DataTable>();
            AutoCompleteStringCollection lst_ItemCode = TDMK_Code.Load_Item_Names("NET_SPEC", "ItemCode", sqlcon_OK2SHIP);
            proc_data. Get_ListTable(-1, myDt, new string[] { "ItemCode" }, ref mylst_dt, "Point+V");
            foreach (DataTable dt in mylst_dt)
            {
                bool pro_en = false;
                string itemcode = dt.Rows[0][1].ToString();
                if (TDMK_Code.check_exist_list(itemcode, lst_ItemCode))
                {
                    if (MessageBox.Show("ItemCode " + itemcode + " is existed! \r\nDo you want to update?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        pro_en = true;
                    }
                    else
                    {
                        pro_en = false;
                    }
                }
                else
                {
                    pro_en = true;
                }
                if (pro_en)
                {
                    TDMK_Code.Delelte_FilteredItem_arr("NET_SPEC", sqlcon_OK2SHIP, TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { txtItemCode_NetSel.Text, cbProcess_NetSel.Text }));
                    int id = TDMK_Code.SQL_MAX("NET_SPEC", "ID", sqlcon_OK2SHIP);
                    int r_inx = 0;
                    foreach (DataRow dr in dt.Rows)
                    {
                        dr[0] = (id + r_inx + 1).ToString();
                        r_inx++;
                    }
                    proc_data. BatchBulkCopy(sqlcon_OK2SHIP, dt, "NET_SPEC");
                }

            }
        }

        private void DGV_NetSel_DataSourceChanged(object sender, EventArgs e)
        {
            if(DGV_NetSel.DataSource != null)
            {
                myCode.Disable_Sort_DGV(DGV_NetSel);
                foreach(DataGridViewColumn dgv_c in DGV_NetSel.Columns)
                {
                    if(dgv_c.Name=="Sel_Report")
                    {
                        dgv_c.ReadOnly = false;
                    }
                    else
                    {
                        dgv_c.ReadOnly = true;
                    }
                }
            }
        }

        private void DGV_NetSel_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            string cell_val = myCode.checkDBNull(DGV_NetSel.CurrentCell.Value);
            List<string> accept_lst = new List<string>() { "YES", "NO" };
            if(cell_val != "")
            {
                if (accept_lst.IndexOf(cell_val.ToUpper())==-1)
                {
                    DGV_NetSel.CurrentCell.Value = "";
                    MessageBox.Show(new Form { TopMost = true }, "Giá trị nhập không hợp lệ \r\nHãy nhập Yes hoặc No", "Thông báo");
                }
            }

        }

        private void txtLotNo_pcs_Validated(object sender, EventArgs e)
        {
            txtLotNo_pcs.Text = proc_data.Lotno_Formated(txtLotNo_pcs.Text);
        }
    }
}
