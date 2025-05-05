using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using TDMK_SEEV_DLL;
using TDMK_SQL;
using IniLibs;
using System.Windows.Forms;
using TDMK_EPPLUS_7;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace Bending_Export
{
    public class Bending_Export_EPPLUS_Lib
    {
        public TDMK_EPPLUS7_lib Excel_Lib = new TDMK_EPPLUS7_lib();
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SEI_Lib myCode = new SEI_Lib();

        IniFile TDMK_init;// = new IniFile("Config.ini");

        //IniFile TDMK_init;// = new IniFile();

        public string Data_Location { get; set; }
        public string format_folder { get; set; }
        public string log_folder { get; set; }
        public string Report_location { get; set; }
        public SqlConnection initial_data(string DB_name, bool sa_en)
        {
            SqlConnection _sqlcon_OK2SHIP;
            string program_loc = find_config_path(Application.StartupPath, "TDMK Program");
            string config_path = Path.Combine(program_loc, "Config.ini");
            //string config_path = Path.Combine(Application.StartupPath, "Config.ini");
            TDMK_init = new IniFile(config_path);
            if (!File.Exists(config_path))
            {
                TDMK_init.Write("Server", "10.212.1.243", "SMT_Config");
                TDMK_init.Write("Account", "sa", "SMT_Config");
                TDMK_init.Write("Password", "seev@123;", "SMT_Config");
                TDMK_init.Write("Report_Location", @"\\10.212.1.200\rgo\OK2SHIP automation system", "SMT_Config");
                TDMK_init.Write("Data_Location", @"\\10.212.1.200\rgo\OK2SHIP automation system\TDMK_Data\TestAreas\OK2SHIP_Measurements\OMM", "SMT_Config");
                TDMK_init.Write("Format_Folder", @"\\10.212.1.200\rgo\OK2SHIP automation system\0.OK2SHIP report format", "SMT_Config");
                TDMK_init.Write("Log_folder", @"\\10.212.1.200\rgo\OK2SHIP automation system\TDMK_Data\TestAreas\OK2SHIP_Measurements", "SMT_Config");
            }
            string server_name = TDMK_init.Read("Server", "SMT_Config");
            string server_acc = TDMK_init.Read("Account", "SMT_Config");
            string server_pass = TDMK_init.Read("Password", "SMT_Config");
            Data_Location = TDMK_init.Read("Data_Location", "SMT_Config");
            format_folder = TDMK_init.Read("Format_Folder", "SMT_Config");
            log_folder = TDMK_init.Read("Log_folder", "SMT_Config");
            Report_location = TDMK_init.Read("Report_Location", "SMT_Config");
            if (sa_en)
            {
                string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, DB_name, server_acc, server_pass).ConnectionString;
                _sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);
            }
            else
            {
                string _strcon = TDMK_Code.data_connection2(server_name, DB_name).ConnectionString;
                _sqlcon_OK2SHIP = new SqlConnection(_strcon);
            }
            return _sqlcon_OK2SHIP;
        }
        public string find_format(string in_data_loc, string ItemCode, List<string> extensions)
        {
            string result = "";

            DirectoryInfo directory = new DirectoryInfo(in_data_loc);
            var files = extensions.SelectMany(e => directory.EnumerateFiles(e, SearchOption.AllDirectories)).Where(x => x.FullName.Contains(ItemCode)).ToList();
            if (files.Count > 0)
            {
                return files[0].FullName;
            }
            return result;
        }
        public SortedDictionary<int, string> Get_bending_Net(SqlConnection sqlcon, string ItemCode, string Process_name)
        {
            DataTable temp_spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, Process_name }));
            DataTable spec_dt = temp_spec_dt.AsDataView().ToTable(false, new string[] { "Point+V", "Point-V", "LSL", "USL" });
            SortedDictionary<int, string> dic_index_lst = new SortedDictionary<int, string>();
            int inx = 0;
            foreach (DataRow dr in temp_spec_dt.Rows)
            {
                if (myCode.checkDBNull(dr["Sel_Report"]) != "No")
                {
                    string cur_net = dr["Net_Name"].ToString();
                    dic_index_lst.Add(inx, cur_net);
                }
                inx++;
            }
            return dic_index_lst;
        }
        public void Get_ListTable(int col_inx, DataTable myDt, string[] src_arr, ref List<DataTable> src_lst_tbl, string tar_item)
        {
            DataRow[] temp_dr;
            if (col_inx < src_arr.Length - 1)
            {
                if (src_arr[col_inx + 1] != "")
                {
                    string[] sel_val = myDt.AsEnumerable().Select(x => x.Field<string>(src_arr[col_inx + 1])).Distinct().ToArray();
                    if (sel_val.Length != 0)
                    {
                        foreach (string sv in sel_val)
                        {
                            if (sv != null)
                            {
                                DataTable curTbl = myDt.AsEnumerable().Where(r => r.Field<string>(src_arr[col_inx + 1]) == sv).CopyToDataTable();
                                Get_ListTable(col_inx + 1, curTbl, src_arr, ref src_lst_tbl, tar_item);
                            }
                        }
                    }
                    else
                    {
                        temp_dr = myDt.AsEnumerable().Where(x => x.Field<string>(tar_item) != null).ToArray();
                        if (temp_dr.Length > 0)
                        {
                            src_lst_tbl.Add(myDt);
                        }
                    }
                }
            }
            else
            {
                temp_dr = myDt.AsEnumerable().Where(x => x.Field<string>(tar_item) != null).ToArray();
                if (temp_dr.Length > 0)
                {
                    src_lst_tbl.Add(myDt);
                }
            }
        }
        public string remove_special_char(string src_str, List<char> remove_char)
        {
            return new string(src_str.Where(x => remove_char.IndexOf(x) == -1).ToArray());
        }
        public string Extract_Num_from_String(string str_in)
        {
            string result = "";
            result = new string(str_in.Where(x => char.IsDigit(x)).ToArray());
            return result;
        }
        public Dictionary<string, List<string>> Get_List_Pair_data(DataTable src_tbl)
        {
            Dictionary<string, List<string>> data_col = new Dictionary<string, List<string>>();
            foreach (DataColumn dc in src_tbl.Columns)
            {
                List<string> col_data = new List<string>();
                if ((dc.ColumnName == "Before") || (dc.ColumnName.Contains("After")))
                {
                    foreach (DataRow dr in src_tbl.Rows)
                    {
                        string cell_val = myCode.checkDBNull(dr[dc]);
                        if (cell_val != "")
                        {
                            col_data.Add(dr[dc].ToString());
                        }
                    }
                }
                if (col_data.Count > 0)
                {
                    data_col.Add(dc.ColumnName, col_data);
                }
            }
            return data_col;
        }
        public bool check_in_limit(string src_UL, string src_LL, string src_act_val)
        {
            bool _result = false;
            bool UL_equal_comp = true;
            bool LL_equal_comp = true;
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

                            _result = true;
                        }
                        else
                        {
                            if (act_val > UL)
                            {
                                _result = false;
                            }
                            else
                            {
                                _result = false;
                            }
                        }
                    }
                    else
                    {
                        if (act_val < UL)
                        {
                            _result = true;
                        }
                        else
                        {
                            if (UL_equal_comp)
                            {
                                if (act_val == UL)
                                {
                                    _result = true;
                                }
                                else
                                {
                                    _result = false;
                                }
                            }
                            else
                            {
                                _result = false;
                            }
                        }

                    }
                }
                else
                {
                    if (src_LL != "")
                    {
                        double LL = Convert.ToDouble(src_LL);
                        double act_val = Convert.ToDouble(src_act_val);
                        if (act_val > LL)
                        {
                            _result = true;
                        }
                        else
                        {
                            if (LL_equal_comp)
                            {
                                if (act_val == LL)
                                {
                                    _result = true;
                                }
                                else
                                {
                                    _result = false;
                                }
                            }
                            else
                            {
                                _result = false;
                            }
                        }
                    }
                    else
                    {
                        _result = false;
                    }
                }
            }
            return _result;
        }
        public void Export_Thermal_HeatSoak_Bend_All_old(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name, int qty, string format_type)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            string format_loc = Path.Combine(format_folder, "Format", format_type);
            string format_file = find_format(format_loc, ItemCode, new List<string> { "*.xlsx", "*.xlsm" });
            if (format_file != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable _src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                DataTable _spec_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process_name }));
                SortedDictionary<int, string> sel_index_lst = Get_bending_Net(sqlcon_OK2SHIP, ItemCode, process_name);
                List<int> index_lst = sel_index_lst.Keys.ToList();
                DataTable src_tbl = _src_tbl.AsEnumerable().Where(x => index_lst.Contains(Convert.ToInt32(x.Field<string>("Net_no")) - 1)).CopyToDataTable();
                DataTable spec_tbl = _spec_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_Report") != "No").CopyToDataTable();
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                List<char> reject_char_lst = new List<char> { ' ', '_', '-', '&' };
                List<string> Net_name = spec_tbl.AsEnumerable().Select(x => x.Field<string>("Net_Name")).ToList();
                ExcelWorkbook report_wrk = null; //create_export_wrk(format_file, format_name);
                ExcelWorksheet tar_wrksht = null;// report_wrk.Sheets[1];
                FileInfo format_file_info = new FileInfo(format_file);
                ExcelPackage report_pack = null;
                string tar_wrksht_name = "";
                if (format_type == "NPI")
                {
                    report_wrk = Excel_Lib.open_excel_file(format_file);
                    foreach (ExcelWorksheet sht in report_wrk.Worksheets)
                    {
                        if (remove_special_char(sht.Name.ToUpper(), reject_char_lst) == remove_special_char(format_name.ToUpper(), reject_char_lst))
                        {
                            tar_wrksht_name = sht.Name;
                            break;
                        }
                    }
                }
                else
                {
                    report_wrk = Excel_Lib.open_excel_file(format_file);
                    foreach (ExcelWorksheet sht in report_wrk.Worksheets)
                    {
                        if (remove_special_char(sht.Name, reject_char_lst) == remove_special_char("Bending < 10%", reject_char_lst))
                        {
                            tar_wrksht_name = sht.Name;
                            break;
                        }
                    }
                }
                if (tar_wrksht_name != "")
                {
                    string[] temp = Path.GetFileNameWithoutExtension(format_file_info.Name).Split('-');
                    string itemname = "";
                    if (temp.Length > 1)
                    {
                        itemname = temp[1];
                    }
                    string report_name = process_name + " " + itemname + "-" + ItemCode + "-" + LotNo + " " + DateTime.Now.ToString("yyyyMMdd") + Path.GetExtension(format_file_info.Name);
                    if (format_type == "NPI")
                    {
                        report_name = Path.GetFileNameWithoutExtension(format_file_info.Name) + "-" + LotNo + " " + DateTime.Now.ToString("yyyyMMdd") + Path.GetExtension(format_file_info.Name);
                    }
                    string daily_folder = Path.Combine(Report_location, "OK2SHIP_report");
                    if (!Directory.Exists(daily_folder))
                    {
                        Directory.CreateDirectory(daily_folder);
                    }
                    string daily_report_name = Path.Combine(daily_folder, report_name);
                    if (File.Exists(daily_report_name))
                    {
                        report_pack = Excel_Lib.open_excel(daily_report_name);
                    }
                    else
                    {
                        report_pack = new ExcelPackage(new FileInfo(format_file));
                        report_pack.SaveAs(new FileInfo(daily_report_name));
                    }
                    report_wrk = report_pack.Workbook;
                    tar_wrksht = report_wrk.Worksheets[tar_wrksht_name];
                    string Echeck_start_rgn = Excel_Lib.Find_Start_Addr("A1", tar_wrksht, false, "Condition");// Get_start_range("Condition", "A1", tar_wrksht);
                    ExcelRangeBase Echeck_cycle_rgn = tar_wrksht.Cells[Echeck_start_rgn];
                    Dictionary<string, string> Echeck_Cycle_addr = Get_Echeck_address(Echeck_cycle_rgn);
                    List<string> data_addr_lst = Get_Bending_data_addr("Test item", "Sample", tar_wrksht, "A10", true);
                    string _NET_start_rgn = Excel_Lib.Find_Cell_Addr("Test item", "A10", tar_wrksht, false);
                    string NET_start_rgn = Excel_Lib.get_offset_addr(_NET_start_rgn, tar_wrksht, 1, false);
                    ExcelRangeBase NET_rgn = tar_wrksht.Cells[NET_start_rgn];
                    int sel_qty = src_tbl_lst.Count;
                    for (int col_inx = 0; col_inx < sel_qty; col_inx++)
                    {
                        DataTable tbl = src_tbl_lst[col_inx];
                        Dictionary<string, List<string>> data_lst = Get_List_Pair_data(tbl);
                        List<string> before_data = data_lst["Before"];
                        List<string> last_data = data_lst.Values.ToList().Last();
                        ExcelRangeBase cycle_rgn = tar_wrksht.Cells[data_addr_lst[col_inx]].Offset(2, 0);
                        foreach (var item in data_lst)
                        {
                            List<string> after_data = item.Value;
                            if (Echeck_Cycle_addr.Keys.ToList().IndexOf(item.Key) != -1)
                            {
                                ExcelRangeBase sum_rgn_offset = tar_wrksht.Cells[Echeck_Cycle_addr[item.Key]].Offset(0, 2 + col_inx);
                                bool data_OK = false;
                                for (int r_inx = 0; r_inx < after_data.Count; r_inx++)
                                {
                                    string USL = myCode.checkDBNull(spec_tbl.Rows[r_inx]["USL"]);
                                    string LSL = myCode.checkDBNull(spec_tbl.Rows[r_inx]["LSL"]);
                                    string cur_val = after_data[r_inx];
                                    if (check_in_limit(USL, LSL, cur_val))
                                    {
                                        data_OK = true;
                                    }
                                    else
                                    {
                                        data_OK = false;
                                        break;
                                    }
                                }
                                if (data_OK)
                                {
                                    sum_rgn_offset.Value = "PASS";
                                }
                                else
                                {
                                    sum_rgn_offset.Value = "FAIL";
                                }
                            }
                        }
                        for (int i = 0; i < before_data.Count; i++)
                        {
                            cycle_rgn.Offset(i, 0).Value = before_data[i];
                            if (i < last_data.Count)
                            {
                                cycle_rgn.Offset(i, 1).Value = last_data[i];
                            }
                            cycle_rgn.Offset(i, 2).FormulaR1C1 = "=RC[-1]/RC[-2]-1";
                            cycle_rgn.Offset(i, 2).Style.Numberformat.Format = "#0.00%";
                            if (Math.Abs(Convert.ToDouble(cycle_rgn.Offset(i, 2).Value)) > 0.1)
                            {
                                cycle_rgn.Offset(i, 2).Style.Fill.PatternType = ExcelFillStyle.Solid;
                                cycle_rgn.Offset(i, 2).Style.Fill.BackgroundColor.SetColor(Color.Red);
                            }
                        }
                    }
                    for (int i = 0; i < Net_name.Count; i++)
                    {
                        NET_rgn.Offset(i, 0).Value = Net_name[i];
                    }
                    string last_cycle_addr = data_addr_lst.LastOrDefault();
                    TDMK_EPPLUS7_lib.MergeAreas_Info Cur_Cells_info = Excel_Lib.Get_Cells_Info(tar_wrksht, tar_wrksht.Cells[last_cycle_addr]);
                    int col_num = Cur_Cells_info.col_qty;
                    int total_col_num = tar_wrksht.Cells[last_cycle_addr].Start.Column + col_num - 1;
                    int r_off = Net_name.Count - 1;
                    int start_rgn_row = tar_wrksht.Cells[NET_start_rgn].Start.Row;
                    int start_rgn_col = tar_wrksht.Cells[NET_start_rgn].Start.Column;
                    ExcelRangeBase final_rgn = tar_wrksht.Cells[start_rgn_row, start_rgn_col, start_rgn_row + r_off, total_col_num];
                    Excel_Lib.Draw_border(final_rgn);
                    report_pack.Save();
                }
                else
                {
                    MessageBox.Show("Không tìm thấy sheet " + process_name, "Thông báo");
                }
            }
            else
            {
                MessageBox.Show("Format of ItemCode " + ItemCode + " not found", "Warning");
            }
        }
        public void Export_Thermal_HeatSoak_Bend_All(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name, int qty, string format_type, ExcelWorksheet tar_wrksht)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            DataTable _src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
            DataTable _spec_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process_name }));
            SortedDictionary<int, string> sel_index_lst = Get_bending_Net(sqlcon_OK2SHIP, ItemCode, process_name);
            List<int> index_lst = sel_index_lst.Keys.ToList();
            DataTable src_tbl = _src_tbl.AsEnumerable().Where(x => index_lst.Contains(Convert.ToInt32(x.Field<string>("Net_no")) - 1)).CopyToDataTable();
            DataTable spec_tbl = _spec_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_Report") != "No").CopyToDataTable();
            DataTable log_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name + "_LOGFILE", filter_str);
            List<DataTable> src_tbl_lst = new List<DataTable>();
            Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
            List<char> reject_char_lst = new List<char> { ' ', '_', '-', '&' };
            List<string> Net_name = spec_tbl.AsEnumerable().Select(x => x.Field<string>("Net_Name")).ToList();
            string condition_addr = Excel_Lib.Find_Cell_Addr("Condition", "A1", tar_wrksht, false);
            string Echeck_start_rgn = Excel_Lib.Find_Start_Addr("A1", tar_wrksht, false, "Condition");// Get_start_range("Condition", "A1", tar_wrksht);
            ExcelRangeBase Echeck_cycle_rgn = tar_wrksht.Cells[Echeck_start_rgn];
            string result_rgn_addr = Excel_Lib.Find_Cell_Addr("Due Date",condition_addr, tar_wrksht, true);
            int offset_val = Excel_Lib.Get_Cells_Info(tar_wrksht, tar_wrksht.Cells[result_rgn_addr]).col_qty+1;
            int offset_duedate = tar_wrksht.Cells[result_rgn_addr].End.Column - tar_wrksht.Cells[condition_addr].End.Column;
            Dictionary<string, string> Echeck_Cycle_addr = Get_Echeck_address(Echeck_cycle_rgn);
            List<string> data_addr_lst = Get_Bending_data_addr("Test item", "Sample", tar_wrksht, "A10", true);
            string _NET_start_rgn = Excel_Lib.Find_Cell_Addr("Test item", "A10", tar_wrksht, false);
            string NET_start_rgn = Excel_Lib.get_offset_addr(_NET_start_rgn, tar_wrksht, 1, false);
            ExcelRangeBase NET_rgn = tar_wrksht.Cells[NET_start_rgn];
            int sel_qty = src_tbl_lst.Count;
            for (int col_inx = 0; col_inx < sel_qty; col_inx++)
            {
                DataTable tbl = src_tbl_lst[col_inx];
                Dictionary<string, List<string>> data_lst = Get_List_Pair_data(tbl);
                List<string> before_data = data_lst["Before"];
                List<string> last_data = data_lst.Values.ToList().Last();
                ExcelRangeBase cycle_rgn = tar_wrksht.Cells[data_addr_lst[col_inx]].Offset(2, 0);
                foreach (var item in data_lst)
                {
                    List<string> after_data = item.Value;
                    string cycle_filter = "BF";
                    if (item.Key.Contains("After"))
                    {
                        cycle_filter = "L" + Extract_Num_from_String(item.Key);
                    }
                    List<string> log_time_lst = log_tbl.AsEnumerable().Where(x => x.Field<string>("Cycles_name") == item.Key).Select(x => x.Field<string>("Remark")).Distinct().ToList();
                    string log_time = "";
                    if (log_time_lst.Count > 0)
                    {
                        string _log_time = log_tbl.AsEnumerable().Where(x => x.Field<string>("Cycles_name") == item.Key).Select(x => x.Field<string>("Logfile")).FirstOrDefault();
                        log_time = _log_time.Split('-').First() + " " + log_time_lst.Last();
                    }
                    if (col_inx<1)
                    {
                        ExcelRangeBase time_rgn_offset = tar_wrksht.Cells[Echeck_Cycle_addr[item.Key]].Offset(0, offset_duedate);
                        time_rgn_offset.Value = log_time.Split(' ').FirstOrDefault();
                        time_rgn_offset.Offset(1,0).Value = log_time.Split(' ').FirstOrDefault();
                    }    
                    if (Echeck_Cycle_addr.Keys.ToList().IndexOf(item.Key) != -1)
                    {
                        ExcelRangeBase sum_rgn_offset = tar_wrksht.Cells[Echeck_Cycle_addr[item.Key]].Offset(0, offset_val + col_inx);
                        bool data_OK = false;
                        for (int r_inx = 0; r_inx < after_data.Count; r_inx++)
                        {
                            string USL = myCode.checkDBNull(spec_tbl.Rows[r_inx]["USL"]);
                            string LSL = myCode.checkDBNull(spec_tbl.Rows[r_inx]["LSL"]);
                            string cur_val = after_data[r_inx];
                            if (check_in_limit(USL, LSL, cur_val))
                            {
                                data_OK = true;
                            }
                            else
                            {
                                data_OK = false;
                                break;
                            }
                        }
                        if (data_OK)
                        {
                            sum_rgn_offset.Value = "PASS";
                        }
                        else
                        {
                            sum_rgn_offset.Value = "FAIL";
                        }
                    }
                }
                for (int i = 0; i < before_data.Count; i++)
                {
                    cycle_rgn.Offset(i, 0).Value = before_data[i];
                    if (i < last_data.Count)
                    {
                        cycle_rgn.Offset(i, 1).Value = last_data[i];
                    }
                    cycle_rgn.Offset(i, 2).FormulaR1C1 = "=RC[-1]/RC[-2]-1";
                    cycle_rgn.Offset(i, 2).Style.Numberformat.Format = "#0.00%";
                    if (Math.Abs(Convert.ToDouble(cycle_rgn.Offset(i, 2).Value)) > 0.1)
                    {
                        cycle_rgn.Offset(i, 2).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cycle_rgn.Offset(i, 2).Style.Fill.BackgroundColor.SetColor(Color.Red);
                    }
                }
            }
            for (int i = 0; i < Net_name.Count; i++)
            {
                NET_rgn.Offset(i, 0).Value = Net_name[i];
            }
            string last_cycle_addr = data_addr_lst.LastOrDefault();
            TDMK_EPPLUS7_lib.MergeAreas_Info Cur_Cells_info = Excel_Lib.Get_Cells_Info(tar_wrksht, tar_wrksht.Cells[last_cycle_addr]);
            int col_num = Cur_Cells_info.col_qty;
            int total_col_num = tar_wrksht.Cells[last_cycle_addr].Start.Column + col_num - 1;
            int r_off = Net_name.Count - 1;
            int start_rgn_row = tar_wrksht.Cells[NET_start_rgn].Start.Row;
            int start_rgn_col = tar_wrksht.Cells[NET_start_rgn].Start.Column;
            ExcelRangeBase final_rgn = tar_wrksht.Cells[start_rgn_row, start_rgn_col, start_rgn_row + r_off, total_col_num];
            Excel_Lib.Draw_border(final_rgn);

        }
        public void Export_Thermal_HeatSoak_Bend_MultiType(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name, int tar_pcs, string format_type)
        {
            string format_loc = Path.Combine(format_folder, format_type);
            //string format_file = find_format(format_folder, ItemCode, new List<string> { "*.xlsx", "*.xlsm" });
            List<string> format_lst = get_multiple_files(format_loc, new List<string> { "*.xlsx", "*.xlsm" }, ItemCode, process_name);
            string format_file = "";
            if (format_lst.Count > 0)
            {
                format_file = format_lst[0];
            }
            if (format_file != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable _src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                DataTable _spec_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process_name }));
                DataTable log_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name + "_LOGFILE", filter_str);
                SortedDictionary<int, string> sel_index_lst = Get_bending_Net(sqlcon_OK2SHIP, ItemCode, process_name);
                List<int> index_lst = sel_index_lst.Keys.ToList();
                DataTable src_tbl = _src_tbl.AsEnumerable().Where(x => index_lst.Contains(Convert.ToInt32(x.Field<string>("Net_no")) - 1)).CopyToDataTable();
                DataTable spec_tbl = _spec_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_Report") != "No").CopyToDataTable();
                List<string> NET_name = spec_tbl.AsEnumerable().Select(x => x.Field<string>("Net_Name")).ToList();
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                List<char> reject_char_lst = new List<char> { ' ', '_', '-', '\r', '\n' };
                ExcelWorkbook report_wrk = null;// 
                ExcelWorksheet tar_wrksht = null;
                ExcelWorksheet result_sht = null;
                ExcelPackage report_pack = null;
                string tar_wrksht_name = "";
                string result_sht_name = "";
                string find_cycle_key = "Bending";
                if (format_type == "NPI")
                {
                    find_cycle_key = "Sample";
                    report_wrk = Excel_Lib.open_excel_file(format_file);
                    foreach (ExcelWorksheet sht in report_wrk.Worksheets)
                    {
                        if (remove_special_char(sht.Name, reject_char_lst) == remove_special_char(process_name, reject_char_lst))
                        {
                            tar_wrksht_name = sht.Name;
                            break;
                        }
                    }
                }
                else
                {
                    report_wrk = Excel_Lib.open_excel_file(format_file);
                    foreach (ExcelWorksheet sht in report_wrk.Worksheets)
                    {
                        if (remove_special_char(sht.Name, reject_char_lst) == remove_special_char("Bending < 10%", reject_char_lst))
                        {
                            tar_wrksht_name = sht.Name;
                            break;
                        }
                    }
                    foreach (ExcelWorksheet sht in report_wrk.Worksheets)
                    {
                        if (remove_special_char(sht.Name, reject_char_lst) == remove_special_char("Result", reject_char_lst))
                        {
                            result_sht_name = sht.Name;
                            break;
                        }
                    }
                }
                if (tar_wrksht_name != "")
                {
                    if (tar_pcs <= src_tbl_lst.Count)
                    {
                        string itemname = "";
                        string UserID = "";
                        itemname = myCode.checkDBNull(log_tbl.Rows[0]["ItemName"]);
                        UserID = myCode.checkDBNull(log_tbl.Rows[0]["UserID"]);
                        FileInfo report_wrk_info = new FileInfo(format_file);
                        string report_name = process_name + " " + itemname + "-" + ItemCode + "-" + LotNo + " " + DateTime.Now.ToString("yyyyMMdd") + Path.GetExtension(report_wrk_info.Name);
                        if (format_type == "NPI")
                        {
                            report_name = Path.GetFileNameWithoutExtension(report_wrk_info.Name) + "-" + LotNo + " " + DateTime.Now.ToString("yyyyMMdd") + Path.GetExtension(report_wrk_info.Name);
                        }
                        string daily_folder = Path.Combine(Report_location, "Daily_report");
                        if (!Directory.Exists(daily_folder))
                        {
                            Directory.CreateDirectory(daily_folder);
                        }
                        string daily_report = Path.Combine(daily_folder, report_name);
                        if (File.Exists(daily_report))
                        {
                            report_pack = Excel_Lib.open_excel(daily_report);
                        }
                        else
                        {
                            report_pack = new ExcelPackage(new FileInfo(format_file));
                            report_pack.SaveAs(new FileInfo(daily_report));
                        }
                        report_wrk = report_pack.Workbook;
                        tar_wrksht = report_wrk.Worksheets[tar_wrksht_name];
                        Dictionary<int, string> cycle_data_addr_lst = Get_Bending_Cycle_addr("Test item", find_cycle_key, tar_wrksht, "A10", false);
                        string NET_start_rgn = Excel_Lib.Find_Start_Addr("A10", tar_wrksht, false, "Test item");// Get_start_range("Test item", "A10", tar_wrksht);
                        ExcelRangeBase NET_rgn = tar_wrksht.Cells[NET_start_rgn];
                        string Echeck_start_rgn = Excel_Lib.Find_Start_Addr("A1", tar_wrksht, false, "Condition");// Get_start_range("Condition", "A1", tar_wrksht);
                        ExcelRangeBase Echeck_cycle_rgn = tar_wrksht.Cells[Echeck_start_rgn];
                        Dictionary<string, string> Echeck_Cycle_addr = Get_Echeck_address(Echeck_cycle_rgn);
                        string proId_addr = "";
                        string itemname_addr = "";
                        string itemcode_addr = "";
                        string lotno_addr = "";
                        string process_addr = "";
                        string cycle_addr = "";
                        string time_addr = "";
                        string result_addr = "";
                        string operator_addr = "";
                        if (result_sht_name != "")
                        {
                            result_sht = report_wrk.Worksheets[result_sht_name];
                            proId_addr = Excel_Lib.Find_Cell_Addr("ProductID", "A1", result_sht, true);// Get_start_range("ProductID", "A1", result_sht, true);
                            itemname_addr = Excel_Lib.Find_Cell_Addr("Item name", "A1", result_sht, true);// Get_start_range("Item name", "A1", result_sht, true);
                            itemcode_addr = Excel_Lib.Find_Cell_Addr("Item code", "A1", result_sht, true); //Get_start_range("Item code", "A1", result_sht, true);
                            lotno_addr = Excel_Lib.Find_Cell_Addr("Item lot", "A1", result_sht, true); //Get_start_range("Item lot", "A1", result_sht, true);
                            process_addr = Excel_Lib.Find_Cell_Addr("Test Condition", "A1", result_sht, true); //Get_start_range("Test Condition", "A1", result_sht, true);
                            cycle_addr = Excel_Lib.Find_Cell_Addr("Bending cycle", "A1", result_sht, true); //Get_start_range("Bending cycle", "A1", result_sht, true);
                            time_addr = Excel_Lib.Find_Cell_Addr(@"ICT Datetime Finish", "A1", result_sht, true); //Get_start_range(@"ICT Datetime Finish", "A1", result_sht, true);
                            result_addr = Excel_Lib.Find_Cell_Addr("Result", "A1", result_sht, true); //Get_start_range("Result", "A1", result_sht, true);
                            operator_addr = Excel_Lib.Find_Cell_Addr("Output OperatorID", "A1", result_sht, true); //Get_start_range("Output OperatorID", "A1", result_sht, true); 
                        }
                        int col_inx = tar_pcs - 1;
                        DataTable tbl = src_tbl_lst[col_inx];
                        string prod_id = tbl.Rows[0]["Pcs_No"].ToString();
                        Dictionary<string, List<string>> data_lst = Get_List_Pair_data(tbl);
                        List<string> before_data = data_lst["Before"];
                        int r_offset = 0;
                        foreach (var item in data_lst)
                        {
                            List<string> sel_data = item.Value;
                            string cycle_filter = "BF";
                            if (item.Key.Contains("After"))
                            {
                                cycle_filter = "L" + Extract_Num_from_String(item.Key);
                            }
                            List<string> log_time_lst = log_tbl.AsEnumerable().Where(x => x.Field<string>("Cycles_name") == item.Key).Select(x => x.Field<string>("Remark")).Distinct().ToList();
                            string log_time = "";
                            if (log_time_lst.Count > 0)
                            {
                                string _log_time = log_tbl.AsEnumerable().Where(x => x.Field<string>("Cycles_name") == item.Key).Select(x => x.Field<string>("Logfile")).FirstOrDefault();
                                log_time = _log_time.Split('-').First() + " " + log_time_lst.Last();
                            }
                            bool data_OK = false;
                            if (Echeck_Cycle_addr.Keys.ToList().IndexOf(item.Key) != -1)
                            {
                                ExcelRangeBase sum_rgn_offset = tar_wrksht.Cells[Echeck_Cycle_addr[item.Key]].Offset(0, 2 + col_inx);
                                ExcelRangeBase time_rgn_offset = tar_wrksht.Cells[Echeck_Cycle_addr[item.Key]].Offset(0, 1 + col_inx);
                                time_rgn_offset.Value = log_time;

                                for (int r_inx = 0; r_inx < sel_data.Count; r_inx++)
                                {
                                    string USL = myCode.checkDBNull(spec_tbl.Rows[r_inx]["USL"]);
                                    string LSL = myCode.checkDBNull(spec_tbl.Rows[r_inx]["LSL"]);
                                    string cur_val = sel_data[r_inx];
                                    if (check_in_limit(USL, LSL, cur_val))
                                    {
                                        data_OK = true;
                                    }
                                    else
                                    {
                                        data_OK = false;
                                        break;
                                    }
                                }
                                if (data_OK)
                                {
                                    sum_rgn_offset.Value = "OK";
                                }
                                else
                                {
                                    sum_rgn_offset.Value = "FAIL";
                                }
                            }
                            if (item.Key.Contains("After"))
                            {
                                int cycle_no = Convert.ToInt32(Extract_Num_from_String(item.Key));
                                if (cycle_data_addr_lst.Keys.ToList().IndexOf(cycle_no) != -1)
                                {
                                    ExcelRangeBase cycle_rgn = tar_wrksht.Cells[cycle_data_addr_lst[cycle_no]].Offset(2, 0);
                                    for (int r_inx = 0; r_inx < sel_data.Count; r_inx++)
                                    {
                                        cycle_rgn.Offset(r_inx, 0).Value = before_data[r_inx];
                                        cycle_rgn.Offset(r_inx, 1).Value = sel_data[r_inx];
                                        cycle_rgn.Offset(r_inx, 2).FormulaR1C1 = "=RC[-1]/RC[-2]-1";
                                        cycle_rgn.Offset(r_inx, 2).Style.Numberformat.Format = "#0.00%";
                                        if (Math.Abs(Convert.ToDouble(cycle_rgn.Offset(r_inx, 2).Value)) > 0.1)
                                        {
                                            cycle_rgn.Offset(r_inx, 2).Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            cycle_rgn.Offset(r_inx, 2).Style.Fill.BackgroundColor.SetColor(Color.Red);
                                        }
                                    }
                                }
                            }
                            if (result_sht_name != "")
                            {

                                if (item.Key.Contains("After"))
                                {
                                    result_sht.Cells[proId_addr].Offset(r_offset + 1, 0).Value = prod_id;
                                    result_sht.Cells[itemname_addr].Offset(r_offset + 1, 0).Value = itemname;
                                    result_sht.Cells[itemcode_addr].Offset(r_offset + 1, 0).Value = ItemCode;
                                    result_sht.Cells[lotno_addr].Offset(r_offset + 1, 0).Value = LotNo;
                                    result_sht.Cells[process_addr].Offset(r_offset + 1, 0).Value = process_name;
                                    result_sht.Cells[cycle_addr].Offset(r_offset + 1, 0).Value = cycle_filter.Replace("L", "");
                                    result_sht.Cells[time_addr].Offset(r_offset + 1, 0).Value = log_time;
                                    result_sht.Cells[operator_addr].Offset(r_offset + 1, 0).Value = UserID;
                                    if (data_OK)
                                    {
                                        result_sht.Cells[result_addr].Offset(r_offset + 1, 0).Value = "Pass";
                                    }
                                    else
                                    {
                                        result_sht.Cells[result_addr].Offset(r_offset + 1, 0).Value = "NG";
                                    }
                                    r_offset++;
                                }
                            }

                        }
                        for (int i = 0; i < NET_name.Count; i++)
                        {
                            NET_rgn.Offset(i, 0).Value = NET_name[i];
                        }
                        string last_cycle_addr = cycle_data_addr_lst.LastOrDefault().Value;
                        TDMK_EPPLUS7_lib.MergeAreas_Info Cur_Cells_info = Excel_Lib.Get_Cells_Info(tar_wrksht, tar_wrksht.Cells[last_cycle_addr]);
                        int col_num = Cur_Cells_info.col_qty;
                        int total_col_num = tar_wrksht.Cells[last_cycle_addr].Start.Column + col_num - 1;
                        int r_off = NET_name.Count - 1;
                        int start_rgn_row = tar_wrksht.Cells[NET_start_rgn].Start.Row;
                        int start_rgn_col = tar_wrksht.Cells[NET_start_rgn].Start.Column;
                        ExcelRangeBase final_rgn = tar_wrksht.Cells[start_rgn_row, start_rgn_col, start_rgn_row + r_off, total_col_num];
                        Excel_Lib.Draw_border(final_rgn);
                        report_pack.Save();
                    }
                }
                else
                {
                    MessageBox.Show("Format của ItemCode " + ItemCode + " không đúng", "Cảnh báo");
                }
            }
            else
            {
                MessageBox.Show("Format of ItemCode " + ItemCode + " not found", "Warning");
            }
        }
        public List<string> get_multiple_files(string src_path, List<string> extensions, string tar_ItemCode, string tar_process)
        {
            List<string> result = new List<string>();
            List<char> remove_char = new List<char> { ' ', '-', '_' };
            string process = new string(tar_process.Where(x => remove_char.IndexOf(x) == -1).ToArray());
            DirectoryInfo directory = new DirectoryInfo(src_path);
            var files = extensions.SelectMany(e => directory.EnumerateFiles(e, SearchOption.AllDirectories)).Where(x => x.FullName.Contains(tar_ItemCode) && remove_special_char(x.FullName, remove_char).ToUpper().Contains(process.ToUpper()));
            foreach (var item in files)
            {
                result.Add(item.FullName);
            }
            return result;
        }
        public List<string> Get_Bending_data_addr(string search_item, string search_key, ExcelWorksheet tar_wrksht, string start_range_addr, bool col_direction)
        {
            List<string> result = new List<string>();
            string tar_addr = Excel_Lib.Find_Cell_Addr(search_item, start_range_addr, tar_wrksht, false);
            ExcelRangeBase sample_rgn = tar_wrksht.Cells[tar_addr];
            ExcelRangeBase sel_rgn = sample_rgn;
            int count_null = 0;
            while (true)
            {
                string rgn_addr = sel_rgn.Address;
                ExcelRangeBase temp_rgn = tar_wrksht.Cells[Excel_Lib.Find_Start_Addr(rgn_addr, tar_wrksht, col_direction)];
                string temp_addr = temp_rgn.Address;
                string temp_rgn_val = myCode.checkDBNull(temp_rgn.Value);
                if (temp_rgn_val.ToUpper().Contains(search_key.ToUpper()))
                {
                    result.Add(temp_rgn.Address);
                    count_null = 0;
                }
                else
                {
                    if (temp_rgn_val == "")
                    {
                        count_null++;
                    }
                }
                if (count_null >= 2)
                {
                    break;
                }
                sel_rgn = temp_rgn;
            }
            return result;
        }
        public Dictionary<string, string> Get_Echeck_address(ExcelRangeBase cycle_rgn)
        {
            Dictionary<string, string> bending_Cycle_addr = new Dictionary<string, string>();
            int row_inx = 0;
            while (myCode.checkDBNull(cycle_rgn.Offset(row_inx, 0).Value) != "")
            {
                string rgn_val = myCode.checkDBNull(cycle_rgn.Offset(row_inx, 0).Value);
                if (rgn_val.ToUpper().Contains("E-CHECK"))
                {
                    string cycle = Extract_Num_from_String(rgn_val);// Get_Number_String(cycle_rgn.Offset(row_inx, 0).Value, ' ');
                    string cycle_name = "";
                    if (cycle != "")
                    {
                        cycle_name = "After_" + cycle;
                    }
                    else
                    {
                        cycle_name = "Before";
                    }
                    if ((cycle_name != "") && (bending_Cycle_addr.Keys.ToList().IndexOf(cycle_name) == -1))
                    {
                        bending_Cycle_addr.Add(cycle_name, cycle_rgn.Offset(row_inx, 0).Address);
                    }
                }
                row_inx++;
            }
            return bending_Cycle_addr;
        }
        public Dictionary<int, string> Get_Bending_Cycle_addr(string search_item, string search_key, ExcelWorksheet tar_wrksht, string start_range_addr, bool row_direction)
        {
            Dictionary<int, string> result = new Dictionary<int, string>();
            string tar_addr = Excel_Lib.Find_Cell_Addr(search_item, start_range_addr, tar_wrksht, false);
            ExcelRangeBase sample_rgn = tar_wrksht.Cells[tar_addr];
            ExcelRangeBase sel_rgn = sample_rgn;
            int count_null = 0;
            while (true)
            {
                string rgn_addr = sel_rgn.Address;
                ExcelRangeBase temp_rgn = tar_wrksht.Cells[Excel_Lib.Find_Start_Addr(rgn_addr, tar_wrksht, !row_direction)];// sel_rgn.Offset(0, 1);
                //if (row_direction)
                //{
                //    temp_rgn = sel_rgn.Offset(1, 0);
                //}
                string temp_addr = temp_rgn.Address;
                string temp_rgn_val = myCode.checkDBNull(temp_rgn.Value);
                if (temp_rgn_val.ToUpper().Contains(search_key.ToUpper()))
                {
                    int key = Convert.ToInt32(Extract_Num_from_String(temp_rgn_val));
                    if (result.Keys.ToList().IndexOf(key) == -1)
                    {
                        result.Add(key, temp_rgn.Address);
                    }
                    else
                    {
                        result[key] = temp_rgn.Address;
                    }
                    count_null = 0;
                }
                else
                {
                    if (temp_rgn_val == "")
                    {
                        count_null++;
                    }
                }
                if (count_null >= 2)
                {
                    break;
                }
                sel_rgn = temp_rgn;
            }
            return result;
        }
        public void Export_Thermal_HeatSoak_Bend_NPI(ExcelWorkbook report_wrk, SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string sheet_name, int qty)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelWorksheet tar_wrksht = null;
            string tar_wrksht_name = "";
            List<char> reject_char_lst = new List<char> { ' ', '_', '-', '&' };
            foreach (ExcelWorksheet sht in report_wrk.Worksheets)
            {
                if (remove_special_char(sht.Name.ToUpper(), reject_char_lst) == remove_special_char(sheet_name.ToUpper(), reject_char_lst))
                {
                    tar_wrksht_name = sht.Name;
                    break;
                }
            }
            if (tar_wrksht_name != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable _src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                DataTable _spec_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process_name }));
                SortedDictionary<int, string> sel_index_lst = Get_bending_Net(sqlcon_OK2SHIP, ItemCode, process_name);
                List<int> index_lst = sel_index_lst.Keys.ToList();
                DataTable src_tbl = _src_tbl.AsEnumerable().Where(x => index_lst.Contains(Convert.ToInt32(x.Field<string>("Net_no")) - 1)).CopyToDataTable();
                DataTable spec_tbl = _spec_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_Report") != "No").CopyToDataTable();
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                List<string> Net_name = spec_tbl.AsEnumerable().Select(x => x.Field<string>("Net_Name")).ToList();
                tar_wrksht = report_wrk.Worksheets[tar_wrksht_name];
                string Echeck_start_rgn = Excel_Lib.Find_Start_Addr("A1", tar_wrksht, false, "Condition");
                ExcelRangeBase Echeck_cycle_rgn = tar_wrksht.Cells[Echeck_start_rgn];
                Dictionary<string, string> Echeck_Cycle_addr = Get_Echeck_address(Echeck_cycle_rgn);
                List<string> data_addr_lst = Get_Bending_data_addr("Test item", "Sample", tar_wrksht, "A10", true);
                string _NET_start_rgn = Excel_Lib.Find_Cell_Addr("Test item", "A10", tar_wrksht, false);
                string NET_start_rgn = Excel_Lib.get_offset_addr(_NET_start_rgn, tar_wrksht, 1, false);
                ExcelRangeBase NET_rgn = tar_wrksht.Cells[NET_start_rgn];
                int sel_qty = src_tbl_lst.Count;
                for (int col_inx = 0; col_inx < sel_qty; col_inx++)
                {
                    DataTable tbl = src_tbl_lst[col_inx];
                    Dictionary<string, List<string>> data_lst = Get_List_Pair_data(tbl);
                    List<string> before_data = data_lst["Before"];
                    List<string> last_data = data_lst.Values.ToList().Last();
                    ExcelRangeBase cycle_rgn = tar_wrksht.Cells[data_addr_lst[col_inx]].Offset(2, 0);
                    foreach (var item in data_lst)
                    {
                        List<string> after_data = item.Value;
                        if (Echeck_Cycle_addr.Keys.ToList().IndexOf(item.Key) != -1)
                        {
                            ExcelRangeBase sum_rgn_offset = tar_wrksht.Cells[Echeck_Cycle_addr[item.Key]].Offset(0, 2 + col_inx);
                            bool data_OK = false;
                            for (int r_inx = 0; r_inx < after_data.Count; r_inx++)
                            {
                                string USL = myCode.checkDBNull(spec_tbl.Rows[r_inx]["USL"]);
                                string LSL = myCode.checkDBNull(spec_tbl.Rows[r_inx]["LSL"]);
                                string cur_val = after_data[r_inx];
                                if (check_in_limit(USL, LSL, cur_val))
                                {
                                    data_OK = true;
                                }
                                else
                                {
                                    data_OK = false;
                                    break;
                                }
                            }
                            if (data_OK)
                            {
                                sum_rgn_offset.Value = "PASS";
                            }
                            else
                            {
                                sum_rgn_offset.Value = "FAIL";
                            }
                        }
                    }
                    for (int i = 0; i < before_data.Count; i++)
                    {
                        cycle_rgn.Offset(i, 0).Value = before_data[i];
                        if (i < last_data.Count)
                        {
                            cycle_rgn.Offset(i, 1).Value = last_data[i];
                        }
                        cycle_rgn.Offset(i, 2).FormulaR1C1 = "=RC[-1]/RC[-2]-1";
                        cycle_rgn.Offset(i, 2).Style.Numberformat.Format = "#0.00%";
                        if (Math.Abs(Convert.ToDouble(cycle_rgn.Offset(i, 2).Value)) > 0.1)
                        {
                            cycle_rgn.Offset(i, 2).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            cycle_rgn.Offset(i, 2).Style.Fill.BackgroundColor.SetColor(Color.Red);
                        }
                    }
                }
                for (int i = 0; i < Net_name.Count; i++)
                {
                    NET_rgn.Offset(i, 0).Value = Net_name[i];
                }
                string last_cycle_addr = data_addr_lst.LastOrDefault();
                TDMK_EPPLUS7_lib.MergeAreas_Info Cur_Cells_info = Excel_Lib.Get_Cells_Info(tar_wrksht, tar_wrksht.Cells[last_cycle_addr]);
                int col_num = Cur_Cells_info.col_qty;
                int total_col_num = tar_wrksht.Cells[last_cycle_addr].Start.Column + col_num - 1;
                int r_off = Net_name.Count - 1;
                int start_rgn_row = tar_wrksht.Cells[NET_start_rgn].Start.Row;
                int start_rgn_col = tar_wrksht.Cells[NET_start_rgn].Start.Column;
                ExcelRangeBase final_rgn = tar_wrksht.Cells[start_rgn_row, start_rgn_col, start_rgn_row + r_off, total_col_num];
                Excel_Lib.Draw_border(final_rgn);
            }
            else
            {
                MessageBox.Show("Không tìm thấy sheet " + process_name, "Thông báo");
            }
        }
        public void Export_Thermal_HeatSoak_Bend_MASS(ExcelWorkbook report_wrk, SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, int tar_pcs)
        {
            List<char> reject_char_lst = new List<char> { ' ', '_', '-', '\r', '\n' };
            ExcelWorksheet tar_wrksht = null;
            ExcelWorksheet result_sht = null;
            string tar_wrksht_name = "";
            string result_sht_name = "";
            string find_cycle_key = "Sample";
            foreach (ExcelWorksheet sht in report_wrk.Worksheets)
            {
                if (remove_special_char(sht.Name, reject_char_lst) == remove_special_char("Bending < 10%", reject_char_lst))
                {
                    tar_wrksht_name = sht.Name;
                    break;
                }
            }
            foreach (ExcelWorksheet sht in report_wrk.Worksheets)
            {
                if (remove_special_char(sht.Name, reject_char_lst) == remove_special_char("Result", reject_char_lst))
                {
                    result_sht_name = sht.Name;
                    break;
                }
            }
            if (tar_wrksht_name != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable _src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                DataTable _spec_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process_name }));
                DataTable log_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name + "_LOGFILE", filter_str);
                SortedDictionary<int, string> sel_index_lst = Get_bending_Net(sqlcon_OK2SHIP, ItemCode, process_name);
                List<int> index_lst = sel_index_lst.Keys.ToList();
                DataTable src_tbl = _src_tbl.AsEnumerable().Where(x => index_lst.Contains(Convert.ToInt32(x.Field<string>("Net_no")) - 1)).CopyToDataTable();
                DataTable spec_tbl = _spec_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_Report") != "No").CopyToDataTable();
                List<string> NET_name = spec_tbl.AsEnumerable().Select(x => x.Field<string>("Net_Name")).ToList();
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                string itemname = "";
                string UserID = "";
                itemname = myCode.checkDBNull(log_tbl.Rows[0]["ItemName"]);
                UserID = myCode.checkDBNull(log_tbl.Rows[0]["UserID"]);
                if (tar_pcs <= src_tbl_lst.Count)
                {
                    tar_wrksht = report_wrk.Worksheets[tar_wrksht_name];
                    Dictionary<int, string> cycle_data_addr_lst = Get_Bending_Cycle_addr("Test item", find_cycle_key, tar_wrksht, "A10", false);
                    string NET_start_rgn = Excel_Lib.Find_Start_Addr("A10", tar_wrksht, false, "Test item");// Get_start_range("Test item", "A10", tar_wrksht);
                    ExcelRangeBase NET_rgn = tar_wrksht.Cells[NET_start_rgn];
                    string Echeck_start_rgn = Excel_Lib.Find_Start_Addr("A1", tar_wrksht, false, "Condition");// Get_start_range("Condition", "A1", tar_wrksht);
                    ExcelRangeBase Echeck_cycle_rgn = tar_wrksht.Cells[Echeck_start_rgn];
                    Dictionary<string, string> Echeck_Cycle_addr = Get_Echeck_address(Echeck_cycle_rgn);
                    string proId_addr = "";
                    string itemname_addr = "";
                    string itemcode_addr = "";
                    string lotno_addr = "";
                    string process_addr = "";
                    string cycle_addr = "";
                    string time_addr = "";
                    string result_addr = "";
                    string operator_addr = "";
                    if (result_sht_name != "")
                    {
                        result_sht = report_wrk.Worksheets[result_sht_name];
                        proId_addr = Excel_Lib.Find_Cell_Addr("ProductID", "A1", result_sht, true);// Get_start_range("ProductID", "A1", result_sht, true);
                        itemname_addr = Excel_Lib.Find_Cell_Addr("Item name", "A1", result_sht, true);// Get_start_range("Item name", "A1", result_sht, true);
                        itemcode_addr = Excel_Lib.Find_Cell_Addr("Item code", "A1", result_sht, true); //Get_start_range("Item code", "A1", result_sht, true);
                        lotno_addr = Excel_Lib.Find_Cell_Addr("Item lot", "A1", result_sht, true); //Get_start_range("Item lot", "A1", result_sht, true);
                        process_addr = Excel_Lib.Find_Cell_Addr("Test Condition", "A1", result_sht, true); //Get_start_range("Test Condition", "A1", result_sht, true);
                        cycle_addr = Excel_Lib.Find_Cell_Addr("Bending cycle", "A1", result_sht, true); //Get_start_range("Bending cycle", "A1", result_sht, true);
                        time_addr = Excel_Lib.Find_Cell_Addr(@"ICT Datetime Finish", "A1", result_sht, true); //Get_start_range(@"ICT Datetime Finish", "A1", result_sht, true);
                        result_addr = Excel_Lib.Find_Cell_Addr("Result", "A1", result_sht, true); //Get_start_range("Result", "A1", result_sht, true);
                        operator_addr = Excel_Lib.Find_Cell_Addr("Output OperatorID", "A1", result_sht, true); //Get_start_range("Output OperatorID", "A1", result_sht, true); 
                    }
                    int col_inx = tar_pcs - 1;
                    DataTable tbl = src_tbl_lst[col_inx];
                    string prod_id = tbl.Rows[0]["Pcs_No"].ToString();
                    Dictionary<string, List<string>> data_lst = Get_List_Pair_data(tbl);
                    List<string> before_data = data_lst["Before"];
                    int r_offset = 0;
                    foreach (var item in data_lst)
                    {
                        List<string> sel_data = item.Value;
                        string cycle_filter = "BF";
                        if (item.Key.Contains("After"))
                        {
                            cycle_filter = "L" + Extract_Num_from_String(item.Key);
                        }
                        List<string> log_time_lst = log_tbl.AsEnumerable().Where(x => x.Field<string>("Logfile").Contains(cycle_filter)).Select(x => x.Field<string>("Remark")).Distinct().ToList();
                        string log_time = "";
                        if (log_time_lst.Count > 0)
                        {
                            log_time = log_tbl.Rows[0]["Logfile"].ToString().Split('-').First() + " " + log_time_lst.Last();
                        }
                        bool data_OK = false;
                        if (Echeck_Cycle_addr.Keys.ToList().IndexOf(item.Key) != -1)
                        {
                            ExcelRangeBase sum_rgn_offset = tar_wrksht.Cells[Echeck_Cycle_addr[item.Key]].Offset(0, 2 + col_inx);
                            ExcelRangeBase time_rgn_offset = tar_wrksht.Cells[Echeck_Cycle_addr[item.Key]].Offset(0, 1 + col_inx);
                            time_rgn_offset.Value = log_time;

                            for (int r_inx = 0; r_inx < sel_data.Count; r_inx++)
                            {
                                string USL = myCode.checkDBNull(spec_tbl.Rows[r_inx]["USL"]);
                                string LSL = myCode.checkDBNull(spec_tbl.Rows[r_inx]["LSL"]);
                                string cur_val = sel_data[r_inx];
                                if (check_in_limit(USL, LSL, cur_val))
                                {
                                    data_OK = true;
                                }
                                else
                                {
                                    data_OK = false;
                                    break;
                                }
                            }
                            if (data_OK)
                            {
                                sum_rgn_offset.Value = "OK";
                            }
                            else
                            {
                                sum_rgn_offset.Value = "FAIL";
                            }
                        }
                        if (item.Key.Contains("After"))
                        {
                            int cycle_no = Convert.ToInt32(Extract_Num_from_String(item.Key));
                            if (cycle_data_addr_lst.Keys.ToList().IndexOf(cycle_no) != -1)
                            {
                                ExcelRangeBase cycle_rgn = tar_wrksht.Cells[cycle_data_addr_lst[cycle_no]].Offset(2, 0);
                                for (int r_inx = 0; r_inx < sel_data.Count; r_inx++)
                                {
                                    cycle_rgn.Offset(r_inx, 0).Value = before_data[r_inx];
                                    cycle_rgn.Offset(r_inx, 1).Value = sel_data[r_inx];
                                    cycle_rgn.Offset(r_inx, 2).FormulaR1C1 = "=RC[-1]/RC[-2]-1";
                                    cycle_rgn.Offset(r_inx, 2).Style.Numberformat.Format = "#0.00%";
                                    if (Math.Abs(Convert.ToDouble(cycle_rgn.Offset(r_inx, 2).Value)) > 0.1)
                                    {
                                        cycle_rgn.Offset(r_inx, 2).Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        cycle_rgn.Offset(r_inx, 2).Style.Fill.BackgroundColor.SetColor(Color.Red);
                                    }
                                }
                            }
                        }
                        if (result_sht_name != "")
                        {

                            if (item.Key.Contains("After"))
                            {
                                result_sht.Cells[proId_addr].Offset(r_offset + 1, 0).Value = prod_id;
                                result_sht.Cells[itemname_addr].Offset(r_offset + 1, 0).Value = itemname;
                                result_sht.Cells[itemcode_addr].Offset(r_offset + 1, 0).Value = ItemCode;
                                result_sht.Cells[lotno_addr].Offset(r_offset + 1, 0).Value = LotNo;
                                result_sht.Cells[process_addr].Offset(r_offset + 1, 0).Value = process_name;
                                result_sht.Cells[cycle_addr].Offset(r_offset + 1, 0).Value = cycle_filter.Replace("L", "");
                                result_sht.Cells[time_addr].Offset(r_offset + 1, 0).Value = log_time;
                                result_sht.Cells[operator_addr].Offset(r_offset + 1, 0).Value = UserID;
                                if (data_OK)
                                {
                                    result_sht.Cells[result_addr].Offset(r_offset + 1, 0).Value = "Pass";
                                }
                                else
                                {
                                    result_sht.Cells[result_addr].Offset(r_offset + 1, 0).Value = "NG";
                                }
                                r_offset++;
                            }
                        }

                    }
                    for (int i = 0; i < NET_name.Count; i++)
                    {
                        NET_rgn.Offset(i, 0).Value = NET_name[i];
                    }
                    string last_cycle_addr = cycle_data_addr_lst.LastOrDefault().Value;
                    TDMK_EPPLUS7_lib.MergeAreas_Info Cur_Cells_info = Excel_Lib.Get_Cells_Info(tar_wrksht, tar_wrksht.Cells[last_cycle_addr]);
                    int col_num = Cur_Cells_info.col_qty;
                    int total_col_num = tar_wrksht.Cells[last_cycle_addr].Start.Column + col_num - 1;
                    int r_off = NET_name.Count - 1;
                    int start_rgn_row = tar_wrksht.Cells[NET_start_rgn].Start.Row;
                    int start_rgn_col = tar_wrksht.Cells[NET_start_rgn].Start.Column;
                    ExcelRangeBase final_rgn = tar_wrksht.Cells[start_rgn_row, start_rgn_col, start_rgn_row + r_off, total_col_num];
                    Excel_Lib.Draw_border(final_rgn);
                }
            }
            else
            {
                MessageBox.Show("Format của ItemCode " + ItemCode + " không đúng", "Cảnh báo");
            }
        }
        public void Export_Thermal_HeatSoak_Bend_NPI(ExcelWorksheet tar_wrksht, SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, int qty)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            List<char> reject_char_lst = new List<char> { ' ', '_', '-', '&' };
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            DataTable _src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
            DataTable _spec_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process_name }));
            SortedDictionary<int, string> sel_index_lst = Get_bending_Net(sqlcon_OK2SHIP, ItemCode, process_name);
            List<int> index_lst = sel_index_lst.Keys.ToList();
            DataTable src_tbl = _src_tbl.AsEnumerable().Where(x => index_lst.Contains(Convert.ToInt32(x.Field<string>("Net_no")) - 1)).CopyToDataTable();
            DataTable spec_tbl = _spec_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_Report") != "No").CopyToDataTable();
            List<DataTable> src_tbl_lst = new List<DataTable>();
            Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
            List<string> Net_name = spec_tbl.AsEnumerable().Select(x => x.Field<string>("Net_Name")).ToList();
            string Echeck_start_rgn = Excel_Lib.Find_Start_Addr("A1", tar_wrksht, false, "Condition");
            ExcelRangeBase Echeck_cycle_rgn = tar_wrksht.Cells[Echeck_start_rgn];
            Dictionary<string, string> Echeck_Cycle_addr = Get_Echeck_address(Echeck_cycle_rgn);
            List<string> data_addr_lst = Get_Bending_data_addr("Test item", "Sample", tar_wrksht, "A10", true);
            string _NET_start_rgn = Excel_Lib.Find_Cell_Addr("Test item", "A10", tar_wrksht, false);
            string NET_start_rgn = Excel_Lib.get_offset_addr(_NET_start_rgn, tar_wrksht, 1, false);
            ExcelRangeBase NET_rgn = tar_wrksht.Cells[NET_start_rgn];
            int sel_qty = src_tbl_lst.Count;
            for (int col_inx = 0; col_inx < sel_qty; col_inx++)
            {
                DataTable tbl = src_tbl_lst[col_inx];
                Dictionary<string, List<string>> data_lst = Get_List_Pair_data(tbl);
                List<string> before_data = data_lst["Before"];
                List<string> last_data = data_lst.Values.ToList().Last();
                ExcelRangeBase cycle_rgn = tar_wrksht.Cells[data_addr_lst[col_inx]].Offset(2, 0);
                foreach (var item in data_lst)
                {
                    List<string> after_data = item.Value;
                    if (Echeck_Cycle_addr.Keys.ToList().IndexOf(item.Key) != -1)
                    {
                        ExcelRangeBase sum_rgn_offset = tar_wrksht.Cells[Echeck_Cycle_addr[item.Key]].Offset(0, 2 + col_inx);
                        bool data_OK = false;
                        for (int r_inx = 0; r_inx < after_data.Count; r_inx++)
                        {
                            string USL = myCode.checkDBNull(spec_tbl.Rows[r_inx]["USL"]);
                            string LSL = myCode.checkDBNull(spec_tbl.Rows[r_inx]["LSL"]);
                            string cur_val = after_data[r_inx];
                            if (check_in_limit(USL, LSL, cur_val))
                            {
                                data_OK = true;
                            }
                            else
                            {
                                data_OK = false;
                                break;
                            }
                        }
                        if (data_OK)
                        {
                            sum_rgn_offset.Value = "PASS";
                        }
                        else
                        {
                            sum_rgn_offset.Value = "FAIL";
                        }
                    }
                }
                for (int i = 0; i < before_data.Count; i++)
                {
                    cycle_rgn.Offset(i, 0).Value = before_data[i];
                    if (i < last_data.Count)
                    {
                        cycle_rgn.Offset(i, 1).Value = last_data[i];
                    }
                    cycle_rgn.Offset(i, 2).FormulaR1C1 = "=RC[-1]/RC[-2]-1";
                    cycle_rgn.Offset(i, 2).Style.Numberformat.Format = "#0.00%";
                    if (Math.Abs(Convert.ToDouble(cycle_rgn.Offset(i, 2).Value)) > 0.1)
                    {
                        cycle_rgn.Offset(i, 2).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cycle_rgn.Offset(i, 2).Style.Fill.BackgroundColor.SetColor(Color.Red);
                    }
                }
            }
            for (int i = 0; i < Net_name.Count; i++)
            {
                NET_rgn.Offset(i, 0).Value = Net_name[i];
            }
            string last_cycle_addr = data_addr_lst.LastOrDefault();
            TDMK_EPPLUS7_lib.MergeAreas_Info Cur_Cells_info = Excel_Lib.Get_Cells_Info(tar_wrksht, tar_wrksht.Cells[last_cycle_addr]);
            int col_num = Cur_Cells_info.col_qty;
            int total_col_num = tar_wrksht.Cells[last_cycle_addr].Start.Column + col_num - 1;
            int r_off = Net_name.Count - 1;
            int start_rgn_row = tar_wrksht.Cells[NET_start_rgn].Start.Row;
            int start_rgn_col = tar_wrksht.Cells[NET_start_rgn].Start.Column;
            ExcelRangeBase final_rgn = tar_wrksht.Cells[start_rgn_row, start_rgn_col, start_rgn_row + r_off, total_col_num];
            Excel_Lib.Draw_border(final_rgn);
        }
        public bool check_limit_vary(List<string> before_data, List<string> last_data, DataTable spec_tbl)
        {
            bool result = false;
            for (int inx = 0; inx < before_data.Count; inx++)
            {
                string USL = myCode.checkDBNull(spec_tbl.Rows[inx]["USL"]);
                string LSL = myCode.checkDBNull(spec_tbl.Rows[inx]["LSL"]);
                string cur_before_val = before_data[inx];
                string cur_after_val = last_data[inx];
                if (myCode.IsNumeric(cur_before_val) && myCode.IsNumeric(cur_after_val))
                {
                    if (check_in_limit(USL, LSL, cur_before_val) && check_in_limit(USL, LSL, cur_after_val))
                    {
                        double bef_val = Convert.ToDouble(cur_before_val);
                        double aft_val = Convert.ToDouble(cur_after_val);
                        double vary_R = Math.Abs(aft_val - bef_val) / bef_val;
                        if (vary_R < 0.1)
                        {
                            result = true;
                        }
                        else
                        {
                            result = false;
                            break;
                        }
                    }
                    else
                    {
                        result = false;
                        break;
                    }
                }
                else
                {
                    result = false;
                    break;
                }
            }
            return result;
        }
        public bool check_bending_data(SqlConnection sqlcon, string ItemCode, string LotNo, string process_name)
        {
            bool data_OK = false;
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            DataTable _src_tbl = TDMK_Code.Datatable_Filter(sqlcon, process_name, filter_str);
            DataTable _spec_tbl = TDMK_Code.Datatable_Filter(sqlcon, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process_name }));
            SortedDictionary<int, string> sel_index_lst = Get_bending_Net(sqlcon, ItemCode, process_name);
            List<int> index_lst = sel_index_lst.Keys.ToList();
            if (_src_tbl.Rows.Count > 0)
            {
                DataTable src_tbl = _src_tbl.AsEnumerable().Where(x => index_lst.Contains(Convert.ToInt32(x.Field<string>("Net_no")) - 1)).CopyToDataTable();
                DataTable spec_tbl = _spec_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_Report") != "No").CopyToDataTable();
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                List<char> reject_char_lst = new List<char> { ' ', '_', '-', '&' };
                List<string> Net_name = spec_tbl.AsEnumerable().Select(x => x.Field<string>("Net_Name")).ToList();
                int sel_qty = src_tbl_lst.Count;
                for (int col_inx = 0; col_inx < sel_qty; col_inx++)
                {
                    DataTable tbl = src_tbl_lst[col_inx];
                    Dictionary<string, List<string>> data_lst = Get_List_Pair_data(tbl);
                    List<string> before_data = data_lst["Before"];
                    List<string> last_data = data_lst.Values.ToList().Last();
                    if (!check_limit_vary(before_data, last_data, spec_tbl))
                    {
                        break;
                    }
                    else
                    {
                        data_OK = true;
                    }
                }
            }

            return data_OK;
        }
        public string find_config_path(string src_string, string f_name)
        {
            DirectoryInfo di = new DirectoryInfo(src_string);
            var folder_lst = di.GetDirectories().Select(x => x.FullName).ToList();
            var temp2 = folder_lst.Where(x => new DirectoryInfo(x).Name == f_name).ToList();
            if (temp2.Count != 0)
            {
                return temp2.FirstOrDefault();
            }
            else
            {
                if (di.Parent != null)
                {
                    return find_config_path(di.Parent.FullName, f_name);
                }
                else
                {
                    return string.Empty;
                }
            }
        }
    }
}
