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
using myExcel = Microsoft.Office.Interop.Excel;
using IniLibs;
using System.Windows.Forms;

namespace Bending_Export
{
    public class Bending_Export_Lib
    {
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SEI_Lib myCode = new SEI_Lib();
        IniFile TDMK_init = new IniFile("Config.ini");
        public string Data_Location { get; set; }
        public string format_folder { get; set; }
        public string log_folder { get; set; }
        public string Report_location { get; set; }
        public SqlConnection initial_data(string DB_name, bool sa_en)
        {
            SqlConnection _sqlcon_OK2SHIP;

            if (!File.Exists("Config.ini"))
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
        public string Get_start_range(string search_key, string start_addr, myExcel.Worksheet tar_wrksht, bool search_by_col)
        {
            string result = "";
            List<char> reject_char_lst = new List<char> { ' ', '_', '-', '\r', '\n' };
            myExcel.Range cycle_rgn = tar_wrksht.Range[start_addr];
            for (int i = 0; i < 100; i++)
            {
                myExcel.Range sel_rgn = cycle_rgn.Offset[i, 0];
                if (search_by_col)
                {
                    sel_rgn = cycle_rgn.Offset[0, i];
                }
                string sel_rgn_val = myCode.checkDBNull(sel_rgn.Value);

                if (remove_special_char(sel_rgn_val, reject_char_lst).ToUpper() == remove_special_char(search_key, reject_char_lst).ToUpper())
                {
                    result = cycle_rgn.Offset[i + sel_rgn.MergeArea.Rows.Count, 0].AddressLocal;
                    if (search_by_col)
                    {
                        result = cycle_rgn.Offset[sel_rgn.MergeArea.Rows.Count, i].AddressLocal;
                    }
                    break;
                }
            }
            return result;
        }
        public string Find_Cell_Addr(string search_key, string start_addr, myExcel.Worksheet tar_wrksht)
        {
            string result = "";
            myExcel.Range cycle_rgn = tar_wrksht.Range[start_addr];
            for (int i = 0; i < 100; i++)
            {
                myExcel.Range sel_rgn = cycle_rgn.Offset[i, 0];
                string sel_rgn_val = myCode.checkDBNull(sel_rgn.Value);
                if (sel_rgn_val.Replace(" ", "").ToUpper() == search_key.Replace(" ", "").ToUpper())
                {
                    result = sel_rgn.AddressLocal;
                    break;
                }
            }
            return result;
        }
        public string Get_start_range(string search_key, string start_addr, myExcel.Worksheet tar_wrksht)
        {
            string result = "";
            myExcel.Range cycle_rgn = tar_wrksht.Range[start_addr];
            for (int i = 0; i < 100; i++)
            {
                myExcel.Range sel_rgn = cycle_rgn.Offset[i, 0];
                string sel_rgn_val = myCode.checkDBNull(sel_rgn.Value);
                if (sel_rgn_val.Replace(" ", "").ToUpper() == search_key.Replace(" ", "").ToUpper())
                {
                    result = cycle_rgn.Offset[i + sel_rgn.MergeArea.Rows.Count, 0].AddressLocal;
                    break;
                }
            }
            return result;
        }
        public Dictionary<string, string> Get_Echeck_address(myExcel.Range cycle_rgn)
        {
            Dictionary<string, string> bending_Cycle_addr = new Dictionary<string, string>();
            int row_inx = 0;
            while (myCode.checkDBNull(cycle_rgn.Offset[row_inx, 0].Value) != "")
            {
                string rgn_val = myCode.checkDBNull(cycle_rgn.Offset[row_inx, 0].Value);
                if (rgn_val.ToUpper().Contains("E-CHECK"))
                {
                    string cycle = Extract_Num_from_String(cycle_rgn.Offset[row_inx, 0].Value);// Get_Number_String(cycle_rgn.Offset[row_inx, 0].Value, ' ');
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
                        bending_Cycle_addr.Add(cycle_name, cycle_rgn.Offset[row_inx, 0].AddressLocal);
                    }
                }
                row_inx++;
            }
            return bending_Cycle_addr;
        }
        public string Extract_Num_from_String(string str_in)
        {
            string result = "";
            result = new string(str_in.Where(x => char.IsDigit(x)).ToArray());
            return result;
        }
        public List<string> Get_Bending_data_addr(string search_item, string search_key, myExcel.Worksheet tar_wrksht, string start_range_addr, bool col_direction)
        {
            List<string> result = new List<string>();
            string tar_addr = Find_Cell_Addr(search_item, start_range_addr, tar_wrksht);
            myExcel.Range sample_rgn = tar_wrksht.Range[tar_addr];
            myExcel.Range sel_rgn = sample_rgn;
            int count_null = 0;
            while (true)
            {
                string rgn_addr = sel_rgn.AddressLocal;
                myExcel.Range temp_rgn = sel_rgn.Offset[0, 1];
                if (col_direction)
                {
                    temp_rgn = sel_rgn.Offset[1, 0];
                }
                string temp_addr = temp_rgn.AddressLocal;
                string temp_rgn_val = myCode.checkDBNull(temp_rgn.Value);
                if (temp_rgn_val.ToUpper().Contains(search_key.ToUpper()))
                {
                    result.Add(temp_rgn.AddressLocal);
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
        public void Export_Thermal_HeatSoak_Bend_All(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, myExcel.Worksheet tar_wrksht)
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



            string Echeck_start_rgn = Get_start_range("Condition", "A1", tar_wrksht);
            myExcel.Range Echeck_cycle_rgn = tar_wrksht.Range[Echeck_start_rgn];
            Dictionary<string, string> Echeck_Cycle_addr = Get_Echeck_address(Echeck_cycle_rgn);
            List<string> data_addr_lst = Get_Bending_data_addr("Test item", "Sample", tar_wrksht, "A10", false);
            myExcel.Range NET_rgn = tar_wrksht.Range[Get_start_range("Test item", "A10", tar_wrksht)];
            //int sel_qty = Math.Min(qty, src_tbl_lst.Count);
            int sel_qty = src_tbl_lst.Count;
            for (int col_inx = 0; col_inx < sel_qty; col_inx++)
            {
                DataTable tbl = src_tbl_lst[col_inx];
                Dictionary<string, List<string>> data_lst = Get_List_Pair_data(tbl);
                List<string> before_data = data_lst["Before"];
                List<string> last_data = data_lst.Values.ToList().Last();
                myExcel.Range cycle_rgn = tar_wrksht.Range[data_addr_lst[col_inx]].Offset[2, 0];
                foreach (var item in data_lst)
                {
                    List<string> after_data = item.Value;
                    if (Echeck_Cycle_addr.Keys.ToList().IndexOf(item.Key) != -1)
                    {
                        myExcel.Range sum_rgn_offset = tar_wrksht.Range[Echeck_Cycle_addr[item.Key]].Offset[0, 2 + col_inx];
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
                    cycle_rgn.Offset[i, 0].Value = before_data[i];
                    if (i < last_data.Count)
                    {
                        cycle_rgn.Offset[i, 1].Value = last_data[i];
                    }
                    cycle_rgn.Offset[i, 2].FormulaR1C1 = "=RC[-1]/RC[-2]-1";
                    if (Math.Abs(Convert.ToDouble(cycle_rgn.Offset[i, 2].Value)) > 0.1)
                    {
                        cycle_rgn.Offset[i, 2].Interior.Color = 255;
                    }
                    cycle_rgn.Offset[i, 2].NumberFormat = "0.00 %";
                }
            }
            for (int i = 0; i < Net_name.Count; i++)
            {
                NET_rgn.Offset[i, 0].Value = Net_name[i];
            }
            string last_cycle_addr = data_addr_lst.LastOrDefault();
            int col_num = tar_wrksht.Range[last_cycle_addr].MergeArea.Columns.Count;
            int total_col_num = tar_wrksht.Range[last_cycle_addr].Column + col_num - 1;
            int r_off = Net_name.Count - 1;
            myExcel.Range final_rgn = tar_wrksht.Range[NET_rgn, NET_rgn.Offset[r_off, total_col_num - 1]];
            List<myExcel.XlBordersIndex> st_lst = new List<myExcel.XlBordersIndex>() { myExcel.XlBordersIndex.xlEdgeLeft, myExcel.XlBordersIndex.xlEdgeTop, myExcel.XlBordersIndex.xlEdgeBottom, myExcel.XlBordersIndex.xlEdgeRight, myExcel.XlBordersIndex.xlInsideHorizontal, myExcel.XlBordersIndex.xlInsideVertical };//myExcel.XlBordersIndex.xlDiagonalDown, myExcel.XlBordersIndex.xlDiagonalUp,
            foreach (var st in st_lst)
            {
                final_rgn.Borders[st].LineStyle = myExcel.XlLineStyle.xlContinuous;
            }


        }
        public void Export_Thermal_HeatSoak_Bend_All_old(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name, int qty, string format_type)
        {
            //string format_loc = Path.Combine(format_folder, format_type);
            string format_loc = format_folder;
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
                myExcel.Workbook report_wrk = null; //create_export_wrk(format_file, format_name);
                myExcel.Worksheet tar_wrksht = null;// report_wrk.Sheets[1];
                if (format_type == "NPI")
                {
                    report_wrk = TDMK_Code.open_excel_file(format_file, "", "");
                    foreach (myExcel.Worksheet sht in report_wrk.Worksheets)
                    {
                       
                        if (remove_special_char(sht.Name.ToUpper(), reject_char_lst) == remove_special_char(format_name.ToUpper(), reject_char_lst))
                        {
                            tar_wrksht = sht;
                            break;
                        }
                    }
                }
                else
                {
                    report_wrk = TDMK_Code.open_excel_file(format_file, "", "");
                    foreach (myExcel.Worksheet sht in report_wrk.Worksheets)
                    {
                        if (remove_special_char(sht.Name, reject_char_lst) == remove_special_char("Bending < 10%", reject_char_lst))
                        {
                            tar_wrksht = sht;
                            break;
                        }
                    }
                }
                if (tar_wrksht != null)
                {
                    string[] temp = Path.GetFileNameWithoutExtension(report_wrk.Name).Split('-');
                    string itemname = "";
                    if (temp.Length > 1)
                    {
                        itemname = temp[1];
                    }
                    string report_name = process_name + " " + itemname + "-" + ItemCode + "-" + LotNo + " " + DateTime.Now.ToString("yyyyMMdd") + Path.GetExtension(report_wrk.Name);
                    if (format_type == "NPI")
                    {
                        report_name = Path.GetFileNameWithoutExtension(report_wrk.Name) + "-" + LotNo + Path.GetExtension(report_wrk.Name);
                    }
                    string daily_folder = Path.Combine(Report_location, "OK2SHIP_report");
                    if (!Directory.Exists(daily_folder))
                    {
                        Directory.CreateDirectory(daily_folder);
                    }
                    report_wrk.SaveAs(Path.Combine(daily_folder, report_name));
                    tar_wrksht.Activate();
                    string Echeck_start_rgn = Get_start_range("Condition", "A1", tar_wrksht);
                    myExcel.Range Echeck_cycle_rgn = tar_wrksht.Range[Echeck_start_rgn];
                    Dictionary<string, string> Echeck_Cycle_addr = Get_Echeck_address(Echeck_cycle_rgn);
                    List<string> data_addr_lst = Get_Bending_data_addr("Test item", "Sample", tar_wrksht, "A10", false);
                    myExcel.Range NET_rgn = tar_wrksht.Range[Get_start_range("Test item", "A10", tar_wrksht)];
                    //int sel_qty = Math.Min(qty, src_tbl_lst.Count);
                    int sel_qty = src_tbl_lst.Count;
                    for (int col_inx = 0; col_inx < sel_qty; col_inx++)
                    {
                        DataTable tbl = src_tbl_lst[col_inx];
                        Dictionary<string, List<string>> data_lst = Get_List_Pair_data(tbl);
                        List<string> before_data = data_lst["Before"];
                        List<string> last_data = data_lst.Values.ToList().Last();
                        myExcel.Range cycle_rgn = tar_wrksht.Range[data_addr_lst[col_inx]].Offset[2, 0];
                        foreach (var item in data_lst)
                        {
                            List<string> after_data = item.Value;
                            if (Echeck_Cycle_addr.Keys.ToList().IndexOf(item.Key) != -1)
                            {
                                myExcel.Range sum_rgn_offset = tar_wrksht.Range[Echeck_Cycle_addr[item.Key]].Offset[0, 2 + col_inx];
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
                            cycle_rgn.Offset[i, 0].Value = before_data[i];
                            if (i < last_data.Count)
                            {
                                cycle_rgn.Offset[i, 1].Value = last_data[i];
                            }
                            cycle_rgn.Offset[i, 2].FormulaR1C1 = "=RC[-1]/RC[-2]-1";
                            if (Math.Abs(Convert.ToDouble(cycle_rgn.Offset[i, 2].Value)) > 0.1)
                            {
                                cycle_rgn.Offset[i, 2].Interior.Color = 255;
                            }
                            cycle_rgn.Offset[i, 2].NumberFormat = "0.00 %";
                        }
                    }
                    for (int i = 0; i < Net_name.Count; i++)
                    {
                        NET_rgn.Offset[i, 0].Value = Net_name[i];
                    }
                    string last_cycle_addr = data_addr_lst.LastOrDefault();
                    int col_num = tar_wrksht.Range[last_cycle_addr].MergeArea.Columns.Count;
                    int total_col_num = tar_wrksht.Range[last_cycle_addr].Column + col_num - 1;
                    int r_off = Net_name.Count - 1;
                    myExcel.Range final_rgn = tar_wrksht.Range[NET_rgn, NET_rgn.Offset[r_off, total_col_num - 1]];
                    List<myExcel.XlBordersIndex> st_lst = new List<myExcel.XlBordersIndex>() { myExcel.XlBordersIndex.xlEdgeLeft, myExcel.XlBordersIndex.xlEdgeTop, myExcel.XlBordersIndex.xlEdgeBottom, myExcel.XlBordersIndex.xlEdgeRight, myExcel.XlBordersIndex.xlInsideHorizontal, myExcel.XlBordersIndex.xlInsideVertical };//myExcel.XlBordersIndex.xlDiagonalDown, myExcel.XlBordersIndex.xlDiagonalUp,
                    foreach (var st in st_lst)
                    {
                        final_rgn.Borders[st].LineStyle = myExcel.XlLineStyle.xlContinuous;
                    }
                    report_wrk.Save();
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
    }
}
