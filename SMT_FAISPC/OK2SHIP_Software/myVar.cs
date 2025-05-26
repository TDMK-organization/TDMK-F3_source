using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.IO;
using System.Data;
using TDMK_SQL;
using myExcel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using OK2SHIP_Measurements;

using OK2SHIP_Lib;
using IniLibs;

namespace OK2SHIP_Software
{
    public class myVar
    {
        public static SqlConnection sqlcon_SMT;
        public static SqlConnection sqlcon_Declare;
        public static SqlConnection sqlcon_Materials;
        public static SqlConnection sqlcon_Recycle;
        public static SqlConnection sqlcon_IPQC;
        public static SqlConnection sqlcon_OK2SHIP;
        public static SqlConnection sqlcon_OK2SHIP_Period2;
        public static string server_name;
        public static string server_acc;
        public static string server_pass;
        public static string DB_name;
        public static string app_path;
        public static string data_loc;
        public static string format_loc;
        public static string log_loc;
        public static string report_loc;
        public static  TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public TDMK_OK2SHIP myCode2 = new TDMK_OK2SHIP();
        public static FrmExport frm_export;// = new FrmMain();
        public static MainScreen frm_mainscreen;// = new MainScreen();
        public static bool confirm_mode = false;
        public static string edit_val = "Yes";
        public static string UserID = "";
        public static string USerDepart = "";
        public IniFile TDMK_init = new IniFile("Config.ini");
        public void initial_data()
        {
            //string config_file = Path.Combine(app_path, "Config", "config.txt");
            //string[] my_config = read_config_arr(config_file);
            //foreach (string c in my_config)
            //{
            //    if (c.Contains("Server"))
            //    {
            //        server_name = c.Split(':')[1].Trim();
            //    }
            //    if (c.Contains("Account"))
            //    {
            //        server_acc = c.Split(':')[1].Trim();
            //    }
            //    if (c.Contains("Password"))
            //    {
            //        server_pass = c.Split(':')[1].Trim();
            //    }
            //    if (c.Contains("DB_name"))
            //    {
            //        DB_name = c.Split(':')[1].Trim();
            //    }
            //    if (c.Contains("Report_Location"))
            //    {
            //        data_loc = c.Split('#')[1].Trim();
            //    }
            //}


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
            server_name = TDMK_init.Read("Server", "SMT_Config");
            server_acc = TDMK_init.Read("Account", "SMT_Config");
            server_pass = TDMK_init.Read("Password", "SMT_Config");
            report_loc = TDMK_init.Read("Report_Location", "SMT_Config");
            data_loc = TDMK_init.Read("Data_Location", "SMT_Config");
            log_loc = TDMK_init.Read("Log_folder", "SMT_Config");
            format_loc = TDMK_init.Read("Format_Folder", "SMT_Config");
            //string connstr_Declare = TDMK_Code.data_connection(myVar.server_name, "Declaration", myVar.server_acc, myVar.server_pass).ConnectionString;
            //string connstr_Materials = TDMK_Code.data_connection(myVar.server_name, "Materials", myVar.server_acc, myVar.server_pass).ConnectionString;
            //string connstr_Recycle = TDMK_Code.data_connection(myVar.server_name, "Recycled_PGC_Copper", myVar.server_acc, myVar.server_pass).ConnectionString;
            //string connstr_FAI = TDMK_Code.data_connection(myVar.server_name, "SEI_FAI", myVar.server_acc, myVar.server_pass).ConnectionString;
            //string connstr_IPQC = TDMK_Code.data_connection(myVar.server_name, "IPQC_Data", myVar.server_acc, myVar.server_pass).ConnectionString;
            //string connstr_OK2SHIP = TDMK_Code.data_connection(myVar.server_name, "OK2SHIP_Items", myVar.server_acc, myVar.server_pass).ConnectionString;
            //string connstr_OK2SHIP_Period2 = TDMK_Code.data_connection(myVar.server_name, "OK2SHIP_Period2", myVar.server_acc, myVar.server_pass).ConnectionString;
            //sqlcon_Declare = new SqlConnection(connstr_Declare);
            //sqlcon_Materials = new SqlConnection(connstr_Materials);
            //sqlcon_Recycle = new SqlConnection(connstr_Recycle);
            //sqlcon_SMT = new SqlConnection(connstr_FAI);
            //sqlcon_IPQC = new SqlConnection(connstr_IPQC);
            //sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);
            //sqlcon_OK2SHIP_Period2 = new SqlConnection(connstr_OK2SHIP_Period2);
        }
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
            server_name = TDMK_init.Read("Server", "SMT_Config");
            server_acc = TDMK_init.Read("Account", "SMT_Config");
            server_pass = TDMK_init.Read("Password", "SMT_Config");
            data_loc = TDMK_init.Read("Data_Location", "SMT_Config");
            format_loc = TDMK_init.Read("Format_Folder", "SMT_Config");
            log_loc = TDMK_init.Read("Log_folder", "SMT_Config");
            report_loc = TDMK_init.Read("Report_Location", "SMT_Config");
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
        public string[] read_config_arr(string src_config_file)
        {
            string[] _result = new string[10];
            StreamReader reader;
            int inx = 0;
            reader = new StreamReader(src_config_file);
            while ((!reader.EndOfStream) && (inx < 10))
            {
                string temp = reader.ReadLine();
                if (temp != "")
                {
                    _result[inx] = temp;
                    inx++;
                }
            }
            Array.Resize<string>(ref _result, inx);
            reader.Close();
            reader.Dispose();
            return _result;
        }
        public string[] read_config_arr2(string src_config_file)
        {
            string[] _result = new string[50];
            StreamReader reader;
            int inx = 0;
            reader = new StreamReader(src_config_file);
            while ((!reader.EndOfStream) && (inx < 50))
            {
                string temp = reader.ReadLine();
                if (temp != "")
                {
                    _result[inx] = temp.Replace(" ","");
                    inx++;
                }
            }
            Array.Resize<string>(ref _result, inx);
            reader.Close();
            reader.Dispose();
            return _result;
        }
        public string[] GetAllTables(SqlConnection connection)
        {
            List<string> result = new List<string>();
            if (connection.State != ConnectionState.Open)
            {
                connection.Open();
            }
            SqlCommand cmd = new SqlCommand("SELECT name FROM sys.Tables", connection);
            System.Data.SqlClient.SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
                result.Add(reader["name"].ToString());
            connection.Close();
            return result.ToArray();
        }
        public void Export_IPQC_Data(string f_format, string tar_ItemCode, string tar_LotNo)
        {
            myExcel.Workbook export_wrkbook = TDMK_Code.open_excel_file(f_format, "", "");
            myExcel.Worksheet tar_wrksheet = new myExcel.Worksheet();
            foreach (myExcel.Worksheet t in export_wrkbook.Sheets)
            {
                if (t.Name.Contains("IPQC"))
                {
                    tar_wrksheet = t;
                    break;
                }
            }
            tar_wrksheet.Select();
            string[] IPQC_items = GetAllTables(myVar.sqlcon_IPQC);
            foreach (string t in IPQC_items)
            {
                if (t != "SpecList")
                {
                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Remark" }, new string[] { tar_ItemCode, tar_LotNo, "%" + t + "%" });
                    AutoCompleteStringCollection addr_list = TDMK_Code.Load_Item_Names_Filter_notOrder(myVar.sqlcon_IPQC, "SpecList", "Address", filter_str);
                    string filter_str_data = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { tar_ItemCode, tar_LotNo });
                    DataTable data_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_IPQC, t, filter_str_data);
                    if (data_tbl.Rows.Count > 0)
                    {
                        for (int i = 0; i < data_tbl.Rows.Count; i++)
                        {
                            for (int j = 3; j < data_tbl.Columns.Count; j++)
                            {
                                string rgn_addr = addr_list[j - 3];
                                myExcel.Range cur_rgn = tar_wrksheet.Range[rgn_addr].Offset[i, 0];
                                cur_rgn.Value = data_tbl.Rows[i][j].ToString();
                            }
                        }
                    }
                }
            }
        }
       
        public void Export_IPQC_Data2(myExcel.Workbook export_wrkbook, string tar_ItemCode, string tar_LotNo)
        {            
            myExcel.Worksheet tar_wrksheet = new myExcel.Worksheet();
            foreach (myExcel.Worksheet t1 in export_wrkbook.Sheets)
            {
                if (t1.Name.Contains("IPQC"))
                {
                    tar_wrksheet = t1;
                    break;
                }
            }
            tar_wrksheet.Select();
            string[] IPQC_items = GetAllTables(myVar.sqlcon_IPQC);
            AutoCompleteStringCollection ignored_lst = new AutoCompleteStringCollection() { "All_Items", "SpecList", "Spec_List", "Sequence" };
            foreach (string t in IPQC_items)
            {
                if (!TDMK_Code.check_exist_list(t, ignored_lst))
                {
                    string filter_str;
                    switch (t)
                    {
                        case "Roughness":
                            filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Col_Name" }, new string[] { tar_ItemCode, "%" + t + "%" });
                            break;
                        case "AU_NI":
                            filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Col_Name" }, new string[] { tar_ItemCode, "%THICKNESS%" });
                            break;
                        default:
                            filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { tar_ItemCode, "%" + t + "%" });
                            break;
                    }

                    DataTable spec_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_IPQC, "SpecList", filter_str);
                    string[] addr_list = spec_tbl.AsEnumerable().Select(r => r.Field<string>("Address")).ToArray();
                    string[] colname_list = spec_tbl.AsEnumerable().Select(r => r.Field<string>("Col_Name")).ToArray();
                    string filter_str_data = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { tar_ItemCode, tar_LotNo });
                    DataTable src_data_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_IPQC, t, filter_str_data);
                    DataTable data_tbl = src_data_tbl.AsDataView().ToTable(false, colname_list);
                    char[] trim_char = new char[] { '\r', '\n', ' ' };
                    int exp_num = 32;
                    if(data_tbl.Rows.Count <= exp_num)
                    {
                        exp_num = data_tbl.Rows.Count;
                    }
                    if (data_tbl.Rows.Count > 0)
                    {
                        for (int i = 0; i < exp_num; i++)
                        {
                            for (int j = 0; j < data_tbl.Columns.Count; j++)
                            {
                                string rgn_addr = addr_list[j];
                                myExcel.Range cur_rgn = tar_wrksheet.Range[rgn_addr].Offset[i, 0];
                                cur_rgn.Value = data_tbl.Rows[i][j].ToString().Trim(trim_char);
                            }
                        }
                    }
                }
            }
        }
        public void Export_To_FAI(string f_format,string _ItemCode, string _LotNo, myExcel.Workbook src_data_wrkbook)
        {
            myExcel.Workbook format_wrk = TDMK_Code.open_excel_file(f_format,"", "");
            //format_wrk.SaveAs(Path.Combine(app_path, "Export", _ItemCode + _LotNo + "_temp.xlsx"));
            myExcel.Worksheet sel_wrksheet;
            myExcel.Worksheet data_sheet = src_data_wrkbook.Sheets[1];
            myExcel.Range data_rgn = data_sheet.Range["A1"];
            for (int i = 0; i < format_wrk.Sheets.Count; i++)
            {
                sel_wrksheet = format_wrk.Sheets[i + 1];
                sel_wrksheet.Activate();
                string wrk_name = sel_wrksheet.Name;
                if (wrk_name.Contains("FAI") || wrk_name.Contains("CPK"))
                {
                    myExcel.Range FAI_rgn = sel_wrksheet.Range["C17"];
                    myExcel.Range dest_FAI_rgn = sel_wrksheet.Range["C259"];
                    myExcel.Range NormDim_val_rgn = sel_wrksheet.Range["C18"];
                    int k = 0;
                    while (checkDBNull(FAI_rgn.Offset[0, k].Value) != "")
                    {
                        int j = 0;
                        string curr_rgn_val = FAI_rgn.Offset[0, k].Value;
                        curr_rgn_val.Replace(" ", "");
                        string[] temp = curr_rgn_val.Split('/');
                        string FAI_rgn_val = checkDBNull(temp[0]);
                        if (wrk_name.Contains("CPK"))
                        {
                            string[] temp2 = temp[1].Split('-');
                            if (TDMK_Code.IsNumeric(temp2[temp2.Length - 1]))
                            {
                                FAI_rgn_val = FAI_rgn_val + "-" + temp2[temp2.Length - 1];
                            }
                        }
                        FAI_rgn_val = FAI_rgn_val + "_" + NormDim_val_rgn.Offset[0, k].Value;
                        while (checkDBNull(data_sheet.Range["A1"].Offset[0, j].Value) != "")
                        {
                            string data_rgn_val = checkDBNull(data_sheet.Range["A1"].Offset[0, j].Value);
                            if (FAI_rgn_val == data_rgn_val)
                            {
                                int row_off = 0;
                                while (checkDBNull(data_rgn.Offset[1 + row_off, j].Value) != "")
                                {
                                    row_off++;
                                }
                                myExcel.Range from_rgn = data_sheet.Range[data_rgn.Offset[1, j], data_rgn.Offset[row_off, j]];
                                myExcel.Range to_rgn = dest_FAI_rgn.Offset[0, k];
                                from_rgn.Copy(to_rgn);
                                break;
                            }
                            j++;
                        }
                        k++;
                    }
                }
            }

        }
        public void Export_To_FAI2(myExcel.Workbook src_format_wrk, string _ItemCode, string _LotNo, myExcel.Workbook src_data_wrkbook)
        {
            DataTable FAI_Data_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { _ItemCode, _LotNo }));
            DataTable FAI_Spec_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "FAI_Spec", "ItemCode = '" + _ItemCode + "'");
            List<string> FAI_keys_lst = new List<string> { "FAI", "SPC", "CPK", "parentheses" };
            foreach (myExcel.Worksheet sht in src_data_wrkbook.Worksheets)
            {
                try
                {
                    if (FAI_keys_lst.Any(x=>sht.Name.Contains(x)))   //if ((sht.Name.Contains("FAI")) || (sht.Name.Contains("SPC")))
                    {
                        string sel_sht_name = sht.Name;
                        myExcel.Worksheet format_wrksheet = src_format_wrk.Sheets[sel_sht_name];
                        format_wrksheet.Activate();
                        string dim_no_addr = Find_Cell_Addr("Dim. No.", "B10", format_wrksheet, false);
                        string instrument_addr = Find_Cell_Addr("instrument", "B10", format_wrksheet, false);
                        string FAI_data_addr = Find_Offset(format_wrksheet.Range[dim_no_addr].Offset[0, 1].AddressLocal, format_wrksheet, true, "");
                        int off_set = format_wrksheet.Range[FAI_data_addr].Column - format_wrksheet.Range[instrument_addr].Column;
                        Dictionary<string, int> fai_loc = Find_FAI_addr_qty(dim_no_addr, format_wrksheet, false);
                        if (fai_loc.Count > 0)
                        {
                            string _data_addr = fai_loc.Keys.ToList()[0];
                            myExcel.Range format_rgn = format_wrksheet.Range[_data_addr].Offset[0, off_set];
                            resume_label: if (format_wrksheet.ProtectContents)
                            {
                                format_wrksheet.Unprotect("Histogram_123");
                            }
                            try
                            {
                                int col_num = sht.UsedRange.Columns.Count;
                                int row_num = Math.Min( sht.UsedRange.Rows.Count, fai_loc[_data_addr]+1);
                                myExcel.Range data_rgn = sht.Range[sht.Cells[2, 1], sht.Cells[row_num, col_num]];
                                data_rgn.Copy();
                                format_rgn.PasteSpecial(myExcel.XlPasteType.xlPasteValues, myExcel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
                            }
                            catch
                            {

                                goto resume_label;
                            }
                        }
                    }
                }
                catch
                {
                    continue;
                }
                
            }
            src_data_wrkbook.Close(false);            
        }
        public string checkDBNull(object src_str)
        {
            string _result;
            if (src_str != null)
            {
                _result = src_str.ToString();
            }
            else
            {
                _result = "";
            }
            return _result;
        }
        public DataTable Load_FAI_DataTable(SqlConnection tar_sqlcon, string tar_tbl, string FAI_No_name, string FAI_Data_Col_name, string tar_ItemCode, string tar_LotNo)
        {
            DataTable tbl_data = new DataTable();
            string[] flt_items = new string[3];
            string[] flt_item_vals = new string[3];
            AutoCompleteStringCollection FAI_lst;
            flt_items[0] = "ItemCode";
            flt_items[1] = "LotNo";
            flt_items[2] = "FAI_No";
            flt_item_vals[0] = tar_ItemCode;
            flt_item_vals[1] = tar_LotNo;
            flt_item_vals[2] = "";
            FAI_lst = TDMK_Code.Load_Item_Names_Filter(tar_sqlcon, tar_tbl, FAI_No_name, TDMK_Code.filter_str(flt_items, flt_item_vals));
            foreach (string c in FAI_lst)
            {
                AutoCompleteStringCollection FAI_vals;
                flt_item_vals[2] = c;
                FAI_vals = TDMK_Code.Load_Item_Filter_str(tar_sqlcon, tar_tbl, FAI_Data_Col_name, TDMK_Code.filter_str(flt_items, flt_item_vals));
                if (!check_columns_existed(tbl_data, c))
                {
                    tbl_data.Columns.Add(c);
                }
                int tbl_row_count = tbl_data.Rows.Count;
                if (tbl_row_count < FAI_vals.Count)
                {
                    for (int i = 0; i < FAI_vals.Count - tbl_row_count; i++)
                    {
                        tbl_data.Rows.Add();
                    }
                }
                int inx = 0;
                foreach (string t in FAI_vals)
                {
                    tbl_data.Rows[inx][c] = t;
                    inx++;
                }
            }
            return tbl_data;
        }
        public bool check_columns_existed(DataTable src_tbl, string find_col_name)
        {
            bool _result = false;
            foreach (DataColumn c in src_tbl.Columns)
            {
                if (c.ColumnName.ToString() == find_col_name)
                {
                    _result = true;
                    break;
                }
            }
            return _result;
        }
        public void Export_FAI_Data(string f_format, string tar_ItemCode, string tar_LotNo)
        {
            DataTable FAI_tbl = Load_FAI_DataTable(myVar.sqlcon_SMT, "FAI_Auto", "FAI_No", "FAI_Data", tar_ItemCode, tar_LotNo);
            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            xlsApp.Visible = true;
            myExcel.Workbook export_wrkbook = xlsApp.Workbooks.Add();
            myExcel.Worksheet exp_wrksheet = export_wrkbook.Sheets[1];
            for (int i = 0; i < FAI_tbl.Columns.Count; i++)
            {
                exp_wrksheet.Range["A1"].Offset[0, i].Value = FAI_tbl.Columns[i].ColumnName;
            }
            int row_inx = 0;
            foreach (DataRow dr in FAI_tbl.Rows)
            {
                int col_inx = 0;
                foreach (DataColumn dcol in FAI_tbl.Columns)
                {
                    exp_wrksheet.Range["A2"].Offset[row_inx, col_inx].Value = dr[dcol].ToString();
                    col_inx++;
                }
                row_inx++;
            }
            Export_To_FAI(f_format, "", "", export_wrkbook);
        }
        public void Export_FAI_Data2(myExcel.Workbook src_format_wrkbook, string tar_ItemCode, string tar_LotNo)
        {
            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            xlsApp.Visible = true;
            myExcel.Workbook exp_file;
            exp_file = xlsApp.Workbooks.Add();
            Export_FAI_Batch(tar_ItemCode, tar_LotNo, exp_file);
            Export_To_FAI2(src_format_wrkbook, tar_ItemCode, tar_LotNo, exp_file);
        }
        public void Export_Materials_Data(myExcel.Workbook tar_wrkbook, string tar_ItemCode, string tar_LotNo)
        {
           // myExcel.Workbook tar_wrkbook = TDMK_Code.open_excel_file(src_file, "", "");
            myExcel._Worksheet tar_wrksheet = tar_wrkbook.Sheets["Materials"];
            myExcel.Range tar_rgn = tar_wrksheet.Range["A12"];
            char[] split_char = new char[] { '-', ' ' };
            string[] Materials_tbl_lst = GetAllTables(myVar.sqlcon_Materials);
            foreach (string c in Materials_tbl_lst)
            {
                string sel_tbl = c;
                sel_tbl = sel_tbl.Trim(split_char).ToUpper();
                tar_wrksheet.Select();
                int row_inx = 0;
                while (row_inx < 200)
                {
                    myExcel.Range sel_rgn = tar_rgn.Offset[row_inx, 0];
                    string sel_rgn_val = sel_rgn.Value;
                    if (checkDBNull(sel_rgn_val) != "")
                    {
                        sel_rgn_val = sel_rgn_val.Trim(split_char).ToUpper();
                        if (sel_rgn_val == sel_tbl)
                        {
                            DataTable cur_dt = TDMK_Code.Datatable_Filter(myVar.sqlcon_Materials, sel_tbl, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { tar_ItemCode, tar_LotNo }));
                            int dt_row = cur_dt.Rows.Count;
                            int dt_col = cur_dt.Columns.Count;
                            for (int i = 0; i < dt_row; i++)
                            {
                                int row_offset_val = i * (dt_col - 2);
                                myExcel.Range data_rgn = sel_rgn.Offset[row_offset_val, 3];
                                for (int j = 3; j < dt_col; j++)
                                {
                                    data_rgn.Offset[j - 3, 0].Value = cur_dt.Rows[i][j].ToString();
                                }
                            }
                            break;
                        }
                    }
                    row_inx++;
                }
            }
        }
        public void Export_Recycle_Data(myExcel.Workbook tar_wrkbook, string tar_ItemCode, string tar_LotNo)
        {
            //myExcel.Workbook tar_wrkbook = TDMK_Code.open_excel_file(src_file, "", "");
            myExcel._Worksheet tar_wrksheet = tar_wrkbook.Sheets["Recycled PGC & copper"];
            myExcel.Range tar_rgn = tar_wrksheet.Range["B2"];
            char[] split_char = new char[] { '-', ' ', '_' };
            string[] Recycle_tbl_lst = GetAllTables(myVar.sqlcon_Recycle);
            tar_wrksheet.Select();
            int col_inx = 0;
            foreach (string c in Recycle_tbl_lst)
            {
                string sel_tbl = c;
                sel_tbl = sel_tbl.Trim(split_char).ToUpper();
                for (int k = 0; k < 4; k++)
                {
                    myExcel.Range sel_rgn = tar_rgn.Offset[0, k];
                    string sel_rgn_val = sel_rgn.Value;
                    if (checkDBNull(sel_rgn_val) != "")
                    {
                        sel_rgn_val = sel_rgn_val.Trim(split_char).ToUpper();
                        if (sel_rgn_val == sel_tbl)
                        {
                            //DataTable cur_dt = TDMK_Code.Datatable_Filter(myVar.sqlcon_Recycle, sel_tbl, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { tar_ItemCode, tar_LotNo }));
                            DataTable cur_dt = TDMK_Code.Datatable_Filter(myVar.sqlcon_Recycle, sel_tbl, "ItemCode = '"+tar_ItemCode+"'");
                            int dt_row = cur_dt.Rows.Count;
                            int dt_col = cur_dt.Columns.Count;
                            for (int i = 0; i < dt_row; i++)
                            {
                                //int row_offset_val = i * (dt_col - 2);
                                myExcel.Range data_rgn = sel_rgn.Offset[i + 1, 0];
                                for (int j = 3; j < dt_col; j++)
                                {
                                    data_rgn.Offset[j - 3, 0].Value = cur_dt.Rows[i][j].ToString();
                                }
                            }
                            break;
                        }
                    }

                }
                col_inx++;
            }
        }
        public void Export_fromExcel(myExcel.Workbook myWrkbook, string ItemCode, string LotNo, string xlFile_path)
        {
            //string src_file = f_file;
           // myExcel.Workbook myWrkbook = TDMK_Code.open_excel_file(src_file, "", "");
            string tar_folder = xlFile_path;
            string[] folders = System.IO.Directory.GetDirectories(tar_folder);
            string tar_file_name = ItemCode + "-" + LotNo;
            List<string> ignoredList = new List<string>() { "FAI","SPC", "CPK", "IPQC", "Recycle", "Materials", "Coverpage", "User Guidelines", "Declaration and Contents" };
            string ignoredSheet = Path.Combine(myVar.app_path, "IgnoredSheet.txt");
            string[] ignoredList_arr = read_config_arr2(ignoredSheet);
            ignoredList.AddRange(ignoredList_arr);
            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            xlsApp.DisplayAlerts = false;
            foreach (myExcel.Worksheet wrksheet in myWrkbook.Sheets)
            {
                string sel_tbl = wrksheet.Name;
                foreach (string c in folders)
                {
                    if (c.Contains(sel_tbl.Trim()))
                    {
                        if (TDMK_Code.check_exist_list_index(c, ignoredList)==-1)
                        {
                            if (!wrksheet.ProtectContents)
                            {
                                Clear_data(myWrkbook, sel_tbl);
                                string[] list_file = Directory.GetFiles(c, "*.xlsx").Where(s => s.Contains(tar_file_name)).ToArray();
                                if (list_file.Length > 0)
                                {
                                    myExcel.Worksheet mywrksheet = myWrkbook.Sheets[sel_tbl];
                                    myExcel.Workbook sel_wrkbook = TDMK_Code.open_excel_file(list_file[0], "", "");
                                    myExcel.Worksheet sel_wrksheet = sel_wrkbook.Sheets[1];
                                    Copy_data(mywrksheet, sel_wrksheet);
                                    sel_wrkbook.Close(false);
                                }
                            }
                            else
                            {
                                string mess_cont = wrksheet.Name + " is protected contents. Cannot modified!";
                                Error_log(app_path, mess_cont);
                            }
                        }
                        break;
                    }
                }
            }
            xlsApp.DisplayAlerts = true;
        }
        public void Export_fromExcel2(myExcel.Workbook myWrkbook, string ItemCode, string LotNo, string xlFile_path)
        {
            //string src_file = f_file;
            // myExcel.Workbook myWrkbook = TDMK_Code.open_excel_file(src_file, "", "");
            string tar_folder = xlFile_path;
            string[] folders = System.IO.Directory.GetDirectories(tar_folder);
            string tar_file_name = ItemCode + "-" + LotNo;
            List<string> ignoredList = new List<string>() { "FAI", "SPC", "CPK", "IPQC", "Recycle", "Materials", "Coverpage", "User Guidelines", "Declaration and Contents" };
            string ignoredSheet = Path.Combine(myVar.app_path, "IgnoredSheet.txt");
            string[] ignoredList_arr = read_config_arr2(ignoredSheet);
            ignoredList.AddRange(ignoredList_arr);
            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            //xlsApp.DisplayAlerts = false;
            foreach (myExcel.Worksheet wrksheet in myWrkbook.Sheets)
            {
                string sel_tbl = wrksheet.Name;
                int sht_inx = wrksheet.Index;
                foreach (string c in folders)
                {
                    if (c.Contains(sel_tbl.Trim()))
                    {
                        if (TDMK_Code.check_exist_list_index(c, ignoredList) == -1)
                        {

                            string[] list_file = Directory.GetFiles(c, "*.xlsx").Where(s => s.Contains(tar_file_name)).ToArray();
                            if (list_file.Length > 0)
                            {
                                myExcel.Worksheet mywrksheet = myWrkbook.Sheets[sel_tbl];
                                myExcel.Workbook sel_wrkbook = TDMK_Code.open_excel_file(list_file[0], "", "");
                                myExcel.Worksheet sel_wrksheet = sel_wrkbook.Sheets[1];
                                xlsApp.DisplayAlerts = false;
                                mywrksheet.Delete();
                                xlsApp.DisplayAlerts = true;
                                sel_wrksheet.Copy(Before: myWrkbook.Sheets[sht_inx]);
                                sel_wrkbook.Close(false);
                            }
                        }
                        break;
                    }
                }
            }
            //xlsApp.DisplayAlerts = true;
        }


        public void Export_all_from_excel(myExcel.Workbook myWrkbook,List<myExcel.Workbook> lst_wb)
        {
            //string src_file = f_file;
            // myExcel.Workbook myWrkbook = TDMK_Code.open_excel_file(src_file, "", "");
          
            List<string> ignoredList = new List<string>() { "FAI", "SPC", "CPK", "IPQC", "Recycle", "Materials", "Coverpage", "User Guidelines", "Declaration and Contents" };
            string ignoredSheet = Path.Combine(myVar.app_path, "IgnoredSheet.txt");
            string[] ignoredList_arr = read_config_arr2(ignoredSheet);
            ignoredList.AddRange(ignoredList_arr);
            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            //xlsApp.DisplayAlerts = false;
            foreach (myExcel.Workbook currwkb in lst_wb)
            {
                //    foreach (myExcel.Worksheet wrksheet in myWrkbook.Sheets)
                //{
                string sel_sheet = currwkb.Sheets[1].Name;
                foreach (myExcel.Worksheet wrksheet in myWrkbook.Sheets)
                {
                    if (sel_sheet == wrksheet.Name)
                    {


                        myExcel.Worksheet mywrksheet = myWrkbook.Sheets[sel_sheet];
                        myExcel.Worksheet sel_wrksheet = currwkb.Sheets[1];
                        xlsApp.DisplayAlerts = false;
                        mywrksheet.Delete();
                        xlsApp.DisplayAlerts = true;
                        sel_wrksheet.Copy(Before: myWrkbook.Sheets[sel_sheet]);
                        currwkb.Close(false);
                       
                        break;
                    }
                }
            }
            //xlsApp.DisplayAlerts = true;
        }



        public void Copy_data(myExcel.Worksheet dest_wrksheet, myExcel.Worksheet src_wrksheet)
        {
            dest_wrksheet.Activate();
            myExcel.Range dest_rgn;
            if (dest_wrksheet.UsedRange.Rows.Count > 1)
            {
                dest_rgn = dest_wrksheet.Range[dest_wrksheet.Cells[2, 1], dest_wrksheet.Cells[dest_wrksheet.UsedRange.Rows.Count, dest_wrksheet.UsedRange.Columns.Count]];
                dest_rgn.Clear();
            }
            else
            {
                dest_rgn = dest_wrksheet.Range["A2"];
            }
            myExcel.Range src_rgn = src_wrksheet.Range[src_wrksheet.Cells[2, 1], src_wrksheet.Cells[src_wrksheet.UsedRange.Rows.Count, src_wrksheet.UsedRange.Columns.Count]];
            src_rgn.Copy(dest_rgn);
        }
        public void Copy_data2(myExcel.Worksheet dest_wrksheet, myExcel.Worksheet src_wrksheet)
        {
            dest_wrksheet.Activate();
            myExcel.Range dest_rgn;
            if (dest_wrksheet.UsedRange.Rows.Count > 1)
            {
                dest_rgn = dest_wrksheet.Range[dest_wrksheet.Cells[2, 1], dest_wrksheet.Cells[dest_wrksheet.UsedRange.Rows.Count, dest_wrksheet.UsedRange.Columns.Count]];
                dest_rgn.Clear();
            }
            else
            {
                dest_rgn = dest_wrksheet.Range["A2"];
            }
            myExcel.Range src_rgn = src_wrksheet.Range[src_wrksheet.Cells[2, 1], src_wrksheet.Cells[src_wrksheet.UsedRange.Rows.Count, src_wrksheet.UsedRange.Columns.Count]];
            src_rgn.Copy();
            dest_rgn.PasteSpecial(myExcel.XlPasteType.xlPasteColumnWidths);
            //src_rgn.Copy(dest_rgn);
            dest_rgn.PasteSpecial(myExcel.XlPasteType.xlPasteAll,myExcel.XlPasteSpecialOperation.xlPasteSpecialOperationNone);
        }
        public void Clear_data(myExcel.Workbook tar_wrk, string sheet_name)
        {
            myExcel.Worksheet dest_wrksheet = tar_wrk.Sheets[sheet_name];
            dest_wrksheet.Activate();            
            int used_row = dest_wrksheet.UsedRange.Rows.Count;
            int used_col = dest_wrksheet.UsedRange.Columns.Count;
            if (used_row > 1)
            {
                myExcel.Range dest_rgn = dest_wrksheet.Range[dest_wrksheet.Cells[2, 1], dest_wrksheet.Cells[used_row, used_col]];
                dest_rgn.Clear();
            }
            foreach (myExcel.Shape sh in dest_wrksheet.Shapes)
            {
                sh.Delete();
            }
        }
        public void fill_data(string tbl_name, myExcel.Workbook mywrkbook, string _itemcode, string _lotno)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { _itemcode, _lotno });
            //string src_file = f_name;// Path.Combine(myVar.app_path, "Temp_Folder", "NPI OK2Ship Report for Bare FPC_" + src_Itemcode + ".xlsm");
            DataTable Result_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_Declare, tbl_name, filter_str);
            DataSet temp_DS = new DataSet();
            TDMK_Code.fill_dataset(temp_DS, tbl_name + "_Address", myVar.sqlcon_Declare);
            DataTable results_addr = temp_DS.Tables[0];
            //myExcel.Workbook mywrkbook = TDMK_Code.open_excel_file(src_file, "", "");
            myExcel.Worksheet mywrksheet = mywrkbook.Sheets["Declaration and Contents"];
            foreach (DataRow t in Result_tbl.Rows)
            {
                for (int i = 3; i < Result_tbl.Columns.Count; i++)
                {
                    string result_item = Result_tbl.Columns[i].ColumnName;
                    string rgn_addr = TDMK_Code.Get_item_val(tbl_name + "_Address", "Items like '" + result_item + "'", "Address", myVar.sqlcon_Declare);
        clear_pass: if(mywrksheet.ProtectContents)
                    {
                        mywrksheet.Unprotect("Histogram_123");
                    }
                    try
                    {
                        mywrksheet.Range[rgn_addr].Value = t.Field<string>(result_item);
                    }
                    catch
                    {
                        goto clear_pass;
                    }
                    
                }
            }
        }
        public void Export_Declaration_Data(myExcel.Workbook mywrkbook, string _itemcode, string _lotno)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { _itemcode, _lotno });
            DataTable Result_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_Declare, "Results", filter_str);
            DataSet temp_DS = new DataSet();
            TDMK_Code.fill_dataset(temp_DS, "Results_Address", myVar.sqlcon_Declare);
            DataTable results_addr = temp_DS.Tables[0];
            myExcel.Worksheet mywrksheet = mywrkbook.Sheets["Declaration and Contents"];
            mywrksheet.Activate();
            foreach (DataRow t in Result_tbl.Rows)
            {
                string result_item = t.Field<string>("Items");
                string rgn_addr = TDMK_Code.Get_item_val("Results_Address", "Items like '" + result_item + "'", "Address", myVar.sqlcon_Declare);
    resume_lable: if (mywrksheet.ProtectContents)
                {
                    mywrksheet.Unprotect("Histogram_123");
                }
                try
                {
                    mywrksheet.Range[rgn_addr].Value = t.Field<string>("Results");
                }
                catch
                {
                    goto resume_lable;
                }
                
            }
            //fill_data("DE_Data", mywrkbook, _itemcode, _lotno);
            //fill_data("User_Data", mywrkbook, _itemcode, _lotno);
        }
        public void Export_Declaration_Data2(myExcel.Workbook mywrkbook, string _itemcode, string _lotno)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { _itemcode, _lotno });
            DataTable Result_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_Declare, "Results", filter_str);            
            myExcel.Workbook temp_Declare = TDMK_Code.Create_workbook();
            myExcel.Worksheet temp_sht = temp_Declare.Sheets[1];
            myExcel.Range temp_rgn = temp_sht.Range["A1"];
            for (int i=0;i< Result_tbl.Rows.Count;i++)
            {
                temp_rgn.Offset[i, 0].Value = Result_tbl.Rows[i]["Results"];
            }
            temp_sht.UsedRange.Copy();
            myExcel.Worksheet mywrksheet = mywrkbook.Sheets["Declaration and Contents"];
            mywrksheet.Activate();
            mywrksheet.Range["F44"].PasteSpecial(myExcel.XlPasteType.xlPasteValues, myExcel.XlPasteSpecialOperation.xlPasteSpecialOperationNone);
            temp_Declare.Close(false);

        }
        public string Create_date_string(DateTime src_date)
        {
            string _result = "";
            string tg_day;
            string tg_month;
            string tg_year;
            tg_day = src_date.Day.ToString();
            tg_month = src_date.Month.ToString();
            if (src_date.Day < 10)
            {
                tg_day = "0" + src_date.Day.ToString();
            }
            else
            {
                tg_day = src_date.Day.ToString();
            }
            if (src_date.Month < 10)
            {
                tg_month = "0" + src_date.Month.ToString();
            }
            else
            {
                tg_month = src_date.Month.ToString();
            }
            tg_year = src_date.Year.ToString();
            _result = tg_year + tg_month + tg_day + "_" + src_date.Hour.ToString() + "h" + src_date.Minute.ToString() + "m" + src_date.Second.ToString() + "s_";
            return _result;
        }
        public void Export_FAI_Batch(string tar_ItemCode, string tar_LotNo, myExcel.Workbook tar_wrkbook, int qty=32)
        {
            char[] trim_char = new char[] { ' ', '\r', '\n' };
            DataTable FAI_Data_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { tar_ItemCode, tar_LotNo }));
            DataTable FAI_Spec_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "FAI_Spec", "ItemCode = '" + tar_ItemCode + "'");
            string[] sheetno = FAI_Spec_tbl.AsEnumerable().Select(r => r.Field<string>("SheetNo")).Distinct().ToArray();
            foreach (string sht in sheetno)
            {
                string[] FAI_No = FAI_Spec_tbl.AsEnumerable().Where(r => r.Field<string>("SheetNo") == sht).Select(r => r.Field<string>("FAI_No")).ToArray();

                if(sht.Contains("SPC"))
                {
                    int inx = 0;
                    foreach(string t in FAI_No)
                    {
                        string[] temp = t.Split('_');
                        string act_val = temp[1];
                        string[] temp2 = temp[0].Split('/');
                        FAI_No[inx] = temp2[0] + "_" + act_val;
                        inx++;
                    }
                }
                myExcel._Worksheet cur_wrksht = tar_wrkbook.Worksheets.Add(After: tar_wrkbook.Sheets[1]);
                cur_wrksht.Name = sht;
                myExcel.Range cur_rgn = cur_wrksht.Range["A1"];
                int c_inx = 0;
                foreach (string fai_no in FAI_No)
                {
                    cur_rgn.Offset[0, c_inx].Value = fai_no;
                    string[] fai_val = FAI_Data_tbl.AsEnumerable().Where(r => r.Field<string>("FAI_No") == fai_no).Select(r => r.Field<string>("FAI_Data")).ToArray();
                    if(fai_val.Length>0)
                    {
                        int r_inx = 0;
                        foreach (string _fai_val in fai_val)
                        {
                            cur_rgn.Offset[r_inx + 1, c_inx].Value = _fai_val.Trim(trim_char);
                            r_inx++;
                        }
                    }
                    else
                    {
                        for(int i=0;i<qty;i++)
                        {
                            cur_rgn.Offset[i + 1, c_inx].Value = "N/A";
                        }
                    }

                    c_inx++;
                }
            }
        }
        public void Error_log(string src_path, string err_msg)
        {
            string Log_dir = Path.Combine(src_path, "Error_Logging");
            if (!Directory.Exists(Log_dir))
            {
                Directory.CreateDirectory(Log_dir);
            }
            string log_name = Create_date_string(DateTime.Now).Trim('_');
            string log_file = Path.Combine(Log_dir, log_name + ".txt");
            StreamWriter log_writer = new StreamWriter(log_file, true);
            log_writer.WriteLine(DateTime.Now.ToString() + '\t' + err_msg);
            log_writer.Close();
        }
        public void Fill_Build_Info(myExcel.Workbook tar_wrkbook)
        {
            AutoCompleteStringCollection ignored_lst = new AutoCompleteStringCollection() { "Coverpage", "User Guidelines", "FAI", "CPK", "SPC", "Declaration and Contents" };
            myExcel.Worksheet declare_sht = tar_wrkbook.Sheets["Declaration and Contents"];
            string[] declare_items = new string[5];
            char[] trim_chr = { ':', ' ', '\r', '\n' };
            for (int i = 0; i < 5; i++)
            {
                declare_items[i] = declare_sht.Range["A26"].Offset[i, 0].Value;
            }
            foreach (myExcel.Worksheet sht in tar_wrkbook.Sheets)
            {
                if (!TDMK_Code.check_exist_list(sht.Name, ignored_lst))
                {
                    if (!sht.ProtectContents)
                    {
                        sht.Activate();
                        int col_num = sht.UsedRange.Columns.Count;
                        if (col_num > 20)
                        {
                            col_num = 20;
                        }
                        myExcel.Range tar_rgn = sht.Range["A2"];
                        int item_inx = 0;
                        foreach (string t in declare_items)
                        {
                            int col_inx = 0;
                            while (col_inx < col_num)
                            {
                                int row_inx = 0;
                                while (row_inx < 10)
                                {
                                    string sel_rgn_val = checkDBNull(tar_rgn.Offset[row_inx, col_inx].Value);
                                    if (sel_rgn_val != "")
                                    {
                                        if (sel_rgn_val.Trim(trim_chr).ToUpper() == t.Trim(trim_chr).ToUpper())
                                        {
                                resume_label: if (sht.ProtectContents)
                                            {
                                             sht.Unprotect("Histogram_123");
                                            }                                            
                                            try
                                            {
                                                tar_rgn.Offset[row_inx, col_inx + 1].Value = declare_sht.Range["A26"].Offset[item_inx, 1].Value;
                                                break;
                                            }
                                            catch
                                            {
                                                
                                                goto resume_label;
                                            }                                            
                                        }
                                    }
                                    row_inx++;
                                }
                                col_inx++;
                            }
                            item_inx++;
                        }
                    }
                }
            }
        }
        public bool Recycle_data_process(string src_master_file)
        {
            bool _result = false;
            try
            {
                AutoCompleteStringCollection items_lst = new AutoCompleteStringCollection();
                AutoCompleteStringCollection PGC_items = TDMK_Code.Load_Item_Names("PGC", "ItemCode", myVar.sqlcon_Recycle);
                AutoCompleteStringCollection Copper_Foil_items = TDMK_Code.Load_Item_Names("Copper Foil", "ItemCode", myVar.sqlcon_Recycle);
                AutoCompleteStringCollection Copper_Salt_items = TDMK_Code.Load_Item_Names("Copper Salt", "ItemCode", myVar.sqlcon_Recycle);
                AutoCompleteStringCollection Copper_Pallet_items = TDMK_Code.Load_Item_Names("Copper Pallet", "ItemCode", myVar.sqlcon_Recycle);
                AutoCompleteStringCollection[] all_item_lst = new AutoCompleteStringCollection[] { PGC_items, Copper_Foil_items, Copper_Salt_items, Copper_Pallet_items };
                string[] item = new string[6];
                item[0] = "ID";
                item[1] = "ItemCode";
                item[2] = "LotNo";
                item[3] = "Vendor";
                item[4] = "PartNumber";
                item[5] = "LotNumber";
                string[] tbl_item = new string[] { "PGC", "Copper Foil", "Copper Salt", "Copper Pallet" };
                myExcel.Workbook export_wrkbook = TDMK_Code.open_excel_file(src_master_file, "", "");
                myExcel.Worksheet tar_wrksht = export_wrkbook.Sheets[1];
                myExcel.Range itemcode_rgn = tar_wrksht.Range["D4"];
                myExcel.Range data_rgn = tar_wrksht.Range["F4"];
                int r_inx = 0;
                while (checkDBNull(itemcode_rgn.Offset[r_inx, 0].Value) != "")
                {
                    string _item = checkDBNull(itemcode_rgn.Offset[r_inx, 0].Value);
                    if (!TDMK_Code.check_exist(_item, items_lst))
                    {
                        string[] item_val = new string[6];
                        for (int i = 0; i < 4; i++)
                        {
                            item_val[1] = _item;
                            for (int j = 0; j < 3; j++)
                            {
                                item_val[3 + j] = data_rgn.Offset[r_inx, 3 * i + j].Value;
                            }
                            if (!TDMK_Code.check_exist(_item, all_item_lst[i]))
                            {
                                item_val[0] = (TDMK_Code.SQL_MAX(tbl_item[i], "ID", myVar.sqlcon_Recycle) + 1).ToString();

                                TDMK_Code.insert_val_arr(tbl_item[i], myVar.sqlcon_Recycle, item, item_val);
                            }
                            else
                            {
                                string ID = TDMK_Code.Get_item_val(tbl_item[i], "ItemCode ='" + _item + "'", "ID", myVar.sqlcon_Recycle);
                                string[] update_item = new string[3];
                                string[] update_item_val = new string[3];
                                Array.Copy(item, 3, update_item, 0, 3);
                                Array.Copy(item_val, 3, update_item_val, 0, 3);
                                TDMK_Code.updatebyID_val_arr(tbl_item[i], myVar.sqlcon_Recycle, Convert.ToInt32(ID), update_item, update_item_val);
                            }
                        }
                        items_lst.Add(_item);
                    }
                    r_inx++;
                }
                export_wrkbook.Close();
                _result = true;
            }
            catch(Exception ex)
            {
                _result = false;
                Error_log(app_path, ex.ToString());
            }
            return _result;
        }
        public string[] summary_FAI(string tar_ItemCode, string tar_LotNo, SqlConnection tar_sqlcon)
        {
            char[] split_char = new char[] { '/', '_' };
            DataTable spec_dt = Load_Spec_Table(tar_sqlcon, tar_ItemCode);
            DataTable FAI_tbl = Load_FAI_Table(tar_sqlcon, "FAI_Auto", "FAI_No", "FAI_Data", tar_ItemCode, tar_LotNo);
            DataTable CPK_tbl = Load_CPK_Table(FAI_tbl, spec_dt);
            DataTable FAI_Spec_tbl = TDMK_Code.Datatable_Filter(tar_sqlcon, "FAI_Spec", "ItemCode = '" + tar_ItemCode + "'");
            string[] sheetno = FAI_Spec_tbl.AsEnumerable().Select(r => r.Field<string>("SheetNo")).Distinct().ToArray();
            string[] result = new string[sheetno.Length];
            int result_inx = 0;
            bool sum_en = false;
            if(FAI_tbl.Rows.Count >0)
            {
                sum_en = true;
            }
            else
            {
                sum_en = false;
            }
            foreach (string sht in sheetno)
            {
                if(sum_en)
                {
                    string[] FAI_NO_lst = FAI_Spec_tbl.AsEnumerable().Where(r => r.Field<string>("SheetNo") == sht).Select(r => r.Field<string>("FAI_No")).ToArray();
                    for (int i = 0; i < FAI_NO_lst.Length; i++)
                    {
                        if (FAI_NO_lst[i].Contains("SPC") || FAI_NO_lst[i].Contains("CPK"))
                        {
                            string[] temp2 = FAI_NO_lst[i].Split(split_char);
                            FAI_NO_lst[i] = temp2[0] + "_" + temp2[2];
                        }
                    }
                    int sum = 0;
                    foreach (string t in FAI_NO_lst)
                    {
                        bool cpk_ok = false;
                        bool fai_ok = false;
                        string sv = spec_dt.Rows[0][t].ToString();
                        string TolMax = spec_dt.Rows[1][t].ToString();
                        string TolMin = spec_dt.Rows[2][t].ToString();
                        string UL = "";
                        string LL = "";
                        if(myCode2.IsNumeric(TolMax))
                        {
                            UL = (Convert.ToDouble(sv) + Convert.ToDouble(TolMax)).ToString();
                        }
                        if (myCode2.IsNumeric(TolMin))
                        {
                            LL = (Convert.ToDouble(sv) - Convert.ToDouble(TolMin)).ToString();
                        }
                        //double _UL = Convert.ToDouble(sv) + Convert.ToDouble(TolMax);
                        //double _LL = Convert.ToDouble(sv) - Convert.ToDouble(TolMin);
                        foreach (DataRow dr in FAI_tbl.Rows)
                        {
                            string act_val = dr[t].ToString();
                            if (myCode2.check_in_limit(UL, LL, act_val,sv))
                            {
                                fai_ok = true;
                            }
                            else
                            {
                                fai_ok = false;
                                break;
                            }
                        }
                        if (myCode2.check_columns_existed(CPK_tbl, t))
                        {
                            if (Convert.ToDouble(CPK_tbl.Rows[7][t]) > 1.67)
                            {
                                cpk_ok = true;
                            }
                            else
                            {
                                cpk_ok = false;
                            }
                        }
                        else
                        {
                            cpk_ok = true;
                            fai_ok = true;
                        }
                        if (!(fai_ok && cpk_ok))
                        {
                            sum++;
                        }
                    }
                    result[result_inx] = sum.ToString() + "F/" + FAI_NO_lst.Length;
                }
                else
                {
                    result[result_inx] = "No data";
                }

                result_inx++;
            }
            return result;
        }
        public DataTable Load_Spec_Table(SqlConnection sqlcon, string tar_ItemCode)
        {
            AutoCompleteStringCollection FAI_SheetNo_list = new AutoCompleteStringCollection();
            FAI_SheetNo_list = TDMK_Code.Load_Item_Names_Filter(sqlcon, "FAI_Spec", "SheetNo", TDMK_Code.filter_str(new string[] { "ItemCode", "SheetNo" }, new string[] { tar_ItemCode, "%FAI%" }));
            string[] items = new string[6];
            string[] items_val = new string[6];
            items[0] = "ItemCode";
            items[1] = "LotNo";
            items[2] = "FAI_No";
            items[3] = "Instrument";
            items[4] = "NormDim";
            items[5] = "SheetNo";
            items_val[0] = tar_ItemCode;
            DataTable cur_dt = new DataTable();
            //int spec_col_inx = 0;
            foreach (string sht in FAI_SheetNo_list)
            {
                items_val[5] = sht;
                items_val[2] = "";
                AutoCompleteStringCollection FAI_No_lst = TDMK_Code.Load_Item_Filter_str(sqlcon, "FAI_Spec", "FAI_No", TDMK_Code.filter_str(items, items_val));
                foreach (string t in FAI_No_lst)
                {
                    items_val[2] = t;
                    string c = TDMK_Code.Get_item_val("FAI_Spec", TDMK_Code.filter_str(items, items_val), "NormDim", sqlcon);
                    string col_val = t;// + "_" + c;
                    if (!myCode2.check_columns_existed(cur_dt, col_val))
                    {

                        cur_dt.Columns.Add(col_val);
                    }
                    if (cur_dt.Rows.Count == 0)
                    {
                        for (int i = 0; i < 6; i++)
                        {
                            cur_dt.Rows.Add();
                        }
                    }
                    DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Spec", TDMK_Code.filter_str(items, items_val));
                    cur_dt.Rows[0][col_val] = spec_dt.Rows[0]["NormDim"];
                    cur_dt.Rows[1][col_val] = spec_dt.Rows[0]["TolMax"];
                    cur_dt.Rows[2][col_val] = spec_dt.Rows[0]["TolMin"];
                    cur_dt.Rows[3][col_val] = spec_dt.Rows[0]["Distribution"];
                    cur_dt.Rows[4][col_val] = spec_dt.Rows[0]["Instrument"];
                    cur_dt.Rows[5][col_val] = spec_dt.Rows[0]["SheetNo"];
                    //spec_col_inx++;
                }
            }
            return cur_dt;
        }
        public DataTable Load_FAI_Table(SqlConnection tar_sqlcon, string tar_tbl, string FAI_No_name, string FAI_Data_Col_name, string tar_ItemCode, string tar_LotNo)
        {
            string[] flt_items = new string[4];
            string[] flt_item_vals = new string[4];
            AutoCompleteStringCollection FAI_lst;
            AutoCompleteStringCollection SheetNo_lst;
            flt_items[0] = "ItemCode";
            flt_items[1] = "LotNo";
            flt_items[2] = "FAI_No";
            flt_items[3] = "SheetNo";
            flt_item_vals[0] = tar_ItemCode;
            flt_item_vals[3] = "%FAI%";
            SheetNo_lst = TDMK_Code.Load_Item_Names_Filter(tar_sqlcon, "FAI_Spec", "SheetNo", TDMK_Code.filter_str(flt_items, flt_item_vals));
            DataTable tbl_data = new DataTable();
            foreach (string t_ShtNo in SheetNo_lst)
            {
                flt_item_vals[1] = "";
                flt_item_vals[2] = "";
                flt_item_vals[3] = t_ShtNo;
                FAI_lst = TDMK_Code.Load_Item_Filter_str(tar_sqlcon, "FAI_Spec", FAI_No_name, TDMK_Code.filter_str(flt_items, flt_item_vals));
                foreach (string c in FAI_lst)
                {
                    AutoCompleteStringCollection FAI_vals;
                    flt_item_vals[1] = tar_LotNo;
                    flt_item_vals[2] = c;
                    flt_item_vals[3] = "";
                    FAI_vals = TDMK_Code.Load_Item_Filter_str(tar_sqlcon, tar_tbl, FAI_Data_Col_name, TDMK_Code.filter_str(flt_items, flt_item_vals));
                    if (!myCode2.check_columns_existed(tbl_data, c))
                    {
                        tbl_data.Columns.Add(c);
                    }
                    int tbl_row_count = tbl_data.Rows.Count;
                    if (tbl_row_count < FAI_vals.Count)
                    {
                        for (int i = 0; i < FAI_vals.Count - tbl_row_count; i++)
                        {
                            tbl_data.Rows.Add();
                        }
                    }
                    int inx = 0;
                    foreach (string t in FAI_vals)
                    {
                        tbl_data.Rows[inx][c] = t;
                        inx++;
                    }
                }
            }
            return tbl_data;
        }
        public DataTable Load_CPK_Table(DataTable src_FAI_Data, DataTable spec_dt)
        {
            DataTable test_tbl = src_FAI_Data;// DGV_To_Table(DGV_DataView);
            int arr_num = test_tbl.Columns.Count;// col_list.Count;
            TDMK_OK2SHIP.Calculate_CPK[] myCalc_CPK = new TDMK_OK2SHIP.Calculate_CPK[arr_num];
            TDMK_OK2SHIP.FAI_Histogram_Data[] myHistogram_data = new TDMK_OK2SHIP.FAI_Histogram_Data[arr_num];
            int col_list_inx = 0;
            foreach (DataColumn tbl_col in test_tbl.Columns)
            {
                TDMK_OK2SHIP.FAI_Spec sel_FAI_test;
                string t = tbl_col.ColumnName;
                foreach (DataColumn dt_c in spec_dt.Columns)
                {
                    if (dt_c.ColumnName == t)
                    {
                        string t_FAIName = t;
                        string t_FAI_Setval = spec_dt.Rows[0][dt_c].ToString();
                        string t_FAI_UL = spec_dt.Rows[1][dt_c].ToString();
                        string t_FAI_LL = spec_dt.Rows[2][dt_c].ToString();
                        string t_checkside = spec_dt.Rows[3][dt_c].ToString();
                        string t_FAI_sheetno = spec_dt.Rows[4][dt_c].ToString();
                        string t_FAI_instrument = spec_dt.Rows[5][dt_c].ToString();
                        sel_FAI_test = new TDMK_OK2SHIP.FAI_Spec(t_FAIName, t_checkside, t_FAI_Setval, t_FAI_UL, t_FAI_LL, t_FAI_sheetno, t_FAI_instrument);
                        string[] temp_FAI_data = test_tbl.AsEnumerable().Select(r => r.Field<string>(t)).ToArray();
                        Double[] data_arr = new double[temp_FAI_data.Length];
                        int inx = 0;
                        foreach (string c in temp_FAI_data)
                        {
                            if ((myCode2.checkDBNull(c) != "") && myCode2.IsNumeric(c))
                            {
                                data_arr[inx] = Convert.ToDouble(c);
                                inx++;
                            }
                        }
                        Array.Resize(ref data_arr, inx);
                        if (inx > 0)
                        {
                            myCode2.Calcul_CPK_FAI2(sel_FAI_test, data_arr, ref myCalc_CPK[col_list_inx], ref myHistogram_data[col_list_inx]);
                        }
                        break;
                    }
                }
                col_list_inx++;
            }
            DataTable tbl_CPK = new DataTable();
            foreach (TDMK_OK2SHIP.Calculate_CPK item_CPK in myCalc_CPK)
            {
                string CPK_col_name = item_CPK.FAI_No;
                if (CPK_col_name != null)
                {
                    if (!myCode2.check_columns_existed(tbl_CPK, CPK_col_name))
                    {
                        tbl_CPK.Columns.Add(CPK_col_name);
                    }
                    if (tbl_CPK.Rows.Count == 0)
                    {
                        for (int i = 0; i < 9; i++)
                        {
                            tbl_CPK.Rows.Add();
                        }
                    }
                    tbl_CPK.Rows[0][CPK_col_name] = Math.Round(item_CPK.stdev, 4).ToString();
                    tbl_CPK.Rows[1][CPK_col_name] = Math.Round(item_CPK.mean, 4).ToString();
                    tbl_CPK.Rows[2][CPK_col_name] = Math.Round(item_CPK.max, 3).ToString();
                    tbl_CPK.Rows[3][CPK_col_name] = Math.Round(item_CPK.min, 3).ToString();
                    tbl_CPK.Rows[4][CPK_col_name] = item_CPK.CP;// Math.Round(item_CPK.CP, 3).ToString();
                    tbl_CPK.Rows[5][CPK_col_name] = item_CPK.CPKL;// Math.Round(item_CPK.CPKL, 3).ToString();
                    tbl_CPK.Rows[6][CPK_col_name] = item_CPK.CPKU;// Math.Round(item_CPK.CPKU, 3).ToString();
                    tbl_CPK.Rows[7][CPK_col_name] = Math.Round(item_CPK.CPK, 3).ToString();
                    tbl_CPK.Rows[8][CPK_col_name] = item_CPK.CPKM;
                }
            }
            return tbl_CPK;
        }
        public string get_target_Table(RadioButton src_RB)
        {
            string sel_tbl = "";
            if (src_RB.Checked)
            {
                sel_tbl = src_RB.Text;
            }
            return sel_tbl;
        }
        public bool IPQC_Process_result_all(string ItemCode, string LotNo, SqlConnection tar_sqlcon_IPQC, string[] all_proc)
        {
            bool _result = true;
            foreach (string proc in all_proc)
            {
                bool cur_proc = myCode2.IPQC_Process_result(ItemCode, LotNo, sqlcon_IPQC, proc);
                _result = _result && cur_proc;
                if (!_result)
                {
                    break;
                }
            }
            return _result;
        }
        public DataTable Declaration_Results(string ItemCode, string LotNo)
        {
            DataTable item_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_OK2SHIP, "Items_Details", "ItemCode ='" + ItemCode + "'");
            string flt_cmd = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            start_label: DataTable result_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_Declare, "Results", flt_cmd);
            if ((item_tbl.Rows.Count > 0) && (result_tbl.Rows.Count == 0))
            {
                foreach (DataRow dr in item_tbl.Rows)
                {
                    string items = dr["Items"].ToString();
                    string ID = (TDMK_Code.SQL_MAX("Results", "ID", myVar.sqlcon_Declare) + 1).ToString();
                    TDMK_Code.insert_val_arr("Results", myVar.sqlcon_Declare, new string[] { "ID", "ItemCode", "LotNo", "Items" }, new string[] { ID, ItemCode, LotNo, items });
                }
                goto start_label;
            }
            else
            {
                //if (item_tbl.Rows.Count > 0)
                //{
                int[] ID_lst = result_tbl.AsEnumerable().Where(r => r.Field<string>("Items").Contains("FAI") || r.Field<string>("Items").Contains("SPC")).Select(r => r.Field<int>("ID")).ToArray();
                string[] FAI_Result = summary_FAI(ItemCode, LotNo, myVar.sqlcon_SMT);
                for (int i = 0; i < ID_lst.Length; i++)
                {
                    string Id = ID_lst[i].ToString();
                    string result_val = FAI_Result[i];
                    TDMK_Code.update_item_val_filter("Results", myVar.sqlcon_Declare, "ID = " + Id, "Results", result_val);
                }
                int IPQC_Id = result_tbl.AsEnumerable().Where(r => r.Field<string>("Items").Contains("IPQC")).Select(r => r.Field<int>("ID")).ToArray()[0];
                string[] IPQC_process = new string[] { "Etching_Process", "Copper_Plating_Process", "Printing_Process", "UV_Process", "Cover_Lay_Process", "Roughness", "AU_NI" };

                string IPQC_result;// = (IPQC_Process_result_all(ItemCode, LotNo, myVar.sqlcon_IPQC, IPQC_process)).ToString();
                if (IPQC_Process_result_all(ItemCode, LotNo, myVar.sqlcon_IPQC, IPQC_process))
                {
                    IPQC_result = "Pass";
                }
                else
                {
                    IPQC_result = "Fail";
                }
                TDMK_Code.update_item_val_filter("Results", myVar.sqlcon_Declare, "ID = " + IPQC_Id, "Results", IPQC_result);
                /********************************* For VHX *************************/
                string[] arr_Item_VHX = { "Stack-up", "Impedance", "BVH & PTH", "CQRA - bHast", "CQRA - Thermal stress", "Solder Mask", "CQRA - Moisture absorption" };
                foreach (string item in arr_Item_VHX)
                {
                    int id = result_tbl.AsEnumerable().Where(r => r.Field<string>("Items").Contains(item)).Select(r => r.Field<int>("ID")).ToArray()[0];
                    string item_result = result_VHX(item, ItemCode, LotNo);
                    TDMK_Code.update_item_val_filter("Results", myVar.sqlcon_Declare, "ID = " + id, "Results", item_result);
                }

                /**********************************************************************/
                /************************************** For Echeck *********************/
                char[] trim_char = new char[] { ' ', '-', '_' };
                List<string> tar_process = new List<string>() { "ELECTRICAL", "CQRA_REFLOW", "CQRA_HOT_OIL", "CQRA_THERMAL_CYCLING", "CQRA_HEAT_SOAK", "CQRA_THERMAL_SHOCK", "CQRA_BENDING", "CQRA_THERMAL_CYCLING_AND_BEND", "CQRA_HEAT_SOAK_AND_BEND", "CQRA_SURVIVAL_REFLOW", "CQRA_SURVIVAL_HOT_OIL" };
                var query = from p in result_tbl.AsEnumerable().Where(x => tar_process.Any(t => myString(t, trim_char).ToUpper() == myString(x.Field<string>("Items"), trim_char).ToUpper())) select p;
                List<string> mylist = new List<string>();
                string ignoredSheet = Path.Combine(myVar.app_path, "IgnoredSheet.txt");
                string[] ignoredList_arr = read_config_arr2(ignoredSheet);
                foreach (var q in query)
                {
                    int inx = result_tbl.Rows.IndexOf(q);
                    var myProcess = from p in tar_process where myString(p, trim_char).ToUpper() == myString(q.Field<string>("Items"), trim_char).ToUpper() select p;
                    string sel_process = myProcess.First().ToString();
                    string item_result = "On-going";
                    if (ignoredList_arr.Any(x => myString(x, trim_char).ToUpper() == myString(q.Field<string>("Items"), trim_char).ToUpper()))
                    {
                        item_result = "On-going";
                    }
                    else
                    {
                        item_result = OK2SHIP_Process_Summary(ItemCode, LotNo, sel_process, myVar.sqlcon_OK2SHIP_Period2);
                    }
                    string id = result_tbl.Rows[inx]["ID"].ToString();
                    TDMK_Code.update_item_val_filter("Results", myVar.sqlcon_Declare, "ID = " + id, "Results", item_result);
                }
                /***************************************************************************/
                /**************************** For Auto fill content judgment ************************************************/
                List<string> auto_content_lst = new List<string>() { "Coverpage", "User guidelines", "Declaration and Contents", "Deviation summary", "Materials", "Recycled PGC and Copper", "Yield", "Hot bar", "Process Flow", "PMP", "UL", "FPC marking", "Via quality", "Solder mask", "Packaging" };
                List<string> auto_content_data = new List<string>() { "Pass", "Pass", "Info", "Deviation summary", "Material", "Not recycled copper", "Yeild", "NA", "Process flow", "PMP", "Pass", "Pass", "Pass", "Pass", "Pass" };
                var query2 = from p in result_tbl.AsEnumerable().Where(x => auto_content_lst.Any(t => myString(t, trim_char).ToUpper() == myString(x.Field<string>("Items"), trim_char).ToUpper())) select p;
                foreach (var q in query2)
                {
                    int inx = result_tbl.Rows.IndexOf(q);
                    var myProcess = from p in auto_content_lst where myString(p, trim_char).ToUpper() == myString(q.Field<string>("Items"), trim_char).ToUpper() select p;
                    string sel_process = myProcess.First().ToString();
                    int auto_inx = auto_content_lst.IndexOf(sel_process);
                    string item_result = auto_content_data[auto_inx];
                    string id = result_tbl.Rows[inx]["ID"].ToString();
                    TDMK_Code.update_item_val_filter("Results", myVar.sqlcon_Declare, "ID = " + id, "Results", item_result);
                }
                /***************************************************************************************************************/
                result_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_Declare, "Results", flt_cmd);

                //}
            }
            return result_tbl;
        }
        public string result_VHX(string process, string ItemCode, string LotNo)
        {
            string slt = "";
            switch (process)
            {
                case "Stack-up":
                    SortedDictionary<string, string> result_stackup = Stackup_result(ItemCode, LotNo, myVar.sqlcon_OK2SHIP_Period2);
                    foreach (var zone_result in result_stackup)
                    {
                        slt += zone_result.Key + ":" + zone_result.Value + ";";

                    }
                    break;


                case "Impedance":
                    SortedDictionary<string, List<string>> result_Impedance = Impedance_result(ItemCode, LotNo, myVar.sqlcon_OK2SHIP_Period2);
                    foreach (var zone_result in result_Impedance)
                    {
                        string result_region = "";
                        foreach (string item in zone_result.Value)
                        {
                            result_region += item + "_";
                        }
                        slt += zone_result.Key + ":" + result_region + ";";

                    }
                    break;

                case "BVH & PTH":
                    SortedDictionary<string, string> result_BVHPTH = BVH_PTH_result(ItemCode, LotNo, myVar.sqlcon_OK2SHIP_Period2);
                    foreach (var zone_result in result_BVHPTH)
                    {
                        slt += zone_result.Key + ":" + zone_result.Value + ";";

                    }
                    break;

                case "CQRA - bHast":
                    SortedDictionary<string, string> result_BHAST = BHAST_result(ItemCode, LotNo, myVar.sqlcon_OK2SHIP_Period2);
                    foreach (var zone_result in result_BHAST)
                    {
                        slt += zone_result.Key + ":" + zone_result.Value + ";";

                    }
                    break;

                case "CQRA - Thermal stress":
                    DataTable dt_data_Thermalstress = TDMK_Code.Datatable_Filter(myVar.sqlcon_OK2SHIP_Period2, "THERMAL_STRESS_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                    if(dt_data_Thermalstress.Rows.Count > 0)
                    {
                        slt = "YES";
                    }
                    else
                    {
                        slt = "NO";
                    }
                    break;

                case "Solder Mask":
                    DataTable dt_data_SolđerMask = TDMK_Code.Datatable_Filter(myVar.sqlcon_OK2SHIP_Period2, "SOLDERMASK_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                    if ( dt_data_SolđerMask.Rows.Count > 0)
                    {
                        slt = "YES";
                    }
                    else
                    {
                        slt = "NO";
                    }
                    break;

                case "CQRA - Moisture absorption":
                    slt = Moisture_result(ItemCode, LotNo, myVar.sqlcon_OK2SHIP_Period2);
                    break;


                default: break;
            }


            return slt;
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
        public SortedDictionary<string, string> Stackup_result(string ItemCode, string LotNo, SqlConnection sqlcon)
        {
            DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
            //DataTable tbl_spec = TDMK_Code.Datatable_Filter(myVar.sqlcon_OK2SHIP, "STACKUP_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { ItemCode }));
            string[] arr_zone = dt_data_summary.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();
            SortedDictionary<string, string> result_all_zone = new SortedDictionary<string, string> { };
            foreach (string zone in arr_zone)
            {
                DataTable dt_zone = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo, zone }));
                DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Zone" }, new string[] { ItemCode, zone.Replace("Zone_", "") }));
                if (dt_spec.Rows.Count > 0)
                {
                    Double USL = Double.Parse(dt_spec.Rows[0]["USL"].ToString());
                    Double LSL = Double.Parse(dt_spec.Rows[0]["LSL"].ToString());

                    string[] arr_spec_detail = dt_spec.Rows[0]["Spec_detail"].ToString().Split(';');

                    string filter = TDMK_Code.filter_str(new string[] { "Zone" }, new string[] { zone });
                    DataView dv_zone = dt_zone.AsDataView();
                    List<DataTable> tbl_pcs = new List<DataTable> { };
                    Get_ListTable(-1, dv_zone.ToTable(), new string[] { "Pcs_No" }, ref tbl_pcs, "Data");
                    SortedDictionary<int, List<double>> dic_spec = new SortedDictionary<int, List<double>> { };
                    bool check = true;
                    int k = 1;
                    foreach (string nominal_value in arr_spec_detail)
                    {
                        if (nominal_value != "")
                        {
                            List<double> lst_spec = new List<double> { };
                            double min = 0;
                            double max = 0;
                            if (nominal_value.Contains("+") || nominal_value.Contains("-"))
                            {
                                string nominal_val = nominal_value.Replace("+", ";").Replace("-", ";").Replace("/", ";");
                                max = double.Parse(nominal_val.Split(';')[0]) + double.Parse(nominal_val.Split(';')[1]);
                                min = double.Parse(nominal_val.Split(';')[0]) - double.Parse(nominal_val.Split(';')[3]);

                            }
                            else if (nominal_value.Contains("±"))
                            {
                                max = double.Parse(nominal_value.Split('±')[0]) + double.Parse(nominal_value.Split('±')[1]);
                                min = double.Parse(nominal_value.Split('±')[0]) - double.Parse(nominal_value.Split('±')[1]);
                            }
                            else
                            {
                                min = double.Parse(nominal_value) * 0.9;
                                max = double.Parse(nominal_value) * 1.1;
                            }

                            lst_spec.Add(max);
                            lst_spec.Add(min);
                            dic_spec.Add(k, lst_spec);
                        }

                        k++;


                    }
                    int count_pcs_NG = 0;

                    foreach (DataTable dt in tbl_pcs)
                    {
                        int s = 1;
                        Double total_pcs = 0;

                        foreach (DataRow dr in dt.Rows)
                        {

                            if (s <= dic_spec.Count)
                            {
                                double value = Double.Parse(dr["Data"].ToString());
                                total_pcs += value;
                                double max = dic_spec[s][0];
                                double min = dic_spec[s][1];

                                if (value > max || value < min)
                                {
                                    check = false;
                                    count_pcs_NG++;
                                    break;

                                }
                                s++;
                            }
                        }
                        if (check)
                        {
                            if (total_pcs > USL || total_pcs < LSL)
                            {
                                check = false;
                                count_pcs_NG++;
                                break;

                            }

                        }

                    }


                    string result_pcs = count_pcs_NG.ToString() + "F/" + tbl_pcs.Count.ToString();
                    result_all_zone.Add(zone, result_pcs);
                }
            }
            return result_all_zone;


        }
       
        public SortedDictionary<string, List<string>> Impedance_result(string ItemCode, string LotNo, SqlConnection sqlcon)
        {
            string[] arr_zone = { "IMPEDANCE", "TRACEWIDTH" };
            SortedDictionary<string, List<string>> result_all_zone = new SortedDictionary<string, List<string>> { };
            foreach (string process in arr_zone)
            {
                DataTable dt_process = TDMK_Code.Datatable_Filter(sqlcon, process + "_VAL", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, process + "_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { ItemCode }));
                List<string> lst_result_region = new List<string> { };
                if (process == "IMPEDANCE")
                {
                    if (dt_spec.Rows.Count > 0)
                    {
                        List<DataTable> tbl_region = new List<DataTable>() { };

                        Get_ListTable(-1, dt_process, new string[] { "Region" }, ref tbl_region, "Data");
                        foreach (DataTable dt in tbl_region)
                        {
                            int count_pcs_NG = 0;
                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                bool check = true;
                                int region = int.Parse(dt_process.Rows[i]["Region"].ToString());

                                if (region <= dt_spec.Rows.Count)
                                {
                                    Double max_target = Double.Parse(dt_spec.Rows[region - 1]["USL"].ToString());
                                    Double min_target = Double.Parse(dt_spec.Rows[region - 1]["LSL"].ToString());
                                    Double data = Double.Parse(dt.Rows[i]["Data"].ToString());
                                    if (data > max_target || data < min_target)
                                    {
                                        check = false;
                                        count_pcs_NG++;
                                    }
                                }
                            }

                            string result_pcs = count_pcs_NG.ToString() + "F/" + dt.Rows.Count.ToString();
                            lst_result_region.Add(result_pcs);
                        }
                    }
                }
                else if (process == "TRACEWIDTH")
                {
                    if (dt_spec.Rows.Count > 0)
                    {
                        List<DataTable> tbl_region = new List<DataTable>() { };

                        Get_ListTable(-1, dt_process, new string[] { "Region" }, ref tbl_region, "Data");
                        foreach (DataTable dt in tbl_region)
                        {
                            int count_pcs_NG = 0;
                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                bool check = true;
                                int region = int.Parse(dt.Rows[i]["Data_For"].ToString().Replace("Tracewidth_", ""));

                                if (region <= dt_spec.Rows.Count)
                                {
                                    Double max_target = Double.Parse(dt_spec.Rows[region - 1]["USL"].ToString());
                                    Double min_target = Double.Parse(dt_spec.Rows[region - 1]["LSL"].ToString());
                                    Double data = Double.Parse(dt.Rows[i]["Data"].ToString());
                                    if (data > max_target || data < min_target)
                                    {
                                        check = false;
                                        count_pcs_NG++;

                                    }
                                }
                            }

                            string result_pcs = count_pcs_NG.ToString() + "/" + dt.Rows.Count.ToString();
                            lst_result_region.Add(result_pcs);
                        }
                    }

                }

                result_all_zone.Add(process, lst_result_region);
            }
            return result_all_zone;


        }

        public SortedDictionary<string, string> BVH_PTH_result(string ItemCode, string LotNo, SqlConnection sqlcon)
        {

            string[] arr_zone = { "BVH_WITH_BONDING_SHEET", "BVH_WITHOUT_BONDING_SHEET", "PLATED_THROUGH_HOLE" };
            SortedDictionary<string, string> result_all_zone = new SortedDictionary<string, string> { };
            foreach (string process in arr_zone)
            {
                DataTable dt_process = TDMK_Code.Datatable_Filter(sqlcon, process, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_" + process.ToUpper(), TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { ItemCode }));
                if (dt_spec.Rows.Count > 0)
                {
                    int count_pcs_NG = 0;
                    for (int i = 0; i < dt_process.Rows.Count; i++)
                    {
                        bool check = true;
                        for (int k = 4; k < 12; k++)
                        {
                            string USL = dt_spec.Rows[0][k - 1].ToString().Replace("/", string.Empty);
                            string LSL = dt_spec.Rows[1][k - 1].ToString().Replace("/", string.Empty);
                            if (checkDBNull(dt_process.Rows[i][k]) != "")
                            {
                                if (!USL.Contains("NA"))
                                {
                                    if (Double.TryParse(dt_process.Rows[i][k].ToString(), out Double data))
                                    {
                                        if (data > Double.Parse(USL))
                                        {
                                            check = false;
                                            count_pcs_NG++;
                                            break;
                                        }
                                    }

                                }
                                if (!LSL.Contains("NA"))
                                {
                                    if (Double.TryParse(dt_process.Rows[i][k].ToString(), out Double data))
                                    {
                                        if (data < Double.Parse(LSL))
                                        {
                                            check = false;
                                            count_pcs_NG++;
                                            break;
                                        }
                                    }

                                }
                            }
                        }

                    }
                    string result_pcs = count_pcs_NG.ToString() + "F/" + dt_process.Rows.Count.ToString();
                    result_all_zone.Add(process, result_pcs);
                }
            }
            return result_all_zone;


        }

        public SortedDictionary<string, string> BHAST_result(string ItemCode, string LotNo, SqlConnection sqlcon)
        {
            SortedDictionary<string, string> result_all_zone = new SortedDictionary<string, string> { };
            DataTable dt_process = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_BHAST_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));


            for (int i = 4; i < dt_process.Columns.Count; i++)
            {
                bool check = true;
                foreach (DataRow dr in dt_process.Rows)
                {
                    if (checkDBNull(dr[i]) != "")
                    {
                        // Double data = Double.Parse(myCode.checkDBNull(dr[i]));
                        if (Double.TryParse(checkDBNull(dr[i]), out Double data))
                        {
                            if (data > Math.Pow(10, 4))
                            {
                                check = false;
                                break;
                            }
                        }
                    }
                }
                if (!check)
                {
                    result_all_zone.Add(dt_process.Columns[i].ColumnName, "False");
                }
            }

            return result_all_zone;
        }

        public string Moisture_result(string ItemCode, string LotNo, SqlConnection sqlcon)
        {
            string result = "";
            DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_MOISTURE_ABSORPTION", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
            int count_pcs_NG = 0;
            foreach(DataRow dr in dt_data_summary.Rows)
            {
                double value_before = Double.Parse(dr["Weight_before"].ToString());
                double value_after = Double.Parse(dr["Weight_after"].ToString());
                double val_judge = (value_after - value_before) / value_before;
                if(val_judge > 0.008)
                {
                    count_pcs_NG++;
                }

            }
            result = count_pcs_NG.ToString() + "F/" + dt_data_summary.Rows.Count;
            return result;
        }

        public int count_excel_data(string ItemCode, string LotNo)
        {
            int excel_file_sum = 0;
            string ignoredSheet = Path.Combine(myVar.app_path, "IgnoredSheet.txt");
            List<string> ignoredList = new List<string>() { "FAI", "SPC", "CPK", "IPQC", "Recycle", "Materials", "Coverpage", "User Guidelines", "Declaration and Contents" };
            string[] ignoredList_arr = read_config_arr2(ignoredSheet);
            ignoredList.AddRange(ignoredList_arr);
            excel_file_sum += ignoredList_arr.Length;
            string tar_item_lot = ItemCode + "-" + LotNo;
            string tar_folder = myVar.data_loc;
            string[] folders = System.IO.Directory.GetDirectories(tar_folder);
            string[] main_list = new string[folders.Length];
            int f_inx = 0;
            foreach (string c in folders)
            {
                FileInfo f = new FileInfo(c);
                string[] temp = f.Name.Split('.');
                if(myCode2.IsNumeric(temp[0])&&(temp[0]!="0"))
                {
                    main_list[f_inx] = f.Name;
                    f_inx++;
                }
            }
            Array.Resize(ref main_list, f_inx);
            foreach (string t in main_list)
            {
                if (TDMK_Code.check_exist_list_index(t, ignoredList) ==-1)
                {
                    string node_f = Path.Combine(tar_folder, t);
                    string[] node_name = Directory.GetFiles(node_f, "*.xlsx").Where(s => s.Contains(tar_item_lot)).ToArray();
                    if (node_name.Length > 0)
                    {
                        excel_file_sum++;
                    }
                }
            }
            return excel_file_sum;
        }
        public int count_excel_data2(string ItemCode, string LotNo, List<string> auto_pro_lst)
        {
            char[] trim_char = new char[] { ' ', '-', '_','&' };
            int excel_file_sum = 0;
            string ignoredSheet = Path.Combine(myVar.app_path, "IgnoredSheet.txt");
            List<string> ignoredList = new List<string>() { "FAI", "SPC", "CPK", "IPQC", "Recycle", "Materials", "Coverpage", "User Guidelines", "Declaration and Contents" };
            string[] ignoredList_arr = read_config_arr2(ignoredSheet);
            ignoredList.AddRange(ignoredList_arr);
            excel_file_sum += ignoredList_arr.Length;
            string tar_item_lot = ItemCode + "-" + LotNo;
            string tar_folder = myVar.data_loc;
            string[] folders = System.IO.Directory.GetDirectories(tar_folder);
            string[] main_list = new string[folders.Length];
            int f_inx = 0;
            foreach (string c in folders)
            {
                FileInfo f = new FileInfo(c);
                string[] temp = f.Name.Split('.');
                if (myCode2.IsNumeric(temp[0]) && (temp[0] != "0"))
                {
                    main_list[f_inx] = f.Name;
                    f_inx++;
                }
            }
            Array.Resize(ref main_list, f_inx);
            foreach(var m in main_list)
            {
                if(auto_pro_lst.Any(x=>myString(x.Replace("IMAGE",""),trim_char).ToUpper()== myString(m, trim_char).ToUpper()))
                {
                    if (TDMK_Code.check_exist_list_index(m, ignoredList) == -1)
                    {
                        ignoredList.Add(m);
                    }
                }
            }
            foreach (string t in main_list)
            {
                if (TDMK_Code.check_exist_list_index(t, ignoredList) == -1)
                {
                    string node_f = Path.Combine(tar_folder, t);
                    string[] node_name = Directory.GetFiles(node_f, "*.xlsx").Where(s => s.Contains(tar_item_lot)).ToArray();
                    if (node_name.Length > 0)
                    {
                        excel_file_sum++;
                    }
                }
            }
            return excel_file_sum;
        }
        public bool check_data_ready(SqlConnection tar_sqlcon, string ItemCode, string LotNo, AutoCompleteStringCollection src_ignoredlst)
        {
            bool _result = false;
            string[] src_tbl = GetAllTables(tar_sqlcon);
            string flt_str = "";
            if (LotNo != "")
            {
                flt_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            }
            else
            {
                flt_str = "ItemCode = '" + ItemCode + "'";
            }
            foreach (string sel_tbl in src_tbl)
            {
                if(src_ignoredlst!=null)
                {
                    if (!TDMK_Code.check_exist_list(sel_tbl, src_ignoredlst))
                    {
                        DataTable cur_sel_data = TDMK_Code.Datatable_Filter(tar_sqlcon, sel_tbl, flt_str);
                        if (cur_sel_data.Rows.Count > 0)
                        {
                            _result = true;
                        }
                        else
                        {
                            _result = false;
                            break;
                        }
                    }
                }
                else
                {
                    DataTable cur_sel_data = TDMK_Code.Datatable_Filter(tar_sqlcon, sel_tbl, flt_str);
                    if (cur_sel_data.Rows.Count > 0)
                    {
                        _result = true;
                    }
                    else
                    {
                        _result = false;
                        break;
                    }
                }

            }
            return _result;
        }
        public int Count_AutoData(string ItemCode, string LotNo)
        {
            int sum = 0;
            AutoCompleteStringCollection IPQC_ignored_lst = new AutoCompleteStringCollection { "All_Items", "SpecList","Sequence"};
            AutoCompleteStringCollection Declare_ignored_lst = new AutoCompleteStringCollection { "DE_Data","DE_Data_Address", "User_Data_Address", "Results_Address" };
            DataTable fai_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
            bool FAI_data_ready = false;
            if (fai_tbl.Rows.Count > 0)
            {
                FAI_data_ready = true;
            }
            else
            {
                FAI_data_ready = false;
            }
            bool IPQC_data_ready = check_data_ready(myVar.sqlcon_IPQC, ItemCode, LotNo, IPQC_ignored_lst);
            //bool Declare_data_ready = check_data_ready(myVar.sqlcon_Declare, ItemCode, "", Declare_ignored_lst);
            //bool Material_data_ready = check_data_ready(myVar.sqlcon_Materials, ItemCode, "", null);
            bool Recycled_data_ready = check_data_ready(myVar.sqlcon_Recycle, ItemCode, "", null);
            sum = sum + Convert.ToInt32(FAI_data_ready) + Convert.ToInt32(IPQC_data_ready) + Convert.ToInt32(Recycled_data_ready);            
            return sum;
        }
        public Dictionary<string, bool> FAI_IPQC_Recycle_process(string ItemCode, string LotNo)
        {
            Dictionary<string, bool> result = new Dictionary<string, bool>();
            AutoCompleteStringCollection IPQC_ignored_lst = new AutoCompleteStringCollection { "All_Items", "SpecList", "Sequence" };
            AutoCompleteStringCollection Declare_ignored_lst = new AutoCompleteStringCollection { "DE_Data", "DE_Data_Address", "User_Data_Address", "Results_Address" };
            DataTable fai_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
            bool FAI_data_ready = false;
            if (fai_tbl.Rows.Count > 0)
            {
                FAI_data_ready = true;
            }
            else
            {
                FAI_data_ready = false;
            }
            bool IPQC_data_ready = check_data_ready(myVar.sqlcon_IPQC, ItemCode, LotNo, IPQC_ignored_lst);
            bool Recycled_data_ready = check_data_ready(myVar.sqlcon_Recycle, ItemCode, "", null);
            result.Add("FAI", FAI_data_ready);
            result.Add("IPQC", IPQC_data_ready);
            result.Add("Recycled_PGC_copper", Recycled_data_ready);
            return result;
        }
        public void DES_Data_Factor(myExcel.Workbook tar_wrkbook)
        {
            myExcel.Worksheet cur_wrksht = tar_wrkbook.Sheets["IPQC Data"];
            myExcel.Range cur_rgn = cur_wrksht.Range["B13"];
            myExcel.Range data_rgn = cur_wrksht.Range["B15"];
            int col_inx = 0;
            while (checkDBNull(cur_rgn.Offset[0, col_inx].Value) != "")
            {
                string rgn_val = checkDBNull(cur_rgn.Offset[0, col_inx].Value);
                if (rgn_val.ToUpper().Contains("FACTOR"))
                {
                    for (int i = 0; i < 32; i++)
                    {
                        string B = checkDBNull(data_rgn.Offset[i, col_inx - 3].Value);
                        string D = checkDBNull(data_rgn.Offset[i, col_inx - 1].Value);
                        string C = checkDBNull(data_rgn.Offset[i, col_inx - 2].Value);
                        if (myCode2.IsNumeric(B) && (myCode2.IsNumeric(D)) && myCode2.IsNumeric(C))
                        {
                            double _B = Convert.ToDouble(B);
                            double _D = Convert.ToDouble(D);
                            double _C = Convert.ToDouble(C);
                            data_rgn.Offset[i, col_inx].Value = (2 * _B / (_D - _C)).ToString();
                        }
                    }
                }
                col_inx++;
            }
        }
        public void Fill_NA_Data(myExcel.Workbook tar_wrkbook, string wrksheet_name, string tar_data_rgn, int row_num)
        {
            myExcel.Worksheet cur_wrksht = tar_wrkbook.Sheets[wrksheet_name];
            myExcel.Range data_rgn = cur_wrksht.Range[tar_data_rgn];
            int col_inx = 0;
            while (checkDBNull(data_rgn.Offset[-2, col_inx].Value) != "")
            {

                for (int i = 0; i < row_num; i++)
                {
                    string rgn_val = checkDBNull(data_rgn.Offset[i, col_inx].Value);
                    if (rgn_val == "")
                    {
                        data_rgn.Offset[i, col_inx].Value = "N/A";
                    }
                }

                col_inx++;
            }
        }
        public bool Check_excel_data_ready(string ItemCode, string LotNo)
        {
            bool excel_file_rdy = true;
            List<string> ignoredList = new List<string>() { "FAI", "SPC", "CPK", "IPQC", "Recycle", "Materials", "Coverpage", "User Guidelines", "Declaration and Contents" };
            string ignoredSheet = Path.Combine(myVar.app_path, "IgnoredSheet.txt");
            string[] ignoredList_arr = read_config_arr2(ignoredSheet);
            ignoredList.AddRange(ignoredList_arr);
            string tar_item_lot = ItemCode + "-" + LotNo;
            string tar_folder = myVar.data_loc;
            string[] folders = System.IO.Directory.GetDirectories(tar_folder);
            string[] main_list = new string[folders.Length - 2];
            string[] tg_main_list = new string[folders.Length + 2];
            tg_main_list[0] = "ItemCode";
            tg_main_list[1] = "LotNo";
            for (int i = 0; i < folders.Length; i++)
            {
                foreach (string c in folders)
                {
                    FileInfo f = new FileInfo(c);
                    if (f.Name.Split('.')[0] == (i + 1).ToString())
                    {
                        main_list[i] = f.Name;
                        tg_main_list[i + 2] = f.Name;
                        break;
                    }
                }
            }
            foreach (string t in main_list)
            {
                if (TDMK_Code.check_exist_list_index(t,ignoredList) !=-1)
                {
                    string node_f = Path.Combine(tar_folder, t);
                    string[] node_name = Directory.GetFiles(node_f, "*.xlsx").Where(s => s.Contains(tar_item_lot)).ToArray();
                    if (node_name.Length == 0)
                    {
                        excel_file_rdy = false;
                        break;
                    }
                }
            }
            return excel_file_rdy;
        }
        public void break_links(myExcel.Workbook src_wrkbook)
        {
            Array links = src_wrkbook.LinkSources(myExcel.XlLinkType.xlLinkTypeExcelLinks) as Array;
            if (links != null)
            {
                for (int i = 1; i <= links.Length; i++)
                {
                    string _link = links.GetValue(i).ToString();
                    src_wrkbook.BreakLink(_link, myExcel.XlLinkType.xlLinkTypeExcelLinks);
                }
            }
        }
        public void remove_NamedRange(myExcel.Workbook src_wrkbook)
        {
            var ranges = src_wrkbook.Names;
            int leftoveritems;
            leftoveritems = ranges.Count;
            int error_count = 0;
            while (leftoveritems - error_count > 0)
            {
                int i = 1;
                try
                {
                    while (i <= leftoveritems)
                    {
                        var currentName = ranges.Item(i, Type.Missing, Type.Missing);
                        if (currentName.Name != "SheetNames")
                        {
                            currentName.Delete();
                        }
                        else
                        {
                            error_count++;
                        }
                        i++;
                    }
                }
                catch (Exception ex)
                {
                    error_count++;
                }
                
                ranges = src_wrkbook.Names;
                leftoveritems = ranges.Count;
            }
            MessageBox.Show("Finished");
        }
        //edit
        public void fill_PushingDate(myExcel.Workbook src_wrkbk, string ItemCode, string LotNo)
        {
            myExcel.Worksheet src_wrksht = src_wrkbk.Sheets["Declaration and Contents"];
            myExcel.Range src_rgn = src_wrksht.Range["D44"];
            myExcel.Range src_date = src_wrksht.Range["B34"];
            myExcel.Range dest_rgn = src_wrksht.Range["E44"];
            string ignoredSheet = Path.Combine(myVar.app_path, "IgnoredSheet.txt");
            List<string> ignoredList_lst = read_config_arr2(ignoredSheet).ToList();
            DataTable result_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_Declare, "Results", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
            string[] sheet_items = result_tbl.AsEnumerable().Select(r => r.Field<string>("Items")).ToArray();
            int row_inx = 0;
            while (checkDBNull(src_rgn.Offset[row_inx, 0].Value) != "")
            {
                string t = checkDBNull(src_rgn.Offset[row_inx, 0].Value);
                DateTime temp_val;
                if(row_inx < sheet_items.Length)
                {
                    string cur_sht = sheet_items[row_inx].Trim();
                    if (TDMK_Code.check_exist_list_index(cur_sht, ignoredList_lst) == -1)
                    {
                        temp_val = Convert.ToDateTime(src_date.Value);
                    }
                    else
                    {
                        if (t.Contains("5"))
                        {
                            temp_val = Convert.ToDateTime(src_date.Value).AddDays(5);
                        }
                        else
                        {
                            if (t.Contains("30"))
                            {
                                temp_val = Convert.ToDateTime(src_date.Value).AddDays(30);
                            }
                            else
                            {
                                temp_val = Convert.ToDateTime(src_date.Value);
                            }
                        }
                    }

                resume_label: if (src_wrksht.ProtectContents)
                    {
                        src_wrksht.Unprotect("Histogram_123");
                    }
                    try
                    {
                        dest_rgn.Offset[row_inx, 0].Value = temp_val.ToShortDateString();
                    }
                    catch
                    {
                        goto resume_label;
                    }
                    
                }
                row_inx++;

            }
        }
        public void Fill_BuildInfo_FAI(myExcel.Workbook src_wrkbk)
        {
            char[] trim_char = new char[] { ' ', ':', '\r', '\n' };
            myExcel.Worksheet src_wrksht = src_wrkbk.Sheets["Declaration and Contents"];
            myExcel.Range src_rgn = src_wrksht.Range["B28"];
            string[] tar_items = new string[] { "Part Number", "Part Description", "Revision", "ISR No", "Supplier", "Cavity / Tool #", "Inspector", "Date", "SMT Lot #", "Bare Flex Lot #" };
            string[] FAI_BuildInfo = new string[] { "C2", "C3", "F2", "F3", "J2", "J3" };
            string[] Process_flow_BuildInfo = new string[] { "D3", "D4", "F3", "F4", "H3", "H4", "J3", "J4", "L3", "L4" };
            string temp = (checkDBNull(src_rgn.Value));
            string[] temp2 = temp.Split('-');
            string part_num = "";// = temp2[0] + "-" + temp2[1];
            string revision = "";// = temp2[2]
            if (temp2.Length >= 3)
            {
                part_num = temp2[0] + "-" + temp2[1];
                revision = temp2[2];
            }
            List<string> src_items = new List<string>() { part_num, myCode2.checkDBNull(src_wrksht.Range["B26"].Value), revision, "NA", "Sumitomo Viet Nam", "1#", "SEEV", myCode2.checkDBNull(src_wrksht.Range["B34"].Value), "NA", myCode2.checkDBNull(src_wrksht.Range["B30"].Value) };
            foreach (myExcel.Worksheet sht in src_wrkbk.Sheets)
            {
                if (sht.Name.Contains("FAI") || sht.Name.Contains("SPC"))
                {
                    sht.Activate();

                    for (int i = 0; i < FAI_BuildInfo.Length; i++)
                    {
                        for (int col_inx = 0; col_inx < 10; col_inx++)
                        {
                            for (int row_inx = 0; row_inx < 3; row_inx++)
                            {
                                string cur_rgn = checkDBNull(sht.Range["B2"].Offset[row_inx, col_inx].Value);
                                if (cur_rgn.Trim(trim_char).ToUpper() == tar_items[i].Trim(trim_char).ToUpper())
                                {
                                    sht.Range["B2"].Offset[row_inx, col_inx + 1].Value = src_items[i];
                                    break;
                                }
                            }
                        }
                    }
                }
                if (sht.Name.Trim() == "Process flow")
                {
                    sht.Activate();
                    for (int i = 0; i < Process_flow_BuildInfo.Length; i++)
                    {
                        for (int col_inx = 0; col_inx < 10; col_inx++)
                        {
                            for (int row_inx = 0; row_inx < 3; row_inx++)
                            {
                                string cur_rgn = checkDBNull(sht.Range["B2"].Offset[row_inx, col_inx].Value);
                                if (cur_rgn.Trim(trim_char).ToUpper() == tar_items[i].Trim(trim_char).ToUpper())
                                {
                                    sht.Range["B2"].Offset[row_inx, col_inx + 1].Value = src_items[i];
                                    break;
                                }
                            }
                        }

                    }
                }
            }
        }
        public KeyValuePair<string, string> Get_Format_Info(string in_ItemCode, string format_loc)
        {
            KeyValuePair<string, string> result = new KeyValuePair<string, string>();

            DirectoryInfo temp = new DirectoryInfo(format_loc);
            FileInfo[] f_lst = temp.GetFiles();
            foreach (var f in f_lst)
            {
                if (f.FullName.Contains(in_ItemCode))
                {

                    string f_name = Path.GetFileNameWithoutExtension(f.Name);
                   // string ItemCode = f_name.Substring(f_name.Length - 7);
                    f_name = f_name.Replace("Ok2ship Request SEEV", "").Replace("Flex", "#");
                    string[] temp_f = f_name.Split('#');
                    //string[] temp_item = temp_f[1].Split(' ');
                    string[] temp_item = temp_f.Last().Split(' ');
                    //string ItemCode = temp_item[temp_item.Length - 1];
                    string ItemCode = temp_item.Last();
                    result = new KeyValuePair<string, string>(ItemCode, temp_f[0].Trim());
                    break;
                }
            }
            return result;
        }
        public bool check_in_spec_tbl(DataTable dt_spec, DataTable dt_data)
        {
            bool result = false;
            int r_inx = 0;
            foreach (DataRow dr in dt_data.Rows)
            {
                string USL = dt_spec.Rows[r_inx]["USL"].ToString();
                string LSL = dt_spec.Rows[r_inx]["LSL"].ToString();
                foreach (DataColumn dc in dt_data.Columns)
                {
                    string col_name = dc.ColumnName;
                    if (col_name.Contains("Before") || col_name.Contains("After"))
                    {
                        var sel_cell = myCode2.checkDBNull(dr[dc]);
                        if (sel_cell != "")
                        {
                            if (check_in_limit(USL, LSL, sel_cell))
                            {
                                result = true;
                            }
                            else
                            {
                                result = false;
                                break;
                            }
                        }
                    }
                }
                r_inx++;
            }
            return result;
        }
        public bool check_in_limit(string src_UL, string src_LL, string src_act_val)
        {
            bool result = false;
            if (myCode2.IsNumeric(src_act_val))
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
                            result = true;
                        }
                        else
                        {
                            result = false;
                        }
                    }
                    else
                    {
                        if (act_val <= UL)
                        {
                            result = true;
                        }
                        else
                        {
                            result = false;
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
                            result = true;
                        }
                        else
                        {
                            result = false;
                        }
                    }
                    else
                    {
                        result = false;
                    }
                }
            }
            return result;
        }
        public string OK2SHIP_Process_Summary(string ItemCode, string LotNo, string process, SqlConnection sqlcon)
        {
            string result = "";
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, process, filter_str); ;
            string spec_filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process });
            DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "NET_SPEC", spec_filter_str);
            double ref_rate = 0.1;
            if (process == "CQRA_SURVIVAL_HOT_OIL")
            {
                ref_rate = 0.05;
            }
            List<string> filter_lst = new List<string>() { "Yes", "A", "B" };
            List<int> tar_lst = OK2SHIP_NET_List(spec_dt, filter_lst, "Sel_Report");
            DataTable tar_spec_dt = OK2SHIP_Data_from_NET(spec_dt, tar_lst);
            List<DataTable> pcs_dt_lst = new List<DataTable>();
            Get_ListTable(-1, src_dt, new string[] { "PCS_No" }, ref pcs_dt_lst, "Before");
            if (pcs_dt_lst.Count > 0)
            {
                Dictionary<string, bool> cycle_result = new Dictionary<string, bool>();
                foreach (var pcs_dt in pcs_dt_lst)
                {
                    DataTable tar_pcs_dt = OK2SHIP_Data_from_NET(pcs_dt, tar_lst);
                    if(tar_pcs_dt.Rows.Count>0)
                    {
                        if (check_in_spec_tbl(tar_spec_dt, tar_pcs_dt))
                        {
                            cycle_result.Add(tar_pcs_dt.Rows[0]["Pcs_No"].ToString(), OK2SHIP_Check_Cycles_Data(tar_pcs_dt, ref_rate));
                        }
                        else
                        {
                            cycle_result.Add(tar_pcs_dt.Rows[0]["Pcs_No"].ToString(), false);
                        }
                    }
                }
                if(cycle_result.Count>0)
                {
                    int sum = 0;
                    foreach (var cycle in cycle_result.Values)
                    {
                        if (!cycle)
                        {
                            sum++;
                        }
                    }
                    result = sum.ToString() + "F/" + cycle_result.Count.ToString();
                }
                else
                {
                    result = "Cycle Data Error";
                }

            }
            else
            {
                MessageBox.Show( "Process "+process+" : Before Data not found!", "Warning");
                result = "Before Data Error";
            }
            return result;
        }
        public List<int> OK2SHIP_NET_List(DataTable spec_dt, List<string> filter_lst, string filter_col)
        {
            DataView dv;
            var query = from sp in spec_dt.AsEnumerable() where filter_lst.Any(lst => lst == sp.Field<string>(filter_col)) select sp;
            if (query.Count() > 0)
            {
                dv = query.AsDataView();
            }
            else
            {
                dv = spec_dt.AsDataView();
            }
            List<int> lst_index = new List<int>();
            for (int i = 0; i < dv.Count; i++)
            {
                DataRow dr = dv[i].Row;
                lst_index.Add(spec_dt.Rows.IndexOf(dr));
            }
            return lst_index;
        }
        public DataTable OK2SHIP_Data_from_NET(DataTable src_data, List<int> NET_list)
        {
            DataTable result = new DataTable();
            var query = from sp in src_data.AsEnumerable() where NET_list.Any(x => x == src_data.Rows.IndexOf(sp)) select sp;
            if (query.Count() > 0)
            {
                result = query.CopyToDataTable();
            }
            return result;
        }
        public bool OK2SHIP_Check_Cycles_Data(DataTable src_dt, double ref_rate)
        {
            bool result = false;
            try
            {
                foreach (DataRow dr in src_dt.Rows)
                {
                    foreach (DataColumn dc in src_dt.Columns)
                    {
                        if (dc.ColumnName.Contains("After"))
                        {
                            if (myCode2.checkDBNull(dr[dc]) != "")
                            {
                                if (myCode2.IsNumeric(dr[dc].ToString()))
                                {
                                    double bef_val = Convert.ToDouble(dr["Before"]);
                                    double aft_val = Convert.ToDouble(dr[dc]);
                                    double rate = Math.Abs(bef_val - aft_val) / bef_val;
                                    if (rate > ref_rate)
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
                        }
                    }
                }
                return result;
            }
            catch
            {
                MessageBox.Show(new Form { TopMost = true }, "Before Data not found", "Warning");
                return false;
            }

        }
        public string myString(string src_str, char[] replace_char)
        {
            string result = ""; //replace_char.Aggregate(src_str,(c1,c2)=>c1.Replace(c2,'\0'));
            char[] src_char = src_str.ToArray();
            foreach (char t in src_char)
            {
                if (replace_char.ToList().IndexOf(t) == -1)
                {
                    result = result + t.ToString();
                }
            }
            return result;
        }
        public Dictionary<string, bool> process_data_ready(string ItemCode, string LotNo, SqlConnection sqlcon, string[] ignoredList_arr)
        {
            List<string> echeck_process = new List<string>() { "ELECTRICAL", "CQRA_REFLOW", "CQRA_HOT_OIL", "CQRA_THERMAL_CYCLING", "CQRA_HEAT_SOAK", "CQRA_THERMAL_SHOCK", "CQRA_BENDING", "CQRA_THERMAL_CYCLING_AND_BEND", "CQRA_HEAT_SOAK_AND_BEND", "CQRA_SURVIVAL_REFLOW", "CQRA_SURVIVAL_HOT_OIL" };
            List<string> VHX_process = new List<string>() { "THERMAL_STRESS_IMAGE", "SOLDERMASK_IMAGE", "CQRA_MOISTURE_ABSORPTION" };
            List<string> BVH_process = new List<string>() { "BVH_WITH_BONDING_SHEET", "BVH_WITHOUT_BONDING_SHEET", "PLATED_THROUGH_HOLE", "BVH_PTH_LOGFILE" };
            List<string> stackup_process = new List<string>() { "STACKUP_IMAGE", "STACKUP_LOGFILE" };
            List<string> impedance_process = new List<string>() { "IMPEDANCE_VAL_LOGFILE", "IMPEDANCE_GRAPH", "IMPEDANCE_IMAGE", "TRACEWIDTH_IMAGE", "TRACEWIDTH_VAL" };
            List<string> bHast_process = new List<string>() { "CQRA_BHAST_IMAGE", "CQRA_BHAST_LOGFILE" };
            List<string> Camera_process = new List<string>() { "CQRA_IR_VIA_TO_VIA_IMAGE", "CQRA_CHEMICAL_RESISTANCE_IMAGE", "CQRA_IR_TRACE_TO_TRACE_IMAGE", "CQRA_IR_LAYER_TO_LAYER_IMAGE", "CQRA_FLUX_RESIST_IMAGE", "CQRA_HIGH_SPEED_BALL_SHEAR_IMAGE", "CQRA_DIELECTRIC_WITHSTANDING_IMAGE", "CQRA_SOLDERABILITY_IMAGE", "OQC_TEST_IMAGE" };
            List<string> ACF_process = new List<string>() { "ACF_CLEANING_IMAGE", "ACF_BEFORE_TAPE_TEST_IMAGE", "ACF_AFTER_TAPE_TEST_IMAGE", "ACF_GRAPH_FORCE_IMAGE","ACF_FLATNESS" };
            List<List<string>> calc_sum_lst = new List<List<string>>() { BVH_process, stackup_process, impedance_process, bHast_process, ACF_process };
            List<List<string>> calc_single_lst = new List<List<string>>() { echeck_process, VHX_process, Camera_process };
            List<string> calc_sum_name = new List<string>() { "BVH_PTH", "STACKUP", "IMPEDANCE", "CQRA_BHAST","ACF" };
            Dictionary<string, bool> process_result_lst = new Dictionary<string, bool>();
            for(int i=0;i< calc_sum_lst.Count;i++)
            {
                var lst = calc_sum_lst[i];
                bool data_OK = false;
                foreach (string t in lst)
                {
                    DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, t, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                    if (dt.Rows.Count > 0)
                    {
                        data_OK = true;
                    }
                    else
                    {
                        data_OK = false;
                        break;
                    }
                }
                process_result_lst.Add(calc_sum_name[i], data_OK);
            }
            foreach(var lst in calc_single_lst)
            {
                foreach (string t in lst)
                {
                    DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, t, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                    if (dt.Rows.Count > 0)
                    {
                        process_result_lst.Add(t, true);
                    }
                    else
                    {
                        process_result_lst.Add(t, false);
                    }
                }
            }
            char[] trim_char = new char[] { ' ', '-', '_' };
            foreach (string t in process_result_lst.Keys.ToList())
            {
                if( ignoredList_arr.Any(x=>myString(x, trim_char).ToUpper()==myString(t.Replace("IMAGE",""),trim_char).ToUpper()))
                {
                    process_result_lst[t] = true;
                }
            }    
            return process_result_lst;
        }
        public List<string> OK2SHIP_Item_list(string ItemCode, string LotNo, string tar_table, SqlConnection sqlcon)
        {
            List<string> result = new List<string>();
            DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, tar_table, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
            List<string> main_items = dt.AsEnumerable().Select(x => x.Field<string>("Items")).Distinct().ToList();
            foreach (string item in main_items)
            {
                string temp = item;
                if (item.Contains("FAI") || item.Contains("SPC"))
                {
                    temp = "FAI";
                }
                if (result.IndexOf(temp) == -1)
                {
                    result.Add(temp);
                }
            }
            return result;
        }
        public string Find_Cell_Addr(string search_key, string start_addr, myExcel.Worksheet tar_wrksht, bool left_to_right)
        {
            string result = "";
            myExcel.Range cycle_rgn = tar_wrksht.Range[start_addr];
            for (int i = 0; i < 100; i++)
            {
                myExcel.Range sel_rgn = cycle_rgn.Offset[i, 0];
                if (left_to_right)
                {
                    sel_rgn = cycle_rgn.Offset[0, i];
                }
                string sel_rgn_val = checkDBNull(sel_rgn.Value);
                if (sel_rgn_val.Replace(" ", "").ToUpper() == search_key.Replace(" ", "").ToUpper())
                {
                    result = sel_rgn.AddressLocal;
                    break;
                }
            }
            return result;
        }
        public string Find_Offset(string start_addr, myExcel.Worksheet tar_wrksht, bool left_to_right, string search_key = "")
        {
            string result = "";
            myExcel.Range cycle_rgn = tar_wrksht.Range[start_addr];
            int c_offset = 0;
            int empty_count = 0;
            while (true)
            {
                myExcel.Range sel_rgn = cycle_rgn.Offset[c_offset, 0];
                if (left_to_right)
                {
                    sel_rgn = cycle_rgn.Offset[0, c_offset];
                }
                string sel_rgn_val = checkDBNull(sel_rgn.Value);
                if (sel_rgn_val == "")
                {
                    empty_count++;
                    if (empty_count > 3)
                    {
                        break;
                    }
                }
                else
                {
                    empty_count = 0;
                }
                if (search_key != "")
                {
                    if (sel_rgn_val.Replace(" ", "").ToUpper() == search_key.Replace(" ", "").ToUpper())
                    {
                        result = sel_rgn.AddressLocal;
                        break;
                    }
                }
                else
                {
                    if (empty_count == 0)
                    {
                        result = sel_rgn.AddressLocal;
                        break;
                    }
                }
                c_offset++;
            }
            return result;
        }
        public Dictionary<string, int> Find_FAI_addr_qty(string start_addr, myExcel.Worksheet tar_wrksht, bool left_to_right)
        {
            Dictionary<string, int> result = new Dictionary<string, int>();
            myExcel.Range cycle_rgn = tar_wrksht.Range[start_addr];
            int offset = 0;
            int empty_count = 0;
            string data_addr = "";
            int qty = 0;
            while (true)
            {
                myExcel.Range sel_rgn = cycle_rgn.Offset[offset, 0];
                if (left_to_right)
                {
                    sel_rgn = cycle_rgn.Offset[0, offset];
                }
                string sel_rgn_val = checkDBNull(sel_rgn.Value);
                if (sel_rgn_val == "")
                {
                    empty_count++;
                    if (empty_count > 3)
                    {
                        break;
                    }
                }
                else
                {
                    if (myCode2.IsNumeric(sel_rgn_val))
                    {
                        if (data_addr == "")
                        {
                            data_addr = sel_rgn.AddressLocal;
                        }
                        qty++;
                    }
                    empty_count = 0;
                }
                offset++;
            }
            if (data_addr != "")
            {
                result.Add(data_addr, qty);
            }
            return result;
        }
    }
}