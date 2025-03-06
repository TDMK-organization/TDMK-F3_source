using OK2SHIP;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SEEV_DLL;
using TDMK_SQL;
using OfficeOpenXml;
using TDMK_EPPLUS_7;
using System.Diagnostics.Eventing.Reader;

namespace OK2SHIP_Software
{
    public class Echeck_Export
    {
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SEI_Lib myCode = new SEI_Lib();
        TDMK_EPPLUS7_lib Excel_Lib = new TDMK_EPPLUS7_lib();
        public string Export_HotOil(ExcelWorkbook report_wrkbk,SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string sheet_name)
        {
            string sel_sheet = find_sheet(report_wrkbk, sheet_name);
            if (sel_sheet != "")
            {
                ExcelWorksheet tar_sht = report_wrkbk.Worksheets[sel_sheet];
                string data_addr = Excel_Lib.Find_Start_Addr("A10", tar_sht, false, "Sample no.");// Get_start_range("Sample no.", "A10", tar_sht);
                ExcelRangeBase PTH_rgn = tar_sht.Cells[data_addr].Offset(1, 1);
                ExcelRangeBase BVH_rgn = tar_sht.Cells[data_addr].Offset(1, 33);
                ExcelRangeBase tar_rgn;
                List<DataTable> src_tbl_lst = new List<DataTable>();
                DataTable dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                List<string> pcs_sel_lst = Get_SelectedPcs_List(dt, "Sel_report", "Yes");// src_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_report").ToUpper() == "YES").Select(x => x.Field<string>("Pcs_No")).Distinct().ToList();
                if (pcs_sel_lst.Count != 0)
                {
                    dt = dt.AsEnumerable().Where(x => x.Field<string>("Sel_report").ToUpper() == "YES").CopyToDataTable();
                }
                Get_ListTable(-1, dt, new string[] { "ItemCode", "LotNo", "Region" }, ref src_tbl_lst, "Before");
                foreach (DataTable sel_dt in src_tbl_lst)
                {
                    if (sel_dt.Rows.Count > 0)
                    {
                        string region = sel_dt.Rows[0]["Region"].ToString();
                        if (region == "BVH")
                        {
                            tar_rgn = BVH_rgn;
                        }
                        else
                        {
                            tar_rgn = PTH_rgn;
                        }
                        DataTable _sel_dt = sel_dt.AsDataView().ToTable(false, new string[] { "Before", "After" });
                        int r_inx = 0;
                        foreach (DataRow dr in _sel_dt.Rows)
                        {
                            int c_inx = 0;
                            foreach (DataColumn dc in _sel_dt.Columns)
                            {
                                string cur_val = myCode.checkDBNull(dr[dc]);
                                if (TDMK_Code.IsNumeric(cur_val))
                                {
                                    tar_rgn.Offset(c_inx, r_inx).Value = Convert.ToDecimal(cur_val);
                                    tar_rgn.Offset(c_inx, r_inx).Style.Numberformat.Format = "#0.000";
                                }
                                c_inx++;
                            }
                            r_inx++;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Process of ItemCode: " + ItemCode + process_name + " not found!", "Warning");
            }
            return sel_sheet;

        }
        //public ExcelWorkbook create_export_wrk_old(string format_file, string process_name, string ItemCode, string Lotno)
        //{
        //    myExcel.Application xlsApp = TDMK_Code.StartExcel();
        //    xlsApp.DisplayAlerts = false;
        //    ExcelWorkbook wb = TDMK_Code.open_excel_file(format_file, "", "");
        //    string _process = process_name.Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
        //    string mySheet = "";
        //    foreach (ExcelWorksheet tg_sht in wb.Worksheets)
        //    {
        //        string cur_sht_name = tg_sht.Name.ToUpper().Replace("-", "").Replace(" ", "");
        //        if (cur_sht_name == _process)
        //        {
        //            mySheet = tg_sht.Name;
        //            break;
        //        }
        //    }
        //    if (mySheet != "")
        //    {

        //        ExcelWorkbook save_wb = TDMK_Code.Create_workbook();
        //        wb.Worksheets[mySheet].Copy(After: save_wb.Worksheets[1]);
        //        ExcelWorksheet del_sht = save_wb.Worksheets[1];
        //        del_sht.Delete();
        //        wb.Close();
        //        xlsApp.DisplayAlerts = true;
        //        return save_wb;
        //    }
        //    else
        //    {
        //        wb.Close();
        //        xlsApp.DisplayAlerts = true;
        //        return null;
        //    }
        //}
        //public string Get_start_range(string search_key, string start_addr, ExcelWorksheet tar_wrksht)
        //{
        //    string result = "";
        //    ExcelRangeBase cycle_rgn = tar_wrksht.Cells[start_addr];
        //    for (int i = 0; i < 100; i++)
        //    {
        //        ExcelRangeBase sel_rgn = cycle_rgn.Offset(i, 0];
        //        string sel_rgn_val = myCode.checkDBNull(sel_rgn.Value);
        //        if (sel_rgn_val.Replace(" ", "").ToUpper() == search_key.Replace(" ", "").ToUpper())
        //        {
        //            result = cycle_rgn.Offset(i + sel_rgn.MergeArea.Rows.Count, 0].Address;
        //            break;
        //        }
        //    }
        //    return result;
        //}
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
        public string find_format(string in_data_loc, string ItemCode)
        {
            string result = "";
            string[] files = Directory.GetFiles(Path.Combine(in_data_loc, "0.OK2SHIP report format"), "*" + ItemCode + "*.xlsm");
            if (files.Length > 0)
            {
                result = files[0];
            }
            return result;
        }
        public string find_sheet(ExcelWorkbook wb, string tar_process)
        {
            string mySheet = "";
            List<char> remove_char = new List<char> { ' ', '-','_','\r','\n' };
            foreach (ExcelWorksheet tg_sht in wb.Worksheets)
            {
                string cur_sht_name = remove_special_char(tg_sht.Name, remove_char).ToUpper();// tg_sht.Name.ToUpper().Replace("-", "").Replace(" ", "");
                string _tar_process = remove_special_char(tar_process, remove_char).ToUpper();
                if (cur_sht_name == remove_special_char(tar_process, remove_char).ToUpper())
                {
                    mySheet = tg_sht.Name;
                    break;
                }
            }
            return mySheet;
        }
        public string remove_special_char(string src_str, List<char> remove_char)
        {
            return new string(src_str.Where(x => remove_char.IndexOf(x) == -1).ToArray());
        }
        public string Export_Bending(ExcelWorkbook report_wrk,SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string sheet_name)
        {
            string sel_sheet = find_sheet(report_wrk, sheet_name);
            if (sel_sheet != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                string spec_filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process_name });
                DataTable spec_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", spec_filter_str);
                List<string> pcs_sel_lst = Get_SelectedPcs_List(src_tbl, "Sel_report", "Yes");// src_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_report").ToUpper() == "YES").Select(x => x.Field<string>("Pcs_No")).Distinct().ToList();
                if (pcs_sel_lst.Count != 0)
                {
                    src_tbl = src_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_report").ToUpper() == "YES").CopyToDataTable();
                }
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                ExcelWorksheet tar_wrksht = report_wrk.Worksheets[sel_sheet];
                string tar_rgn_addr = Excel_Lib.Find_Cell_Addr("Test item", "A10", tar_wrksht, false);// Get_start_range("Test item", "A10", tar_wrksht);
                int tar_qty_off = Excel_Lib.Get_Cells_Info(tar_wrksht, tar_wrksht.Cells[tar_rgn_addr]).row_qty;
                ExcelRangeBase tar_rgn = tar_wrksht.Cells[tar_rgn_addr].Offset(tar_qty_off, 1);
                string sum_rgn_addr = Excel_Lib.Find_Cell_Addr("Condition", "A10", tar_wrksht,false);// Get_start_range("Condition", "A10", tar_wrksht);
                int sum_qty_off = Excel_Lib.Get_Cells_Info(tar_wrksht, tar_wrksht.Cells[sum_rgn_addr]).row_qty;
                ExcelRangeBase sum_rgn = tar_wrksht.Cells[sum_rgn_addr].Offset(sum_qty_off+1, 2);
                ExcelRangeBase sum_rgn_before = tar_wrksht.Cells[sum_rgn_addr].Offset(sum_qty_off, 2);
                int col_inx = 0;
                foreach (var tbl in src_tbl_lst)
                {
                    string sum_before = "OK";
                    foreach(DataRow dr in tbl.Rows)
                    {
                        string Net_No = dr["Net_No"].ToString();
                        string before_data_val = dr["Before"].ToString();
                        string UL = spec_tbl.AsEnumerable().Where(x => x.Field<string>("Net_Name") == Net_No).Select(x => x.Field<string>("USL")).FirstOrDefault();
                        string LL = spec_tbl.AsEnumerable().Where(x => x.Field<string>("Net_Name") == Net_No).Select(x => x.Field<string>("LSL")).FirstOrDefault();
                        if(myCode.check_in_limit(UL,LL,before_data_val,""))
                        {
                            sum_before = "OK";
                        }
                        else
                        {
                            sum_before = "NG";
                            break;
                        }
                    }
                    sum_rgn_before.Offset(0, col_inx).Value = sum_before;
                    Dictionary<string, List<string>> data_lst = Get_List_Pair_data(tbl);
                    List<string> before_data = data_lst["Before"];
                    int r_offset = 0;
                    foreach (var item in data_lst)
                    {
                        if (item.Key != "Before")
                        {
                            ExcelRangeBase tar_rgn_offset = tar_rgn.Offset(r_offset * tbl.Rows.Count, 0);
                            ExcelRangeBase sum_rgn_offset = sum_rgn.Offset(r_offset, 0);
                            List<string> after_data = item.Value;
                            string[] sum_arr = new string[2];
                            for (int r_inx = 0; r_inx < before_data.Count; r_inx++)
                            {
                                int sum_inx = r_inx % 2;
                                double bef_val = Convert.ToDouble(before_data[r_inx]);
                                double aft_val = Convert.ToDouble(after_data[r_inx]);
                                double rate = Math.Abs(bef_val - aft_val) / bef_val;
                                tar_rgn_offset.Offset(r_inx, 4 * col_inx).Value = bef_val;// before_data[r_inx];
                                tar_rgn_offset.Offset(r_inx, 4 * col_inx + 1).Value = aft_val;// after_data[r_inx];
                                tar_rgn_offset.Offset(r_inx, 4 * col_inx).Style.Numberformat.Format = "#0.000";
                                tar_rgn_offset.Offset(r_inx, 4 * col_inx + 1).Style.Numberformat.Format = "#0.000";
                                if (rate < 0.1)
                                {
                                    sum_arr[sum_inx] = "OK";
                                }
                                else
                                {
                                    sum_arr[sum_inx] = "NG";
                                }
                                if (sum_inx == 1)
                                {
                                    if (TDMK_Code.check_exist_list_index("NG", sum_arr.ToList()) == -1)
                                    {
                                        sum_rgn_offset.Offset(r_inx / 2, col_inx).Value = "OK";
                                    }
                                    else
                                    {
                                        sum_rgn_offset.Offset(r_inx / 2, col_inx).Value = "NG";
                                    }
                                }
                            }
                            r_offset++;
                        }
                    }
                    col_inx++;
                }
            }
            else
            {
                MessageBox.Show("Process of ItemCode: " + ItemCode + process_name + " not found!", "Warning");
            }
            return sel_sheet;

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
        public string Export_Thermal_HeatSoak_Bend(ExcelWorkbook report_wrk,SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string sheet_name, int qty)
        {
            string sel_sheet = find_sheet(report_wrk, sheet_name);
            if (sel_sheet != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable _src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                List<string> pcs_sel_lst = Get_SelectedPcs_List(_src_tbl, "Sel_report", "Yes");//;//_src_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_report").ToUpper() == "YES").Select(x => x.Field<string>("Pcs_No")).Distinct().ToList();
                if (pcs_sel_lst.Count != 0)
                {
                    _src_tbl = _src_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_report").ToUpper() == "YES").CopyToDataTable();
                }
                SortedDictionary<int, string> sel_index_lst = Get_bending_Net(sqlcon_OK2SHIP, ItemCode, process_name);
                List<int> index_lst = sel_index_lst.Keys.ToList();
                if(_src_tbl.Rows.Count>0)
                {
                    DataTable src_tbl = _src_tbl.AsEnumerable().Where(x => index_lst.Contains(Convert.ToInt32(x.Field<string>("Net_no")) - 1)).CopyToDataTable();
                    List<DataTable> src_tbl_lst = new List<DataTable>();
                    Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                    ExcelWorksheet tar_wrksht = report_wrk.Worksheets[sel_sheet];
                    string Echeck_start_rgn = Excel_Lib.Find_Cell_Addr("Condition", "A10", tar_wrksht, false);// Get_start_range("Condition", "A10", tar_wrksht);
                    ExcelRangeBase Echeck_cycle_rgn = tar_wrksht.Cells[Echeck_start_rgn].Offset(Excel_Lib.Get_Cells_Info(tar_wrksht, tar_wrksht.Cells[Echeck_start_rgn]).row_qty,0);
                    Dictionary<string, string> Echeck_Cycle_addr = Get_Echeck_address(Echeck_cycle_rgn);
                    string start_rgn = Excel_Lib.Find_Cell_Addr("Test item", "A10", tar_wrksht, false);// Get_start_range("Test item", "A10", tar_wrksht);
                    ExcelRangeBase cycle_rgn = tar_wrksht.Cells[start_rgn].Offset(Excel_Lib.Get_Cells_Info(tar_wrksht, tar_wrksht.Cells[start_rgn]).row_qty,0);
                    Dictionary<string, string> bending_Cycle_addr = Get_Echeck_address(cycle_rgn);
                    int sel_qty = Math.Min(qty, src_tbl_lst.Count);
                    for (int col_inx = 0; col_inx < sel_qty; col_inx++)
                    {
                        DataTable tbl = src_tbl_lst[col_inx];
                        Dictionary<string, List<string>> data_lst = Get_List_Pair_data(tbl);
                        List<string> before_data = data_lst["Before"];
                        List<string> Echeck_data = data_lst["After"];
                        data_lst.Remove("Before");
                        foreach (var item in data_lst)
                        {
                            List<string> after_data = item.Value;
                            if (Echeck_Cycle_addr.Keys.ToList().IndexOf(item.Key) != -1)
                            {
                                string[] sum_arr = new string[2];
                                ExcelRangeBase sum_rgn_offset = tar_wrksht.Cells[Echeck_Cycle_addr[item.Key]].Offset(0, 2);
                                for (int r_inx = 0; r_inx < before_data.Count; r_inx++)
                                {
                                    int sum_inx = r_inx % 2;
                                    double bef_val = Convert.ToDouble(before_data[r_inx]);
                                    double aft_val = Convert.ToDouble(after_data[r_inx]);
                                    double rate = Math.Abs(bef_val - aft_val) / bef_val;
                                    if (rate < 0.1)
                                    {
                                        sum_arr[sum_inx] = "OK";
                                    }
                                    else
                                    {
                                        sum_arr[sum_inx] = "NG";
                                    }
                                    if (sum_inx == 1)
                                    {
                                        if (TDMK_Code.check_exist_list_index("NG", sum_arr.ToList()) == -1)
                                        {
                                            sum_rgn_offset.Offset(r_inx / 2, col_inx).Value = "OK";
                                        }
                                        else
                                        {
                                            sum_rgn_offset.Offset(r_inx / 2, col_inx).Value = "NG";
                                        }
                                    }
                                }
                            }
                            if (bending_Cycle_addr.Keys.ToList().IndexOf(item.Key) != -1)
                            {
                                ExcelRangeBase tar_rgn_offset = tar_wrksht.Cells[bending_Cycle_addr[item.Key]].Offset(0, 1);
                                for (int r_inx = 0; r_inx < before_data.Count; r_inx++)
                                {
                                    try
                                    {
                                        tar_rgn_offset.Offset(r_inx, 4 * col_inx).Value = Convert.ToDecimal(before_data[r_inx]);
                                        tar_rgn_offset.Offset(r_inx, 4 * col_inx).Style.Numberformat.Format = "0.000";
                                    }
                                    catch
                                    {
                                        tar_rgn_offset.Offset(r_inx, 4 * col_inx).Value = before_data[r_inx];
                                    }
                                    try
                                    {
                                        tar_rgn_offset.Offset(r_inx, 4 * col_inx + 1).Value = Convert.ToDecimal(after_data[r_inx]);
                                        tar_rgn_offset.Offset(r_inx, 4 * col_inx + 1).Style.Numberformat.Format = "0.000";
                                    }
                                    catch
                                    {
                                        tar_rgn_offset.Offset(r_inx, 4 * col_inx + 1).Value = after_data[r_inx];
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Process of ItemCode: " + ItemCode + process_name + ": no data!", "Warning");
                }
            }
            else
            {
                MessageBox.Show("Process of ItemCode: " + ItemCode + process_name + " not found!", "Warning");
            }
            return sel_sheet;
        }
        public Dictionary<string, string> Get_Echeck_address(ExcelRangeBase cycle_rgn)
        {
            Dictionary<string, string> bending_Cycle_addr = new Dictionary<string, string>();
            int row_inx = 0;
            while (myCode.checkDBNull(cycle_rgn.Offset(row_inx, 0).Value) != "")
            {
                string cycle = Get_Number_String( myCode.checkDBNull(cycle_rgn.Offset(row_inx, 0).Value), ' ');
                string cycle_name = "";
                if (cycle != "")
                {
                    cycle_name = "After_" + cycle + "_Bending";
                }
                else
                {
                    cycle_name = "After";
                }
                if ((cycle_name != "") && (bending_Cycle_addr.Keys.ToList().IndexOf(cycle_name) == -1))
                {
                    bending_Cycle_addr.Add(cycle_name, cycle_rgn.Offset(row_inx, 0).Address);
                }
                row_inx++;
            }
            return bending_Cycle_addr;
        }
        public SortedDictionary<int, string> Get_bending_Net(SqlConnection sqlcon, string ItemCode, string Process_name)
        {
            DataTable temp_spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, Process_name }));
            DataTable spec_dt = temp_spec_dt.AsDataView().ToTable(false, new string[] { "Point+V", "Point-V", "LSL", "USL" });
            SortedDictionary<int, string> dic_index_lst = new SortedDictionary<int, string>();
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
            return dic_index_lst;
        }
        public string Get_Number_String(string src_str, char split_chr)
        {
            string _result = "";
            string[] temp = src_str.Split(split_chr);
            foreach (var t in temp)
            {
                if (myCode.IsNumeric(t))
                {
                    _result = t;
                    break;
                }
            }
            return _result;
        }
        public string Export_Reflow_new(ExcelWorkbook report_wrk,SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string sheet_name)
        {
            string sel_sheet = find_sheet(report_wrk, sheet_name);
            if (sel_sheet != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                List<string> pcs_sel_lst = Get_SelectedPcs_List(src_tbl, "Sel_report", "Yes");// src_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_report").ToUpper() == "YES").Select(x => x.Field<string>("Pcs_No")).Distinct().ToList();
                if (pcs_sel_lst.Count != 0)
                {
                    src_tbl = src_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_report").ToUpper() == "YES").CopyToDataTable();
                }                
                string spec_filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process_name });
                DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", spec_filter_str);
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                ExcelWorksheet tar_wrksht = report_wrk.Worksheets[sel_sheet];
                string data_rgn_addr = Excel_Lib.Find_Cell_Addr("S/N", "A10", tar_wrksht, false);
                string sum_rgn_addr = Excel_Lib.Find_Cell_Addr("Unit S/N", "B10", tar_wrksht, false);
                int data_r_offset = Excel_Lib.Get_Cells_Info(tar_wrksht, tar_wrksht.Cells[data_rgn_addr]).row_qty;
                int sum_r_offset = Excel_Lib.Get_Cells_Info(tar_wrksht, tar_wrksht.Cells[sum_rgn_addr]).row_qty;
                ExcelRangeBase tar_rgn = tar_wrksht.Cells[data_rgn_addr].Offset(data_r_offset, 2);// tar_wrksht.Cells["C70"];
                ExcelRangeBase sum_rgn = tar_wrksht.Cells[sum_rgn_addr].Offset(sum_r_offset, 1);// tar_wrksht.Cells["C30"];
                int col_inx = 0;
                foreach (var tbl in src_tbl_lst)
                {
                    List<string> before_data = new List<string>();
                    List<string> after_data = new List<string>();
                    Get_Pair_data_new(spec_dt, tbl, ref before_data, ref after_data);
                    if (before_data.Count > 0)
                    {
                        sum_rgn.Offset(col_inx, 0).Value = Calcu_process.ConvertToDouble(before_data).Max();
                        sum_rgn.Offset(col_inx, 1).Value = Calcu_process.ConvertToDouble(after_data).Max();
                        for (int r_inx = 0; r_inx < before_data.Count; r_inx++)
                        {
                            tar_rgn.Offset(r_inx, 6 * col_inx).Value = Convert.ToDecimal(before_data[r_inx]);
                            tar_rgn.Offset(r_inx, 6 * col_inx + 1).Value = Convert.ToDecimal(after_data[r_inx]);
                            tar_rgn.Offset(r_inx, 6 * col_inx).Style.Numberformat.Format = "#0.000";
                            tar_rgn.Offset(r_inx, 6 * col_inx + 1).Style.Numberformat.Format = "#0.000";
                        }
                    }
                    col_inx++;
                }
            }
            else
            {
                MessageBox.Show("Process of ItemCode: " + ItemCode + process_name + " not found!", "Warning");
            }
            return sel_sheet;
        }
        public void Get_Pair_data_new(DataTable spec_dt, DataTable src_tbl, ref List<string> before_data, ref List<string> after_data)
        {
            List<string> _before_data = new List<string>();
            List<string> _after_data = new List<string>();
            Dictionary<int, string> sel_net = new Dictionary<int, string>();
            int index = 0;
            foreach (DataRow dr in spec_dt.Rows)
            {
                if (myCode.checkDBNull(dr["Sel_Report"]) == "Yes")
                {
                    sel_net.Add(index, myCode.checkDBNull(dr["Net_Name"]));
                }
                index++;
            }
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
            if (data_col.Count >= 2)
            {
                _before_data = data_col.ElementAt(0).Value;
                _after_data = data_col.ElementAt(data_col.Count - 1).Value;
            }
            foreach (var t in sel_net.Keys.ToList())
            {
                before_data.Add(_before_data[t]);
                after_data.Add(_after_data[t]);
            }
        }
        public string Export_ThermalCycling(ExcelWorkbook report_wrk,SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string sheet_name)
        {
            string sel_sheet = find_sheet(report_wrk, sheet_name);
            if (sel_sheet != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                List<string> pcs_sel_lst = Get_SelectedPcs_List(src_tbl, "Sel_report", "Yes");//;// src_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_report").ToUpper() == "YES").Select(x => x.Field<string>("Pcs_No")).Distinct().ToList();
                if (pcs_sel_lst.Count != 0)
                {
                    src_tbl = src_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_report").ToUpper() == "YES").CopyToDataTable();
                }
                DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process_name }));
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                ExcelWorksheet tar_wrksht = report_wrk.Worksheets[sel_sheet];

                string data_rgn_addr = Excel_Lib.Find_Cell_Addr("Sample 1", "C10", tar_wrksht,false);

                ExcelRangeBase sum_rgn = tar_wrksht.Cells[data_rgn_addr].Offset(1,0) ;// tar_wrksht.Cells["C17"];
                int col_inx = 0;
                foreach (var tbl in src_tbl_lst)
                {
                    Dictionary<string, List<string>> data_lst = Get_List_Pair_data(tbl);
                    int item_inx = 0;
                    foreach (var item in data_lst)
                    {
                        bool data_OK = false;
                        for (int i = 0; i < item.Value.Count; i++)
                        {
                            string USL = myCode.checkDBNull(spec_dt.Rows[i]["USL"]);
                            string LSL = myCode.checkDBNull(spec_dt.Rows[i]["LSL"]);
                            string cur_val = item.Value[i];
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
                            sum_rgn.Offset(item_inx, col_inx).Value = "OK";
                        }
                        else
                        {
                            sum_rgn.Offset(item_inx, col_inx).Value = "NG";
                        }
                        item_inx++;
                    }
                    col_inx++;
                }
            }
            else
            {
                MessageBox.Show("Process of ItemCode: " + ItemCode + process_name + " not found!", "Warning");
            }
            return sel_sheet;

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
        public string Export_Survival(ExcelWorkbook report_wrk,SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string sheet_name)
        {
            string sel_sheet = find_sheet(report_wrk, sheet_name);
            if (sel_sheet != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                List<string> pcs_sel_lst = Get_SelectedPcs_List(src_tbl, "Sel_report", "Yes");// src_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_report").ToUpper() == "YES").Select(x => x.Field<string>("Pcs_No")).Distinct().ToList();
                if (pcs_sel_lst.Count != 0)
                {
                    src_tbl = src_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_report").ToUpper() == "YES").CopyToDataTable();
                }
                string spec_filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process_name });
                DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", spec_filter_str);
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                ExcelWorksheet tar_wrksht = report_wrk.Worksheets[sel_sheet];
                string data_rgn_addr = Excel_Lib.Find_Cell_Addr("Before", "C10", tar_wrksht,false);
                int data_r_offset = Excel_Lib.Get_Cells_Info(tar_wrksht, tar_wrksht.Cells[data_rgn_addr]).row_qty;
                ExcelRangeBase tar_rgn = tar_wrksht.Cells[data_rgn_addr].Offset(data_r_offset,0) ;// tar_wrksht.Cells["C16"];
                int off_set = 0;
                foreach (var tbl in src_tbl_lst)
                {
                    Dictionary<string, List<string>> result_data = Get_List_Pair_data_spec(spec_dt, tbl);
                    int col_inx = 0;
                    foreach (var item_col in result_data.Keys.ToList())
                    {
                        List<string> cur_data = result_data[item_col];
                        for (int r_inx = 0; r_inx < cur_data.Count; r_inx++)
                        {
                            //tar_rgn.Offset(off_set * cur_data.Count + r_inx, col_inx).Value = 
                            try
                            {
                                tar_rgn.Offset(off_set * cur_data.Count + r_inx, col_inx).Value = Convert.ToDecimal(cur_data[r_inx]);
                                tar_rgn.Offset(off_set * cur_data.Count + r_inx, col_inx).Style.Numberformat.Format = "0.000";
                            }
                            catch
                            {
                                tar_rgn.Offset(off_set * cur_data.Count + r_inx, col_inx).Value = cur_data[r_inx];
                            }

                        }
                        col_inx++;
                    }
                    off_set++;
                }
            }
            else
            {
                MessageBox.Show("Process of ItemCode: " + ItemCode + process_name + " not found!", "Warning");
            }
            return sel_sheet;
        }
        public Dictionary<string, List<string>> Get_List_Pair_data_spec(DataTable spec_dt, DataTable src_tbl)
        {
            Dictionary<int, string> sel_net = new Dictionary<int, string>();
            int index = 0;
            foreach (DataRow dr in spec_dt.Rows)
            {
                if (myCode.checkDBNull(dr["Sel_Report"]) == "Yes")
                {
                    sel_net.Add(index, myCode.checkDBNull(dr["Net_Name"]));
                }
                index++;
            }
            Dictionary<string, List<string>> data_col = new Dictionary<string, List<string>>();
            foreach (DataColumn dc in src_tbl.Columns)
            {
                List<string> _col_data = new List<string>();
                List<string> col_data = new List<string>();
                if ((dc.ColumnName == "Before") || (dc.ColumnName.Contains("After")))
                {
                    foreach (DataRow dr in src_tbl.Rows)
                    {
                        string cell_val = myCode.checkDBNull(dr[dc]);
                        if (cell_val != "")
                        {
                            _col_data.Add(dr[dc].ToString());
                        }
                    }
                }
                if (_col_data.Count > 0)
                {
                    foreach (var t in sel_net.Keys.ToList())
                    {
                        col_data.Add(_col_data[t]);
                    }
                    data_col.Add(dc.ColumnName, col_data);
                }
            }
            return data_col;
        }
        public string Export_Surv_HotOil_Multi(ExcelWorkbook report_wrk,SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string sheet_name)
        {
            //ExcelWorkbook report_wrk = TDMK_Code.Create_Report_new(Application.StartupPath, format_name, "", ItemCode + "-" + LotNo);
            string sel_sheet = find_sheet(report_wrk, sheet_name);
            if (sel_sheet != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                List<string> pcs_sel_lst = Get_SelectedPcs_List(src_tbl, "Sel_report", "Yes");// src_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_report").ToUpper() == "YES").Select(x => x.Field<string>("Pcs_No")).Distinct().ToList();
                if (pcs_sel_lst.Count != 0)
                {
                    src_tbl = src_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_report").ToUpper() == "YES").CopyToDataTable();
                }
                if (src_tbl.Rows.Count > 0)
                {
                    ExcelWorksheet tar_wrksht = report_wrk.Worksheets[sel_sheet];
                    ExcelRangeBase start_tar_rgn = tar_wrksht.Cells["A15"];
                    List<DataTable> _src_tbl_lst = new List<DataTable>();
                    Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Logfile" }, ref _src_tbl_lst, "Before");
                    //int inx = 0;
                    foreach (var temp_src_tbl in _src_tbl_lst)
                    {
                        string logfile = temp_src_tbl.Rows[0]["Logfile"].ToString();
                        string title_rgn = "";
                        if (logfile.Contains("BVH"))
                        {
                            title_rgn = Excel_Lib.Find_Cell_Addr("Resistance for BVH", start_tar_rgn.Address, tar_wrksht, false);// Get_start_range("Resistance for BVH", start_tar_rgn.Address, tar_wrksht);
                        }
                        else
                        {
                            title_rgn = Excel_Lib.Find_Cell_Addr("Resistance for PTH", start_tar_rgn.Address, tar_wrksht, false);// Get_start_range("Resistance for PTH", start_tar_rgn.Address, tar_wrksht);
                        }
                        if (title_rgn != "")
                        {
                            string data_rgn = Excel_Lib.Find_Cell_Addr("Sample No", title_rgn, tar_wrksht, false);// Get_start_range("Sample No", title_rgn, tar_wrksht);
                            int qty_off = Excel_Lib.Get_Cells_Info(tar_wrksht, tar_wrksht.Cells[data_rgn]).row_qty;
                            ExcelRangeBase tar_rgn = tar_wrksht.Cells[data_rgn].Offset(qty_off, 1);
                            List<DataTable> src_tbl_lst = new List<DataTable>();
                            Get_ListTable(-1, temp_src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                            int off_set = 0;
                            foreach (var tbl in src_tbl_lst)
                            {
                                Dictionary<string, List<string>> result_data = Get_List_Pair_data(tbl);
                                int col_inx = 0;
                                foreach (var item_col in result_data.Keys.ToList())
                                {
                                    List<string> cur_data = result_data[item_col];
                                    for (int r_inx = 0; r_inx < cur_data.Count; r_inx++)
                                    {
                                        //tar_rgn.Offset(off_set * cur_data.Count + r_inx, col_inx).Value = Convert.ToDecimal(cur_data[r_inx]);
                                        try
                                        {
                                            tar_rgn.Offset(off_set * cur_data.Count + r_inx, col_inx).Value = Convert.ToDecimal(cur_data[r_inx]);
                                            tar_rgn.Offset(off_set * cur_data.Count + r_inx, col_inx).Style.Numberformat.Format = "0.000";
                                        }
                                        catch
                                        {
                                            tar_rgn.Offset(off_set * cur_data.Count + r_inx, col_inx).Value = cur_data[r_inx];
                                        }

                                    }
                                    col_inx++;
                                }
                                off_set++;
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Process of ItemCode: " + ItemCode + process_name + ": no data!", "Warning");
                }
            }
            else
            {
                MessageBox.Show("Process of ItemCode: " + ItemCode + process_name + " not found!", "Warning");
            }
            return sel_sheet;

        }
        public string Export_Electrical(ExcelWorkbook report_wrk,SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string sheet_name)
        {
            string sel_sheet = find_sheet(report_wrk, sheet_name);
            if (sel_sheet != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                List<string> pcs_sel_lst = Get_SelectedPcs_List(src_tbl, "Sel_report", "Yes");// src_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_report").ToUpper() == "YES").Select(x => x.Field<string>("Pcs_No")).Distinct().ToList();
                if (pcs_sel_lst.Count != 0)
                {
                    src_tbl = src_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_report").ToUpper() == "YES").CopyToDataTable();
                }
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Data");
                ExcelWorksheet tar_wrksht = report_wrk.Worksheets[sel_sheet];
                string data_rgn_addr = Excel_Lib.Find_Cell_Addr("Data No", "G10", tar_wrksht,false);
                int data_r_offset = Excel_Lib.Get_Cells_Info(tar_wrksht, tar_wrksht.Cells[data_rgn_addr]).row_qty;
                ExcelRangeBase tar_rgn = tar_wrksht.Cells[data_rgn_addr].Offset(data_r_offset+1,0) ;// tar_wrksht.Cells["G35"];
                int col_inx = 0;
                foreach (var tbl in src_tbl_lst)
                {
                    for (int r_inx = 0; r_inx < tbl.Rows.Count; r_inx++)
                    {
                        //tar_rgn.Offset(r_inx, col_inx).Value = Convert.ToDecimal(tbl.Rows[r_inx]["Data"]);
                        if(myCode.IsNumeric(myCode.checkDBNull(tbl.Rows[r_inx]["Data"].ToString())))
                        {
                            tar_rgn.Offset(r_inx, col_inx).Value = Convert.ToDecimal(tbl.Rows[r_inx]["Data"]);
                            tar_rgn.Offset(r_inx, col_inx).Style.Numberformat.Format = "0.000";
                        }
                        else
                        {
                            tar_rgn.Offset(r_inx, col_inx).Value = tbl.Rows[r_inx]["Data"];
                        }    
                    }
                    col_inx++;
                }
                
            }
            else
            {
                MessageBox.Show("Process of ItemCode: " + ItemCode + process_name + " not found!", "Warning");
                
            }
            return sel_sheet;

        }
        public string Export_ThermalCycling_On_Coupon(ExcelWorkbook report_wrk,SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string sheet_name)
        {
            string sel_sheet = find_sheet(report_wrk, sheet_name);
            if (sel_sheet != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                List<string> pcs_sel_lst = Get_SelectedPcs_List(src_tbl, "Sel_report", "Yes");// src_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_report").ToUpper() == "YES").Select(x => x.Field<string>("Pcs_No")).Distinct().ToList();
                if (pcs_sel_lst.Count != 0)
                {
                    src_tbl = src_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_report").ToUpper() == "YES").CopyToDataTable();
                }
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                ExcelWorksheet tar_wrksht = report_wrk.Worksheets[sel_sheet];
                string data_rgn_addr = Excel_Lib.Find_Cell_Addr("Coupon 1", "C10", tar_wrksht,false);
                ExcelRangeBase sum_rgn = tar_wrksht.Cells[data_rgn_addr].Offset(1, 0);// tar_wrksht.Cells["C17"];
                int col_inx = 0;
                foreach (var tbl in src_tbl_lst)
                {
                    Dictionary<string, List<string>> data_lst = Get_List_Pair_data(tbl);
                    int item_inx = 0;
                    foreach (var item in data_lst)
                    {
                        bool data_OK = false;
                        for (int i = 0; i < item.Value.Count; i++)
                        {
                            string cur_val = item.Value[i];
                            if (myCode.IsNumeric(cur_val))
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
                            sum_rgn.Offset(item_inx, col_inx).Value = "OK";
                        }
                        else
                        {
                            sum_rgn.Offset(item_inx, col_inx).Value = "NG";
                        }
                        item_inx++;
                    }
                    col_inx++;
                    if (col_inx == 10)
                    {
                        break;
                    }
                }
            }
            else
            {                
                MessageBox.Show("Process of ItemCode: " + ItemCode + process_name + " not found!", "Warning");
            }
            return sel_sheet;
        }
        public List<string> Get_SelectedPcs_List(DataTable src_dt, string sel_col_name, string filter_col_val)
        {
            List<string> result = new List<string>();
            string filter_str = TDMK_Code.filter_str(new string[] { sel_col_name }, new string[] { filter_col_val });
            DataView dv = src_dt.AsDataView();
            dv.RowFilter = filter_str;
            if (dv.Count != 0)
            {
                result = dv.ToTable().AsEnumerable().Select(x => x.Field<string>("Pcs_No")).Distinct().ToList();
            }
            return result;
        }
    }
}
