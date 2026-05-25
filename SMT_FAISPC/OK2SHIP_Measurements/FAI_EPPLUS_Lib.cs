using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TDMK_EPPLUS_7;
using TDMK_SEEV_DLL;
using TDMK_SQL;
using OK2SHIP_Lib;
using OfficeOpenXml;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace FAI_Export_EPPLUS
{
    public class FAI_EPPLUS_Lib
    {
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SEI_Lib myCode = new SEI_Lib();
        TDMK_EPPLUS7_lib Excel_Lib = new TDMK_EPPLUS7_lib();
        TDMK_OK2SHIP TDMK_OK2SHIP = new TDMK_OK2SHIP();
        public Dictionary<string, DataTable> Export_FAI_Batch(SqlConnection sqlcon, string tar_ItemCode, string tar_LotNo, string format_type, string shift)
        {
            Dictionary<string, DataTable> dt_lst = new Dictionary<string, DataTable>();
            char[] trim_char = new char[] { ' ', '\r', '\n' };
            DataTable FAI_Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Remark","Shift" }, new string[] { tar_ItemCode, tar_LotNo, format_type, shift }));
            DataTable FAI_Spec_tbl = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Spec", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { tar_ItemCode, format_type }));
            string[] sheetno = FAI_Spec_tbl.AsEnumerable().Select(r => r.Field<string>("SheetNo")).Distinct().ToArray();
            foreach (string sht in sheetno)
            {
                DataTable dt = new DataTable();
                Dictionary<string, List<string>> dic_FAI = new Dictionary<string, List<string>>();
                string[] FAI_No = FAI_Spec_tbl.AsEnumerable().Where(r => r.Field<string>("SheetNo") == sht).Select(r => r.Field<string>("FAI_No")).ToArray();
                if (sht.Contains("SPC"))
                {
                    int inx = 0;
                    foreach (string t in FAI_No)
                    {
                        string[] temp = t.Split('_');
                        string act_val = temp[1];
                        string[] temp2 = temp[0].Split('/');
                        FAI_No[inx] = temp2[0] + "_" + act_val;
                        inx++;
                    }
                }
                foreach (string fai_no in FAI_No)
                {
                    List<string> fai_val = FAI_Data_tbl.AsEnumerable().Where(r => r.Field<string>("FAI_No") == fai_no).Select(r => r.Field<string>("FAI_Data")).ToList();
                    if (dic_FAI.Keys.ToList().IndexOf(fai_no) == -1)
                    {
                        dic_FAI.Add(fai_no, fai_val);
                    }
                }
                foreach (var fai in dic_FAI)
                {

                    if (!myCode.check_columns_existed(dt, fai.Key))
                    {
                        if (fai.Value.Count != 0)
                        {
                            dt.Columns.Add(fai.Key, typeof(double));
                        }
                        else
                        {
                            dt.Columns.Add(fai.Key);
                        }
                    }
                    if (dt.Rows.Count < fai.Value.Count)
                    {
                        int r = fai.Value.Count - dt.Rows.Count;
                        for (int i = 0; i < r; i++)
                        {
                            dt.Rows.Add();
                        }
                    }
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        if (fai.Value.Count != 0)
                        {
                            if (j < fai.Value.Count)
                            {
                                dt.Rows[j][fai.Key] = fai.Value[j];
                            }
                        }
                        else
                        {
                            dt.Rows[j][fai.Key] = "N/A";
                        }
                    }

                }
                if (dt.Rows.Count > 0)
                {
                    dt_lst.Add(sht, dt);
                }
            }
            return dt_lst;
        }
        public void Export_To_FAI(SqlConnection sqlcon, ExcelWorkbook src_format_wrk, string _ItemCode, string _LotNo, Dictionary<string, DataTable> dic_data, string format_type)
        {
            DataTable FAI_Spec_tbl = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Spec", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { _ItemCode, format_type }));
            List<string> FAI_keys_lst = new List<string> { "FAI", "SPC", "CPK", "parentheses" };
            foreach (var sht in dic_data)
            {
                try
                {
                    if (FAI_keys_lst.Any(x => sht.Key.Contains(x)))
                    {
                        string sel_sht_name = sht.Key;
                        ExcelWorksheet format_wrksheet = src_format_wrk.Worksheets[sel_sht_name];
                        string dim_no_addr = Excel_Lib.Find_Cell_Addr("Dim. No.", "B10", format_wrksheet, false);
                        string instrument_addr = Excel_Lib.Find_Cell_Addr("instrument", "B10", format_wrksheet, false);
                        string FAI_data_addr = Find_Offset(format_wrksheet.Cells[dim_no_addr].Offset(0, 1).Address, format_wrksheet, true, "");
                        int off_set = format_wrksheet.Cells[FAI_data_addr].Start.Column - format_wrksheet.Cells[instrument_addr].Start.Column;
                        Dictionary<string, int> fai_loc = Find_FAI_addr_qty(dim_no_addr, format_wrksheet, false);
                        if (fai_loc.Count > 0)
                        {
                            string _data_addr = fai_loc.Keys.ToList()[0];
                            ExcelRangeBase format_rgn = format_wrksheet.Cells[_data_addr].Offset(0, off_set);
                            format_rgn.LoadFromDataTable(sht.Value);
                            int total_col_num = sht.Value.Columns.Count - 1;
                            int r_off = sht.Value.Rows.Count - 1;
                            int start_rgn_row = format_rgn.Start.Row;
                            int start_rgn_col = format_rgn.Start.Column;
                            ExcelRangeBase final_rgn = format_wrksheet.Cells[start_rgn_row, start_rgn_col, start_rgn_row + r_off, total_col_num];
                            final_rgn.Style.Numberformat.Format = "#0.000";
                        }
                        //if (format_type == "MASS")
                        //{
                        //    string itemcode_loc = Excel_Lib.Find_Cell_Addr("ITEM CODE:", "B1", format_wrksheet, false);
                        //    string lotno_loc = Excel_Lib.Find_Cell_Addr("LOT:", "B1", format_wrksheet, false);
                        //    string itemname_loc = Excel_Lib.Find_Cell_Addr("ITEM NAME:", "B1", format_wrksheet, false);
                        //    string date_loc = Excel_Lib.Find_Cell_Addr("NGÀY:", "B1", format_wrksheet, false);

                        //    string format_name = Path.GetFileNameWithoutExtension(src_format_wrk);
                        //    string[] temp = format_name.Split('-');
                        //    string item_name = "";
                        //    if (temp.Length > 1)
                        //    {
                        //        item_name = temp[1];
                        //    }
                        //    format_wrksheet.Range[itemname_loc].Offset[0, 1].Value = item_name;
                        //    format_wrksheet.Range[itemcode_loc].Offset[0, 1].Value = _ItemCode;
                        //    format_wrksheet.Range[lotno_loc].Offset[0, 1].Value = _LotNo;
                        //    format_wrksheet.Range[date_loc].Offset[0, 1].Value = DateTime.Now.ToShortDateString();
                        //}
                    }
                }
                catch
                {
                    continue;
                }

            }
        }
        public void Export_To_FAI(SqlConnection sqlcon, ExcelPackage src_pack, string _ItemCode, string _LotNo, Dictionary<string, DataTable> dic_data, string format_type)
        {
            ExcelWorkbook src_format_wrk = src_pack.Workbook;
            DataTable FAI_Spec_tbl = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Spec", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { _ItemCode, format_type }));
            List<string> FAI_keys_lst = new List<string> { "FAI", "SPC", "CPK", "parentheses" };
            foreach (var sht in dic_data)
            {
                try
                {
                    if (FAI_keys_lst.Any(x => sht.Key.Contains(x)))
                    {
                        string sel_sht_name = sht.Key;
                        ExcelWorksheet format_wrksheet = src_format_wrk.Worksheets[sel_sht_name];
                        string dim_no_addr = Excel_Lib.Find_Cell_Addr("Dim. No.", "B10", format_wrksheet, false);
                        string instrument_addr = Excel_Lib.Find_Cell_Addr("instrument", "B10", format_wrksheet, false);
                        string FAI_data_addr = Find_Offset(format_wrksheet.Cells[dim_no_addr].Offset(0, 1).Address, format_wrksheet, true, "");
                        int off_set = format_wrksheet.Cells[FAI_data_addr].Start.Column - format_wrksheet.Cells[instrument_addr].Start.Column;
                        Dictionary<string, int> fai_loc = Find_FAI_addr_qty(dim_no_addr, format_wrksheet, false);
                        if (fai_loc.Count > 0)
                        {
                            string _data_addr = fai_loc.Keys.ToList()[0];
                            ExcelRangeBase format_rgn = format_wrksheet.Cells[_data_addr].Offset(0, off_set);
                            format_rgn.LoadFromDataTable(sht.Value);
                            int total_col_num = sht.Value.Columns.Count - 1;
                            int r_off = sht.Value.Rows.Count - 1;
                            int start_rgn_row = format_rgn.Start.Row;
                            int start_rgn_col = format_rgn.Start.Column;
                            ExcelRangeBase final_rgn = format_wrksheet.Cells[start_rgn_row, start_rgn_col, start_rgn_row + r_off, total_col_num];
                            final_rgn.Style.Numberformat.Format = "#0.000";
                        }
                        if (format_type == "MASS")
                        {
                            string itemcode_loc = Excel_Lib.Find_Cell_Addr("ITEM CODE:", "B1", format_wrksheet, false);
                            string lotno_loc = Excel_Lib.Find_Cell_Addr("LOT:", "B1", format_wrksheet, false);
                            string itemname_loc = Excel_Lib.Find_Cell_Addr("ITEM NAME:", "B1", format_wrksheet, false);
                            string date_loc = Excel_Lib.Find_Cell_Addr("NGÀY:", "B1", format_wrksheet, false);

                            string format_name = Path.GetFileNameWithoutExtension(src_pack.File.Name);
                            string[] temp = format_name.Split('-');
                            string item_name = "";
                            if (temp.Length > 1)
                            {
                                item_name = temp[1];
                            }
                            format_wrksheet.Cells[itemname_loc].Offset(0, 1).Value = item_name;
                            format_wrksheet.Cells[itemcode_loc].Offset(0, 1).Value = _ItemCode;
                            format_wrksheet.Cells[lotno_loc].Offset(0, 1).Value = _LotNo;
                            format_wrksheet.Cells[date_loc].Offset(0, 1).Value = DateTime.Now.ToShortDateString();
                        }
                    }
                }
                catch
                {
                    continue;
                }

            }

            /***********/
            List<string> FAI_keys_lst_new = new List<string> { "FAI", "SPC" };
            if (format_type == "NPI")
            {
                int error_count = 0;
                while (src_format_wrk.Worksheets.Any(x=> !FAI_keys_lst_new.Any(y=> x.Name.Contains(y))) && error_count <5)
                {
                    try
                    {
                        foreach (ExcelWorksheet sht in src_format_wrk.Worksheets)
                        {
                            if (sht!=null)
                            {
                                if (!FAI_keys_lst_new.Any(x => sht.Name.Contains(x)))
                                {
                                    src_format_wrk.Worksheets.Delete(sht);
                                }
                            }    
                        }
                    }
                    catch
                    {
                        error_count++;
                    }

                }
            }
            /***********/
 
        }
        public Dictionary<string, int> Find_FAI_addr_qty(string start_addr, ExcelWorksheet tar_wrksht, bool left_to_right)
        {
            Dictionary<string, int> result = new Dictionary<string, int>();
            ExcelRangeBase cycle_rgn = tar_wrksht.Cells[start_addr];
            int offset = 0;
            int empty_count = 0;
            string data_addr = "";
            int qty = 0;
            while (true)
            {
                ExcelRangeBase sel_rgn = cycle_rgn.Offset(offset, 0);
                if (left_to_right)
                {
                    sel_rgn = cycle_rgn.Offset(0, offset);
                }
                string sel_rgn_val = myCode.checkDBNull(sel_rgn.Value);
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
                    if (myCode.IsNumeric(sel_rgn_val))
                    {
                        if (data_addr == "")
                        {
                            data_addr = sel_rgn.Address;
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
        public string Find_Offset(string start_addr, ExcelWorksheet tar_wrksht, bool left_to_right, string search_key = "")
        {
            string result = "";
            ExcelRangeBase cycle_rgn = tar_wrksht.Cells[start_addr];
            int c_offset = 0;
            int empty_count = 0;
            while (true)
            {
                ExcelRangeBase sel_rgn = cycle_rgn.Offset(c_offset, 0);
                if (left_to_right)
                {
                    sel_rgn = cycle_rgn.Offset(0, c_offset);
                }
                string sel_rgn_val = myCode.checkDBNull(sel_rgn.Value);
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
                        result = sel_rgn.Address;
                        break;
                    }
                }
                else
                {
                    if (empty_count == 0)
                    {
                        result = sel_rgn.Address;
                        break;
                    }
                }
                c_offset++;
            }
            return result;
        }
        public bool Check_FAI_data(string tar_ItemCode, string tar_LotNo,string shift, SqlConnection tar_sqlcon, double tar_CPK)
        {
            char[] split_char = new char[] { '/', '_' };
            DataTable spec_dt = Load_FAI_Spec_ToTable(tar_sqlcon, tar_ItemCode, "NPI");
            DataTable FAI_tbl = TDMK_OK2SHIP.Load_FAI_ToTable(tar_sqlcon, tar_ItemCode, tar_LotNo, "NPI", shift);
            DataTable CPK_tbl = Load_CPK_Table(FAI_tbl, spec_dt);
            DataTable FAI_Spec_tbl = TDMK_Code.Datatable_Filter(tar_sqlcon, "FAI_Spec", "ItemCode = '" + tar_ItemCode + "'");
            string[] sheetno = FAI_Spec_tbl.AsEnumerable().Select(r => r.Field<string>("SheetNo")).Distinct().ToArray();
            bool sum_en = false;
            bool result = false;
            if (FAI_tbl.Rows.Count > 0)
            {
                sum_en = true;
            }
            else
            {
                sum_en = false;
            }
            foreach (string sht in sheetno)
            {
                if (sum_en)
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
                    foreach (string t in FAI_NO_lst)
                    {
                        bool cpk_ok = false;
                        bool fai_ok = false;
                        string sv = spec_dt.Rows[0][t].ToString();
                        string TolMax = spec_dt.Rows[1][t].ToString();
                        string TolMin = spec_dt.Rows[2][t].ToString();
                        string UL = "";
                        string LL = "";
                        if (myCode.IsNumeric(TolMax))
                        {
                            UL = TolMax;// (Convert.ToDouble(sv) + Convert.ToDouble(TolMax)).ToString();
                        }
                        if (myCode.IsNumeric(TolMin))
                        {
                            LL = TolMin;// (Convert.ToDouble(sv) - Convert.ToDouble(TolMin)).ToString();
                        }
                        foreach (DataRow dr in FAI_tbl.Rows)
                        {
                            string act_val = dr[t].ToString();
                            if (myCode.check_in_limit(UL, LL, act_val, sv))
                            {
                                fai_ok = true;
                            }
                            else
                            {
                                fai_ok = false;
                                break;
                            }
                        }
                        if (myCode.check_columns_existed(CPK_tbl, t))
                        {
                            if (Convert.ToDouble(CPK_tbl.Rows[7][t]) > tar_CPK)
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
                        if (fai_ok && cpk_ok)
                        {
                            result = true;
                        }
                        else
                        {
                            result = false;
                            break;
                        }
                    }
                    if (!result)
                    {
                        break;
                    }
                }
            }
            return result;
        }
        public DataTable Load_CPK_Table(DataTable src_FAI_Data, DataTable spec_dt)
        {
            DataTable test_tbl = src_FAI_Data;// DGV_To_Table(DGV_DataView);
            int arr_num = test_tbl.Columns.Count;// col_list.Count;
            SEI_Lib.Calculate_CPK[] myCalc_CPK = new SEI_Lib.Calculate_CPK[arr_num];
            SEI_Lib.FAI_Histogram_Data[] myHistogram_data = new SEI_Lib.FAI_Histogram_Data[arr_num];
            int col_list_inx = 0;
            foreach (DataColumn tbl_col in test_tbl.Columns)
            {
                SEI_Lib.FAI_Spec sel_FAI_test;
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
                        sel_FAI_test = new SEI_Lib.FAI_Spec(t_FAIName, t_checkside, t_FAI_Setval, t_FAI_UL, t_FAI_LL, t_FAI_sheetno, t_FAI_instrument);
                        string[] temp_FAI_data = test_tbl.AsEnumerable().Select(r => r.Field<string>(t)).ToArray();
                        Double[] data_arr = new double[temp_FAI_data.Length];
                        int inx = 0;
                        foreach (string c in temp_FAI_data)
                        {
                            if ((myCode.checkDBNull(c) != "") && myCode.IsNumeric(c))
                            {
                                data_arr[inx] = Convert.ToDouble(c);
                                inx++;
                            }
                        }
                        Array.Resize(ref data_arr, inx);
                        if (inx > 0)
                        {
                            myCode.Calcul_CPK_FAI2(sel_FAI_test, data_arr, ref myCalc_CPK[col_list_inx], ref myHistogram_data[col_list_inx]);
                        }
                        break;
                    }
                }
                col_list_inx++;
            }
            DataTable tbl_CPK = new DataTable();
            foreach (SEI_Lib.Calculate_CPK item_CPK in myCalc_CPK)
            {
                string CPK_col_name = item_CPK.FAI_No;
                if (CPK_col_name != null)
                {
                    if (!myCode.check_columns_existed(tbl_CPK, CPK_col_name))
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
                    tbl_CPK.Rows[8][CPK_col_name] = Math.Round(item_CPK.CPKM, 3).ToString();
                }
            }
            return tbl_CPK;
        }
        public DataTable Load_FAI_Spec_ToTable(SqlConnection sqlcon, string tar_ItemCode, string format_type)
        {
            DataTable cur_dt = new DataTable();
            DataTable FAI_spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Spec", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { tar_ItemCode, format_type }));
            List<string> search_key_lst = new List<string>() { "FAI", "Dimension" };
            List<string> FAI_SheetNo_list = FAI_spec_dt.AsEnumerable().Where(x => search_key_lst.Any(y => x.Field<string>("SheetNo").Contains(y))).Select(x => x.Field<string>("SheetNo")).Distinct().ToList();
            if (FAI_SheetNo_list.Count > 0)
            {
                List<DataTable> dt_lst = new List<DataTable>();
                TDMK_OK2SHIP.Get_ListTable(-1, FAI_spec_dt, new string[] { "SheetNo", "FAI_No" }, ref dt_lst, "Distribution");
                foreach (DataTable dt in dt_lst)
                {
                    if (dt.Rows.Count > 0)
                    {
                        string sheetno = TDMK_OK2SHIP.checkDBNull(dt.Rows[0]["SheetNo"]);
                        if (FAI_SheetNo_list.IndexOf(sheetno) != -1)
                        {
                            string fai_no = TDMK_OK2SHIP.checkDBNull(dt.Rows[0]["FAI_No"]);
                            if (!TDMK_OK2SHIP.check_columns_existed(cur_dt, fai_no))
                            {
                                cur_dt.Columns.Add(fai_no);
                            }
                            if (cur_dt.Rows.Count == 0)
                            {
                                for (int i = 0; i < 6; i++)
                                {
                                    cur_dt.Rows.Add();
                                }
                            }
                            cur_dt.Rows[0][fai_no] = dt.Rows[0]["NormDim"];
                            cur_dt.Rows[1][fai_no] = dt.Rows[0]["TolMax"];
                            cur_dt.Rows[2][fai_no] = dt.Rows[0]["TolMin"];
                            cur_dt.Rows[3][fai_no] = dt.Rows[0]["Distribution"];
                            cur_dt.Rows[4][fai_no] = dt.Rows[0]["Instrument"];
                            cur_dt.Rows[5][fai_no] = dt.Rows[0]["SheetNo"];
                        }
                    }
                }
            }
            return cur_dt;
        }
        public void Export_FAI_Process(SqlConnection sqlcon, ExcelWorkbook report_saved, string tar_ItemCode, string tar_LotNo, string format_type,string shift, double tar_CPK)
        {
            if (Check_FAI_data(tar_ItemCode, tar_LotNo, shift,sqlcon, tar_CPK))
            {
                Dictionary<string, DataTable> dic_data = Export_FAI_Batch(sqlcon, tar_ItemCode, tar_LotNo, format_type, shift);
                Export_To_FAI(sqlcon, report_saved, tar_ItemCode, tar_LotNo, dic_data, format_type);
            }
        }
        public ExcelPackage open_excel(string file_name)
        {
            ExcelPackage myexcel = null;
            FileInfo excel_file = new FileInfo(file_name);
            if (File.Exists(file_name))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                myexcel = new ExcelPackage(excel_file);
            }
            return myexcel;
        }
        public Dictionary<string, DataTable> Export_FAI_Batch(SqlConnection sqlcon, ExcelWorkbook src_format_wrk, string tar_ItemCode, string tar_LotNo, string format_type, string shift="1")
        {
            Dictionary<string, DataTable> dt_lst = new Dictionary<string, DataTable>();
            char[] trim_char = new char[] { ' ', '\r', '\n' };
            DataTable FAI_Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Remark","Shift" }, new string[] { tar_ItemCode, tar_LotNo, format_type,shift }));
            List<string> sht_keys = new List<string> { "FAI", "SPC", "parentheses" };
            Dictionary<string, List<string>> sheet_FAI_dic = new Dictionary<string, List<string>>();
            foreach (ExcelWorksheet tg in src_format_wrk.Worksheets)
            {
                if (sht_keys.Any(x => tg.Name.Contains(x)))
                {
                    string dim_no_addr = Find_Cell_Addr("Dim. No.", "B10", tg, false);
                    string instrument_addr = Find_Cell_Addr("instrument", "B10", tg, false);
                    string FAI_data_addr = Find_Offset(tg.Cells[dim_no_addr].Offset(0, 1).Address, tg, true, "");
                    int off_set = tg.Cells[FAI_data_addr].End.Column - tg.Cells[instrument_addr].End.Column;
                    ExcelRangeBase sel_rgn = tg.Cells[dim_no_addr].Offset(0, off_set);// tg.Range["D19"];    //tg.Range["C17"];
                    sheet_FAI_dic.Add(tg.Name, new List<string>());
                    int sel_inx = 0;
                    while (myCode.checkDBNull(sel_rgn.Offset(0, sel_inx).Value) != "")
                    {
                        string t_checkside = myCode.checkDBNull(sel_rgn.Offset(-1, sel_inx).Value);
                        string t_FAIName = myCode.checkDBNull(sel_rgn.Offset(0, sel_inx).Value).Replace(" ", "");
                        string t_FAI_Setval = myCode.checkDBNull(sel_rgn.Offset(1, sel_inx).Value);
                        string col_name = t_FAIName.Split('/').FirstOrDefault() + "_" + t_FAI_Setval;
                        string t_FAI_sheetno = tg.Name;
                        sheet_FAI_dic[tg.Name].Add(col_name);
                        sel_inx++;
                    }
                }
            }
            foreach (var _sht in sheet_FAI_dic)
            {
                DataTable dt = new DataTable();
                Dictionary<string, List<string>> dic_FAI = new Dictionary<string, List<string>>();
                string[] FAI_No = _sht.Value.Distinct().ToArray();
                string sht = _sht.Key;
                int inx = 0;
                foreach (string t in FAI_No)
                {

                    FAI_No[inx] = t.Split(new char[] {'/','_' }).FirstOrDefault() + "_" + t.Split(new char[] { '/', '_' }).LastOrDefault();
                    inx++;
                }
                foreach (string fai_no in FAI_No)
                {
                    List<string> fai_val = FAI_Data_tbl.AsEnumerable().Where(r => r.Field<string>("FAI_No") == fai_no).Select(r => r.Field<string>("FAI_Data")).ToList();
                    if (dic_FAI.Keys.ToList().IndexOf(fai_no) == -1)
                    {
                        dic_FAI.Add(fai_no, fai_val);
                    }
                }
                foreach (var fai in dic_FAI)
                {

                    if (!myCode.check_columns_existed(dt, fai.Key))
                    {
                        if (fai.Value.Count != 0)
                        {
                            dt.Columns.Add(fai.Key, typeof(double));
                        }
                        else
                        {
                            dt.Columns.Add(fai.Key);
                        }
                    }
                    if (dt.Rows.Count < fai.Value.Count)
                    {
                        int r = fai.Value.Count - dt.Rows.Count;
                        for (int i = 0; i < r; i++)
                        {
                            dt.Rows.Add();
                        }
                    }
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        if (fai.Value.Count != 0)
                        {
                            if (j < fai.Value.Count)
                            {
                                dt.Rows[j][fai.Key] = fai.Value[j];
                            }
                        }
                        else
                        {
                            dt.Rows[j][fai.Key] = "N/A";
                        }
                    }

                }
                if (dt.Rows.Count > 0)
                {
                    dt_lst.Add(sht, dt);
                }
            }
            return dt_lst;
        }
        public string Find_Cell_Addr(string search_key, string start_addr, ExcelWorksheet tar_wrksht, bool left_to_right)
        {
            string result = "";

            ExcelRangeBase cycle_rgn = tar_wrksht.Cells[start_addr];

            for (int i = 0; i < 100; i++)
            {
                ExcelRangeBase sel_rgn = cycle_rgn.Offset(i, 0);
                if (left_to_right)
                {
                    sel_rgn = cycle_rgn.Offset(0, i);
                }
                string sel_rgn_val = myCode.checkDBNull(sel_rgn.Value);
                if (sel_rgn_val.Replace(" ", "").ToUpper() == search_key.Replace(" ", "").ToUpper())
                {
                    result = sel_rgn.Address;
                    break;
                }
            }
            return result;
        }

    }
}
