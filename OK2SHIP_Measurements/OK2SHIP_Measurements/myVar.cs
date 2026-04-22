using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TDMK_SQL;
using myExcel = Microsoft.Office.Interop.Excel;
using OK2SHIP_Lib;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace OK2SHIP_Measurements
{
    public class myVar
    {
        public static FrmFAI FAI_form;
        public static FrmHistogram Chart_Form;
        public static FrmIPQC_Man IPQC_Man_form;
        public static string curr_user;
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
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public TDMK_OK2SHIP myCode2 = new TDMK_OK2SHIP();
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
        public string Get_start_range(string search_key, string start_addr, myExcel.Worksheet tar_wrksht)
        {
            string result = "";
            myExcel.Range cycle_rgn = tar_wrksht.Range[start_addr];
            for (int i = 0; i < 100; i++)
            {
                myExcel.Range sel_rgn = cycle_rgn.Offset[i, 0];
                string sel_rgn_val = checkDBNull(sel_rgn.Value);
                if (sel_rgn_val.Replace(" ", "").ToUpper() == search_key.Replace(" ", "").ToUpper())
                {
                    result = cycle_rgn.Offset[i + sel_rgn.MergeArea.Rows.Count, 0].AddressLocal;
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
                string sel_rgn_val = checkDBNull(sel_rgn.Value);
                if (sel_rgn_val.Replace(" ", "").ToUpper() == search_key.Replace(" ", "").ToUpper())
                {
                    result = sel_rgn.AddressLocal;
                    break;
                }
            }
            return result;
        }
        public List<string> Get_Listdata_addr(string search_item, string search_key, myExcel.Worksheet tar_wrksht, string start_range_addr, bool col_direction, ref Dictionary<string, List<string>> dic_ACF_locate)
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
                string temp_rgn_val = checkDBNull(temp_rgn.Value);
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
                    else
                    {
                        count_null = 0;
                        if (result.Count > 0)
                        {
                            if (dic_ACF_locate.Keys.ToList().IndexOf(temp_rgn_val) == -1)
                            {
                                dic_ACF_locate.Add(temp_rgn_val, new List<string> { temp_addr });
                            }
                            else
                            {
                                dic_ACF_locate[temp_rgn_val].Add(temp_addr);
                            }
                        }
                    }
                }
                if (count_null >= 5)
                {
                    break;
                }
                sel_rgn = temp_rgn;
            }
            return result;
        }
        public DataTable ACF_GetSpec_Process(string f_name)
        {
            char[] split_chr = { '~', '<', '\u2264', '\u2265', '\u00B1', '>' };
            char[] trim_chr = { 'u', 'm', '\u03BC', '\u00B5', ' ', '=' };
            DataTable mySpec_dgv = new DataTable();
            myExcel.Workbook wrkbk = TDMK_Code.open_excel_file(f_name, "", "");
            foreach (myExcel.Worksheet sht in wrkbk.Sheets)
            {
                if (sht.Name.Contains("ACF"))
                {
                    myExcel.Worksheet wrksht = sht;// wrkbk.Sheets["ACF"];
                    wrksht.Activate();
                    string start_rgn = Get_start_range("ACF pad location", "A10", wrksht);
                    myExcel.Range sel_rgn = wrksht.Range[start_rgn].Offset[1, 0];
                    string stop_rgn = Get_start_range("ACF pad location", sel_rgn.AddressLocal, wrksht);
                    Dictionary<string, List<string>> ACF_addr_lst = new Dictionary<string, List<string>>();
                    Get_Listdata_addr("Surface Roughness Measurement", "ACF pad", wrksht, "A10", true, ref ACF_addr_lst);//"ACF pads Surface Roughness"
                    Dictionary<string, Dictionary<string, string>> dic_ACF_Spec = new Dictionary<string, Dictionary<string, string>>();
                    foreach (var spec in ACF_addr_lst)
                    {
                        string spec_rgn_addr = spec.Value[0];
                        myExcel.Range spec_rgn = wrksht.Range[spec_rgn_addr].Offset[0, 1];
                        Dictionary<string, string> spec_ACF = new Dictionary<string, string>();
                        for (int i = 0; i < wrksht.Range[spec_rgn_addr].MergeArea.Rows.Count; i++)
                        {
                            string temp_val = spec_rgn.Offset[i, 0].Value;
                            List<string> lst = temp_val.Split(new char[] { '(', ')' }).ToList();
                            string item = lst.First().Trim();
                            lst.Remove(lst.First());
                            lst.Remove(lst.Last());
                            string item_val = lst.First();
                            spec_ACF.Add(item, item_val);
                        }
                        dic_ACF_Spec.Add(spec.Key, spec_ACF);
                    }
                    int inx = 0;
                    foreach (var t in dic_ACF_Spec)
                    {
                        string col_name = "L" + (inx + 1).ToString() + " Roughness ";
                        foreach (var temp in t.Value)
                        {
                            string cur_col = col_name + temp.Key;
                            mySpec_dgv.Columns.Add(cur_col);
                            if (mySpec_dgv.Rows.Count == 0)
                            {
                                for (int j = 0; j < 3; j++)
                                {
                                    mySpec_dgv.Rows.Add();
                                }
                            }
                            string _limit = temp.Value;
                            DataRow dr_SV = mySpec_dgv.Rows[0];
                            DataRow dr_UL = mySpec_dgv.Rows[1];
                            DataRow dr_LL = mySpec_dgv.Rows[2];
                            if (_limit.Contains('/'))
                            {
                                dr_SV[cur_col] = _limit.Split('-')[0].Trim(trim_chr);
                                string[] temp_limit = _limit.Split('-')[1].Trim(trim_chr).Split('/');
                                double sv = Convert.ToDouble(dr_SV[cur_col]);
                                double ul = Convert.ToDouble(temp_limit[1].Trim(trim_chr));
                                double ll = Convert.ToDouble(temp_limit[0].Trim(trim_chr));
                                dr_UL[cur_col] = (sv + ul).ToString();
                                dr_LL[cur_col] = (sv - ll).ToString();
                            }
                            else
                            {

                                string[] UL_LL = _limit.Split(split_chr);
                                if (_limit.Contains('>'))
                                {
                                    dr_SV[cur_col] = _limit;
                                    dr_UL[cur_col] = UL_LL[0].Trim(trim_chr);
                                    dr_LL[cur_col] = UL_LL[1].Trim(trim_chr);
                                }
                                else
                                {
                                    if (_limit.Contains('\u00B1'))
                                    {
                                        dr_SV[cur_col] = UL_LL[0];
                                        dr_UL[cur_col] = (Convert.ToDouble(UL_LL[0].Trim(trim_chr)) + Convert.ToDouble(UL_LL[1].Trim(trim_chr))).ToString();
                                        dr_LL[cur_col] = (Convert.ToDouble(UL_LL[0].Trim(trim_chr)) - Convert.ToDouble(UL_LL[1].Trim(trim_chr))).ToString();
                                    }
                                    else
                                    {
                                        dr_SV[cur_col] = _limit;
                                        dr_LL[cur_col] = UL_LL[0].Trim(trim_chr);
                                        dr_UL[cur_col] = UL_LL[1].Trim(trim_chr);
                                    }
                                }

                            }
                        }
                        inx++;
                    }
                    break;
                }
            }
            wrkbk.Close();
            TDMK_Code.releaseObject(wrkbk);
            return mySpec_dgv;
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
        public Dictionary<string, DataTable> Export_FAI_Batch2(string tar_ItemCode, string tar_LotNo, string format_type, string shift)
        {
            Dictionary<string, DataTable> dt_lst = new Dictionary<string, DataTable>();
            char[] trim_char = new char[] { ' ', '\r', '\n' };
            DataTable FAI_Data_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Remark", "Shift" }, new string[] { tar_ItemCode, tar_LotNo, format_type, shift }));
            DataTable FAI_Spec_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "FAI_Spec", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { tar_ItemCode, format_type }));
            string[] sheetno = FAI_Spec_tbl.AsEnumerable().Select(r => r.Field<string>("SheetNo")).Distinct().ToArray();
            foreach (string sht in sheetno)
            {
                DataTable dt = new DataTable();
                Dictionary<string, List<string>> dic_FAI = new Dictionary<string, List<string>>();
                string[] FAI_No = FAI_Spec_tbl.AsEnumerable().Where(r => r.Field<string>("SheetNo") == sht).Select(r => r.Field<string>("FAI_No")).ToArray();
                //if (sht.Contains("SPC"))
                //{
                //    int inx = 0;
                //    foreach (string t in FAI_No)
                //    {
                //        string[] temp = t.Split('_');
                //        string act_val = temp[1];
                //        string[] temp2 = temp[0].Split('/');
                //        FAI_No[inx] = temp2[0] + "_" + act_val;
                //        inx++;
                //    }
                //}
                int inx = 0;
                foreach (string t in FAI_No)
                {
                    FAI_No[inx] =t.Split(new char[] { '/','_'}).FirstOrDefault()+"_"+ t.Split(new char[] { '/', '_' }).LastOrDefault();
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
                    if (!myCode2.check_columns_existed(dt, fai.Key))
                    {
                        dt.Columns.Add(fai.Key);
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
                    //string strFilePath = Path.Combine(Application.StartupPath, sht + ".csv");
                    //StreamWriter sw = new StreamWriter(strFilePath, false);
                    //string line = "";
                    //foreach(DataColumn dc in dt.Columns)
                    //{
                    //    line += dc.ColumnName + ",";
                    //}
                    //line = line.Trim(',');
                    //sw.WriteLine(line);
                    //foreach (DataRow dr in dt.Rows)
                    //{
                    //    line = "";
                    //    foreach (DataColumn dc in dt.Columns)
                    //    {
                    //        if (myCode2.checkDBNull(dr[dc]) == "")
                    //        {
                    //            line += "N/A,";
                    //        }
                    //        else
                    //        {
                    //            line += dr[dc].ToString() + ",";
                    //        }
                    //    }
                    //    line = line.Trim(',');
                    //    sw.WriteLine(line);
                    //}
                    //sw.Close();
                    dt_lst.Add(sht, dt);
                }
            }
            return dt_lst;
        }
        public myExcel.Workbook Table_To_CSV(string tbl_name, DataTable src_dt)
        {
            myExcel.Workbook result = null;
            if (src_dt.Rows.Count > 0)
            {
                string strFilePath = Path.Combine(Application.StartupPath, tbl_name + ".csv");
                StreamWriter sw = new StreamWriter(strFilePath, false);
                string line = "";
                foreach (DataColumn dc in src_dt.Columns)
                {
                    line += dc.ColumnName + ",";
                }
                line = line.Trim(',');
                sw.WriteLine(line);
                foreach (DataRow dr in src_dt.Rows)
                {
                    line = "";
                    foreach (DataColumn dc in src_dt.Columns)
                    {
                        if (myCode2.checkDBNull(dr[dc]) == "")
                        {
                            line += "N/A,";
                        }
                        else
                        {
                            line += dr[dc].ToString() + ",";
                        }
                    }
                    line = line.Trim(',');
                    sw.WriteLine(line);
                }
                sw.Close();
                result = TDMK_Code.open_excel_file(strFilePath, "", "");
            }
            return result;
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
                    if (FAI_keys_lst.Any(x => sht.Name.Contains(x)))   //if ((sht.Name.Contains("FAI")) || (sht.Name.Contains("SPC")))
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
                                int row_num = Math.Min(sht.UsedRange.Rows.Count, fai_loc[_data_addr] + 1);
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
        public void Export_To_FAI2(myExcel.Workbook src_format_wrk, string _ItemCode, string _LotNo, myExcel.Workbook src_data_wrkbook, string format_type)
        {
            //DataTable FAI_Data_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" ,"Remark"}, new string[] { _ItemCode, _LotNo,format_type }));
            DataTable FAI_Spec_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "FAI_Spec", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { _ItemCode, format_type }));
            List<string> FAI_keys_lst = new List<string> { "FAI", "SPC", "CPK", "parentheses" };
            foreach (myExcel.Worksheet sht in src_data_wrkbook.Worksheets)
            {
                try
                {
                    if (FAI_keys_lst.Any(x => sht.Name.Contains(x)))   //if ((sht.Name.Contains("FAI")) || (sht.Name.Contains("SPC")))
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
                                int row_num = Math.Min(sht.UsedRange.Rows.Count, fai_loc[_data_addr] + 1);
                                myExcel.Range data_rgn = sht.Range[sht.Cells[2, 1], sht.Cells[row_num, col_num]];
                                data_rgn.Copy();
                                format_rgn.PasteSpecial(myExcel.XlPasteType.xlPasteValues, myExcel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
                                if (format_type == "MASS")
                                {
                                    string itemcode_loc = myCode2.Find_Cell_Addr("ITEM CODE:", "B1", format_wrksheet, false);
                                    string lotno_loc = myCode2.Find_Cell_Addr("LOT:", "B1", format_wrksheet, false);
                                    string itemname_loc = myCode2.Find_Cell_Addr("ITEM NAME:", "B1", format_wrksheet, false);
                                    string date_loc = myCode2.Find_Cell_Addr("NGÀY:", "B1", format_wrksheet, false);

                                    string format_name = Path.GetFileNameWithoutExtension(src_format_wrk.Name);
                                    string[] temp = format_name.Split('-');
                                    string item_name = "";
                                    if (temp.Length > 1)
                                    {
                                        item_name = temp[1];
                                    }
                                    format_wrksheet.Range[itemname_loc].Offset[0, 1].Value = item_name;
                                    format_wrksheet.Range[itemcode_loc].Offset[0, 1].Value = _ItemCode;
                                    format_wrksheet.Range[lotno_loc].Offset[0, 1].Value = _LotNo;
                                    format_wrksheet.Range[date_loc].Offset[0, 1].Value = DateTime.Now.ToShortDateString();
                                }
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
    }
}
