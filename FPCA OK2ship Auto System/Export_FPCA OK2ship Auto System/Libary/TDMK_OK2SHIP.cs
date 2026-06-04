using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SQL;

namespace Export_FPCA_OK2ship_Auto_System.Libary
{
    public class TDMK_OK2SHIP
    {
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public struct FAI_Spec
        {
            public string FAI_Name;
            public string check_side;
            public string SetVal;
            public string UL;
            public string LL;
            public string sheetno;
            public string instrument;
            public FAI_Spec(string _FAI_Name, string _check_side, string _SetVal, string _UL, string _LL, string _sheetno, string _instrument)
            {
                FAI_Name = _FAI_Name;
                check_side = _check_side;
                SetVal = _SetVal;
                UL = _UL;
                LL = _LL;
                sheetno = _sheetno;
                instrument = _instrument;
            }
        }
        public struct Calculate_CPK
        {
            public string FAI_No;
            public double stdev;
            public double mean;
            public double max;
            public double min;
            public string CP;
            public string CPKL;
            public string CPKU;
            public double CPK;
            public double CPKM;
        }
        public struct FAI_Histogram_Data
        {
            public double[] _Bin_data;
            public int[] _Freq_bin_data;
            public double[] _Modified_NormDist_data;
            public string _FAI_No;
            public double[] _FAI_Data;
        }
        public bool IsNumeric(string text)
        {
            double test;
            return double.TryParse(text, out test);
        }
        public string IsNumeric_Val(string text)
        {
            double test;
            if (double.TryParse(text, out test))
            {
                return test.ToString();
            }
            else
            {
                return "";
            }
            //return double.TryParse(text, out test);
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
        public bool check_columns_existed(DataTable src_tbl, string find_col_name)
        {
            bool _result = false;
            foreach (DataColumn c in src_tbl.Columns)
            {
                //string col_name = c.ColumnName.Split('_')[0];
                string col_name = c.ColumnName;
                if (col_name == find_col_name)
                {
                    _result = true;
                    break;
                }
            }
            return _result;
        }
        public bool check_columns_existed_inx(DataTable src_tbl, string find_col_name, ref int col_inx)
        {
            bool _result = false;
            int _inx = 0;
            foreach (DataColumn c in src_tbl.Columns)
            {
                string col_name = c.ColumnName.Split('_')[0];
                //string col_name = c.ColumnName;
                if (col_name == find_col_name)
                {
                    _result = true;
                    col_inx = _inx;
                    break;
                }
                else
                {
                    _result = false;
                }
                _inx++;
            }
            return _result;
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
                Get_ListTable(-1, FAI_spec_dt, new string[] { "SheetNo", "FAI_No" }, ref dt_lst, "Distribution");
                foreach (DataTable dt in dt_lst)
                {
                    if (dt.Rows.Count > 0)
                    {
                        string sheetno = checkDBNull(dt.Rows[0]["SheetNo"]);
                        if (FAI_SheetNo_list.IndexOf(sheetno) != -1)
                        {
                            string fai_no = checkDBNull(dt.Rows[0]["FAI_No"]);
                            if (!check_columns_existed(cur_dt, fai_no))
                            {
                                cur_dt.Columns.Add(fai_no);
                            }
                            if (cur_dt.Rows.Count == 0)
                            {
                                for (int i = 0; i < 4; i++)
                                {
                                    cur_dt.Rows.Add();
                                }
                            }
                            cur_dt.Rows[0][fai_no] = dt.Rows[0]["NormDim"];
                            cur_dt.Rows[1][fai_no] = dt.Rows[0]["TolMax"];
                            cur_dt.Rows[2][fai_no] = dt.Rows[0]["TolMin"];
                            cur_dt.Rows[3][fai_no] = dt.Rows[0]["Distribution"];
                        }
                    }
                }
            }
            return cur_dt;
        }
        public void Calcul_CPK_FAI_Located_NG(FAI_Spec src_FAI_spec, string[] tar_FAI_Data, ref Calculate_CPK tar_CPK_Result, ref FAI_Histogram_Data tar_Histogram_Data, List<int> Item_Located)
        {
            List<SortedDictionary<uint, double>> myData = new List<SortedDictionary<uint, double>>();
            List<SortedDictionary<uint, double>> myResult = new List<SortedDictionary<uint, double>>();

            Double[] data_arr = new double[tar_FAI_Data.Length];
            int inx = 0;
            foreach (string c in tar_FAI_Data)
            {
                if (checkDBNull(c) != "")
                {
                    data_arr[inx] = Convert.ToDouble(c);
                    inx++;
                }
            }
            Array.Resize(ref data_arr, inx);
            string Side_check = src_FAI_spec.check_side;//cbSideCheck.SelectedItem.ToString();
            double Norminal_Dim = Convert.ToDouble(src_FAI_spec.SetVal);// Convert.ToDouble(txtNorminalDim.Text);

            double UL = 0;// Norminal_Dim + Convert.ToDouble(src_FAI_spec.UL);
            double LL = 0;// Norminal_Dim - Convert.ToDouble(src_FAI_spec.LL);

            if (IsNumeric_Val(src_FAI_spec.UL) != "")
            {
                UL = Convert.ToDouble(src_FAI_spec.UL); // Norminal_Dim + Convert.ToDouble(src_FAI_spec.UL);
            }
            if (IsNumeric_Val(src_FAI_spec.LL) != "")
            {
                LL = Convert.ToDouble(src_FAI_spec.LL); // Norminal_Dim - Convert.ToDouble(src_FAI_spec.LL);
            }
            double stdev = CalculateStandardDeviation(data_arr);// tg.StDev(data_arr);
            double mean = data_arr.Average();
            double max = data_arr.Max();
            double min = data_arr.Min();

            double CP = (UL - LL) / (6 * stdev);
            double CPKL = (mean - LL) / (3 * stdev);
            double CPKU = (UL - mean) / (3 * stdev);
            double CPK = new double[] { CPKL, CPKU }.Min();
            double tg1 = (mean - Norminal_Dim) / stdev;
            double CPKM = CPK / Math.Sqrt(1 + Math.Pow(tg1, 2));

            double margin = 2 * stdev;
            double mean_minus_7sig = mean - (7 * stdev);
            double mean_plus_7sig = mean + (7 * stdev);

            double bin_start;
            double bin_end;
            if (Side_check == "SingleSide-USL")
            {
                bin_start = new double[] { mean_minus_7sig, UL - margin }.Min();
            }
            else
            {
                bin_start = new double[] { LL - margin, mean_minus_7sig }.Min();
            }

            if (Side_check == "SingleSide-LSL")
            {
                bin_end = new double[] { mean_plus_7sig, LL + margin }.Max();
            }
            else
            {
                bin_end = new double[] { UL + margin, mean_minus_7sig }.Max();
            }

            double bin_range = bin_end - bin_start;
            double Qty = 50;
            double bin_step = bin_range / Qty;
            double[] bin_data = new double[51];
            int[] freq_bin_data = new int[50];
            double[] modified_NormDist = new double[50];
            double[] NormDist_Bin = new double[50];
            double XiShu;

            for (int i = 0; i < 51; i++)
            {
                bin_data[i] = bin_start + i * bin_step;
            }
            for (int i = 0; i < 50; i++)
            {
                freq_bin_data[i] = Countif(data_arr, bin_data[i + 1]) - Countif(data_arr, bin_data[i]);
            }
            for (int i = 0; i < 50; i++)
            {
                NormDist_Bin[i] = normdist(bin_data[i], mean, stdev, false);
            }
            myData = data_arr.Bucketize6(50, bin_start, bin_end, ref freq_bin_data);
            XiShu = freq_bin_data.Max() / NormDist_Bin.Max();
            for (int i = 0; i < 50; i++)
            {
                modified_NormDist[i] = NormDist_Bin[i] * XiShu;
            }
            //var first_LL = bin_data.Where(x => x >= LL).FirstOrDefault();
            //var first_LL_inx = bin_data.ToList().IndexOf(first_LL);
            //myResult.Add(myData[first_LL_inx]);
            //var first_UL = bin_data.Where(x => x > UL).FirstOrDefault();
            //var first_UL_inx = bin_data.ToList().IndexOf(first_LL);
            //myResult.Add(myData[first_UL_inx-1]);

            var targetIndices = freq_bin_data.Select((val, i) => new { Value = val, Index = i }).Where(x => x.Value > 0);
            var sel_inx = targetIndices.Select(x => x.Index).ToList();
            var tar_norm = modified_NormDist.Select((val, i) => new { Value = val, Index = i }).Where(x => sel_inx.IndexOf(x.Index) != -1).ToList();
            var norm_max = tar_norm.Select(x => x.Value).Max();
            var max_inx = tar_norm.Where(x => x.Value == norm_max).Select(x => x.Index).FirstOrDefault();
            int qty = Math.Min(Math.Abs(max_inx - sel_inx.Max()), Math.Abs(max_inx - sel_inx.Min()));
            List<int> sel_lst_inx = new List<int>();
            List<int> ignored_lst = new List<int>();
            if (qty > 1)
            {
                sel_lst_inx.Add(max_inx);
                for (int i = 1; i < qty; i++)
                {
                    int tg = max_inx - i;
                    int tg_r = max_inx + i;
                    if (sel_inx.IndexOf(tg) != -1)
                    {
                        sel_lst_inx.Add(tg);
                    }
                    if (sel_inx.IndexOf(tg_r) != -1)
                    {
                        sel_lst_inx.Add(tg_r);
                    }
                }
                ignored_lst = sel_inx.Except(sel_lst_inx).ToList();
            }
            foreach (var ig_inx in ignored_lst)
            {
                myResult.Add(myData[ig_inx]);
            }
            foreach (var item in myResult)
            {
                foreach (KeyValuePair<uint, double> x in item)
                {
                    if (x.Value < UL && x.Value > LL)
                    {
                        Item_Located.Add((int)x.Key);
                    }
                }
            }

            //var targetIndices = freq_bin_data.Select((o, i) => new { Value = o, Index = i }).Where(o => o.Value>0);
            //var sel_inx = targetIndices.Select(x => x.Index).ToList();

            tar_CPK_Result.FAI_No = src_FAI_spec.FAI_Name;

            switch (Side_check)
            {
                case "SingleSide-USL":
                    tar_CPK_Result.CP = "NA";
                    tar_CPK_Result.CPKL = "NA";
                    tar_CPK_Result.CPK = CPKU;
                    break;
                case "SingleSide-LSL":
                    tar_CPK_Result.CP = "NA";
                    tar_CPK_Result.CPKU = "NA";
                    tar_CPK_Result.CPK = CPKL;
                    break;
                default:
                    tar_CPK_Result.CP = CP.ToString("#0.00#");
                    tar_CPK_Result.CPKL = CPKL.ToString("#0.00#");
                    tar_CPK_Result.CPKU = CPKU.ToString("#0.00#");
                    tar_CPK_Result.CPK = CPK;
                    break;
            }

            tar_CPK_Result.CPKM = CPKM;
            tar_CPK_Result.max = max;
            tar_CPK_Result.min = min;
            tar_CPK_Result.stdev = stdev;
            tar_CPK_Result.mean = mean;

            tar_Histogram_Data._FAI_No = src_FAI_spec.FAI_Name;
            tar_Histogram_Data._Bin_data = bin_data;
            tar_Histogram_Data._Freq_bin_data = freq_bin_data;
            tar_Histogram_Data._Modified_NormDist_data = modified_NormDist;
            tar_Histogram_Data._FAI_Data = data_arr;
        }
        public double CalculateStandardDeviation(IEnumerable<double> values)
        {
            double standardDeviation = 0;

            if (values.Any())
            {
                // Compute the average.     
                double avg = values.Average();

                // Perform the Sum of (value-avg)_2_2.      
                double sum = values.Sum(d => Math.Pow(d - avg, 2));

                // Put it all together.      
                standardDeviation = Math.Sqrt((sum) / (values.Count() - 1));
            }

            return standardDeviation;
        }
        public int Countif(double[] src_data, double ref_point)
        {
            int _result = 0;
            foreach (double c in src_data)
            {
                if (c < ref_point)
                {
                    _result++;
                }
            }
            return _result;
        }
        public static double normdist(double x, double mean, double standard_dev, bool cumalative)
        {
            if (cumalative == false)
            {
                double fact = standard_dev * Math.Sqrt(2.0 * Math.PI);
                double expo = (x - mean) * (x - mean) / (2.0 * standard_dev * standard_dev);
                return Math.Exp(-expo) / fact;
            }
            else
            {
                x = (x - mean) / standard_dev;
                if (x == 0)
                    return 0.5;
                double t = 1.0 / (1.0 + 0.2316419 * Math.Abs(x));
                double cdf = t * (1.0 / (Math.Sqrt(2.0 * Math.PI)))
                                * Math.Exp(-0.5 * x * x)
                                * (0.31938153 + t
                                * (-0.356563782 + t
                                * (1.781477937 + t
                                * (-1.821255978 + t * 1.330274429))));
                return x >= 0 ? 1.0 - cdf : cdf;
            }
        }
        public DataTable Load_FAI_ToTable(SqlConnection tar_sqlcon, string tar_ItemCode, string tar_LotNo, string format_type)
        {
            string[] flt_items = new string[5];
            string[] flt_item_vals = new string[5];
            AutoCompleteStringCollection FAI_lst;
            flt_items[0] = "ItemCode";
            flt_items[1] = "LotNo";
            flt_items[2] = "FAI_No";
            flt_items[3] = "SheetNo";
            flt_items[4] = "Remark";
            flt_item_vals[0] = tar_ItemCode;
            flt_item_vals[4] = format_type;
            DataTable FAI_spec_dt = TDMK_Code.Datatable_Filter(tar_sqlcon, "FAI_Spec", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { tar_ItemCode, format_type }));
            List<string> search_key_lst = new List<string>() { "FAI", "Dimension" };
            List<string> SheetNo_lst = FAI_spec_dt.AsEnumerable().Where(x => search_key_lst.Any(y => x.Field<string>("SheetNo").Contains(y))).Select(x => x.Field<string>("SheetNo")).Distinct().ToList();
            DataTable tbl_data = new DataTable();
            foreach (string t_ShtNo in SheetNo_lst)
            {
                flt_item_vals[1] = "";
                flt_item_vals[2] = "";
                flt_item_vals[3] = t_ShtNo;
                FAI_lst = TDMK_Code.Load_Item_Filter_str(tar_sqlcon, "FAI_Spec", "FAI_No", TDMK_Code.filter_str(flt_items, flt_item_vals));
                foreach (string c in FAI_lst)
                {
                    AutoCompleteStringCollection FAI_vals;
                    flt_item_vals[1] = tar_LotNo;
                    flt_item_vals[2] = c;
                    flt_item_vals[3] = "";
                    FAI_vals = TDMK_Code.Load_Item_Filter_str(tar_sqlcon, "FAI_Auto", "FAI_Data", TDMK_Code.filter_str(flt_items, flt_item_vals));
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
            }
            return tbl_data;
        }
        public DataTable Load_FAI_ToTable(SqlConnection tar_sqlcon, string tar_ItemCode, string tar_LotNo, string format_type, string tar_shift)
        {
            string[] flt_items = new string[6];
            string[] flt_item_vals = new string[6];
            AutoCompleteStringCollection FAI_lst;
            flt_items[0] = "ItemCode";
            flt_items[1] = "LotNo";
            flt_items[2] = "FAI_No";
            flt_items[3] = "SheetNo";
            flt_items[4] = "Remark";
            flt_items[5] = "Shift";
            flt_item_vals[0] = tar_ItemCode;
            flt_item_vals[4] = format_type;
            DataTable FAI_spec_dt = TDMK_Code.Datatable_Filter(tar_sqlcon, "FAI_Spec", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { tar_ItemCode, format_type }));
            List<string> search_key_lst = new List<string>() { "FAI", "Dimension" };
            List<string> SheetNo_lst = FAI_spec_dt.AsEnumerable().Where(x => search_key_lst.Any(y => x.Field<string>("SheetNo").Contains(y))).Select(x => x.Field<string>("SheetNo")).Distinct().ToList();
            DataTable tbl_data = new DataTable();
            foreach (string t_ShtNo in SheetNo_lst)
            {
                flt_item_vals[1] = "";
                flt_item_vals[2] = "";
                flt_item_vals[3] = t_ShtNo;
                flt_item_vals[5] = "";
                FAI_lst = TDMK_Code.Load_Item_Filter_str(tar_sqlcon, "FAI_Spec", "FAI_No", TDMK_Code.filter_str(flt_items, flt_item_vals));
                foreach (string c in FAI_lst)
                {
                    AutoCompleteStringCollection FAI_vals;
                    flt_item_vals[1] = tar_LotNo;
                    flt_item_vals[2] = c;
                    flt_item_vals[3] = "";
                    flt_item_vals[5] = tar_shift;
                    FAI_vals = TDMK_Code.Load_Item_Filter_str(tar_sqlcon, "FAI_Auto", "FAI_Data", TDMK_Code.filter_str(flt_items, flt_item_vals));
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
            }
            return tbl_data;
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
        public void Get_List_data(int col_inx, DataTable myDt, string[] src_arr, ref List<string> src_lst_data, string tar_item, bool distinct_en)
        {
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
                                Get_List_data(col_inx + 1, curTbl, src_arr, ref src_lst_data, tar_item, distinct_en);
                            }
                        }
                    }
                    else
                    {
                        if (distinct_en)
                        {
                            src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Distinct().ToArray());
                        }
                        else
                        {
                            src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).ToArray());
                        }
                    }
                }
            }
            else
            {
                if (distinct_en)
                {
                    src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Distinct().ToArray());
                }
                else
                {
                    src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).ToArray());
                }
            }
        }
        public void Get_List_data2(int col_inx, DataTable myDt, string[] src_arr, ref List<string> src_lst_data, string tar_item, bool distinct_en)
        {
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
                                Get_List_data2(col_inx + 1, curTbl, src_arr, ref src_lst_data, tar_item, distinct_en);
                            }
                        }
                    }
                    else
                    {
                        if (distinct_en)
                        {
                            //src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Distinct().ToArray());
                            src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Where(r => r != null).Distinct().ToArray());
                        }
                        else
                        {
                            //src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).ToArray());
                            src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Where(r => r != null).ToArray());
                        }
                    }
                }
            }
            else
            {
                if (distinct_en)
                {
                    //src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Distinct().ToArray());
                    src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Where(r => r != null).Distinct().ToArray());
                }
                else
                {
                    //src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).ToArray());
                    src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Where(r => r != null).ToArray());
                }
            }
        }
        public void Calcul_CPK_FAI2(FAI_Spec src_FAI_spec, Double[] data_arr, ref Calculate_CPK tar_CPK_Result, ref FAI_Histogram_Data tar_Histogram_Data)
        {
            //Double[] data_arr = new double[tar_FAI_Data.Length];
            //bool calc_en = false;
            //int inx = 0;
            //foreach (string c in tar_FAI_Data)
            //{
            //    if ((checkDBNull(c) != "") && (IsNumeric(c)))
            //    {
            //        calc_en = true;
            //        data_arr[inx] = Convert.ToDouble(c);
            //        inx++;
            //    }
            //    else
            //    {
            //        calc_en = false;
            //        ///break;
            //    }
            //}

            //Array.Resize(ref data_arr, inx);
            string Side_check = src_FAI_spec.check_side;//cbSideCheck.SelectedItem.ToString();
            double Norminal_Dim = Convert.ToDouble(src_FAI_spec.SetVal);// Convert.ToDouble(txtNorminalDim.Text);
            //double UL = Norminal_Dim + Convert.ToDouble(src_FAI_spec.UL);
            //double LL = Norminal_Dim - Convert.ToDouble(src_FAI_spec.LL);

            double UL = 0;// Norminal_Dim + Convert.ToDouble(src_FAI_spec.UL);
            double LL = 0;// Norminal_Dim - Convert.ToDouble(src_FAI_spec.LL);
            double CPKL = 0;// (mean - LL) / (3 * stdev);
            double CPKU = 0;// (UL - mean) / (3 * stdev);

            double stdev = CalculateStandardDeviation(data_arr);// tg.StDev(data_arr);
            double mean = data_arr.Average();
            double max = data_arr.Max();
            double min = data_arr.Min();

            if (IsNumeric_Val(src_FAI_spec.UL) != "")
            {
                UL = Convert.ToDouble(src_FAI_spec.UL);//Norminal_Dim + Convert.ToDouble(src_FAI_spec.UL);
                CPKU = (UL - mean) / (3 * stdev);
            }
            if (IsNumeric_Val(src_FAI_spec.LL) != "")
            {
                LL = Convert.ToDouble(src_FAI_spec.LL);//Norminal_Dim - Convert.ToDouble(src_FAI_spec.LL);
                CPKL = (mean - LL) / (3 * stdev);
            }

            double CP = (UL - LL) / (6 * stdev);
            double CPK = new double[] { CPKL, CPKU }.Min();
            double tg1 = (mean - Norminal_Dim) / stdev;
            double CPKM = CPK / Math.Sqrt(1 + Math.Pow(tg1, 2));

            double margin = 2 * stdev;
            double mean_minus_7sig = mean - (7 * stdev);
            double mean_plus_7sig = mean + (7 * stdev);

            double bin_start;
            double bin_end;
            if (Side_check == "SingleSide-USL")
            {
                bin_start = new double[] { mean_minus_7sig, UL - margin }.Min();
            }
            else
            {
                bin_start = new double[] { LL - margin, mean_minus_7sig }.Min();
            }

            if (Side_check == "SingleSide-LSL")
            {
                bin_end = new double[] { mean_plus_7sig, LL + margin }.Max();
            }
            else
            {
                bin_end = new double[] { UL + margin, mean_minus_7sig }.Max();
            }

            double bin_range = bin_end - bin_start;
            double Qty = 50;
            double bin_step = bin_range / Qty;
            double[] bin_data = new double[51];
            int[] freq_bin_data = new int[50];
            double[] modified_NormDist = new double[50];
            double[] NormDist_Bin = new double[50];
            double XiShu;

            for (int i = 0; i < 51; i++)
            {
                bin_data[i] = bin_start + i * bin_step;
            }
            for (int i = 0; i < 50; i++)
            {
                freq_bin_data[i] = Countif(data_arr, bin_data[i + 1]) - Countif(data_arr, bin_data[i]);
            }
            for (int i = 0; i < 50; i++)
            {

                NormDist_Bin[i] = normdist(bin_data[i], mean, stdev, false);
            }
            XiShu = freq_bin_data.Max() / NormDist_Bin.Max();
            for (int i = 0; i < 50; i++)
            {
                modified_NormDist[i] = NormDist_Bin[i] * XiShu;
            }
            tar_CPK_Result.FAI_No = src_FAI_spec.FAI_Name;

            switch (Side_check)
            {
                case "SingleSide-USL":
                    tar_CPK_Result.CP = "NA";
                    tar_CPK_Result.CPKL = "NA";
                    tar_CPK_Result.CPK = CPKU;
                    break;
                case "SingleSide-LSL":
                    tar_CPK_Result.CP = "NA";
                    tar_CPK_Result.CPKU = "NA";
                    tar_CPK_Result.CPK = CPKL;
                    break;
                default:
                    tar_CPK_Result.CP = CP.ToString("#0.00#");
                    tar_CPK_Result.CPKL = CPKL.ToString("#0.00#");
                    tar_CPK_Result.CPKU = CPKU.ToString("#0.00#");
                    tar_CPK_Result.CPK = CPK;
                    break;
            }

            //tar_CPK_Result.CP = CP;

            //tar_CPK_Result.CPKL = CPKL;
            tar_CPK_Result.CPKM = CPKM;
            //tar_CPK_Result.CPKU = CPKU;
            tar_CPK_Result.max = max;
            tar_CPK_Result.min = min;
            tar_CPK_Result.stdev = stdev;
            tar_CPK_Result.mean = mean;

            tar_Histogram_Data._FAI_No = src_FAI_spec.FAI_Name;
            tar_Histogram_Data._Bin_data = bin_data;
            tar_Histogram_Data._Freq_bin_data = freq_bin_data;
            tar_Histogram_Data._Modified_NormDist_data = modified_NormDist;
            tar_Histogram_Data._FAI_Data = data_arr;
        }
        public void check_FAIdata_inSpec(DataGridView src_DGV_data, DataGridView src_DGV_Spec)
        {
            List<string> spec_col_lst = new List<string>();

            foreach (DataGridViewColumn dc in src_DGV_Spec.Columns)
            {
                spec_col_lst.Add(dc.Name);
            }
            foreach (DataGridViewColumn dc in src_DGV_data.Columns)
            {
                string cur_col = dc.Name;
                if (spec_col_lst.IndexOf(cur_col) != -1)
                {
                    string Tol_Max = checkDBNull(src_DGV_Spec.Rows[1].Cells[cur_col].Value);
                    string Tol_Min = checkDBNull(src_DGV_Spec.Rows[2].Cells[cur_col].Value);
                    string SetVal = checkDBNull(src_DGV_Spec.Rows[0].Cells[cur_col].Value);
                    double UL = Convert.ToDouble(SetVal) + Convert.ToDouble(Tol_Max);
                    double LL = Convert.ToDouble(SetVal) - Convert.ToDouble(Tol_Min);
                    foreach (DataGridViewRow dr in src_DGV_data.Rows)
                    {
                        string src_act_val = checkDBNull(dr.Cells[cur_col].Value);
                        dr.Cells[cur_col].Style.BackColor = check_in_limit_Color(UL.ToString(), LL.ToString(), src_act_val, SetVal);
                    }
                }
                else
                {
                    foreach (DataGridViewRow dr in src_DGV_data.Rows)
                    {
                        string src_act_val = checkDBNull(dr.Cells[cur_col].Value);
                        if (src_act_val != "")
                        {
                            dr.Cells[cur_col].Style.BackColor = Color.Gray;
                        }
                        else
                        {
                            dr.Cells[cur_col].Style.BackColor = Color.LightPink;
                        }
                    }
                }
            }
        }
        public void check_FAIdata_inSpec2(DataGridView src_DGV_data, DataGridView src_DGV_Spec)
        {
            List<string> spec_col_lst = new List<string>();

            foreach (DataGridViewColumn dc in src_DGV_Spec.Columns)
            {
                spec_col_lst.Add(dc.Name);
            }
            foreach (DataGridViewColumn dc in src_DGV_data.Columns)
            {
                string cur_col = dc.Name;
                if (spec_col_lst.IndexOf(cur_col) != -1)
                {
                    string Tol_Max = checkDBNull(src_DGV_Spec.Rows[1].Cells[cur_col].Value);
                    string Tol_Min = checkDBNull(src_DGV_Spec.Rows[2].Cells[cur_col].Value);
                    string SetVal = checkDBNull(src_DGV_Spec.Rows[0].Cells[cur_col].Value);
                    string UL = IsNumeric_Val(Tol_Max);
                    string LL = IsNumeric_Val(Tol_Min);
                    foreach (DataGridViewRow dr in src_DGV_data.Rows)
                    {
                        string src_act_val = checkDBNull(dr.Cells[cur_col].Value);
                        dr.Cells[cur_col].Style.BackColor = check_in_limit_Color(UL, LL, src_act_val, SetVal);
                    }
                }
                else
                {
                    foreach (DataGridViewRow dr in src_DGV_data.Rows)
                    {
                        string src_act_val = checkDBNull(dr.Cells[cur_col].Value);
                        if (src_act_val != "")
                        {
                            dr.Cells[cur_col].Style.BackColor = Color.Gray;
                        }
                        else
                        {
                            dr.Cells[cur_col].Style.BackColor = Color.LightPink;
                        }
                    }
                }
            }
        }
        public bool check_FAIdata_inSpec(DataTable src_data, DataTable src_Spec)
        {
            bool result = false;
            List<string> spec_col_lst = new List<string>();
            foreach (DataColumn dc in src_Spec.Columns)
            {
                spec_col_lst.Add(dc.ColumnName);
            }
            foreach (DataColumn dc in src_data.Columns)
            {
                string cur_col = dc.ColumnName;
                if (spec_col_lst.IndexOf(cur_col) != -1)
                {
                    string USL = checkDBNull(src_Spec.Rows[1][cur_col]);
                    string LSL = checkDBNull(src_Spec.Rows[2][cur_col]);
                    string SetVal = IsNumeric_Val(checkDBNull(src_Spec.Rows[0][cur_col]));
                    string UL = IsNumeric_Val(USL);
                    string LL = IsNumeric_Val(LSL);
                    foreach (DataRow dr in src_data.Rows)
                    {
                        string src_act_val = IsNumeric_Val(checkDBNull(dr[cur_col]));
                        if (!check_in_limit(UL, LL, src_act_val, SetVal))
                        {
                            result = false;
                            return result;

                        }
                        else
                        {
                            result = true;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Không tìm thấy Spec của " + cur_col, "Thông báo");
                    break;
                }
            }
            return result;
        }
        public Color check_in_limit_Color(string src_UL, string src_LL, string src_act_val, string SetVal)
        {
            Color tar_color = Color.LightPink;
            bool UL_equal_comp = true;
            bool LL_equal_comp = true;

            if (!IsNumeric(SetVal) && (!SetVal.Contains("=")))
            {
                UL_equal_comp = false;
                LL_equal_comp = false;
            }
            if (IsNumeric(src_act_val))
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
                        if (act_val < UL)
                        {
                            tar_color = Color.White;
                        }
                        else
                        {
                            if (UL_equal_comp)
                            {
                                if (act_val == UL)
                                {
                                    tar_color = Color.White;
                                }
                                else
                                {
                                    tar_color = Color.Yellow;
                                }
                            }
                            else
                            {
                                tar_color = Color.Yellow;
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
                            tar_color = Color.White;
                        }
                        else
                        {
                            if (LL_equal_comp)
                            {
                                if (act_val == LL)
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
                                tar_color = Color.LightBlue;
                            }
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
        public bool check_in_limit(string src_UL, string src_LL, string src_act_val, string SetVal)
        {
            bool _result = false;
            bool UL_equal_comp = true;
            bool LL_equal_comp = true;

            if (!IsNumeric(SetVal) && (!SetVal.Contains("=")))
            {
                UL_equal_comp = false;
                LL_equal_comp = false;
            }
            if (IsNumeric(src_act_val))
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
            else
            {
                _result = true;
            }
            return _result;
        }
    }
    public static class Calcu_process
    {
        public static List<SortedDictionary<uint, double>> Group_data_Double(this IEnumerable<double> source, int totalBuckets)
        {
            var _result = new List<SortedDictionary<uint, double>>();
            for (int i = 0; i < totalBuckets; i++)
            {
                _result.Add(new SortedDictionary<uint, double>());
            }
            var min = source.Min();
            var max = source.Max();
            var bucketSize = (max - min) / totalBuckets;
            uint id = 0;
            foreach (var value in source)
            {
                int bucketIndex = 0;
                if (bucketSize > 0.0)
                {
                    bucketIndex = (int)((value - min) / bucketSize);
                    if (bucketIndex == totalBuckets)
                    {
                        bucketIndex--;
                    }
                }
                _result[bucketIndex].Add(id, value);
                id++;
            }
            return _result;
        }
        public static List<SortedDictionary<uint, double>> Group_data_String(this IEnumerable<string> source, int totalBuckets)
        {
            var _result = new List<SortedDictionary<uint, double>>();
            for (int i = 0; i < totalBuckets; i++)
            {
                _result.Add(new SortedDictionary<uint, double>());
            }
            IEnumerable<double> _source = ConvertToDouble(source);
            var min = _source.Min();
            var max = _source.Max();
            var bucketSize = (max - min) / totalBuckets;
            uint id = 0;
            foreach (var value in _source)
            {
                int bucketIndex = 0;
                if (bucketSize > 0.0)
                {
                    bucketIndex = (int)((value - min) / bucketSize);
                    if (bucketIndex == totalBuckets)
                    {
                        bucketIndex--;
                    }
                }
                _result[bucketIndex].Add(id, value);
                id++;
            }
            return _result;
        }
        public static IEnumerable<string> ConvertToString(IEnumerable<double> doubles)
        {
            return doubles.Select(ConvertToString);
        }
        public static string ConvertToString(double d)
        {
            return string.Format("{0:0.000}", d);
        }
        public static IEnumerable<double> ConvertToDouble(IEnumerable<string> src_string)
        {
            return src_string.Select(ConvertToDouble);
        }
        public static Double ConvertToDouble(string d)
        {
            double test;
            if (double.TryParse(d, out test))
            {
                return test;
            }
            else
            {
                return 0;
            }
        }
        public static int[] Bucketize_test(this IEnumerable<double> source, int totalBuckets, double bin_start, double bin_end)
        {
            var min = bin_start;
            var max = bin_end;
            var buckets = new int[totalBuckets];
            var bucketSize = (max - min) / totalBuckets;
            foreach (var value in source)
            {
                int bucketIndex = 0;
                if (bucketSize > 0.0)
                {
                    bucketIndex = (int)((value - min) / bucketSize);
                    if (bucketIndex == totalBuckets)
                    {
                        bucketIndex--;
                    }
                }
                buckets[bucketIndex]++;
            }
            return buckets;
        }
        public static List<SortedDictionary<uint, double>> Bucketize6(this IEnumerable<double> source, int totalBuckets, double bin_start, double bin_end, ref int[] item_freq)
        {
            var _result = new List<SortedDictionary<uint, double>>();
            for (int i = 0; i <= totalBuckets; i++)
            {
                _result.Add(new SortedDictionary<uint, double>());
            }
            var min = bin_start;
            var max = bin_end;
            var bucketSize = (max - min) / totalBuckets;
            uint id = 0;
            foreach (var value in source)
            {
                int bucketIndex = 0;
                if (bucketSize > 0.0)
                {
                    bucketIndex = (int)((value - min) / bucketSize);
                    if (bucketIndex == totalBuckets)
                    {
                        bucketIndex--;
                    }
                }
                if (bucketIndex <= totalBuckets)
                {
                    _result[bucketIndex].Add(id, value);
                    item_freq[bucketIndex]++;
                    id++;
                }
            }
            return _result;
        }
        public static List<SortedDictionary<uint, double>> Group_data_Freq(this IEnumerable<double> source, int totalBuckets, double bin_start, double bin_end, ref int[] item_freq)
        {
            var _result = new List<SortedDictionary<uint, double>>();
            for (int i = 0; i <= totalBuckets; i++)
            {
                _result.Add(new SortedDictionary<uint, double>());
            }
            var min = bin_start;
            var max = bin_end;
            var bucketSize = (max - min) / totalBuckets;
            uint id = 0;
            foreach (var value in source)
            {
                int bucketIndex = 0;
                if (bucketSize > 0.0)
                {
                    bucketIndex = (int)((value - min) / bucketSize);
                    if (bucketIndex == totalBuckets)
                    {
                        bucketIndex--;
                    }

                }
                if ((bucketIndex <= totalBuckets) && (bucketIndex >= 0))
                {
                    _result[bucketIndex].Add(id, value);
                    item_freq[bucketIndex]++;
                    id++;
                }
            }
            return _result;
        }
        public static double CPK(this IEnumerable<double> source, double USL, double LSL)
        {
            //double _result=0;
            double UL = USL;
            double LL = LSL;
            double CPKL = 0;
            double CPKU = 0;
            double stdev = CalculateStandardDeviation(source);// tg.StDev(data_arr);
            double mean = source.Average();
            CPKU = (UL - mean) / (3 * stdev);
            CPKL = (mean - LL) / (3 * stdev);
            double _result = new double[] { CPKL, CPKU }.Min();
            return _result;
        }
        public static double CalculateStandardDeviation(IEnumerable<double> values)
        {
            double standardDeviation = 0;

            if (values.Any())
            {
                // Compute the average.     
                double avg = values.Average();

                // Perform the Sum of (value-avg)_2_2.      
                double sum = values.Sum(d => Math.Pow(d - avg, 2));

                // Put it all together.      
                standardDeviation = Math.Sqrt((sum) / (values.Count() - 1));
            }

            return standardDeviation;
        }
        public static double normdist(double x, double mean, double standard_dev, bool cumalative)
        {
            if (cumalative == false)
            {
                double fact = standard_dev * Math.Sqrt(2.0 * Math.PI);
                double expo = (x - mean) * (x - mean) / (2.0 * standard_dev * standard_dev);
                return Math.Exp(-expo) / fact;
            }
            else
            {
                x = (x - mean) / standard_dev;
                if (x == 0)
                    return 0.5;
                double t = 1.0 / (1.0 + 0.2316419 * Math.Abs(x));
                double cdf = t * (1.0 / (Math.Sqrt(2.0 * Math.PI)))
                                * Math.Exp(-0.5 * x * x)
                                * (0.31938153 + t
                                * (-0.356563782 + t
                                * (1.781477937 + t
                                * (-1.821255978 + t * 1.330274429))));
                return x >= 0 ? 1.0 - cdf : cdf;
            }
        }
        public static List<SortedDictionary<uint, double>> Calcul_CPK2(Double[] data_arr, string USL, string LSL, int sigma, int count)
        {
            List<SortedDictionary<uint, double>> myData1 = new List<SortedDictionary<uint, double>>();
            List<SortedDictionary<uint, double>> myResult1 = new List<SortedDictionary<uint, double>>();
            double UL = Convert.ToDouble(USL);
            double LL = Convert.ToDouble(LSL);
            double CPKL = 0;
            double CPKU = 0;
            double stdev = CalculateStandardDeviation(data_arr);
            double mean = data_arr.Average();
            double max = data_arr.Max();
            double min = data_arr.Min();
            CPKU = (UL - mean) / (3 * stdev);
            CPKL = (mean - LL) / (3 * stdev);
            double CP = (UL - LL) / (6 * stdev);
            double CPK = new double[] { CPKL, CPKU }.Min();
            double margin = 2 * stdev;
            double mean_minus_7sig = mean - (sigma * stdev);
            double mean_plus_7sig = mean + (sigma * stdev);
            double bin_start;
            double bin_end;
            bin_start = new double[] { mean_minus_7sig, UL - margin }.Min();
            bin_end = new double[] { UL + margin, mean_minus_7sig }.Max();
            double bin_range = bin_end - bin_start;
            double Qty = count;
            double bin_step = bin_range / Qty;
            double[] bin_data = new double[count + 1];
            int[] freq_bin_data = new int[count];
            double[] modified_NormDist = new double[count];
            double[] NormDist_Bin = new double[count];
            double XiShu;
            for (int i = 0; i < count + 1; i++)
            {
                bin_data[i] = bin_start + i * bin_step;
            }
            myData1 = data_arr.Bucketize6(count, bin_start, bin_end, ref freq_bin_data);
            for (int i = 0; i < count; i++)
            {

                NormDist_Bin[i] = normdist(bin_data[i], mean, stdev, false);
            }
            XiShu = freq_bin_data.Max() / NormDist_Bin.Max();
            for (int i = 0; i < count; i++)
            {
                modified_NormDist[i] = NormDist_Bin[i] * XiShu;
                if ((modified_NormDist[i] >= freq_bin_data[i]) && (freq_bin_data[i] > 0))
                {
                    myResult1.Add(myData1[i]);
                }
            }
            return myResult1;
        }
        public static List<string> Items_OK_CPK_List(Double[] data_arr, string USL, string LSL, int sigma, int count)
        {
            List<SortedDictionary<uint, double>> myData1 = new List<SortedDictionary<uint, double>>();
            List<SortedDictionary<uint, double>> myResult1 = new List<SortedDictionary<uint, double>>();
            List<string> result = new List<string>();
            double UL = Convert.ToDouble(USL);
            double LL = Convert.ToDouble(LSL);
            double CPKL = 0;
            double CPKU = 0;
            double stdev = CalculateStandardDeviation(data_arr);
            double mean = data_arr.Average();
            double max = data_arr.Max();
            double min = data_arr.Min();
            CPKU = (UL - mean) / (3 * stdev);
            CPKL = (mean - LL) / (3 * stdev);
            double CP = (UL - LL) / (6 * stdev);
            double CPK = new double[] { CPKL, CPKU }.Min();
            double margin = 2 * stdev;
            double mean_minus_7sig = mean - (sigma * stdev);
            double mean_plus_7sig = mean + (sigma * stdev);
            double bin_start;
            double bin_end;
            bin_start = new double[] { mean_minus_7sig, UL - margin }.Min();
            bin_end = new double[] { UL + margin, mean_minus_7sig }.Max();
            double bin_range = bin_end - bin_start;
            double Qty = count;
            double bin_step = bin_range / Qty;
            double[] bin_data = new double[count + 1];
            int[] freq_bin_data = new int[count];
            double[] modified_NormDist = new double[count];
            double[] NormDist_Bin = new double[count];
            double XiShu;
            for (int i = 0; i < count + 1; i++)
            {
                bin_data[i] = bin_start + i * bin_step;
            }
            myData1 = data_arr.Bucketize6(count, bin_start, bin_end, ref freq_bin_data);
            for (int i = 0; i < count; i++)
            {

                NormDist_Bin[i] = normdist(bin_data[i], mean, stdev, false);
            }
            XiShu = freq_bin_data.Max() / NormDist_Bin.Max();
            for (int i = 0; i < count; i++)
            {
                modified_NormDist[i] = NormDist_Bin[i] * XiShu;
                if ((modified_NormDist[i] >= freq_bin_data[i]) && (freq_bin_data[i] > 0))
                {
                    myResult1.Add(myData1[i]);
                }
            }
            foreach (var item in myResult1)
            {
                if (item.Values.ToList().Count > 3)
                {
                    foreach (KeyValuePair<uint, double> x in item)
                    {
                        result.Add(x.Key.ToString());
                    }
                }
            }
            return result;
        }
        public static void calc_CPK(double[] data_arr, double UL, double LL)
        {
            double stdev = CalculateStandardDeviation(data_arr);// tg.StDev(data_arr);
            double mean = data_arr.Average();
            double max = data_arr.Max();
            double min = data_arr.Min();

            double CP = (UL - LL) / (6 * stdev);
            double CPKL = (mean - LL) / (3 * stdev);
            double CPKU = (UL - mean) / (3 * stdev);
            double CPK = new double[] { CPKL, CPKU }.Min();


        }

    }
}
