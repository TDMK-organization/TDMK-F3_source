using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FAI_Data_Collect
{
    public class SEI_Lib
    {
       public struct FAI_Spec
        {
            public string FAI_Name;
            public string check_side;
            public string SetVal;
            public string UL;
            public string LL;
            public FAI_Spec(string _FAI_Name, string _check_side, string _SetVal, string _UL, string _LL)
            {
                FAI_Name = _FAI_Name;
                check_side = _check_side;
                SetVal = _SetVal;
                UL = _UL;
                LL = _LL;
            }
        }
        public struct Calculate_CPK
        {
            public string FAI_No;
            public  double stdev;
            public double mean;
            public double max;
            public  double min;
            public double CP;
            public  double CPKL;
            public double CPKU;
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
        public static bool admin_mode = false;
        public Thread export_data;
        //public static Form1 main_form;
        //public static Form2 Chart_Form;
        public static string curr_user;
        public static string scan_file;
        public void Nikon_Data_display(string _filename, DataGridView tar_DGV, DataGridView DGV_Spec, bool in_FAI)
        {
            string tar_file;
            char[] split_char = { '\t', ',' };
            string[] src_data = new string[1000];
            tar_file = _filename;
            System.IO.StreamReader file = new System.IO.StreamReader(tar_file);
            string newline;
            int i = 0;
            while ((newline = file.ReadLine()) != null)
            {
                src_data[i] = newline.Replace("\"", "");
                i++;
            }
            file.Close();
            Array.Resize<string>(ref src_data, i);
            src_data = src_data.Skip(1).ToArray();
            int col_inx = src_data.Length;
            string[] col_name = new string[col_inx];
            string[] Setval = new string[col_inx];
            string[] UL_val = new string[col_inx];
            string[] LL_val = new string[col_inx];
            tar_DGV.Columns.Clear();
            DataTable tbl_spec = new DataTable();
            tbl_spec = DGV_To_Table(DGV_Spec);
            if (tbl_spec.Rows.Count == 0)
            {
                tbl_spec.Rows.Add();
                tbl_spec.Rows.Add();
                tbl_spec.Rows.Add();
            }
            for (i = 0; i < col_inx; i++)
            {
                if (in_FAI)
                {
                    if (src_data[i].Contains("FAI"))
                    {
                        string[] temp_arr = src_data[i].Split(',');
                        col_name[i] = temp_arr[0];
                        double SV = Convert.ToDouble(temp_arr[3]);
                        double UL = Convert.ToDouble(temp_arr[4]);
                        double LL = Convert.ToDouble(temp_arr[5]);
                        tar_DGV.Columns.Add(col_name[i], col_name[i]);
                        string[] col_val;
                        col_val = src_data[i].Split(',').Skip(6).ToArray();
                        if (tar_DGV.Rows.Count < col_val.Length)
                        {
                            int dgv_row = col_val.Length - tar_DGV.Rows.Count;
                            for (int k = 0; k < dgv_row; k++)
                            {
                                tar_DGV.Rows.Add();
                            }
                        }
                        for (int j = 0; j < col_val.Length; j++)
                        {
                            tar_DGV.Rows[j].Cells[col_name[i]].Value = col_val[j];
                            if ((SV <= SV + UL) && (SV >= SV + LL))
                            {
                                tar_DGV.Rows[j].Cells[col_name[i]].Style.BackColor = Color.White;
                            }
                            else
                            {
                                tar_DGV.Rows[j].Cells[col_name[i]].Style.BackColor = Color.Red;
                            }

                        }
                        if (!check_columns_existed(tbl_spec, col_name[i]))
                        {
                            tbl_spec.Columns.Add(col_name[i]);
                            tbl_spec.Rows[0][col_name[i]] = temp_arr[3];
                            tbl_spec.Rows[1][col_name[i]] = (SV + UL).ToString();
                            tbl_spec.Rows[2][col_name[i]] = (SV + LL).ToString();
                        }
                    }
                }
                else
                {
                    if (!src_data[i].Contains("FAI"))
                    {
                        string[] temp_arr = src_data[i].Split(',');
                        col_name[i] = temp_arr[0];
                        double SV = Convert.ToDouble(temp_arr[3]);
                        double UL = Convert.ToDouble(temp_arr[4]);
                        double LL = Convert.ToDouble(temp_arr[5]);
                        tar_DGV.Columns.Add(col_name[i], col_name[i]);
                        string[] col_val;
                        col_val = src_data[i].Split(',').Skip(6).ToArray();
                        if (tar_DGV.Rows.Count < col_val.Length)
                        {
                            int dgv_row = col_val.Length - tar_DGV.Rows.Count;
                            for (int k = 0; k < dgv_row; k++)
                            {
                                tar_DGV.Rows.Add();
                            }
                        }
                        for (int j = 0; j < col_val.Length; j++)
                        {
                            tar_DGV.Rows[j].Cells[col_name[i]].Value = col_val[j];
                            if ((SV <= SV + UL) && (SV >= SV + LL))
                            {
                                tar_DGV.Rows[j].Cells[col_name[i]].Style.BackColor = Color.White;
                            }
                            else
                            {
                                tar_DGV.Rows[j].Cells[col_name[i]].Style.BackColor = Color.Red;
                            }

                        }
                        if (!check_columns_existed(tbl_spec, col_name[i]))
                        {
                            tbl_spec.Columns.Add(col_name[i]);
                            tbl_spec.Rows[0][col_name[i]] = temp_arr[3];
                            tbl_spec.Rows[1][col_name[i]] = (SV + UL).ToString();
                            tbl_spec.Rows[2][col_name[i]] = (SV + LL).ToString();
                        }
                    }
                }
            }
            if (tar_DGV.ColumnCount > 0)
            {
                for (int k = 0; k < tar_DGV.RowCount; k++)
                {
                    tar_DGV.Rows[k].HeaderCell.Value = (k + 1).ToString();
                }
                tar_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            }
            DGV_Spec.DataSource = tbl_spec;
            if (DGV_Spec.Columns.Count > 0)
            {
                DGV_Spec.Rows[0].HeaderCell.Value = "SV";
                DGV_Spec.Rows[1].HeaderCell.Value = "UL";
                DGV_Spec.Rows[2].HeaderCell.Value = "LL";
                DGV_Spec.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                DGV_Spec.AutoResizeRows();
                DGV_Spec.AutoResizeColumns();
            }
            Disable_Sort_DGV(DGV_Spec);
            Disable_Sort_DGV(tar_DGV);
        }
        public DataTable DGV_To_Table(DataGridView src_DGV)
        {
            //DataTable _result = new DataTable();
            int col_inx = src_DGV.Columns.Count;
            string[] col_name = new string[col_inx];
            DataTable dt = new DataTable();
            for (int i = 0; i < col_inx; i++)
            {
                col_name[i] = src_DGV.Columns[i].Name;
                dt.Columns.Add(col_name[i]);
            }

            for (int i = 0; i < src_DGV.Rows.Count; i++)
            {
                DataRow dr = dt.NewRow();
                for (int j = 0; j < col_inx; j++)
                {
                    dr[j] = src_DGV.Rows[i].Cells[j].Value;
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
        public DataTable Get_data_detail(DataTable src_tbl, string col_name, string col_filter_name)
        {
            DataTable _result = new DataTable();
            string filter_cmd;
            if (col_name == "All")
            {
                filter_cmd = "";
            }
            else
            {
                filter_cmd = col_name + " = '" + col_filter_name + "'";
            }
            //filter_cmd = col_name + " = '" + col_filter_name + "'";
            _result = src_tbl.Select(filter_cmd).CopyToDataTable();
            return _result;
        }
        public void Disable_Sort_DGV(DataGridView sel_DGV)
        {
            foreach (DataGridViewColumn column in sel_DGV.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
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
        public bool IsNumeric(string text)
        {
            double test;
            return double.TryParse(text, out test);
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
        public void Mitutoyo_Data_display(string _filename, DataGridView tar_DGV, DataGridView DGV_Spec, bool in_FAI)
        {
            DataTable mytbl = new DataTable();
            DataTable tbl_spec = new DataTable();
            DataTable FAI_tbl = new DataTable();
            DataTable non_FAI = new DataTable();
            string[] col_name;
            mytbl = Mitutoyo_Get_raw_data(_filename);
            if (in_FAI)
            {
                FAI_tbl = mytbl.Select("FeatureLaber like '%FAI%'").CopyToDataTable();
            }
            else
            {
                FAI_tbl = mytbl.Select("FeatureLaber not like '%FAI%'").CopyToDataTable();
            }
            col_name = FAI_tbl.AsEnumerable().Select(r => r.Field<string>("FeatureLaber")).Distinct().ToArray();
            tar_DGV.Columns.Clear();
            if (DGV_Spec.Columns.Count > 0)
            {
                tbl_spec = DGV_To_Table(DGV_Spec);
            }
            foreach (string tg in col_name)
            {
                tar_DGV.Columns.Add(tg, tg);
                if (!check_columns_existed(tbl_spec, tg))
                {
                    tbl_spec.Columns.Add(tg);
                }
            }
            if (tbl_spec.Rows.Count == 0)
            {
                tbl_spec.Rows.Add();
                tbl_spec.Rows.Add();
                tbl_spec.Rows.Add();
            }
            foreach (string tg in col_name)
            {
                DataTable temp = new DataTable();
                temp = Get_data_detail(mytbl, "FeatureLaber", tg);
                int row_inx = temp.Rows.Count;
                if (tar_DGV.Rows.Count < row_inx)
                {
                    int dgv_row = row_inx - tar_DGV.Rows.Count;
                    for (int i = 0; i < dgv_row; i++)
                    {
                        tar_DGV.Rows.Add();
                    }
                }
                string[] item_val;
                string[] UL_val;
                string[] LL_val;
                string[] SetVal;
                double SV;
                double UL;
                double LL;
                item_val = temp.AsEnumerable().Select(r => r.Field<string>("Actual")).ToArray();
                SetVal = temp.AsEnumerable().Select(r => r.Field<string>("Nominal")).Distinct().ToArray();
                UL_val = temp.AsEnumerable().Select(r => r.Field<string>("Upper Tol")).Distinct().ToArray();
                LL_val = temp.AsEnumerable().Select(r => r.Field<string>("LowerTol")).Distinct().ToArray();
                SV = Convert.ToDouble(SetVal[0]);
                UL = Convert.ToDouble(SetVal[0]) + Convert.ToDouble(UL_val[0]);
                LL = Convert.ToDouble(SetVal[0]) + Convert.ToDouble(LL_val[0]);
                tbl_spec.Rows[0][tg] = SetVal[0];
                tbl_spec.Rows[1][tg] = UL.ToString();
                tbl_spec.Rows[2][tg] = LL.ToString();
                int j = 0;
                foreach (string c in item_val)
                {
                    tar_DGV.Rows[j].Cells[tg].Value = c;
                    if ((Convert.ToDouble(c) <= UL) && (Convert.ToDouble(c) >= LL))
                    {
                        tar_DGV.Rows[j].Cells[tg].Style.BackColor = Color.White;
                    }
                    else
                    {
                        tar_DGV.Rows[j].Cells[tg].Style.BackColor = Color.Red;
                    }
                    j++;
                }
            }
            for (int k = 0; k < tar_DGV.RowCount; k++)
            {
                tar_DGV.Rows[k].HeaderCell.Value = (k + 1).ToString();
            }
            tar_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            DGV_Spec.DataSource = tbl_spec;
            DGV_Spec.Rows[0].HeaderCell.Value = "SV";
            DGV_Spec.Rows[1].HeaderCell.Value = "UL";
            DGV_Spec.Rows[2].HeaderCell.Value = "LL";
            DGV_Spec.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            DGV_Spec.AutoResizeRows();
            DGV_Spec.AutoResizeColumns();
            Disable_Sort_DGV(DGV_Spec);
            Disable_Sort_DGV(tar_DGV);
        }
        public DataTable Mitutoyo_Get_raw_data(string _filename)
        {
            //DataTable _result = new DataTable();
            string tar_file;
            char[] split_char = { '\t', ',' };
            string[] src_data = new string[1000];
            tar_file = _filename;
            System.IO.StreamReader file = new System.IO.StreamReader(tar_file);
            string newline;
            int i = 0;
            while ((newline = file.ReadLine()) != null)
            {
                if (newline.Contains('"'))
                {
                    string[] temp = newline.Split('"');
                    temp[1] = temp[1].Split('[', ']')[0];
                    src_data[i] = string.Concat(temp);
                }
                else
                {
                    src_data[i] = newline;
                }
                i++;
            }
            Array.Resize<string>(ref src_data, i);
            string[] columnnames = src_data[0].Split(split_char);
            DataTable dt = new DataTable();
            foreach (string c in columnnames)
            {
                dt.Columns.Add(c);
            }
            for (int j = 1; j < src_data.Length; j++)
            {
                DataRow dr = dt.NewRow();
                string[] values = src_data[j].Split(split_char);
                for (int k = 0; k < values.Length; k++)
                {
                    dr[k] = values[k];
                }
                dt.Rows.Add(dr);
            }
            file.Close();
            return dt;

        }
       
    }
}
