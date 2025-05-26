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
using myExcel = Microsoft.Office.Interop.Excel;
using TDMK_SQL;
using System.Data.SqlClient;
using ZedGraph;
using OK2SHIP;

namespace OK2SHIP_Measurements
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
            public string sheetno;
            public string instrument;
            public FAI_Spec(string _FAI_Name, string _check_side, string _SetVal, string _UL, string _LL,string _sheetno, string _instrument)
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
            public  double stdev;
            public double mean;
            public double max;
            public  double min;
            public string CP;
            public  string CPKL;
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
        public static bool admin_mode = false;
        public Thread export_data;
        public static FrmFAI FAI_form;
        public static FrmHistogram Chart_Form;
        public static FrmIPQC_Man IPQC_Man_form;        
        public static string curr_user;
        //public static string scan_file;
        public static bool confirm_request = false;
        public static string operate_mode = "Auto";
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public List<FAI_Spec> testFAI_spec = new List<FAI_Spec>();  
        public DataTable Mitutoyo_Get_raw_data(string _filename)
        {
            //DataTable _result = new DataTable();
            string tar_file;
            char[] split_char = { '\t', ',' };
            string[] src_data = new string[10000];
            //tar_file = Path.Combine(app_path, "Original", _filename);
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
        public DataTable Mitutoyo_Get_raw_data2(string _filename)
        {
            //DataTable _result = new DataTable();
            string tar_file;
            char[] split_char = { '\t', ',' };
            string[] src_data = new string[10000];
            //tar_file = Path.Combine(app_path, "Original", _filename);
            tar_file = _filename;
            System.IO.StreamReader file = new System.IO.StreamReader(tar_file);
            string newline;
            int i = 0;
            while ((newline = file.ReadLine()) != null)
            {
                src_data[i] = newline.Replace("\"", "");
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
                List<string> list_val;
                list_val = values.ToList();
                list_val.RemoveAt(2);
                for (int k = 0; k < list_val.Count; k++)
                {
                    if (list_val[k].Contains("["))
                    {
                        dr[k] = list_val[k].Split('[')[0];
                    }
                    else
                    {
                        dr[k] = list_val[k];
                    }
                }
                dt.Rows.Add(dr);
            }
            file.Close();
            return dt;

        }
        public void Mitutoyo_Data_display(string _filename, DataGridView tar_DGV, DataGridView DGV_Spec, DataGridView DGV_plus, bool edit_mode)
        {
            DataTable mytbl = new DataTable();
            DataTable tbl_spec = new DataTable();
            DataTable FAI_tbl = new DataTable();
            DataTable non_FAI = new DataTable();
            mytbl = Mitutoyo_Get_raw_data2(_filename);
            string tar_Col = mytbl.Columns[1].ColumnName;
            if (!edit_mode)
            {
                try
                {
                    fill_datatable_DGV3(mytbl, tar_DGV, DGV_Spec,DGV_plus);                    
                }
                catch
                {

                }
            }
            else
            {
                DGV_plus.Columns.Clear();
                fill_datatable_DGV(mytbl, DGV_plus, DGV_Spec);
            }
            Disable_Sort_DGV(tar_DGV);
            Disable_Sort_DGV(DGV_plus);
            Disable_Sort_DGV(DGV_Spec);
            //if (!edit_mode)
            //{
            //    if (DGV_plus.Rows.Count>0)
            //    {
            //        SEI_Lib.confirm_request = true;
            //    }
            //}
        }
        public void fill_datatable_DGV(DataTable src_tbl, DataGridView tar_DGV, DataGridView tar_DGV_Spec)
        {
            string[] col_name;
            DataTable tbl_spec = new DataTable();
            DataTable tbl_DGV_data = new DataTable();
            int tar_DGV_rows_count;//= tar_DGV.Rows.Count;
            int curr_DGV_row;
            string tar_Col = src_tbl.Columns[1].ColumnName;
            string vitri_col_name = src_tbl.Columns[0].ColumnName;
            string vitri_act_val = src_tbl.Columns[4].ColumnName;
            string vitri_Set_val = src_tbl.Columns[5].ColumnName;
            string vitri_UL_val = src_tbl.Columns[7].ColumnName;
            string vitri_LL_val = src_tbl.Columns[8].ColumnName;
            col_name = src_tbl.AsEnumerable().Select(r => r.Field<string>(tar_Col)).Distinct().ToArray();
            tbl_DGV_data = DGV_To_Table(tar_DGV);
            tar_DGV_rows_count = tbl_DGV_data.Rows.Count;
            tbl_spec = DGV_To_Table(tar_DGV_Spec);
            //foreach (string tg in col_name)
            //{
            //    if (!check_columns_existed(tbl_DGV_data, tg))
            //    {
            //        tar_DGV.Columns.Add(tg, tg);
            //    }
            //}

            foreach (string tg in col_name)
            {
                //double UL = 0;// = Convert .ToDouble (tbl_spec.Rows [1][tg]);
                //double LL = 0;// Convert.ToDouble(tbl_spec.Rows[2][tg]);

                string UL = "";
                string LL = "";
                double SetVal=0;
                bool spec_check_en = false;
                int sel_col_inx = 0;
                string col_dgv_data = tg;
                if (check_columns_existed_inx(tbl_spec, tg, ref sel_col_inx))
                {
                    //double sv = Convert.ToDouble(tbl_spec.Rows[0][sel_col_inx]);
                    //UL = sv + Convert.ToDouble(tbl_spec.Rows[1][sel_col_inx]);
                    //LL = sv - Convert.ToDouble(tbl_spec.Rows[2][sel_col_inx]);


                    SetVal = Convert.ToDouble(tbl_spec.Rows[0][sel_col_inx]);
                    string _UL =IsNumeric_Val(checkDBNull(tbl_spec.Rows[1][sel_col_inx]));
                    string _LL =IsNumeric_Val(checkDBNull(tbl_spec.Rows[2][sel_col_inx]));

                    if (_UL != "")
                    {
                        UL = (SetVal + Convert.ToDouble(_UL)).ToString();
                    }
                    if (_LL != "")
                    {
                        LL = (SetVal - Convert.ToDouble(_LL)).ToString();
                    }


                    spec_check_en = true;
                    col_dgv_data = tbl_spec.Columns[sel_col_inx].ColumnName;
                    //if (!check_columns_existed(tbl_DGV_data, col_dgv_data))
                    //{
                    //    tar_DGV.Columns.Add(col_dgv_data, tg);
                    //}
                }
                else
                {
                    //tar_DGV.Columns.Add(tg, tg);
                    spec_check_en = false;
                }
                if (!check_columns_existed(tbl_DGV_data, col_dgv_data))
                {
                    tar_DGV.Columns.Add(col_dgv_data, tg);
                }
                DataTable temp = new DataTable();
                temp = Get_data_detail(src_tbl, tar_Col, tg);
                int row_inx = temp.Rows.Count;
                curr_DGV_row = tar_DGV.RowCount;
                int sel_row_inx = curr_DGV_row - tar_DGV_rows_count;
                if (sel_row_inx < row_inx)
                {
                    int dgv_row = row_inx - sel_row_inx;
                    for (int i = 0; i < dgv_row; i++)
                    {
                        tar_DGV.Rows.Add();
                    }
                }
                string[] item_val;
                string[] vitri_val;
                item_val = temp.AsEnumerable().Select(r => r.Field<string>(vitri_act_val)).ToArray();
                vitri_val = temp.AsEnumerable().Select(r => r.Field<string>(vitri_col_name)).ToArray();
                int j = tar_DGV_rows_count;
                int vitri_inx = 0;
                foreach (string c in item_val)
                {
                    if (vitri_val[vitri_inx].ToUpper().Contains("ANGLE"))
                    {
                        tar_DGV.Rows[j].Cells[col_dgv_data].Value = c;
                    }
                    else
                    {
                        tar_DGV.Rows[j].Cells[col_dgv_data].Value = Math.Abs(Convert.ToDouble(c)).ToString();
                    }
                    if (spec_check_en)
                    {
                        //if ((Convert.ToDouble(c) <= UL) && (Convert.ToDouble(c) >= LL))
                        //{
                        //    tar_DGV.Rows[j].Cells[col_dgv_data].Style.BackColor = Color.White;
                        //}
                        //else
                        //{
                        //    tar_DGV.Rows[j].Cells[col_dgv_data].Style.BackColor = Color.Red;
                        //}
                        tar_DGV.Rows[j].Cells[col_dgv_data].Style.BackColor = check_in_limit5(UL, LL, c, SetVal.ToString());
                    }
                    else

                    {
                        tar_DGV.Rows[j].Cells[col_dgv_data].Style.BackColor = Color.Yellow;
                    }

                    j++;
                    vitri_inx++;
                }
            }
            for (int k = 0; k < tar_DGV.RowCount; k++)
            {
                tar_DGV.Rows[k].HeaderCell.Value = (k + 1).ToString();
            }
            tar_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            tar_DGV.AutoResizeColumns();
            Disable_Sort_DGV(tar_DGV);
        }
        public void fill_datatable_DGV2(DataTable src_tbl, DataGridView tar_DGV, DataGridView tar_DGV_Spec)
        {
            string[] col_name;
            DataTable tbl_spec = new DataTable();
            DataTable tbl_DGV_data = new DataTable();
            int tar_DGV_rows_count;//= tar_DGV.Rows.Count;
            int curr_DGV_row;
            string tar_Col = src_tbl.Columns[1].ColumnName;
            string vitri_col_name = src_tbl.Columns[3].ColumnName;
            string vitri_act_val = src_tbl.Columns[4].ColumnName;
            string vitri_Set_val = src_tbl.Columns[5].ColumnName;
            string vitri_UL_val = src_tbl.Columns[7].ColumnName;
            string vitri_LL_val = src_tbl.Columns[8].ColumnName;
            col_name = src_tbl.AsEnumerable().Select(r => r.Field<string>(tar_Col)).Distinct().ToArray();
            tbl_DGV_data = DGV_To_Table(tar_DGV);
            tar_DGV_rows_count = tbl_DGV_data.Rows.Count;
            if (tar_DGV_Spec.Columns.Count > 0)
            {
                tbl_spec = DGV_To_Table(tar_DGV_Spec);
            }
            //tbl_spec = DGV_To_Table(tar_DGV_Spec);
            foreach (string tg in col_name)
            {
                if (!check_columns_existed(tbl_DGV_data, tg))
                {
                    tar_DGV.Columns.Add(tg, tg);
                }
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
                temp = Get_data_detail(src_tbl, tar_Col, tg);
                int row_inx = temp.Rows.Count;
                curr_DGV_row = tar_DGV.RowCount;
                int sel_row_inx = curr_DGV_row - tar_DGV_rows_count;
                if (sel_row_inx < row_inx)
                {
                    int dgv_row = row_inx - sel_row_inx;
                    for (int i = 0; i < dgv_row; i++)
                    {
                        tar_DGV.Rows.Add();
                    }
                }
                string[] item_val;
                string[] UL_val;
                string[] LL_val;
                string[] SetVal;
                string[] vitri_val;
                double SV;
                double UL;// = Convert.ToDouble(tbl_spec.Rows[1][tg]);
                double LL;// = Convert.ToDouble(tbl_spec.Rows[2][tg]); ;
                          //double UL = 0;// = Convert .ToDouble (tbl_spec.Rows [1][tg]);
                          // double LL = 0;// Convert.ToDouble(tbl_spec.Rows[2][tg]);
                          //bool spec_check_en = false;
                          //if (check_columns_existed(tbl_spec, tg))
                          //{
                          //    UL = Convert.ToDouble(tbl_spec.Rows[1][tg]);
                          //    LL = Convert.ToDouble(tbl_spec.Rows[2][tg]);
                          //    spec_check_en = true;
                          //}
                item_val = temp.AsEnumerable().Select(r => r.Field<string>(vitri_act_val)).ToArray();
                vitri_val = temp.AsEnumerable().Select(r => r.Field<string>(vitri_col_name)).ToArray();
                SetVal = temp.AsEnumerable().Select(r => r.Field<string>(vitri_Set_val)).Distinct().ToArray();
                UL_val = temp.AsEnumerable().Select(r => r.Field<string>(vitri_UL_val)).Distinct().ToArray();
                LL_val = temp.AsEnumerable().Select(r => r.Field<string>(vitri_LL_val)).Distinct().ToArray();
                SV = Math.Abs(Convert.ToDouble(SetVal[0]));
                UL = Math.Abs(Convert.ToDouble(SetVal[0]) + Convert.ToDouble(UL_val[0]));
                LL = Math.Abs(Convert.ToDouble(SetVal[0]) + Convert.ToDouble(LL_val[0]));
                tbl_spec.Rows[0][tg] = SetVal[0];
                tbl_spec.Rows[1][tg] = UL.ToString();
                tbl_spec.Rows[2][tg] = LL.ToString();
                int j = tar_DGV_rows_count;
                int vitri_inx = 0;
                foreach (string c in item_val)
                {
                    if (vitri_val[vitri_inx].ToUpper().Contains("ANGLE"))
                    {
                        tar_DGV.Rows[j].Cells[tg].Value = c;
                    }
                    else
                    {
                        tar_DGV.Rows[j].Cells[tg].Value = Math.Abs(Convert.ToDouble(c)).ToString();
                    }

                    if ((Convert.ToDouble(c) <= UL) && (Convert.ToDouble(c) >= LL))
                    {
                        tar_DGV.Rows[j].Cells[tg].Style.BackColor = Color.White;
                    }
                    else
                    {
                        tar_DGV.Rows[j].Cells[tg].Style.BackColor = Color.Red;
                    }


                    j++;
                    vitri_inx++;
                }
            }
            for (int k = 0; k < tar_DGV.RowCount; k++)
            {
                tar_DGV.Rows[k].HeaderCell.Value = (k + 1).ToString();
            }
            tar_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            tar_DGV.AutoResizeColumns();
            tar_DGV_Spec.DataSource = tbl_spec;
            tar_DGV_Spec.Rows[0].HeaderCell.Value = "SV";
            tar_DGV_Spec.Rows[1].HeaderCell.Value = "UL";
            tar_DGV_Spec.Rows[2].HeaderCell.Value = "LL";
            tar_DGV_Spec.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            tar_DGV_Spec.AutoResizeRows();
            tar_DGV_Spec.AutoResizeColumns();
            Disable_Sort_DGV(tar_DGV_Spec);
            Disable_Sort_DGV(tar_DGV);
        }
        public void Nikon_Data_display(string _filename, DataGridView tar_DGV, DataGridView DGV_Spec, DataGridView tar_DGV_Plus, bool edit_mode)
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
            int tar_DGV_row_inx = tar_DGV.RowCount;
            int DGV_plus_row_inx = tar_DGV_Plus.RowCount;
            int row_inx = 0;// = tar_DGV.RowCount;     // Vi tri bat dau insert du lieu
            int curr_DGV_count; // dem so dong cua DGV sau khi insert;
            DataTable tbl_spec = new DataTable();
            tbl_spec = DGV_To_Table(DGV_Spec);
            DataGridView sel_DGV = new DataGridView();
            for (i = 0; i < col_inx; i++)
            {
                if (edit_mode)
                {
                    sel_DGV = tar_DGV_Plus;
                }
                else
                {
                    if (src_data[i].Contains("FAI"))
                    {
                        sel_DGV = tar_DGV;
                        row_inx = tar_DGV_row_inx;
                    }
                    else
                    {
                        sel_DGV = tar_DGV_Plus;
                        row_inx = DGV_plus_row_inx;
                    }
                }
                string[] temp_arr = src_data[i].Split(',');
                col_name[i] = temp_arr[0];
                //double UL = 0;// = Convert.ToDouble(temp_arr[4]);
                //double LL = 0;// = Convert.ToDouble(temp_arr[5]);

                string UL = "";
                string LL = "";
                double SetVal = 0;
                bool spec_check_en = false;
                //if (!check_columns_existed(DGV_To_Table(sel_DGV), col_name[i]))
                //{
                //    sel_DGV.Columns.Add(col_name[i], col_name[i]);
                //}
                DataTable tbl_DGV_data = DGV_To_Table(sel_DGV);
                int sel_col_inx = 0;
                string tg = col_name[i];
                string col_dgv_data = col_name[i];
                if (check_columns_existed_inx(tbl_spec, tg, ref sel_col_inx))
                {
                    col_dgv_data = tbl_spec.Columns[sel_col_inx].ColumnName;
                    //double sv = Convert.ToDouble(tbl_spec.Rows[0][sel_col_inx]);
                    //UL = sv + Convert.ToDouble(tbl_spec.Rows[1][sel_col_inx]);
                    //LL = sv - Convert.ToDouble(tbl_spec.Rows[2][sel_col_inx]);


                    SetVal = Convert.ToDouble(tbl_spec.Rows[0][sel_col_inx]);
                    string _UL = IsNumeric_Val(checkDBNull(tbl_spec.Rows[1][sel_col_inx]));
                    string _LL = IsNumeric_Val(checkDBNull(tbl_spec.Rows[2][sel_col_inx]));

                    if (_UL != "")
                    {
                        UL = (SetVal + Convert.ToDouble(_UL)).ToString();
                    }
                    if (_LL != "")
                    {
                        LL = (SetVal - Convert.ToDouble(_LL)).ToString();
                    }

                    spec_check_en = true;
                }
                if (!check_columns_existed(tbl_DGV_data, col_dgv_data))
                {
                    sel_DGV.Columns.Add(col_dgv_data, col_name[i]);
                }
                string[] col_val;
                col_val = src_data[i].Split(',').Skip(6).ToArray();
                curr_DGV_count = sel_DGV.RowCount - row_inx;
                if (curr_DGV_count < col_val.Length)
                {
                    int dgv_row;
                    dgv_row = col_val.Length - curr_DGV_count;
                    for (int k = 0; k < dgv_row; k++)
                    {
                        sel_DGV.Rows.Add();
                    }
                }
                for (int j = 0; j < col_val.Length; j++)
                {
                    //sel_DGV.Rows[row_inx + j].Cells[col_name[i]].Value = col_val[j];
                    sel_DGV.Rows[row_inx + j].Cells[col_dgv_data].Value = col_val[j];
                    //double act_val = Convert.ToDouble(col_val[j]);
                    string act_val = checkDBNull(col_val[j]);
                    if (spec_check_en)
                    {
                        //if ((act_val <= UL) && (act_val >= LL))
                        //{
                        //    //sel_DGV.Rows[row_inx + j].Cells[col_name[i]].Style.BackColor = Color.White;
                        //    sel_DGV.Rows[row_inx + j].Cells[col_dgv_data].Style.BackColor = Color.White;
                        //}
                        //else
                        //{
                        //    //sel_DGV.Rows[row_inx + j].Cells[col_name[i]].Style.BackColor = Color.Red;
                        //    sel_DGV.Rows[row_inx + j].Cells[col_dgv_data].Style.BackColor = Color.Red;
                        //}
                        sel_DGV.Rows[row_inx + j].Cells[col_dgv_data].Style.BackColor = check_in_limit5(UL, LL, act_val, SetVal.ToString());
                    }
                    else
                    {
                        //sel_DGV.Rows[row_inx + j].Cells[col_name[i]].Style.BackColor = Color.Yellow;
                        sel_DGV.Rows[row_inx + j].Cells[col_dgv_data].Style.BackColor = Color.LightGray;
                    }
                }
            }
            if (tar_DGV.ColumnCount > 0)
            {
                //DataTable tbl_DGV_data = DGV_To_Table(tar_DGV);
                //foreach (DataGridViewColumn dgv_col in tar_DGV.Columns)
                //{
                //    int sel_col_inx = 0;
                //    string sel_col_name = dgv_col.Name;
                //    if (check_columns_existed_inx(tbl_spec, sel_col_name, ref sel_col_inx))
                //    {
                //        dgv_col.Name = tbl_spec.Columns[sel_col_inx].ColumnName;
                //    }                  
                //}
                for (int k = 0; k < tar_DGV.RowCount; k++)
                {
                    tar_DGV.Rows[k].HeaderCell.Value = (k + 1).ToString();
                }
                tar_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            }
            if (tar_DGV_Plus.ColumnCount > 0)
            {
                for (int k = 0; k < tar_DGV_Plus.RowCount; k++)
                {
                    tar_DGV_Plus.Rows[k].HeaderCell.Value = (k + 1).ToString();
                }
                tar_DGV_Plus.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            }            
            Disable_Sort_DGV(DGV_Spec);
            Disable_Sort_DGV(tar_DGV);
            Disable_Sort_DGV(tar_DGV_Plus);

        }
        public void Nikon_Data_display3(string _filename, DataGridView tar_DGV, DataGridView DGV_Spec, DataGridView tar_DGV_Plus, bool edit_mode)
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
            int tar_DGV_row_inx = tar_DGV.RowCount;
            int DGV_plus_row_inx = tar_DGV_Plus.RowCount;
            int row_inx = 0;// = tar_DGV.RowCount;     // Vi tri bat dau insert du lieu
            int curr_DGV_count; // dem so dong cua DGV sau khi insert;
            DataTable tbl_spec = new DataTable();
            tbl_spec = DGV_To_Table(DGV_Spec);
            DataGridView sel_DGV = new DataGridView();
            DataTable tbl_DGV = DGV_To_Table(tar_DGV);
            DataTable tbl_DGV_Plus = DGV_To_Table(tar_DGV_Plus);
            DataTable sel_tbl = new DataTable();
            for (i = 0; i < col_inx; i++)
            {
                string[] temp_arr = src_data[i].Split(',');
                col_name[i] = temp_arr[0];
                //double UL = 0;// = Convert.ToDouble(temp_arr[4]);
                //double LL = 0;// = Convert.ToDouble(temp_arr[5]);
                string UL = "";
                string LL = "";
                double SetVal = 0;
                bool spec_check_en = false;
                //
                int sel_col_inx = 0;
                string tg = col_name[i];
                string col_dgv_data = col_name[i];
                if (check_columns_existed_inx(tbl_spec, tg, ref sel_col_inx))
                {
                    col_dgv_data = tbl_spec.Columns[sel_col_inx].ColumnName;
                    //double sv = Convert.ToDouble(tbl_spec.Rows[0][sel_col_inx]);
                    //UL = sv + Convert.ToDouble(tbl_spec.Rows[1][sel_col_inx]);
                    //LL = sv - Convert.ToDouble(tbl_spec.Rows[2][sel_col_inx]);


                    SetVal = Convert.ToDouble(tbl_spec.Rows[0][sel_col_inx]);
                    string _UL = IsNumeric_Val(checkDBNull(tbl_spec.Rows[1][sel_col_inx]));
                    string _LL = IsNumeric_Val(checkDBNull(tbl_spec.Rows[2][sel_col_inx]));

                    if (_UL != "")
                    {
                        UL = (SetVal + Convert.ToDouble(_UL)).ToString();
                    }
                    if (_LL != "")
                    {
                        LL = (SetVal - Convert.ToDouble(_LL)).ToString();
                    }


                    if (!edit_mode)
                    {
                        sel_tbl = tbl_DGV;
                        sel_DGV = tar_DGV;
                    }
                    else
                    {
                        sel_tbl = tbl_DGV_Plus;
                        sel_DGV = tar_DGV_Plus;
                    }
                    spec_check_en = true;
                }
                else
                {
                    sel_tbl = tbl_DGV_Plus;
                    sel_DGV = tar_DGV_Plus;
                    spec_check_en = false;
                }
                if (!check_columns_existed(sel_tbl, col_dgv_data))
                {
                    sel_DGV.Columns.Add(col_dgv_data, col_name[i]);
                }
                string[] col_val;
                col_val = src_data[i].Split(',').Skip(6).ToArray();
                row_inx = sel_tbl.Rows.Count;
                curr_DGV_count = sel_DGV.RowCount - row_inx;
                if (curr_DGV_count < col_val.Length)
                {
                    int dgv_row;
                    dgv_row = col_val.Length - curr_DGV_count;
                    for (int k = 0; k < dgv_row; k++)
                    {
                        sel_DGV.Rows.Add();
                    }
                }
                for (int j = 0; j < col_val.Length; j++)
                {
                    if (checkDBNull(col_val[j]) != "")
                    {
                        sel_DGV.Rows[row_inx + j].Cells[col_dgv_data].Value = col_val[j];
                        //double act_val = Convert.ToDouble(col_val[j]);
                        string act_val = checkDBNull(col_val[j]);
                        if (spec_check_en)
                        {

                            //if ((act_val <= UL) && (act_val >= LL))
                            //{
                            //    sel_DGV.Rows[row_inx + j].Cells[col_dgv_data].Style.BackColor = Color.White;
                            //}
                            //else
                            //{
                            //    if (act_val > UL)
                            //    {
                            //        sel_DGV.Rows[row_inx + j].Cells[col_dgv_data].Style.BackColor = Color.Yellow;
                            //    }
                            //    if (act_val< LL)
                            //    {
                            //        sel_DGV.Rows[row_inx + j].Cells[col_dgv_data].Style.BackColor = Color.LightBlue;
                            //    }
                            //    //confirm_request = true;
                            //}
                            sel_DGV.Rows[row_inx + j].Cells[col_dgv_data].Style.BackColor = check_in_limit5(UL, LL, act_val, SetVal.ToString());
                        }
                        else
                        {
                            sel_DGV.Rows[row_inx + j].Cells[col_dgv_data].Style.BackColor = Color.LightGray;
                        }
                    }
                }
            }
            DGV_Auto_Resize(tar_DGV);
            DGV_Auto_Resize(tar_DGV_Plus);
        }
        public void Nikon_Data_display2(string _filename, DataGridView tar_DGV, DataGridView DGV_Spec, DataGridView tar_DGV_Plus, bool edit_mode)
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
            int tar_DGV_row_inx = tar_DGV.RowCount;
            int DGV_plus_row_inx = tar_DGV_Plus.RowCount;
            int row_inx = 0;// = tar_DGV.RowCount;     // Vi tri bat dau insert du lieu
            int curr_DGV_count; // dem so dong cua DGV sau khi insert;
            DataTable tbl_spec = new DataTable();
            tbl_spec = DGV_To_Table(DGV_Spec);
            if (tbl_spec.Rows.Count == 0)
            {
                tbl_spec.Rows.Add();
                tbl_spec.Rows.Add();
                tbl_spec.Rows.Add();
            }
            DataGridView sel_DGV = new DataGridView();
            for (i = 0; i < col_inx; i++)
            {
                if (edit_mode)
                {
                    sel_DGV = tar_DGV_Plus;
                }
                else
                {
                    if (src_data[i].Contains("FAI"))
                    {
                        sel_DGV = tar_DGV;
                        row_inx = tar_DGV_row_inx;
                    }
                    else
                    {
                        sel_DGV = tar_DGV_Plus;
                        row_inx = DGV_plus_row_inx;
                    }
                }
                string[] temp_arr = src_data[i].Split(',');
                col_name[i] = temp_arr[0];
                double SV = Convert.ToDouble(temp_arr[3]);
                double UL = Convert.ToDouble(temp_arr[4]);
                double LL = Convert.ToDouble(temp_arr[5]);
                if (!check_columns_existed(DGV_To_Table(sel_DGV), col_name[i]))
                {
                    sel_DGV.Columns.Add(col_name[i], col_name[i]);
                }
                string[] col_val;
                col_val = src_data[i].Split(',').Skip(6).ToArray();
                curr_DGV_count = sel_DGV.RowCount - row_inx;
                if (curr_DGV_count < col_val.Length)
                {
                    int dgv_row;
                    dgv_row = col_val.Length - curr_DGV_count;
                    for (int k = 0; k < dgv_row; k++)
                    {
                        sel_DGV.Rows.Add();
                    }
                }
                for (int j = 0; j < col_val.Length; j++)
                {
                    sel_DGV.Rows[row_inx + j].Cells[col_name[i]].Value = col_val[j];
                    if ((SV <= SV + UL) && (SV >= SV + LL))
                    {
                        sel_DGV.Rows[row_inx + j].Cells[col_name[i]].Style.BackColor = Color.White;
                    }
                    else
                    {
                        sel_DGV.Rows[row_inx + j].Cells[col_name[i]].Style.BackColor = Color.Red;
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
            if (tar_DGV.ColumnCount > 0)
            {
                for (int k = 0; k < tar_DGV.RowCount; k++)
                {
                    tar_DGV.Rows[k].HeaderCell.Value = (k + 1).ToString();
                }
                tar_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            }
            if (tar_DGV_Plus.ColumnCount > 0)
            {
                for (int k = 0; k < tar_DGV_Plus.RowCount; k++)
                {
                    tar_DGV_Plus.Rows[k].HeaderCell.Value = (k + 1).ToString();
                }
                tar_DGV_Plus.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
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
            Disable_Sort_DGV(tar_DGV_Plus);
        }
        public void Nikon_Data_display_old(string _filename, DataGridView tar_DGV, DataGridView DGV_Spec, bool in_FAI)
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
            DGV_Spec.Columns.Clear();
            //int sel_inx = 0;
            for (i = 0; i < col_inx; i++)
            {
                string[] temp_arr = src_data[i].Split(',');
                col_name[i] = temp_arr[0];
                Setval[i] = temp_arr[3];
                double SV = Convert.ToDouble(temp_arr[3]);
                double UL = Convert.ToDouble(temp_arr[4]);
                double LL = Convert.ToDouble(temp_arr[5]);
                UL_val[i] = (SV + UL).ToString();
                LL_val[i] = (SV + LL).ToString();
                tar_DGV.Columns.Add(col_name[i], col_name[i]);
                DGV_Spec.Columns.Add(col_name[i], col_name[i]);
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
            }
            if (tar_DGV.ColumnCount > 0)
            {
                for (int k = 0; k < tar_DGV.RowCount; k++)
                {
                    tar_DGV.Rows[k].HeaderCell.Value = (k + 1).ToString();
                }
                tar_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            }
            if (DGV_Spec.ColumnCount > 0)
            {
                DGV_Spec.Rows.Add(Setval);
                DGV_Spec.Rows.Add(UL_val);
                DGV_Spec.Rows.Add(LL_val);
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
            DataTable dt = new DataTable();
            if (src_DGV.DataSource==null)
            {
                int col_inx = src_DGV.Columns.Count;
                string[] col_name = new string[col_inx];
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
            }
            else
            {
                dt = (DataTable)src_DGV.DataSource;
            }
            return dt;
        }
        public void Keyence_Display_Data(string _filename, DataGridView tar_DGV_Data, DataGridView DGV_Spec, DataGridView tar_DGV_plus, bool edit_mode)
        {
            string tar_file;
            char[] split_char = { '\t', ',' };
            string[] src_data = new string[1000];
            //tar_file = Path.Combine(app_path, "Original", _filename);
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
            DataGridView sel_DGV;
            DataTable orig_tbl = new DataTable();
            if (!edit_mode)
            {
                sel_DGV = tar_DGV_Data;
            }
            else
            {
                sel_DGV = tar_DGV_plus;
                //sel_DGV.Columns.Clear();
            }
            orig_tbl = DGV_To_Table(sel_DGV);
            DataTable tbl_spec = new DataTable();
            tbl_spec = DGV_To_Table(DGV_Spec);
            //DGV_Spec.Columns.Clear();
            Array.Resize<string>(ref src_data, i);
            string[] columnnames = src_data[0].Split(split_char);
            DataTable dt = new DataTable();
            foreach (string c in columnnames)
            {
                string mycol_name;
                if (c.Contains("]"))
                {
                    string[] temp = c.Split(']');
                    //dt.Columns.Add(temp[1].Replace(" ", "").ToUpper());
                    mycol_name = temp[1].Replace(" ", "").ToUpper();
                }
                else
                {
                    //dt.Columns.Add(c);
                    mycol_name = c;
                }
                int sel_col_inx = 0;
                if (check_columns_existed_inx(tbl_spec, mycol_name, ref sel_col_inx))
                {
                    dt.Columns.Add(tbl_spec.Columns[sel_col_inx].ColumnName);
                }
                else
                {
                    dt.Columns.Add(mycol_name);
                }
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
            DataTable tbl2 = new DataTable();

            tbl2 = Get_data_detail(dt, "[OK/NG]", "OK");
            for (int k = 1; k < 7; k++)
            {
                tbl2.Columns.Remove(columnnames[k]);
            }
            tbl2.Columns.RemoveAt(0);
            orig_tbl.Merge(tbl2);
            if (edit_mode)
            {
                sel_DGV.DataSource = tbl2;
            }
            else
            {
                sel_DGV.DataSource = orig_tbl;
            }
            double UL = 0;
            double LL = 0;
            bool spec_check_en = false;
            for (int k = 0; k < sel_DGV.ColumnCount; k++)
            {
                string src_col = sel_DGV.Columns[k].Name.ToString();
                sel_DGV.Columns[k].HeaderText = src_col.Split('_')[0];
                if (check_columns_existed(tbl_spec, src_col))
                {
                    double SV = Convert.ToDouble(tbl_spec.Rows[0][src_col]);
                    UL = SV + Convert.ToDouble(tbl_spec.Rows[1][src_col]);
                    LL = SV- Convert.ToDouble(tbl_spec.Rows[2][src_col]);
                    spec_check_en = true;
                }
                else
                {
                    spec_check_en = false;
                }
                for (int t = 0; t < sel_DGV.RowCount; t++)
                {
                    string sel_curr_cellval = checkDBNull(sel_DGV.Rows[t].Cells[k].Value);
                    if (sel_curr_cellval!="")
                    {
                        double act_val = Convert.ToDouble(sel_DGV.Rows[t].Cells[k].Value);
                        if (spec_check_en)
                        {
                            if ((act_val <= UL) && (act_val >= LL))
                            {
                                sel_DGV.Rows[t].Cells[k].Style.BackColor = Color.White;
                            }
                            else
                            {
                                sel_DGV.Rows[t].Cells[k].Style.BackColor = Color.Red;
                            }
                        }
                        else
                        {
                            sel_DGV.Rows[t].Cells[k].Style.BackColor = Color.Yellow;
                        }
                    }
                }
            }
            for (int k = 0; k < sel_DGV.RowCount; k++)
            {
                sel_DGV.Rows[k].HeaderCell.Value = (k + 1).ToString();
            }
            sel_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            sel_DGV.ColumnHeadersDefaultCellStyle.Font = new Font(sel_DGV.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold);
            sel_DGV.AutoResizeColumns();
            sel_DGV.AutoResizeRows();
            Disable_Sort_DGV(sel_DGV);            
        }
        public void Keyence_Display_Data3(string _filename, DataGridView tar_DGV_Data, DataGridView DGV_Spec, DataGridView tar_DGV_plus, bool edit_mode)
        {
            DataTable tbl_spec = DGV_To_Table(DGV_Spec);
            DataTable tbldata = Keyence_RawData_Process(_filename, tbl_spec);
            Fill_DGV_Keyence(tbldata, tar_DGV_Data, tbl_spec, tar_DGV_plus, edit_mode);
        }
        public DataTable Keyence_RawData_Process(string _filename, DataTable src_tbl_spec)
        {
            DataTable _result = new DataTable();
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
                string mycol_name;
                if (c.Contains("]"))
                {
                    string[] temp = c.Split(']');
                    mycol_name = temp[1].Replace(" ", "").ToUpper();
                }
                else
                {
                    mycol_name = c;
                }
                int sel_col_inx = 0;
                if (check_columns_existed_inx(src_tbl_spec, mycol_name, ref sel_col_inx))
                {
                    dt.Columns.Add(src_tbl_spec.Columns[sel_col_inx].ColumnName);
                }
                else
                {
                    dt.Columns.Add(mycol_name);
                }
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
            DataTable tbldata;// = new DataTable();           
            tbldata = Get_data_detail(dt, "[OK/NG]", "OK");
            for (int k = 1; k < 7; k++)
            {
                tbldata.Columns.Remove(columnnames[k]);
            }
            tbldata.Columns.RemoveAt(0);
            _result = tbldata;
            return _result;
        }
        public void Fill_DGV_Keyence(DataTable src_Keyence_tbl, DataGridView tar_DGV_Data, DataTable src_tbl_spec, DataGridView tar_DGV_plus, bool edit_mode)
        {
            DataGridView sel_DGV;
            if (edit_mode)
            {
                tar_DGV_plus.DataSource = src_Keyence_tbl;
                sel_DGV = tar_DGV_plus;
            }
            else
            {
                int FAI_inx = 0;
                int noneFAI_inx = 0;
                int arr_size = src_Keyence_tbl.Columns.Count;
                string[] FAI_col_arr = new string[arr_size];
                string[] noneFAI_col_arr = new string[arr_size];
                foreach (DataColumn c in src_Keyence_tbl.Columns)
                {
                    if (check_columns_existed(src_tbl_spec, c.ColumnName))
                    {
                        FAI_col_arr[FAI_inx] = c.ColumnName;
                        FAI_inx++;
                    }
                    else
                    {
                        noneFAI_col_arr[noneFAI_inx] = c.ColumnName;
                        noneFAI_inx++;
                    }
                }
                Array.Resize(ref FAI_col_arr, FAI_inx);
                Array.Resize(ref noneFAI_col_arr, noneFAI_inx);
                DataTable FAI_tbl;// = new DataTable();
                DataTable noneFAI_tbl;// = new DataTable();
                FAI_tbl = src_Keyence_tbl.AsDataView().ToTable(false, FAI_col_arr);
                noneFAI_tbl = src_Keyence_tbl.AsDataView().ToTable(false, noneFAI_col_arr);
                DataTable tbl_DGV = DGV_To_Table(tar_DGV_Data);
                DataTable tbl_DGV_Plus = DGV_To_Table(tar_DGV_plus);
                tbl_DGV.Merge(FAI_tbl);
                tbl_DGV_Plus.Merge(noneFAI_tbl);
                tar_DGV_Data.DataSource = tbl_DGV;
                tar_DGV_plus.DataSource = tbl_DGV_Plus;
                sel_DGV = tar_DGV_Data;
                for (int r = 0; r < tar_DGV_plus.Rows.Count; r++)
                {
                    tar_DGV_plus.Rows[r].HeaderCell.Value = (r + 1).ToString();
                    foreach (DataGridViewColumn c in tar_DGV_plus.Columns)
                    {
                        tar_DGV_plus.Rows[r].Cells[c.Index].Style.BackColor = Color.Yellow;
                    }
                }
                DGV_Auto_Resize(tar_DGV_plus);
            }
            //double UL = 0;
            //double LL = 0;
            string UL = "";
            string LL = "";
            double SetVal = 0;
            bool spec_check_en = false;
            for (int k = 0; k < sel_DGV.ColumnCount; k++)
            {
                string src_col = sel_DGV.Columns[k].Name.ToString();
                if (check_columns_existed(src_tbl_spec, src_col))
                {
                    //double SV = Convert.ToDouble(src_tbl_spec.Rows[0][src_col]);
                    //UL = SV + Convert.ToDouble(src_tbl_spec.Rows[1][src_col]);
                    //LL = SV - Convert.ToDouble(src_tbl_spec.Rows[2][src_col]);


                    SetVal = Convert.ToDouble(src_tbl_spec.Rows[0][src_col]);
                    string _UL = IsNumeric_Val(checkDBNull(src_tbl_spec.Rows[1][src_col]));
                    string _LL = IsNumeric_Val(checkDBNull(src_tbl_spec.Rows[2][src_col]));

                    if (_UL != "")
                    {
                        UL = (SetVal + Convert.ToDouble(_UL)).ToString();
                    }
                    if (_LL != "")
                    {
                        LL = (SetVal - Convert.ToDouble(_LL)).ToString();
                    }


                    sel_DGV.Columns[k].HeaderText = src_col.Split('_')[0];
                    spec_check_en = true;
                }
                else
                {
                    spec_check_en = false;
                }
                for (int t = 0; t < sel_DGV.RowCount; t++)
                {
                    string sel_curr_cellval = checkDBNull(sel_DGV.Rows[t].Cells[k].Value);
                    if (sel_curr_cellval != "")
                    {
                        //double act_val = Convert.ToDouble(sel_DGV.Rows[t].Cells[k].Value);
                        string act_val = checkDBNull(sel_DGV.Rows[t].Cells[k].Value);
                        if (spec_check_en)
                        {
                            //if ((act_val <= UL) && (act_val >= LL))
                            //{
                            //    sel_DGV.Rows[t].Cells[k].Style.BackColor = Color.White;
                            //}
                            //else
                            //{
                            //    if (act_val > UL)
                            //    {
                            //        sel_DGV.Rows[t].Cells[k].Style.BackColor = Color.Yellow;
                            //    }
                            //    if (act_val < LL)
                            //    {
                            //        sel_DGV.Rows[t].Cells[k].Style.BackColor = Color.LightBlue;
                            //    }
                            //    //confirm_request = true;
                            //}
                            sel_DGV.Rows[t].Cells[k].Style.BackColor = check_in_limit5(UL, LL, act_val, SetVal.ToString());
                        }
                        else
                        {
                            sel_DGV.Rows[t].Cells[k].Style.BackColor = Color.LightGray;
                        }
                    }
                }
            }
            DGV_Auto_Resize(sel_DGV);
        }
        public void Keyence_Display_Data2(string _filename, DataGridView tar_DGV_Data, DataGridView DGV_Spec, DataGridView tar_DGV_plus, bool edit_mode)
        {
            string tar_file;
            char[] split_char = { '\t', ',' };
            string[] src_data = new string[1000];
            //tar_file = Path.Combine(app_path, "Original", _filename);
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
            DataGridView sel_DGV;
            DataTable orig_tbl = new DataTable();
            if (!edit_mode)
            {
                sel_DGV = tar_DGV_Data;
            }
            else
            {
                sel_DGV = tar_DGV_plus;
                //sel_DGV.Columns.Clear();
            }
            orig_tbl = DGV_To_Table(sel_DGV);
            //DGV_Spec.Columns.Clear();
            Array.Resize<string>(ref src_data, i);
            string[] columnnames = src_data[0].Split(split_char);
            DataTable dt = new DataTable();
            foreach (string c in columnnames)
            {
                if (c.Contains("]"))
                {
                    string[] temp = c.Split(']');
                    dt.Columns.Add(temp[1].Replace(" ", "").ToUpper());
                }
                else
                {
                    dt.Columns.Add(c);
                }
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
            DataTable tbl2 = new DataTable();
            DataTable tbl_spec = new DataTable();
            tbl2 = Get_data_detail(dt, "[OK/NG]", "OK");
            tbl_spec = Get_data_detail(dt, "[OK/NG]", "");
            for (int k = 1; k < 7; k++)
            {
                tbl2.Columns.Remove(columnnames[k]);
                tbl_spec.Columns.Remove(columnnames[k]);
            }
            DataRow dr_spec_SetVal = tbl_spec.NewRow();
            DataRow dr_spec_ul = tbl_spec.NewRow();
            DataRow dr_spec_ll = tbl_spec.NewRow();
            for (int k = 0; k < tbl_spec.Columns.Count; k++)
            {
                dr_spec_SetVal[k] = tbl_spec.Rows[5][k];
                dr_spec_ul[k] = tbl_spec.Rows[6][k];
                dr_spec_ll[k] = tbl_spec.Rows[7][k];
            }
            tbl_spec.Rows.Clear();
            tbl_spec.Rows.Add(dr_spec_SetVal);
            tbl_spec.Rows.Add(dr_spec_ul);
            tbl_spec.Rows.Add(dr_spec_ll);
            tbl_spec.Columns.RemoveAt(0);
            tbl2.Columns.RemoveAt(0);
            orig_tbl.Merge(tbl2);
            if (edit_mode)
            {
                sel_DGV.DataSource = tbl2;
            }
            else
            {
                sel_DGV.DataSource = orig_tbl;
            }
            for (int k = 0; k < sel_DGV.ColumnCount; k++)
            {
                string src_col = sel_DGV.Columns[k].Name.ToString();
                double UL = Convert.ToDouble(tbl_spec.Rows[1][src_col]);
                double LL = Convert.ToDouble(tbl_spec.Rows[2][src_col]);
                for (int t = 0; t < sel_DGV.RowCount; t++)
                {
                    double act_val = Convert.ToDouble(sel_DGV.Rows[t].Cells[k].Value);
                    if ((act_val <= UL) && (act_val >= LL))
                    {
                        sel_DGV.Rows[t].Cells[k].Style.BackColor = Color.White;
                    }
                    else
                    {
                        sel_DGV.Rows[t].Cells[k].Style.BackColor = Color.Red;
                    }
                }

            }
            for (int k = 0; k < sel_DGV.RowCount; k++)
            {
                sel_DGV.Rows[k].HeaderCell.Value = (k + 1).ToString();
            }
            sel_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            sel_DGV.ColumnHeadersDefaultCellStyle.Font = new Font(sel_DGV.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold);
            sel_DGV.AutoResizeColumns();
            sel_DGV.AutoResizeRows();
            DGV_Spec.DataSource = tbl_spec;
            DGV_Spec.ColumnHeadersDefaultCellStyle.Font = new Font(DGV_Spec.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold);
            DGV_Spec.Rows[0].HeaderCell.Value = "SV";
            DGV_Spec.Rows[1].HeaderCell.Value = "UL";
            DGV_Spec.Rows[2].HeaderCell.Value = "LL";
            DGV_Spec.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            DGV_Spec.AutoResizeRows();
            DGV_Spec.AutoResizeColumns();
            Disable_Sort_DGV(DGV_Spec);
            Disable_Sort_DGV(sel_DGV);
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
        public void Disable_Sort_DGV(DataGridView sel_DGV)
        {
            foreach (DataGridViewColumn column in sel_DGV.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
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
        public void Export_To_FAI(string app_path ,string _ItemCode, string _LotNo, myExcel.Workbook src_data_wrkbook)
        {
            myExcel.Workbook format_wrk = TDMK_Code.open_excel_file(Path.Combine(app_path, "Format"), _ItemCode + _LotNo + "_Format.xlsx", "");
            format_wrk.SaveAs(Path.Combine(app_path, "Format", _ItemCode + _LotNo + "_temp.xlsx"));
            myExcel.Worksheet sel_wrksheet;
            myExcel.Worksheet data_sheet = src_data_wrkbook.Sheets[1];
            myExcel.Range data_rgn = data_sheet.Range["B1"];
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
                        while (checkDBNull(data_sheet.Range["B1"].Offset[0, j].Value) != "")
                        {
                            string data_rgn_val = checkDBNull(data_sheet.Range["B1"].Offset[0, j].Value);
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
        public void Load_Test_Data(string format_loc, string tar_ItemCode, string tar_LotNo, string file_extension, DataGridView tar_DGV_Data)
        {
            string tar_format = tar_ItemCode + tar_LotNo;
            string tar_format_file = Path.Combine(format_loc, tar_format + file_extension);
            string[] spec_item_str = new string[13];
            spec_item_str[0] = "SetVal";
            spec_item_str[1] = "UL";
            spec_item_str[2] = "LL";
            spec_item_str[3] = "CheckSide";
            spec_item_str[4] = "STDEV";
            spec_item_str[5] = "Mean";
            spec_item_str[6] = "MAX";
            spec_item_str[7] = "MIN";
            spec_item_str[8] = "Cp";
            spec_item_str[9] = "Cpkl";
            spec_item_str[10] = "Cpku";
            spec_item_str[11] = "Cpk";
            spec_item_str[12] = "Cpkm";
            if (File.Exists(tar_format_file))
            {
                myExcel.Workbook tar_wkbook = TDMK_Code.open_excel_file(tar_format_file, "", "");
                int wrksheet_num = tar_wkbook.Worksheets.Count;
                int spec_col_inx = 0;
                foreach (myExcel.Worksheet tg in tar_wkbook.Worksheets)
                {
                    if (tg.Name.Contains("FAI"))
                    {
                        myExcel.Range sel_rgn = tg.Range["C14"];
                        myExcel.Range NormDim_rgn = tg.Range["C15"];
                        myExcel.Range data_rgn = tg.Range["C255"];
                        int sel_inx = 0;
                        while (checkDBNull(sel_rgn.Offset[0, sel_inx].Value) != "")
                        {
                            string t_FAIName = checkDBNull(sel_rgn.Offset[0, sel_inx].Value).Replace(" ", "");
                            string col_dgv_name = t_FAIName + "_" + NormDim_rgn.Offset[0, sel_inx].Value;
                            if (!check_columns_existed(DGV_To_Table(tar_DGV_Data), col_dgv_name))
                            {
                                tar_DGV_Data.Columns.Add(col_dgv_name, t_FAIName);
                                int i = 0;
                                while (checkDBNull(data_rgn.Offset[i, sel_inx].Value) != "")
                                {
                                    if (tar_DGV_Data.RowCount < i + 1)
                                    {
                                        tar_DGV_Data.Rows.Add();
                                        tar_DGV_Data.Rows[i].HeaderCell.Value = (i + 1).ToString();
                                    }
                                    tar_DGV_Data.Rows[i].Cells[col_dgv_name].Value = data_rgn.Offset[i, sel_inx].Value;
                                    i++;
                                }
                            }
                            sel_inx++;
                        }
                    }
                }
                tar_DGV_Data.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                Disable_Sort_DGV(tar_DGV_Data);                
                MessageBox.Show(new Form { TopMost = true },"Finished!", "Thong bao");
                //this.Activate();
                //tar_wkbook.Close();
                //TDMK_Code.releaseObject(tar_wkbook);

            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },"File format khong tim thay");
            }
        }
        public void Export_Data_Process(string app_path, DataGridView tar_DGV, string tar_ItemCode, string tar_LotNo)
        {
            myExcel.Workbook tar_wrkbook;
            myExcel.Worksheet tar_wrksheet;
            myExcel.Range FAI_No_rgn;
            myExcel.Range FAI_data_rgn;
            string format_path = Path.Combine(app_path, "Format");
            string temp_path = Path.Combine(app_path, "Temp_file");
            tar_wrkbook = TDMK_Code.open_excel_file(format_path, tar_ItemCode + tar_LotNo + ".xlsm", "");
            tar_wrksheet = tar_wrkbook.Sheets["FAI (1)"];
            FAI_No_rgn = tar_wrksheet.Range["C17"];
            string[] FAI_lst = new string[100];
            int col_inx = 0;
            while (checkDBNull(FAI_No_rgn.Offset[0, col_inx].Value) != "")
            {
                string sel_FAI_val;
                sel_FAI_val = checkDBNull(FAI_No_rgn.Offset[0, col_inx].Value);
                FAI_lst[col_inx] = sel_FAI_val.Replace(" ", "");
                col_inx++;
            }
            Array.Resize<string>(ref FAI_lst, col_inx);
            tar_wrkbook.SaveAs(Path.Combine(temp_path, Create_date_string(DateTime.Now) + "_" + tar_ItemCode + "_" + tar_LotNo + ".xlsm"));
            DataTable result_FAI = DGV_To_Table(tar_DGV);
            int offset_val = 0;
            foreach (string c in FAI_lst)
            {
                if (check_columns_existed(result_FAI, c))
                {
                    FAI_data_rgn = tar_wrksheet.Range["C259"].Offset[0, offset_val];
                    for (int k = 0; k < result_FAI.Rows.Count; k++)
                    {
                        FAI_data_rgn.Offset[k, 0].Value = result_FAI.Rows[k][c].ToString();
                    }
                }
                offset_val++;
            }
            MessageBox.Show(new Form { TopMost = true },"Hoàn thành xuất dữ liệu", "Thông báo");
        }
        public void Load_FAI_DGV(DataGridView tar_DGV, SqlConnection tar_sqlcon, string tar_tbl, string FAI_No_name, string FAI_Data_Col_name, string tar_ItemCode, string tar_LotNo)
        {
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
            DataTable tbl_data = new DataTable();
            tbl_data = DGV_To_Table(tar_DGV);
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
            tar_DGV.DataSource = tbl_data;
            foreach (DataGridViewColumn t in tar_DGV.Columns)
            {
                t.HeaderText = t.HeaderText.Split('_')[0];
            }
            Disable_Sort_DGV(tar_DGV);
        }
        public void Load_FAI_DGV2(DataGridView tar_DGV, SqlConnection tar_sqlcon, string tar_tbl, string FAI_No_name, string FAI_Data_Col_name, string tar_ItemCode, string tar_LotNo, string format_type)
        {
            string[] flt_items = new string[5];
            string[] flt_item_vals = new string[5];
            AutoCompleteStringCollection FAI_lst;
            //AutoCompleteStringCollection SheetNo_lst;
            flt_items[0] = "ItemCode";
            flt_items[1] = "LotNo";
            flt_items[2] = "FAI_No";
            flt_items[3] = "SheetNo";
            flt_items[4] = "Remark";
            flt_item_vals[0] = tar_ItemCode;
            flt_item_vals[4] = format_type;
            //SheetNo_lst = TDMK_Code.Load_Item_Names_Filter(tar_sqlcon, "FAI_Spec", "SheetNo", TDMK_Code.filter_str(flt_items, flt_item_vals));

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
                    flt_item_vals[2] = c;
                    flt_item_vals[3] = "";
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
            }
            tar_DGV.DataSource = tbl_data;
            foreach (DataGridViewColumn t in tar_DGV.Columns)
            {
                t.HeaderText = t.HeaderText.Split('_')[0];
            }
            Disable_Sort_DGV(tar_DGV);
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
        public bool Load_Spec(SqlConnection sqlcon, string format_loc, string tar_ItemCode, string tar_LotNo, string file_extension, DataGridView tar_DGV_Spec, string tar_Dev)
        {
            bool _result = false;
            testFAI_spec.Clear();
            //AutoCompleteStringCollection FAI_SheetNo_list = new AutoCompleteStringCollection();
            //FAI_SheetNo_list = TDMK_Code.Load_Item_Names_Filter(sqlcon, "FAI_Spec", "SheetNo", TDMK_Code.filter_str(new string[] { "ItemCode", "Instrument","SheetNo" }, new string[] { tar_ItemCode, tar_Dev,"%FAI%" }));
            DataTable FAI_spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Spec", TDMK_Code.filter_str(new string[] { "ItemCode", "Instrument"}, new string[] { tar_ItemCode, tar_Dev }));
            List<string> search_key_lst = new List<string>() { "FAI", "Dimension" };
            List<string> FAI_SheetNo_list = FAI_spec_dt.AsEnumerable().Where(x => search_key_lst.Any(y => x.Field<string>("SheetNo").Contains(y))).Select(x => x.Field<string>("SheetNo")).Distinct().ToList();
            if (FAI_SheetNo_list.Count == 0)
            {
                MessageBox.Show(new Form { TopMost = true },"Spec not found!", "Warning");
                _result = false;
            }
            else
            {
                string[] items = new string[6];
                string[] items_val = new string[6];
                items[0] = "ItemCode";
                items[1] = "LotNo";
                items[2] = "FAI_No";
                items[3] = "Instrument";
                items[4] = "NormDim";
                items[5] = "SheetNo";
                items_val[0] = tar_ItemCode;
                items_val[3] = tar_Dev;
                DataTable cur_dt = new DataTable();
                int spec_col_inx = 0;
                foreach (string sht in FAI_SheetNo_list)
                {
                    items_val[5] = sht;
                    items_val[2] = "";
                    AutoCompleteStringCollection FAI_No_lst = TDMK_Code.Load_Item_Filter_str(sqlcon, "FAI_Spec", "FAI_No", TDMK_Code.filter_str(items, items_val));
                    foreach (string t in FAI_No_lst)
                    {
                        items_val[2] = t;
                        //string c = TDMK_Code.Get_item_val("FAI_Spec", TDMK_Code.filter_str(items, items_val), "NormDim", sqlcon);
                        string col_val = t;// + "_" + c;
                        int dgv_row = tar_DGV_Spec.Rows.Count;
                        if (!check_columns_existed(cur_dt, col_val))
                        {

                            cur_dt.Columns.Add(col_val);
                        }
                        if (cur_dt.Rows.Count == 0)
                        {
                            for (int i = 0; i < 4; i++)
                            {
                                cur_dt.Rows.Add();
                            }
                        }
                        DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Spec", TDMK_Code.filter_str(items, items_val));
                        cur_dt.Rows[0][col_val] = spec_dt.Rows[0]["NormDim"];
                        cur_dt.Rows[1][col_val] = spec_dt.Rows[0]["TolMax"];
                        cur_dt.Rows[2][col_val] = spec_dt.Rows[0]["TolMin"];//Math .Round ( TolMin,3).ToString  ();//                         
                        cur_dt.Rows[3][col_val] = spec_dt.Rows[0]["Distribution"];
                        string t_checkside = spec_dt.Rows[0]["Distribution"].ToString(); // checkDBNull(sel_rgn.Offset[-1, sel_inx].Value);
                        string t_FAIName = t; //checkDBNull(sel_rgn.Offset[0, sel_inx].Value).Replace(" ", "");
                        string t_FAI_Setval = spec_dt.Rows[0]["NormDim"].ToString(); // checkDBNull(sel_rgn.Offset[1, sel_inx].Value);
                        string t_FAI_UL = spec_dt.Rows[0]["TolMax"].ToString(); ;/// checkDBNull(sel_rgn.Offset[6, sel_inx].Value);
                        string t_FAI_LL = spec_dt.Rows[0]["TolMin"].ToString(); // checkDBNull(sel_rgn.Offset[7, sel_inx].Value);
                        string t_FAI_sheetno = spec_dt.Rows[0]["SheetNo"].ToString();
                        string t_FAI_instrument = spec_dt.Rows[0]["Instrument"].ToString();
                        testFAI_spec.Add(new SEI_Lib.FAI_Spec(col_val, t_checkside, t_FAI_Setval, t_FAI_UL, t_FAI_LL, t_FAI_sheetno, t_FAI_instrument));
                        spec_col_inx++;
                    }
                }
                if (cur_dt.Columns.Count > 0)
                {
                    tar_DGV_Spec.DataSource = cur_dt;
                    foreach (DataGridViewColumn t_col in tar_DGV_Spec.Columns)
                    {
                        t_col.HeaderText = t_col.HeaderText.Split('_')[0];
                    }
                    if (tar_DGV_Spec.Rows.Count > 0)
                    {
                        tar_DGV_Spec.Rows[0].HeaderCell.Value = "NormDim";
                        tar_DGV_Spec.Rows[1].HeaderCell.Value = "Tol_Max";
                        tar_DGV_Spec.Rows[2].HeaderCell.Value = "Tol_Min";
                        tar_DGV_Spec.Rows[3].HeaderCell.Value = "Distribution";
                        tar_DGV_Spec.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                        Disable_Sort_DGV(tar_DGV_Spec);
                    }
                }

                //Array.Resize<SEI_Lib.FAI_Spec>(ref testFAI_spec, spec_col_inx);
                _result = true;
            }
            return _result;
        }
        public bool Load_Spec(SqlConnection sqlcon, string format_loc, string tar_ItemCode, string tar_LotNo, string file_extension, DataGridView tar_DGV_Spec, string tar_Dev, string format_type)
        {
            bool _result = false;
            testFAI_spec.Clear();
            //AutoCompleteStringCollection FAI_SheetNo_list = new AutoCompleteStringCollection();
            //FAI_SheetNo_list = TDMK_Code.Load_Item_Names_Filter(sqlcon, "FAI_Spec", "SheetNo", TDMK_Code.filter_str(new string[] { "ItemCode", "Instrument","SheetNo" }, new string[] { tar_ItemCode, tar_Dev,"%FAI%" }));
            DataTable FAI_spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Spec", TDMK_Code.filter_str(new string[] { "ItemCode", "Instrument", "Remark" }, new string[] { tar_ItemCode, tar_Dev, format_type }));
            List<string> search_key_lst = new List<string>() { "FAI", "Dimension" };
            List<string> FAI_SheetNo_list = FAI_spec_dt.AsEnumerable().Where(x => search_key_lst.Any(y => x.Field<string>("SheetNo").Contains(y))).Select(x => x.Field<string>("SheetNo")).Distinct().ToList();
            if (FAI_SheetNo_list.Count == 0)
            {
                MessageBox.Show(new Form { TopMost = true }, "Spec not found!", "Warning");
                _result = false;
            }
            else
            {
                string[] items = new string[7];
                string[] items_val = new string[7];
                items[0] = "ItemCode";
                items[1] = "LotNo";
                items[2] = "FAI_No";
                items[3] = "Instrument";
                items[4] = "NormDim";
                items[5] = "SheetNo";
                items[6] = "Remark";
                items_val[0] = tar_ItemCode;
                items_val[3] = tar_Dev;
                items_val[6] = format_type;
                DataTable cur_dt = new DataTable();
                int spec_col_inx = 0;
                foreach (string sht in FAI_SheetNo_list)
                {
                    items_val[5] = sht;
                    items_val[2] = "";
                    AutoCompleteStringCollection FAI_No_lst = TDMK_Code.Load_Item_Filter_str(sqlcon, "FAI_Spec", "FAI_No", TDMK_Code.filter_str(items, items_val));
                    foreach (string t in FAI_No_lst)
                    {
                        items_val[2] = t;
                       
                        //string c = TDMK_Code.Get_item_val("FAI_Spec", TDMK_Code.filter_str(items, items_val), "NormDim", sqlcon);
                        string col_val = t;// + "_" + c;
                        int dgv_row = tar_DGV_Spec.Rows.Count;
                        if (!check_columns_existed(cur_dt, col_val))
                        {

                            cur_dt.Columns.Add(col_val);
                        }
                        if (cur_dt.Rows.Count == 0)
                        {
                            for (int i = 0; i < 4; i++)
                            {
                                cur_dt.Rows.Add();
                            }
                        }
                        DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Spec", TDMK_Code.filter_str(items, items_val));
                        cur_dt.Rows[0][col_val] = spec_dt.Rows[0]["NormDim"];
                        cur_dt.Rows[1][col_val] = spec_dt.Rows[0]["TolMax"];
                        cur_dt.Rows[2][col_val] = spec_dt.Rows[0]["TolMin"];//Math .Round ( TolMin,3).ToString  ();//                         
                        cur_dt.Rows[3][col_val] = spec_dt.Rows[0]["Distribution"];
                        string t_checkside = spec_dt.Rows[0]["Distribution"].ToString(); // checkDBNull(sel_rgn.Offset[-1, sel_inx].Value);
                        string t_FAIName = t; //checkDBNull(sel_rgn.Offset[0, sel_inx].Value).Replace(" ", "");
                        string t_FAI_Setval = spec_dt.Rows[0]["NormDim"].ToString(); // checkDBNull(sel_rgn.Offset[1, sel_inx].Value);
                        string t_FAI_UL = spec_dt.Rows[0]["TolMax"].ToString(); ;/// checkDBNull(sel_rgn.Offset[6, sel_inx].Value);
                        string t_FAI_LL = spec_dt.Rows[0]["TolMin"].ToString(); // checkDBNull(sel_rgn.Offset[7, sel_inx].Value);
                        string t_FAI_sheetno = spec_dt.Rows[0]["SheetNo"].ToString();
                        string t_FAI_instrument = spec_dt.Rows[0]["Instrument"].ToString();
                        testFAI_spec.Add(new SEI_Lib.FAI_Spec(col_val, t_checkside, t_FAI_Setval, t_FAI_UL, t_FAI_LL, t_FAI_sheetno, t_FAI_instrument));
                        spec_col_inx++;
                    }
                }
                if (cur_dt.Columns.Count > 0)
                {
                    tar_DGV_Spec.DataSource = cur_dt;
                    foreach (DataGridViewColumn t_col in tar_DGV_Spec.Columns)
                    {
                        t_col.HeaderText = t_col.HeaderText.Split('_')[0];
                    }
                    if (tar_DGV_Spec.Rows.Count > 0)
                    {
                        tar_DGV_Spec.Rows[0].HeaderCell.Value = "NormDim";
                        tar_DGV_Spec.Rows[1].HeaderCell.Value = "Tol_Max";
                        tar_DGV_Spec.Rows[2].HeaderCell.Value = "Tol_Min";
                        tar_DGV_Spec.Rows[3].HeaderCell.Value = "Distribution";
                        tar_DGV_Spec.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                        Disable_Sort_DGV(tar_DGV_Spec);
                    }
                }

                //Array.Resize<SEI_Lib.FAI_Spec>(ref testFAI_spec, spec_col_inx);
                _result = true;
            }
            return _result;
        }
        public DataTable Load_FAI_Spec_ToTable(SqlConnection sqlcon, string tar_ItemCode,string format_type)
        {
            DataTable cur_dt = new DataTable();
            DataTable FAI_spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Spec", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { tar_ItemCode, format_type }));
            List<string> search_key_lst = new List<string>() { "FAI", "Dimension" };
            List<string> FAI_SheetNo_list = FAI_spec_dt.AsEnumerable().Where(x => search_key_lst.Any(y => x.Field<string>("SheetNo").Contains(y))).Select(x => x.Field<string>("SheetNo")).Distinct().ToList();
            if (FAI_SheetNo_list.Count == 0)
            {
                MessageBox.Show(new Form { TopMost = true }, "Spec not found!", "Warning");
            }
            else
            {
                List<DataTable> dt_lst = new List<DataTable>();
                Get_ListTable(-1, FAI_spec_dt, new string[] { "SheetNo", "FAI_No" }, ref dt_lst, "Distribution");
                foreach(DataTable dt in dt_lst)
                {
                    if(dt.Rows.Count>0)
                    {
                        string sheetno = checkDBNull(dt.Rows[0]["SheetNo"]);
                        if(FAI_SheetNo_list.IndexOf(sheetno)!=-1)
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
        public bool Load_Spec_IPQC(SqlConnection tar_sqlcon, string spec_tbl_name, string tar_ItemCode, string tar_itemcheck, DataGridView tar_DGV_Spec, string format_loc, string file_extension)
        {
            bool _result = false;
 start_lbl: DataTable Spec_tbl = TDMK_Code.Datatable_Filter(tar_sqlcon, spec_tbl_name, TDMK_Code.filter_str(new string[] { "ItemCode", "Item", "Remark" }, new string[] { tar_ItemCode, "%" + tar_itemcheck + "%", "%AU_Plating%" })); //Roughness AU_NI            
            if (Spec_tbl.Rows.Count == 0)
            {
                string tar_format_file;
                string[] file_format_lst = Directory.GetFiles(format_loc, "*" + file_extension).Where(s => s.Contains(tar_ItemCode)).ToArray();
                if (file_format_lst.Length > 0)
                {
                    tar_format_file = file_format_lst[0];
                    Setup_IPQC_Spec(tar_format_file, tar_ItemCode, tar_sqlcon);
                    DataTable tbl_spec = TDMK_Code.Datatable_Filter(tar_sqlcon, spec_tbl_name, TDMK_Code.filter_str(new string[] { "ItemCode", "Item", "Remark" }, new string[] { tar_ItemCode, "%" + tar_itemcheck + "%", "%AU_Plating%" }));
                    if (tbl_spec.Rows.Count>0)
                    {
                        _result = true;
                        goto start_lbl;
                    }
                    else
                    {
                        _result = false;
                        MessageBox.Show(new Form { TopMost = true },"File Format not found and no data of ItemCode "+ tar_ItemCode, "Warning");
                    }
                }
                else
                {
                    _result = false;
                    MessageBox.Show(new Form { TopMost = true },"No Spec data in database");
                }                                 
            }
            else
            {
                IPQC_Spec_Process(Spec_tbl, tar_DGV_Spec);
                _result = true;
            }
            return _result;
        }
        public bool Load_Spec_IPQC2(SqlConnection tar_sqlcon, string spec_tbl_name, string tar_ItemCode, string tar_itemcheck, DataGridView tar_DGV_Spec, string format_loc, string file_extension)
        {
            bool _result = false;
        start_lbl: DataTable Spec_tbl = TDMK_Code.Datatable_Filter(tar_sqlcon, spec_tbl_name, TDMK_Code.filter_str(new string[] { "ItemCode", "Item", "Remark" }, new string[] { tar_ItemCode, "%" + tar_itemcheck + "%", "%AU_Plating%" })); //Roughness AU_NI            
            if (Spec_tbl.Rows.Count == 0)
            {
                string tar_format_file;
                string[] file_format_lst = Directory.GetFiles(format_loc, "*" + file_extension).Where(s => s.Contains(tar_ItemCode)).ToArray();
                if (file_format_lst.Length > 0)
                {
                    tar_format_file = file_format_lst[0];
                    Setup_IPQC_Spec(tar_format_file, tar_ItemCode, tar_sqlcon);
                    DataTable tbl_spec = TDMK_Code.Datatable_Filter(tar_sqlcon, spec_tbl_name, TDMK_Code.filter_str(new string[] { "ItemCode", "Item", "Remark" }, new string[] { tar_ItemCode, "%" + tar_itemcheck + "%", "%AU_Plating%" }));
                    if (tbl_spec.Rows.Count > 0)
                    {
                        _result = true;
                        goto start_lbl;
                    }
                    else
                    {
                        _result = false;
                        MessageBox.Show(new Form { TopMost = true },"File Format not found and no data of ItemCode " + tar_ItemCode, "Warning");
                    }
                }
                else
                {
                    _result = false;
                    MessageBox.Show(new Form { TopMost = true },"No Spec data in database");
                }
            }
            else
            {
                IPQC_Spec_Process2(Spec_tbl, tar_DGV_Spec);
                _result = true;
            }
            return _result;
        }
        public void IPQC_Spec_Process(DataTable src_Spec_tbl, DataGridView tar_DGV_Spec)
        {
            string[] items_list = src_Spec_tbl.AsEnumerable().Select(r => r.Field<string>("Item")).ToArray();
            string[] limit_lst = src_Spec_tbl.AsEnumerable().Select(r => r.Field<string>("Spec")).ToArray();
            char[] split_chr = { '~', '<', '\u2264', '\u2265', '\u00B1', '>' };
            char[] trim_chr = { 'u', 'm', '\u03BC', '\u00B5', ' ', '=' };
            DataTable mySpec_dgv = new DataTable();
            foreach (string t in items_list)
            {
                mySpec_dgv.Columns.Add(t);
            }
            DataRow dr_SV = mySpec_dgv.NewRow();
            DataRow dr_UL = mySpec_dgv.NewRow();
            DataRow dr_LL = mySpec_dgv.NewRow();
            int inx = 0;
            foreach (string _limit in limit_lst)
            {
                if (_limit.Contains('/'))
                {
                    dr_SV[inx] = _limit.Split('-')[0].Trim(trim_chr);
                    string[] temp_limit = _limit.Split('-')[1].Trim(trim_chr).Split('/');
                    double sv = Convert.ToDouble(dr_SV[inx]);
                    double ul = Convert.ToDouble(temp_limit[1].Trim(trim_chr));
                    double ll = Convert.ToDouble(temp_limit[0].Trim(trim_chr));
                    dr_UL[inx] = (sv + ul).ToString();
                    dr_LL[inx] = (sv - ll).ToString();
                }
                else
                {
                    string[] UL_LL = _limit.Split(split_chr);
                    if (_limit.Contains('\u2265'))
                    {
                        dr_SV[inx] = "";
                        dr_UL[inx] = UL_LL[0].Trim(trim_chr);
                        dr_LL[inx] = UL_LL[1].Trim(trim_chr);
                    }
                    else
                    {
                        if (_limit.Contains('\u00B1'))
                        {
                            dr_SV[inx] = UL_LL[0];
                            dr_UL[inx] = (Convert.ToDouble(UL_LL[0].Trim(trim_chr)) + Convert.ToDouble(UL_LL[1].Trim(trim_chr))).ToString();
                            dr_LL[inx] = (Convert.ToDouble(UL_LL[0].Trim(trim_chr)) - Convert.ToDouble(UL_LL[1].Trim(trim_chr))).ToString();
                        }
                        else
                        {
                            dr_SV[inx] = "";
                            dr_LL[inx] = UL_LL[0].Trim(trim_chr);
                            dr_UL[inx] = UL_LL[1].Trim(trim_chr);
                        }
                    }
                }
                inx++;
            }
            mySpec_dgv.Rows.Add(dr_SV);
            mySpec_dgv.Rows.Add(dr_UL);
            mySpec_dgv.Rows.Add(dr_LL);
            tar_DGV_Spec.DataSource = mySpec_dgv;
            tar_DGV_Spec.Rows[0].HeaderCell.Value = "SV";
            tar_DGV_Spec.Rows[1].HeaderCell.Value = "UL";
            tar_DGV_Spec.Rows[2].HeaderCell.Value = "LL";
            tar_DGV_Spec.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            tar_DGV_Spec.AutoResizeColumns();
            Disable_Sort_DGV(tar_DGV_Spec);
        }
        public void IPQC_Spec_Process2(DataTable src_Spec_tbl, DataGridView tar_DGV_Spec)
        {
            string[] items_list = src_Spec_tbl.AsEnumerable().Select(r => r.Field<string>("Item")).ToArray();
            string[] limit_lst = src_Spec_tbl.AsEnumerable().Select(r => r.Field<string>("Spec")).ToArray();
            char[] split_chr = { '~', '<', '\u2264', '\u2265', '\u00B1', '>' };
            char[] trim_chr = { 'u', 'm', '\u03BC', '\u00B5', ' ', '=' };
            DataTable mySpec_dgv = new DataTable();
            foreach (string t in items_list)
            {
                mySpec_dgv.Columns.Add(t);
            }
            DataRow dr_SV = mySpec_dgv.NewRow();
            DataRow dr_UL = mySpec_dgv.NewRow();
            DataRow dr_LL = mySpec_dgv.NewRow();
            int inx = 0;
            foreach (string _limit in limit_lst)
            {
                if (_limit.Contains('/'))
                {
                    dr_SV[inx] = _limit.Split('-')[0].Trim(trim_chr);
                    string[] temp_limit = _limit.Split('-')[1].Trim(trim_chr).Split('/');
                    double sv = Convert.ToDouble(dr_SV[inx]);
                    double ul = Convert.ToDouble(temp_limit[1].Trim(trim_chr));
                    double ll = Convert.ToDouble(temp_limit[0].Trim(trim_chr));
                    dr_UL[inx] = (sv + ul).ToString();
                    dr_LL[inx] = (sv - ll).ToString();
                }
                else
                {
                    //string[] UL_LL = _limit.Split(split_chr);
                    //if (_limit.Contains('>'))
                    //{
                    //    dr_SV[inx] = "";
                    //    dr_LL[inx] = _limit;
                    //}
                    //else
                    //{
                    //    if (_limit.Contains('\u00B1'))
                    //    {
                    //        dr_SV[inx] = UL_LL[0];
                    //        dr_UL[inx] = (Convert.ToDouble(UL_LL[0].Trim(trim_chr)) + Convert.ToDouble(UL_LL[1].Trim(trim_chr))).ToString();
                    //        dr_LL[inx] = (Convert.ToDouble(UL_LL[0].Trim(trim_chr)) - Convert.ToDouble(UL_LL[1].Trim(trim_chr))).ToString();
                    //    }
                    //    else
                    //    {
                    //        dr_SV[inx] = "";                            
                    //        dr_UL[inx] = _limit;
                    //    }
                    //}

                    string[] UL_LL = _limit.Split(split_chr);
                    if (_limit.Contains('>'))
                    {
                        dr_SV[inx] = _limit;
                        dr_UL[inx] = UL_LL[0].Trim(trim_chr);
                        dr_LL[inx] = UL_LL[1].Trim(trim_chr);
                    }
                    else
                    {
                        if (_limit.Contains('\u00B1'))
                        {
                            dr_SV[inx] = UL_LL[0];
                            dr_UL[inx] = (Convert.ToDouble(UL_LL[0].Trim(trim_chr)) + Convert.ToDouble(UL_LL[1].Trim(trim_chr))).ToString();
                            dr_LL[inx] = (Convert.ToDouble(UL_LL[0].Trim(trim_chr)) - Convert.ToDouble(UL_LL[1].Trim(trim_chr))).ToString();
                        }
                        else
                        {
                            dr_SV[inx] = _limit;
                            dr_LL[inx] = UL_LL[0].Trim(trim_chr);
                            dr_UL[inx] = UL_LL[1].Trim(trim_chr);
                        }
                    }

                }
                inx++;
            }
            mySpec_dgv.Rows.Add(dr_SV);
            mySpec_dgv.Rows.Add(dr_UL);
            mySpec_dgv.Rows.Add(dr_LL);
            tar_DGV_Spec.DataSource = mySpec_dgv;
            tar_DGV_Spec.Rows[0].HeaderCell.Value = "SV";
            tar_DGV_Spec.Rows[1].HeaderCell.Value = "UL";
            tar_DGV_Spec.Rows[2].HeaderCell.Value = "LL";
            tar_DGV_Spec.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            tar_DGV_Spec.AutoResizeColumns();
            Disable_Sort_DGV(tar_DGV_Spec);
        }
        public void IPQC_Spec_Process_Man(DataTable src_Spec_tbl, DataGridView tar_DGV_Spec, string[] tar_col_name, string [] col_header)
        {
            //string[] items_list = src_Spec_tbl.AsEnumerable().Select(r => r.Field<string>("Item")).ToArray();
            string[] limit_lst = src_Spec_tbl.AsEnumerable().Select(r => r.Field<string>("Spec")).ToArray();
            string[] colname_list = src_Spec_tbl.AsEnumerable().Select(r => r.Field<string>("Col_Name")).ToArray();
            char[] split_chr = { '~', '<', '\u2264', '\u2265', '\u00B1', '>' };
            char[] trim_chr = { 'u', 'm', '\u03BC', '\u00B5', ' ', '=' };
            DataTable mySpec_dgv = new DataTable();
            int inx = 0;
            foreach (string t in tar_col_name)
            {

                mySpec_dgv.Columns.Add(t);
                string[] temp = src_Spec_tbl.AsEnumerable().Where(r => r.Field<string>("Col_Name") == t).Select(r => r.Field<string>("Item")).ToArray();
                string[] temp2 = src_Spec_tbl.AsEnumerable().Where(r => r.Field<string>("Col_Name") == t).Select(r => r.Field<string>("Spec")).ToArray();
                limit_lst[inx] = temp2[0];
                inx++;
            }
            Array.Resize(ref limit_lst, inx);
            DataRow dr_SV = mySpec_dgv.NewRow();
            DataRow dr_UL = mySpec_dgv.NewRow();
            DataRow dr_LL = mySpec_dgv.NewRow();
            inx = 0;
            foreach (string _limit in limit_lst)
            {
                if (_limit.Contains('/'))
                {
                    dr_SV[inx] = _limit.Split('-')[0].Trim(trim_chr);
                    string[] temp_limit = _limit.Split('-')[1].Trim(trim_chr).Split('/');
                    double sv = Convert.ToDouble(dr_SV[inx]);
                    double ul = Convert.ToDouble(temp_limit[1].Trim(trim_chr));
                    double ll = Convert.ToDouble(temp_limit[0].Trim(trim_chr));
                    dr_UL[inx] = (sv + ul).ToString();
                    dr_LL[inx] = (sv - ll).ToString();
                }
                else
                {
                    string[] UL_LL = _limit.Split(split_chr);
                    if (_limit.Contains('\u2265'))
                    {
                        dr_SV[inx] = "";
                        dr_UL[inx] = UL_LL[0].Trim(trim_chr);
                        dr_LL[inx] = UL_LL[1].Trim(trim_chr);
                    }
                    else
                    {
                        if (_limit.Contains('\u00B1'))
                        {
                            dr_SV[inx] = UL_LL[0];
                            dr_UL[inx] = (Convert.ToDouble(UL_LL[0].Trim(trim_chr)) + Convert.ToDouble(UL_LL[1].Trim(trim_chr))).ToString();
                            dr_LL[inx] = (Convert.ToDouble(UL_LL[0].Trim(trim_chr)) - Convert.ToDouble(UL_LL[1].Trim(trim_chr))).ToString();
                        }
                        else
                        {
                            dr_SV[inx] = "";
                            dr_LL[inx] = UL_LL[0].Trim(trim_chr);
                            dr_UL[inx] = UL_LL[1].Trim(trim_chr);
                        }
                    }
                }
                inx++;
            }
            mySpec_dgv.Rows.Add(dr_SV);
            mySpec_dgv.Rows.Add(dr_UL);
            mySpec_dgv.Rows.Add(dr_LL);
            tar_DGV_Spec.Columns.Clear();
            tar_DGV_Spec.DataSource = mySpec_dgv;
            tar_DGV_Spec.Rows[0].HeaderCell.Value = "SV";
            tar_DGV_Spec.Rows[1].HeaderCell.Value = "UL";
            tar_DGV_Spec.Rows[2].HeaderCell.Value = "LL";
            int c_inx = 0;
            foreach (DataGridViewColumn c in tar_DGV_Spec.Columns)
            {
                c.HeaderText = col_header[c_inx];
                c_inx++;
            }
            tar_DGV_Spec.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            tar_DGV_Spec.AutoResizeColumns();
            Disable_Sort_DGV(tar_DGV_Spec);
            tar_DGV_Spec.EditMode = DataGridViewEditMode.EditProgrammatically;
        }

        public void IPQC_Spec_Process_Man2(DataTable src_Spec_tbl, DataGridView tar_DGV_Spec)
        {
            string[] limit_lst = src_Spec_tbl.AsEnumerable().Select(r => r.Field<string>("Spec")).ToArray();
            string[] colname_list = src_Spec_tbl.AsEnumerable().Select(r => r.Field<string>("Col_Name")).ToArray();
            string[] col_header = new string[tar_DGV_Spec.Columns.Count];
            char[] split_chr = { '~', '<', '\u2264', '\u2265', '\u00B1', '>' };
            char[] trim_chr = { 'u', 'm', '\u03BC', '\u00B5', ' ', '=','-' };
            DataTable mySpec_dgv = DGV_To_Table(tar_DGV_Spec);            
            int inx = 0;
            foreach (DataGridViewColumn c in tar_DGV_Spec.Columns)
            {
                string t = c.Name;
                string[] temp2 = src_Spec_tbl.AsEnumerable().Where(r => r.Field<string>("Col_Name") == t).Select(r => r.Field<string>("Spec")).ToArray();
                if(temp2.Length > 0)
                {
                    limit_lst[inx] = temp2[0];
                    col_header[inx] = c.HeaderText;
                    inx++;
                }
                else
                {
                    mySpec_dgv.Columns.Remove(t);
                }
            }
            Array.Resize(ref limit_lst, inx);
            DataRow dr_SV = mySpec_dgv.NewRow();
            DataRow dr_UL = mySpec_dgv.NewRow();
            DataRow dr_LL = mySpec_dgv.NewRow();
            inx = 0;
            foreach (string _limit in limit_lst)
            {
                if (_limit.Contains('/'))
                {
                    string[] temp = _limit.Split('/');
                    double sv;
                    double ul;
                    double ll;
                    if (temp[0].Contains('-'))
                    {
                        sv = Convert.ToDouble(temp[0].Split('-')[0].Trim(trim_chr));
                        ll = Convert.ToDouble(temp[0].Split('-')[1].Trim(trim_chr));
                        ul = Convert.ToDouble(temp[1].Trim(trim_chr));
                    }
                    else
                    {
                        sv = Convert.ToDouble(temp[0].Split('+')[0].Trim(trim_chr));
                        ul = Convert.ToDouble(temp[0].Split('+')[1].Trim(trim_chr));
                        ll = Convert.ToDouble(temp[1].Trim(trim_chr));
                    }
                    dr_SV[inx] = sv;
                    dr_UL[inx] = (sv + ul).ToString();
                    dr_LL[inx] = (sv - ll).ToString();
                }
                else
                {
                    string[] UL_LL = _limit.Split(split_chr);
                    if (_limit.Contains('>'))
                    {
                        dr_SV[inx] = "";
                        dr_UL[inx] = UL_LL[0].Trim(trim_chr);
                        dr_LL[inx] = UL_LL[1].Trim(trim_chr);
                    }
                    else
                    {
                        if (_limit.Contains('\u00B1'))
                        {
                            dr_SV[inx] = UL_LL[0];
                            dr_UL[inx] = (Convert.ToDouble(UL_LL[0].Trim(trim_chr)) + Convert.ToDouble(UL_LL[1].Trim(trim_chr))).ToString();
                            dr_LL[inx] = (Convert.ToDouble(UL_LL[0].Trim(trim_chr)) - Convert.ToDouble(UL_LL[1].Trim(trim_chr))).ToString();
                        }
                        else
                        {
                            dr_SV[inx] = "";
                            dr_LL[inx] = UL_LL[0].Trim(trim_chr);
                            dr_UL[inx] = UL_LL[1].Trim(trim_chr);
                        }
                    }
                }
                inx++;
            }
            if(mySpec_dgv.Rows.Count>0)
            {
                mySpec_dgv.Rows.Clear();
            }
            mySpec_dgv.Rows.Add(dr_SV);
            mySpec_dgv.Rows.Add(dr_UL);
            mySpec_dgv.Rows.Add(dr_LL);
            tar_DGV_Spec.Columns.Clear();
            tar_DGV_Spec.DataSource = mySpec_dgv;
            tar_DGV_Spec.Rows[0].HeaderCell.Value = "SV";
            tar_DGV_Spec.Rows[1].HeaderCell.Value = "UL";
            tar_DGV_Spec.Rows[2].HeaderCell.Value = "LL";
            int c_inx = 0;
            foreach (DataGridViewColumn c in tar_DGV_Spec.Columns)
            {
                c.HeaderText = col_header[c_inx];
                c_inx++;
            }
            tar_DGV_Spec.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            tar_DGV_Spec.AutoResizeColumns();
            Disable_Sort_DGV(tar_DGV_Spec);
            tar_DGV_Spec.EditMode = DataGridViewEditMode.EditProgrammatically;
        }
        public void IPQC_Spec_Process_Man3(DataTable src_Spec_tbl, DataGridView tar_DGV_Spec)
        {
            string[] limit_lst = src_Spec_tbl.AsEnumerable().Select(r => r.Field<string>("Spec")).ToArray();
            string[] colname_list = src_Spec_tbl.AsEnumerable().Select(r => r.Field<string>("Col_Name")).ToArray();
            string[] col_header = new string[tar_DGV_Spec.Columns.Count];
            char[] split_chr = { '~', '<', '\u2264', '\u2265', '\u00B1', '>' };
            char[] trim_chr = { 'u', 'm', '\u03BC', '\u00B5', ' ', '=', '-' };
            DataTable mySpec_dgv = DGV_To_Table(tar_DGV_Spec);
            int inx = 0;
            foreach (DataGridViewColumn c in tar_DGV_Spec.Columns)
            {
                string t = c.Name;
                string[] temp2 = src_Spec_tbl.AsEnumerable().Where(r => r.Field<string>("Col_Name") == t).Select(r => r.Field<string>("Spec")).ToArray();
                if (temp2.Length > 0)
                {
                    limit_lst[inx] = temp2[0];
                    col_header[inx] = c.HeaderText;
                    inx++;
                }
                else
                {
                    mySpec_dgv.Columns.Remove(t);
                }
            }
            Array.Resize(ref limit_lst, inx);
            DataRow dr_SV = mySpec_dgv.NewRow();
            DataRow dr_UL = mySpec_dgv.NewRow();
            DataRow dr_LL = mySpec_dgv.NewRow();
            inx = 0;
            foreach (string _limit in limit_lst)
            {
                if (_limit.Contains('/'))
                {
                    string[] temp = _limit.Split('/');
                    double sv;
                    double ul;
                    double ll;
                    if (temp[0].Contains('-'))
                    {
                        sv = Convert.ToDouble(temp[0].Split('-')[0].Trim(trim_chr));
                        ll = Convert.ToDouble(temp[0].Split('-')[1].Trim(trim_chr));
                        ul = Convert.ToDouble(temp[1].Trim(trim_chr));
                    }
                    else
                    {
                        sv = Convert.ToDouble(temp[0].Split('+')[0].Trim(trim_chr));
                        ul = Convert.ToDouble(temp[0].Split('+')[1].Trim(trim_chr));
                        ll = Convert.ToDouble(temp[1].Trim(trim_chr));
                    }
                    dr_SV[inx] = sv;
                    dr_UL[inx] = (sv + ul).ToString();
                    dr_LL[inx] = (sv - ll).ToString();
                }
                else
                {
                    string[] UL_LL = _limit.Split(split_chr);
                    if (_limit.Contains('>'))
                    {
                        dr_SV[inx] = _limit;
                        dr_UL[inx] = UL_LL[0].Trim(trim_chr);
                        dr_LL[inx] = UL_LL[1].Trim(trim_chr);
                    }
                    else
                    {
                        if (_limit.Contains('\u00B1'))
                        {
                            dr_SV[inx] = UL_LL[0];
                            dr_UL[inx] = (Convert.ToDouble(UL_LL[0].Trim(trim_chr)) + Convert.ToDouble(UL_LL[1].Trim(trim_chr))).ToString();
                            dr_LL[inx] = (Convert.ToDouble(UL_LL[0].Trim(trim_chr)) - Convert.ToDouble(UL_LL[1].Trim(trim_chr))).ToString();
                        }
                        else
                        {
                            if(_limit!="NA")
                            {
                                dr_SV[inx] = _limit;
                                dr_LL[inx] = UL_LL[0].Trim(trim_chr);
                                dr_UL[inx] = UL_LL[1].Trim(trim_chr);
                            }
                            else
                            {
                                dr_SV[inx] = _limit;
                                dr_LL[inx] = _limit;
                                dr_UL[inx] = _limit;
                            }

                        }
                    }


                }
                inx++;
            }
            if (mySpec_dgv.Rows.Count > 0)
            {
                mySpec_dgv.Rows.Clear();
            }
            mySpec_dgv.Rows.Add(dr_SV);
            mySpec_dgv.Rows.Add(dr_UL);
            mySpec_dgv.Rows.Add(dr_LL);
            tar_DGV_Spec.Columns.Clear();
            tar_DGV_Spec.DataSource = null;
            tar_DGV_Spec.DataSource = mySpec_dgv;
            tar_DGV_Spec.Rows[0].HeaderCell.Value = "SV";
            tar_DGV_Spec.Rows[1].HeaderCell.Value = "UL";
            tar_DGV_Spec.Rows[2].HeaderCell.Value = "LL";
            int c_inx = 0;
            foreach (DataGridViewColumn c in tar_DGV_Spec.Columns)
            {
                c.HeaderText = col_header[c_inx];
                c_inx++;
            }
            tar_DGV_Spec.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            tar_DGV_Spec.AutoResizeColumns();
            Disable_Sort_DGV(tar_DGV_Spec);
            tar_DGV_Spec.EditMode = DataGridViewEditMode.EditProgrammatically;
        }
        public void save_FAI_spec(SEI_Lib.FAI_Spec[] src_FAI_spec, string tar_ItemCode, SqlConnection sqlcon)
        {
            string[] items = new string[11];
            string[] item_vals = new string[11];
            items[0] = "ID";
            items[1] = "ItemCode";
            items[2] = "LotNo";
            items[3] = "NormDim";
            items[4] = "TolMax";
            items[5] = "TolMin";
            items[6] = "Instrument";
            items[7] = "FAI_No";
            items[8] = "Distribution";
            items[9] = "SheetNo";
            item_vals[1] = tar_ItemCode;//"ItemCode";
            //item_vals[2] = txtLotNo.Text;//"LotNo";
            
            if (src_FAI_spec.Length > 0)
            {
                foreach (SEI_Lib.FAI_Spec t in src_FAI_spec)
                {
                    AutoCompleteStringCollection FAI_No_list = new AutoCompleteStringCollection();
                    item_vals[3] = "";
                    item_vals[4] = "";
                    item_vals[5] = "";
                    item_vals[8] = "";
                    item_vals[6] = t.instrument;
                    item_vals[7] = t.FAI_Name;
                    FAI_No_list = TDMK_Code.Load_Item_Filter_str(sqlcon, "FAI_Spec", "NormDim", TDMK_Code.filter_str(items, item_vals));
                    if (FAI_No_list.Count == 0)
                    {
                        item_vals[0] = (TDMK_Code.SQL_MAX("FAI_Spec", "ID", sqlcon) + 1).ToString();
                        item_vals[3] = t.SetVal;
                        item_vals[4] = t.UL;
                        item_vals[5] = t.LL;
                        item_vals[8] = t.check_side;
                        item_vals[9] = t.sheetno;
                        TDMK_Code.insert_val_arr("FAI_Spec", sqlcon, items, item_vals);
                    }
                }
            }
        }        
        public void Calcul_CPK_FAI(SEI_Lib.FAI_Spec src_FAI_spec, string[] tar_FAI_Data, ref SEI_Lib.Calculate_CPK tar_CPK_Result, ref SEI_Lib.FAI_Histogram_Data tar_Histogram_Data)
        {
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

            if(IsNumeric_Val(src_FAI_spec.UL)!="")
            {
                UL = Norminal_Dim + Convert.ToDouble(src_FAI_spec.UL);
            }
            if(IsNumeric_Val(src_FAI_spec.LL)!="")
            {
                LL = Norminal_Dim - Convert.ToDouble(src_FAI_spec.LL);
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
            XiShu = freq_bin_data.Max() / NormDist_Bin.Max();
            for (int i = 0; i < 50; i++)
            {
                modified_NormDist[i] = NormDist_Bin[i] * XiShu;
            }
            tar_CPK_Result.FAI_No = src_FAI_spec.FAI_Name;

            
            switch(Side_check)
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

            //Form2 chart_form = new Form2();
            //chart_form.src_bin_data = bin_data;
            //chart_form.src_freq_bin_data = freq_bin_data;
            //chart_form.src_FAI_spec_Data = data_arr;
            //chart_form.FAI_Spec_val = testFAI[0];
            //chart_form.src_modified_NormDist = modified_NormDist;
            //chart_form.Show();
        }
        public void Calcul_CPK_FAI_Located_NG(SEI_Lib.FAI_Spec src_FAI_spec, string[] tar_FAI_Data, ref SEI_Lib.Calculate_CPK tar_CPK_Result, ref SEI_Lib.FAI_Histogram_Data tar_Histogram_Data, List<int> Item_Located)
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
                UL = Norminal_Dim + Convert.ToDouble(src_FAI_spec.UL);
            }
            if (IsNumeric_Val(src_FAI_spec.LL) != "")
            {
                LL = Norminal_Dim - Convert.ToDouble(src_FAI_spec.LL);
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
                double cur_mod = Math.Round(modified_NormDist[i],4);
                if (cur_mod < 0.01)
                {
                    myResult.Add(myData[i]);
                }
            }
            foreach (var item in myResult)
            {
                foreach (KeyValuePair<uint, double> x in item)
                {
                    Item_Located.Add((int)x.Key);
                }
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
        public void Load_CPK_DGV(DataTable src_FAI_Data, DataGridView tar_DGV, SEI_Lib.FAI_Spec[] src_FAI_Spec)
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
                foreach (SEI_Lib.FAI_Spec ref_spec in src_FAI_Spec)
                {
                    if (ref_spec.FAI_Name == t)
                    {
                        sel_FAI_test = ref_spec;
                        string[] temp_FAI_data = test_tbl.AsEnumerable().Select(r => r.Field<string>(t)).ToArray();
                        Double[] data_arr = new double[temp_FAI_data.Length];
                        int inx = 0;
                        foreach (string c in temp_FAI_data)
                        {
                            if (checkDBNull(c) != "")
                            {
                                data_arr[inx] = Convert.ToDouble(c);
                                inx++;
                            }
                        }
                        Array.Resize(ref data_arr, inx);
                        Calcul_CPK_FAI(ref_spec, temp_FAI_data, ref myCalc_CPK[col_list_inx], ref myHistogram_data[col_list_inx]);
                        break;
                    }
                }
                col_list_inx++;
            }
            DataTable tbl_CPK = new DataTable();
            foreach (SEI_Lib.Calculate_CPK item_CPK in myCalc_CPK)
            {
                string CPK_col_name = item_CPK.FAI_No;
                if (CPK_col_name !=null)
                {
                    if (!check_columns_existed(tbl_CPK, CPK_col_name))
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
            tar_DGV.DataSource = null;
            tar_DGV.DataSource = tbl_CPK;
            if (tar_DGV.Rows.Count > 0)
            {
                tar_DGV.Rows[0].HeaderCell.Value = "STDEV";
                tar_DGV.Rows[1].HeaderCell.Value = "MEAN";
                tar_DGV.Rows[2].HeaderCell.Value = "MAX";
                tar_DGV.Rows[3].HeaderCell.Value = "MIN";
                tar_DGV.Rows[4].HeaderCell.Value = "CP";
                tar_DGV.Rows[5].HeaderCell.Value = "CPKL";
                tar_DGV.Rows[6].HeaderCell.Value = "CPKU";
                tar_DGV.Rows[7].HeaderCell.Value = "CPK";
                tar_DGV.Rows[8].HeaderCell.Value = "CPKM";
                tar_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                Disable_Sort_DGV(tar_DGV);
                foreach (DataGridViewColumn c in tar_DGV.Columns)
                {
                    double cpk_val = Convert.ToDouble(tar_DGV.Rows[7].Cells[c.Index].Value);
                    if (cpk_val < 1.67)
                    {
                        tar_DGV.Rows[7].Cells[c.Index].Style.BackColor = Color.Yellow;
                        tar_DGV.Rows[7].Cells[c.Index].Style.ForeColor = Color.Red;
                    }
                }
            }

        }
        public void Load_CPK_DGV_Histogram(DataGridView src_DGV, DataGridView tar_DGV, SEI_Lib.FAI_Spec[] src_FAI_Spec, DataGridView tar_Hist_DGV)
        {
            DataTable test_tbl =  DGV_To_Table(src_DGV);//src_FAI_Data;//
            int arr_num = test_tbl.Columns.Count;// col_list.Count;
            SEI_Lib.Calculate_CPK[] myCalc_CPK = new SEI_Lib.Calculate_CPK[arr_num];
            SEI_Lib.FAI_Histogram_Data[] myHistogram_data = new SEI_Lib.FAI_Histogram_Data[arr_num];
            int col_list_inx = 0;
            DataTable histo_tbl = new DataTable();
            Dictionary<string, Bitmap> Histo_result_lst = new Dictionary<string, Bitmap>();
            foreach (DataColumn tbl_col in test_tbl.Columns)
            {
                SEI_Lib.FAI_Spec sel_FAI_test;
                string t = tbl_col.ColumnName;
                foreach (SEI_Lib.FAI_Spec ref_spec in src_FAI_Spec)
                {
                    if (ref_spec.FAI_Name == t)
                    {
                        sel_FAI_test = ref_spec;
                        string[] temp_FAI_data = test_tbl.AsEnumerable().Where(x=>x.Field<string>(t)!=null).Select(r => r.Field<string>(t)).ToArray();
                        if(temp_FAI_data.Length>0)
                        {
                            if(!temp_FAI_data.Any(x=>!IsNumeric(x)))
                            {
                                Double[] data_arr = new double[temp_FAI_data.Length];
                                int inx = 0;
                                foreach (string c in temp_FAI_data)
                                {
                                    if (checkDBNull(c) != "")
                                    {
                                        data_arr[inx] = Convert.ToDouble(c);
                                        inx++;
                                    }
                                }
                                Array.Resize(ref data_arr, inx);
                                //Calcul_CPK_FAI(ref_spec, temp_FAI_data, ref myCalc_CPK[col_list_inx], ref myHistogram_data[col_list_inx]);
                                List<int> item_NG_lst = new List<int>();
                                Calcul_CPK_FAI_Located_NG(ref_spec, temp_FAI_data, ref myCalc_CPK[col_list_inx], ref myHistogram_data[col_list_inx], item_NG_lst);
                                if (myCalc_CPK[col_list_inx].CPK<1.67)
                                {
                                    foreach (int item_inx in item_NG_lst)
                                    {
                                        src_DGV.Rows[item_inx].Cells[t].Style.BackColor = Color.Aqua;
                                    }
                                }    
                                Bitmap cur_hist = Draw_Histogram(myHistogram_data[col_list_inx]._Bin_data, myHistogram_data[col_list_inx]._Freq_bin_data, myHistogram_data[col_list_inx]._Modified_NormDist_data, ref_spec);
                                Histo_result_lst.Add(t, cur_hist);
                            }
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
                    if (!check_columns_existed(tbl_CPK, CPK_col_name))
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
            tar_DGV.DataSource = null;
            tar_DGV.DataSource = tbl_CPK;
            if (tar_DGV.Rows.Count > 0)
            {
                tar_DGV.Rows[0].HeaderCell.Value = "STDEV";
                tar_DGV.Rows[1].HeaderCell.Value = "MEAN";
                tar_DGV.Rows[2].HeaderCell.Value = "MAX";
                tar_DGV.Rows[3].HeaderCell.Value = "MIN";
                tar_DGV.Rows[4].HeaderCell.Value = "CP";
                tar_DGV.Rows[5].HeaderCell.Value = "CPKL";
                tar_DGV.Rows[6].HeaderCell.Value = "CPKU";
                tar_DGV.Rows[7].HeaderCell.Value = "CPK";
                tar_DGV.Rows[8].HeaderCell.Value = "CPKM";
                tar_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                Disable_Sort_DGV(tar_DGV);
                foreach (DataGridViewColumn c in tar_DGV.Columns)
                {
                    double cpk_val = Convert.ToDouble(tar_DGV.Rows[7].Cells[c.Index].Value);
                    if (cpk_val < 1.67)
                    {
                        tar_DGV.Rows[7].Cells[c.Index].Style.BackColor = Color.Yellow;
                        tar_DGV.Rows[7].Cells[c.Index].Style.ForeColor = Color.Red;
                    }
                }
            }
            foreach(var histo in Histo_result_lst)
            {
                DataColumn column = new DataColumn(histo.Key);
                column.DataType = System.Type.GetType("System.Byte[]");
                histo_tbl.Columns.Add(column);
                if(histo_tbl.Rows.Count==0)
                {
                    histo_tbl.Rows.Add();
                }
                histo_tbl.Rows[0][histo.Key] = imgToByteConverter(histo.Value);
            }
            tar_Hist_DGV.DataSource = histo_tbl;
            DGV_Auto_Resize(tar_Hist_DGV);
            tar_Hist_DGV.Rows[0].Height = 100;
        }
        public void Load_CPK_DGV_ColName(DataTable src_FAI_Data, DataGridView tar_DGV, SEI_Lib.FAI_Spec[] src_FAI_Spec)
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
                foreach (SEI_Lib.FAI_Spec ref_spec in src_FAI_Spec)
                {
                    if (ref_spec.FAI_Name == t)
                    {
                        sel_FAI_test = ref_spec;
                        string[] temp_FAI_data = test_tbl.AsEnumerable().Select(r => r.Field<string>(t)).ToArray();
                        Double[] data_arr = new double[temp_FAI_data.Length];
                        int inx = 0;
                        foreach (string c in temp_FAI_data)
                        {
                            if (checkDBNull(c) != "")
                            {
                                data_arr[inx] = Convert.ToDouble(c);
                                inx++;
                            }
                        }
                        Array.Resize(ref data_arr, inx);
                        if(inx>0)
                        {
                            Calcul_CPK_FAI(ref_spec, temp_FAI_data, ref myCalc_CPK[col_list_inx], ref myHistogram_data[col_list_inx]);
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
                    if (!check_columns_existed(tbl_CPK, CPK_col_name))
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
            tar_DGV.DataSource = tbl_CPK;

            if (tar_DGV.Rows.Count > 0)
            {
                tar_DGV.Rows[0].HeaderCell.Value = "STDEV";
                tar_DGV.Rows[1].HeaderCell.Value = "MEAN";
                tar_DGV.Rows[2].HeaderCell.Value = "MAX";
                tar_DGV.Rows[3].HeaderCell.Value = "MIN";
                tar_DGV.Rows[4].HeaderCell.Value = "CP";
                tar_DGV.Rows[5].HeaderCell.Value = "CPKL";
                tar_DGV.Rows[6].HeaderCell.Value = "CPKU";
                tar_DGV.Rows[7].HeaderCell.Value = "CPK";
                tar_DGV.Rows[8].HeaderCell.Value = "CPKM";
                tar_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                Disable_Sort_DGV(tar_DGV);
                foreach (DataGridViewColumn c in tar_DGV.Columns)
                {
                    double cpk_val = Convert.ToDouble(tar_DGV.Rows[7].Cells[c.Index].Value);
                    if (cpk_val < 1.67)
                    {
                        tar_DGV.Rows[7].Cells[c.Index].Style.BackColor = Color.Yellow;
                        tar_DGV.Rows[7].Cells[c.Index].Style.ForeColor = Color.Red;
                    }
                }
            }

        }
        public bool check_data_OK(DataGridView src_DGV_data, DataGridView src_DGV_Spec)
        {
            bool _result = false;
            for (int c = 0; c < src_DGV_Spec.Columns.Count; c++)
            {
                string src_UL = checkDBNull(src_DGV_Spec.Rows[1].Cells[c].Value);
                string src_LL = checkDBNull(src_DGV_Spec.Rows[2].Cells[c].Value);
                string SV = checkDBNull(src_DGV_Spec.Rows[0].Cells[c].Value);
                for (int r = 0; r < src_DGV_data.Rows.Count; r++)
                {
                    string src_act_val = checkDBNull(src_DGV_data.Rows[r].Cells[c].Value);
                    if (check_in_limit(src_UL, src_LL, src_act_val,SV))
                    {
                        _result = true;
                    }
                    else
                    {
                        _result = false;
                        goto end_label;
                    }

                }
            }
        end_label: return _result;
        }
        public bool check_FAIdata_OK(DataGridView src_DGV_data, DataGridView src_DGV_Spec)
        {
            bool _result = false;
            for (int c = 0; c < src_DGV_Spec.Columns.Count; c++)
            {
                string Tol_Max = checkDBNull(src_DGV_Spec.Rows[1].Cells[c].Value);
                string Tol_Min = checkDBNull(src_DGV_Spec.Rows[2].Cells[c].Value);
                string SetVal = checkDBNull(src_DGV_Spec.Rows[0].Cells[c].Value);
                double UL = Convert.ToDouble(SetVal) + Convert.ToDouble(Tol_Max);
                double LL = Convert.ToDouble(SetVal) - Convert.ToDouble(Tol_Min);
                for (int r = 0; r < src_DGV_data.Rows.Count; r++)
                {
                    string src_act_val = checkDBNull(src_DGV_data.Rows[r].Cells[c].Value);
                    if (check_in_limit(UL.ToString(), LL.ToString(), src_act_val,SetVal))
                    {
                        _result = true;
                    }
                    else
                    {
                        _result = false;
                        goto end_label;
                    }

                }
            }
        end_label: return _result;
        }
        public bool check_FAIdata_OK2(DataGridView src_DGV_data, DataGridView src_DGV_Spec)
        {
            bool _result = false;
            for (int c = 0; c < src_DGV_data.Columns.Count; c++)
            {
                string col_name = src_DGV_data.Columns[c].Name;
                string Tol_Max = checkDBNull(src_DGV_Spec.Rows[1].Cells[col_name].Value);
                string Tol_Min = checkDBNull(src_DGV_Spec.Rows[2].Cells[col_name].Value);
                string SetVal = checkDBNull(src_DGV_Spec.Rows[0].Cells[col_name].Value);
                double UL = Convert.ToDouble(SetVal) + Convert.ToDouble(Tol_Max);
                double LL = Convert.ToDouble(SetVal) - Convert.ToDouble(Tol_Min);
                for (int r = 0; r < src_DGV_data.Rows.Count; r++)
                {
                    string src_act_val = checkDBNull(src_DGV_data.Rows[r].Cells[c].Value);
                    if (check_in_limit(UL.ToString(), LL.ToString(), src_act_val, SetVal))
                    {
                        _result = true;
                    }
                    else
                    {
                        _result = false;
                        goto end_label;
                    }

                }
            }
            end_label: return _result;
        }
        //public bool check_in_limit(string src_UL, string src_LL, string src_act_val)
        //{
        //    bool _result = false;

        //    if (IsNumeric(src_act_val))
        //    {
        //        if (src_UL != "")
        //        {
        //            double UL = Convert.ToDouble(src_UL);
        //            double act_val = Convert.ToDouble(src_act_val);
        //            if (src_LL != "")
        //            {
        //                double LL = Convert.ToDouble(src_LL);
        //                if ((act_val <= UL) && (act_val >= LL))
        //                {
        //                    _result = true;
        //                }
        //                else
        //                {
        //                    _result = false;
        //                }
        //            }
        //            else
        //            {
        //                if (act_val <= UL)
        //                {
        //                    _result = true;
        //                }
        //                else
        //                {
        //                    _result = false;
        //                }
        //            }
        //        }
        //        else
        //        {
        //            if (src_LL != "")
        //            {
        //                double LL = Convert.ToDouble(src_LL);
        //                double act_val = Convert.ToDouble(src_act_val);
        //                if (act_val >= LL)
        //                {
        //                    _result = true;
        //                }
        //                else
        //                {
        //                    _result = false;
        //                }
        //            }
        //            else
        //            {
        //                _result = false;
        //            }
        //        }
        //    }
        //    return _result;
        //}
        public Color check_in_limit2(string src_UL, string src_LL, string src_act_val)
        {
            Color tar_color = Color.LightPink;
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
        public Color check_in_limit3(string src_UL, string src_LL, string src_act_val)
        {
            Color tar_color = Color.LightPink;
            bool UL_equal_comp = true;
            bool LL_equal_comp = true;
            char[] trim_chr = { 'u', 'm', '\u03BC', '\u00B5', ' ', '=', '>','<' };
            if (IsNumeric(src_act_val))
            {
                if (src_UL != "")
                {
                    if(!src_UL.Contains('='))
                    {
                        UL_equal_comp = false;
                    }
                    else
                    {
                        UL_equal_comp = true;
                    }
                    double UL = Convert.ToDouble(src_UL.Trim(trim_chr));
                    double act_val = Convert.ToDouble(src_act_val);
                    if (src_LL != "")
                    {
                        double LL = Convert.ToDouble(src_LL.Trim(trim_chr));
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
                            if(UL_equal_comp)
                            {
                                if(act_val == UL)
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
                        if (!src_LL.Contains('='))
                        {
                            LL_equal_comp = false;
                        }
                        else
                        {
                            LL_equal_comp = true;
                        }
                        double LL = Convert.ToDouble(src_LL.Trim(trim_chr));
                        double act_val = Convert.ToDouble(src_act_val);
                        if (act_val > LL)
                        {
                            tar_color = Color.White;
                        }
                        else
                        {
                            if(LL_equal_comp)
                            {
                                if(act_val == LL)
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
        public Color check_in_limit4(string src_UL, string src_LL, string src_act_val, bool UL_equal_comp, bool LL_equal_comp)
        {
            Color tar_color = Color.LightPink;
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
        public Color check_in_limit5(string src_UL, string src_LL, string src_act_val, string SetVal)
        {
            Color tar_color = Color.LightPink;
            bool UL_equal_comp = true;
            bool LL_equal_comp = true;

            if(!IsNumeric(SetVal) && (!SetVal.Contains("=")))
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
            return _result;
        }
        public bool check_data_enough(DataGridView src_DGV, int row_num)
        {
            bool _result = false;
            if (src_DGV .Rows.Count >= row_num )
            {
                for (int r = 0; r < row_num; r++)
                {
                    for (int c = 0; c < src_DGV.Columns.Count; c++)
                    {
                        if (checkDBNull(src_DGV.Rows[r].Cells[c].Value) == "")
                        {
                            _result = false;
                            goto end_label;
                        }
                        else
                        {
                            _result = true;
                        }
                    }
                }
            }
        end_label: return _result;
        }
        public bool check_ColumnData_enough(DataGridView src_DGV, int row_num, string col_name)
        {
            bool _result = false;
            if (src_DGV.Rows.Count >= row_num)
            {
                for (int r = 0; r < row_num; r++)
                {
                    if (checkDBNull(src_DGV.Rows[r].Cells[col_name].Value) == "")
                    {
                        _result = false;
                        goto end_label;
                    }
                    else
                    {
                        _result = true;
                    }

                }
            }
        end_label: return _result;
        }
        public void AU_IN_Scan(string src_file, DataGridView tar_DGV_Data, DataGridView tar_DGV_Spec, DataGridView tar_DGV_Data_Plus, bool _editMode, int qty)
        {
            if (tar_DGV_Spec.RowCount > 0)
            {
                myExcel.Workbook src_wrkbk = TDMK_Code.open_excel_file(src_file, "", "");
                myExcel.Worksheet src_wrksht = src_wrkbk.Sheets[1];
                myExcel.Range src_AU_rgn = src_wrksht.Range["I9"];
                myExcel.Range src_NI_rgn = src_wrksht.Range["J9"];
                string[] L1_AU = new string[qty];
                string[] L2_AU = new string[qty];
                string[] L1_NI = new string[qty];
                string[] L2_NI = new string[qty];
                int inx = 0;
                while (checkDBNull(src_AU_rgn.Offset[inx, 0].Value) != "")
                {
                    string curr_AU_val = checkDBNull(src_AU_rgn.Offset[inx, 0].Value);
                    string curr_NI_val = checkDBNull(src_NI_rgn.Offset[inx, 0].Value);
                    if (inx < qty)
                    {
                        L1_AU[inx] = Math.Round(Convert.ToDouble(curr_AU_val), 3).ToString();
                        L1_NI[inx] = Math.Round(Convert.ToDouble(curr_NI_val), 3).ToString();
                    }
                    else
                    {
                        if (inx < 2*qty)
                        {
                            L2_AU[inx - qty] = Math.Round(Convert.ToDouble(curr_AU_val), 3).ToString();
                            L2_NI[inx - qty] = Math.Round(Convert.ToDouble(curr_NI_val), 3).ToString();
                        }
                        else
                        {
                            break;
                        }
                    }
                    inx++;
                }
                DataGridView sel_DGV;
                if (_editMode)
                {
                    sel_DGV = tar_DGV_Data_Plus;
                }
                else
                {
                    sel_DGV = tar_DGV_Data;
                }
                sel_DGV.Columns.Clear();
                for (int i = 0; i < tar_DGV_Spec.Columns.Count; i++)
                {
                    sel_DGV.Columns.Add(tar_DGV_Spec.Columns[i].Name, tar_DGV_Spec.Columns[i].Name);
                }
                for (int i = 0; i < qty; i++)
                {
                    sel_DGV.Rows.Add(L1_AU[i], L2_AU[i], L1_NI[i], L2_NI[i]);
                    for (int j = 0; j < sel_DGV.Columns.Count; j++)
                    {
                        string act_val = checkDBNull(sel_DGV.Rows[i].Cells[j].Value);
                        if (act_val != "")
                        {
                            string UL = checkDBNull(tar_DGV_Spec.Rows[1].Cells[j].Value);
                            string LL = checkDBNull(tar_DGV_Spec.Rows[2].Cells[j].Value);
                            string SV = checkDBNull(tar_DGV_Spec.Rows[0].Cells[j].Value);
                            sel_DGV.Rows[i].Cells[j].Style.BackColor = check_in_limit5(UL, LL, act_val, SV);
                        }
                        else
                        {
                            sel_DGV.Rows[i].Cells[j].Style.BackColor = Color.Yellow;
                        }
                    }

                    sel_DGV.Rows[i].HeaderCell.Value = (i + 1).ToString();
                }
                sel_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                Disable_Sort_DGV(sel_DGV);
                src_wrkbk.Close();
            }
        }
        public void Roughness_Scan(string src_file, DataGridView tar_DGV_Data, DataGridView tar_DGV_Spec, DataGridView tar_DGV_Data_Plus, bool _editMode, int qty)
        {
            if (tar_DGV_Spec.RowCount > 0)
            {
                myExcel.Workbook src_wrkbk = TDMK_Code.open_excel_file(src_file, "", "");
                myExcel.Worksheet src_wrksht = src_wrkbk.Sheets[1];
                myExcel.Range src_Sa_rgn = src_wrksht.Range["B2"];
                myExcel.Range src_Sq_rgn = src_wrksht.Range["C2"];
                string[] L1_Sa = new string[qty];
                string[] L2_Sa = new string[qty];
                string[] L1_Sq = new string[qty];
                string[] L2_Sq = new string[qty];

                int inx = 0;
                while (checkDBNull(src_Sa_rgn.Offset[inx, 0].Value) != "")
                {
                    string curr_Sa_val = checkDBNull(src_Sa_rgn.Offset[inx, 0].Value);
                    string curr_Sq_val = checkDBNull(src_Sq_rgn.Offset[inx, 0].Value);
                    if (inx < qty)
                    {
                        L1_Sa[inx] = Math.Round(Convert.ToDouble(curr_Sa_val), 3).ToString();
                        L1_Sq[inx] = Math.Round(Convert.ToDouble(curr_Sq_val), 3).ToString();
                    }
                    else
                    {
                        if (inx < 2 * qty)
                        {
                            L2_Sa[inx - qty] = Math.Round(Convert.ToDouble(curr_Sa_val), 3).ToString();
                            L2_Sq[inx - qty] = Math.Round(Convert.ToDouble(curr_Sq_val), 3).ToString();
                        }
                        else
                        {
                            break;
                        }
                    }
                    inx++;
                }
                DataGridView sel_DGV;
                if (_editMode)
                {
                    sel_DGV = tar_DGV_Data_Plus;
                }
                else
                {
                    sel_DGV = tar_DGV_Data;
                }
                sel_DGV.Columns.Clear();
                for (int i = 0; i < tar_DGV_Spec.Columns.Count; i++)
                {
                    sel_DGV.Columns.Add(tar_DGV_Spec.Columns[i].Name, tar_DGV_Spec.Columns[i].Name);
                }
                for (int i = 0; i < qty; i++)
                {
                    sel_DGV.Rows.Add(L1_Sa[i], L1_Sq[i], L2_Sa[i], L2_Sq[i]);
                    for (int j = 0; j < sel_DGV.Columns.Count; j++)
                    {
                        string act_val = checkDBNull(sel_DGV.Rows[i].Cells[j].Value);
                        if (act_val != "")
                        {
                            string UL = checkDBNull(tar_DGV_Spec.Rows[1].Cells[j].Value);
                            string LL = checkDBNull(tar_DGV_Spec.Rows[2].Cells[j].Value);
                            string SV = checkDBNull(tar_DGV_Spec.Rows[0].Cells[j].Value);
                            sel_DGV.Rows[i].Cells[j].Style.BackColor = check_in_limit5(UL, LL, act_val, SV);
                        }
                        else
                        {
                            sel_DGV.Rows[i].Cells[j].Style.BackColor = Color.Yellow;
                        }

                    }
                    sel_DGV.Rows[i].HeaderCell.Value = (i + 1).ToString();
                }
                sel_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                Disable_Sort_DGV(sel_DGV);
                src_wrkbk.Close();
            }

        }
        public void Roughness_Scan2(string src_file, DataGridView tar_DGV_Data, DataGridView tar_DGV_Spec, DataGridView tar_DGV_Data_Plus, bool _editMode, int qty, int MeasLoc)
        {
            if (tar_DGV_Spec.RowCount > 0)
            {
                string[] L1_Sa = new string[qty];
                string[] L2_Sa = new string[qty];
                string[] L1_Sq = new string[qty];
                string[] L2_Sq = new string[qty];

                string[][] data_arr = new string[][] { L1_Sa, L2_Sa, L1_Sq, L2_Sq };
                List<string>[] Result = Roughness_Result(src_file, MeasLoc);
                for(int i=0; i<4;i++)
                {
                    if( Result[i].Count !=0)
                    {
                        if (Result[i].Count > qty)
                        {
                            Result[i].CopyTo(0, data_arr[i], 0, qty);
                        }
                        else
                        {
                            Result[i].CopyTo(data_arr[i]);
                        }
                    }
                    else
                    {
                        for(int j = 0; j< qty; j ++)
                        {
                            data_arr[i][j] = "NA";
                        }
                    }

                }
                DataGridView sel_DGV;
                if (_editMode)
                {
                    sel_DGV = tar_DGV_Data_Plus;
                }
                else
                {
                    sel_DGV = tar_DGV_Data;
                }
                sel_DGV.Columns.Clear();
                for (int i = 0; i < tar_DGV_Spec.Columns.Count; i++)
                {
                    sel_DGV.Columns.Add(tar_DGV_Spec.Columns[i].Name, tar_DGV_Spec.Columns[i].Name);
                }
                for (int i = 0; i < qty; i++)
                {
                    sel_DGV.Rows.Add(L1_Sa[i], L1_Sq[i], L2_Sa[i], L2_Sq[i]);
                    for (int j = 0; j < sel_DGV.Columns.Count; j++)
                    {
                        string act_val = checkDBNull(sel_DGV.Rows[i].Cells[j].Value);
                        if (act_val != "")
                        {
                            string UL = checkDBNull(tar_DGV_Spec.Rows[1].Cells[j].Value);
                            string LL = checkDBNull(tar_DGV_Spec.Rows[2].Cells[j].Value);
                            string SV = checkDBNull(tar_DGV_Spec.Rows[0].Cells[j].Value);
                            sel_DGV.Rows[i].Cells[j].Style.BackColor = check_in_limit5(UL, LL, act_val, SV);
                        }
                        else
                        {
                            sel_DGV.Rows[i].Cells[j].Style.BackColor = Color.Yellow;
                        }

                    }
                    sel_DGV.Rows[i].HeaderCell.Value = (i + 1).ToString();
                }
                sel_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                Disable_Sort_DGV(sel_DGV);
            }

        }
        public void Roughness_Scan_ACF(string src_file, DataGridView tar_DGV_Data, DataGridView tar_DGV_Spec, DataGridView tar_DGV_Data_Plus, bool _editMode, int qty, int MeasLoc)
        {
            if (tar_DGV_Spec.RowCount > 0)
            {
                string[] L1_Sa = new string[qty];
                string[] L2_Sa = new string[qty];
                string[] L1_Sq = new string[qty];
                string[] L2_Sq = new string[qty];
                string[] L1_Sdr = new string[qty];
                string[] L2_Sdr = new string[qty];
                string[] L3_Sa = new string[qty];
                string[] L3_Sq = new string[qty];
                string[] L3_Sdr = new string[qty];
                string[][] data_arr = new string[][] { L1_Sa, L2_Sa, L1_Sq, L2_Sq, L1_Sdr, L2_Sdr, L3_Sa, L3_Sq,L3_Sdr };
                List<string>[] Result = Roughness_Result_ACF(src_file, MeasLoc);
                for (int i = 0; i < 9; i++)
                {
                    if (Result[i].Count != 0)
                    {
                        if (Result[i].Count > qty)
                        {
                            Result[i].CopyTo(0, data_arr[i], 0, qty);
                        }
                        else
                        {
                            Result[i].CopyTo(data_arr[i]);
                        }
                    }
                    else
                    {
                        for (int j = 0; j < qty; j++)
                        {
                            data_arr[i][j] = "NA";
                        }
                    }

                }
                DataGridView sel_DGV;
                if (_editMode)
                {
                    sel_DGV = tar_DGV_Data_Plus;
                }
                else
                {
                    sel_DGV = tar_DGV_Data;
                }
                sel_DGV.Columns.Clear();
                for (int i = 0; i < tar_DGV_Spec.Columns.Count; i++)
                {
                    sel_DGV.Columns.Add(tar_DGV_Spec.Columns[i].Name, tar_DGV_Spec.Columns[i].Name);
                }
                //********************* Add new columns for Sdr ***********************************************
                sel_DGV.Columns.Add("L1 Roughness Sdr", "L1 Roughness Sdr");
                sel_DGV.Columns.Add("L2 Roughness Sdr", "L2 Roughness Sdr");
                //*********************************************************************************************
                //********************* Add new columns for L3 Sa / Sq / Sdr ***********************************************
                sel_DGV.Columns.Add("L3 Roughness Sa", "L3 Roughness Sa");
                sel_DGV.Columns.Add("L3 Roughness Sq", "L3 Roughness Sq");
                sel_DGV.Columns.Add("L3 Roughness Sdr", "L3 Roughness Sdr");
                //*********************************************************************************************
                for (int i = 0; i < qty; i++)
                {
                    sel_DGV.Rows.Add(L1_Sa[i], L1_Sq[i], L2_Sa[i], L2_Sq[i], L1_Sdr[i], L2_Sdr[i], L3_Sa[i],L3_Sq[i],L3_Sdr[i]);
                    for (int j = 0; j < sel_DGV.Columns.Count-5; j++)
                    {
                        string act_val = checkDBNull(sel_DGV.Rows[i].Cells[j].Value);
                        if (act_val != "")
                        {
                            string UL = checkDBNull(tar_DGV_Spec.Rows[1].Cells[j].Value);
                            string LL = checkDBNull(tar_DGV_Spec.Rows[2].Cells[j].Value);
                            string SV = checkDBNull(tar_DGV_Spec.Rows[0].Cells[j].Value);
                            sel_DGV.Rows[i].Cells[j].Style.BackColor = check_in_limit5(UL, LL, act_val, SV);
                        }
                        else
                        {
                            sel_DGV.Rows[i].Cells[j].Style.BackColor = Color.Yellow;
                        }

                    }
                    sel_DGV.Rows[i].HeaderCell.Value = (i + 1).ToString();
                }
                sel_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                Disable_Sort_DGV(sel_DGV);
            }

        }
        public void Roughness_Scan_ACF_SMT(string src_file, DataGridView tar_DGV_Data, DataGridView tar_DGV_Spec, DataGridView tar_DGV_Data_Plus, bool _editMode, int qty, int MeasLoc)
        {
            if (tar_DGV_Spec.RowCount > 0)
            {
                string[] L1_Sa = new string[qty];
                string[] L2_Sa = new string[qty];
                string[] L1_Sq = new string[qty];
                string[] L2_Sq = new string[qty];
                string[] L1_Sdr = new string[qty];
                string[] L2_Sdr = new string[qty];
                string[] L3_Sa = new string[qty];
                string[] L3_Sq = new string[qty];
                string[] L3_Sdr = new string[qty];
                string[][] data_arr = new string[][] { L1_Sa, L2_Sa, L1_Sq, L2_Sq, L1_Sdr, L2_Sdr, L3_Sa, L3_Sq, L3_Sdr };
                List<string>[] Result = Roughness_Result_ACF(src_file, MeasLoc);
                for (int i = 0; i < 9; i++)
                {
                    if (Result[i].Count != 0)
                    {
                        if (Result[i].Count > qty)
                        {
                            Result[i].CopyTo(0, data_arr[i], 0, qty);
                        }
                        else
                        {
                            Result[i].CopyTo(data_arr[i]);
                        }
                    }
                    else
                    {
                        for (int j = 0; j < qty; j++)
                        {
                            data_arr[i][j] = "NA";
                        }
                    }

                }
                DataGridView sel_DGV;
                if (_editMode)
                {
                    sel_DGV = tar_DGV_Data_Plus;
                }
                else
                {
                    sel_DGV = tar_DGV_Data;
                }
                sel_DGV.Columns.Clear();
                for (int i = 0; i < tar_DGV_Spec.Columns.Count; i++)
                {
                    sel_DGV.Columns.Add(tar_DGV_Spec.Columns[i].Name, tar_DGV_Spec.Columns[i].Name);
                }
                ////********************* Add new columns for Sdr ***********************************************
                //sel_DGV.Columns.Add("L1 Roughness Sdr", "L1 Roughness Sdr");
                //sel_DGV.Columns.Add("L2 Roughness Sdr", "L2 Roughness Sdr");
                ////*********************************************************************************************
                ////********************* Add new columns for L3 Sa / Sq / Sdr ***********************************************
                //sel_DGV.Columns.Add("L3 Roughness Sa", "L3 Roughness Sa");
                //sel_DGV.Columns.Add("L3 Roughness Sq", "L3 Roughness Sq");
                //sel_DGV.Columns.Add("L3 Roughness Sdr", "L3 Roughness Sdr");
                //*********************************************************************************************
                for (int i = 0; i < qty; i++)
                {
                    sel_DGV.Rows.Add(L1_Sa[i], L1_Sq[i], L2_Sa[i], L2_Sq[i], L1_Sdr[i], L2_Sdr[i], L3_Sa[i], L3_Sq[i], L3_Sdr[i]);
                    for (int j = 0; j < sel_DGV.Columns.Count - 5; j++)
                    {
                        string act_val = checkDBNull(sel_DGV.Rows[i].Cells[j].Value);
                        if (act_val != "")
                        {
                            string UL = checkDBNull(tar_DGV_Spec.Rows[1].Cells[j].Value);
                            string LL = checkDBNull(tar_DGV_Spec.Rows[2].Cells[j].Value);
                            string SV = checkDBNull(tar_DGV_Spec.Rows[0].Cells[j].Value);
                            sel_DGV.Rows[i].Cells[j].Style.BackColor = check_in_limit5(UL, LL, act_val, SV);
                        }
                        else
                        {
                            sel_DGV.Rows[i].Cells[j].Style.BackColor = Color.Yellow;
                        }

                    }
                    sel_DGV.Rows[i].HeaderCell.Value = (i + 1).ToString();
                }
                sel_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                Disable_Sort_DGV(sel_DGV);
            }

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
        public void fill_datatable_DGV3(DataTable src_tbl, DataGridView tar_DGV, DataGridView tar_DGV_Spec, DataGridView tar_DGV_Plus)
        {
            string[] col_name;
            DataTable tbl_spec = new DataTable();
            DataTable tbl_DGV_data = new DataTable();
            DataTable tbl_DGV_plus = new DataTable();
            DataTable sel_tbl = new DataTable();
            DataGridView sel_DGV;
            int tar_DGV_rows_count;//= tar_DGV.Rows.Count;
            int tar_DGV_Plus_rows_count;
            int sel_DGV_rows_count;
            int curr_DGV_row;
            string tar_Col = src_tbl.Columns[1].ColumnName;
            string vitri_col_name = src_tbl.Columns[0].ColumnName;
            string vitri_act_val = src_tbl.Columns[4].ColumnName;
            string vitri_Set_val = src_tbl.Columns[5].ColumnName;
            string vitri_UL_val = src_tbl.Columns[7].ColumnName;
            string vitri_LL_val = src_tbl.Columns[8].ColumnName;
            col_name = src_tbl.AsEnumerable().Select(r => r.Field<string>(tar_Col)).Distinct().ToArray();
            tbl_DGV_data = DGV_To_Table(tar_DGV);
            tbl_DGV_plus = DGV_To_Table(tar_DGV_Plus);
            tar_DGV_rows_count = tbl_DGV_data.Rows.Count;
            tar_DGV_Plus_rows_count = tbl_DGV_plus.Rows.Count;
            tbl_spec = DGV_To_Table(tar_DGV_Spec);
            foreach (string tg in col_name)
            {
                double UL = 0;// = Convert .ToDouble (tbl_spec.Rows [1][tg]);
                double LL = 0;// Convert.ToDouble(tbl_spec.Rows[2][tg]);
                bool spec_check_en = false;
                int sel_col_inx = 0;
                string col_dgv_data = tg;
                if (check_columns_existed_inx(tbl_spec, tg, ref sel_col_inx))
                {
                    double sv = Convert.ToDouble(tbl_spec.Rows[0][sel_col_inx]);
                    UL = sv + Convert.ToDouble(tbl_spec.Rows[1][sel_col_inx]);
                    LL = sv - Convert.ToDouble(tbl_spec.Rows[2][sel_col_inx]);
                    spec_check_en = true;
                    col_dgv_data = tbl_spec.Columns[sel_col_inx].ColumnName;
                    sel_DGV = tar_DGV;
                    sel_tbl = tbl_DGV_data;
                }
                else
                {
                    spec_check_en = false;
                    sel_DGV = tar_DGV_Plus;
                    sel_tbl = tbl_DGV_plus;
                }
                if (!check_columns_existed(sel_tbl, col_dgv_data))
                {
                    sel_DGV.Columns.Add(col_dgv_data, tg);
                }
                sel_DGV_rows_count = sel_tbl.Rows.Count;
                DataTable temp = new DataTable();
                temp = Get_data_detail(src_tbl, tar_Col, tg);
                int row_inx = temp.Rows.Count;
                curr_DGV_row = sel_DGV.RowCount;
                int sel_row_inx = curr_DGV_row - sel_DGV_rows_count;
                if (sel_row_inx < row_inx)
                {
                    int dgv_row = row_inx - sel_row_inx;
                    for (int i = 0; i < dgv_row; i++)
                    {
                        sel_DGV.Rows.Add();
                    }
                }
                string[] item_val;
                string[] vitri_val;
                item_val = temp.AsEnumerable().Select(r => r.Field<string>(vitri_act_val)).ToArray();
                vitri_val = temp.AsEnumerable().Select(r => r.Field<string>(vitri_col_name)).ToArray();
                int j = sel_DGV_rows_count;
                int vitri_inx = 0;
                foreach (string c in item_val)
                {
                    if (checkDBNull(c) != "")
                    {
                        if (vitri_val[vitri_inx].ToUpper().Contains("ANGLE"))
                        {
                            sel_DGV.Rows[j].Cells[col_dgv_data].Value = c;
                        }
                        else
                        {
                            sel_DGV.Rows[j].Cells[col_dgv_data].Value = Math.Abs(Convert.ToDouble(c)).ToString();
                        }
                        if (spec_check_en)
                        {
                            double act_val = Convert.ToDouble(sel_DGV.Rows[j].Cells[col_dgv_data].Value); 
                            if ((act_val <= UL) && (act_val >= LL))
                            {
                                sel_DGV.Rows[j].Cells[col_dgv_data].Style.BackColor = Color.White;
                            }
                            else
                            {
                                if (act_val > UL)
                                {
                                    sel_DGV.Rows[j].Cells[col_dgv_data].Style.BackColor = Color.Yellow;
                                }
                                if (act_val < LL)
                                {
                                    sel_DGV.Rows[j].Cells[col_dgv_data].Style.BackColor = Color.LightBlue;
                                }
                                //confirm_request = true;
                            }
                        }
                        else

                        {
                            sel_DGV.Rows[j].Cells[col_dgv_data].Style.BackColor = Color.LightGray;
                        }
                    }
                    j++;
                    vitri_inx++;
                }
            }
            DGV_Auto_Resize(tar_DGV);
            DGV_Auto_Resize(tar_DGV_Plus);
        }
        public void fill_datatable_DGV_Testing(DataTable src_tbl, DataGridView tar_DGV, DataGridView tar_DGV_Spec, DataGridView tar_DGV_Plus)
        {
            string[] col_name;
            DataTable tbl_spec = new DataTable();
            DataTable tbl_DGV_data = new DataTable();
            DataTable tbl_DGV_plus = new DataTable();
            DataTable sel_tbl = new DataTable();
            DataGridView sel_DGV;
            int tar_DGV_rows_count;//= tar_DGV.Rows.Count;
            int tar_DGV_Plus_rows_count;
            int sel_DGV_rows_count;
            int curr_DGV_row;
            string tar_Col = src_tbl.Columns[1].ColumnName;
            string vitri_col_name = src_tbl.Columns[0].ColumnName;
            string vitri_act_val = src_tbl.Columns[4].ColumnName;
            string vitri_Set_val = src_tbl.Columns[5].ColumnName;
            string vitri_UL_val = src_tbl.Columns[7].ColumnName;
            string vitri_LL_val = src_tbl.Columns[8].ColumnName;
            col_name = src_tbl.AsEnumerable().Select(r => r.Field<string>(tar_Col)).Distinct().ToArray();
            tbl_DGV_data = DGV_To_Table(tar_DGV);
            tbl_DGV_plus = DGV_To_Table(tar_DGV_Plus);
            tar_DGV_rows_count = tbl_DGV_data.Rows.Count;
            tar_DGV_Plus_rows_count = tbl_DGV_plus.Rows.Count;
            tbl_spec = DGV_To_Table(tar_DGV_Spec);
            foreach (string tg in col_name)
            {
                double UL = 0;// = Convert .ToDouble (tbl_spec.Rows [1][tg]);
                double LL = 0;// Convert.ToDouble(tbl_spec.Rows[2][tg]);
                bool spec_check_en = false;
                int sel_col_inx = 0;
                string col_dgv_data = tg;
                if (check_columns_existed_inx(tbl_spec, tg, ref sel_col_inx))
                {
                    double sv = Convert.ToDouble(tbl_spec.Rows[0][sel_col_inx]);
                    UL = sv + Convert.ToDouble(tbl_spec.Rows[1][sel_col_inx]);
                    LL = sv - Convert.ToDouble(tbl_spec.Rows[2][sel_col_inx]);
                    spec_check_en = true;
                    col_dgv_data = tbl_spec.Columns[sel_col_inx].ColumnName;
                    sel_DGV = tar_DGV;
                    sel_tbl = tbl_DGV_data;
                }
                else
                {
                    spec_check_en = false;
                    sel_DGV = tar_DGV_Plus;
                    sel_tbl = tbl_DGV_plus;
                }
                if (!check_columns_existed(sel_tbl, col_dgv_data))
                {
                    sel_DGV.Columns.Add(col_dgv_data, tg);
                }
                sel_DGV_rows_count = sel_tbl.Rows.Count;
                DataTable temp = new DataTable();
                temp = Get_data_detail(src_tbl, tar_Col, tg);
                int row_inx = temp.Rows.Count;
                curr_DGV_row = sel_DGV.RowCount;
                int sel_row_inx = curr_DGV_row - sel_DGV_rows_count;
                if (sel_row_inx < row_inx)
                {
                    int dgv_row = row_inx - sel_row_inx;
                    for (int i = 0; i < dgv_row; i++)
                    {
                        sel_DGV.Rows.Add();
                    }
                }
                string[] item_val;
                string[] vitri_val;
                item_val = temp.AsEnumerable().Select(r => r.Field<string>(vitri_act_val)).ToArray();
                vitri_val = temp.AsEnumerable().Select(r => r.Field<string>(vitri_col_name)).ToArray();
                int j = sel_DGV_rows_count;
                int vitri_inx = 0;
                foreach (string c in item_val)
                {
                    if (checkDBNull(c) != "")
                    {
                        if (vitri_val[vitri_inx].ToUpper().Contains("ANGLE"))
                        {
                            sel_DGV.Rows[j].Cells[col_dgv_data].Value = c;
                        }
                        else
                        {
                            sel_DGV.Rows[j].Cells[col_dgv_data].Value = Math.Abs(Convert.ToDouble(c)).ToString();
                        }
                        if (spec_check_en)
                        {
                            double act_val = Convert.ToDouble(sel_DGV.Rows[j].Cells[col_dgv_data].Value);
                            if ((act_val <= UL) && (act_val >= LL))
                            {
                                sel_DGV.Rows[j].Cells[col_dgv_data].Style.BackColor = Color.White;
                            }
                            else
                            {
                                if (act_val > UL)
                                {
                                    sel_DGV.Rows[j].Cells[col_dgv_data].Style.BackColor = Color.Yellow;
                                }
                                if (act_val < LL)
                                {
                                    sel_DGV.Rows[j].Cells[col_dgv_data].Style.BackColor = Color.LightBlue;
                                }
                                //confirm_request = true;
                            }
                        }
                        else

                        {
                            sel_DGV.Rows[j].Cells[col_dgv_data].Style.BackColor = Color.LightGray;
                        }
                    }
                    j++;
                    vitri_inx++;
                }
            }
            DGV_Auto_Resize(tar_DGV);
            DGV_Auto_Resize(tar_DGV_Plus);
        }
        public void DGV_Auto_Resize(DataGridView src_DGV)
        {
            for (int k = 0; k < src_DGV.RowCount; k++)
            {
                src_DGV.Rows[k].HeaderCell.Value = (k + 1).ToString();
            }
            src_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            src_DGV.AutoResizeColumns();
            Disable_Sort_DGV(src_DGV);
        }
        public bool insert_FAI_data_DGV(DataGridView src_DGV, DataGridView tar_DGV, DataTable src_tbl_spec)
        {
            bool _result = false;
            int sel_col_inx = src_DGV.SelectedCells[0].ColumnIndex;
            string src_col_name = src_DGV.Columns[sel_col_inx].Name;
            string src_col_header = src_col_name.Split('_')[0];
            if (check_columns_existed(src_tbl_spec, src_col_name))
            {
                if (!check_columns_existed(DGV_To_Table(tar_DGV), src_col_name))
                {
                    string UL = "";
                    string LL = "";
                    double SetVal = 0;
                    SetVal = Convert.ToDouble(src_tbl_spec.Rows[0][src_col_name]);
                    string _UL = IsNumeric_Val(checkDBNull(src_tbl_spec.Rows[1][src_col_name]));
                    string _LL = IsNumeric_Val(checkDBNull(src_tbl_spec.Rows[2][src_col_name]));

                    if (_UL != "")
                    {
                        UL = (SetVal + Convert.ToDouble(_UL)).ToString();
                    }
                    if (_LL != "")
                    {
                        LL = (SetVal - Convert.ToDouble(_LL)).ToString();
                    }
                    if (tar_DGV.Rows.Count > 0)
                    {
                        if (src_DGV.Rows.Count == tar_DGV.Rows.Count)
                        {
                            tar_DGV.Columns.Add(src_col_name, src_col_header);
                            for (int i = 0; i < src_DGV.Rows.Count; i++)
                            {
                                string act_val = checkDBNull(src_DGV.Rows[i].Cells[sel_col_inx].Value);
                                tar_DGV.Rows[i].Cells[src_col_name].Value = act_val;//src_DGV.Rows[i].Cells[sel_col_inx].Value;
                                tar_DGV.Rows[i].Cells[src_col_name].Style.BackColor = check_in_limit5(UL, LL, act_val, SetVal.ToString());
                            }
                            _result = true;
                        }
                        else
                        {
                            MessageBox.Show(new Form { TopMost = true },"Data size not suitable!", "Warning");
                            return false;
                        }
                    }
                    else
                    {
                        DataTable tbl_data = DGV_To_Table(tar_DGV);
                        tbl_data.Columns.Add(src_col_name);
                        for (int i = 0; i < src_DGV.Rows.Count; i++)
                        {
                            tbl_data.Rows.Add();
                            tbl_data.Rows[i][src_col_name] = src_DGV.Rows[i].Cells[sel_col_inx].Value;
                        }
                        tar_DGV.DataSource = tbl_data;
                        tar_DGV.Columns[src_col_name].HeaderText = src_col_header;
                        for (int i = 0; i < tar_DGV.RowCount; i++)
                        {
                            tar_DGV.Rows[i].HeaderCell.Value = (i + 1).ToString();
                            string act_val = checkDBNull(tar_DGV.Rows[i].Cells[src_col_name].Value);
                            tar_DGV.Rows[i].Cells[src_col_name].Style.BackColor = check_in_limit5(UL, LL, act_val, SetVal.ToString());                            
                        }
                        _result = true;
                    }
                    Disable_Sort_DGV(tar_DGV);
                    src_DGV.Columns.Remove(src_col_name);
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true },src_col_name + " is already existed, cannot add more!", "Warning");
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },"Please, change to suitable items!","Warning");
            }
            return _result;
        }
        public bool insert_FAI_data_DGV2(DataGridView src_DGV, DataGridView tar_DGV, DataTable src_tbl_spec)
        {
            bool _result = false;
            int sel_col_inx = src_DGV.SelectedCells[0].ColumnIndex;
            string src_col_name = src_DGV.Columns[sel_col_inx].Name;
            DataTable tar_tbl = DGV_To_Table(tar_DGV);
            if (check_columns_existed(src_tbl_spec, src_col_name))
            {
                if (!check_columns_existed(tar_tbl, src_col_name))
                {
                    if (tar_tbl.Rows.Count > 0)
                    {
                        if (tar_tbl.Rows.Count == tar_DGV.Rows.Count)
                        {
                            tar_tbl.Columns.Add(src_col_name);
                            for (int i = 0; i < tar_tbl.Rows.Count; i++)
                            {
                                string act_val = checkDBNull(src_DGV.Rows[i].Cells[sel_col_inx].Value);
                                tar_tbl.Rows[i][src_col_name]= act_val;
                            }
                            _result = true;
                        }
                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Data size not suitable!", "Warning");
                            return false;
                        }
                    }
                    else
                    {
                        tar_tbl.Columns.Add(src_col_name);
                        for (int i = 0; i < src_DGV.Rows.Count; i++)
                        {
                            tar_tbl.Rows.Add();
                            tar_tbl.Rows[i][src_col_name] = src_DGV.Rows[i].Cells[sel_col_inx].Value;
                        }
                        _result = true;
                    }
                    tar_DGV.DataSource = null;
                    tar_DGV.DataSource = Order_table_inSpec(tar_tbl, src_tbl_spec);
                    for (int i = 0; i < tar_DGV.RowCount; i++)
                    {
                        tar_DGV.Rows[i].HeaderCell.Value = (i + 1).ToString();
                    }
                    Disable_Sort_DGV(tar_DGV);
                    src_DGV.Columns.Remove(src_col_name);
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, src_col_name + " is already existed, cannot add more!", "Warning");
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Please, change to suitable items!", "Warning");
            }
            return _result;
        }
        public bool insert_IPQC_data_DGV(DataGridView src_DGV, DataGridView tar_DGV, DataTable src_tbl_spec)
        {
            bool _result = false;
            int sel_col_inx = src_DGV.SelectedCells[0].ColumnIndex;
            string src_col_name = src_DGV.Columns[sel_col_inx].Name;
            string src_col_header = src_DGV.Columns[sel_col_inx].HeaderText;
            bool spec_check_en = false;
            if (check_columns_existed(src_tbl_spec, src_col_name) && !src_DGV.Columns[src_col_name].ReadOnly)
            {
                if (!check_columns_existed(DGV_To_Table(tar_DGV), src_col_name))
                {
                    string sv = checkDBNull(src_tbl_spec.Rows[0][src_col_name]);
                    string TolMax = checkDBNull(src_tbl_spec.Rows[1][src_col_name]);
                    string TolMin = checkDBNull(src_tbl_spec.Rows[2][src_col_name]);
                    if ((TolMax=="")&&(TolMin==""))
                    {
                        spec_check_en = false;
                    }
                    else
                    {
                        spec_check_en = true;
                    }
                    string UL="";// =Convert.ToDouble(TolMax);
                    string LL="";// = Convert.ToDouble(TolMin);
                    bool insert_en = true;
                    for (int i = 0; i < src_DGV.Rows.Count; i++)
                    {
                        string act_val = checkDBNull(src_DGV.Rows[i].Cells[sel_col_inx].Value);
                        if (act_val == "")
                        {
                            insert_en = false;
                            break;
                        }
                    }
                    if(insert_en)
                    {
                        if (tar_DGV.Rows.Count > 0)
                        {
                            if (src_DGV.Rows.Count == tar_DGV.Rows.Count)
                            {

                                tar_DGV.Columns.Add(src_col_name, src_col_header);
                                for (int i = 0; i < src_DGV.Rows.Count; i++)
                                {
                                    string act_val = checkDBNull(src_DGV.Rows[i].Cells[sel_col_inx].Value);
                                    tar_DGV.Rows[i].Cells[src_col_name].Value = act_val;//src_DGV.Rows[i].Cells[sel_col_inx].Value;
                                    if ((spec_check_en) && (act_val != ""))
                                    {
                                        if (TolMax != "")
                                        {
                                            UL = TolMax;
                                        }
                                        if (TolMin != "")
                                        {
                                            LL = TolMin;
                                        }
                                        //tar_DGV.Rows[i].Cells[src_col_name].Style.BackColor = check_in_limit2(UL, LL, act_val);
                                        tar_DGV.Rows[i].Cells[src_col_name].Style.BackColor = check_in_limit5(UL, LL, act_val, sv);
                                    }

                                }
                                _result = true;
                            }
                            else
                            {
                                MessageBox.Show(new Form { TopMost = true },"Data size not suitable!", "Warning");
                                return false;
                            }
                        }
                        else
                        {
                            DataTable tbl_data = DGV_To_Table(tar_DGV);
                            tbl_data.Columns.Add(src_col_name);
                            for (int i = 0; i < src_DGV.Rows.Count; i++)
                            {
                                tbl_data.Rows.Add();
                                tbl_data.Rows[i][src_col_name] = src_DGV.Rows[i].Cells[sel_col_inx].Value;
                            }
                            tar_DGV.DataSource = tbl_data;
                            tar_DGV.Columns[src_col_name].HeaderText = src_col_header;
                            for (int i = 0; i < tar_DGV.RowCount; i++)
                            {
                                tar_DGV.Rows[i].HeaderCell.Value = (i + 1).ToString();
                                string act_val = checkDBNull(tar_DGV.Rows[i].Cells[src_col_name].Value);
                                if (spec_check_en && (act_val != ""))
                                {
                                    tar_DGV.Rows[i].Cells[src_col_name].Style.BackColor = check_in_limit5(UL, LL, act_val, sv);
                                }

                            }
                            _result = true;
                        }
                        Disable_Sort_DGV(tar_DGV);
                        src_DGV.Columns.Remove(src_col_name);
                    }
                    else
                    {
                        _result = false;
                        MessageBox.Show(new Form { TopMost = true },src_col_header + " is not filled enough data!", "Warning");
                    }
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true },src_col_header + " is already existed, cannot add more!", "Warning");
                }
            }
            else
            {
                if(src_DGV.Columns[src_col_name].ReadOnly)
                {
                    MessageBox.Show(new Form { TopMost = true },"Data is existed, cannot add more!", "Warning");
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true },"Please, change to suitable items!", "Warning");
                }
                
            }
            return _result;
        }
        public void Setup_IPQC_Spec(string f_name, string tar_ItemCode, SqlConnection tar_sqlcon)
        {
            myExcel.Workbook cur_wrk = TDMK_Code.open_excel_file(f_name, "", "");
            myExcel.Worksheet cur_wrksht = cur_wrk.Sheets["IPQC Data"];
            myExcel.Range item_rgn = cur_wrksht.Range["B13"];
            myExcel.Range spec_rgn = cur_wrksht.Range["B14"];
            
            DataSet myDS = new DataSet();
            TDMK_Code.fill_dataset(myDS, "All_Items", tar_sqlcon);
            DataTable Spec_addr_tbl = myDS.Tables[0];
            List<string> itemLst = Spec_addr_tbl.AsEnumerable().Select(r => r.Field<string>("Report_Name")).ToList();
            List<string> ColName_List = Spec_addr_tbl.AsEnumerable().Select(r => r.Field<string>("Col_Name")).ToList();
            List<string> Process_name_lst = Spec_addr_tbl.AsEnumerable().Select(r => r.Field<string>("Process_Name")).ToList();
            string[] items = new string[] { "ID", "ItemCode", "Item", "Spec", "Address", "Remark", "Col_Name" };
            int inx = 0;
            int arr_inx = 0;
            AutoCompleteStringCollection ItemCode_lst = TDMK_Code.Load_Item_Names_Filter(tar_sqlcon, "SpecList", "ItemCode", "ItemCode = '" + tar_ItemCode + "'");
            if (ItemCode_lst.Count == 0)
            {
                while (checkDBNull(item_rgn.Offset[0, inx].Value) != "")
                {
                    string rgn_val = item_rgn.Offset[0, inx].Value;
                    if (checkDBNull(spec_rgn.Offset[0, inx].Value)!="")
                    {
                        string spec_val = spec_rgn.Offset[0, inx].Value;
                        spec_val = spec_val.Replace("\u2265", ">=");
                        spec_val = spec_val.Replace("\u2264", "<=");
                        int sel_inx = TDMK_Code.check_exist_list_index2(rgn_val, itemLst);
                        if (sel_inx != -1)
                        {
                            string[] item_vals = new string[7];
                            item_vals[1] = tar_ItemCode;
                            item_vals[0] = (TDMK_Code.SQL_MAX("SpecList", "ID", tar_sqlcon) + 1).ToString();
                            item_vals[2] = itemLst[sel_inx];
                            item_vals[3] = spec_val;
                            item_vals[4] = spec_rgn.Offset[1, inx].AddressLocal.ToString();
                            item_vals[5] = Process_name_lst[sel_inx];
                            item_vals[6] = ColName_List[sel_inx];
                            TDMK_Code.insert_val_arr("SpecList", tar_sqlcon, items, item_vals);
                            itemLst.RemoveAt(sel_inx);
                            ColName_List.RemoveAt(sel_inx);
                            Process_name_lst.RemoveAt(sel_inx);
                            arr_inx++;
                        }
                    }
                    inx++;
                }
            }
            cur_wrk.Close();
        }
        public bool IsNumeric(string text)
        {
            double test;
            return double.TryParse(text, out test);
        }
        public string IsNumeric_Val(string text)
        {
            double test;
            if(double.TryParse(text, out test))
            {
                return test.ToString();
            }
            else
            {
                return "";
            }
            //return double.TryParse(text, out test);
        }
        public void CopyToClipboard(DataGridView tar_DGV)
        {
            //Copy to clipboard
            DataObject dataObj = tar_DGV.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }
        public void PasteClipboardValue(bool _transpose, DataGridView tar_DGV, DataGridView tar_DGV_Spec)
        {
            //Show Error if no cell is selected
            if (tar_DGV.SelectedCells.Count == 0)
            {
                MessageBox.Show(new Form { TopMost = true },"Please select a cell", "Paste", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            //Get the satring Cell
            DataGridViewCell startCell = GetStartCell(tar_DGV);
            //Get the clipboard value in a dictionary
            Dictionary<int, Dictionary<int, string>> cbValue = ClipBoardValues(Clipboard.GetText());
            if (cbValue.Count > 0)
            {
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
                                DataGridViewCell cell = tar_DGV[iColIndex, iRowIndex];
                                //Copy to selected cells if 'chkPasteToSelectedCells' is checked
                                if (cell.Selected)
                                {
                                    cell.Value = cbValue[rowKey][cellKey];
                                    try
                                    {
                                        string curr_col_name = tar_DGV.Columns[iColIndex].Name.ToString();
                                        if (check_columns_existed(DGV_To_Table(tar_DGV_Spec), curr_col_name))
                                        {
                                            string UL = checkDBNull(tar_DGV_Spec.Rows[1].Cells[curr_col_name].Value);
                                            string LL = checkDBNull(tar_DGV_Spec.Rows[2].Cells[curr_col_name].Value);
                                            string SV = checkDBNull(tar_DGV_Spec.Rows[0].Cells[curr_col_name].Value);
                                            double _UL = Convert.ToDouble(SV) + Convert.ToDouble(UL);
                                            double _LL = Convert.ToDouble(SV) - Convert.ToDouble(LL);
                                            string act_val = checkDBNull(cell.Value);
                                            cell.Style.BackColor = check_in_limit5(UL, LL, act_val, SV);
                                        }
                                    }
                                    catch
                                    {

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
                                DataGridViewCell cell = tar_DGV[iColIndex, iRowIndex];
                                //Copy to selected cells if 'chkPasteToSelectedCells' is checked
                                if (cell.Selected)
                                {
                                    cell.Value = cbValue[rowKey][cellKey];
                                    try
                                    {
                                        string curr_col_name = tar_DGV.Columns[iColIndex].Name.ToString();
                                        if (check_columns_existed(DGV_To_Table(tar_DGV_Spec), curr_col_name))
                                        {
                                            string UL = checkDBNull(tar_DGV_Spec.Rows[1].Cells[curr_col_name].Value);
                                            string LL = checkDBNull(tar_DGV_Spec.Rows[2].Cells[curr_col_name].Value);
                                            string SV = checkDBNull(tar_DGV_Spec.Rows[0].Cells[curr_col_name].Value);
                                            double _UL = Convert.ToDouble(SV) + Convert.ToDouble(UL);
                                            double _LL = Convert.ToDouble(SV) - Convert.ToDouble(LL);
                                            string act_val = checkDBNull(cell.Value);
                                            cell.Style.BackColor = check_in_limit5(UL, LL, act_val, SV);
                                        }
                                    }
                                    catch
                                    {

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
                MessageBox.Show(new Form { TopMost = true },"PLease, select data!", "Warning");
            }

        }
        public DataGridViewCell GetStartCell(DataGridView dgView)
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
        public Dictionary<int, Dictionary<int, string>> ClipBoardValues(string clipboardValue)
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
        public List<string>[] Roughness_Result(string f_name, int MeasLoc)
        {
            List<string> mylstSa1 = new List<string>();
            List<string> mylstSa2 = new List<string>();
            List<string> mylstSq2 = new List<string>();
            List<string> mylstSq1 = new List<string>();
            List<string>[] Sa_result = new List<string>[] { mylstSa1, mylstSa2 };
            List<string>[] Sq_result = new List<string>[] { mylstSq1, mylstSq2 };
            List<string>[] _result = new List<string>[] { mylstSa1, mylstSa2, mylstSq1, mylstSq2 };
            myExcel.Workbook tar_wrkbook = TDMK_Code.open_excel_file(f_name, "", "");
            myExcel.Worksheet tar_wrksht = tar_wrkbook.Sheets[1];
            myExcel.Range tar_rgn = tar_wrksht.Range["B2"];
            int inx = 0;
            while (checkDBNull(tar_rgn.Offset[inx, 0].Value) != "")
            {
                for (int i = 0; i <= MeasLoc - 1; i++)
                {
                    if (i < Sa_result.Length)
                    {
                        int cur_val = inx - 2 * i;
                        int div_val = cur_val / (2 * MeasLoc);
                        int hieuso = cur_val - div_val * 2 * MeasLoc;
                        if (hieuso == 0)
                        {
                            double Sa_val = Convert.ToDouble(checkDBNull(tar_rgn.Offset[inx, 0].Value));
                            double Sq_val = Convert.ToDouble(checkDBNull(tar_rgn.Offset[inx, 1].Value));
                            Sa_result[i].Add(Math.Round(Sa_val, 3).ToString());//"#0.##0"
                            Sq_result[i].Add(Math.Round(Sq_val, 3).ToString());//"#0.##0"
                        }
                    }
                }
                inx++;
            }
            tar_wrkbook.Close();
            return _result;
        }
        public List<string>[] Roughness_Result_ACF(string f_name, int MeasLoc)
        {
            List<string> mylstSa1 = new List<string>();
            List<string> mylstSa2 = new List<string>();
            List<string> mylstSa3 = new List<string>();

            List<string> mylstSq1 = new List<string>();
            List<string> mylstSq2 = new List<string>();
            List<string> mylstSq3 = new List<string>();

            List<string> mylstSdr1 = new List<string>();
            List<string> mylstSdr2 = new List<string>();
            List<string> mylstSdr3 = new List<string>();

            List<string>[] Sa_result = new List<string>[] { mylstSa1, mylstSa2, mylstSa3 };
            List<string>[] Sq_result = new List<string>[] { mylstSq1, mylstSq2, mylstSq3 };
            List<string>[] Sdr_result = new List<string>[] { mylstSdr1, mylstSdr2, mylstSdr3 };

            List<string>[] _result = new List<string>[] { mylstSa1, mylstSa2, mylstSq1, mylstSq2, mylstSdr1, mylstSdr2, mylstSa3,mylstSq3,mylstSdr3 };
            myExcel.Workbook tar_wrkbook = TDMK_Code.open_excel_file(f_name, "", "");
            myExcel.Worksheet tar_wrksht = tar_wrkbook.Sheets[1];
            myExcel.Range tar_rgn = tar_wrksht.Range["B2"];
            int inx = 0;
            while (checkDBNull(tar_rgn.Offset[inx, 0].Value) != "")
            {
                for (int i = 0; i <= MeasLoc - 1; i++)
                {
                    int cur_val = inx - 2 * i;
                    int div_val = cur_val / (2 * MeasLoc);
                    int hieuso = cur_val - div_val * 2 * MeasLoc;
                    if (hieuso == 0)
                    {
                        double Sa_val = Convert.ToDouble(checkDBNull(tar_rgn.Offset[inx, 0].Value));
                        double Sq_val = Convert.ToDouble(checkDBNull(tar_rgn.Offset[inx, 1].Value));
                        double Sdr_val = Convert.ToDouble(checkDBNull(tar_rgn.Offset[inx+1, 3].Value));
                        Sa_result[i].Add(Math.Round(Sa_val, 3).ToString());//"#0.##0"
                        Sq_result[i].Add(Math.Round(Sq_val, 3).ToString());//"#0.##0"
                        Sdr_result[i].Add(Math.Round(Sdr_val, 3).ToString());//"#0.##0"
                    }
                }
                inx++;
            }
            tar_wrkbook.Close();
            return _result;
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

        public void Mitutoyo_process(string _filename, DataGridView tar_DGV, DataGridView DGV_Spec, DataGridView DGV_plus, bool edit_mode)
        {
            string f_name = _filename;
            DataTable logfile_dt = new DataTable();
            logfile_dt = Mitutoyo_Get_raw_data2(f_name);
            DataTable FAI_logfile_tbl = FAI_raw_tbl(logfile_dt);
            DataTable tbl_spec = (DataTable)DGV_Spec.DataSource;
            DataTable FAI_tbl = new DataTable();
            DataTable non_FAI = new DataTable();
            DataTable FAI_arranged_tbl = filter_table(FAI_logfile_tbl, tbl_spec, ref FAI_tbl, ref non_FAI);
            if (!edit_mode)
            {
                DataTable dest_FAI_tbl = result_add((DataTable)tar_DGV.DataSource, FAI_tbl, true);
                DataTable dest_nonFAI_tbl = result_add((DataTable)DGV_plus.DataSource, non_FAI, true);
                tar_DGV.DataSource = null;
                DGV_plus.DataSource = null;
                tar_DGV.DataSource = Order_table_inSpec(dest_FAI_tbl, tbl_spec);
                DGV_plus.DataSource = dest_nonFAI_tbl;
            }
            else
            {
                DGV_plus.DataSource = null;
                DGV_plus.DataSource = Order_table_inSpec(FAI_arranged_tbl, tbl_spec);

            }
            check_FAIdata_inSpec(tar_DGV, DGV_Spec);
            check_FAIdata_inSpec(DGV_plus, DGV_Spec);
            DGV_Auto_Resize(tar_DGV);
            DGV_Auto_Resize(DGV_plus);

        }
        public DataTable result_add(DataTable dest_tbl, DataTable src_inData, bool col_add_en)
        {
            if (dest_tbl != null)
            {
                int dest_row_count = dest_tbl.Rows.Count;
                foreach (DataColumn dc in src_inData.Columns)
                {
                    int start_row = 0;
                    if (check_columns_existed(dest_tbl, dc.ColumnName))
                    {
                        DataView dv = dest_tbl.AsDataView();
                        string filter_str = "[" + dc.ColumnName + "] is null";
                        dv.RowFilter = filter_str;
                        if (dv.Count > 0)
                        {
                            DataRow dr = dv[0].Row;
                            start_row = dest_tbl.Rows.IndexOf(dr);
                        }
                        else
                        {
                            start_row = dest_row_count;
                        }
                    }
                    else
                    {
                        if (col_add_en)
                        {
                            dest_tbl.Columns.Add(dc.ColumnName);
                        }
                        else
                        {
                            continue;
                        }
                    }
                    for (int i = 0; i < src_inData.Rows.Count; i++)
                    {
                        int cur_row_inx = start_row + i;
                        if (dest_tbl.Rows.Count <= cur_row_inx)
                        {
                            dest_tbl.Rows.Add();
                        }
                        dest_tbl.Rows[cur_row_inx][dc.ColumnName] = src_inData.Rows[i][dc];
                    }
                }
                return dest_tbl;
            }
            else
            {
                return src_inData;
            }
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
                        dr.Cells[cur_col].Style.BackColor = check_in_limit5(UL.ToString(), LL.ToString(), src_act_val, SetVal);
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
                    string Tol_Max = checkDBNull(src_Spec.Rows[1][cur_col]);
                    string Tol_Min = checkDBNull(src_Spec.Rows[2][cur_col]);
                    string SetVal = checkDBNull(src_Spec.Rows[0][cur_col]);
                    double UL = Convert.ToDouble(SetVal) + Convert.ToDouble(Tol_Max);
                    double LL = Convert.ToDouble(SetVal) - Convert.ToDouble(Tol_Min);
                    foreach (DataRow dr in src_data.Rows)
                    {
                        string src_act_val = checkDBNull(dr[cur_col]);
                        if(!check_in_limit(UL.ToString(), LL.ToString(), src_act_val, SetVal))
                        {
                            result = false;
                            break;
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
        public DataTable FAI_raw_tbl(DataTable logfile_dt)
        {
            DataTable result_tbl = new DataTable();
            string[] col_name;
            string tar_Col = logfile_dt.Columns[1].ColumnName;
            string vitri_col_name = logfile_dt.Columns[0].ColumnName;
            string vitri_act_val = logfile_dt.Columns[4].ColumnName;
            string vitri_Set_val = logfile_dt.Columns[5].ColumnName;
            string vitri_UL_val = logfile_dt.Columns[7].ColumnName;
            string vitri_LL_val = logfile_dt.Columns[8].ColumnName;
            //foreach (DataRow dr in logfile_dt.Rows)
            //{
            //    try
            //    {
            //        dr[vitri_Set_val] = Convert.ToDouble(dr[vitri_Set_val]);
            //    }
            //    catch
            //    {
            //        dr[vitri_Set_val] = "Error";
            //    }
            //}
            col_name = logfile_dt.AsEnumerable().Select(r => r.Field<string>(tar_Col)).Distinct().ToArray();
            List<string> col_lst = col_name.ToList();
            List<DataTable> src_tbl_lst = new List<DataTable>();
            Get_ListTable(-1, logfile_dt, new string[] { tar_Col }, ref src_tbl_lst, vitri_act_val);
            Dictionary<string, List<string>> dicFAI_data = new Dictionary<string, List<string>>();
            foreach (DataTable dt in src_tbl_lst)
            {
                List<string> setval_lst = dt.AsEnumerable().Where(x => x.Field<string>(vitri_Set_val) != "").Select(x => x.Field<string>(vitri_Set_val)).Distinct().ToList();
                string fai_name = dt.Rows[0][tar_Col].ToString();
                string setval = setval_lst[0];// dt.Rows[0][vitri_Set_val].ToString();
                if(IsNumeric(dt.Rows[0][vitri_Set_val].ToString()))
                {
                    setval= Math.Abs(Convert.ToDecimal(dt.Rows[0][vitri_Set_val])).ToString();
                }
                string FAI = fai_name + "_" + setval;
                string vitri_val = dt.Rows[0][vitri_col_name].ToString();
                result_tbl.Columns.Add(FAI);
                List<string> data = dt.AsEnumerable().Select(x => x.Field<string>(vitri_act_val)).ToList();
                if (!vitri_val.ToUpper().Contains("ANGLE"))
                {
                    for (int i = 0; i < data.Count; i++)
                    {
                        if(IsNumeric(data[i]))
                        {
                            data[i] = Math.Abs(Convert.ToDouble(data[i])).ToString();
                        }
                    }
                }
                if (dicFAI_data.ContainsKey(FAI))
                {
                    var temp_dic = dicFAI_data[FAI];
                    temp_dic.AddRange(data);
                }
                else
                {
                    dicFAI_data.Add(FAI, data);
                }
            }
            foreach (var fai in dicFAI_data)
            {
                int row_add = fai.Value.Count - result_tbl.Rows.Count;
                if (row_add > 0)
                {
                    for (int i = 0; i < row_add; i++)
                    {
                        result_tbl.Rows.Add();
                    }
                }
                int r_inx = 0;
                foreach (var item in fai.Value)
                {
                    result_tbl.Rows[r_inx][fai.Key] = item;
                    r_inx++;
                }
            }
            return result_tbl;
        }
        public DataTable filter_table(DataTable FAI_dt, DataTable spec_dt, ref DataTable outFAI_dt, ref DataTable outUnknowndt)
        {
            List<string> spec_col_lst = new List<string>();
            foreach (DataColumn dc in spec_dt.Columns)
            {
                spec_col_lst.Add(dc.ColumnName);
            }
            start_label: List<string> data_col_lst = new List<string>();
            foreach (DataColumn dc in FAI_dt.Columns)
            {
                data_col_lst.Add(dc.ColumnName);
            }
            var tar_col_lst = spec_col_lst.Intersect(data_col_lst).ToList();
            var ex_col = data_col_lst.Except(spec_col_lst).ToList();
            bool en = false;
            foreach (var t in ex_col)
            {
                string t_sv = IsNumeric_Val(t.Split('_').Last());
                string t_fai = t.Split('_').First();
                int col_inx = -1;
                if (check_columns_existed_inx(spec_dt, t_fai, ref col_inx))
                {
                    string tar_sv = IsNumeric_Val(spec_col_lst[col_inx].Split('_').Last());
                    string tar_fai = spec_col_lst[col_inx].Split('_').First();
                    if ((t_sv != "") && (tar_sv != "") && (tar_fai == t_fai) && (t_sv == tar_sv))
                    {
                        int fai_col_inx = data_col_lst.IndexOf(t);
                        FAI_dt.Columns[fai_col_inx].ColumnName = spec_col_lst[col_inx];
                        en = true;
                    }
                }
            }
            if (en)
            {
                goto start_label;
            }
            //for (int i = 0; i < tar_col_lst.Count; i++)
            //{
            //    string fai_col = tar_col_lst[i];
            //    List<string> sel_fai_col = FAI_dt.AsEnumerable().Select(x => x.Field<string>(fai_col)).ToList();
            //    if (sel_fai_col.Any(x => !IsNumeric(x)))
            //    {
            //        tar_col_lst.Remove(fai_col);
            //        if (ex_col.IndexOf(fai_col) == -1)
            //        {
            //            ex_col.Add(fai_col);
            //        }
            //    }
            //}
            if (tar_col_lst.Count!=0)
            {
                outFAI_dt = FAI_dt.AsDataView().ToTable(false, tar_col_lst.ToArray());
            }
            if(ex_col.Count!=0)
            {
                outUnknowndt = FAI_dt.AsDataView().ToTable(false, ex_col.ToArray());
            }
            tar_col_lst.AddRange(ex_col);
            return FAI_dt.AsDataView().ToTable(false, tar_col_lst.ToArray());
        }
        public DataTable Order_table_inSpec(DataTable src_dt, DataTable spec_dt)
        {
            List<string> spec_col_lst = new List<string>();
            foreach (DataColumn dc in spec_dt.Columns)
            {
                spec_col_lst.Add(dc.ColumnName);
            }
            start_label: List<string> data_col_lst = new List<string>();
            foreach (DataColumn dc in src_dt.Columns)
            {
                data_col_lst.Add(dc.ColumnName);
            }
            var tar_col_lst = spec_col_lst.Intersect(data_col_lst).ToList();
            var ex_col = data_col_lst.Except(spec_col_lst).ToList();
            bool en = false;
            foreach (var t in ex_col)
            {
                string t_sv = IsNumeric_Val(t.Split('_').Last());
                string t_fai = t.Split('_').First();
                int col_inx = -1;
                if (check_columns_existed_inx(spec_dt, t_fai, ref col_inx))
                {
                    string tar_sv = IsNumeric_Val(spec_col_lst[col_inx].Split('_').Last());
                    string tar_fai = spec_col_lst[col_inx].Split('_').First();
                    if ((t_sv != "") && (tar_sv != "") && (t_fai == tar_fai) && (t_sv == tar_sv))
                    {
                        int fai_col_inx = data_col_lst.IndexOf(t);
                        src_dt.Columns[fai_col_inx].ColumnName = spec_col_lst[col_inx];
                        en = true;
                    }
                }
            }
            if (en)
            {
                goto start_label;
            }
            tar_col_lst.AddRange(ex_col);
            return src_dt.AsDataView().ToTable(false, tar_col_lst.ToArray());
        }
        public void Clear_DGV(DataGridView src_DGV)
        {
            if(src_DGV.DataSource!=null)
            {
                src_DGV.DataSource = null;
            }
            else
            {
                if(src_DGV.Columns.Count>0)
                {
                    src_DGV.Rows.Clear();
                    src_DGV.Columns.Clear();
                }
            }
        }
        public string Lotno_Formated(string lotno)
        {
            string result = "";
            if (lotno.All(char.IsDigit))
            {
                result = string.Format("{0:00000}", Convert.ToInt32(lotno));
                result = result.Substring(0, 5);
            }
            else
            {
                if (lotno.Contains('-'))
                {
                    string lotno1 = lotno.Split('-')[0];
                    string cutno = lotno.Split('-')[1];
                    if (lotno1.All(char.IsDigit) && cutno.All(char.IsDigit))
                    {
                        string lotno2 = string.Format("{0:00000}", Convert.ToInt32(lotno1));
                        string cutno2 = string.Format("{0:00}", Convert.ToInt32(cutno));
                        result = lotno2.Substring(0, 5) + "-" + cutno2.Substring(0, 2);
                    }
                }
            }
            return result;
        }
        public Bitmap Draw_Histogram(double[] src_bin, int[] src_freq_bin, double[] src_modified, SEI_Lib.FAI_Spec src_FAI_spec)
        {
            ZedGraph.ZedGraphControl tar_Graph = new ZedGraphControl();
            GraphPane myPane = tar_Graph.GraphPane;
            PointPairList NormDistList = new PointPairList();
            PointPairList FAIPointList = new PointPairList();
            PointPairList USL_list = new PointPairList();
            PointPairList LSL_list = new PointPairList();

            for (int i = 0; i < 50; i++)
            {
                double x = src_bin[i];
                double y = src_freq_bin[i];
                double y2 = src_modified[i];
                FAIPointList.Add(x, y);
                NormDistList.Add(x, y2);
            }
            double freq_bin_Min = src_freq_bin.Min();
            double freq_bin_Max = src_freq_bin.Max();
            if (IsNumeric_Val(src_FAI_spec.UL) != "")
            {
                double UL = Convert.ToDouble(src_FAI_spec.SetVal) + Convert.ToDouble(src_FAI_spec.UL);
                USL_list.Add(UL, freq_bin_Min);
                USL_list.Add(UL, freq_bin_Max);
                LineItem UL_curve = myPane.AddCurve("USL", USL_list, Color.Red, SymbolType.None);
            }
            if (IsNumeric_Val(src_FAI_spec.LL) != "")
            {
                double LL = Convert.ToDouble(src_FAI_spec.SetVal) - Convert.ToDouble(src_FAI_spec.LL);
                LSL_list.Add(LL, freq_bin_Min);
                LSL_list.Add(LL, freq_bin_Max);
                LineItem LL_curve = myPane.AddCurve("LSL", LSL_list, Color.Blue, SymbolType.None);
            }
            BarItem FAI_Curve = myPane.AddBar(src_FAI_spec.FAI_Name, FAIPointList, Color.Gray);
            LineItem NormDist_Curve = myPane.AddCurve("NormDist", NormDistList, Color.DarkBlue, SymbolType.Square);
            NormDist_Curve.Line.Width = 3;
            tar_Graph.IsShowPointValues = true;
            myPane.XAxis.Scale.Min = src_bin.Min();
            myPane.XAxis.Scale.Max = src_bin.Max();
            myPane.XAxis.Scale.MajorStep = (src_bin.Max() - src_bin.Min()) / 50;
            myPane.Title.Text = src_FAI_spec.FAI_Name + "_Histogram";
            tar_Graph.AxisChange();
            tar_Graph.Invalidate();
            return myPane.GetImage(300, 100, 1200);
        }
        public byte[] imgToByteConverter(Image inImg)
        {
            ImageConverter imgCon = new ImageConverter();
            return (byte[])imgCon.ConvertTo(inImg, typeof(byte[]));
        }
        public DataTable Result_Histogram(DataTable src_dt)
        {
            DataTable histo_dt = new DataTable();
            List<string> col_list = new List<string>();
            foreach (DataColumn dc in src_dt.Columns)
            {
                if (!col_list.Contains(dc.ColumnName))
                {
                    col_list.Add(dc.ColumnName);
                    DataColumn column = new DataColumn(dc.ColumnName);
                    column.DataType = System.Type.GetType("System.Byte[]");
                    histo_dt.Columns.Add(column);
                }
            }
            histo_dt.Rows.Add();
            int arr_num = col_list.Count;
            SEI_Lib.Calculate_CPK[] myCalc_CPK = new SEI_Lib.Calculate_CPK[arr_num];
            SEI_Lib.FAI_Histogram_Data[] myHistogram_data = new SEI_Lib.FAI_Histogram_Data[arr_num];
            int col_list_inx = 0;
            List<Bitmap> histo_lt = new List<Bitmap>();
            foreach (string t in col_list)
            {
                foreach (SEI_Lib.FAI_Spec ref_spec in testFAI_spec)
                {
                    if (ref_spec.FAI_Name == t)
                    {
                        string[] temp_FAI_data = src_dt.AsEnumerable().Where(x => x.Field<string>(t) != null).Select(r => r.Field<string>(t)).ToArray();
                        if (temp_FAI_data.Length > 0)
                        {
                            Double[] data_arr = new double[temp_FAI_data.Length];
                            int inx = 0;
                            foreach (string c in temp_FAI_data)
                            {
                                if (checkDBNull(c) != "")
                                {
                                    data_arr[inx] = Convert.ToDouble(c);
                                    inx++;
                                }
                            }
                            Array.Resize(ref data_arr, inx);
                            Calcul_CPK_FAI(ref_spec, temp_FAI_data, ref myCalc_CPK[col_list_inx], ref myHistogram_data[col_list_inx]);
                            Bitmap cur_hist = Draw_Histogram(myHistogram_data[col_list_inx]._Bin_data, myHistogram_data[col_list_inx]._Freq_bin_data, myHistogram_data[col_list_inx]._Modified_NormDist_data, ref_spec);
                            histo_dt.Rows[0][t] = imgToByteConverter(cur_hist);
                            histo_lt.Add(cur_hist);
                        }
                        break;
                    }
                }
                col_list_inx++;
            }
            return histo_dt;
        }
        public void Calcul_CPK_FAI2(SEI_Lib.FAI_Spec src_FAI_spec, Double[] data_arr, ref SEI_Lib.Calculate_CPK tar_CPK_Result, ref SEI_Lib.FAI_Histogram_Data tar_Histogram_Data)
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

            if (IsNumeric(src_FAI_spec.UL))
            {
                UL = Norminal_Dim + Convert.ToDouble(src_FAI_spec.UL);
                CPKU = (UL - mean) / (3 * stdev);
            }
            if (IsNumeric(src_FAI_spec.LL))
            {
                LL = Norminal_Dim - Convert.ToDouble(src_FAI_spec.LL);
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
        public void Check_Data_inSpec(DataGridView src_DGV_Data, DataGridView src_DGV_Spec)
        {
            foreach (DataGridViewColumn dgv_col in src_DGV_Data.Columns)
            {
                string col_name = dgv_col.Name;
                //double sv = Convert.ToDouble(src_DGV_Spec.Rows[0].Cells[col_name].Value);
                //double UL = sv + Convert.ToDouble(src_DGV_Spec.Rows[1].Cells[col_name].Value);
                //double LL = sv - Convert.ToDouble(src_DGV_Spec.Rows[2].Cells[col_name].Value);
                string sv = checkDBNull(src_DGV_Spec.Rows[0].Cells[col_name].Value);
                string UL = "";
                string LL = "";
                if (IsNumeric(sv))
                {
                    double SetVal = Convert.ToDouble(sv);
                    string _UL = checkDBNull(src_DGV_Spec.Rows[1].Cells[col_name].Value);
                    string _LL = checkDBNull(src_DGV_Spec.Rows[2].Cells[col_name].Value);
                    if (IsNumeric(_UL))
                    {
                        UL = (SetVal + Convert.ToDouble(_UL)).ToString();
                    }
                    if (IsNumeric(_LL))
                    {
                        LL = (SetVal - Convert.ToDouble(_LL)).ToString();
                    }
                }
                foreach (DataGridViewRow dgv_row in src_DGV_Data.Rows)
                {
                    string act_val = checkDBNull(dgv_row.Cells[col_name].Value);
                    //dgv_row.Cells[col_name].Style.BackColor = check_in_limit2(UL.ToString(), LL.ToString(), act_val);
                    dgv_row.Cells[col_name].Style.BackColor = check_in_limit5(UL, LL, act_val, sv);
                }
            }
        }
        public bool IPQC_Process_result(string tar_ItemCode, string tar_LotNo, SqlConnection tar_sqlcon, string tar_process)
        {
            bool _result = false;
            string flt_str = "";
            DataTable tbl_all = new DataTable();
            switch (tar_process)
            {
                case "Roughness":
                    flt_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Col_Name" }, new string[] { tar_ItemCode, "%Roughness%" });
                    tbl_all = TDMK_Code.Datatable_Filter(tar_sqlcon, "All_Items", "Col_Name like '%Roughness%'");
                    break;
                case "AU_NI":
                    flt_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Col_Name" }, new string[] { tar_ItemCode, "%THICKNESS%" });
                    tbl_all = TDMK_Code.Datatable_Filter(tar_sqlcon, "All_Items", "Col_Name like '%THICKNESS%'");
                    break;
                default:
                    flt_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { tar_ItemCode, tar_process });
                    tbl_all = TDMK_Code.Datatable_Filter(tar_sqlcon, "All_Items", "Process_Name = '" + tar_process + "'");
                    break;
            }
            DataTable sel_dt = TDMK_Code.Datatable_Filter(tar_sqlcon, "SpecList", flt_str);
            DataTable spec_tbl = new DataTable();
            if (sel_dt.Rows.Count > 0)
            {
                string[] Col_name = sel_dt.AsEnumerable().Select(r => r.Field<string>("Col_name")).ToArray();
                string[] col_header = new string[Col_name.Length];
                int inx = 0;
                foreach (string c_name in Col_name)
                {
                    string[] item_name = tbl_all.AsEnumerable().Where(r => r.Field<string>("Col_name") == c_name).Select(r => r.Field<string>("Item_Name")).ToArray();
                    col_header[inx] = item_name[0];
                    if (!check_columns_existed(spec_tbl, c_name))
                    {
                        spec_tbl.Columns.Add(c_name);
                    }
                    inx++;
                }
                Fill_IPQC_Spec_Table(sel_dt, spec_tbl);
                DataTable sel_tbl_data = TDMK_Code.Datatable_Filter(tar_sqlcon, tar_process, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { tar_ItemCode, tar_LotNo }));
                DataTable main_data = sel_tbl_data.AsDataView().ToTable(false, Col_name);
                _result = Check_inSpec_IPQC_Data(main_data, spec_tbl);
            }
            return _result;
        }
        public void Fill_IPQC_Spec_Table(DataTable src_Spec_tbl, DataTable tar_Spec_tbl)
        {
            string[] limit_lst = src_Spec_tbl.AsEnumerable().Select(r => r.Field<string>("Spec")).ToArray();
            string[] colname_list = src_Spec_tbl.AsEnumerable().Select(r => r.Field<string>("Col_Name")).ToArray();
            string[] col_header = new string[tar_Spec_tbl.Columns.Count];
            //char[] split_chr = { '~', '<', '\u2264', '\u2265', '\u00B1', '>' };
            //char[] trim_chr = { 'u', 'm', '\u03BC', '\u00B5', ' ', '=', '-', '+' };

            char[] split_chr = { '~', '<', '\u2264', '\u2265', '\u00B1', '>' };
            char[] trim_chr = { 'u', 'm', '\u03BC', '\u00B5', ' ', '=', '-' };

            int inx = 0;
            foreach (DataColumn c in tar_Spec_tbl.Columns)
            {
                string t = c.ColumnName;
                string[] temp2 = src_Spec_tbl.AsEnumerable().Where(r => r.Field<string>("Col_Name") == t).Select(r => r.Field<string>("Spec")).ToArray();
                limit_lst[inx] = temp2[0];
                inx++;
            }
            Array.Resize(ref limit_lst, inx);
            DataRow dr_SV = tar_Spec_tbl.NewRow();
            DataRow dr_UL = tar_Spec_tbl.NewRow();
            DataRow dr_LL = tar_Spec_tbl.NewRow();
            inx = 0;
            foreach (string _limit in limit_lst)
            {
                if (_limit.Contains('/'))
                {
                    string[] temp = _limit.Split('/');
                    double sv;
                    double ul;
                    double ll;
                    if (temp[0].Contains('-'))
                    {
                        sv = Convert.ToDouble(temp[0].Split('-')[0].Trim(trim_chr));
                        ll = Convert.ToDouble(temp[0].Split('-')[1].Trim(trim_chr));
                        ul = Convert.ToDouble(temp[1].Trim(trim_chr));
                    }
                    else
                    {
                        sv = Convert.ToDouble(temp[0].Split('+')[0].Trim(trim_chr));
                        ul = Convert.ToDouble(temp[0].Split('+')[1].Trim(trim_chr));
                        ll = Convert.ToDouble(temp[1].Trim(trim_chr));
                    }
                    dr_SV[inx] = sv;
                    dr_UL[inx] = (sv + ul).ToString();
                    dr_LL[inx] = (sv - ll).ToString();
                }
                else
                {
                    string[] UL_LL = _limit.Split(split_chr);
                    if (_limit.Contains('>'))
                    {
                        dr_SV[inx] = _limit;
                        dr_UL[inx] = UL_LL[0].Trim(trim_chr);
                        dr_LL[inx] = UL_LL[1].Trim(trim_chr);
                    }
                    else
                    {
                        if (_limit.Contains('\u00B1'))
                        {
                            dr_SV[inx] = UL_LL[0];
                            dr_UL[inx] = (Convert.ToDouble(UL_LL[0].Trim(trim_chr)) + Convert.ToDouble(UL_LL[1].Trim(trim_chr))).ToString();
                            dr_LL[inx] = (Convert.ToDouble(UL_LL[0].Trim(trim_chr)) - Convert.ToDouble(UL_LL[1].Trim(trim_chr))).ToString();
                        }
                        else
                        {
                            dr_SV[inx] = _limit;
                            dr_LL[inx] = UL_LL[0].Trim(trim_chr);
                            dr_UL[inx] = UL_LL[1].Trim(trim_chr);
                        }
                    }
                }
                inx++;
            }
            tar_Spec_tbl.Rows.Add(dr_SV);
            tar_Spec_tbl.Rows.Add(dr_UL);
            tar_Spec_tbl.Rows.Add(dr_LL);
        }
        public bool Check_inSpec_IPQC_Data(DataTable src_tbl_Data, DataTable src_tbl_Spec)
        {
            bool _result = false;
            char[] trim_char = new char[] { '\r', '\n', ' ' };
            foreach (DataColumn dgv_col in src_tbl_Data.Columns)
            {
                string col_name = dgv_col.ColumnName;
                string UL = checkDBNull(src_tbl_Spec.Rows[1][col_name]).Trim(trim_char);
                string LL = checkDBNull(src_tbl_Spec.Rows[2][col_name]).Trim(trim_char);
                string SV = checkDBNull(src_tbl_Spec.Rows[0][col_name]).Trim(trim_char);
                foreach (DataRow dgv_row in src_tbl_Data.Rows)
                {
                    string act_val = checkDBNull(dgv_row[col_name]).Trim(trim_char);
                    if (act_val != "")
                    {
                        if (!check_in_limit(UL, LL, act_val, SV))
                        {
                            _result = false;
                            break;
                        }
                        else
                        {
                            _result = true;
                        }
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
    }
}
