using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using myExcel = Microsoft.Office.Interop.Excel;
using TDMK_SQL;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Data.SqlClient;
using OK2SHIP_Lib;

namespace OK2SHIP_Measurements
{
    public partial class FrmIPQC_Man : Form
    {
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public TDMK_OK2SHIP myCode = new TDMK_OK2SHIP();
        IPQC_LogFile proc_data = new IPQC_LogFile();
        public string app_path;
        public string origin_cell_val;
        public bool edit_en = false;
        public bool pro_en = false;
        public delegate void Data_process(string _filename, DataGridView tar_DGV, DataGridView DGV_Spec, DataGridView tar_DGV_plus, bool edit_mode);
        public delegate void SetText(Button tar_btn, string btn_str);
        public delegate void checkTextbox(TextBox src_txt);
        public Color curr_color;
        public SqlConnection sqlcon;
        public string sel_tbl;
        public DataTable sel_dt = new DataTable();
        public DataTable cur_sel_dt = new DataTable();
        public bool update_en = false;
        public AutoCompleteStringCollection LAB_Printing_IgnoredList = new AutoCompleteStringCollection() { "LPI_Open_B1", "LPI_Open_B2" };
        public AutoCompleteStringCollection Etching_DESData_list = new AutoCompleteStringCollection() { "DES_DATA_4", "DES_DATA_7", "DES_DATA_11", "DES_DATA_14", "DES_DATA_18", "DES_DATA_21", "DES_DATA_25", "DES_DATA_28" };
        public string Sel_Depart = "LAB";
        public bool system_ready = false;
        public int sel_spec_col = 0;
        public int start_ID = 0;
        public List<string> sel_ColName_lst = new List<string>();
        struct config_info
        {
            public string tar_app;
            public string tar_Dev;
            public config_info(string app, string Dev)
            {
                this.tar_app = app;
                this.tar_Dev = Dev;
            }

        }
        struct FAI_Spec
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
        public string Server_Account;
        public string Server_Pass;
        public string Dev;
        public string format_folder;
        public string Data_Location;
        public string Server_name;
        public bool EditMode;
        public string DB_name;
        public string f_ext;
        public string log_folder;

        //FAI_Spec[] testFAI = new FAI_Spec[1000];
        TDMK_OK2SHIP.FAI_Spec[] testFAI_spec = new TDMK_OK2SHIP.FAI_Spec[1000];
        public FrmIPQC_Man()
        {
            InitializeComponent();
        }
        //public FrmIPQC_Man(string _Server_Account, string _Server_Pass, string _Server_name, string _log_folder, SqlConnection _sqlcon)
        //{
        //    InitializeComponent();
        //    Server_Account = _Server_Account;
        //    Server_Pass = _Server_Pass;
        //    Server_name = _Server_name;
        //    log_folder = _log_folder;
        //    sqlcon = _sqlcon;
        //}
        public FrmIPQC_Man(string _log_folder, SqlConnection _sqlcon)
        {
            InitializeComponent();
            log_folder = _log_folder;
            sqlcon = _sqlcon;
        }

        private void DGV_DataView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            
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
                    FAI_tbl = mytbl.Select(tar_Col + " like '%FAI%'").CopyToDataTable();
                    fill_datatable_DGV(FAI_tbl, tar_DGV, DGV_Spec);
                }
                catch
                {

                }
                try
                {
                    non_FAI = mytbl.Select(tar_Col + " not like '%FAI%'").CopyToDataTable();
                    fill_datatable_DGV(non_FAI, DGV_plus, DGV_Spec);
                }
                catch
                {

                }
            }
            else
            {
                DGV_plus.Columns.Clear();
                fill_datatable_DGV(mytbl, DGV_plus, DGV_Spec);
                //fill_datatable_DGV(FAI_tbl, DGV_plus, DGV_Spec);
            }
        }
        public void fill_datatable_DGV(DataTable src_tbl, DataGridView tar_DGV, DataGridView tar_DGV_Spec)
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
            //if (tar_DGV_Spec.Columns.Count > 0)
            //{
            //    tbl_spec = DGV_To_Table(tar_DGV_Spec);
            //}
            tbl_spec = DGV_To_Table(tar_DGV_Spec);
            foreach (string tg in col_name)
            {
                if (!check_columns_existed(tbl_DGV_data, tg))
                {
                    tar_DGV.Columns.Add(tg, tg);
                }
                //if (!check_columns_existed(tbl_spec, tg))
                //{
                //    tbl_spec.Columns.Add(tg);
                //}
            }
            //if (tbl_spec.Rows.Count == 0)
            //{
            //    tbl_spec.Rows.Add();
            //    tbl_spec.Rows.Add();
            //    tbl_spec.Rows.Add();
            //}
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
                //string[] UL_val;
                //string[] LL_val;
                //string[] SetVal;
                string[] vitri_val;
                //double SV;
                //double UL;// = Convert.ToDouble(tbl_spec.Rows[1][tg]);
                //double LL;// = Convert.ToDouble(tbl_spec.Rows[2][tg]); ;
                double UL = 0;// = Convert .ToDouble (tbl_spec.Rows [1][tg]);
                double LL = 0;// Convert.ToDouble(tbl_spec.Rows[2][tg]);
                bool spec_check_en = false;
                if (check_columns_existed(tbl_spec, tg))
                {
                    UL = Convert.ToDouble(tbl_spec.Rows[1][tg]);
                    LL = Convert.ToDouble(tbl_spec.Rows[2][tg]);
                    spec_check_en = true;
                }
                item_val = temp.AsEnumerable().Select(r => r.Field<string>(vitri_act_val)).ToArray();
                vitri_val = temp.AsEnumerable().Select(r => r.Field<string>(vitri_col_name)).ToArray();
                //SetVal = temp.AsEnumerable().Select(r => r.Field<string>(vitri_Set_val)).Distinct().ToArray();
                //UL_val = temp.AsEnumerable().Select(r => r.Field<string>(vitri_UL_val)).Distinct().ToArray();
                //LL_val = temp.AsEnumerable().Select(r => r.Field<string>(vitri_LL_val)).Distinct().ToArray();
                //SV = Math.Abs(Convert.ToDouble(SetVal[0]));
                //UL = Math.Abs(Convert.ToDouble(SetVal[0]) + Convert.ToDouble(UL_val[0]));
                //LL = Math.Abs(Convert.ToDouble(SetVal[0]) + Convert.ToDouble(LL_val[0]));
                //tbl_spec.Rows[0][tg] = SetVal[0];
                //tbl_spec.Rows[1][tg] = UL.ToString();
                //tbl_spec.Rows[2][tg] = LL.ToString();
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
                    if (spec_check_en)
                    {
                        if ((Convert.ToDouble(c) <= UL) && (Convert.ToDouble(c) >= LL))
                        {
                            tar_DGV.Rows[j].Cells[tg].Style.BackColor = Color.White;
                        }
                        else
                        {
                            tar_DGV.Rows[j].Cells[tg].Style.BackColor = Color.Red;
                        }
                    }
                    else

                    {
                        tar_DGV.Rows[j].Cells[tg].Style.BackColor = Color.Yellow;
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
            //tar_DGV_Spec.DataSource = tbl_spec;
            //tar_DGV_Spec.Rows[0].HeaderCell.Value = "SV";
            //tar_DGV_Spec.Rows[1].HeaderCell.Value = "UL";
            //tar_DGV_Spec.Rows[2].HeaderCell.Value = "LL";
            //tar_DGV_Spec.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            //tar_DGV_Spec.AutoResizeRows();
            //tar_DGV_Spec.AutoResizeColumns();
            //Disable_Sort_DGV(tar_DGV_Spec);
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
            //if (tbl_spec.Rows.Count == 0)
            //{
            //    tbl_spec.Rows.Add();
            //    tbl_spec.Rows.Add();
            //    tbl_spec.Rows.Add();
            //}
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
                //double SV = Convert.ToDouble(temp_arr[3]);
                //double UL = Convert.ToDouble(temp_arr[4]);
                //double LL = Convert.ToDouble(temp_arr[5]);

                double UL = 0;// = Convert.ToDouble(temp_arr[4]);
                double LL = 0;// = Convert.ToDouble(temp_arr[5]);
                bool spec_check_en = false;
                if (!check_columns_existed(DGV_To_Table(sel_DGV), col_name[i]))
                {
                    sel_DGV.Columns.Add(col_name[i], col_name[i]);
                }
                if (check_columns_existed(tbl_spec, col_name[i]))
                {
                    UL = Convert.ToDouble(tbl_spec.Rows[1][col_name[i]]);
                    LL = Convert.ToDouble(tbl_spec.Rows[2][col_name[i]]);
                    spec_check_en = true;
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
                    double act_val = Convert.ToDouble(col_val[j]);
                    if (spec_check_en)
                    {
                        if ((act_val <= UL) && (act_val >= LL))
                        {
                            sel_DGV.Rows[row_inx + j].Cells[col_name[i]].Style.BackColor = Color.White;
                        }
                        else
                        {
                            sel_DGV.Rows[row_inx + j].Cells[col_name[i]].Style.BackColor = Color.Red;
                        }
                    }
                    else
                    {
                        sel_DGV.Rows[row_inx + j].Cells[col_name[i]].Style.BackColor = Color.Yellow;
                    }
                }
                //if (!check_columns_existed(tbl_spec, col_name[i]))
                //{
                //    tbl_spec.Columns.Add(col_name[i]);
                //    tbl_spec.Rows[0][col_name[i]] = temp_arr[3];
                //    tbl_spec.Rows[1][col_name[i]] = (SV + UL).ToString();
                //    tbl_spec.Rows[2][col_name[i]] = (SV + LL).ToString();
                //}
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
            //DGV_Spec.DataSource = tbl_spec;
            //if (DGV_Spec.Columns.Count > 0)
            //{
            //    DGV_Spec.Rows[0].HeaderCell.Value = "SV";
            //    DGV_Spec.Rows[1].HeaderCell.Value = "UL";
            //    DGV_Spec.Rows[2].HeaderCell.Value = "LL";
            //    DGV_Spec.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            //    DGV_Spec.AutoResizeRows();
            //    DGV_Spec.AutoResizeColumns();
            //}
            Disable_Sort_DGV(DGV_Spec);
            Disable_Sort_DGV(tar_DGV);
            Disable_Sort_DGV(tar_DGV_Plus);
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
            tar_file = Path.Combine(app_path, "Original", _filename);
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
            tbl_spec = DGV_To_Table(DGV_Spec);
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
                if (check_columns_existed(tbl_spec, src_col))
                {
                    UL = Convert.ToDouble(tbl_spec.Rows[1][src_col]);
                    LL = Convert.ToDouble(tbl_spec.Rows[2][src_col]);
                    spec_check_en = true;
                }

                for (int t = 0; t < sel_DGV.RowCount; t++)
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

        private void Form1_Load(object sender, EventArgs e)
        {
            //app_path = @"D:\Customer Projects\SEEV\TestAreas";
            app_path = Application.StartupPath;
            curr_color = this.BackColor;
            EditMode = false;
            RB_Etching_Process.Checked = true;
            RBNormal.Checked = true;
            RB_Lab.Checked = true;
            sel_tbl = "Etching_Process";
            system_ready = true;
        }
        private void Watchfolder_Created(object sender, FileSystemEventArgs e)
        {
            if (e.ChangeType == WatcherChangeTypes.Created)
            {
                try
                {
                    string itemcode_folder;
                    string f_info = e.FullPath;
                    string done_file;
                    string f_name = Path.GetFileName(f_info);
                    string save_time = Create_date_string(DateTime.Now);
                    if (pro_en)
                    {
                        itemcode_folder = Path.Combine(app_path, "Scan_Finished", txtItemCode.Text);
                        done_file = Path.Combine(itemcode_folder, "done_" + save_time + f_name);
                        switch (Dev)
                        {
                            case "Mitutoyo":
                                Invoke(new Data_process(Mitutoyo_Data_display), f_info, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode);
                                break;
                            case "Nikon":
                                Invoke(new Data_process(Nikon_Data_display), f_info, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode);
                                break;
                            case "Keyence":
                                Invoke(new Data_process(Keyence_Display_Data), f_info, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode);
                                break;
                        }
                        pro_en = false;
                        Invoke(new SetText(SetTextButton), btnRun, "Start");
                    }
                    else
                    {
                        itemcode_folder = Path.Combine(app_path, "Ignored_Scan", txtItemCode.Text);
                        done_file = Path.Combine(itemcode_folder, "Ignored_" + save_time + f_name);
                    }
                    if (!Directory.Exists(itemcode_folder))
                    {
                        Directory.CreateDirectory(itemcode_folder);
                    }
                    File.Copy(f_info, done_file, true);
                    File.Delete(f_info);
                    MessageBox.Show(new Form { TopMost = true },"Hoàn thành lấy dữ liệu", "Thông tin");
                }
                catch (Exception ex)
                {
                    Error_log(app_path, ex.Message);
                }
            }
        }
        private void btnRun_Click(object sender, EventArgs e)
        {

            /****************************************** Main Program **************************************************************/
            Main_Program();
            /*********************************************************************************************************************/

            /****************************************** Test Areas **************************************************************/
            //sel_tbl = "Roughness";
            //sel_tbl = "AU_NI";
            //Load_Test_Data2(@"D:\Customer Projects\SEEV\TestAreas\TestData", txtItemCode.Text, txtLotNo.Text, ".xlsm", DGV_DataView);


            //Load_Test_IPQC_Data(@"D:\Customer Projects\SEEV\TestAreas\TestData", txtItemCode.Text, txtLotNo.Text, ".xlsm", DGV_DataView);

            /*********************************************************************************************************************/

        }
        public void Main_Program()
        {
            DataGridView sel_DGV;
            if (EditMode)
            {
                sel_DGV = DGV_Data_Plus;
            }
            else
            {
                sel_DGV = DGV_DataView;
            }
            if (btnRun.Text == "Start")
            {
                if (init_data3())
                {
                    btnRun.Text = "Stop";

                    DataTable dt_result = TDMK_Code.Datatable_Filter(sqlcon, sel_tbl, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));

                    if (dt_result.Rows.Count == 0)
                    {
                        foreach (DataGridViewColumn t in DGV_SpecView.Columns)
                        {
                            if(!myCode.check_columns_existed(myCode.DGV_To_Table(sel_DGV),t.Name))
                            {
                                sel_DGV.Columns.Add(t.Name, t.HeaderText);
                            }                            
                        }
                        if(sel_DGV.Rows.Count!= Convert.ToInt32(numQty.Value))
                        {
                            for (int i = 0; i < Convert.ToInt32(numQty.Value); i++)
                            {
                                sel_DGV.Rows.Add();
                                sel_DGV.Rows[i].HeaderCell.Value = (i + 1).ToString();
                            }
                        }
                        foreach (DataGridViewColumn t in sel_DGV.Columns)
                        {
                            if(t.Name.Contains("Rate"))
                            {
                                foreach(DataGridViewRow r in sel_DGV.Rows)
                                {
                                    r.Cells[t.Index].ReadOnly = true;
                                }
                            }
                        }
                        sel_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                        update_en = false;
                    }
                    else
                    {
                        if(sel_DGV.Rows.Count==0)
                        {
                            string[] col_name_lst = new string[DGV_SpecView.Columns.Count];
                            string[] col_header_lst = new string[DGV_SpecView.Columns.Count];
                            List<string> cur_col_arr = cur_sel_dt.AsEnumerable().Select(r => r.Field<string>("Col_Name")).ToList();
                            int inx = 0;
                            foreach (DataGridViewColumn t in DGV_SpecView.Columns)
                            {
                                if (TDMK_Code.check_exist_list_index2(t.Name, cur_col_arr) > -1)
                                {
                                    col_name_lst[inx] = t.Name;
                                    col_header_lst[inx] = t.HeaderText;
                                    inx++;
                                }
                            }
                            Array.Resize(ref col_name_lst, inx);
                            Array.Resize(ref col_header_lst, inx);
                            sel_DGV.DataSource = dt_result.AsDataView().ToTable(false, col_name_lst);
                            inx = 0;
                            foreach (DataGridViewColumn t in sel_DGV.Columns)
                            {
                                t.HeaderText = col_header_lst[inx];
                                inx++;
                            }
                            if ((sel_tbl == "UV_Process") && (EditMode))
                            {
                                sel_DGV.Columns.Add("Inner_Rate", "Inner_Rate");
                                sel_DGV.Columns.Add("Outer_Rate", "Outer_Rate");
                            }

                            start_ID = Convert.ToInt32(dt_result.Rows[0]["ID"]);
                            for (int r = 0; r < dt_result.Rows.Count; r++)
                            {

                                int cur_ID = Convert.ToInt32(dt_result.Rows[r]["ID"]);
                                //sel_DGV.Rows[r].HeaderCell.Value = dt_result.Rows[r]["ID"].ToString();
                                sel_DGV.Rows[r].HeaderCell.Value = (cur_ID - start_ID + 1).ToString();
                            }
                            sel_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                            foreach (DataGridViewColumn c in sel_DGV.Columns)
                            {
                                if (check_columns_existed(DGV_To_Table(DGV_SpecView), c.Name))
                                {
                                    string UL = myCode.checkDBNull(DGV_SpecView.Rows[1].Cells[c.Name].Value);
                                    string LL = myCode.checkDBNull(DGV_SpecView.Rows[2].Cells[c.Name].Value);
                                    string SV = myCode.checkDBNull(DGV_SpecView.Rows[0].Cells[c.Name].Value);
                                    foreach (DataGridViewRow r in sel_DGV.Rows)
                                    {
                                        string act_val = myCode.checkDBNull(r.Cells[c.Index].Value);
                                        if (act_val != "")
                                        {
                                            //r.Cells[c.Index].ReadOnly = true;
                                            c.ReadOnly = true;
                                        }
                                        //r.Cells[c.Index].Style.BackColor = myCode.check_in_limit2(UL, LL, act_val);
                                        r.Cells[c.Index].Style.BackColor = myCode.check_in_limit_Color(UL, LL, act_val, SV);
                                    }
                                }

                            }
                            update_en = true;
                        }
                    }
                    Disable_Sort_DGV(sel_DGV);
                    btnSave.Enabled = false;
                    sel_DGV.EditMode = DataGridViewEditMode.EditOnEnter;
                    timer1.Enabled = true;
                    edit_en = true;
                    pro_en = true;
                }
            }
            else
            {
                DGV_Data_Plus.EditMode = DataGridViewEditMode.EditProgrammatically;
                DGV_DataView.EditMode = DataGridViewEditMode.EditProgrammatically;
                btnRun.Text = "Start";
                pro_en = false;
                timer1.Enabled = false;
                btnRun.BackColor = curr_color;
                edit_en = false;
                btnSave.Enabled = true;
            }
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

        private void btnReset_Click(object sender, EventArgs e)
        {
            DGV_DataView.Columns.Clear();
            DGV_Data_Plus.Columns.Clear();
            DGV_SpecView.Columns.Clear();
            DGV_SpecView.DataSource = null;
            DGV_DataView.DataSource = null;
            DGV_DataView.Columns.Clear();
            DGV_Data_Plus.DataSource = null;
            chklstPart.Items.Clear();
            pro_en = false;
            Invoke(new SetText(SetTextButton), btnRun, "Start");
        }
        //private void SetText1()
        //{
        //    txtItemCode.Text = "Test";
        //}

        //private void SetText2()
        //{
        //    txtLotNo.Invoke(new Action(() => txtLotNo.Text = "Test"));
        //}
        public void Disable_Sort_DGV(DataGridView sel_DGV)
        {
            foreach (DataGridViewColumn column in sel_DGV.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
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
        public void SetTextButton(Button src_btn, string txt_str)
        {
            src_btn.Text = txt_str;
            timer1.Enabled = false;
            src_btn.BackColor = curr_color;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (myCode.check_data_enough(DGV_DataView, Convert.ToInt32(numQty.Value)))
            {
                bool en_save = false;
                if (myCode.check_data_OK(DGV_DataView, DGV_SpecView))
                {
                    en_save = true;
                }
                else
                {
                    if ((MessageBox.Show("Data out of spec. Do you want to save it ?", "Warning", MessageBoxButtons.YesNo)) == DialogResult.Yes)
                    {
                        en_save = true;
                    }
                    else
                    {
                        en_save = false;
                    }
                }
                if (en_save)
                {
                    //Save_data();
                    if (TDMK_OK2SHIP.admin_mode)
                    {
                        Save_Testdata();
                        TDMK_OK2SHIP.admin_mode = false;
                    }
                    else
                    {
                        MessageBox.Show("Please, Login to save data");
                        FrmLogin frmLogin = new FrmLogin();
                        frmLogin.ShowDialog();
                    }
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },"Not enough data", "Warning");
            }
        }
        public bool init_data()
        {
            bool _result = false;
            if (chklstPart.CheckedItems.Count > 0)
            {
                if (DGV_SpecView.Columns.Count == 0)
                {
                    string[] header_arr = new string[] { "SetVal", "UL", "LL" };
                    foreach (int check_inx in chklstPart.CheckedIndices)
                    {
                        string col_name = cur_sel_dt.Rows[check_inx]["Col_Name"].ToString();
                        DGV_SpecView.Columns.Add(col_name, chklstPart.Items[check_inx].ToString());
                    }
                    for (int i = 0; i < 3; i++)
                    {
                        DGV_SpecView.Rows.Add();
                        DGV_SpecView.Rows[i].HeaderCell.Value = header_arr[i];
                    }
                    DGV_SpecView.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                    Disable_Sort_DGV(DGV_SpecView);
                }
                _result = true;
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },"Please, select items check!", "Warning!");
                _result = false;
            }
            return _result;
        }
        public bool init_data2()
        {
            bool _result = false;
            if (chklstPart.CheckedItems.Count > 0)
            {
                if (sel_dt.Rows.Count > 0)
                {
                    string[] col_arr = new string[chklstPart.CheckedIndices.Count];
                    string[] col_header = new string[chklstPart.CheckedIndices.Count];
                    int col_inx = 0;
                    foreach (int check_inx in chklstPart.CheckedIndices)
                    {
                        col_arr[col_inx] = cur_sel_dt.Rows[check_inx]["Col_Name"].ToString();
                        col_header[col_inx] = chklstPart.Items[check_inx].ToString();
                        col_inx++;
                    }
                    myCode.IPQC_Spec_Process_Man(cur_sel_dt, DGV_SpecView, col_arr, col_header);
                }
                else
                {
                    DGV_SpecView.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
                    if (DGV_SpecView.Columns.Count == 0)
                    {
                        string[] header_arr = new string[] { "SetVal", "UL", "LL" };
                        foreach (int check_inx in chklstPart.CheckedIndices)
                        {
                            string col_name = cur_sel_dt.Rows[check_inx]["Col_Name"].ToString();
                            DGV_SpecView.Columns.Add(col_name, chklstPart.Items[check_inx].ToString());
                        }
                        for (int i = 0; i < 3; i++)
                        {
                            DGV_SpecView.Rows.Add();
                            DGV_SpecView.Rows[i].HeaderCell.Value = header_arr[i];
                        }
                        DGV_SpecView.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                        Disable_Sort_DGV(DGV_SpecView);
                    }
                }
                _result = true;
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },"Please, select items check!", "Warning!");
                _result = false;
            }
            return _result;
        }
        public bool init_data3()
        {
            bool _result = false;
            if (DGV_SpecView.Columns.Count>0)
            {
                if ((sel_dt.Rows.Count > 0)&& (!EditMode))
                {
                    myCode.IPQC_Spec_Process_Man3(cur_sel_dt, DGV_SpecView);
                }
                else
                {
                    DGV_SpecView.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
                    string[] header_arr = new string[] { "SetVal", "UL", "LL" };
                    if (DGV_SpecView.Columns.Count == 0)
                    {
                        foreach (int check_inx in chklstPart.CheckedIndices)
                        {
                            string col_name = cur_sel_dt.Rows[check_inx]["Col_Name"].ToString();
                            DGV_SpecView.Columns.Add(col_name, chklstPart.Items[check_inx].ToString());
                        }

                    }
                    if (DGV_SpecView.RowCount ==0)
                    {
                        for (int i = 0; i < 3; i++)
                        {
                            DGV_SpecView.Rows.Add();
                            DGV_SpecView.Rows[i].HeaderCell.Value = header_arr[i];
                        }
                    }

                    DGV_SpecView.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                    Disable_Sort_DGV(DGV_SpecView);
                }
            start_label: DataTable sequence_tbl = TDMK_Code.Datatable_Filter(sqlcon, "Sequence", TDMK_Code.filter_str(new string[] { "ItemCode", "Process", "Depart" }, new string[] { txtItemCode.Text, sel_tbl, Sel_Depart }));
                if(sequence_tbl.Rows.Count ==0)
                {
                    for (int i = 0; i < DGV_SpecView.Columns.Count; i++)
                    {
                        string ID = (TDMK_Code.SQL_MAX("Sequence", "ID", sqlcon) + 1).ToString();
                        TDMK_Code.insert_val_arr2("Sequence", sqlcon, new string[] { "ID", "ItemCode", "Col_Name", "Item_Name", "Process","Depart" }, new string[] { ID, txtItemCode.Text, DGV_SpecView.Columns[i].Name, DGV_SpecView.Columns[i].HeaderText, sel_tbl, Sel_Depart });
                    }
                }
                else
                {
                    if(DGV_SpecView.Columns.Count != sequence_tbl.Rows.Count)
                    {
                        if (MessageBox.Show("Sequence is created! Do you want to update again", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            TDMK_Code.Delelte_FilteredItem_arr("Sequence", sqlcon, "ItemCode = '" + txtItemCode.Text + "'");
                            goto start_label;
                        }
                    }
                }
                _result = true;
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },"Please, select items check!", "Warning!");
                _result = false;
            }
            return _result;
        }
        public void save_spec()
        {
            string[] items = new string[10];
            string[] item_vals = new string[10];
            items[0] = "ID";
            items[1] = "ItemCode";
            items[2] = "LotNo";
            items[3] = "NormDim";
            items[4] = "TolMax";
            items[5] = "TolMin";
            items[6] = "Instrument";
            items[7] = "FAI_No";
            items[8] = "Distribution";
            item_vals[1] = txtItemCode.Text;
            item_vals[2] = txtLotNo.Text;
            item_vals[6] = txtMachine.Text;
            if (DGV_SpecView.Rows.Count > 0)
            {
                for (int i = 0; i < DGV_SpecView.ColumnCount; i++)
                {
                    AutoCompleteStringCollection FAI_No_list = new AutoCompleteStringCollection();
                    item_vals[3] = "";
                    item_vals[4] = "";
                    item_vals[5] = "";
                    item_vals[8] = "";
                    item_vals[7] = DGV_SpecView.Columns[i].Name;
                    FAI_No_list = TDMK_Code.Load_Item_Filter_str(sqlcon, "FAI_Spec", "NormDim", TDMK_Code.filter_str(items, item_vals));
                    if (FAI_No_list.Count == 0)
                    {
                        item_vals[0] = (TDMK_Code.SQL_MAX("FAI_Spec", "ID", sqlcon) + 1).ToString();
                        item_vals[3] = DGV_SpecView.Rows[0].Cells[i].Value.ToString();
                        item_vals[4] = DGV_SpecView.Rows[1].Cells[i].Value.ToString();
                        item_vals[5] = DGV_SpecView.Rows[2].Cells[i].Value.ToString();
                        item_vals[8] = DGV_SpecView.Rows[3].Cells[i].Value.ToString();
                        TDMK_Code.insert_val_arr("FAI_Spec", sqlcon, items, item_vals);
                    }
                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (btnRun.BackColor == curr_color)
            {
                btnRun.BackColor = Color.YellowGreen;
            }
            else
            {
                btnRun.BackColor = curr_color;
            }
        }
        public void check_txt(TextBox tar_txt)
        {
            if (tar_txt.Text != "")
            {
                tar_txt.BackColor = Color.White;
            }
            else
            {
                tar_txt.BackColor = Color.Red;
            }
        }
        private void txtItemCode_TextChanged(object sender, EventArgs e)
        {
            Invoke(new checkTextbox(check_txt), txtItemCode);
        }

        private void txtLotNo_TextChanged(object sender, EventArgs e)
        {
            Invoke(new checkTextbox(check_txt), txtLotNo);
        }

        private void txtOperator_TextChanged(object sender, EventArgs e)
        {
            Invoke(new checkTextbox(check_txt), txtOperator);
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            DataGridView sel_DGV;
            if(EditMode)
            {
                sel_DGV = DGV_Data_Plus;
            }
            else
            {
                sel_DGV = DGV_DataView;
            }
            if (sel_DGV.RowCount > 0)
            {
                myCode.export_data = new Thread(export_to_Excel);
                myCode.export_data.Start();
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },"No Data to export", "Warning");
            }


        }
        //public bool Load_Spec_Manual(string format_loc, string tar_ItemCode, string tar_LotNo,string file_extension, DataGridView tar_DGV_Spec, string tar_device)
        //{
        //    bool _result = false;
        //    string tar_format = tar_ItemCode + tar_LotNo;
        //    string tar_format_file = Path.Combine(format_loc, tar_format + file_extension);
        //    string[] spec_item_str = new string[13];
        //    AutoCompleteStringCollection FAI_No_list = new AutoCompleteStringCollection();
        //    FAI_No_list = TDMK_Code.Load_Item_Filter_str(sqlcon, "FAI_Spec", "FAI_No", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Instrument" }, new string[] { txtItemCode.Text, txtLotNo.Text, txtMachine.Text  }));
        //    if (FAI_No_list.Count == 0)
        //    {
        //        spec_item_str[0] = "SetVal";
        //        spec_item_str[1] = "UL";
        //        spec_item_str[2] = "LL";
        //        spec_item_str[3] = "CheckSide";
        //        spec_item_str[4] = "STDEV";
        //        spec_item_str[5] = "Mean";
        //        spec_item_str[6] = "MAX";
        //        spec_item_str[7] = "MIN";
        //        spec_item_str[8] = "Cp";
        //        spec_item_str[9] = "Cpkl";
        //        spec_item_str[10] = "Cpku";
        //        spec_item_str[11] = "Cpk";
        //        spec_item_str[12] = "Cpkm";
        //        if (File.Exists(tar_format_file))
        //        {
        //            myExcel.Workbook tar_wkbook = TDMK_Code.open_excel_file(tar_format_file, "", "");
        //            int wrksheet_num = tar_wkbook.Worksheets.Count;
        //            int spec_col_inx = 0;
        //            string src_dev = txtMachine.Text;
        //            foreach (myExcel.Worksheet tg in tar_wkbook.Worksheets)
        //            {
        //                if (tg.Name.Contains("FAI"))
        //                {
        //                    myExcel.Range sel_rgn = tg.Range["C17"];
        //                    myExcel.Range dev_rgn = tg.Range["C21"];
        //                    int sel_inx = 0;
        //                    while (checkDBNull(sel_rgn.Offset[0, sel_inx].Value) != "")
        //                    {
        //                        string cur_dev = dev_rgn.Offset[0, sel_inx].Value;
        //                        if (cur_dev.Contains(src_dev))
        //                        {
        //                            string t_checkside = checkDBNull(sel_rgn.Offset[-1, sel_inx].Value);
        //                            string t_FAIName = checkDBNull(sel_rgn.Offset[0, sel_inx].Value).Replace(" ", "");
        //                            string t_FAI_Setval = checkDBNull(sel_rgn.Offset[1, sel_inx].Value);
        //                            string t_FAI_UL = checkDBNull(sel_rgn.Offset[6, sel_inx].Value);
        //                            string t_FAI_LL = checkDBNull(sel_rgn.Offset[7, sel_inx].Value);
        //                            testFAI_spec[spec_col_inx] = new TDMK_OK2SHIP.FAI_Spec(t_FAIName, t_checkside, t_FAI_Setval, t_FAI_UL, t_FAI_LL);
        //                            if (!check_columns_existed(DGV_To_Table(tar_DGV_Spec), t_FAIName))
        //                            {
        //                                tar_DGV_Spec.Columns.Add(t_FAIName, t_FAIName);

        //                                if (tar_DGV_Spec.Rows.Count == 0)
        //                                {
        //                                    for (int i = 0; i < 4; i++)
        //                                    {
        //                                        tar_DGV_Spec.Rows.Add();
        //                                        tar_DGV_Spec.Rows[i].HeaderCell.Value = spec_item_str[i];
        //                                    }
        //                                }
        //                                tar_DGV_Spec.Rows[0].Cells[t_FAIName].Value = t_FAI_Setval; // Set val at Row =0
        //                                tar_DGV_Spec.Rows[1].Cells[t_FAIName].Value = t_FAI_UL; // UL at Row =1
        //                                tar_DGV_Spec.Rows[2].Cells[t_FAIName].Value = t_FAI_LL; // LL at Row =2
        //                                tar_DGV_Spec.Rows[3].Cells[t_FAIName].Value = t_checkside; // Check Side at Row =3
        //                                spec_col_inx++;
        //                            }
        //                        }
        //                        sel_inx++;
        //                    }
        //                }
        //            }
        //            Array.Resize<TDMK_OK2SHIP.FAI_Spec>(ref testFAI_spec, spec_col_inx);
        //            tar_DGV_Spec.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        //            Disable_Sort_DGV(tar_DGV_Spec);

        //            tar_wkbook.Close();
        //            if (DGV_SpecView.Rows.Count >0)
        //            {
        //                save_spec();
        //                _result = true;
        //            }
        //            else
        //            {
        //                _result = false;
        //                MessageBox.Show("Không có thông tin cài đặt cho thiết bị: " + txtMachine .Text , "Thông báo");
        //            }

        //        }
        //        else
        //        {
        //            MessageBox.Show("File format khong tim thay");
        //            _result = false;
        //        }
        //    }
        //    else
        //    {
        //        string[] items = new string[4];
        //        string[] items_val = new string[4];
        //        items[0] = "ItemCode";
        //        items[1] = "LotNo";
        //        items[2] = "FAI_No";
        //        items[3] = "Instrument";
        //        items_val[0] = txtItemCode.Text;
        //        items_val[1] = txtLotNo.Text;
        //        items_val[3] = txtMachine.Text;
        //        foreach (string t in FAI_No_list)
        //        {
        //            items_val[2] = t;
        //            DataTable cur_dt = new DataTable();
        //            cur_dt = DGV_To_Table(DGV_SpecView);
        //            int dgv_row = cur_dt.Rows.Count;
        //            if (!check_columns_existed(cur_dt, t))
        //            {
        //                cur_dt.Columns.Add(t);
        //            }
        //            if (dgv_row == 0)
        //            {
        //                for (int i = 0; i < 4; i++)
        //                {
        //                    cur_dt.Rows.Add();
        //                }
        //            }
        //            DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Spec", TDMK_Code.filter_str(items, items_val));
        //            cur_dt.Rows[0][t] = spec_dt.Rows[0]["NormDim"];
        //            cur_dt.Rows[1][t] = spec_dt.Rows[0]["TolMax"];
        //            cur_dt.Rows[2][t] = spec_dt.Rows[0]["TolMin"];
        //            cur_dt.Rows[3][t] = spec_dt.Rows[0]["Distribution"];
        //            DGV_SpecView.DataSource = cur_dt;
        //        }
        //        if (DGV_SpecView.Rows.Count > 0)
        //        {
        //            DGV_SpecView.Rows[0].HeaderCell.Value = "NormDim";
        //            DGV_SpecView.Rows[1].HeaderCell.Value = "USL";
        //            DGV_SpecView.Rows[2].HeaderCell.Value = "LSL";
        //            DGV_SpecView.Rows[3].HeaderCell.Value = "Distribution";
        //            DGV_SpecView.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        //            _result = true;
        //        }
        //    }

        //    return _result;
        //}
        //public bool Load_Spec(string format_loc, string tar_ItemCode, string tar_LotNo, string file_extension, DataGridView tar_DGV_Spec)
        //{
        //    bool _result = false;
        //    string tar_format = tar_ItemCode + tar_LotNo;
        //    string tar_format_file = Path.Combine(format_loc, tar_format + file_extension);
        //    string[] spec_item_str = new string[13];
        //    spec_item_str[0] = "SetVal";
        //    spec_item_str[1] = "UL";
        //    spec_item_str[2] = "LL";
        //    spec_item_str[3] = "CheckSide";
        //    spec_item_str[4] = "STDEV";
        //    spec_item_str[5] = "Mean";
        //    spec_item_str[6] = "MAX";
        //    spec_item_str[7] = "MIN";
        //    spec_item_str[8] = "Cp";
        //    spec_item_str[9] = "Cpkl";
        //    spec_item_str[10] = "Cpku";
        //    spec_item_str[11] = "Cpk";
        //    spec_item_str[12] = "Cpkm";
        //    if (File.Exists(tar_format_file))
        //    {
        //        myExcel.Workbook tar_wkbook = TDMK_Code.open_excel_file(tar_format_file, "", "");
        //        int wrksheet_num = tar_wkbook.Worksheets.Count;
        //        int spec_col_inx = 0;
        //        foreach (myExcel.Worksheet tg in tar_wkbook.Worksheets)
        //        {
        //            if (tg.Name.Contains("FAI"))
        //            {
        //                myExcel.Range sel_rgn = tg.Range["C17"];
        //                int sel_inx = 0;
        //                while (checkDBNull(sel_rgn.Offset[0, sel_inx].Value) != "")
        //                {
        //                    string t_checkside = checkDBNull(sel_rgn.Offset[-1, sel_inx].Value);
        //                    string t_FAIName = checkDBNull(sel_rgn.Offset[0, sel_inx].Value).Replace(" ", "");
        //                    string t_FAI_Setval = checkDBNull(sel_rgn.Offset[1, sel_inx].Value);
        //                    string t_FAI_UL = checkDBNull(sel_rgn.Offset[6, sel_inx].Value);
        //                    string t_FAI_LL = checkDBNull(sel_rgn.Offset[7, sel_inx].Value);
        //                    testFAI_spec[spec_col_inx] = new TDMK_OK2SHIP.FAI_Spec(t_FAIName, t_checkside, t_FAI_Setval, t_FAI_UL, t_FAI_LL);
        //                    if (!check_columns_existed(DGV_To_Table(tar_DGV_Spec), t_FAIName))
        //                    {
        //                        tar_DGV_Spec.Columns.Add(t_FAIName, t_FAIName);

        //                        if (tar_DGV_Spec.Rows.Count == 0)
        //                        {
        //                            for (int i = 0; i < 4; i++)
        //                            {
        //                                tar_DGV_Spec.Rows.Add();
        //                                tar_DGV_Spec.Rows[i].HeaderCell.Value = spec_item_str[i];
        //                            }
        //                        }
        //                        tar_DGV_Spec.Rows[0].Cells[t_FAIName].Value = t_FAI_Setval; // Set val at Row =0
        //                        tar_DGV_Spec.Rows[1].Cells[t_FAIName].Value = t_FAI_UL; // UL at Row =1
        //                        tar_DGV_Spec.Rows[2].Cells[t_FAIName].Value = t_FAI_LL; // LL at Row =2
        //                        tar_DGV_Spec.Rows[3].Cells[t_FAIName].Value = t_checkside; // Check Side at Row =3
        //                        spec_col_inx++;
        //                    }
        //                    sel_inx++;
        //                }
        //            }
        //        }
        //        Array.Resize<TDMK_OK2SHIP.FAI_Spec>(ref testFAI_spec, spec_col_inx);
        //        tar_DGV_Spec.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        //        Disable_Sort_DGV(tar_DGV_Spec);
        //        _result = true;
        //        tar_wkbook.Close();               
        //    }
        //    else
        //    {
        //        MessageBox.Show("File format khong tim thay");
        //        _result = false;
        //    }
        //    return _result;
        //}
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
                        myExcel.Range sel_rgn = tg.Range["C17"];
                        myExcel.Range data_rgn = tg.Range["C259"];
                        int sel_inx = 0;
                        while (checkDBNull(sel_rgn.Offset[0, sel_inx].Value) != "")
                        {
                            string t_FAIName = checkDBNull(sel_rgn.Offset[0, sel_inx].Value).Replace(" ", "");
                            if (!check_columns_existed(DGV_To_Table(tar_DGV_Data), t_FAIName))
                            {
                                tar_DGV_Data.Columns.Add(t_FAIName, t_FAIName);
                                for (int i = 0; i < 32; i++)
                                {
                                    if (tar_DGV_Data.RowCount < 32)
                                    {
                                        tar_DGV_Data.Rows.Add();
                                        tar_DGV_Data.Rows[i].HeaderCell.Value = (i + 1).ToString();
                                    }
                                    tar_DGV_Data.Rows[i].Cells[t_FAIName].Value = data_rgn.Offset[i, sel_inx].Value;
                                }
                            }
                            sel_inx++;
                        }
                    }
                }
                tar_DGV_Data.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                Disable_Sort_DGV(tar_DGV_Data);
                txtOperator.Text = spec_col_inx.ToString();
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
        public void export_to_Excel()
        {
            //Export_Data_Process(DGV_DataView, txtItemCode.Text, txtLotNo.Text);
            DataGridView sel_DGV;
            if (EditMode)
            {
                sel_DGV = DGV_Data_Plus;
            }
            else
            {
                sel_DGV = DGV_DataView;
            }
            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            xlsApp.Visible = true;
            myExcel.Workbook export_wrkbook = xlsApp.Workbooks.Add();
            TDMK_Code.Export_DGV_Excel3(sel_DGV, export_wrkbook,true);
            //Export_To_FAI(txtItemCode.Text, txtLotNo.Text, export_wrkbook);
        }
        public void Export_To_FAI(string _ItemCode, string _LotNo, myExcel.Workbook src_data_wrkbook)
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
        public void Export_Data_Process(DataGridView tar_DGV, string tar_ItemCode, string tar_LotNo)
        {
            myExcel.Workbook tar_wrkbook;
            myExcel.Worksheet tar_wrksheet;
            myExcel.Range FAI_No_rgn;
            myExcel.Range FAI_data_rgn;
            string format_path = Path.Combine(app_path, "Format");
            string temp_path = Path.Combine(app_path, "Temp_file");
            tar_wrkbook = TDMK_Code.open_excel_file(format_path, tar_ItemCode + "_" + tar_LotNo + ".xlsm", "");
            tar_wrksheet = tar_wrkbook.Sheets[1];
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
        }
        public void Save_data()
        {
            bool data_rec_en = true;
            int col_num = DGV_SpecView.Columns.Count + 6;
            string[] items = new string[col_num];
            string[] item_vals = new string[col_num];
            items[0] = "ID";
            items[1] = "ItemCode";
            items[2] = "LotNo";
            items[3] = "Operator";
            items[4] = "Machine";
            items[5] = "MDate";
            for (int i = 0; i < col_num - 6; i++)
            {
                string cur_col = DGV_SpecView.Columns[i].Name;
                if (cur_col.Contains(" "))
                {
                    items[6 + i] = "[" + cur_col + "]";
                }
                else
                {
                    items[6 + i] = cur_col;
                }
            }
            item_vals[1] = txtItemCode.Text;//"ItemCode";
            item_vals[2] = txtLotNo.Text;//"LotNo";
            item_vals[3] = txtOperator.Text;//"Operator";
            item_vals[4] = txtMachine.Text;//"Machine";
            item_vals[5] = DateTime.Now.ToString();//"MDate";
            foreach (Control c in GBInfo.Controls)
            {
                if ((c.Text == "") && (c.GetType().ToString().Contains("TextBox")))
                {
                    c.BackColor = Color.Red;
                    data_rec_en = false;
                }
            }
            if (data_rec_en)
            {
                if (DGV_DataView.Rows.Count > 0)
                {
                    for (int r_inx = 0; r_inx < DGV_DataView.RowCount; r_inx++)
                    {
                        for (int c_inx = 0; c_inx < DGV_DataView.ColumnCount; c_inx++)
                        {
                            item_vals[6 + c_inx] = checkDBNull(DGV_DataView.Rows[r_inx].Cells[c_inx].Value);

                        }
                        if (!update_en)
                        {
                            item_vals[0] = (TDMK_Code.SQL_MAX(sel_tbl, "ID", sqlcon) + 1).ToString();//"ID";
                            TDMK_Code.insert_val_arr(sel_tbl, sqlcon, items, item_vals);
                        }
                        else
                        {
                            int col_count = DGV_DataView.ColumnCount;
                            string[] update_item = new string[col_count];
                            string[] update_item_val = new string[col_count];
                            Array.Copy(items, 6, update_item, 0, col_count);
                            Array.Copy(item_vals, 6, update_item_val, 0, col_count);
                            int ID = Convert.ToInt32(DGV_DataView.Rows[r_inx].HeaderCell.Value);
                            TDMK_Code.updatebyID_val_arr(sel_tbl, sqlcon, ID, update_item, update_item_val);
                        }

                    }
                    DGV_DataView.Columns.Clear();
                    MessageBox.Show(new Form { TopMost = true },"Ghi dữ liệu thành công", "Thông báo");
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true },"Không có dữ liệu!", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },"Nhập dữ liệu thông tin!", "Chú ý");
            }
        }
        public void Save_Testdata()
        {
            bool data_rec_en = true;
            int col_num = DGV_DataView.Columns.Count + 6;
            string[] items = new string[col_num];
            string[] item_vals = new string[col_num];
            char[] trim_char = new char[] { '\r', '\n', ' ' };
            items[0] = "ID";
            items[1] = "ItemCode";
            items[2] = "LotNo";
            items[3] = "Operator";
            items[4] = "Machine";
            items[5] = "MDate";
            for (int i = 0; i < col_num - 6; i++)
            {
                string cur_col = DGV_DataView.Columns[i].Name;
                if (cur_col.Contains(" "))
                {
                    items[6 + i] = "[" + cur_col + "]";
                }
                else
                {
                    items[6 + i] = cur_col;
                }
            }
            item_vals[1] = txtItemCode.Text;//"ItemCode";
            item_vals[2] = txtLotNo.Text;//"LotNo";
            item_vals[3] = txtOperator.Text;//"Operator";
            item_vals[4] = txtMachine.Text;//"Machine";
            item_vals[5] = DateTime.Now.ToString();//"MDate";
            foreach (Control c in GBInfo.Controls)
            {
                if ((c.Text == "") && (c.GetType().ToString().Contains("TextBox")))
                {
                    c.BackColor = Color.Red;
                    data_rec_en = false;
                }
            }
            if (data_rec_en)
            {
                if (DGV_DataView.Rows.Count > 0)
                {
                    for (int r_inx = 0; r_inx < DGV_DataView.RowCount; r_inx++)
                    {
                        for (int c_inx = 0; c_inx < DGV_DataView.ColumnCount; c_inx++)
                        {
                            item_vals[6 + c_inx] = checkDBNull(DGV_DataView.Rows[r_inx].Cells[c_inx].Value).Trim(trim_char);

                        }
                        if (!update_en)
                        {
                            item_vals[0] = (TDMK_Code.SQL_MAX(sel_tbl, "ID", sqlcon) + 1).ToString();//"ID";
                            TDMK_Code.insert_val_arr(sel_tbl, sqlcon, items, item_vals);
                        }
                        else
                        {
                            int col_count = DGV_DataView.ColumnCount;
                            string[] update_item = new string[col_count];
                            string[] update_item_val = new string[col_count];
                            Array.Copy(items, 6, update_item, 0, col_count);
                            Array.Copy(item_vals, 6, update_item_val, 0, col_count);
                            int ID = start_ID-1+ Convert.ToInt32(DGV_DataView.Rows[r_inx].HeaderCell.Value);
                            TDMK_Code.updatebyID_val_arr(sel_tbl, sqlcon, ID, update_item, update_item_val);
                        }

                    }
                    DGV_DataView.Columns.Clear();
                    MessageBox.Show(new Form { TopMost = true },"Ghi dữ liệu thành công", "Thông báo");
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true },"Không có dữ liệu!", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },"Nhập dữ liệu thông tin!", "Chú ý");
            }
        }

        private void RBNormal_CheckedChanged(object sender, EventArgs e)
        {
            EditMode = false;
            if ((RBNormal.Checked)&& system_ready)
            {
                DGV_SpecView.EditMode = DataGridViewEditMode.EditProgrammatically;
                GB_Additional.Enabled = false;
                btnLoadSpec.PerformClick();
            }
        }

        private void RBEdit_CheckedChanged(object sender, EventArgs e)
        {
            EditMode = true;
            if (RBEdit.Checked)
            {
                DGV_SpecView.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
                GB_Additional.Enabled = true;
                btnLoadSpec.PerformClick();
            }

        }

        private void DGV_Data_Plus_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                cmsAction.Show(DGV_Data_Plus, e.Location);
            }
        }

        private void tsmReplace_Click(object sender, EventArgs e)
        {
            if (TDMK_OK2SHIP.admin_mode)
            {
                if (DGV_DataView.SelectedRows.Count > 0)
                {
                    int sel_row_inx = DGV_DataView.SelectedRows[0].Index;
                    foreach (DataGridViewCell c in DGV_Data_Plus.SelectedCells)
                    {
                        int src_col_inx = c.ColumnIndex;
                        string src_col_name = DGV_Data_Plus.Columns[src_col_inx].Name;
                        try
                        {
                            DGV_DataView.Rows[sel_row_inx].Cells[src_col_name].Value = c.Value;
                            if (check_columns_existed(DGV_To_Table(DGV_SpecView), src_col_name))
                            {
                                double UL = Convert.ToDouble(DGV_SpecView.Rows[1].Cells[src_col_name].Value);
                                double LL = Convert.ToDouble(DGV_SpecView.Rows[2].Cells[src_col_name].Value);
                                double Act_val = Convert.ToDouble(c.Value);
                                if ((Act_val <= UL) && (Act_val >= LL))
                                {
                                    DGV_DataView.Rows[sel_row_inx].Cells[src_col_name].Style.BackColor = Color.White;
                                }
                                else
                                {
                                    DGV_DataView.Rows[sel_row_inx].Cells[src_col_name].Style.BackColor = Color.Red;
                                }
                            }
                        }
                        catch
                        {
                            MessageBox.Show(new Form { TopMost = true },"Sai vi tri", "Thong bao");
                        }
                    }
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true },"Chọn dòng");
                }
            }
            else
            {
                MessageBox.Show("Bạn cần đăng nhập để chỉnh sửa", "Thông báo");
            }
        }
        //public void Calcul_CPK()
        //{
        //    TDMK_OK2SHIP. FAI_Spec sel_FAI_test = testFAI_spec[0];
        //    DataTable test_tbl = DGV_To_Table(DGV_DataView);
        //    string[] temp_FAI_data = test_tbl.AsEnumerable().Select(r => r.Field<string>(sel_FAI_test.FAI_Name)).ToArray();
        //    Double[] data_arr = new double[temp_FAI_data.Length];
        //    int inx = 0;
        //    foreach (string c in temp_FAI_data )
        //    {
        //        data_arr[inx] = Convert.ToDouble(c);
        //        inx++;
        //    }
        //    string Side_check = sel_FAI_test.check_side;//cbSideCheck.SelectedItem.ToString();
        //    double Norminal_Dim = Convert.ToDouble( sel_FAI_test.SetVal);// Convert.ToDouble(txtNorminalDim.Text);
        //    double UL = Convert.ToDouble(sel_FAI_test.UL);
        //    double LL = Convert.ToDouble(sel_FAI_test.LL);

        //    double stdev = CalculateStandardDeviation(data_arr);// tg.StDev(data_arr);
        //    double mean = data_arr.Average();
        //    double max = data_arr.Max();
        //    double min = data_arr.Min();

        //    double CP = (UL - LL) / (6 * stdev);
        //    double CPKL = (mean - LL) / (3 * stdev);
        //    double CPKU = (UL - mean) / (3 * stdev);
        //    double CPK = new double[] { CPKL, CPKU }.Min();
        //    double tg1 = (mean - Norminal_Dim) / stdev;
        //    double CPKM = CPK / Math.Sqrt(1 + Math.Pow(tg1, 2));

        //    double margin = 2 * stdev;
        //    double mean_minus_7sig = mean - (7 * stdev);
        //    double mean_plus_7sig = mean + (7 * stdev);

        //    double bin_start;
        //    double bin_end;
        //    if (Side_check == "SingleSide-USL")
        //    {
        //        bin_start = new double[] { mean_minus_7sig, UL - margin }.Min();
        //    }
        //    else
        //    {
        //        bin_start = new double[] { LL - margin, mean_minus_7sig }.Min();
        //    }

        //    if (Side_check == "SingleSide-LSL")
        //    {
        //        bin_end = new double[] { mean_plus_7sig, LL + margin }.Max();
        //    }
        //    else
        //    {
        //        bin_end = new double[] { UL + margin, mean_minus_7sig }.Max();
        //    }

        //    double bin_range = bin_end - bin_start;
        //    double Qty = 50;
        //    double bin_step = bin_range / Qty;
        //    DGV_SpecView.Rows[4].Cells[sel_FAI_test.FAI_Name].Value =Math.Round (stdev,4).ToString();
        //    DGV_SpecView.Rows[5].Cells[sel_FAI_test.FAI_Name].Value =Math.Round ( mean,4).ToString();
        //    DGV_SpecView.Rows[6].Cells[sel_FAI_test.FAI_Name].Value = Math.Round(max,3).ToString();
        //    DGV_SpecView.Rows[7].Cells[sel_FAI_test.FAI_Name].Value =Math.Round ( min,3).ToString();
        //    DGV_SpecView.Rows[8].Cells[sel_FAI_test.FAI_Name].Value = Math.Round( CP,3).ToString();
        //    DGV_SpecView.Rows[9].Cells[sel_FAI_test.FAI_Name].Value = Math.Round ( CPKL,3).ToString();
        //    DGV_SpecView.Rows[10].Cells[sel_FAI_test.FAI_Name].Value =Math.Round ( CPKU,3).ToString();
        //    DGV_SpecView.Rows[11].Cells[sel_FAI_test.FAI_Name].Value =Math.Round ( CPK,3).ToString();
        //    DGV_SpecView.Rows[12].Cells[sel_FAI_test.FAI_Name].Value =Math.Round( CPKM,3).ToString();
        //    double[] bin_data = new double[51];
        //    int[] freq_bin_data = new int[50];
        //    double[] modified_NormDist = new double[50];
        //    double[] NormDist_Bin = new double[50];
        //    double XiShu;

        //    for (int i = 0; i < 51; i++)
        //    {
        //        bin_data[i] = bin_start + i * bin_step;
        //    }
        //    for (int i = 0; i < 50; i++)
        //    {
        //        freq_bin_data[i] = Countif(data_arr, bin_data[i + 1]) - Countif(data_arr, bin_data[i]);
        //    }
        //    for (int i = 0; i < 50; i++)
        //    {

        //        NormDist_Bin[i] = normdist(bin_data[i], mean, stdev, false);
        //    }
        //    XiShu = freq_bin_data.Max() / NormDist_Bin.Max();
        //    for (int i = 0; i < 50; i++)
        //    {
        //        modified_NormDist[i] = NormDist_Bin[i] * XiShu;
        //    }
        //    DGV_Data_Plus.Columns.Add("Bin", "Bin");
        //    DGV_Data_Plus.Columns.Add("Freq_Bin", "Freq_Bin");
        //    DGV_Data_Plus.Columns.Add("NormDist", "NormDist");
        //    DGV_Data_Plus.Columns.Add("Modified_NormDist", "Modified_NormDist");
        //    for (int i = 0; i < 50; i++)
        //    {
        //        DGV_Data_Plus.Rows.Add(bin_data[i], freq_bin_data[i], NormDist_Bin[i], modified_NormDist[i]);
        //        DGV_Data_Plus.Rows[i].HeaderCell.Value = (i + 1).ToString();
        //    }
        //    Form2 chart_form = new Form2();
        //    chart_form.src_bin_data = bin_data;
        //    chart_form.src_freq_bin_data = freq_bin_data;
        //    chart_form.src_FAI_Data = data_arr;
        //    chart_form.FAI_Spec_val = testFAI_spec[0];
        //    chart_form.src_modified_NormDist = modified_NormDist;
        //    chart_form.Show();
        //}
        //public void Calcul_CPK_FAI(TDMK_OK2SHIP.FAI_Spec src_FAI_spec, string[] tar_FAI_Data, ref TDMK_OK2SHIP.Calculate_CPK tar_CPK_Result, ref TDMK_OK2SHIP.FAI_Histogram_Data tar_Histogram_Data)
        //{
        //    Double[] data_arr = new double[tar_FAI_Data.Length];
        //    int inx = 0;
        //    foreach (string c in tar_FAI_Data)
        //    {
        //        data_arr[inx] = Convert.ToDouble(c);
        //        inx++;
        //    }
        //    string Side_check = src_FAI_spec.check_side;//cbSideCheck.SelectedItem.ToString();
        //    double Norminal_Dim = Convert.ToDouble(src_FAI_spec.SetVal);// Convert.ToDouble(txtNorminalDim.Text);
        //    double UL = Convert.ToDouble(src_FAI_spec.UL);
        //    double LL = Convert.ToDouble(src_FAI_spec.LL);

        //    double stdev = CalculateStandardDeviation(data_arr);// tg.StDev(data_arr);
        //    double mean = data_arr.Average();
        //    double max = data_arr.Max();
        //    double min = data_arr.Min();

        //    double CP = (UL - LL) / (6 * stdev);
        //    double CPKL = (mean - LL) / (3 * stdev);
        //    double CPKU = (UL - mean) / (3 * stdev);
        //    double CPK = new double[] { CPKL, CPKU }.Min();
        //    double tg1 = (mean - Norminal_Dim) / stdev;
        //    double CPKM = CPK / Math.Sqrt(1 + Math.Pow(tg1, 2));

        //    double margin = 2 * stdev;
        //    double mean_minus_7sig = mean - (7 * stdev);
        //    double mean_plus_7sig = mean + (7 * stdev);

        //    double bin_start;
        //    double bin_end;
        //    if (Side_check == "SingleSide-USL")
        //    {
        //        bin_start = new double[] { mean_minus_7sig, UL - margin }.Min();
        //    }
        //    else
        //    {
        //        bin_start = new double[] { LL - margin, mean_minus_7sig }.Min();
        //    }

        //    if (Side_check == "SingleSide-LSL")
        //    {
        //        bin_end = new double[] { mean_plus_7sig, LL + margin }.Max();
        //    }
        //    else
        //    {
        //        bin_end = new double[] { UL + margin, mean_minus_7sig }.Max();
        //    }

        //    double bin_range = bin_end - bin_start;
        //    double Qty = 50;
        //    double bin_step = bin_range / Qty;
        //    double[] bin_data = new double[51];
        //    int[] freq_bin_data = new int[50];
        //    double[] modified_NormDist = new double[50];
        //    double[] NormDist_Bin = new double[50];
        //    double XiShu;

        //    for (int i = 0; i < 51; i++)
        //    {
        //        bin_data[i] = bin_start + i * bin_step;
        //    }
        //    for (int i = 0; i < 50; i++)
        //    {
        //        freq_bin_data[i] = Countif(data_arr, bin_data[i + 1]) - Countif(data_arr, bin_data[i]);
        //    }
        //    for (int i = 0; i < 50; i++)
        //    {

        //        NormDist_Bin[i] = normdist(bin_data[i], mean, stdev, false);
        //    }
        //    XiShu = freq_bin_data.Max() / NormDist_Bin.Max();
        //    for (int i = 0; i < 50; i++)
        //    {
        //        modified_NormDist[i] = NormDist_Bin[i] * XiShu;
        //    }
        //    tar_CPK_Result.FAI_No = src_FAI_spec.FAI_Name;
        //    tar_CPK_Result.CP = CP;
        //    tar_CPK_Result.CPK = CPK;
        //    tar_CPK_Result.CPKL = CPKL;
        //    tar_CPK_Result.CPKM = CPKM;
        //    tar_CPK_Result.CPKU = CPKU;
        //    tar_CPK_Result.max = max;
        //    tar_CPK_Result.min = min;
        //    tar_CPK_Result.stdev = stdev;
        //    tar_CPK_Result.mean = mean;

        //    tar_Histogram_Data._FAI_No = src_FAI_spec.FAI_Name;
        //    tar_Histogram_Data._Bin_data = bin_data;
        //    tar_Histogram_Data._Freq_bin_data = freq_bin_data;
        //    tar_Histogram_Data._Modified_NormDist_data = modified_NormDist;
        //    tar_Histogram_Data._FAI_Data = data_arr;

        //    //Form2 chart_form = new Form2();
        //    //chart_form.src_bin_data = bin_data;
        //    //chart_form.src_freq_bin_data = freq_bin_data;
        //    //chart_form.src_FAI_spec_Data = data_arr;
        //    //chart_form.FAI_Spec_val = testFAI[0];
        //    //chart_form.src_modified_NormDist = modified_NormDist;
        //    //chart_form.Show();
        //}
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

        private void DGV_DataView_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {


        }

        private void btnLoadSpec_Click(object sender, EventArgs e)
        {
            chklstPart.Items.Clear();
            DataTable tbl_all = TDMK_Code.Datatable_Filter(sqlcon, "All_Items", "Process_Name = '" + sel_tbl + "'");
            DataTable sequence_tbl = TDMK_Code.Datatable_Filter(sqlcon, "Sequence",TDMK_Code.filter_str(new string[] { "ItemCode","Process","Depart"},new string[] { txtItemCode.Text, sel_tbl, Sel_Depart}));
            sel_dt = TDMK_Code.Datatable_Filter(sqlcon, "SpecList", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { txtItemCode.Text, sel_tbl }));
            bool en_add = false;
            sel_ColName_lst.Clear();
            if ( (sel_dt.Rows.Count > 0) && !EditMode)
            {
                string[] Col_name = sel_dt.AsEnumerable().Select(r => r.Field<string>("Col_name")).ToArray();
                foreach (string c_name in Col_name)
                {
                    switch (Sel_Depart)
                    {
                        case "LAB":
                            switch (sel_tbl)
                            {
                                case "Etching_Process":
                                    if (!TDMK_Code.check_exist_list(c_name, Etching_DESData_list))
                                    {
                                        en_add = true;
                                    }
                                    else
                                    {
                                        en_add = false;
                                    }
                                    break;
                                case "Printing_Process":
                                    if (!TDMK_Code.check_exist_list(c_name, LAB_Printing_IgnoredList))
                                    {
                                        en_add = true;
                                    }
                                    else
                                    {
                                        en_add = false;
                                    }
                                    break;
                                case "Copper_Plating_Process":
                                    en_add = true;
                                    break;
                                case "Cover_Lay_Process":
                                    en_add = true;
                                    break;
                                default:
                                    en_add = false;
                                    break;
                            }
                            break;
                        case "Production":
                            switch (sel_tbl)
                            {
                                case "Etching_Process":
                                    if (TDMK_Code.check_exist_list(c_name, Etching_DESData_list))
                                    {
                                        en_add = true;
                                    }
                                    else
                                    {
                                        en_add = false;
                                    }
                                    break;
                                case "Printing_Process":
                                    if (TDMK_Code.check_exist_list(c_name, LAB_Printing_IgnoredList))
                                    {
                                        en_add = true;
                                    }
                                    else
                                    {
                                        en_add = false;
                                    }
                                    break;
                                default:
                                    en_add = false;
                                    break;
                            }
                            break;
                        case "UV_Process":
                            en_add = true;
                            break;
                    }
                    if (en_add)
                    {
                        DataView dv = tbl_all.AsDataView();
                        dv.RowFilter = "Col_name = '" + c_name + "'";
                        string sel_item = checkDBNull(dv[0]["Item_Name"]);
                        string internal_item = checkDBNull(dv[0]["Internal_Point"]);
                        //string[] item_name = tbl_all.AsEnumerable().Where(r => r.Field<string>("Col_name") == c_name).Select(r => r.Field<string>("Item_Name")).ToArray();
                        //chklstPart.Items.Add(item_name[0]);
                        chklstPart.Items.Add(internal_item + "-" + sel_item);
                        sel_ColName_lst.Add(c_name);
                    }
                }
                if ((sequence_tbl.Rows.Count > 0) && chklstPart.Items.Count > 0)
                {
                    foreach (DataRow dr in sequence_tbl.Rows)
                    {
                        string col_name = dr["Col_Name"].ToString();
                        string header_text = dr["Item_Name"].ToString();
                        if (!check_columns_existed(DGV_To_Table(DGV_SpecView), col_name))
                        {
                            DGV_SpecView.Columns.Add(col_name, header_text);
                        }
                    }
                }
                cur_sel_dt = sel_dt;
               
                /********************************** Test areas **********************************/

                //DGV_Data_Plus.DataSource = cur_sel_dt;

                /********************************** Finish Test areas **********************************/
            }
            else
            {
                if(EditMode)
                {
                    DataView dv = tbl_all.AsDataView();
                    for(int i=0;i<dv.Count;i++)
                    {
                        string sel_item = checkDBNull(dv[i]["Item_Name"]);
                        string internal_item = checkDBNull(dv[i]["Internal_Point"]);
                        chklstPart.Items.Add(internal_item + "-" + sel_item);
                    }
                    //string[] items = tbl_all.AsEnumerable().Select(r => r.Field<string>("Item_Name")).ToArray();
                    //foreach (string t in items)
                    //{
                    //    chklstPart.Items.Add(t);
                    //}
                    if (sequence_tbl.Rows.Count > 0)
                    {
                        foreach (DataRow dr in sequence_tbl.Rows)
                        {
                            string col_name = dr["Col_Name"].ToString();
                            string header_text = dr["Item_Name"].ToString();
                            if (!check_columns_existed(DGV_To_Table(DGV_SpecView), col_name))
                            {
                                DGV_SpecView.Columns.Add(col_name, header_text);
                            }
                        }
                    }
                    cur_sel_dt = tbl_all;
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true },"No data", "Warning");
                }
            }
        }

        //private void btnLogin_Click(object sender, EventArgs e)
        //{
        //    if (btnLogin .Text == "Logout")
        //    {
        //        btnLogin.Text = "Login";
        //        TDMK_OK2SHIP.admin_mode = false;
        //        btnLogin.BackColor = curr_color;
        //        txtOperator.Text = TDMK_OK2SHIP.curr_user;
        //    }
        //    else
        //    {
        //        TDMK_OK2SHIP.curr_user = txtOperator.Text;
        //        Form3 frm_login = new Form3();
        //        frm_login.ShowDialog();
        //    }
        //}

        private void cmsAction_Opening(object sender, CancelEventArgs e)
        {

        }

        private void tsmSelectedColumn_Click(object sender, EventArgs e)
        {
            int sel_col_inx = DGV_Data_Plus.SelectedCells[0].ColumnIndex;
            int src_col_inx = DGV_SpecView.SelectedCells[0].ColumnIndex;
            string src_col_name = DGV_SpecView.Columns[src_col_inx].Name;
            string src_col_header = src_col_name.Split('_')[0];
            if (!myCode.check_columns_existed(myCode.DGV_To_Table(DGV_Data_Plus), src_col_name))
            {
                DGV_Data_Plus.Columns[sel_col_inx].HeaderText = src_col_header;
                DGV_Data_Plus.Columns[sel_col_inx].Name = src_col_name;
                double sv = Convert.ToDouble(DGV_SpecView.Rows[0].Cells[src_col_name].Value);
                double UL = sv + Convert.ToDouble(DGV_SpecView.Rows[1].Cells[src_col_name].Value);
                double LL = sv - Convert.ToDouble(DGV_SpecView.Rows[2].Cells[src_col_name].Value);
                foreach (DataGridViewRow r in DGV_Data_Plus.Rows)
                {
                    string act_val = checkDBNull(r.Cells[src_col_name].Value);
                    if (myCode.check_in_limit(UL.ToString(), LL.ToString(), act_val,sv.ToString()))
                    {
                        r.Cells[src_col_name].Style.BackColor = Color.White;
                    }
                    else
                    {
                        if (Convert.ToDouble(act_val) > UL)
                        {
                            r.Cells[src_col_name].Style.BackColor = Color.Blue;
                        }
                        if (Convert.ToDouble(act_val) < UL)
                        {
                            r.Cells[src_col_name].Style.BackColor = Color.Gray;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },"Cannot rename");
            }
        }
        private void tsmDefault_Click(object sender, EventArgs e)
        {

            int sel_col_inx = DGV_Data_Plus.SelectedCells[0].ColumnIndex;
            string default_name = "Default" + sel_col_inx.ToString();
            if (!check_columns_existed(DGV_To_Table(DGV_Data_Plus), default_name))
            {
                DGV_Data_Plus.Columns[sel_col_inx].HeaderText = default_name;
                DGV_Data_Plus.Columns[sel_col_inx].Name = default_name;
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },"Cannot rename");
            }
        }

        private void DGV_Data_Plus_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            DGV_Cell_edit();
            //if (edit_en)
            //{
            //    DataGridViewCell curr_cell;
            //    string curr_col_name;
            //    curr_cell = DGV_Data_Plus.CurrentCell;
            //    curr_col_name = DGV_Data_Plus.Columns[curr_cell.ColumnIndex].Name.ToString();
            //    DataTable spec_tbl = DGV_To_Table(DGV_SpecView);
            //    if (checkDBNull(curr_cell.Value) != "")
            //    {
            //        if (!IsNumeric(checkDBNull(curr_cell.Value)))
            //        {
            //            MessageBox.Show("You have to enter a number!", "Warning");
            //            DGV_Data_Plus.CurrentCell.Value = "";
            //        }
            //        else
            //        {
            //            if(check_columns_existed(spec_tbl, "UV_DATA_INNER_1")&& check_columns_existed(spec_tbl, "UV_DATA_INNER_2"))
            //            {
            //                if(curr_col_name!="Inner_Rate")
            //                {
            //                    int r_inx = curr_cell.RowIndex;
            //                    string _x1 = checkDBNull(DGV_Data_Plus.Rows[r_inx].Cells["UV_DATA_INNER_1"].Value);
            //                    string _x2 = checkDBNull(DGV_Data_Plus.Rows[r_inx].Cells["UV_DATA_INNER_2"].Value);
            //                    if (IsNumeric(_x1) && IsNumeric(_x2))
            //                    {
            //                        double x1 = Convert.ToDouble(_x1);
            //                        double x2 = Convert.ToDouble(_x2);
            //                        DGV_Data_Plus.Rows[r_inx].Cells["Inner_Rate"].Value = (x1 / x2).ToString();
            //                    }

            //                }
            //            }
            //            if (check_columns_existed(spec_tbl, "UV_DATA_OUTER_1") && check_columns_existed(spec_tbl, "UV_DATA_OUTER_2"))
            //            {
            //                if (curr_col_name != "Outer_Rate")
            //                {
            //                    int r_inx = curr_cell.RowIndex;
            //                    string _x1 = checkDBNull(DGV_Data_Plus.Rows[r_inx].Cells["UV_DATA_OUTER_1"].Value);
            //                    string _x2 = checkDBNull(DGV_Data_Plus.Rows[r_inx].Cells["UV_DATA_OUTER_2"].Value);
            //                    if (IsNumeric(_x1) && IsNumeric(_x2))
            //                    {
            //                        double x1 = Convert.ToDouble(_x1);
            //                        double x2 = Convert.ToDouble(_x2);
            //                        DGV_Data_Plus.Rows[r_inx].Cells["Outer_Rate"].Value = (x1 / x2).ToString();
            //                    }

            //                }
            //            }
            //            if (check_columns_existed(spec_tbl, curr_col_name))
            //            {
            //                double UL = Convert.ToDouble(DGV_SpecView.Rows[1].Cells[curr_col_name].Value);
            //                double LL = Convert.ToDouble(DGV_SpecView.Rows[2].Cells[curr_col_name].Value);
            //                double act_val = Convert.ToDouble(DGV_Data_Plus.CurrentCell.Value);
            //                if ((act_val <= UL) && (act_val >= LL))
            //                {
            //                    curr_cell.Style.BackColor = Color.White;
            //                }
            //                else
            //                {
            //                    curr_cell.Style.BackColor = Color.Red;
            //                }
            //            }

            //        }
            //    }
            //}
        }

        public void Set_target_Table(RadioButton src_RB)
        {
            if ((src_RB.Checked)&& system_ready)
            {
                sel_tbl = src_RB.Text;
                DGV_SpecView.DataSource = null;
                DGV_SpecView.Columns.Clear();
                btnLoadSpec.PerformClick();
            }
        }

        private void RB_Etching_Process_CheckedChanged(object sender, EventArgs e)
        {
            Set_target_Table(RB_Etching_Process);
        }

        private void RB_CopperPlating_CheckedChanged(object sender, EventArgs e)
        {
            Set_target_Table(RB_CopperPlating);
        }

        private void RB_Printing_Process_CheckedChanged(object sender, EventArgs e)
        {
            Set_target_Table(RB_Printing_Process);
        }

        private void RB_UV_Process_CheckedChanged(object sender, EventArgs e)
        {
            if(RB_UV_Process.Checked)
            {
                Sel_Depart = RB_UV_Process.Text;
                GBIPQC_Items.Enabled = false;
            }
            Set_target_Table(RB_UV_Process);
        }

        private void RB_CoverLay_Process_CheckedChanged(object sender, EventArgs e)
        {
            Set_target_Table(RB_CoverLay_Process);
        }
        private void lblQty_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            FrmLogin frm_login = new FrmLogin();
            frm_login.ShowDialog();
        }

        private void numQty_Validated(object sender, EventArgs e)
        {
            numQty.Enabled = false;
        }

        private void chklstPart_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                chklstPart.ContextMenuStrip = cmsAddItem;
            }
        }

        private void tsmSelectItem_Click(object sender, EventArgs e)
        {
            int cur_inx = chklstPart.SelectedIndex;
            string col_name;
            string col_header;
            if (chklstPart.GetItemChecked(cur_inx))
            {
                col_name = cur_sel_dt.Rows[cur_inx]["Col_Name"].ToString();
                col_header = chklstPart.Items[cur_inx].ToString();

                if (!myCode.check_columns_existed(myCode.DGV_To_Table(DGV_SpecView), col_name))
                {
                    DGV_SpecView.Columns.Add(col_name, col_header);
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true },"This item already selected!", "Warning");
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },"Please, select checked item", "Warning");
            }
        }

        private void tsmRemove_Click(object sender, EventArgs e)
        {
            if (DGV_Data_Plus.SelectedRows.Count > 0)
            {
                foreach (DataGridViewRow dgv_r in DGV_Data_Plus.SelectedRows)
                {
                    DGV_Data_Plus.Rows.Remove(dgv_r);
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },"Please, select rows to remove");
            }
        }

        private void tsmInsert_Click(object sender, EventArgs e)
        {

                if (myCode.insert_IPQC_data_DGV(DGV_Data_Plus, DGV_DataView, myCode.DGV_To_Table(DGV_SpecView)))
                {

                }


        }
        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            myCode.CopyToClipboard(DGV_Data_Plus);
        }

        private void selectedRowsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            myCode.PasteClipboardValue(false, DGV_Data_Plus, DGV_SpecView);
        }

        private void selectedColumnsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            myCode.PasteClipboardValue(true, DGV_Data_Plus, DGV_SpecView);
        }
        public void Load_Test_IPQC_Data(string format_loc, string tar_ItemCode, string tar_LotNo, string file_extension, DataGridView tar_DGV_Data)
        {
            string tar_format = tar_ItemCode + tar_LotNo;
            string tar_format_file = Path.Combine(format_loc, tar_format + file_extension);
            DataTable spec_tbl = TDMK_Code.Datatable_Filter(sqlcon, "SpecList", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { txtItemCode.Text, sel_tbl }));
            string[] addr_lst = spec_tbl.AsEnumerable().Select(r => r.Field<string>("Address")).ToArray();
            string[] Colname_lst = spec_tbl.AsEnumerable().Select(r => r.Field<string>("Col_Name")).ToArray();
            string [] item_lst = spec_tbl.AsEnumerable().Select(r => r.Field<string>("Item")).ToArray();
            DataTable data_tbl = new DataTable();
            btnSave.Enabled = true;
            foreach(string _Colname in Colname_lst)
            {
                //DGV_DataView.Columns.Add(_addr, _addr);
                data_tbl.Columns.Add(_Colname);
            }
            

            if (File.Exists(tar_format_file))
            {
                myExcel.Workbook tar_wkbook = TDMK_Code.open_excel_file(tar_format_file, "", "");
                int wrksheet_num = tar_wkbook.Worksheets.Count;
                foreach (myExcel.Worksheet tg in tar_wkbook.Worksheets)
                {
                    if (tg.Name.Contains("IPQC"))
                    {                                             
                        for(int r_inx =0;r_inx< 32; r_inx++)
                        {
                            int c_inx = 0;
                            DataRow dr = data_tbl.NewRow();
                            foreach (string _addr in addr_lst)
                            {
                                myExcel.Range data_rgn = tg.Range[_addr].Offset[r_inx,0];
                                dr[c_inx] = data_rgn.Value;
                                c_inx++;
                            }
                            data_tbl.Rows.Add(dr);
                        }
                    }
                }
                tar_DGV_Data.DataSource = data_tbl;
                for(int i=0;i< tar_DGV_Data.Columns.Count;i++)
                {
                    tar_DGV_Data.Columns[i].HeaderText = item_lst[i];
                }
                int temp_r = 0;
                foreach(DataGridViewRow dgv_r in tar_DGV_Data.Rows)
                {
                    dgv_r.HeaderCell.Value = (temp_r + 1).ToString();
                    temp_r++;
                }
                tar_DGV_Data.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                Disable_Sort_DGV(tar_DGV_Data);
                MessageBox.Show(new Form { TopMost = true },"Finished!", "Thong bao");
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },"File format khong tim thay");
            }
        }
        public void Load_Test_Data2(string format_loc, string tar_ItemCode, string tar_LotNo, string file_extension, DataGridView tar_DGV_Data)
        {
            string tar_format = tar_ItemCode + tar_LotNo;
            string tar_format_file = Path.Combine(format_loc, tar_format + file_extension);
            DataTable spec_tbl = new DataTable();
            if ((sel_tbl=="Roughness")|| (sel_tbl == "AU_NI"))
            {
                if (sel_tbl == "AU_NI")
                {
                    spec_tbl = TDMK_Code.Datatable_Filter(sqlcon, "SpecList", TDMK_Code.filter_str(new string[] { "ItemCode", "Col_Name" }, new string[] { txtItemCode.Text, "%THICKNESS%" }));
                }
                else
                {
                    spec_tbl = TDMK_Code.Datatable_Filter(sqlcon, "SpecList", TDMK_Code.filter_str(new string[] { "ItemCode", "Col_Name" }, new string[] { txtItemCode.Text, "%" + sel_tbl + "%" }));
                }
                
            }
            else
            {
                spec_tbl = TDMK_Code.Datatable_Filter(sqlcon, "SpecList", TDMK_Code.filter_str(new string[] { "ItemCode", "Col_Name" }, new string[] { txtItemCode.Text, sel_tbl }));
            }
            string[] addr_lst = spec_tbl.AsEnumerable().Select(r => r.Field<string>("Address")).ToArray();
            string[] Colname_lst = spec_tbl.AsEnumerable().Select(r => r.Field<string>("Col_Name")).ToArray();
            string[] item_lst = spec_tbl.AsEnumerable().Select(r => r.Field<string>("Item")).ToArray();
            DataTable data_tbl = new DataTable();
            btnSave.Enabled = true;
            foreach (string _Colname in Colname_lst)
            {
                //DGV_DataView.Columns.Add(_addr, _addr);
                data_tbl.Columns.Add(_Colname);
            }


            if (File.Exists(tar_format_file))
            {
                myExcel.Workbook tar_wkbook = TDMK_Code.open_excel_file(tar_format_file, "", "");
                int wrksheet_num = tar_wkbook.Worksheets.Count;
                foreach (myExcel.Worksheet tg in tar_wkbook.Worksheets)
                {
                    if (tg.Name.Contains("IPQC"))
                    {
                        for (int r_inx = 0; r_inx < 32; r_inx++)
                        {
                            int c_inx = 0;
                            DataRow dr = data_tbl.NewRow();
                            foreach (string _addr in addr_lst)
                            {
                                myExcel.Range data_rgn = tg.Range[_addr].Offset[r_inx, 0];
                                dr[c_inx] = data_rgn.Value;
                                c_inx++;
                            }
                            data_tbl.Rows.Add(dr);
                        }
                    }
                }
                tar_DGV_Data.DataSource = data_tbl;
                for (int i = 0; i < tar_DGV_Data.Columns.Count; i++)
                {
                    tar_DGV_Data.Columns[i].HeaderText = item_lst[i];
                }
                int temp_r = 0;
                foreach (DataGridViewRow dgv_r in tar_DGV_Data.Rows)
                {
                    dgv_r.HeaderCell.Value = (temp_r + 1).ToString();
                    temp_r++;
                }
                tar_DGV_Data.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                Disable_Sort_DGV(tar_DGV_Data);
                MessageBox.Show(new Form { TopMost = true },"Finished!", "Thong bao");
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },"File format khong tim thay");
            }
        }

        private void chklstPart_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            int cur_inx = chklstPart.SelectedIndex;
            string col_name;
            string col_header;
            try
            {
                if(EditMode)
                {
                    col_name = cur_sel_dt.Rows[cur_inx]["Col_Name"].ToString();
                }
                else
                {
                    col_name = sel_ColName_lst[cur_inx];
                }               
            }
            catch
            {
                col_name = chklstPart.Items[cur_inx].ToString(); // "Column_" + (DGV_SpecView.Columns.Count + 1).ToString();
            }           
            col_header = chklstPart.Items[cur_inx].ToString();
            if (!myCode.check_columns_existed(myCode.DGV_To_Table(DGV_SpecView), col_name))
            {
                DGV_SpecView.Columns.Add(col_name, col_header);
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },"This item already selected!", "Warning");
            }
        }

        private void RB_Lab_CheckedChanged(object sender, EventArgs e)
        {
            if((RB_Lab.Checked)&& system_ready)
            {
                Sel_Depart = RB_Lab.Text;
                GBIPQC_Items.Enabled = true;
                DGV_SpecView.DataSource = null;
                DGV_SpecView.Columns.Clear();
                foreach (RadioButton c in GBIPQC_Items.Controls)
                {
                    if (c.Checked)
                    {
                        Set_target_Table(c);
                        break;
                    }
                }
                btnLoadSpec.PerformClick();
            }
        }

        private void RB_Etching_CheckedChanged(object sender, EventArgs e)
        {
            if((RB_Etching.Checked)&& system_ready)
            {
                Sel_Depart = RB_Etching.Text;
                GBIPQC_Items.Enabled = true;
                DGV_SpecView.DataSource = null;
                DGV_SpecView.Columns.Clear();
                foreach (RadioButton c in GBIPQC_Items.Controls)
                {
                    if (c.Checked)
                    {
                        Set_target_Table(c);
                        break;
                    }
                }
                btnLoadSpec.PerformClick();
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if(txtAdd.Text!="")
            {
                chklstPart.Items.Add(txtAdd.Text);
                
            }
        }

        private void DGV_Data_Plus_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            DGV_Cell_edit();
        }
        public void DGV_Cell_edit()
        {
            if (edit_en)
            {
                DataGridViewCell curr_cell;
                string curr_col_name;
                curr_cell = DGV_Data_Plus.CurrentCell;
                curr_col_name = DGV_Data_Plus.Columns[curr_cell.ColumnIndex].Name.ToString();
                DataTable spec_tbl = DGV_To_Table(DGV_SpecView);
                if (checkDBNull(curr_cell.Value) != "")
                {
                    if (!IsNumeric(checkDBNull(curr_cell.Value)))
                    {
                        MessageBox.Show(new Form { TopMost = true },"You have to enter a number!", "Warning");
                        DGV_Data_Plus.CurrentCell.Value = "";
                    }
                    else
                    {
                        if (check_columns_existed(spec_tbl, "UV_DATA_INNER_1") && check_columns_existed(spec_tbl, "UV_DATA_INNER_2"))
                        {
                            if (curr_col_name != "Inner_Rate")
                            {
                                int r_inx = curr_cell.RowIndex;
                                string _x1 = checkDBNull(DGV_Data_Plus.Rows[r_inx].Cells["UV_DATA_INNER_1"].Value);
                                string _x2 = checkDBNull(DGV_Data_Plus.Rows[r_inx].Cells["UV_DATA_INNER_2"].Value);
                                if (IsNumeric(_x1) && IsNumeric(_x2))
                                {
                                    double x1 = Convert.ToDouble(_x1);
                                    double x2 = Convert.ToDouble(_x2);
                                    double result = (x2 / x1) * 100;
                                    DGV_Data_Plus.Rows[r_inx].Cells["Inner_Rate"].Value = result.ToString("#0.#0");
                                    if (result <= 80)
                                    {
                                        DGV_Data_Plus.Rows[r_inx].Cells["Inner_Rate"].Style.BackColor = Color.Yellow;
                                    }
                                    else
                                    {
                                        DGV_Data_Plus.Rows[r_inx].Cells["Inner_Rate"].Style.BackColor = Color.White;
                                    }
                                }

                            }
                        }
                        if (check_columns_existed(spec_tbl, "UV_DATA_OUTER_1") && check_columns_existed(spec_tbl, "UV_DATA_OUTER_2"))
                        {
                            if (curr_col_name != "Outer_Rate")
                            {
                                int r_inx = curr_cell.RowIndex;
                                string _x1 = checkDBNull(DGV_Data_Plus.Rows[r_inx].Cells["UV_DATA_OUTER_1"].Value);
                                string _x2 = checkDBNull(DGV_Data_Plus.Rows[r_inx].Cells["UV_DATA_OUTER_2"].Value);
                                if (IsNumeric(_x1) && IsNumeric(_x2))
                                {
                                    double x1 = Convert.ToDouble(_x1);
                                    double x2 = Convert.ToDouble(_x2);
                                    double result = (x2 / x1) * 100;
                                    DGV_Data_Plus.Rows[r_inx].Cells["Outer_Rate"].Value = result.ToString("#0.#0");
                                    if(result <=80)
                                    {
                                        DGV_Data_Plus.Rows[r_inx].Cells["Outer_Rate"].Style.BackColor = Color.Yellow;
                                    }
                                    else
                                    {
                                        DGV_Data_Plus.Rows[r_inx].Cells["Outer_Rate"].Style.BackColor = Color.White;
                                    }
                                }

                            }
                        }
                        if ((check_columns_existed(spec_tbl, curr_col_name))&&(!curr_col_name.Contains("Rate")))
                        {
                            string UL = checkDBNull( DGV_SpecView.Rows[1].Cells[curr_col_name].Value);
                            string LL = checkDBNull( DGV_SpecView.Rows[2].Cells[curr_col_name].Value);
                            string act_val = checkDBNull( DGV_Data_Plus.CurrentCell.Value);
                            string SV = checkDBNull(DGV_SpecView.Rows[0].Cells[curr_col_name].Value);
                            //curr_cell.Style.BackColor = myCode.check_in_limit2(UL, LL, act_val);
                            //bool UL_equal_comp = true;
                            //bool LL_equal_comp = true;
                            //if (!myCode.IsNumeric(SV))
                            //{
                            //    if (!SV.Contains("="))
                            //    {
                            //        UL_equal_comp = false;
                            //        LL_equal_comp = false;
                            //    }
                            //}
                            curr_cell.Style.BackColor = myCode.check_in_limit_Color(UL, LL, act_val, SV);

                        }
                    }
                }
            }
        }
        private void DGV_SpecView_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            sel_spec_col = e.ColumnIndex;
        }

        private void DGV_SpecView_MouseDown(object sender, MouseEventArgs e)
        {
            if ((DGV_SpecView.Columns.Count > 0) && EditMode && (e.Button == MouseButtons.Right) && (DGV_SpecView.Rows.Count == 0))
            {
                
                cmsSpec_Remove.Show(DGV_SpecView, e.Location);
            }
        }

        private void tsmRemove_Spec_Click(object sender, EventArgs e)
        {
            
            string item_remove = DGV_SpecView.Columns[sel_spec_col].Name;
            TDMK_Code.Delelte_FilteredItem_arr("Sequence", sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode", "Col_Name" }, new string[] { txtItemCode.Text, item_remove }));
            DGV_SpecView.Columns.RemoveAt(sel_spec_col);
            //DGV_SpecView.Columns.Clear();
            //DGV_SpecView.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "Sequence", "ItemCode = '" + txtItemCode.Text + "'");
        }

        private void DGV_DataView_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            if (edit_en)
            {
                DataGridViewCell curr_cell;
                string curr_col_name;
                curr_cell = DGV_DataView.CurrentCell;
                curr_col_name = DGV_DataView.Columns[curr_cell.ColumnIndex].Name.ToString();
                if (checkDBNull(curr_cell.Value) != "")
                {
                    if (!IsNumeric(checkDBNull(curr_cell.Value)))
                    {
                        MessageBox.Show(new Form { TopMost = true },"You have to enter a number!", "Warning");
                        DGV_DataView.CurrentCell.Value = "";
                    }
                    else
                    {
                        if (check_columns_existed(DGV_To_Table(DGV_SpecView), curr_col_name))
                        {
                            string UL = checkDBNull(DGV_SpecView.Rows[1].Cells[curr_col_name].Value);
                            string LL = checkDBNull(DGV_SpecView.Rows[2].Cells[curr_col_name].Value);
                            string act_val = checkDBNull(DGV_DataView.CurrentCell.Value);
                            string SV = checkDBNull(DGV_SpecView.Rows[0].Cells[curr_col_name].Value);                           
                            curr_cell.Style.BackColor = myCode.check_in_limit_Color(UL, LL, act_val, SV);
                        }
                    }
                }

            }
        }

        private void DGV_DataView_MouseClick(object sender, MouseEventArgs e)
        {
            if(e.Button == MouseButtons.Right)
            {
                cmsPaste.Show(DGV_DataView, e.Location);
            }
        }

        private void tsmPasteRow_Click(object sender, EventArgs e)
        {
            myCode.PasteClipboardValue(false, DGV_DataView, DGV_SpecView);
        }

        private void tsmPasteColumns_Click(object sender, EventArgs e)
        {
            myCode.PasteClipboardValue(true, DGV_DataView, DGV_SpecView);
        }
        public DataTable Load_Setting_Process(string ItemCode, string Process_name)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, Process_name });
            DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, "SpecList", filter_str);
            DataTable _src_dt = src_dt.Clone();
            bool en_add = false;
            foreach (DataRow dgv_r in src_dt.Rows)
            {
                string cell_val = dgv_r["Col_Name"].ToString();

                switch (Sel_Depart)
                {
                    case "LAB":
                        switch (sel_tbl)
                        {
                            case "Etching_Process":
                                if (!TDMK_Code.check_exist_list(cell_val, Etching_DESData_list))
                                {
                                    en_add = true;
                                }
                                else
                                {
                                    en_add = false;
                                }
                                break;
                            case "Printing_Process":
                                if (!TDMK_Code.check_exist_list(cell_val, LAB_Printing_IgnoredList))
                                {
                                    en_add = true;
                                }
                                else
                                {
                                    en_add = false;
                                }
                                break;
                            case "Copper_Plating_Process":
                                en_add = true;
                                break;
                            case "Cover_Lay_Process":
                                en_add = true;
                                break;
                            default:
                                en_add = true;
                                break;
                        }
                        break;
                    case "Production":
                        switch (sel_tbl)
                        {
                            case "Etching_Process":
                                if (TDMK_Code.check_exist_list(cell_val, Etching_DESData_list))
                                {
                                    en_add = true;
                                }
                                else
                                {
                                    en_add = false;
                                }
                                break;
                            case "Printing_Process":
                                if (TDMK_Code.check_exist_list(cell_val, LAB_Printing_IgnoredList))
                                {
                                    en_add = true;
                                }
                                else
                                {
                                    en_add = false;
                                }
                                break;
                            default:
                                en_add = true;
                                break;
                        }

                        break;
                    case "UV_Process":
                        en_add = true;
                        break;
                }
                if (en_add)
                {
                    DataRow dr = _src_dt.NewRow();
                    for (int i = 0; i < src_dt.Columns.Count; i++)
                    {
                        dr[i] = dgv_r[i];
                    }
                    _src_dt.Rows.Add(dr);
                }
            }
            return _src_dt;
        }

        private void ckb_Logfile_en_CheckedChanged(object sender, EventArgs e)
        {
            if(ckb_Logfile_en.Checked)
            {
                txtLogfile_Loc.Enabled = true;
            }
            else
            {
                txtLogfile_Loc.Enabled = false;
                txtLogfile_Loc.Clear();
            }
        }

        private void txtLogfile_Loc_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void txtLogfile_Loc_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            FolderBrowserDialog f_open = new FolderBrowserDialog();
            f_open.SelectedPath = Application.StartupPath;
            if (f_open.ShowDialog() == DialogResult.OK)
            {
                txtLogfile_Loc.Text = f_open.SelectedPath;
                string logfile_name = f_open.SelectedPath;
                if(logfile_name.Contains(txtItemCode.Text)&&(logfile_name.Contains(txtLotNo.Text)))
                {
                    if (RBEdit.Checked)
                    {
                        proc_data.Fill_LogFile_Data_CheckSpec(f_open.SelectedPath, sqlcon, txtItemCode.Text, sel_tbl, DGV_Data_Plus, DGV_SpecView);
                    }
                    else
                    {
                        proc_data.Fill_LogFile_Data_CheckSpec(f_open.SelectedPath, sqlcon, txtItemCode.Text, sel_tbl, DGV_DataView, DGV_SpecView);
                    }
                }
                else
                {
                    MessageBox.Show("Please, select corrected logfile folder of ItemCode: " + txtItemCode.Text + " / Lotno: " + txtLotNo.Text, "Warning");
                }
            }
        }

        private void txtLogfile_Loc_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }
        private void txtLogfile_Loc_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                if (RBEdit.Checked)
                {
                    proc_data.Fill_LogFile_Data_CheckSpec(txtLogfile_Loc.Text, sqlcon, txtItemCode.Text, sel_tbl, DGV_Data_Plus, DGV_SpecView);
                }
                else
                {
                    proc_data.Fill_LogFile_Data_CheckSpec(txtLogfile_Loc.Text, sqlcon, txtItemCode.Text, sel_tbl, DGV_DataView, DGV_SpecView);
                }
            }
        }
    }
}
