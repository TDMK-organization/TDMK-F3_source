using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using myExcel = Microsoft.Office.Interop.Excel;
using TDMK_SQL;
using System.IO;
using TDMK_SEEV_DLL;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Drawing;
using OK2SHIP;
using VHX;
using System.Diagnostics;
//using Microsoft.Office.Interop.Excel;
namespace Echeck_LogFile_Process
{
    public class ECheck_Process
    {
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SEI_Lib myCode = new SEI_Lib();
        myVar exp_proc = new myVar();
        //public List<SortedDictionary<uint, double>> myData = new List<SortedDictionary<uint, double>>();
        //public List<SortedDictionary<uint, double>> myResult = new List<SortedDictionary<uint, double>>();
        public struct block_data
        {
            public int row { get; set; }
            public int col { get; set; }
            public block_data(int in_row, int in_col)
            {
                row = in_row;
                col = in_col;
            }
        }
        public struct NG_list
        {
            public int row_inx { get; set; }
            public string col_val { get; set; }
            public NG_list(int _rinx, string _col)
            {
                row_inx = _rinx;
                col_val = _col;
            }
        }
        public struct NG_list2
        {
            public List< int> lst_row_inx { get; set; }
            public string col_val { get; set; }
            public NG_list2(List<int> _lst_rinx, string _col)
            {
                lst_row_inx = _lst_rinx;
                col_val = _col;
            }
        }
        public DataTable Tayo_3GMRD_Excel(string logfile, string tar_range, int num_NET)
        {
            myExcel.Workbook tar_wrkbk = TDMK_Code.open_excel_file(logfile, "", "");
            myExcel.Worksheet tar_wrksht = tar_wrkbk.Sheets[1];
            List<List<string>> tar_lst_Pattern = new List<List<string>>();
            DataTable _result = new DataTable();
            myExcel.Range sel_rgn = tar_wrksht.Range[tar_range];
            int r_inx = 0;
            int lst_inx = 0;
            int col_inx = 0;
            int item_inx = 0;
            while (Convert.ToString(sel_rgn.Offset[0, col_inx].Value) != null)
            {
                while (Convert.ToString(sel_rgn.Offset[r_inx, col_inx].Value) != null)
                {
                    int mod = item_inx % num_NET;
                    if (mod == 0)
                    {
                        tar_lst_Pattern.Add(new List<string>());
                    }
                    lst_inx = item_inx / num_NET;
                    tar_lst_Pattern[lst_inx].Add(Convert.ToString(sel_rgn.Offset[r_inx, col_inx].Value));
                    r_inx += 1;
                    item_inx += 1;
                }
                r_inx = 0;
                col_inx += 1;
            }
            _result = LogData_To_Datatable(tar_lst_Pattern, "ItemNo");
            return _result;
        }
        public DataTable Tayo_3GMRD_Process(string logfile, int num_PCS, ref DataTable tar_NET_spec)
        {
            DataTable dt = Tayo_LogFile(logfile);
            int num_NET = dt.Rows.Count * (dt.Columns.Count-6) / num_PCS;
            DataTable dt_spec = new DataTable();
            List<string> spec_col = new List<string>();
            for (int i = 0; i < 2; i++)
            {
                spec_col.Add(dt.Columns[i].ColumnName);
            }
            for (int i = 5; i > 3; i--)
            {
                spec_col.Add(dt.Columns[i].ColumnName);
            }
            dt_spec = dt.AsDataView().ToTable(false, spec_col.ToArray());
            while (dt_spec.Rows.Count > num_NET)
            {
                int inx = dt_spec.Rows.Count - 1;
                DataRow dr = dt_spec.Rows[inx];
                dt_spec.Rows.Remove(dr);
            }
            List<List<string>> tar_lst_Pattern = new List<List<string>>();
            DataTable _result = new DataTable();
            int lst_inx = 0;
            int item_inx = 0;
            for (int col = 6; col < dt.Columns.Count; col++)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    int mod = item_inx % num_NET;
                    if (mod == 0)
                    {
                        tar_lst_Pattern.Add(new List<string>());
                    }
                    lst_inx = item_inx / num_NET;
                    tar_lst_Pattern[lst_inx].Add(dr[col].ToString());
                    item_inx += 1;
                }
            }
            _result = LogData_To_Datatable(tar_lst_Pattern, "ItemNo");
            tar_NET_spec = dt_spec;
            return _result;
        }
        public int count_total_row(myExcel.Worksheet src_wrksht, string start_addr)
        {
            int _result = 0;
            myExcel.Range sel_rgn = src_wrksht.Range[start_addr];
            int r_inx = 0;
            while (sel_rgn.Offset[r_inx, 0].Value != null)
            {
                r_inx++;
            }
            _result = r_inx;
            return _result;
        }
        public int count_total_col(myExcel.Worksheet src_wrksht, string start_addr)
        {
            int _result = 0;
            myExcel.Range sel_rgn = src_wrksht.Range[start_addr];
            int col_inx = 0;
            while (sel_rgn.Offset[0, col_inx].Value != null)
            {
                col_inx++;
            }
            _result = col_inx;
            return _result;
        }
        public int count_item_total(myExcel.Worksheet src_wrksht, string start_addr, string item_val)
        {
            int _result = 0;
            myExcel.Range sel_rgn = src_wrksht.Range[start_addr];
            int col_inx = 0;
            string item = Convert.ToString(sel_rgn.Offset[0, col_inx].Value);
            while (item != null && item.Contains(item_val))
            {
                col_inx++;
                item = Convert.ToString(sel_rgn.Offset[0, col_inx].Value);
            }
            _result = col_inx;
            return _result;
        }
        public List<List<string>> Yamaha_3GSPD2(myExcel.Worksheet tar_wrksht)
        {
            myExcel.Range tar_rgn = tar_wrksht.Range["F20"];
            int data_row = count_total_row(tar_wrksht, "F20");
            int list_count = count_item_total(tar_wrksht, "G18", "Piece");
            List<List<string>> src_lst = new List<List<string>>();
            block_data myblock = new block_data(count_total_row(tar_wrksht, "G20"), count_total_col(tar_wrksht, "G20"));
            for (int i = 0; i < list_count; i++)
            {
                src_lst.Add(new List<string>());
                int step = 0;
                while (myblock.row * step < data_row)
                {
                    if (tar_rgn.Offset[myblock.row * step, i + 1].Value != null)
                    {
                        myExcel.Range sel_rgn = tar_rgn.Offset[myblock.row * step, i + 1];
                        for (int j = 0; j < myblock.row; j++)
                        {
                            src_lst[i].Add(Convert.ToString(sel_rgn.Offset[j, 0].Value));
                        }
                        break;
                    }
                    else
                    {
                        step++;
                    }
                }
            }
            return src_lst;
        }
        public DataTable LogData_To_Datatable(List<List<string>> src_lst, string col_title)
        {
            DataTable dt = new DataTable();
            for (int i = 0; i < src_lst.Count; i++)
            {
                string col_name = col_title + "_" + (i + 1).ToString();
                dt.Columns.Add(col_name);
                int r_inx = 0;
                foreach (string t in src_lst[i])
                {
                    if (dt.Rows.Count <= r_inx)
                    {
                        dt.Rows.Add();
                    }
                    dt.Rows[r_inx][col_name] = t;
                    r_inx++;
                }
            }
            return dt;
        }
        public DataTable Yamaha_3GSPD_Excel(string logfile)
        {
            DataTable _result = new DataTable();
            myExcel.Workbook tar_wrkbk = TDMK_Code.open_excel_file(logfile, "", "");
            myExcel.Worksheet tar_wrksht = tar_wrkbk.Sheets[1];
            myExcel.Range tar_rgn = tar_wrksht.Range["F20"];
            int data_row = count_total_row(tar_wrksht, "F20");
            int list_count = count_item_total(tar_wrksht, "G18", "Piece");
            List<List<string>> src_lst = new List<List<string>>();
            block_data myblock = new block_data(count_total_row(tar_wrksht, "G20"), count_total_col(tar_wrksht, "G20"));
            for (int i = 0; i < list_count; i++)
            {
                src_lst.Add(new List<string>());
                int step = 0;
                while (myblock.row * step < data_row)
                {
                    if (tar_rgn.Offset[myblock.row * step, i + 1].Value != null)
                    {
                        myExcel.Range sel_rgn = tar_rgn.Offset[myblock.row * step, i + 1];
                        for (int j = 0; j < myblock.row; j++)
                        {
                            src_lst[i].Add(Convert.ToString(sel_rgn.Offset[j, 0].Value));
                        }
                        break;
                    }
                    else
                    {
                        step++;
                    }
                }
            }
            _result = LogData_To_Datatable(src_lst, "ItemNo");
            return _result;
        }
        public DataTable Yamaha_3GSPD_Process(string logfile, ref DataTable tar_NET_spec)
        {
            DataTable _result = new DataTable();
            DataTable dt = Yamaha_LogFile(logfile, ref tar_NET_spec);
            List<List<string>> src_lst = new List<List<string>>();
            int inx = 0;
            foreach (DataColumn dc in dt.Columns)
            {
                src_lst.Add(new List<string>());
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr[dc].ToString() != "")
                    {
                        src_lst[inx].Add(dr[dc].ToString());
                    }
                }
                inx++;
            }
            _result = LogData_To_Datatable(src_lst, "ItemNo");
            return _result;
        }
        public DataTable Tayo_LogFile(string filename)
        {
            string CSVFilePathName = filename;
            string[] Lines = File.ReadAllLines(CSVFilePathName);
            string[] Fields;
            Fields = Lines[0].Split(new char[] { ',' });
            int Cols = 0;
            DataTable dt = new DataTable();
            for (int i = 0; i < Fields.GetLength(0); i++)
            {
                if (Fields[i] != "")
                {
                    string col_name = "ItemNo_" + (i + 1).ToString();
                    dt.Columns.Add(col_name, typeof(string));
                    Cols++;
                }
                else
                {
                    break;
                }
            }
            DataRow Row;
            for (int i = 0; i < Lines.GetLength(0); i++)
            {
                Fields = Lines[i].Split(new char[] { ',' });
                Row = dt.NewRow();
                for (int f = 0; f < Cols; f++)
                {
                    Row[f] = Fields[f];
                }
                dt.Rows.Add(Row);
            }
            return dt;
        }
        public DataTable Yamaha_LogFile_old(string filename)
        {
            string CSVFilePathName = filename;
            string[] Lines = File.ReadAllLines(CSVFilePathName);
            string[] Fields;
            List<string> data_col = new List<string>();
            int inx = 0;
            foreach (string li in Lines)
            {
                if (li.Contains("Pin") && li.Contains("Piece"))
                {
                    break;
                }
                else
                {
                    inx++;
                }
            }
            Fields = Lines[inx].Split(new char[] { ',' });
            int Cols = Fields.GetLength(0);
            DataTable dt = new DataTable();
            for (int i = 0; i < Cols; i++)
            {
                dt.Columns.Add(Fields[i], typeof(string));
                if (Fields[i].Contains("Piece"))
                {
                    data_col.Add(Fields[i]);
                }
            }
            DataRow Row;
            for (int i = inx + 2; i < Lines.GetLength(0); i++)
            {
                Fields = Lines[i].Split(new char[] { ',' });
                if ((Fields.Length == dt.Columns.Count) && (Fields.Distinct().Count() > 1))
                {
                    Row = dt.NewRow();
                    for (int f = 0; f < Cols; f++)
                        Row[f] = checkNumeric(Fields[f]);
                    dt.Rows.Add(Row);
                }
                else
                {
                    break;
                }
            }
            return dt.AsDataView().ToTable(false, data_col.ToArray());
        }
        public DataTable Yamaha_LogFile(string filename, ref DataTable tbl_Net_Spec)
        {
            string CSVFilePathName = filename;
            string[] Lines = File.ReadAllLines(CSVFilePathName);
            string[] Fields;
            List<string> data_col = new List<string>();
            int inx = 0;
            foreach (string li in Lines)
            {
                if (li.Contains("Pin") && li.Contains("Piece"))
                {
                    break;
                }
                else
                {
                    inx++;
                }
            }
            Fields = Lines[inx].Split(new char[] { ',' });
            int Cols = Fields.GetLength(0);
            DataTable dt = new DataTable();
            for (int i = 0; i < Cols; i++)
            {
                dt.Columns.Add(Fields[i], typeof(string));

                if (Fields[i].Contains("Piece"))
                {
                    data_col.Add(Fields[i]);
                }
            }
            DataRow Row;
            for (int i = inx + 2; i < Lines.GetLength(0); i++)
            {
                Fields = Lines[i].Split(new char[] { ',' });
                if ((Fields.Length == dt.Columns.Count) && (Fields.Distinct().Count() > 1))
                {
                    Row = dt.NewRow();
                    for (int f = 0; f < Cols; f++)
                    {
                        Row[f] = checkNumeric(Fields[f]);
                    }    
                        
                    dt.Rows.Add(Row);
                }
                else
                {
                    if(dt.Rows.Count==0)
                    {
                        continue;
                    }
                    else
                    {
                        break;
                    }
                }
            }
            DataTable dt1 = dt.AsEnumerable().Where(x => x.Field<string>("Method").Equals("3")).CopyToDataTable();
            DataTable dt2 = dt1.AsEnumerable().Where(x => x.Field<string>("Piece 1") != "").CopyToDataTable();
            tbl_Net_Spec = dt2.AsDataView().ToTable(false, new string[] { "PinA", "PinB", "Threshold L", "Threshold U" });
            foreach(DataRow dr in tbl_Net_Spec.Rows)
            {
                string LL = myCode.checkDBNull(dr["Threshold L"]);
                string UL = myCode.checkDBNull(dr["Threshold U"]);
                if(TDMK_Code.IsNumeric(LL))
                {
                    dr["Threshold L"] = Convert.ToDouble(LL) * 1000;
                }
                if (TDMK_Code.IsNumeric(UL))
                {
                    dr["Threshold U"] = Convert.ToDouble(UL) * 1000;
                }
            }
            DataTable result_dt = dt1.AsDataView().ToTable(false, data_col.ToArray());
            foreach(DataRow dr in result_dt.Rows)
            {
                foreach (DataColumn dc in result_dt.Columns)
                {
                    string sel_data = myCode.checkDBNull(dr[dc]);
                    if(TDMK_Code.IsNumeric(sel_data))
                    {
                        dr[dc] = Convert.ToDouble(sel_data) * 1000;
                    }
                }
            }
            return result_dt;
        }
        public string checkNumeric(string text)
        {
            string _result;// = text;
            double test;
            if (double.TryParse(text, out test))
            {
                _result = test.ToString();
            }
            else
            {
                _result = text;
            }
            return _result;
        }
        public DataTable Save_LogFile(SqlConnection sqlcon_OK2SHIP,  DataTable logfile_tbl, string ItemCode, string LotNo, string process_name, string cycle_name)
        {
            start_lbl: string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo"}, new string[] { ItemCode, LotNo });
            DataTable temp_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
            if (cycle_name == "")
            {
                cycle_name = "Data";
            }
            if (temp_dt.Rows.Count > 0)
            {
                if (temp_dt.Rows.Count == logfile_tbl.Rows.Count * logfile_tbl.Columns.Count)
                {
                    int r_inx = 0;
                    foreach (DataColumn dc in logfile_tbl.Columns)
                    {
                        foreach (DataRow dr in logfile_tbl.Rows)
                        {
                            if (r_inx < temp_dt.Rows.Count)
                            {
                                temp_dt.Rows[r_inx][cycle_name] = dr[dc].ToString();
                            }
                            r_inx++;
                        }
                    }
                    TDMK_Code.Delelte_FilteredItem_arr(process_name, sqlcon_OK2SHIP, filter_str);
                    BatchBulkCopy(sqlcon_OK2SHIP, temp_dt, process_name);
                }
                else
                {
                    TDMK_Code.Delelte_FilteredItem_arr(process_name, sqlcon_OK2SHIP, filter_str);
                    goto start_lbl;
                }

            }
            else
            {
                int id = TDMK_Code.SQL_MAX(process_name, "ID", sqlcon_OK2SHIP);
                int col_inx = 0;
                foreach (DataColumn dc in logfile_tbl.Columns)
                {
                    int r_inx = 0;
                    foreach (DataRow dr in logfile_tbl.Rows)
                    {
                        id++;
                        TDMK_Code.insert_val_arr2(process_name, sqlcon_OK2SHIP, new string[] { "ID", "ItemCode", "LotNo", "Net_No", "Pcs_No", cycle_name }, new string[] { id.ToString(), ItemCode, LotNo, (r_inx + 1).ToString(), (col_inx + 1).ToString(), dr[dc].ToString() });
                        r_inx++;
                    }
                    col_inx++;
                }
            }
            return TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
        }
        public DataTable Save_LogFile_detail(SqlConnection sqlcon_OK2SHIP, DataTable logfile_tbl, string ItemCode, string LotNo, string process_name, string cycle_name, string Logfile_name)
        {
            process_name = process_name + "_LOGFILE";
            start_lbl: string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Logfile","Cycles_name" }, new string[] { ItemCode, LotNo, Logfile_name, cycle_name });
            DataTable temp_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
            if (temp_dt.Rows.Count > 0)
            {
                if(MessageBox.Show("Dữ liệu của logfile: "+Logfile_name+" đã có. Bạn muốn cập nhật lại?","Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    if (temp_dt.Rows.Count == logfile_tbl.Rows.Count * logfile_tbl.Columns.Count)
                    {
                        int r_inx = 0;
                        foreach (DataColumn dc in logfile_tbl.Columns)
                        {
                            foreach (DataRow dr in logfile_tbl.Rows)
                            {
                                if (r_inx < temp_dt.Rows.Count)
                                {
                                    temp_dt.Rows[r_inx]["Data"] = dr[dc].ToString();
                                }
                                r_inx++;
                            }
                        }
                        TDMK_Code.Delelte_FilteredItem_arr(process_name, sqlcon_OK2SHIP, filter_str);
                        BatchBulkCopy(sqlcon_OK2SHIP, temp_dt, process_name);
                    }
                    else
                    {
                        TDMK_Code.Delelte_FilteredItem_arr(process_name, sqlcon_OK2SHIP, filter_str);
                        goto start_lbl;
                    }
                }
            }
            else
            {
                int id = TDMK_Code.SQL_MAX(process_name, "ID", sqlcon_OK2SHIP);
                int col_inx = 0;
                string[] col_name_arr = new string[] { "ID", "ItemCode", "LotNo", "Net_No", "Pcs_No", "Logfile", "Cycles_name", "Data" };
                int d_row = 0;
                foreach (DataColumn dc in logfile_tbl.Columns)
                {
                    int r_inx = 0;
                    foreach (DataRow dr in logfile_tbl.Rows)
                    {
                        id++;
                        string[] col_val_arr = new string[] { id.ToString(), ItemCode, LotNo, (r_inx + 1).ToString(), (col_inx + 1).ToString(), Logfile_name, cycle_name, dr[dc].ToString() };
                        temp_dt.Rows.Add();
                        for(int i=0;i<col_name_arr.Length;i++)
                        {
                            temp_dt.Rows[d_row][col_name_arr[i]] = col_val_arr[i];
                        }
                        //TDMK_Code.insert_val_arr2(process_name, sqlcon_OK2SHIP, new string[] { "ID", "ItemCode", "LotNo", "Net_No", "Pcs_No", "Logfile", "Cycles_name", "Data" }, new string[] { id.ToString(), ItemCode, LotNo, (r_inx + 1).ToString(), (col_inx + 1).ToString(),Logfile_name, cycle_name,dr[dc].ToString() });
                        r_inx++;
                        d_row++;
                    }
                    col_inx++;
                }
                BatchBulkCopy(sqlcon_OK2SHIP, temp_dt, process_name);
            }
            return TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
        }
        public DataTable Load_Cycles_Data(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string cycle_name)
        {
            DataTable dgv_dt = new DataTable();
            List<DataTable> myLst_tbl = new List<DataTable>();
            DataSet DS = new DataSet();
            if (cycle_name == "")
            {
                cycle_name = "Data";
            }
            TDMK_Code.fill_dataset_Filter_arr(DS, sqlcon_OK2SHIP, process_name, new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            DataTable dt = DS.Tables[0];
            Get_ListTable(-1, dt, new string[] { "ItemCode", "LotNo", "Net_No" }, ref myLst_tbl, cycle_name);
            foreach (DataTable temp_dt in myLst_tbl)
            {
                List<string> temp_data = new List<string>();
                Get_List_data(-1, temp_dt, new string[] { "ItemCode", "LotNo" }, ref temp_data, cycle_name, false);
                if (dgv_dt.Columns.Count < temp_data.Count)
                {
                    for (int i = 0; i < temp_data.Count; i++)
                    {
                        string col_name = "ItemNo_" + (i + 1).ToString();
                        if (!myCode.check_columns_existed(dgv_dt, col_name))
                        {
                            dgv_dt.Columns.Add(col_name);
                        }
                    }
                }
                DataRow dr = dgv_dt.NewRow();
                int col_inx = 0;
                foreach (string col in temp_data)
                {
                    dr[col_inx] = col;
                    col_inx++;
                }
                dgv_dt.Rows.Add(dr);
            }
            return dgv_dt;
        }
        public DataTable Load_Log_Data(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string cycle_name, string logfile_name)
        {
            DataTable dgv_dt = new DataTable();
            List<DataTable> myLst_tbl = new List<DataTable>();
            DataSet DS = new DataSet();
            process_name = process_name + "_LOGFILE";
            if(cycle_name =="")
            {
                TDMK_Code.fill_dataset_Filter_arr(DS, sqlcon_OK2SHIP, process_name, new string[] { "ItemCode", "LotNo","Logfile" }, new string[] { ItemCode, LotNo, logfile_name });
            }
            else
            {
                TDMK_Code.fill_dataset_Filter_arr(DS, sqlcon_OK2SHIP, process_name, new string[] { "ItemCode", "LotNo", "Cycles_name", "Logfile" }, new string[] { ItemCode, LotNo, cycle_name, logfile_name });
            }
            
            DataTable dt = DS.Tables[0];
            Get_ListTable(-1, dt, new string[] { "ItemCode", "LotNo", "Net_No" }, ref myLst_tbl, "Data");
            foreach (DataTable temp_dt in myLst_tbl)
            {
                List<string> temp_data = new List<string>();
                Get_List_data(-1, temp_dt, new string[] { "ItemCode", "LotNo" }, ref temp_data, "Data", false);
                if (dgv_dt.Columns.Count < temp_data.Count)
                {
                    for (int i = 0; i < temp_data.Count; i++)
                    {
                        string col_name = "ItemNo_" + (i + 1).ToString();
                        if (!myCode.check_columns_existed(dgv_dt, col_name))
                        {
                            dgv_dt.Columns.Add(col_name);
                        }
                    }
                }
                DataRow dr = dgv_dt.NewRow();
                int col_inx = 0;
                foreach (string col in temp_data)
                {
                    dr[col_inx] = col;
                    col_inx++;
                }
                dgv_dt.Rows.Add(dr);
            }
            return dgv_dt;
        }
        public DataTable Load_Log_Data_fromTable(DataTable src_dt, string logfile_name)
        {
            DataTable dgv_dt = new DataTable();
            List<DataTable> myLst_tbl = new List<DataTable>();
            DataTable dt = src_dt.AsEnumerable().Where(x=>x.Field<string>("Logfile")==logfile_name).CopyToDataTable();
            Get_ListTable(-1, dt, new string[] { "ItemCode", "LotNo", "Net_No" }, ref myLst_tbl, "Data");
            foreach (DataTable temp_dt in myLst_tbl)
            {
                List<string> temp_data = new List<string>();
                Get_List_data(-1, temp_dt, new string[] { "ItemCode", "LotNo" }, ref temp_data, "Data", false);
                if (dgv_dt.Columns.Count < temp_data.Count)
                {
                    for (int i = 0; i < temp_data.Count; i++)
                    {
                        string col_name = "ItemNo_" + (i + 1).ToString();
                        if (!myCode.check_columns_existed(dgv_dt, col_name))
                        {
                            dgv_dt.Columns.Add(col_name);
                        }
                    }
                }
                DataRow dr = dgv_dt.NewRow();
                int col_inx = 0;
                foreach (string col in temp_data)
                {
                    dr[col_inx] = col;
                    col_inx++;
                }
                dgv_dt.Rows.Add(dr);
            }
            return dgv_dt;
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
        public void BatchBulkCopy(SqlConnection sqlcon_OK2SHIP, DataTable dataTable, string DestinationTbl)
        {
            DataTable dtInsertRows = dataTable;
            using (SqlBulkCopy sbc = new SqlBulkCopy(sqlcon_OK2SHIP))
            {
                if (sqlcon_OK2SHIP.State != ConnectionState.Open)
                {
                    sqlcon_OK2SHIP.Open();
                }
                sbc.DestinationTableName = DestinationTbl;
                foreach (DataColumn dc in dataTable.Columns)
                {
                    sbc.ColumnMappings.Add(dc.ColumnName, dc.ColumnName);
                }
                sbc.WriteToServer(dtInsertRows);
                sqlcon_OK2SHIP.Close();
            }
        }
        public void Calcul_CPK(Double[] data_arr, ref SEI_Lib.Calculate_CPK tar_CPK_Result, ref SEI_Lib.FAI_Histogram_Data tar_Histogram_Data, string USL, string LSL, int sigma, int count)
        {

            double UL = Convert.ToDouble(USL);
            double LL = Convert.ToDouble(LSL);
            double CPKL = 0;
            double CPKU = 0;

            double stdev = Calcu_process.CalculateStandardDeviation(data_arr);// tg.StDev(data_arr);
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
            List<SortedDictionary<uint, double>> myData = new List<SortedDictionary<uint, double>>();
            List<SortedDictionary<uint, double>> myResult = new List<SortedDictionary<uint, double>>();
            for (int i = 0; i < count + 1; i++)
            {
                bin_data[i] = bin_start + i * bin_step;
            }
            myData = data_arr.Group_data_Freq(count, bin_start, bin_end, ref freq_bin_data);
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
                    myResult.Add(myData[i]);
                }
            }
            tar_CPK_Result.max = max;
            tar_CPK_Result.min = min;
            tar_CPK_Result.stdev = stdev;
            tar_CPK_Result.mean = mean;
            tar_CPK_Result.CPK = CPK;
            tar_Histogram_Data._Bin_data = bin_data;
            tar_Histogram_Data._Freq_bin_data = freq_bin_data;
            tar_Histogram_Data._Modified_NormDist_data = modified_NormDist;
            tar_Histogram_Data._FAI_Data = data_arr;
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
        //public DataTable Detail_NET_PCS(DataTable src_tbl)
        //{
        //    DataTable dgv_dt = new DataTable();
        //    List<DataTable> myLst_tbl = new List<DataTable>();
        //    int col_count = src_tbl.Columns.Count;
        //    string cycle_name = src_tbl.Columns[col_count - 1].ColumnName;
        //    DataTable dt = src_tbl;
        //    Get_ListTable(-1, dt, new string[] { "ItemCode", "LotNo", "Net_No" }, ref myLst_tbl, cycle_name);
        //    foreach (DataTable temp_dt in myLst_tbl)
        //    {
        //        List<string> temp_data = new List<string>();
        //        Get_List_data(-1, temp_dt, new string[] { "ItemCode", "LotNo" }, ref temp_data, cycle_name, false);
        //        if (dgv_dt.Columns.Count < temp_data.Count)
        //        {
        //            for (int i = 0; i < temp_data.Count; i++)
        //            {
        //                string col_name = "ItemNo_" + (i + 1).ToString();
        //                if (!myCode.check_columns_existed(dgv_dt, col_name))
        //                {
        //                    dgv_dt.Columns.Add(col_name);
        //                }
        //            }
        //        }
        //        DataRow dr = dgv_dt.NewRow();
        //        int col_inx = 0;
        //        foreach (string col in temp_data)
        //        {
        //            dr[col_inx] = col;
        //            col_inx++;
        //        }
        //        dgv_dt.Rows.Add(dr);
        //    }
        //    return dgv_dt;
        //}
        public DataTable Detail_NET_PCS(DataTable src_tbl)
        {
            DataTable dgv_dt = new DataTable();
            List<DataTable> myLst_tbl = new List<DataTable>();
            int col_count = src_tbl.Columns.Count;
            string cycle_name = src_tbl.Columns[col_count - 1].ColumnName;
            DataTable dt = src_tbl;
            Get_ListTable(-1, dt, new string[] { "ItemCode", "LotNo", "Net_No" }, ref myLst_tbl, cycle_name);
            foreach (DataTable temp_dt in myLst_tbl)
            {
                List<string> temp_data = new List<string>();
                Get_List_data(-1, temp_dt, new string[] { "ItemCode", "LotNo" }, ref temp_data, cycle_name, false);
                if (dgv_dt.Columns.Count < temp_data.Count)
                {
                    List<string> col_data = new List<string>();
                    Get_List_data(-1, temp_dt, new string[] { "ItemCode", "LotNo" }, ref col_data, "Pcs_No", false);
                    for (int i = 0; i < col_data.Count; i++)
                    {
                        string col_name = col_data[i];
                        if (!myCode.check_columns_existed(dgv_dt, col_name))
                        {
                            dgv_dt.Columns.Add(col_name);
                        }
                    }
                }
                DataRow dr = dgv_dt.NewRow();
                int col_inx = 0;
                foreach (string col in temp_data)
                {
                    dr[col_inx] = col;
                    col_inx++;
                }
                dgv_dt.Rows.Add(dr);
            }
            return dgv_dt;
        }
        public List<NG_list> Get_NG_point(DataTable src_dt)
        {
            List<NG_list> _result = new List<NG_list>();
            foreach (DataColumn dc in src_dt.Columns)
            {
                int r_inx = 0;
                foreach (DataRow dr in src_dt.Rows)
                {
                    if (!myCode.IsNumeric(dr[dc].ToString()))
                    {
                        NG_list sel_NG = new NG_list(r_inx, dc.ColumnName);
                        _result.Add(sel_NG);
                    }
                    r_inx++;
                }
            }
            return _result;
        }
        public List<NG_list> Get_NG_point2(DataGridView dgv_spec, DataGridView dgv_data)
        {
            List<NG_list> _result = new List<NG_list>();
            //foreach (DataColumn dc in src_dt.Columns)
            //{
            //    int r_inx = 0;
            //    foreach (DataRow dr in src_dt.Rows)
            //    {
            //        if (!myCode.IsNumeric(dr[dc].ToString()))
            //        {
            //            NG_list sel_NG = new NG_list(r_inx, dc.ColumnName);
            //            _result.Add(sel_NG);
            //        }
            //        r_inx++;
            //    }
            //}
            foreach (DataGridViewColumn dc in dgv_data.Columns) //foreach (DataGridViewRow dr in dgv_data.Rows)
            {
                foreach (DataGridViewRow dr in dgv_data.Rows) //foreach (DataGridViewColumn dc in dgv_data.Columns)
                {
                    string USL = dgv_spec.Rows[dr.Index].Cells[3].Value.ToString();
                    string LSL = dgv_spec.Rows[dr.Index].Cells[2].Value.ToString();
                    DataGridViewCell sel_cell = dr.Cells[dc.Index];
                    System.Drawing.Color Result_Color = myCode.check_in_limit2(USL, LSL, sel_cell.Value.ToString());
                    sel_cell.Style.BackColor = Result_Color;
                    if(Result_Color!=System.Drawing.Color.White)
                    {

                        NG_list sel_NG = new NG_list(dr.Index, dc.Name);
                        _result.Add(sel_NG);
                    }
                }
            }
            return _result;
        }
        public List<NG_list2> Get_NG_point3(DataGridView dgv_spec, DataGridView dgv_data)
        {
            List<NG_list2> _result = new List<NG_list2>();
            foreach (DataGridViewColumn dc in dgv_data.Columns)
            {
                NG_list2 _myNGlst = new NG_list2(new List<int>(), dc.Name);
                foreach (DataGridViewRow dr in dgv_data.Rows)
                {
                    string USL = dgv_spec.Rows[dr.Index].Cells[3].Value.ToString();
                    string LSL = dgv_spec.Rows[dr.Index].Cells[2].Value.ToString();
                    DataGridViewCell sel_cell = dr.Cells[dc.Index];
                    System.Drawing.Color Result_Color = myCode.check_in_limit2(USL, LSL, sel_cell.Value.ToString());
                    sel_cell.Style.BackColor = Result_Color;
                    if (Result_Color != System.Drawing.Color.White)
                    {
                        _myNGlst.lst_row_inx.Add(dr.Index);
                        sel_cell.Style.ForeColor = System.Drawing.Color.Red;
                    }
                }
                if (_myNGlst.lst_row_inx.Count > 0)
                {
                    _result.Add(_myNGlst);
                }
            }
            return _result;
        }
        public List<NG_list2> Get_NG_point_tbl(DataTable dgv_spec, DataTable dgv_data)
        {
            List<NG_list2> _result = new List<NG_list2>();
            foreach (DataColumn dc in dgv_data.Columns)
            {
                NG_list2 _myNGlst = new NG_list2(new List<int>(), dc.ColumnName);
                int r_inx = 0;
                foreach (DataRow dr in dgv_data.Rows)
                {
                    string USL = dgv_spec.Rows[r_inx][3].ToString();
                    string LSL = dgv_spec.Rows[r_inx][2].ToString();
                    object sel_cell = dr[dc];
                    System.Drawing.Color Result_Color = myCode.check_in_limit2(USL, LSL, sel_cell.ToString());                    
                    if (Result_Color != System.Drawing.Color.White)
                    {
                        _myNGlst.lst_row_inx.Add(r_inx);
                    }
                    r_inx++;
                }
                if (_myNGlst.lst_row_inx.Count > 0)
                {
                    _result.Add(_myNGlst);
                }
            }
            return _result;
        }
        public DataTable NG_NET_detail(DataTable Spec_tbl, NG_list2 NG_point)
        {
            DataTable _result = new DataTable();
            _result = Spec_tbl.Copy();
            _result.Rows.Clear();
            foreach(int r_inx in NG_point.lst_row_inx)
            {
                DataRow src_dr = Spec_tbl.Rows[r_inx];
                DataRow dr = _result.NewRow();
                for(int i=0;i< _result.Columns.Count;i++)
                {
                    dr[i] = src_dr[i];
                }
                _result.Rows.Add(dr);
            }
            return _result;
        }
        public DataTable Select_table_IndexList(DataTable Spec_tbl, List<int> NG_point)
        {
            DataTable _result = new DataTable();
            _result = Spec_tbl.Copy();
            _result.Rows.Clear();
            foreach (int r_inx in NG_point)
            {
                DataRow src_dr = Spec_tbl.Rows[r_inx];
                DataRow dr = _result.NewRow();
                for (int i = 0; i < _result.Columns.Count; i++)
                {
                    dr[i] = src_dr[i];
                }
                _result.Rows.Add(dr);
            }
            return _result;
        }
        public void Summary_NG_list(List<ECheck_Process.NG_list2> templst, ref List<string> tar_item_NG, ref DataTable NG_NET_Table, DataTable Spec_tbl, ref List<int> index_lst)
        {          
            if(Spec_tbl!=null)
            {
                foreach (ECheck_Process.NG_list2 item in templst)
                {
                    string col_name = item.col_val;
                    index_lst.AddRange(item.lst_row_inx);
                    tar_item_NG.Add(col_name);
                }
                int[] temp = index_lst.Distinct().ToArray();
                Array.Sort(temp);
                index_lst = temp.ToList();
                NG_NET_Table = Select_table_IndexList(Spec_tbl, index_lst);
            }
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
        public DataTable load_spec_from_Logfile(DataTable src_Spec_tbl, string ItemCode, string Process_name, SqlConnection sqlcon, bool overwrite_en)
        {
start_lbl: string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, Process_name });
            DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "NET_SPEC", filter_str);
            if (spec_dt.Rows.Count > 0) 
            {
                if ((src_Spec_tbl.Rows.Count > 0)&& (overwrite_en))
                {
                    if (MessageBox.Show("Do you want to update NET Spec table?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        TDMK_Code.Delelte_FilteredItem_arr("NET_SPEC", sqlcon, filter_str);
                        goto start_lbl;
                    }
                }
            }
            else
            {
                int inx = TDMK_Code.SQL_MAX("NET_SPEC", "ID", sqlcon) + 1;
                foreach (DataRow dr in src_Spec_tbl.Rows)
                {
                    List<string> item_val = new List<string>();
                    string ID = inx.ToString();
                    item_val.Add(ID);
                    item_val.Add(ItemCode);
                    foreach (DataColumn dc in src_Spec_tbl.Columns)
                    {
                        item_val.Add(dr[dc].ToString());
                    }
                    item_val.Add(Process_name);
                    TDMK_Code.insert_val_arr2("NET_SPEC", sqlcon, new string[] { "ID", "ItemCode", "[Point+V]", "[Point-V]", "LSL", "USL", "Remark" }, item_val.ToArray());
                    inx++;
                }
                spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "NET_SPEC", filter_str);
            }
            return spec_dt.AsDataView().ToTable(false, new string[] { "Point+V", "Point-V", "LSL", "USL" });
        }
        public DataTable Summary_Selected(DataTable src_Sum,List<NG_list2> src_NG, int inQty)
        {
            DataTable _result = new DataTable();
            int qty = 0;
            List<string> sel_item_lst = new List<string>();
            if (src_NG.Count != 0)
            {
                qty = Math.Min(inQty, src_NG.Count);
                List<int> A = new List<int>();
                foreach (ECheck_Process.NG_list2 lst in src_NG)
                {
                    A.Add(lst.lst_row_inx.Count);
                }
                var sorted = A
                    .Select((x, i) => new KeyValuePair<int, int>(x, i))
                    .OrderBy(x => x.Key)
                    .ToList();
                List<int> idx = sorted.Select(x => x.Value).ToList();
                for (int i = 0; i < qty; i++)
                {
                    int cur_inx = idx[i];
                    sel_item_lst.Add(src_NG[cur_inx].col_val);
                }
                
            }
            else
            {
                qty = inQty;
                for (int i = 0; i < qty; i++)
                {
                    sel_item_lst.Add(src_Sum.Columns[i].ColumnName);
                }
            }
            _result = src_Sum.AsDataView().ToTable(false, sel_item_lst.ToArray());
            return _result;
        }
        public DataTable Summary_Selected_FromExisted(DataTable src_Sum, List<NG_list2> src_NG, int inQty)
        {
            DataTable _result = new DataTable();
            int qty = 0;
            List<string> sel_item_lst = new List<string>();
            List<string> ignored_lst = new List<string>();
            foreach (ECheck_Process.NG_list2 lst in src_NG)
            {
                ignored_lst.Add(lst.col_val);
            }
            int item_OK = src_Sum.Columns.Count - ignored_lst.Count;
            if(item_OK>0)
            {
                int sel_Qty = Math.Min(item_OK, inQty);
                int inx = 0;
                for (int i = 0; i < src_Sum.Columns.Count; i++)
                {
                    string cur_col = src_Sum.Columns[i].ColumnName;
                    if (TDMK_Code.check_exist_list_index(cur_col, ignored_lst) == -1)
                    {
                        sel_item_lst.Add(cur_col);
                        inx++;
                        if (inx == sel_Qty)
                        {
                            break;
                        }
                    }
                }
                if(sel_Qty< inQty)
                {
                    int sel_NG_qty = inQty - sel_Qty;
                    if (src_NG.Count != 0)
                    {
                        qty = Math.Min(sel_NG_qty, src_NG.Count);
                        List<int> A = new List<int>();
                        foreach (ECheck_Process.NG_list2 lst in src_NG)
                        {
                            A.Add(lst.lst_row_inx.Count);
                        }
                        var sorted = A
                            .Select((x, i) => new KeyValuePair<int, int>(x, i))
                            .OrderBy(x => x.Key)
                            .ToList();
                        List<int> idx = sorted.Select(x => x.Value).ToList();
                        for (int i = 0; i < qty; i++)
                        {
                            int cur_inx = idx[i];
                            sel_item_lst.Add(src_NG[cur_inx].col_val);
                        }

                    }
                    else
                    {
                        qty = sel_NG_qty;
                        for (int i = 0; i < qty; i++)
                        {
                            sel_item_lst.Add(src_Sum.Columns[i].ColumnName);
                        }
                    }
                }
            }
            _result = src_Sum.AsDataView().ToTable(false, sel_item_lst.ToArray());
            return _result;
        }
        public void Submit_Data(DataGridView src_DGV_spec, DataGridView src_DGV_data, ref DataTable tbl_NG_NET, ref List<int> NET_index_lst, DataGridView tar_DGV_Net, ref List<string> item_NG_lst, ref List<NG_list2> _myNGlst2)
        {
            _myNGlst2 = Get_NG_point3(src_DGV_spec, src_DGV_data);
            //List<string> item_NG_lst = new List<string>();
            Summary_NG_list(_myNGlst2, ref item_NG_lst, ref tbl_NG_NET, (DataTable)src_DGV_spec.DataSource, ref NET_index_lst);
            tar_DGV_Net.DataSource = tbl_NG_NET;
            int r_inx1 = 0;
            foreach (int inx in NET_index_lst)
            {
                tar_DGV_Net.Rows[r_inx1].HeaderCell.Value = (inx + 1).ToString();
                r_inx1++;
            }
            tar_DGV_Net.AutoResizeColumns();
            tar_DGV_Net.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }
        public void Submit_Data_2(DataTable src_DGV_spec, DataTable src_DGV_data, ref DataTable tbl_NG_NET, ref List<int> NET_index_lst, DataGridView tar_DGV_Net, ref List<string> item_NG_lst, ref List<NG_list2> _myNGlst2)
        {
            _myNGlst2 = Get_NG_point_tbl(src_DGV_spec, src_DGV_data);
            Summary_NG_list(_myNGlst2, ref item_NG_lst, ref tbl_NG_NET, src_DGV_spec, ref NET_index_lst);
            tar_DGV_Net.DataSource = tbl_NG_NET;
            int r_inx1 = 0;
            foreach (int inx in NET_index_lst)
            {
                tar_DGV_Net.Rows[r_inx1].HeaderCell.Value = (inx + 1).ToString();
                r_inx1++;
            }
            tar_DGV_Net.AutoResizeColumns();
            tar_DGV_Net.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }
        public DataTable Save_Submit_data(SqlConnection sqlcon_OK2SHIP, DataTable submit_tbl, string ItemCode, string LotNo, string process_name, string cycle_name)
        {
            start_lbl: string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            DataTable temp_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
            if (cycle_name == "")
            {
                cycle_name = "Data";
            }
            if (temp_dt.Rows.Count > 0)
            {
                List<string> cycle_data_lst = new List<string>();
                Get_List_data2(-1, temp_dt, new string[] { "ItemCode", "LotNo" }, ref cycle_data_lst, cycle_name, false);
                if (cycle_data_lst.Count>0)
                {
                    if(MessageBox.Show("Dữ liệu đã được lưu.Bạn muốn ghi đè?", "Warning",MessageBoxButtons.YesNo)==DialogResult.Yes)
                    {
                        goto next_act;
                    }
                    else
                    {
                        return temp_dt;
                    }    
                }
    next_act:  if(temp_dt.Rows.Count == submit_tbl.Rows.Count * submit_tbl.Columns.Count)
                {
                    int r_inx = 0;
                    foreach (DataColumn dc in submit_tbl.Columns)
                    {
                        foreach (DataRow dr in submit_tbl.Rows)
                        {
                            if (r_inx < temp_dt.Rows.Count)
                            {
                                temp_dt.Rows[r_inx][cycle_name] = dr[dc].ToString();
                            }
                            r_inx++;
                        }
                    }
                    TDMK_Code.Delelte_FilteredItem_arr(process_name, sqlcon_OK2SHIP, filter_str);
                    BatchBulkCopy(sqlcon_OK2SHIP, temp_dt, process_name);
                }
                else
                {
                    TDMK_Code.Delelte_FilteredItem_arr(process_name, sqlcon_OK2SHIP, filter_str);
                    goto start_lbl;
                }
                //else
                //{

                //}
            }
            else
            {
                bool save_en = false;
                if ((cycle_name=="Data")||(cycle_name=="Before"))
                {
                    save_en = true;
                }
                else
                {
                    List<string> before_data_lst = new List<string>();
                    Get_List_data2(-1, temp_dt, new string[] { "ItemCode", "LotNo" }, ref before_data_lst, "Before", false);
                    if (before_data_lst.Count > 0)
                    {
                        save_en = true;
                    }
                    else
                    {
                        save_en = false;
                        MessageBox.Show("Before data not existed!\r\nCannot save", "Warning");                        
                    }
                }
                if(save_en)
                {
                    int id = TDMK_Code.SQL_MAX(process_name, "ID", sqlcon_OK2SHIP);
                    int col_inx = 0;
                    foreach (DataColumn dc in submit_tbl.Columns)
                    {
                        int r_inx = 0;
                        string pcs_no = dc.ColumnName;
                        foreach (DataRow dr in submit_tbl.Rows)
                        {
                            id++;
                            TDMK_Code.insert_val_arr2(process_name, sqlcon_OK2SHIP, new string[] { "ID", "ItemCode", "LotNo", "Net_No", "Pcs_No", cycle_name }, new string[] { id.ToString(), ItemCode, LotNo, (r_inx + 1).ToString(), pcs_no, dr[dc].ToString() });
                            r_inx++;
                        }
                        col_inx++;
                    }
                }
            }
            return TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
        }
        public DataTable HotOil_process(string logfile)
        {
            DataTable _result = new DataTable();
            myExcel.Workbook tar_wrkbk = TDMK_Code.open_excel_file(logfile, "", "");
            myExcel.Worksheet tar_wrksht = tar_wrkbk.Sheets["500 CHUKY"];
            myExcel.Range tar_rgn = tar_wrksht.Range["D11"];
            myExcel.Range data_rgn = tar_wrksht.Range["C17"];
            int col_inx = 0;
            while (!data_rgn.Offset[0,col_inx].HasFormula)
            {
                col_inx++;
            }
            for(int i = 0;i< col_inx;i++)
            {
                if(i==0)
                {
                    _result.Columns.Add("Before");
                }
                else
                {
                    string _col = "After_" + tar_wrksht.Range["C12"].Offset[0, i].Value;
                    _result.Columns.Add(_col);
                }
                int r_inx = 0;
                
                while(myCode.checkDBNull(data_rgn.Offset[r_inx,i].Value)!="")
                {
                    string cell_val = myCode.checkDBNull(data_rgn.Offset[r_inx, i].Value);
                    if (_result.Rows.Count<r_inx+1)
                    {
                        _result.Rows.Add();
                    }
                    _result.Rows[r_inx][i] = cell_val;
                    
                    r_inx++;
                }
            }           
            return _result;
        }
        public bool Check_HotOil_data(DataGridView src_dgv)
        {
            bool _result = true;
            foreach (DataGridViewRow dgv_r in src_dgv.Rows)
            {
                double before_val = Convert.ToDouble(dgv_r.Cells[0].Value);
                for (int i = 1; i < src_dgv.Columns.Count; i++)
                {
                    double after_val = Convert.ToDouble(dgv_r.Cells[i].Value);
                    double rate_val = Math.Abs(after_val - before_val) / before_val;
                    if (rate_val > 0.05)
                    {
                        dgv_r.Cells[i].Style.BackColor = Color.Yellow;
                        _result = false;
                    }
                    else
                    {
                        dgv_r.Cells[i].Style.BackColor = Color.White;
                    }
                }
            }
            return _result;
        }
        public DataTable Save_LogFile_HotOil(SqlConnection sqlcon_OK2SHIP, DataTable logfile_tbl, string ItemCode, string LotNo, string process_name,  string Logfile_name)
        {
            process_name = process_name + "_LOGFILE";
            start_lbl: string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Logfile" }, new string[] { ItemCode, LotNo, Logfile_name });
            DataTable temp_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
            if (temp_dt.Rows.Count > 0)
            {
                if (MessageBox.Show("Data is existed in Database. Do you want to update again?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    if (temp_dt.Rows.Count == logfile_tbl.Rows.Count * logfile_tbl.Columns.Count)
                    {
                        int r_inx = 0;
                        foreach (DataColumn dc in logfile_tbl.Columns)
                        {
                            foreach (DataRow dr in logfile_tbl.Rows)
                            {
                                if (r_inx < temp_dt.Rows.Count)
                                {
                                    temp_dt.Rows[r_inx]["Data"] = dr[dc].ToString();
                                }
                                r_inx++;
                            }
                        }
                        TDMK_Code.Delelte_FilteredItem_arr(process_name, sqlcon_OK2SHIP, filter_str);
                        BatchBulkCopy(sqlcon_OK2SHIP, temp_dt, process_name);
                    }
                    else
                    {
                        TDMK_Code.Delelte_FilteredItem_arr(process_name, sqlcon_OK2SHIP, filter_str);
                        goto start_lbl;
                    }
                }
            }
            else
            {
                int id = TDMK_Code.SQL_MAX(process_name, "ID", sqlcon_OK2SHIP);
                string region = "";
                string f_name = Path.GetFileNameWithoutExtension(Logfile_name);
                string[] temp = f_name.Split('-');
                region = temp[temp.Length - 1];
                int col_inx = 0;
                foreach (DataColumn dc in logfile_tbl.Columns)
                {
                    int r_inx = 0;
                    foreach (DataRow dr in logfile_tbl.Rows)
                    {
                        id++;
                        TDMK_Code.insert_val_arr2(process_name, sqlcon_OK2SHIP, new string[] { "ID", "ItemCode", "LotNo", "Region", "Pcs_No", "Data", "Logfile", "Cycles_name" }, new string[] { (id + 1).ToString(), ItemCode, LotNo, region, (r_inx + 1).ToString(), dr[dc].ToString(), Logfile_name, dc.ColumnName });
                        r_inx++;
                    }
                    col_inx++;
                }
            }
            return TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
        }
        public DataTable Load_Log_Data_HotOil(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string region, string logfile_name)
        {
            DataTable dgv_dt = new DataTable();
            List<DataTable> myLst_tbl = new List<DataTable>();
            DataSet DS = new DataSet();
            process_name = process_name + "_LOGFILE";
            //TDMK_Code.fill_dataset_Filter_arr(DS, sqlcon_OK2SHIP, process_name, new string[] { "ItemCode", "LotNo", "Region", "Logfile" }, new string[] { ItemCode, LotNo, region, logfile_name });
            DataTable dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region", "Logfile" }, new string[] { ItemCode, LotNo, region, logfile_name }));
            Get_ListTable(-1, dt, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref myLst_tbl, "Data");
            foreach (DataTable temp_dt in myLst_tbl)
            {
                List<string> temp_data = new List<string>();
                List<string> temp_col = new List<string>();
                Get_List_data(-1, temp_dt, new string[] { "ItemCode", "LotNo" }, ref temp_data, "Data", false);
                Get_List_data(-1, temp_dt, new string[] { "ItemCode", "LotNo" }, ref temp_col, "Cycles_name", false);
                if (dgv_dt.Columns.Count < temp_data.Count)
                {
                    for (int i = 0; i < temp_col.Count; i++)
                    {
                        string col_name = temp_col[i];
                        if (!myCode.check_columns_existed(dgv_dt, col_name))
                        {
                            dgv_dt.Columns.Add(col_name);
                        }
                    }
                }
                DataRow dr = dgv_dt.NewRow();
                int col_inx = 0;
                foreach (string col in temp_data)
                {
                    dr[col_inx] = col;
                    col_inx++;
                }
                dgv_dt.Rows.Add(dr);
            }
            return dgv_dt;
        }
        public void export_HotOil_fromLog(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name)
        {
            myExcel.Workbook report_wrkbk = TDMK_Code.Create_Report_new(Application.StartupPath, format_name, "", "");
            myExcel.Worksheet tar_sht = report_wrkbk.Sheets[1];
            myExcel.Range BVH_rgn = tar_sht.Range["AH25"];
            myExcel.Range PTH_rgn = tar_sht.Range["B25"];
            myExcel.Range tar_rgn;
            List<string> logfile_lst = new List<string>();
            DataTable dt1 = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name + "_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
            Get_List_data(-1, dt1, new string[] { "ItemCode", "LotNo" }, ref logfile_lst, "Logfile", true);
            foreach (var lst in logfile_lst)
            {
                string region = "";
                string f_name = Path.GetFileNameWithoutExtension(lst);
                string[] temp = f_name.Split('-');
                region = temp[temp.Length - 1];
                if (region == "BVH")
                {
                    tar_rgn = BVH_rgn;
                }
                else
                {
                    tar_rgn = PTH_rgn;
                }
                DataTable dt = Load_Log_Data_HotOil(sqlcon_OK2SHIP, ItemCode, LotNo, process_name, region, lst);
                string sel_col = "";
                foreach (DataColumn dc in dt.Columns)
                {
                    if (dc.ColumnName.Contains("After"))
                    {
                        if (myCode.checkDBNull(dt.Rows[0][dc]) != "")
                        {
                            sel_col = dc.ColumnName;
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                DataTable sel_dt = dt.AsDataView().ToTable(false, new string[] { "Before", sel_col });
                int r_inx = 0;
                foreach (DataRow dr in sel_dt.Rows)
                {
                    int c_inx = 0;
                    foreach (DataColumn dc in sel_dt.Columns)
                    {
                        tar_rgn.Offset[c_inx, r_inx].Value = dr[dc].ToString();
                        c_inx++;
                    }
                    r_inx++;
                }
            }
        }
        public void Save_HotOil(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            DataTable tar_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
            if(tar_tbl.Rows.Count>0)
            {
                MessageBox.Show("Data existed!","Warning");
            }
            else
            {
                List<string> logfile_lst = new List<string>();
                DataTable dt1 = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name + "_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                Get_List_data(-1, dt1, new string[] { "ItemCode", "LotNo" }, ref logfile_lst, "Logfile", true);
                foreach (var lst in logfile_lst)
                {
                    string region = "";
                    string f_name = Path.GetFileNameWithoutExtension(lst);
                    string[] temp = f_name.Split('-');
                    region = temp[temp.Length - 1];
                    DataTable dt = Load_Log_Data_HotOil(sqlcon_OK2SHIP, ItemCode, LotNo, process_name, region, lst);
                    string sel_col = "";
                    foreach (DataColumn dc in dt.Columns)
                    {
                        if (dc.ColumnName.Contains("After"))
                        {
                            if (myCode.checkDBNull(dt.Rows[0][dc]) != "")
                            {
                                sel_col = dc.ColumnName;
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                    DataTable sel_dt = dt.AsDataView().ToTable(false, new string[] { "Before", sel_col });
                    int id = TDMK_Code.SQL_MAX(process_name, "ID", sqlcon_OK2SHIP);
                    int r_inx = 0;
                    foreach (DataRow dr in sel_dt.Rows)
                    {
                        TDMK_Code.insert_val_arr2(process_name, sqlcon_OK2SHIP, new string[] { "ID", "ItemCode", "LotNo", "Pcs_No", "Region", "Before", "After" }, new string[] { (id + 1).ToString(), ItemCode, LotNo, (r_inx + 1).ToString(), region, dr["Before"].ToString(), dr[sel_col].ToString() });
                        id++;
                        r_inx++;
                    }
                }
                MessageBox.Show("Completed!", "Warning");
            }

        }
        public void Export_HotOil(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name)
        {
            string format_file = exp_proc.find_format(myVar.data_loc,ItemCode);
            if(format_file!="")
            {
                myExcel.Workbook report_wrkbk = exp_proc.create_export_wrk(format_file, format_name, ItemCode, LotNo);
                //myExcel.Workbook report_wrkbk = TDMK_Code.Create_Report_new(Application.StartupPath, format_name, "", ItemCode+"-"+LotNo);
                myExcel.Worksheet tar_sht = report_wrkbk.Sheets[1];
                //myExcel.Range BVH_rgn = tar_sht.Range["AH25"];
                //myExcel.Range PTH_rgn = tar_sht.Range["B25"];

                string data_addr = Get_start_range("Sample no.", "A10", tar_sht);
                myExcel.Range PTH_rgn = tar_sht.Range[data_addr].Offset[0,1];
                myExcel.Range BVH_rgn = tar_sht.Range[data_addr].Offset[0, 33];

                myExcel.Range tar_rgn;
                List<DataTable> src_tbl_lst = new List<DataTable>();
                DataTable dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
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
                                tar_rgn.Offset[c_inx, r_inx].Value = dr[dc].ToString();
                                c_inx++;
                            }
                            r_inx++;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Format of ItemCode: " + ItemCode + " not found!", "Warning");
            }

        }
        public void Get_Pair_data(DataTable src_tbl, ref List<string> before_data, ref List<string> after_data)
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
            if (data_col.Count >= 2)
            {
                before_data = data_col.ElementAt(0).Value;
                after_data = data_col.ElementAt(data_col.Count - 1).Value;
            }
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
            foreach(var t in sel_net.Keys.ToList())
            {
                before_data.Add(_before_data[t]);
                after_data.Add(_after_data[t]);
            }
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
        public Dictionary<string, List<string>> Get_List_Pair_data_bending(DataTable src_tbl)
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
            data_col.Remove("After");
            return data_col;
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
        public void Export_Reflow(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
            List<DataTable> src_tbl_lst = new List<DataTable>();
            Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
            myExcel.Workbook report_wrk = TDMK_Code.Create_Report_new(Application.StartupPath, format_name, "", ItemCode+"-"+LotNo);
            myExcel.Worksheet tar_wrksht = report_wrk.Sheets[1];
            myExcel.Range tar_rgn = tar_wrksht.Range["C70"];
            myExcel.Range sum_rgn = tar_wrksht.Range["C30"];
            int col_inx = 0;
            foreach (var tbl in src_tbl_lst)
            {
                List<string> before_data = new List<string>();
                List<string> after_data = new List<string>();
                Get_Pair_data(tbl, ref before_data, ref after_data);
                sum_rgn.Offset[col_inx, 0].Value = Calcu_process.ConvertToDouble(before_data).Max();
                sum_rgn.Offset[col_inx, 1].Value = Calcu_process.ConvertToDouble(after_data).Max();
                for (int r_inx = 0; r_inx < before_data.Count; r_inx++)
                {
                    tar_rgn.Offset[r_inx, 6 * col_inx].Value = before_data[r_inx];
                    tar_rgn.Offset[r_inx, 6 * col_inx + 1].Value = after_data[r_inx];
                }
                col_inx++;
            }
        }
        public void Export_Reflow_new(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name)
        {

            //myExcel.Workbook report_wrk = TDMK_Code.Create_Report_new(Application.StartupPath, format_name, "", ItemCode + "-" + LotNo);
            string format_file = exp_proc.find_format(myVar.data_loc, ItemCode);

            if (format_file != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                string spec_filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process_name });
                DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", spec_filter_str);
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                myExcel.Workbook report_wrk = exp_proc.create_export_wrk(format_file, format_name, ItemCode, LotNo);
                myExcel.Worksheet tar_wrksht = report_wrk.Sheets[1];
                myExcel.Range tar_rgn = tar_wrksht.Range["C70"];
                myExcel.Range sum_rgn = tar_wrksht.Range["C30"];
                int col_inx = 0;
                foreach (var tbl in src_tbl_lst)
                {
                    List<string> before_data = new List<string>();
                    List<string> after_data = new List<string>();
                    Get_Pair_data_new(spec_dt, tbl, ref before_data, ref after_data);
                    sum_rgn.Offset[col_inx, 0].Value = Calcu_process.ConvertToDouble(before_data).Max();
                    sum_rgn.Offset[col_inx, 1].Value = Calcu_process.ConvertToDouble(after_data).Max();
                    for (int r_inx = 0; r_inx < before_data.Count; r_inx++)
                    {
                        tar_rgn.Offset[r_inx, 6 * col_inx].Value = before_data[r_inx];
                        tar_rgn.Offset[r_inx, 6 * col_inx + 1].Value = after_data[r_inx];
                    }
                    col_inx++;
                }
            }
            else
            {
                MessageBox.Show("Format of ItemCode " + ItemCode + " not found", "Warning");
            }

        }
        public string Get_Number_String(string src_str, char split_chr)
        {
            string _result = "";
            string[] temp = src_str.Split(split_chr);
            foreach(var t in temp)
            {
                if(myCode.IsNumeric(t))
                {
                    _result = t;
                    break;
                }
            }
            return _result;
        }
        public DataTable Save_Bending(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string logfile)
        {
            Dictionary<string, List<string>> NET_data_lst = new Dictionary<string, List<string>>();
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            List<string> item_lst = new List<string>() { "ID", "ItemCode", "LotNo", "Net_No", "Pcs_No" };
start_lbl:  DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
            if(src_dt.Rows.Count>0)
            {
                if (MessageBox.Show("Data is existed in Database. Do you want to update again?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    TDMK_Code.Delelte_FilteredItem_arr(process_name, sqlcon_OK2SHIP, filter_str);
                    goto start_lbl;
                }
            }
            else
            {
                myExcel.Workbook report_wrk = TDMK_Code.open_excel_file(logfile, "", "");
                myExcel.Worksheet tar_wrksht = report_wrk.Sheets[1];
                myExcel.Range tar_rgn = tar_wrksht.Range["C35"];
                int col_inx = 0;
                while (myCode.checkDBNull(tar_rgn.Offset[-1, col_inx].Value) != "")
                {
                    string cell_val = myCode.checkDBNull(tar_rgn.Offset[-1, col_inx].Value);
                    if (cell_val.Contains("Before") || (cell_val.Contains("After")))
                    {
                        string _col_name = Get_Number_String(cell_val, ' ');
                        string col_name = "";
                        if (cell_val.Contains("Before"))
                        {
                            col_name = "Before";
                        }
                        else
                        {
                            if (_col_name != "")
                            {
                                col_name = "After_" + _col_name;
                            }
                        }
                        List<string> NET_lst = new List<string>();
                        int r_inx = 0;
                        while (myCode.checkDBNull(tar_rgn.Offset[r_inx, col_inx].Value) != "")
                        {
                            NET_lst.Add(myCode.checkDBNull(tar_rgn.Offset[r_inx, col_inx].Value));
                            r_inx++;
                        }
                        NET_data_lst.Add(col_name, NET_lst);
                    }
                    col_inx++;
                }
                item_lst.AddRange(NET_data_lst.Keys.ToList());
                int row_count = NET_data_lst.ElementAt(0).Value.Count;
                int id = TDMK_Code.SQL_MAX(process_name, "ID", sqlcon_OK2SHIP);
                for (int i = 0; i < row_count; i++)
                {
                    string net = "";
                    string pcs = ((int)(i / 2) + 1).ToString();
                    if ((i % 2) == 0)
                    {
                        net = "A";
                    }
                    else
                    {
                        net = "B";
                    }
                    List<string> item_val_lst = new List<string>() { (id + i + 1).ToString(), ItemCode, LotNo, net, pcs };
                    List<string> temp_lst = new List<string>();
                    foreach (var item in NET_data_lst)
                    {
                        temp_lst.Add(item.Value[i]);
                    }
                    item_val_lst.AddRange(temp_lst.ToList());
                    TDMK_Code.insert_val_arr2(process_name, sqlcon_OK2SHIP, item_lst.ToArray(), item_val_lst.ToArray());
                }
            }
            return TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
        }
        public void Export_Bending(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name)
        {
            string format_file = exp_proc.find_format(myVar.data_loc, ItemCode);
            if(format_file!="")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                List<string> pcs_sel_lst = Get_SelectedPcs_List(src_tbl,"Sel_report","Yes");// src_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_report").ToUpper() == "YES").Select(x => x.Field<string>("Pcs_No")).Distinct().ToList();
                if(pcs_sel_lst.Count!=0)
                {
                    src_tbl = src_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_report").ToUpper() == "YES").CopyToDataTable();
                }
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                myExcel.Workbook report_wrk = exp_proc.create_export_wrk(format_file, format_name, ItemCode, LotNo);
                //myExcel.Workbook report_wrk = TDMK_Code.Create_Report_new(Application.StartupPath, format_name, "", ItemCode+"-"+LotNo);
                myExcel.Worksheet tar_wrksht = report_wrk.Sheets[1];

                //myExcel.Range tar_rgn = tar_wrksht.Range["B34"];
                //myExcel.Range sum_rgn = tar_wrksht.Range["C17"];
                string tar_rgn_addr = Get_start_range("Test item", "A10", tar_wrksht);
                myExcel.Range tar_rgn = tar_wrksht.Range[tar_rgn_addr].Offset[0, 1];
                string sum_rgn_addr = Get_start_range("Condition", "A10", tar_wrksht);
                myExcel.Range sum_rgn = tar_wrksht.Range[sum_rgn_addr].Offset[0,2];
                int col_inx = 0;
                foreach (var tbl in src_tbl_lst)
                {
                    Dictionary<string, List<string>> data_lst = Get_List_Pair_data(tbl);
                    List<string> before_data = data_lst["Before"];
                    int r_offset = 0;
                    foreach (var item in data_lst)
                    {
                        if (item.Key != "Before")
                        {
                            myExcel.Range tar_rgn_offset = tar_rgn.Offset[r_offset * tbl.Rows.Count, 0];
                            myExcel.Range sum_rgn_offset = sum_rgn.Offset[r_offset, 0];
                            List<string> after_data = item.Value;
                            string[] sum_arr = new string[2];
                            for (int r_inx = 0; r_inx < before_data.Count; r_inx++)
                            {
                                int sum_inx = r_inx % 2;
                                tar_rgn_offset.Offset[r_inx, 4 * col_inx].Value = before_data[r_inx];
                                tar_rgn_offset.Offset[r_inx, 4 * col_inx + 1].Value = after_data[r_inx];
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
                                        sum_rgn_offset.Offset[r_inx / 2, col_inx].Value = "OK";
                                    }
                                    else
                                    {
                                        sum_rgn_offset.Offset[r_inx / 2, col_inx].Value = "NG";
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
                MessageBox.Show("Format of ItemCode " + ItemCode + " not found", "Warning");
            }
            
        }
        public DataTable Bending_process(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string logfile)
        {
            Dictionary<string, List<string>> NET_data_lst = new Dictionary<string, List<string>>();
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            List<string> item_lst = new List<string>() { "ID", "ItemCode", "LotNo", "Net_No", "Pcs_No","Logfile" };
            start_lbl: DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
            if (src_dt.Rows.Count > 0)
            {
                if (MessageBox.Show("Data is existed in Database. Do you want to update again?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    TDMK_Code.Delelte_FilteredItem_arr(process_name, sqlcon_OK2SHIP, filter_str);
                    goto start_lbl;
                }
            }
            else
            {
                myExcel.Workbook report_wrk = TDMK_Code.open_excel_file(logfile, "", "");
                myExcel.Worksheet tar_wrksht = report_wrk.Sheets[1];
                myExcel.Range tar_rgn = tar_wrksht.Range["Z35"];
                int col_inx = 0;
                while (myCode.checkDBNull(tar_rgn.Offset[-1, col_inx].Value) != "")
                {
                    string cell_val = myCode.checkDBNull(tar_rgn.Offset[-1, col_inx].Value);
                    if (cell_val.Contains("Before") || (cell_val.Contains("After")))
                    {
                        string _col_name = Get_Number_String(cell_val, ' ');
                        string col_name = "";
                        if (cell_val.Contains("Before"))
                        {
                            col_name = "Before";
                        }
                        else
                        {
                            if (_col_name != "")
                            {
                                col_name = "After_" + _col_name;
                            }
                        }
                        List<string> NET_lst = new List<string>();
                        int r_inx = 0;
                        while (myCode.checkDBNull(tar_rgn.Offset[r_inx, col_inx].Value) != "")
                        {
                            NET_lst.Add(myCode.checkDBNull(tar_rgn.Offset[r_inx, col_inx].Value));
                            r_inx++;
                        }
                        NET_data_lst.Add(col_name, NET_lst);
                    }
                    col_inx++;
                }
                item_lst.AddRange(NET_data_lst.Keys.ToList());
                int row_count = NET_data_lst.ElementAt(0).Value.Count;
                int id = TDMK_Code.SQL_MAX(process_name, "ID", sqlcon_OK2SHIP);
                for (int i = 0; i < row_count; i++)
                {
                    string net = "";
                    string pcs = ((int)(i / 2) + 1).ToString();
                    if ((i % 2) == 0)
                    {
                        net = "A";
                    }
                    else
                    {
                        net = "B";
                    }
                    List<string> item_val_lst = new List<string>() { (id + i + 1).ToString(), ItemCode, LotNo, net, pcs,Path.GetFileName(logfile) };
                    List<string> temp_lst = new List<string>();
                    foreach (var item in NET_data_lst)
                    {
                        temp_lst.Add(item.Value[i]);
                    }
                    item_val_lst.AddRange(temp_lst.ToList());
                    //src_dt.Rows.Add(item_val_lst.ToArray());
                    TDMK_Code.insert_val_arr2(process_name, sqlcon_OK2SHIP, item_lst.ToArray(), item_val_lst.ToArray());
                }
            }
            return TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
        }
        public void Export_Thermal_HeatSoak_Bend_old(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name)
        {
            //myExcel.Workbook report_wrk = TDMK_Code.Create_Report_new(Application.StartupPath, format_name, "", ItemCode + "-" + LotNo);
            string format_file = exp_proc.find_format(myVar.data_loc, ItemCode);
            if (format_file != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                myExcel.Workbook report_wrk = exp_proc.create_export_wrk(format_file, format_name, ItemCode, LotNo);
                myExcel.Worksheet tar_wrksht = report_wrk.Sheets[1];
                myExcel.Range tar_rgn = tar_wrksht.Range["B33"];
                myExcel.Range sum_rgn = tar_wrksht.Range["C17"];
                int col_inx = 0;
                foreach (var tbl in src_tbl_lst)
                {
                    Dictionary<string, List<string>> data_lst = Get_List_Pair_data(tbl);
                    List<string> before_data = data_lst["Before"];
                    int r_offset = 0;
                    foreach (var item in data_lst)
                    {
                        if (item.Key != "Before")
                        {
                            myExcel.Range tar_rgn_offset = tar_rgn.Offset[r_offset * tbl.Rows.Count, 0];
                            myExcel.Range sum_rgn_offset = sum_rgn.Offset[r_offset, 0];
                            List<string> after_data = item.Value;
                            string[] sum_arr = new string[2];
                            for (int r_inx = 0; r_inx < before_data.Count; r_inx++)
                            {
                                int sum_inx = r_inx % 2;
                                tar_rgn_offset.Offset[r_inx, 4 * col_inx].Value = before_data[r_inx];
                                tar_rgn_offset.Offset[r_inx, 4 * col_inx + 1].Value = after_data[r_inx];
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
                                        sum_rgn_offset.Offset[r_inx / 2, col_inx].Value = "OK";
                                    }
                                    else
                                    {
                                        sum_rgn_offset.Offset[r_inx / 2, col_inx].Value = "NG";
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
                MessageBox.Show("Format of ItemCode " + ItemCode + " not found", "Warning");
            }
            
        }
        public void Export_Thermal_HeatSoak_Bend_old_2(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name, int qty)
        {
            //myExcel.Workbook report_wrk = TDMK_Code.Create_Report_new(Application.StartupPath, format_name, "", ItemCode + "-" + LotNo);
            string format_file = exp_proc.find_format(myVar.data_loc, ItemCode);
            if (format_file != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable _src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                SortedDictionary<int, string> sel_index_lst = Get_bending_Net(sqlcon_OK2SHIP, ItemCode, process_name);
                List<int> index_lst = sel_index_lst.Keys.ToList();
                DataTable src_tbl = _src_tbl.AsEnumerable().Where(x => index_lst.Contains(Convert.ToInt32(x.Field<string>("Net_no")) - 1)).CopyToDataTable();
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                myExcel.Workbook report_wrk = exp_proc.create_export_wrk(format_file, format_name, ItemCode, LotNo);
                myExcel.Worksheet tar_wrksht = report_wrk.Sheets[1];
                myExcel.Range tar_rgn = tar_wrksht.Range["B33"];
                myExcel.Range sum_rgn = tar_wrksht.Range["C17"];
                int col_inx = 0;
                foreach (var tbl in src_tbl_lst)
                {
                    Dictionary<string, List<string>> data_lst = Get_List_Pair_data_bending(tbl);
                    List<string> before_data = data_lst["Before"];
                    int r_offset = 0;
                    foreach (var item in data_lst)
                    {
                        if (item.Key != "Before")
                        {
                            myExcel.Range tar_rgn_offset = tar_rgn.Offset[r_offset * tbl.Rows.Count, 0];
                            myExcel.Range sum_rgn_offset = sum_rgn.Offset[r_offset, 0];
                            List<string> after_data = item.Value;
                            string[] sum_arr = new string[2];
                            int sel_qty = Math.Min(qty * 2, before_data.Count);
                            for (int r_inx = 0; r_inx < sel_qty; r_inx++)
                            {
                                int sum_inx = r_inx % 2;
                                tar_rgn_offset.Offset[r_inx, 4 * col_inx].Value = before_data[r_inx];
                                tar_rgn_offset.Offset[r_inx, 4 * col_inx + 1].Value = after_data[r_inx];
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
                                        sum_rgn_offset.Offset[r_inx / 2, col_inx].Value = "OK";
                                    }
                                    else
                                    {
                                        sum_rgn_offset.Offset[r_inx / 2, col_inx].Value = "NG";
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
                MessageBox.Show("Format of ItemCode " + ItemCode + " not found", "Warning");
            }

        }
        public void Export_ThermalCycling(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name)
        {
            //myExcel.Workbook report_wrk = TDMK_Code.Create_Report_new(Application.StartupPath, format_name, "", ItemCode + "-" + LotNo);
            string format_file = exp_proc.find_format(myVar.data_loc, ItemCode);
            if (format_file != "")
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
                myExcel.Workbook report_wrk = exp_proc.create_export_wrk(format_file, format_name, ItemCode, LotNo);
                myExcel.Worksheet tar_wrksht = report_wrk.Sheets[1];
                myExcel.Range sum_rgn = tar_wrksht.Range["C17"];
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
                            sum_rgn.Offset[item_inx, col_inx].Value = "OK";
                        }
                        else
                        {
                            sum_rgn.Offset[item_inx, col_inx].Value = "NG";
                        }
                        item_inx++;
                    }
                    col_inx++;
                }
            }
            else
            {
                MessageBox.Show("Format of ItemCode " + ItemCode + " not found", "Warning");
            }
            
        }
        public void Export_Survival(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name)
        {
            
            //myExcel.Workbook report_wrk = TDMK_Code.Create_Report_new(Application.StartupPath, format_name, "", ItemCode + "-" + LotNo);
            string format_file = exp_proc.find_format(myVar.data_loc, ItemCode);
            if (format_file != "")
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
                myExcel.Workbook report_wrk = exp_proc.create_export_wrk(format_file, format_name, ItemCode, LotNo);
                myExcel.Worksheet tar_wrksht = report_wrk.Sheets[1];
                myExcel.Range tar_rgn = tar_wrksht.Range["C16"];
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
                            tar_rgn.Offset[off_set * cur_data.Count + r_inx, col_inx].Value = cur_data[r_inx];
                        }
                        col_inx++;
                    }
                    off_set++;
                }
            }
            else
            {
                MessageBox.Show("Format of ItemCode " + ItemCode + " not found", "Warning");
            }           
        }
        public DataTable Survival_HotOil_process(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string logfile)
        {
            Dictionary<string, List<string>> NET_data_lst = new Dictionary<string, List<string>>();
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo","Logfile" }, new string[] { ItemCode, LotNo, Path.GetFileName(logfile) });
            List<string> item_lst = new List<string>() { "ID", "ItemCode", "LotNo", "Pcs_No", "Logfile" };
            start_lbl: DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
            if (src_dt.Rows.Count > 0)
            {
                if (MessageBox.Show("Data is existed in Database. Do you want to update again?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    TDMK_Code.Delelte_FilteredItem_arr(process_name, sqlcon_OK2SHIP, filter_str);
                    goto start_lbl;
                }
            }
            else
            {
                myExcel.Workbook tar_wrkbk = TDMK_Code.open_excel_file(logfile, "", "");
                myExcel.Worksheet tar_wrksht = tar_wrkbk.Sheets["500 CHUKY"];
                myExcel.Range tar_rgn = tar_wrksht.Range["C11"];
                myExcel.Range data_rgn = tar_wrksht.Range["C16"];
                int col_inx = 0;
                while (!data_rgn.Offset[0, col_inx].HasFormula)
                {
                    string cell_val = myCode.checkDBNull(tar_rgn.Offset[0, col_inx].Value);
                    if (cell_val.Contains("Before") || (cell_val.Contains("After")))
                    {
                        string _col_name = Get_Number_String(cell_val, ' ');
                        string col_name = "";
                        if (cell_val.Contains("Before"))
                        {
                            col_name = "Before";
                        }
                        else
                        {
                            if (_col_name != "")
                            {
                                col_name = "After_" + _col_name;
                            }
                        }
                        List<string> NET_lst = new List<string>();
                        int r_inx = 0;
                        while (myCode.checkDBNull(data_rgn.Offset[r_inx, col_inx].Value) != "")
                        {
                            NET_lst.Add(myCode.checkDBNull(data_rgn.Offset[r_inx, col_inx].Value));
                            r_inx++;
                        }
                        NET_data_lst.Add(col_name, NET_lst);
                    }
                    col_inx++;
                }
                item_lst.AddRange(NET_data_lst.Keys.ToList());
                int row_count = NET_data_lst.ElementAt(0).Value.Count;
                int id = TDMK_Code.SQL_MAX(process_name, "ID", sqlcon_OK2SHIP);
                for (int i = 0; i < row_count; i++)
                {
                    string pcs = (i+1).ToString();
                    List<string> item_val_lst = new List<string>() { (id + i + 1).ToString(), ItemCode, LotNo,  pcs, Path.GetFileName(logfile) };
                    List<string> temp_lst = new List<string>();
                    foreach (var item in NET_data_lst)
                    {
                        temp_lst.Add(item.Value[i]);
                    }
                    item_val_lst.AddRange(temp_lst.ToList());
                    TDMK_Code.insert_val_arr2(process_name, sqlcon_OK2SHIP, item_lst.ToArray(), item_val_lst.ToArray());
                }
            }
            return TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
        }
        public void Export_Surv_HotOil(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name)
        {
            //myExcel.Workbook report_wrk = TDMK_Code.Create_Report_new(Application.StartupPath, format_name, "", ItemCode + "-" + LotNo);
            string format_file = exp_proc.find_format(myVar.data_loc, ItemCode);
            if (format_file != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                myExcel.Workbook report_wrk = exp_proc.create_export_wrk(format_file, format_name, ItemCode, LotNo);
                myExcel.Worksheet tar_wrksht = report_wrk.Sheets[1];
                myExcel.Range tar_rgn = tar_wrksht.Range["B16"];
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
                            tar_rgn.Offset[off_set * cur_data.Count + r_inx, col_inx].Value = cur_data[r_inx];
                        }
                        col_inx++;
                    }
                    off_set++;
                }
            }
            else
            {
                MessageBox.Show("Format of ItemCode " + ItemCode + " not found", "Warning");
            }
            
        }
        public void Export_Surv_HotOil_Multi(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name)
        {
            //myExcel.Workbook report_wrk = TDMK_Code.Create_Report_new(Application.StartupPath, format_name, "", ItemCode + "-" + LotNo);
            string format_file = exp_proc.find_format(myVar.data_loc, ItemCode);
            if (format_file != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                if(src_tbl.Rows.Count>0)
                {
                    myExcel.Workbook report_wrk = exp_proc.create_export_wrk(format_file, format_name, ItemCode, LotNo);
                    myExcel.Worksheet tar_wrksht = report_wrk.Sheets[1];
                    myExcel.Range start_tar_rgn = tar_wrksht.Range["A15"];
                    List<DataTable> _src_tbl_lst = new List<DataTable>();
                    Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Logfile" }, ref _src_tbl_lst, "Before");
                    //int inx = 0;
                    foreach (var temp_src_tbl in _src_tbl_lst)
                    {
                        string logfile = temp_src_tbl.Rows[0]["Logfile"].ToString();
                        string title_rgn = "";
                        if (logfile.Contains("BVH"))
                        {
                            title_rgn = Get_start_range("Resistance for BVH", start_tar_rgn.AddressLocal, tar_wrksht);
                        }
                        else
                        {
                            title_rgn = Get_start_range("Resistance for PTH", start_tar_rgn.AddressLocal, tar_wrksht);
                        }
                        if(title_rgn!="")
                        {
                            string data_rgn = Get_start_range("Sample No", title_rgn, tar_wrksht);
                            myExcel.Range tar_rgn = tar_wrksht.Range[data_rgn].Offset[0, 1];
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
                                        tar_rgn.Offset[off_set * cur_data.Count + r_inx, col_inx].Value = cur_data[r_inx];
                                    }
                                    col_inx++;
                                }
                                off_set++;
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Format of ItemCode " + ItemCode + " not found", "Warning");
            }

        }
        public void Export_Electrical(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name)
        {
            //myExcel.Workbook report_wrk = TDMK_Code.Create_Report_new(Application.StartupPath, format_name, "", ItemCode + "-" + LotNo);
            string format_file = exp_proc.find_format(myVar.data_loc, ItemCode);
            if (format_file != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Data");
                myExcel.Workbook report_wrk = exp_proc.create_export_wrk(format_file, format_name, ItemCode, LotNo);
                myExcel.Worksheet tar_wrksht = report_wrk.Sheets[1];
                myExcel.Range tar_rgn = tar_wrksht.Range["G35"];
                int col_inx = 0;
                foreach (var tbl in src_tbl_lst)
                {
                    for (int r_inx = 0; r_inx < tbl.Rows.Count; r_inx++)
                    {
                        tar_rgn.Offset[r_inx, col_inx].Value = tbl.Rows[r_inx]["Data"].ToString();
                    }
                    col_inx++;
                }
            }
            else
            {
                MessageBox.Show("Format of ItemCode " + ItemCode + " not found", "Warning");
            }
            
        }
        public List<string> Get_intersec(List<List<string>> src_lst)
        {
            List<string> result = new List<string>();
            if (src_lst.Count > 0)
            {
                result = src_lst[0];
                for (int i = 0; i < src_lst.Count; i++)
                {
                    var temp1 = result.Intersect(src_lst[i]);
                    result = temp1.ToList();
                }
            }
            return result;
        }
        public List<int> Get_NG_CPK_NET(SqlConnection sqlcon_OK2SHIP, DataTable src_tbl, string ItemCode, string process_name, double tar_CPK)
        {
            List<int> result = new List<int>();
            List<List<string>> src_NET_data_lst = new List<List<string>>();
            SortedDictionary<int, double> CPK_result_lst = new SortedDictionary<int, double>();
            DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process_name }));
            foreach (DataRow dr in src_tbl.Rows)
            {
                List<string> cur_Net_lst = new List<string>();
                foreach (DataColumn dc in src_tbl.Columns)
                {
                    cur_Net_lst.Add(dr[dc].ToString());
                }
                src_NET_data_lst.Add(cur_Net_lst);
            }
            int r_inx = 0;
            foreach (List<string> lst in src_NET_data_lst)
            {
                IEnumerable<double> cur_net_val = Calcu_process.ConvertToDouble(lst);
                double USL = Convert.ToDouble(spec_dt.Rows[r_inx]["USL"]);
                double LSL = Convert.ToDouble(spec_dt.Rows[r_inx]["LSL"]);
                double CPK = cur_net_val.ToArray().CPK(USL, LSL);
                CPK_result_lst.Add(r_inx, CPK);
                r_inx++;
            }
            List<int> ignored_lst = new List<int>();
            for (int i = 0; i < spec_dt.Rows.Count; i++)
            {
                if (spec_dt.Rows[i]["Sel_Report"].ToString() == "No")
                {
                    ignored_lst.Add(i);
                }
            }
            foreach (var item in CPK_result_lst.ToList())
            {
                if ((ignored_lst.FindIndex(x => x == item.Key) != -1) || (item.Value > tar_CPK))
                {
                    CPK_result_lst.Remove(item.Key);
                }
            }
            foreach (var item in CPK_result_lst.ToList())
            {
                result.Add(item.Key);
            }
            return result;
        }
        public List<string> Sum_OK_Item(SqlConnection sqlcon_OK2SHIP, DataTable src_tbl, string ItemCode, string process_name, double tar_CPK)
        {
            List<string> result = new List<string>();
            DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process_name }));
            List<int> CPK_NG_NET_lst = Get_NG_CPK_NET(sqlcon_OK2SHIP, src_tbl, ItemCode, process_name, tar_CPK);
            if(CPK_NG_NET_lst.Count>0)
            {
                List<List<string>> sel_item_lst = new List<List<string>>();
                foreach (int t in CPK_NG_NET_lst)
                {
                    List<string> cur_NET = new List<string>();
                    string USL = spec_dt.Rows[t]["USL"].ToString();
                    string LSL = spec_dt.Rows[t]["LSL"].ToString();
                    foreach (DataColumn dc in src_tbl.Columns)
                    {
                        cur_NET.Add(src_tbl.Rows[t][dc].ToString());
                    }
                    sel_item_lst.Add(Calcu_process.Items_OK_CPK_List(Calcu_process.ConvertToDouble(cur_NET).ToArray(), USL, LSL, 7, 50));
                }
                result = Get_intersec(sel_item_lst);
            }
            else
            {
                for(int t =0; t<src_tbl.Columns.Count;t++)
                {
                    result.Add(t.ToString());
                }
            }
            return result;
        }
        public void Check_Cycles_Data(DataGridView src_DGV, double ref_rate)
        {
            //try
            //{
                foreach (DataGridViewRow dr in src_DGV.Rows)
                {
                    foreach (DataGridViewColumn dc in src_DGV.Columns)
                    {
                        if (dc.Name.Contains("After"))
                        {
                            if (myCode.checkDBNull(dr.Cells[dc.Index].Value) != "")
                            {
                                if (myCode.IsNumeric(dr.Cells[dc.Index].Value.ToString())&&myCode.IsNumeric(dr.Cells["Before"].Value.ToString()))
                                {
                                    double bef_val = Convert.ToDouble(dr.Cells["Before"].Value);
                                    double aft_val = Convert.ToDouble(dr.Cells[dc.Index].Value);
                                    double rate = Math.Abs(bef_val - aft_val) / bef_val;
                                    if (rate > ref_rate)
                                    {
                                        dr.Cells[dc.Index].Style.BackColor = Color.Red;
                                    }
                                }
                                else
                                {
                                    dr.Cells[dc.Index].Style.BackColor = Color.Red;
                                }
                            }
                        }
                    }
                }
            //}
            //catch
            //{
            //    MessageBox.Show("Before Data not found", "Warning");
            //}
            
        }
        public void check_in_spec_tbl(DataTable dgv_spec, DataGridView dgv_data)
        {
            int r_inx = 0;
            foreach (DataGridViewRow dr in dgv_data.Rows)
            {
                string USL = dgv_spec.Rows[r_inx][3].ToString();
                string LSL = dgv_spec.Rows[r_inx][2].ToString();
                foreach (DataGridViewColumn dc in dgv_data.Columns)
                {
                    DataGridViewCell sel_cell = dr.Cells[dc.Index];
                    sel_cell.Style.BackColor = myCode.check_in_limit2(USL, LSL, sel_cell.Value.ToString());
                }
                r_inx++;
            }
        }
        public DataTable Electric_Summary(SqlConnection sqlcon_OK2SHIP, List<DataTable> tbl_lst,  string ItemCode, string process_name, int sel_qty)
        {
            DataTable dest_tbl = new DataTable();
            List<List<string>> src_NET_lst = new List<List<string>>();
            foreach (var tbl in tbl_lst)
            {
                int dr_inx = 0;
                foreach (DataRow dr in tbl.Rows)
                {
                    if (src_NET_lst.Count < dr_inx + 1)
                    {
                        src_NET_lst.Add(new List<string>());
                    }
                    foreach (DataColumn dc in tbl.Columns)
                    {
                        src_NET_lst[dr_inx].Add(dr[dc].ToString());
                    }
                    dr_inx++;
                }
            }
            //MessageBox.Show("Fill datatable");
            for (int i = 0; i < src_NET_lst[0].Count; i++)
            {
                dest_tbl.Columns.Add("ItemNo_" + (i + 1).ToString());
            }
            foreach (var n_lst in src_NET_lst)
            {
                DataRow dr = dest_tbl.NewRow();
                int c_inx = 0;
                foreach (string t in n_lst)
                {
                    dr[c_inx] = t;
                    c_inx++;
                }
                dest_tbl.Rows.Add(dr);
            }
            List<string> sel_Col = Sum_OK_Item(sqlcon_OK2SHIP, dest_tbl, ItemCode, process_name, 1.67);
            List<string> col_name_sel = new List<string>();
            int qty = Math.Min(sel_Col.Count, sel_qty);
            for (int i = 0; i < qty; i++)
            {
                string t = sel_Col[i];
                col_name_sel.Add(dest_tbl.Columns[int.Parse(t)].ColumnName);
            }
            DataTable result_tbl = dest_tbl.AsDataView().ToTable(false, col_name_sel.ToArray());
            return result_tbl;
        }
        public DataTable ThermalCycling_On_Coupon(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string logfile)
        {
            Dictionary<string, List<string>> NET_data_lst = new Dictionary<string, List<string>>();
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo","Logfile" }, new string[] { ItemCode, LotNo, Path.GetFileName(logfile) });
            List<string> item_lst = new List<string>() { "ID", "ItemCode", "LotNo", "Pcs_No", "Logfile" };
            start_lbl: DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
            if (src_dt.Rows.Count > 0)
            {
                if (MessageBox.Show("Data is existed in Database. Do you want to update again?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    TDMK_Code.Delelte_FilteredItem_arr(process_name, sqlcon_OK2SHIP, filter_str);
                    goto start_lbl;
                }
            }
            else
            {
                myExcel.Workbook report_wrk = TDMK_Code.open_excel_file(logfile, "", "");
                myExcel.Worksheet tar_wrksht = report_wrk.Sheets[1];
                myExcel.Range tar_rgn = tar_wrksht.Range["C14"];
                List<string> col_lst = new List<string>() { "Before", "After_100", "After_200", "After_300", "After_400", "After_500" };
                int col_inx = 0;
                while (!tar_rgn.Offset[0, col_inx].HasFormula)
                {
                    string cell_val = myCode.checkDBNull(tar_rgn.Offset[-7, col_inx].Value);
                    if (cell_val.Contains("Before") || (cell_val.Contains("After")))
                    {
                        string _col_name = Get_Number_String(cell_val, ' ');
                        string col_name = "";
                        if (cell_val.Contains("Before"))
                        {
                            col_name = "Before";
                        }
                        else
                        {
                            if (_col_name != "")
                            {
                                col_name = "After_" + _col_name;
                            }
                        }
                        if (col_name != "")
                        {
                            if (TDMK_Code.check_exist_list_index2(col_name, col_lst) != -1)
                            {
                                List<string> NET_lst = new List<string>();
                                int r_inx = 0;
                                while (myCode.checkDBNull(tar_rgn.Offset[r_inx, col_inx].Value) != "")
                                {
                                    NET_lst.Add(myCode.checkDBNull(tar_rgn.Offset[r_inx, col_inx].Value));
                                    r_inx++;
                                }
                                NET_data_lst.Add(col_name, NET_lst);
                            }
                        }

                    }
                    col_inx++;
                }
                item_lst.AddRange(NET_data_lst.Keys.ToList());
                int row_count = NET_data_lst.ElementAt(0).Value.Count;
                int id = TDMK_Code.SQL_MAX(process_name, "ID", sqlcon_OK2SHIP);
                for (int i = 0; i < row_count; i++)
                {
                    string pcs = (i + 1).ToString();
                    List<string> item_val_lst = new List<string>() { (id + i + 1).ToString(), ItemCode, LotNo, pcs, Path.GetFileName(logfile) };
                    List<string> temp_lst = new List<string>();
                    foreach (var item in NET_data_lst)
                    {
                        temp_lst.Add(item.Value[i]);
                    }
                    item_val_lst.AddRange(temp_lst.ToList());
                    TDMK_Code.insert_val_arr2(process_name, sqlcon_OK2SHIP, item_lst.ToArray(), item_val_lst.ToArray());
                }
                report_wrk.Close();
            }
            return TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
        }
        public void Export_ThermalCycling_On_Coupon(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name)
        {
            string format_file = exp_proc.find_format(myVar.data_loc, ItemCode);
            if (format_file != "")
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
                myExcel.Workbook report_wrk = exp_proc.create_export_wrk(format_file, format_name, ItemCode, LotNo);
                myExcel.Worksheet tar_wrksht = report_wrk.Sheets[1];
                myExcel.Range sum_rgn = tar_wrksht.Range["C26"];
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
                            sum_rgn.Offset[item_inx, col_inx].Value = "OK";
                        }
                        else
                        {
                            sum_rgn.Offset[item_inx, col_inx].Value = "NG";
                        }
                        item_inx++;
                    }
                    col_inx++;
                    if(col_inx ==10)
                    {
                        break;
                    }
                }
            }
            else
            {
                MessageBox.Show("Format of ItemCode " + ItemCode + " not found", "Warning");
            }
        }
        public Dictionary<string, List<string>> Get_BlockofLogFile(List<string> logfile_lst)
        {
            Dictionary<string, List<string>> block_logfile = new Dictionary<string, List<string>>();
            foreach (string t in logfile_lst)
            {
                string[] temp_name = t.Split('_');
                string item_name = temp_name[1];
                string[] temp = item_name.Split('-');
                string blockno = temp[temp.Length - 1];
                string block_name = item_name.Substring(0, item_name.Length - blockno.Length - 1);
                if (TDMK_Code.check_exist_list_index2(block_name, block_logfile.Keys.ToList()) != -1)
                {
                    block_logfile[block_name].Add(t);
                }
                else
                {
                    block_logfile.Add(block_name, new List<string>() { t });
                }
            }
            return block_logfile;
        }
        public Dictionary<string, List<string>> Get_BlockofLogFile_Taiyo(List<string> logfile_lst)
        {
            Dictionary<string, List<string>> block_logfile = new Dictionary<string, List<string>>();
            foreach (string t in logfile_lst)
            {
                string block_name = "";
                if (t.Contains("_"))
                {
                    string[] temp_name = t.Split('_');
                    string item_name = temp_name[1];
                    string[] temp = item_name.Split('-');
                    string blockno = temp[temp.Length - 1];
                    block_name = item_name.Substring(0, item_name.Length - blockno.Length - 1);
                }
                else
                {
                    string[] temp = t.Split('-');
                    string blockno = temp[temp.Length - 1];
                    block_name = temp[temp.Length - 4] + "-" + temp[temp.Length - 3] + "-" + temp[temp.Length - 2];// +"#"+  blockno.Substring(0, 1);
                }
                if (TDMK_Code.check_exist_list_index2(block_name, block_logfile.Keys.ToList()) != -1)
                {
                    block_logfile[block_name].Add(t);
                }
                else
                {
                    block_logfile.Add(block_name, new List<string>() { t });
                }
            }
            return block_logfile;
        }
        public DataTable Summary_Data_Table_from_list(List<DataTable> tbl_data_lst, DataTable spec_dt)
        {
            DataTable final_tbl = new DataTable();
            if (tbl_data_lst.Count > 1)
            {
                final_tbl = tbl_data_lst[0];
                int c_inx = 0;
                foreach (DataColumn dc in final_tbl.Columns)
                {
                    int r_inx = 0;
                    foreach (DataRow dr in final_tbl.Rows)
                    {
                        string USL = spec_dt.Rows[r_inx][3].ToString();
                        string LSL = spec_dt.Rows[r_inx][2].ToString();
                        string cur_val = dr[dc].ToString();
                        if (!check_in_limit(USL, LSL, cur_val))
                        {
                            for (int i = 1; i < tbl_data_lst.Count; i++)
                            {
                                DataTable cur_dt = tbl_data_lst[i];
                                if (check_in_limit(USL, LSL, cur_dt.Rows[r_inx][c_inx].ToString()))
                                {
                                    dr[dc] = cur_dt.Rows[r_inx][c_inx];
                                }
                            }
                        }
                        r_inx++;
                    }
                    c_inx++;
                }
            }
            else
            {
                final_tbl = tbl_data_lst[0];
            }
            return final_tbl;
        }
        public void Electrical_CPK_Details(DataTable src_tbl, DataGridView DGV_Data, SqlConnection sqlcon_OK2SHIP, DataGridView DGV_NG_detail, string ItemCode, string process_name)
        {
            DGV_Data.DataSource = null;
            DGV_Data.Rows.Clear();
            DGV_Data.Columns.Clear();
            DGV_NG_detail.DataSource = null;
            DGV_NG_detail.Rows.Clear();
            DGV_NG_detail.Columns.Clear();
            List<List<string>> src_NET_data_lst = new List<List<string>>();
            List<List<string>> _src_NET_data_lst = new List<List<string>>();
            SortedDictionary<int, double> CPK_result_lst = new SortedDictionary<int, double>();
            foreach (DataRow dr in src_tbl.Rows)
            {
                List<string> cur_Net_lst = new List<string>();
                foreach (DataColumn dc in src_tbl.Columns)
                {
                    cur_Net_lst.Add(dr[dc].ToString());
                }
                src_NET_data_lst.Add(cur_Net_lst);
            }
            int col_inx = 0;
            DGV_Data.Columns.Add("CPK", "CPK");
            DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process_name }));
            DataTable _spec_dt = spec_dt.AsDataView().ToTable(false, new string[] { "Point+V", "Point-V", "LSL", "USL" });
            foreach (List<string> lst in src_NET_data_lst)
            {
                IEnumerable<double> cur_net_val = Calcu_process.ConvertToDouble(lst);
                double USL = Convert.ToDouble(_spec_dt.Rows[col_inx][3].ToString());
                double LSL = Convert.ToDouble(_spec_dt.Rows[col_inx][2].ToString());
                double CPK = cur_net_val.ToArray().CPK(USL, LSL);
                DGV_Data.Rows.Add(CPK.ToString("#0.000#"));
                if (CPK < 1.67)
                {
                    DGV_Data.Rows[col_inx].Cells[0].Style.BackColor = Color.Red;
                }
                CPK_result_lst.Add(col_inx, CPK);
                col_inx++;
            }
            myCode.DGV_Auto_Resize(DGV_Data);
            List<int> ignored_lst = new List<int>();
            for (int i = 0; i < spec_dt.Rows.Count; i++)
            {
                if (spec_dt.Rows[i]["Net_Internal"].ToString() == "No")
                {
                    ignored_lst.Add(i);
                    DGV_Data.Rows[i].Cells[0].Style.BackColor = Color.Blue;
                }
                else
                {
                    _src_NET_data_lst.Add(src_NET_data_lst[i]);
                }
            }
            foreach (var item in CPK_result_lst.ToList())
            {
                if ((ignored_lst.FindIndex(x => x == item.Key) != -1) || (item.Value > 1.67))
                {
                    CPK_result_lst.Remove(item.Key);
                }
            }
            int r_inx = 0;
            DGV_NG_detail.Columns.Add("CPK", "CPK");
            foreach (KeyValuePair<int, double> k in CPK_result_lst)
            {
                DGV_NG_detail.Rows.Add(k.Value.ToString("#0.000#"));
                DGV_NG_detail.Rows[r_inx].HeaderCell.Value = (k.Key + 1).ToString();
                r_inx++;
            }
            DGV_NG_detail.AutoResizeColumns();
            DGV_NG_detail.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }
        public DataTable Electric_Summary_Block(SqlConnection sqlcon_OK2SHIP, Dictionary<string, DataTable> dic_tbl_lst, string ItemCode, string process_name, int sel_qty)
        {
            DataTable dest_tbl = new DataTable();
            List<List<string>> src_NET_lst = new List<List<string>>();
            List<string> tar_col_name = new List<string>();
            foreach (var dic_tbl in dic_tbl_lst)
            {
                var tbl = dic_tbl.Value;
                string bl_name = dic_tbl.Key.Split('-')[2];
                int dr_inx = 0;
                foreach (DataRow dr in tbl.Rows)
                {
                    if (src_NET_lst.Count < dr_inx + 1)
                    {
                        src_NET_lst.Add(new List<string>());
                    }
                    foreach (DataColumn dc in tbl.Columns)
                    {
                        src_NET_lst[dr_inx].Add(dr[dc].ToString());
                    }
                    dr_inx++;
                }
                foreach (DataColumn dc in tbl.Columns)
                {
                    tar_col_name.Add(dc.ColumnName + "_BL" + bl_name);
                }
            }
            //MessageBox.Show("Fill datatable");
            for (int i = 0; i < tar_col_name.Count; i++)
            {
                dest_tbl.Columns.Add(tar_col_name[i]);
            }
            foreach (var n_lst in src_NET_lst)
            {
                DataRow dr = dest_tbl.NewRow();
                int c_inx = 0;
                foreach (string t in n_lst)
                {
                    dr[c_inx] = t;
                    c_inx++;
                }
                dest_tbl.Rows.Add(dr);
            }
            List<string> sel_Col = Sum_OK_Item(sqlcon_OK2SHIP, dest_tbl, ItemCode, process_name, 1.67);
            List<string> col_name_sel = new List<string>();
            int qty = Math.Min(sel_Col.Count, sel_qty);
            for (int i = 0; i < qty; i++)
            {
                string t = sel_Col[i];
                col_name_sel.Add(dest_tbl.Columns[int.Parse(t)].ColumnName);
            }
            DataTable result_tbl = dest_tbl.AsDataView().ToTable(false, col_name_sel.ToArray());
            return result_tbl;
        }
        public DataTable Electric_Data_Summary(Dictionary<string, DataTable> dic_tbl_lst)
        {
            DataTable dest_tbl = new DataTable();
            List<List<string>> src_NET_lst = new List<List<string>>();
            List<string> tar_col_name = new List<string>();
            foreach (var dic_tbl in dic_tbl_lst)
            {
                var tbl = dic_tbl.Value;
                string bl_name = dic_tbl.Key.Split('-')[2];
                int dr_inx = 0;
                foreach (DataRow dr in tbl.Rows)
                {
                    if (src_NET_lst.Count < dr_inx + 1)
                    {
                        src_NET_lst.Add(new List<string>());
                    }
                    foreach (DataColumn dc in tbl.Columns)
                    {
                        src_NET_lst[dr_inx].Add(dr[dc].ToString());
                    }
                    dr_inx++;
                }
                foreach (DataColumn dc in tbl.Columns)
                {
                    tar_col_name.Add(dc.ColumnName + "_BL" + bl_name);
                }
            }
            //MessageBox.Show("Fill datatable");
            for (int i = 0; i < tar_col_name.Count; i++)
            {
                dest_tbl.Columns.Add(tar_col_name[i]);
            }
            foreach (var n_lst in src_NET_lst)
            {
                DataRow dr = dest_tbl.NewRow();
                int c_inx = 0;
                foreach (string t in n_lst)
                {
                    dr[c_inx] = t;
                    c_inx++;
                }
                dest_tbl.Rows.Add(dr);
            }
            return dest_tbl;
        }
        public Color Check_Variability_Resistance(double ref_rate, int cell_row, int cell_col, DataTable src_tbl, DataTable cycle_tbl)
        {
            Color result_color = Color.Gray;
            if ((cell_row > -1) && (cell_col > -1))
            {
                string col_name = cycle_tbl.Columns[cell_col].ColumnName;
                string[] sel_data = src_tbl.AsEnumerable().Where(x => x.Field<string>("Pcs_No") == col_name && x.Field<string>("Net_No") == (cell_row + 1).ToString()).Select(x => x.Field<string>("Before")).ToArray();
                string cur_val = myCode.checkDBNull(cycle_tbl.Rows[cell_row][cell_col]);
                if (myCode.IsNumeric(cur_val) && (myCode.IsNumeric(sel_data[0])))
                {
                    double _cur_val = Convert.ToDouble(cur_val);
                    double _sel_val = Convert.ToDouble(sel_data[0]);
                    double rate = Math.Abs(_cur_val - _sel_val) / _sel_val;
                    if(rate < ref_rate)
                    {
                        result_color = Color.White;
                    }
                    else
                    {
                        result_color = Color.Red;
                    }
                }
            }
            return result_color;
        }
        public string Details_Variability_Resistance(int cell_row, int cell_col, DataTable src_tbl, DataTable cycle_tbl)
        {
            string result = "Error";
            if ((cell_row > -1) && (cell_col > -1))
            {
                string col_name = cycle_tbl.Columns[cell_col].ColumnName;
                string[] sel_data = src_tbl.AsEnumerable().Where(x => x.Field<string>("Pcs_No") == col_name && x.Field<string>("Net_No") == (cell_row + 1).ToString()).Select(x => x.Field<string>("Before")).ToArray();
                string cur_val = myCode.checkDBNull(cycle_tbl.Rows[cell_row][cell_col]);
                string rate_val = "Invalid";
                if (myCode.IsNumeric(cur_val) && (myCode.IsNumeric(sel_data[0])))
                {
                    double _cur_val = Convert.ToDouble(cur_val);
                    double _sel_val = Convert.ToDouble(sel_data[0]);
                    double rate = Math.Abs(_cur_val - _sel_val) / _sel_val;
                    rate_val = (rate*100).ToString("#0.00#") + "%";                    
                }
                result = "Before Value = " + sel_data[0] + "\r\n" + "\u0394" + "R = " + rate_val;
            }
            return result;
        }
        public string Details_Variability_Resistance_bending(int cell_row, int cell_col, DataTable src_tbl, DataTable cycle_tbl, string sel_cycle, SortedDictionary<int, string>dic_Net)
        {
            string result = "Error";
            if ((cell_row > -1) && (cell_col > -1))
            {
                string col_name = cycle_tbl.Columns[cell_col].ColumnName;
                int net_no = cell_row;
                if (sel_cycle.Contains("Bending"))
                {
                    if(dic_Net.Count >0)
                    {
                        net_no = dic_Net.Keys.ToArray()[cell_row];
                    }
                    else
                    {
                        return result;
                    }
                }
                string[] sel_data = src_tbl.AsEnumerable().Where(x => x.Field<string>("Pcs_No") == col_name && x.Field<string>("Net_No") == (net_no + 1).ToString()).Select(x => x.Field<string>("Before")).ToArray();
                if(sel_data.Length>0)
                {
                    string cur_val = myCode.checkDBNull(cycle_tbl.Rows[cell_row][cell_col]);
                    string rate_val = "Invalid";
                    if (myCode.IsNumeric(cur_val) && (myCode.IsNumeric(sel_data[0])))
                    {
                        double _cur_val = Convert.ToDouble(cur_val);
                        double _sel_val = Convert.ToDouble(sel_data[0]);
                        double rate = Math.Abs(_cur_val - _sel_val) / _sel_val;
                        rate_val = (rate * 100).ToString("#0.00#") + "%";
                    }
                    result = "Before Value = " + sel_data[0] + "\r\n" + "\u0394" + "R = " + rate_val;
                }

            }
            return result;
        }
        public Color CheckCell_Variability_Resistance(double ref_rate, DataTable src_tbl, DataGridViewCell tar_cell, string col_name)
        {
            Color result_color = Color.Gray;
            int cell_row = tar_cell.RowIndex;
            int cell_col = tar_cell.ColumnIndex;
            if ((cell_row > -1) && (cell_col > -1))
            {
                string[] sel_data = src_tbl.AsEnumerable().Where(x => x.Field<string>("Pcs_No") == col_name && x.Field<string>("Net_No") == (cell_row + 1).ToString()).Select(x => x.Field<string>("Before")).ToArray();
                string cur_val = myCode.checkDBNull(tar_cell.Value);
                if (myCode.IsNumeric(cur_val) && (myCode.IsNumeric(sel_data[0])))
                {
                    double _cur_val = Convert.ToDouble(cur_val);
                    double _sel_val = Convert.ToDouble(sel_data[0]);
                    double rate = Math.Abs(_cur_val - _sel_val) / _sel_val;
                    if (rate < ref_rate)
                    {
                        result_color = Color.White;
                    }
                    else
                    {
                        result_color = Color.Red;
                    }
                }
            }
            return result_color;
        }
        public Color CheckCell_Variability_Resistance_bending(double ref_rate, DataTable src_tbl, DataGridViewCell tar_cell, string col_name, string sel_cycle, SortedDictionary<int, string>dic_Net)
        {
            Color result_color = Color.Gray;
            int cell_row = tar_cell.RowIndex;
            int cell_col = tar_cell.ColumnIndex;
            if ((cell_row > -1) && (cell_col > -1))
            {
                
                int net_no = cell_row;
                if (sel_cycle.Contains("Bending"))
                {
                    if (dic_Net.Count > 0)
                    {
                        net_no = dic_Net.Keys.ToArray()[cell_row];
                    }
                    else
                    {
                        return result_color;
                    }
                }
                string[] sel_data = src_tbl.AsEnumerable().Where(x => x.Field<string>("Pcs_No") == col_name && x.Field<string>("Net_No") == (net_no + 1).ToString()).Select(x => x.Field<string>("Before")).ToArray();
                string cur_val = myCode.checkDBNull(tar_cell.Value);
                if (myCode.IsNumeric(cur_val) && (myCode.IsNumeric(sel_data[0])))
                {
                    double _cur_val = Convert.ToDouble(cur_val);
                    double _sel_val = Convert.ToDouble(sel_data[0]);
                    double rate = Math.Abs(_cur_val - _sel_val) / _sel_val;
                    if (rate < ref_rate)
                    {
                        result_color = Color.White;
                    }
                    else
                    {
                        result_color = Color.Red;
                    }
                }
            }
            return result_color;
        }
        public Dictionary<string, List<int>> Check_Cycles_Data_Vary(DataGridView src_DGV, double ref_rate)
        {
            Dictionary<string, List<int>> result = new Dictionary<string, List<int>>();
            try
            {
                foreach (DataGridViewRow dr in src_DGV.Rows)
                {
                    foreach (DataGridViewColumn dc in src_DGV.Columns)
                    {
                        if (dc.Name.Contains("After"))
                        {
                            
                            if (myCode.checkDBNull(dr.Cells[dc.Index].Value) != "")
                            {
                                bool add_en = false;
                                if (myCode.IsNumeric(dr.Cells[dc.Index].Value.ToString()))
                                {
                                    double bef_val = Convert.ToDouble(dr.Cells["Before"].Value);
                                    double aft_val = Convert.ToDouble(dr.Cells[dc.Index].Value);
                                    double rate = Math.Abs(bef_val - aft_val) / bef_val;
                                    if (rate > ref_rate)
                                    {
                                        dr.Cells[dc.Index].Style.BackColor = Color.Red;
                                        add_en = true;
                                    }
                                }
                                else
                                {
                                    dr.Cells[dc.Index].Style.BackColor = Color.Red;
                                    add_en = true;
                                }
                                if(add_en)
                                {
                                    string col_name = dr.Cells["Pcs_No"].Value.ToString();
                                    int r_no = dr.Index + 1;
                                    if (TDMK_Code.check_exist_list_index2(col_name, result.Keys.ToList()) == -1)
                                    {
                                        result.Add(col_name, new List<int>());
                                    }
                                    result[col_name].Add(r_no);
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("Before Data not found", "Warning");
            }
            return result;
        }
        public DataTable Filter_table_by_index(DataTable src_dt, List<int> src_index)
        {
            var allowedIndices = src_index.ToArray();
            DataTable tblAllowedRows = src_dt.AsEnumerable().Where((r, i) => allowedIndices.Contains(i)).CopyToDataTable();
            return tblAllowedRows;
        }
        public void update_bending_table(SqlConnection sqlcon, DataTable src_dt, DataTable data_tbl, string ItemCode, string Process_name, string cycle_name)
        {
            List<string> Net_lst = new List<string>();
            Net_lst = src_dt.AsEnumerable().Select(x => x.Field<string>("Net_No")).Distinct().ToList();
            DataTable temp_spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, Process_name }));
            List<int> index_lst = new List<int>();
            int inx = 0;
            foreach (DataRow dr in temp_spec_dt.Rows)
            {
                if (myCode.checkDBNull(dr["Sel_Report"]) != "")
                {
                    index_lst.Add(inx);
                }
                inx++;
            }
            for (int i = 0; i < data_tbl.Rows.Count; i++)
            {
                int net_no = index_lst[i];
                foreach (DataColumn dc in data_tbl.Columns)
                {
                    string pcs_no = dc.ColumnName;
                    int ng_item_inx = data_tbl.Columns.IndexOf(pcs_no);
                    int tar_row = ng_item_inx * Net_lst.Count + net_no;
                    src_dt.Rows[tar_row][cycle_name] = data_tbl.Rows[i][dc];
                }
            }
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
        public void Export_Thermal_HeatSoak_Bend(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name, int qty)
        {
            string format_file = exp_proc.find_format(myVar.data_loc, ItemCode);
            if (format_file != "")
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
                DataTable src_tbl = _src_tbl.AsEnumerable().Where(x => index_lst.Contains(Convert.ToInt32(x.Field<string>("Net_no")) - 1)).CopyToDataTable();
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                myExcel.Workbook report_wrk = exp_proc.create_export_wrk(format_file, format_name, ItemCode, LotNo);
                myExcel.Worksheet tar_wrksht = report_wrk.Sheets[1];
                string Echeck_start_rgn = Get_start_range("Condition", "A10", tar_wrksht);
                myExcel.Range Echeck_cycle_rgn = tar_wrksht.Range[Echeck_start_rgn];
                Dictionary<string, string> Echeck_Cycle_addr = Get_Echeck_address(Echeck_cycle_rgn);
                string start_rgn = Get_start_range("Test item", "A10", tar_wrksht);
                myExcel.Range cycle_rgn = tar_wrksht.Range[start_rgn];
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
                            myExcel.Range sum_rgn_offset = tar_wrksht.Range[Echeck_Cycle_addr[item.Key]].Offset[0, 2];
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
                                        sum_rgn_offset.Offset[r_inx / 2, col_inx].Value = "OK";
                                    }
                                    else
                                    {
                                        sum_rgn_offset.Offset[r_inx / 2, col_inx].Value = "NG";
                                    }
                                }
                            }
                        }
                        if (bending_Cycle_addr.Keys.ToList().IndexOf(item.Key) != -1)
                        {
                            myExcel.Range tar_rgn_offset = tar_wrksht.Range[bending_Cycle_addr[item.Key]].Offset[0, 1];
                            for (int r_inx = 0; r_inx < before_data.Count; r_inx++)
                            {
                                tar_rgn_offset.Offset[r_inx, 4 * col_inx].Value = before_data[r_inx];
                                tar_rgn_offset.Offset[r_inx, 4 * col_inx + 1].Value = after_data[r_inx];
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Format of ItemCode " + ItemCode + " not found", "Warning");
            }
        }
        public Dictionary<string, string> Get_Echeck_address(myExcel.Range cycle_rgn)
        {
            Dictionary<string, string> bending_Cycle_addr = new Dictionary<string, string>();
            int row_inx = 0;
            while (myCode.checkDBNull(cycle_rgn.Offset[row_inx, 0].Value) != "")
            {
                string cycle = Get_Number_String(cycle_rgn.Offset[row_inx, 0].Value, ' ');
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
                    bending_Cycle_addr.Add(cycle_name, cycle_rgn.Offset[row_inx, 0].AddressLocal);
                }
                row_inx++;
            }
            return bending_Cycle_addr;
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
        public string Get_blockname_from_Logfile(string src_logfile)
        {
            string result = "Error";
            try
            {
                if (src_logfile.Contains("_"))
                {
                    string[] temp_name = src_logfile.Split('_');
                    if (temp_name.Length == 3)
                    {
                        string item_name = temp_name[1];
                        string[] temp = item_name.Split('-');
                        string blockno = temp[temp.Length - 1];
                        result = item_name.Substring(0, item_name.Length - blockno.Length - 1);
                    }
                }
                else
                {
                    string[] temp = src_logfile.Split('-');
                    if (temp.Length >= 9)
                    {
                        string blockno = temp[temp.Length - 1];
                        result = temp[temp.Length - 4] + "-" + temp[temp.Length - 3] + "-" + temp[temp.Length - 2];// + "#" + blockno.Substring(0, 1);
                    }
                }
            }
            catch
            {

            }
            return result;
        }
        public bool Logfile_ItemCode_LotNo_Matching(string src_logfile, string ItemCode, string LotNo)
        {
            bool result = false;
            string blockname = Get_blockname_from_Logfile(src_logfile);
            if(!myVar_ECheck.CSV_en)
            {
                blockname = src_logfile;
            }
            if(blockname!="Error")
            {
                if(blockname.Contains(ItemCode) && blockname.Contains(LotNo.Split('-')[0]))
                {
                    result = true;
                }
            }
            return result;
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
        public DataTable Summary_Logfile_All(SqlConnection sqlcon, string itemcode, string lotno, string process, string tar_cycle,int sel_num = 1)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcode, lotno });
            DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, process, filter_str);
            int id = TDMK_Code.SQL_MAX(process, "ID", sqlcon);            
            if (process != "ELECTRICAL")
            {
                Dictionary<string, DataTable> dic_sum = Summary_Cycles_Logfile(sqlcon, itemcode, lotno, process, tar_cycle, sel_num);
                foreach (var cycle_dt in dic_sum)
                {
                    string cycle_name = cycle_dt.Key;
                    foreach (DataColumn dc in cycle_dt.Value.Columns)
                    {
                        int r_offset = src_dt.Rows.Count;
                        string pcs = dc.ColumnName;
                        DataView dv = src_dt.AsDataView();
                        dv.RowFilter = "Pcs_No = '" + pcs + "'";
                        if (dv.Count > 0)
                        {
                            if (dv.Count == cycle_dt.Value.Rows.Count)
                            {
                                for (int i = 0; i < dv.Count; i++)
                                {
                                    dv[i][cycle_name] = cycle_dt.Value.Rows[i][dc];
                                }
                            }
                            else
                            {
                                MessageBox.Show("Dữ liệu chu kỳ " + cycle_name + " không đủ", "Thông báo");
                            }
                        }
                        else
                        {
                            for (int i = 0; i < cycle_dt.Value.Rows.Count; i++)
                            {
                                src_dt.Rows.Add((id + 1).ToString(), itemcode, lotno, (i + 1), pcs);
                                src_dt.Rows[i + r_offset][cycle_name] = cycle_dt.Value.Rows[i][dc];
                                id++;
                            }
                        }
                    }
                }
                TDMK_Code.Delelte_FilteredItem_arr(process, sqlcon, filter_str);
                BatchBulkCopy(sqlcon, src_dt, process);
            }
            else
            {
                string spec_filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { itemcode, process });
                DataTable src_log_dt = TDMK_Code.Datatable_Filter(sqlcon, process + "_LOGFILE", filter_str);
                DataTable src_data_dt = TDMK_Code.Datatable_Filter(sqlcon, process, filter_str);
                DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "NET_SPEC", spec_filter_str);
                List<string> lst_log = src_log_dt.AsEnumerable().Select(x=>x.Field<string>("Logfile")).Distinct().ToList();
                Dictionary<string, DataTable>  result_table_lst = new Dictionary<string, DataTable>();
                Dictionary<string, List<DataTable>> dic_tbl_data_lst = new Dictionary<string, List<DataTable>>();
                Dictionary<string, List<string>> block_logfile = Get_BlockofLogFile_Taiyo(lst_log);
                foreach (var block in block_logfile)
                {
                    foreach (var logfile in block.Value)
                    {
                        //DataTable dt = Load_Log_Data(sqlcon, itemcode, lotno, process, tar_cycle, logfile);
                        DataTable dt = Load_Log_Data_fromTable(src_log_dt, logfile);
                        if (dt.Rows.Count > 0)
                        {
                            if (TDMK_Code.check_exist_list_index2(block.Key, dic_tbl_data_lst.Keys.ToList()) == -1)
                            {
                                dic_tbl_data_lst.Add(block.Key, new List<DataTable>() { dt });
                            }
                            else
                            {
                                dic_tbl_data_lst[block.Key].Add(dt);
                            }
                        }
                    }
                }
                foreach (var tbl in dic_tbl_data_lst)
                {
                    DataTable block_tbl = Summary_Data_Table_from_list(tbl.Value, spec_dt);
                    result_table_lst.Add(tbl.Key, block_tbl);
                }
                DataTable _src_dt = Electric_Summary_Block(sqlcon, result_table_lst, itemcode, process, sel_num);
                src_dt = Save_Submit_data(sqlcon, _src_dt, itemcode, lotno, process, tar_cycle);
            }
            MessageBox.Show("Lưu dữ liệu hoàn thành", "Thông báo");
            return src_dt;
        }
        public Dictionary<string, DataTable> Summary_Cycles_Logfile(SqlConnection sqlcon, string itemcode, string lotno, string process, string tar_cycle="All", int sel_num = 1)
        {
            Dictionary<string, DataTable> dic_sum = new Dictionary<string, DataTable>();
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo"}, new string[] { itemcode, lotno});
            string spec_filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { itemcode, process });
            DataTable src_log_dt = TDMK_Code.Datatable_Filter(sqlcon, process + "_LOGFILE", filter_str);
            DataTable src_data_dt = TDMK_Code.Datatable_Filter(sqlcon, process , filter_str);
            List<string> _after_pcs_lst = src_log_dt.AsEnumerable().Where(x => x.Field<string>("Cycles_name").Contains("After")).Select(x => x.Field<string>("Pcs_No")).Distinct().ToList();
            List<string> after_pcs_lst = _after_pcs_lst.Where(x => !x.Contains("_")).Select(x => x.ToUpper()).Distinct().ToList();
            List<string> cycle_lst = src_log_dt.AsEnumerable().Select(x => x.Field<string>("Cycles_name")).Distinct().ToList();
            if(tar_cycle!="All" )
            {
                if (cycle_lst.IndexOf(tar_cycle) != -1)
                {
                    cycle_lst = new List<string>() { tar_cycle };
                }
                else
                {
                    cycle_lst = new List<string>();
                }

            }
            foreach (string cycle in cycle_lst)
            {
                dic_sum.Add(cycle, new DataTable());
                List<string> logfile_lst = src_log_dt.AsEnumerable().Where(x => x.Field<string>("Cycles_name") == cycle).Select(x => x.Field<string>("Logfile")).Distinct().ToList();
                Dictionary<string, List<string>> dic_logfile = Get_BlockofLogFile_Taiyo(logfile_lst);
                Dictionary<string, DataTable> result_table_lst = new Dictionary<string, DataTable>();
                Dictionary<string, List<DataTable>> dic_tbl_data_lst = new Dictionary<string, List<DataTable>>();
                DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "NET_SPEC", spec_filter_str);
                foreach (var block in dic_logfile)
                {
                    foreach (var logfile in block.Value)
                    {
                        DataTable dt = new DataTable();
                        List<string> temp_net_lst = new List<string>();
                        dt = Load_Log_Data(sqlcon, itemcode, lotno, process, cycle, logfile);
                        //dt = Load_Log_Data_fromTable(src_log_dt, logfile);
                        if (dt.Rows.Count > 0)
                        {
                            if (TDMK_Code.check_exist_list_index2(block.Key, dic_tbl_data_lst.Keys.ToList()) == -1)
                            {
                                dic_tbl_data_lst.Add(block.Key, new List<DataTable>() { dt });
                            }
                            else
                            {
                                dic_tbl_data_lst[block.Key].Add(dt);
                            }
                        }
                    }
                }
                foreach (var tbl in dic_tbl_data_lst)
                {
                    DataTable block_tbl = Summary_Data_Table_from_list(tbl.Value, spec_dt);
                    result_table_lst.Add(tbl.Key, block_tbl);
                }
                List<string> pcs_lst = new List<string>();
                Get_List_data2(-1, src_data_dt, new string[] { "Pcs_No" }, ref pcs_lst, "Pcs_No", true);
                //pcs_lst = Get_column_name(dic_sum["Before"]);
                if (pcs_lst.Count > 0)
                {
                    DataTable cur_dt = new DataTable();
                    for (int i = 0; i < pcs_lst.Count; i++)
                    {
                        cur_dt.Columns.Add(pcs_lst[i]);
                    }
                    foreach (var tbl in result_table_lst)
                    {
                        string[] bl_name_arr = tbl.Key.Split('-');
                        string bl_name = bl_name_arr[bl_name_arr.Length - 1];
                        for (int r_inx = 0; r_inx < tbl.Value.Rows.Count; r_inx++)
                        {
                            if (cur_dt.Rows.Count < r_inx + 1)
                            {
                                cur_dt.Rows.Add();
                            }
                            for (int c_inx = 0; c_inx < tbl.Value.Columns.Count; c_inx++)
                            {
                                string col_name = tbl.Value.Columns[c_inx].ColumnName + "_" + "BL" + bl_name;
                                if (TDMK_Code.check_exist_list_index2(col_name, pcs_lst) != -1)
                                {
                                    cur_dt.Rows[r_inx][col_name] = tbl.Value.Rows[r_inx][c_inx];
                                }
                            }
                        }
                    }
                    dic_sum[cycle] = cur_dt;
                }
                else
                {
                    int div_factor = result_table_lst.Count;
                    int fact = sel_num / div_factor;
                    List<int> item_qty = new List<int>();
                    List<DataTable> summ_tbl_lst = new List<DataTable>();
                    if (sel_num > 1)
                    {
                        for (int i = 0; i < div_factor - 1; i++)
                        {
                            item_qty.Add(fact);
                        }
                        item_qty.Add(sel_num - (div_factor - 1) * fact);
                    }
                    else
                    {
                        for (int i = 0; i < div_factor; i++)
                        {
                            item_qty.Add(1);
                        }
                    }
                    int inx = 0;
                    foreach (var tbl in result_table_lst)
                    {
                        List<string> sel_before_pcs_lst = Get_column_name(tbl.Value).Select(x => x.ToUpper()).ToList();
                        List<string> sel_pcs_lst = Get_intersec(new List<List<string>>() { sel_before_pcs_lst, after_pcs_lst });
                        DataTable dest_dt = tbl.Value.AsDataView().ToTable(false, sel_pcs_lst.ToArray());
                        List<ECheck_Process.NG_list2> cur_NGList = Get_NG_point_tbl(spec_dt, dest_dt);
                        DataTable sel_dt = Summary_Selected_FromExisted(dest_dt, cur_NGList, item_qty[inx]);
                        summ_tbl_lst.Add(sel_dt);
                        inx++;
                    }
                    DataTable disp_result = new DataTable();
                    string bl_name;
                    for (int i = 0; i < summ_tbl_lst.Count; i++)
                    {
                        string[] bl_name_arr = result_table_lst.Keys.ToList()[i].Split('-');
                        bl_name = bl_name_arr[bl_name_arr.Length - 1];
                        for (int j = 0; j < summ_tbl_lst[i].Columns.Count; j++)
                        {

                            string col_name = summ_tbl_lst[i].Columns[j].ColumnName + "_" + "BL" + bl_name;
                            disp_result.Columns.Add(col_name);
                        }
                    }
                    for (int i = 0; i < summ_tbl_lst.Count; i++)
                    {
                        string[] bl_name_arr = result_table_lst.Keys.ToList()[i].Split('-');
                        bl_name = bl_name_arr[bl_name_arr.Length - 1];
                        for (int r_inx = 0; r_inx < summ_tbl_lst[i].Rows.Count; r_inx++)
                        {
                            if (disp_result.Rows.Count < r_inx + 1)
                            {
                                disp_result.Rows.Add();
                            }
                            for (int c_inx = 0; c_inx < summ_tbl_lst[i].Columns.Count; c_inx++)
                            {
                                string col_name = summ_tbl_lst[i].Columns[c_inx].ColumnName + "_" + "BL" + bl_name;
                                disp_result.Rows[r_inx][col_name] = summ_tbl_lst[i].Rows[r_inx][c_inx];
                            }
                        }
                    }
                    dic_sum[cycle] = disp_result;
                }
            }
            return dic_sum;
        }
        public List<string> Get_column_name(DataTable src_dt)
        {
            List<string> result = new List<string>();
            foreach (DataColumn dc in src_dt.Columns)
            {
                result.Add(dc.ColumnName);
            }
            return result;
        }
        public void Save_logfile_all(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process, string cycle, int PCS_num, string sel_dev, List<string> logfile_lst, string log_locate)
        {
            var temp = Get_BlockofLogFile_Taiyo(logfile_lst).Keys.ToList();
            List<string> before_item_lst = new List<string>();
            bool en_save = Check_Block_Matching(sqlcon_OK2SHIP, ItemCode,LotNo, process, temp, out before_item_lst);
 start_lbl: if (!en_save)
            {
                if (cycle == "Before")
                {
                    if (before_item_lst.Count > 0)
                    {
                        if (MessageBox.Show("Số pcs và block khác với dữ liệu đã lưu. Bạn muốn cập nhật lại toàn bộ dữ liệu?", "Cảnh báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            TDMK_Code.Delelte_FilteredItem_arr(process, sqlcon_OK2SHIP, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                        }
                        else
                        {
                            return;
                        }
                    }
                    //before_item_lst = temp;
                    en_save = true;
                    goto start_lbl;
                }
                else
                {
                    MessageBox.Show("Dữ liệu block không đủ hoặc không đúng", "Thông báo");
                }    
            }
            else
            {    
                foreach (var log_f in logfile_lst)
                {
                    string f_name = Path.Combine(log_locate, log_f);
                    DataTable netSpec_tbl = new DataTable();
                    DataTable result_data = new DataTable();
                    try
                    {
                        switch (sel_dev)
                        {
                            case "YAMAHA":
                                result_data = Yamaha_3GSPD_Process(f_name, ref netSpec_tbl);
                                break;
                            case "TAIYO":
                                result_data = Tayo_3GMRD_Process(f_name, PCS_num, ref netSpec_tbl);
                                break;
                        }
                    }
                    catch
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Lỗi do chọn sai loại máy!", "Thông báo");
                        return;
                    }
                    if (result_data.Rows.Count > 0)
                    {
                        Save_LogFile_detail(sqlcon_OK2SHIP, result_data, ItemCode, LotNo, process, cycle, log_f);
                    }
                    load_spec_from_Logfile(netSpec_tbl, ItemCode, process, sqlcon_OK2SHIP, false);
                }
                MessageBox.Show(new Form { TopMost = true }, "Lưu dữ liệu logfile hoàn thành", "Thông báo");
            }
        }
        public bool Save_logfile_all_en(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process, string cycle, int PCS_num, string sel_dev, List<string> logfile_lst, string log_locate)
        {
            var temp = Get_BlockofLogFile_Taiyo(logfile_lst).Keys.ToList();
            List<string> before_item_lst = new List<string>();
            bool en_save = Check_Block_Matching(sqlcon_OK2SHIP, ItemCode, LotNo, process, temp, out before_item_lst);
            start_lbl: if (!en_save)
            {
                if (cycle == "Before")
                {
                    if (before_item_lst.Count > 0)
                    {
                        if (MessageBox.Show("Số pcs và block khác với dữ liệu đã lưu. Bạn muốn cập nhật lại toàn bộ dữ liệu?", "Cảnh báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            TDMK_Code.Delelte_FilteredItem_arr(process, sqlcon_OK2SHIP, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                            TDMK_Code.Delelte_FilteredItem_arr(process+"_LOGFILE", sqlcon_OK2SHIP, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                        }
                        else
                        {
                            return false;
                        }
                    }
                    //before_item_lst = temp;
                    en_save = true;
                    goto start_lbl;
                }
                else
                {
                    MessageBox.Show("Dữ liệu block không đủ hoặc không đúng", "Thông báo");
                    return false;
                }
            }
            else
            {
                foreach (var log_f in logfile_lst)
                {
                    var f_bl = Get_blockname_from_Logfile(log_f);
        redo_lbl: if(before_item_lst.IndexOf(f_bl)!=-1)
                    {
                        string f_name = Path.Combine(log_locate, log_f);
                        DataTable netSpec_tbl = new DataTable();
                        DataTable result_data = new DataTable();
                        try
                        {
                            switch (sel_dev)
                            {
                                case "YAMAHA":
                                    result_data = Yamaha_3GSPD_Process(f_name, ref netSpec_tbl);
                                    break;
                                case "TAIYO":
                                    result_data = Tayo_3GMRD_Process(f_name, PCS_num, ref netSpec_tbl);
                                    break;
                            }
                        }
                        catch
                        {
                            MessageBox.Show("Lỗi do chọn sai loại máy!", "Thông báo");
                            return false;
                        }
                        if (result_data.Rows.Count > 0)
                        {
                            Save_LogFile_detail(sqlcon_OK2SHIP, result_data, ItemCode, LotNo, process, cycle, log_f);
                        }
                        load_spec_from_Logfile(netSpec_tbl, ItemCode, process, sqlcon_OK2SHIP, false);
                    }
                    else
                    {
                        if((cycle=="Before")||(cycle==""))
                        {
                            before_item_lst.Add(f_bl);
                            goto redo_lbl;
                        }
                    }

                }
                MessageBox.Show("Lưu dữ liệu logfile hoàn thành", "Thông báo");
                return true;
            }
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
        public bool Check_Block_Matching(SqlConnection sqlcon, string ItemCode, string LotNo, string process, List<string> src_block_lst, out List<string> sel_lst)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Cycles_name" }, new string[] { ItemCode, LotNo, "Before" });
            DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, process + "_LOGFILE", filter_str);
            List<string> before_item_lst = new List<string>();
            Get_List_data2(-1, dt, new string[] { "Logfile" }, ref before_item_lst, "Logfile", true);
            before_item_lst = Get_BlockofLogFile_Taiyo(before_item_lst).Keys.ToList();
            var lst = Get_intersec(new List<List<string>>() { src_block_lst, before_item_lst });
            if(lst.Count==before_item_lst.Count)
            {
                sel_lst = lst;
                return true;
            }
            else
            {
                sel_lst = src_block_lst;
                return false;
            }
        }
        public List<string> Get_block_from_Report_table(SqlConnection sqlcon, string ItemCode, string LotNo, string process)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, process, filter_str);
            List<string> pcs_lst = dt.AsEnumerable().Select(x => x.Field<string>("Pcs_No")).Distinct().ToList();
            List<string> block_lst = pcs_lst.Select(x => x.Split('_').LastOrDefault().Replace("BL","")).Distinct().ToList();
            return block_lst;
        }
        public List<string>Get_block_from_logfile_table(SqlConnection sqlcon, string ItemCode, string LotNo, string process)
        {
            List<string> logfile_lst = new List<string>();
            DataTable logfile_dt = TDMK_Code.Datatable_Filter(sqlcon, process + "_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Cycles_name" }, new string[] { ItemCode, LotNo, "Before" }));
            Get_List_data2(-1, logfile_dt, new string[] { "Logfile" }, ref logfile_lst, "Logfile", true);
            List<string> _block_of_logfile = Get_BlockofLogFile_Taiyo(logfile_lst).Keys.ToList();
            List<string> block_of_logfile = _block_of_logfile.Select(x => x.Split('-').LastOrDefault()).Distinct().ToList();
            return block_of_logfile;
        }
    }
}
