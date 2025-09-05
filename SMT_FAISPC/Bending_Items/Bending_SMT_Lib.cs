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
using myExcel = Microsoft.Office.Interop.Excel;
using IniLibs;
using OfficeOpenXml;

namespace Bending_Items
{
    public class Bending_SMT_Lib
    {
        SEI_Lib myCode = new SEI_Lib();
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        IniFile TDMK_init;// = new IniFile("Config.ini");
        ECheck_Process exp_process = new ECheck_Process();
        string Data_Location;
        string format_folder;
        string log_folder;
        string Report_location;
        public struct bending_data
        {
            public string before_data;
            public string after_data;
            public string resistance_vary;
            public bending_data(string in_before, string in_after, string in_res_vary)
            {
                before_data = in_before;
                after_data = in_after;
                resistance_vary = in_res_vary;
            }
        }
        public struct Log_bending_data
        {
            public string log_time;
            public List<string> log_data;
            public string ItemName;
            public string UserID;
            public Log_bending_data(string in_log_time, List<string> in_log_data, string in_ItemName, string in_UserID)
            {
                log_time = in_log_time;
                log_data = in_log_data;
                ItemName = in_ItemName;
                UserID = in_UserID;
            }
        }
        public struct Cell_data_details
        {
            public string before_edit_data { get; set; }
            public string after_edit_data { get; set; }
            public string Net_No;
            public string USL;
            public string LSL;
            public Cell_data_details(string in_before, string in_after, string in_Net, string in_USL, string in_LSL)
            {
                before_edit_data = in_before;
                after_edit_data = in_after;
                Net_No = in_Net;
                USL = in_USL;
                LSL = in_LSL;
            }
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
        public DataTable Bending_result(string logfile)
        {
            DataTable dt = CSV_To_DataTable(logfile);
            List<string> Net_name = new List<string>();
            foreach (DataColumn dc in dt.Columns)
            {
                string col_name = dc.ColumnName;
                if (!col_name.Contains("SUS") && !col_name.Contains("OPEN") && col_name.Contains("("))
                {
                    Net_name.Add(col_name);
                }
            }
            DataTable filter_dt = dt.AsEnumerable().Where(x => x.Field<string>("UUT") != "No_UUT_ID" && x.Field<string>("Fails") == "0").CopyToDataTable();
            return filter_dt.AsDataView().ToTable(false, Net_name.ToArray());
        }
        public DataTable Spec_from_Logfile(string logfile)
        {
            DataTable dt = CSV_To_DataTable(logfile);
            List<string> Net_name = new List<string>();
            foreach (DataColumn dc in dt.Columns)
            {
                string col_name = dc.ColumnName;
                if (!col_name.Contains("SUS") && !col_name.Contains("OPEN") && col_name.Contains("("))
                {
                    Net_name.Add(col_name);
                }
            }
            DataTable filter_dt = dt.AsEnumerable().Where(x => !x.Field<string>("Test Title").Contains("*") && x.Field<string>("Test Title") != "Type").CopyToDataTable();
            return filter_dt.AsDataView().ToTable(false, Net_name.ToArray());
        }
        public DataTable CSV_To_DataTable(string _filename)
        {
            string tar_file;
            char[] split_char = { '\t', ',' };
            List<string> src_data = new List<string>();
            tar_file = _filename;
            System.IO.StreamReader file = new System.IO.StreamReader(tar_file);
            string newline;
            while ((newline = file.ReadLine()) != null)
            {
                if (newline.Contains('"'))
                {
                    string[] temp = newline.Split('"');
                    temp[1] = temp[1].Split('[', ']')[0];
                    src_data.Add(string.Concat(temp));
                }
                else
                {
                    src_data.Add(newline);
                }
            }

            string[] columnnames = src_data[0].Split(split_char);
            DataTable dt = new DataTable();
            foreach (string c in columnnames)
            {
                dt.Columns.Add(c);
            }
            for (int j = 1; j < src_data.Count; j++)
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
        public DataTable DAT_To_DataTable(string _filename)
        {
            string tar_file;
            char[] split_char = { '\t', ',' };
            List<string> src_data = new List<string>();
            tar_file = _filename;
            System.IO.StreamReader file = new System.IO.StreamReader(tar_file);
            string newline;
            while ((newline = file.ReadLine()) != null)
            {
                if (newline.Contains('"'))
                {
                    string[] temp = newline.Split('"');
                    temp[1] = temp[1].Split('[', ']')[0];
                    src_data.Add(string.Concat(temp));
                }
                else
                {
                    src_data.Add(newline);
                }
            }

            string[] columnnames = src_data[0].Split(split_char);
            DataTable dt = new DataTable();
            foreach (string c in columnnames)
            {
                if (c != "")
                {
                    dt.Columns.Add(c);
                }
            }
            bool rec_en = false;
            bool proc_en = false;
            Dictionary<int, List<string>> Raw_Bending_data = new Dictionary<int, List<string>>();
            int item_count = 0;
            for (int j = 2; j < src_data.Count; j++)
            {
                string[] values = src_data[j].Split(split_char);
                if (values.Length > dt.Columns.Count)
                {
                    rec_en = true;
                    if (proc_en)
                    {
                        proc_en = false;
                    }
                    item_count++;
                    Raw_Bending_data.Add(item_count, new List<string>() { values[5] });
                }
                else
                {
                    if (rec_en)
                    {
                        if (!proc_en)
                        {
                            proc_en = true;
                        }
                    }
                    else
                    {
                        proc_en = false;
                    }
                }
                if (proc_en)
                {
                    if (values.Length == dt.Columns.Count)
                    {
                        Raw_Bending_data[item_count].Add(src_data[j]);
                    }
                }
            }
            file.Close();
            dt.Columns.Add("UUT");
            foreach (var data in Raw_Bending_data.Values)
            {
                if (!data[0].Contains("No_UUT_ID"))
                {
                    for (int k = 1; k < data.Count; k++)
                    {
                        if (!data[k].Contains("OPEN") && !data[k].Contains("SUS"))
                        {
                            DataRow dr = dt.NewRow();
                            string[] temp_arr = data[k].Split(split_char);
                            for (int x = 0; x < temp_arr.Length; x++)
                            {
                                dr[x] = temp_arr[x];
                            }
                            dr[temp_arr.Length] = data[0];
                            dt.Rows.Add(dr);
                        }
                    }
                }
            }
            return dt;
        }
        public Dictionary<string, Dictionary<int, bending_data>> Bending_process(DataTable dt_before, DataTable dt_after)
        {
            Dictionary<string, Dictionary<int, bending_data>> result = new Dictionary<string, Dictionary<int, bending_data>>();
            List<string> bef_lst = Get_column_name(dt_before);
            List<string> aft_lst = Get_column_name(dt_after);
            var net_lst = bef_lst.Intersect(aft_lst).ToList();
            int row_inx = Math.Min(dt_before.Rows.Count, dt_after.Rows.Count);
            foreach (var net in net_lst)
            {
                result.Add(net, new Dictionary<int, bending_data>());
                for (int r = 0; r < row_inx; r++)
                {
                    string bef_val = dt_before.Rows[r][net].ToString();
                    string aft_val = dt_after.Rows[r][net].ToString();
                    result[net].Add(r + 1, new bending_data(bef_val, aft_val, Resistance_vary(bef_val, aft_val)));
                }
            }
            return result;
        }
        public string Resistance_vary(string bef_val, string aft_val)
        {
            string result = "Data Error";
            try
            {
                double before = Convert.ToDouble(bef_val);
                double after = Convert.ToDouble(aft_val);
                double _result = Math.Abs(before - after) * 100 / before;
                result = _result.ToString("#0.000");
            }
            catch
            {

            }
            return result;
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
        public DataTable DAT_To_DataTable_details_old(string _filename, ref List<string> Net_list, ref DataTable tar_spec_dt)
        {
            string tar_file;
            char[] split_char = { '\t', ',' };
            List<string> src_data = new List<string>();
            tar_file = _filename;
            System.IO.StreamReader file = new System.IO.StreamReader(tar_file);
            string newline;
            while ((newline = file.ReadLine()) != null)
            {
                if (newline.Contains('"'))
                {
                    string[] temp = newline.Split('"');
                    temp[1] = temp[1].Split('[', ']')[0];
                    src_data.Add(string.Concat(temp));
                }
                else
                {
                    src_data.Add(newline);
                }
            }
            List<string> columnnames = src_data[0].Split(split_char).Where(x => x != "").Select(x => x.Trim()).ToList();
            //Net_list = columnnames;
            int col_num = columnnames.Count;
            bool rec_en = false;
            bool proc_en = false;
            Dictionary<int, List<string>> Raw_Bending_data = new Dictionary<int, List<string>>();
            int item_count = 0;
            for (int j = 2; j < src_data.Count; j++)
            {
                string[] values = src_data[j].Split(split_char);
                if (values.Length > col_num)
                {
                    rec_en = true;
                    if (proc_en)
                    {
                        proc_en = false;
                    }
                    item_count++;
                    Raw_Bending_data.Add(item_count, new List<string>() { values[5] });
                }
                else
                {
                    if (rec_en)
                    {
                        if (!proc_en)
                        {
                            proc_en = true;
                        }
                    }
                    else
                    {
                        proc_en = false;
                    }
                }
                if (proc_en)
                {
                    if (values.Length == col_num)
                    {
                        Raw_Bending_data[item_count].Add(src_data[j]);
                    }
                }
            }
            file.Close();
            DataTable result_dt = new DataTable();
            DataTable spec_dt = new DataTable();
            List<string> spec_col_lst = new List<string>() { "Net_Name", "Point+V", "Point-V", "LSL", "USL" };
            foreach (string col in spec_col_lst)
            {
                spec_dt.Columns.Add(col);
            }
            foreach (var data in Raw_Bending_data.Values)
            {
                if (!data[0].Contains("No_UUT_ID"))
                {
                    for (int k = 1; k < data.Count; k++)
                    {
                        if (!data[k].Contains("OPEN") && !data[k].Contains("SUS"))
                        {
                            string col_name = data[0];
                            string[] detail_data = data[k].Split(split_char);

                            if (!myCode.check_columns_existed(result_dt, col_name))
                            {
                                result_dt.Columns.Add(col_name);
                            }
                            if (result_dt.Rows.Count < k)
                            {
                                result_dt.Rows.Add();
                            }
                            result_dt.Rows[k - 1][col_name] = detail_data[5];
                            if (Net_list.IndexOf(detail_data[6]) == -1)
                            {
                                Net_list.Add(detail_data[6]);
                                string net = detail_data[6];
                                string ul = detail_data[7];
                                string ll = detail_data[8];
                                string[] points = net.Split('(').First().Replace("S", "").Split('-');
                                string point1 = points.First();
                                string point2 = points.Last();
                                spec_dt.Rows.Add(net, point1, point2, ul, ll);
                            }
                        }
                    }
                }
            }
            tar_spec_dt = spec_dt;
            return result_dt;

        }
        public DataTable DAT_To_DataTable_details(string _filename, ref List<string> Net_list, ref DataTable tar_spec_dt)
        {
            string tar_file;
            char[] split_char = { '\t', ',' };
            List<string> src_data = new List<string>();
            tar_file = _filename;
            System.IO.StreamReader file = new System.IO.StreamReader(tar_file, Encoding.UTF8);
            string newline;
            while ((newline = file.ReadLine()) != null)
            {
                if (newline.Contains('"'))
                {
                    string[] temp = newline.Split('"');
                    temp[1] = temp[1].Split('[', ']')[0];
                    src_data.Add(string.Concat(temp));
                }
                else
                {
                    src_data.Add(newline);
                }
            }
            List<string> columnnames = src_data[0].Split(split_char).Where(x => x != "").Select(x => x.Trim()).ToList();
            int col_num = columnnames.Count;
            bool rec_en = false;
            bool proc_en = false;
            Dictionary<int, List<string>> Raw_Bending_data = new Dictionary<int, List<string>>();
            int item_count = 0;
            for (int j = 2; j < src_data.Count; j++)
            {
                string[] values = src_data[j].Split(split_char);
                if (values.Length > col_num)
                {
                    string fail_val = Extract_Num_from_String(values[2]);
                    if (fail_val == "0")
                    {
                        rec_en = true;
                        if (proc_en)
                        {
                            proc_en = false;
                        }
                        item_count++;
                        Raw_Bending_data.Add(item_count, new List<string>() { values[5].Trim() });
                    }
                    else
                    {
                        rec_en = false;
                        proc_en = false;
                    }
                }
                else
                {
                    if (rec_en)
                    {
                        if (!proc_en)
                        {
                            proc_en = true;
                        }
                    }
                    else
                    {
                        proc_en = false;
                    }
                }
                if (proc_en)
                {
                    if (values.Length == col_num)
                    {
                        Raw_Bending_data[item_count].Add(src_data[j]);
                    }
                }
            }
            file.Close();
            DataTable result_dt = new DataTable();
            DataTable spec_dt = new DataTable();
            try
            {
                List<string> spec_col_lst = new List<string>() { "Net_Name", "Point+V", "Point-V", "LSL", "USL" };
                foreach (string col in spec_col_lst)
                {
                    spec_dt.Columns.Add(col);
                }
                foreach (var data in Raw_Bending_data.Values)
                {
                    if (!data[0].Contains("No_UUT_ID"))
                    {
                        int r_count = 0;
                        for (int k = 1; k < data.Count; k++)
                        {
                            if (!data[k].Contains(">"))
                            {
                                r_count++;
                                string col_name = data[0];
                                string[] detail_data = data[k].Split(split_char);
                                if (!myCode.check_columns_existed(result_dt, col_name))
                                {
                                    result_dt.Columns.Add(col_name);
                                }
                                if (result_dt.Rows.Count < r_count)
                                {
                                    result_dt.Rows.Add();
                                }
                                result_dt.Rows[r_count - 1][col_name] = detail_data[5];
                                if (Net_list.IndexOf(detail_data[6]) == -1)
                                {
                                    Net_list.Add(detail_data[6].Trim());
                                    string net = detail_data[6].Trim();
                                    string ul = detail_data[7].Trim();
                                    string ll = detail_data[8].Trim();
                                    string[] points = Get_Net_Points(net, new char[] { ' ', '(', '-', '_' });// net.Split('(').First().Replace("S", "").Split('-');
                                    string point1 = points.First();
                                    string point2 = points.Last();
                                    spec_dt.Rows.Add(net, point1, point2, ul, ll);
                                }
                            }
                        }
                    }
                }
                tar_spec_dt = spec_dt;
            }
            catch
            {
                MessageBox.Show(new Form { TopMost = true }, "Lỗi tên chân Pin", "Thông báo");
            }
            return result_dt;
        }
        public DataTable DAT_To_DataTable_details_time(string _filename, ref List<string> Net_list, ref DataTable tar_spec_dt, ref Dictionary<string, string> Log_info)
        {
            char[] split_char = { '\t', ',' };
            Dictionary<int, Log_bending_data> Raw_Bending_data = Get_Raw_Bending_data(_filename);
            DataTable result_dt = new DataTable();
            DataTable spec_dt = new DataTable();
            try
            {
                List<string> spec_col_lst = new List<string>() { "Net_Name", "Point+V", "Point-V", "LSL", "USL" };
                foreach (string col in spec_col_lst)
                {
                    spec_dt.Columns.Add(col);
                }
                int UUT_inx = 0;
                foreach (var data in Raw_Bending_data.Values)
                {
                    if (!data.log_data[0].Contains("No_UUT_ID"))
                    {
                        int r_count = 0;
                        string col_name = data.log_data[0];
                        if (myCode.check_columns_existed(result_dt, col_name))
                        {
                            col_name = col_name + "_" + (UUT_inx + 1).ToString();
                            UUT_inx++;
                        }
                        for (int k = 1; k < data.log_data.Count; k++)
                        {
                            if (!data.log_data[k].Contains(">"))
                            {
                                string[] detail_data = data.log_data[k].Split(split_char);
                                if (!myCode.check_columns_existed(result_dt, col_name))
                                {
                                    result_dt.Columns.Add(col_name);
                                }
                                int r = Net_list.IndexOf(detail_data[6]);
                                if (r == -1)
                                {
                                    if (result_dt.Rows.Count < r_count + 1)
                                    {
                                        result_dt.Rows.Add();
                                    }
                                    result_dt.Rows[r_count][col_name] = detail_data[5];
                                    r_count++;
                                    Net_list.Add(detail_data[6].Trim());
                                    string net = detail_data[6].Trim();
                                    string ul = detail_data[7].Trim();
                                    string ll = detail_data[8].Trim();
                                    string[] points = Get_Net_Points(net, new char[] { ' ', '(', '-', '_' });// net.Split('(').First().Replace("S", "").Split('-');
                                    string point1 = points.FirstOrDefault();
                                    string point2 = points.LastOrDefault();
                                    spec_dt.Rows.Add(net, point1, point2, ul, ll);
                                }
                                else
                                {
                                    result_dt.Rows[r][col_name] = detail_data[5];
                                }
                            }
                        }
                        if (Log_info.Keys.ToList().IndexOf(data.log_data[0]) == -1)
                        {
                            Log_info.Add(data.log_data[0], data.log_time);
                        }
                    }
                }
                tar_spec_dt = spec_dt;
            }
            catch
            {
                MessageBox.Show(new Form { TopMost = true }, "Lỗi tên chân Pin", "Thông báo");
            }
            return result_dt;
        }
        public DataTable DAT_To_DataTable_details_time(string _filename, ref List<string> Net_list, ref DataTable tar_spec_dt, ref Dictionary<string, string> Log_info, ref string ItemName, ref string UserId)
        {
            char[] split_char = { '\t', ',' };
            Dictionary<int, Log_bending_data> Raw_Bending_data = Get_Raw_Bending_data(_filename);
            DataTable result_dt = new DataTable();
            DataTable spec_dt = new DataTable();
            List<string> Ignored_NET_lst = new List<string>() { "SHORT", "OPEN" };
            try
            {
                List<string> spec_col_lst = new List<string>() { "Net_Name", "Point+V", "Point-V", "LSL", "USL" };
                foreach (string col in spec_col_lst)
                {
                    spec_dt.Columns.Add(col);
                }
                int UUT_inx = 0;
                foreach (var data in Raw_Bending_data.Values)
                {
                    if (!data.log_data[0].Contains("No_UUT_ID"))
                    {
                        int r_count = 0;
                        string col_name = data.log_data[0];
                        List<string> result_col_lst = Get_column_name(result_dt);
                        UUT_inx = result_col_lst.Where(x => x.Contains(col_name)).Count();
                        if (UUT_inx != 0)
                        {
                            col_name = col_name + "_" + UUT_inx.ToString();
                        }
                        for (int k = 1; k < data.log_data.Count; k++)
                        {
                            if (!data.log_data[k].Contains(">"))
                            {
                                string[] detail_data = data.log_data[k].Split(split_char);
                                if (!myCode.check_columns_existed(result_dt, col_name))
                                {
                                    result_dt.Columns.Add(col_name);
                                }
                                int r = Net_list.IndexOf(detail_data[6]);
                                if (r == -1)
                                {
                                    if (Ignored_NET_lst.IndexOf(detail_data[6].ToUpper()) == -1)
                                    {
                                        if (result_dt.Rows.Count < r_count + 1)
                                        {
                                            result_dt.Rows.Add();
                                        }
                                        result_dt.Rows[r_count][col_name] = detail_data[5];
                                        r_count++;
                                        Net_list.Add(detail_data[6].Trim());
                                        string net = detail_data[6].Trim();
                                        string ul = detail_data[7].Trim();
                                        string ll = detail_data[8].Trim();
                                        string[] points = Get_Net_Points(net, new char[] { ' ', '(', '-', '_' });// net.Split('(').First().Replace("S", "").Split('-');
                                        string point1 = points.FirstOrDefault();
                                        string point2 = points.LastOrDefault();
                                        spec_dt.Rows.Add(net, point1, point2, ul, ll);
                                    }
                                }
                                else
                                {
                                    result_dt.Rows[r][col_name] = detail_data[5];
                                }
                            }
                        }
                        if (Log_info.Keys.ToList().IndexOf(data.log_data[0]) == -1)
                        {
                            Log_info.Add(data.log_data[0], data.log_time);
                        }
                        if (ItemName == "")
                        {
                            ItemName = data.ItemName;
                        }
                        if (UserId == "")
                        {
                            UserId = data.UserID;
                        }
                    }
                }
                tar_spec_dt = spec_dt;
            }
            catch
            {
                MessageBox.Show(new Form { TopMost = true }, "Lỗi tên chân Pin", "Thông báo");
            }
            return result_dt;
        }
        public Dictionary<int, Log_bending_data> Get_Raw_Bending_data(string tar_file)
        {
            char[] split_char = { '\t', ',' };
            List<string> src_data = new List<string>();
            System.IO.StreamReader file = new System.IO.StreamReader(tar_file, Encoding.UTF8);
            string newline;
            while ((newline = file.ReadLine()) != null)
            {
                if (newline.Contains('"'))
                {
                    string[] temp = newline.Split('"');
                    temp[1] = temp[1].Split('[', ']')[0];
                    src_data.Add(string.Concat(temp));
                }
                else
                {
                    if (newline != "")
                    {
                        src_data.Add(newline);
                    }
                }
            }
            List<string> columnnames = src_data[0].Split(split_char).Where(x => x != "").Select(x => x.Trim()).ToList();
            int col_num = columnnames.Count;
            bool rec_en = false;
            bool proc_en = false;
            Dictionary<int, Log_bending_data> Raw_Bending_data = new Dictionary<int, Log_bending_data>();
            int item_count = 0;
            for (int j = 2; j < src_data.Count; j++)
            {
                string[] values = src_data[j].Split(split_char);
                if (values.Length > col_num)
                {
                    string fail_val = Extract_Num_from_String(values[2]);
                    if (fail_val == "0")
                    {
                        rec_en = true;
                        if (proc_en)
                        {
                            proc_en = false;
                        }
                        item_count++;
                        string itemname = values[0].Split('-').FirstOrDefault();
                        string userid = values[10];
                        Raw_Bending_data.Add(item_count, new Log_bending_data(values[1].Trim(), new List<string>() { values[5].Trim() }, itemname, userid));
                    }
                    else
                    {
                        rec_en = false;
                        proc_en = false;
                    }
                }
                else
                {
                    if (rec_en)
                    {
                        if (!proc_en)
                        {
                            proc_en = true;
                        }
                    }
                    else
                    {
                        proc_en = false;
                    }
                }
                if (proc_en)
                {
                    if (values.Length == col_num)
                    {
                        Raw_Bending_data[item_count].log_data.Add(src_data[j]);
                    }
                }
            }
            file.Close();
            return Raw_Bending_data;
        }
        public void Marking_DGV_Rows_Net(DataGridView src_DGV, List<string> src_Net = null)
        {
            int r = 0;
            foreach (DataGridViewRow dgr in src_DGV.Rows)
            {
                if (src_Net != null)
                {
                    if (r < src_Net.Count)
                    {
                        dgr.HeaderCell.Value = src_Net[r];
                    }
                    else
                    {
                        dgr.HeaderCell.Value = (r + 1).ToString();
                    }
                }
                else
                {
                    dgr.HeaderCell.Value = (r + 1).ToString();
                }
                r++;
            }
            src_DGV.AutoResizeColumns();
            src_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            myCode.Disable_Sort_DGV(src_DGV);
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
        public DataTable Save_LogFile_detail(SqlConnection sqlcon_OK2SHIP, DataTable logfile_tbl, string ItemCode, string LotNo, string process_name, string cycle_name, string Logfile_name, List<string> Net_lst)
        {
            process_name = process_name + "_LOGFILE";
        start_lbl: string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Logfile", "Cycles_name" }, new string[] { ItemCode, LotNo, Logfile_name, cycle_name });
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
                int col_inx = 0;
                foreach (DataColumn dc in logfile_tbl.Columns)
                {
                    int r_inx = 0;
                    foreach (DataRow dr in logfile_tbl.Rows)
                    {
                        id++;
                        string net = (r_inx + 1).ToString();
                        if (r_inx < Net_lst.Count)
                        {
                            net = Net_lst[r_inx];
                        }
                        string pcs = dc.ColumnName;
                        TDMK_Code.insert_val_arr2(process_name, sqlcon_OK2SHIP, new string[] { "ID", "ItemCode", "LotNo", "Net_No", "Pcs_No", "Logfile", "Cycles_name", "Data" }, new string[] { id.ToString(), ItemCode, LotNo, net, pcs, Logfile_name, cycle_name, dr[dc].ToString() });
                        r_inx++;
                    }
                    col_inx++;
                }
            }
            return TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
        }
        public DataTable Save_LogFile_detail_time(SqlConnection sqlcon_OK2SHIP, DataTable logfile_tbl, string ItemCode, string LotNo, string process_name, string cycle_name, string Logfile_name, List<string> Net_lst, Dictionary<string, string> Log_details)
        {
            process_name = process_name + "_LOGFILE";
            string _Logfile_name = "%" + Path.GetFileNameWithoutExtension(Logfile_name) + "%";
        start_lbl: string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Logfile", "Cycles_name" }, new string[] { ItemCode, LotNo, _Logfile_name.ToUpper(), cycle_name });
            DataTable temp_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
            if (temp_dt.Rows.Count > 0)
            {
                if (MessageBox.Show("Data of log file: " + Logfile_name + " is existed in Database. Do you want to update again?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
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
                //int col_inx = 0;
                foreach (DataColumn dc in logfile_tbl.Columns)
                {
                    int r_inx = 0;
                    string pcs = dc.ColumnName;
                    string log_name = Logfile_name;
                    string[] temp_pcs = pcs.Split('_');
                    string log_time = Log_details[temp_pcs.First()];
                    if (temp_pcs.Length > 1)
                    {
                        log_name = Path.GetFileNameWithoutExtension(Logfile_name) + "_" + temp_pcs.Last() + Path.GetExtension(Logfile_name);
                        pcs = temp_pcs.First();
                    }
                    foreach (DataRow dr in logfile_tbl.Rows)
                    {
                        id++;
                        string net = (r_inx + 1).ToString();
                        if (r_inx < Net_lst.Count)
                        {
                            net = Net_lst[r_inx];
                        }
                        TDMK_Code.insert_val_arr2(process_name, sqlcon_OK2SHIP, new string[] { "ID", "ItemCode", "LotNo", "Net_No", "Pcs_No", "Logfile", "Cycles_name", "Data", "Remark" }, new string[] { id.ToString(), ItemCode, LotNo, net, pcs, log_name.ToUpper(), cycle_name, dr[dc].ToString(), log_time });
                        r_inx++;
                    }
                    //col_inx++;
                }
            }
            return TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
        }
        public DataTable Save_LogFile_detail_time(SqlConnection sqlcon_OK2SHIP, DataTable logfile_tbl, string ItemCode, string LotNo, string process_name, string cycle_name, string shift, string userid, string itemname, string Logfile_name, List<string> Net_lst, Dictionary<string, string> Log_details)
        {
            process_name = process_name + "_LOGFILE";
            string _Logfile_name = "%" + Path.GetFileNameWithoutExtension(Logfile_name) + "%";
        start_lbl: string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Logfile", "Cycles_name", "Shift" }, new string[] { ItemCode, LotNo, _Logfile_name.ToUpper(), cycle_name, shift });
            DataTable temp_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
            if (temp_dt.Rows.Count > 0)
            {
                if (MessageBox.Show("Data of log file: " + Logfile_name + " is existed in Database. Do you want to update again?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
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
                        foreach (DataRow dr in temp_dt.Rows)
                        {
                            dr["Shift"] = shift;
                            dr["UserID"] = userid;
                            dr["ItemName"] = itemname;
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
                //int col_inx = 0;
                List<string> col_lst = Get_column_name(logfile_tbl).Where(x => !x.Contains("_")).ToList();
                foreach (DataColumn dc in logfile_tbl.Columns)
                {
                    int r_inx = 0;
                    string pcs = dc.ColumnName;
                    string log_name = Logfile_name;
                    string[] temp_pcs = pcs.Split('_');
                    string log_time = Log_details[temp_pcs.First()];
                    if (col_lst.Count == 1)
                    {
                        if (temp_pcs.Length > 1)
                        {
                            log_name = Path.GetFileNameWithoutExtension(Logfile_name) + "_" + temp_pcs.Last() + Path.GetExtension(Logfile_name);
                            pcs = temp_pcs.First();
                        }
                    }
                    foreach (DataRow dr in logfile_tbl.Rows)
                    {
                        id++;
                        string net = (r_inx + 1).ToString();
                        if (r_inx < Net_lst.Count)
                        {
                            net = Net_lst[r_inx];
                        }
                        TDMK_Code.insert_val_arr2(process_name, sqlcon_OK2SHIP, new string[] { "ID", "ItemCode", "LotNo", "Net_No", "Pcs_No", "Logfile", "Cycles_name", "Data", "Remark", "Shift", "UserID", "ItemName" }, new string[] { id.ToString(), ItemCode, LotNo, net, pcs, log_name.ToUpper(), cycle_name, dr[dc].ToString(), log_time, shift, userid, itemname });
                        r_inx++;
                    }
                    //col_inx++;
                }
            }
            return TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
        }
        public bool Save_LogFile_detail_time_result(SqlConnection sqlcon_OK2SHIP, DataTable logfile_tbl, string ItemCode, string LotNo, string process_name, string cycle_name, string shift, string userid, string itemname, string Logfile_name, List<string> Net_lst, Dictionary<string, string> Log_details, string season = "")
        {
            try
            {
                process_name = process_name + "_LOGFILE" + season;
                string _Logfile_name = "%" + Path.GetFileNameWithoutExtension(Logfile_name) + "%";
            start_lbl: string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Logfile", "Cycles_name", "Shift" }, new string[] { ItemCode, LotNo, _Logfile_name.ToUpper(), cycle_name, shift });
                DataTable temp_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                if (temp_dt.Rows.Count > 0)
                {
                    if (MessageBox.Show("Data of log file: " + Logfile_name + " is existed in Database. Do you want to update again?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
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
                            foreach (DataRow dr in temp_dt.Rows)
                            {
                                dr["Shift"] = shift;
                                dr["UserID"] = userid;
                                dr["ItemName"] = itemname;
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
                    //int col_inx = 0;
                    List<string> col_lst = Get_column_name(logfile_tbl).Where(x => !x.Contains("_")).ToList();
                    foreach (DataColumn dc in logfile_tbl.Columns)
                    {
                        int r_inx = 0;
                        string pcs = dc.ColumnName;
                        string log_name = Logfile_name;
                        string[] temp_pcs = pcs.Split('_');
                        string log_time = Log_details[temp_pcs.First()];
                        if (col_lst.Count == 1)
                        {
                            if (temp_pcs.Length > 1)
                            {
                                log_name = Path.GetFileNameWithoutExtension(Logfile_name) + "_" + temp_pcs.Last() + Path.GetExtension(Logfile_name);
                                pcs = temp_pcs.First();
                            }
                        }
                        foreach (DataRow dr in logfile_tbl.Rows)
                        {
                            id++;
                            string net = (r_inx + 1).ToString();
                            if (r_inx < Net_lst.Count)
                            {
                                net = Net_lst[r_inx];
                            }
                            TDMK_Code.insert_val_arr2(process_name, sqlcon_OK2SHIP, new string[] { "ID", "ItemCode", "LotNo", "Net_No", "Pcs_No", "Logfile", "Cycles_name", "Data", "Remark", "Shift", "UserID", "ItemName" }, new string[] { id.ToString(), ItemCode, LotNo, net, pcs, log_name.ToUpper(), cycle_name, dr[dc].ToString(), log_time, shift, userid, itemname });
                            r_inx++;
                        }
                        //col_inx++;
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Thong bao");
                return false;
            }
        }
        public DataTable Load_Log_Data(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string cycle_name, string logfile_name, List<string> Net_lst_name)
        {
            DataTable dgv_dt = new DataTable();
            List<DataTable> myLst_tbl = new List<DataTable>();
            DataSet DS = new DataSet();
            process_name = process_name + "_LOGFILE";
            if (cycle_name == "")
            {
                TDMK_Code.fill_dataset_Filter_arr(DS, sqlcon_OK2SHIP, process_name, new string[] { "ItemCode", "LotNo", "Logfile" }, new string[] { ItemCode, LotNo, logfile_name });
            }
            else
            {
                TDMK_Code.fill_dataset_Filter_arr(DS, sqlcon_OK2SHIP, process_name, new string[] { "ItemCode", "LotNo", "Cycles_name", "Logfile" }, new string[] { ItemCode, LotNo, cycle_name, logfile_name });
            }

            DataTable dt = DS.Tables[0];
            List<string> col_lst = new List<string>();
            Get_List_data(-1, dt, new string[] { "ItemCode", "LotNo" }, ref col_lst, "Pcs_No", true);
            foreach (string col in col_lst)
            {
                if (!myCode.check_columns_existed(dgv_dt, col))
                {
                    dgv_dt.Columns.Add(col);
                }
            }
            Get_ListTable(-1, dt, new string[] { "ItemCode", "LotNo", "Net_No" }, ref myLst_tbl, "Data");
            foreach (DataTable temp_dt in myLst_tbl)
            {
                Net_lst_name.Add(temp_dt.Rows[0]["Net_No"].ToString());
                List<string> temp_data = new List<string>();
                Get_List_data(-1, temp_dt, new string[] { "ItemCode", "LotNo" }, ref temp_data, "Data", false);
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
        public DataTable Load_Log_Data(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string cycle_name, string shift, ref string userid, ref string itemname, string logfile_name, List<string> Net_lst_name)
        {
            DataTable dgv_dt = new DataTable();
            List<DataTable> myLst_tbl = new List<DataTable>();
            DataSet DS = new DataSet();
            process_name = process_name + "_LOGFILE";
            if (cycle_name == "")
            {
                TDMK_Code.fill_dataset_Filter_arr(DS, sqlcon_OK2SHIP, process_name, new string[] { "ItemCode", "LotNo", "Logfile" }, new string[] { ItemCode, LotNo, logfile_name });
            }
            else
            {
                TDMK_Code.fill_dataset_Filter_arr(DS, sqlcon_OK2SHIP, process_name, new string[] { "ItemCode", "LotNo", "Cycles_name", "Shift", "Logfile" }, new string[] { ItemCode, LotNo, cycle_name, shift, logfile_name });
            }

            DataTable dt = DS.Tables[0];
            if (dt.Rows.Count > 0)
            {
                userid = myCode.checkDBNull(dt.Rows[0]["UserID"]);
                itemname = myCode.checkDBNull(dt.Rows[0]["ItemName"]);
            }
            List<string> col_lst = new List<string>();
            Get_List_data(-1, dt, new string[] { "ItemCode", "LotNo" }, ref col_lst, "Pcs_No", true);
            foreach (string col in col_lst)
            {
                if (!myCode.check_columns_existed(dgv_dt, col))
                {
                    dgv_dt.Columns.Add(col);
                }
            }
            Get_ListTable(-1, dt, new string[] { "ItemCode", "LotNo", "Net_No" }, ref myLst_tbl, "Data");
            foreach (DataTable temp_dt in myLst_tbl)
            {
                Net_lst_name.Add(temp_dt.Rows[0]["Net_No"].ToString());
                List<string> temp_data = new List<string>();
                Get_List_data(-1, temp_dt, new string[] { "ItemCode", "LotNo" }, ref temp_data, "Data", false);
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
        public Dictionary<string, List<string>> Get_BlockofLogFile(List<string> logfile_lst)
        {
            Dictionary<string, List<string>> block_logfile = new Dictionary<string, List<string>>();
            char[] split_char = new char[] { ' ', '(', '_' };
            foreach (string t in logfile_lst)
            {
                string bl = Path.GetFileNameWithoutExtension(t);

                string block_name = bl.Split('-').Last().Split(split_char).First();
                //string num = Extract_Num_from_String(block_name);
                //string bl1 = block_name.Substring(0, block_name.Length - num.Length);
                if (block_name.ToUpper().Contains("BF"))
                {
                    block_name = "BF";
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
                        string USL = spec_dt.Rows[r_inx]["USL"].ToString();
                        string LSL = spec_dt.Rows[r_inx]["LSL"].ToString();
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
        public List<string> Sum_OK_Item(SqlConnection sqlcon_OK2SHIP, DataTable src_tbl, string ItemCode, string process_name, double tar_CPK)
        {
            List<string> result = new List<string>();
            DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process_name }));
            List<int> CPK_NG_NET_lst = Get_NG_CPK_NET(sqlcon_OK2SHIP, src_tbl, ItemCode, process_name, tar_CPK);
            if (CPK_NG_NET_lst.Count > 0)
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
                for (int t = 0; t < src_tbl.Columns.Count; t++)
                {
                    result.Add(src_tbl.Columns[t].ColumnName);
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
        public DataTable Save_Submit_data(SqlConnection sqlcon_OK2SHIP, DataTable submit_tbl, string ItemCode, string LotNo, string process_name, string cycle_name, List<string> Net_name)
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
                Get_List_data_notNull(-1, temp_dt, new string[] { "ItemCode", "LotNo" }, ref cycle_data_lst, cycle_name, false);
                if (cycle_data_lst.Count > 0)
                {
                    MessageBox.Show("Data existed!\r\nCannot overwrite", "Warning");
                }
                else
                {
                    if (temp_dt.Rows.Count == submit_tbl.Rows.Count * submit_tbl.Columns.Count)
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
                }
            }
            else
            {
                bool save_en = false;
                if ((cycle_name == "Data") || (cycle_name == "Before"))
                {
                    save_en = true;
                }
                else
                {
                    List<string> before_data_lst = new List<string>();
                    Get_List_data_notNull(-1, temp_dt, new string[] { "ItemCode", "LotNo" }, ref before_data_lst, "Before", false);
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
                if (save_en)
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
                            string net = (r_inx + 1).ToString();
                            if (r_inx < Net_name.Count)
                            {
                                net = Net_name[r_inx];
                            }
                            TDMK_Code.insert_val_arr2(process_name, sqlcon_OK2SHIP, new string[] { "ID", "ItemCode", "LotNo", "Net_No", "Pcs_No", cycle_name }, new string[] { id.ToString(), ItemCode, LotNo, net, pcs_no, dr[dc].ToString() });
                            r_inx++;
                        }
                        col_inx++;
                    }
                }
            }
            return TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
        }
        public void Get_List_data_notNull(int col_inx, DataTable myDt, string[] src_arr, ref List<string> src_lst_data, string tar_item, bool distinct_en)
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
                                Get_List_data_notNull(col_inx + 1, curTbl, src_arr, ref src_lst_data, tar_item, distinct_en);
                            }
                        }
                    }
                    else
                    {
                        if (distinct_en)
                        {
                            src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Where(r => r != null).Distinct().ToArray());
                        }
                        else
                        {
                            src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Where(r => r != null).ToArray());
                        }
                    }
                }
            }
            else
            {
                if (distinct_en)
                {
                    src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Where(r => r != null).Distinct().ToArray());
                }
                else
                {
                    src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Where(r => r != null).ToArray());
                }
            }
        }
        public myExcel.Workbook create_export_wrk(string format_file, string process_name)
        {
            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            xlsApp.DisplayAlerts = false;
            List<char> split_char = new List<char> { '_', ' ', '-' };
            myExcel.Workbook wb = TDMK_Code.open_excel_file(format_file, "", "");
            string _process = remove_special_char(process_name, split_char).ToUpper();// process_name.Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
            string mySheet = "";
            foreach (myExcel.Worksheet tg_sht in wb.Worksheets)
            {
                string cur_sht_name = remove_special_char(tg_sht.Name, split_char).ToUpper();// tg_sht.Name.ToUpper().Replace("-", "").Replace(" ", "");
                if (cur_sht_name == _process)
                {
                    mySheet = tg_sht.Name;
                    break;
                }
            }
            if (mySheet != "")
            {
                myExcel.Workbook save_wb = TDMK_Code.Create_workbook();
                wb.Sheets[mySheet].Copy(After: save_wb.Sheets[1]);
                myExcel.Worksheet del_sht = save_wb.Sheets[1];
                del_sht.Delete();
                wb.Close();
                xlsApp.DisplayAlerts = true;
                return save_wb;
            }
            else
            {
                wb.Close();
                xlsApp.DisplayAlerts = true;
                return null;
            }
        }
        public string find_format(string in_data_loc, string ItemCode)
        {
            string result = "";
            string[] files = Directory.GetFiles(in_data_loc, "*" + ItemCode + "*.xlsm");

            if (files.Length > 0)
            {
                result = files[0];
            }
            return result;
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
        public SqlConnection initial_data(string DB_name, bool sa_en)
        {
            SqlConnection _sqlcon_OK2SHIP;
            string temp = find_config_path(Application.StartupPath, "TDMK Program");// Path.GetDirectoryName(Application.StartupPath);
            string config_file = Path.Combine(temp, "Config.ini");
            TDMK_init = new IniFile(config_file);
            if (!File.Exists(config_file))
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
                string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, DB_name, server_acc, server_pass).ConnectionString + ";Connection Timeout=60000";
                _sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);
            }
            else
            {
                string _strcon = TDMK_Code.data_connection2(server_name, DB_name).ConnectionString;
                _sqlcon_OK2SHIP = new SqlConnection(_strcon);
            }
            return _sqlcon_OK2SHIP;
        }
        public void Export_Bending(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name)
        {
            string f_loca = "";
            string format_file = find_format(format_folder, ItemCode);//@"D:\Customer Projects\SEEV\SMT Project\temp\Flex Bending_Export.xlsx";//
            if (format_file != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                myExcel.Workbook report_wrk = create_export_wrk(format_file, format_name);
                myExcel.Worksheet tar_wrksht = report_wrk.Sheets[1];
                List<string> data_addr_lst = Get_Bending_data_addr("Test item", "Sample", tar_wrksht, "A10", false);
                int col_inx = 0;
                foreach (var tbl in src_tbl_lst)
                {
                    Dictionary<string, List<string>> data_lst = Get_List_Pair_data(tbl);
                    List<string> before_data = data_lst["Before"];
                    string cur_addr = data_addr_lst[col_inx];
                    myExcel.Range dest_rgn = tar_wrksht.Range[cur_addr].Offset[2, 0];
                    foreach (var item in data_lst)
                    {
                        int col_off = 0;

                        if (item.Key != "Before")
                        {
                            col_off = 1;
                        }
                        int r_offset = 0;
                        foreach (var d in item.Value)
                        {
                            dest_rgn.Offset[r_offset, col_off].Value = d;
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
        public string[] Get_Net_Points(string src_Net, char[] split_char)
        {
            var arr = src_Net.Replace("S", "").Split(split_char).Distinct().ToArray();
            return arr.Where(x => myCode.IsNumeric(x)).ToArray();
        }
        public string Extract_Num_from_String(string str_in)
        {
            string result = "";
            result = new string(str_in.Where(x => char.IsDigit(x)).ToArray());
            return result;
        }
        public void Save_Log(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process, List<string> logfile_lst, string log_locate)
        {
            char[] split_char = new char[] { ' ', '(', '_' };
            foreach (string log in logfile_lst)
            {
                string f_name = Path.Combine(log_locate, log);
                string cycle = "After";
                string bl = Path.GetFileNameWithoutExtension(log);
                string block_name = bl.Split('-').Last().Split(split_char).First();
                string num = Extract_Num_from_String(block_name);
                string bl1 = block_name.Substring(0, block_name.Length - num.Length);
                if (bl1.ToUpper().Contains("BF"))
                {
                    cycle = "Before";
                }
                else
                {
                    cycle = cycle + "_" + num;
                }
                DataTable netSpec_tbl = new DataTable();
                DataTable result_data = new DataTable();
                List<string> Net_lst = new List<string>();
                Dictionary<string, string> Log_info = new Dictionary<string, string>();
                result_data = DAT_To_DataTable_details_time(f_name, ref Net_lst, ref netSpec_tbl, ref Log_info);
                DataTable dt = Save_LogFile_detail_time(sqlcon_OK2SHIP, result_data, ItemCode, LotNo, process, cycle, log, Net_lst, Log_info);
                DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process }));
                if (spec_dt.Rows.Count == 0)
                {
                    List<string> items = new List<string>() { "ID", "ItemCode" };
                    DataTable temp_spec = netSpec_tbl;
                    foreach (DataColumn dc in temp_spec.Columns)
                    {
                        items.Add("[" + dc.ColumnName + "]");
                    }
                    items.Add("Remark");
                    int id = TDMK_Code.SQL_MAX("NET_SPEC", "ID", sqlcon_OK2SHIP);
                    foreach (DataRow dr in temp_spec.Rows)
                    {
                        id++;
                        List<string> items_val = new List<string>() { id.ToString(), ItemCode };
                        foreach (DataColumn dc in temp_spec.Columns)
                        {
                            items_val.Add(dr[dc].ToString());
                        }
                        items_val.Add(process);
                        TDMK_Code.insert_val_arr2("NET_SPEC", sqlcon_OK2SHIP, items.ToArray(), items_val.ToArray());
                    }
                }
            }
            MessageBox.Show(new Form { TopMost = true }, "Lưu dữ liệu hoàn thành", "Thông báo");
        }
        public void Save_Log(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process, string shift, List<string> logfile_lst, string log_locate)
        {
            char[] split_char = new char[] { ' ', '(', '_' };
            bool save_spec_en = true;
            foreach (string log in logfile_lst)
            {
                string f_name = Path.Combine(log_locate, log);
                string cycle = "After";
                string bl = Path.GetFileNameWithoutExtension(log);
                string block_name = bl.Split('-').Last().Split(split_char).First();
                string num = Extract_Num_from_String(block_name);
                string bl1 = block_name.Substring(0, block_name.Length - num.Length);
                if (bl1.ToUpper().Contains("BF"))
                {
                    cycle = "Before";
                }
                else
                {
                    cycle = cycle + "_" + num;
                }
                DataTable netSpec_tbl = new DataTable();
                DataTable result_data = new DataTable();
                List<string> Net_lst = new List<string>();
                Dictionary<string, string> Log_info = new Dictionary<string, string>();
                string userid = "";
                string itemname = "";
                result_data = DAT_To_DataTable_details_time(f_name, ref Net_lst, ref netSpec_tbl, ref Log_info, ref itemname, ref userid);
                bool save_result = Save_LogFile_detail_time_result(sqlcon_OK2SHIP, result_data, ItemCode, LotNo, process, cycle, shift, userid, itemname, log, Net_lst, Log_info);
            start_label: DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process }));
                if (spec_dt.Rows.Count == 0)
                {
                    List<string> items = new List<string>() { "ID", "ItemCode" };
                    DataTable temp_spec = netSpec_tbl;
                    foreach (DataColumn dc in temp_spec.Columns)
                    {
                        items.Add("[" + dc.ColumnName + "]");
                    }
                    items.Add("Remark");
                    int id = TDMK_Code.SQL_MAX("NET_SPEC", "ID", sqlcon_OK2SHIP);
                    foreach (DataRow dr in temp_spec.Rows)
                    {
                        id++;
                        List<string> items_val = new List<string>() { id.ToString(), ItemCode };
                        foreach (DataColumn dc in temp_spec.Columns)
                        {
                            items_val.Add(dr[dc].ToString());
                        }
                        items_val.Add(process);
                        TDMK_Code.insert_val_arr2("NET_SPEC", sqlcon_OK2SHIP, items.ToArray(), items_val.ToArray());
                    }
                    save_spec_en = false;
                }
                else
                {
                    if (save_spec_en)
                    {
                        if (netSpec_tbl.Rows.Count != spec_dt.Rows.Count)
                        {
                            if (MessageBox.Show("Dữ liệu NET Spec của logfile khác với dữ liệu đã lưu. Bạn muốn cập nhật lại ?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                            {
                                save_spec_en = false;
                                TDMK_Code.Delelte_FilteredItem_arr("NET_SPEC", sqlcon_OK2SHIP, TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process }));
                                goto start_label;
                            }
                        }
                    }
                }
            }
            MessageBox.Show(new Form { TopMost = true }, "Lưu dữ liệu logfile hoàn thành", "Thông báo");
        }
        public void Save_Log_submit(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process, List<string> logfile_lst, string log_locate)
        {
            char[] split_char = new char[] { ' ', '(', '_' };
            foreach (string log in logfile_lst)
            {
                string f_name = Path.Combine(log_locate, log);
                string cycle = "After";
                string bl = Path.GetFileNameWithoutExtension(log);
                string block_name = bl.Split('-').Last().Split(split_char).First();
                string num = Extract_Num_from_String(block_name);
                string bl1 = block_name.Substring(0, block_name.Length - num.Length);
                if (bl1.ToUpper().Contains("BF"))
                {
                    cycle = "Before";
                }
                else
                {
                    cycle = cycle + "_" + num;
                }
                DataTable netSpec_tbl = new DataTable();
                DataTable result_data = new DataTable();
                List<string> Net_lst = new List<string>();
                Dictionary<string, string> Log_info = new Dictionary<string, string>();
                result_data = DAT_To_DataTable_details_time(f_name, ref Net_lst, ref netSpec_tbl, ref Log_info);
                DataTable dt = Save_LogFile_detail_time(sqlcon_OK2SHIP, result_data, ItemCode, LotNo, process, cycle, log, Net_lst, Log_info);
                DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process }));
                if (spec_dt.Rows.Count == 0)
                {
                    List<string> items = new List<string>() { "ID", "ItemCode" };
                    DataTable temp_spec = netSpec_tbl;
                    foreach (DataColumn dc in temp_spec.Columns)
                    {
                        items.Add("[" + dc.ColumnName + "]");
                    }
                    items.Add("Remark");
                    int id = TDMK_Code.SQL_MAX("NET_SPEC", "ID", sqlcon_OK2SHIP);
                    foreach (DataRow dr in temp_spec.Rows)
                    {
                        id++;
                        List<string> items_val = new List<string>() { id.ToString(), ItemCode };
                        foreach (DataColumn dc in temp_spec.Columns)
                        {
                            items_val.Add(dr[dc].ToString());
                        }
                        items_val.Add(process);
                        TDMK_Code.insert_val_arr2("NET_SPEC", sqlcon_OK2SHIP, items.ToArray(), items_val.ToArray());
                    }
                }
            }
            MessageBox.Show(new Form { TopMost = true }, "Lưu dữ liệu hoàn thành", "Thông báo");
        }
        public void Export_Thermal_HeatSoak_Bend_All(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name, int qty, string format_type)
        {
            //string format_loc = Path.Combine(format_folder, format_type);
            string format_loc = Path.Combine(find_config_path(Application.StartupPath, "SEEV Data"), "Format", format_type);
            Report_location = Path.Combine(find_config_path(Application.StartupPath, "SEEV Data"), "Report");
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
                DataTable log_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name + "_LOGFILE", filter_str);
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
                        report_name = Path.GetFileNameWithoutExtension(report_wrk.Name) + "-" + LotNo + " " + DateTime.Now.ToString("yyyyMMdd") + Path.GetExtension(report_wrk.Name);
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
                            string cycle_filter = "BF";
                            if (item.Key.Contains("After"))
                            {
                                cycle_filter = "L" + Extract_Num_from_String(item.Key);
                            }
                            List<string> log_time_lst = log_tbl.AsEnumerable().Where(x => x.Field<string>("Cycles_name") == item.Key).Select(x => x.Field<string>("Remark")).Distinct().ToList();
                            string log_time = "";
                            if (log_time_lst.Count > 0)
                            {
                                string _log_time = log_tbl.AsEnumerable().Where(x => x.Field<string>("Cycles_name") == item.Key).Select(x => x.Field<string>("Logfile")).FirstOrDefault();
                                log_time = _log_time.Split('-').First() + " " + log_time_lst.Last();
                            }
                            if (col_inx < 1)
                            {
                                myExcel.Range time_rgn_offset = tar_wrksht.Range[Echeck_Cycle_addr[item.Key]].Offset[0, 1];
                                time_rgn_offset.Value = log_time.Split(' ');
                                time_rgn_offset.Offset[1, 0].Value = log_time.Split(' ');
                            }
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

            //ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

        }
        public void Export_Thermal_HeatSoak_Bend(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name, int tar_pcs)
        {
            string format_file = ""; //find_format(format_folder, ItemCode);
            List<string> format_lst = get_multiple_files(format_folder, new List<string> { "*.xlsx", "*.xlsm" }, ItemCode, process_name);
            if (format_lst.Count > 0)
            {
                format_file = format_lst[0];
            }
            if (format_file != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable _src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                DataTable _spec_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process_name }));
                DataTable log_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name + "_LOGFILE", filter_str);
                SortedDictionary<int, string> sel_index_lst = Get_bending_Net(sqlcon_OK2SHIP, ItemCode, process_name);
                List<int> index_lst = sel_index_lst.Keys.ToList();
                DataTable src_tbl = _src_tbl.AsEnumerable().Where(x => index_lst.Contains(Convert.ToInt32(x.Field<string>("Net_no")) - 1)).CopyToDataTable();
                DataTable spec_tbl = _spec_tbl.AsEnumerable().Where(x => x.Field<string>("Net_Name") != "No").CopyToDataTable();
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                myExcel.Workbook report_wrk = create_export_wrk(format_file, "Bending < 10%");
                myExcel.Worksheet tar_wrksht = report_wrk.Sheets[1];
                string Echeck_start_rgn = Get_start_range("Condition", "A1", tar_wrksht);
                myExcel.Range Echeck_cycle_rgn = tar_wrksht.Range[Echeck_start_rgn];
                Dictionary<string, string> Echeck_Cycle_addr = Get_Echeck_address(Echeck_cycle_rgn);
                Dictionary<int, string> cycle_data_addr_lst = Get_Bending_Cycle_addr("Test item", "Bending", tar_wrksht, "A10", false);
                if (tar_pcs <= src_tbl_lst.Count)
                {
                    int col_inx = tar_pcs - 1;
                    DataTable tbl = src_tbl_lst[col_inx];
                    Dictionary<string, List<string>> data_lst = Get_List_Pair_data(tbl);
                    List<string> before_data = data_lst["Before"];
                    foreach (var item in data_lst)
                    {
                        //List<string> after_data = item.Value;
                        List<string> sel_data = item.Value;
                        string cycle_filter = "BF";
                        if (item.Key.Contains("After"))
                        {
                            cycle_filter = "L" + Extract_Num_from_String(item.Key);
                        }
                        List<string> log_time_lst = log_tbl.AsEnumerable().Where(x => x.Field<string>("Logfile").Contains(cycle_filter)).Select(x => x.Field<string>("Remark")).Distinct().ToList();
                        string log_time = "";
                        if (log_time_lst.Count > 0)
                        {
                            log_time = log_tbl.Rows[0]["Logfile"].ToString().Split('-').First() + " " + log_time_lst.Last();
                        }
                        if (Echeck_Cycle_addr.Keys.ToList().IndexOf(item.Key) != -1)
                        {
                            myExcel.Range sum_rgn_offset = tar_wrksht.Range[Echeck_Cycle_addr[item.Key]].Offset[0, 2 + col_inx];
                            myExcel.Range time_rgn_offset = tar_wrksht.Range[Echeck_Cycle_addr[item.Key]].Offset[0, 1 + col_inx];
                            time_rgn_offset.Value = log_time;
                            bool data_OK = false;
                            for (int r_inx = 0; r_inx < sel_data.Count; r_inx++)
                            {
                                string USL = myCode.checkDBNull(spec_tbl.Rows[r_inx]["USL"]);
                                string LSL = myCode.checkDBNull(spec_tbl.Rows[r_inx]["LSL"]);
                                string cur_val = sel_data[r_inx];
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
                                sum_rgn_offset.Value = "OK";
                            }
                            else
                            {
                                sum_rgn_offset.Value = "FAIL";
                            }
                        }
                        if (item.Key.Contains("After"))
                        {
                            int cycle_no = Convert.ToInt32(Extract_Num_from_String(item.Key));
                            if (cycle_data_addr_lst.Keys.ToList().IndexOf(cycle_no) != -1)
                            {
                                myExcel.Range cycle_rgn = tar_wrksht.Range[cycle_data_addr_lst[cycle_no]].Offset[2, 0];
                                for (int r_inx = 0; r_inx < sel_data.Count; r_inx++)
                                {
                                    cycle_rgn.Offset[r_inx, 0].Value = before_data[r_inx];
                                    cycle_rgn.Offset[r_inx, 1].Value = sel_data[r_inx];
                                }
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
        public void Export_Thermal_HeatSoak_Bend_MultiType(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name, int tar_pcs, string format_type, string shift)
        {
            //string format_loc = Path.Combine(format_folder, format_type);
            string format_loc = Path.Combine(find_config_path(Application.StartupPath, "SEEV Data"), "Format", format_type);
            Report_location = Path.Combine(find_config_path(Application.StartupPath, "SEEV Data"), "Report");
            //string format_file = find_format(format_folder, ItemCode, new List<string> { "*.xlsx", "*.xlsm" });
            List<string> format_lst = get_multiple_files(format_loc, new List<string> { "*.xlsx", "*.xlsm" }, ItemCode, process_name);
            string format_file = "";
            if (format_lst.Count > 0)
            {
                format_file = format_lst[0];
            }
            if (format_file != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Shift" }, new string[] { ItemCode, LotNo, shift });
                DataTable _src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                DataTable _spec_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process_name }));
                DataTable log_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name + "_LOGFILE", filter_str);
                SortedDictionary<int, string> sel_index_lst = Get_bending_Net(sqlcon_OK2SHIP, ItemCode, process_name);
                List<int> index_lst = sel_index_lst.Keys.ToList();
                DataTable src_tbl = _src_tbl.AsEnumerable().Where(x => index_lst.Contains(Convert.ToInt32(x.Field<string>("Net_no")) - 1)).CopyToDataTable();
                DataTable spec_tbl = _spec_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_Report") != "No").CopyToDataTable();

                List<string> NET_name = spec_tbl.AsEnumerable().Select(x => x.Field<string>("Net_Name")).ToList();
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                List<char> reject_char_lst = new List<char> { ' ', '_', '-', '\r', '\n' };
                myExcel.Workbook report_wrk = null;// 
                myExcel.Worksheet tar_wrksht = null;
                myExcel.Worksheet result_sht = null;
                string find_cycle_key = "Bending";
                if (format_type == "NPI")
                {
                    find_cycle_key = "Sample";
                    report_wrk = TDMK_Code.open_excel_file(format_file, "", "");
                    foreach (myExcel.Worksheet sht in report_wrk.Worksheets)
                    {
                        if (remove_special_char(sht.Name, reject_char_lst) == remove_special_char(process_name, reject_char_lst))
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
                    foreach (myExcel.Worksheet sht in report_wrk.Worksheets)
                    {
                        if (remove_special_char(sht.Name, reject_char_lst) == remove_special_char("Result", reject_char_lst))
                        {
                            result_sht = sht;
                            break;
                        }
                    }
                }
                if (tar_wrksht != null)
                {
                    Dictionary<int, string> cycle_data_addr_lst = Get_Bending_Cycle_addr("Test item", find_cycle_key, tar_wrksht, "A10", false);
                    string NET_start_rgn = Get_start_range("Test item", "A10", tar_wrksht);
                    myExcel.Range NET_rgn = tar_wrksht.Range[NET_start_rgn];
                    string Echeck_start_rgn = Get_start_range("Condition", "A1", tar_wrksht);
                    myExcel.Range Echeck_cycle_rgn = tar_wrksht.Range[Echeck_start_rgn];
                    Dictionary<string, string> Echeck_Cycle_addr = Get_Echeck_address(Echeck_cycle_rgn);
                    string proId_addr = "";
                    string itemname_addr = "";
                    string itemcode_addr = "";
                    string lotno_addr = "";
                    string process_addr = "";
                    string cycle_addr = "";
                    string time_addr = "";
                    string result_addr = "";
                    string operator_addr = "";
                    if (result_sht != null)
                    {
                        proId_addr = Get_start_range("ProductID", "A1", result_sht, true);
                        itemname_addr = Get_start_range("Item name", "A1", result_sht, true);
                        itemcode_addr = Get_start_range("Item code", "A1", result_sht, true);
                        lotno_addr = Get_start_range("Item lot", "A1", result_sht, true);
                        process_addr = Get_start_range("Test Condition", "A1", result_sht, true);
                        cycle_addr = Get_start_range("Bending cycle", "A1", result_sht, true);
                        time_addr = Get_start_range(@"ICT Datetime Finish", "A1", result_sht, true);
                        result_addr = Get_start_range("Result", "A1", result_sht, true);
                        operator_addr = Get_start_range("Output OperatorID", "A1", result_sht, true);
                    }
                    if (tar_pcs <= src_tbl_lst.Count)
                    {
                        //string[] temp = Path.GetFileNameWithoutExtension(report_wrk.Name).Split('-');
                        string itemname = "";
                        string UserID = "";
                        //if (temp.Length > 1)
                        //{
                        //    itemname = temp[1];
                        //}
                        itemname = myCode.checkDBNull(log_tbl.Rows[0]["ItemName"]);
                        UserID = myCode.checkDBNull(log_tbl.Rows[0]["UserID"]);
                        string report_name = process_name + " " + itemname + "-" + ItemCode + "-" + LotNo + "-C" + shift + " " + DateTime.Now.ToString("yyyyMMdd") + Path.GetExtension(report_wrk.Name);
                        if (format_type == "NPI")
                        {
                            report_name = Path.GetFileNameWithoutExtension(report_wrk.Name) + "-" + LotNo + " " + DateTime.Now.ToString("yyyyMMdd") + Path.GetExtension(report_wrk.Name);
                        }
                        string daily_folder = Path.Combine(Report_location, "Daily_report");
                        if (!Directory.Exists(daily_folder))
                        {
                            Directory.CreateDirectory(daily_folder);
                        }
                        report_wrk.SaveAs(Path.Combine(daily_folder, report_name));
                        int col_inx = tar_pcs - 1;
                        DataTable tbl = src_tbl_lst[col_inx];
                        string prod_id = tbl.Rows[0]["Pcs_No"].ToString();
                        Dictionary<string, List<string>> data_lst = Get_List_Pair_data(tbl);
                        List<string> before_data = data_lst["Before"];
                        int r_offset = 0;
                        foreach (var item in data_lst)
                        {
                            List<string> sel_data = item.Value;
                            string cycle_filter = "BF";
                            if (item.Key.Contains("After"))
                            {
                                cycle_filter = "L" + Extract_Num_from_String(item.Key);
                            }
                            List<string> log_time_lst = log_tbl.AsEnumerable().Where(x => x.Field<string>("Cycles_name") == item.Key).Select(x => x.Field<string>("Remark")).Distinct().ToList();
                            string log_time = "";
                            if (log_time_lst.Count > 0)
                            {
                                string _log_time = log_tbl.AsEnumerable().Where(x => x.Field<string>("Cycles_name") == item.Key).Select(x => x.Field<string>("Logfile")).FirstOrDefault();
                                log_time = _log_time.Split('-').First() + " " + log_time_lst.Last();
                            }
                            bool data_OK = false;
                            if (Echeck_Cycle_addr.Keys.ToList().IndexOf(item.Key) != -1)
                            {
                                myExcel.Range sum_rgn_offset = tar_wrksht.Range[Echeck_Cycle_addr[item.Key]].Offset[0, 2 + col_inx];
                                myExcel.Range time_rgn_offset = tar_wrksht.Range[Echeck_Cycle_addr[item.Key]].Offset[0, 1 + col_inx];
                                time_rgn_offset.Value = log_time;

                                for (int r_inx = 0; r_inx < sel_data.Count; r_inx++)
                                {
                                    string USL = myCode.checkDBNull(spec_tbl.Rows[r_inx]["USL"]);
                                    string LSL = myCode.checkDBNull(spec_tbl.Rows[r_inx]["LSL"]);
                                    string cur_val = sel_data[r_inx];
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
                                    sum_rgn_offset.Value = "OK";
                                }
                                else
                                {
                                    sum_rgn_offset.Value = "FAIL";
                                }
                            }
                            if (item.Key.Contains("After"))
                            {
                                int cycle_no = Convert.ToInt32(Extract_Num_from_String(item.Key));
                                if (cycle_data_addr_lst.Keys.ToList().IndexOf(cycle_no) != -1)
                                {
                                    myExcel.Range cycle_rgn = tar_wrksht.Range[cycle_data_addr_lst[cycle_no]].Offset[2, 0];
                                    for (int r_inx = 0; r_inx < sel_data.Count; r_inx++)
                                    {
                                        cycle_rgn.Offset[r_inx, 0].Value = before_data[r_inx];
                                        cycle_rgn.Offset[r_inx, 1].Value = sel_data[r_inx];
                                        cycle_rgn.Offset[r_inx, 2].FormulaR1C1 = "=RC[-1]/RC[-2]-1";
                                        if (Math.Abs(Convert.ToDouble(cycle_rgn.Offset[r_inx, 2].Value)) > 0.1)
                                        {
                                            cycle_rgn.Offset[r_inx, 2].Interior.Color = 255;
                                        }
                                        //cycle_rgn.Offset[r_inx, 2].Style = "Percent";
                                        cycle_rgn.Offset[r_inx, 2].NumberFormat = "0.00 %";
                                    }
                                }
                            }
                            if (result_sht != null)
                            {
                                if (item.Key.Contains("After"))
                                {
                                    //result_sht.Range["E2"].Offset[r_offset, 0].Value = prod_id;
                                    //result_sht.Range["F2"].Offset[r_offset, 0].Value = itemname;
                                    //result_sht.Range["G2"].Offset[r_offset, 0].Value = ItemCode;
                                    //result_sht.Range["H2"].Offset[r_offset, 0].Value = LotNo;
                                    //result_sht.Range["Q2"].Offset[r_offset, 0].Value = process_name;
                                    //result_sht.Range["R2"].Offset[r_offset, 0].Value = cycle_filter.Replace("L", "");
                                    //result_sht.Range["T2"].Offset[r_offset, 0].Value = log_time;

                                    result_sht.Range[proId_addr].Offset[r_offset, 0].Value = prod_id;
                                    result_sht.Range[itemname_addr].Offset[r_offset, 0].Value = itemname;
                                    result_sht.Range[itemcode_addr].Offset[r_offset, 0].Value = ItemCode;
                                    result_sht.Range[lotno_addr].Offset[r_offset, 0].Value = LotNo;
                                    result_sht.Range[process_addr].Offset[r_offset, 0].Value = process_name;
                                    result_sht.Range[cycle_addr].Offset[r_offset, 0].Value = cycle_filter.Replace("L", "");
                                    result_sht.Range[time_addr].Offset[r_offset, 0].Value = log_time;
                                    result_sht.Range[operator_addr].Offset[r_offset, 0].Value = UserID;
                                    if (data_OK)
                                    {
                                        result_sht.Range[result_addr].Offset[r_offset, 0].Value = "Pass";
                                    }
                                    else
                                    {
                                        result_sht.Range[result_addr].Offset[r_offset, 0].Value = "NG";
                                    }
                                    r_offset++;
                                }
                            }

                        }
                        for (int i = 0; i < NET_name.Count; i++)
                        {
                            NET_rgn.Offset[i, 0].Value = NET_name[i];
                        }
                        string last_cycle_addr = cycle_data_addr_lst.LastOrDefault().Value;
                        int col_num = tar_wrksht.Range[last_cycle_addr].MergeArea.Columns.Count;
                        int total_col_num = tar_wrksht.Range[last_cycle_addr].Column + col_num - 1;
                        int r_off = NET_name.Count - 1;
                        myExcel.Range final_rgn = tar_wrksht.Range[tar_wrksht.Range[NET_start_rgn], tar_wrksht.Range[NET_start_rgn].Offset[r_off, total_col_num - 1]];
                        List<myExcel.XlBordersIndex> st_lst = new List<myExcel.XlBordersIndex>() { myExcel.XlBordersIndex.xlEdgeLeft, myExcel.XlBordersIndex.xlEdgeTop, myExcel.XlBordersIndex.xlEdgeBottom, myExcel.XlBordersIndex.xlEdgeRight, myExcel.XlBordersIndex.xlInsideHorizontal, myExcel.XlBordersIndex.xlInsideVertical };//myExcel.XlBordersIndex.xlDiagonalDown, myExcel.XlBordersIndex.xlDiagonalUp,
                        foreach (var st in st_lst)
                        {
                            final_rgn.Borders[st].LineStyle = myExcel.XlLineStyle.xlContinuous;
                        }
                        report_wrk.Save();
                    }
                }
                else
                {
                    MessageBox.Show("Format của ItemCode " + ItemCode + " không đúng", "Cảnh báo");
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
        public List<string> get_multiple_files(string src_path, List<string> extensions, string tar_ItemCode, string tar_process)
        {
            List<string> result = new List<string>();
            List<char> remove_char = new List<char> { ' ', '-', '_' };
            string process = new string(tar_process.Where(x => remove_char.IndexOf(x) == -1).ToArray());
            DirectoryInfo directory = new DirectoryInfo(src_path);
            var files = extensions.SelectMany(e => directory.EnumerateFiles(e, SearchOption.AllDirectories)).Where(x => x.FullName.Contains(tar_ItemCode) && remove_special_char(x.FullName, remove_char).ToUpper().Contains(process.ToUpper()));
            foreach (var item in files)
            {
                result.Add(item.FullName);
            }
            return result;
        }
        public string remove_special_char(string src_str, List<char> remove_char)
        {
            return new string(src_str.Where(x => remove_char.IndexOf(x) == -1).ToArray());
        }
        public Dictionary<int, string> Get_Bending_Cycle_addr(string search_item, string search_key, myExcel.Worksheet tar_wrksht, string start_range_addr, bool row_direction)
        {
            Dictionary<int, string> result = new Dictionary<int, string>();
            string tar_addr = Find_Cell_Addr(search_item, start_range_addr, tar_wrksht);
            myExcel.Range sample_rgn = tar_wrksht.Range[tar_addr];
            myExcel.Range sel_rgn = sample_rgn;
            int count_null = 0;
            while (true)
            {
                string rgn_addr = sel_rgn.AddressLocal;
                myExcel.Range temp_rgn = sel_rgn.Offset[0, 1];
                if (row_direction)
                {
                    temp_rgn = sel_rgn.Offset[1, 0];
                }
                string temp_addr = temp_rgn.AddressLocal;
                string temp_rgn_val = myCode.checkDBNull(temp_rgn.Value);
                if (temp_rgn_val.ToUpper().Contains(search_key.ToUpper()))
                {
                    int key = Convert.ToInt32(Extract_Num_from_String(temp_rgn_val));
                    if (result.Keys.ToList().IndexOf(key) == -1)
                    {
                        result.Add(key, temp_rgn.AddressLocal);
                    }
                    else
                    {
                        result[key] = temp_rgn.AddressLocal;
                    }
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
        public Dictionary<string, DataTable> Summary_Cycles_Logfile(SqlConnection sqlcon, string itemcode, string lotno, string process, int sel_num = 1)
        {
            Dictionary<string, DataTable> dic_sum = new Dictionary<string, DataTable>();
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcode, lotno });
            string spec_filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { itemcode, process });
            DataTable src_log_dt = TDMK_Code.Datatable_Filter(sqlcon, process + "_LOGFILE", filter_str);
            List<string> cycle_lst = src_log_dt.AsEnumerable().Select(x => x.Field<string>("Cycles_name")).Distinct().ToList();
            foreach (string cycle in cycle_lst)
            {
                dic_sum.Add(cycle, new DataTable());
                List<string> logfile_lst = src_log_dt.AsEnumerable().Where(x => x.Field<string>("Cycles_name") == cycle).Select(x => x.Field<string>("Logfile")).Distinct().ToList();
                Dictionary<string, List<string>> dic_logfile = Get_BlockofLogFile(logfile_lst);
                Dictionary<string, DataTable> result_table_lst = new Dictionary<string, DataTable>();
                Dictionary<string, List<DataTable>> dic_tbl_data_lst = new Dictionary<string, List<DataTable>>();
                DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "NET_SPEC", spec_filter_str);
                foreach (var block in dic_logfile)
                {
                    foreach (var logfile in block.Value)
                    {
                        DataTable dt = new DataTable();
                        List<string> temp_net_lst = new List<string>();
                        dt = Load_Log_Data(sqlcon, itemcode, lotno, process, cycle, logfile, temp_net_lst);

                        if (dt.Rows.Count > 0)
                        {
                            if (TDMK_Code.check_exist_list_index2(cycle, dic_tbl_data_lst.Keys.ToList()) == -1)
                            {
                                dic_tbl_data_lst.Add(cycle, new List<DataTable>() { dt });
                            }
                            else
                            {
                                dic_tbl_data_lst[cycle].Add(dt);
                            }
                        }
                    }
                }
                foreach (var tbl in dic_tbl_data_lst)
                {
                    DataTable block_tbl = Summary_Data_Table_from_list(tbl.Value, spec_dt);
                    List<string> col_lst = Get_column_name(block_tbl);
                    if (col_lst.Count > 1)
                    {
                        List<string> main_col_lst = col_lst.Where(x => !x.Contains("_")).ToList();
                        DataTable final_dt = new DataTable();
                        foreach (string sel_col in main_col_lst)
                        {
                            string[] cur_col = col_lst.Where(x => x.Contains(sel_col)).ToArray();
                            DataTable cur_dt = block_tbl.AsDataView().ToTable(false, cur_col);
                            final_dt.Columns.Add(sel_col);
                            int r_inx = 0;
                            foreach (DataRow dr in cur_dt.Rows)
                            {
                                string USL = spec_dt.Rows[r_inx]["USL"].ToString();
                                string LSL = spec_dt.Rows[r_inx]["LSL"].ToString();
                                List<double> col_data_lst = new List<double>();
                                foreach (DataColumn dc in cur_dt.Columns)
                                {
                                    string cur_val = myCode.checkDBNull(dr[dc]);
                                    if (check_in_limit(USL, LSL, cur_val))
                                    {
                                        col_data_lst.Add(Convert.ToDouble(cur_val));
                                    }
                                }
                                if (final_dt.Rows.Count <= r_inx)
                                {
                                    final_dt.Rows.Add();
                                }
                                if (col_data_lst.Count > 0)
                                {
                                    final_dt.Rows[r_inx][sel_col] = col_data_lst.Min();
                                }
                                r_inx++;
                            }
                        }
                        result_table_lst.Add(tbl.Key, final_dt);
                    }
                    else
                    {
                        result_table_lst.Add(tbl.Key, block_tbl);
                    }
                }
                List<string> pcs_lst = new List<string>();
                pcs_lst = Get_column_name(dic_sum["Before"]);
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
                                string col_name = tbl.Value.Columns[c_inx].ColumnName;// + "_" + "BL" + bl_name;
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
                        List<ECheck_Process.NG_list2> cur_NGList = exp_process.Get_NG_point_tbl(spec_dt, tbl.Value);
                        DataTable sel_dt = exp_process.Summary_Selected_FromExisted(tbl.Value, cur_NGList, item_qty[inx]);
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

                            string col_name = summ_tbl_lst[i].Columns[j].ColumnName;// + "_" + "BL" + bl_name;
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
                                string col_name = summ_tbl_lst[i].Columns[c_inx].ColumnName;// + "_" + "BL" + bl_name;
                                disp_result.Rows[r_inx][col_name] = summ_tbl_lst[i].Rows[r_inx][c_inx];
                            }
                        }
                    }
                    dic_sum[cycle] = disp_result;
                }
            }
            return dic_sum;
        }
        public Dictionary<string, DataTable> Summary_Cycles_Logfile(SqlConnection sqlcon, string itemcode, string lotno, string process, string shift, int sel_num = 1, string season = "")
        {
            Dictionary<string, DataTable> dic_sum = new Dictionary<string, DataTable>();
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Shift" }, new string[] { itemcode, lotno, shift });
            string spec_filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { itemcode, process });
            DataTable src_log_dt = TDMK_Code.Datatable_Filter(sqlcon, process + "_LOGFILE" + season, filter_str);
            List<string> _after_pcs_lst = src_log_dt.AsEnumerable().Where(x => x.Field<string>("Cycles_name").Contains("After")).Select(x => x.Field<string>("Pcs_No")).Distinct().ToList();
            List<string> after_pcs_lst = _after_pcs_lst.Where(x => !x.Contains("_")).Select(x => x.ToUpper()).Distinct().ToList();
            List<string> _cycle_lst = src_log_dt.AsEnumerable().Select(x => x.Field<string>("Cycles_name").Trim()).Distinct().ToList();
            List<string> cycle_lst = new List<string>();
            if (_cycle_lst.IndexOf("Before") != -1)
            {
                cycle_lst.Add("Before");
                _cycle_lst.Remove("Before");
                cycle_lst.AddRange(_cycle_lst);
            }
            else
            {
                return null;
            }
            foreach (string cycle in cycle_lst)
            {
                dic_sum.Add(cycle, new DataTable());
                List<string> logfile_lst = src_log_dt.AsEnumerable().Where(x => x.Field<string>("Cycles_name").Trim() == cycle).Select(x => x.Field<string>("Logfile")).Distinct().ToList();
                Dictionary<string, List<string>> dic_logfile = Get_BlockofLogFile(logfile_lst);
                Dictionary<string, DataTable> result_table_lst = new Dictionary<string, DataTable>();
                Dictionary<string, List<DataTable>> dic_tbl_data_lst = new Dictionary<string, List<DataTable>>();
                DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "NET_SPEC", spec_filter_str);
                foreach (var block in dic_logfile)
                {
                    foreach (var logfile in block.Value)
                    {
                        DataTable dt = new DataTable();
                        List<string> temp_net_lst = new List<string>();
                        string userid = "";
                        string itemname = "";
                        dt = Load_Log_Data(sqlcon, itemcode, lotno, process, cycle, shift, ref userid, ref itemname, logfile, temp_net_lst);

                        if (dt.Rows.Count > 0)
                        {
                            if (TDMK_Code.check_exist_list_index2(cycle, dic_tbl_data_lst.Keys.ToList()) == -1)
                            {
                                dic_tbl_data_lst.Add(cycle, new List<DataTable>() { dt });
                            }
                            else
                            {
                                dic_tbl_data_lst[cycle].Add(dt);
                            }
                        }
                    }
                }
                foreach (var tbl in dic_tbl_data_lst)
                {
                    DataTable block_tbl = Summary_Data_Table_from_list(tbl.Value, spec_dt);
                    List<string> col_lst = Get_column_name(block_tbl);
                    if (col_lst.Count > 1)
                    {
                        List<string> main_col_lst = col_lst.Where(x => !x.Contains("_")).ToList();
                        DataTable final_dt = new DataTable();
                        foreach (string sel_col in main_col_lst)
                        {
                            string[] cur_col = col_lst.Where(x => x.Contains(sel_col)).ToArray();
                            DataTable cur_dt = block_tbl.AsDataView().ToTable(false, cur_col);
                            final_dt.Columns.Add(sel_col);
                            int r_inx = 0;
                            foreach (DataRow dr in cur_dt.Rows)
                            {
                                string USL = spec_dt.Rows[r_inx]["USL"].ToString();
                                string LSL = spec_dt.Rows[r_inx]["LSL"].ToString();
                                List<double> col_data_lst = new List<double>();
                                foreach (DataColumn dc in cur_dt.Columns)
                                {
                                    string cur_val = myCode.checkDBNull(dr[dc]);
                                    if (check_in_limit(USL, LSL, cur_val))
                                    {
                                        col_data_lst.Add(Convert.ToDouble(cur_val));
                                    }
                                }
                                if (final_dt.Rows.Count <= r_inx)
                                {
                                    final_dt.Rows.Add();
                                }
                                if (col_data_lst.Count > 0)
                                {
                                    final_dt.Rows[r_inx][sel_col] = col_data_lst.Min();
                                }
                                r_inx++;
                            }
                        }
                        result_table_lst.Add(tbl.Key, final_dt);
                    }
                    else
                    {
                        result_table_lst.Add(tbl.Key, block_tbl);
                    }
                }
                List<string> pcs_lst = new List<string>();
                pcs_lst = Get_column_name(dic_sum["Before"]);
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
                                string col_name = tbl.Value.Columns[c_inx].ColumnName;// + "_" + "BL" + bl_name;
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
                        List<ECheck_Process.NG_list2> cur_NGList = exp_process.Get_NG_point_tbl(spec_dt, dest_dt);
                        DataTable sel_dt = exp_process.Summary_Selected_FromExisted(dest_dt, cur_NGList, item_qty[inx]);
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

                            string col_name = summ_tbl_lst[i].Columns[j].ColumnName;// + "_" + "BL" + bl_name;
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
                                string col_name = summ_tbl_lst[i].Columns[c_inx].ColumnName;// + "_" + "BL" + bl_name;
                                disp_result.Rows[r_inx][col_name] = summ_tbl_lst[i].Rows[r_inx][c_inx];
                            }
                        }
                    }
                    dic_sum[cycle] = disp_result;
                }
            }
            return dic_sum;
        }
        public DataTable Summary_Logfile_All(SqlConnection sqlcon, string itemcode, string lotno, string process, int sel_num = 1)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcode, lotno });
            Dictionary<string, DataTable> dic_sum = Summary_Cycles_Logfile(sqlcon, itemcode, lotno, process, sel_num);
        start_lable: DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, process, filter_str);
            if (src_dt.Rows.Count > 0)
            {
                if (MessageBox.Show("Dữ liệu ItemCode / LotNo " + itemcode + " / " + lotno + "đã có. Bạn muốn cập nhật ?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    TDMK_Code.Delelte_FilteredItem_arr(process, sqlcon, filter_str);
                    goto start_lable;
                }
            }
            else
            {
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
                            int id = TDMK_Code.SQL_MAX(process, "ID", sqlcon);
                            for (int i = 0; i < cycle_dt.Value.Rows.Count; i++)
                            {
                                src_dt.Rows.Add((id + 1 + i).ToString(), itemcode, lotno, (i + 1), pcs);
                                src_dt.Rows[i + r_offset][cycle_name] = cycle_dt.Value.Rows[i][dc];
                            }
                        }
                    }
                }
                BatchBulkCopy(sqlcon, src_dt, process);
            }
            return src_dt;
        }
        public DataTable Summary_Logfile_All(SqlConnection sqlcon, string itemcode, string lotno, string process, string shift, int sel_num = 1, string season = "")
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Shift" }, new string[] { itemcode, lotno, shift });
            Dictionary<string, DataTable> dic_sum = Summary_Cycles_Logfile(sqlcon, itemcode, lotno, process, shift, sel_num, season);
            if (dic_sum != null)
            {
            start_lable: DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, process, filter_str);
                if (src_dt.Rows.Count > 0)
                {
                    if (MessageBox.Show("Dữ liệu ItemCode / LotNo / Shift " + itemcode + " / " + lotno + " / " + shift + " đã có. Bạn muốn cập nhật ?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        TDMK_Code.Delelte_FilteredItem_arr(process, sqlcon, filter_str);
                        goto start_lable;
                    }
                }
                else
                {
                    int id = TDMK_Code.SQL_MAX(process, "ID", sqlcon) + 1;
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
                                    src_dt.Rows.Add((id + i).ToString(), itemcode, lotno, (i + 1), pcs);
                                    src_dt.Rows[i + r_offset][cycle_name] = cycle_dt.Value.Rows[i][dc];
                                    src_dt.Rows[i + r_offset]["Shift"] = shift;
                                }
                                id += cycle_dt.Value.Rows.Count;
                            }
                        }
                    }
                    BatchBulkCopy(sqlcon, src_dt, process);
                    MessageBox.Show("Lưu dữ liệu hoàn thành", "Thông báo");
                }
                return src_dt;
            }
            else
            {
                MessageBox.Show("Thiếu dữ liệu Before", "Thông báo");
                return null;
            }

        }

        public void Export_Thermal_HeatSoak_Bend_MultiType_Type2(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_type)//string format_name, ,  
        {
            string format_loc = Path.Combine(format_folder, format_type);
            List<string> format_lst = get_multiple_files(format_loc, new List<string> { "*.xlsx", "*.xlsm" }, ItemCode, process_name);
            string format_file = "";
            if (format_lst.Count > 0)
            {
                format_file = format_lst[0];
            }
            if (format_file != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable _src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                DataTable _spec_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process_name }));
                DataTable log_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name + "_LOGFILE", filter_str);
                SortedDictionary<int, string> sel_index_lst = Get_bending_Net(sqlcon_OK2SHIP, ItemCode, process_name);
                List<int> index_lst = sel_index_lst.Keys.ToList();
                DataTable src_tbl = _src_tbl.AsEnumerable().Where(x => index_lst.Contains(Convert.ToInt32(x.Field<string>("Net_no")) - 1)).CopyToDataTable();
                DataTable spec_tbl = _spec_tbl.AsEnumerable().Where(x => x.Field<string>("Sel_Report") != "No").CopyToDataTable();
                List<string> NET_name = spec_tbl.AsEnumerable().Select(x => x.Field<string>("Net_Name")).ToList();
                List<DataTable> src_tbl_lst = new List<DataTable>();
                Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                List<char> reject_char_lst = new List<char> { ' ', '_', '-' };
                myExcel.Workbook report_wrk = null;// 
                myExcel.Worksheet tar_wrksht = null;
                string find_cycle_key = "After";
                report_wrk = TDMK_Code.open_excel_file(format_file, "", "");
                foreach (myExcel.Worksheet sht in report_wrk.Worksheets)
                {
                    if (remove_special_char(sht.Name, reject_char_lst).ToUpper() == remove_special_char(process_name, reject_char_lst).ToUpper())
                    {
                        tar_wrksht = sht;
                        break;
                    }
                }
                if (tar_wrksht != null)
                {
                    Dictionary<int, string> cycle_data_addr_lst = Get_Bending_Cycle_addr("T0 - Before (Ω)", find_cycle_key, tar_wrksht, "J1", false);
                    string NET_start_rgn = Get_start_range("Net", "D4", tar_wrksht);
                    string before_addr = Find_Cell_Addr("T0 - Before (Ω)", "J1", tar_wrksht);
                    int row_num = tar_wrksht.Range[before_addr].MergeArea.Rows.Count;
                    int before_column = tar_wrksht.Range[before_addr].Column;
                    myExcel.Range NET_rgn = tar_wrksht.Range[NET_start_rgn];
                    myExcel.Range No_rgn = NET_rgn.Offset[0, -3];
                    myExcel.Range pattern_rgn = NET_rgn.Offset[0, -2];
                    cycle_data_addr_lst.Add(0, before_addr);
                    string[] temp = Path.GetFileNameWithoutExtension(report_wrk.Name).Split('-');
                    string itemname = "";
                    if (temp.Length > 1)
                    {
                        itemname = temp[1];
                    }
                    string report_name = process_name + " " + itemname + "-" + ItemCode + "-" + LotNo + " " + DateTime.Now.ToString("yyyyMMdd") + Path.GetExtension(report_wrk.Name);
                    if (format_type == "NPI")
                    {
                        report_name = Path.GetFileNameWithoutExtension(report_wrk.Name) + "-" + LotNo + " " + DateTime.Now.ToString("yyyyMMdd") + Path.GetExtension(report_wrk.Name);
                    }
                    string daily_folder = Path.Combine(Report_location, "Daily_report");
                    if (!Directory.Exists(daily_folder))
                    {
                        Directory.CreateDirectory(daily_folder);
                    }
                    report_wrk.SaveAs(Path.Combine(daily_folder, report_name));
                    int col_inx = 0;
                    foreach (DataTable tbl in src_tbl_lst)
                    {
                        int r_offset = col_inx * NET_name.Count;
                        Dictionary<string, List<string>> data_lst = Get_List_Pair_data(tbl);
                        foreach (var item in data_lst)
                        {
                            List<string> sel_data = item.Value;
                            int cycle_no = 0;
                            if (item.Key.Contains("After"))
                            {
                                cycle_no = Convert.ToInt32(Extract_Num_from_String(item.Key));
                            }
                            if (cycle_data_addr_lst.Keys.ToList().IndexOf(cycle_no) != -1)
                            {
                                myExcel.Range cycle_rgn = tar_wrksht.Range[cycle_data_addr_lst[cycle_no]].Offset[row_num, 0];
                                if (cycle_no == 0)
                                {
                                    cycle_rgn = tar_wrksht.Range[cycle_data_addr_lst[cycle_no]].Offset[1, 0];
                                }
                                for (int r_inx = 0; r_inx < sel_data.Count; r_inx++)
                                {
                                    cycle_rgn.Offset[r_inx + r_offset, 0].Value = sel_data[r_inx];

                                    if (cycle_no == 0)
                                    {
                                        NET_rgn.Offset[r_inx + r_offset, 0].Value = NET_name[r_inx];
                                        No_rgn.Offset[r_inx + r_offset, 0].Value = r_inx + r_offset + 1;
                                        pattern_rgn.Offset[r_inx + r_offset, 0].Value = col_inx + 1;
                                    }
                                    else
                                    {
                                        cycle_rgn.Offset[r_inx + r_offset, 1].FormulaR1C1 = "=+(RC[-1]-RC" + before_column.ToString() + ")/RC" + before_column.ToString();
                                        cycle_rgn.Offset[r_inx + r_offset, 2].FormulaR1C1 = @"=+IFERROR(IF(ABS(RC[-1])>10%,IF(ABS(RC[-1])<100%,""10% ~ 100%"",IF(ABS(RC[-1])<500%,""100% ~ 500%"",""Over 500%"")),""Under 10%""),""Under 10%"")";
                                    }
                                }
                            }

                        }
                        col_inx++;
                    }
                    report_wrk.Save();

                }
                else
                {
                    MessageBox.Show("Format của ItemCode " + ItemCode + " không đúng", "Cảnh báo");
                }
            }
            else
            {
                MessageBox.Show("Format of ItemCode " + ItemCode + " not found", "Warning");
            }
        }
        public string find_config_path(string src_string, string f_name)
        {
            DirectoryInfo di = new DirectoryInfo(src_string);
            var folder_lst = di.GetDirectories().Select(x => x.FullName).ToList();
            var temp2 = folder_lst.Where(x => new DirectoryInfo(x).Name == f_name).ToList();
            if (temp2.Count != 0)
            {
                return temp2.FirstOrDefault();
            }
            else
            {
                if (di.Parent != null)
                {
                    return find_config_path(di.Parent.FullName, f_name);
                }
                else
                {
                    return string.Empty;
                }
            }
        }

    }
}
