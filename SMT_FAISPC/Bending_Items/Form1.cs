using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SEEV_DLL;
using TDMK_SQL;
//using Echeck_LogFile_Process;
using myExcel = Microsoft.Office.Interop.Excel;


namespace Bending_Items
{
    public partial class Form1 : Form
    {
        SEI_Lib myCode = new SEI_Lib();
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        ECheck_Process exp_process = new ECheck_Process();
        Bending_SMT_Lib bending_proc = new Bending_SMT_Lib();
        SqlConnection sqlcon_OK2SHIP;
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
        public Form1()
        {
            InitializeComponent();
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            List<string> Net = new List<string>();
            DataTable dt = new DataTable();
            Dictionary<string, string> Log_details = new Dictionary<string, string>();
            //DGV_Data.DataSource = bending_proc.DAT_To_DataTable(txtLogfile_before.Text);
            string userid = "";
            string itemname = "";
            DGV_Data.DataSource = bending_proc.DAT_To_DataTable_details_time(txtlogfile_After.Text,ref Net,ref dt, ref Log_details,ref itemname, ref userid);
            myCode.DGV_Auto_Resize(DGV_Data);
            DGV_Data2.DataSource = dt;
            //txtUUT_No.Text = log_time;
            /*char[] split_char = new char[] { ' ', '(', '-','_' };
            string src_txt1 = "S1-55(I/O-I/O)_S1-55(I/O-I/O)";
            string src_txt2 = "OPEN B2B(401-66";
            var arr = src_txt1.Replace("S","").Split(split_char).Distinct().ToArray();
            var arr2 = src_txt2.Replace("S", "").Split(split_char).Distinct().ToArray();
            lstData_Addr.DataSource = arr.Where(x=> myCode.IsNumeric(x)).ToList();
            lstNet.DataSource = arr2.Where(x => myCode.IsNumeric(x)).ToList();*/
            //lstData_Addr.DataSource = get_multiple_files(txtLogfile_before.Text,new List<string> { "*.csv","*.dat","*.xls","*.xlsx"}, "7S0030","Flex bending");
            //lstData_Addr.DataSource = get_multiple_files(txtLogfile_before.Text, new List<string> { "*.*" }, "7S0030", "Flex bending");
        }
        public List<string> get_multiple_files(string src_path, List<string> extensions, string tar_ItemCode, string tar_process)
        {
            List<string> result = new List<string>();
            List<char> remove_char = new List<char> { ' ', '-', '_' };
            string process =new string(tar_process.Where(x => remove_char.IndexOf(x) == -1).ToArray());
            DirectoryInfo directory = new DirectoryInfo(src_path);
            var files = extensions.SelectMany(e => directory.EnumerateFiles(e, SearchOption.AllDirectories)).Where(x=>x.FullName.Contains(tar_ItemCode)&& remove_special_char(Path.GetFileNameWithoutExtension(x.Name), remove_char).ToUpper()== process.ToUpper());
            foreach(var item in files)
            {
                result.Add(item.FullName);
            }
            return result;
        }
        public string remove_special_char(string src_str, List<char> remove_char)
        {
            return new string(src_str.Where(x => remove_char.IndexOf(x) == -1).ToArray());
        }
        public void Marking_DGV_Rows_Net(DataGridView src_DGV, List<string> src_Net=null)
        {

            int r = 0;
            foreach (DataGridViewRow dgr in src_DGV.Rows)
            {
                if(src_Net!=null)
                {
                    if(r<src_Net.Count)
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

        }
        public List<string> Get_Bending_data_addr(string search_item,string search_key, myExcel.Worksheet tar_wrksht, string start_range_addr)
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
            DataTable filter_dt = dt.AsEnumerable().Where(x => !x.Field<string>("Test Title").Contains("*") && x.Field<string>("Test Title")!= "Type").CopyToDataTable();
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
                    src_data.Add( newline);
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
                if(c!="")
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
                    if(proc_en)
                    {
                        proc_en = false;
                    }
                    item_count++;
                    Raw_Bending_data.Add(item_count, new List<string>() { values[5] });
                }
                else
                {
                    if(rec_en)
                    {
                        if(!proc_en)
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
            foreach(var data in Raw_Bending_data.Values)
            {
                if(!data[0].Contains("No_UUT_ID"))
                {
                    for (int k = 1; k < data.Count; k++)
                    {
                        if(!data[k].Contains("OPEN") && !data[k].Contains("SUS"))
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
        public DataTable DAT_To_DataTable_details(string _filename, ref List<string> Net_list, ref DataTable tar_spec_dt)
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
            foreach(string col in spec_col_lst)
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
                            if(!myCode.check_columns_existed(result_dt,col_name))
                            {
                                result_dt.Columns.Add(col_name);
                            }
                            if(result_dt.Rows.Count<k)
                            {
                                result_dt.Rows.Add();
                            }
                            result_dt.Rows[k - 1][col_name] = detail_data[5];
                            if(Net_list.IndexOf(detail_data[6])==-1)
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
        public Dictionary<string, Dictionary<int, bending_data>> Bending_process(DataTable dt_before, DataTable dt_after)
        {
            Dictionary<string, Dictionary<int, bending_data>> result = new Dictionary<string, Dictionary<int, bending_data>>();
            List<string> bef_lst = Get_column_name(dt_before);
            List<string> aft_lst = Get_column_name(dt_after);
            var net_lst = bef_lst.Intersect(aft_lst).ToList();
            int row_inx = Math.Min(dt_before.Rows.Count, dt_after.Rows.Count);
            foreach(var net in net_lst)
            {
                result.Add(net, new Dictionary<int, bending_data>());
                for(int r=0;r< row_inx; r++)
                {
                    string bef_val = dt_before.Rows[r][net].ToString();
                    string aft_val = dt_after.Rows[r][net].ToString();
                    result[net].Add(r + 1, new bending_data(bef_val, aft_val,Resistance_vary(bef_val,aft_val)));
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
            foreach(DataColumn dc in src_dt.Columns)
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

        private void Form1_Load(object sender, EventArgs e)
        {
            sqlcon_OK2SHIP = bending_proc.initial_data("OK2SHIP_SMT", true);
        }
        public void Export_Bending(SqlConnection sqlcon_OK2SHIP, string ItemCode, string LotNo, string process_name, string format_name)
        {
            string f_loca = "";
            string format_file = @"D:\Customer Projects\SEEV\SMT Project\temp\Flex Bending_Export.xlsx";// bending_proc.find_format(f_loca, ItemCode);
            if (format_file != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
                DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process_name, filter_str);
                List<DataTable> src_tbl_lst = new List<DataTable>();
                bending_proc.Get_ListTable(-1, src_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_tbl_lst, "Before");
                myExcel.Workbook report_wrk = bending_proc.create_export_wrk(format_file, format_name);
                myExcel.Worksheet tar_wrksht = report_wrk.Sheets[1];
                List<string> data_addr_lst = bending_proc.Get_Bending_data_addr("Test item", "Sample", tar_wrksht, "A10",false);
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
        public myExcel.Workbook create_export_wrk(string format_file, string process_name, string ItemCode, string Lotno)
        {
            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            xlsApp.DisplayAlerts = false;
            myExcel.Workbook wb = TDMK_Code.open_excel_file(format_file, "", "");
            string _process = process_name.Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
            string mySheet = "";
            foreach (myExcel.Worksheet tg_sht in wb.Worksheets)
            {
                string cur_sht_name = tg_sht.Name.ToUpper().Replace("-", "").Replace(" ", "");
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

        private void DGV_Data2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnTestSum_Click(object sender, EventArgs e)
        {
            string itemcode = "71804P";//"71370";//
            string lotno = "00324";
            string process = "FLEX_BENDING";

            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcode, lotno });
            
            DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process, filter_str);
            //bending_proc.Export_Thermal_HeatSoak_Bend_MultiType_Type2(sqlcon_OK2SHIP, itemcode, lotno, process, "Other");

            DGV_Data.DataSource =  bending_proc.Summary_Logfile_All(sqlcon_OK2SHIP, itemcode, lotno, process,"1",20);

            //Dictionary<string, DataTable> dic_sum = bending_proc.Summary_Cycles_Logfile(sqlcon_OK2SHIP, itemcode, lotno, process, 20);
            //string cycle_name = "Before";
            //DataTable cycle_dt = dic_sum[cycle_name];
            //DGV_Data2.DataSource = cycle_dt;
            //foreach (DataColumn dc in cycle_dt.Columns)
            //{
            //    int r_offset = src_dt.Rows.Count;
            //    string pcs = dc.ColumnName;
            //    DataView dv = src_dt.AsDataView();
            //    dv.RowFilter = "Pcs_No = '" + pcs + "'";
            //    if (dv.Count > 0)
            //    {
            //        if (dv.Count == cycle_dt.Rows.Count)
            //        {
            //            for (int i = 0; i < dv.Count; i++)
            //            {
            //                dv[i][cycle_name] = cycle_dt.Rows[i][dc];
            //            }
            //        }
            //        else
            //        {
            //            MessageBox.Show("Dữ liệu chu kỳ " + cycle_name + " không đủ", "Thông báo");
            //        }
            //    }
            //    else
            //    {
            //        int id = TDMK_Code.SQL_MAX(process, "ID", sqlcon_OK2SHIP);
            //        for (int i = 0; i < cycle_dt.Rows.Count; i++)
            //        {
            //            src_dt.Rows.Add((id + 1 + i).ToString(), itemcode, lotno, (i + 1), pcs);
            //            src_dt.Rows[i+r_offset][cycle_name] = cycle_dt.Rows[i][dc];
            //        }
            //    }
            //}
            //DGV_Data.DataSource = src_dt;
            myCode.DGV_Auto_Resize(DGV_Data);

            //myExcel.Workbook src_wrkbk = TDMK_Code.open_excel_file(txtFormatFile.Text, "", "");
            //myExcel.Worksheet src_wrksht = src_wrkbk.Sheets["Thermal cycling ICT  data"];
            //string before_addr = bending_proc.Find_Cell_Addr("T0 - Before (Ω)", "J1", src_wrksht);
            //Dictionary<int, string> cycle_addr_lst =  bending_proc.Get_Bending_Cycle_addr("T0 - Before (Ω)", "After", src_wrksht, "J1", false);
            //int row_num = src_wrksht.Range[before_addr].MergeArea.Rows.Count;
            //string NET_start_rgn = Get_start_range("Test item", "A10", src_wrksht);
            //string last_cycle_addr = cycle_addr_lst.LastOrDefault().Value;
            //int col_num = src_wrksht.Range[last_cycle_addr].MergeArea.Columns.Count;
            //int total_col_num = src_wrksht.Range[last_cycle_addr].Column + col_num - 1;
            //int r_off = 10;
            //myExcel.Range final_rgn = src_wrksht.Range[src_wrksht.Range[NET_start_rgn], src_wrksht.Range[NET_start_rgn].Offset[r_off, total_col_num - 1]];
            //List<myExcel.XlBordersIndex> st_lst = new List<myExcel.XlBordersIndex>() {  myExcel.XlBordersIndex.xlEdgeLeft, myExcel.XlBordersIndex.xlEdgeTop, myExcel.XlBordersIndex.xlEdgeBottom, myExcel.XlBordersIndex.xlEdgeRight, myExcel.XlBordersIndex.xlInsideHorizontal, myExcel.XlBordersIndex.xlInsideVertical };//myExcel.XlBordersIndex.xlDiagonalDown, myExcel.XlBordersIndex.xlDiagonalUp,
            //foreach (var st in st_lst)
            //{
            //    final_rgn.Borders[st].LineStyle = myExcel.XlLineStyle.xlContinuous;
            //}
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
        public Dictionary<string, DataTable> Summary_Cycles_Logfile(string itemcode, string lotno, string process)
        {
            Dictionary<string, DataTable> dic_sum = new Dictionary<string, DataTable>();
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcode, lotno });
            string spec_filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { itemcode, process });
            DataTable src_log_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process + "_LOGFILE", filter_str);
            List<string> cycle_lst = src_log_dt.AsEnumerable().Select(x => x.Field<string>("Cycles_name")).Distinct().ToList();
            foreach (string cycle in cycle_lst)
            {
                dic_sum.Add(cycle, new DataTable());
                List<string> logfile_lst = src_log_dt.AsEnumerable().Where(x => x.Field<string>("Cycles_name") == cycle).Select(x => x.Field<string>("Logfile")).Distinct().ToList();
                Dictionary<string, List<string>> dic_logfile = bending_proc.Get_BlockofLogFile(logfile_lst);
                Dictionary<string, DataTable> result_table_lst = new Dictionary<string, DataTable>();
                Dictionary<string, List<DataTable>> dic_tbl_data_lst = new Dictionary<string, List<DataTable>>();
                DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", spec_filter_str);
                foreach (var block in dic_logfile)
                {
                    foreach (var logfile in block.Value)
                    {
                        DataTable dt = new DataTable();
                        List<string> temp_net_lst = new List<string>();
                        dt = bending_proc.Load_Log_Data(sqlcon_OK2SHIP, itemcode, lotno, process, cycle, logfile, temp_net_lst);

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
                    DataTable block_tbl = bending_proc.Summary_Data_Table_from_list(tbl.Value, spec_dt);
                    result_table_lst.Add(tbl.Key, block_tbl);
                }
                List<string> pcs_lst = new List<string>();
                pcs_lst = bending_proc.Get_column_name(dic_sum["Before"]);
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
                    int sel_num = 1;
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
        public Dictionary<string, DataTable> Summary_Cycles_Logfile(string itemcode, string lotno, string process, string shift)
        {
            Dictionary<string, DataTable> dic_sum = new Dictionary<string, DataTable>();
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcode, lotno });
            string spec_filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { itemcode, process });
            DataTable src_log_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, process + "_LOGFILE", filter_str);
            List<string> cycle_lst = src_log_dt.AsEnumerable().Select(x => x.Field<string>("Cycles_name")).Distinct().ToList();
            foreach (string cycle in cycle_lst)
            {
                dic_sum.Add(cycle, new DataTable());
                List<string> logfile_lst = src_log_dt.AsEnumerable().Where(x => x.Field<string>("Cycles_name") == cycle).Select(x => x.Field<string>("Logfile")).Distinct().ToList();
                Dictionary<string, List<string>> dic_logfile = bending_proc.Get_BlockofLogFile(logfile_lst);
                Dictionary<string, DataTable> result_table_lst = new Dictionary<string, DataTable>();
                Dictionary<string, List<DataTable>> dic_tbl_data_lst = new Dictionary<string, List<DataTable>>();
                DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP, "NET_SPEC", spec_filter_str);
                foreach (var block in dic_logfile)
                {
                    foreach (var logfile in block.Value)
                    {
                        DataTable dt = new DataTable();
                        List<string> temp_net_lst = new List<string>();
                        string userid = "";
                        string itemname = "";
                        dt = bending_proc.Load_Log_Data(sqlcon_OK2SHIP, itemcode, lotno, process, cycle,shift,ref userid,ref itemname, logfile, temp_net_lst);

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
                    DataTable block_tbl = bending_proc.Summary_Data_Table_from_list(tbl.Value, spec_dt);
                    result_table_lst.Add(tbl.Key, block_tbl);
                }
                List<string> pcs_lst = new List<string>();
                pcs_lst = bending_proc.Get_column_name(dic_sum["Before"]);
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
                    int sel_num = 1;
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
    }
}
