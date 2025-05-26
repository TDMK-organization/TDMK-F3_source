using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using myExcel = Microsoft.Office.Interop.Excel;
using TDMK_SQL;
using System.Drawing;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using TDMK_SEEV_DLL;
using Microsoft.Office.Core;
using OK2SHIP;
using Echeck_LogFile_Process;
using Manual_Input;

namespace VHX
{
    public class myVar
    {
        public static FrmMain _frmMain;
        public static FrmCamera _frmCAM;
        public static FrmMainVHX _frmVHX;
        public static ECheck_Main _frmECheck;
        public static Frm_Manual_input _frmManual;
        public static Check_Impedance_logfile _frmCheckLogfile;
        public static bHast_Graph _frmbHastGraph;
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SEI_Lib myCode = new SEI_Lib();
        public static string data_loc;
        Export_Image_Class img_proc = new Export_Image_Class();
        public static SqlConnection mysqlcon;
        public static string server_name = "";
        public static string server_acc = "";
        public static string server_pass = "";
        public static string sel_DB = "OK2SHIP_Period2";
       // public static string sel_DB = "tdm06349_DB";
        public struct Impedance_data
        {
            public string data_val { get; set; }
            public byte[] grap_data { get; set; }
            public Impedance_data( string in_data, byte[] in_graph)
            {
                data_val = in_data;
                grap_data = in_graph;
            }
        }
        public struct Impedance_Graph
        {
            public byte[] img_data { get; set; }
            public byte[] grap_data { get; set; }
            public Impedance_Graph(byte[] in_data, byte[] in_graph)
            {
                img_data = in_data;
                grap_data = in_graph;
            }
        }
        public string Find_Export_Path(string format_file, string tar_process)
        {
            string result = "";
            DirectoryInfo d = new DirectoryInfo(format_file);
            DirectoryInfo root_d = d.Parent.Parent;
            DirectoryInfo[] d_arr = root_d.GetDirectories();
            string[] myLst = d_arr.Where(x => !x.Name.Contains("TDMK")).Where(x => x.Name.ToUpper().Split('.')[1].Replace("-", "").Replace(" ", "") == tar_process.Replace("_", "").Replace("-", "").Replace(" ","").ToUpper()).Select(x => x.Name).ToArray();
            if (myLst.Length > 0)
            {
               result = myLst[0];
            }
            return result;
        }
        public string find_format(string in_data_loc, string ItemCode)
        {
            string result = "";
            string[] files = Directory.GetFiles(Path.Combine(in_data_loc, "0.OK2SHIP report format"), "*" + ItemCode + "*.xlsm");
            if(files.Length>0)
            {
                result = files[0];
            }
            return result;
        }
        public myExcel.Workbook create_export_wrk_old(string format_file, string process_name, string ItemCode, string Lotno)
        {
            string export_path = Path.Combine(data_loc, Find_Export_Path(format_file, process_name));
            string savedfile = Path.Combine(export_path, ItemCode + "-" + Lotno) + ".xlsx";
            if (File.Exists(savedfile))
            {
                return TDMK_Code.open_excel_file(savedfile, "", "");
            }
            else
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
                    if (!Directory.Exists(export_path))
                    {
                        Directory.CreateDirectory(export_path);
                    }
                    myExcel.Workbook save_wb = TDMK_Code.Create_workbook();
                    wb.Sheets[mySheet].Copy(After: save_wb.Sheets[1]);
                    save_wb.SaveAs(savedfile);
                    myExcel.Worksheet del_sht = save_wb.Sheets[1];                    
                    del_sht.Delete();
                   // save_wb.Save();
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
        public byte[] get_image_excel(myExcel.Worksheet wrk_sheet)
        {
            myExcel.Shape cur_image = wrk_sheet.Shapes.Item("graph1");
            cur_image.Copy();
            Byte[] data = new Byte[0];
            Image myImg = Clipboard.GetImage();
            ImageConverter imgCon = new ImageConverter();
            data = (byte[])imgCon.ConvertTo(myImg, typeof(byte[]));
            return data;
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
        public SqlConnection initial_data(string DB_name, bool sa_en)
        {
            SqlConnection _sqlcon_OK2SHIP;
            string app_path = Application.StartupPath;
            string config_file = Path.Combine(app_path.Replace("\\FPCA OK2SHIP Auto System\\VHX-IMADA", ""), "Config.ini");
            string[] my_config = myCode.read_config_arr(config_file);

            foreach (string c in my_config)
            {
                if (c.Contains("Server"))
                {
                    server_name = c.Split(':')[1].Trim();
                }
                if (c.Contains("Account"))
                {
                    server_acc = c.Split(':')[1].Trim();
                }
                if (c.Contains("Password"))
                {
                    server_pass = c.Split(':')[1].Trim();
                }
                if (c.Contains("Data_Location"))
                {
                    data_loc = c.Split('#')[1].Trim();
                }
            }
            if (sa_en)
            {
                string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, DB_name, server_acc, server_pass).ConnectionString;
                _sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);
            }
            else
            {
                string _strcon = TDMK_Code.data_connection2(server_name, DB_name).ConnectionString;
                _sqlcon_OK2SHIP = new SqlConnection(_strcon);
            }
            mysqlcon = _sqlcon_OK2SHIP;
            return _sqlcon_OK2SHIP;
        }
        public void Export_DatatableImage_Excel(DataTable dt, string Image_Col_name, myExcel.Worksheet tar_wrksht, string tar_rgn, bool row_offset)
        {
            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = Path.Combine(file_folder, DateTime.Now.ToString("ddMMMyy_HHmmss") +".jpg");
            int r_inx = 0;
            myExcel.Range sel_rgn = tar_wrksht.Range[tar_rgn];
            foreach (DataRow dr in dt.Rows)
            {
                byte[] img_byte = (byte[])(dr[Image_Col_name]);
                using (MemoryStream ms = new MemoryStream(img_byte))
                {
                    Image temp = Image.FromStream(ms);
                    temp.Save(file_dic);
                }
                img_proc.InsertPicture_Name(tar_wrksht, sel_rgn, file_dic, 5);
                int col_offset = sel_rgn.Columns.Count;
                int row_off = sel_rgn.Rows.Count;
                if (row_offset)
                {
                    sel_rgn = sel_rgn.Offset[row_off, 0];
                }
                else
                {
                    sel_rgn = sel_rgn.Offset[0, col_offset];
                }
                r_inx++;
            }
            try
            {
                File.Delete(file_dic);
            }
            catch
            {

            }
            
        }
        public void Export_DatatableImage_Excel_2(DataTable dt, string Image_Col_name, myExcel.Worksheet tar_wrksht, myExcel.Range sel_rgn, bool row_offset)
        {
            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = Path.Combine(file_folder, DateTime.Now.ToString("ddMMMyy_HHmmss") + ".jpg");
            int r_inx = 0;
            // myExcel.Range sel_rgn = tar_wrksht.Range[tar_rgn];
            foreach (DataRow dr in dt.Rows)
            {
                byte[] img_byte = (byte[])(dr[Image_Col_name]);
                using (MemoryStream ms = new MemoryStream(img_byte))
                {
                    Image temp = Image.FromStream(ms);
                    temp.Save(file_dic);
                }
                img_proc.InsertPicture_Name(tar_wrksht, sel_rgn, file_dic, 5);
                int col_offset = sel_rgn.Columns.Count;
                int row_off = sel_rgn.Rows.Count;
                if (row_offset)
                {
                    sel_rgn = sel_rgn.Offset[row_off, 0];
                }
                else
                {
                    sel_rgn = sel_rgn.Offset[0, col_offset];
                }
                r_inx++;
            }
            try
            {
                File.Delete(file_dic);
            }
            catch
            {

            }

        }
        public void Insert_Object_List(string tar_table, SqlConnection in_sqlcon, List<string> item, List<object> item_val)
        {
            string col_name = "";
            string col_val = "";
            int qty = Math.Min(item.Count, item_val.Count);
            for (int i = 0; i < qty; i++)
            {
                if (!(item_val[i] is DBNull))
                {
                    col_name = col_name + item[i] + ",";
                    col_val = col_val + "@" + item[i] + ",";
                }
            }
            col_name = col_name.TrimEnd(',');
            col_val = col_val.TrimEnd(',');
            string query = "Insert into " + tar_table + "(" + col_name + " ) values (" + col_val + ")";
            using (SqlCommand cmd = new SqlCommand(query))
            {
                cmd.Connection = in_sqlcon;
                for (int i = 0; i < qty; i++)
                {
                    if (!(item_val[i] is DBNull))
                    {
                        cmd.Parameters.AddWithValue("@" + item[i], item_val[i]);
                    }
                }
                in_sqlcon.Open();
                cmd.ExecuteNonQuery();
                in_sqlcon.Close();
            }
        }
        public string Lotno_Formated(string lotno)
        {
            string result = "";
            if (lotno.All(char.IsDigit))
            {
                try
                {
                    result = string.Format("{0:00000}", Convert.ToInt32(lotno));
                    result = result.Substring(0, 5);
                }
                catch { }
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
        public void DGV_Image_Fitting(DataGridView src_DGV)
        {
            foreach (DataGridViewColumn dc in src_DGV.Columns)
            {
                string col_name = dc.Name;
                if (col_name.ToUpper().Contains("IMAGE"))
                {
                    ((DataGridViewImageColumn)dc).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    dc.Width = 200;
                }
            }
            foreach (DataGridViewRow dr in src_DGV.Rows)
            {
                dr.Height = 150;
            }
        }
        public myExcel.Workbook Table_To_CSV(string tbl_name, DataTable src_dt)
        {
            myExcel.Workbook result = null;
            if (src_dt.Rows.Count > 0)
            {
                string temp_folder = Path.Combine(Application.StartupPath, "Temp");
                if (!Directory.Exists(temp_folder))
                {
                    Directory.CreateDirectory(temp_folder);
                }
                string strFilePath = Path.Combine(temp_folder, tbl_name + ".csv");
                StreamWriter sw = new StreamWriter(strFilePath, false);
                string line = "No,";
                foreach (DataColumn dc in src_dt.Columns)
                {
                    line += dc.ColumnName + ",";
                }
                line = line.Trim(',');
                sw.WriteLine(line);
                int r_inx = 0;
                foreach (DataRow dr in src_dt.Rows)
                {
                    line = (r_inx+1).ToString()+",";
                    foreach (DataColumn dc in src_dt.Columns)
                    {
                        line += myCode.checkDBNull(dr[dc]) + ",";
                    }
                    line = line.Trim(',');
                    sw.WriteLine(line);
                    r_inx++;
                }
                sw.Close();
                result = TDMK_Code.open_excel_file(strFilePath, "", "");
            }
            return result;
        }
        public bool check_column_exited(DataTable dt, string find_col_name, ref int col_inx)
        {
            bool result = false;
            for(int i=0;i<dt.Columns.Count;i++)
            {
                string dt_col_name = dt.Columns[i].ColumnName;
                if(dt_col_name.Replace(" ","").ToUpper()==find_col_name.Replace(" ","").ToUpper())
                {
                    result = true;
                    col_inx = i;
                    break;
                }
            }
            return result;
        }
    }
}
