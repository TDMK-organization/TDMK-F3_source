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
using TDMK_SQL;
using myExcel = Microsoft.Office.Interop.Excel;
using OK2SHIP_Lib;
using OfficeOpenXml;
using TDMK_SEEV_DLL;
using FAI_Export_EPPLUS;
using TDMK_EPPLUS_7;
using IniLibs;
using System.Diagnostics;

namespace OK2SHIP_Measurements
{
    public partial class TestFunction : Form
    {
        IPQC_LogFile IPQC_LogFile = new IPQC_LogFile();
        FAI_EPPLUS_Lib FAI_lib = new FAI_EPPLUS_Lib();
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        //SEI_Lib myCode = new SEI_Lib();
        public TDMK_OK2SHIP myCode = new TDMK_OK2SHIP();
        public TDMK_EPPLUS7_lib TDMK_EPPLUS7 = new TDMK_EPPLUS7_lib();
        public IniFile TDMK_init = new IniFile("Config.ini");
        public static SqlConnection sqlcon_FAI;
        public static SqlConnection sqlcon_Declare;
        public static SqlConnection sqlcon_Materials;
        public static SqlConnection sqlcon_Recycle;
        public static SqlConnection sqlcon_IPQC;
        public static SqlConnection sqlcon_OK2SHIP;
        public static SqlConnection sel_sqlcon;
        public string server_name;
        public string server_acc;
        public string server_pass;
        public string DB_name;
        public string app_path;
        public string Data_Location;
        public string format_folder;
        public string log_folder;
        public string f_ext;
        public bool editmode = false;
        public TestFunction()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private void TestFunction_Load(object sender, EventArgs e)
        {
            sel_sqlcon = initial_data("OK2SHIP_Period2", true);//initial_data();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog f_file = new OpenFileDialog();
            if(f_file.ShowDialog() == DialogResult.OK)
            {
                //myCode.Mitutoyo_process(f_file.FileName, DGV_DataView, DGV_SpecView, DGV_Data_Plus, editmode);
                string f_name = f_file.FileName;
                DataTable FAI_logfile_tbl = myCode.PTH_Diameter_Logfile(f_name);
                DataTable tbl_spec = (DataTable)DGV_SpecView.DataSource;
                DataTable FAI_tbl = new DataTable();
                DataTable non_FAI = new DataTable();
                DataTable FAI_arranged_tbl = myCode.filter_table(FAI_logfile_tbl, tbl_spec, ref FAI_tbl, ref non_FAI);
                if (!editmode)
                {
                    DataTable dest_FAI_tbl = myCode. result_add((DataTable)DGV_DataView.DataSource, FAI_tbl, true);
                    DataTable dest_nonFAI_tbl = myCode.result_add((DataTable)DGV_Data_Plus.DataSource, non_FAI, true);
                    DGV_DataView.DataSource = null;
                    DGV_Data_Plus.DataSource = null;
                    DGV_DataView.DataSource = myCode.Order_table_inSpec(dest_FAI_tbl, tbl_spec);
                    DGV_Data_Plus.DataSource = dest_nonFAI_tbl;
                }
                else
                {
                    DGV_Data_Plus.DataSource = null;
                    DGV_Data_Plus.DataSource = myCode.Order_table_inSpec(FAI_arranged_tbl, tbl_spec);

                }
                myCode.check_FAIdata_inSpec(DGV_DataView, DGV_SpecView);
                myCode.check_FAIdata_inSpec(DGV_Data_Plus, DGV_SpecView);

            }
        }
        public void Mitutoyo_process(string _filename, DataGridView tar_DGV, DataGridView DGV_Spec, DataGridView DGV_plus, bool edit_mode)
        {
            string f_name = _filename;
            DataTable logfile_dt = new DataTable();
            logfile_dt = myCode.Mitutoyo_Get_raw_data2(f_name);
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
            myCode.DGV_Auto_Resize(tar_DGV);
            myCode.DGV_Auto_Resize(DGV_plus);

        }
        public DataTable result_add(DataTable dest_tbl, DataTable src_inData, bool col_add_en)
        {
            if(dest_tbl!=null)
            {
                int dest_row_count = dest_tbl.Rows.Count;
                foreach (DataColumn dc in src_inData.Columns)
                {
                    int start_row = 0;
                    if (myCode.check_columns_existed(dest_tbl, dc.ColumnName))
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
                        if(col_add_en)
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
            foreach(DataGridViewColumn dc in src_DGV_data.Columns )
            {
                string cur_col = dc.Name;
                if(spec_col_lst.IndexOf(cur_col)!=-1)
                {
                    string Tol_Max = myCode.checkDBNull(src_DGV_Spec.Rows[1].Cells[cur_col].Value);
                    string Tol_Min = myCode.checkDBNull(src_DGV_Spec.Rows[2].Cells[cur_col].Value);
                    string SetVal = myCode.checkDBNull(src_DGV_Spec.Rows[0].Cells[cur_col].Value);
                    double UL = Convert.ToDouble(SetVal) + Convert.ToDouble(Tol_Max);
                    double LL = Convert.ToDouble(SetVal) - Convert.ToDouble(Tol_Min);
                    foreach(DataGridViewRow dr in src_DGV_data.Rows)
                    {
                        string src_act_val = myCode.checkDBNull(dr.Cells[cur_col].Value);
                        dr.Cells[cur_col].Style.BackColor = myCode.check_in_limit_Color(UL.ToString(), LL.ToString(), src_act_val, SetVal);
                    }
                }
                else
                {
                    foreach (DataGridViewRow dr in src_DGV_data.Rows)
                    {
                        string src_act_val = myCode.checkDBNull(dr.Cells[cur_col].Value);
                        if(src_act_val!="")
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
            col_name = logfile_dt.AsEnumerable().Select(r => r.Field<string>(tar_Col)).Distinct().ToArray();
            List<string> col_lst = col_name.ToList();
            //lstSTT.DataSource = col_lst.OrderBy(x => x).ToList();
            List<DataTable> src_tbl_lst = new List<DataTable>();
            foreach(DataRow dr in logfile_dt.Rows)
            {
                dr[vitri_Set_val] = Convert.ToDouble(dr[vitri_Set_val]);
            }
            myCode.Get_ListTable(-1, logfile_dt, new string[] { tar_Col, vitri_Set_val }, ref src_tbl_lst, vitri_act_val);
            Dictionary<string, List<string>> dicFAI_data = new Dictionary<string, List<string>>();
            foreach (DataTable dt in src_tbl_lst)
            {
                string fai_name = dt.Rows[0][tar_Col].ToString();
                string setval = dt.Rows[0][vitri_Set_val].ToString();
                string FAI = fai_name + "_" + setval;
                string vitri_val = dt.Rows[0][vitri_col_name].ToString();
                result_tbl.Columns.Add(FAI);
                List<string> data = dt.AsEnumerable().Select(x => x.Field<string>(vitri_act_val)).ToList();
                if (!vitri_val.ToUpper().Contains("ANGLE"))
                {
                    for (int i = 0; i < data.Count; i++)
                    {
                        data[i] = Math.Abs(Convert.ToDouble(data[i])).ToString();
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
                string t_sv = myCode.IsNumeric_Val(t.Split('_').Last());
                string t_fai = t.Split('_').First();
                int col_inx = -1;
                if (myCode.check_columns_existed_inx(spec_dt, t_fai, ref col_inx))
                {
                    string tar_sv = myCode.IsNumeric_Val(spec_col_lst[col_inx].Split('_').Last());
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
            outFAI_dt = FAI_dt.AsDataView().ToTable(false, tar_col_lst.ToArray());
            outUnknowndt = FAI_dt.AsDataView().ToTable(false, ex_col.ToArray());
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
                string t_sv = myCode.IsNumeric_Val(t.Split('_').Last());
                string t_fai = t.Split('_').First();
                int col_inx = -1;
                if (myCode.check_columns_existed_inx(spec_dt, t_fai, ref col_inx))
                {
                    string tar_sv = myCode.IsNumeric_Val(spec_col_lst[col_inx].Split('_').Last());
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
        public List<string>[]Roughness_Result(string f_name, int MeasLoc)
        {
            List<string> mylstSa1 = new List<string>();
            List<string> mylstSa2 = new List<string>();
            List<string> mylstSq2 = new List<string>();
            List<string>[] Sa_result = new List<string>[] { mylstSa1, mylstSa2 };
            List<string> mylstSq1 = new List<string>();
            List<string>[] Sq_result = new List<string>[] { mylstSq1, mylstSq2 };
            List<string>[] _result = new List<string>[] { mylstSa1, mylstSa2, mylstSq1, mylstSq2 };
            myExcel.Workbook tar_wrkbook = TDMK_Code.open_excel_file(f_name, "", "");
            myExcel.Worksheet tar_wrksht = tar_wrkbook.Sheets[1];
            myExcel.Range tar_rgn = tar_wrksht.Range["B2"];
            int inx = 0;
            while (myCode.checkDBNull(tar_rgn.Offset[inx, 0].Value) != "")
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
                            double Sa_val = Convert.ToDouble(myCode.checkDBNull(tar_rgn.Offset[inx, 0].Value));
                            double Sq_val = Convert.ToDouble(myCode.checkDBNull(tar_rgn.Offset[inx, 1].Value));                            
                            Sa_result[i].Add(Math.Round(Sa_val, 3).ToString());//"#0.##0"
                            Sq_result[i].Add(Math.Round(Sq_val, 3).ToString());//"#0.##0"
                        }
                    }
                }
                inx++;
            }
            return _result;
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            string log_path = @"D:\Customer Projects\SEEV\OK2SHIP_GD3\Logfile\Echeck_Pcs_Type\GUITDMK4\GUITDMK4\VHX\422435-02404_506377+BL 154-4";
            DGV_Data.DataSource = VHX_FAI_Data_Process(log_path);
                


            //DataTable logfile_dt = (DataTable)DGV_Data.DataSource;
            //string[] col_name;
            //string tar_Col = logfile_dt.Columns[1].ColumnName;
            //string vitri_col_name = logfile_dt.Columns[0].ColumnName;
            //string vitri_act_val = logfile_dt.Columns[4].ColumnName;
            //string vitri_Set_val = logfile_dt.Columns[5].ColumnName;
            //string vitri_UL_val = logfile_dt.Columns[7].ColumnName;
            //string vitri_LL_val = logfile_dt.Columns[8].ColumnName;
            //col_name = logfile_dt.AsEnumerable().Select(r => r.Field<string>(tar_Col)).Distinct().ToArray();
            //List<string> col_lst = col_name.ToList();
            //lstSTT.DataSource = col_lst.OrderBy(x => x).ToList();
            //List<DataTable> src_tbl_lst = new List<DataTable>();
            //myCode.Get_ListTable(-1, logfile_dt, new string[] { tar_Col, vitri_Set_val }, ref src_tbl_lst, vitri_act_val);
            //Dictionary<string, List<string>> dicFAI_data = new Dictionary<string, List<string>>();
            //DataTable result_tbl = new DataTable();
            //foreach (DataTable dt in src_tbl_lst)
            //{
            //    string fai_name = dt.Rows[0][tar_Col].ToString();
            //    string setval = dt.Rows[0][vitri_Set_val].ToString();
            //    string FAI = fai_name + "_" + setval;
            //    result_tbl.Columns.Add(FAI);
            //    List<string> data = dt.AsEnumerable().Select(x => x.Field<string>(vitri_act_val)).ToList();
            //    if (dicFAI_data.ContainsKey(FAI))
            //    {
            //        var temp_dic = dicFAI_data[FAI];
            //        temp_dic.AddRange(data);
            //    }
            //    else
            //    {
            //        dicFAI_data.Add(FAI,data);
            //    }
            //}
            //foreach(var fai in dicFAI_data)
            //{
            //    int row_add = fai.Value.Count - result_tbl.Rows.Count;
            //    if (row_add > 0)
            //    {
            //        for(int i=0;i<row_add;i++)
            //        {
            //            result_tbl.Rows.Add();
            //        }
            //    }
            //    int r_inx = 0;
            //    foreach(var item in fai.Value)
            //    {
            //        result_tbl.Rows[r_inx][fai.Key] = item;
            //        r_inx++;
            //    }
            //}
            //DGV_DataView.DataSource = result_tbl;

            //DataTable dt = new DataTable();
            //foreach(DataGridViewColumn dc in DGV_DataView.Columns)
            //{
            //    DGV_Data.Columns.Add(dc.Name, dc.Name);
            //    DGV_Data.Columns[dc.Name].DefaultCellStyle.Format = "N3";
            //}
            //int r_inx = 0;
            //foreach (DataGridViewRow dr in DGV_DataView.Rows)
            //{
            //    DGV_Data.Rows.Add();
            //    foreach (DataGridViewColumn dc in DGV_DataView.Columns)
            //    {
            //        DGV_Data.Rows[r_inx].Cells[dc.Name].Value = dr.Cells[dc.Name].Value;
            //    }
            //    r_inx++;
            //}

            //DataGridViewCellStyle style = new DataGridViewCellStyle();
            //style.Format = "N3";
            //foreach (DataGridViewColumn dc in DGV_DataView.Columns)
            //{

            //    dc.DefaultCellStyle = style;
            //}

            //DataTable dt = (DataTable)DGV_DataView.DataSource;
            //DGV_Data.DataSource = dt;

            //ExcelPackage Report_Pack = new ExcelPackage();
            //Report_Pack = FAI_lib.open_excel(@"E:\SEEV_GĐ2\SEEV_Modify\Install for SEEV\Install for SEEV\OK2SHIP automation system\0.OK2SHIP report format\Ok2ship Request SEEV D83 RCAM AVJ Flex C30(POR)_821-04230-04 528797.xlsm");

            //ExcelWorkbook report_saved = Report_Pack.Workbook;
            //List<string> selected_sht_name = new List<string>() { "FAI", "SPC" };
            //List<string> del_name_lst = new List<string>();
            //foreach(ExcelWorksheet cur_sht in report_saved.Worksheets)
            //{
            //    string sht_name = cur_sht.Name;
            //    if (!selected_sht_name.Any(x => sht_name.Contains(x)))
            //    {
            //        del_name_lst.Add(sht_name);
            //    }
            //}
            //foreach(string del_name in del_name_lst)
            //{
            //    report_saved.Worksheets.Delete(del_name);
            //}
            //Report_Pack.SaveAs(new FileInfo(@"E:\TestExcel.xlsm"));
            //Report_Pack.Dispose();
        }
        public DataTable VHX_FAI_Data_Process(string log_path)
        {
            DataTable result_dt = new DataTable();
            DirectoryInfo tar_d = new DirectoryInfo(log_path);
            DirectoryInfo[] inter_type_lst = new DirectoryInfo[] { tar_d };
            Dictionary<string, Dictionary<int, Dictionary<int, string>>> dic_result = new Dictionary<string, Dictionary<int, Dictionary<int, string>>>();
            foreach (var inter_type in inter_type_lst)
            {
                DirectoryInfo sel_dir = new DirectoryInfo(inter_type.FullName);
                IPQC_LogFile.Get_logfile_Multi(inter_type.FullName, ref dic_result);
            }
            foreach (string fai in dic_result.Keys)
            {
                if (result_dt.Columns.IndexOf(fai) == -1)
                {
                    result_dt.Columns.Add(fai);
                }
                List<string> data = dic_result[fai].Values.SelectMany(x => x.Values).ToList();
                if (result_dt.Rows.Count < data.Count)
                {
                    int row_added = data.Count - result_dt.Rows.Count;
                    for (int i = 0; i < row_added; i++)
                    {
                        result_dt.Rows.Add();
                    }
                }
                foreach (DataRow dr in result_dt.Rows)
                {
                    dr[fai] = data[result_dt.Rows.IndexOf(dr)];
                }
            }
            return result_dt;
        }
       
        private void btnLoadSpec_Click(object sender, EventArgs e)
        {
            string ItemCode = "422435";
            string LotNo =  "00003";
            Load_PTH_Diameter_Spec(ItemCode, sel_sqlcon, "SPEC_PTH_Diameter", ref DGV_SpecView);

        }
        public void initial_data()
        {
            app_path = Application.StartupPath;
            string config_file = Path.Combine(app_path, "Config", "config.txt");
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
                    string[] temp = c.Split(':');
                    string tg = "";
                    if (temp.Length > 2)
                    {

                        for (int t = 1; t < temp.Length; t++)
                        {
                            tg = tg + temp[t] + ":";
                        }
                    }
                    Data_Location = tg.TrimEnd(':').Trim();
                }
                if (c.Contains("Format_Folder"))
                {
                    format_folder = c.Split('#')[1].Trim();
                }
                if (c.Contains("Log_folder"))
                {
                    log_folder = c.Split('#')[1].Trim();
                }

            }
            string connstr_Declare = TDMK_Code.data_connection(server_name, "Declaration", server_acc, server_pass).ConnectionString;
            string connstr_Materials = TDMK_Code.data_connection(server_name, "Materials", server_acc, server_pass).ConnectionString;
            string connstr_Recycle = TDMK_Code.data_connection(server_name, "Recycled_PGC_Copper", server_acc, server_pass).ConnectionString;
            //string connstr_FAI = TDMK_Code.data_connection(server_name, "SEI_DB", server_acc, server_pass).ConnectionString;
            string connstr_FAI = TDMK_Code.data_connection(server_name, "SEI_FAI", server_acc, server_pass).ConnectionString;
            string connstr_IPQC = TDMK_Code.data_connection(server_name, "IPQC_Data", server_acc, server_pass).ConnectionString;
            string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, "OK2SHIP_Items", server_acc, server_pass).ConnectionString;
            sqlcon_Declare = new SqlConnection(connstr_Declare);
            sqlcon_Materials = new SqlConnection(connstr_Materials);
            sqlcon_Recycle = new SqlConnection(connstr_Recycle);
            sqlcon_FAI = new SqlConnection(connstr_FAI);
            sqlcon_IPQC = new SqlConnection(connstr_IPQC);
            sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);

        }

        private void lstSTT_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sel_col = lstSTT.SelectedItem.ToString();
            DataTable src_tbl = (DataTable)DGV_Data.DataSource;
            
        }

        private void tsmloadSpec_Click(object sender, EventArgs e)
        {
            string ItemCode = "422435";
            string LotNo = "00003";
            string fil_loc = txtFile.Text;
            string filter_str = TDMK_Code.filter_str(new string[] {"ItemCode"}, new string[] {ItemCode});
            
            if (PTH_Diameter_Spec_Setup(fil_loc,ItemCode,sel_sqlcon,"SPEC_PTH_Diameter"))
            {
                DGV_SpecView.DataSource = TDMK_Code.Datatable_Filter(sel_sqlcon, "SPEC_PTH_Diameter", filter_str);
            }
            
            ////myCode.Load_Spec(sqlcon_FAI, fil_loc, ItemCode, LotNo, ".xlsm", DGV_SpecView, "");
            
            //DataTable format_spec_dt = PTH_Diameter_Load_Spec_fromFile(sel_sqlcon, fil_loc, ItemCode, new List<string> { "*.xlsx", "*.xlsm" });
            //TDMK_Code.Delelte_FilteredItem_arr("SPEC_PTH_Diameter", sel_sqlcon, filter_str);
            //DataTable spec_dt = TDMK_Code.Datatable_Filter(sel_sqlcon, "SPEC_PTH_Diameter", filter_str);
            //int id = Convert.ToInt32(TDMK_Code.SQL_MAX("SPEC_PTH_Diameter","ID",sel_sqlcon))+1;
            //foreach(DataColumn dc in format_spec_dt.Columns)
            //{
            //    DataRow dr = spec_dt.NewRow();
            //    dr[0] = id++;
            //    dr[1] = ItemCode;
            //    dr[2] = TDMK_EPPLUS7.checkDBNull(format_spec_dt.Rows[0][dc]);
            //    dr[3] = TDMK_EPPLUS7.checkDBNull(format_spec_dt.Rows[1][dc]);
            //    dr[4] = TDMK_EPPLUS7.checkDBNull(format_spec_dt.Rows[2][dc]);
            //    spec_dt.Rows.Add(dr);
            //}
            //BatchBulkCopy(sel_sqlcon, spec_dt, "SPEC_PTH_Diameter");
            
        }
        public bool PTH_Diameter_Spec_Setup(string format_file, string _ItemCode, SqlConnection sqlcon, string DB_table_name)
        {
            bool result = false;
            DataTable format_spec_dt = PTH_Diameter_Load_Spec_fromFile(sqlcon, format_file, _ItemCode, new List<string> { "*.xlsx", "*.xlsm" });
            if(format_spec_dt.Rows.Count!=0)
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { _ItemCode });
                TDMK_Code.Delelte_FilteredItem_arr(DB_table_name, sqlcon, filter_str);
                DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon, DB_table_name, filter_str);
                int id = Convert.ToInt32(TDMK_Code.SQL_MAX(DB_table_name, "ID", sqlcon)) + 1;
                foreach (DataColumn dc in format_spec_dt.Columns)
                {
                    DataRow dr = spec_dt.NewRow();
                    dr[0] = id++;
                    dr[1] = _ItemCode;
                    dr[2] = TDMK_EPPLUS7.checkDBNull(format_spec_dt.Rows[0][dc]);
                    dr[3] = TDMK_EPPLUS7.checkDBNull(format_spec_dt.Rows[1][dc]);
                    dr[4] = TDMK_EPPLUS7.checkDBNull(format_spec_dt.Rows[2][dc]);
                    spec_dt.Rows.Add(dr);
                }
                BatchBulkCopy(sqlcon, spec_dt, DB_table_name);
                result = true;
            }
            return result;
        }
        public DataTable Load_PTH_Diameter_Spec(string _ItemCode,SqlConnection sqlcon, string DB_table_name)
        {
            DataTable result = new DataTable();
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { _ItemCode });
            DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon, DB_table_name, filter_str);
            foreach(DataRow dr in spec_dt.Rows)
            {
                string setval = TDMK_EPPLUS7.checkDBNull(dr["Normdim"]);
                string USL = TDMK_EPPLUS7.checkDBNull(dr["USL"]);
                string LSL = TDMK_EPPLUS7.checkDBNull(dr["LSL"]);
                string col_name = "PTH-Diameter_" + setval;
                if(!result.Columns.Contains(col_name))
                {
                    result.Columns.Add(col_name);
                    if(result.Rows.Count==0)
                    {
                        for(int i=0;i<3;i++)
                        {
                            result.Rows.Add();
                        }
                    }
                    result.Rows[0][col_name] = setval;
                    result.Rows[1][col_name] = USL;
                    result.Rows[2][col_name] = LSL;
                }
            }
            return result;
        }
        public bool Load_PTH_Diameter_Spec(string _ItemCode, SqlConnection sqlcon, string DB_table_name, ref DataGridView tar_DGV)
        {
            DataTable result = Load_PTH_Diameter_Spec(_ItemCode,sqlcon,DB_table_name);
            if (result.Rows.Count == 0)
            {
                return false;
            }
            else
            {
                tar_DGV.DataSource = result;
                List<string> header_list = new List<string>() { "Normdim", "USL", "LSL" };
                for (int i = 0; i < 3; i++)
                {
                    tar_DGV.Rows[i].HeaderCell.Value = header_list[i];
                }
                tar_DGV.AutoResizeColumns();
                tar_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                return true;
            }
        }

        private void tsmTestFunction_Click(object sender, EventArgs e)
        {
            DataTable spec_dt = (DataTable)DGV_SpecView.DataSource;
            DataTable FAI_dt = (DataTable)DGV_DataView.DataSource;
            DataTable main_FAI = new DataTable();
            DataTable other_FAI = new DataTable();
            filter_table(FAI_dt, spec_dt, ref main_FAI, ref other_FAI);
            DGV_Data_Plus.DataSource = main_FAI;
            DGV_CPK.DataSource = other_FAI;
        }


        private void tsmTest2_Click(object sender, EventArgs e)
        {
            string format_file = @"D:\Customer Projects\SEEV\OK2SHIP_GD3\Format\Ok2ship Request Sumitomo VN D84 R-Cam T Pismo Flex RR Build 821-04428-A 422435.xlsm";// format_lst[0];
            string export_file = "";
            string type = "MASS";
            string _itemcode = "422435";
            string _lotno = "00003";
            string report_location = Application.StartupPath;
            string report_path = Path.Combine(report_location, type);
            int qty = 32;
            if (!Directory.Exists(report_path))
            {
                Directory.CreateDirectory(report_path);
            }
            ExcelPackage Format_Pack = new ExcelPackage(format_file);
            ExcelPackage Report_Pack = new ExcelPackage();
            if (type.ToUpper() == "MASS")
            {
                qty = 6;
                export_file = Path.Combine(report_path, Path.GetFileNameWithoutExtension(format_file) + "-" + _lotno + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");
                //List<string> selected_sht_name = new List<string>() { "FAI", "SPC", "BVH", "PTH" };
                List<string> selected_sht_name = new List<string>() { "BVH", "PTH" };
                bool same_ext = false;
                if (Path.GetExtension(format_file) == Path.GetExtension(export_file))
                {
                    same_ext = true;
                }
                Report_Pack = myCode.Worksheet_select(Format_Pack, selected_sht_name, same_ext);
            }
            else
            {
                export_file = Path.Combine(report_path, _itemcode + "-" + _lotno + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(format_file));
                Report_Pack = Format_Pack;
            }
            DataTable PTH_data_tbl = TDMK_Code.Datatable_Filter(sel_sqlcon, "PTH_Diameter", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo","Data_Type" }, new string[] { _itemcode, _lotno, type }));
            int qty_export = Math.Min(qty, PTH_data_tbl.Rows.Count);
            string start_data_addr = TDMK_EPPLUS7.Find_Cell_Addr("Sample", "A1", Report_Pack.Workbook.Worksheets[0], false);
            string PTH_Zone = TDMK_EPPLUS7.Find_Cell_Addr("Diameter seen from top", start_data_addr, Report_Pack.Workbook.Worksheets[0], true);
            string PTH_Data_Addr = TDMK_EPPLUS7.Find_Start_Addr(PTH_Zone, Report_Pack.Workbook.Worksheets[0], false);
            ExcelRangeBase PTH_Data_rgn = Report_Pack.Workbook.Worksheets[0].Cells[PTH_Data_Addr];
            for (int i=0;i< qty_export;i++)
            {
                PTH_Data_rgn.Offset(i, 0).Value = PTH_data_tbl.Rows[i]["Data"];
            }
            Report_Pack.SaveAs(new FileInfo(export_file));
            Report_Pack.Dispose();
            Format_Pack.Dispose();
        }
        public ExcelPackage Worksheet_select(ExcelPackage src_pack,List<string> selected_sht_name, bool same_file_ext=true)
        {
            ExcelPackage result = new ExcelPackage();
            List<string> del_name_lst = new List<string>();
            foreach (ExcelWorksheet cur_sht in src_pack.Workbook.Worksheets)
            {
                string sht_name = TDMK_EPPLUS7.remove_special_chars(cur_sht.Name, new List<char> { '-', '&', '_', ' ', '\r', '\n' }).ToUpper();
                if (!selected_sht_name.Any(x => sht_name.Contains(x.ToUpper())))
                {
                    del_name_lst.Add(cur_sht.Name);
                }
            }
            foreach (string del_name in del_name_lst)
            {
                src_pack.Workbook.Worksheets.Delete(del_name);
            }
            if(same_file_ext)
            {
                return src_pack;
            }
            else
            {
                foreach (ExcelWorksheet cur_sht in src_pack.Workbook.Worksheets)
                {
                    result.Workbook.Worksheets.Add(cur_sht.Name, cur_sht);
                }
                return result;
            }
        }

        private void RBNormal_CheckedChanged(object sender, EventArgs e)
        {
            if(RBNormal.Checked)
            {
                editmode = false;
            }
            else
            {
                editmode = true;
            }
        }

        private void DGV_DataView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //if (e.ColumnIndex !=-1  && e.RowIndex != -1)
            //{
            //    double d = double.Parse(e.Value.ToString());
            //    e.Value = d.ToString("N3");
            //}
        }
        public void Export_FAI_Batch(DataTable FAI_Data_tbl, string tar_ItemCode, string tar_LotNo, myExcel.Workbook tar_wrkbook)
        {
            char[] trim_char = new char[] { ' ', '\r', '\n' };
            //DataTable FAI_Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { tar_ItemCode, tar_LotNo }));
            DataTable FAI_Spec_tbl = TDMK_Code.Datatable_Filter(sqlcon_FAI, "FAI_Spec", "ItemCode = '" + tar_ItemCode + "'");
            string[] sheetno = FAI_Spec_tbl.AsEnumerable().Select(r => r.Field<string>("SheetNo")).Distinct().ToArray();
            foreach (string sht in sheetno)
            {
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
                myExcel._Worksheet cur_wrksht = tar_wrkbook.Worksheets.Add(After: tar_wrkbook.Sheets[1]);
                cur_wrksht.Name = sht;
                myExcel.Range cur_rgn = cur_wrksht.Range["A1"];
                int c_inx = 0;
                foreach (string fai_no in FAI_No)
                {
                    cur_rgn.Offset[0, c_inx].Value = fai_no;
                    string[] fai_val = FAI_Data_tbl.AsEnumerable().Select(r => r.Field<string>(fai_no)).ToArray();
                    if (fai_val.Length > 0)
                    {
                        int r_inx = 0;
                        foreach (string _fai_val in fai_val)
                        {
                            if(_fai_val!=null)
                            {
                                cur_rgn.Offset[r_inx + 1, c_inx].Value = _fai_val.Trim(trim_char);
                            }
                            r_inx++;
                        }
                    }
                    else
                    {
                        for (int i = 0; i < 32; i++)
                        {
                            cur_rgn.Offset[i + 1, c_inx].Value = "N/A";
                        }
                    }
                    c_inx++;
                }
            }
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
        public DataTable PTH_Diameter_Load_Spec_fromFile(SqlConnection sqlcon, string format_loc, string tar_ItemCode, List<string> extensions)
        {
            DataTable spec_dt = new DataTable();
            string tar_format_file;
            List<string> file_format_lst = new List<string>();
            DirectoryInfo directory = new DirectoryInfo(format_loc);
            var files = extensions.SelectMany(e => directory.EnumerateFiles(e, SearchOption.AllDirectories)).Where(x => x.FullName.Contains(tar_ItemCode));
            foreach (var item in files)
            {
                file_format_lst.Add(item.FullName);
            }
            if (file_format_lst.Count > 0)
            {
                tar_format_file = file_format_lst[0];
                FileInfo excel_file = new FileInfo(tar_format_file);
                using (ExcelPackage myexcel = new ExcelPackage(excel_file))
                {
                    ExcelWorkbook tar_wkbook = myexcel.Workbook;
                    List<char> rejected_char_lst = new List<char> { ' ','-','&','_','\r','\n' };
                    List<string> sht_keys = new List<string> { "BVH-PTH" };
                    int wrksheet_num = tar_wkbook.Worksheets.Count;
                    foreach (ExcelWorksheet tg in tar_wkbook.Worksheets)
                    {
                        string sht_name = TDMK_EPPLUS7.remove_special_chars(tg.Name, rejected_char_lst).ToUpper();
                        if (sht_keys.Any(x => sht_name.Contains(TDMK_EPPLUS7.remove_special_chars(x, rejected_char_lst))))
                        {
                            string spec_location_address = TDMK_EPPLUS7.Find_Cell_Location("Diameter seen from top", tg, true);
                            if (spec_location_address != "")
                            {
                                var cell_info = TDMK_EPPLUS7.Get_Cells_Info(tg, tg.Cells[spec_location_address]);
                                int row_offset = cell_info.row_qty;
                                int col_offset = cell_info.col_qty;
                                string search_item_address = tg.Cells[spec_location_address].Offset(row_offset, -col_offset).Address;
                                string normdim_addr = TDMK_EPPLUS7.Find_Cell_Addr("Nominal Dim.", search_item_address, tg, false);
                                string tol_max_addr = TDMK_EPPLUS7.Find_Cell_Addr("Tol. Max. (+)", search_item_address, tg, false);
                                string tol_min_addr = TDMK_EPPLUS7.Find_Cell_Addr("Tol. Min. (-)", search_item_address, tg, false);
                                string instrument_addr = TDMK_EPPLUS7.Find_Cell_Addr("instrument", search_item_address, tg, false);
                                string USL_addr = TDMK_EPPLUS7.Find_Cell_Addr("USL", search_item_address, tg, false);
                                string LSL_addr = TDMK_EPPLUS7.Find_Cell_Addr("LSL", search_item_address, tg, false);
                                try
                                {
                                    ExcelRangeBase sel_rgn = tg.Cells[normdim_addr].Offset(0, col_offset);// tg.Range["D19"];    //tg.Range["C17"];
                                    ExcelRangeBase dev_rgn = tg.Cells[instrument_addr].Offset(0, col_offset);// tg.Range["D23"];    //tg.Range["C21"]
                                    int sel_inx = 0;
                                    while (myCode.checkDBNull(sel_rgn.Offset(0, sel_inx).Value) != "")
                                    {
                                        double USL = 0;
                                        double LSL = 0;
                                        string t_FAIName = "PTH-Dimension";
                                        string t_FAI_Setval = myCode.checkDBNull(sel_rgn.Offset(0, sel_inx).Value);
                                        string t_FAI_UL = myCode.checkDBNull(tg.Cells[USL_addr].Offset(0, col_offset + sel_inx).Value); ;// myCode.checkDBNull(sel_rgn.Offset[2, sel_inx].Value);
                                        string t_FAI_LL = myCode.checkDBNull(tg.Cells[LSL_addr].Offset(0, col_offset + sel_inx).Value); ;// myCode.checkDBNull(sel_rgn.Offset[3, sel_inx].Value);
                                        string t_FAI_Tol_plus = myCode.checkDBNull(tg.Cells[tol_max_addr].Offset(0, col_offset + sel_inx).Value);
                                        string t_FAI_Tol_minus = myCode.checkDBNull(tg.Cells[tol_min_addr].Offset(0, col_offset + sel_inx).Value);
                                        if ((t_FAI_Setval != "NA") && (t_FAI_Setval != ""))
                                        {
                                            double sv = Convert.ToDouble(t_FAI_Setval);
                                            if (myCode.IsNumeric(t_FAI_Tol_plus))
                                            {
                                                USL = sv + Convert.ToDouble(t_FAI_Tol_plus);
                                                t_FAI_UL = USL.ToString();
                                            }
                                            else
                                            {
                                                t_FAI_UL = t_FAI_Tol_plus;
                                            }
                                            if (myCode.IsNumeric(t_FAI_Tol_minus))
                                            {
                                                LSL = sv - Convert.ToDouble(t_FAI_Tol_minus);
                                                t_FAI_LL = LSL.ToString();
                                            }
                                            else
                                            {
                                                t_FAI_LL = t_FAI_Tol_minus;
                                            }
                                        }
                                        else
                                        {
                                            t_FAI_LL = t_FAI_Tol_minus;
                                            t_FAI_UL = t_FAI_Tol_plus;
                                        }
                                        string col_name = t_FAIName + "_" + t_FAI_Setval;
                                        string t_instrument = myCode.checkDBNull(dev_rgn.Offset(0, sel_inx).Value);
                                        if (!myCode.check_columns_existed(spec_dt, col_name))
                                        {
                                            spec_dt.Columns.Add(col_name);
                                            if (spec_dt.Rows.Count == 0)
                                            {
                                                for (int i = 0; i < 3; i++)
                                                {
                                                    spec_dt.Rows.Add();
                                                }
                                            }
                                            spec_dt.Rows[0][col_name] = t_FAI_Setval;
                                            spec_dt.Rows[1][col_name] = t_FAI_UL;
                                            spec_dt.Rows[2][col_name] = t_FAI_LL;
                                        }
                                        sel_inx++;
                                    }
                                }
                                catch
                                {

                                }

                            }    
                        }
                    }
                }
            }
            else
            {
                //MessageBox.Show(new Form { TopMost = true }, "File format not found");
            }
            return spec_dt;
        }
        
        public DataTable Load_Spec_fromFile(SqlConnection sqlcon, ExcelWorkbook wb, string tar_ItemCode, List<string> extensions, string format_type = "NPI")
        {
            DataTable spec_dt = new DataTable();
            AutoCompleteStringCollection FAI_SheetNo_list = new AutoCompleteStringCollection();
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { tar_ItemCode, format_type });
            FAI_SheetNo_list = TDMK_Code.Load_Item_Names_Filter(sqlcon, "FAI_Spec", "SheetNo", filter_str);
            List<string> sht_keys = new List<string> { "FAI", "SPC", "parentheses" };
            int wrksheet_num = wb.Worksheets.Count;
            foreach (ExcelWorksheet tg in wb.Worksheets)
            {
                if (sht_keys.Any(x => tg.Name.Contains(x)))
                {
                    string dim_no_addr = TDMK_EPPLUS7.Find_Cell_Addr("Dim. No.", "B10", tg, false);
                    string instrument_addr = TDMK_EPPLUS7.  Find_Cell_Addr("instrument", "B10", tg, false);
                    string FAI_data_addr = TDMK_EPPLUS7.FindByOffset(tg.Cells[dim_no_addr].Offset(0, 1).Address, tg, true, "");
                    int off_set = tg.Cells[FAI_data_addr].End.Column - tg.Cells[instrument_addr].End.Column;
                    ExcelRangeBase sel_rgn = tg.Cells[dim_no_addr].Offset(0, off_set);// tg.Range["D19"];    //tg.Range["C17"];
                    ExcelRangeBase dev_rgn = tg.Cells[instrument_addr].Offset(0, off_set);// tg.Range["D23"];    //tg.Range["C21"]
                    int sel_inx = 0;
                    while (myCode.checkDBNull(sel_rgn.Offset(0, sel_inx).Value) != "")
                    {
                        double USL = 0;
                        double LSL = 0;
                        string t_checkside = myCode.checkDBNull(sel_rgn.Offset(-1, sel_inx).Value);
                        string t_FAIName = myCode.checkDBNull(sel_rgn.Offset(0, sel_inx).Value).Replace(" ", "");
                        string t_FAI_Setval = myCode.checkDBNull(sel_rgn.Offset(1, sel_inx).Value);
                        string t_FAI_UL = myCode.checkDBNull(sel_rgn.Offset(6, sel_inx).Value); ;// myCode.checkDBNull(sel_rgn.Offset[2, sel_inx].Value);
                        string t_FAI_LL = myCode.checkDBNull(sel_rgn.Offset(7, sel_inx).Value); ;// myCode.checkDBNull(sel_rgn.Offset[3, sel_inx].Value);
                        string t_FAI_Tol_plus = myCode.checkDBNull(sel_rgn.Offset(2, sel_inx).Value);
                        string t_FAI_Tol_minus = myCode.checkDBNull(sel_rgn.Offset(3, sel_inx).Value);
                        if ((t_FAI_Setval != "NA") && (t_FAI_Setval != ""))
                        {
                            double sv = Convert.ToDouble(t_FAI_Setval);
                            if (myCode.IsNumeric(t_FAI_Tol_plus))
                            {
                                USL = sv + Convert.ToDouble(t_FAI_Tol_plus);
                                t_FAI_UL = USL.ToString();
                            }
                            else
                            {
                                t_FAI_UL = t_FAI_Tol_plus;
                            }
                            if (myCode.IsNumeric(t_FAI_Tol_minus))
                            {
                                LSL = sv - Convert.ToDouble(t_FAI_Tol_minus);
                                t_FAI_LL = LSL.ToString();
                            }
                            else
                            {
                                t_FAI_LL = t_FAI_Tol_minus;
                            }
                        }
                        else
                        {
                            if (t_checkside == "SingleSide-USL")
                            {
                                //USL = Convert.ToDouble(t_FAI_Tol_plus);
                                t_FAI_UL = t_FAI_Tol_plus;// USL.ToString();
                            }
                            if (t_checkside == "SingleSide-LSL")
                            {
                                //LSL = Convert.ToDouble(t_FAI_Tol_minus);
                                t_FAI_LL = t_FAI_Tol_minus;// LSL.ToString();
                            }
                        }
                        string col_name = t_FAIName + "_" + t_FAI_Setval;
                        string t_FAI_sheetno = tg.Name;
                        string t_instrument = myCode.checkDBNull(dev_rgn.Offset(0, sel_inx).Value);
                        if (!myCode.check_columns_existed(spec_dt, col_name))
                        {
                            spec_dt.Columns.Add(col_name);
                            if (spec_dt.Rows.Count == 0)
                            {
                                for (int i = 0; i < 6; i++)
                                {
                                    spec_dt.Rows.Add();
                                }
                            }
                            spec_dt.Rows[0][col_name] = t_FAI_Setval; // Set val at Row =0
                            spec_dt.Rows[1][col_name] = t_FAI_UL; // UL at Row =1
                            spec_dt.Rows[2][col_name] = t_FAI_LL; // LL at Row =2
                            spec_dt.Rows[3][col_name] = t_checkside; // Check Side at Row =3
                            spec_dt.Rows[4][col_name] = t_instrument; // Instrument at Row =4
                            spec_dt.Rows[5][col_name] = tg.Name; // Instrument at Row =5
                        }
                        sel_inx++;
                    }
                }
            }
            return spec_dt;
        }
        public SqlConnection initial_data(string DB_name, bool sa_en)
        {
            SqlConnection _sqlcon_OK2SHIP;
            if (!File.Exists("Config.ini"))
            {
                TDMK_init.Write("Server", "10.212.1.243", "OK2SHIP_Config");
                TDMK_init.Write("Account", "sa", "OK2SHIP_Config");
                TDMK_init.Write("Password", "seev@123;", "OK2SHIP_Config");
                TDMK_init.Write("Report_Location", @"\\10.212.1.200\rgo\OK2SHIP automation system", "OK2SHIP_Config");
                TDMK_init.Write("Data_Location", @"\\10.212.1.200\rgo\OK2SHIP automation system\TDMK_Data\TestAreas\OK2SHIP_Measurements\OMM", "OK2SHIP_Config");
                TDMK_init.Write("Format_Folder", @"\\10.212.1.200\rgo\OK2SHIP automation system\0.OK2SHIP report format", "OK2SHIP_Config");
                TDMK_init.Write("Log_folder", @"\\10.212.1.200\rgo\OK2SHIP automation system\TDMK_Data\TestAreas\OK2SHIP_Measurements", "OK2SHIP_Config");
            }
            server_name = TDMK_init.Read("Server", "OK2SHIP_Config");
            server_acc = TDMK_init.Read("Account", "OK2SHIP_Config");
            server_pass = TDMK_init.Read("Password", "OK2SHIP_Config");
            Data_Location = TDMK_init.Read("Data_Location", "OK2SHIP_Config");
            format_folder = TDMK_init.Read("Format_Folder", "OK2SHIP_Config");
            log_folder = TDMK_init.Read("Log_folder", "OK2SHIP_Config");
            string report_location = TDMK_init.Read("Report_Location", "OK2SHIP_Config");
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
            return _sqlcon_OK2SHIP;
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
        public void Export_PTH_Diameter_byEPPLUS( SqlConnection sqlcon, string ItemCode, string LotNo, string type="NPI")
        {
            string format_loc = format_folder;
            string _itemcode = ItemCode;
            string _lotno = LotNo;
            string report_location = "";
            DataTable PTH_Diameter_dt = myCode.Load_PTH_Diameter_ToTable(sqlcon,"PTH_Diameter",ItemCode,LotNo) ;// Load_FAI_ToTable(myVar.sqlcon_SMT, "FAI_Auto", "FAI_No", "FAI_Data", _itemcode, _lotno, type); //myCode.Load_FAI_ToTable(myVar.sqlcon_SMT, _itemcode, _lotno, cbType_Sel.SelectedItem.ToString());
            DataTable PTH_Diameter_Spec = myCode.Load_PTH_Diameter_Spec(ItemCode, sqlcon, "SPEC_PTH_Diameter");
            bool export_en = false;
            if (myCode.check_FAIdata_inSpec(PTH_Diameter_dt, PTH_Diameter_Spec))
            {
                export_en = true;
            }
            else
            {
                if (MessageBox.Show("Dữ liệu NG. Tiếp tục xuất dữ liệu?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    export_en = true;
                }
                else
                {
                    export_en = false;
                }

            }
            if (export_en)
            {
                List<string> format_lst = get_multiple_format(format_loc, new List<string> { "*.xlsx", "*.xlsm" }, _itemcode);
                if (format_lst.Count > 0)
                {
                    if (PTH_Diameter_dt.Rows.Count > 0)
                    {
                        string format_file = format_lst[0];
                        string export_file = "";
                        string report_path = Path.Combine(report_location, type);
                        ExcelPackage Format_Pack = new ExcelPackage(format_file);
                        ExcelPackage Report_Pack = null;
                        ExcelWorkbook report_saved = null;
                        if (!Directory.Exists(report_path))
                        {
                            Directory.CreateDirectory(report_path);
                        }
                        if (type.ToUpper() == "MASS")
                        {
                            export_file = Path.Combine(report_path, Path.GetFileNameWithoutExtension(format_file) + "-" + _lotno + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");
                            List<string> selected_sht_name = new List<string>() { "BVH", "PTH" };
                            bool same_ext = false;
                            if (Path.GetExtension(format_file) == Path.GetExtension(export_file))
                            {
                                same_ext = true;
                            }
                            Report_Pack = myCode.Worksheet_select(Format_Pack, selected_sht_name, same_ext);
                        }
                        else
                        {
                            export_file = Path.Combine(report_path, _itemcode + "-" + _lotno + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(format_file));
                            Report_Pack = Format_Pack;
                        }
                        report_saved = Report_Pack.Workbook;

                        //Dictionary<string, DataTable> dic_data = FAI_lib.Export_FAI_Batch(sel_sqlcon, report_saved, _itemcode, _lotno, type, "");
                        //FAI_lib.Export_To_FAI(sel_sqlcon, Report_Pack, _itemcode, _lotno, dic_data, type);
                        Report_Pack.SaveAs(export_file);
                        Report_Pack.Dispose();
                        MessageBox.Show(new Form { TopMost = true }, "Hoàn thành xuất dữ liệu", "Thông báo");
                        ProcessStartInfo pi = new ProcessStartInfo(export_file);
                        Process.Start(pi);
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Không tìm thấy dữ liệu của ItemCode / Lotno / Shift : " + _itemcode + " / " + _lotno , "Thông báo");
                    }
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Không tìm thấy Format", "Cảnh báo");
                }
            }

        }
        public List<string> get_multiple_format(string src_path, List<string> extensions, string tar_ItemCode)
        {
            List<string> result = new List<string>();
            List<char> remove_char = new List<char> { ' ', '-', '_' };
            DirectoryInfo directory = new DirectoryInfo(src_path);
            if (directory.Exists)
            {
                var files = extensions.SelectMany(e => directory.EnumerateFiles(e, SearchOption.TopDirectoryOnly)).Where(x => x.FullName.Contains(tar_ItemCode));
                foreach (var item in files)
                {
                    result.Add(item.FullName);
                }
            }
            else
            {
                MessageBox.Show("Không tìm thấy thư mục : " + src_path, "Thông báo");
            }
            return result;
        }
    }
}
