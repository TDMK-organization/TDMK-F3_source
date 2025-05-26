using IniLibs;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SQL;
using ZedGraph;
using myExcel = Microsoft.Office.Interop.Excel;
//using OK2SHIP;
using System.Threading;
using OK2SHIP_Software;
using OK2SHIP_Lib;

namespace OK2SHIP_Measurements
{
    public partial class TestFunction : Form
    {
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        TDMK_OK2SHIP myCode = new TDMK_OK2SHIP();
        myVar common_lib = new myVar();
        //public static SqlConnection sqlcon_FAI;
        //public static SqlConnection sqlcon_Declare;
        //public static SqlConnection sqlcon_Materials;
        //public static SqlConnection sqlcon_Recycle;
        //public static SqlConnection sqlcon_IPQC;
        public static SqlConnection sqlcon_OK2SHIP;
        public static SqlConnection sel_sqlcon;
        public static SqlConnection sqlcon_SMT;
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
        public IniFile TDMK_init = new IniFile("Config.ini");
        public bool draw_en = false;
        public FileSystemWatcher watchfolder = new FileSystemWatcher();
        public TestFunction()
        {
            InitializeComponent();
        }

        private void TestFunction_Load(object sender, EventArgs e)
        {
            sqlcon_SMT= initial_data("OK2SHIP_SMT",true);
            myVar.sqlcon_SMT = sqlcon_SMT;
            //txtFile.Text = Data_Location;
            string watch_path = Data_Location;
            string[] filter_ext = new string[] { ".csv", ".xlsx" };
            watchfolder.Path = watch_path;
            watchfolder.NotifyFilter = NotifyFilters.DirectoryName | NotifyFilters.FileName | NotifyFilters.Attributes;
            watchfolder.Created += Watchfolder_Created;
            watchfolder.EnableRaisingEvents = true;

        }
        private void Watchfolder_Created(object sender, FileSystemEventArgs e)
        {
            if (e.ChangeType == WatcherChangeTypes.Created)
            {
                Thread.Sleep(1000);
                FileInfo f_info = new FileInfo(e.FullPath);
                MessageBox.Show(e.FullPath);
                //try
                //{
                //    string itemcode_folder;
                //    string f_info = e.FullPath;
                //    string done_file;
                //    string f_name = Path.GetFileName(f_info);
                //    //string save_time = Create_date_string(DateTime.Now);
                //    //if (!e.FullPath.Contains("$"))
                //    //{
                //    //    TDMK_OK2SHIP.scan_file = e.FullPath;
                //    //    if (pro_en)
                //    //    {
                //    //        Thread tg_thread = new Thread(Scan_data);
                //    //        tg_thread.Start();
                //    //    }
                //    //    else
                //    //    {
                //    //        itemcode_folder = Path.Combine(app_path, "Ignored_Scan", txtItemCode.Text);
                //    //        done_file = Path.Combine(itemcode_folder, "Ignored_" + save_time + f_name);
                //    //        if (!Directory.Exists(itemcode_folder))
                //    //        {
                //    //            Directory.CreateDirectory(itemcode_folder);
                //    //        }
                //    //        File.Copy(f_info, done_file, true);
                //    //        File.Delete(f_info);
                //    //        MessageBox.Show(new Form { TopMost = true }, "Bạn chưa ấn nút Start", "Thông tin");
                //    //    }
                //    //}
                //}
                //catch (Exception ex)
                //{
                //    //Error_log(app_path, ex.Message);
                //}
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog f_file = new OpenFileDialog();
            if(f_file.ShowDialog() == DialogResult.OK)
            {
                //myCode.Mitutoyo_process(f_file.FileName, DGV_DataView, DGV_SpecView, DGV_Data_Plus, editmode);
                string f_name = f_file.FileName;



                //DataTable logfile_dt = new DataTable();
                //logfile_dt = myCode.Mitutoyo_Get_raw_data2(f_name);
                //DGV_Data.DataSource = logfile_dt;
                //DataTable FAI_logfile_tbl = FAI_raw_tbl(logfile_dt);
                //DGV_DataView.DataSource = FAI_logfile_tbl;

                //DataTable tbl_spec = (DataTable)DGV_SpecView.DataSource;
                //DataTable FAI_tbl = new DataTable();
                //DataTable non_FAI = new DataTable();
                //DataTable FAI_arranged_tbl = filter_table(FAI_logfile_tbl, tbl_spec, ref FAI_tbl, ref non_FAI);
                //DGV_Data_Plus.DataSource = FAI_tbl;
                //DGV_CPK.DataSource = non_FAI;

                //if (!editmode)
                //{
                //    DataTable dest_FAI_tbl = result_add((DataTable)DGV_DataView.DataSource, FAI_tbl, true);
                //    DataTable dest_nonFAI_tbl = result_add((DataTable)DGV_Data_Plus.DataSource, non_FAI, true);
                //    DGV_DataView.DataSource = null;
                //    DGV_Data_Plus.DataSource = null;
                //    DGV_DataView.DataSource = Order_table_inSpec(dest_FAI_tbl, tbl_spec);
                //    DGV_Data_Plus.DataSource = dest_nonFAI_tbl;
                //}
                //else
                //{
                //    DGV_Data_Plus.DataSource = null;
                //    DGV_Data_Plus.DataSource = Order_table_inSpec(FAI_arranged_tbl, tbl_spec);

                //}
                //check_FAIdata_inSpec(DGV_DataView, DGV_SpecView);
                //check_FAIdata_inSpec(DGV_Data_Plus, DGV_SpecView);

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
            foreach (DataRow dr in logfile_dt.Rows)
            {
                try
                {
                    dr[vitri_Set_val] = Convert.ToDouble(dr[vitri_Set_val]);
                }
                catch
                {
                    //dr[vitri_Set_val] = "Error";
                }

            }
            myCode.Get_ListTable(-1, logfile_dt, new string[] { tar_Col }, ref src_tbl_lst, vitri_act_val);
            Dictionary<string, List<string>> dicFAI_data = new Dictionary<string, List<string>>();
            foreach (DataTable dt in src_tbl_lst)
            {
                List<string> setval_lst = dt.AsEnumerable().Where(x => x.Field<string>(vitri_Set_val)!="").Select(x => x.Field<string>(vitri_Set_val)).Distinct().ToList();
                string fai_name = dt.Rows[0][tar_Col].ToString();
                
                string setval = setval_lst[0];//dt.Rows[0][vitri_Set_val].ToString();                
                string FAI = fai_name + "_" + setval;
                string vitri_val = dt.Rows[0][vitri_col_name].ToString();
                //if(fai_name=="FAI60")
                //{
                //    MessageBox.Show("here");
                //}
                result_tbl.Columns.Add(FAI);
                List<string> data = dt.AsEnumerable().Select(x => x.Field<string>(vitri_act_val)).ToList();
                if (!vitri_val.ToUpper().Contains("ANGLE"))
                {
                    for (int i = 0; i < data.Count; i++)
                    {
                        if(myCode.IsNumeric(data[i]))
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
            for( int i=0;i< tar_col_lst.Count;i++)
            {
                string fai_col = tar_col_lst[i];
                List<string> sel_fai_col =  FAI_dt.AsEnumerable().Select(x => x.Field<string>(fai_col)).ToList();
                if(sel_fai_col.Any(x=>!myCode.IsNumeric(x)))
                {
                    tar_col_lst.Remove(fai_col);
                    if(ex_col.IndexOf(fai_col)==-1)
                    {
                        ex_col.Add(fai_col);
                    }
                }
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
           

        }
        
        private void btnLoadSpec_Click(object sender, EventArgs e)
        {
            string ItemCode = txtItemCode.Text;// "7S0030";// "422292";
            string LotNo =  txtLotNo.Text;// "00005-01";
            //string fil_loc = txtFile.Text;
            //myCode.Load_Spec(sqlcon_FAI, fil_loc, ItemCode, LotNo, ".xlsm", DGV_SpecView, "");
           // myCode.Load_FAI_DGV2(DGV_DataView, sqlcon_SMT, "FAI_AUTO", "FAI_No", "FAI_Data", ItemCode, LotNo);
        }
        public void initial_data()
        {
            app_path = Application.StartupPath;
            if (!File.Exists("Config.ini"))
            {
                TDMK_init.Write("Server", "10.212.1.243", "SMT_Config");
                TDMK_init.Write("Account", "sa", "SMT_Config");
                TDMK_init.Write("Password", "seev@123;", "SMT_Config");
                TDMK_init.Write("Report_Location", @"\\10.212.1.200\rgo\OK2SHIP automation system", "SMT_Config");
                TDMK_init.Write("Data_Location", @"\\10.212.1.200\rgo\OK2SHIP automation system\TDMK_Data\TestAreas\OK2SHIP_Measurements\OMM", "SMT_Config");
                TDMK_init.Write("Format_Folder", @"\\10.212.1.200\rgo\OK2SHIP automation system\0.OK2SHIP report format", "SMT_Config");
                TDMK_init.Write("Log_folder", @"\\10.212.1.200\rgo\OK2SHIP automation system\TDMK_Data\TestAreas\OK2SHIP_Measurements", "SMT_Config");
            }
            server_name = TDMK_init.Read("Server", "SMT_Config");
            server_acc = TDMK_init.Read("Account", "SMT_Config");
            server_pass = TDMK_init.Read("Password", "SMT_Config");
            Data_Location = TDMK_init.Read("Data_Location", "SMT_Config");
            format_folder = TDMK_init.Read("Format_Folder", "SMT_Config");
            log_folder = TDMK_init.Read("Log_folder", "SMT_Config");

            string connstr_SMT = TDMK_Code.data_connection(server_name, "OK2SHIP_SMT", server_acc, server_pass).ConnectionString;
            //string connstr_IPQC = TDMK_Code.data_connection(server_name, "IPQC_Data", server_acc, server_pass).ConnectionString;
            string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, "OK2SHIP_Items", server_acc, server_pass).ConnectionString;
            sqlcon_SMT = new SqlConnection(connstr_SMT);
            //sqlcon_IPQC = new SqlConnection(connstr_IPQC);
            sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);

        }

        private void lstSTT_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sel_col = lstSTT.SelectedItem.ToString();
            DataTable src_tbl = (DataTable)DGV_Data.DataSource;
            
        }

        private void tsmloadSpec_Click(object sender, EventArgs e)
        {
            string ItemCode = txtItemCode.Text;// "422292";
            string LotNo = txtLotNo.Text;// "00003";
            string fil_loc = txtFile.Text;
            //myCode.Load_Spec(sqlcon_SMT, fil_loc, ItemCode, LotNo, ".xlsm", DGV_Data, "","NPI");
            //DGV_DataView.DataSource = myCode.Load_FAI_Spec_ToTable(sqlcon_SMT, ItemCode, "NPI");
            //foreach (DataGridViewColumn t_col in DGV_DataView.Columns)
            //{
            //    t_col.HeaderText = t_col.HeaderText.Split('_')[0];
            //}
            //if (DGV_DataView.Rows.Count > 0)
            //{
            //    DGV_DataView.Rows[0].HeaderCell.Value = "NormDim";
            //    DGV_DataView.Rows[1].HeaderCell.Value = "Tol_Max";
            //    DGV_DataView.Rows[2].HeaderCell.Value = "Tol_Min";
            //    DGV_DataView.Rows[3].HeaderCell.Value = "Distribution";
            //    DGV_DataView.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            //    myCode. Disable_Sort_DGV(DGV_DataView);
            //}
            DGV_DataView.DataSource = Load_Spec_fromFile(txtFile.Text, ItemCode, LotNo, new List<string> { "*.xlsx", "*.xlsm" });

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
            //DataGridViewCell cur_cell = DGV_DataView.CurrentCell;
            //DataTable dt = (DataTable)DGV_DataView.DataSource;
            //txtInput.Text = cur_cell.Value.ToString();
            //txtOutput.Text = dt.Rows[cur_cell.RowIndex][cur_cell.ColumnIndex].ToString();

            //myExcel.Application xlsApp = TDMK_Code.StartExcel();
            //xlsApp.Visible = true;
            //myExcel.Workbook exp_file;
            //exp_file = xlsApp.Workbooks.Add();
            //Export_FAI_Batch((DataTable)DGV_DataView.DataSource, "422292", "00003", exp_file);

            //DGV_Data_Plus.DataSource = Result_Histogram(DGV_DataView);
            //myCode.DGV_Auto_Resize(DGV_Data_Plus);
            //DGV_Data_Plus.Rows[0].Height = 200;

            //string f_name = txtFile.Text;
            //myExcel.Workbook wrkbk = TDMK_Code.open_excel_file(f_name, "", "");
            //myExcel.Worksheet wrksht = wrkbk.Sheets["ACF"];
            //string start_rgn = common_lib.Get_start_range("ACF pad location", "A10", wrksht);
            //myExcel.Range sel_rgn = wrksht.Range[start_rgn].Offset[1, 0];
            //string stop_rgn = common_lib.Get_start_range("ACF pad location", sel_rgn.AddressLocal, wrksht);
            //Dictionary<string, List<string>> ACF_addr_lst = new Dictionary<string, List<string>>();
            //lstSa.DataSource = common_lib.Get_Listdata_addr("Surface Roughness Measurement", "ACF pad", wrksht, "A10", true, ref ACF_addr_lst);//"ACF pads Surface Roughness"
            //Dictionary<string, Dictionary<string, string>> dic_ACF_Spec = new Dictionary<string, Dictionary<string, string>>();
            //foreach(var spec in ACF_addr_lst)
            //{
            //    string spec_rgn_addr = spec.Value[0];
            //    myExcel.Range spec_rgn = wrksht.Range[spec_rgn_addr].Offset[0, 1];
            //    Dictionary<string, string> spec_ACF = new Dictionary<string, string>();
            //    for (int i = 0; i < wrksht.Range[spec_rgn_addr].MergeArea.Rows.Count; i++)
            //    {
            //        string temp_val = spec_rgn.Offset[i, 0].Value;
            //        List<string> lst = temp_val.Split(new char[] { '(', ')' }).ToList();
            //        string item = lst.First().Trim();
            //        lst.Remove(lst.First());
            //        lst.Remove(lst.Last());
            //        string item_val = lst.First();
            //        spec_ACF.Add(item, item_val);
            //    }
            //    dic_ACF_Spec.Add(spec.Key, spec_ACF);
            //}

            //string f_name = txtFile.Text;
            //DGV_SpecView.DataSource= common_lib.ACF_GetSpec_Process(f_name);
            //if(DGV_SpecView.Rows.Count==3)
            //{
            //    DGV_SpecView.Rows[0].HeaderCell.Value = "SetVal";
            //    DGV_SpecView.Rows[1].HeaderCell.Value = "USL";
            //    DGV_SpecView.Rows[2].HeaderCell.Value = "LSL";
            //}
            //DGV_SpecView.AutoResizeColumns();
            //DGV_SpecView.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            //myCode.Disable_Sort_DGV(DGV_SpecView);

            //DGV_SpecView.SelectAll();
            //DataObject data = DGV_SpecView.GetClipboardContent();
            //Clipboard.SetDataObject(data);
            string _itemcode = "71798T";
            string _lotno = "00003";
            ////DGV_DataView.DataSource = common_lib.Export_FAI_Batch2(itemcode, lotno,"NPI")[0];
            ////myCode.DGV_Auto_Resize(DGV_DataView);
            //DataTable spec_dt = myCode.DGV_To_Table(DGV_DataView);
            //Save_FAI_Spec(spec_dt, itemcode, sqlcon_SMT);

            DataTable FAI_dt = myCode.Load_FAI_ToTable(myVar.sqlcon_SMT, _itemcode, _lotno, "MASS");
            DataTable FAI_Spec = myCode.Load_FAI_Spec_ToTable(myVar.sqlcon_SMT, _itemcode, "MASS");
            DGV_Data.DataSource = FAI_Spec;
            DGV_DataView.DataSource = FAI_dt;

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
            DataTable FAI_Spec_tbl = TDMK_Code.Datatable_Filter(sqlcon_SMT, "FAI_Spec", "ItemCode = '" + tar_ItemCode + "'");
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
                myExcel.Worksheet cur_wrksht = tar_wrkbook.Worksheets.Add(After: tar_wrkbook.Sheets[1]);
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
        public SqlConnection initial_data(string DB_name, bool sa_en)
        {
            SqlConnection _sqlcon_OK2SHIP;
            if (!File.Exists("Config.ini"))
            {
                TDMK_init.Write("Server", "10.212.1.243", "SMT_Config");
                TDMK_init.Write("Account", "sa", "SMT_Config");
                TDMK_init.Write("Password", "seev@123;", "SMT_Config");
                TDMK_init.Write("Report_Location", @"\\10.212.1.200\rgo\OK2SHIP automation system", "SMT_Config");
                TDMK_init.Write("Data_Location", @"\\10.212.1.200\rgo\OK2SHIP automation system\TDMK_Data\TestAreas\OK2SHIP_Measurements\OMM", "SMT_Config");
                TDMK_init.Write("Format_Folder", @"\\10.212.1.200\rgo\OK2SHIP automation system\0.OK2SHIP report format", "SMT_Config");
                TDMK_init.Write("Log_folder", @"\\10.212.1.200\rgo\OK2SHIP automation system\TDMK_Data\TestAreas\OK2SHIP_Measurements", "SMT_Config");
            }
            server_name = TDMK_init.Read("Server", "SMT_Config");
            server_acc = TDMK_init.Read("Account", "SMT_Config");
            server_pass = TDMK_init.Read("Password", "SMT_Config");
            Data_Location = TDMK_init.Read("Data_Location", "SMT_Config");
            format_folder = TDMK_init.Read("Format_Folder", "SMT_Config");
            log_folder = TDMK_init.Read("Log_folder", "SMT_Config");
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

        private void DGV_DataView_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //try
            //{
                //DataTable histo_dt = new DataTable();
                List<string> col_list = new List<string>();
                foreach (DataGridViewCell c in DGV_DataView.SelectedCells)
                {
                    int k = c.ColumnIndex;
                    if (!col_list.Contains(DGV_DataView.Columns[k].Name))
                    {
                        col_list.Add(DGV_DataView.Columns[k].Name);
                        DataColumn column = new DataColumn(DGV_DataView.Columns[k].Name);
                        column.DataType = System.Type.GetType("System.Byte[]");
                       // histo_dt.Columns.Add(column);
                    }
                }
                //histo_dt.Rows.Add();
                DataTable test_tbl = myCode.DGV_To_Table(DGV_DataView);
                int arr_num = col_list.Count;
                TDMK_OK2SHIP.Calculate_CPK[] myCalc_CPK = new TDMK_OK2SHIP.Calculate_CPK[arr_num];
                TDMK_OK2SHIP.FAI_Histogram_Data[] myHistogram_data = new TDMK_OK2SHIP.FAI_Histogram_Data[arr_num];
                int col_list_inx = 0;
                List<Bitmap> histo_lt = new List<Bitmap>();
                foreach (string t in col_list)
                {
                    TDMK_OK2SHIP.FAI_Spec sel_FAI_test;
                    foreach (TDMK_OK2SHIP.FAI_Spec ref_spec in myCode.testFAI_spec)
                    {
                        if (ref_spec.FAI_Name == t)
                        {
                            sel_FAI_test = ref_spec;
                            string[] temp_FAI_data = test_tbl.AsEnumerable().Select(r => r.Field<string>(t)).ToArray();
                            Double[] data_arr = new double[temp_FAI_data.Length];
                            int inx = 0;
                            foreach (string c in temp_FAI_data)
                            {
                                if (myCode.checkDBNull(c) != "")
                                {
                                    data_arr[inx] = Convert.ToDouble(c);
                                    inx++;
                                }
                            }
                            Array.Resize(ref data_arr, inx);
                        List<uint> NG_CPK_inx_lst = Items_NG_CPK_List(sel_FAI_test, data_arr, 7, 50);
                        foreach(var item_inx in NG_CPK_inx_lst)
                        {
                            DGV_DataView.Rows[(int)item_inx].Cells[sel_FAI_test.FAI_Name].Style.BackColor = Color.Aqua;
                        }
                            myCode.Calcul_CPK_FAI2(ref_spec, Calcu_process.ConvertToDouble(temp_FAI_data.ToList()).ToArray(), ref myCalc_CPK[col_list_inx], ref myHistogram_data[col_list_inx]);
                            FrmHistogram chart_form = new FrmHistogram();
                            chart_form.src_bin_data = myHistogram_data[col_list_inx]._Bin_data;// bin_data;
                            chart_form.src_freq_bin_data = myHistogram_data[col_list_inx]._Freq_bin_data;//freq_bin_data;
                            chart_form.src_FAI_Data = myHistogram_data[col_list_inx]._FAI_Data;
                            chart_form.FAI_Spec_val = ref_spec;
                            chart_form.src_CPK_result = myCalc_CPK[col_list_inx];
                            chart_form.src_modified_NormDist = myHistogram_data[col_list_inx]._Modified_NormDist_data; //modified_NormDist;
                            chart_form.Show();
                            //Bitmap cur_hist = Draw_Histogram(myHistogram_data[col_list_inx]._Bin_data, myHistogram_data[col_list_inx]._Freq_bin_data, myHistogram_data[col_list_inx]._Modified_NormDist_data, ref_spec);
                            //histo_dt.Rows[0][t] = imgToByteConverter(cur_hist);
                            //histo_lt.Add(cur_hist);
                            break;
                        }
                    }
                    col_list_inx++;
                }
                //DGV_Data_Plus.DataSource = histo_dt;
                //myCode.DGV_Auto_Resize(DGV_Data_Plus);
                //DGV_Data_Plus.Rows[0].Height = 200;
            //DGV_Data_Plus.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);

            //}
            //catch
            //{

            //}
        }
        private void DGV_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {

        }
        public Bitmap ResizeImage(Image image, int width, int height)
        {
            Size tar_size = GetResizeImage(image, width, height);
            var destRect = new Rectangle(0, 0, tar_size.Width, tar_size.Height);
            var destImage = new Bitmap(tar_size.Width, tar_size.Height);
            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, tar_size.Width, tar_size.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }
        public Bitmap Resize2(Bitmap srcImage, double scaleFactor)
        {
                var newWidth = (int)(srcImage.Width * scaleFactor);
                var newHeight = (int)(srcImage.Height * scaleFactor);
                var newImage = new Bitmap(newWidth, newHeight);
                using (var graphics = Graphics.FromImage(newImage))
                {
                    graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    graphics.DrawImage(srcImage, new Rectangle(0, 0, newWidth, newHeight));
                }
            return newImage;
        }
        private Size GetResizeImage(Image img, int maxWidth, int maxHeight)
        {
            int resizeWidth = img.Width;
            int resizeHeight = img.Height;

            double aspect = resizeWidth / resizeHeight;

            if (resizeWidth > maxWidth)
            {
                resizeWidth = maxWidth;
                resizeHeight = (int)(resizeWidth / aspect);
            }
            if (resizeHeight > maxHeight)
            {
                aspect = resizeWidth / resizeHeight;
                resizeHeight = maxHeight;
                resizeWidth = (int)(resizeHeight * aspect);
            }
            return new Size(resizeWidth, resizeHeight);
        }

        public Image resizeImage(int newWidth, int newHeight, Image imgPhoto)
        {
            //Image imgPhoto = Image.FromFile(stPhotoPath);

            int sourceWidth = imgPhoto.Width;
            int sourceHeight = imgPhoto.Height;

            //Consider vertical pics
            if (sourceWidth < sourceHeight)
            {
                int buff = newWidth;

                newWidth = newHeight;
                newHeight = buff;
            }

            int sourceX = 0, sourceY = 0, destX = 0, destY = 0;
            float nPercent = 0, nPercentW = 0, nPercentH = 0;

            nPercentW = ((float)newWidth / (float)sourceWidth);
            nPercentH = ((float)newHeight / (float)sourceHeight);
            if (nPercentH < nPercentW)
            {
                nPercent = nPercentH;
                destX = System.Convert.ToInt16((newWidth -
                          (sourceWidth * nPercent)) / 2);
            }
            else
            {
                nPercent = nPercentW;
                destY = System.Convert.ToInt16((newHeight -
                          (sourceHeight * nPercent)) / 2);
            }

            int destWidth = (int)(sourceWidth * nPercent);
            int destHeight = (int)(sourceHeight * nPercent);


            Bitmap bmPhoto = new Bitmap(newWidth, newHeight,
                          PixelFormat.Format24bppRgb);

            bmPhoto.SetResolution(imgPhoto.HorizontalResolution,
                         imgPhoto.VerticalResolution);

            Graphics grPhoto = Graphics.FromImage(bmPhoto);
            grPhoto.Clear(Color.White);
            grPhoto.InterpolationMode =
                System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

            grPhoto.DrawImage(imgPhoto,
                new Rectangle(destX, destY, destWidth, destHeight),
                new Rectangle(sourceX, sourceY, sourceWidth, sourceHeight),
                GraphicsUnit.Pixel);

            grPhoto.Dispose();
            imgPhoto.Dispose();

            return bmPhoto;
        }
        public Bitmap Draw_Histogram(double[] src_bin, int[] src_freq_bin, double[] src_modified, TDMK_OK2SHIP.FAI_Spec src_FAI_spec)
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
            if (myCode.IsNumeric_Val(src_FAI_spec.UL) != "")
            {
                double UL = Convert.ToDouble(src_FAI_spec.SetVal) + Convert.ToDouble(src_FAI_spec.UL);
                USL_list.Add(UL, freq_bin_Min);
                USL_list.Add(UL, freq_bin_Max);
                LineItem UL_curve = myPane.AddCurve("USL", USL_list, Color.Red, SymbolType.None);
            }
            if (myCode.IsNumeric_Val(src_FAI_spec.LL) != "")
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
            return myPane.GetImage(300, 200, 1200);
        }
        public byte[] imgToByteConverter(Image inImg)
        {
            ImageConverter imgCon = new ImageConverter();
            return (byte[])imgCon.ConvertTo(inImg, typeof(byte[]));
        }
        public DataTable Result_Histogram(DataGridView src_DGV)
        {
            DataTable src_dt = myCode.DGV_To_Table(src_DGV);
            DataTable histo_dt = new DataTable();
            List<string> col_list = new List<string>();
            foreach(DataColumn dc in src_dt.Columns)
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
            TDMK_OK2SHIP.Calculate_CPK[] myCalc_CPK = new TDMK_OK2SHIP.Calculate_CPK[arr_num];
            TDMK_OK2SHIP.FAI_Histogram_Data[] myHistogram_data = new TDMK_OK2SHIP.FAI_Histogram_Data[arr_num];
            int col_list_inx = 0;
            List<Bitmap> histo_lt = new List<Bitmap>();
            foreach (string t in col_list)
            {
                foreach (TDMK_OK2SHIP.FAI_Spec ref_spec in myCode.testFAI_spec)
                {
                    if (ref_spec.FAI_Name == t)
                    {
                        string[] temp_FAI_data = src_dt.AsEnumerable().Where(x=>x.Field<string>(t)!=null).Select(r => r.Field<string>(t)).ToArray();
                        if(temp_FAI_data.Length>0)
                        {
                            Double[] data_arr = new double[temp_FAI_data.Length];
                            int inx = 0;
                            foreach (string c in temp_FAI_data)
                            {
                                if (myCode.checkDBNull(c) != "")
                                {
                                    data_arr[inx] = Convert.ToDouble(c);
                                    inx++;
                                }
                            }
                            Array.Resize(ref data_arr, inx);
                            List<int> item_NG_lst = new List<int>();
                            myCode.Calcul_CPK_FAI_Located_NG(ref_spec, temp_FAI_data, ref myCalc_CPK[col_list_inx], ref myHistogram_data[col_list_inx], item_NG_lst);
                            foreach(int item_inx in item_NG_lst)
                            {
                                src_DGV.Rows[item_inx].Cells[t].Style.BackColor = Color.Aqua;
                            }
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
        public List<string> Items_OK_CPK_List(Double[] data_arr, string USL, string LSL, int sigma, int count)
        {
            List<SortedDictionary<uint, double>> myData1 = new List<SortedDictionary<uint, double>>();
            List<SortedDictionary<uint, double>> myResult1 = new List<SortedDictionary<uint, double>>();
            List<string> result = new List<string>();
            double UL = Convert.ToDouble(USL);
            double LL = Convert.ToDouble(LSL);
            double CPKL = 0;
            double CPKU = 0;
            double stdev =  myCode. CalculateStandardDeviation(data_arr);
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

                NormDist_Bin[i] = TDMK_OK2SHIP.normdist(bin_data[i], mean, stdev, false);
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
        public List<uint> Items_NG_CPK_List(TDMK_OK2SHIP.FAI_Spec src_FAI_spec, Double[] data_arr, int sigma, int count)
        {
            List<SortedDictionary<uint, double>> myData1 = new List<SortedDictionary<uint, double>>();
            List<SortedDictionary<uint, double>> myResult1 = new List<SortedDictionary<uint, double>>();
            List<uint> result = new List<uint>();
            string Side_check = src_FAI_spec.check_side;
            double Norminal_Dim = Convert.ToDouble(src_FAI_spec.SetVal);
            double UL = 0;
            double LL = 0;
            if (myCode. IsNumeric_Val(src_FAI_spec.UL) != "")
            {
                UL = Norminal_Dim + Convert.ToDouble(src_FAI_spec.UL);
            }
            if (myCode.IsNumeric_Val(src_FAI_spec.LL) != "")
            {
                LL = Norminal_Dim - Convert.ToDouble(src_FAI_spec.LL);
            }
            double stdev = Calcu_process. CalculateStandardDeviation(data_arr);// tg.StDev(data_arr);
            double mean = data_arr.Average();
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

                NormDist_Bin[i] = Calcu_process.normdist(bin_data[i], mean, stdev, false);
            }
            XiShu = freq_bin_data.Max() / NormDist_Bin.Max();
            for (int i = 0; i < count; i++)
            {
                modified_NormDist[i] = NormDist_Bin[i] * XiShu;
                if ((modified_NormDist[i] < freq_bin_data[i]) && (freq_bin_data[i]>0))
                {
                    myResult1.Add(myData1[i]);
                }
            }
            foreach (var item in myResult1)
            {
                foreach (KeyValuePair<uint, double> x in item)
                {
                    result.Add(x.Key);
                }
            }
            return result;
        }
        public DataTable Load_Spec_fromFile(string format_loc, string tar_ItemCode, string tar_LotNo, List<string> extensions, string format_type = "NPI")
        {
            DataTable spec_dt = new DataTable();
            string tar_format_file;        
            AutoCompleteStringCollection FAI_SheetNo_list = new AutoCompleteStringCollection();
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { txtItemCode.Text, format_type });
            FAI_SheetNo_list = TDMK_Code.Load_Item_Names_Filter(myVar.sqlcon_SMT, "FAI_Spec", "SheetNo", filter_str);
            List<string> file_format_lst = new List<string>();// = Directory.GetFiles(format_loc, "*" + file_extension).Where(s => s.Contains(tar_ItemCode)).ToArray();
            DirectoryInfo directory = new DirectoryInfo(format_loc);
            var files = extensions.SelectMany(e => directory.EnumerateFiles(e, SearchOption.AllDirectories)).Where(x => x.FullName.Contains(tar_ItemCode));
            foreach (var item in files)
            {
                file_format_lst.Add(item.FullName);
            }
            if (file_format_lst.Count > 0)
            {
                tar_format_file = file_format_lst[0];
                myExcel.Workbook tar_wkbook = TDMK_Code.open_excel_file(tar_format_file, "", "");
                List<string> sht_keys = new List<string> { "FAI", "SPC", "parentheses" };
                int wrksheet_num = tar_wkbook.Worksheets.Count;
                foreach (myExcel.Worksheet tg in tar_wkbook.Worksheets)
                {
                    if(sht_keys.Any(x=>tg.Name.Contains(x)))
                    {
                        tg.Activate();
                        string dim_no_addr = myCode.Find_Cell_Addr("Dim. No.", "B10", tg, false);
                        string instrument_addr = myCode.Find_Cell_Addr("instrument", "B10", tg, false);
                        string FAI_data_addr = myCode.Find_Offset(tg.Range[dim_no_addr].Offset[0, 1].AddressLocal, tg, true, "");
                        int off_set = tg.Range[FAI_data_addr].Column - tg.Range[instrument_addr].Column;
                        myExcel.Range sel_rgn = tg.Range[dim_no_addr].Offset[0, off_set];// tg.Range["D19"];    //tg.Range["C17"];
                        myExcel.Range dev_rgn = tg.Range[instrument_addr].Offset[0, off_set];// tg.Range["D23"];    //tg.Range["C21"]
                        int sel_inx = 0;
                        while (myCode.checkDBNull(sel_rgn.Offset[0, sel_inx].Value) != "")
                        {
                            string t_checkside = myCode.checkDBNull(sel_rgn.Offset[-1, sel_inx].Value);
                            string t_FAIName = myCode.checkDBNull(sel_rgn.Offset[0, sel_inx].Value).Replace(" ", "");
                            string t_FAI_Setval = myCode.checkDBNull(sel_rgn.Offset[1, sel_inx].Value);
                            if (t_FAI_Setval == "")
                            {
                                t_FAI_Setval = "NA";
                            }
                            string t_FAI_UL = myCode.checkDBNull(sel_rgn.Offset[6, sel_inx].Value); ;// myCode.checkDBNull(sel_rgn.Offset[2, sel_inx].Value);
                            string t_FAI_LL = myCode.checkDBNull(sel_rgn.Offset[7, sel_inx].Value); ;// myCode.checkDBNull(sel_rgn.Offset[3, sel_inx].Value);
                            string col_name = t_FAIName + "_" + t_FAI_Setval;
                            string t_FAI_sheetno = tg.Name;
                            string t_instrument = myCode.checkDBNull(dev_rgn.Offset[0, sel_inx].Value);
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
                                spec_dt.Rows[5][col_name] = tg.Name; // Instrument at Row =4
                            }
                            sel_inx++;
                        }
                    }
                }
                tar_wkbook.Close();
                MessageBox.Show(new Form { TopMost = true }, "Completed set up Spec  for ItemCode: " + txtItemCode.Text + "!");
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "File format not found");
            }
            return spec_dt;
        }
        public void Save_FAI_Spec(DataTable src_spec_dt,string tar_ItemCode, SqlConnection sqlcon, string type ="NPI")
        {
            string[] items = new string[11];
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
            items[10] = "Remark";
            int id = TDMK_Code.SQL_MAX("FAI_Spec", "ID", sqlcon);
            foreach (DataColumn dc in src_spec_dt.Columns)
            {
                string[] item_vals = new string[11];
                item_vals[1] = tar_ItemCode;
                item_vals[10] = type;
                item_vals[0] = id++.ToString();
                item_vals[7] = dc.ColumnName;
                item_vals[3] = src_spec_dt.Rows[0][dc].ToString();// SetVal
                item_vals[4] = src_spec_dt.Rows[1][dc].ToString(); ;// t.UL;
                item_vals[5] = src_spec_dt.Rows[2][dc].ToString(); ;// t.LL;
                item_vals[8] = src_spec_dt.Rows[3][dc].ToString(); ;// t.check_side;
                item_vals[6] = src_spec_dt.Rows[4][dc].ToString(); ;// Instruments
                item_vals[9] = src_spec_dt.Rows[5][dc].ToString(); ;// t.sheetno;
                TDMK_Code.insert_val_arr2("FAI_Spec", sqlcon, items, item_vals);
            }
            MessageBox.Show("Lưu spec của ItemCode " + tar_ItemCode + " thành công", "Thông báo");
        }
        //public void save_FAI_spec(FAI_Spec[] src_FAI_spec, string tar_ItemCode, SqlConnection sqlcon, string type = "NPI")
        //{
        //    string[] items = new string[11];
        //    string[] item_vals = new string[11];
        //    items[0] = "ID";
        //    items[1] = "ItemCode";
        //    items[2] = "LotNo";
        //    items[3] = "NormDim";
        //    items[4] = "TolMax";
        //    items[5] = "TolMin";
        //    items[6] = "Instrument";
        //    items[7] = "FAI_No";
        //    items[8] = "Distribution";
        //    items[9] = "SheetNo";
        //    items[10] = "Remark";
        //    item_vals[1] = tar_ItemCode;//"ItemCode";
        //    item_vals[10] = type;
        //    //item_vals[2] = txtLotNo.Text;//"LotNo";
        //    if (src_FAI_spec.Length > 0)
        //    {
        //        foreach (FAI_Spec t in src_FAI_spec)
        //        {
        //            AutoCompleteStringCollection FAI_No_list = new AutoCompleteStringCollection();
        //            item_vals[3] = "";
        //            item_vals[4] = "";
        //            item_vals[5] = "";
        //            item_vals[8] = "";
        //            item_vals[6] = t.instrument;
        //            item_vals[7] = t.FAI_Name;
        //            FAI_No_list = TDMK_Code.Load_Item_Filter_str(sqlcon, "FAI_Spec", "NormDim", TDMK_Code.filter_str(items, item_vals));
        //            if (FAI_No_list.Count == 0)
        //            {
        //                item_vals[0] = (TDMK_Code.SQL_MAX("FAI_Spec", "ID", sqlcon) + 1).ToString();
        //                item_vals[3] = t.SetVal;
        //                item_vals[4] = t.UL;
        //                item_vals[5] = t.LL;
        //                item_vals[8] = t.check_side;
        //                item_vals[9] = t.sheetno;
        //                TDMK_Code.insert_val_arr("FAI_Spec", sqlcon, items, item_vals);
        //            }
        //        }
        //    }
        //}
        public bool check_in_limit(string src_UL, string src_LL, string src_act_val, string SetVal)
        {
            bool _result = false;
            bool UL_equal_comp = true;
            bool LL_equal_comp = true;

            if (!myCode. IsNumeric(SetVal) && (!SetVal.Contains("=")))
            {
                UL_equal_comp = false;
                LL_equal_comp = false;
            }
            if (myCode.IsNumeric_Val(src_act_val)!="")
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
                if(cur_col== "FAI47-14_0.213")
                {
                    MessageBox.Show("Here");
                }
                if (spec_col_lst.IndexOf(cur_col) != -1)
                {
                    string USL = myCode. checkDBNull(src_Spec.Rows[1][cur_col]);
                    string LSL = myCode.checkDBNull(src_Spec.Rows[2][cur_col]);
                    string SetVal = myCode.IsNumeric_Val(myCode.checkDBNull(src_Spec.Rows[0][cur_col]));
                    string UL = myCode.IsNumeric_Val(USL);
                    string LL = myCode.IsNumeric_Val(LSL);
                    foreach (DataRow dr in src_data.Rows)
                    {
                        string src_act_val = myCode.IsNumeric_Val(myCode.checkDBNull(dr[cur_col]));
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

    }
}
