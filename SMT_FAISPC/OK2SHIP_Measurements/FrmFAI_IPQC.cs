using OK2SHIP_Software;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using TDMK_SQL;
using ZedGraph;
using myExcel = Microsoft.Office.Interop.Excel;
using OK2SHIP_Lib;
using FAI_Export_EPPLUS;
using OfficeOpenXml;
//using TDMK_EPPLUS_7;
using FAI_Export_EPPLUS;

namespace OK2SHIP_Measurements
{
    public partial class FrmFAI : Form
    {
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public TDMK_OK2SHIP myCode = new TDMK_OK2SHIP();
        myVar myCode2 = new myVar();
        FAI_EPPLUS_Lib FAI_lib = new FAI_EPPLUS_Lib();
        FAI_EPPLUS_Lib Excel_Lib = new FAI_EPPLUS_Lib();
        public string app_path;
        public string origin_cell_val;
        public bool edit_en = false;
        public bool pro_en = false;
        public delegate void Data_process(string _filename, DataGridView tar_DGV, DataGridView DGV_Spec, DataGridView tar_DGV_plus, bool edit_mode);
        public delegate void Data_process2(string _filename, DataGridView tar_DGV, DataGridView DGV_Spec, DataGridView tar_DGV_plus, bool edit_mode, int qty);
        public delegate void Roughness_process(string _filename, DataGridView tar_DGV, DataGridView DGV_Spec, DataGridView tar_DGV_plus, bool edit_mode, int qty, int MeasLoc);
        //public Data_process my_process;
        public delegate void SetText(Button tar_btn, string btn_str);
        public delegate void checkTextbox(TextBox src_txt);
        public Color curr_color;
        public SqlConnection sqlcon { get; set; }
        public List<string> list_FAI_NG = new List<string>();
        //public List<TDMK_OK2SHIP.FAI_Spec> cur_FAI_Spec = new List<TDMK_OK2SHIP.FAI_Spec>();
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
        public string Dev { get; set; }
        public string DB_name { get; set; }
        public string Data_Location { get; set; }
        public string format_folder { get; set; }
        public string log_folder { get; set; }
        public string report_location { get; set; }
        public string f_ext { get; set; }
        public string Server_name;
        public bool EditMode;
        public FileSystemWatcher watchfolder = new FileSystemWatcher();
        //FAI_Spec[] testFAI = new FAI_Spec[1000];
        //TDMK_OK2SHIP.FAI_Spec[] testFAI_spec;
        public FrmFAI()
        {
            InitializeComponent();
        }
        public FrmFAI(string src_Dev, string src_DB, SqlConnection tar_sqlcon, string tar_datalocation, string tar_formatfolder,string tar_reportlocation, string tar_f_ext, string tar_log_folder)
        {
            InitializeComponent();
            Dev = src_Dev;
            DB_name = src_DB;
            sqlcon = tar_sqlcon;
            Data_Location = tar_datalocation;
            format_folder = tar_formatfolder;
            f_ext = tar_f_ext;
            log_folder = tar_log_folder;
            report_location = tar_reportlocation;
            myVar.sqlcon_SMT = tar_sqlcon;
        }

        private void DGV_DataView_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridViewCell curr_cell;
            string curr_col_name;
            if (e.Button == MouseButtons.Right)
            {
                if (TDMK_OK2SHIP.admin_mode)
                {
                    if (TDMK_OK2SHIP.operate_mode == "Auto")
                    {
                        curr_cell = DGV_DataView.CurrentCell;
                        curr_col_name = DGV_DataView.Columns[curr_cell.ColumnIndex].Name.ToString();
                        origin_cell_val = checkDBNull(curr_cell.Value);
                        DGV_DataView.BeginEdit(true);
                        edit_en = true;
                    }
                    else
                    {
                        if (DGV_DataView.SelectedCells.Count>0)
                        {
                            DGV_DataView.ContextMenuStrip = cmsCopyPaste;
                        }
                    }

                }
                else
                {
                   MessageBox.Show(new Form { TopMost = true },"Bạn cần đăng nhập để chỉnh sửa", "Thông báo");
                }
            }
        }
        private void DGV_DataView_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            if (Dev == "Manual")
            {
                edit_en = true;
            }
            else
            {
                edit_en = false;
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
        private void Form1_Load(object sender, EventArgs e)
        {
            //app_path = @"D:\Customer Projects\SEEV\TestAreas";
            app_path = Application.StartupPath;
            Process[] p;
            p = Process.GetProcessesByName("OK2SHIP_Measurements");
            if (p.Count() > 1)
            {
               MessageBox.Show(new Form { TopMost = true }, "OK2SHIP_Measurements is running");
                this.Close();
            }
            else
            {
                string watch_path = Data_Location;
                string[] filter_ext = new string[] { ".xls", ".xlsx" };
                watchfolder.Path = watch_path;
                watchfolder.NotifyFilter = NotifyFilters.DirectoryName | NotifyFilters.FileName | NotifyFilters.Attributes;
                if(Dev == "Roughness")
                {
                    //watchfolder.Filter = "*.xls";
                    f_ext = ".xls";
                    numMeasLoc.Enabled = true;
                }
                else
                {
                    //watchfolder.Filter = f_ext;// "*.csv";
                    numMeasLoc.Enabled = false;
                }
                
                watchfolder.Created += Watchfolder_Created;
                watchfolder.EnableRaisingEvents = true;
                curr_color = this.BackColor;
                EditMode = false;
                RBNormal.Checked = true;
                txtDataFolder.Text = Data_Location;
                if (Dev == "Manual")
                {
                    cbMachine.Enabled = true;
                    DGV_DataView.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
                    TDMK_OK2SHIP.admin_mode = true;
                }
                else
                {
                    cbMachine.Text = Dev;
                    cbMachine.Enabled = false;
                    DGV_DataView.EditMode = DataGridViewEditMode.EditProgrammatically;
                }
                cbType.SelectedIndex = 1;
                cbShift.SelectedIndex = 0;
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            }
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
                    if (!e.FullPath .Contains ("$"))
                    {
                        //TDMK_OK2SHIP.scan_file = e.FullPath;
                        if (pro_en)
                        {
                            Thread tg_thread = new Thread(Scan_data);
                            tg_thread.Start(e.FullPath);
                        }
                        else
                        {
                            itemcode_folder = Path.Combine(app_path, "Ignored_Scan", txtItemCode.Text);
                            done_file = Path.Combine(itemcode_folder, "Ignored_" + save_time + f_name);
                            if (!Directory.Exists(itemcode_folder))
                            {
                                Directory.CreateDirectory(itemcode_folder);
                            }
                            File.Copy(f_info, done_file, true);
                            File.Delete(f_info);
                           MessageBox.Show(new Form { TopMost = true },"Bạn chưa ấn nút Start", "Thông tin");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Error_log(app_path, ex.Message);
                }
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            /********************************************* Main Program ******************************************************************/
            main_program();

            /**********************************************************************************************************************************************/
            /********************************************* Test  Areas ******************************************************************/

           //myCode. Load_Test_Data(@"D:\Customer Projects\SEEV\TestAreas\TestData", txtItemCode.Text, txtLotNo.Text, ".xlsm", DGV_DataView);

            /**********************************************************************************************************************************************/
        }

        public void main_program()
        {
            if (!TDMK_OK2SHIP.confirm_request)
            {
                if (DGV_SpecView.Rows.Count > 0)
                {
                    if (btnRun.Text == "Start")
                    {
                        if (Dev != "Manual")
                        {
                            watchfolder.Path = txtDataFolder.Text;
                        }
                        else
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
                            if (sel_DGV.ColumnCount == 0)
                            {
                                foreach (DataGridViewColumn t in DGV_SpecView.Columns)
                                {
                                    sel_DGV.Columns.Add(t.Name, t.Name);
                                }
                            }
                            if (sel_DGV.Rows.Count == 0)
                            {
                                for (int i = 0; i < Convert.ToInt32(numQty.Value); i++)
                                {
                                    sel_DGV.Rows.Add();
                                    sel_DGV.Rows[i].HeaderCell.Value = (i + 1).ToString();
                                }
                                sel_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                                Disable_Sort_DGV(sel_DGV);
                            }
                            sel_DGV.EditMode = DataGridViewEditMode.EditOnEnter;
                        }
                        btnRun.Text = "Stop";
                        timer1.Enabled = true;
                        pro_en = true;
                        btnSave.Enabled = false;
                        btnLoadSpec.Enabled = false;
                    }
                    else
                    {
                        DGV_Data_Plus.EditMode = DataGridViewEditMode.EditProgrammatically;
                        DGV_DataView.EditMode = DataGridViewEditMode.EditProgrammatically;
                        btnRun.Text = "Start";
                        pro_en = false;
                        timer1.Enabled = false;
                        btnSave.Enabled = true;
                        btnRun.BackColor = curr_color;
                    }
                }
                else
                {
                   MessageBox.Show(new Form { TopMost = true },"Press Load Spec!", "Warning");
                }

            }
            else
            {
               MessageBox.Show(new Form { TopMost = true },"Please, reset at first!", "Warning");
            }
        }
        private void btnReset_Click(object sender, EventArgs e)
        {           
            if (TDMK_OK2SHIP.confirm_request)
            {
                TDMK_OK2SHIP.confirm_request = false;
            }
            else
            {
                myCode.Clear_DGV(DGV_DataView);
                myCode.Clear_DGV(DGV_Data_Plus);
                //myCode.Clear_DGV(DGV_SpecView);
                myCode.Clear_DGV(DGV_CPK);
                myCode.Clear_DGV(DGV_Histogram);
                //DGV_DataView.Columns.Clear();
                //DGV_Data_Plus.Columns.Clear();
                //DGV_SpecView.Columns.Clear();
                //DGV_CPK.Columns.Clear();
                btnLoadSpec.Enabled = true;
                TDMK_OK2SHIP.confirm_request = false;
                //btnLogin.Text = "Login";
                //btnLogin.BackColor = this.BackColor;
                pro_en = false;
                Invoke(new SetText(SetTextButton), btnRun, "Start");
            }
            TDMK_OK2SHIP.admin_mode = false;
            btnLogin.Text = "Login";
            btnLogin.BackColor = this.BackColor;
        }
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
            if(DGV_DataView.Rows.Count>= Convert.ToInt32(numQty.Value))
            {
                //src_btn.Text = txt_str;
                //src_btn.BackColor = curr_color;
                //pro_en = false;
                //timer1.Enabled = false;
                //btnSave.Enabled = true;
                //myCode.Load_CPK_DGV(myCode.DGV_To_Table(DGV_DataView), DGV_CPK, myCode.testFAI_spec.ToArray());
                if(Dev!="Roughness")
                {
                    myCode.Load_CPK_DGV_Histogram(DGV_DataView, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
                }
                //if ((DGV_Data_Plus.Rows.Count > 0) && (!TDMK_OK2SHIP.admin_mode) && (!EditMode))
                //{
                //    TDMK_OK2SHIP.confirm_request = true;
                //    GBControl.Enabled = false;
                //}
                if(DGV_DataView.Rows.Count > Convert.ToInt32(numQty.Value))
                {
                    MessageBox.Show(new Form { TopMost = true },"input data is more than " + numQty.Value.ToString(), "Warning");
                }
            }
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            bool export_en = false;
            if (myCode.check_data_enough(DGV_DataView, Convert.ToInt32(numQty.Value)))
            {
                export_en = true;
            }
            else
            {
                if( MessageBox.Show(new Form { TopMost = true },"Not enough data", "Warning",MessageBoxButtons.YesNo)==DialogResult.Yes)
                {
                    export_en = true;
                }
                else
                {
                    export_en = false;
                }
            }
            if(export_en)
            {
                bool en_save = false;
                bool data_ok = false;
                if ((Dev == "Roughness") || (Dev == "AU_NI_Thickness"))
                {
                    data_ok = myCode.check_data_OK(DGV_DataView, DGV_SpecView);
                }
                else
                {
                    data_ok = myCode.check_FAIdata_OK2(myCode.DGV_To_Table(DGV_DataView), myCode.DGV_To_Table(DGV_SpecView));
                }
                if (data_ok)
                {
                    en_save = true;
                }
                else
                {
                    if ((MessageBox.Show("Data out of spec. Do you want to save it ?", "Warning", MessageBoxButtons.YesNo)) == DialogResult.Yes)
                    {
                        if (TDMK_OK2SHIP.admin_mode)
                        {
                            en_save = true;
                        }
                        else
                        {
                            en_save = false;
                            MessageBox.Show(new Form { TopMost = true }, "Please, login to save data", "Warning");
                        }

                    }
                    else
                    {
                        en_save = false;
                    }
                }
                if (en_save)
                {
                    if (Dev != "Roughness")
                    {
                        Save_FAI_data();
                    }
                    else
                    {
                        Save_data_Roughness();
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
            if (EditMode)
            {
                myExcel.Application xlsApp = TDMK_Code.StartExcel();
                xlsApp.Visible = true;
                myExcel.Workbook export_wrkbook = xlsApp.Workbooks.Add();
                TDMK_Code.Export_DGV_Excel2(DGV_Data_Plus, export_wrkbook);
            }
            else
            {
                DataTable spec_dt = (DataTable)DGV_SpecView.DataSource;
                List<string> spec_col_list = new List<string>();
                DataTable FAI_dt = myCode.DGV_To_Table(DGV_DataView);
                List<string> FAI_col_list = new List<string>();
                foreach (DataColumn dc in spec_dt.Columns)
                {
                    if (spec_col_list.IndexOf(dc.ColumnName) == -1)
                    {
                        spec_col_list.Add(dc.ColumnName);
                    }
                }
                foreach (DataColumn dc in FAI_dt.Columns)
                {
                    string col = dc.ColumnName;//.Split('_')[0];
                    if (FAI_col_list.IndexOf(col) == -1)
                    {
                        FAI_col_list.Add(col);
                    }
                }
                if ((spec_col_list.Count > 0) && (FAI_col_list.Count > 0))
                {
                    List<string> exp_col_lst = spec_col_list.Intersect(FAI_col_list).ToList();
                    if (exp_col_lst.Count > 0)
                    {
                        DataTable exp_tbl = FAI_dt.AsDataView().ToTable(false, exp_col_lst.ToArray());
                        myExcel.Application xlsApp = TDMK_Code.StartExcel();
                        xlsApp.Visible = true;
                        myExcel.Workbook export_wrkbook = xlsApp.Workbooks.Add();
                        myExcel.Worksheet cur_wrksht = export_wrkbook.Sheets[1];
                        myExcel.Range start_rgn = cur_wrksht.Range["A1"];
                        int c_inx = 0;
                        foreach (DataColumn dc in exp_tbl.Columns)
                        {
                            start_rgn.Offset[0, c_inx].Value = dc.ColumnName;
                            int r_inx = 0;
                            foreach (DataRow dr in exp_tbl.Rows)
                            {
                                start_rgn.Offset[1 + r_inx, c_inx].Value = dr[dc].ToString();
                                r_inx++;
                            }
                            c_inx++;
                        }
                    }
                }
            }
            /*************************** Test Areas **************************************************************************/
            //myCode.export_data = new Thread(export_to_Excel);
            //myCode.export_data.Start();

            //myCode.Load_CPK_DGV(myCode.DGV_To_Table( DGV_DataView), DGV_CPK, myCode.testFAI_spec);
            //string f_name_keyence = @"D:\Customer Projects\SEEV\TestAreas\Raw_data\Keyence.csv";
            //string f_name_nikon = @"D:\Customer Projects\SEEV\TestAreas\Raw_data\Nikon.csv";
            //string f_name_Mitutoyo = @"D:\Customer Projects\SEEV\TestAreas\Raw_data\Mitutoyo.csv";
            //myCode.Keyence_Display_Data(f_name_keyence, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode);
            //myCode.Nikon_Data_display(f_name_nikon, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode);
            //myCode.Mitutoyo_Data_display(f_name_Mitutoyo, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode);

            //myExcel.Application xlsApp = TDMK_Code.StartExcel();
            //xlsApp.Visible = true;
            //myExcel.Workbook export_wrkbook = xlsApp.Workbooks.Add();
            //if (EditMode )
            //{
            //    TDMK_Code.Export_DGV_Excel2(DGV_Data_Plus, export_wrkbook);
            //}
            //else
            //{
            //    TDMK_Code.Export_DGV_Excel2(DGV_DataView, export_wrkbook);
            //}

            /*************************** Finish Test Areas *******************************************************************/
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
        public void Save_FAI_data()
        {
            char[] trim_char = new char[] { ' ', '\r', '\n' };
            bool data_rec_en = true;
            string[] items = new string[10];
            string[] item_vals = new string[10];
            items[0] = "ID";
            items[1] = "ItemCode";
            items[2] = "LotNo";
            items[3] = "Operator";
            items[4] = "Machine";
            items[5] = "MDate";
            items[6] = "FAI_No";
            items[7] = "FAI_Data";
            items[8] = "Remark";
            items[9] = "Shift";
            item_vals[1] = txtItemCode.Text;//"ItemCode";
            item_vals[2] = txtLotNo.Text;//"LotNo";
            item_vals[3] = txtOperator.Text;//"Operator";
            item_vals[4] = cbMachine.Text;//"Machine";
            item_vals[5] = DateTime.Now.ToString();//"MDate";
            item_vals[8] = cbType.SelectedItem.ToString();
            item_vals[9] = cbShift.SelectedItem.ToString();
            DataTable tbl_spec = myCode.DGV_To_Table(DGV_SpecView);
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
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo","Machine", "Remark", "Shift" }, new string[] { txtItemCode.Text, txtLotNo.Text,cbMachine.Text, cbType.SelectedItem.ToString(), cbShift.SelectedItem.ToString() });
    start_label:  DataTable fai_dt = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Auto", filter_str);
                if(fai_dt.Rows.Count>0)
                {
                    if(MessageBox.Show("Bạn muốn ghi đè dữ liệu?", "Thông báo",MessageBoxButtons.YesNo)==DialogResult.Yes)
                    {
                        TDMK_Code.Delelte_FilteredItem_arr("FAI_Auto", sqlcon, filter_str);
                        goto start_label;
                    }    
                }
                if (DGV_DataView.Rows.Count > 0)
                {
                    for (int i = 0; i < DGV_DataView.ColumnCount; i++)
                    {
                        string curr_col_name = DGV_DataView.Columns[i].Name;
                        if (myCode. check_columns_existed(tbl_spec, curr_col_name))
                        {
                            item_vals[6] = curr_col_name;
                            for (int j = 0; j < DGV_DataView.RowCount; j++)
                            {
                                if (checkDBNull(DGV_DataView.Rows[j].Cells[i].Value) != "")
                                {
                                    item_vals[0] = (TDMK_Code.SQL_MAX("FAI_Auto", "ID", sqlcon) + 1).ToString();//"ID";
                                    item_vals[7] = checkDBNull(DGV_DataView.Rows[j].Cells[i].Value).Trim(trim_char);
                                    TDMK_Code.insert_val_arr("FAI_Auto", sqlcon, items, item_vals);
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }
                    }
                    DGV_DataView.Columns.Clear();
                    DGV_CPK.Columns.Clear();
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
            if (RBNormal.Checked)
            {
                EditMode = false;
                DGV_SpecView.ContextMenuStrip = null;
            }
           
        }
        private void RBEdit_CheckedChanged(object sender, EventArgs e)
        {
            if (RBEdit.Checked)
            {
                EditMode = true;
            }
            
        }

        private void DGV_Data_Plus_MouseClick(object sender, MouseEventArgs e)
        {
            if ((e.Button == MouseButtons.Right)&& DGV_Data_Plus.Columns .Count >0)
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
                            if (myCode. check_columns_existed(myCode.DGV_To_Table(DGV_SpecView), src_col_name))
                            {
                                string SetVal = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[0].Cells[src_col_name].Value));
                                string UL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[1].Cells[src_col_name].Value));
                                string LL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[2].Cells[src_col_name].Value));
                                string Act_val = myCode.checkDBNull(c.Value);
                                DGV_DataView.Rows[sel_row_inx].Cells[src_col_name].Value = c.Value;
                                DGV_DataView.Rows[sel_row_inx].Cells[src_col_name].Style.BackColor = myCode.check_in_limit_Color(UL, LL, Act_val, SetVal.ToString());
                                c.Value = "";
                            }
                        }
                        catch
                        {
                           MessageBox.Show(new Form { TopMost = true },"Sai vi tri", "Thong bao");
                        }
                    }
                    //myCode.Load_CPK_DGV_ColName(myCode.DGV_To_Table(DGV_DataView), DGV_CPK, myCode.testFAI_spec.ToArray());
                    myCode.Load_CPK_DGV_Histogram(DGV_DataView, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
                }
                else
                {
                   MessageBox.Show(new Form { TopMost = true },"Chọn dòng");
                }
            }
            else
            {
               MessageBox.Show(new Form { TopMost = true },"Bạn cần đăng nhập để chỉnh sửa", "Thông báo");
            }
        }
        private void DGV_DataView_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                List<string> col_list = new List<string>();
                foreach (DataGridViewCell c in DGV_DataView.SelectedCells)
                {
                    int k = c.ColumnIndex;
                    if (!col_list.Contains(DGV_DataView.Columns[k].Name))
                    {
                        col_list.Add(DGV_DataView.Columns[k].Name);
                    }
                }

                DataTable test_tbl = myCode.DGV_To_Table(DGV_DataView);
                int arr_num = col_list.Count;
                TDMK_OK2SHIP.Calculate_CPK[] myCalc_CPK = new TDMK_OK2SHIP.Calculate_CPK[arr_num];
                TDMK_OK2SHIP.FAI_Histogram_Data[] myHistogram_data = new TDMK_OK2SHIP.FAI_Histogram_Data[arr_num];
                int col_list_inx = 0;
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
                            myCode.Calcul_CPK_FAI2(ref_spec, data_arr, ref myCalc_CPK[col_list_inx], ref myHistogram_data[col_list_inx]);
                            FrmHistogram chart_form = new FrmHistogram();
                            chart_form.src_bin_data = myHistogram_data[col_list_inx]._Bin_data;// bin_data;
                            chart_form.src_freq_bin_data = myHistogram_data[col_list_inx]._Freq_bin_data;//freq_bin_data;
                            chart_form.src_FAI_Data = myHistogram_data[col_list_inx]._FAI_Data;
                            chart_form.FAI_Spec_val = ref_spec;
                            chart_form.src_CPK_result = myCalc_CPK[col_list_inx];
                            chart_form.src_modified_NormDist = myHistogram_data[col_list_inx]._Modified_NormDist_data; //modified_NormDist;
                            chart_form.Show();
                            break;
                        }
                    }
                    col_list_inx++;
                }
            }
            catch
            {

            }


        }

        private void btnLoadSpec_Click(object sender, EventArgs e)
        {
            /******************************************** Main Program *****************************************************************/
            bool en_proc = false;
            switch (Dev)
            {
                case "AU_NI_Thickness":
                    //en_proc = Load_Spec_IPQC(sqlcon, "SpecList", txtItemCode.Text, "THICKNESS", DGV_SpecView);
                    en_proc = myCode.Load_Spec_IPQC(sqlcon, "SpecList", txtItemCode.Text, "THICKNESS", DGV_SpecView, format_folder, ".xlsm");
                    break;
                case "Roughness":
                    //en_proc = Load_Spec_IPQC(sqlcon, "SpecList", txtItemCode.Text, "Roughness", DGV_SpecView);
                    //en_proc = myCode.Load_Spec_IPQC2(sqlcon, "SpecList", txtItemCode.Text, "Roughness", DGV_SpecView, format_folder, ".xlsm");
                    string format_file = Path.Combine(format_folder, cbType.SelectedItem.ToString());
                    string _itemcode = txtItemCode.Text;
                    List<string> format_lst = get_multiple_format(format_file, new List<string> { "*.xlsx", "*.xlsm" }, _itemcode);
                    if (format_lst.Count > 0)
                    {
                        DataTable dt =  myCode2.ACF_GetSpec_Process(format_lst[0]);
                        if(dt!=null)
                        {
                            DGV_SpecView.DataSource = dt;
                            if (DGV_SpecView.Rows.Count == 3)
                            {
                                DGV_SpecView.Rows[0].HeaderCell.Value = "SetVal";
                                DGV_SpecView.Rows[1].HeaderCell.Value = "USL";
                                DGV_SpecView.Rows[2].HeaderCell.Value = "LSL";
                                DGV_SpecView.AutoResizeColumns();
                                DGV_SpecView.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                                myCode.Disable_Sort_DGV(DGV_SpecView);
                            }
                            en_proc = true;
                        }
                        else
                        {
                            en_proc = false;
                            MessageBox.Show(new Form { TopMost = true }, "Không tìm thấy sheet ACF trong format", "Thông báo");
                        }
                    }
                    else
                    {
                        en_proc = false;
                        MessageBox.Show(new Form { TopMost = true }, "Không tìm thấy file format", "Thông báo");
                    }
                    break;
                case "Mitutoyo":
                    en_proc = myCode.Load_Spec(sqlcon, format_folder, txtItemCode.Text, txtLotNo.Text, ".xlsm", DGV_SpecView, "OMM", cbType.SelectedItem.ToString());
                    break;
                case "Nikon":
                    en_proc = myCode.Load_Spec(sqlcon, format_folder, txtItemCode.Text, txtLotNo.Text, ".xlsm", DGV_SpecView, "OMM", cbType.SelectedItem.ToString());
                    break;
                case "Keyence":
                    en_proc = myCode.Load_Spec(sqlcon, format_folder, txtItemCode.Text, txtLotNo.Text, ".xlsm", DGV_SpecView, "OMM", cbType.SelectedItem.ToString());
                    break;
                default:
                    if (cbMachine.Text !="")
                    {
                        en_proc = myCode.Load_Spec(sqlcon, format_folder, txtItemCode.Text, txtLotNo.Text, ".xlsm", DGV_SpecView, cbMachine.Text);
                    }
                    else
                    {
                       MessageBox.Show(new Form { TopMost = true },"Please, select Machine type!","Warning");
                    }
                    break;
            }
            if (en_proc)
            {
                GBControl.Enabled = true;
               MessageBox.Show(new Form { TopMost = true },"Hệ thống sẵn sàng", "Thông báo");
            }
            else
            {
                MessageBox.Show("Dữ liệu Spec chưa được cài đặt hoặc Sai thiết bị", "Thông báo");
                GBControl.Enabled = false;
            }

            /******************************************** Finish Main Program *****************************************************************/


            /***************************************** Testing Areas *********************************************************************************/

            //string test_file = @"D:\Customer Projects\SEEV\TestAreas\Raw_data\Mitutoyo.CSV";
            //DGV_SpecView.DataSource = myCode.Mitutoyo_Get_raw_data2(test_file);
            //myCode.Load_FAI_DGV2(DGV_DataView, sqlcon, "FAI_AUTO", "FAI_No", "FAI_Data", txtItemCode.Text, txtLotNo.Text);
            //myCode.Load_CPK_DGV(myCode.DGV_To_Table(DGV_DataView), DGV_CPK, myCode.testFAI_spec);

            /***************************************** Finish Testing Areas **************************************************************************/


        }
       
        private void btnLogin_Click(object sender, EventArgs e)
        {
            if (btnLogin.Text == "Logout")
            {
                btnLogin.Text = "Login";
                TDMK_OK2SHIP.admin_mode = false;
                numQty.Enabled = false;
                btnLogin.BackColor = curr_color;
                txtOperator.Text = TDMK_OK2SHIP.curr_user;
            }
            else
            {
                TDMK_OK2SHIP.curr_user = txtOperator.Text;
                FrmLogin frm_login = new FrmLogin();
                frm_login.ShowDialog();
            }
        }

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
                DGV_Data_Plus.Columns[sel_col_inx].HeaderText = src_col_name;
                DGV_Data_Plus.Columns[sel_col_inx].Name = src_col_name;
                double SetVal = Convert.ToDouble(DGV_SpecView.Rows[0].Cells[src_col_name].Value);
                string _UL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[1].Cells[src_col_name].Value));
                string _LL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[2].Cells[src_col_name].Value));
                string UL = "";
                string LL = "";
                if (_UL != "")
                {
                    UL = (SetVal + Convert.ToDouble(_UL)).ToString();
                }
                if (_LL != "")
                {
                    LL = (SetVal - Convert.ToDouble(_LL)).ToString();
                }


                foreach (DataGridViewRow r in DGV_Data_Plus.Rows)
                {
                    string act_val = checkDBNull(r.Cells[src_col_name].Value);
                    r.Cells[src_col_name].Style.BackColor = myCode.check_in_limit_Color(UL.ToString(), LL.ToString(), act_val, SetVal.ToString());
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
            if (!myCode.check_columns_existed(myCode.DGV_To_Table(DGV_Data_Plus), default_name))
            {
                DGV_Data_Plus.Columns[sel_col_inx].HeaderText = default_name;
                DGV_Data_Plus.Columns[sel_col_inx].Name = default_name;
            }
            else
            {
               MessageBox.Show(new Form { TopMost = true },"Cannot rename");
            }
        }

        private void DGV_Data_Plus_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            List<string> col_list = new List<string>();
            foreach (DataGridViewCell c in DGV_Data_Plus.SelectedCells)
            {
                int k = c.ColumnIndex;
                if (!col_list.Contains(DGV_Data_Plus.Columns[k].Name))
                {
                    col_list.Add(DGV_Data_Plus.Columns[k].Name);
                }

            }

            DataTable test_tbl = myCode.DGV_To_Table(DGV_Data_Plus);
            int arr_num = col_list.Count;
            TDMK_OK2SHIP.Calculate_CPK[] myCalc_CPK = new TDMK_OK2SHIP.Calculate_CPK[arr_num];
            TDMK_OK2SHIP.FAI_Histogram_Data[] myHistogram_data = new TDMK_OK2SHIP.FAI_Histogram_Data[arr_num];
            int col_list_inx = 0;
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
                            data_arr[inx] = Convert.ToDouble(c);
                            inx++;
                        }
                        myCode.Calcul_CPK_FAI2(ref_spec, Calcu_process.ConvertToDouble(temp_FAI_data.ToList()).ToArray(), ref myCalc_CPK[col_list_inx], ref myHistogram_data[col_list_inx]);
                        //DGV_SpecView.Rows[4].Cells[myCalc_CPK[col_list_inx].FAI_No].Value = Math.Round(myCalc_CPK[col_list_inx].stdev, 4).ToString();
                        //DGV_SpecView.Rows[5].Cells[myCalc_CPK[col_list_inx].FAI_No].Value = Math.Round(myCalc_CPK[col_list_inx].mean, 4).ToString();
                        //DGV_SpecView.Rows[6].Cells[myCalc_CPK[col_list_inx].FAI_No].Value = Math.Round(myCalc_CPK[col_list_inx].max, 3).ToString();
                        //DGV_SpecView.Rows[7].Cells[myCalc_CPK[col_list_inx].FAI_No].Value = Math.Round(myCalc_CPK[col_list_inx].min, 3).ToString();
                        //DGV_SpecView.Rows[8].Cells[myCalc_CPK[col_list_inx].FAI_No].Value = Math.Round(myCalc_CPK[col_list_inx].CP, 3).ToString();
                        //DGV_SpecView.Rows[9].Cells[myCalc_CPK[col_list_inx].FAI_No].Value = Math.Round(myCalc_CPK[col_list_inx].CPKL, 3).ToString();
                        //DGV_SpecView.Rows[10].Cells[myCalc_CPK[col_list_inx].FAI_No].Value = Math.Round(myCalc_CPK[col_list_inx].CPKU, 3).ToString();
                        //DGV_SpecView.Rows[11].Cells[myCalc_CPK[col_list_inx].FAI_No].Value = Math.Round(myCalc_CPK[col_list_inx].CPK, 3).ToString();
                        //DGV_SpecView.Rows[12].Cells[myCalc_CPK[col_list_inx].FAI_No].Value = Math.Round(myCalc_CPK[col_list_inx].CPKM, 3).ToString();
                        FrmHistogram chart_form = new FrmHistogram();
                        chart_form.src_bin_data = myHistogram_data[col_list_inx]._Bin_data;// bin_data;
                        chart_form.src_freq_bin_data = myHistogram_data[col_list_inx]._Freq_bin_data;//freq_bin_data;
                        chart_form.src_FAI_Data = myHistogram_data[col_list_inx]._FAI_Data;
                        chart_form.FAI_Spec_val = ref_spec;
                        chart_form.src_CPK_result = myCalc_CPK[col_list_inx];
                        chart_form.src_modified_NormDist = myHistogram_data[col_list_inx]._Modified_NormDist_data; //modified_NormDist;
                        chart_form.Show();
                        break;
                    }
                }
                col_list_inx++;
            }
        }
        public void Scan_data(object src_log)
        {
            Thread.Sleep(1000);
            string itemcode_folder="";
            string done_file="";
            string logfile = src_log.ToString();
            FileInfo f_info = new FileInfo(src_log.ToString());
            string f_exten = f_info.Extension;
            string save_time = Create_date_string(DateTime.Now);
            string f_name = Path.GetFileName(src_log.ToString());
            if (f_exten==f_ext)
            {
                itemcode_folder = Path.Combine(log_folder, "Scan_Finished", txtItemCode.Text, txtLotNo.Text);
                done_file = Path.Combine(itemcode_folder, "done_" + save_time + f_name);
                try
                {
                    switch (Dev)
                    {
                        case "Mitutoyo":
                            /******************************** Replace on 25Jun23 *****************************************************************************/
                            //Invoke(new Data_process(myCode.Mitutoyo_Data_display), TDMK_OK2SHIP.scan_file, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode);
                            /*************************************************** *****************************************************************************/
                            Invoke(new Data_process(myCode.Mitutoyo_process), logfile, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode);
                            break;
                        case "Nikon":
                            Invoke(new Data_process(myCode.Nikon_Data_display3), logfile, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode);
                            break;
                        case "Keyence":
                            Invoke(new Data_process(myCode.Keyence_Display_Data3), logfile, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode);
                            break;
                        case "AU_NI_Thickness":
                            Invoke(new Data_process2(myCode.AU_IN_Scan), logfile, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode, Convert.ToInt32(numQty.Value));
                            break;
                        case "Roughness":
                            //Invoke(new Data_process2(myCode.Roughness_Scan), TDMK_OK2SHIP.scan_file, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode, Convert.ToInt32(numQty.Value));
                            //Invoke(new Roughness_process(myCode.Roughness_Scan2), TDMK_OK2SHIP.scan_file, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode, Convert.ToInt32(numQty.Value),Convert.ToInt32(numMeasLoc.Value));
                            Invoke(new Roughness_process(myCode.Roughness_Scan_ACF_SMT), logfile, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode, Convert.ToInt32(numQty.Value), Convert.ToInt32(numMeasLoc.Value));
                            break;
                    }
                }
                catch
                {
                    MessageBox.Show(new Form { TopMost = true }, "Invalid format input file", "Warning");
                }
            }
            else
            {
                itemcode_folder = Path.Combine(log_folder, "Ignored", txtItemCode.Text, txtLotNo.Text);
                done_file = Path.Combine(itemcode_folder, "Ignored_" + save_time + f_name);
            }
            try
            {
                if (!Directory.Exists(itemcode_folder))
                {
                    Directory.CreateDirectory(itemcode_folder);
                }
                File.Copy(logfile, done_file, true);
                File.Delete(logfile);
            }
            catch
            {

            }
            //pro_en = false;
            Invoke(new SetText(SetTextButton), btnRun, "Start");
            //MessageBox.Show("Hoàn thành lấy dữ liệu", "Thông tin");
        }

        private void txtDataFolder_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            FolderBrowserDialog folder_data = new FolderBrowserDialog();
            if (folder_data.ShowDialog() == DialogResult.OK)
            {
                txtDataFolder.Text = folder_data.SelectedPath;
            }
        }

        private void txtItemCode_KeyDown(object sender, KeyEventArgs e)
        {
            if(Dev=="Manual")
            {
                if (e.KeyCode == Keys.Enter)
                {
                    if (txtItemCode.Text != "")
                    {
                        cbMachine.DataSource = TDMK_Code.Load_Item_Names_Filter(sqlcon, "FAI_Spec", "Instrument", "ItemCode = '" + txtItemCode.Text + "' and Instrument not like 'OMM'");
                    }
                }
            }
        }

        public bool Load_Spec_IPQC(SqlConnection tar_sqlcon, string spec_tbl_name, string tar_ItemCode, string tar_itemcheck, DataGridView tar_DGV_Spec)
        {
            bool _result = false;
            DataTable Spec_tbl = TDMK_Code.Datatable_Filter(sqlcon, spec_tbl_name, TDMK_Code.filter_str(new string[] { "ItemCode", "Item", "Remark" }, new string[] { tar_ItemCode, "%" + tar_itemcheck + "%","%AU_Plating%" })); //Roughness AU_NI            
            if (Spec_tbl.Rows.Count == 0)
            {
               MessageBox.Show(new Form { TopMost = true },"No Spec data in database");
                _result = false;
            }
            else
            {
                string[] items_list = Spec_tbl.AsEnumerable().Select(r => r.Field<string>("Item")).ToArray();
                string[] limit_lst = Spec_tbl.AsEnumerable().Select(r => r.Field<string>("Spec")).ToArray();
                char[] split_chr = { '~', '<', '\u2264', '\u2265', '\u00B1','>' };
                char[] trim_chr = { 'u', 'm', '\u03BC', '\u00B5', ' ' ,'='};
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
                _result = true;
            }
            return _result;
        }
        public void Save_data_AU_NI()
        {
            bool data_rec_en = true;
            string[] items = new string[11];
            string[] item_vals = new string[11];
            items[0] = "ID";
            items[1] = "ItemCode";
            items[2] = "LotNo";
            items[3] = "Operator";
            items[4] = "Machine";
            items[5] = "MDate";
            items[6] = "L1_AU_THICKNESS";
            items[7] = "L2_AU_THICKNESS";
            items[8] = "L1_NI_THICKNESS";
            items[9] = "L2_NI_THICKNESS";
            item_vals[1] = txtItemCode.Text;//"ItemCode";
            item_vals[2] = txtLotNo.Text;//"LotNo";
            item_vals[3] = txtOperator.Text;//"Operator";
            item_vals[4] = cbMachine.Text;//"Machine";
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
                    for (int j = 0; j < DGV_DataView.RowCount; j++)
                    {
                        for (int i = 0; i < DGV_DataView.Columns.Count; i++)
                        {
                            item_vals[6 + i] = checkDBNull(DGV_DataView.Rows[j].Cells[i].Value); // "L1_AU";
                        }
                        item_vals[0] = (TDMK_Code.SQL_MAX("AU_NI", "ID", sqlcon) + 1).ToString();//"ID";
                        TDMK_Code.insert_val_arr("AU_NI", sqlcon, items, item_vals);
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
        public void Save_data_Roughness()
        {
            bool data_rec_en = true;
            string[] items = new string[15];
            string[] item_vals = new string[15];
            items[0] = "ID";
            items[1] = "ItemCode";
            items[2] = "LotNo";
            items[3] = "Operator";
            items[4] = "Machine";
            items[5] = "MDate";
            items[6] = "L1_Roughness_Sa";
            items[7] = "L1_Roughness_Sq";
            items[8] = "L2_Roughness_Sa";
            items[9] = "L2_Roughness_Sq";

            items[10] = "L1_Roughness_Sdr";
            items[11] = "L2_Roughness_Sdr";

            items[12] = "L3_Roughness_Sa";
            items[13] = "L3_Roughness_Sq";
            items[14] = "L3_Roughness_Sdr";

            item_vals[1] = txtItemCode.Text;//"ItemCode";
            item_vals[2] = txtLotNo.Text;//"LotNo";
            item_vals[3] = txtOperator.Text;//"Operator";
            item_vals[4] = cbMachine.Text;//"Machine";
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
                    for (int j = 0; j < DGV_DataView.RowCount; j++)
                    {
                        for (int i = 0; i < DGV_DataView.Columns.Count; i++)
                        {
                            item_vals[6 + i] = checkDBNull(DGV_DataView.Rows[j].Cells[i].Value); // "L1_AU";
                        }
                        item_vals[0] = (TDMK_Code.SQL_MAX("Roughness", "ID", sqlcon) + 1).ToString();//"ID";
                        TDMK_Code.insert_val_arr("Roughness", sqlcon, items, item_vals);
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

        private void tsmRemove_Click(object sender, EventArgs e)
        {
            if(DGV_Data_Plus.SelectedRows.Count>0)
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
            if (!TDMK_OK2SHIP.confirm_request)
            {
                if (myCode.insert_FAI_data_DGV2(DGV_Data_Plus, DGV_DataView, myCode.DGV_To_Table(DGV_SpecView)))
                {
                    myCode.check_FAIdata_inSpec2(DGV_DataView, DGV_SpecView);
                    myCode.Load_CPK_DGV_Histogram(DGV_DataView, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
                    if (TDMK_OK2SHIP.confirm_request)
                    {
                        GBControl.Enabled = false;
                    }
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },"Need confirmation from authorized PIC");
            }            
        }

        private void DGV_DataView_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            
                            CopyToClipboard(DGV_DataView);
                            break;

                        case Keys.V:                            
                            PasteClipboardValue(true, DGV_DataView);
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(new Form { TopMost = true },"Copy/paste operation failed. " + ex.Message, "Copy/Paste", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void CopyToClipboard(DataGridView tar_DGV)
        {
            //Copy to clipboard
            DataObject dataObj = tar_DGV.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }
        private void PasteClipboardValue(bool _transpose, DataGridView tar_DGV)
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
            if (cbValue.Count>0)
            {
                if (!_transpose)
                {
                    int iRowIndex = startCell.RowIndex;
                    foreach (int rowKey in cbValue.Keys)
                    {
                        int iColIndex = startCell.ColumnIndex;
                        foreach (int cellKey in cbValue[rowKey].Keys)
                        {
                            if (iColIndex <= tar_DGV.Columns.Count - 1 && iRowIndex <= tar_DGV.Rows.Count - 1)
                            {
                                DataGridViewCell cell = tar_DGV[iColIndex, iRowIndex];
                                if (cell.Selected)
                                {
                                    cell.Value = cbValue[rowKey][cellKey];
                                    if (checkDBNull(cell.Value) != "")
                                    {
                                        string curr_col_name = tar_DGV.Columns[iColIndex].Name.ToString();
                                        if (myCode.check_columns_existed(myCode.DGV_To_Table(DGV_SpecView), curr_col_name))
                                        {
                                            string SetVal = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[0].Cells[curr_col_name].Value));
                                            string UL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[1].Cells[curr_col_name].Value));
                                            string LL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[2].Cells[curr_col_name].Value));
                                            string act_val = checkDBNull(cell.Value);
                                            cell.Style.BackColor = myCode.check_in_limit_Color(UL, LL, act_val, SetVal.ToString());
                                        }
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
                            if (iColIndex <= tar_DGV.Columns.Count - 1 && iRowIndex <= tar_DGV.Rows.Count - 1)
                            {
                                DataGridViewCell cell = tar_DGV[iColIndex, iRowIndex];
                                if (cell.Selected)
                                {
                                    cell.Value = cbValue[rowKey][cellKey];
                                    if (checkDBNull(cell.Value) != "")
                                    {
                                        string curr_col_name = tar_DGV.Columns[iColIndex].Name.ToString();
                                        if (myCode.check_columns_existed(myCode.DGV_To_Table(DGV_SpecView), curr_col_name))
                                        {

                                            string SetVal = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[0].Cells[curr_col_name].Value));
                                            string UL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[1].Cells[curr_col_name].Value));
                                            string LL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[2].Cells[curr_col_name].Value));

                                            string act_val = checkDBNull(cell.Value);
                                            cell.Style.BackColor = myCode.check_in_limit_Color(UL, LL, act_val, SetVal);
                                        }
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
        private DataGridViewCell GetStartCell(DataGridView dgView)
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
        private Dictionary<int, Dictionary<int, string>> ClipBoardValues(string clipboardValue)
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

        private void DGV_DataView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
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
                        if (myCode.check_columns_existed(myCode.DGV_To_Table(DGV_SpecView), curr_col_name))
                        {
                            //string UL = checkDBNull(DGV_SpecView.Rows[1].Cells[curr_col_name].Value);
                            //string LL = checkDBNull(DGV_SpecView.Rows[2].Cells[curr_col_name].Value);
                            //
                            //double _UL = Convert.ToDouble(SV) + Convert.ToDouble(UL);
                            //double _LL = Convert.ToDouble(SV) - Convert.ToDouble(LL);

                            string SV = checkDBNull(DGV_SpecView.Rows[0].Cells[curr_col_name].Value);
                            double SetVal = 0;
                            if(myCode.IsNumeric(SV))
                            {
                                SetVal = Convert.ToDouble(SV);
                            }    
                            //Convert.ToDouble(DGV_SpecView.Rows[0].Cells[curr_col_name].Value);
                            string _UL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[1].Cells[curr_col_name].Value));
                            string _LL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[2].Cells[curr_col_name].Value));
                            string UL = "";
                            string LL = "";
                            if (_UL != "")
                            {
                                UL = (SetVal + Convert.ToDouble(_UL)).ToString();
                            }
                            if (_LL != "")
                            {
                                LL = (SetVal - Convert.ToDouble(_LL)).ToString();
                            }

                            string act_val = checkDBNull(DGV_DataView.CurrentCell.Value);
                            curr_cell.Style.BackColor = myCode.check_in_limit_Color(_UL.ToString(), _LL.ToString(), act_val,SetVal.ToString());

                            if (myCode.check_ColumnData_enough(DGV_DataView, Convert.ToInt32(numQty.Value), curr_col_name))
                            {
                                myCode.Load_CPK_DGV_ColName(myCode.DGV_To_Table(DGV_DataView), DGV_CPK, myCode.testFAI_spec.ToArray());
                            }
                        }
                    }
                }

            }
        }

        private void selectedColumnsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PasteClipboardValue(true,DGV_DataView);
        }

        private void tsmCopy_Click(object sender, EventArgs e)
        {
            CopyToClipboard(DGV_Data_Plus);
        }

        private void selectedRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PasteClipboardValue(false, DGV_DataView);
        }

        private void selectedRowsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PasteClipboardValue(false,DGV_Data_Plus);
        }

        private void selectedColumnsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            PasteClipboardValue(true, DGV_Data_Plus);
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyToClipboard(DGV_Data_Plus);
        }

        private void calculateCPKToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                // myCode.Load_CPK_DGV(myCode.DGV_To_Table(DGV_DataView), DGV_CPK, myCode.testFAI_spec.ToArray());
                myCode.Load_CPK_DGV_Histogram(DGV_DataView, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
            }
            catch
            {
                MessageBox.Show(new Form { TopMost = true },"Not enough data", "Warning");
            }
            
        }
        public void Load_Test_FAI_Data(string format_loc, string tar_ItemCode, string tar_LotNo, string file_extension, DataGridView tar_DGV_Data)
        {
            string tar_format = tar_ItemCode + tar_LotNo;
            string tar_format_file = Path.Combine(format_loc, tar_format + file_extension);
            if (File.Exists(tar_format_file))
            {
                myExcel.Workbook tar_wkbook = TDMK_Code.open_excel_file(tar_format_file, "", "");
                int wrksheet_num = tar_wkbook.Worksheets.Count;
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
                            if (!myCode.check_columns_existed(myCode.DGV_To_Table(tar_DGV_Data), col_dgv_name))
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
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },"File format khong tim thay");
            }
        }

        private void DGV_Data_Plus_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            //if ((Dev == "Manual") && (EditMode))
            //{
            //    edit_en = true;
            //}
            //else
            //{
            //    edit_en = false;
            //}
        }

        private void DGV_Data_Plus_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (edit_en)
            {
                DataGridViewCell curr_cell;
                string curr_col_name;
                curr_cell = DGV_Data_Plus.CurrentCell;
                curr_col_name = DGV_Data_Plus.Columns[curr_cell.ColumnIndex].Name.ToString();
                if (checkDBNull(curr_cell.Value) != "")
                {
                    if (!IsNumeric(checkDBNull(curr_cell.Value)))
                    {
                        MessageBox.Show(new Form { TopMost = true },"You have to enter a number!", "Warning");
                        DGV_Data_Plus.CurrentCell.Value = "";
                    }
                    else
                    {
                        if (myCode.check_columns_existed(myCode.DGV_To_Table(DGV_SpecView), curr_col_name))
                        {
                            //string UL = checkDBNull(DGV_SpecView.Rows[1].Cells[curr_col_name].Value);
                            //string LL = checkDBNull(DGV_SpecView.Rows[2].Cells[curr_col_name].Value);
                            //string SV = checkDBNull(DGV_SpecView.Rows[0].Cells[curr_col_name].Value);
                            //double _UL = Convert.ToDouble(SV) + Convert.ToDouble(UL);
                            //double _LL = Convert.ToDouble(SV) - Convert.ToDouble(LL);

                            double SetVal = Convert.ToDouble(DGV_SpecView.Rows[0].Cells[curr_col_name].Value);
                            string _UL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[1].Cells[curr_col_name].Value));
                            string _LL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[2].Cells[curr_col_name].Value));
                            string UL = "";
                            string LL = "";
                            if (_UL != "")
                            {
                                UL = (SetVal + Convert.ToDouble(_UL)).ToString();
                            }
                            if (_LL != "")
                            {
                                LL = (SetVal - Convert.ToDouble(_LL)).ToString();
                            }


                            string act_val = checkDBNull(DGV_Data_Plus.CurrentCell.Value);

                            //curr_cell.Style.BackColor = myCode.check_in_limit2(_UL.ToString(), _LL.ToString(), act_val);
                            curr_cell.Style.BackColor = myCode.check_in_limit_Color(UL, LL, act_val, SetVal.ToString());
                        }
                    }
                }

            }
        }

        private void DGV_Data_Plus_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if ((Dev == "Manual") && (EditMode))
            {
                //edit_en = true;
                DataGridViewCell curr_cell;
                string curr_col_name;
                curr_cell = DGV_Data_Plus.CurrentCell;
                curr_col_name = DGV_Data_Plus.Columns[curr_cell.ColumnIndex].Name.ToString();
                if (checkDBNull(curr_cell.Value) != "")
                {
                    if (!IsNumeric(checkDBNull(curr_cell.Value)))
                    {
                        MessageBox.Show(new Form { TopMost = true },"You have to enter a number!", "Warning");
                        DGV_Data_Plus.CurrentCell.Value = "";
                    }
                    else
                    {
                        if (myCode.check_columns_existed(myCode.DGV_To_Table(DGV_SpecView), curr_col_name))
                        {
                            //string UL = checkDBNull(DGV_SpecView.Rows[1].Cells[curr_col_name].Value);
                            //string LL = checkDBNull(DGV_SpecView.Rows[2].Cells[curr_col_name].Value);
                            //string SV = checkDBNull(DGV_SpecView.Rows[0].Cells[curr_col_name].Value);
                            //double _UL = Convert.ToDouble(SV) + Convert.ToDouble(UL);
                            //double _LL = Convert.ToDouble(SV) - Convert.ToDouble(LL);



                            double SetVal = Convert.ToDouble(DGV_SpecView.Rows[0].Cells[curr_col_name].Value);
                            string _UL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[1].Cells[curr_col_name].Value));
                            string _LL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[2].Cells[curr_col_name].Value));
                            string UL = "";
                            string LL = "";
                            if (_UL != "")
                            {
                                UL = (SetVal + Convert.ToDouble(_UL)).ToString();
                            }
                            if (_LL != "")
                            {
                                LL = (SetVal - Convert.ToDouble(_LL)).ToString();
                            }


                            string act_val = checkDBNull(DGV_Data_Plus.CurrentCell.Value);
                            //curr_cell.Style.BackColor = myCode.check_in_limit2(_UL.ToString(), _LL.ToString(), act_val);
                            curr_cell.Style.BackColor = myCode.check_in_limit_Color(UL, LL, act_val, SetVal.ToString());                           
                        }
                    }
                }
            }
            else
            {
                edit_en = false;
            }
}

        private void DGV_DataView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //if ((Dev=="Manual")||(edit_en))
            //{

            //}
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
                    if (myCode.check_columns_existed(myCode.DGV_To_Table(DGV_SpecView), curr_col_name))
                    {
                        string SetVal = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[0].Cells[curr_col_name].Value));
                        //double SetVal = Convert.ToDouble(DGV_SpecView.Rows[0].Cells[curr_col_name].Value);
                        string _UL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[1].Cells[curr_col_name].Value));
                        string _LL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[2].Cells[curr_col_name].Value));
                        //string UL = "";
                        //string LL = "";
                        //if (_UL != "")
                        //{
                        //    UL = (SetVal + Convert.ToDouble(_UL)).ToString();
                        //}
                        //if (_LL != "")
                        //{
                        //    LL = (SetVal - Convert.ToDouble(_LL)).ToString();
                        //}

                        string act_val = checkDBNull(DGV_DataView.CurrentCell.Value);
                        curr_cell.Style.BackColor = myCode.check_in_limit_Color(_UL, _LL, act_val, SetVal);
                        /*if (myCode.check_ColumnData_enough(DGV_DataView, Convert.ToInt32(numQty.Value), curr_col_name))
                        {
                            myCode.Load_CPK_DGV_ColName(myCode.DGV_To_Table(DGV_DataView), DGV_CPK, myCode.testFAI_spec);
                        }*/
                    }
                }
            }
        }
        public void Alarm_FAI_NG_Saved_Encode_Email()
        {
            string list_FAI_CPK = "";
            for (int i = 0; i < DGV_CPK.Columns.Count; i++)
            {
                var val = DGV_CPK.Rows[7].Cells[i].Value;
                if (val != null)
                {
                    double value = Convert.ToDouble(val);
                    if (value < 1.67)
                    {
                        string FAI_Col = DGV_CPK.Columns[i].HeaderText;
                        list_FAI_NG.Add(FAI_Col);
                        if (FAI_Col.Contains("FAI") == false)
                        {
                            FAI_Col = "FAI-" + FAI_Col;
                        }
                        list_FAI_CPK = list_FAI_CPK + ", " + FAI_Col.Split('_')[0].Trim() + " _ CPK = " + val.ToString();
                    }
                }
            }
            if (list_FAI_NG.Count > 0)
            {
                int count = list_FAI_NG.Count;
                string file_addr = Path.Combine(Application.StartupPath, "Format", "FAI.xlsm");
                string tepm_file = Path.Combine(Application.StartupPath, "Format", "test.xlsx");
                if (File.Exists(tepm_file))
                {
                    File.Delete(tepm_file);
                }
                CreateExcelFile(tepm_file);
                var workbook = TDMK_Code.open_excel_file(tepm_file, "", "");
                myExcel.Worksheet ws = workbook.Worksheets[1];
                if (ws.ProtectContents)
                {
                    ws.Unprotect("Histogram_123");
                }
                for (int i = 0; i < count; i++)
                {
                    double dim_no = Convert.ToDouble(list_FAI_NG[i].Split('_')[0]); // ví dụ 4,5,7,9
                    double normal_dim = Convert.ToDouble(DGV_SpecView.Rows[0].Cells[list_FAI_NG[i]].Value.ToString());
                    double max = Convert.ToDouble(DGV_SpecView.Rows[1].Cells[list_FAI_NG[i]].Value.ToString());
                    double min = Convert.ToDouble(DGV_SpecView.Rows[2].Cells[list_FAI_NG[i]].Value.ToString());
                    ws.Cells[14, i + 3].Value = dim_no;
                    ws.Cells[15, i + 3].Value = normal_dim;
                    ws.Cells[16, i + 3].Value = max;
                    ws.Cells[17, i + 3].Value = min;
                    List<string> list_val = new List<string>();
                    for (int J = 0; J < DGV_DataView.Rows.Count; J++)
                    {
                        double x = Convert.ToDouble(DGV_DataView.Rows[J].Cells[list_FAI_NG[i]].Value.ToString());
                        ws.Cells[255 + J, i + 3].Value = x;//Math.Round(x,3);
                    }
                }
                workbook.Save();
                string new_file_format = Path.Combine(Application.StartupPath, "NG_FAI_Data", txtItemCode.Text + "_FAI_NG.xlsm");
                if (File.Exists(new_file_format))
                {
                    File.Delete(new_file_format);
                }
                File.Copy(file_addr, new_file_format);
                myExcel.Workbook dest_wrk = TDMK_Code.open_excel_file(new_file_format, "", "");
                CopyPasteExcelData(workbook, dest_wrk, "C14:I17", 14, 3);
                CopyPasteExcelData(workbook, dest_wrk, "C255:I286", 255, 3);
                workbook.Close();
                dest_wrk.Close();
                File.Delete(tepm_file);
                string id = (TDMK_Code.SQL_MAX("Alarm_Excel_File", "ID", sqlcon) + 1).ToString();
                string[] item = new string[] { "ID", "ItemCode", "LotNo", "FAI_CPK_NG", "File_Address", "Encode_File", "State" };
                object[] item_val = new object[] { id, (txtItemCode.Text != "") ? txtItemCode.Text : "null", (txtLotNo.Text != "") ? txtLotNo.Text : "null", list_FAI_CPK.Remove(0, 1), new_file_format, File_To_Binary(new_file_format), "null" };
                Insert_SQL_Data("Alarm_Excel_File", item, item_val, sqlcon);
            }
        }
        public void CreateExcelFile(string filePath)
        {
            // Tạo một workbook mới
            var workbook = TDMK_Code.Create_workbook();

            // Thêm một bảng tính mới vào workbook
            //var worksheet = workbook.Worksheets.Add("Sheet1");

            // Lưu workbook thành tệp Excel
            workbook.SaveAs(filePath);

            // Giải phóng tài nguyên
            workbook.Close();
        }
        public void CopyPasteExcelData(myExcel.Workbook src_wrk, myExcel.Workbook dest_wrk, string source_range, int des_start_row, int des_start_col)
        {
            //myExcel.Workbook sourceWorkbook = TDMK_Code.open_excel_file(sourceFilePath, "", "");

            myExcel.Worksheet sourceWorksheet = src_wrk.Worksheets[1]; // Chọn sheet nguồn, ở đây là sheet đầu tiên

            //myExcel.Workbook destinationWorkbook = TDMK_Code.open_excel_file(destinationFilePath, "", "");

            myExcel.Worksheet destinationWorksheet = dest_wrk.Worksheets[1]; // Chọn sheet đích, ở đây là sheet đầu tiên                   
            if (destinationWorksheet.ProtectContents)
            {
                destinationWorksheet.Unprotect("Histogram_123");
            }
            // Lựa chọn vùng dữ liệu nguồn từ dòng 5 đến dòng 10 và cột 2 đến cột 9
            myExcel.Range sourceRange = sourceWorksheet.Range[source_range];

            // Lựa chọn vùng đích để paste dữ liệu
            myExcel.Range destinationRange = destinationWorksheet.Cells[des_start_row, des_start_col]; // Ví dụ: Paste từ ô A1

            // Copy và paste vùng dữ liệu
            sourceRange.Copy();
            destinationRange.PasteSpecial(myExcel.XlPasteType.xlPasteValues, myExcel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            // Lưu tệp đích
            dest_wrk.Save();
            //sourceWorkbook.Close();
            //destinationWorkbook.Close();
        }
        public void Insert_SQL_Data(string tableName, string[] columnNames, object[] values, SqlConnection sqlcon)
        {
            if (IsConnectionOpen(sqlcon) == false)
            {
                sqlcon.Open();
            }
            // Tạo câu lệnh SQL INSERT
            string sql = "INSERT INTO " + tableName + " (";
            for (int i = 0; i < columnNames.Length; i++)
            {
                sql += columnNames[i];
                if (i < columnNames.Length - 1)
                {
                    sql += ", ";
                }
            }
            sql += ") VALUES (";
            for (int i = 0; i < columnNames.Length; i++)
            {
                sql += "@" + columnNames[i];
                if (i < columnNames.Length - 1)
                {
                    sql += ", ";
                }
            }
            sql += ")";

            // Tạo đối tượng SqlCommand và đặt tham số giá trị vào câu lệnh SQL
            using (SqlCommand command = new SqlCommand(sql, sqlcon))
            {
                for (int i = 0; i < columnNames.Length; i++)
                {
                    if (values[i] is byte[])
                    {
                        command.Parameters.AddWithValue("@" + columnNames[i], values[i]);
                    }
                    else
                    {
                        command.Parameters.AddWithValue("@" + columnNames[i], values[i].ToString());
                    }
                }

                // Thực thi câu lệnh SQL để chèn giá trị vào bảng
                command.ExecuteNonQuery();
            }


        }
        public static bool IsConnectionOpen(SqlConnection connection)
        {
            return connection.State == System.Data.ConnectionState.Open;
        }
        public byte[] File_To_Binary(string filePath)
        {
            return File.ReadAllBytes(filePath);
        }

        private void DGV_DataView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex != -1 && e.RowIndex != -1)
            {
                try
                {
                    if(e.Value!=null)
                    {
                        double d = double.Parse(e.Value.ToString());
                        e.Value = d.ToString("N3");
                    }
                }
                catch
                {

                }
            }
        }

        private void DGV_Data_Plus_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex != -1 && e.RowIndex != -1)
            {
                try
                {
                    double d = double.Parse(e.Value.ToString());
                    e.Value = d.ToString("N3");
                }
                catch
                {

                }
            }
        }
        public void Export_To_FAI2(myExcel.Workbook src_format_wrk, string _ItemCode, string _LotNo, myExcel.Workbook src_data_wrkbook)
        {
            DataTable FAI_Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { _ItemCode, _LotNo }));
            DataTable FAI_Spec_tbl = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Spec", "ItemCode = '" + _ItemCode + "'");
            foreach (myExcel.Worksheet sht in src_data_wrkbook.Worksheets)
            {
                if ((sht.Name.Contains("FAI")) || (sht.Name.Contains("SPC")))
                {
                    string sel_sht_name = sht.Name;
                    myExcel.Worksheet format_wrksheet = src_format_wrk.Sheets[sel_sht_name];
                    format_wrksheet.Activate();
                    myExcel.Range format_rgn = format_wrksheet.Range["C255"];
                    int col_num = sht.UsedRange.Columns.Count;
                    int row_num = sht.UsedRange.Rows.Count;
                    myExcel.Range data_rgn = sht.Range[sht.Cells[2, 1], sht.Cells[row_num, col_num]];
                    data_rgn.Copy();
                    format_rgn.PasteSpecial(myExcel.XlPasteType.xlPasteValues, myExcel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
                }
            }
            src_data_wrkbook.Close(false);
        }
        public void Export_FAI_Batch(DataTable FAI_Data_tbl, string tar_ItemCode, string tar_LotNo, myExcel.Workbook tar_wrkbook)
        {
            char[] trim_char = new char[] { ' ', '\r', '\n' };
            //DataTable FAI_Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { tar_ItemCode, tar_LotNo }));
            DataTable FAI_Spec_tbl = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Spec", "ItemCode = '" + tar_ItemCode + "'");
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
                            cur_rgn.Offset[r_inx + 1, c_inx].Value = _fai_val.Trim(trim_char);
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

        private void txtLotNo_Validated(object sender, EventArgs e)
        {
            txtLotNo.Text = myCode.Lotno_Formated(txtLotNo.Text);
        }

        private void cbType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(cbType.SelectedIndex==1)
            {
                numQty.Value = 6;
            }
            else
            {
                numQty.Value = 32;
            }
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
                        string[] temp_FAI_data = src_dt.AsEnumerable().Where(x => x.Field<string>(t) != null).Select(r => r.Field<string>(t)).ToArray();
                        if (temp_FAI_data.Length > 0)
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
                            myCode.Calcul_CPK_FAI2(ref_spec, Calcu_process.ConvertToDouble(temp_FAI_data.ToList()).ToArray(), ref myCalc_CPK[col_list_inx], ref myHistogram_data[col_list_inx]);
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

        private void btnLoadData_Click(object sender, EventArgs e)
        {
start_lbl: if (spData.Panel2.Controls.Count == 0)
            {
                FAI_Data frmFAI = new FAI_Data(txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbType_Sel.SelectedItem.ToString(),cbShift_Sel.SelectedItem.ToString()) { Dock = DockStyle.Fill, TopLevel = false, TopMost = true };
                frmFAI.Show();
                spData.Panel2.Controls.Add(frmFAI);
            }
            else
            {
                spData.Panel2.Controls.Clear();
                goto start_lbl;
            }
        }

        private void btnExport_Data_Click(object sender, EventArgs e)
        {

            /***************************************** Main Process **********************************************************************/
            Export_FAI_Data_byEPPLUS();
            /***************************************** End Main Process **********************************************************************/

        }
        public void Export_FAI_data_byExcel()
        {
            string format_file = Path.Combine(format_folder, cbType_Sel.SelectedItem.ToString());
            string _itemcode = txtItemCode_Sel.Text;
            string _lotno = txtLotNo_Sel.Text;
            string _shift = cbShift_Sel.SelectedItem.ToString();
            DataTable FAI_dt = myCode.Load_FAI_ToTable(myVar.sqlcon_SMT, _itemcode, _lotno, cbType_Sel.SelectedItem.ToString(), cbShift_Sel.SelectedItem.ToString());
            DataTable FAI_Spec = myCode.Load_FAI_Spec_ToTable(myVar.sqlcon_SMT, _itemcode, cbType_Sel.SelectedItem.ToString());
            bool export_en = false;
            if (myCode.check_FAIdata_inSpec(FAI_dt, FAI_Spec))
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
                List<string> format_lst = get_multiple_format(format_file, new List<string> { "*.xlsx", "*.xlsm" }, _itemcode);
                if (format_lst.Count > 0)
                {
                    string report_type = cbType_Sel.SelectedItem.ToString();
                    DataTable FAI_Data_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Remark", "Shift" }, new string[] { _itemcode, _lotno, cbType_Sel.SelectedItem.ToString(), cbShift_Sel.SelectedItem.ToString() }));
                    if (FAI_Data_tbl.Rows.Count > 0)
                    {
                        format_file = format_lst[0];
                        string export_file = Path.Combine(report_location, Path.GetFileNameWithoutExtension(format_file) + "-" + _lotno + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(format_file));
                        myExcel.Workbook report_saved = null;
                        if (System.IO.File.Exists(export_file))
                        {
                            report_saved = TDMK_Code.open_excel_file(export_file, "", "");
                        }
                        else
                        {
                            report_saved = TDMK_Code.open_excel_file(format_file, "", "");
                            report_saved.SaveAs(export_file);
                        }
                        Dictionary<string, DataTable> result_dt = myCode2.Export_FAI_Batch2(_itemcode, _lotno, cbType_Sel.SelectedItem.ToString(), _shift);
                        myExcel.Application xlsApp = TDMK_Code.StartExcel();
                        xlsApp.DisplayAlerts = false;
                        foreach (var dt in result_dt)
                        {
                            myExcel.Workbook exp_file = myCode2.Table_To_CSV(dt.Key, dt.Value);
                            if (exp_file != null)
                            {
                                myCode2.Export_To_FAI2(report_saved, _itemcode, _lotno, exp_file, cbType_Sel.SelectedItem.ToString());
                            }
                        }
                        xlsApp.DisplayAlerts = true;
                    }
                    else
                    {
                        MessageBox.Show("Không tìm thấy dữ liệu của ItemCode / Lotno / Shift : " + _itemcode + " / " + _lotno + " / " + _shift, "Thông báo");
                    }
                }
                else
                {
                    MessageBox.Show("Không tìm thấy Format", "Cảnh báo");
                }
            }
        }
        public void Export_FAI_Data_byEPPLUS()
        {
            string type = cbType_Sel.SelectedItem.ToString();
            //string format_loc = Path.Combine(format_folder, "MASS","FAI");
            format_folder = Path.Combine(find_config_path(Application.StartupPath,"SEEV Data"),"Format");
            report_location = Path.Combine(find_config_path(Application.StartupPath, "SEEV Data"), "Report");
            string format_loc = "";
            if (type == "NPI")
            {
                format_loc = Path.Combine(format_folder, "NPI");
            }
            else
            {
                format_loc = Path.Combine(format_folder, "MASS", "FAI");
            }
            
            string _itemcode = txtItemCode_Sel.Text;
            string _lotno = txtLotNo_Sel.Text;
            string _shift = cbShift_Sel.SelectedItem.ToString();
            DataTable FAI_dt = myCode.Load_FAI_ToTable(myVar.sqlcon_SMT, _itemcode, _lotno, cbType_Sel.SelectedItem.ToString(), cbShift_Sel.SelectedItem.ToString());
            DataTable FAI_Spec = myCode.Load_FAI_Spec_ToTable(myVar.sqlcon_SMT, _itemcode, cbType_Sel.SelectedItem.ToString());
            bool export_en = false;
            if (myCode.check_FAIdata_inSpec(FAI_dt, FAI_Spec))
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
                    string report_type = cbType_Sel.SelectedItem.ToString();
                    DataTable FAI_Data_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Remark", "Shift" }, new string[] { _itemcode, _lotno, cbType_Sel.SelectedItem.ToString(), cbShift_Sel.SelectedItem.ToString() }));
                    if (FAI_Data_tbl.Rows.Count > 0)
                    {
                        string format_file = format_lst[0];
                        string export_file = "";
                        if (type == "MASS")
                        {
                            string report_path = Path.Combine(report_location, "OMM_Dimension");
                            if (!Directory.Exists(report_path))
                            {
                                Directory.CreateDirectory(report_path);
                            }
                            export_file = Path.Combine(report_path, Path.GetFileNameWithoutExtension(format_file) + "-" + _lotno + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(format_file));

                        }
                        else
                        {
                            string report_path = Path.Combine(report_location, "NPI", "FAI");
                            if (!Directory.Exists(report_path))
                            {
                                Directory.CreateDirectory(report_path);
                            } 
                            export_file = Path.Combine(report_path, _itemcode + "-" + _lotno + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(format_file));
                        } 
                        ExcelPackage Report_Pack = new ExcelPackage();
                        ExcelWorkbook report_saved = null;
                        bool file_existed = false;
                        if (System.IO.File.Exists(export_file))
                        {
                            Report_Pack = Excel_Lib.open_excel(export_file);
                            file_existed = true;
                        }
                        else
                        {
                            Report_Pack = Excel_Lib.open_excel(format_file);
                            file_existed = false;
                            //Report_Pack.SaveAs(new FileInfo(export_file));
                        }
                        report_saved = Report_Pack.Workbook;
                        int error_count = 0;
                        while (report_saved.Names.Count > 0 && error_count < 5)
                        {
                            for (int i = 0; i < report_saved.Names.Count; i++)
                            {
                                try
                                {
                                    string cur_name = report_saved.Names[i].Name;
                                    report_saved.Names.Remove(cur_name);
                                }
                                catch
                                {
                                    error_count++;
                                    continue;
                                }
                            }
                        }
                        try
                        {
                            report_saved.ExternalLinks.Clear();
                        }
                        catch
                        {

                        }
                        Dictionary<string, DataTable> dic_data = FAI_lib.Export_FAI_Batch(sqlcon,report_saved, _itemcode, _lotno, report_type,_shift);
                        FAI_lib.Export_To_FAI(sqlcon, Report_Pack, _itemcode, _lotno, dic_data, report_type);
                        if(file_existed)
                        {
                            Report_Pack.Save();
                        }
                        else
                        {
                            Report_Pack.SaveAs(new FileInfo(export_file));
                        }
                        Report_Pack.Dispose();
                        MessageBox.Show(new Form { TopMost=true},"Hoàn thành xuất dữ liệu", "Thông báo");
                        ProcessStartInfo pi = new ProcessStartInfo(export_file);
                        Process.Start(pi);
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true },"Không tìm thấy dữ liệu của ItemCode / Lotno / Shift : " + _itemcode + " / " + _lotno + " / " + _shift, "Thông báo");
                    }
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true },"Không tìm thấy Format", "Cảnh báo");
                }
            }
        }

        public List<string> get_multiple_format(string src_path, List<string> extensions, string tar_ItemCode, string tar_process)
        {
            List<string> result = new List<string>();
            List<char> remove_char = new List<char> { ' ', '-', '_' };
            string process = new string(tar_process.Where(x => remove_char.IndexOf(x) == -1).ToArray());
            DirectoryInfo directory = new DirectoryInfo(src_path);

            var files = extensions.SelectMany(e => directory.EnumerateFiles(e, SearchOption.AllDirectories)).Where(x => x.FullName.Contains(tar_ItemCode) && x.FullName.Contains(tar_process));
            foreach (var item in files)
            {
                result.Add(item.FullName);
            }
            return result;
        }
        public List<string> get_multiple_format(string src_path, List<string> extensions, string tar_ItemCode)
        {
            List<string> result = new List<string>();
            List<char> remove_char = new List<char> { ' ', '-', '_' };
            DirectoryInfo directory = new DirectoryInfo(src_path);
            var files = extensions.SelectMany(e => directory.EnumerateFiles(e, SearchOption.TopDirectoryOnly)).Where(x => x.FullName.Contains(tar_ItemCode));
            foreach (var item in files)
            {
                result.Add(item.FullName);
            }
            return result;
        }

        private void txtLotNo_Sel_TextChanged(object sender, EventArgs e)
        {
            Invoke(new checkTextbox(check_txt), txtLotNo_Sel) ;
        }

        private void txtLotNo_Sel_Validated(object sender, EventArgs e)
        {
            txtLotNo_Sel.Text = myCode.Lotno_Formated(txtLotNo_Sel.Text);
        }

        private void tabMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            string tar_tab = tabMain.SelectedTab.Text;
            if(tar_tab=="Data")
            {
                txtItemCode_Sel.Text = txtItemCode.Text;
                txtLotNo_Sel.Text = txtLotNo.Text;
                cbShift_Sel.SelectedIndex = cbShift.SelectedIndex;
                cbType_Sel.SelectedIndex = cbType.SelectedIndex;
            }
        }
        private void btnEditData_Click(object sender, EventArgs e)
        {
            start_lbl: if (spData.Panel2.Controls.Count == 0)
            {
                FAI_Edit frmFAIEdit = new FAI_Edit(txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbType_Sel.SelectedItem.ToString(), cbShift_Sel.SelectedItem.ToString(),txtOperator.Text,cbMachine.Text) { Dock = DockStyle.Fill, TopLevel = false, TopMost = true };
                frmFAIEdit.Show();
                spData.Panel2.Controls.Add(frmFAIEdit);
            }
            else
            {
                spData.Panel2.Controls.Clear();
                goto start_lbl;
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
