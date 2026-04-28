using Funtion_F3_SMT;
using OfficeOpenXml;
using OK2SHIP_Lib;
using OK2SHIP_Measurements.Services;
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
using TDMK_EPPLUS_7;
using TDMK_Helper;
using TDMK_SQL;
using ZedGraph;
using myExcel = Microsoft.Office.Interop.Excel;

namespace OK2SHIP_Measurements
{
    public partial class FrmFAI : Form
    {
        //public PTH_Diameter_Process PTH_Diameter_proc = new PTH_Diameter_Process();
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        //public SEI_Lib myCode = new SEI_Lib();
        public TDMK_OK2SHIP myCode = new TDMK_OK2SHIP();
        TDMK_EPPLUS7_lib FAI_lib = new TDMK_EPPLUS7_lib();
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
        myVar myCode2 = new myVar();
        IPQC_LogFile IPQC_LogFile = new IPQC_LogFile();
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
        //SEI_Lib.FAI_Spec[] testFAI_spec;
        public FrmFAI()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            checkLogin();
        }
        //public FrmFAI(string src_Dev, string src_DB, SqlConnection tar_sqlcon, string tar_datalocation, string tar_formatfolder, string tar_f_ext, string tar_log_folder)
        //{
        //    InitializeComponent();
        //    Dev = src_Dev;
        //    DB_name = src_DB;
        //    sqlcon = tar_sqlcon;
        //    Data_Location = tar_datalocation;
        //    format_folder = tar_formatfolder;
        //    f_ext = tar_f_ext;
        //    log_folder = tar_log_folder;
        //}

        public bool check_FAIdata_OK2(DataTable src_DGV_data, DataTable src_DGV_Spec)
        {
            bool result = false;
            for (int i = 0; i < src_DGV_data.Columns.Count; i++)
            {
                string columnName = src_DGV_data.Columns[i].ColumnName;
                string text = checkDBNull(src_DGV_Spec.Rows[1][columnName]);
                string text2 = checkDBNull(src_DGV_Spec.Rows[2][columnName]);
                string setVal = checkDBNull(src_DGV_Spec.Rows[0][columnName]);
                string src_UL = myCode.IsNumeric_Val(text);
                string src_LL = myCode.IsNumeric_Val(text2);
                int num = 0;
                while (num < src_DGV_data.Rows.Count)
                {
                    string src_act_val = checkDBNull(src_DGV_data.Rows[num][i]);
                    if (string.IsNullOrEmpty(src_act_val))
                        continue;
                    if (check_in_limit(src_UL, src_LL, src_act_val, setVal))
                    {
                        result = true;
                        num++;
                        continue;
                    }

                    goto end_label;
                }

                continue;
            end_label:
                result = false;
                break;
            }
            return result;
        }

        public FrmFAI(string src_Dev, string src_DB, SqlConnection tar_sqlcon, string tar_datalocation, string tar_formatfolder, string tar_reportlocation, string tar_f_ext, string tar_log_folder)
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
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
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
                        if (DGV_DataView.SelectedCells.Count > 0)
                        {
                            DGV_DataView.ContextMenuStrip = cmsCopyPaste;
                        }
                    }

                }
                else
                {
                    TDMK_Message.MessageBoxTDMK_Warning("Bạn cần đăng nhập để chỉnh sửa");
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
            string watch_path = Data_Location;
            string[] filter_ext = new string[] { ".xls", ".xlsx" };
            watchfolder.Path = watch_path;
            watchfolder.NotifyFilter = NotifyFilters.DirectoryName | NotifyFilters.FileName | NotifyFilters.Attributes;
            if (Dev == "Roughness")
            {
                watchfolder.Filter = "*.xls";
                numMeasLoc.Enabled = true;
            }
            else
            {
                watchfolder.Filter = f_ext;// "*.csv";
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
            cbType.SelectedIndex = 0;
            cbShift.SelectedIndex = 0;
            cbType_Sel.SelectedIndex = 0;
            cbShift_Sel.SelectedIndex = 0;
            //if(DB_name!="OK2SHIP_SMT")
            //{
            //    tabMain.TabPages.RemoveAt(1);
            //}
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
                    if (!e.FullPath.Contains("$"))
                    {
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
                            TDMK_Message.MessageBoxTDMK_Warning("Bạn chưa ấn nút Start");
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
                            DataTable sel_dt = new DataTable();
                            if (EditMode)
                            {
                                sel_DGV = DGV_Data_Plus;
                            }
                            else
                            {
                                sel_DGV = DGV_DataView;
                            }
                            if (sel_dt.Columns.Count == 0)
                            {
                                foreach (DataGridViewColumn t in DGV_SpecView.Columns)
                                {
                                    if (!t.Name.Contains("Format"))
                                        sel_dt.Columns.Add(t.Name);
                                }
                            }
                            if (sel_dt.Rows.Count == 0)
                            {
                                for (int i = 0; i < Convert.ToInt32(numQty.Value); i++)
                                {
                                    sel_dt.Rows.Add();
                                    //sel_DGV.Rows[i].HeaderCell.Value = (i + 1).ToString();
                                }
                                //sel_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                                //Disable_Sort_DGV(sel_DGV);
                            }
                            sel_DGV.DataSource = sel_dt;
                            myCode.DGV_Auto_Resize(sel_DGV);
                            Disable_Sort_DGV(sel_DGV);
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
                    TDMK_Message.MessageBoxTDMK_Warning("Hãy nhấn nút Load Spec!");
                }

            }
            else
            {
                TDMK_Message.MessageBoxTDMK_Warning("Hãy ấn reset trước!");
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
                myCode.Clear_DGV(DGV_SpecView);
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
            if (DGV_DataView.Rows.Count >= Convert.ToInt32(numQty.Value))
            {
                //src_btn.Text = txt_str;
                //src_btn.BackColor = curr_color;
                //pro_en = false;
                //timer1.Enabled = false;
                //btnSave.Enabled = true;
                //myCode.Load_CPK_DGV(myCode.DGV_To_Table(DGV_DataView), DGV_CPK, myCode.testFAI_spec.ToArray());
                if (Dev != "Roughness")
                {
                    myCode.Load_CPK_DGV_Histogram(DGV_DataView, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
                }
                //if ((DGV_Data_Plus.Rows.Count > 0) && (!TDMK_OK2SHIP.admin_mode) && (!EditMode))
                //{
                //    TDMK_OK2SHIP.confirm_request = true;
                //    GBControl.Enabled = false;
                //}
                if (DGV_DataView.Rows.Count > Convert.ToInt32(numQty.Value))
                {
                    TDMK_Message.MessageBoxTDMK_Warning("Dữ liệu vào cần nhiều hơn " + numQty.Value.ToString());
                }
            }
        }
        public bool check_data_OK(DataGridView src_DGV_data, DataGridView src_DGV_Spec)
        {
            bool result = false;
            for (int i = 0; i < src_DGV_Spec.Columns.Count; i++)
            {
                string src_UL = checkDBNull(src_DGV_Spec.Rows[1].Cells[i].Value);
                string src_LL = checkDBNull(src_DGV_Spec.Rows[2].Cells[i].Value);
                string setVal = checkDBNull(src_DGV_Spec.Rows[0].Cells[i].Value);
                int num = 0;
                while (num < src_DGV_data.Rows.Count)
                {
                    string src_act_val = checkDBNull(src_DGV_data.Rows[num].Cells[i].Value);
                    if (string.IsNullOrEmpty(src_act_val))
                        continue;
                    if (check_in_limit(src_UL, src_LL, src_act_val, setVal))
                    {
                        result = true;
                        num++;
                        continue;
                    }

                    goto IL_00b8;
                }

                continue;
            IL_00b8:
                result = false;
                break;
            }

            return result;
        }
        public bool check_FAIdata_OK2(DataGridView src_DGV_data, DataGridView src_DGV_Spec)
        {
            bool result = false;
            for (int i = 0; i < src_DGV_data.Columns.Count; i++)
            {
                string name = src_DGV_data.Columns[i].Name;
                string value = checkDBNull(src_DGV_Spec.Rows[1].Cells[name].Value);
                string value2 = checkDBNull(src_DGV_Spec.Rows[2].Cells[name].Value);
                string text = checkDBNull(src_DGV_Spec.Rows[0].Cells[name].Value);
                double num = Convert.ToDouble(text) + Convert.ToDouble(value);
                double num2 = Convert.ToDouble(text) - Convert.ToDouble(value2);
                int num3 = 0;
                while (num3 < src_DGV_data.Rows.Count)
                {
                    string src_act_val = checkDBNull(src_DGV_data.Rows[num3].Cells[i].Value);
                    if (check_in_limit(num.ToString(), num2.ToString(), src_act_val, text))
                    {
                        result = true;
                        num3++;
                        continue;
                    }

                    goto IL_00f8;
                }

                continue;
            IL_00f8:
                result = false;
                break;
            }

            return result;
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
                if (TDMK_Message.MessageBoxDialog("Không đủ dữ liệu. Bạn muốn tiếp tục lưu?") == DialogResult.Yes)
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
                bool en_save = false;
                bool data_ok = false;
                if ((Dev == "Roughness") || (Dev == "AU_NI_Thickness"))
                {
                    data_ok = check_data_OK(DGV_DataView, DGV_SpecView);
                }
                else
                {
                    data_ok = check_FAIdata_OK2(myCode.DGV_To_Table(DGV_DataView), myCode.DGV_To_Table(DGV_SpecView));
                }
                if (data_ok)
                {
                    en_save = true;
                }
                else
                {
                    //Alt + O để bật login
                    bool isOK = TDMK_Message.MessageWarningLogin("Dữ liệu NG, hãy xác nhận với Leader");
                    if (isOK)
                    {
                        en_save = true;
                    }
                    //if ((MessageBox.Show("Dữ liệu NG, hãy xác nhận với Leader", "Warning", MessageBoxButtons.YesNo)) == DialogResult.Yes)
                    //{
                    //    if (TDMK_OK2SHIP.admin_mode)
                    //    {
                    //        en_save = true;
                    //    }
                    //    else
                    //    {
                    //        en_save = false;
                    //        MessageBox.Show(new Form { TopMost = true }, "Please, login to save data", "Warning");
                    //    }

                    //}
                    //else
                    //{
                    //    en_save = false;
                    //}
                }
                if (en_save)
                {
                    switch (Dev)
                    {
                        case "Roughness":
                            Save_data_Roughness(OK2SHIP_Measurements.sqlcon_IPQC);
                            break;
                        case "AU_NI_Thickness":
                            Save_data_AU_NI();
                            break;
                        default:
                            //try
                            //{
                            //    Alarm_FAI_NG_Saved_Encode_Email();
                            //}
                            //catch
                            //{

                            //}
                            if (DB_name == "OK2SHIP_Period2")
                            {
                                Save_PTH_Diameter();
                            }
                            else
                            {
                                Save_FAI_data(EditMode);
                            }

                            break;
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
        public void Save_FAI_data(bool _editted = false)
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
            item_vals[8] = cbType.SelectedItem.ToString().ToUpper();
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
                if (DGV_DataView.Rows.Count > 0)
                {
                    //DataTable logfile_data_tbl = myCode.DGV_To_Table(DGV_DataView);
                    foreach (DataGridViewRow row in DGV_DataView.Rows)
                    {
                        if (row.IsNewRow) continue; // bỏ qua dòng trống cuối

                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            if (cell.Value == null || string.IsNullOrWhiteSpace(cell.Value.ToString()))
                            {
                                TDMK_Message.MessageBoxTDMK_Warning("Chưa điền đủ dữ liệu.");
                                return;
                            }
                        }
                    }
                begin_lbl: string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Machine", "Remark", "Shift" }, new string[] { txtItemCode.Text, txtLotNo.Text, cbMachine.Text, cbType.SelectedItem.ToString().ToUpper(), cbShift.SelectedItem.ToString() });
                    DataTable fai_dt = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Auto", filter_str);
                    List<string> existed_FAI_list = fai_dt.AsEnumerable().Select(x => x.Field<string>("FAI_No")).Distinct().ToList();
                start_label: List<string> added_FAI_lst = new List<string>();
                    int ID = TDMK_Code.SQL_MAX("FAI_Auto", "ID", sqlcon);
                    for (int i = 0; i < DGV_DataView.ColumnCount; i++)
                    {
                        string curr_col_name = DGV_DataView.Columns[i].Name;
                        if (existed_FAI_list.IndexOf(curr_col_name) == -1)
                        {
                            item_vals[6] = curr_col_name;
                            for (int j = 0; j < DGV_DataView.RowCount; j++)
                            {
                                if (checkDBNull(DGV_DataView.Rows[j].Cells[i].Value) != "")
                                {
                                    item_vals[0] = ID++.ToString();// (TDMK_Code.SQL_MAX("FAI_Auto", "ID", sqlcon) + 1).ToString();//"ID";
                                    item_vals[7] = checkDBNull(DGV_DataView.Rows[j].Cells[i].Value).Trim(trim_char);
                                    TDMK_Code.insert_val_arr("FAI_Auto", sqlcon, items, item_vals);
                                }
                                else
                                {
                                    break;
                                }
                            }
                            added_FAI_lst.Add(curr_col_name);
                        }
                    }
                    foreach (string col in added_FAI_lst)
                    {
                        if (DGV_DataView.Columns.Contains(col))
                            DGV_DataView.Columns.Remove(col);
                        if (DGV_CPK.Columns.Contains(col))
                            DGV_CPK.Columns.Remove(col);
                        if (DGV_Histogram.Columns.Contains(col))
                            DGV_Histogram.Columns.Remove(col);
                    }
                    if (DGV_DataView.Columns.Count > 0)
                    {
                        if (!_editted)
                        {
                            if (TDMK_Message.MessageBoxDialog("Dữ liệu đã tồn tại. Bạn muốn cộng dồn dữ liệu?") == DialogResult.Yes)
                            {
                                if (TDMK_OK2SHIP.admin_mode)
                                {
                                    existed_FAI_list.Clear();
                                    goto start_label;
                                }
                                else
                                {
                                    TDMK_Message.MessageBoxTDMK_Warning("Bạn cần đăng nhập để cộng dồn dữ liệu đã có");
                                }
                            }
                        }
                        else
                        {
                            if (TDMK_Message.MessageBoxDialog("Dữ liệu đã tồn tại. Bạn muốn thay bằng dữ liệu mới?") == DialogResult.Yes)
                            {
                                if (TDMK_OK2SHIP.admin_mode)
                                {
                                    foreach (DataGridViewColumn dgv_c in DGV_DataView.Columns)
                                    {
                                        string col_name = dgv_c.Name;
                                        string del_filter = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Machine", "Remark", "FAI_No" }, new string[] { txtItemCode.Text, txtLotNo.Text, cbMachine.Text, cbType.SelectedItem.ToString().ToUpper(), col_name });
                                        TDMK_Code.Delelte_FilteredItem_arr("FAI_Auto", sqlcon, del_filter);

                                    }
                                    goto begin_lbl;
                                }
                                else
                                {
                                    TDMK_Message.MessageBoxTDMK_Warning("Bạn cần đăng nhập để thay mới dữ liệu đã có");
                                }
                            }
                        }
                    }
                    else
                    {
                        DGV_DataView.Columns.Clear();
                        DGV_CPK.Columns.Clear();
                        DGV_Histogram.Columns.Clear();
                        TDMK_Message.MessageBoxTDMK_Info("Ghi dữ liệu thành công");
                    }
                }
                else
                {
                    TDMK_Message.MessageBoxTDMK_Warning("Không có dữ liệu!");
                }
            }
            else
            {
                TDMK_Message.MessageBoxTDMK_Warning("Nhập dữ liệu thông tin!");
            }
        }
        private void RBNormal_CheckedChanged(object sender, EventArgs e)
        {
            if (RBNormal.Checked)
            {
                EditMode = false;
                DGV_SpecView.ContextMenuStrip = null;
                btnLoadEditData.Enabled = false;
                DGV_DataView.DataSource = null;
                DGV_CPK.DataSource = null;
                DGV_Histogram.DataSource = null;
                DGV_Data_Plus.DataSource = null;
            }
        }

        private void RBEdit_CheckedChanged(object sender, EventArgs e)
        {
            if (RBEdit.Checked)
            {
                EditMode = true;
                btnLoadEditData.Enabled = true;
            }
        }

        private void DGV_Data_Plus_MouseClick(object sender, MouseEventArgs e)
        {
            if ((e.Button == MouseButtons.Right) && DGV_Data_Plus.Columns.Count > 0)
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
                            if (myCode.check_columns_existed(myCode.DGV_To_Table(DGV_SpecView), src_col_name))
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
                            TDMK_Message.MessageBoxTDMK_Warning("Sai vị trí");
                        }
                    }
                    myCode.Load_CPK_DGV_Histogram(DGV_DataView, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
                }
                else
                {
                    TDMK_Message.MessageBoxTDMK_Warning("Chọn dòng");
                }
            }
            else
            {
                TDMK_Message.MessageBoxTDMK_Warning("Bạn cần đăng nhập để chỉnh sửa");
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
            bool en_proc = false;
            switch (Dev)
            {
                case "AU_NI_Thickness":
                    en_proc = Load_Spec_IPQC(sqlcon, "SpecList", txtItemCode.Text, "THICKNESS", DGV_SpecView);
                    break;
                case "Roughness":
                    en_proc = Get_Roughness_Spec(txtItemCode.Text, sqlcon, DGV_SpecView);

                    /*********************** Load Spec from ACF Sheet****************************************************/
                    break;
                case "Mitutoyo":
                    if (DB_name == "OK2SHIP_Period2")
                    {
                        en_proc = myCode.Load_PTH_Diameter_Spec(txtItemCode.Text, sqlcon, "SPEC_PTH_Diameter", ref DGV_SpecView);
                    }
                    else
                    {
                        en_proc = myCode.Load_Spec(sqlcon, txtItemCode.Text, DGV_SpecView, "OMM", "");
                    }
                    break;
                case "Nikon":
                    if (DB_name == "OK2SHIP_Period2")
                    {
                        en_proc = myCode.Load_PTH_Diameter_Spec(txtItemCode.Text, sqlcon, "SPEC_PTH_Diameter", ref DGV_SpecView);
                    }
                    else
                    {
                        en_proc = myCode.Load_Spec(sqlcon, format_folder, txtItemCode.Text, txtLotNo.Text, ".xlsm", DGV_SpecView, "OMM", "");
                    }
                    break;
                case "Keyence":
                    if (DB_name == "OK2SHIP_Period2")
                    {
                        en_proc = myCode.Load_PTH_Diameter_Spec(txtItemCode.Text, sqlcon, "SPEC_PTH_Diameter", ref DGV_SpecView);
                    }
                    else
                    {
                        en_proc = myCode.Load_Spec(sqlcon, format_folder, txtItemCode.Text, txtLotNo.Text, ".xlsm", DGV_SpecView, "OMM", "");
                    }
                    break;
                default:
                    if (cbMachine.Text != "")
                    {
                        en_proc = myCode.Load_Spec(sqlcon, format_folder, txtItemCode.Text, txtLotNo.Text, ".xlsm", DGV_SpecView, cbMachine.Text);
                    }
                    else
                    {
                        TDMK_Message.MessageBoxTDMK_Warning("Hãy chọn loại máy");
                    }
                    break;
            }
            if (en_proc)
            {
                GBControl.Enabled = true;
                TDMK_Message.MessageBoxTDMK_Info("Hệ thống sẵn sàng!");
            }
            else
            {
                TDMK_Message.MessageBoxTDMK_Warning("Dữ liệu Spec chưa được cài đặt hoặc Sai thiết bị");
                GBControl.Enabled = false;
            }

            /***************************************** Finish Testing Areas **************************************************************************/
        }
        private void btnLogin_Click(object sender, EventArgs e)
        {
            if (btnLogin.Text == "Logout")
            {
                UserSession.Instance.Logout();
                btnLogin.Text = "Login";
                numQty.Enabled = false;
                btnLogin.BackColor = curr_color;
            }
            else
            {
                Login frm_login = new Login();
                frm_login.ShowDialog();

            }
            checkLogin();
        }
        private void checkLogin()
        {
            txtOperator.Text = UserSession.Instance.User_ID;
            if (UserSession.Instance.IsLoggedIn)
            {
                btnLogin.Text = "Logout";
                numQty.Enabled = true;
                btnLogin.BackColor = Color.LightGreen;
            }
            else
            {
                btnLogin.Text = "Login";
                numQty.Enabled = false;
                btnLogin.BackColor = curr_color;
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
                string SetVal = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[0].Cells[src_col_name].Value)); //Convert.ToDouble(DGV_SpecView.Rows[0].Cells[src_col_name].Value);
                string _UL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[1].Cells[src_col_name].Value));
                string _LL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[2].Cells[src_col_name].Value));
                foreach (DataGridViewRow r in DGV_Data_Plus.Rows)
                {
                    string act_val = checkDBNull(r.Cells[src_col_name].Value);
                    r.Cells[src_col_name].Style.BackColor = myCode.check_in_limit_Color(_UL, _LL, act_val, SetVal);
                }
            }
            else
            {
                TDMK_Message.MessageBoxTDMK_Warning("Không thể đổi tên");
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
                TDMK_Message.MessageBoxTDMK_Warning("Không thể đổi tên");
            }
        }

        public void Scan_data(object src_log)
        {
            Thread.Sleep(1000);
            string itemcode_folder = "";
            string done_file = "";
            string logfile = src_log.ToString();
            FileInfo f_info = new FileInfo(src_log.ToString());
            string f_exten = f_info.Extension;
            string save_time = Create_date_string(DateTime.Now);
            string f_name = Path.GetFileName(src_log.ToString());
            if (f_exten.ToUpper() == f_ext.Replace("*", "").ToUpper())
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
                            if (DB_name == "OK2SHIP_Period2")
                            {
                                Invoke(new Data_process(myCode.PTH_Diameter_process), logfile, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode);
                            }
                            else
                            {
                                Invoke(new Data_process(myCode.Mitutoyo_process), logfile, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode);
                            }
                            break;
                        case "Nikon":
                            if (DB_name == "OK2SHIP_Period2")
                            {
                                Invoke(new Data_process(myCode.PTH_Diameter_process), logfile, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode);
                            }
                            else
                            {
                                Invoke(new Data_process(myCode.Nikon_process), logfile, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode);
                            }
                            break;
                        case "Keyence":
                            if (DB_name == "OK2SHIP_Period2")
                            {
                                Invoke(new Data_process(myCode.PTH_Diameter_process), logfile, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode);
                            }
                            else
                            {
                                Invoke(new Data_process(myCode.Keyence_Display_Data3), logfile, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode);
                            }
                            break;
                        case "AU_NI_Thickness":
                            Invoke(new Data_process2(myCode.AU_IN_Scan), logfile, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode, Convert.ToInt32(numQty.Value));
                            break;
                        case "Roughness":
                            //Invoke(new Data_process2(myCode.Roughness_Scan), TDMK_OK2SHIP.scan_file, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode, Convert.ToInt32(numQty.Value));
                            //Invoke(new Roughness_process(myCode.Roughness_Scan2), TDMK_OK2SHIP.scan_file, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode, Convert.ToInt32(numQty.Value),Convert.ToInt32(numMeasLoc.Value));
                            Invoke(new Roughness_process(Roughness_Scan_ACF_SMT), logfile, DGV_DataView, DGV_SpecView, DGV_Data_Plus, EditMode, Convert.ToInt32(numQty.Value), Convert.ToInt32(numMeasLoc.Value));
                            break;
                    }
                }
                catch
                {
                    TDMK_Message.MessageBoxTDMK_Warning("Không đúng định dạng file.");
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
            Invoke(new SetText(SetTextButton), btnRun, "Start");

        }
        private void txtDataFolder_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            FolderBrowserDialog folder_data = new FolderBrowserDialog();
            if (folder_data.ShowDialog() == DialogResult.OK)
            {
                string log_path = folder_data.SelectedPath;
                txtDataFolder.Text = log_path;
                Manual_Data_Process(log_path);
            }
        }

        private void txtItemCode_KeyDown(object sender, KeyEventArgs e)
        {
            //if((DB_name=="SEI_FAI")&&(Dev=="Manual"))
            if ((myVar.FAI_form != null) && (Dev == "Manual"))
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
            DataTable Spec_tbl = TDMK_Code.Datatable_Filter(sqlcon, spec_tbl_name, TDMK_Code.filter_str(new string[] { "ItemCode", "Item", "Remark" }, new string[] { tar_ItemCode, "%" + tar_itemcheck + "%", "%AU_Plating%" })); //Roughness AU_NI            
            if (Spec_tbl.Rows.Count == 0)
            {
                TDMK_Message.MessageBoxDialog("Chưa tồn tại dữ liệu spec trong cơ sở dữ liệu");
                _result = false;
            }
            else
            {
                string[] items_list = Spec_tbl.AsEnumerable().Select(r => r.Field<string>("Item")).ToArray();
                string[] limit_lst = Spec_tbl.AsEnumerable().Select(r => r.Field<string>("Spec")).ToArray();
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
                    TDMK_Message.MessageBoxTDMK_Info("Ghi dữ liệu thành công");
                }
                else
                {
                    TDMK_Message.MessageBoxTDMK_Warning("Không có dữ liệu!");
                }
            }
            else
            {
                TDMK_Message.MessageBoxTDMK_Warning("Nhập dữ liệu thông tin!");
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
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
            start_label: DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "Roughness", filter_str);
                if (dt.Rows.Count > 0)
                {
                    if (TDMK_Message.MessageBoxDialog("Dữ liệu Roughness đã có. Bạn muốn cập nhật lại ?") == DialogResult.Yes)
                    {
                        TDMK_Code.Delelte_FilteredItem_arr("Roughness", sqlcon, filter_str);
                        goto start_label;
                    }
                }
                else
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
                        TDMK_Message.MessageBoxTDMK_Info("Ghi dữ liệu thành công");
                    }
                    else
                    {
                        TDMK_Message.MessageBoxTDMK_Warning("Không có dữ liệu!");
                    }
                }
            }
            else
            {
                TDMK_Message.MessageBoxTDMK_Warning("Nhập dữ liệu thông tin!");
            }
        }
        public void Save_data_Roughness(SqlConnection tar_sqlcon)
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
            items[8] = "L1_Roughness_Sdr";
            items[9] = "L2_Roughness_Sa";
            items[10] = "L2_Roughness_Sq";
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
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
            start_label: DataTable dt = TDMK_Code.Datatable_Filter(tar_sqlcon, "Roughness", filter_str);
                if (dt.Rows.Count > 0)
                {
                    if (TDMK_Message.MessageBoxDialog("Dữ liệu Roughness đã có. Bạn muốn cập nhật lại ?") == DialogResult.Yes)
                    {
                        TDMK_Code.Delelte_FilteredItem_arr("Roughness", tar_sqlcon, filter_str);
                        goto start_label;
                    }
                }
                else
                {
                    if (DGV_DataView.Rows.Count > 0)
                    {
                        for (int j = 0; j < DGV_DataView.RowCount; j++)
                        {
                            for (int i = 0; i < DGV_DataView.Columns.Count; i++)
                            {
                                item_vals[6 + i] = checkDBNull(DGV_DataView.Rows[j].Cells[i].Value); // "L1_AU";
                            }
                            item_vals[0] = (TDMK_Code.SQL_MAX("Roughness", "ID", tar_sqlcon) + 1).ToString();//"ID";
                            TDMK_Code.insert_val_arr("Roughness", tar_sqlcon, items, item_vals);
                        }
                        DGV_DataView.Columns.Clear();
                        TDMK_Message.MessageBoxTDMK_Info("Ghi dữ liệu thành công");
                    }
                    else
                    {
                        TDMK_Message.MessageBoxTDMK_Warning("Không có dữ liệu!");
                    }
                }
            }
            else
            {
                TDMK_Message.MessageBoxTDMK_Warning("Nhập dữ liệu thông tin!");
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
                TDMK_Message.MessageBoxTDMK_Warning("Chọn dòng để loại bỏ");
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
                TDMK_Message.MessageBoxTDMK_Warning("Cần có sự xác nhận từ PIC");
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
                TDMK_Message.MessageBoxTDMK_Error("Copy/paste lỗi. " + ex.Message);
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
                TDMK_Message.MessageBoxTDMK_Warning("Hãy chọn ô");
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
                TDMK_Message.MessageBoxTDMK_Warning("Hãy chọn dữ liệu!");
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
                        TDMK_Message.MessageBoxTDMK_Warning("Bạn cần nhập giá trị là số!");
                        DGV_DataView.CurrentCell.Value = "";
                    }
                    else
                    {
                        if (myCode.check_columns_existed(myCode.DGV_To_Table(DGV_SpecView), curr_col_name))
                        {
                            string SetVal = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[0].Cells[curr_col_name].Value));
                            string _UL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[1].Cells[curr_col_name].Value));
                            string _LL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[2].Cells[curr_col_name].Value));
                            string act_val = checkDBNull(DGV_DataView.CurrentCell.Value);
                            curr_cell.Style.BackColor = myCode.check_in_limit_Color(_UL.ToString(), _LL.ToString(), act_val, SetVal);

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
            PasteClipboardValue(true, DGV_DataView);
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
            PasteClipboardValue(false, DGV_Data_Plus);
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
                TDMK_Message.MessageBoxTDMK_Warning("Không đủ dữ liệu");
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
                TDMK_Message.MessageBoxTDMK_Info("Hoàn thành!");
            }
            else
            {
                TDMK_Message.MessageBoxTDMK_Warning("Không tìm thấy file format");
            }
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
                        TDMK_Message.MessageBoxTDMK_Warning("Bạn cần nhập giá trị là số");
                        DGV_Data_Plus.CurrentCell.Value = "";
                    }
                    else
                    {
                        if (myCode.check_columns_existed(myCode.DGV_To_Table(DGV_SpecView), curr_col_name))
                        {
                            string SetVal = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[0].Cells[curr_col_name].Value));
                            string _UL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[1].Cells[curr_col_name].Value));
                            string _LL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[2].Cells[curr_col_name].Value));
                            string act_val = checkDBNull(DGV_Data_Plus.CurrentCell.Value);
                            curr_cell.Style.BackColor = myCode.check_in_limit_Color(_UL, _LL, act_val, SetVal);
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
                        TDMK_Message.MessageBoxTDMK_Warning("Bạn cần nhập giá trị là số");
                        DGV_Data_Plus.CurrentCell.Value = "";
                    }
                    else
                    {
                        if (myCode.check_columns_existed(myCode.DGV_To_Table(DGV_SpecView), curr_col_name))
                        {
                            string SetVal = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[0].Cells[curr_col_name].Value));
                            string _UL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[1].Cells[curr_col_name].Value));
                            string _LL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[2].Cells[curr_col_name].Value));
                            string act_val = checkDBNull(DGV_Data_Plus.CurrentCell.Value);
                            curr_cell.Style.BackColor = myCode.check_in_limit_Color(_UL, _LL, act_val, SetVal);
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
            DataGridViewCell curr_cell;
            string curr_col_name;
            curr_cell = DGV_DataView.CurrentCell;
            curr_col_name = DGV_DataView.Columns[curr_cell.ColumnIndex].Name.ToString();
            if (checkDBNull(curr_cell.Value) != "")
            {
                if (!IsNumeric(checkDBNull(curr_cell.Value)))
                {
                    TDMK_Message.MessageBoxTDMK_Warning("Bạn cần nhập giá trị là số");
                    DGV_DataView.CurrentCell.Value = "";
                }
                else
                {
                    if (myCode.check_columns_existed(myCode.DGV_To_Table(DGV_SpecView), curr_col_name))
                    {
                        string SetVal = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[0].Cells[curr_col_name].Value));
                        string _UL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[1].Cells[curr_col_name].Value));
                        string _LL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_SpecView.Rows[2].Cells[curr_col_name].Value));
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
                    if (e.Value != null)
                    {
                        double d = double.Parse(e.Value.ToString());
                        //e.Value = d.ToString("N3");
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
                    //e.Value = d.ToString("N3");
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
            if (cbType.SelectedIndex == 1)
            {
                numQty.Value = 6;
            }
            else
            {
                numQty.Value = 32;
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
                TDMK_Message.MessageBoxTDMK_Warning("Không tìm thấy thư mục : " + src_path);
            }
            return result;
        }
        private void btnLoadData_Click(object sender, EventArgs e)
        {
        start_lbl: if (spData.Panel2.Controls.Count == 0)
            {
                bool FAI_mode = true;
                if (DB_name == "OK2SHIP_Period2")
                {
                    FAI_mode = false;
                }
                FAI_Data frmFAI = new FAI_Data(txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbType_Sel.SelectedItem.ToString(), cbShift_Sel.SelectedItem.ToString(), FAI_mode) { Dock = DockStyle.Fill, TopLevel = false, TopMost = true };
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
            if (DB_name == "OK2SHIP_Period2")
            {
                //Export_PTH_Diameter(format_folder, txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbType_Sel.Text.ToUpper());
            }
            else
            {
                Export_FAI_Data_byEPPLUS((int)numExport.Value);
            }
        }
        public void Export_FAI_Data_byExcel()
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
                if (TDMK_Message.MessageBoxDialog("Dữ liệu NG. Tiếp tục xuất dữ liệu?") == DialogResult.Yes)
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
                        TDMK_Message.MessageBoxTDMK_Warning("Không tìm thấy dữ liệu của ItemCode / Lotno / Shift : " + _itemcode + " / " + _lotno + " / " + _shift);
                    }
                }
                else
                {
                    TDMK_Message.MessageBoxTDMK_Warning("Không tìm thấy Format");
                }
            }
        }
        public DataTable Load_FAI_ToTable(SqlConnection tar_sqlcon, string tar_tbl, string FAI_No_name, string FAI_Data_Col_name, string tar_ItemCode, string tar_LotNo, string format_type)
        {
            DataTable FAI_spec_dt = TDMK_Code.Datatable_Filter(tar_sqlcon, "FAI_Spec", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { tar_ItemCode, format_type }));
            DataTable FAI_data_dt = TDMK_Code.Datatable_Filter(tar_sqlcon, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Remark" }, new string[] { tar_ItemCode, tar_LotNo, format_type }));
            List<string> temp_lst = FAI_spec_dt.AsEnumerable().Select(x => x.Field<string>("FAI_No")).Distinct().ToList();
            List<string> main_FAI = new List<string>();
            char[] split_chars = new char[] { '_', '/', '\\' };
            foreach (string t in temp_lst)
            {
                string setval = t.Split(split_chars).LastOrDefault().Trim();
                string fai_name = t.Split(split_chars).FirstOrDefault().Trim();
                string sel_FAI = fai_name + "_" + setval;
                if (main_FAI.IndexOf(sel_FAI) == -1)
                {
                    main_FAI.Add(sel_FAI);
                }
            }
            DataTable tbl_data = new DataTable();
            foreach (string c in main_FAI)
            {
                List<string> FAI_vals = FAI_data_dt.AsEnumerable().Where(x => x.Field<string>(FAI_No_name) == c).Select(x => x.Field<string>(FAI_Data_Col_name)).ToList();

                if (!myCode.check_columns_existed(tbl_data, c))
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
            return tbl_data;
        }
        private void txtLotNo_Sel_Validated(object sender, EventArgs e)
        {
            txtLotNo_Sel.Text = myCode.Lotno_Formated(txtLotNo_Sel.Text);
        }

        private void txtLotNo_Sel_TextChanged(object sender, EventArgs e)
        {
            Invoke(new checkTextbox(check_txt), txtLotNo_Sel);
        }
        public bool check_FAIdata_inSpec(DataTable src_data, DataTable src_Spec)
        {
            bool result = false;
            List<string> list = new List<string>();
            foreach (DataColumn column in src_Spec.Columns)
            {
                list.Add(column.ColumnName);
            }

            foreach (DataColumn column2 in src_data.Columns)
            {
                string columnName = column2.ColumnName;
                if (list.IndexOf(columnName) != -1)
                {
                    string text = checkDBNull(src_Spec.Rows[1][columnName]);
                    string text2 = checkDBNull(src_Spec.Rows[2][columnName]);
                    string setVal = myCode.IsNumeric_Val(checkDBNull(src_Spec.Rows[0][columnName]));
                    string src_UL = myCode.IsNumeric_Val(text);
                    string src_LL = myCode.IsNumeric_Val(text2);
                    foreach (DataRow row in src_data.Rows)
                    {
                        string src_act_val = myCode.IsNumeric_Val(checkDBNull(row[columnName]));
                        if (!check_in_limit(src_UL, src_LL, src_act_val, setVal))
                        {
                            result = false;
                            return result;
                        }

                        result = true;
                    }

                    continue;
                }

                TDMK_Message.MessageBoxTDMK_Warning("Không tìm thấy Spec của " + columnName);
                break;
            }

            return result;
        }

        public bool check_in_limit(string src_UL, string src_LL, string src_act_val, string SetVal)
        {
            if (string.IsNullOrEmpty(src_UL) && string.IsNullOrEmpty(src_LL))
            {
                return true;
            }
            bool flag2 = true;
            bool flag3 = true;
            if (!IsNumeric(SetVal) && !SetVal.Contains("="))
            {
                flag2 = false;
                flag3 = false;
            }

            if (IsNumeric(src_act_val))
            {
                if (src_UL != "")
                {
                    double num = Convert.ToDouble(src_UL);
                    double num2 = Convert.ToDouble(src_act_val);
                    if (src_LL != "")
                    {
                        double num3 = Convert.ToDouble(src_LL);
                        if (num2 <= num && num2 >= num3)
                        {
                            return true;
                        }

                        if (num2 > num)
                        {
                            return false;
                        }

                        return false;
                    }

                    if (num2 < num)
                    {
                        return true;
                    }

                    if (flag2)
                    {
                        if (num2 == num)
                        {
                            return true;
                        }

                        return false;
                    }

                    return false;
                }

                if (src_LL != "")
                {
                    double num4 = Convert.ToDouble(src_LL);
                    double num5 = Convert.ToDouble(src_act_val);
                    if (num5 > num4)
                    {
                        return true;
                    }

                    if (flag3)
                    {
                        if (num5 == num4)
                        {
                            return true;
                        }

                        return false;
                    }

                    return false;
                }

                return false;
            }

            return false;
        }


        public void Export_FAI_Data_byEPPLUS(int qty)
        {
            string type = cbType_Sel.SelectedItem.ToString().ToUpper();
            string format_loc = format_folder;
            string _itemcode = txtItemCode_Sel.Text;
            string _lotno = txtLotNo_Sel.Text;
            string _shift = cbShift_Sel.SelectedItem.ToString();
            DataTable FAI_dt = Load_FAI_ToTable(myVar.sqlcon_SMT, "FAI_Auto", "FAI_No", "FAI_Data", _itemcode, _lotno, cbType_Sel.SelectedItem.ToString()); //myCode.Load_FAI_ToTable(myVar.sqlcon_SMT, _itemcode, _lotno, cbType_Sel.SelectedItem.ToString());
            DataTable FAI_Spec = myCode.Load_FAI_Spec_ToTable(myVar.sqlcon_SMT, _itemcode, "");
            bool export_en = false;
            if (check_FAIdata_inSpec(FAI_dt, FAI_Spec))
            {
                export_en = true;
            }
            else
            {
                if (TDMK_Message.MessageBoxDialog("Dữ liệu NG. Tiếp tục xuất dữ liệu?") == DialogResult.Yes)
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
                    string report_type = cbType_Sel.SelectedItem.ToString().ToUpper();
                    DataTable FAI_Data_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Remark" }, new string[] { _itemcode, _lotno, cbType_Sel.SelectedItem.ToString() }));
                    if (FAI_Data_tbl.Rows.Count > 0)
                    {
                        string format_file = format_lst[0];
                        string export_file = "";
                        string report_path = Path.Combine(report_location, type);
                        if (!Directory.Exists(report_path))
                        {
                            Directory.CreateDirectory(report_path);
                        }
                        ExcelPackage Format_Pack = new ExcelPackage(format_file);
                        ExcelPackage Report_Pack = new ExcelPackage();
                        if (type.ToUpper() == "MASS")
                        {
                            export_file = Path.Combine(report_path, Path.GetFileNameWithoutExtension(format_file) + "-" + _lotno + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(format_file));//".xlsx"
                            List<string> selected_sht_name = new List<string>() { "FAI", "SPC", "Dimension" };
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
                        ExcelWorkbook report_saved = Report_Pack.Workbook;
                        Dictionary<string, DataTable> dic_data = Export_FAI_Batch(sqlcon, report_saved, _itemcode, _lotno, report_type, "", qty);
                        Export_To_FAI(sqlcon, report_saved, _itemcode, _lotno, dic_data, report_type);
                        Report_Pack.SaveAs(export_file);
                        Report_Pack.Dispose();
                        TDMK_Message.MessageBoxTDMK_Info("Hoàn thành xuất dữ liệu");
                        ProcessStartInfo pi = new ProcessStartInfo(export_file);
                        Process.Start(pi);
                    }
                    else
                    {
                        TDMK_Message.MessageBoxTDMK_Warning("Không tìm thấy dữ liệu của ItemCode / Lotno / Shift : " + _itemcode + " / " + _lotno + " / " + _shift);
                    }
                }
                else
                {
                    TDMK_Message.MessageBoxTDMK_Warning("Không tìm thấy Format");
                }
            }
        }

        public Dictionary<string, DataTable> Export_FAI_Batch(SqlConnection sqlcon, ExcelWorkbook src_format_wrk, string tar_ItemCode, string tar_LotNo, string format_type, string shift = "1", int qty = 0)
        {
            Dictionary<string, DataTable> dictionary = new Dictionary<string, DataTable>();
            char[] array = new char[3] { ' ', '\r', '\n' };
            //DataTable FAI_dt = Load_FAI_ToTable(myVar.sqlcon_SMT, "FAI_Auto", "FAI_No", "FAI_Data", _itemcode, _lotno, cbType_Sel.SelectedItem.ToString());
            DataTable source = null;
            if (format_type == "MASS")
            {
                source = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Auto", TDMK_Code.filter_str(new string[4] { "ItemCode", "LotNo", "Remark", "Shift" }, new string[4] { tar_ItemCode, "", format_type, shift }));
            }
            else
            {
                source = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Auto", TDMK_Code.filter_str(new string[4] { "ItemCode", "LotNo", "Remark", "Shift" }, new string[4] { tar_ItemCode, tar_LotNo, format_type, shift }));
            }

            DataTable Fai_Spec = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Spec", TDMK_Code.filter_str(new string[2] { "ItemCode", "Remark" }, new string[2] { tar_ItemCode, format_type }));

            List<string> source2 = new List<string> { "FAI", "SPC", "parentheses" };
            Dictionary<string, List<string>> dictionary2 = new Dictionary<string, List<string>>();
            int num = 32;
            //if (format_type.ToUpper() != "MASS")
            //{
            //    num = 32;
            //}

            foreach (ExcelWorksheet tg in src_format_wrk.Worksheets)
            {
                if (source2.Any((string x) => tg.Name.Contains(x)))
                {
                    string address = TDMK_EPPLUS7.Find_Cell_Addr("Dim. No.", "B10", tg, left_to_right: false);
                    string address2 = TDMK_EPPLUS7.Find_Cell_Addr("instrument", "B10", tg, left_to_right: false);
                    string address3 = TDMK_EPPLUS7.FindByOffset(tg.Cells[address].Offset(0, 1).Address, tg, left_to_right: true);
                    int columnOffset = tg.Cells[address3].End.Column - tg.Cells[address2].End.Column;
                    ExcelRangeBase excelRangeBase = tg.Cells[address].Offset(0, columnOffset);
                    dictionary2.Add(tg.Name, new List<string>());
                    for (int i = 0; checkDBNull(excelRangeBase.Offset(0, i).Value) != ""; i++)
                    {
                        string text = checkDBNull(excelRangeBase.Offset(-1, i).Value);
                        string text2 = checkDBNull(excelRangeBase.Offset(0, i).Value).Replace(" ", "");
                        string text3 = checkDBNull(excelRangeBase.Offset(1, i).Value);
                        string item = text2.Split('/').FirstOrDefault() + "_" + text3;
                        string name = tg.Name;
                        dictionary2[tg.Name].Add(item);
                    }
                }
            }

            foreach (KeyValuePair<string, List<string>> item2 in dictionary2)
            {
                List<double> error_position = new List<double>();

                DataTable dataTable = new DataTable();
                Dictionary<string, List<string>> dictionary3 = new Dictionary<string, List<string>>();

                string key = item2.Key;

                // Tối ưu xử lý chuỗi - chỉ Split 1 lần
                string[] array2 = item2.Value.Distinct()
                    .Select(text4 =>
                    {
                        var parts = text4.Split('/', '_');
                        return parts.FirstOrDefault() + "_" + parts.LastOrDefault();
                    })
                    .ToArray();

                // Pre-filter và cache data để tránh truy vấn lặp lại
                var sourceByFaiNo = source.AsEnumerable()
                    .Where(r => array2.Contains(r.Field<string>("FAI_No")))
                    .ToLookup(r => r.Field<string>("FAI_No"));

                var specDataCache = Fai_Spec.AsEnumerable()
                    .Where(r => array2.Contains(r.Field<string>("FAI_No")))
                    .ToDictionary(r => r.Field<string>("FAI_No"),
                                 r => new
                                 {
                                     TolMax = r.Field<string>("TolMax"),
                                     TolMin = r.Field<string>("TolMin")
                                 });

                foreach (string fai_no in array2)
                {
                    List<string> value = sourceByFaiNo[fai_no]
                        .Select(r => r.Field<string>("FAI_Data"))
                        .ToList();

                    if (format_type == "MASS")
                    {
                        // Sử dụng cache thay vì truy vấn lại
                        value = sourceByFaiNo[fai_no]
                            .Where(r => r.Field<string>("LotNo") == tar_LotNo)
                            .Select(r => r.Field<string>("FAI_Data"))
                            .ToList();

                        if (value.Count < num && value.Count == 6)
                        {
                            int needed = num - value.Count;

                            // Tối ưu truy vấn - lấy data 1 lần
                            var extraData = sourceByFaiNo[fai_no]
                                //.OrderByDescending(r => r.Field<string>("LotNo")) // Thay vì Reverse()
                                .Where(r => r.Field<string>("LotNo") != tar_LotNo)
                                .Reverse()
                                .Select(r => r.Field<string>("FAI_Data"))
                                .ToList();

                            // Kiểm tra spec data một lần
                            if (specDataCache.ContainsKey(fai_no))
                            {
                                var specData = specDataCache[fai_no];
                                Double.TryParse(specData.TolMax, out double tolMax);
                                Double.TryParse(specData.TolMin, out double tolMin);

                                // Tìm vị trí lỗi một lần
                                for (int i = 0; i < extraData.Count; i++)
                                {
                                    if (Double.TryParse(extraData[i], out double result))
                                    {
                                        if (tolMax < result || tolMin > result)
                                        {
                                            if (!error_position.Contains(i))
                                                error_position.Add(i + 6);
                                        }
                                    }
                                }
                            }

                            value.AddRange(extraData);
                        }
                    }

                    // Sử dụng TryAdd thay vì kiểm tra IndexOf
                    if (!dictionary3.ContainsKey(fai_no))
                    {
                        dictionary3.Add(fai_no, value);
                    }
                }

                // Tối ưu việc tạo DataTable
                foreach (KeyValuePair<string, List<string>> item3 in dictionary3)
                {
                    if (!dataTable.Columns.Contains(item3.Key))
                    {
                        if (item3.Value.Count != 0)
                        {
                            dataTable.Columns.Add(item3.Key, typeof(double));
                        }
                        else
                        {
                            dataTable.Columns.Add(item3.Key);
                        }
                    }
                }

                // Tính toán số rows cần thiết một lần
                int maxRows = 0;
                foreach (var item3 in dictionary3)
                {
                    int num3 = qty > 0 ? Math.Min(qty, item3.Value.Count) : item3.Value.Count;
                    maxRows = Math.Max(maxRows, num3);
                }

                // Thêm tất cả rows cần thiết một lần
                for (int i = 0; i < maxRows; i++)
                {
                    dataTable.Rows.Add();
                }

                // Populate data
                foreach (KeyValuePair<string, List<string>> item3 in dictionary3)
                {
                    int num3 = qty > 0 ? Math.Min(qty, item3.Value.Count) : item3.Value.Count;
                    int valueIndex = 0;

                    for (int m = 0; m < dataTable.Rows.Count; m++)
                    {
                        if (item3.Value.Count != 0)
                        {
                            if (m < num3)
                            {
                                // Skip error positions
                                while (error_position.Contains(valueIndex))
                                {
                                    valueIndex++;
                                }

                                if (valueIndex >= item3.Value.Count)
                                {
                                    TDMK_Message.MessageBoxTDMK_Warning($"Không đủ dữ liệu pcs tại Fai_No {item3.Key}. Vui lòng kiểm tra lại!");
                                    break;
                                }

                                dataTable.Rows[m][item3.Key] = item3.Value[valueIndex];
                                valueIndex++;
                            }
                            else
                            {
                                dataTable.Rows[m][item3.Key] = "N/A";
                            }
                        }
                    }
                }

                if (dataTable.Rows.Count > 0)
                {
                    dictionary.Add(key, dataTable);
                }
            }

            return dictionary;
        }

        public bool check_columns_existed(DataTable src_tbl, string find_col_name)
        {
            bool result = false;
            foreach (DataColumn column in src_tbl.Columns)
            {
                string columnName = column.ColumnName;
                if (columnName == find_col_name)
                {
                    result = true;
                    break;
                }
            }

            return result;
        }

        public Dictionary<string, int> Find_FAI_addr_qty(string start_addr, ExcelWorksheet tar_wrksht, bool left_to_right)
        {
            Dictionary<string, int> dictionary = new Dictionary<string, int>();
            ExcelRangeBase excelRangeBase = tar_wrksht.Cells[start_addr];
            int num = 0;
            int num2 = 0;
            string text = "";
            int num3 = 0;
            while (true)
            {
                ExcelRangeBase excelRangeBase2 = excelRangeBase.Offset(num, 0);
                if (left_to_right)
                {
                    excelRangeBase2 = excelRangeBase.Offset(0, num);
                }

                string text2 = checkDBNull(excelRangeBase2.Value);
                if (text2 == "")
                {
                    num2++;
                    if (num2 > 3)
                    {
                        break;
                    }
                }
                else
                {
                    if (IsNumeric(text2))
                    {
                        if (text == "")
                        {
                            text = excelRangeBase2.Address;
                        }

                        num3++;
                    }

                    num2 = 0;
                }

                num++;
            }

            if (text != "")
            {
                dictionary.Add(text, num3);
            }

            return dictionary;
        }

        TDMK_EPPLUS7_lib TDMK_EPPLUS7 = new TDMK_EPPLUS7_lib();

        public void Export_To_FAI(SqlConnection sqlcon, ExcelWorkbook src_format_wrk, string _ItemCode, string _LotNo, Dictionary<string, DataTable> dic_data, string format_type)
        {
            DataTable dataTable = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Spec", TDMK_Code.filter_str(new string[2] { "ItemCode", "Remark" }, new string[2] { _ItemCode, format_type }));
            List<string> source = new List<string> { "FAI", "SPC", "CPK", "parentheses" };
            foreach (KeyValuePair<string, DataTable> sht in dic_data)
            {
                try
                {
                    if (source.Any((string x) => sht.Key.Contains(x)))
                    {
                        string key = sht.Key;
                        ExcelWorksheet excelWorksheet = src_format_wrk.Worksheets[key];
                        string text = TDMK_EPPLUS7.Find_Cell_Addr("Dim. No.", "B10", excelWorksheet, left_to_right: false);
                        string address = TDMK_EPPLUS7.Find_Cell_Addr("instrument", "B10", excelWorksheet, left_to_right: false);
                        string address2 = TDMK_EPPLUS7.FindByOffset(excelWorksheet.Cells[text].Offset(0, 1).Address, excelWorksheet, left_to_right: true);
                        int columnOffset = excelWorksheet.Cells[address2].Start.Column - excelWorksheet.Cells[address].Start.Column;
                        Dictionary<string, int> dictionary = Find_FAI_addr_qty(text, excelWorksheet, left_to_right: false);
                        if (dictionary.Count > 0)
                        {
                            string address3 = dictionary.Keys.ToList()[0];
                            ExcelRangeBase excelRangeBase = excelWorksheet.Cells[address3].Offset(0, columnOffset);
                            excelRangeBase.LoadFromDataTable(sht.Value);
                            int toCol = sht.Value.Columns.Count - 1;
                            int num = sht.Value.Rows.Count - 1;
                            int row = excelRangeBase.Start.Row;
                            int column = excelRangeBase.Start.Column;
                            ExcelRangeBase excelRangeBase2 = excelWorksheet.Cells[row, column, row + num, toCol];
                            excelRangeBase2.Style.Numberformat.Format = "#0.000";
                        }
                    }
                }
                catch
                {
                }
            }
        }

        private void btnEditData_Click(object sender, EventArgs e)
        {
        start_lbl: if (spData.Panel2.Controls.Count == 0)
            {
                FAI_Edit frmFAIEdit = new FAI_Edit(txtItemCode_Sel.Text, txtLotNo_Sel.Text, cbType_Sel.SelectedItem.ToString(), cbShift_Sel.SelectedItem.ToString(), txtOperator.Text, cbMachine.Text) { Dock = DockStyle.Fill, TopLevel = false, TopMost = true };
                frmFAIEdit.Show();
                spData.Panel2.Controls.Add(frmFAIEdit);
            }
            else
            {
                spData.Panel2.Controls.Clear();
                goto start_lbl;
            }
        }

        private void tabMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabMain.SelectedIndex == 1)
            {
                txtItemCode_Sel.Text = txtItemCode.Text;
                txtLotNo_Sel.Text = txtLotNo.Text;
                cbShift_Sel.SelectedIndex = cbShift.SelectedIndex;
                cbType_Sel.SelectedIndex = cbType.SelectedIndex;
            }
        }
        public bool Get_Roughness_Spec(string ItemCode, SqlConnection tar_sqlcon, DataGridView tar_DGV)
        {
            DataTable ACF_Spec_dt = TDMK_Code.Datatable_Filter(tar_sqlcon, "ACF_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { ItemCode }));
            if (ACF_Spec_dt.Columns.Contains("Format_Type"))
            {
                ACF_Spec_dt.Columns.Remove("Format_Type");
            }
            if (ACF_Spec_dt.Rows.Count > 0)
            {
                DataTable result = new DataTable();
                foreach (DataColumn dc in ACF_Spec_dt.Columns)
                {
                    string col_name = dc.ColumnName.Replace("_", " Roughness ");
                    if ((col_name != "ID") && (col_name != "ItemCode"))
                    {
                        if (result.Columns.IndexOf(col_name) == -1)
                        {
                            result.Columns.Add(col_name);
                        }
                        if ((result.Rows.Count == 0) && (result.Columns.Count > 0))
                        {
                            for (int i = 0; i < 3; i++)
                            {
                                result.Rows.Add();
                            }
                        }
                        string spec_val = myCode.checkDBNull(ACF_Spec_dt.Rows[0][dc]);
                        if (spec_val.Contains("USL"))
                        {
                            result.Rows[0][col_name] = spec_val.Replace("USL:", "< ");
                            result.Rows[1][col_name] = spec_val.Replace("USL:", "");
                        }
                        else
                        {
                            if (spec_val.Contains("LSL"))
                            {
                                result.Rows[0][col_name] = spec_val.Replace("LSL:", "> ");
                                result.Rows[2][col_name] = spec_val.Replace("LSL:", "");
                            }
                            else if (col_name.Contains("Format Roughness Type"))
                            {
                                result.Rows[0][col_name] = spec_val;
                                result.Rows[1][col_name] = spec_val;
                                result.Rows[2][col_name] = spec_val;
                            }
                            else
                            {
                                result.Rows[0][col_name] = spec_val;
                            }
                        }

                    }
                }
                tar_DGV.DataSource = result;
                List<string> header_dgv = new List<string> { "SV", "UL", "LL" };
                for (int r = 0; r < 3; r++)
                {
                    tar_DGV.Rows[r].HeaderCell.Value = header_dgv[r];
                }
                tar_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                myCode.Disable_Sort_DGV(tar_DGV);
                return true;
            }
            else
            {
                return false;
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
                string[][] data_arr = new string[][] { L1_Sa, L1_Sq, L1_Sdr, L2_Sa, L2_Sq, L2_Sdr, L3_Sa, L3_Sq, L3_Sdr };
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
                    sel_DGV.Rows.Add(L1_Sa[i], L1_Sq[i], L1_Sdr[i], L2_Sa[i], L2_Sq[i], L2_Sdr[i], L3_Sa[i], L3_Sq[i], L3_Sdr[i]);
                    for (int j = 0; j < sel_DGV.Columns.Count - 5; j++)
                    {
                        string act_val = checkDBNull(sel_DGV.Rows[i].Cells[j].Value);
                        if (act_val != "")
                        {
                            string UL = checkDBNull(tar_DGV_Spec.Rows[1].Cells[j].Value);
                            string LL = checkDBNull(tar_DGV_Spec.Rows[2].Cells[j].Value);
                            string SV = checkDBNull(tar_DGV_Spec.Rows[0].Cells[j].Value);
                            sel_DGV.Rows[i].Cells[j].Style.BackColor = myCode.check_in_limit_Color(UL, LL, act_val, SV);
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

            List<string>[] _result = new List<string>[] { mylstSa1, mylstSq1, mylstSdr1, mylstSa2, mylstSq2, mylstSdr2, mylstSa3, mylstSq3, mylstSdr3 };
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
                        double Sdr_val = Convert.ToDouble(checkDBNull(tar_rgn.Offset[inx + 1, 3].Value));
                        //Sa_result[i].Add(Math.Round(Sa_val, 3).ToString());//"#0.##0"
                        //Sq_result[i].Add(Math.Round(Sq_val, 3).ToString());//"#0.##0"
                        //Sdr_result[i].Add(Math.Round(Sdr_val, 3).ToString());//"#0.##0"
                        Sa_result[i].Add(Sa_val.ToString());//"#0.##0"
                        Sq_result[i].Add(Sq_val.ToString());//"#0.##0"
                        Sdr_result[i].Add(Sdr_val.ToString());//"#0.##0"
                    }
                }
                inx++;
            }
            tar_wrkbook.Close();
            return _result;
        }
        public void Save_PTH_Diameter()
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Data_Type" }, new string[] { txtItemCode.Text, txtLotNo.Text, "" });
        start_lbl: DataTable PTH_diameter_dt = TDMK_Code.Datatable_Filter(sqlcon, "PTH_Diameter", filter_str);
            if (PTH_diameter_dt.Rows.Count > 0)
            {
                if (TDMK_Message.MessageBoxDialog("Dữ liệu đã có. Bạn muốn cộng dồn dữ liệu?") == DialogResult.No)
                {
                    TDMK_Code.Delelte_FilteredItem_arr("PTH_Diameter", sqlcon, filter_str);
                    goto start_lbl;
                }
            }
            DataTable src_dt = myCode.DGV_To_Table(DGV_DataView);
            int id = TDMK_Code.SQL_MAX("PTH_Diameter", "ID", sqlcon) + 1;
            foreach (DataColumn dc in src_dt.Columns)
            {
                foreach (DataRow dr in src_dt.Rows)
                {
                    DataRow PTH_dr = PTH_diameter_dt.NewRow();
                    PTH_dr[0] = id++;
                    PTH_dr[1] = txtItemCode.Text;
                    PTH_dr[2] = txtLotNo.Text;
                    PTH_dr[3] = dr[dc];
                    PTH_dr[4] = dc.ColumnName.Split('_').LastOrDefault();
                    PTH_dr[5] = txtOperator.Text;
                    PTH_dr[6] = Dev;
                    PTH_dr[7] = DateTime.Now.ToString();
                    PTH_dr[8] = cbType.Text.ToUpper();
                    PTH_diameter_dt.Rows.Add(PTH_dr);
                }
            }
            TDMK_Code.Delelte_FilteredItem_arr("PTH_Diameter", sqlcon, filter_str);
            myCode.BatchBulkCopy(sqlcon, PTH_diameter_dt, "PTH_Diameter");
            DGV_DataView.Columns.Clear();
            DGV_CPK.Columns.Clear();
            DGV_Data_Plus.Columns.Clear();
            DGV_Histogram.Columns.Clear();
            TDMK_Message.MessageBoxTDMK_Info("Ghi dữ liệu thành công");
        }
        public void Export_PTH_Diameter(string format_loc, string _itemcode, string _lotno, string type)
        {
            DataTable PTH_Diameter_dt = myCode.Load_PTH_Diameter_ToTable(myVar.sqlcon_SMT, "PTH_Diameter", _itemcode, _lotno);// Load_FAI_ToTable(myVar.sqlcon_SMT, "FAI_Auto", "FAI_No", "FAI_Data", _itemcode, _lotno, cbType_Sel.SelectedItem.ToString()); //myCode.Load_FAI_ToTable(myVar.sqlcon_SMT, _itemcode, _lotno, cbType_Sel.SelectedItem.ToString());
            DataTable PTH_Diameter_Spec_dt = myCode.Load_PTH_Diameter_Spec(_itemcode, myVar.sqlcon_SMT, "SPEC_PTH_Diameter");
            bool export_en = false;
            if (myCode.check_FAIdata_inSpec(PTH_Diameter_dt, PTH_Diameter_Spec_dt))
            {
                export_en = true;
            }
            else
            {
                if (TDMK_Message.MessageBoxDialog("Dữ liệu NG. Tiếp tục xuất dữ liệu?") == DialogResult.Yes)
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
                    List<string> selected_sht_name = new List<string>() { "BVH", "PTH" };
                    //string report_type = cbType_Sel.SelectedItem.ToString();
                    //DataTable FAI_Data_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Remark" }, new string[] { _itemcode, _lotno, cbType_Sel.SelectedItem.ToString() }));
                    DataTable PTH_data_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "PTH_Diameter", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Data_Type" }, new string[] { _itemcode, _lotno, type }));
                    if (PTH_data_tbl.Rows.Count > 0)
                    {
                        string format_file = format_lst[0];
                        string export_file = "";
                        string report_path = Path.Combine(report_location, type);
                        int qty = 32;
                        if (!Directory.Exists(report_path))
                        {
                            Directory.CreateDirectory(report_path);
                        }
                        ExcelPackage Format_Pack = new ExcelPackage(format_file);
                        ExcelPackage Report_Pack = new ExcelPackage();
                        ExcelWorksheet wrk_sht = null;
                        if (type.ToUpper() == "MASS")
                        {
                            qty = 6;
                            export_file = Path.Combine(report_path, Path.GetFileNameWithoutExtension(format_file) + "-" + _lotno + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");
                            //List<string> selected_sht_name = new List<string>() { "FAI", "SPC", "BVH", "PTH" };
                            bool same_ext = false;
                            if (Path.GetExtension(format_file) == Path.GetExtension(export_file))
                            {
                                same_ext = true;
                            }
                            Report_Pack = myCode.Worksheet_select(Format_Pack, selected_sht_name, same_ext);
                            wrk_sht = Report_Pack.Workbook.Worksheets[0];
                        }
                        else
                        {
                            export_file = Path.Combine(report_path, _itemcode + "-" + _lotno + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(format_file));
                            Report_Pack = Format_Pack;
                            foreach (ExcelWorksheet sht in Report_Pack.Workbook.Worksheets)
                            {
                                string sht_name = FAI_lib.remove_special_chars(sht.Name, new List<char> { '-', '&', '_', ' ', '\r', '\n' }).ToUpper();
                                if (selected_sht_name.Any(x => sht_name.Contains(x.ToUpper())))
                                {
                                    wrk_sht = sht;
                                    break;
                                }
                            }
                        }
                        int qty_export = Math.Min(qty, PTH_data_tbl.Rows.Count);
                        string start_data_addr = FAI_lib.Find_Cell_Addr("Sample", "A1", wrk_sht, false);
                        string PTH_Zone = FAI_lib.Find_Cell_Addr("Diameter seen from top", start_data_addr, wrk_sht, true);
                        string PTH_Data_Addr = FAI_lib.Find_Start_Addr(PTH_Zone, wrk_sht, false);
                        ExcelRangeBase PTH_Data_rgn = wrk_sht.Cells[PTH_Data_Addr];
                        for (int i = 0; i < qty_export; i++)
                        {
                            PTH_Data_rgn.Offset(i, 0).Value = PTH_data_tbl.Rows[i]["Data"];
                        }
                        Report_Pack.SaveAs(new FileInfo(export_file));
                        Report_Pack.Dispose();
                        Format_Pack.Dispose();
                        TDMK_Message.MessageBoxTDMK_Info("Hoàn thành xuất dữ liệu");
                        ProcessStartInfo pi = new ProcessStartInfo(export_file);
                        Process.Start(pi);
                    }
                    else
                    {
                        TDMK_Message.MessageBoxTDMK_Warning("Không tìm thấy dữ liệu của ItemCode / Lotno : " + _itemcode + " / " + _lotno);
                    }
                }
                else
                {
                    TDMK_Message.MessageBoxTDMK_Warning("Không tìm thấy Format");
                }
            }
        }

        private void btnLoadEditData_Click(object sender, EventArgs e)
        {
            if (DB_name != "OK2SHIP_Period2")
            {
                myCode.Load_FAI_DGV(DGV_DataView, myVar.sqlcon_SMT, "FAI_Auto", "FAI_No", "FAI_Data", txtItemCode.Text, txtLotNo.Text, cbType.Text.ToUpper(), false);//format_type.ToUpper()
                myCode.Load_CPK_DGV_Histogram(DGV_DataView, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
                myCode.Check_Data_inSpec2(DGV_DataView, DGV_SpecView);
                for (int r = 0; r < DGV_DataView.Rows.Count; r++)
                {
                    DGV_DataView.Rows[r].HeaderCell.Value = (r + 1).ToString();
                }
                DGV_DataView.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            }
            else
            {
                if (myCode.Load_PTH_Diameter_Spec(txtItemCode.Text, myVar.sqlcon_SMT, "SPEC_PTH_Diameter", ref DGV_SpecView))
                {
                    myCode.Load_PTH_Diameter_DGV(DGV_DataView, myVar.sqlcon_SMT, "PTH_Diameter", txtItemCode.Text, txtLotNo.Text);
                    myCode.Load_CPK_DGV_Histogram(DGV_DataView, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
                    myCode.Check_Data_inSpec2(DGV_DataView, DGV_SpecView);
                    for (int r = 0; r < DGV_DataView.Rows.Count; r++)
                    {
                        DGV_DataView.Rows[r].HeaderCell.Value = (r + 1).ToString();
                    }
                    DGV_DataView.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                }
            }
        }
        public void Load_FAI_DGV(DataGridView tar_DGV, SqlConnection tar_sqlcon, string tar_tbl, string FAI_No_name, string FAI_Data_Col_name, string tar_ItemCode, string tar_LotNo, string format_type = "", string tar_shift = "")
        {
            DataTable FAI_spec_dt = TDMK_Code.Datatable_Filter(tar_sqlcon, "FAI_Spec", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { tar_ItemCode, format_type }));
            DataTable FAI_data_dt = TDMK_Code.Datatable_Filter(tar_sqlcon, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Remark", "Shift" }, new string[] { tar_ItemCode, tar_LotNo, format_type, tar_shift }));
            List<string> temp_lst = FAI_spec_dt.AsEnumerable().Select(x => x.Field<string>("FAI_No")).Distinct().ToList();
            List<string> main_FAI = new List<string>();
            char[] split_chars = new char[] { '_', '/', '\\' };
            foreach (string t in temp_lst)
            {
                string setval = t.Split(split_chars).LastOrDefault().Trim();
                string fai_name = t.Split(split_chars).FirstOrDefault().Trim();
                string sel_FAI = fai_name + "_" + myCode.IsNumeric_Val(setval);
                if (main_FAI.IndexOf(sel_FAI) == -1)
                {
                    main_FAI.Add(sel_FAI);
                }
            }
            DataTable tbl_data = new DataTable();
            foreach (string c in main_FAI)
            {
                List<string> FAI_vals = FAI_data_dt.AsEnumerable().Where(x => x.Field<string>(FAI_No_name) == c).Select(x => x.Field<string>(FAI_Data_Col_name)).ToList();

                if (!myCode.check_columns_existed(tbl_data, c))
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
            //foreach (DataGridViewColumn t in tar_DGV.Columns)
            //{
            //    t.HeaderText = t.HeaderText.Split('_')[0];
            //}
            Disable_Sort_DGV(tar_DGV);
        }

        private void numExport_DoubleClick(object sender, EventArgs e)
        {

        }

        private void numExport_ValueChanged(object sender, EventArgs e)
        {

        }

        private void numExport_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void cbType_Sel_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbType_Sel.SelectedIndex != -1)
            {
                if (cbType_Sel.Text == "NPI")
                {
                    numExport.Value = 32;
                }
                else
                {
                    numExport.Value = 32;
                }
            }
        }

        private void lblQty_export_DoubleClick(object sender, EventArgs e)
        {
            if (numExport.Enabled)
            {
                numExport.Enabled = false;
            }
            else
            {
                numExport.Enabled = true;
            }
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
        public void Manual_Data_Process(string log_path)
        {
            if (Dev == "Manual")
            {
                if (btnRun.Text == "Stop")
                {
                    DataTable data_dt = (DataTable)(DGV_DataView.DataSource);
                    DataTable log_dt = VHX_FAI_Data_Process(log_path);
                    foreach (DataColumn dc in log_dt.Columns)
                    {
                        int tar_col_inx = -1;
                        if (myCode.check_columns_existed_inx(data_dt, dc.ColumnName, ref tar_col_inx))
                        {
                            for (int r_inx = 0; r_inx < log_dt.Rows.Count; r_inx++)
                            {
                                if (r_inx < data_dt.Rows.Count)
                                {
                                    data_dt.Rows[r_inx][tar_col_inx] = log_dt.Rows[r_inx][dc];
                                }
                            }
                        }
                    }
                    //DGV_DataView.DataSource= null  ;
                    DGV_DataView.DataSource = data_dt;
                    myCode.Load_CPK_DGV_Histogram(DGV_DataView, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
                }
                else
                {
                    MessageBox.Show("Ấn nút Start để bắt đầu", "Thông báo");
                }
            }
        }
        private void txtDataFolder_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                string _log_path = txtDataFolder.Text;
                if (_log_path != "")
                {
                    string log_path = TDMK_Code.remove_char(_log_path, new string[] { '\r'.ToString(), '\n'.ToString() });
                    Manual_Data_Process(log_path);
                    txtDataFolder.Text = log_path;
                }
            }
        }
        private List<FAI_Spec> testFAI_spec = new List<FAI_Spec>();
    }
}
