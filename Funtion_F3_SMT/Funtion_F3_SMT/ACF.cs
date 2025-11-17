using Funtion_F3_SMT;
using IniLibs;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OK2SHIP_SMT.Libary;
using OK2SHIP_SMT.Repositories;
using OK2SHIP_SMT.Services;
using OK2SHIP_SMT.ToolBoxs;
using OK2SHIP_SMT.UserControls;
using OK2SHIP_SMT.Views;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using TDMK_SEEV_DLL;
using TDMK_SQL;
using DataTable = System.Data.DataTable;
using myExcel = Microsoft.Office.Interop.Excel;

namespace OK2SHIP_SMT
{
    public partial class ACF : Form
    {

        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        //public EPPlus_Lib TDMK_EPPLUS = new EPPlus_Lib();

        public SEI_Lib myCode = new SEI_Lib();
        string admin_mode = "LOGIN";
        public SqlConnection sqlcon = null;
        public string strcon = "";
        string DB_name = "OK2SHIP_SMT";
        string data_loc = "";
        string spec_roughness = "Sa(<0.5um)^Sq(<0.75um)^Sdr(>0.0075)";
        bool judgeall_rougness = true;
        bool judgeall_peel = true;
        string server_name = "";
        string server_acc = "";
        string server_pass = "";
        IniFile TDMK_init;
        Funtion_export_file_le_EPPlus F_export_EPPlus = new Funtion_export_file_le_EPPlus();

        bool hide_mode_wetting = true;
        public ACF()
        {
            InitializeComponent();
        }

        private void ACF_Load(object sender, EventArgs e)
        {
            UpdateLogin();
            string app_path = System.Windows.Forms.Application.StartupPath;
            //app_path = @"\\10.212.6.212\Saomai\QA\TDMK_DATA\Test_Areas\FPCA OK2SHIP Auto System(temp2)\VHX-IMADA";
            string program_loc = F_export_EPPlus.find_config_path(app_path, "TDMK Program");
            string config_path = Path.Combine(program_loc, "Config.ini");
            //string config_path = Path.Combine(app_path.Replace(@"\VHX-IMADA", ""), "Config.ini");
            TDMK_init = new IniFile(config_path);
            data_loc = TDMK_init.Read("Format_Folder", "SMT_Config") + "\\SEEV Data";
            //data_loc = F_export_EPPlus.find_config_path(app_path.Replace("\\FPCA OK2SHIP Auto System\\VHX-IMADA", ""), "Format_Folder");
            //data_loc = Path.Combine(System.Windows.Forms.Application.StartupPath.Replace(@"\VHX-IMADA", ""));
            server_name = TDMK_init.Read("Server", "SMT_Config");
            server_acc = TDMK_init.Read("Account", "SMT_Config");
            server_pass = TDMK_init.Read("Password", "SMT_Config");
            sqlcon = initial_data(DB_name, true);
            //data_loc = data_loc.Replace(@"\VHX-IMADA", "");


        }
        public SqlConnection initial_data_old(string DB_name, bool sa_en)
        {
            SqlConnection _sqlcon_OK2SHIP;
            string app_path = System.Windows.Forms.Application.StartupPath;
            string config_file = Path.Combine(app_path, "Config", "config.txt");
            string[] my_config = myCode.read_config_arr(config_file);
            string server_name = "";
            string server_acc = "";
            string server_pass = "";


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

            return _sqlcon_OK2SHIP;
        }

        public SqlConnection initial_data(string DB_name, bool sa_en)
        {
            SqlConnection _sqlcon_OK2SHIP;
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
        public void insert_row_excel(myExcel.Worksheet ws, int rcopy_begin, int rcopy_end)
        {
            // Get the rows to copy (rows 4157 to 4178 in your example).
            myExcel.Range copyRange = ws.Rows[rcopy_begin + ":" + rcopy_end];

            // Insert enough new rows to fit the rows we're copying.
            copyRange.Insert(myExcel.XlInsertShiftDirection.xlShiftDown);

            // The copied data will be put in the same place (starting at row 4157 in
            // your example).
            myExcel.Range dest = ws.Rows[rcopy_begin + ":" + rcopy_end];

            copyRange.Copy(dest);
        }


        private void btn_load_Click(object sender, EventArgs e)
        {

        }

        public string get_value_cell(object src_str)
        {
            if (src_str != null)
            {
                return src_str.ToString();
            }

            return "";
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
        public void Hide_column(DataGridView dgv, ref bool en_)
        {
            string[] col_hide = { "ItemCode", "LotNo", "Sheet", "Operator", "Time_Update", "Remark" };
            if (en_)
            {
                for (int i = 0; i < dgv.Columns.Count; i++)
                {
                    if (col_hide.Contains(dgv.Columns[i].Name))
                    {
                        dgv.Columns[i].Visible = false;
                    }
                }
            }
            else
            {
                for (int i = 0; i < dgv.Columns.Count; i++)
                {
                    if (col_hide.Contains(dgv.Columns[i].Name))
                    {
                        dgv.Columns[i].Visible = true;
                    }
                }
            }

            en_ = !en_;

        }

        public void highlight_row_old(DataGridView dgv)
        {
            int k = 0;
            foreach (DataGridViewRow dgv_row in dgv.Rows)
            {
                for (int i = 0; i < dgv.Columns.Count; i++)
                {
                    if (k % 2 == 0)
                    {
                        dgv_row.Cells[i].Style.BackColor = Color.LightBlue;
                    }
                    else
                    {

                        dgv_row.Cells[i].Style.BackColor = Color.White;
                    }
                }
                k++;
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

        public double CalculateStandardDeviation(IEnumerable<double> values)
        {
            double result = 0.0;
            if (values.Any())
            {
                double avg = values.Average();
                double num = values.Sum((double d) => Math.Pow(d - avg, 2.0));
                result = Math.Sqrt(num / (double)(values.Count() - 1));
            }

            return result;
        }
        public void fill_dgv_spec_wetting()
        {

        }
        private void FillData()
        {
            DateTime dateTime = DateTime.Now;
            foreach (string key in dic.Keys)
            {
                if (dic_list.TryGetValue(key, out List<DateTime> list))
                {

                }
                else
                {
                    // Tạo một List<DateTime> gồm 6 phần tử, mỗi phần tử cách nhau 10 ngày
                    List<DateTime> dateList = new List<DateTime>();
                    DateTime startDate = DateTime.Now;
                    for (int i = 0; i < 6; i++)
                    {
                        dateList.Add(startDate.AddDays(i * 10));
                    }
                    dic_list.Add(key, dateList);
                }
            }
            listBox1.Items.Clear();
            listBox1.Items.AddRange(dic.Keys.ToArray());
            dataGridView.DataSource = new DataTable();

        }
        private Dictionary<string, DataTable> dic = new Dictionary<string, DataTable>();
        private Dictionary<string, List<DateTime>> dic_list = new Dictionary<string, List<DateTime>>();
        private void btn_loadwetting_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtOperator.Text != "")
                {
                    UC_AddDataMachine uc = new UC_AddDataMachine();
                    uc.DICTIONARY = dic;
                    NormalForm form = new NormalForm(uc);
                    form.ShowDialog();

                    dic = uc.DICTIONARY;
                    FillData();
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Please, fill in ItemCode / LotNo / Operator / Type / Logfile Location", "Warning");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ACF_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //    OpenFileDialog f_open = new OpenFileDialog(); 
            //    f_open.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            //    f_open.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            //    if (f_open.ShowDialog() == DialogResult.OK)
            //    {
            //        if (f_open.FileName != "")
            //        {
            //            txtLogfile_wetting.Text = f_open.FileName;
            //        }

            //}
        }

        private void txtLogfile_wetting_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //OpenFileDialog f_open = new OpenFileDialog();
            //f_open.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            //f_open.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            //if (f_open.ShowDialog() == DialogResult.OK)
            //{
            //    if (f_open.FileName != "")
            //    {
            //        txtLogfile_wetting.Text = f_open.FileName;
            //    }

            //}
        }
        public string find_format(string in_data_loc, string ItemCode)
        {
            //string result = "";
            //string[] files = Directory.GetFiles(Path.Combine(in_data_loc, "Format"), "*" + ItemCode + "*.xlsm");
            //if (files.Length > 0)
            //{
            //    result = files[0];
            //}
            //return result;
            string result = "";
            if (System.IO.Directory.Exists(Path.Combine(in_data_loc, "Format", cb_Type.SelectedItem.ToString())))
            {
                string path = "";
                if (cb_Type.SelectedItem.ToString() == "NPI")
                {
                    path = Path.Combine(in_data_loc, "Format", cb_Type.SelectedItem.ToString());
                }
                else
                {
                    path = Path.Combine(in_data_loc, "Format", cb_Type.SelectedItem.ToString(), "ACF");
                }

                if (System.IO.Directory.Exists(path))
                {
                    string[] file_xlsm = Directory.GetFiles(path, "*" + ItemCode + "*.xlsm");

                    if (file_xlsm.Length > 0)
                    {
                        result = file_xlsm[0];
                    }
                    else
                    {
                        string[] file_xlsx = Directory.GetFiles(path, "*" + ItemCode + "*.xlsx");
                        if (file_xlsx.Length > 0)
                        {
                            result = file_xlsx[0];
                            foreach (string f in file_xlsx)
                            {
                                if (f.ToUpper().Replace(" ", "_").Contains("ACF"))
                                {
                                    result = f;
                                }
                            }
                        }
                    }
                }
                else
                {
                    Directory.CreateDirectory(path);
                }


            }
            return result;
        }
        //public string find_format(string in_data_loc, string ItemCode)
        //{

        //    string result = "";
        //    string[] file_xlsm = Directory.GetFiles(Path.Combine(in_data_loc, "Format"), "*" + ItemCode + "*.xlsm");
        //    if (file_xlsm.Length > 0)
        //    {
        //        result = file_xlsm[0];

        //    }
        //    else
        //    {
        //        string[] file_xlsx = Directory.GetFiles(Path.Combine(in_data_loc, "Format"), "*" + ItemCode + "*.xlsx");
        //        if (file_xlsx.Length > 0)
        //        {
        //            result = file_xlsx[0];

        //            foreach (string f in file_xlsx)
        //            {
        //                if (f.ToUpper().Replace(" ", "_").Contains("ACF"))
        //                {
        //                    result = f;
        //                }
        //            }
        //        }
        //    }
        //    return result;
        //}

        public void Export_ACFFlatness(myExcel.Worksheet ws, DataTable tbl_data)
        {
            DataTable src_dt = tbl_data;


            if (src_dt.Rows.Count > 0)
            {
                myExcel.Range curr_rgn = ws.Range["A1"];
                for (int i = 0; i < 100; i++)
                {
                    if (curr_rgn.Offset[i, 0].Value != null)
                    {
                        if (curr_rgn.Offset[i, 0].Value.ToString().Replace(" ", "").ToUpper().Contains("ACF Flatness".Replace(" ", "").ToUpper()) && !curr_rgn.Offset[i, 0].Value.ToString().Replace(" ", "").ToUpper().Contains("Measurement".Replace(" ", "").ToUpper()))
                        {
                            curr_rgn = curr_rgn.Offset[i, 0];
                            //break;
                            //for(int r = 1; r < 5; r++)
                            //{
                            for (int c = 1; c < 5; c++)
                            {
                                if (myCode.checkDBNull(curr_rgn.Offset[1, c].Offset[-1, 0].Value).Contains("Point"))
                                {
                                    curr_rgn = curr_rgn.Offset[1, c];
                                    goto lbl_export;
                                }
                            }
                            // }
                        }
                    }
                }
            lbl_export:
                for (int i = 0; i < 5; i++)
                {
                    for (int j = 0; j < 12; j++)
                    {
                        curr_rgn.Offset[i, j].Value = src_dt.Rows[i][4 + j].ToString();
                    }
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "ACF_Flatness: No data", "Warning");
            }

        }

        public List<List<string>> lst_roughness(DataTable dt)
        {
            List<List<string>> lst_roughness = new List<List<string>>() { };
            string[] arr_name = { "L1_Roughness_Sa", "L1_Roughness_Sq", "L1_Roughness_Sdr", "L2_Roughness_Sa", "L2_Roughness_Sq", "L2_Roughness_Sdr", "L3_Roughness_Sa", "L3_Roughness_Sq", "L3_Roughness_Sdr" };
            foreach (var item in arr_name)
            {
                lst_roughness.Add(dt.AsEnumerable().Select(x => x.Field<string>(item)).ToList());
            }

            return lst_roughness;
        }

        public string ACF_pad_location(myExcel.Worksheet ws)
        {

            string get_info = "";
            myExcel.Range curr_rgn_1 = ws.Range["A1"];

            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 10; i < 80; i++)
            {
                if (curr_rgn_1.Offset[i, 0].Value != null)
                {
                    if (curr_rgn_1.Offset[i, 0].Value.ToString().Replace(" ", "").ToUpper() == "ACF pad location".Replace(" ", "").ToUpper())
                    {
                        int k = 1;
                        get_info = i.ToString();
                        while (curr_rgn_1.Offset[i + k, 1].Value.ToString() != "Roughness")
                        {
                            if (curr_rgn_1.Offset[i + k, 1].Value.ToString().Contains("Sa"))
                            {

                                get_info += "+ sa";
                            }
                            if (curr_rgn_1.Offset[i + k, 1].Value.ToString().Contains("Sq"))
                            {

                                get_info += "+ sq";
                            }
                            if (curr_rgn_1.Offset[i + k, 1].Value.ToString().Contains("Sdr"))
                            {

                                get_info += "+ sdr";
                            }
                            k++;
                        }


                        get_info += "+" + k.ToString();
                        break;
                    }
                }
            }
            return get_info;


        }

        public string ACF_pad_location_mass(myExcel.Worksheet ws)
        {

            string get_info = "";
            myExcel.Range curr_rgn_1 = ws.Range["A1"];

            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 10; i < 50; i++)
            {
                if (curr_rgn_1.Offset[i, 0].Value != null)
                {
                    if (curr_rgn_1.Offset[i, 0].Value.ToString().Replace(" ", "").ToUpper() == "ACF pad location".Replace(" ", "").ToUpper())
                    {
                        int k = 1;
                        get_info = i.ToString();
                        while (myCode.checkDBNull(curr_rgn_1.Offset[i + k, 1].Value) != "")
                        {
                            if (myCode.checkDBNull(curr_rgn_1.Offset[i + k, 1].Value).Contains("Sa"))
                            {
                                get_info += "+ sa";
                            }
                            if (myCode.checkDBNull(curr_rgn_1.Offset[i + k, 1].Value).Contains("Sq"))
                            {
                                get_info += "+ sq";
                            }
                            if (myCode.checkDBNull(curr_rgn_1.Offset[i + k, 1].Value).Contains("Sdr"))
                            {
                                get_info += "+ sdr";
                            }
                            k++;
                        }

                        get_info += "+" + k.ToString();
                        break;
                    }
                }
            }
            return get_info;


        }

        public void check_data_bonding_mass(DataGridView dgv)
        {
            if (cb_Type.SelectedIndex != -1 && cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, "ACF", "MASS" }));
                for (int i = 0; i < dgv.Rows.Count; i++)
                {
                    dgv.Rows[i].Cells["Data"].Style.BackColor = Color.White;
                }
                if (dt_spec.Rows.Count > 0)
                {
                    string[] spec_all = dt_spec.Rows[0]["Location"].ToString().Split('^');
                    //  string[] region = ((DataTable)dgv.DataSource).AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                    int k = 0;
                    foreach (string spec in spec_all)
                    {
                        if (spec != "")
                        {
                            double R = 0;
                            double UCL = 0;
                            double LCL = 0;

                            if (myCode.IsNumeric(spec.Split(';')[0]))
                            {
                                R = double.Parse(spec.Split(';')[0]);
                            }
                            if (myCode.IsNumeric(spec.Split(';')[1]))
                            {
                                UCL = double.Parse(spec.Split(';')[1]);
                            }
                            if (myCode.IsNumeric(spec.Split(';')[2]))
                            {
                                LCL = double.Parse(spec.Split(';')[2]);
                            }

                            SortedDictionary<int, string> dic_data = new SortedDictionary<int, string> { };
                            List<double> lst_data = new List<double> { };
                            for (int i = k; i < k + 5 && i < dgv.Rows.Count; i++)
                            {
                                string val = dgv.Rows[i].Cells["Data"].Value.ToString().Replace(" ", "");
                                if (myCode.IsNumeric(val))
                                {
                                    dic_data.Add(i, val);
                                    lst_data.Add(double.Parse(val));
                                    if (Double.Parse(val) > UCL && UCL != 0)
                                    {
                                        dgv.Rows[i].Cells["Data"].Style.BackColor = Color.Red;

                                    }
                                    if (Double.Parse(val) < LCL && LCL != 0)
                                    {
                                        dgv.Rows[i].Cells["Data"].Style.BackColor = Color.Red;

                                    }

                                }
                            }

                            if (R != 0)
                            {
                                if (lst_data.Count > 0)
                                {
                                    double tb = lst_data.Average();
                                    foreach (int r1 in dic_data.Keys)
                                    {
                                        foreach (int r2 in dic_data.Keys)
                                        {
                                            if (r2 < r1)
                                            {
                                                double sub_data = Math.Abs(double.Parse(dic_data[r1]) - double.Parse(dic_data[r2]));
                                                if (sub_data > R && tb != 0)
                                                {

                                                    double a1 = Math.Abs(double.Parse(dic_data[r1]) - tb);
                                                    double a2 = Math.Abs(double.Parse(dic_data[r2]) - tb);
                                                    if (a1 > a2)
                                                    {
                                                        dgv.Rows[r1].Cells["Data"].Style.BackColor = Color.Red;
                                                    }
                                                    else
                                                    {
                                                        dgv.Rows[r2].Cells["Data"].Style.BackColor = Color.Red;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                            }

                            k = k + 5;
                        }
                    }
                }
            }
        }


        public void Export_ACF_roughness_mass(myExcel.Worksheet ws, DataTable tbl_data)
        {
            List<List<string>> lst_rghness = lst_roughness(tbl_data);

            myExcel.Range curr_rgn_1 = ws.Range["A1"];

            List<int> vitri_sa = new List<int> { };
            List<int> vitri_sq = new List<int> { };
            List<int> vitri_sdr = new List<int> { };
            string[] arr_info = ACF_pad_location_mass(ws).Split('+');
            string vitri_ACF_pad = arr_info[0];

            int row = int.Parse(arr_info[0]);
            int r_count = int.Parse(arr_info[arr_info.Length - 1]);
            curr_rgn_1 = curr_rgn_1.Offset[row + 1, 2];


            for (int i = 1; i < arr_info.Length; i++)
            {
                if (arr_info[i].Contains("sa"))
                {
                    vitri_sa.Add(i);
                }
                if (arr_info[i].Contains("sq"))
                {
                    vitri_sq.Add(i);
                }
                if (arr_info[i].Contains("sdr"))
                {
                    vitri_sdr.Add(i);
                }
            }

            int count = 0;
            for (int i = 1; i < 100; i++)
            {
                for (int j = 1; j < 5; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Contains("Sample"))
                    {
                        while (myCode.checkDBNull(ws.Cells[i, j + count].Value).Contains("Sample"))
                        {
                            count++;
                        }
                        goto lbl_continue;


                    }
                }
            }

        lbl_continue:

            int count_sample = new int[] { count, lst_rghness[0].Count }.Min();

            for (int m = 0; m < count_sample; m++)
            {
                int k = 0;
                foreach (int j in vitri_sa)
                {
                    curr_rgn_1.Offset[j, m].Value = lst_rghness[k][m];
                    k += 3;

                }
                k = 1;
                foreach (int j in vitri_sq)
                {
                    curr_rgn_1.Offset[j, m].Value = lst_rghness[k][m];
                    k += 3;

                }
                k = 2;
                foreach (int j in vitri_sdr)
                {
                    curr_rgn_1.Offset[j, m].Value = lst_rghness[k][m];
                    k += 3;

                }
            }
            //int count_sample = tbl_data.Rows.Count;

            int count_offset = vitri_sa.Count + vitri_sq.Count + vitri_sdr.Count + 1;

            List<string> lst_judge = new List<string> { };
            for (int i = 0; i < count_sample; i++)
            {
                bool check = true;
                for (int c = 6; c < tbl_data.Columns.Count; c++)
                {
                    if (dgv_roughness_data.Rows[i].Cells[c].Style.BackColor == Color.Red)
                    {
                        check = false;
                        break;
                    }

                }
                if (!check)
                {
                    lst_judge.Add("Fail");
                    judgeall_rougness = false;

                }
                else
                {
                    lst_judge.Add("Pass");
                }
            }


            for (int smp = 0; smp < count_sample; smp++)
            {
                curr_rgn_1.Offset[count_offset, smp].Value = lst_judge[smp];
            }

            //for(int i = 1; i < count_offset; i++)
            //{
            //    curr_rgn_1.Offset[i, count_sample].FormulaR1C1 = "=MIN(R[0]C[-" + count_sample.ToString() + "]:R[0]C[-1])";
            //    curr_rgn_1.Offset[i, count_sample+1].FormulaR1C1 = "=MAX(R[0]C[-" + (count_sample +1).ToString() + "]:R[0]C[-2])";
            //    curr_rgn_1.Offset[i, count_sample + 2].FormulaR1C1 = "=AVERAGE(R[0]C[-" + (count_sample + 2).ToString() + "]:R[0]C[-3])";
            //    curr_rgn_1.Offset[i, count_sample + 3].FormulaR1C1 = "=STDEV(R[0]C[-" + (count_sample + 3).ToString() + "]:R[0]C[-4])";
            //    curr_rgn_1.Offset[i, count_sample + 4].FormulaR1C1 = "=(0.5 - (R[0]C[-2])/( 3 * R[0]C[-1])";
            //}


            if (cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                myExcel.Range cur_rgn_judge = find_cell(ws, "JUDGEMENT");
                if (cur_rgn_judge != null)
                {
                    if (judgeall_rougness)
                    {
                        cur_rgn_judge.Offset[1, 0].Value = "OK";
                        ws.Cells[1, 1].Value = "OK";
                    }
                    else
                    {
                        cur_rgn_judge.Offset[1, 0].Value = "NG";
                        ws.Cells[1, 1].Value = "NG";
                    }
                }
            }
        }


        public void Export_ACF_Roughness(myExcel.Worksheet ws, DataTable tbl_data)
        {
            DataTable src_dt = new DataTable();
            //if (txtItemCode_FPC.Text != "" && txtLotNo_FPC.Text != "")
            //{
            //    SqlConnection sqlcon_IPQC = initial_data("IPQC_Data", true);
            //    src_dt = TDMK_Code.Datatable_Filter(sqlcon_IPQC, "Roughness", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtLotNo_FPC.Text, txtLotNo_FPC.Text }));
            //}
            //else if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtItemCode_FPC.Text == "" && txtLotNo_FPC.Text == "")
            //{
            src_dt = tbl_data;
            //}

            if (src_dt.Rows.Count > 0)
            {
                List<List<string>> lst_rghness = lst_roughness(src_dt);

                myExcel.Range curr_rgn_1 = ws.Range["A1"];

                List<int> vitri_sa = new List<int> { };
                List<int> vitri_sq = new List<int> { };
                List<int> vitri_sdr = new List<int> { };
                string[] arr_info = ACF_pad_location(ws).Split('+');
                string vitri_ACF_pad = arr_info[0];

                int row = int.Parse(arr_info[0]);
                int r_count = int.Parse(arr_info[arr_info.Length - 1]);
                curr_rgn_1 = curr_rgn_1.Offset[row, 0];

                for (int i = 1; i < arr_info.Length; i++)
                {
                    if (arr_info[i].Contains("sa"))
                    {
                        vitri_sa.Add(i);
                    }
                    if (arr_info[i].Contains("sq"))
                    {
                        vitri_sq.Add(i);
                    }
                    if (arr_info[i].Contains("sdr"))
                    {
                        vitri_sdr.Add(i);
                    }
                }

                int m = 0;
                int k;
                while (m < 32)
                {
                    for (int c = 0; c < 4; c++)
                    {
                        for (int i = 2; i <= 9; i++)
                        {
                            k = 0;
                            foreach (int j in vitri_sa)
                            {
                                curr_rgn_1.Offset[j, i].Value = lst_rghness[k][m];
                                k += 3;
                            }
                            k = 1;
                            foreach (int j in vitri_sq)
                            {
                                curr_rgn_1.Offset[j, i].Value = lst_rghness[k][m];
                                k += 3;
                            }
                            k = 2;
                            foreach (int j in vitri_sdr)
                            {
                                curr_rgn_1.Offset[j, i].Value = lst_rghness[k][m];
                                k += 3;
                            }
                            m++;
                        }

                        curr_rgn_1 = curr_rgn_1.Offset[r_count, 0];

                    }
                }

            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "ACF_Roughness: No data", "Warning");
            }
        }
        public myExcel.Workbook create_export_wrk(string format_file, string process_name)
        {
            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            xlsApp.DisplayAlerts = false;
            myExcel.Workbook wb = TDMK_Code.open_excel_file(format_file, "", "");
            if (wb != null)
            {
                string _process = process_name.Replace("_", "").Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "").ToUpper();
                string mySheet = "";
                foreach (myExcel.Worksheet tg_sht in wb.Worksheets)
                {
                    string cur_sht_name = tg_sht.Name.ToUpper().Replace("_", "").Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "");
                    if (_process == cur_sht_name)
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
            else
            {
                return null;
            }
        }


        private void txtLogfile_wetting_TextChanged(object sender, EventArgs e)
        {

        }
        public string Find_addr(string in_item, myExcel.Range start_rgn)
        {
            //string result = start_rgn.AddressLocal.ToString();
            string result = "";
            int r_inx = 0;
            while (r_inx < 30)
            {
                if (myCode.checkDBNull(start_rgn.Offset[r_inx, 0].Value).Replace(" ", "").ToUpper().Contains(in_item.Replace(" ", "").ToUpper().Replace("_", "").Replace("ACF", "").Replace("IMAGE", "")))
                {
                    result = start_rgn.Offset[r_inx, 0].AddressLocal;
                    break;
                }
                r_inx++;
            }
            return result;
        }

        public string Find_addr_peel(string in_item, myExcel.Range start_rgn)
        {
            //string result = start_rgn.AddressLocal.ToString();
            string result = "";
            int r_inx = 0;
            while (r_inx < 60)
            {
                if (myCode.checkDBNull(start_rgn.Offset[r_inx, 0].Value).Replace(" ", "").ToUpper().Contains(in_item.Replace(" ", "").ToUpper()))
                {
                    result = start_rgn.Offset[r_inx, 0].AddressLocal;
                    break;
                }
                r_inx++;
            }
            return result;
        }

        public void InsertPicture_Name(myExcel.Worksheet tar_wrksht, myExcel.Range tar_range, string picFile, int margin)
        {
            float left;
            float top;
            float width;
            float height;
            string pic_name = Path.GetFileName(picFile);

            if (tar_range.MergeCells)
            {
                myExcel.Range refer_range = tar_range.MergeArea;//Range["G16"];
                left = (float)(tar_range.Left) + margin;
                top = (float)(tar_range.Top) + margin;
                width = (float)(refer_range.Width) - 2 * margin;
                height = (float)(refer_range.Height) - 2 * margin;
                //width = (float)(refer_range.Width * 2.8) + 2 * margin;
                //height = (float)(refer_range.Height * 8.9) + 2 * margin;
            }
            else
            {
                left = (float)tar_range.Left + margin;
                top = (float)tar_range.Top + margin;
                width = (float)tar_range.Width - 2 * margin;
                height = (float)tar_range.Height - 2 * margin;
            }
            //  myExcel.Shape sel_picture = tar_wrksht.Shapes.AddPicture(picFile, MsoTriState.msoFalse, MsoTriState.msoTrue, left, top, width, height);
            myExcel.Shape sel_picture = tar_wrksht.Shapes.AddPicture2(picFile, MsoTriState.msoFalse, MsoTriState.msoTrue, left, top, width, height, MsoPictureCompress.msoPictureCompressFalse);
            sel_picture.LockAspectRatio = MsoTriState.msoTrue;
            sel_picture.Placement = myExcel.XlPlacement.xlMoveAndSize;
            sel_picture.Name = pic_name;

        }
        public void Export_DatatableImage_Excel_Graph(DataTable dt, string Image_Col_name, myExcel.Worksheet tar_wrksht, myExcel.Range sel_rgn, bool row_offset)
        {
            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = Path.Combine(file_folder, DateTime.Now.ToString("ddMMMyy_HHmmss") + ".jpg");
            int r_inx = 0;
            // myExcel.Range sel_rgn = tar_wrksht.Range[tar_rgn];
            myExcel.Range rgn_begin = sel_rgn;
            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {

                byte[] img_byte = (byte[])(dr[Image_Col_name]);
                using (MemoryStream ms = new MemoryStream(img_byte))
                {
                    Image _temp = Image.FromStream(ms);
                    //Image temp = ChangeColor((Bitmap)_temp, Color.FromArgb(243, 219, 203));
                    _temp.Save(file_dic);
                }
                InsertPicture_Name(tar_wrksht, sel_rgn, file_dic, 5);
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

                count++;
            }

            try
            {
                File.Delete(file_dic);
            }
            catch
            {
            }

        }
        public Bitmap ChangeColor(Bitmap scrBitmap, System.Drawing.Color newColor)
        {
            System.Drawing.Color actualColor;
            Color ref_color = Color.FromArgb(255, 255, 255);
            Bitmap newBitmap = new Bitmap(scrBitmap.Width, scrBitmap.Height);

            for (int i = 0; i < scrBitmap.Width; i++)
            {
                for (int j = 0; j < scrBitmap.Height; j++)
                {
                    actualColor = scrBitmap.GetPixel(i, j);

                    if (actualColor.R > 240 && actualColor.G > 240 && actualColor.B > 240)
                    {
                        newBitmap.SetPixel(i, j, newColor);
                    }
                    else
                    {
                        newBitmap.SetPixel(i, j, actualColor);
                    }
                }
            }
            return newBitmap;

        }

        public void Export_DatatableImage_Excel(DataTable dt, string Image_Col_name, myExcel.Worksheet tar_wrksht, myExcel.Range sel_rgn, bool row_offset)
        {
            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = Path.Combine(file_folder, DateTime.Now.ToString("ddMMMyy_HHmmss") + ".jpg");
            int r_inx = 0;
            // myExcel.Range sel_rgn = tar_wrksht.Range[tar_rgn];
            myExcel.Range rgn_begin = sel_rgn;

            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                byte[] img_byte = (byte[])(dr[Image_Col_name]);
                using (MemoryStream ms = new MemoryStream(img_byte))
                {
                    Image temp = Image.FromStream(ms);
                    temp.Save(file_dic);
                }

                InsertPicture_Name(tar_wrksht, sel_rgn, file_dic, 5);
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
                count++;

            }

            try
            {
                File.Delete(file_dic);
            }
            catch
            {


            }

        }
        public void Export_ACF_Peel(myExcel.Worksheet ws, DataTable tbl_Data)
        {
            myExcel.Range sel_rgn = ws.Range["A1"];
            // string find_addr = Find_addr("Before tape test", sel_rgn);
            string find_addr = Find_addr_peel("ACF Bonding", sel_rgn);
            if (find_addr != "")
            {
                myExcel.Range rgn_ = ws.Range[find_addr];
                myExcel.Range cur_rgn = ws.Range[find_addr].Offset[1, 1];
                if (myCode.checkDBNull(rgn_.Offset[1, 0].Value).Contains("Product"))
                    cur_rgn = cur_rgn.Offset[1, 0];


                if (cur_rgn != null)
                {
                    Export_DatatableImage_Excel(tbl_Data, "Image_Before", ws, cur_rgn, false);
                    Export_DatatableImage_Excel(tbl_Data, "Image_After", ws, cur_rgn.Offset[1, 0], false);
                    // Export_DatatableImage_Excel(tbl_Data, "Graph", ws, cur_rgn.Offset[2, 0], false);
                    Export_DatatableImage_Excel_Graph(tbl_Data, "Graph", ws, cur_rgn.Offset[2, 0], false);
                    List<string> lst_data = tbl_Data.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                    for (int i = 0; i < tbl_Data.Rows.Count; i++)
                    {
                        cur_rgn.Offset[3, i].Value = lst_data[i];
                    }


                    int count_sample = tbl_Data.Rows.Count;

                    List<string> lst_judge = new List<string> { };
                    for (int i = 0; i < count_sample; i++)
                    {

                        if (dgv_peel.Rows[i].Cells["Data"].Style.BackColor == Color.Red)
                        {
                            lst_judge.Add("Fail");
                            judgeall_peel = false;

                        }
                        else
                        {
                            lst_judge.Add("Pass");
                        }
                    }


                    for (int i = 1; i < 13; i++)
                    {
                        if (myCode.checkDBNull(cur_rgn.Offset[i, -1].Value).Replace(" ", "").ToUpper().Contains("MINFORCE"))
                        {

                            if (count_sample == 5)
                            {
                                myExcel.Range cell_min = cur_rgn.Offset[i, 0];
                                myExcel.Range cell_max = cur_rgn.Offset[i + 1, 0];
                                myExcel.Range cell_ave = cur_rgn.Offset[i + 2, 0];

                                int c_offset = count_sample;
                                cell_min.FormulaR1C1 = "=MIN(R[-" + (i - 3).ToString() + "]C[0]:R[-" + (i - 3).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                cell_max.FormulaR1C1 = "=MAX(R[-" + (i - 2).ToString() + "]C[0]:R[-" + (i - 2).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                cell_ave.FormulaR1C1 = "=AVERAGE(R[-" + (i - 1).ToString() + "]C[0]:R[-" + (i - 1).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                            }
                            else if (count_sample == 15)
                            {
                                for (int k = 0; k < 3; k++)
                                {
                                    if (k == 0)
                                    {
                                        myExcel.Range cell_min = cur_rgn.Offset[i, 5 * k];
                                        myExcel.Range cell_max = cur_rgn.Offset[i + 1, 5 * k];
                                        myExcel.Range cell_ave = cur_rgn.Offset[i + 2, 5 * k];

                                        int c_offset = 5;
                                        cell_min.FormulaR1C1 = "=MIN(R[-" + (i - 3).ToString() + "]C[0]:R[-" + (i - 3).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                        cell_max.FormulaR1C1 = "=MAX(R[-" + (i - 2).ToString() + "]C[0]:R[-" + (i - 2).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                        cell_ave.FormulaR1C1 = "=AVERAGE(R[-" + (i - 1).ToString() + "]C[0]:R[-" + (i - 1).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                    }
                                    else
                                    {
                                        myExcel.Range cell_min = cur_rgn.Offset[i, 5 * k + 1];
                                        myExcel.Range cell_max = cur_rgn.Offset[i + 1, 5 * k + 1];
                                        myExcel.Range cell_ave = cur_rgn.Offset[i + 2, 5 * k + 1];

                                        int c_offset = 5;
                                        cell_min.FormulaR1C1 = "=MIN(R[-" + (i - 3).ToString() + "]C[-1]:R[-" + (i - 3).ToString() + "]C[" + (c_offset - 2).ToString() + "])";
                                        cell_max.FormulaR1C1 = "=MAX(R[-" + (i - 2).ToString() + "]C[-1]:R[-" + (i - 2).ToString() + "]C[" + (c_offset - 2).ToString() + "])";
                                        cell_ave.FormulaR1C1 = "=AVERAGE(R[-" + (i - 1).ToString() + "]C[-1]:R[-" + (i - 1).ToString() + "]C[" + (c_offset - 2).ToString() + "])";
                                    }

                                }
                            }

                            break;
                        }
                    }

                    if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                    {
                        for (int smp = 0; smp < count_sample; smp++)
                        {
                            cur_rgn.Offset[5, smp].Value = lst_judge[smp];
                        }
                    }

                    if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                    {
                        myExcel.Range cur_rgn_judge = find_cell(ws, "JUDGEMENT");
                        if (cur_rgn_judge != null)
                        {
                            if (judgeall_peel)
                            {
                                cur_rgn_judge.Offset[1, 0].Value = "OK";
                                ws.Cells[1, 1].Value = "OK";
                            }
                            else
                            {
                                cur_rgn_judge.Offset[1, 0].Value = "NG";
                                ws.Cells[1, 1].Value = "NG";
                            }
                        }
                    }
                }

            }

            //  }
            //else
            //{
            //    MessageBox.Show(new Form { TopMost = true }, "ACF_Bonding: No data", "Warning");
            //}
        }

        public void export_ACF_Wetting(myExcel.Worksheet ws)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "")
            {
                DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet" }, new string[] { txtItemCode.Text, "ACF" }));
                if (dt_spec.Rows.Count > 0)
                {
                    int count_sample = int.Parse(dt_spec.Rows[0]["Count_sample"].ToString());
                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet" }, new string[] { txtItemCode.Text, "ACF" });

                    DataTable Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, "ACF_WETTING", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                    if (Data_tbl.Rows.Count > 0)
                    {
                        List<string> lst_machine = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Machine")).Distinct().ToList();
                        int count_machine_DB = lst_machine.Count();
                        int count_machine_format = dt_spec.Rows[0]["Location"].ToString().Split('_').Length - 1;

                        int r_begin = int.Parse(dt_spec.Rows[0]["Location"].ToString().Split('_')[0].Split(';')[0]);
                        int c_begin = int.Parse(dt_spec.Rows[0]["Location"].ToString().Split('_')[0].Split(';')[1]);

                        Range line = (Range)ws.Rows[r_begin];
                        line.Insert();
                        if (count_machine_DB > count_machine_format)
                        {
                            insert_row_excel(ws, r_begin, r_begin + 6);
                        }
                        foreach (string machine in lst_machine)
                        {

                            Dictionary<int, List<Double>> dic_list = new Dictionary<int, List<double>> { };
                            List<double> lst_after_plasma_ACF = new List<double> { };
                            List<double> lst_after_plasma_GND = new List<double> { };
                            List<double> lst_Before_Packing_ACF = new List<double> { };
                            List<double> lst_Before_Packing_GND = new List<double> { };


                            ws.Cells[r_begin, 1].Value = "Result data machine " + machine;

                            DataTable tbl_machine = TDMK_Code.Datatable_Filter(sqlcon, "ACF_WETTING", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Machine" }, new string[] { txtItemCode.Text, txtLotNo.Text, machine }));
                            int col_data = c_begin + 9;
                            for (int i = 0; i < tbl_machine.Rows.Count; i = i + 2)
                            {
                                if (col_data < col_data + count_sample)
                                {

                                    if (myCode.checkDBNull(tbl_machine.Rows[i]["After_Plasma"]) != "")
                                    {
                                        ws.Cells[r_begin + 2, col_data].Value = tbl_machine.Rows[i]["After_Plasma"].ToString();
                                        lst_after_plasma_ACF.Add(double.Parse(myCode.checkDBNull(myCode.checkDBNull(tbl_machine.Rows[i]["After_Plasma"]))));
                                    }

                                    if (myCode.checkDBNull(tbl_machine.Rows[i + 1]["After_Plasma"]) != "")
                                    {
                                        ws.Cells[r_begin + 3, col_data].Value = tbl_machine.Rows[i + 1]["After_Plasma"].ToString();
                                        lst_after_plasma_GND.Add(double.Parse(myCode.checkDBNull(tbl_machine.Rows[i + 1]["After_Plasma"])));
                                    }

                                    if (myCode.checkDBNull(tbl_machine.Rows[i]["Before_Packing"]) != "")
                                    {
                                        ws.Cells[r_begin + 4, col_data].Value = tbl_machine.Rows[i]["Before_Packing"].ToString();
                                        lst_Before_Packing_ACF.Add(double.Parse(myCode.checkDBNull(tbl_machine.Rows[i]["Before_Packing"])));
                                    }

                                    if (myCode.checkDBNull(tbl_machine.Rows[i + 1]["Before_Packing"]) != "")
                                    {
                                        ws.Cells[r_begin + 5, col_data].Value = tbl_machine.Rows[i + 1]["Before_Packing"].ToString();
                                        lst_Before_Packing_GND.Add(double.Parse(myCode.checkDBNull(tbl_machine.Rows[i + 1]["Before_Packing"])));
                                    }

                                    col_data++;
                                }
                            }
                            dic_list.Add(2, lst_after_plasma_ACF);
                            dic_list.Add(3, lst_after_plasma_GND);
                            dic_list.Add(4, lst_Before_Packing_ACF);
                            dic_list.Add(5, lst_Before_Packing_GND);


                            foreach (var lst in dic_list)
                            {
                                string max = lst.Value.Max().ToString();
                                string min = lst.Value.Min().ToString();
                                string Average = Math.Round(lst.Value.Average(), 2).ToString();
                                double STDEV = CalculateStandardDeviation(lst.Value);

                                string _spec = ws.Cells[r_begin + lst.Key, c_begin + 2].Value;
                                double Mean = double.Parse(_spec.Replace("<", "").Replace(">", "").Replace("°", "").Replace(" ", ""));
                                double cpk = Math.Round((Mean - lst.Value.Average()) / (3 * STDEV), 2);

                                ws.Cells[r_begin + lst.Key, c_begin + 3].Value = min;
                                ws.Cells[r_begin + lst.Key, c_begin + 4].Value = max;
                                ws.Cells[r_begin + lst.Key, c_begin + 5].Value = Average;
                                ws.Cells[r_begin + lst.Key, c_begin + 6].Value = STDEV;
                                ws.Cells[r_begin + lst.Key, c_begin + 7].Value = cpk;
                                if (cpk > 1.33)
                                {
                                    ws.Cells[r_begin + lst.Key, c_begin + 8].Value = "OK";
                                }
                                else
                                {
                                    ws.Cells[r_begin + lst.Key, c_begin + 8].Value = "NG";
                                }

                            }

                            r_begin = r_begin + 7;
                        }
                    }
                }


            }
        }

        public string find_format_mass(string in_data_loc, string ItemCode)
        {
            string result = "";
            if (System.IO.Directory.Exists(in_data_loc))
            {
                string[] file_xlsm = Directory.GetFiles(in_data_loc, "*" + ItemCode + "*.xlsm");

                if (file_xlsm.Length > 0)
                {
                    result = file_xlsm[0];
                }
                else
                {
                    string[] file_xlsx = Directory.GetFiles(in_data_loc, "*" + ItemCode + "*.xlsx");
                    if (file_xlsx.Length > 0)
                    {
                        result = file_xlsx[0];

                    }
                }

            }
            result = result.Replace("~$", "");
            return result;
        }
        //public Workbook Create_workbook()
        //{
        //    Microsoft.Office.Interop.Excel.Application application = ((!TDMK_Code.IsExcelRunning()) ? ((Microsoft.Office.Interop.Excel.Application)Interaction.CreateObject("Excel.Application", "")) : ((Microsoft.Office.Interop.Excel.Application)Interaction.GetObject((string)null, "Excel.Application")));
        //    Workbook result = application.Workbooks.Add(RuntimeHelpers.GetObjectValue(Missing.Value));
        //    result = application.Workbooks.Add(MsoPictureCompress.msoPictureCompressFalse);
        //    application.Visible = true;
        //    return result;
        //}

        private void btn_export_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "")
            {
                if (cb_Type.SelectedIndex != -1)
                {
                    //btnLoaddb_wetting.PerformClick();
                    if (cb_Type.SelectedItem.ToString() == "NPI")
                    {
                        string file_format = find_format(data_loc, txtItemCode.Text);

                        if (file_format != "")
                        {
                            string report_folder = Path.Combine(data_loc, "Report", cb_Type.SelectedItem.ToString(), "ACF");
                            if (!System.IO.Directory.Exists(report_folder))
                                System.IO.Directory.CreateDirectory(report_folder);
                            F_export_EPPlus.export_NPI_ACF(file_format, report_folder, txtItemCode.Text, txtLotNo.Text, "ACF", sqlcon, !Legacy.Checked);
                        }
                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Format of " + txtItemCode.Text + " not found", "Thông báo");
                        }
                    }


                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Please select type of format to export", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Please fill in ItemCode(SMT)", "Thông báo");
            }

        }



        public void BatchBulkCopy(SqlConnection sqlcon, DataTable dataTable, string tablename)
        {
            DataTable dtInsertRows = dataTable;
            using (SqlBulkCopy sbc = new SqlBulkCopy(sqlcon))
            {
                if (sqlcon.State != ConnectionState.Open)
                {
                    sqlcon.Open();
                }
                sbc.DestinationTableName = tablename;
                foreach (DataColumn dc in dataTable.Columns)
                {
                    sbc.ColumnMappings.Add(dc.ColumnName, dc.ColumnName);
                }
                sbc.WriteToServer(dtInsertRows);
                sqlcon.Close();
            }
        }

        private void btn_save_wetting_Click(object sender, EventArgs e)
        {
            try
            {
                new ACFService().Save(txtItemCode.Text, txtLotNo.Text, dic, dic_list, cb_Type.Text);
                MessageBox.Show("Save successfully!");

            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                if (msg.Split('_').Count() >= 2 && msg.Split('_')[0].Contains("1231"))
                {
                    bool prime = MessageBox.Show($"{msg.Split('_')[1]}", "Thông báo!", MessageBoxButtons.YesNo) == DialogResult.Yes;
                    new ACFService().Save(txtItemCode.Text, txtLotNo.Text, dic, dic_list, cb_Type.Text, prime);
                }
                else
                {
                    MessageBox.Show(msg);
                }
            }
        }

        public void resize_column_image(DataGridView dgv)
        {
            ((DataGridViewImageColumn)dgv.Columns["Image_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            ((DataGridViewImageColumn)dgv.Columns["Image_Data"]).Width = 100;

            foreach (DataGridViewRow dr in dgv.Rows)
            {
                dr.Height = 70;
            }

        }

        public void check_roughness(DataGridView dgv)
        {
            string[] arr_spec = spec_roughness.Split('^');
            double sa = 0;
            double sq = 0;
            double sdr = 0;
            foreach (string spec in arr_spec)
            {
                if (spec.Contains("Sa"))
                {
                    sa = double.Parse(spec.Replace(" ", "").Replace("Sa(", "").Replace(")", "").Replace("um", "").Replace("<", ""));
                }
                if (spec.Contains("sq"))
                {
                    sq = double.Parse(spec.Replace(" ", "").Replace("Sq(", "").Replace(")", "").Replace("um", "").Replace("<", ""));
                }
                if (spec.Contains("Sdr"))
                {
                    sdr = double.Parse(spec.Replace(" ", "").Replace("Sdr(", "").Replace(")", "").Replace("um", "").Replace(">", ""));
                }
            }
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                for (int j = 0; j < dgv.Columns.Count; j++)
                {
                    if (dgv.Columns[j].Name.Contains("Sa") && sa != 0)
                    {

                        if (myCode.IsNumeric(myCode.checkDBNull(dgv.Rows[i].Cells[j].Value)))
                        {
                            double data = double.Parse(myCode.checkDBNull(dgv.Rows[i].Cells[j].Value));
                            if (data >= sa)
                            {
                                dgv.Rows[i].Cells[j].Style.BackColor = Color.Red;
                            }
                            else
                            {
                                dgv.Rows[i].Cells[j].Style.BackColor = Color.White;
                            }

                        }
                    }

                    if (dgv.Columns[j].Name.Contains("Sq") && sq != 0)
                    {
                        if (myCode.IsNumeric(myCode.checkDBNull(dgv.Rows[i].Cells[j].Value)))
                        {
                            double data = double.Parse(myCode.checkDBNull(dgv.Rows[i].Cells[j].Value));
                            if (data >= sq)
                            {
                                dgv.Rows[i].Cells[j].Style.BackColor = Color.Red;
                            }
                            else
                            {
                                dgv.Rows[i].Cells[j].Style.BackColor = Color.White;
                            }

                        }
                    }

                    if (dgv.Columns[j].Name.Contains("Sdr") && sdr != 0)
                    {

                        if (myCode.IsNumeric(myCode.checkDBNull(dgv.Rows[i].Cells[j].Value)))
                        {
                            double data = double.Parse(myCode.checkDBNull(dgv.Rows[i].Cells[j].Value));
                            if (data <= sdr)
                            {
                                dgv.Rows[i].Cells[j].Style.BackColor = Color.Red;
                            }
                            else
                            {
                                dgv.Rows[i].Cells[j].Style.BackColor = Color.White;
                            }

                        }
                    }
                }
            }
        }

        public void check_roughness_cell(DataGridView dgv, int row_indx, int col_indx)
        {
            string[] arr_spec = spec_roughness.Split('^');
            double sa = 0;
            double sq = 0;
            double sdr = 0;
            foreach (string spec in arr_spec)
            {
                if (spec.Contains("Sa"))
                {
                    sa = double.Parse(spec.Replace(" ", "").Replace("Sa(", "").Replace(")", "").Replace("um", "").Replace("<", ""));
                }
                if (spec.Contains("Sq"))
                {
                    sq = double.Parse(spec.Replace(" ", "").Replace("Sq(", "").Replace(")", "").Replace("um", "").Replace("<", ""));
                }
                if (spec.Contains("Sdr"))
                {
                    sdr = double.Parse(spec.Replace(" ", "").Replace("Sdr(", "").Replace(")", "").Replace("um", "").Replace(">", ""));
                }
            }
            int i = row_indx;
            int j = col_indx;

            if (dgv.Columns[j].Name.Contains("Sa") && sa != 0)
            {

                if (myCode.checkDBNull(dgv.Rows[i].Cells[j].Value) != "")
                {
                    try
                    {

                        double data = double.Parse(myCode.checkDBNull(dgv.Rows[i].Cells[j].Value));
                        if (data >= sa)
                        {
                            dgv.Rows[i].Cells[j].Style.BackColor = Color.Red;
                        }
                        else
                        {
                            dgv.Rows[i].Cells[j].Style.BackColor = Color.White;
                        }

                    }
                    catch
                    {
                        return;
                    }
                }
            }

            if (dgv.Columns[j].Name.Contains("Sq") && sq != 0)
            {

                if (myCode.checkDBNull(dgv.Rows[i].Cells[j].Value) != "")
                {
                    double data = double.Parse(myCode.checkDBNull(dgv.Rows[i].Cells[j].Value));
                    if (data >= sq)
                    {
                        dgv.Rows[i].Cells[j].Style.BackColor = Color.Red;
                    }
                    else
                    {
                        dgv.Rows[i].Cells[j].Style.BackColor = Color.White;
                    }

                }
            }

            if (dgv.Columns[j].Name.Contains("Sdr") && sdr != 0)
            {

                if (myCode.checkDBNull(dgv.Rows[i].Cells[j].Value) != "")
                {
                    double data = double.Parse(myCode.checkDBNull(dgv.Rows[i].Cells[j].Value));
                    if (data <= sdr)
                    {
                        dgv.Rows[i].Cells[j].Style.BackColor = Color.Red;
                    }
                    else
                    {
                        dgv.Rows[i].Cells[j].Style.BackColor = Color.White;
                    }

                }
            }

        }

        public void Load_data_SMT(string tbl_name, DataGridView dgv_data)
        {
            string itemCode = txtItemCode.Text.ToString();
            string lotNo = txtLotNo.Text.ToString();
            if (tbl_name == "ACF_FLATNESS" && !Legacy.Checked)
            {
                tbl_name = "ACF_FLATNESS_NAS";
                itemCode = itemCode.PadRight(10);
                lotNo = lotNo.PadRight(10);
            }
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemCode, lotNo });

            DataTable dt_analysis = tbl_name != "ACF_BONDING" ? TDMK_Code.Datatable_Filter(sqlcon, tbl_name, filter_str) : new DataTable();

            if (tbl_name == "ACF_BONDING")
            {
                dt_analysis = new ACFService().loadPeel(txtItemCode.Text, txtLotNo.Text, Legacy.Checked);

            }

            if (tbl_name == "Roughness")
            {
                itemCode = itemCode.PadRight(10);
                lotNo = lotNo.PadRight(10);
                filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemCode, lotNo });

                dt_analysis = TDMK_Code.Datatable_Filter(sqlcon, tbl_name + "_NAS", filter_str);
                if(dt_analysis.Rows.Count > 0)
                {
                    string json = dt_analysis.Rows[0]["Data"].ToString();
                   dt_analysis = ConverterService.JsonToDataTable(json);
                }
            }

            if (tbl_name == "ACF_FLATNESS_NAS")
            {
                string json = dt_analysis.Rows[0]["Data"].ToString();
                dt_analysis = ConverterService.JsonToDataTable(json);
            }
            if (dt_analysis.Rows.Count > 0)
            {
                if (!dt_analysis.Columns.Contains("ProductID"))
                {
                    try
                    {

                        ProductIDService.FillProductID(dt_analysis, txtItemCode.Text, txtLotNo.Text, tbl_name);
                    }
                    catch { }
                }
                int ID = 1;
                foreach (DataRow dr in dt_analysis.Rows)
                {
                    dr["ID"] = ID;
                    ID++;
                }

                dgv_data.DataSource = dt_analysis;
                if (tbl_name == "ACF_BONDING")
                {
                    if (dgv_data.Rows.Count > 0)
                    {
                        resize_column_image_bonding(dgv_data);
                    }

                }

                myCode.Disable_Sort_DGV(dgv_data);
                if (tbl_name == "Roughness")
                {
                    check_roughness(dgv_data);
                }
            }
            else
            {
                throw new Exception(tbl_name + ": Không có dữ liệu");
            }
        }


        private void btnLoaddb_wetting_Click(object sender, EventArgs e)
        {
            try
            {
                ACFService service = new ACFService();
                txt_ItemName.Text = service.loadItemNamebyItemCode(txtItemCode.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            #region WCA
            try
            {
                KeyValuePair<Dictionary<string, DataTable>, Dictionary<string, List<DateTime>>> z = new ACFService().LoadWCA(txtItemCode.Text, txtLotNo.Text, cb_Type.Text);
                dic = z.Key;
                dic_list = z.Value;
                listBox1.Items.Clear();
                listBox1.Items.AddRange(dic.Keys.ToArray());
                dataGridView.DataSource = new DataTable();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"WCA: {ex.Message}");
            }
            #endregion
            #region FLATNESS
            try
            {
                Load_data_SMT("ACF_FLATNESS", dgv_Flatness);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ACF_FLATNESS: {ex.Message}");
            }
            try
            {
                Load_data_SMT("ACF_BONDING", dgv_peel);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ACF_BONDING: {ex.Message}");
            }
            try
            {
                Load_data_SMT("Roughness", dgv_roughness_data);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Roughness: {ex.Message}");
            }

            #endregion
        }


        private void dgv_Wetting_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //highlight_row(dgv_Wetting);
        }

        private void btnEdit_wetting_Click(object sender, EventArgs e)
        {

        }


        private void UpdateLogin()
        {
            if (UserSession.Instance.IsLoggedIn)
            {
                lbl_Login.Text = "Logout";
                lbl_Login.ForeColor = Color.Red;
                txtOperator.Text = UserSession.Instance.User_ID;
            }
            else
            {
                lbl_Login.Text = "Login";
                lbl_Login.ForeColor = Color.SteelBlue;
                txtOperator.Text = "";
            }
        }
        private void lbl_Login_Click(object sender, EventArgs e)
        {
            if (UserSession.Instance.IsLoggedIn)
            {
                UserSession.Instance.Logout();
            }
            else
            {
                Login fr1 = new Login();
                fr1.ShowDialog();
            }
            UpdateLogin();
        }

        private void dgv_Wetting_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {

        }

        private void dgv_Wetting_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void dgv_roughness_data_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            if (e.RowIndex != -1 && e.ColumnIndex != -1)
            {
                if (dgv_roughness_data.Columns[e.ColumnIndex].Name.Contains("Roughness"))
                {
                    string a = "Sa(<0.5um)^Sq(<0.75um)^Sdr(>0.0075)";
                    dgv_roughness_data.CurrentCell.ToolTipText = a.Replace("^", "\n");
                }

            }
        }

        private void dgv_roughness_data_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1 && e.ColumnIndex != -1)
            {
                if (dgv_roughness_data.Columns[e.ColumnIndex].Name.Contains("Roughness"))
                {
                    check_roughness_cell(dgv_roughness_data, e.RowIndex, e.ColumnIndex);
                }

            }

        }

        private void txtLotNo_Validated(object sender, EventArgs e)
        {
            txtLotNo.Text = Lotno_Formated(txtLotNo.Text);
        }

        private void txtLotNo_Mass_Validated(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void cb_Type_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_Type.SelectedIndex != -1)
            {
                txt_qty_peel.Text = "";
                DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, "ACF", cb_Type.SelectedItem.ToString() }));
                if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                {
                    dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, "ACF", "MASS" }));
                }

                if (dt_spec.Rows.Count == 0 && cb_Type.SelectedItem.ToString() == "NPI")
                {
                    MessageBox.Show(new Form { TopMost = true }, "Chưa cài đặt format", "Thông báo");
                }
                //else
                //{ 
                if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                {
                    btn_export_bonding.Visible = true;
                    btn_export_roughness.Visible = true;
                    btn_export.Visible = false;


                    if (dt_spec.Rows.Count > 0)
                    {
                        string q = dt_spec.Rows[0]["Count_sample"].ToString();
                        if (q == "15")
                        {
                            txt_qty_peel.Text = "5";
                        }
                        else if (q == "5" || q == "32")
                        {
                            txt_qty_peel.Text = q;
                        }


                        if (dgv_peel.DataSource != null)
                        {
                            check_data_bonding_mass(dgv_peel);
                        }
                    }
                }
                else
                {

                    btn_export_bonding.Visible = false;
                    btn_export_roughness.Visible = false;
                    btn_export.Visible = true;

                }
                //}
            }
        }

        public double flatness(Double[] arr_sample)
        {
            double a = arr_sample.Min();
            double b = arr_sample.Max();
            return b - a;
        }
        public void check_Flatness(DataGridView dgv)
        {
            double val = 0;
            if (dgv.Rows.Count > 0)
            {
                for (int row = 0; row < dgv.Rows.Count; row++)
                {
                    List<double> arr_sample = new List<double> { };
                    for (int i = 4; i < dgv.Columns.Count - 2; i++)
                    {
                        double value;
                        if (Double.TryParse(dgv.Rows[row].Cells[i].Value.ToString(), out value))
                        {
                            arr_sample.Add(value);
                        }
                        else
                        {
                            break;
                        }
                    }
                    if (arr_sample.Count > 0)
                    {
                        val = flatness(arr_sample.ToArray());
                        dgv.Rows[row].Cells["Flatness"].Value = val.ToString();
                        if (val < 15)
                        {
                            dgv.Rows[row].Cells["Judgement"].Value = "OK";
                            dgv.Rows[row].Cells["Judgement"].Style.BackColor = Color.White;
                        }
                        else
                        {
                            dgv.Rows[row].Cells["Judgement"].Value = "NG";
                            dgv.Rows[row].Cells["Judgement"].Style.BackColor = Color.Red;
                        }
                    }

                }

            }
        }

        public void check_Flatness_inrow(DataGridView dgv, int row)
        {
            double val = 0;
            if (dgv.Rows.Count > 0)
            {
                List<double> arr_sample = new List<double> { };
                for (int i = 4; i < dgv.Columns.Count - 3; i++)
                {
                    double value;
                    if (Double.TryParse(dgv.Rows[row].Cells[i].Value.ToString(), out value))
                    {
                        arr_sample.Add(value);
                    }
                    else
                    {
                        break;
                    }
                }
                if (arr_sample.Count > 0)
                {
                    val = flatness(arr_sample.ToArray());
                    dgv.Rows[row].Cells["Flatness"].Value = val.ToString();
                    if (val < 15)
                    {
                        dgv.Rows[row].Cells["Judgement"].Value = "OK";
                        dgv.Rows[row].Cells["Judgement"].Style.BackColor = Color.White;
                    }
                    else
                    {
                        dgv.Rows[row].Cells["Judgement"].Value = "NG";
                        dgv.Rows[row].Cells["Judgement"].Style.BackColor = Color.Red;
                    }
                }

            }
        }


        private void btnLoad_Flatness_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtLogfile_Flatness.Text != "" && txtOperator.Text != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
                DataTable Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, "ACF_Flatness", filter_str).Clone();
                ExcelWorkbook wb = TDMK_EPPLUS.open_excel_file(txtLogfile_Flatness.Text);
                ExcelWorksheet ws = wb.Worksheets[0];
                for (int i = 1; i < 20; i++)
                {
                    for (int j = 1; j < 10; j++)
                    {
                        if (myCode.checkDBNull(ws.Cells[i, j].Value).Replace(" ", "").ToUpper().Contains("POINT"))
                        {
                            ExcelRange curr_rng = ws.Cells[i + 1, j];
                            for (int k = 0; k < 5; k++)
                            {
                                DataRow dr = Data_tbl.NewRow();
                                dr[0] = k + 1;
                                dr[1] = txtItemCode.Text;
                                dr[2] = txtLotNo.Text;
                                dr[3] = "Sample " + (k + 1).ToString();
                                for (int t = 4; t < 14; t++)
                                {
                                    dr[t] = myCode.checkDBNull(curr_rng.Offset(k, t - 4).Value).ToString();
                                }
                                dr[14] = "";
                                dr[15] = "";

                                Data_tbl.Rows.Add(dr);
                            }
                            if (Data_tbl.Rows.Count > 0)
                            {
                                try
                                {

                                    new ACFService().GetProductID(txtItemCode.Text, txtLotNo.Text, textBox1.Text, "Flatness", Data_tbl);
                                }
                                catch
                                {
                                    MessageBox.Show("ProductID Null");
                                }
                            }
                            dgv_Flatness.DataSource = Data_tbl;
                            myCode.Disable_Sort_DGV(dgv_Flatness);
                            check_Flatness(dgv_Flatness);
                            return;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Vui lòng điền đầy đủ thông tin ItemCode / LotNo / Operator / Type / Logfile Location", "Warning");
            }
        }

        private void txtLogfile_Flatness_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            OpenFileDialog f_open = new OpenFileDialog();
            f_open.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            f_open.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            if (f_open.ShowDialog() == DialogResult.OK)
            {
                if (f_open.FileName != "")
                {
                    txtLogfile_Flatness.Text = f_open.FileName;
                }

            }
        }

        private void dgv_Flatness_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            check_Flatness_inrow(dgv_Flatness, e.RowIndex);
        }

        private void txtLogfile_Flatness_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnsave_Flatness_Click(object sender, EventArgs e)
        {
            try
            {

                if (txtItemCode.Text != "" && txtLotNo.Text != "" && dgv_Flatness.DataSource != null)
                {
                    DataTable tbl_data = (DataTable)dgv_Flatness.DataSource;

                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                lblsave:
                    DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "ACF_FLATNESS_NAS", filter_str);
                    if (dt.Rows.Count == 0)
                    {
                        if (tbl_data.Rows.Count > 0)
                        {
                            int i = TDMK_Code.SQL_MAX("ACF_FLATNESS", "ID", sqlcon) + 1;
                            foreach (DataRow dr in tbl_data.Rows)
                            {
                                dr[0] = i;
                                i++;
                            }
                            if (tbl_data.Columns.Contains("ProductID"))
                            {
                                string pid = ProductIDService.ConverterProductID(tbl_data, "Id");
                                tbl_data.Columns.Remove("ProductID");
                                ProductIDService.InsertProductID(txtItemCode.Text, txtLotNo.Text, "ACF_FLATNESS", pid);
                            }
                            DataRow row = dt.NewRow();
                            row["Data"] = ConverterService.DataTableToJson(tbl_data);
                            row["ItemCode"] = txtItemCode.Text;
                            row["LotNo"] = txtLotNo.Text;
                            dt.Rows.Add(row);
                            new DBContext().BuckDataTable(dt, "ACF_FLATNESS_NAS", new[] { "ItemCode", "LotNo" }, null, "ID");
                            //BatchBulkCopy(sqlcon, (DataTable)dgv_Flatness.DataSource, "ACF_FLATNESS");
                            MessageBox.Show(new Form { TopMost = true }, "Lưu dữ liệu thành công!", "Warning");
                        }
                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Không có dữ liệu", "Warning");
                        }
                    }
                    else
                    {

                        if (MessageBox.Show(new Form { TopMost = true }, "Dữ liệu đã tồn tại. Bạn có muốn cập nhật không?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            if (UserSession.Instance.IsLoggedIn)
                            {
                                TDMK_Code.Delelte_FilteredItem_arr("ACF_FLATNESS", sqlcon, filter_str);
                                goto lblsave;
                            }
                            else
                            {
                                MessageBox.Show("Vui lòng đăng nhập để cập nhật dữ liệu", "Warning");

                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error Save: {ex.Message}");
            }
        }

        private void txtItemCode_TextChanged(object sender, EventArgs e)
        {
            txt_qty_peel.Text = "";
            //txtLogfile_Flatness.Text = txtLogfile_peel.Text = txtLogfile_wetting.Text = txt_logfile_roughness.Text = "";

        }

        public void Get_Image_comment3(string in_src, ref SortedDictionary<int, byte[]> lst_result_sorted)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            SortedDictionary<int, byte[]> lst_result = new SortedDictionary<int, byte[]> { };
            FileInfo[] temp_lst = tar_d.GetFiles("*.jpg").Concat(tar_d.GetFiles("*.jpeg")).ToArray();
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    var sel_img = Bitmap.FromFile(temp_lst[i].FullName);
                    ImageConverter imgcon = new ImageConverter();
                    byte[] img_data = (byte[])imgcon.ConvertTo(sel_img, typeof(byte[]));
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });

                    char[] ch_arr = f_na.ToCharArray();
                    for (int t = 0; t < ch_arr.Length; t++)
                    {
                        string a = ch_arr[t].ToString();
                        if (!myCode.IsNumeric(a))
                        {
                            f_na = f_na.Replace(a, "");
                        }
                    }

                    if (!lst_result.ContainsKey(Convert.ToInt32(f_na)))
                    {
                        lst_result.Add(Convert.ToInt32(f_na), img_data);
                    }

                }
                int k = 0;

                foreach (var item in lst_result)
                {
                    lst_result_sorted.Add(k, item.Value);
                    k++;
                }
            }

        }

        //public string get_data_val_comment3(ExcelWorksheet ws, string textfind)
        //{
        //    string val = "";
        //    // string a = myCode.checkDBNull(ws.Cells[].Value);
        //    for (int i = 5; i < 30; i++)
        //    {
        //        for (int j = 4; j < 15; j++)
        //        {
        //            if (myCode.checkDBNull(ws.Cells[i, j].Value).Replace(" ", string.Empty).ToUpper().Contains(textfind.Replace(" ", string.Empty).ToUpper()))
        //            {
        //                for (int k = 1; k < 5; k++)
        //                {
        //                    if (myCode.checkDBNull(ws.Cells[i, j + k].Value) != "")
        //                    {
        //                        val = myCode.checkDBNull(ws.Cells[i, j + k].Value);
        //                        break;
        //                    }
        //                }
        //                //val = myCode.checkDBNull(ws.Cells[i, j + 1].Value);



        //                return val;
        //            }
        //        }
        //    }
        //    return val;
        //}

        public string get_data_val_comment3(ExcelWorksheet ws, string textfind)
        {
            string val = "";

            for (int i = 5; i < 30; i++)
            {
                for (int j = 4; j < 15; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Replace(" ", string.Empty).ToUpper().Contains(textfind.Replace(" ", string.Empty).ToUpper()))
                    {
                        string cur_addr = ws.Cells[i, j].Address;
                        string tar_addr = get_offset_addr(cur_addr, ws, 1, true);
                        val = myCode.checkDBNull(ws.Cells[tar_addr].Value);
                        if (!myCode.IsNumeric(val))
                        {
                            string xml = @"<root>" + myCode.checkDBNull(ws.Cells[tar_addr].Value) + @"</root>";

                            XmlDocument xmlDoc = new XmlDocument();
                            xmlDoc.LoadXml(xml);
                            XmlNodeList textNodes = xmlDoc.GetElementsByTagName("t");

                            foreach (XmlNode textNode in textNodes)
                            {
                                string textValue = textNode.InnerText.Trim();
                                if (myCode.IsNumeric(textValue))
                                {
                                    val = textValue;
                                    break;
                                }

                            }
                        }
                        return val;

                    }
                }
            }
            return val;
        }
        public string get_offset_addr(string start_addr, ExcelWorksheet tar_wrksht, int offset_val, bool left_to_right)
        {
            string result = start_addr;
            ExcelRangeBase tar_rgn = tar_wrksht.Cells[start_addr];
            int inx = 0;
            while (inx < offset_val)
            {
                int r_count = tar_rgn.End.Row;
                int col_count = tar_rgn.End.Column;
                var idx = tar_wrksht.GetMergeCellId(r_count, col_count);
                string mergedCellAddress = tar_rgn.Address;
                if (idx > 0)
                {
                    mergedCellAddress = tar_wrksht.MergedCells[idx - 1];
                }
                int r = tar_wrksht.Cells[mergedCellAddress].Rows;
                int c = tar_wrksht.Cells[mergedCellAddress].Columns;
                if (left_to_right)
                {
                    tar_rgn = tar_rgn.Offset(0, c);
                }
                else
                {
                    tar_rgn = tar_rgn.Offset(r, 0);
                }
                inx++;
            }
            result = tar_rgn.Address;
            return result;
        }

        public Image get_image_excel(ExcelWorksheet wrk_sheet)
        {
            Byte[] data = new Byte[0];
            Image myImg = TDMK_EPPLUS.get_pic(wrk_sheet, "Picture 1");// Clipboard.GetImage();
            return myImg;
        }
        public void Get_Image_bonding(string in_src, ref SortedDictionary<int, Image> lst_result_sorted)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            SortedDictionary<int, Image> lst_result = new SortedDictionary<int, Image> { };
            FileInfo[] temp_lst = tar_d.GetFiles("*.jpg").Concat(tar_d.GetFiles("*.jpeg")).ToArray();
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    var sel_img = Bitmap.FromFile(temp_lst[i].FullName);
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });

                    char[] ch_arr = f_na.ToCharArray();
                    for (int t = 0; t < ch_arr.Length; t++)
                    {
                        string a = ch_arr[t].ToString();
                        if (!myCode.IsNumeric(a))
                        {
                            f_na = f_na.Replace(a, "");
                        }
                    }

                    if (!lst_result.ContainsKey(Convert.ToInt32(f_na)))
                    {
                        lst_result.Add(Convert.ToInt32(f_na), sel_img);
                    }
                }

                int k = 1;
                foreach (var item in lst_result)
                {
                    lst_result_sorted.Add(k, item.Value);
                    k++;
                }
            }
        }

        public void get_data_bonding(DirectoryInfo tar_d, ref SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_lst_result)
        {
            SortedDictionary<int, byte[]> dic_image = new SortedDictionary<int, byte[]>();
            Get_Image_comment3(tar_d.FullName, ref dic_image);
            FileInfo[] temp_lst = tar_d.GetFiles("*.xlsx");

            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
                    f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                    if (myCode.IsNumeric(f_na))
                    {
                        ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                        ExcelWorksheet wrksht = wrkbk.Worksheets[0];

                        Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                        data.data_val = get_data_val_comment3(wrksht, "Max");
                        data.grap_data = get_image_excel(wrksht);

                        int f_inx = Convert.ToInt32(f_na);
                        if (!dic_lst_result.ContainsKey(f_inx))
                        {
                            dic_lst_result.Add(f_inx, data);
                        }
                    }
                }

            }
        }

        public void resize_column_image_bonding(DataGridView dgv)
        {
            ((DataGridViewImageColumn)dgv.Columns["Image_Before"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            ((DataGridViewImageColumn)dgv.Columns["Image_Before"]).Width = 100;

            ((DataGridViewImageColumn)dgv.Columns["Image_After"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            ((DataGridViewImageColumn)dgv.Columns["Image_After"]).Width = 100;

            ((DataGridViewImageColumn)dgv.Columns["Graph"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            ((DataGridViewImageColumn)dgv.Columns["Graph"]).Width = 100;

            foreach (DataGridViewRow dr in dgv.Rows)
            {
                dr.Height = 70;
            }

        }

        public DataTable load_data_logfile_peel(string in_src)
        {
            if (string.IsNullOrEmpty(txt_ItemName.Text))
            {
                MessageBox.Show("Thiếu ITEMNAME");
            }
            int qty = int.Parse(txt_qty_peel.Text);

            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
            DataTable Data_tbl = ACFService.getStructorPeel();
            DirectoryInfo tar_parent = new DirectoryInfo(in_src);
            DirectoryInfo[] arr_dir_child = tar_parent.GetDirectories();
            string[] ListFolderName = new string[arr_dir_child.Length];

            for (int i = 0; i < arr_dir_child.Length; i++)
            {
                ListFolderName[i] = arr_dir_child[i].Name.ToUpper().Replace(" ", "");
            }

            if (arr_dir_child.Length > 0)
            {
                Dictionary<string, SortedDictionary<int, Image>> lst_result = new Dictionary<string, SortedDictionary<int, Image>> { };
                Dictionary<string, SortedDictionary<int, Funtion_SMT.Peeltest_data>> dic_grp_data_all = new Dictionary<string, SortedDictionary<int, Funtion_SMT.Peeltest_data>> { };
                foreach (DirectoryInfo tar_d2 in arr_dir_child)
                {
                    if (tar_d2.Name.ToUpper() == "BEFORE" || tar_d2.Name.ToUpper() == "AFTER")
                    {
                        SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image> { };
                        Get_Image_bonding(tar_d2.FullName, ref dic_image);
                        lst_result.Add(tar_d2.Name.ToUpper(), dic_image);

                    }
                    else if (tar_d2.Name.ToUpper() == "A" || tar_d2.Name.ToUpper() == "B" || tar_d2.Name.ToUpper() == "C")
                    {
                        SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_grp_data = new SortedDictionary<int, Funtion_SMT.Peeltest_data> { };
                        get_data_bonding(tar_d2, ref dic_grp_data);
                        dic_grp_data_all.Add(tar_d2.Name.ToUpper(), dic_grp_data);
                    }
                }

                if (dic_grp_data_all.Count == 0)
                {
                    SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_grp_data = new SortedDictionary<int, Funtion_SMT.Peeltest_data> { };
                    get_data_bonding(tar_parent, ref dic_grp_data);
                    dic_grp_data_all.Add("A", dic_grp_data);
                }

                int ID = 1;
                foreach (var val in dic_grp_data_all)
                {
                    int count_sample = 1;
                    foreach (var log in val.Value)
                    {
                        if (count_sample <= qty && lst_result["BEFORE"].ContainsKey(ID) && lst_result["AFTER"].ContainsKey(ID))
                        {
                            Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, count_sample, lst_result["BEFORE"][ID], lst_result["AFTER"][ID], log.Value.grap_data, log.Value.data_val, txtOperator.Text, DateTime.Now.ToString(), txt_ItemName.Text + "_" + txt_date_peel.Text + "_" + txt_worker_peel.Text + "_" + cb_Type.SelectedItem.ToString());
                            count_sample++;
                            ID++;
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }
            else
            {
                SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image> { };
                Get_Image_bonding(tar_parent.FullName, ref dic_image);
                SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_grp_data = new SortedDictionary<int, Funtion_SMT.Peeltest_data> { };
                get_data_bonding(tar_parent, ref dic_grp_data);


                for (int i = 1; i <= dic_grp_data.Count; i++)
                {
                    if (dic_image.ContainsKey(i) && dic_image.ContainsKey(i + qty))
                    {
                        Data_tbl.Rows.Add(i, txtItemCode.Text, txtLotNo.Text, i, dic_image[i], dic_image[i + qty], dic_grp_data[i].grap_data, dic_grp_data[i].data_val, txtOperator.Text, DateTime.Now.ToString(), txt_ItemName.Text + "_" + txt_date_peel.Text + "_" + txt_worker_peel.Text + "_" + cb_Type.SelectedItem.ToString());
                    }

                }
            }
            return Data_tbl;
        }

        private void btn_loadpeel_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtLogfile_peel.Text != "" && txtOperator.Text != "" && cb_Type.SelectedIndex != -1 && myCode.IsNumeric(txt_qty_peel.Text))
            {
                DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, "ACF", cb_Type.SelectedItem.ToString() }));
                if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                {
                    dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, "ACF", "MASS" }));
                }

                if (dt_spec.Rows.Count > 0)
                {
                    DataTable dt = load_data_logfile_peel(txtLogfile_peel.Text);
                    try
                    {

                        ProductIDService service = new ProductIDService(txtItemCode.Text, txtLotNo.Text, textBox2.Text, new[] { "OQC", "ACF" }, new[] { "Bonding" });
                        if (service._listFile.Count > 0)
                        {
                            List<string> s = service.getListProductID(service._listFile["Bonding"]);
                            if (!dt.Columns.Contains("ProductID"))
                            {
                                dt.Columns.Add("ProductID", typeof(string));
                            }
                            int i = 0;
                            foreach (DataRow row in dt.Rows)
                            {
                                if (i < s.Count)
                                {
                                    row["ProductID"] = s[i++];
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Product ID: {ex.Message}");
                    }
                    dgv_peel.DataSource = dt;
                    if (dgv_peel.DataSource != null)
                    {
                        check_data_bonding_mass(dgv_peel);
                    }
                    myCode.Disable_Sort_DGV(dgv_peel);
                    resize_column_image_bonding(dgv_peel);
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Chưa cài đặt format", "Thông báo");
                }

            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Vui lòng điền đầy đủ thông tin! ", "Thông báo");
            }


        }

        public void save_data(string tbl_name, DataGridView dgv_data)
        {
            DataTable tbl_data = (DataTable)dgv_data.DataSource;
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && cb_Type.SelectedIndex != -1)
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                if (tbl_name == "ACF_BONDING")
                {
                    tbl_name += "_NAS";
                    filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Remark" }, new string[] { txtItemCode.Text, txtLotNo.Text, txt_ItemName.Text + "_" + txt_date_peel.Text + "_" + txt_worker_peel.Text + "_" + cb_Type.SelectedItem.ToString() });
                }
                if (tbl_name == "Roughness")
                {
                    tbl_name = "Roughness_NAS";
                }
            lblsave:
                DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, tbl_name, filter_str);
                if (dt.Rows.Count == 0)
                {
                    if (tbl_data.Rows.Count > 0)
                    {

                        if (tbl_data.Columns.Contains("ProductID"))
                        {
                            string content = ProductIDService.ConverterProductID(tbl_data, "Id");
                            ProductIDService.InsertProductID(txtItemCode.Text, txtLotNo.Text, tbl_name, content);
                            tbl_data.Columns.Remove("ProductID");
                        }
                        if (tbl_name == "ACF_BONDING_NAS")
                        {
                            NasRepository nas = new NasRepository();
                            string location = nas.HandleImageDataTable(tbl_data, tbl_name, txtItemCode.Text, txtLotNo.Text);
                            string json = ConverterService.DataTableToJson(tbl_data);
                            DataRow row = dt.NewRow();
                            row["ItemCode"] = txtItemCode.Text;
                            row["LotNo"] = txtLotNo.Text;
                            row["Remark"] = tbl_data.Rows[0]["Remark"];
                            row["Data"] = json;
                            row["LocationImg"] = location;
                            dt.Rows.Add(row);
                        }
                        else
                        {
                            NasRepository nas = new NasRepository();
                            string location = nas.HandleImageDataTable(tbl_data, tbl_name, txtItemCode.Text, txtLotNo.Text);
                            string json = ConverterService.DataTableToJson(tbl_data);
                            DataRow row = dt.NewRow();
                            row["ItemCode"] = txtItemCode.Text;
                            row["LotNo"] = txtLotNo.Text;
                            row["Data"] = json;
                            dt.Rows.Add(row);
                        }
                        new DBContext().BuckDataTable(dt, $"{tbl_name}", new[] { "ItemCode", "LotNo" }, null, "ID");
                        MessageBox.Show(new Form { TopMost = true }, "Lưu thành công", "Thông báo");
                        dgv_data.DataSource = new DataTable();
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Không có dữ liệu", "Thông báo");
                    }
                }
                else
                {

                    if (MessageBox.Show(new Form { TopMost = true }, "Dữ liệu đã được lưu. Bạn có muốn cập nhật không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        if (UserSession.Instance.IsLoggedIn)
                        {
                            TDMK_Code.Delelte_FilteredItem_arr(tbl_name, sqlcon, filter_str);
                            goto lblsave;
                        }
                        else
                        {
                            MessageBox.Show("Vui lòng đăng nhập để cập nhật", "Thông báo");

                        }
                    }

                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Vui lòng điền đầy đủ thông tin! ", "Thông báo");
            }
        }

        private void btn_save_peel_Click(object sender, EventArgs e)
        {
            if (dgv_peel.DataSource != null)
            {
                bool chk = true;
                if (cb_Type.SelectedIndex != -1)
                {
                    if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                    {
                        DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, "ACF", "MASS" }));
                        if (dt_spec.Rows.Count > 0)
                        {
                            if (myCode.IsNumeric(txt_qty_peel.Text))
                            {
                                int r_setup = int.Parse(txt_qty_peel.Text);
                                if (r_setup > dgv_peel.Rows.Count)
                                {
                                    chk = false;
                                }
                            }
                        }

                    }
                    if (chk)
                    {
                        bool chk_NG = true;
                        foreach (DataGridViewRow dr in dgv_peel.Rows)
                        {
                            if (dr.Cells["Data"].Style.BackColor == Color.Red)
                            {
                                chk_NG = false;
                                break;
                            }
                        }
                    lbl_save:
                        if (chk_NG)
                        {

                            save_data("ACF_BONDING", dgv_peel);
                            MessageBox.Show("Save successfully!");

                        }
                        else
                        {
                            if (MessageBox.Show(new Form { TopMost = true }, "Dữ liệu NG. Bạn có muốn lưu không ?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                            {
                                chk_NG = true;
                                goto lbl_save;

                            }
                        }

                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Không đủ dữ liệu", "Thông báo");
                    }
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Vui lòng chọn Type", "Thông báo");
                }


            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Không có dữ liệu", "Thông báo");
            }

        }

        private void txtItemCode_FPC_TextChanged(object sender, EventArgs e)
        {


        }

        private void txtLotNo_FPC_TextChanged(object sender, EventArgs e)
        {

        }

        private void dgv_peel_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int col_inx = e.ColumnIndex;
            int r_inx = e.RowIndex;

            if (col_inx != -1 && r_inx != -1)
            {
                DataGridViewCell cur_cell = dgv_peel.CurrentCell;
                if (dgv_peel.Columns[col_inx].Name.Contains("Image") || dgv_peel.Columns[col_inx].Name.Contains("Graph"))
                {
                    if (myCode.checkDBNull(cur_cell.Value) != "")
                    {
                        View_detail_Image fr1 = new View_detail_Image() { TopMost = true };
                        fr1.data = (byte[])cur_cell.Value;
                        fr1.Show();
                    }
                }
            }
        }

        private void tab_main_Click(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel5_TabIndexChanged(object sender, EventArgs e)
        {

        }

        private void tab_main_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (dgv_peel.Rows.Count > 0)
            {
                if (dgv_peel.Columns.Contains("Image_Before"))
                {
                    resize_column_image_bonding(dgv_peel);
                }
                else if (dgv_peel.Columns.Contains("Image_Data"))
                {
                    resize_column_image(dgv_peel);
                }

            }
        }

        private void txtLogfile_peel_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //btn_loadpeel.PerformClick();
                string str_infor = Path.GetFileNameWithoutExtension(txtLogfile_peel.Text);
                if (str_infor.Contains("_"))
                    str_infor = str_infor.Split('_')[0];


                string[] arr_infor = str_infor.Split('-');
                if (arr_infor.Length == 6)
                {
                    txt_ItemName.Text = arr_infor[0];
                    txt_date_peel.Text = arr_infor[3] + "-" + arr_infor[4];
                    txt_worker_peel.Text = arr_infor[5];

                }
                else if (arr_infor.Length == 7)
                {
                    txt_ItemName.Text = arr_infor[0];
                    txt_date_peel.Text = arr_infor[3] + "-" + arr_infor[4] + "-" + arr_infor[5];
                    txt_worker_peel.Text = arr_infor[6];

                }
                else
                {
                    txt_ItemName.Text = arr_infor[0];

                }
            }
        }

        private void txtLogfile_peel_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "")
            {
                FolderBrowserDialog f_open = new FolderBrowserDialog();
                f_open.SelectedPath = System.Windows.Forms.Application.StartupPath;
                if (f_open.ShowDialog() == DialogResult.OK)
                {
                    txtLogfile_peel.Text = f_open.SelectedPath;

                    string str_infor = Path.GetFileNameWithoutExtension(txtLogfile_peel.Text);
                    if (str_infor.Contains("_"))
                        str_infor = str_infor.Split('_')[0];


                    string[] arr_infor = str_infor.Split('-');
                    if (arr_infor.Length == 6)
                    {
                        txt_ItemName.Text = arr_infor[0];
                        txt_date_peel.Text = arr_infor[3] + "-" + arr_infor[4];
                        txt_worker_peel.Text = arr_infor[5];

                    }
                    else if (arr_infor.Length == 7)
                    {
                        txt_ItemName.Text = arr_infor[0];
                        txt_date_peel.Text = arr_infor[3] + "-" + arr_infor[4] + "-" + arr_infor[5];
                        txt_worker_peel.Text = arr_infor[6];

                    }
                    else
                    {
                        txt_ItemName.Text = arr_infor[0];

                    }


                }

            }
        }

        private void btnLoad_roughness_Click(object sender, EventArgs e)
        {

        }



        private void btn_export_bonding_Click(object sender, EventArgs e)
        {
            if (cb_Type.SelectedItem.ToString().Contains("NPI"))
            {
                return;
            }
            if (cb_Type.SelectedItem.ToString().Contains("MASS") && txtItemCode.Text != "" && txtLotNo.Text != "" && txt_qty_peel.Text != "")
            {
                if (dgv_peel.DataSource != null)
                {
                    string report_folder_bonding = Path.Combine(data_loc, "Report", cb_Type.SelectedItem.ToString(), "ACF", "ACF_BONDING");
                    if (!System.IO.Directory.Exists(report_folder_bonding))
                        System.IO.Directory.CreateDirectory(report_folder_bonding);

                    string format_bonding = find_format_mass(Path.Combine(data_loc, "Format", "MASS", "ACF", "ACF_BONDING"), txtItemCode.Text);
                    if (format_bonding != "")
                    {
                        if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                        {
                            if (myCode.IsNumeric(txt_qty_peel.Text))
                            {
                                F_export_EPPlus.export_MASS_ACFpeel(format_bonding, report_folder_bonding, txtItemCode.Text, txtLotNo.Text, dgv_peel, cb_Type.SelectedItem.ToString(), txtOperator.Text);
                            }
                            else
                            {
                                MessageBox.Show(new Form { TopMost = true }, "Chưa nhập Qty", "Thông báo");
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Không tìm thấy format", "Thông báo");
                    }
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Không có dữ liệu", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Vui lòng điền đầy đủ thông tin", "Thông báo");
            }
        }


        private void btn_export_roughness_Click(object sender, EventArgs e)
        {
            if (dgv_roughness_data.DataSource == null)
            {
                dgv_roughness_data.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "Roughness", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));

                if (spec_roughness != "")
                {
                    check_roughness(dgv_roughness_data);
                }

                myCode.Disable_Sort_DGV(dgv_roughness_data);
            }

            if (cb_Type.SelectedItem.ToString().Contains("MASS") && txtItemCode.Text != "" && txtLotNo.Text != "" && dgv_roughness_data.DataSource != null)
            {
                string report_folder = Path.Combine(data_loc, "Report", cb_Type.SelectedItem.ToString(), "ACF", "ACF_ROUGHNESS");
                if (!System.IO.Directory.Exists(report_folder))
                    System.IO.Directory.CreateDirectory(report_folder);


                string export_path = Path.Combine(report_folder, cb_Type.SelectedItem.ToString().Replace("MASS", "").Replace("(", "").Replace(")", "") + " Roughness " + txt_ItemName.Text + "-" + txtItemCode.Text + "-" + txtLotNo.Text + ".xlsx");

                if (!System.IO.File.Exists(export_path))
                {
                    string format_path = find_format_mass(Path.Combine(data_loc, "Format", "MASS", "ACF", "ACF_ROUGHNESS"), txtItemCode.Text);
                    if (format_path != "")
                    {
                        myExcel.Workbook report_format = TDMK_Code.open_excel_file(format_path, "", "");
                        report_format.SaveAs(export_path);
                        report_format.Close();
                    }
                    else
                    {

                        export_path = "";
                    }
                }

                if (export_path != "")
                {
                    myExcel.Workbook wb = TDMK_Code.open_excel_file(export_path, "", "");
                    myExcel.Worksheet ws = wb.Sheets[1];

                    DataTable dt_roughness = (DataTable)dgv_roughness_data.DataSource;
                    if (dt_roughness != null)
                    {
                        judgeall_rougness = true;
                        string itemname = txt_ItemName.Text;
                        string date = dt_roughness.Rows[0]["MDate"].ToString();
                        string worker = dt_roughness.Rows[0]["Operator"].ToString();
                        string machine = dt_roughness.Rows[0]["Machine"].ToString();
                        List<string> lst_infor = new List<string> { itemname, txtItemCode.Text, txtLotNo.Text, date, machine, worker };
                        export_info_mass(ws, lst_infor);

                        Export_ACF_roughness_mass(ws, dt_roughness);
                        MessageBox.Show(new Form { TopMost = true }, "Xuất báo cáo thành công!", "Warning");
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Không có dữ liệu", "Warning");
                    }
                }
                else
                {
                }
            }
        }

        public void export_info_mass(myExcel.Worksheet ws, List<string> lst_infor)
        {

            string[] arr_item = new string[5] { "ITEMNAME:", "ITEMCODE:", "LOTNO:", "DATE:", "MACHINE:" };
            Dictionary<string, string> dic_item = new Dictionary<string, string>();

            int i = 0;
            foreach (string item in arr_item)
            {
                dic_item.Add(item, lst_infor[i]);
                i++;

            }

            foreach (var val in dic_item)
            {
                myExcel.Range cur_rgn = find_cell(ws, val.Key);
                if (cur_rgn != null)
                {
                    if (val.Key == "LOTNO:")
                    {
                        cur_rgn.Offset[0, 1].FormulaR1C1 = "'" + val.Value;
                    }
                    else
                    {
                        cur_rgn.Offset[0, 1].Value = val.Value;
                    }

                }
            }


            string leader = txtOperator.Text;
            myExcel.Range cur_rgn_worker = find_cell(ws, "WORKER");
            if (cur_rgn_worker != null)
                cur_rgn_worker.Offset[1, 0].Value = lst_infor[5];

            myExcel.Range cur_rgn_leader = find_cell(ws, "LEADER");
            if (cur_rgn_leader != null)
                cur_rgn_leader.Offset[1, 0].Value = leader;


        }
        public void export_info_peel(myExcel.Worksheet ws, List<string> lst_infor)
        {

            string[] arr_item = new string[5] { "ITEMNAME:", "ITEMCODE:", "LOTNO:", "DATE:", "PROCESS:" };
            Dictionary<string, string> dic_item = new Dictionary<string, string>();

            int i = 0;
            foreach (string item in arr_item)
            {
                dic_item.Add(item, lst_infor[i]);
                i++;

            }

            foreach (var val in dic_item)
            {
                myExcel.Range cur_rgn = find_cell(ws, val.Key);
                if (cur_rgn != null)
                {
                    if (val.Key == "LOTNO:")
                    {
                        cur_rgn.Offset[0, 1].FormulaR1C1 = "'" + val.Value;
                    }
                    else
                    {
                        cur_rgn.Offset[0, 1].Value = val.Value;
                    }

                }
            }


            string leader = txtOperator.Text;
            myExcel.Range cur_rgn_worker = find_cell(ws, "WORKER");
            if (cur_rgn_worker != null)
                cur_rgn_worker.Offset[1, 0].Value = lst_infor[5];

            myExcel.Range cur_rgn_leader = find_cell(ws, "LEADER");
            if (cur_rgn_leader != null)
                cur_rgn_leader.Offset[1, 0].Value = leader;


        }

        public myExcel.Range find_cell(myExcel.Worksheet ws, string find_item)
        {
            myExcel.Range cur_cell = null;

            for (int j = 1; j < 15; j++)
            {
                for (int i = 1; i < 15; i++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).ToUpper().Replace(" ", "").Contains(find_item))
                    {
                        cur_cell = ws.Cells[i, j];
                        return cur_cell;

                    }
                }
            }

            return cur_cell;

        }

        private void dgv_peel_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (dgv_peel.DataSource != null)
            {
                check_data_bonding_mass(dgv_peel);
            }
        }

        private void dgv_peel_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            if (cb_Type.SelectedIndex != -1 && cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, "ACF", "MASS" }));
                if (e.RowIndex != -1)
                {
                    string ID = dgv_peel.Rows[e.RowIndex].Cells["ID"].Value.ToString();
                    int idx = 0;
                    if (int.Parse(ID) <= 5)
                    {
                        idx = 0;
                    }
                    else if (int.Parse(ID) > 6 && int.Parse(ID) <= 10)
                    {
                        idx = 1;
                    }
                    else if (int.Parse(ID) > 10 && int.Parse(ID) <= 15)
                    {
                        idx = 2;
                    }

                    if (dt_spec.Rows.Count > 0)
                    {
                        string[] spec_all = dt_spec.Rows[0]["Location"].ToString().Split('^');
                        if (idx < spec_all.Length - 1)
                        {
                            string spec = spec_all[idx];
                            string R = spec.Split(';')[0];
                            string UCL = spec.Split(';')[1];
                            string LCL = spec.Split(';')[2];

                            dgv_peel.CurrentCell.ToolTipText = "R: " + R + "\n" + "UCL: " + UCL + "\n" + "LCL: " + LCL + "\n";
                        }

                    }
                }


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
            while (myCode.checkDBNull(tar_rgn.Offset[inx, 0].Value) != "")
            {
                for (int i = 0; i <= MeasLoc - 1; i++)
                {
                    int cur_val = inx - 2 * i;
                    int div_val = cur_val / (2 * MeasLoc);
                    int hieuso = cur_val - div_val * 2 * MeasLoc;
                    if (hieuso == 0)
                    {
                        if (myCode.IsNumeric(myCode.checkDBNull(tar_rgn.Offset[inx, 0].Value)))
                        {
                            double Sa_val = Convert.ToDouble(myCode.checkDBNull(tar_rgn.Offset[inx, 0].Value));
                            Sa_result[i].Add(Math.Round(Sa_val, 3).ToString());//"#0.##0"
                        }

                        if (myCode.IsNumeric(myCode.checkDBNull(tar_rgn.Offset[inx, 1].Value)))
                        {
                            double Sq_val = Convert.ToDouble(myCode.checkDBNull(tar_rgn.Offset[inx, 1].Value));
                            Sq_result[i].Add(Math.Round(Sq_val, 3).ToString());//"#0.##0"
                        }
                        for (int t = 2; t < 5; t++)
                        {
                            if (myCode.checkDBNull(tar_rgn.Offset[-1, t].Value).Contains("Sdr"))
                            {
                                if (myCode.IsNumeric(myCode.checkDBNull(tar_rgn.Offset[inx + 1, t].Value)))
                                {
                                    double Sdr_val = Convert.ToDouble(myCode.checkDBNull(tar_rgn.Offset[inx + 1, t].Value));
                                    Sdr_result[i].Add(Math.Round(Sdr_val, 3).ToString());//"#0.##0"
                                }
                                break;
                            }

                        }

                    }
                }
                inx++;
            }
            tar_wrkbook.Close();
            return _result;
        }
        public DataTable Roughness_Scan_ACF_SMT(string src_file, int qty, int MeasLoc)
        {
            //if (tar_DGV_Spec.RowCount > 0)
            //{
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
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
            // DataTable dt_analysis = TDMK_Code.Datatable_Filter(sqlcon, "Roughness", filter_str);
            DataTable dt_data = TDMK_Code.Datatable_Filter(sqlcon, "Roughness", filter_str).Clone();
            for (int i = 0; i < qty; i++)
            {
                DataRow dr = dt_data.NewRow();
                dr[0] = i + 1;
                dr[1] = txtItemCode.Text;
                dr[2] = txtLotNo.Text;
                dr[3] = txt_worker_roughness.Text;
                dr[4] = txt_machine.Text;
                dr[5] = txt_date_roughness.Text;

                for (int c = 6; c < dt_data.Columns.Count; c++)
                {
                    dr[c] = data_arr[c - 6][i];
                }
                dt_data.Rows.Add(dr);
            }

            return dt_data;

            ////********************* Add new columns for Sdr ***********************************************
            //sel_DGV.Columns.Add("L1 Roughness Sdr", "L1 Roughness Sdr");
            //sel_DGV.Columns.Add("L2 Roughness Sdr", "L2 Roughness Sdr");
            ////*********************************************************************************************
            ////********************* Add new columns for L3 Sa / Sq / Sdr ***********************************************
            //sel_DGV.Columns.Add("L3 Roughness Sa", "L3 Roughness Sa");
            //sel_DGV.Columns.Add("L3 Roughness Sq", "L3 Roughness Sq");
            //sel_DGV.Columns.Add("L3 Roughness Sdr", "L3 Roughness Sdr");
            //*********************************************************************************************

            //sel_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            //myCode.Disable_Sort_DGV(sel_DGV);


        }

        public void err_color(Control e)
        {
            if (!myCode.IsNumeric(e.Text))
            {
                e.BackColor = Color.Yellow;
            }
            else
            {
                e.BackColor = Color.White;
            }
        }

        private void btn_loadroughness_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && txt_qty.Text != "" && txt_measloc.Text != "" && cb_Type.SelectedIndex != -1)
            {
                if (int.TryParse(txt_measloc.Text, out int num) && int.Parse(txt_measloc.Text) <= 3)
                {
                    if (myCode.IsNumeric(txt_qty.Text) && myCode.IsNumeric(txt_measloc.Text))
                    {
                        string logfile = "";
                        FileAttributes attr = File.GetAttributes(txt_logfile_roughness.Text);
                        if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
                        {
                            string[] file_format_lst_xls = Directory.GetFiles(txt_logfile_roughness.Text, "*" + ".xls");
                            if (file_format_lst_xls.Length > 0)
                            {
                                logfile = file_format_lst_xls[0];
                            }
                        }
                        else
                        {
                            logfile = txt_logfile_roughness.Text;
                        }


                        if (logfile != "")
                        {
                            DataTable dt = Roughness_Scan_ACF_SMT(logfile, int.Parse(txt_qty.Text), int.Parse(txt_measloc.Text));

                            if (!string.IsNullOrEmpty(textBox3.Text))
                            {
                                ProductIDService pid = new ProductIDService(txtItemCode.Text, txtLotNo.Text, textBox3.Text, new[] { "OQC", "ACF" }, new[] { "roughness" });
                                if (pid._listFile.Count > 0)
                                {
                                    List<string> s = pid.getListProductID(pid._listFile["roughness"]);
                                    if (!dt.Columns.Contains("ProductID"))
                                    {
                                        dt.Columns.Add("ProductID", typeof(string));
                                    }
                                    int i = 0;
                                    foreach (DataRow row in dt.Rows)
                                    {
                                        if (i < s.Count)
                                        {
                                            row["ProductID"] = s[i++];
                                        }
                                    }
                                }
                            }
                            dgv_roughness_data.DataSource = dt;
                        }
                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Đường dẫn không đúng", "Thông báo");
                        }

                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Nhập đầy đủ Qty và MeasLoc", "Thông báo");
                    }
                    if (spec_roughness != "")
                    {
                        check_roughness(dgv_roughness_data);
                    }
                    myCode.Disable_Sort_DGV(dgv_roughness_data);

                }
                else
                {
                    txt_measloc.BackColor = Color.Red;
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Vui lòng điền đầy đủ thông tin", "Thông báo");
            }

        }

        private void txt_logfile_roughness_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            OpenFileDialog f_open = new OpenFileDialog();
            f_open.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            f_open.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            if (f_open.ShowDialog() == DialogResult.OK)
            {
                if (f_open.FileName != "")
                {
                    txt_logfile_roughness.Text = f_open.FileName;

                    string str_infor = Path.GetFileNameWithoutExtension(txt_logfile_roughness.Text);
                    if (str_infor.Contains("_"))
                        str_infor = str_infor.Split('_')[0];


                    string[] arr_infor = str_infor.Split('-');
                    if (arr_infor.Length == 7)
                    {
                        txt_ItemName.Text = arr_infor[0];
                        txt_date_roughness.Text = arr_infor[3] + "-" + arr_infor[4];
                        txt_worker_roughness.Text = arr_infor[5];
                        txt_machine.Text = arr_infor[6];
                    }
                    else if (arr_infor.Length == 8)
                    {
                        txt_ItemName.Text = arr_infor[0];
                        txt_date_roughness.Text = arr_infor[3] + "-" + arr_infor[4] + "-" + arr_infor[5];
                        txt_worker_roughness.Text = arr_infor[6];
                        txt_machine.Text = arr_infor[7];
                    }
                    else
                    {
                        txt_ItemName.Text = arr_infor[0];

                    }
                }

            }
        }

        private void txt_qty_TextChanged(object sender, EventArgs e)
        {
            err_color(txt_qty);
        }

        private void txt_measloc_TextChanged(object sender, EventArgs e)
        {
            err_color(txt_measloc);
        }

        private void txt_logfile_roughness_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void btn_save_roughness_Click(object sender, EventArgs e)
        {
            if (dgv_roughness_data.DataSource != null)
            {
                save_data("Roughness", dgv_roughness_data);
                MessageBox.Show("Save successfully!");

            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Không có dữ liệu", "Waring");
            }
        }

        private void txt_machine_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            err_color(txt_qty_peel);
        }

        private void txt_ItemName_TextChanged(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

                ACFService service = new ACFService();
                txt_ItemName.Text = service.loadItemNamebyItemCode(txtItemCode.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (checkRowSpec((DataGridView)sender, ((DataGridView)sender).Rows[e.RowIndex]))
            {

                ((DataGridView)sender).Rows[e.RowIndex].Cells["JUDGE"].Value = "PASS";
            }
            else
            {
                ((DataGridView)sender).Rows[e.RowIndex].Cells["JUDGE"].Value = "FAIL";
            }
        }
        private void FillDataList(List<DateTime> dateTime)
        {
            dateTimePicker1.Value = dateTime[0];
            dateTimePicker2.Value = dateTime[1];
            dateTimePicker3.Value = dateTime[2];
            dateTimePicker4.Value = dateTime[3];
            dateTimePicker5.Value = dateTime[4];
            dateTimePicker6.Value = dateTime[5];
        }
        private void FillDataList()
        {
            string i = listBox1.Items[listBox1.SelectedIndex].ToString();
            dic_list[i][0] = dateTimePicker1.Value;
            dic_list[i][1] = dateTimePicker2.Value;
            dic_list[i][2] = dateTimePicker3.Value;
            dic_list[i][3] = dateTimePicker4.Value;
            dic_list[i][4] = dateTimePicker5.Value;
            dic_list[i][5] = dateTimePicker6.Value;
            dic[i] = dataGridView.DataSource as DataTable;
        }
        private bool checkRowSpec(DataGridView dgv, DataGridViewRow row)
        {

            bool prime = true;
            if (row.Cells["Spec"].Value == null)
            {
                prime = false;
            }
            else if (double.TryParse(row.Cells["Spec"].Value.ToString(), out double d) && d == 0)
            {
                prime = false;
            }
            if (!prime)
            {
                row.Cells["Spec"].Style.BackColor = Color.Red;
            }

            else
            {

                row.Cells["Spec"].Style.BackColor = Color.White;
                if (double.TryParse(row.Cells["Spec"].Value.ToString(), out double d))
                {
                    foreach (DataGridViewColumn col in dgv.Columns)
                    {
                        prime = true;
                        if (col.Name.Contains("Sample"))
                        {
                            if (double.TryParse(row.Cells[col.Name].Value.ToString(), out double dz))
                            {
                                if (dz >= d)
                                {

                                    prime = false;
                                    row.Cells[col.Name].Style.BackColor = Color.Red;

                                }
                                else
                                {
                                    row.Cells[col.Name].Style.BackColor = Color.White;
                                }
                            }
                            else
                            {

                                prime = false;
                                row.Cells[col.Name].Style.BackColor = Color.Red;

                            }
                        }
                    }
                }
            }
            return prime;
        }
        private void checkSpecWCA(DataGridView dgv)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.Index == dgv.Rows.Count - 1)
                {
                    break;
                }
                if (!checkRowSpec(dgv, row))
                {
                    row.Cells["JUDGE"].Value = "FAIL";
                }
                else
                {
                    row.Cells["JUDGE"].Value = "PASS";
                }

            }
        }
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (tdmK_Button3.Enabled == true)
                {
                    if (MessageBox.Show("Bạn có muốn lưu thay đổi hay không", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        FillDataList();
                    }
                }
                string i = listBox1.Items[listBox1.SelectedIndex].ToString();
                dataGridView.DataSource = dic[i];
                checkSpecWCA(dataGridView);
                FillDataList(dic_list[i]);
                tdmK_Button3.Enabled = false;
            }
            catch
            {

            }
        }
        private void remove()
        {
            if (listBox1.SelectedIndex != -1)
            {
                string i = listBox1.Items[listBox1.SelectedIndex].ToString();
                dic.Remove(i);
                dic_list.Remove(i);
                listBox1.Items.RemoveAt(listBox1.SelectedIndex);
                FillData();
            }
        }
        private void tdmK_Button2_Click(object sender, EventArgs e)
        {
            remove();
        }
        private void rename()
        {
            if (listBox1.SelectedIndex != -1)
            {
                string i = listBox1.Items[listBox1.SelectedIndex].ToString();
                Form inputForm = new Form();
                TableLayoutPanel tbl = new TableLayoutPanel() { Dock = DockStyle.Fill };
                TableLayoutPanel tbl1 = new TableLayoutPanel() { Dock = DockStyle.Fill };
                System.Windows.Forms.Button button = new System.Windows.Forms.Button() { Text = "Ok", Dock = DockStyle.Fill, Width = 300, Height = 200 };


                tbl.ColumnCount = 2;
                tbl1.RowCount = 2;

                tbl1.RowStyles.Add(new ColumnStyle(SizeType.Percent, 20));
                tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
                tbl1.Controls.Add(button, 0, 1);
                tbl1.Controls.Add(tbl, 0, 0);
                System.Windows.Forms.TextBox tb = new System.Windows.Forms.TextBox() { Text = i, Dock = DockStyle.Fill };
                TDMK_Label lb = new TDMK_Label() { Text = "Nhập tên mới" };
                tbl.Controls.Add(tb, 1, 0);
                tbl.Controls.Add(lb, 0, 0);
                inputForm.Controls.Add(tbl1);
                inputForm.ShowDialog();

                string new_name = tb.Text;
                if (new_name != "")
                {
                    if (!dic.ContainsKey(new_name))
                    {
                        dic.Add(new_name, dic[i]);
                        dic_list.Add(new_name, dic_list[i]);
                        dic.Remove(i);
                        dic_list.Remove(i);
                        listBox1.Items[listBox1.SelectedIndex] = new_name;
                    }
                    else
                    {
                        MessageBox.Show("Tên đã tồn tại");
                    }
                }
            }
        }


        private void tdmK_Button1_Click(object sender, EventArgs e)
        {
            rename();
        }


        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            primeNow();
        }
        private void primeNow()
        {
            tdmK_Button3.Enabled = true;
        }
        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

            primeNow();
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {

            primeNow();
        }

        private void dateTimePicker4_ValueChanged(object sender, EventArgs e)
        {

            primeNow();
        }

        private void dateTimePicker6_ValueChanged(object sender, EventArgs e)
        {

            primeNow();
        }

        private void dateTimePicker5_ValueChanged(object sender, EventArgs e)
        {
            primeNow();
        }

        private void dataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            primeNow();
        }

        private void tdmK_Button3_Click(object sender, EventArgs e)
        {
            tdmK_Button3.Enabled = false;
            FillDataList();
        }

        private void textBox1_DoubleClick(object sender, EventArgs e)
        {
            OpenFileDialog f_open = new OpenFileDialog();
            f_open.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            f_open.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            if (f_open.ShowDialog() == DialogResult.OK)
            {
                if (f_open.FileName != "")
                {
                    txtLogfile_Flatness.Text = f_open.FileName;
                }

            }
        }

        private void textBox2_DoubleClick(object sender, EventArgs e)
        {
            OpenFileDialog f_open = new OpenFileDialog();
            f_open.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            f_open.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            if (f_open.ShowDialog() == DialogResult.OK)
            {
                if (f_open.FileName != "")
                {
                    txtLogfile_Flatness.Text = f_open.FileName;
                }

            }
        }

        private void textBox3_DoubleClick(object sender, EventArgs e)
        {
            OpenFileDialog f_open = new OpenFileDialog();
            f_open.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            f_open.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            if (f_open.ShowDialog() == DialogResult.OK)
            {
                if (f_open.FileName != "")
                {
                    txtLogfile_Flatness.Text = f_open.FileName;
                }

            }
        }
    }

}
