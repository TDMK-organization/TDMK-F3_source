using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SEEV_DLL;
using myExcel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using Image = System.Drawing.Image;
using Microsoft.Office.Core;
using Bending_Export;
using FAI_Export;
using OK2SHIP_Lib;
using System.Diagnostics;
using OfficeOpenXml;

using IniLibs;
using Export_FPCA_OK2ship_Auto_System.Services;
using Export_FPCA_OK2ship_Auto_System.Libary;
using TDMK_SQL;
using Export_FPCA_OK2ship_Auto_System.Repositories;
using Export_FPCA_OK2ship_Auto_System.Views;



namespace Export_FPCA_OK2ship_Auto_System
{
    public partial class Form1 : Form
    {
        private AirBubbleService airService = null;
        //public Bending_Export_Lib Bending_Exp = new Bending_Export_Lib();
        public Bending_Export_EPPLUS_Lib Bending_Exp = new Bending_Export_EPPLUS_Lib();
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public SEI_Lib myCode = new SEI_Lib();
        public TDMK_EPPLUS TDMK_Code2 = new TDMK_EPPLUS();
        public SqlConnection sqlcon = null;
        public string data_loc = "";
        public Funtion_export_FPCA F_expNPI = new Funtion_export_FPCA();
        List<string> lst_sheet_export = new List<string> { };
        List<string> lst_auto = new List<string> { };
        List<string> lst_auto_new = new List<string> { };
        bool hide_mode = true;
        string sheet_select = "";
        public TDMK_OK2SHIP myCode2 = new TDMK_OK2SHIP();
        FAI_EPPLUS_Lib FAI_lib = new FAI_EPPLUS_Lib();
        string app_path = "";
        List<string> lst_manual = new List<string> { };
        string str_infor = "";
        string Itemcode_unmating = "";
        string LotNo_unmating = "";
        string ItemCode_liner_coupon = "";
        string LotNo_liner_coupon = "";
        string ItemCode_psa_coupon = "";
        string LotNo_psa_coupon = "";
        string ItemCode_crosscut = "";
        string lotNo_crosscut = "";


        public Form1()
        {
            InitializeComponent();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void txtLotNo_TextChanged(object sender, EventArgs e)
        {
            dgv_progress.DataSource = null;
        }

        public string progress(string sheet, string ItemCode, string LotNo)
        {
            string dr_progress = "";
            if (ItemCode != "" && LotNo != "")
            {
                DataTable dt_analysis = TDMK_Code.Datatable_Filter(sqlcon, sheet, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                if (dt_analysis.Rows.Count > 0)
                {
                    string[] arr_val = dt_analysis.AsEnumerable().Select(x => x.Field<string>("Sheet")).Distinct().ToArray();
                    string filter = "";
                    foreach (string item in arr_val)
                    {
                        if (item.Contains("NPI"))
                        {
                            filter = item;
                            break;
                        }
                    }
                    if (filter != "")
                    {
                        dr_progress = "x";
                    }
                }
            }
            return dr_progress;

        }



        private void btn_check_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "")
            {
                DataTable dt_sum = new DataTable();
                dt_sum.Columns.Add("ID");
                dt_sum.Columns.Add("Sheet");
                dt_sum.Columns.Add("Progress");
                dt_sum.Columns.Add("Export Success");

                SortedDictionary<int, string> dic_sheet = new SortedDictionary<int, string> { };
                string file_format = find_format(data_loc, txtItemCode.Text);
                if (file_format != "")
                {
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    ExcelPackage xlPackage = open_excel(file_format);
                    ExcelWorkbook wb = xlPackage.Workbook;
                    int i = 1;
                    foreach (ExcelWorksheet ws in wb.Worksheets)
                    {
                        dic_sheet.Add(i, ws.Name.TrimEnd(' '));
                        i++;
                    }
                }

                foreach (var isheet in dic_sheet)
                {
                    string sheet = isheet.Value.Replace(" ", "_").Replace("(", "").Replace(")", "").ToUpper();
                    string dr_progress = "";
                    DataRow dr = dt_sum.NewRow();
                    dr[0] = isheet.Key;
                    dr[1] = isheet.Value;
                    Dictionary<string, DataTable> dic_data_FAI = FAI_lib.Export_FAI_Batch(sqlcon, txtItemCode.Text, txtLotNo.Text, "NPI");

                    if (sheet == "ACF")
                    {
                        string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                        DataTable dt_wetting = TDMK_Code.Datatable_Filter(sqlcon, "ACF_WETTING", filter_str);

                        DataTable dt_analysis = TDMK_Code.Datatable_Filter(sqlcon, "ACF_BONDING", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                        DataTable dt_bonding = new DataTable();
                        if (dt_analysis.Rows.Count > 0)
                        {
                            string[] arr_val = dt_analysis.AsEnumerable().Select(x => x.Field<string>("Remark")).Distinct().ToArray();
                            string filter = "";
                            foreach (string item in arr_val)
                            {
                                if (item.Contains("NPI"))
                                {
                                    filter = item;
                                    dt_bonding = dt_analysis.AsEnumerable().Where(r => r.Field<string>("Remark").Contains("NPI")).CopyToDataTable();
                                    break;
                                }
                            }
                        }

                        DataTable dt_flatness = TDMK_Code.Datatable_Filter(sqlcon, "ACF_FLATNESS", filter_str);
                        DataTable dt_roughness = TDMK_Code.Datatable_Filter(sqlcon, "Roughness", filter_str);

                        if (dt_wetting.Rows.Count > 0 && dt_bonding.Rows.Count > 0 && dt_flatness.Rows.Count > 0 && dt_roughness.Rows.Count > 0)
                        {
                            dr_progress = "x";
                        }
                    }
                    else if (dic_data_FAI.ContainsKey(isheet.Value))
                    {
                        //Dictionary<string, DataTable> dic_data = FAI_lib.Export_FAI_Batch(sqlcon, txtItemCode.Text, txtLotNo.Text, "NPI");
                        //if (dic_data.Count > 0)
                        //{
                        dr_progress = "x";
                        //}
                    }
                    else if (sheet.Contains("BENDING"))
                    {
                        string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                        string s = "";
                        switch (isheet.Value.Replace(" ", "").Replace("&", "").ToUpper())
                        {
                            case "FLEXBENDING":
                                s = "FLEX_BENDING";
                                break;

                            case "THERMALCYCLINGBENDING":
                                s = "THERMAL_CYCLING_AND_BEND";
                                break;

                            case "HEATSOAKBENDING":
                                s = "HEAT_SOAK_AND_BEND";
                                break;

                        }
                        if (s != "")
                        {
                            DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon, s, filter_str);
                            if (src_tbl.Rows.Count > 0)
                            {
                                dr_progress = "x";
                            }
                        }

                    }
                    else if (sheet.Contains("UNMATING") && !sheet.Contains("OQC"))
                    {
                        if (Itemcode_unmating != "" && LotNo_unmating != "")
                        {
                            dr_progress = progress(sheet, Itemcode_unmating, LotNo_unmating);
                        }
                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Chưa nhập ItemCode(NVL) và LotNo(NVL) của (IQC Unmating) Pull Test", "Thông báo");
                        }

                    }
                    else if (sheet.Contains("COUPON") && sheet.Contains("LINER"))
                    {
                        if (ItemCode_liner_coupon != "" && LotNo_liner_coupon != "")
                        {
                            dr_progress = progress(sheet, ItemCode_liner_coupon, LotNo_liner_coupon);
                        }
                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Chưa nhập ItemCode(NVL) và LotNo(NVL) của IQC Liner peeling (Coupon)", "Thông báo");
                        }
                    }
                    else if (sheet.Contains("COUPON") && sheet.Contains("PSA"))
                    {
                        if (ItemCode_psa_coupon != "" && LotNo_psa_coupon != "")
                        {
                            dr_progress = progress(sheet, ItemCode_psa_coupon, LotNo_psa_coupon);
                        }
                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Chưa nhập ItemCode(NVL) và LotNo(NVL) của IQC PSA peeling (Coupon)", "Thông báo");
                        }
                    }
                    //else if (sheet.Contains("CROSS") || sheet.Contains("GAP"))
                    //{
                    //    if (ItemCode_crosscut != "" && lotNo_crosscut != "")
                    //    {
                    //        dr_progress = progress(sheet, ItemCode_crosscut, lotNo_crosscut);
                    //    }
                    //    else
                    //    {
                    //        MessageBox.Show(new Form { TopMost = true }, "Chưa nhập ItemCode/LOTNO (OK2BUILD) của " + sheet, "Thông báo");
                    //    }
                    //}
                    else if (lst_auto_new.Contains(isheet.Value.Replace(" ", "").Replace("(", "").Replace(")", "").ToUpper()))
                    {
                        DataTable dt_analysis = TDMK_Code.Datatable_Filter(sqlcon, sheet, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                        if (dt_analysis.Rows.Count > 0)
                        {
                            string[] arr_val = dt_analysis.AsEnumerable().Select(x => x.Field<string>("Sheet")).Distinct().ToArray();
                            string filter = "";
                            foreach (string item in arr_val)
                            {
                                if (item.Contains("NPI"))
                                {
                                    filter = item;
                                    break;
                                }
                            }
                            if (filter != "")
                            {
                                dr_progress = "x";
                            }
                        }

                    }
                    else if (sheet.Contains("SEM"))
                    {
                        SEMServices sem = new SEMServices();
                        DataTable dt = sem.LoadDataProcess(txtItemCode.Text.Trim(), txtLotNo.Text.Trim());
                        if (dt.Rows.Count > 0)
                        {
                            dr_progress = "x";
                        }
                    }
                    else if (sheet.Contains("BAR_CODE"))
                    {
                        OQCB2BMatingUnmatting oqc = new OQCB2BMatingUnmatting();
                        DataTable dt = oqc.LoadProcess(txtItemCode.Text.Trim(), txtLotNo.Text.Trim());
                        if (dt.Rows.Count > 0)
                        {
                            dr_progress = "x";
                        }
                    }
                    else if (sheet.Contains("UNMATING") && sheet.Contains("OQC"))
                    {
                        OQCB2BMatingUnmatting oqc = new OQCB2BMatingUnmatting();
                        DataTable dt = oqc.LoadProcess(txtItemCode.Text.Trim(), txtLotNo.Text.Trim());
                        if (dt.Rows.Count > 0)
                        {
                            dr_progress = "x";
                        }
                    }
                    else
                    {
                        string sht = isheet.Value.Replace(" ", "").Replace("(", "").Replace(")", "").ToUpper();
                        string report_manual_path = Path.Combine(data_loc, "Report", "Report_Manual");
                        if (!System.IO.Directory.Exists(report_manual_path))
                            System.IO.Directory.CreateDirectory(report_manual_path);
                        if (Check_Exist_report_Manual(sht, txtItemCode.Text, txtLotNo.Text, report_manual_path))
                        {
                            dr_progress = "x";
                        }
                    }
                    dr[2] = dr_progress;
                    dr[3] = "";

                    dt_sum.Rows.Add(dr);
                }
                dgv_progress.DataSource = dt_sum;
                dgv_progress.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                myCode.Disable_Sort_DGV(dgv_progress);
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Vui lòng điền đầy đủ thông tin!", "Thông báo");
            }
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
        private void Form1_Load(object sender, EventArgs e)
        {
            //sqlcon = initial_data("OK2SHIP_SMT", true);
            lst_auto = new List<string> { "FAI", "Cross section", "GAP Connector", "Peel Test", "(Mating) Pull Test", "(IQC Unmating) Pull Test", "Shear test", "IQC Liner peeling (Coupon)", "IQC PSA peeling (Coupon)", "Flex bending", "Thermal Cycling & bending", "Heat Soak & bending", "ACF" };

            foreach (string item in lst_auto)
            {
                lst_auto_new.Add(item.Replace(" ", "").Replace("(", "").Replace(")", "").ToUpper());
            }
            //lst_manual = cbl_sheet_manual.Items.Cast<string>().ToList();

            sqlcon = Bending_Exp.initial_data("OK2SHIP_SMT", true);
            string app_path = System.Windows.Forms.Application.StartupPath;
            string config_path = Path.Combine(app_path.Replace("\\FPCA OK2SHIP Auto System", ""), "config.ini");
            //app_path = @"\\10.212.6.212\Saomai\QA\TDMK_DATA\Test_Areas\OK2SHIP_SMT\TDMK Program\FPCA OK2SHIP Auto System";
            IniFile za = new IniFile(config_path);
            data_loc = za.Read("Format_Folder", "SMT_Config") + $"\\SEEV Data";
            // app_path = System.Windows.Forms.Application.StartupPath; 
            SortedDictionary<int, string> dic_sheet = new SortedDictionary<int, string> { };

        }


        private void getval(string value)
        {
            if (value != "")
            {
                string[] arr_info = value.Split('^');
                Itemcode_unmating = arr_info[0];
                LotNo_unmating = arr_info[1];
                ItemCode_liner_coupon = arr_info[2];
                LotNo_liner_coupon = arr_info[3];
                ItemCode_psa_coupon = arr_info[4];
                LotNo_psa_coupon = arr_info[5];
            }

        }

        //public SqlConnection initial_data(string DB_name, bool sa_en)
        //{
        //    SqlConnection _sqlcon_OK2SHIP;
        //    string app_path = System.Windows.Forms.Application.StartupPath;
        //    string config_file = Path.Combine(app_path, "Config", "config.txt");
        //    string[] my_config = myCode.read_config_arr(config_file);
        //    string server_name = "";
        //    string server_acc = "";
        //    string server_pass = "";


        //    foreach (string c in my_config)
        //    {
        //        if (c.Contains("Server"))
        //        {
        //            server_name = c.Split(':')[1].Trim();
        //        }
        //        if (c.Contains("Account"))
        //        {
        //            server_acc = c.Split(':')[1].Trim();
        //        }
        //        if (c.Contains("Password"))
        //        {
        //            server_pass = c.Split(':')[1].Trim();
        //        }
        //        if (c.Contains("Data_Location"))
        //        {
        //            data_loc = c.Split('#')[1].Trim();
        //        }
        //    }
        //    if (sa_en)
        //    {
        //        string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, DB_name, server_acc, server_pass).ConnectionString;
        //        _sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);
        //    }
        //    else
        //    {
        //        string _strcon = TDMK_Code.data_connection2(server_name, DB_name).ConnectionString;
        //        _sqlcon_OK2SHIP = new SqlConnection(_strcon);
        //    }

        //    return _sqlcon_OK2SHIP;
        //}
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

        private void txtLotNo_Validated(object sender, EventArgs e)
        {
            txtLotNo.Text = Lotno_Formated(txtLotNo.Text);
        }

        private void cbAll_CheckedChanged(object sender, EventArgs e)
        {
            if (cbAll.Checked == true)
            {
                for (int i = 0; i < cbl_sheet.Items.Count; i++)
                {
                    cbl_sheet.SetItemChecked(i, true);
                }
                //  lst_sheet_export = lst_auto;
            }
            else
            {
                for (int i = 0; i < cbl_sheet.Items.Count; i++)
                {
                    cbl_sheet.SetItemChecked(i, false);
                }
                // lst_sheet_export = new List<string> { };
            }
        }
        public string find_format(string in_data_loc, string ItemCode)
        {
            string result = "";
            string path = Path.Combine(in_data_loc, "Format", "NPI");
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

                    }
                }
            }

            return result;
        }

        public void Update_dgv_complete(string sheet, bool en_exp)
        {
            sheet = sheet.Replace(" ", "").Replace("(", "").Replace(")", "").ToUpper();

            if (dgv_progress.DataSource != null)
            {
                for (int i = 0; i < dgv_progress.Rows.Count; i++)
                {
                    if (dgv_progress.Rows[i].Cells["Sheet"].Value.ToString().Replace(" ", "").Replace("(", "").Replace(")", "").ToUpper() == sheet && en_exp == true && dgv_progress.Rows[i].Cells["Progress"].Value.ToString() == "x")
                    {
                        dgv_progress.Rows[i].Cells["Export Success"].Style.BackColor = Color.Green;
                        break;
                    }
                    else if (dgv_progress.Rows[i].Cells["Sheet"].Value.ToString().Replace(" ", "").Replace("(", "").Replace(")", "").ToUpper() == sheet && en_exp == false && dgv_progress.Rows[i].Cells["Progress"].Value.ToString() == "x")
                    {
                        dgv_progress.Rows[i].Cells["Export Success"].Style.BackColor = Color.Red;
                        break; ;
                    }
                }
            }

        }
        public void Update_dgv_complete_FAI(string sheet, List<string> lst_sht_OK)
        {
            List<string> FAI_keys_lst = new List<string> { "FAI", "SPC", "CPK", "parentheses" };

            for (int i = 0; i < dgv_progress.Rows.Count; i++)
            {
                string sht = dgv_progress.Rows[i].Cells["Sheet"].Value.ToString();
                if (FAI_keys_lst.Any(x => sht.Contains(x) && lst_sht_OK.Contains(sht) && dgv_progress.Rows[i].Cells["Progress"].Value.ToString() == "x"))
                {
                    dgv_progress.Rows[i].Cells["Export Success"].Style.BackColor = Color.Green;
                }
                else if (FAI_keys_lst.Any(x => sht.Contains(x)) && !lst_sht_OK.Contains(sht) && dgv_progress.Rows[i].Cells["Progress"].Value.ToString() == "x")
                {
                    dgv_progress.Rows[i].Cells["Export Success"].Style.BackColor = Color.Red;
                }
            }
        }



        public void resize_column_image(DataGridView dgv, string sheet)
        {
            if (dgv.Rows.Count > 0)
            {
                if (sheet == "CROSS_SECTION")
                {
                    ((DataGridViewImageColumn)dgv.Columns["Image1"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)dgv.Columns["Image1"]).Width = 100;

                    ((DataGridViewImageColumn)dgv.Columns["Image2"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)dgv.Columns["Image2"]).Width = 100;

                    foreach (DataGridViewRow dr in dgv.Rows)
                    {
                        dr.Height = 70;
                    }
                }
                else if (sheet == "GAP_CONNECTOR")
                {
                    ((DataGridViewImageColumn)dgv.Columns["Image"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)dgv.Columns["Image"]).Width = 100;

                    ((DataGridViewImageColumn)dgv.Columns["Image1"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)dgv.Columns["Image1"]).Width = 100;

                    ((DataGridViewImageColumn)dgv.Columns["Image2"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)dgv.Columns["Image2"]).Width = 100;

                    foreach (DataGridViewRow dr in dgv.Rows)
                    {
                        dr.Height = 70;
                    }
                }
                else
                {

                    ((DataGridViewImageColumn)dgv.Columns["Image"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)dgv.Columns["Image"]).Width = 100;

                    ((DataGridViewImageColumn)dgv.Columns["Graph"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)dgv.Columns["Graph"]).Width = 100;

                    foreach (DataGridViewRow dr in dgv.Rows)
                    {
                        dr.Height = 70;
                    }

                }
            }
        }
        public string remove_special_char(string src_str, List<char> remove_char)
        {
            return new string(src_str.Where(x => remove_char.IndexOf(x) == -1).ToArray());
        }
        public ExcelPackage open_excel(string file_name)
        {
            ExcelPackage myexcel = null;
            FileInfo excel_file = new FileInfo(file_name);
            if (File.Exists(file_name))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                myexcel = new ExcelPackage(excel_file);
            }
            return myexcel;
        }
        public bool CHECK_PEELTEST = false;

        private void btn_export_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "")
            {
                string file_format = find_format(data_loc, txtItemCode.Text);
                if (file_format != "")
                {
                    string f_name = Path.GetFileNameWithoutExtension(file_format);
                    string report_folder = Path.Combine(data_loc, "Report", "NPI");
                    string export_path = Path.Combine(report_folder, f_name + "-" + txtLotNo.Text + Path.GetExtension(file_format));

                    ExcelWorkbook report_saved = null;
                    ExcelPackage sourcePackage = null;
                    ExcelPackage xlPackage = null;
                lbl_export:
                    if (System.IO.File.Exists(export_path))
                    {
                        DialogResult result = MessageBox.Show(new Form { TopMost = true }, "Báo cáo của " + txtItemCode.Text + "-" + txtLotNo.Text + " đã tồn tại. Bạn có muốn cập nhật không?\nYes: Cập nhật\nNo: Tạo mới\nCancel: Thoát", "Thông báo", MessageBoxButtons.YesNoCancel);

                        if (result == DialogResult.Yes)
                        {
                            try
                            {
                                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                                xlPackage = open_excel(export_path);
                                report_saved = xlPackage.Workbook;
                            }
                            catch (Exception ex)
                            {
                                DialogResult resultz = MessageBox.Show("Hiện tại không thể mở thư mục này để chỉnh sửa bạn có muốn ghi đè không?", "Thông báo", MessageBoxButtons.YesNoCancel);
                                if (resultz == DialogResult.Yes)
                                {
                                    System.IO.File.Delete(export_path);
                                    goto lbl_export;
                                }
                                return;
                            }
                        }
                        else if (result == DialogResult.No)
                        {
                            System.IO.File.Delete(export_path);
                            goto lbl_export;
                        }
                        else
                        {
                            return;
                        }
                    }
                    else
                    {
                        if (!System.IO.Directory.Exists(report_folder))
                            System.IO.Directory.CreateDirectory(report_folder);

                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        sourcePackage = new ExcelPackage(new FileInfo(file_format));
                        //sourcePackage.Settings.ImageSettings.PrimaryImageHandler = new OfficeOpenXml.Drawing.GenericImageHandler();
                        report_saved = sourcePackage.Workbook;

                    }

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

                    //MessageBox.Show(new Form { TopMost = true }, "Đang xuất báo cáo. Vui lòng đợi!", "Thông báo", MessageBoxButtons.OK);

                    lst_sheet_export = new List<string> { };
                    string[] skipSheet = new[] { "Coverpage", "Process Comparison", "Mishandling test", "Rev History", "User Guideline", "Low CPK Action", "Declaration", "Deviation summary", "OQC Test", "ORT-Assy", "Electrical", "Switch Quality", "Process flow", "Process Comparison" };

                    foreach (string sheet in cbl_sheet.CheckedItems)
                    {
                        if (!skipSheet.Contains(sheet))
                        {
                            lst_sheet_export.Add(sheet);
                        }
                    }
                    List<string> lst_manual_exp = new List<string> { };
                    List<string> lst_auto_exp = new List<string> { };

                    foreach (string sheet in cbl_sheet.CheckedItems)
                    {
                        lst_auto_exp.Add(sheet.Replace(" ", "").Replace("(", "").Replace(")", "").ToUpper());
                    }
                    //foreach (string sheet in cbl_sheet_manual.CheckedItems)
                    //{
                    //    lst_manual_exp.Add(sheet.Replace(" ", "").Replace("(", "").Replace(")", "").ToUpper());
                    //}

                    for (int i = 0; i < dgv_progress.Rows.Count; i++)
                    {
                        string s = dgv_progress.Rows[i].Cells["Sheet"].Value.ToString().Replace(" ", "").Replace("(", "").Replace(")", "").ToUpper();
                        if (lst_manual_exp.IndexOf(s) != -1 || lst_auto_exp.IndexOf(s) != -1)
                        {
                            dgv_progress.Rows[i].Cells["Export Success"].Style.BackColor = Color.White;

                        }
                        else if (lst_auto_exp.IndexOf("FAI") != -1 && (s.Contains("FAI") || s.Contains("SPC")))
                        {
                            dgv_progress.Rows[i].Cells["Export Success"].Style.BackColor = Color.White;
                        }
                        else
                        {
                            dgv_progress.Rows[i].Cells["Export Success"].Style.BackColor = Color.LightGray;
                        }
                    }

                    DataTable dt_spec_all = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { txtItemCode.Text, "NPI" }));
                    foreach (string item in lst_sheet_export)
                    {
                        if (item != "FAI")
                        {
                            //try
                            //{
                            string sheet = item.Replace(" ", "_").Replace("(", "").Replace(")", "").ToUpper();
                            //DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, sheet, "NPI" }));
                            DataRow rowZ = dt_spec_all.AsEnumerable().FirstOrDefault(x => x["Sheet"].ToString().Trim().Contains(sheet));
                            DataTable dt_spec = dt_spec_all.Clone();
                            if (rowZ != null)
                            {
                                dt_spec.ImportRow(rowZ);
                            }
                            DataTable Data_all = new DataTable();
                            string _process = item.Replace("_", "").Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("&", "").ToUpper();
                            string mySheet = "";
                            foreach (ExcelWorksheet tg_sht in report_saved.Worksheets)
                            {
                                string cur_sht_name = tg_sht.Name.ToUpper().Replace("_", "").Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("&", "");
                                if (_process == cur_sht_name)
                                {
                                    mySheet = tg_sht.Name;
                                    break;
                                }
                            }
                            if (mySheet != "")
                            {
                                ExcelWorksheet ws = report_saved.Worksheets[mySheet];
                                //ws.Activate();
                                bool export_ok = true;
                                try
                                {

                                    switch (mySheet.Trim())
                                    {
                                        case "ACF":
                                            string msgACF = "";
                                            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                                            DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, "Roughness", filter_str);
                                            if (F_expNPI.check_roughness_data(src_dt))
                                            {
                                                bool chk = F_expNPI.check_cpk_roughness(sqlcon, txtItemCode.Text, txtLotNo.Text);
                                            lbl_continue_exp:
                                                if (chk)
                                                {
                                                    try
                                                    {

                                                        F_expNPI.export_ACF_Wetting(ws, txtItemCode.Text, txtLotNo.Text, sqlcon, !cb_legacy.Checked);
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        msgACF += $"ACF WETTING: {ex.Message}";
                                                    }
                                                    try
                                                    {

                                                        F_expNPI.Export_ACF_Peel(ws, txtItemCode.Text, txtLotNo.Text, sqlcon, !cb_legacy.Checked);
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        msgACF += $"ACF peel: {ex.Message}";
                                                    }
                                                    try
                                                    {

                                                        F_expNPI.Export_ACFFlatness(ws, txtItemCode.Text, txtLotNo.Text, sqlcon, !cb_legacy.Checked);
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        msgACF += $"ACF flatness: {ex.Message}";
                                                    }
                                                    try
                                                    {
                                                        F_expNPI.Export_ACF_Roughness(ws, txtItemCode.Text, txtLotNo.Text, sqlcon, !cb_legacy.Checked);
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        msgACF += $"ACF roughness: {ex.Message}";
                                                    }
                                                }
                                                else
                                                {
                                                    if (MessageBox.Show(new Form { TopMost = true }, "Roughness: cpk < 1.33. Tiếp tục xuất dữ liệu ?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                                                    {
                                                        chk = true;
                                                        goto lbl_continue_exp;
                                                    }
                                                    else
                                                    {
                                                        export_ok = false;
                                                    }

                                                }
                                            }
                                            else
                                            {
                                                export_ok = false;
                                            }
                                            break;
                                        case "Flex bending":
                                            if (Bending_Exp.check_bending_data(sqlcon, txtItemCode.Text, txtLotNo.Text, "FLEX_BENDING"))
                                            {
                                                Bending_Exp.Export_Thermal_HeatSoak_Bend_All(sqlcon, txtItemCode.Text, txtLotNo.Text, "FLEX_BENDING", mySheet, 20, "NPI", ws);
                                            }
                                            else
                                            {
                                                export_ok = false;
                                            }

                                            break;
                                        case "Thermal Cycling & bending":
                                            if (Bending_Exp.check_bending_data(sqlcon, txtItemCode.Text, txtLotNo.Text, "THERMAL_CYCLING_AND_BEND"))
                                            {
                                                Bending_Exp.Export_Thermal_HeatSoak_Bend_All(sqlcon, txtItemCode.Text, txtLotNo.Text, "THERMAL_CYCLING_AND_BEND", mySheet, 20, "NPI", ws);
                                            }
                                            else
                                            {
                                                export_ok = false;
                                            }

                                            break;
                                        case "Heat Soak & bending":
                                            if (Bending_Exp.check_bending_data(sqlcon, txtItemCode.Text, txtLotNo.Text, "HEAT_SOAK_AND_BEND"))
                                            {
                                                Bending_Exp.Export_Thermal_HeatSoak_Bend_All(sqlcon, txtItemCode.Text, txtLotNo.Text, "HEAT_SOAK_AND_BEND", mySheet, 20, "NPI", ws);
                                            }
                                            else
                                            {
                                                export_ok = false;
                                            }

                                            break;
                                        case "Packaging":
                                            new PackagingService().ExportToExcel(ws, txtItemCode.Text.Trim());
                                            break;
                                        case "Impedance":
                                            new ImpedanceService().Export(ws, txtItemCode.Text.Trim(), txtLotNo.Text.Trim());
                                            break;
                                        case "SEM BSE & Binarization":
                                            SEMServices.Export(sqlcon, ws, txtItemCode.Text.Trim(), txtLotNo.Text.Trim());
                                            break;
                                        case "OQC B2B Mating-Unmating":
                                            OQCB2BMatingUnmatting.Export(sqlcon, ws, txtItemCode.Text.Trim(), txtLotNo.Text.Trim());
                                            break;
                                        //case 
                                        case "Table of Contents":
                                            TableOfContentService.Export(sqlcon, ws, txtItemCode.Text.Trim(), txtLotNo.Text.Trim());
                                            break;
                                        case "Environment en-durance":
                                            new EEDService(sqlcon).Export(ws, txtItemCode.Text.Trim(), txtLotNo.Text.Trim());
                                            break;
                                        case "Assy Yield":
                                            new AssyYieldService(sqlcon).Export(ws, txtItemCode.Text.Trim(), txtLotNo.Text.Trim());
                                            break;
                                        case "Bar Code Verification":
                                            new BarCodeVertification().Export(sqlcon, ws, txtItemCode.Text.Trim(), txtLotNo.Text.Trim());
                                            break;
                                        case "Air bubble btw Liner-PSA":
                                            if (airService == null)
                                            {
                                                airService = new AirBubbleService(txtItemCode.Text.Trim(), txtLotNo.Text.Trim());
                                                airService.LoadDataRefer();
                                            }

                                            airService.Export(ws, "Liner");

                                            break;
                                        case "Air bubble btw PSA-FPC":
                                            if (airService == null)
                                            {
                                                airService = new AirBubbleService(txtItemCode.Text.Trim(), txtLotNo.Text.Trim());
                                                airService.LoadDataRefer();
                                            }
                                            airService.Export(ws, "PSA");
                                            break;
                                        case "Thermal Cycling":
                                        case "Thermal Shock":
                                        case "Heat Soak and Recovery":
                                            new TCHSTSService().Export(ws, txtItemCode.Text.Trim(), txtLotNo.Text.Trim(), mySheet);
                                            break;
                                        case "X-Ray picture":
                                            string[] types = new[] { "Flex bending", "Thermal Cycling & bending", "Heat Soak & bending" };
                                            foreach (string type in types)
                                            {
                                                string nameSheet = "";
                                                switch (type)
                                                {
                                                    case "Flex bending":
                                                        nameSheet = "X-ray after Bending";
                                                        break;
                                                    case "Thermal Cycling & bending":
                                                        nameSheet = "X-ray after TC & Bending";
                                                        break;
                                                    case "Heat Soak & bending":
                                                        nameSheet = "X-ray after HS & Bending";
                                                        break;
                                                }
                                                sourcePackage.Workbook.Worksheets.Add(nameSheet, ws);
                                                //new XRayPictureService().Export(sourcePackage.Workbook.Worksheets[nameSheet], txtItemCode.Text.Trim(), txtLotNo.Text.Trim(), type);
                                            }
                                            sourcePackage.Workbook.Worksheets.Delete(ws);
                                            break;
                                        default:
                                            string str_filter = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                                            if (sheet.Contains("UNMATING"))
                                            {
                                                if (Itemcode_unmating != "" && LotNo_unmating != "")
                                                {
                                                    str_filter = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { Itemcode_unmating, LotNo_unmating });
                                                }
                                                else
                                                {
                                                    MessageBox.Show(new Form { TopMost = true }, "Chưa nhập ItemCode(NVL) và LotNo(NVL) của (IQC Unmating) Pull Test", "Thông báo");
                                                }

                                            }
                                            else if (sheet.Contains("COUPON") && sheet.Contains("LINER"))
                                            {
                                                if (ItemCode_liner_coupon != "" && LotNo_liner_coupon != "")
                                                {
                                                    str_filter = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode_liner_coupon, LotNo_liner_coupon });
                                                }
                                                else
                                                {
                                                    MessageBox.Show(new Form { TopMost = true }, "Chưa nhập ItemCode(NVL) và LotNo(NVL) của IQC Liner peeling (Coupon)", "Thông báo");
                                                }
                                            }
                                            else if (sheet.Contains("COUPON") && sheet.Contains("PSA"))
                                            {
                                                if (ItemCode_psa_coupon != "" && LotNo_psa_coupon != "")
                                                {
                                                    str_filter = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode_psa_coupon, LotNo_psa_coupon });
                                                }
                                                else
                                                {
                                                    MessageBox.Show(new Form { TopMost = true }, "Chưa nhập ItemCode(NVL) và LotNo(NVL) của IQC PSA peeling (Coupon)", "Thông báo");
                                                }
                                            }
                                            else if (sheet.Contains("ON_PRODUCT"))
                                            {
                                                try
                                                {

                                                    new PeelTestOnProductService(txtItemCode.Text, txtLotNo.Text).Export(ws, !sheet.Contains("LINER"));
                                                    export_ok = true;
                                                    //F_expNPI.export_excel_onproduct(ws, Data_all, dt_spec, sheet, ref export_ok);
                                                }
                                                catch (Exception exZ)
                                                {
                                                    throw new Exception(exZ.Message);
                                                }
                                                break;
                                            }
                                            switch (sheet)
                                            {
                                                case "OQC_B2B_MATING-UNMATING":
                                                    sheet = "OQC_B2B_Mating_Unmating";
                                                    break;
                                                case "SEM_BSE_&_BINARIZATION":
                                                    sheet = "SEM_BSE_Binarization_Logfile";
                                                    break;
                                                default:
                                                    break;
                                            }

                                            DataTable dt_analysis = TDMK_Code.Datatable_Filter(sqlcon, sheet, str_filter);

                                            if (dt_analysis.Rows.Count > 0)
                                            {
                                                string[] arr_val = dt_analysis.AsEnumerable().Select(x => x.Field<string>("Sheet")).Distinct().ToArray();
                                                string filter = "";
                                                foreach (string i in arr_val)
                                                {
                                                    if (i.Contains("NPI"))
                                                    {
                                                        filter = i;
                                                        break;
                                                    }
                                                }
                                                if (filter != "")
                                                {
                                                    Data_all = dt_analysis.AsEnumerable().Where(r => r.Field<string>("Sheet") == filter).CopyToDataTable();

                                                }
                                            }
                                            if (Data_all.Rows.Count > 0 && dt_spec.Rows.Count > 0)
                                            {
                                                try
                                                {

                                                    PIDService.FillProductID(Data_all, txtItemCode.Text, txtLotNo.Text, sheet);
                                                }
                                                catch
                                                {

                                                }

                                                int st = 1;
                                                foreach (DataRow dr in Data_all.Rows)
                                                {
                                                    dr["ID"] = st;
                                                    st++;
                                                }
                                                if (sheet == "GAP_CONNECTOR")
                                                {
                                                    new GAPConnectorService().Export(ws, Data_all, dt_spec);
                                                }
                                                else if (sheet == "CROSS_SECTION")
                                                {
                                                    F_expNPI.export_excel_cross_section(txtItemCode.Text, txtLotNo.Text, ws, Data_all, dt_spec, ref export_ok);

                                                }
                                                //else if (sheet.Contains("ON_PRODUCT"))
                                                //{
                                                //    F_expNPI.export_excel_onproduct(ws, Data_all, dt_spec, sheet, ref export_ok);
                                                //}
                                                else if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST" || sheet == "SHEAR_TEST" || sheet == "Peel test without SUS")
                                                {
                                                    if (sheet == "PEEL_TEST")
                                                    {
                                                        if (Data_all.Rows.Count <= 0)
                                                        {
                                                            CHECK_PEELTEST = false;
                                                            break;
                                                        }
                                                        foreach (DataRow row in Data_all.Rows)
                                                        {
                                                            string numStr = row["Mode 1: Solder joint crack"].ToString().Trim().Split('%')[0];
                                                            double num = double.Parse(numStr);
                                                            if (num > 50)
                                                            {
                                                                CHECK_PEELTEST = true;
                                                                sourcePackage.Workbook.Worksheets.Add("Peel test without SUS", ws);
                                                                break;
                                                            }
                                                        }
                                                    }
                                                    F_expNPI.export_excel_peel_pull_shear(ws, sheet, Data_all, dt_spec, ref export_ok);
                                                }
                                                else if (sheet == "IQC_UNMATING_PULL_TEST")
                                                {
                                                    F_expNPI.export_excel_unmating(ws, Data_all, dt_spec, ref export_ok);
                                                }
                                                else
                                                {
                                                    F_expNPI.export_excel_coupon(ws, Data_all, dt_spec, sheet, ref export_ok);
                                                }

                                            }
                                            else
                                            {
                                                export_ok = false;
                                            }
                                            break;

                                    }

                                    Update_dgv_complete(item, export_ok);
                                    if (export_ok)
                                        ws.TabColor = Color.Green;

                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show($"{mySheet}: {ex.Message}");
                                }


                            }

                            //}
                            //catch
                            //{
                            //    Update_dgv_complete(item, false);
                            //    MessageBox.Show(new Form { TopMost = true }, item + " : Lỗi!", "Thông báo");
                            //}
                        }
                        else
                        {
                            bool chk = FAI_lib.Check_FAI_data(txtItemCode.Text, txtLotNo.Text, sqlcon, 1.67);
                        lbl_export_FAI:
                            if (chk)
                            {
                                List<string> lst_sht_OK = new List<string>();
                                FAI_lib.Export_FAI_Process(sqlcon, report_saved, txtItemCode.Text, txtLotNo.Text, "NPI", 1.67, ref lst_sht_OK);
                                Update_dgv_complete_FAI(item, lst_sht_OK);
                            }
                            else
                            {
                                if (MessageBox.Show(new Form { TopMost = true }, "FAI: cpk < 1.67. Tiếp tục xuất dữ liệu ?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                                {
                                    chk = true;
                                    goto lbl_export_FAI;
                                }
                                else
                                {
                                    // Update_dgv_complete_FAI(item, lst_sht_OK);
                                }
                            }
                        }
                    }
                    string xlFile_path = Path.Combine(data_loc, "Report", "Report_Manual");
                    if (!System.IO.Directory.Exists(xlFile_path))
                        System.IO.Directory.CreateDirectory(xlFile_path);
                    Export_fromExcel_Manual(report_saved, txtItemCode.Text, txtLotNo.Text, xlFile_path);
                    ExportPeelTest(sourcePackage);
                    if (sourcePackage != null)
                    {
                        try
                        {
                            sourcePackage.SaveAs(new FileInfo(export_path));
                        }
                        catch (Exception Ex)
                        {
                            MessageBox.Show($"Error Save: {Ex.Message}");
                        }
                    }
                    else if (xlPackage != null)
                    {
                        xlPackage.Save();
                    }

                    try
                    {
                        myExcel.Workbook wb = TDMK_Code.open_excel_file(export_path, "", "");
                        if (wb != null)
                        {
                            Export_fromExcel_Manual_ignore(wb, txtItemCode.Text, txtLotNo.Text, xlFile_path);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                    MessageBox.Show(new Form { TopMost = true }, "Xuất dữ liệu thành công!", "Thông báo");
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Không tìm thấy format " + txtItemCode.Text, "Thông báo");
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Vui lòng điền đầy đủ thông tin!", "Thông báo");
            }
        }

        private void ExportPeelTest(ExcelPackage package)
        {
            try
            {

                if (CHECK_PEELTEST)
                {
                    foreach (var ws in package.Workbook.Worksheets)
                    {
                        if (ws.Name.Contains("Peel Test"))
                        {

                            int index = ws.Index;
                            package.Workbook.Worksheets.MoveAfter(package.Workbook.Worksheets.Count() - 1, index + 1);

                        }
                        if (ws.Name.Contains("Peel test without SUS"))
                        {
                            DataTable data_all = new DBContext().LoadDataTable("PEEL_TEST_WITHOUT_SUS", new[] { "ItemCode", "LotNo" }, new[] { txtItemCode.Text, txtLotNo.Text });
                            try
                            {
                                PIDService.FillProductID(data_all, txtItemCode.Text, txtLotNo.Text, "PEEL_TEST_WITHOUT_SUS");
                            }
                            catch
                            {

                            }
                            if (data_all.Rows.Count <= 0)
                            {
                                MessageBox.Show("Cần đánh giá peeling without SUS");
                                return;
                            }
                            DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, "PEEL_TEST", "NPI" }));
                            bool export_ok = true;
                            F_expNPI.export_excel_peel_pull_shear(ws, ws.Name, data_all, dt_spec, ref export_ok);
                        }
                    }
                }
            }
            catch
            {

            }
        }

        private void txtItemCode_TextChanged(object sender, EventArgs e)
        {
            txtItemCode.Text = txtItemCode.Text.Trim();
            dgv_progress.DataSource = null;
        }

        private void dgv_progress_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int col_inx = e.ColumnIndex;
            if (col_inx != -1)
            {
                if (dgv_progress.Columns[col_inx].Name == "Sheet")
                {
                    string sheet = dgv_progress.Rows[e.RowIndex].Cells["Sheet"].Value.ToString().Replace(" ", "_").Replace("(", "").Replace(")", "").ToUpper();

                    DataTable Data_all = new DataTable();
                    if (sheet != "ACF" && !sheet.Contains("BENDING") && sheet != "FAI" && lst_auto_new.Contains(dgv_progress.Rows[e.RowIndex].Cells["Sheet"].Value.ToString().Replace(" ", "").Replace("(", "").Replace(")", "").ToUpper()))
                    {
                        DataTable dt_analysis = TDMK_Code.Datatable_Filter(sqlcon, sheet, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                        if (dt_analysis.Rows.Count > 0)
                        {
                            string[] arr_val = dt_analysis.AsEnumerable().Select(x => x.Field<string>("Sheet")).Distinct().ToArray();
                            string filter = "";
                            foreach (string i in arr_val)
                            {
                                if (i.Contains("NPI"))
                                {
                                    filter = i;
                                    break;
                                }
                            }
                            if (filter != "")
                            {
                                Data_all = dt_analysis.AsEnumerable().Where(r => r.Field<string>("Sheet") == filter).CopyToDataTable();

                            }
                        }
                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Không có dữ liệu", "Thông báo");
                        }
                    }
                    else if (sheet.Contains("BENDING"))
                    {
                        string process_name = "";
                        switch (sheet.Replace("&", "").Replace("_", ""))
                        {
                            case "FLEXBENDING":
                                process_name = "FLEX_BENDING";
                                break;

                            case "THERMALCYCLINGBENDING":
                                process_name = "THERMAL_CYCLING_AND_BEND";
                                break;

                            case "HEATSOAKBENDING":
                                process_name = "HEAT_SOAK_AND_BEND";
                                break;

                        }
                        //if(process_name != "")
                        //{
                        //    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                        //    DataTable _src_tbl = TDMK_Code.Datatable_Filter(sqlcon, process_name, filter_str);
                        //    DataTable _spec_tbl = TDMK_Code.Datatable_Filter(sqlcon, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { txtItemCode.Text, process_name }));
                        //    SortedDictionary<int, string> sel_index_lst = Bending_Exp.Get_bending_Net(sqlcon, txtItemCode.Text, process_name);
                        //    List<int> index_lst = sel_index_lst.Keys.ToList();
                        //    Data_all = _src_tbl.AsEnumerable().Where(x => index_lst.Contains(Convert.ToInt32(x.Field<string>("Net_no")) - 1)).CopyToDataTable();
                        //} 
                    }

                    dgv_View.DataSource = Data_all;
                    myCode.Disable_Sort_DGV(dgv_View);
                    resize_column_image(dgv_View, sheet);
                    sheet_select = sheet;
                }
            }
        }

        private void dgv_View_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dgv_View.Rows.Count > 0 && sheet_select != "")
            {
                Hide_column(dgv_View, ref hide_mode);
                resize_column_image(dgv_View, sheet_select);
            }
        }

        private void lblFAI_Click(object sender, EventArgs e)
        {

            Get_Latest_Version(Path.Combine(app_path, "OK2SHIP_Measurements"), "OK2SHIP_Measurements");

        }
        private void lblBending_Click(object sender, EventArgs e)
        {
            Get_Latest_Version(Path.Combine(app_path, "Bending_Items"), "Bending_Items");

        }

        private void lblType3_Click(object sender, EventArgs e)
        {

            Get_Latest_Version(Path.Combine(app_path, "VHX-IMADA"), "VHX-IMADA");

        }


        public void Get_Latest_Version(string folder_, string name)
        {
            DirectoryInfo tar_parent = new DirectoryInfo(folder_);
            FileInfo[] temp_lst = tar_parent.GetFiles("*.exe");
            string file_name = "";
            double ver = 0;

            foreach (FileInfo f in temp_lst)
            {
                if (f.Name.Contains(name))
                {
                    string v = "";
                    char[] ch_arr = Path.GetFileNameWithoutExtension(f.Name).ToCharArray();
                    for (int t = 0; t < ch_arr.Length; t++)
                    {
                        string a = ch_arr[t].ToString();
                        if (myCode.IsNumeric(a))
                        {
                            v += a;
                        }
                    }
                    if (Double.Parse(v) > ver)
                    {
                        ver = Double.Parse(v);
                        file_name = f.Name;
                    }
                }
            }

            if (file_name != "")
            {
                string process_name = Path.Combine(folder_, file_name);
                if (File.Exists(process_name))
                {
                    Process.Start(process_name);
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Không tìm thấy chương trình " + name, "Thông báo");
            }
        }

        private void lblSetting_Click(object sender, EventArgs e)
        {
            Setup_Spec_SMT frmsetup = new Setup_Spec_SMT();
            frmsetup.Show();
            // this.Hide();
        }
        public Boolean Check_Exist_report_Manual(string sheet, string ItemCode, string LotNo, string tar_folder)
        {
            bool chk = false;
            string[] folders = System.IO.Directory.GetDirectories(tar_folder);
            string tar_file_name = ItemCode + "-" + LotNo;
            foreach (string f in folders)
            {
                if (Path.GetFileName(f).Replace(" ", "").Replace("(", "").Replace(")", "").ToUpper() == sheet)
                {
                    string[] list_file = Directory.GetFiles(f, "*.xlsx").Where(s => s.Contains(tar_file_name)).ToArray();
                    if (list_file.Length > 0)
                        return true;

                    break;
                }
            }
            return chk;
        }

        public void Export_fromExcel_Manual(ExcelWorkbook myWrkbook, string ItemCode, string LotNo, string xlFile_path)
        {
            string tar_folder = xlFile_path;
            string[] folders = System.IO.Directory.GetDirectories(tar_folder);
            string tar_file_name = ItemCode + "-" + LotNo;
            /***************************************************************** Old Version **************************************************************************************/
            List<string> auto_lst = new List<string>();// { "FAI", "SPC", "CPK", "IPQC", "Recycle", "Materials", "Coverpage", "User Guidelines", "Declaration and Contents" };
            auto_lst = lst_auto;
            List<string> lst_checked_manual = new List<string>();
            foreach (string sheet in cbl_sheet.CheckedItems)
            {
                lst_checked_manual.Add(sheet.Replace(" ", "").Replace("(", "").Replace(")", "").ToUpper());
            }
            List<string> lst_ignore = new List<string>() { "Thermal Cycling", "Thermal Shock", "Heat Soak and Recovery" };

            foreach (ExcelWorksheet wrksheet in myWrkbook.Worksheets)
            {
                string sel_tbl = wrksheet.Name;
                int sht_inx = wrksheet.Index;
                if (sht_inx > 0)
                {
                    string bef_sht_name = myWrkbook.Worksheets[sht_inx - 1].Name;
                    foreach (string c in folders)
                    {
                        string f_sheet = Path.GetFileName(c).Replace(" ", "").Replace("(", "").Replace(")", "").ToUpper();
                        if (f_sheet.Contains(sel_tbl.Replace(" ", "").Replace("(", "").Replace(")", "").ToUpper().Trim()))
                        {
                            int iz = TDMK_Code.check_exist_list_index(f_sheet, auto_lst);
                            int izz = TDMK_Code.check_exist_list_index(f_sheet, lst_checked_manual);
                            if (TDMK_Code.check_exist_list_index(f_sheet, auto_lst) == -1 && TDMK_Code.check_exist_list_index(f_sheet, lst_checked_manual) != -1) //new version: if (TDMK_Code.check_exist_list_index(c, ignoredList) != -1)
                            {
                                string[] list_file = Directory.GetFiles(c, "*.xlsx").Where(s => s.Contains(tar_file_name)).ToArray();
                                if (list_file.Length > 0 && lst_ignore.IndexOf(sel_tbl) == -1)
                                {
                                    try
                                    {
                                        ExcelWorksheet mywrksheet = myWrkbook.Worksheets[sel_tbl];
                                        ExcelWorkbook sel_wrkbook = TDMK_EPPLUS.open_excel_file(list_file[0]);
                                        ExcelWorksheet sel_wrksheet = sel_wrkbook.Worksheets[0];
                                        myWrkbook.Worksheets.Delete(sel_tbl);
                                        var c1 = myWrkbook.Worksheets.Add(sel_tbl, sel_wrksheet);
                                        c1.View.SetTabSelected(false);
                                        myWrkbook.Worksheets.MoveAfter(sel_tbl, bef_sht_name);
                                        myWrkbook.Worksheets[sel_tbl].TabColor = Color.Green;
                                        Update_dgv_complete(sel_tbl, true);
                                    }
                                    catch
                                    {
                                        MessageBox.Show(new Form { TopMost = true }, sel_tbl + " :Lỗi!", "Thông báo");
                                    }
                                }
                            }
                            break;
                        }
                    }
                }
            }
        }

        public void Export_fromExcel_Manual_ignore(myExcel.Workbook myWrkbook, string ItemCode, string LotNo, string xlFile_path)
        {
            string tar_folder = xlFile_path;
            string[] folders = System.IO.Directory.GetDirectories(tar_folder);
            string tar_file_name = ItemCode + "-" + LotNo;
            /***************************************************************** Old Version **************************************************************************************/
            List<string> auto_lst = new List<string>();// { "FAI", "SPC", "CPK", "IPQC", "Recycle", "Materials", "Coverpage", "User Guidelines", "Declaration and Contents" };
            auto_lst = lst_auto;
            List<string> lst_checked_manual = new List<string>();
            //foreach (string sheet in cbl_sheet_manual.CheckedItems)
            //{
            //    lst_checked_manual.Add(sheet.Replace(" ", "").Replace("(", "").Replace(")", "").ToUpper());
            //}
            List<string> ignoredList = new List<string>() { "Thermal Cycling", "Thermal Shock", "Heat Soak and Recovery" };

            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            foreach (myExcel.Worksheet wrksheet in myWrkbook.Sheets)
            {
                string sel_tbl = wrksheet.Name;
                int sht_inx = wrksheet.Index;
                foreach (string c in folders)
                {
                    if (c.Contains(sel_tbl.Trim()))
                    {
                        if (TDMK_Code.check_exist_list_index(c, ignoredList) != -1 && TDMK_Code.check_exist_list_index(c, lst_checked_manual) != -1)
                        {
                            string[] list_file = Directory.GetFiles(c, "*.xlsx").Where(s => s.Contains(tar_file_name)).ToArray();
                            if (list_file.Length > 0)
                            {
                                myExcel.Worksheet mywrksheet = myWrkbook.Sheets[sel_tbl];
                                myExcel.Workbook sel_wrkbook = TDMK_Code.open_excel_file(list_file[0], "", "");
                                myExcel.Worksheet sel_wrksheet = sel_wrkbook.Sheets[1];
                                xlsApp.DisplayAlerts = false;
                                mywrksheet.Delete();
                                xlsApp.DisplayAlerts = true;
                                sel_wrksheet.Copy(Before: myWrkbook.Sheets[sht_inx]);
                                myWrkbook.Sheets[sel_tbl].Tab.Color = Color.Green;
                                sel_wrkbook.Close(false);
                                Update_dgv_complete(sel_tbl, true);
                            }
                        }
                        break;
                    }
                }
            }
        }


        private void GBInfo_Enter(object sender, EventArgs e)
        {

        }

        private void cbl_sheet_manual_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cbAll_manual_CheckedChanged(object sender, EventArgs e)
        {
            //if (cbAll_manual.Checked == true)
            //{
            //    for (int i = 0; i < cbl_sheet_manual.Items.Count; i++)
            //    {
            //        cbl_sheet_manual.SetItemChecked(i, true);
            //    }

            //}
            //else
            //{
            //    for (int i = 0; i < cbl_sheet_manual.Items.Count; i++)
            //    {
            //        cbl_sheet_manual.SetItemChecked(i, false);
            //    }
            //    // lst_sheet_export = new List<string> { };
            //}
        }

        private void btn_convert_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "")
            {
                NVL_SP fr1 = new NVL_SP(getval) { TopMost = true };
                fr1.ItemCodeSMT_ = txtItemCode.Text;
                fr1.LotNoSMT_ = txtLotNo.Text;
                fr1.infor_item_lot_ = Itemcode_unmating + "^" + LotNo_unmating + "^" + ItemCode_liner_coupon + "^" + LotNo_liner_coupon + "^" + ItemCode_psa_coupon + "^" + LotNo_psa_coupon;
                fr1.Show();
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Vui lòng điền đầy đủ thông tin!", "Thông báo");
            }
        }

        private void lbltitle_Click(object sender, EventArgs e)
        {

        }

        private void tabMain_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void btn_open_Click(object sender, EventArgs e)
        {
            using (ExportManager form = new ExportManager())
            {
                form.ShowDialog();
            }

        }
    }
}