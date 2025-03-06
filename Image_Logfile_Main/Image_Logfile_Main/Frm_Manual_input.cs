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
using myExcel = Microsoft.Office.Interop.Excel;
using VHX;
using OK2SHIP;
using static System.ComponentModel.Design.ObjectSelectorEditor;
using System.Security.Policy;
using System.Diagnostics.Eventing.Reader;
using System.Diagnostics;
using OK2SHIP_SMT;
using OfficeOpenXml;
using TDMK_EPPLUS_7;
using OK2SHIP_Lib;

namespace Manual_Input
{
    public partial class Frm_Manual_input : Form
    {
        myVar exp_proc = new myVar();
        public Frm_Manual_input()
        {
            InitializeComponent();
        }
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public SEI_Lib myCode = new SEI_Lib();
        public EPPlus_Lib ExportLib = new EPPlus_Lib();
        TDMK_EPPLUS7_lib Excel_lib = new TDMK_EPPLUS7_lib();
        TDMK_OK2SHIP TDMK_OK2SHIP = new TDMK_OK2SHIP();
        SqlConnection sqlcon = null;
        Dictionary<string, DataTable> result_bHast;
        string logfile_path = "";
        int indx_select = 0;
        public string depart = "QA";
        public bool admin_mode = false;
        public List<string> sel_pcs_lst = new List<string> { };
        bool update_mode_bhast = false;

        private void Frm_Manual_input_Load(object sender, EventArgs e)
        {
            cb_sheet.Items.Add("ACF Flatness");
            //cb_sheet.Items.Add("ACF Roughness");
            cb_sheet.Items.Add("Moisture Absorption");
            cb_sheet.Items.Add("CQRA_BHAST");
            sqlcon = exp_proc.initial_data(myVar.sel_DB, true);
        }
        public void PasteClipboardValue(bool _transpose, DataGridView tar_DGV)
        {
            if (tar_DGV.SelectedCells.Count == 0)
            {
                MessageBox.Show("Please select a cell", "Paste", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            DataGridViewCell startCell = myCode.GetStartCell(tar_DGV);
            Dictionary<int, Dictionary<int, string>> dictionary = myCode.ClipBoardValues(Clipboard.GetText());
            if (dictionary.Count > 0)
            {
                if (!_transpose)
                {
                    int num = startCell.RowIndex;
                    foreach (int key in dictionary.Keys)
                    {
                        int num2 = startCell.ColumnIndex;
                        foreach (int key2 in dictionary[key].Keys)
                        {
                            if (num2 <= tar_DGV.Columns.Count - 1 && num <= tar_DGV.Rows.Count - 1)
                            {
                                DataGridViewCell dataGridViewCell = tar_DGV[num2, num];
                                if (dataGridViewCell.Selected)
                                {
                                    dataGridViewCell.Value = dictionary[key][key2];

                                }
                            }

                            num2++;
                        }

                        num++;
                    }

                    return;
                }

                int num5 = startCell.ColumnIndex;
                foreach (int key3 in dictionary.Keys)
                {
                    int num6 = startCell.RowIndex;
                    foreach (int key4 in dictionary[key3].Keys)
                    {
                        if (num5 <= tar_DGV.Columns.Count - 1 && num6 <= tar_DGV.Rows.Count - 1)
                        {
                            DataGridViewCell dataGridViewCell2 = tar_DGV[num5, num6];
                            if (dataGridViewCell2.Selected)
                            {
                                dataGridViewCell2.Value = dictionary[key3][key4];

                            }
                        }
                        num6++;
                    }

                    num5++;
                }
            }
            else
            {
                MessageBox.Show("PLease, select data!", "Warning");
            }
        }


        private void btn_load_input_Click(object sender, EventArgs e)
        {
            indx_select = cb_sheet.SelectedIndex;
            DGV_input.DataSource = null;
            //txtqty.Text = ""; 

            if (txtItemCode.Text != "" && txtLotNo.Text != "" && txt_Operator.Text != "")
            {
                if (cb_sheet.SelectedIndex != -1)
                {
                    switch (indx_select)
                    {
                        case 0:
                            DataTable dt = new DataTable();
                            dt = TDMK_Code.Datatable_Filter(sqlcon, "ACF_Flatness", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                            if (dt.Rows.Count == 0)
                            {
                                for (int i = 0; i < 5; i++)
                                {
                                    DataRow dr = dt.NewRow();
                                    dr[0] = (i + 1).ToString();
                                    dr[1] = txtItemCode.Text;
                                    dr[2] = txtLotNo.Text;
                                    dr[3] = "Sample" + (i + 1).ToString();
                                    dt.Rows.Add(dr);
                                }
                                DGV_input.ReadOnly = false;
                            }
                            DGV_input.DataSource = dt;
                            DGV_input.Columns["Flatness"].ReadOnly = true;
                            DGV_input.Columns["Judgement"].ReadOnly = true;
                            btncheck.Visible = true;
                            break;
                        case 1:
                            //DataTable dt2 = new DataTable();
                            //dt2 = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_MOISTURE_ABSORPTION", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                            //if (dt2.Rows.Count == 0)
                            //{
                            //    for (int i = 0; i < 5; i++)
                            //    {
                            //        DataRow dr = dt2.NewRow();
                            //        dr[0] = (i + 1).ToString();
                            //        dr[1] = txtItemCode.Text;
                            //        dr[2] = txtLotNo.Text;

                            //        dt2.Rows.Add(dr);
                            //    }
                            //    DGV_input.ReadOnly = false;
                            //}
                            //DGV_input.DataSource = dt2;

                            DataTable dt2 = new DataTable();
                            dt2 = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_MOISTURE_ABSORPTION", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                            dt2.Columns.Add("%");
                            dt2.Columns.Add("Judgement");
                            if (dt2.Rows.Count == 0)
                            {
                                for (int i = 0; i < 5; i++)
                                {
                                    DataRow dr = dt2.NewRow();
                                    dr[0] = (i + 1).ToString();
                                    dr[1] = txtItemCode.Text;
                                    dr[2] = txtLotNo.Text;


                                    dt2.Rows.Add(dr);
                                }
                                DGV_input.ReadOnly = false;
                            }
                            DGV_input.DataSource = dt2;


                            break;

                        case 2:
                            DataTable bHast_dt = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_BHAST_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));

                            if (bHast_dt.Rows.Count > 0)
                            {
                                bool check_enough = true;
                                for (int i = 1; i <= 8; i++)
                                {
                                    if (myCode.checkDBNull(bHast_dt.Rows[0]["Ch0" + i.ToString()]) == "")
                                    {
                                        check_enough = false;
                                        break;

                                    }
                                }
                                if (!check_enough)
                                {
                                    OpenFileDialog f_open = new OpenFileDialog();
                                    f_open.Filter = "Excel(*.csv)|*.csv";
                                    f_open.InitialDirectory = Application.StartupPath;
                                    if (f_open.ShowDialog() == DialogResult.OK)
                                    {
                                        if (f_open.FileName != "")
                                        {
                                            logfile_path = f_open.FileName;
                                            string logfile_name = Path.GetFileNameWithoutExtension(f_open.FileName);
                                            string[] ItemCode_arr = logfile_name.Split('+');
                                            bool check_name = false;
                                            foreach (var x in ItemCode_arr)
                                            {
                                                string itemlot = x.Split('(')[0];
                                                string item = itemlot.Split('-')[0];
                                                //  string lot = itemlot.Substring(item.Length + 1);
                                                string lot = itemlot.Replace(item + "-", string.Empty);
                                                lot = exp_proc.Lotno_Formated(lot);

                                                if (item == txtItemCode.Text && (lot == txtLotNo.Text))
                                                {
                                                    check_name = true;
                                                    break;
                                                }
                                            }


                                            if (check_name)
                                            {
                                                result_bHast = bHast_Logfile_Process(f_open.FileName);
                                                lstItem.DataSource = result_bHast.Keys.ToList();
                                            }
                                            else
                                            {
                                                MessageBox.Show("Logfile does not contain data for ItemCode: " + txtItemCode.Text + "_LotNo: " + txtLotNo.Text, "Warning", MessageBoxButtons.YesNo);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    DGV_input.DataSource = bHast_dt;
                                    MessageBox.Show("Data is exist. Select chanel to update ?", "Warning", MessageBoxButtons.YesNo);
                                }

                            }
                            else
                            {
                                OpenFileDialog f_open = new OpenFileDialog();
                                f_open.Filter = "Excel(*.csv)|*.csv";
                                f_open.InitialDirectory = Application.StartupPath;
                                if (f_open.ShowDialog() == DialogResult.OK)
                                {
                                    if (f_open.FileName != "")
                                    {
                                        logfile_path = f_open.FileName;
                                        string logfile_name = Path.GetFileNameWithoutExtension(f_open.FileName);
                                        string[] ItemCode_arr = logfile_name.Split('+');
                                        bool check_name = false;
                                        foreach (var x in ItemCode_arr)
                                        {
                                            string itemlot = x.Split('(')[0];
                                            string item = itemlot.Split('-')[0];
                                            // string lot = itemlot.Substring(item.Length + 1);
                                            string lot = itemlot.Replace(item + "-", string.Empty);
                                            lot = exp_proc.Lotno_Formated(lot);

                                            if (item == txtItemCode.Text && (lot == txtLotNo.Text))
                                            {
                                                check_name = true;
                                                break;
                                            }
                                        }


                                        if (check_name)
                                        {
                                            result_bHast = bHast_Logfile_Process(f_open.FileName);
                                            lstItem.DataSource = result_bHast.Keys.ToList();
                                        }
                                        else
                                        {
                                            MessageBox.Show("Logfile does not constain data for ItemCode: " + txtItemCode.Text + "_LotNo: " + txtLotNo.Text, "Warning", MessageBoxButtons.YesNo);
                                        }

                                    }
                                }
                            }


                            break;
                    }
                }
                else
                {
                    MessageBox.Show("Please, Choose region", "Warning");
                }


            }

            else
            {
                MessageBox.Show("Please, fill in ItemCode / LotNo / Operator ID", "Warning");
            }



            //string targetPath = @"E:\Sample.xlsx";
            //SqlConnection sql_IPQC =  exp_proc.initial_data("IPQC_Data", true);
            //myExcel.Workbook curr_wrk = TDMK_Code.open_excel_file(targetPath, "", "");
            //string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode"}, new string[] { txtItemCode.Text});
            //DataTable dt = new DataTable();
            //dt = TDMK_Code.Datatable_Filter(sql_IPQC, "Roughness", filter_str).Clone();
            //myExcel.Worksheet ws = curr_wrk.Sheets[1];
            //myExcel.Range rgn = ws.Range["A1"];

            //for (int i = 1; i < 161; i++)
            //{
            //    DataRow dr = dt.NewRow();
            //    for (int j = 0; j < dt.Columns.Count; j++)
            //    {                    
            //        dr[j] = myCode.checkDBNull(rgn.Offset[i, j].Value);
            //    }
            //    dt.Rows.Add(dr);
            //}

            //exp_proc.BatchBulkCopy(sql_IPQC, dt, "Roughness");



        }
        public List<string> get_list_lotno_roughness()
        {

            SqlConnection sqlcon_IPQC = exp_proc.initial_data("IPQC_Data", true);
            DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon_IPQC, "Roughness", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
            sqlcon = exp_proc.initial_data(myVar.sel_DB, true);
            List<string> lst_lotno = new List<string> { };
            lst_lotno.AddRange(src_dt.AsEnumerable().Select(r => r.Field<string>("LotNo")).Distinct().ToArray());
            return lst_lotno;
        }

        private void cb_sheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            btncheck.Visible = false;
            DGV_input.DataSource = null;
            DGV_input.ReadOnly = true;
            // txtqty.Text = "";
            // label2.Visible = false;
            // txtqty.Visible = false;
            switch (cb_sheet.SelectedIndex)
            {
                case 0:
                    lbltitle.Text = "ACF Flatness";
                    // cb_lotno_roughness.Visible = false;
                    btn_load_input.Text = "INPUT";
                    lstItem.DataSource = null;
                    lstItem.Enabled = false;
                    result_bHast = null;
                    btn_load_input.Enabled = true;
                    btn_save_db.Enabled = true;
                    break;


                case 1:
                    lbltitle.Text = "Moisture Absorption";
                    btn_load_input.Text = "INPUT";
                    lstItem.DataSource = null;
                    lstItem.Enabled = false;
                    result_bHast = null;
                    btn_load_input.Enabled = true;
                    btn_save_db.Enabled = true;
                    //  label2.Visible = true;
                    // txtqty.Visible = true;

                    //DataTable dt_setting = TDMK_Code.Datatable_Filter(sqlcon, "SETTING_PCS", TDMK_Code.filter_str(new string[] { "ItemCode", "Data_For" }, new string[] { txtItemCode.Text, "CQRA - Moisture absorption" }));
                    //if (dt_setting.Rows.Count > 0)
                    //{
                    //    txtqty.Text = dt_setting.Rows[0]["Pcs_setting"].ToString();
                    //}

                    break;
                case 2:
                    lbltitle.Text = "BHAST Log file data";
                    // cb_lotno_roughness.Visible = false;
                    btn_load_input.Text = "Browse....";
                    lstItem.DataSource = null;
                    result_bHast = null;
                    lstItem.Enabled = true;
                    btn_load_input.Enabled = true;
                    btn_save_db.Enabled = true;
                    break;
            }
        }
        public double flatness(Double[] arr_sample)
        {
            double a = arr_sample.Min();
            double b = arr_sample.Max();
            return b - a;
        }

        private void btncheck_Click(object sender, EventArgs e)
        {
            double val = 0;
            if (DGV_input.Rows.Count > 0)
            {
                for (int row = 0; row < DGV_input.Rows.Count; row++)
                {
                    List<double> arr_sample = new List<double> { };
                    for (int i = 4; i < DGV_input.Columns.Count - 2; i++)
                    {
                        double value;
                        if (Double.TryParse(DGV_input.Rows[row].Cells[i].Value.ToString(), out value))
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
                        DGV_input.Rows[row].Cells["Flatness"].Value = val;
                        if (val < 15)
                        {
                            DGV_input.Rows[row].Cells["Judgement"].Value = "Pass";
                            DGV_input.Rows[row].Cells["Judgement"].Style.BackColor = Color.White;
                        }
                        else
                        {
                            DGV_input.Rows[row].Cells["Judgement"].Value = "Fail";
                            DGV_input.Rows[row].Cells["Judgement"].Style.BackColor = Color.Red;
                        }
                    }

                }
            }
        }
        public void save_flatness_db()
        {
            string[] arr_item = new string[16];
            arr_item[0] = "ID";
            arr_item[1] = "ItemCode";
            arr_item[2] = "LotNo";
            arr_item[3] = "Sample";
            for (int i = 4; i < 14; i++)
            {

                arr_item[i] = "Point" + (i - 3).ToString();
            }
            arr_item[14] = "Flatness";
            arr_item[15] = "Judgement";


            string[] arr_item_val = new string[16];

            bool check = true;
            for (int i = 0; i < 5; i++)
            {
                for (int j = 4; j < 14; j++)
                {
                    if (DGV_input.Rows[i].Cells[j].Value.ToString() == "")
                    {
                        check = false;
                        break;
                    }
                }
                if (!check)
                {
                    MessageBox.Show("Not enough data ", "Warning");
                    break;
                }

            }
            if (check)
            {
                for (int i = 0; i < 5; i++)
                {
                    arr_item_val[0] = (TDMK_Code.SQL_MAX("ACF_Flatness", "ID", sqlcon) + 1).ToString();
                    for (int j = 1; j < DGV_input.Columns.Count; j++)
                    {
                        arr_item_val[j] = DGV_input.Rows[i].Cells[j].Value.ToString();

                    }

                    TDMK_Code.insert_val_arr("ACF_Flatness", sqlcon, arr_item, arr_item_val);
                }
                MessageBox.Show(new Form { TopMost = true }, "Save data completed", "Warning");

            }
        }


        public void save_Moisture_db_old()
        {
            string[] arr_item = { "ID", "ItemCode", "LotNo", "Weight_before", "Weight_after" };
            string[] arr_item_val = new string[5];


            bool check = true;
            for (int i = 0; i < 5; i++)
            {
                for (int j = 3; j < 5; j++)
                {
                    if (DGV_input.Rows[i].Cells[j].Value.ToString() == "")
                    {
                        check = false;
                        break;
                    }
                }
                if (!check)
                {
                    MessageBox.Show("Not enough data ", "Warning");
                    break;
                }

            }
            if (check)
            {
                for (int i = 0; i < 5; i++)
                {
                    arr_item_val[0] = (TDMK_Code.SQL_MAX("CQRA_MOISTURE_ABSORPTION", "ID", sqlcon) + 1).ToString();
                    for (int j = 1; j < DGV_input.Columns.Count; j++)
                    {
                        arr_item_val[j] = DGV_input.Rows[i].Cells[j].Value.ToString();

                    }

                    TDMK_Code.insert_val_arr("CQRA_MOISTURE_ABSORPTION", sqlcon, arr_item, arr_item_val);
                }
                MessageBox.Show(new Form { TopMost = true }, "Save data completed", "Warning");
            }
        }
        public void save_Moisture_db()
        {
            string[] arr_item = { "ID", "ItemCode", "LotNo", "Weight_before", "Weight_after" };
            string[] arr_item_val = new string[5];


            bool check = true;
            for (int i = 0; i < 5; i++)
            {
                for (int j = 3; j < 5; j++)
                {
                    if (DGV_input.Rows[i].Cells[j].Value.ToString() == "")
                    {
                        check = false;
                        break;
                    }
                }
                if (!check)
                {
                    MessageBox.Show("Not enough data ", "Warning");
                    break;
                }

            }
            if (check)
            {
                for (int i = 0; i < 5; i++)
                {
                    arr_item_val[0] = (TDMK_Code.SQL_MAX("CQRA_MOISTURE_ABSORPTION", "ID", sqlcon) + 1).ToString();
                    for (int j = 1; j < 4; j++)
                    {
                        arr_item_val[j] = DGV_input.Rows[i].Cells[j].Value.ToString();

                    }
                    arr_item_val[4] = DGV_input.Rows[i].Cells[4].Value.ToString() + "(" + DGV_input.Rows[i].Cells["%"].Value.ToString() + ";" + DGV_input.Rows[i].Cells["Judgement"].Value.ToString() + ")";

                    TDMK_Code.insert_val_arr("CQRA_MOISTURE_ABSORPTION", sqlcon, arr_item, arr_item_val);
                }
                MessageBox.Show(new Form { TopMost = true }, "Save data completed", "Warning");
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        public void Export_ACFFlatness_Excel(SqlConnection sqlcon, string ItemCode, string LotNo)
        {
            if (ItemCode != "" && LotNo != "")
            {
                DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, "ACF_FLATNESS", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                if (src_dt.Rows.Count > 0)
                {
                    string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                    if (format_file != "")
                    {
                        myExcel.Workbook curr_wrkbook = exp_proc.create_export_wrk(format_file, "OQC ACF", txtItemCode.Text, txtLotNo.Text);
                        myExcel.Worksheet curr_wrksheet = curr_wrkbook.Sheets[1];
                        myExcel.Range curr_rgn = curr_wrksheet.Range["A1"];
                        for (int j = 0; j < 5; j++)
                        {
                            curr_rgn = curr_rgn.Offset[0, j];
                            for (int i = 0; i < 100; i++)
                            {

                                if (myCode.checkDBNull(curr_rgn.Offset[i, 0].Value).Contains("Point"))
                                {
                                    curr_rgn = curr_rgn.Offset[i, 0];
                                    goto lbl_export;
                                }
                            }
                        }


                    lbl_export:
                        for (int i = 0; i < 5; i++)
                        {
                            for (int j = 0; j < 10; j++)
                            {
                                string sel_val = myCode.checkDBNull(src_dt.Rows[i][4 + j]);
                                if (TDMK_Code.IsNumeric(sel_val))
                                {
                                    curr_rgn.Offset[i + 1, j].Value = Convert.ToDecimal(sel_val) * 1000;
                                }
                                else
                                {
                                    curr_rgn.Offset[i + 1, j].Value = sel_val;
                                }
                            }
                        }
                        MessageBox.Show("Export to Checksheet completed!", "Warning");

                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Format of " + txtItemCode.Text + " not found", "Warning");
                    }


                }
                else
                {
                    MessageBox.Show("No data", "Warning");
                }
            }
        }
        public void Export_ACFFlatness(SqlConnection sqlcon, string ItemCode, string LotNo)
        {
            if (ItemCode != "" && LotNo != "")
            {
                DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, "ACF_FLATNESS", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                if (src_dt.Rows.Count > 0)
                {
                    string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                    if (format_file != "")
                    {

                        string export_path = Path.Combine(Application.StartupPath, "Export", cb_sheet.Text);
                        if (!Directory.Exists(export_path))
                        {
                            Directory.CreateDirectory(export_path);
                        }
                        ExcelPackage format_pack = Excel_lib.open_excel(format_file);
                        ExcelWorkbook curr_wrkbook = format_pack.Workbook;
                        ExcelWorksheet ws = null;
                        string _process = "OQC ACF".Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
                        foreach (ExcelWorksheet tg_sht in curr_wrkbook.Worksheets)
                        {
                            string cur_sht_name = tg_sht.Name.ToUpper().Replace("-", "").Replace(" ", "");
                            if (cur_sht_name == _process)
                            {
                                ws = tg_sht;
                                break;
                            }
                        }
                        if (ws != null)
                        {
                            ExcelRangeBase curr_rgn = ws.Cells["A1"];
                            for (int j = 0; j < 5; j++)
                            {
                                curr_rgn = curr_rgn.Offset(0, j);
                                for (int i = 0; i < 100; i++)
                                {

                                    if (myCode.checkDBNull(curr_rgn.Offset(i, 0).Value).Contains("Point"))
                                    {
                                        curr_rgn = curr_rgn.Offset(i, 0);
                                        goto lbl_export;
                                    }
                                }
                            }

                        lbl_export:
                            for (int i = 0; i < 5; i++)
                            {
                                for (int j = 0; j < 10; j++)
                                {
                                    string sel_val = myCode.checkDBNull(src_dt.Rows[i][4 + j]);
                                    if (TDMK_Code.IsNumeric(sel_val))
                                    {
                                        curr_rgn.Offset(i + 1, j).Value = Convert.ToDecimal(sel_val) * 1000;
                                    }
                                    else
                                    {
                                        curr_rgn.Offset(i + 1, j).Value = sel_val;
                                    }
                                }
                            }
                            string export_file = Path.Combine(export_path, cb_sheet.Text + "_" + ItemCode + "-" + LotNo + "_" + txt_Operator.Text + "_" + DateTime.Now.ToString("ddMMMyyyy HHmmss") + ".xlsx");
                            ExcelPackage report_pack = new ExcelPackage(export_file);
                            ExcelWorkbook report_wrkbk = report_pack.Workbook;
                            report_wrkbk.Worksheets.Add(ws.Name, ws);
                            report_pack.Save();
                            report_pack = null;
                            report_wrkbk = null;
                            try
                            {
                                ProcessStartInfo pi = new ProcessStartInfo(export_file);
                                Process.Start(pi);
                            }
                            catch { };
                            MessageBox.Show(new Form { TopMost = true }, "Hoàn thành xuất dữ liệu", "Thông báo");
                        }
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Không tìm thấy Format của ItemCode:  " + txtItemCode.Text, "Thông báo");
                    }
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true },"Không có dữ liệu", "Thông báo");
                }
            }
        }
        public void Export_Moisture_Excel(SqlConnection sqlcon, string ItemCode, string LotNo)
        {
            if (ItemCode != "" && LotNo != "")
            {
                DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_MOISTURE_ABSORPTION", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                if (src_dt.Rows.Count > 0)
                {
                    string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                    if (format_file != "")
                    {
                        myExcel.Workbook curr_wrkbook = exp_proc.create_export_wrk(format_file, "CQRA - Moisture absorption", txtItemCode.Text, txtLotNo.Text);
                        myExcel.Worksheet curr_wrksheet = curr_wrkbook.Sheets[1];


                        for (int col = 1; col < 5; col++)
                        {
                            for (int row = 2; row <= 50; row++)
                            {
                                myExcel.Range cell = curr_wrksheet.Cells[row, col];
                                if (myCode.checkDBNull(cell.Value).ToUpper().Contains("Sample".ToUpper()))
                                {
                                    // myExcel.Range cell2 = wb.Sheets[sheet_name].Cells[row + 2, col];
                                    int i = 0;
                                    while (myCode.IsNumeric(myCode.checkDBNull(curr_wrksheet.Cells[row + 2 + i, col].Value)))
                                    {
                                        curr_wrksheet.Cells[row + 2 + i, col + 1].Value = src_dt.Rows[i]["Weight_before"].ToString();
                                        curr_wrksheet.Cells[row + 2 + i, col + 2].Value = src_dt.Rows[i]["Weight_after"].ToString().Split('(')[0];

                                        string val_after = src_dt.Rows[i]["Weight_after"].ToString();
                                        
                                        if (val_after.Contains("("))
                                        {
                                            curr_wrksheet.Cells[row + 2 + i, col + 3] = val_after.Split('(')[1].Split(';')[0];
                                            curr_wrksheet.Cells[row + 2 + i, col + 4] = val_after.Split('(')[1].Split(';')[1].Replace(")", "");
                                        } 

                                        i++;
                                    }

                                    break;
                                }

                            }

                        }
                         
                        MessageBox.Show("Export to Checksheet completed!", "Warning");
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Format of " + txtItemCode.Text + " not found", "Warning");
                    }


                }
                else
                {
                    MessageBox.Show("No data", "Warning");
                }
            }
        }
        public void Export_Moisture(SqlConnection sqlcon, string ItemCode, string LotNo)
        {
            if (ItemCode != "" && LotNo != "")
            {
                DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_MOISTURE_ABSORPTION", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                if (src_dt.Rows.Count > 0)
                {
                    string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                    if (format_file != "")
                    {
                        string export_path = Path.Combine(Application.StartupPath, "Export", cb_sheet.Text);
                        if (!Directory.Exists(export_path))
                        {
                            Directory.CreateDirectory(export_path);
                        }
                        ExcelPackage format_pack = Excel_lib.open_excel(format_file);
                        ExcelWorkbook curr_wrkbook = format_pack.Workbook;
                        ExcelWorksheet ws = null;
                        string _process = "CQRA - Moisture absorption".Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
                        foreach (ExcelWorksheet tg_sht in curr_wrkbook.Worksheets)
                        {
                            string cur_sht_name = tg_sht.Name.ToUpper().Replace("-", "").Replace(" ", "");
                            if (cur_sht_name == _process)
                            {
                                ws = tg_sht;
                                break;
                            }
                        }
                        if (ws != null)
                        {
                            for (int col = 1; col < 5; col++)
                            {
                                for (int row = 2; row <= 50; row++)
                                {
                                    ExcelRangeBase cell = ws.Cells[row, col];
                                    if (myCode.checkDBNull(cell.Value).ToUpper().Contains("Sample".ToUpper()))
                                    {
                                        // myExcel.Range cell2 = wb.Sheets[sheet_name].Cells[row + 2, col];
                                        int i = 0;
                                        while (myCode.IsNumeric(myCode.checkDBNull(ws.Cells[row + 2 + i, col].Value)))
                                        {
                                            ws.Cells[row + 2 + i, col + 1].Value = src_dt.Rows[i]["Weight_before"].ToString();
                                            ws.Cells[row + 2 + i, col + 2].Value = src_dt.Rows[i]["Weight_after"].ToString().Split('(')[0];

                                            string val_after = src_dt.Rows[i]["Weight_after"].ToString();

                                            if (val_after.Contains("("))
                                            {
                                                ws.Cells[row + 2 + i, col + 3].Value = val_after.Split('(')[1].Split(';')[0];
                                                ws.Cells[row + 2 + i, col + 4].Value = val_after.Split('(')[1].Split(';')[1].Replace(")", "");
                                            }

                                            i++;
                                        }

                                        break;
                                    }

                                }

                            }
                            string export_file = Path.Combine(export_path, cb_sheet.Text + "_" + ItemCode + "-" + LotNo + "_" + txt_Operator.Text + "_" + DateTime.Now.ToString("ddMMMyyyy HHmmss") + ".xlsx");
                            ExcelPackage report_pack = new ExcelPackage(export_file);
                            ExcelWorkbook report_wrkbk = report_pack.Workbook;
                            report_wrkbk.Worksheets.Add(ws.Name, ws);
                            report_pack.Save();
                            report_pack = null;
                            report_wrkbk = null;
                            try
                            {
                                ProcessStartInfo pi = new ProcessStartInfo(export_file);
                                Process.Start(pi);
                            }
                            catch { };
                            MessageBox.Show(new Form { TopMost = true }, "Hoàn thành xuất dữ liệu", "Thông báo");
                        }
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Format of " + txtItemCode.Text + " not found", "Warning");
                    }
                }
                else
                {
                    MessageBox.Show("No data", "Warning");
                }
            }
        }

        public void DB_to_excel(string path, string itemcode, string lotno)
        {

            myExcel.Workbook curr_wrkbook = TDMK_Code.open_excel_file(path, "", "");
            DataTable dt1 = new DataTable();
            dt1 = TDMK_Code.Datatable_Filter(sqlcon, "Roughness", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcode, lotno }));
            List<List<string>> lst_rghness = lst_roughness(dt1);

            myExcel.Worksheet curr_wrksheet_1 = curr_wrkbook.Sheets["ACF"];
            myExcel.Range curr_rgn_1 = curr_wrksheet_1.Range["A1"];

            List<int> vitri_sa = new List<int> { };
            List<int> vitri_sq = new List<int> { };
            List<int> vitri_sdr = new List<int> { };
            string[] arr_info = ACF_pad_location(curr_wrkbook).Split('+');
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

        //public List<List<string>> lst_roughness(DataTable dt)
        //{

        //    List<List<string>> lst_roughness = new List<List<string>>() { };
        //    string[] arr_name = { "L1_Roughness_Sa", "L1_Roughness_Sq", "L1_Roughness_Sdr", "L2_Roughness_Sa", "L2_Roughness_Sq", "L12_Roughness_Sdr", "L3_Roughness_Sa", "L3_Roughness_Sq", "L3_Roughness_Sdr" };
        //    for (int i = 0; i < arr_name.Length; i++)
        //    {
        //        List<string> lst_item = new List<string>() { };
        //        for (int j = 0; j < dt.Rows.Count; j++)
        //        {
        //            lst_item.Add(dt.Rows[j][i + 6].ToString());
        //        }
        //        lst_roughness.Add(lst_item);
        //    }
        //    return lst_roughness;
        //}

        private void button6_Click(object sender, EventArgs e)
        {

        }
        private void btnexport_Click(object sender, EventArgs e)
        {
            indx_select = cb_sheet.SelectedIndex;
            if (cb_sheet.SelectedIndex != -1)
            {
                if (txtItemCode.Text != "" && txtLotNo.Text != "" && txt_Operator.Text != "")
                {

                    switch (indx_select)
                    {
                        case 0:
                            Export_ACFFlatness(sqlcon, txtItemCode.Text, txtLotNo.Text);
                            break;
                        //case 1:
                        //    Export_ACF_Roughness(txtItemCode.Text, txtLotNo.Text);
                        //    break;
                        case 1:
                            Export_Moisture(sqlcon, txtItemCode.Text, txtLotNo.Text);
                            break;
                        case 2:
                            Export_Bhast_Logfile(txtItemCode.Text, txtLotNo.Text, sqlcon);
                            break;
                    }


                }
                else
                {
                    MessageBox.Show("Please, fill in ItemCode / LotNo / Operator ID", "Warning");
                }
            }
            else
            {
                MessageBox.Show("Please, Choose region", "Warning");
            }
        }
        public void new_workbook(string format_path, string file_export_path, string name_sheet)
        {

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook curr_wrkbook = app.Workbooks.Open(format_path, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            myExcel.Worksheet curr_wrksheet = curr_wrkbook.Sheets[name_sheet];

            var newbook = app.Workbooks.Add(1);
            curr_wrksheet.Copy(newbook.Sheets[1]);
            newbook.SaveAs(file_export_path);
            newbook.Close();

            // curr_wrkbook.Close();
            // return curr_wrkbook;
        }

        private void btn_loaddata_Click(object sender, EventArgs e)
        {
            indx_select = cb_sheet.SelectedIndex;
            if (cb_sheet.SelectedIndex != -1)
            {
                if (txtItemCode.Text != "" && txtLotNo.Text != "" && txt_Operator.Text != "")
                {

                    switch (indx_select)
                    {
                        case 0:

                            DGV_input.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "ACF_Flatness", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                            if (DGV_input.Rows.Count == 0)
                            {
                                MessageBox.Show("No data", "Warning");
                            }
                            break;

                        //case 1:

                        //    SqlConnection sqlcon_IPQC = exp_proc.initial_data("IPQC_Data", true);
                        //    DGV_input.DataSource = TDMK_Code.Datatable_Filter(sqlcon_IPQC, "Roughness", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                        //    sqlcon = exp_proc.initial_data(myVar.sel_DB, true);
                        //    if (DGV_input.Rows.Count == 0)
                        //    {
                        //        MessageBox.Show("No data", "Warning");
                        //    }
                        //    break;


                        case 1:
                            //DGV_input.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_MOISTURE_ABSORPTION", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                            DataTable dt2 = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_MOISTURE_ABSORPTION", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                            if (dt2.Rows.Count == 0)
                            {
                                MessageBox.Show("No data", "Warning");
                            }
                            else
                            { 
                                dt2.Columns.Add("%");
                                dt2.Columns.Add("Judgement");
                                for (int i = 0; i < dt2.Rows.Count; i++)
                                {
                                    string val_after = dt2.Rows[i]["Weight_after"].ToString();
                                    dt2.Rows[i]["Weight_after"] = val_after.Split('(')[0];
                                    if (val_after.Contains("("))
                                    {
                                        dt2.Rows[i]["%"] = val_after.Split('(')[1].Split(';')[0];
                                        dt2.Rows[i]["Judgement"] = val_after.Split('(')[1].Split(';')[1].Replace(")", "");
                                    }
                                    else
                                    {
                                        dt2.Rows[i]["%"] = "";
                                        dt2.Rows[i]["Judgement"] = "";
                                    }
                                }
                                DGV_input.DataSource = dt2;

                            }
                            break;
                        case 2:
                            DGV_input.DataSource = null;
                            DataTable bHast_dt = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_BHAST_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                            if (bHast_dt.Rows.Count > 0)
                            {
                                DGV_input.DataSource = bHast_dt;
                                lstItem.DataSource = null;
                                result_bHast = null;
                                myVar._frmbHastGraph = new bHast_Graph(bHast_dt, txtItemCode.Text, txtLotNo.Text);
                                myVar._frmbHastGraph.Show();
                            }
                            else
                            {
                                MessageBox.Show("No data", "Warning");
                            }
                            break;

                    }
                    DGV_input.ReadOnly = true;


                }

                else
                {
                    MessageBox.Show("Please, fill in ItemCode / LotNo / Operator ID", "Warning");
                }
            }
            else
            {
                MessageBox.Show("Please, Choose region", "Warning");
            }
        }
        public void Update_history_ACF_Flatness()
        {
            SortedDictionary<string, List<int>> edited_lst = new SortedDictionary<string, List<int>> { };
            List<int> lst_point = new List<int> { };
            DataTable dt_before = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "ACF_Flatness", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
            DataTable dt_after = (DataTable)DGV_input.DataSource;
            for (int i = 0; i < DGV_input.Rows.Count; i++)
            {
                lst_point = new List<int> { };
                for (int j = 4; j < DGV_input.Columns.Count - 2; j++)
                {
                    if (dt_after.Rows[i][j].ToString() != dt_before.Rows[i][j].ToString())
                    {
                        lst_point.Add(j);
                    }

                }
                edited_lst.Add(dt_after.Rows[i]["Sample"].ToString(), lst_point);
            }

            List<string> item_arr = new List<string>() { "ID", "ItemCode", "LotNo", "Process", "PCS_No", "Region", "Zone", "Before_Data", "After_Data", "Time_Update", "Operator", "Depart", "Remark" };
            List<object> item_arr_val = new List<object>();
            DateTime tim_up = DateTime.Now;
            int id_edit = TDMK_Code.SQL_MAX("VHX_Edit_History", "ID", myVar.mysqlcon);
            int count = 0;
            int no_sample = 0;


            foreach (var item in edited_lst)
            {
                string zone = item.Key.ToString();
                DataView dv = dt_before.AsDataView();

                foreach (var i_point in item.Value)
                {

                    dv.RowFilter = "Sample = '" + zone + "'";
                    if (dv.Count > 0)
                    {
                        DataRow dr = dv[0].Row;
                        int r_x = dt_before.Rows.IndexOf(dr);

                        string before_data = dt_before.Rows[r_x][i_point].ToString();
                        string after_data = dt_after.Rows[no_sample][i_point].ToString();
                        byte[] bef_data = System.Text.Encoding.UTF8.GetBytes(before_data);
                        byte[] aft_data = System.Text.Encoding.UTF8.GetBytes(after_data);

                        item_arr_val = new List<object>() { (id_edit + count + 1), txtItemCode.Text, txtLotNo.Text, "ACF_Flatness", "Point " + (i_point - 3).ToString(), "Logfile_data", zone, bef_data, aft_data, tim_up, txtUsername.Text, depart, "" };
                        exp_proc.Insert_Object_List("VHX_Edit_History", sqlcon, item_arr, item_arr_val);
                        count++;


                    }
                    count++;

                }
                no_sample++;

            }

            //DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { txtItemCode.Text, txtLotNo.Text, "ACF_Flatness" }));
        }
        public void Update_history_Moisture()
        {
            SortedDictionary<string, List<string>> edited_lst = new SortedDictionary<string, List<string>> { };
            List<string> lst_point = new List<string> { };
            DataTable dt_before = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "CQRA_MOISTURE_ABSORPTION", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
            DataTable dt_after = (DataTable)DGV_input.DataSource;


            for (int j = 3; j < 5; j++)
            {
                lst_point = new List<string> { };
                for (int i = 0; i < DGV_input.Rows.Count; i++)
                {
                    if (dt_after.Rows[i][j].ToString() != dt_before.Rows[i][j].ToString())
                    {
                        string id = dt_after.Rows[i]["ID"].ToString();
                        lst_point.Add(id);

                    }
                }
                edited_lst.Add(dt_after.Columns[j].ColumnName, lst_point);

            }


            List<string> item_arr = new List<string>() { "ID", "ItemCode", "LotNo", "Process", "PCS_No", "Region", "Zone", "Before_Data", "After_Data", "Time_Update", "Operator", "Depart", "Remark" };
            List<object> item_arr_val = new List<object>();
            DateTime tim_up = DateTime.Now;
            int id_edit = TDMK_Code.SQL_MAX("VHX_Edit_History", "ID", myVar.mysqlcon);
            int count = 0;




            foreach (var item in edited_lst)
            {
                string zone = item.Key.ToString();
                int num = 0;

                DataView dv = dt_before.AsDataView();

                foreach (var i_point in item.Value)
                {

                    dv.RowFilter = "ID = '" + i_point + "'";
                    if (dv.Count > 0)
                    {
                        DataRow dr = dv[0].Row;
                        int r_x = dt_before.Rows.IndexOf(dr);

                        string before_data = dt_before.Rows[r_x][zone].ToString();
                        string after_data = dt_after.Rows[num][zone].ToString();
                        byte[] bef_data = System.Text.Encoding.UTF8.GetBytes(before_data);
                        byte[] aft_data = System.Text.Encoding.UTF8.GetBytes(after_data);

                        item_arr_val = new List<object>() { (id_edit + count + 1), txtItemCode.Text, txtLotNo.Text, "CQRA_MOISTURE_ABSORPTION", (num + 1).ToString(), "Logfile_data", zone, bef_data, aft_data, tim_up, txtUsername.Text, depart, "" };
                        exp_proc.Insert_Object_List("VHX_Edit_History", sqlcon, item_arr, item_arr_val);
                        count++;
                    }
                    num++;

                }
            }

            // DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { txtItemCode.Text, txtLotNo.Text, "CQRA_MOISTURE_ABSORPTION" }));
        }



        private void btn_savedb_Click(object sender, EventArgs e)
        {
            indx_select = cb_sheet.SelectedIndex;
            if (cb_sheet.SelectedIndex != -1)
            {
                if (txtItemCode.Text != "" && txtLotNo.Text != "" && txt_Operator.Text != "")
                {

                    switch (indx_select)
                    {
                        case 0:
                            if (DGV_input.Rows.Count > 0)
                            {
                                bool chk = true;

                                for (int k = 0; k < DGV_input.Rows.Count; k++)
                                {
                                    if (DGV_input.Rows[k].Cells["Judgement"].Style.BackColor == Color.Red)
                                    {
                                        chk = false;
                                        break;
                                    }

                                }

                            lblsave: if (chk)
                                {
                                    DataTable dt1 = new DataTable();

                                    dt1 = TDMK_Code.Datatable_Filter(sqlcon, "ACF_Flatness", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                                    var result = DialogResult.No;
                                    if (dt1.Rows.Count > 0)
                                    {
                                        result = MessageBox.Show("Data is exist. Do you want to replace ?", "Warning", MessageBoxButtons.YesNo);
                                        if (admin_mode)
                                        {
                                            if (result == DialogResult.Yes)
                                            {
                                                Update_history_ACF_Flatness();
                                                TDMK_Code.Delelte_FilteredItem_arr("ACF_Flatness", sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                                                save_flatness_db();
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show("Please, login to save data", "Warning");
                                        }


                                    }
                                    else
                                    {
                                        save_flatness_db();
                                    }

                                }
                                else
                                {
                                    if (MessageBox.Show(new Form { TopMost = true }, "Data is out of spec. Do you want to save data", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                                    {

                                        if (admin_mode)
                                        {
                                            chk = true;
                                            goto lblsave;

                                        }
                                        else
                                        {
                                            MessageBox.Show("Please, login to save data", "Warning");
                                        }


                                    }
                                }

                            }
                            break;
                        case 1:
                            if (DGV_input.Rows.Count > 0)
                            {

                                DataTable dt1 = new DataTable();

                                dt1 = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_MOISTURE_ABSORPTION", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                                var result = DialogResult.No;
                                if (dt1.Rows.Count > 0)
                                {
                                    result = MessageBox.Show("Data is exist. Do you want to replace ?", "Warning", MessageBoxButtons.YesNo);

                                    if (result == DialogResult.Yes)
                                    {
                                        if (admin_mode)
                                        {
                                            if (result == DialogResult.Yes)
                                            {
                                                Update_history_Moisture();
                                                TDMK_Code.Delelte_FilteredItem_arr("CQRA_MOISTURE_ABSORPTION", sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                                                save_Moisture_db();
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show("Please, login to save data", "Warning");
                                        }

                                    }
                                }
                                else
                                {
                                    save_Moisture_db();

                                }
                            }
                            break;
                        case 2:
                            save_bHast_data2();
                            break;
                    }

                }
                else
                {
                    MessageBox.Show("Please, fill in ItemCode / LotNo / Operator ID", "Warning");
                }
            }
            else
            {

                MessageBox.Show("Please, choose region ", "Warning");
            }
        }
        public void save_bHast_data()
        {
            if (lstItem.DataSource != null)
            {
                string sel_item = lstItem.SelectedItem.ToString();
                string itemlot = sel_item.Split('(')[0];
                string item = itemlot.Split('-')[0];
                string lot = itemlot.Substring(item.Length + 1);
                if ((item == txtItemCode.Text) && (lot == txtLotNo.Text))
                {
                    DataTable sel_bHast = result_bHast[sel_item];
                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { item, lot });
                start_labl: DataTable src_tar_dt = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_BHAST_LOGFILE", filter_str);
                    if (src_tar_dt.Rows.Count > 0)
                    {
                        if (MessageBox.Show("bHast data of ItemCode:  " + item + "-" + lot + " is existed\r\nDo you want to overwrite?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            TDMK_Code.Delelte_FilteredItem_arr("CQRA_BHAST_LOGFILE", sqlcon, filter_str);
                            goto start_labl;
                        }
                    }
                    else
                    {
                        DataTable tar_dt = src_tar_dt.Clone();
                        int id = TDMK_Code.SQL_MAX("CQRA_BHAST_LOGFILE", "ID", sqlcon);
                        int r_inx = 0;
                        foreach (DataRow dr in sel_bHast.Rows)
                        {
                            List<object> insert_lst = new List<object>();
                            insert_lst.Add(id + r_inx + 1);
                            insert_lst.Add(item);
                            insert_lst.Add(lot);
                            foreach (DataColumn dc in sel_bHast.Columns)
                            {
                                insert_lst.Add(dr[dc]);
                            }
                            tar_dt.Rows.Add(insert_lst.ToArray());
                            r_inx++;
                        }
                        exp_proc.BatchBulkCopy(sqlcon, tar_dt, "CQRA_BHAST_LOGFILE");
                        MessageBox.Show("Insert bHast data of ItemCode: " + item + " - " + lot + " is completed!", "Information");
                    }
                }
                else
                {
                    MessageBox.Show("ItemCode and LotNo of Log file not matched!\r\nPlease, choose corrected log file");
                }
            }
        }
        public void save_bHast_data2()
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
            if (lstItem.DataSource != null && update_mode_bhast == false)
            {
                string save_item = "";
                bool check_name_list = false;
                foreach (string _item in lstItem.Items)
                {
                    //string sel_item = lstItem.SelectedItem.ToString();
                    string itemlot = _item.Split('(')[0];
                    string item = itemlot.Split('-')[0];
                    // string lot = itemlot.Substring(item.Length + 1);
                    string lot = itemlot.Replace(item + "-", string.Empty);
                    lot = exp_proc.Lotno_Formated(lot);

                    if (item == txtItemCode.Text && (lot == txtLotNo.Text))
                    {
                        check_name_list = true;
                        save_item = _item;
                        break;
                    }
                }


                if (check_name_list = true && save_item != "")
                {
                    DataTable sel_bHast = result_bHast[save_item];
                    DataTable src_tar_dt = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_BHAST_LOGFILE", filter_str);

                    // var result = DialogResult.No;
                    if (src_tar_dt.Rows.Count > 0)
                    {
                        bool check_enough = true;
                        for (int i = 1; i <= 8; i++)
                        {
                            if (myCode.checkDBNull(src_tar_dt.Rows[0]["Ch0" + i.ToString()]) == "")
                            {
                                check_enough = false;
                                break;
                            }
                        }
                        if (!check_enough)
                        {
                            List<string> _logfile_lst = src_tar_dt.AsEnumerable().Select(x => x.Field<string>("Remark")).Distinct().ToList();
                            List<string> logfile_lst = new List<string>();
                            foreach (var lst in _logfile_lst)
                            {
                                logfile_lst.AddRange(lst.Split('$').ToList());
                            }
                            if (logfile_lst.IndexOf(logfile_path) == -1)
                            {
                                if (src_tar_dt.Rows.Count == sel_bHast.Rows.Count)
                                {
                                    List<string> empty_col = new List<string>();
                                    foreach (DataColumn dc in src_tar_dt.Columns)
                                    {
                                        if (myCode.checkDBNull(src_tar_dt.Rows[0][dc]) == "")
                                        {
                                            empty_col.Add(dc.ColumnName);
                                        }
                                    }
                                    if (empty_col.Count > 0)
                                    {
                                        List<string> sel_bHast_col = new List<string>();
                                        foreach (DataColumn dc in sel_bHast.Columns)
                                        {
                                            if (dc.ColumnName.ToUpper().Contains("CH"))
                                            {
                                                sel_bHast_col.Add(dc.ColumnName);
                                            }
                                        }
                                        int col_num = Math.Min(empty_col.Count, Math.Min(sel_bHast_col.Count, 8));
                                        int r_inx = 0;
                                        foreach (DataRow dr in src_tar_dt.Rows)
                                        {
                                            for (int i = 0; i < col_num; i++)
                                            {
                                                dr[empty_col[i]] = sel_bHast.Rows[r_inx][sel_bHast_col[i]];
                                            }
                                            dr["Remark"] = dr["Remark"].ToString() + "$" + logfile_path;
                                            r_inx++;
                                        }

                                        TDMK_Code.Delelte_FilteredItem_arr("CQRA_BHAST_LOGFILE", sqlcon, filter_str);
                                        exp_proc.BatchBulkCopy(sqlcon, src_tar_dt, "CQRA_BHAST_LOGFILE");
                                        MessageBox.Show("Save data completed", "Warning");
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Data not enough!", "Warning");
                                }
                            }
                            else
                            {
                                MessageBox.Show("Data of this logfile is inserted. Cannot insert again!", "Warning");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Data is exist. Please update data", "Warning");
                        }

                    }
                    else
                    {
                        DataTable tar_dt = src_tar_dt.Clone();
                        int id = TDMK_Code.SQL_MAX("CQRA_BHAST_LOGFILE", "ID", sqlcon);
                        int r_inx = 0;
                        foreach (DataRow dr in sel_bHast.Rows)
                        {
                            List<object> insert_lst = new List<object>();
                            insert_lst.Add(id + r_inx + 1);
                            insert_lst.Add(txtItemCode.Text);
                            insert_lst.Add(txtLotNo.Text);

                            int count_CH = 0;
                            foreach (DataColumn dc in sel_bHast.Columns)
                            {
                                if (count_CH < 9)
                                {
                                    insert_lst.Add(dr[dc]);
                                    count_CH++;
                                }
                            }
                            tar_dt.Rows.Add(insert_lst.ToArray());
                            tar_dt.Rows[r_inx]["Remark"] = logfile_path;
                            r_inx++;
                        }
                        exp_proc.BatchBulkCopy(sqlcon, tar_dt, "CQRA_BHAST_LOGFILE");
                        MessageBox.Show("Data of this logfile is inserted", "Information");
                    }
                }
                else
                {
                    //MessageBox.Show("ItemCode and LotNo of Log file not matched!\r\nPlease, choose corrected log file");
                    MessageBox.Show("Log file does not constain data for ItemCode: " + txtItemCode.Text + " and LotNo: " + txtLotNo.Text + "\r\nPlease, choose corrected log file");
                }
            }
            else if (update_mode_bhast)
            {
                DataTable dt_before = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_BHAST_LOGFILE", filter_str);
                DataTable src_tbl = TDMK_Code.Datatable_Filter(sqlcon, "BHAST_EDIT_HISTORY", filter_str);

                string[] remark_all = src_tbl.AsEnumerable().Select(x => x.Field<string>("Remark")).Distinct().ToArray();
                foreach (DataRow dr in dt_before.Rows)
                {
                    dr["Remark"] = "Modified: " + (remark_all.Length + 1).ToString() + "_Depart: " + depart + "_TimeUpdate: " + DateTime.Now;
                }
                exp_proc.BatchBulkCopy(sqlcon, dt_before, "BHAST_EDIT_HISTORY");

                TDMK_Code.Delelte_FilteredItem_arr("CQRA_BHAST_LOGFILE", sqlcon, filter_str);
                exp_proc.BatchBulkCopy(sqlcon, (DataTable)DGV_input.DataSource, "CQRA_BHAST_LOGFILE");
                MessageBox.Show("Update data completed", "Warning");
                update_mode_bhast = false;
            }
        }

        public void update_history_BHAST(string ItemCode, string LotNo)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            DataTable src_tar_dt = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_BHAST_LOGFILE", filter_str);
            string[] remark_all = src_tar_dt.AsEnumerable().Select(x => x.Field<string>("Remark")).Distinct().ToArray();

            foreach (DataRow row in src_tar_dt.Rows)
            {
                string remark = row["Remark"].ToString();
                int No_F = remark_all.Length;
                remark = remark + "_NGF" + No_F.ToString();
                row["Remark"] = remark;
            }
        }


        private void Frm_Manual_input_FormClosed(object sender, FormClosedEventArgs e)
        {
            myVar._frmMain.Show();
        }
        public DataTable bHast_LogFile(string filename)
        {
            string CSVFilePathName = filename;
            string[] Lines = File.ReadAllLines(CSVFilePathName);
            string[] Fields;
            List<string> data_col = new List<string>();
            int inx = 0;
            for (int i = 0; i < Lines.Length; i++)
            {
                Lines[i] = Lines[i].Replace("\"", "");
            }
            foreach (string _li in Lines)
            {

                string li = _li.Replace("\"", "");

                if (li.Contains("Total") && li.Contains("Ch01"))
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
                if ((Fields[i].Contains("Total") || Fields[i].Contains("Ch")) && (!Fields[i].Contains("Temperature")) && (!Fields[i].Contains("Humidity")))
                {
                    if (Fields[i].Contains("Total"))
                    {
                        Fields[i] = "Time";
                    }

                    if (!myCode.check_columns_existed(dt, Fields[i]))
                    {
                        dt.Columns.Add(Fields[i], typeof(string));
                    }
                }
            }
            DataRow Row;
            for (int i = inx + 1; i < Lines.GetLength(0); i++)
            {
                Fields = Lines[i].Split(new char[] { ',' });
                Row = dt.NewRow();
                for (int f = 0; f < dt.Columns.Count; f++)
                {
                    if (IsNumeric(Fields[f]))
                    {
                        Row[f] = Fields[f];
                    }
                    else
                    {
                        Row[f] = "<1E+03";
                    }
                }
                dt.Rows.Add(Row);
            }
            return dt;
        }
        public Dictionary<string, DataTable> bHast_Logfile_Process_old(string src_path)
        {
            Dictionary<string, DataTable> result = new Dictionary<string, DataTable>();
            DataTable src_tbl = bHast_LogFile(src_path);
            string file_name = Path.GetFileNameWithoutExtension(src_path);
            string[] ItemCode_arr = file_name.Split('+');
            foreach (var x in ItemCode_arr)
            {
                string chanel_start = (x.Split('(')[1].Replace(")", "")).Split('-')[0];
                string chanel_end = (x.Split('(')[1].Replace(")", "")).Split('-')[1];
                List<string> sel_col = new List<string>();
                sel_col.Add(src_tbl.Columns[0].ColumnName);
                for (int i = Convert.ToInt32(chanel_start); i <= Convert.ToInt32(chanel_end); i++)
                {
                    if (i < 10)
                    {
                        sel_col.Add("Ch0" + i.ToString());

                    }
                    else
                    {
                        sel_col.Add("Ch" + i.ToString());
                    }
                }
                result.Add(x, src_tbl.AsDataView().ToTable(false, sel_col.ToArray()));
            }
            return result;
        }

        public Dictionary<string, DataTable> bHast_Logfile_Process(string src_path)
        {
            Dictionary<string, DataTable> result = new Dictionary<string, DataTable>();
            DataTable src_tbl = bHast_LogFile(src_path);
            string file_name = Path.GetFileNameWithoutExtension(src_path);
            string[] ItemCode_arr = file_name.Split('_')[0].Split('+');
            foreach (var x in ItemCode_arr)
            {
                if (!x.Contains(";"))
                {
                    string chanel_start = (x.Split('(')[1].Replace(")", "")).Split('-')[0];
                    string chanel_end = (x.Split('(')[1].Replace(")", "")).Split('-')[1];
                    List<string> sel_col = new List<string>();
                    sel_col.Add(src_tbl.Columns[0].ColumnName);
                    for (int i = Convert.ToInt32(chanel_start); i <= Convert.ToInt32(chanel_end); i++)
                    {
                        if (i < 10)
                        {
                            sel_col.Add("Ch0" + i.ToString());
                        }
                        else
                        {
                            sel_col.Add("Ch" + i.ToString());
                        }
                    }
                    result.Add(x, src_tbl.AsDataView().ToTable(false, sel_col.ToArray()));
                }
                else
                {
                    string[] arr_chanel = (x.Split('(')[1].Replace(")", "")).Split(';');
                    List<string> sel_col = new List<string>();
                    sel_col.Add(src_tbl.Columns[0].ColumnName);
                    foreach (string c in arr_chanel)
                    {
                        if (c.Contains("-"))
                        {
                            string chanel_start = c.Split('-')[0];
                            string chanel_end = c.Split('-')[1];


                            for (int i = Convert.ToInt32(chanel_start); i <= Convert.ToInt32(chanel_end); i++)
                            {
                                if (i < 10)
                                {
                                    sel_col.Add("Ch0" + i.ToString());
                                }
                                else
                                {
                                    sel_col.Add("Ch" + i.ToString());
                                }
                            }
                        }
                        else
                        {
                            int i = Convert.ToInt32(c);
                            if (i < 10)
                            {
                                sel_col.Add("Ch0" + i.ToString());
                            }
                            else
                            {
                                sel_col.Add("Ch" + i.ToString());
                            }
                        }
                    }
                    result.Add(x, src_tbl.AsDataView().ToTable(false, sel_col.ToArray()));
                }
            }
            return result;
        }
        public DataTable bHast_Logfile_Process_Update(string src_path, string itemcode, List<string> sel_col_update)
        {
            Dictionary<string, DataTable> result = new Dictionary<string, DataTable>();
            DataTable dt = (DataTable)DGV_input.DataSource;
            DataTable src_tbl = bHast_LogFile(src_path);
            string file_name = Path.GetFileNameWithoutExtension(src_path);
            string[] ItemCode_arr = file_name.Split('+');
            foreach (var x in ItemCode_arr)
            {
                if (x.Contains(itemcode))
                {
                    //string chanel_start = (x.Split('(')[1].Replace(")", "")).Split('-')[0];
                    //string chanel_end = (x.Split('(')[1].Replace(")", "")).Split('-')[1];
                    List<string> sel_col = new List<string> { };
                    sel_col.Add(src_tbl.Columns[0].ColumnName);
                    foreach (string item in sel_col_update)
                    {
                        sel_col.Add(item);
                    }

                    DataTable tbl_update = src_tbl.AsDataView().ToTable(false, sel_col.ToArray());
                    foreach (DataColumn dc in dt.Columns)
                    {
                        if (sel_col_update.IndexOf(dc.ColumnName) != -1)
                        {
                            int indx = 0;
                            foreach (DataRow dr in tbl_update.Rows)
                            {
                                if (dt.Rows.Count > indx)
                                {
                                    dt.Rows[indx][dc.ColumnName] = dr[dc.ColumnName];
                                    dt.Rows[indx]["Remark"] = logfile_path;

                                }
                                else
                                {
                                    DataRow drow = dt.NewRow();
                                    drow[dc.ColumnName] = dr[dc.ColumnName];
                                    drow["Remark"] = logfile_path;
                                    dt.Rows.Add(drow);
                                }
                                indx++;

                            }
                        }
                    }
                    update_mode_bhast = true;

                    // dt_before.AsDataView().ToTable(false, col_before.ToArray());
                    break;
                }
            }
            return dt;
        }
        public static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum;
        }
        private void lstItem_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstItem.SelectedIndex > -1)
            {
                string sel_item = lstItem.SelectedItem.ToString();
                string itemlot = sel_item.Split('(')[0];
                string item = itemlot.Split('-')[0];
                // string lot = itemlot.Substring(item.Length + 1);
                string lot = itemlot.Replace(item + "-", string.Empty);
                DGV_input.DataSource = result_bHast[sel_item];
                myVar._frmbHastGraph = new bHast_Graph(result_bHast[sel_item], item, lot);
                myVar._frmbHastGraph.Show();
            }
        }
        public void Export_Bhast_Logfile_excel(string itemcode, string lotno, SqlConnection sqlcon)
        {
            string[] item = { "ItemCode", "LotNo" };
            string[] item_val = { itemcode, lotno };
            string[] list_ch = { "Ch01", "Ch02", "Ch03", "Ch04", "Ch05", "Ch06", "Ch07", "Ch08" };
            int start_col = 1;
            int start_row = 1;
            string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
            if (format_file != "")
            {
                myExcel.Workbook curr_wrkbook = exp_proc.create_export_wrk(format_file, "CQRA - bHast", txtItemCode.Text, txtLotNo.Text);
                myExcel.Worksheet ws = curr_wrkbook.Sheets[1];
                // tìm vị trí Ch01, vị trí khác offset
                myExcel.Application xlsApp = TDMK_Code.StartExcel();
                xlsApp.DisplayAlerts = false;
                for (int i = 15; i < 35; i++)
                {
                    for (int j = 4; j < 20; j++)
                    {
                        if (ws.Cells[i, j].Value != null)
                        {
                            if (ws.Cells[i, j].Value.ToString() == "Ch01")
                            {
                                start_row = i;
                                start_col = j; break;
                            }
                        }
                    }
                }
                foreach (string chanel in list_ch)
                {
                    try
                    {
                        AutoCompleteStringCollection list = TDMK_Code.Load_Item_Filter_str(sqlcon, "CQRA_BHAST_LOGFILE", chanel, TDMK_Code.filter_str(item, item_val));

                        string number = chanel.Replace("Ch0", "");
                        int col_offset = Int32.Parse(number);

                        for (int y = 1; y <= list.Count; y++)
                        {
                            ws.Cells[start_row + y, start_col + col_offset - 1].Value = list[y - 1];
                        }
                    }
                    catch
                    {

                    }

                }
                xlsApp.DisplayAlerts = true;
                MessageBox.Show("Export data completed!", "Information");
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Format of " + txtItemCode.Text + " not found", "Warning");
            }

        }
        public void Export_Bhast_Logfile(string itemcode, string lotno, SqlConnection sqlcon)
        {
            string[] item = { "ItemCode", "LotNo" };
            string[] item_val = { itemcode, lotno };
            string[] list_ch = {"Time", "Ch01", "Ch02", "Ch03", "Ch04", "Ch05", "Ch06", "Ch07", "Ch08" };
            int start_col = 1;
            int start_row = 1;
            string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
            if(format_file!="")
            {
                string export_path = Path.Combine(Application.StartupPath, "Export", cb_sheet.Text);
                if (!Directory.Exists(export_path))
                {
                    Directory.CreateDirectory(export_path);
                }
                ExcelPackage format_pack = Excel_lib.open_excel(format_file);
                ExcelWorkbook curr_wrkbook = format_pack.Workbook;
                ExcelWorksheet ws = null;
                string _process = "CQRA - bHast".Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
                foreach (ExcelWorksheet tg_sht in curr_wrkbook.Worksheets)
                {
                    string cur_sht_name = tg_sht.Name.ToUpper().Replace("-", "").Replace(" ", "");
                    if (cur_sht_name == _process)
                    {
                        ws = tg_sht;
                        break;
                    }
                }
                if (ws != null)
                {
                    DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_BHAST_LOGFILE", TDMK_Code.filter_str(item, item_val));
                    for (int i = 15; i < 35; i++)
                    {
                        for (int j = 4; j < 20; j++)
                        {
                            if (ws.Cells[i, j].Value != null)
                            {
                                if (ws.Cells[i, j].Value.ToString().ToUpper() == "CH01")
                                {
                                    start_row = i;
                                    start_col = j; break;
                                }
                            }
                        }
                    }
                    ExcelRangeBase channel_rgn = ws.Cells[start_row, start_col];
                    int inx = 0;
                    List<string> channel_lst = new List<string>() { "Time"};
                    while (myCode.checkDBNull( channel_rgn.Offset(0,inx).Value).ToUpper().Contains("CH"))
                    {
                        channel_lst.Add(channel_rgn.Offset(0, inx).Value.ToString());
                        inx++;
                    }
                    foreach (string chanel in list_ch)
                    {
                        try
                        {
                            // AutoCompleteStringCollection list = TDMK_Code.Load_Item_Filter_str(sqlcon, "CQRA_BHAST_LOGFILE", chanel, TDMK_Code.filter_str(item, item_val));
                            if(channel_lst.IndexOf(chanel)!=-1)
                            {
                                List<string> list = dt.AsEnumerable().Select(x => x.Field<string>(chanel)).ToList();
                                string number = chanel.Replace("Ch0", "");
                                int col_offset = 0;
                                if (TDMK_Code.IsNumeric(number))
                                {
                                    col_offset = Int32.Parse(number);
                                }
                                for (int y = 1; y <= list.Count; y++)
                                {

                                    if (TDMK_Code.IsNumeric(list[y - 1]))
                                    {
                                        ws.Cells[start_row + y, start_col + col_offset - 1].Value = Convert.ToDouble(list[y - 1]);
                                        if (chanel != "Time")
                                        {
                                            ws.Cells[start_row + y, start_col + col_offset - 1].Style.Numberformat.Format = "0.00E+00";
                                        }
                                    }
                                    else
                                    {
                                        ws.Cells[start_row + y, start_col + col_offset - 1].Value = list[y - 1];
                                    }
                                }
                            }

                        }
                        catch
                        {

                        }
                    }
                    string export_file = Path.Combine(export_path, cb_sheet.Text + "_" + itemcode + "-" + lotno + "_" + txt_Operator.Text + "_" + DateTime.Now.ToString("ddMMMyyyy HHmmss") + ".xlsx");
                    ExcelPackage report_pack = new ExcelPackage(export_file);
                    ExcelWorkbook report_wrkbk = report_pack.Workbook;
                    report_wrkbk.Worksheets.Add(ws.Name, ws);
                    report_pack.Save();
                    report_pack = null;
                    report_wrkbk = null;
                    try
                    {
                        ProcessStartInfo pi = new ProcessStartInfo(export_file);
                        Process.Start(pi);
                    }
                    catch { };
                    MessageBox.Show(new Form { TopMost = true }, "Hoàn thành xuất dữ liệu", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Không tìm thấy format", "Thông báo");
            }
        }
        public void Export_ACF_Roughness(string itemcode, string lotno)
        {
            SqlConnection sqlcon_IPQC = exp_proc.initial_data("IPQC_Data", true);

            DataTable dt1 = new DataTable();
            dt1 = TDMK_Code.Datatable_Filter(sqlcon_IPQC, "Roughness", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcode, lotno }));
            if (dt1.Rows.Count > 0)
            {
                List<List<string>> lst_rghness = lst_roughness(dt1);
                string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                if (format_file != "")
                {
                    myExcel.Workbook curr_wrkbook = exp_proc.create_export_wrk(format_file, "ACF", txtItemCode.Text, txtLotNo.Text);
                    myExcel.Worksheet curr_wrksheet_1 = curr_wrkbook.Sheets["ACF"];
                    myExcel.Range curr_rgn_1 = curr_wrksheet_1.Range["A1"];

                    List<int> vitri_sa = new List<int> { };
                    List<int> vitri_sq = new List<int> { };
                    List<int> vitri_sdr = new List<int> { };
                    string[] arr_info = ACF_pad_location(curr_wrkbook).Split('+');
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
                    MessageBox.Show("Export to Checksheet completed!", "Warning");
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Format of " + txtItemCode.Text + " not found", "Warning");
                }
            }
            else
            {
                MessageBox.Show("No data", "Warning");
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
        public string ACF_pad_location(myExcel.Workbook curr_wrkbook)
        {
            // List<int> lst_info = new List<int> { };
            string get_info = "";
            myExcel.Worksheet curr_wrksheet_1 = curr_wrkbook.Sheets["ACF"];
            myExcel.Range curr_rgn_1 = curr_wrksheet_1.Range["A1"];

            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 30; i < 100; i++)
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

        private void cb_lotno_roughness_SelectedIndexChanged(object sender, EventArgs e)
        {
            // txtLotNo.Text = cb_lotno_roughness.SelectedItem.ToString();
        }

        private void rbQA_CheckedChanged(object sender, EventArgs e)
        {
            if (rbQA.Checked)
            {
                depart = rbQA.Text;
            }
        }

        private void rbDE_CheckedChanged(object sender, EventArgs e)
        {
            if (rbDE.Checked)
            {
                depart = rbDE.Text;
            }
        }

        private void rbNPI_CheckedChanged(object sender, EventArgs e)
        {
            if (rbNPI.Checked)
            {
                depart = rbNPI.Text;
            }
        }

        private void rbPro_CheckedChanged(object sender, EventArgs e)
        {
            if (rbPro.Checked)
            {
                depart = rbPro.Text;
            }
        }

        private void rbPE_CheckedChanged(object sender, EventArgs e)
        {
            if (rbPE.Checked)
            {
                depart = rbPE.Text;
            }
        }

        private void Cancel_Click(object sender, EventArgs e)
        {

        }

        private void btnLogin_Click(object sender, EventArgs e)
        {

        }

        private void cbProcess_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbProcess_Edit.SelectedIndex != -1 && txtItemCode_Edit.Text != "" && txtLotNo_Edit.Text != "")
            {
                if (txtLotNo.Text.Length == 5)
                {


                    DGV_DataEdit.DataSource = null;
                    if (cbProcess_Edit.SelectedIndex != 2)
                    {
                        Search_Editted_Data();
                    }
                    else
                    {
                        DataTable bHast_dt = TDMK_Code.Datatable_Filter(sqlcon, "BHAST_EDIT_HISTORY", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                        DGV_DataEdit.DataSource = bHast_dt;
                    }
                }
                else
                {
                    MessageBox.Show("Please, Fill in the correct number of characters of lotno", "Warning");
                }

            }
        }
        public void Search_Editted_Data()
        {
            DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { txtItemCode_Edit.Text, txtLotNo_Edit.Text, cbProcess_Edit.Text }));
            if (dt.Rows.Count > 0)
            {
                DataTable dt_data = new DataTable();
                DataTable dt_img = new DataTable();
                Split_Data_Image(dt, ref dt_data, ref dt_img);
                DGV_DataEdit.DataSource = dt_data;

            }
            else
            {
                DGV_DataEdit.DataSource = null;

                MessageBox.Show("No data", "Warning");
            }
        }
        public void Split_Data_Image(DataTable src_dt, ref DataTable data_dt, ref DataTable img_dt)
        {
            List<string> col_name = new List<string>();
            foreach (DataColumn dc in src_dt.Columns)
            {
                if (!dc.ColumnName.Contains("Before") && (!dc.ColumnName.Contains("After")))
                {
                    col_name.Add(dc.ColumnName);
                }
            }
            // DataTable dt_img = src_dt.AsEnumerable().Where(x => x.Field<string>("Region") == "Image_data").CopyToDataTable();
            DataTable dt_data = src_dt.AsEnumerable().Where(x => x.Field<string>("Region") != "Image_data").CopyToDataTable();
            // img_dt = dt_img.AsDataView().ToTable(false, new string[] { "Time_Update", "PCS_No", "Before_Data", "After_Data" });
            DataTable disp_dt = dt_data.AsDataView().ToTable(false, col_name.ToArray());
            disp_dt.Columns.Add("Before_Data", typeof(string));
            disp_dt.Columns.Add("After_Data", typeof(string));
            int r_inx = 0;
            foreach (DataRow dr in dt_data.Rows)
            {
                byte[] bef_val = (byte[])dr["Before_Data"];
                byte[] aft_val = (byte[])dr["After_Data"];
                disp_dt.Rows[r_inx]["Before_Data"] = Encoding.UTF8.GetString(bef_val, 0, bef_val.Length);
                disp_dt.Rows[r_inx]["After_Data"] = Encoding.UTF8.GetString(aft_val, 0, aft_val.Length);
                r_inx++;
            }
            data_dt = disp_dt.AsDataView().ToTable(false, new string[] { "Zone", "PCS_No", "Before_Data", "After_Data", "Time_Update", "Depart" });
        }

        private void DGV_input_MouseDown(object sender, MouseEventArgs e)
        {
            if (cb_sheet.SelectedIndex == 2)
            {
                if (e.Button == MouseButtons.Right)
                {

                    if (txtItemCode.Text != "" && txtLotNo.Text != "")
                    {
                        DataTable logfile_before = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_BHAST_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                        if (logfile_before.Rows.Count > 0)
                        {
                            if (admin_mode)
                            {
                                if (DGV_input.SelectedCells.Count > 0)
                                {

                                    sel_pcs_lst = new List<string> { };
                                    List<string> ignored_col = new List<string>() { "ID", "ItemCode", "LotNo", "Time", "Remark" };
                                    foreach (DataGridViewCell cells in DGV_input.SelectedCells)
                                    {

                                        int c_inx = cells.ColumnIndex;
                                        string col_name = DGV_input.Columns[c_inx].Name;
                                        //string _pcs_no = new string(col_name.Where(char.IsDigit).ToArray());
                                        if (ignored_col.IndexOf(col_name) == -1)
                                        {

                                            if (sel_pcs_lst.IndexOf(col_name) == -1)
                                            {
                                                if (myCode.checkDBNull(DGV_input.Rows[0].Cells[col_name].Value) != "")
                                                {
                                                    sel_pcs_lst.Add(col_name);
                                                }
                                            }
                                            //  int pcs_no = Convert.ToInt32(_pcs_no);
                                        }
                                    }
                                    if (sel_pcs_lst.Count > 0)
                                    {

                                        OpenFileDialog f_open = new OpenFileDialog();
                                        f_open.Filter = "Excel(*.csv)|*.csv";
                                        f_open.InitialDirectory = Application.StartupPath;
                                        if (f_open.ShowDialog() == DialogResult.OK)
                                        {
                                            if (f_open.FileName != "")
                                            {
                                                logfile_path = f_open.FileName;
                                                string logfile_name = Path.GetFileNameWithoutExtension(f_open.FileName);
                                                string[] ItemCode_arr = logfile_name.Split('+');
                                                bool check_name = false;
                                                foreach (var x in ItemCode_arr)
                                                {

                                                    string itemlot = x.Split('(')[0];
                                                    string item = itemlot.Split('-')[0];
                                                    // string lot = itemlot.Substring(item.Length + 1);
                                                    string lot = itemlot.Replace(item + "-", string.Empty);
                                                    lot = exp_proc.Lotno_Formated(lot);

                                                    if (item == txtItemCode.Text && (lot == txtLotNo.Text))
                                                    {
                                                        check_name = true;
                                                        break;
                                                    }
                                                }
                                                if (check_name)
                                                {
                                                    DataTable src_dt = bHast_Logfile_Process_Update(f_open.FileName, txtItemCode.Text, sel_pcs_lst);
                                                    DGV_input.DataSource = src_dt;
                                                    myVar._frmbHastGraph = new bHast_Graph(src_dt, txtItemCode.Text, txtLotNo.Text);
                                                    myVar._frmbHastGraph.Show();
                                                }
                                                else
                                                {
                                                    MessageBox.Show("Logfile does not constain data for ItemCode: " + txtItemCode.Text + "_LotNo: " + txtLotNo.Text, "Warning", MessageBoxButtons.YesNo);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Please, login to edit data", "Warning");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Data is not exist to update. Must save data before", "Warning");
                        }
                    }

                }
            }

        }

        private void btnLogin_Click_1(object sender, EventArgs e)
        {
            if (btnLogin.Text == "Login")
            {
                bool login_en = false;
                if ((txtUsername.Text == "Admin") && (txtPassword.Text == "TDMK"))
                {
                    login_en = true;
                }
                else
                {
                    string sqllogin_constr = TDMK_Code.data_connection(myVar.server_name, "OK2SHIP_Items", myVar.server_acc, myVar.server_pass).ConnectionString;
                    SqlConnection sql_login = new SqlConnection(sqllogin_constr);
                    DataTable info = TDMK_Code.Datatable_Filter(sql_login, "USerInfo", TDMK_Code.filter_str(new string[] { "User_Name", "User_Password", "Department" }, new string[] { txtUsername.Text, txtPassword.Text, depart }));
                    if (info.Rows.Count == 0)
                    {
                        login_en = false;
                    }
                    else
                    {
                        login_en = true;
                    }
                }
                if (login_en)
                {
                    tabLogin.Text = "Admin_mode";
                    //tabLogin.BackColor = Color.GreenYellow;
                    btnLogin.Text = "Logout";
                    admin_mode = true;
                }
                else
                {
                    MessageBox.Show("Sai mật khẩu!");
                }
            }
            else
            {
                admin_mode = false;
                btnLogin.Text = "Login";
                tabLogin.Text = "Login";
                //tabLogin.BackColor = this.BackColor;
            }
        }

        private void txtItemCode_Edit_TextChanged(object sender, EventArgs e)
        {
            DGV_DataEdit.DataSource = null;

        }

        private void txtLotNo_Edit_TextChanged(object sender, EventArgs e)
        {
            DGV_DataEdit.DataSource = null;
        }

        private void txtItemCode_TextChanged(object sender, EventArgs e)
        {
            DGV_input.DataSource = null;
            DGV_input.ReadOnly = true;
            txtLotNo.Text = "";
            cb_sheet.SelectedIndex = -1;
            //txtqty.Text = "";

        }

        private void txtLotNo_TextChanged(object sender, EventArgs e)
        {
            DGV_input.DataSource = null;
            DGV_input.ReadOnly = true;

        }

        private void DGV_input_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (cb_sheet.SelectedItem.ToString() != "CQRA_BHAST")
                {
                    cmsPaste.Show(DGV_input, e.Location);
                }

            }
        }

        private void patToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void selectedRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //if(txtItemCode.Text != "" && txtLotNo.Text != "")
            //{
            //   int indx_select = cb_sheet.SelectedIndex;
            //    DataTable dt = new DataTable();
            //    switch (indx_select)
            //    {
            //        case 0:
            //            dt = TDMK_Code.Datatable_Filter(sqlcon, "ACF_Flatness", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
            //            break;

            //        case 1:
            //            dt = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_MOISTURE_ABSORPTION", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
            //            break;

            //        case 2:
            //            dt = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_BHAST_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
            //            break; 
            //    }

            //    if (dt.Rows.Count > 0)
            //    {
            //        if (admin_mode)
            //        {
            PasteClipboardValue(false, DGV_input);
            //        }
            //        else
            //        {
            //            MessageBox.Show("Please, login to update data", "Warning");
            //        }
            //    }
            //    else
            //    {
            //        PasteClipboardValue(false, DGV_input);
            //    }


            //}


        }

        private void selectedColumnToolStripMenuItem_Click(object sender, EventArgs e)
        {

            PasteClipboardValue(true, DGV_input);

        }

        private void DGV_input_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (admin_mode)
            {
                DGV_input.ReadOnly = false;
            }
        }

        private void DGV_input_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (cb_sheet.SelectedIndex == 0)
            {
                int c_inx = e.ColumnIndex;
                int r_inx = e.RowIndex;
                List<double> data_lst = new List<double>();
                bool en_calc = true;
                for (int i = 0; i < 10; i++)
                {
                    string col_name = "Point" + (i + 1).ToString();
                    string sel_cell_val = myCode.checkDBNull(DGV_input.Rows[r_inx].Cells[col_name].Value);
                    if (sel_cell_val != "")
                    {
                        data_lst.Add(Convert.ToDouble(sel_cell_val));
                    }
                    else
                    {
                        en_calc = false;
                        break;
                    }
                }
                if (en_calc)
                {
                    double min_val = data_lst.Min();
                    double max_val = data_lst.Max();
                    double flatness_val = max_val - min_val;
                    DGV_input.Rows[r_inx].Cells["Flatness"].Value = flatness_val;
                    if (flatness_val < 15)
                    {
                        DGV_input.Rows[r_inx].Cells["Judgement"].Value = "Pass";
                        DGV_input.Rows[r_inx].Cells["Judgement"].Style.BackColor = Color.White;
                    }
                    else
                    {
                        DGV_input.Rows[r_inx].Cells["Judgement"].Value = "Fail";
                        DGV_input.Rows[r_inx].Cells["Judgement"].Style.BackColor = Color.Red;
                    }
                }
            }
        }

        private void DGV_input_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && cb_sheet.SelectedItem.ToString() != "CQRA_BHAST")
            {
                int indx_select = cb_sheet.SelectedIndex;
                DataTable dt = new DataTable();
                switch (indx_select)
                {
                    case 0:
                        dt = TDMK_Code.Datatable_Filter(sqlcon, "ACF_Flatness", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                        break;

                    case 1:
                        dt = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_MOISTURE_ABSORPTION", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                        break;

                        //case 2:
                        //    dt = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_BHAST_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                        //    break;


                }

                if (dt.Rows.Count > 0)
                {
                    if (admin_mode)
                    {
                        DGV_input.ReadOnly = false;
                    }
                    else
                    {
                        MessageBox.Show("Please, login to update data", "Warning");
                    }
                }
                else
                {
                    DGV_input.ReadOnly = false;
                }

            }

        }

        private void cmsPaste_Opening(object sender, CancelEventArgs e)
        {

        }

        private void txtLotNo_Validated(object sender, EventArgs e)
        {
            txtLotNo.Text = exp_proc.Lotno_Formated(txtLotNo.Text);
        }

        private void txtLotNo_Edit_Validated(object sender, EventArgs e)
        {
            txtLotNo_Edit.Text = exp_proc.Lotno_Formated(txtLotNo_Edit.Text);
        }

        private void tableLayoutPanel10_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void DGV_input_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            int r_inx = e.RowIndex;
            int c_inx = e.ColumnIndex;
            if (r_inx != -1 && c_inx != -1 && cb_sheet.SelectedItem.ToString() == "Moisture Absorption" && DGV_input.Rows.Count > 0 && (DGV_input.Columns[c_inx].Name.Contains("after") || DGV_input.Columns[c_inx].Name.Contains("before")))
            {
                string val_before = DGV_input.Rows[r_inx].Cells["Weight_before"].Value.ToString();
                string val_after = DGV_input.Rows[r_inx].Cells["Weight_after"].Value.ToString();


                if (myCode.IsNumeric(val_before) && myCode.IsNumeric(val_after))
                {
                    double percent = (double.Parse(val_after) - double.Parse(val_before)) / double.Parse(val_before);
                    DGV_input.Rows[r_inx].Cells["%"].Value = Math.Round(percent * 100, 2).ToString() + "%";
                    DGV_input.Rows[r_inx].Cells["Judgement"].Value = "OK";
                    //if (percent < 0.008)
                    //{
                    //    DGV_input.Rows[r_inx].Cells["Judgement"].Value = "OK";
                    //}
                    //else
                    //{
                    //    DGV_input.Rows[r_inx].Cells["Judgement"].Value = "NG";
                    //}
                }
            }
        }
    }
}
