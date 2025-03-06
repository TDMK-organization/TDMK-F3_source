using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using VHX;
using myExcel = Microsoft.Office.Interop.Excel;
using OK2SHIP;
using TDMK_SQL;
using TDMK_SEEV_DLL;
using System.Runtime.InteropServices;
using System.Net.NetworkInformation;

namespace OK2SHIP
{
    public partial class Check_Impedance_logfile : Form
    {
        public Check_Impedance_logfile()
        {
            InitializeComponent();
        }
        myVar exp_proc = new myVar();
        Impedance_data proc_data = new Impedance_data();

        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SEI_Lib myCode = new SEI_Lib();
        Export_Image_Class img_proc = new Export_Image_Class();
        Impedance_main impedance_code = new Impedance_main();

        public string strcon = "";
        public SqlConnection sqlcon = null;




        
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

        public void Logfile_Data_Process()
        {
            FolderBrowserDialog f_open = new FolderBrowserDialog();
            f_open.SelectedPath = Application.StartupPath;
            if (f_open.ShowDialog() == DialogResult.OK)
            {
                DGV_impedance.DataSource = Impedance_Data(f_open.SelectedPath);
                DGV_tracewidth.DataSource = Tracewidth_Data(f_open.SelectedPath);
                MessageBox.Show("Hoàn thành", "Thông báo");
            }
        }
        public DataTable dt_check_logfile(string ItemCode, string LotNo, List<Double[]> lst_regions)
        {

            DataTable dt = new DataTable();
            string[] arr_spec = { "Region","Max_target","Min_targer", "Max_Data", "Avg", "Min_Data", "std Dev", "Cp", "Cpkl", "Cpku", "Cpk" };


            for (int i = 0; i < 11; i++)
            {
                dt.Columns.Add(arr_spec[i]);
            }

            int j = 0;
            foreach (var item in lst_regions)
            {
                Double nom = Double.Parse(txt_nom_imp.Text);
                Double tol = Double.Parse(txt_tol_imp.Text);
                Double max_target = nom + (tol * nom)/100;
                Double min_target = nom - (tol * nom)/100;


                Double STDEV = CalculateStandardDeviation(item);
                Double Mean = item.Average();
                Double cp = (max_target - min_target) / (6 * STDEV);
                Double cpkl = (Mean - min_target) / (3 * STDEV);
                Double cpku = (max_target - Mean) / (3 * STDEV);
                DataRow dr = dt.NewRow();
                dr[0] = "Impedance" + j.ToString();
                dr[1] = max_target;
                dr[2] = min_target;
                dr[3] = item.Max();
                dr[4] = nom;
                dr[5] = item.Min();
                dr[6] = STDEV;
                dr[7] = cp;
                dr[8] = cpkl;
                dr[9] = cpku;
                dr[10] = new double[] { cpkl, cpku }.Min();
                dt.Rows.Add(dr);
                j++;
            }
            return dt;
        }
        public DataTable Impedance_Data(string src_path)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
            DataTable Imp_data_dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_VAL", filter_str).Clone();
            if (src_path != "")
            {
                Dictionary<string, SortedDictionary<int, myVar.Impedance_data>> Impedance_result = new Dictionary<string, SortedDictionary<int, myVar.Impedance_data>>();

                impedance_code.Get_Impedance_logfile_Multi2(src_path, ref Impedance_result);
                string sel_graph = Impedance_result.Keys.ToArray()[0];
                SortedDictionary<int, myVar.Impedance_data> graph_data = Impedance_result[sel_graph];

                int r_inx = 0;
                foreach (var item in Impedance_result)
                {
                    var cur_val = item.Value;
                    int pcs_no = 0;
                    foreach (var t in cur_val)
                    {
                        Imp_data_dt.Rows.Add(r_inx + 1, txtItemCode.Text, txtLotNo.Text, t.Key, item.Key, t.Value.data_val, "");
                        pcs_no++;
                        r_inx++;
                    }
                }

            }
            return Imp_data_dt;
        }
        public DataTable Tracewidth_Data(string src_path)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
            DataTable VHX_log_tbl = TDMK_Code.Datatable_Filter(sqlcon, "DATA_LOG_FILE", filter_str).Clone();
            if (src_path != "")
            {
                
                SortedDictionary<int, Dictionary<int, string>> logfile_result_VHX = new SortedDictionary<int, Dictionary<int, string>>();
                impedance_code.Get_logfile_Multi(src_path, ref logfile_result_VHX);

                int log_inx = 0;
                foreach (var log in logfile_result_VHX)
                {
                    foreach (var log_val in log.Value)
                    {
                        VHX_log_tbl.Rows.Add((log_inx + 1), txtItemCode.Text, txtLotNo.Text, "Sample" + log.Key.ToString(), log_val.Key.ToString(), log_val.Value, DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), txtOperator.Text, "IMPEDANCE");
                        log_inx++;
                    }
                }

            }
            return VHX_log_tbl;
        }

        public void load_spec_from_excel(string ItemCode, string LotNo, DataGridView dgv)
        {
            string format_file = exp_proc.find_format(myVar.data_loc, ItemCode);
            myExcel.Workbook curr_wrkbook = exp_proc.create_export_wrk(format_file, "Impedance", ItemCode, LotNo);
            myExcel.Worksheet curr_wrksheet = curr_wrkbook.Sheets[1];
            myExcel.Range curr_rgn = curr_wrksheet.Range["B0"];

            DataTable dt1 = new DataTable();
            dt1.Columns.Add("/");
            dt1.Columns.Add("NOM");
            dt1.Columns.Add("TOL");
            dt1.Columns.Add("MIN");
            dt1.Columns.Add("MEAN");
            dt1.Columns.Add("MAX");
            dt1.Columns.Add("RANGE");

            DataRow dr = dt1.NewRow();
            dr[0] = "Target";

            dt1.Rows.Add(dr);
            dgv.DataSource = dt1;




            for (int i = 0; i < 100; i++)
            {
                if (curr_rgn.Offset[i, 1].Value != null)
                {
                    if (curr_rgn.Offset[i, 1].Value.ToString().Contains("NOM"))
                    {
                        curr_rgn = curr_rgn.Offset[i, 0];
                        break;

                    }
                }
            }
            if (dgv.Rows.Count > 0)
            {
                dgv.Rows[0].Cells[1].Value = curr_rgn.Value.ToString();
                if (dgv.Rows[0].Cells[2].Value != null)
                {

                    Double nom = Double.Parse(dgv.Rows[0].Cells[1].Value.ToString());
                    Double tol = Double.Parse(dgv.Rows[0].Cells[2].Value.ToString());
                    dgv.Rows[0].Cells["MIN"].Value = (nom - (tol * nom)).ToString();
                    dgv.Rows[0].Cells["MEAN"].Value = nom.ToString();
                    dgv.Rows[0].Cells["MAX"].Value = (nom + (tol * nom)).ToString();
                }
            }

        }
        public List<String> info_from_excel(myExcel.Workbook curr_wrkbook)
        {
            List<String> lst_val = new List<string> { };
            myExcel.Worksheet curr_wrksheet = curr_wrkbook.Sheets[1];
            myExcel.Range curr_rgn = curr_wrksheet.Range["A1"];

            for (int i = 0; i < 100; i++)
            {
                if (curr_rgn.Offset[i, 1].Value != null)
                {
                    if (curr_rgn.Offset[i, 1].Value.ToString().Contains("NOM"))
                    {
                        curr_rgn = curr_rgn.Offset[i, 0];
                        break;

                    }
                }
            }
            lst_val.Add(curr_rgn.Value.ToString());


            if (curr_rgn.Offset[1, 0].Value != null)
            {
                string value = curr_rgn.Offset[1, 0].Value.ToString();
                if (value.Contains("%"))
                {
                    lst_val.Add(value);
                }
            }

            return lst_val;


        }

        private void Check_Impedance_logfile_Load(object sender, EventArgs e)
        {
            sqlcon = exp_proc.initial_data(myVar.sel_DB, true);
        }

        private void DGV_spec_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void txtlogfille_TextChanged(object sender, EventArgs e)
        {

        }



        private void txtLogfile_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Logfile_Data_Process();
        }
        public void reset_DGV_color(DataGridView DGV)
        {
            for (int i = 0; i < DGV.Rows.Count; i++)
            {
                for (int j = 0; j < DGV.Columns.Count; j++)
                {

                    DGV.Rows[i].Cells[j].Style.BackColor = Color.White;

                }
            }
        }

      

        private void btn_get_infor_excel_Click(object sender, EventArgs e)
        {
            string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
            myExcel.Workbook curr_wrkbook = exp_proc.create_export_wrk(format_file, "Impedance", txtItemCode.Text, txtLotNo.Text);
            List<string> lst_info = info_from_excel(curr_wrkbook);
            if (lst_info.Count() == 1)
            {
                txt_nom_imp.Text = lst_info[0];
            }
            else
            {
                txt_nom_imp.Text = lst_info[0];
                txt_tol_imp.Text = lst_info[1];
            }
            btn_get_infor_excel.Text = "SPEC";


        }

        private void btn_check_Click(object sender, EventArgs e)
        {

            reset_DGV_color(DGV_impedance);
            if (txt_nom_imp.Text != "" && txt_tol_imp.Text != "")
            {

                List<Double[]> lst_regions = new List<Double[]> { };
                List<Double> arr_Impedance1 = new List<Double> { };

                List<Double> arr_Impedance2 = new List<Double> { };

                List<Double> arr_Impedance3 = new List<Double> { };

                List<Double> arr_Impedance4 = new List<Double> { };

                if (DGV_impedance.Rows.Count > 0)
                {
                    for (int i = 0; i < DGV_impedance.Rows.Count; i++)
                    {
                        if (DGV_impedance.Rows[i].Cells["Region"].Value.ToString() == "1")
                        {
                            arr_Impedance1.Add(Double.Parse(DGV_impedance.Rows[i].Cells["Data"].Value.ToString()));
                        }
                        if (DGV_impedance.Rows[i].Cells["Region"].Value.ToString() == "2")
                        {
                            arr_Impedance2.Add(Double.Parse(DGV_impedance.Rows[i].Cells["Data"].Value.ToString()));
                        }
                        if (DGV_impedance.Rows[i].Cells["Region"].Value.ToString() == "3")
                        {
                            arr_Impedance3.Add(Double.Parse(DGV_impedance.Rows[i].Cells["Data"].Value.ToString()));
                        }
                        if (DGV_impedance.Rows[i].Cells["Region"].Value.ToString() == "4")
                        {
                            arr_Impedance4.Add(Double.Parse(DGV_impedance.Rows[i].Cells["Data"].Value.ToString()));
                        }
                    }
                }

                if (arr_Impedance1.Count > 0)
                {
                    lst_regions.Add(arr_Impedance1.ToArray());
                }
                if (arr_Impedance2.Count > 0)
                {
                    lst_regions.Add(arr_Impedance2.ToArray());
                }
                if (arr_Impedance3.Count > 0)
                {
                    lst_regions.Add(arr_Impedance3.ToArray());
                }
                if (arr_Impedance4.Count > 0)
                {
                    lst_regions.Add(arr_Impedance4.ToArray());
                }


                DGV_spec_impedance.DataSource = dt_check_logfile(txtItemCode.Text, txtLotNo.Text, lst_regions);
                Double nom = Double.Parse(txt_nom_imp.Text);
                Double tol = Double.Parse(txt_tol_imp.Text);
                Double max_target = nom + (tol * nom)/100;
                Double min_target = nom - (tol * nom)/100;


                for (int i = 0; i < DGV_impedance.Rows.Count; i++)
                {

                    Double data = Double.Parse(DGV_impedance.Rows[i].Cells["Data"].Value.ToString());
                    if (data > max_target || data < min_target)
                    {
                        for (int j = 0; j < DGV_impedance.Columns.Count; j++)
                        {

                            DGV_impedance.Rows[i].Cells[j].Style.BackColor = Color.Red;
                        }
                    }

                }
               
            }
        }

        private void txtLogfile_TextChanged(object sender, EventArgs e)
        {

        }

        private void btn_check_tracewidth_Click(object sender, EventArgs e)
        {
            if (txt_nom_tracewidth.Text != "" && txt_tol_tracewidth.Text != "")
            {
                reset_DGV_color(DGV_tracewidth);
                DataTable dt = new DataTable();
                dt.Columns.Add("MAX");
                dt.Columns.Add("MIN");
                
                Double nom = Double.Parse(txt_nom_tracewidth.Text);
                Double tol = Double.Parse(txt_tol_tracewidth.Text);
                Double max_target = nom + (tol * nom) / 100;
                Double min_target = nom - (tol * nom) / 100;

                DataRow dr = dt.NewRow();
                dr[0] = max_target;
                dr[1] = min_target;

                dt.Rows.Add(dr);


                dgv_spec_tracewidth.DataSource = dt;


                for (int i = 0; i < DGV_tracewidth.Rows.Count; i++)
                {

                    Double data = Double.Parse(DGV_tracewidth.Rows[i].Cells["Data"].Value.ToString());
                    if (data > max_target || data < min_target)
                    {
                        for (int j = 0; j < DGV_tracewidth.Columns.Count; j++)
                        {

                            DGV_tracewidth.Rows[i].Cells[j].Style.BackColor = Color.Red;
                        }

                    }

                }
               
            }
            else
            {
                MessageBox.Show("Nhập đầy đủ thông tin", "Thông báo");
            }
        }

        private void txtLogfile_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                DGV_impedance.DataSource = Impedance_Data(txtLogfile.Text);
                DGV_tracewidth.DataSource = Tracewidth_Data(txtLogfile.Text);
                MessageBox.Show("Hoàn thành", "Thông báo");
            }
        }

        private void btnsave_Click(object sender, EventArgs e)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
            DGV_impedance.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_VAL", filter_str);
            DGV_tracewidth.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "DATA_LOG_FILE", filter_str);
        }

        private void Check_Impedance_logfile_FormClosed(object sender, FormClosedEventArgs e)
        {
            myVar._frmMain.Show();
        }
        public void Save_data()
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                List<string> tar_table_name = new List<string>() { "DATA_LOG_FILE", "IMPEDANCE_VAL" };
                List<DataTable> tar_table = new List<DataTable>() { (DataTable)DGV_tracewidth.DataSource, (DataTable)DGV_impedance.DataSource };
                int tbl_inx = 0;
                foreach (var t in tar_table_name)
                {
                    start_lbl: DataTable cur_dt = TDMK_Code.Datatable_Filter(sqlcon, t, filter_str);
                    if (cur_dt.Rows.Count > 0)
                    {
                        if (MessageBox.Show("Table " + t + ": Data is existed. Do you want to overwrite?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            TDMK_Code.Delelte_FilteredItem_arr(t, sqlcon, filter_str);
                            goto start_lbl;
                        }
                    }
                    else
                    {
                        int id = TDMK_Code.SQL_MAX(t, "ID", sqlcon);
                        int r_inx = 0;
                        foreach (DataRow dr in cur_dt.Rows)
                        {
                            dr["ID"] = id + 1 + r_inx;
                            r_inx++;
                        }
                        exp_proc.BatchBulkCopy(sqlcon, tar_table[tbl_inx], t);
                    }
                    tbl_inx++;

                }
                MessageBox.Show("Hoàn thành", "Thông báo");
            }
        }
    }
}
