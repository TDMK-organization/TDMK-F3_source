using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SEEV_DLL;
using TDMK_SQL;
using System.Data.SqlClient;
using myExcel = Microsoft.Office.Interop.Excel;
using OK2SHIP;
using System.IO;
using VHX;
using System.Diagnostics;
using System.Xml.Serialization;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace OK2SHIP
{
    public partial class ACF : Form
    {
        public ACF()
        {
            InitializeComponent();
        }

        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SEI_Lib myCode = new SEI_Lib();
        myVar exp_proc = new myVar();
        Export_Image_Class diem = new Export_Image_Class();
        Funtion_export F_export_new = new Funtion_export();
        ExportEPPlus F_exportEPPlus = new ExportEPPlus();
        SqlConnection sqlcon = null;
        SqlConnection sqlcon_IPQC = null;

        private void ACF_Load(object sender, EventArgs e)
        {
            sqlcon = exp_proc.initial_data(myVar.sel_DB, true);
            sqlcon_IPQC = exp_proc.initial_data("IPQC_Data", true);
        }

        public void Clear_DGV(DataGridView DGV)
        {
            DGV.DataSource = null;
            DGV.Columns.Clear();
            DGV.Rows.Clear();
        }



        private void tsmLoadData_Click(object sender, EventArgs e)
        {
            if (DGV_Data.DataSource != null)
            {
                Clear_DGV(DGV_Data);
            }
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && cb_data_for.SelectedIndex != -1)
            {
                if (cb_data_for.SelectedIndex != -1)
                {
                    lbltitle.Text = "Data For " + cb_data_for.Text;
                    string[] item = { "ItemCode", "LotNo" };
                    string[] item_val = { txtItemCode.Text, txtLotNo.Text };

                    if (cb_data_for.SelectedItem.ToString() == "ACF_ROUGHNESS")
                    {
                        //  SqlConnection sqlcon_IPQC = exp_proc.initial_data("IPQC_Data", true);
                        DGV_Data.DataSource = TDMK_Code.Datatable_Filter(sqlcon_IPQC, "Roughness", TDMK_Code.filter_str(item, item_val));
                        myCode.Disable_Sort_DGV(DGV_Data);
                        // sqlcon = exp_proc.initial_data(myVar.sel_DB, true);
                        DGV_Roughness_spec.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "ACF_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
                        if (DGV_Roughness_spec.Rows.Count == 0)
                        {
                            if (MessageBox.Show(new Form { TopMost = true }, "No Spec . Do you want to Load spec?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                            {
                                string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                                if (format_file != "")
                                {
                                    myExcel.Worksheet ws = TDMK_Code.open_excel_file(format_file, "", "").Sheets["OQC ACF"];
                                    DataTable dt = dt_roughness_spec_newformat(ws);
                                    DGV_Roughness_spec.DataSource = dt;
                                    exp_proc.BatchBulkCopy(sqlcon, dt, "ACF_SPEC");

                                }
                                else
                                {
                                    MessageBox.Show(new Form { TopMost = true }, "Format of " + txtItemCode.Text + " not found", "Warning");
                                }
                            }
                        }

                        Check_roughness(DGV_Data);
                    }
                    else
                    {
                        // sqlcon = exp_proc.initial_data(myVar.sel_DB, true);
                        DGV_Data.DataSource = TDMK_Code.Datatable_Filter(sqlcon, cb_data_for.SelectedItem.ToString(), TDMK_Code.filter_str(item, item_val));
                        myCode.Disable_Sort_DGV(DGV_Data);
                        if (DGV_Data.Rows.Count > 0)
                        {
                            if (cb_data_for.SelectedItem.ToString() == "ACF_GRAPH_FORCE_IMAGE")
                            {

                                ((DataGridViewImageColumn)DGV_Data.Columns["Image_Graph"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                                ((DataGridViewImageColumn)DGV_Data.Columns["Image_Graph"]).Width = 100;

                                foreach (DataGridViewRow dr in DGV_Data.Rows)
                                {
                                    dr.Height = 70;
                                }
                            }
                            else if (cb_data_for.SelectedItem.ToString() != "ACF_FLATNESS")
                            {

                                ((DataGridViewImageColumn)DGV_Data.Columns["Image_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                                ((DataGridViewImageColumn)DGV_Data.Columns["Image_Data"]).Width = 100;

                                foreach (DataGridViewRow dr in DGV_Data.Rows)
                                {
                                    dr.Height = 70;
                                }

                            }
                            change_ID_DGV(DGV_Data);
                        }
                        else
                        {
                            MessageBox.Show("No data", "Warning");
                        } 
                    } 
                }
                else
                {
                    MessageBox.Show("Please, Select Data For", "Warning");
                }
            }
            else
            {
                MessageBox.Show("Please, fill in ItemCode / LotNo ", "Warning");
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
                        temp_dr = myDt.AsEnumerable().Where(x => x.Field<byte[]>(tar_item) != null).ToArray();
                        if (temp_dr.Length > 0)
                        {
                            src_lst_tbl.Add(myDt);
                        }
                    }
                }
            }
            else
            {
                temp_dr = myDt.AsEnumerable().Where(x => x.Field<byte[]>(tar_item) != null).ToArray();
                if (temp_dr.Length > 0)
                {
                    src_lst_tbl.Add(myDt);
                }
            }
        }

        private void cb_data_for_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_data_for.SelectedIndex != -1)
            {

                cb_data_for.Text = cb_data_for.SelectedItem.ToString();
                DGV_Data.DataSource = null;
                picdetail.Image = null;
                if (cb_data_for.Text == "ALL")
                {
                    tsmLoadData.Enabled = false;
                }
                else
                {
                    tsmLoadData.Enabled = true;
                }
                if (cb_data_for.Text == "ACF_ROUGHNESS")
                {
                    DGV_Roughness_spec.Visible = true;
                    lbl_title_spec.Visible = true;
                    btn_load_spec.Visible = true;
                }
                else
                {
                    DGV_Roughness_spec.Visible = false;
                    lbl_title_spec.Visible = false;
                    btn_load_spec.Visible = false;
                }
                lbltitle.Text = "Data For ";
                lblImage.Text = "Image / Graph Details";
            }

        }
        public string Find_addr(string in_item, myExcel.Range start_rgn)
        {
            string result = start_rgn.AddressLocal.ToString();
            int r_inx = 0;
            while (r_inx < 30)
            {
                if (myCode.checkDBNull(start_rgn.Offset[r_inx, 0].Value) == in_item)
                {
                    result = start_rgn.Offset[r_inx, 0].AddressLocal;
                    break;
                }
                r_inx++;
            }
            return result;
        }


        public void Export_ACF_select(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno, List<string> search_item)
        {
            myExcel.Worksheet ws = wb.Sheets["ACF"];


            myExcel.Range sel_rgn = ws.Range["A16"];
            foreach (var item in search_item)
            {
                int num_insert = 5;
                string tar_table = "";
                string tar_col_data = "Image_Data";
                if (item.Contains("cleaning") || (item.Contains("OQC")))
                {
                    tar_table = "ACF_CLEANING_IMAGE";
                    num_insert = 10;
                }
                else
                {
                    if (item.Contains("picture"))
                    {
                        tar_table = "ACF_GRAPH_FORCE_IMAGE";
                        tar_col_data = "Image_Graph";
                    }
                    else
                    {
                        tar_table = "ACF_" + item.Trim().Replace(" ", "_").ToUpper() + "_IMAGE";
                    }
                }
                myExcel.Range data_rgn = ws.Range[Find_addr(item, sel_rgn)].Offset[0, 1];
                DataTable data_tbl = TDMK_Code.Datatable_Filter(sqlcon, tar_table, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { Itemcode, Lotno }));
                List<DataTable> src_lst_tbl = new List<DataTable>();
                Get_ListTable(-1, data_tbl, new string[] { "ItemCode", "LotNo", "Region" }, ref src_lst_tbl, tar_col_data);
                foreach (var dt in src_lst_tbl)
                {
                    if (dt.Rows.Count > 0)
                    {
                        string region = dt.Rows[0]["Region"].ToString();
                        int off_set = 0;
                        switch (region)
                        {
                            case "Picture":
                                off_set = 0;
                                break;
                            case "TRAI":
                                off_set = 0;
                                break;
                            case "GIUA":
                                off_set = 1;
                                break;
                            case "PHAI":
                                off_set = 2;
                                break;
                            default:
                                off_set = 0;
                                break;
                        }
                        int qty = Math.Min(dt.Rows.Count, num_insert);
                        for (int i = 0; i < qty; i++)
                        {
                            myExcel.Range cur_rgn = data_rgn.Offset[0, i + 5 * off_set];
                            byte[] data = (byte[])dt.Rows[i][tar_col_data];
                            using (MemoryStream ms = new MemoryStream(data))
                            {
                                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                                string file_dic = Path.Combine(file_folder, "1.jpg");
                                Image image = Image.FromStream(ms);
                                image.Save(file_dic);
                                diem.InsertPicture_Name(ws, cur_rgn, file_dic, 10);
                            }
                            if (tar_col_data.Contains("Graph"))
                            {
                                myExcel.Range force_rgn = data_rgn.Offset[1, i + 5 * off_set];
                                force_rgn.Value = dt.Rows[i]["Force_Data"];
                            }

                        }
                    }
                }
                sel_rgn = data_rgn.Offset[0, -1];
            }
        }

        //public void Export_ACF_Cleaning(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno, List<string> search_item)
        //{
        //    myExcel.Worksheet ws = wb.Sheets["ACF"];

        //    myExcel.Range sel_rgn = ws.Range["A16"];
        //    int no_sample = 0;
        //    int i = 1;

        //    while (sel_rgn.Offset[-1, i].Value.ToString() != null)
        //    {
        //        no_sample++;
        //        i++;
        //    }


        //    foreach (var item in search_item)
        //    {

        //        string tar_table = "ACF_CLEANING_IMAGE";
        //        string tar_col_data = "Image_Data";

        //        myExcel.Range data_rgn = ws.Range[Find_addr(item, sel_rgn)].Offset[0, 1];
        //        DataTable data_tbl = TDMK_Code.Datatable_Filter(sqlcon, tar_table, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { Itemcode, Lotno }));
        //        List<DataTable> src_lst_tbl = new List<DataTable>();
        //        Get_ListTable(-1, data_tbl, new string[] { "ItemCode", "LotNo", "Region" }, ref src_lst_tbl, tar_col_data);
        //        foreach (var dt in src_lst_tbl)
        //        {
        //            if (dt.Rows.Count > 0)
        //            {
        //                string region = dt.Rows[0]["Region"].ToString();
        //                int off_set = 0;

        //                int qty = Math.Min(dt.Rows.Count, 5);
        //                for (int i = 0; i < qty; i++)
        //                {
        //                    myExcel.Range cur_rgn = data_rgn.Offset[0, i + 5 * off_set];
        //                    byte[] data = (byte[])dt.Rows[i][tar_col_data];
        //                    using (MemoryStream ms = new MemoryStream(data))
        //                    {
        //                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
        //                        string file_dic = Path.Combine(file_folder, "1.jpg");
        //                        Image image = Image.FromStream(ms);
        //                        image.Save(file_dic);
        //                        diem.InsertPicture_Name(ws, cur_rgn, file_dic, 10);
        //                    }
        //                    if (tar_col_data.Contains("Graph"))
        //                    {
        //                        myExcel.Range force_rgn = data_rgn.Offset[1, i + 5 * off_set];
        //                        force_rgn.Value = dt.Rows[i]["Force_Data"];
        //                    }

        //                }
        //            }
        //        }
        //        sel_rgn = data_rgn.Offset[0, -1];
        //    }
        //}
        public void Export_ACFFlatness(SqlConnection sqlcon, myExcel.Worksheet ws, string ItemCode, string LotNo)
        {
            if (ItemCode != "" && LotNo != "")
            {
                DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, "ACF_FLATNESS", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                if (src_dt.Rows.Count > 0)
                {
                    myExcel.Range curr_rgn = ws.Range["A1"];
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
                    for (int i = 0; i < src_dt.Rows.Count; i++)
                    {
                        for (int j = 0; j < 10; j++)
                        {
                            string val = src_dt.Rows[i][4 + j].ToString();
                            if (myCode.IsNumeric(val))
                            {
                                //curr_rgn.Offset[i + 1, j].Value = src_dt.Rows[i][4 + j].ToString();
                                curr_rgn.Offset[i + 1, j].Value = double.Parse(val) * 1000;
                            }

                        }
                    }

                }

            }
        }




        public string ACF_pad_location(myExcel.Worksheet ws)
        {
            // List<int> lst_info = new List<int> { };
            string get_info = "";
            myExcel.Range curr_rgn_1 = ws.Range["A1"];

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

        public string ACF_pad_location_newformat(myExcel.Worksheet ws)
        {
            // List<int> lst_info = new List<int> { };
            string get_info = "";
            myExcel.Range curr_rgn_1 = ws.Range["A1"];

            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 30; i < 200; i++)
            {
                if (curr_rgn_1.Offset[i, 0].Value != null)
                {
                    if (curr_rgn_1.Offset[i, 0].Value.ToString().Replace(" ", "").ToUpper() == "ACF pads Surface Roughness".Replace(" ", "").ToUpper())
                    {
                        int k = 1;
                        int count_sample = 0;
                        get_info = i.ToString();
                        for (int m = 1; m < 20; m++)
                        {
                            if (myCode.checkDBNull(curr_rgn_1.Offset[i + m + count_sample, 0].Value).Replace(" ", string.Empty).ToUpper().Contains("Sample".Replace(" ", string.Empty).ToUpper()))
                            {
                                get_info = (i + m).ToString();

                                while (myCode.checkDBNull(curr_rgn_1.Offset[i + m + count_sample, 0].Value).Replace(" ", string.Empty).ToUpper().Contains("Sample".Replace(" ", string.Empty).ToUpper()))
                                {
                                    count_sample++;
                                }

                                break;
                            }
                        }


                        while (myCode.checkDBNull(curr_rgn_1.Offset[i + 2, k].Value).Replace(" ", string.Empty).ToUpper().Contains("surface".Replace(" ", string.Empty).ToUpper()))
                        {
                            if (myCode.checkDBNull(curr_rgn_1.Offset[i + 2, k].Value).Contains("Sa"))
                            {

                                get_info += "+ sa";
                            }
                            if (myCode.checkDBNull(curr_rgn_1.Offset[i + 2, k].Value).Contains("Sq"))
                            {

                                get_info += "+ sq";
                            }
                            if (myCode.checkDBNull(curr_rgn_1.Offset[i + 2, k].Value).Contains("Sdr"))
                            {

                                get_info += "+ sdr";
                            }
                            k++;
                        }

                        get_info += "+" + count_sample.ToString();



                        //for(int m = 1; m < 20; m++)
                        //{
                        //    if (curr_rgn_1.Offset[i + m, 0].Value.ToString().Replace(" ", "").ToUpper().Constain("Sample".Replace(" ", "").ToUpper()))
                        //    {

                        //    }
                        break;
                    }
                }
            }
            return get_info;


        }
        public void Check_roughness(DataGridView dgv)
        {
            string[] item = { "ItemCode" };
            string[] item_val = { txtItemCode.Text };
            DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "ACF_SPEC", TDMK_Code.filter_str(item, item_val));
            if (dt_spec.Rows.Count > 0)
            {
                for (int i = 0; i < DGV_Data.Rows.Count; i++)
                {
                    for (int k = 3; k < 11; k++)
                    {
                        string val = dt_spec.Rows[0][k].ToString();
                        if (val.Contains("USL"))
                        {
                            val = val.Replace("USL:", "");
                            if (Double.TryParse(DGV_Data.Rows[i].Cells[k + 3].Value.ToString(), out Double data))
                            {
                                if (data > Double.Parse(val) || data == Double.Parse(val))
                                {
                                    DGV_Data.Rows[i].Cells[k + 3].Style.BackColor = Color.Red;
                                }
                            }

                        }
                        if (val.Contains("LSL"))
                        {
                            val = val.Replace("LSL:", "");

                            if (Double.TryParse(DGV_Data.Rows[i].Cells[k + 3].Value.ToString(), out Double data))
                            {
                                if (data < Double.Parse(val) || data == Double.Parse(val))
                                {
                                    DGV_Data.Rows[i].Cells[k + 3].Style.BackColor = Color.Red;
                                }
                            }
                        }
                    }
                }
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
        public void Export_ACF_Roughness(string itemcode, string lotno, myExcel.Worksheet ws)
        {
            SqlConnection sqlcon_IPQC = exp_proc.initial_data("IPQC_Data", true);

            DataTable dt1 = new DataTable();
            dt1 = TDMK_Code.Datatable_Filter(sqlcon_IPQC, "Roughness", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcode, lotno }));
            if (dt1.Rows.Count > 0)
            {
                List<List<string>> lst_rghness = lst_roughness(dt1);

                myExcel.Range curr_rgn_1 = ws.Range["A1"];

                List<int> vitri_sa = new List<int> { };
                List<int> vitri_sq = new List<int> { };
                List<int> vitri_sdr = new List<int> { };
                //string[] arr_info = ACF_pad_location(ws).Split('+');
                string[] arr_info = ACF_pad_location_newformat(ws).Split('+');
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
                sqlcon = exp_proc.initial_data(myVar.sel_DB, true);
            }
        }


        public void Export_ACF_Roughness_newformat(string itemcode, string lotno, myExcel.Worksheet ws)
        {
            SqlConnection sqlcon_IPQC = exp_proc.initial_data("IPQC_Data", true);

            DataTable dt1 = new DataTable();
            dt1 = TDMK_Code.Datatable_Filter(sqlcon_IPQC, "Roughness", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcode, lotno }));
            if (dt1.Rows.Count > 0)
            {
                List<List<string>> lst_rghness = lst_roughness(dt1);

                myExcel.Range curr_rgn_1 = ws.Range["A1"];

                List<int> vitri_sa = new List<int> { };
                List<int> vitri_sq = new List<int> { };
                List<int> vitri_sdr = new List<int> { };
                string[] arr_info = ACF_pad_location_newformat(ws).Split('+');
                string vitri_ACF_pad = arr_info[0];

                if (myCode.IsNumeric(arr_info[0]))
                {
                    int row = int.Parse(arr_info[0]);
                    curr_rgn_1 = curr_rgn_1.Offset[row, 0];
                    int r = 1;
                    for (int i = 0; i < 9; i++)
                    {
                        if (r < arr_info.Length - 1)
                        {
                            if (arr_info[r].Contains("sa"))
                            {
                                vitri_sa.Add(i);
                            }
                            if (arr_info[r].Contains("sq"))
                            {
                                vitri_sq.Add(i);
                            }
                            if (arr_info[r].Contains("sdr"))
                            {
                                vitri_sdr.Add(i);
                            }
                        }

                        r++;
                    }

                    int m = 0;
                    int k;

                    int count_pt = dt1.Rows.Count;
                    while (m < count_pt)
                    {
                        int c = 0;
                        foreach (int j in vitri_sa)
                        {
                            curr_rgn_1.Offset[m, j + 1].Value = lst_rghness[c * 3][m];
                            c++;

                        }
                        c = 0;
                        foreach (int j in vitri_sq)
                        {
                            curr_rgn_1.Offset[m, j + 1].Value = lst_rghness[1 + c * 3][m];
                            c++;
                        }
                        c = 0;
                        foreach (int j in vitri_sdr)
                        {
                            curr_rgn_1.Offset[m, j + 1].Value = lst_rghness[2 + c * 3][m];
                            c++;
                        }
                        m++;

                        // curr_rgn_1 = curr_rgn_1.Offset[r_count, 0];


                    }
                }


                //  int count_sample = int.Parse(arr_info[arr_info.Length - 1]);

                sqlcon = exp_proc.initial_data(myVar.sel_DB, true);
            }
        }
        public void Export_ACF_Image(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            myExcel.Worksheet ws = wb.Sheets["ACF"];
            List<string> search_item = new List<string>() { "Photo before cleaning", "Photo after cleaning", "Photo after OQC testing", "Before tape test", "After tape test", "Tape test picture" };//
            myExcel.Range sel_rgn = ws.Range["A16"];

            foreach (var item in search_item)
            {
                int num_insert = 5;
                string tar_table = "";
                string tar_col_data = "Image_Data";
                if (item.Contains("cleaning") || (item.Contains("OQC")))
                {
                    tar_table = "ACF_CLEANING_IMAGE";
                    num_insert = 10;
                }
                else
                {
                    if (item.Contains("picture"))
                    {
                        tar_table = "ACF_GRAPH_FORCE_IMAGE";
                        tar_col_data = "Image_Graph";
                    }
                    else
                    {
                        tar_table = "ACF_" + item.Trim().Replace(" ", "_").ToUpper() + "_IMAGE";
                    }
                }
                myExcel.Range data_rgn = ws.Range[Find_addr(item, sel_rgn)].Offset[0, 1];
                DataTable data_tbl = TDMK_Code.Datatable_Filter(sqlcon, tar_table, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { Itemcode, Lotno }));
                List<DataTable> src_lst_tbl = new List<DataTable>();
                Get_ListTable(-1, data_tbl, new string[] { "ItemCode", "LotNo", "Region" }, ref src_lst_tbl, tar_col_data);
                foreach (var dt in src_lst_tbl)
                {
                    if (dt.Rows.Count > 0)
                    {
                        string region = dt.Rows[0]["Region"].ToString();
                        int off_set = 0;
                        switch (region)
                        {
                            case "Picture":
                                off_set = 0;
                                break;
                            case "TRAI":
                                off_set = 0;
                                break;
                            case "GIUA":
                                off_set = 1;
                                break;
                            case "PHAI":
                                off_set = 2;
                                break;
                            default:
                                off_set = 0;
                                break;
                        }
                        int qty = Math.Min(dt.Rows.Count, num_insert);
                        for (int i = 0; i < qty; i++)
                        {
                            myExcel.Range cur_rgn = data_rgn.Offset[0, i + 5 * off_set];
                            byte[] data = (byte[])dt.Rows[i][tar_col_data];
                            using (MemoryStream ms = new MemoryStream(data))
                            {
                                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                                string file_dic = Path.Combine(file_folder, "1.jpg");
                                Image image = Image.FromStream(ms);
                                image.Save(file_dic);
                                diem.InsertPicture_Name(ws, cur_rgn, file_dic, 10);
                            }
                            if (tar_col_data.Contains("Graph"))
                            {
                                myExcel.Range force_rgn = data_rgn.Offset[1, i + 5 * off_set];
                                force_rgn.Value = dt.Rows[i]["Force_Data"];
                            }
                        }
                    }
                }
                sel_rgn = data_rgn.Offset[0, -1];
            }
        }
        public void save_report_new(string ItemCode, string LotNo, string id, myExcel.Workbook curr_wrkbook, string sheet, string fpath)
        {
            string report_path = "";
            DirectoryInfo tar_parent = new DirectoryInfo(fpath);
            DirectoryInfo[] arr_dic_child = tar_parent.GetDirectories();
            foreach (DirectoryInfo folder_sheet in arr_dic_child)
            {
                string f_name = folder_sheet.Name;
                if (f_name.Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper().Contains(sheet.Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper()))
                {
                    report_path = Path.Combine(fpath, f_name, ItemCode + "-" + LotNo + "_" + id);
                    if (File.Exists(report_path))
                        File.Delete(report_path);

                    try
                    {
                        curr_wrkbook.SaveAs(report_path);
                    }
                    catch
                    {

                    }

                    break;
                }
            }

        }

        private void tsmExport_Click(object sender, EventArgs e)
        {
            if (cb_data_for.SelectedIndex != -1)
            {
                if (cb_data_for.Text == "ALL")
                {
                    if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtOperator.Text != "")
                    {
                        string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                        if (format_file != "")
                        {
                            string export_path = F_exportEPPlus.get_reportpath(txtItemCode.Text, txtLotNo.Text, txtOperator.Text, "OQC ACF", myVar.data_loc);
                            if (export_path != "")
                            {
                                // SqlConnection sqlcon_IPQC = null;

                                F_exportEPPlus.Export_EPPlus(sqlcon, format_file, export_path, txtItemCode.Text, txtLotNo.Text, "OQC ACF", sqlcon_IPQC);
                            }
                            else
                            {
                                MessageBox.Show(new Form { TopMost = true }, "Folder OQC ACF not found!", "Warning");
                            }
                        }
                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Format of " + txtItemCode.Text + " not found", "Warning");
                        }
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Please, fill in ItemCode / LotNo ", "Warning");
                    }
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Please select ALL to export", "Warning");
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Please, Select Data For", "Warning");
            }

        }


        private void DGV_Data_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int col_inx = e.ColumnIndex;
            int r_inx = e.RowIndex;
            if (col_inx != -1 && r_inx != -1)
            {
                DataGridViewCell cur_cell = DGV_Data.CurrentCell;
                if (DGV_Data.Columns[col_inx].Name.Contains("Image"))
                {
                    byte[] data = (byte[])cur_cell.Value;
                    using (MemoryStream ms = new MemoryStream(data))
                    {
                        picdetail.Image = Image.FromStream(ms);
                    }
                    string region = DGV_Data.Rows[r_inx].Cells["Region"].Value.ToString();
                    string pcs = DGV_Data.Rows[r_inx].Cells["Pcs_No"].Value.ToString();
                    lblImage.Text = "Image / Graph Details ---> Region: " + region + " / " + "Pcs_No: " + pcs;
                }
            }

        }
        public List<string> List_rougness_spec(myExcel.Worksheet ws)
        {
            List<string> lst_info = new List<string> { };
            //string get_info = "";
            myExcel.Range curr_rgn_1 = ws.Range["A1"];

            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 30; i < 100; i++)
            {
                if (curr_rgn_1.Offset[i, 0].Value != null)
                {
                    if (curr_rgn_1.Offset[i, 0].Value.ToString().Replace(" ", "").ToUpper() == "ACF pad location".Replace(" ", "").ToUpper())
                    {
                        int k = 1;
                        // get_info = i.ToString();
                        while (curr_rgn_1.Offset[i + k, 1].Value.ToString() != "Roughness")
                        {
                            int L1 = 0;
                            int L2 = 0;
                            int L3 = 0;
                            if (curr_rgn_1.Offset[i + k, 1].Value.ToString().Contains("Sa"))
                            {
                                L1++;
                                string value = curr_rgn_1.Offset[i + k, 1].Value.ToString().Replace("Sa(", "").Replace(")", "");
                                lst_info.Add("L" + L1.ToString() + "_Sa+" + value);
                            }
                            if (curr_rgn_1.Offset[i + k, 1].Value.ToString().Contains("Sq"))
                            {
                                L2++;
                                string value = curr_rgn_1.Offset[i + k, 1].Value.ToString().Replace("Sq(", "").Replace(")", "");
                                lst_info.Add("L" + L2.ToString() + "_Sq+" + value);
                            }
                            if (curr_rgn_1.Offset[i + k, 1].Value.ToString().Contains("Sdr"))
                            {
                                L3++;
                                string value = curr_rgn_1.Offset[i + k, 1].Value.ToString().Replace("Sdr(", "").Replace(")", "");
                                lst_info.Add("L" + L3.ToString() + "_Sdr+" + value);
                            }
                            k++;
                        }


                        break;
                    }
                }
            }
            return lst_info;


        }

        public DataTable dt_roughness_spec(myExcel.Worksheet ws)
        {
            List<string> lst_info = new List<string> { };

            myExcel.Range curr_rgn_1 = ws.Range["A1"];

            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 30; i < 100; i++)
            {
                if (curr_rgn_1.Offset[i, 0].Value != null)
                {
                    if (curr_rgn_1.Offset[i, 0].Value.ToString().Replace(" ", "").ToUpper() == "ACF pad location".Replace(" ", "").ToUpper())
                    {
                        int k = 1;
                        // get_info = i.ToString();
                        int L1 = 0;
                        int L2 = 0;
                        int L3 = 0;
                        while (curr_rgn_1.Offset[i + k, 1].Value.ToString() != "Roughness")
                        {

                            if (curr_rgn_1.Offset[i + k, 1].Value.ToString().Contains("Sa"))
                            {
                                L1++;
                                string value = curr_rgn_1.Offset[i + k, 1].Value.ToString().Replace(" ", "").Replace("Sa(", "").Replace(")", "").Replace("um", "");
                                lst_info.Add("L" + L1.ToString() + "_Sa+" + value);
                            }
                            if (curr_rgn_1.Offset[i + k, 1].Value.ToString().Contains("Sq"))
                            {
                                L2++;
                                string value = curr_rgn_1.Offset[i + k, 1].Value.ToString().Replace(" ", "").Replace("Sq(", "").Replace(")", "").Replace("um", "");
                                lst_info.Add("L" + L2.ToString() + "_Sq+" + value);
                            }
                            if (curr_rgn_1.Offset[i + k, 1].Value.ToString().Contains("Sdr"))
                            {
                                L3++;
                                string value = curr_rgn_1.Offset[i + k, 1].Value.ToString().Replace(" ", "").Replace("Sdr(", "").Replace(")", "").Replace("um", "");
                                lst_info.Add("L" + L3.ToString() + "_Sdr+" + value);
                            }
                            k++;
                        }


                        break;
                    }
                }
            }

            string[] item = { "ItemCode" };
            string[] item_val = { txtItemCode.Text };
            DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "ACF_SPEC", TDMK_Code.filter_str(item, item_val)).Clone();
            DataRow dr = dt_spec.NewRow();
            dr[0] = TDMK_Code.SQL_MAX("ACF_SPEC", "ID", sqlcon) + 1;
            dr[1] = txtItemCode.Text;
            //dr[2] = txtLotNo.Text;


            foreach (string itm in lst_info)
            {
                if (itm.Contains("L1_Sa"))
                    dr[2] = itm.Split('+')[1];
                if (itm.Contains("L1_Sq"))
                    dr[3] = itm.Split('+')[1];
                if (itm.Contains("L1_Sdr"))
                    dr[4] = itm.Split('+')[1];
                if (itm.Contains("L2_Sa"))
                    dr[5] = itm.Split('+')[1];
                if (itm.Contains("L2_Sq"))
                    dr[6] = itm.Split('+')[1];
                if (itm.Contains("L2_Sdr"))
                    dr[7] = itm.Split('+')[1];
                if (itm.Contains("L3_Sa"))
                    dr[8] = itm.Split('+')[1];
                if (itm.Contains("L3_Sq"))
                    dr[9] = itm.Split('+')[1];
                if (itm.Contains("L3_Sdr"))
                    dr[10] = itm.Split('+')[1];

            }
            dt_spec.Rows.Add(dr);


            return dt_spec;

        }


        public DataTable dt_roughness_spec_newformat(myExcel.Worksheet ws)
        {
            List<string> lst_info = new List<string> { };

            myExcel.Range curr_rgn_1 = ws.Range["A1"];

            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 30; i < 200; i++)
            {
                if (curr_rgn_1.Offset[i, 0].Value != null)
                {
                    if (curr_rgn_1.Offset[i, 0].Value.ToString().Replace(" ", "").ToUpper() == "ACF pads Surface Roughness".Replace(" ", "").ToUpper())
                    {
                        int k = 1;
                        // get_info = i.ToString();
                        int L1 = 0;
                        int L2 = 0;
                        int L3 = 0;

                        while (myCode.checkDBNull(curr_rgn_1.Offset[i + 2, k].Value).Replace(" ", string.Empty).ToUpper().Contains("surface".Replace(" ", string.Empty).ToUpper()))
                        {
                            string type = myCode.checkDBNull(curr_rgn_1.Offset[i + 3, k].Value);
                            string USL = myCode.checkDBNull(curr_rgn_1.Offset[i + 4, k].Value);
                            string LSL = myCode.checkDBNull(curr_rgn_1.Offset[i + 5, k].Value);
                            string f = "USL";
                            if (type.ToUpper().Contains("LSL"))
                            {
                                f = "LSL";
                            }
                            if (myCode.checkDBNull(curr_rgn_1.Offset[i + 2, k].Value).Contains("Sa"))
                            {
                                L1++;
                                string value = f + ":" + myCode.checkDBNull(curr_rgn_1.Offset[i + 4, k].Value);
                                lst_info.Add("L" + L1.ToString() + "_Sa+" + value);
                            }
                            if (myCode.checkDBNull(curr_rgn_1.Offset[i + 2, k].Value).Contains("Sq"))
                            {
                                L2++;
                                string value = f + ":" + myCode.checkDBNull(curr_rgn_1.Offset[i + 4, k].Value);
                                lst_info.Add("L" + L2.ToString() + "_Sq+" + value);
                            }
                            if (myCode.checkDBNull(curr_rgn_1.Offset[i + 2, k].Value).Contains("Sdr"))
                            {
                                L3++;
                                string value = f + ":" + myCode.checkDBNull(curr_rgn_1.Offset[i + 5, k].Value);
                                lst_info.Add("L" + L3.ToString() + "_Sdr+" + value);
                            }
                            k++;
                        }



                    }
                }
            }

            string[] item = { "ItemCode" };
            string[] item_val = { txtItemCode.Text };
            DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "ACF_SPEC", TDMK_Code.filter_str(item, item_val)).Clone();
            DataRow dr = dt_spec.NewRow();
            dr[0] = TDMK_Code.SQL_MAX("ACF_SPEC", "ID", sqlcon) + 1;
            dr[1] = txtItemCode.Text;
            //dr[2] = txtLotNo.Text;


            foreach (string itm in lst_info)
            {
                if (itm.Contains("L1_Sa"))
                    dr[2] = itm.Split('+')[1];
                if (itm.Contains("L1_Sq"))
                    dr[3] = itm.Split('+')[1];
                if (itm.Contains("L1_Sdr"))
                    dr[4] = itm.Split('+')[1];
                if (itm.Contains("L2_Sa"))
                    dr[5] = itm.Split('+')[1];
                if (itm.Contains("L2_Sq"))
                    dr[6] = itm.Split('+')[1];
                if (itm.Contains("L2_Sdr"))
                    dr[7] = itm.Split('+')[1];
                if (itm.Contains("L3_Sa"))
                    dr[8] = itm.Split('+')[1];
                if (itm.Contains("L3_Sq"))
                    dr[9] = itm.Split('+')[1];
                if (itm.Contains("L3_Sdr"))
                    dr[10] = itm.Split('+')[1];

            }
            dt_spec.Rows.Add(dr);


            return dt_spec;

        }

        private void mnuAction_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void btn_load_spec_Click(object sender, EventArgs e)
        {
            string[] item = { "ItemCode" };
            string[] item_val = { txtItemCode.Text };
start_label: DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "ACF_SPEC", TDMK_Code.filter_str(item, item_val));
            string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
            if (dt_spec.Rows.Count > 0)
            {
                if (MessageBox.Show(new Form { TopMost = true }, "Spec is existed.Do you want to overwrite ? ", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    
                    if (format_file != "")
                    {
                        TDMK_Code.Delelte_FilteredItem_arr("ACF_SPEC", sqlcon, TDMK_Code.filter_str(item, item_val));
                        goto start_label;
                        //myExcel.Worksheet ws = TDMK_Code.open_excel_file(format_file, "", "").Sheets["ACF"];
                        //DataTable dt = dt_roughness_spec_newformat(ws);
                        //DGV_Roughness_spec.DataSource = dt;

                        //exp_proc.BatchBulkCopy(sqlcon, dt, "ACF_SPEC");

                        //MessageBox.Show("Setup Completed", "Warning");

                    }
                    //else
                    //{
                    //    MessageBox.Show(new Form { TopMost = true }, "Format of " + txtItemCode.Text + " not found", "Warning");
                    //}
                }
            }
            else
            {
                //string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                if (format_file != "")
                {
                    myExcel.Workbook wb = TDMK_Code.open_excel_file(format_file, "", "");
                    if (wb != null)
                    {
                        foreach (myExcel.Worksheet tg_sht in wb.Worksheets)
                        {
                            string cur_sht_name = tg_sht.Name.ToUpper().Replace("-", "").Replace(" ", "");
                            if (cur_sht_name.Contains("ACF"))
                            {
                                myExcel.Worksheet ws = tg_sht;
                                DataTable dt = dt_roughness_spec_newformat(ws);
                                DGV_Roughness_spec.DataSource = dt;

                                exp_proc.BatchBulkCopy(sqlcon, dt, "ACF_SPEC");
                                MessageBox.Show("Setup Completed", "Warning");
                                break;
                            }
                        }

                    }
                    else
                    {
                        MessageBox.Show("Error", "Warning");
                    }

                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Format of " + txtItemCode.Text + " not found", "Warning");
                }
            }
            Check_roughness(DGV_Data);
        }

        private void txtLotNo_Validated(object sender, EventArgs e)
        {
            txtLotNo.Text = exp_proc.Lotno_Formated(txtLotNo.Text);
        }
        public void change_ID_DGV(DataGridView dgv)
        {
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                dgv.Rows[i].Cells["ID"].Value = i + 1;
            }

        }

        private void txtItemCode_TextChanged(object sender, EventArgs e)
        {
            DGV_Data.DataSource = DGV_Roughness_spec.DataSource = null;
            picdetail.Image = null;
            //cb_data_for.SelectedIndex = -1;
        }

        private void txtLotNo_TextChanged(object sender, EventArgs e)
        {
            DGV_Data.DataSource = DGV_Roughness_spec.DataSource = null;
            picdetail.Image = null;
            // cb_data_for.SelectedIndex = -1;
        }
    }
}
