using Echeck_LogFile_Process;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using OK2SHIP;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;
using System.Xml.Linq;
using TDMK_SEEV_DLL;
using TDMK_SQL;
using VHX;
using ZedGraph;
using DataTable = System.Data.DataTable;
using myExcel = Microsoft.Office.Interop.Excel;

namespace VHX
{
    public partial class BVH_PTH_Update : Form
    {
        Impedance_data proc_data = new Impedance_data();
        myVar exp_proc = new myVar();
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SEI_Lib myCode = new SEI_Lib();
        ExportEPPlus F_exportEPPlus = new ExportEPPlus();
        Impedance_main Imp_code = new Impedance_main();
        bool check_save_more = false;
        string[] arr_item_BVH_with = new string[13];
        string[] arr_item_BVH_without = new string[13];
        string[] arr_item_PTH = new string[12];

        public string strcon = "";
        public SqlConnection sqlcon = null;
        public SortedDictionary<int, byte[]> Image_result = null;
        public SortedDictionary<int, byte[]> Image_result_2 = null;


        public string depart = "QA";
        public bool admin_mode = false;
        public SortedDictionary<int, List<string>> edited_lst;
        SortedDictionary<int, List<string>> sel_cells_lst = new SortedDictionary<int, List<string>>();
        bool update_mode = false;
        bool insert_mode = false;

        int count_doday = 0;

        public BVH_PTH_Update()
        {
            InitializeComponent();
        }

        private void tsmSaveData_Click(object sender, EventArgs e)
        {

        }

        private void tsmExport_Click(object sender, EventArgs e)
        {

        }

        private void cb_process_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "")
            {

                DGV_logfile.DataSource = null;
                DGV_data_process.DataSource = null;

                //btn_load.Enabled = true;
                btn_insert.Enabled = false;
                if (cb_process.SelectedIndex != -1 && txtLotNo.Text != "" && txtItemCode.Text != "")
                {
                    DGV_spec.DataSource = null;
                    lbl_title.Text = "Data Analysis";
                    //DGV_spec.DataSource = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SPEC_" + cb_process.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
                    DataTable dt_spec = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SPEC_" + cb_process.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
                    if (dt_spec.Rows.Count > 0)
                    {
                        int count_pcs = int.Parse(dt_spec.Rows[0]["Remark"].ToString().Split('+')[2]);
                        if (count_pcs != 0)
                        {
                            DGV_spec.DataSource = dt_spec;
                        }
                        else
                        {
                            MessageBox.Show("This process can not input data", "Warning");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please click SETUP to setup spec ", "Warning");
                    }

                    // DataTable dt = TDMK_Code.Datatable_Filter(myVar.mysqlcon, cb_process.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));


                }
            }


        }
        public void insert_data_update(string tbl_process, SortedDictionary<int, byte[]> Image_result, DataTable tbl_data_all)
        {

            DataTable dt_master = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_MASTERLIST", TDMK_Code.filter_str(new string[] { "ItemCode", "Process" }, new string[] { txtItemCode.Text, cb_process.SelectedItem.ToString() }));
            if (dt_master.Rows.Count == 0)
            {
                dt_master = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_MASTERLIST", TDMK_Code.filter_str(new string[] { "ItemCode", "Process" }, new string[] { "22B0500", cb_process.SelectedItem.ToString() }));
            }
            if (dt_master.Rows.Count > 0)
            {
                List<int> postno = new List<int> { };
                for (int i = 0; i < dt_master.Rows.Count; i++)
                {
                    if (dt_master.Rows[i]["Post_No"].ToString() != "")
                    {
                        postno.Add(int.Parse(dt_master.Rows[i]["Post_No"].ToString()));
                    }
                    else
                    {
                        postno.Add(0);
                    }

                }


                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                DataTable dt_process = TDMK_Code.Datatable_Filter(sqlcon, tbl_process, filter_str).Clone();
                DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, tbl_process, filter_str);
                string[] remark_all = dt.AsEnumerable().Select(x => x.Field<string>("Remark")).Distinct().ToArray();

                int r = 0;
                foreach (var log in Image_result)
                {
                    while (r < dt.Rows.Count)
                    {
                        if (dt.Rows[r]["PCS_No"].ToString() == log.Key.ToString() && dt.Rows[r]["Remark"].ToString() == "")
                        {

                            DataRow dr = dt.NewRow();
                            int count_row = dt.Rows.Count;

                            dr[0] = count_row;
                            for (int j = 1; j < dt.Columns.Count - 1; j++)
                            {
                                dr[j] = dt.Rows[r][j];
                            }

                            dr["Remark"] = "NG_F" + remark_all.Length.ToString();
                            dt.Rows.Add(dr);

                            dt.Rows[r]["Image_data"] = log.Value;
                            break;
                        }
                        r++;

                    }

                    DGV_data_process.DataSource = dt;
                    ((DataGridViewImageColumn)DGV_data_process.Columns["Image_data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)DGV_data_process.Columns["Image_data"]).Width = 100;

                    foreach (DataGridViewRow dr in DGV_data_process.Rows)
                    {
                        dr.Height = 70;
                    }

                }
            }
        }
        public void insert_data(string tbl_process, SortedDictionary<int, byte[]> Image_result, DataTable tbl_data_all)
        {

            DataTable dt_master = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_MASTERLIST", TDMK_Code.filter_str(new string[] { "ItemCode", "Process" }, new string[] { txtItemCode.Text, cb_process.SelectedItem.ToString() }));
            if (dt_master.Rows.Count == 0)
            {
                dt_master = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_MASTERLIST", TDMK_Code.filter_str(new string[] { "ItemCode", "Process" }, new string[] { "22B0500", cb_process.SelectedItem.ToString() }));
            }
            if (dt_master.Rows.Count > 0)
            {
                List<int> postno = new List<int> { };
                for (int i = 0; i < dt_master.Rows.Count; i++)
                {
                    if (dt_master.Rows[i]["Post_No"].ToString() != "")
                    {
                        postno.Add(int.Parse(dt_master.Rows[i]["Post_No"].ToString()));
                    }
                    else
                    {
                        postno.Add(0);
                    }

                }


                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                DataTable dt_process = TDMK_Code.Datatable_Filter(sqlcon, tbl_process, filter_str).Clone();

                int indx = 0;
                foreach (var log in Image_result)
                {

                    DataTable tbl_data = tbl_data_all.Clone();
                    foreach (DataRow dr in tbl_data_all.Rows)
                    {
                        if (dr["PCS_No"].ToString() == log.Key.ToString())
                        {
                            DataRow row = tbl_data.NewRow();
                            for (int i = 0; i < tbl_data.Columns.Count; i++)
                            {
                                row[i] = dr[i];
                            }
                            tbl_data.Rows.Add(row);
                        }

                    }

                    //tbl_data = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process", "PCS_No" }, new string[] { txtItemCode.Text, txtLotNo.Text, cb_process.SelectedItem.ToString(), (indx + 1).ToString() }));
                    if (tbl_data.Rows.Count > 0)
                    {
                        DataRow dr = dt_process.NewRow();
                        dr[0] = (indx + 1).ToString();
                        dr[1] = txtItemCode.Text;
                        dr[2] = txtLotNo.Text;
                        dr[3] = log.Key;
                        for (int m = 4; m < dt_process.Columns.Count - 2; m++)
                        {
                            if (m - 4 < postno.Count)
                            {
                                if (postno[m - 4] != 0 && postno[m - 4] <= tbl_data.Rows.Count)
                                {
                                    dr[m] = tbl_data.Rows[postno[m - 4] - 1][6].ToString();
                                }
                            }


                        }
                        dr["Image_data"] = log.Value;
                        dr["Remark"] = tbl_data_all.Rows[0]["Remark"].ToString();

                        dt_process.Rows.Add(dr);
                    }
                    indx++;
                    //if (indx > log.Key)
                    //    break;
                }

                DGV_data_process.DataSource = dt_process;
                ((DataGridViewImageColumn)DGV_data_process.Columns["Image_data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)DGV_data_process.Columns["Image_data"]).Width = 100;

                foreach (DataGridViewRow dr in DGV_data_process.Rows)
                {
                    dr.Height = 70;
                }


            }
        }
        public void insert_data_2_old(string tbl_process, SortedDictionary<int, byte[]> Image_result, SortedDictionary<int, byte[]> Image_result_2, DataTable tbl_data_all)
        {
            DataTable dt_master = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_MASTERLIST", TDMK_Code.filter_str(new string[] { "ItemCode", "Process" }, new string[] { txtItemCode.Text, cb_process.SelectedItem.ToString() }));
            if (dt_master.Rows.Count == 0)
            {
                dt_master = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_MASTERLIST", TDMK_Code.filter_str(new string[] { "ItemCode", "Process" }, new string[] { "22B0500", cb_process.SelectedItem.ToString() }));
            }
            if (dt_master.Rows.Count > 0)
            {
                List<int> postno = new List<int> { };
                //for (int i = 0; i < dt_master.Rows.Count; i++)
                //{
                //    if (dt_master.Rows[i]["Post_No"].ToString() != "")
                //    {
                //        postno.Add(int.Parse(dt_master.Rows[i]["Post_No"].ToString()));
                //    }
                //    else
                //    {
                //        postno.Add(0);
                //    }
                //}

                for (int i = 0; i < dt_master.Rows.Count; i++)
                {
                    if (dt_master.Rows[i]["Post_No"].ToString().ToUpper().Replace(" ", "").Contains("DODAY"))
                    {
                        string a = dt_master.Rows[i]["Post_No"].ToString().ToUpper().Replace(" ", "").Replace("DODAY", "").Replace("_", "").Replace("-", "");
                        if (myCode.IsNumeric(a))
                        {
                            postno.Add(int.Parse(a));
                        }     
                    }
                    if (dt_master.Rows[i]["Post_No"].ToString() == "")
                    {
                        postno.Add(0);
                    }
                }
                //int count_pcs_doday = 0;
                //foreach (int i in postno)
                //{
                //    if(i != 0)
                //    {
                //        count_pcs_doday++;
                //    } 
                //}

               
                for(int i = 0; i < dt_master.Rows.Count; i++)
                {
                    if (dt_master.Rows[i]["Post_No"].ToString().ToUpper().Replace(" ", "").Contains("DK"))
                    {
                        string a = dt_master.Rows[i]["Post_No"].ToString().ToUpper().Replace(" ", "").Replace("DK", "").Replace("_", "").Replace("-", "");
                        if (myCode.IsNumeric(a))
                        {
                            postno.Add(int.Parse(a) + count_doday);
                        }
                    }
                       
                }
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                DataTable dt_process = TDMK_Code.Datatable_Filter(sqlcon, tbl_process, filter_str).Clone();

                int indx = 0;
                foreach (var log in Image_result)
                {

                    DataTable tbl_data = tbl_data_all.Clone();
                    foreach (DataRow dr in tbl_data_all.Rows)
                    {
                        if (dr["PCS_No"].ToString() == log.Key.ToString())
                        {
                            DataRow row = tbl_data.NewRow();
                            for (int i = 0; i < tbl_data.Columns.Count; i++)
                            {
                                row[i] = dr[i];
                            }
                            tbl_data.Rows.Add(row);
                        } 
                    }

                    //tbl_data = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process", "PCS_No" }, new string[] { txtItemCode.Text, txtLotNo.Text, cb_process.SelectedItem.ToString(), (indx + 1).ToString() }));
                    if (tbl_data.Rows.Count > 0)
                    {
                        DataRow dr = dt_process.NewRow();
                        dr[0] = (indx + 1).ToString();
                        dr[1] = txtItemCode.Text;
                        dr[2] = txtLotNo.Text;
                        dr[3] = log.Key;
                        for (int m = 4; m < dt_process.Columns.Count - 3; m++)
                        {
                            if (postno[m - 4] != 0 && postno[m - 4] <= tbl_data.Rows.Count)
                            {
                                dr[m] = tbl_data.Rows[postno[m - 4] - 1][6].ToString();
                            }

                        }

                        dr["Image_data"] = log.Value;
                        int key_image_2 = log.Key;
                        if (Image_result_2.ContainsKey(key_image_2))
                        {
                            dr["Image_data_2"] = Image_result_2[key_image_2];
                        }
                        else
                        {
                            dr["Image_data_2"] = null;
                        }

                        dr["Remark"] = tbl_data.Rows[0]["Remark"].ToString();

                        dt_process.Rows.Add(dr);
                    }
                    indx++;
                    //if (indx > log.Key)
                    //    break;
                }

                DGV_data_process.DataSource = dt_process;
                ((DataGridViewImageColumn)DGV_data_process.Columns["Image_data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)DGV_data_process.Columns["Image_data"]).Width = 100;
                ((DataGridViewImageColumn)DGV_data_process.Columns["Image_data_2"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)DGV_data_process.Columns["Image_data_2"]).Width = 100;

                foreach (DataGridViewRow dr in DGV_data_process.Rows)
                {
                    dr.Height = 70;
                }
            }
        }
        public void insert_data_2(string tbl_process, SortedDictionary<int, byte[]> Image_result, SortedDictionary<int, byte[]> Image_result_2, DataTable tbl_data_all)
        {
            DataTable dt_master = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_MASTERLIST", TDMK_Code.filter_str(new string[] { "ItemCode", "Process" }, new string[] { txtItemCode.Text, cb_process.SelectedItem.ToString() }));
            if (dt_master.Rows.Count == 0)
            {
                dt_master = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_MASTERLIST", TDMK_Code.filter_str(new string[] { "ItemCode", "Process" }, new string[] { "22B0500", cb_process.SelectedItem.ToString() }));
            }
            if (dt_master.Rows.Count > 0)
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                DataTable dt_process = TDMK_Code.Datatable_Filter(sqlcon, tbl_process, filter_str).Clone();
                List<DataTable> pcs_dt_list = new List<DataTable>();
                exp_proc.Get_ListTable(-1, tbl_data_all, new string[] { "PCS_No" }, ref pcs_dt_list, "Post_No");
                int indx = 0;
                foreach (DataTable dt in pcs_dt_list)
                {
                    if(dt.Rows.Count > 0)
                    {
                        string cur_pcs_no = myCode.checkDBNull( dt.Rows[0]["PCS_No"]);
                        DataRow dr = dt_process.NewRow();
                        dr[0] = (indx++).ToString();
                        dr[1] = txtItemCode.Text;
                        dr[2] = txtLotNo.Text;
                        dr[3] = cur_pcs_no;
                        foreach (DataRow dr_src in dt.Rows)
                        {
                            string post_no = myCode.checkDBNull(dr_src["Post_No"]);
                            foreach(DataRow master_dr in dt_master.Rows)
                            {
                                string temp_dr_val = myCode.checkDBNull(master_dr["Post_No"]).Replace(" ", "").ToUpper();
                                if(temp_dr_val == post_no.ToUpper())
                                {
                                    string tar_col = myCode.checkDBNull(master_dr["ItemName"]);
                                    int tar_col_inx = -1;
                                    if (exp_proc.check_column_exited(dt_process, tar_col, ref tar_col_inx))
                                    {
                                        dr[tar_col_inx] = dr_src["Data"];
                                    }
                                    break;
                                }
                            }
                        }
                        if(myCode.IsNumeric(cur_pcs_no))
                        {
                            int inx = Convert.ToInt32(cur_pcs_no);
                            if (Image_result.ContainsKey(inx))
                            {
                                dr["Image_data"] = Image_result[inx];
                            }
                            if (Image_result_2.ContainsKey(inx))
                            {
                                dr["Image_data_2"] = Image_result_2[inx];
                            }
                        }
                        dr["Remark"] = dt.Rows[0]["Remark"].ToString();
                        dt_process.Rows.Add(dr);
                    }
                }
                DGV_data_process.DataSource = dt_process;
                ((DataGridViewImageColumn)DGV_data_process.Columns["Image_data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)DGV_data_process.Columns["Image_data"]).Width = 100;
                ((DataGridViewImageColumn)DGV_data_process.Columns["Image_data_2"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)DGV_data_process.Columns["Image_data_2"]).Width = 100;

                foreach (DataGridViewRow dr in DGV_data_process.Rows)
                {
                    dr.Height = 70;
                }
            }
        }
        public AutoCompleteStringCollection location_process(myExcel.Worksheet ws)
        {
            AutoCompleteStringCollection list = new AutoCompleteStringCollection();

            for (int j = 1; j < 30; j++)
            {
                for (int i = 1; i < 200; i++)
                {
                    if (ws.Cells[i, j].Value != null)
                    {
                        if (ws.Cells[i, j].Value.ToString() == "S/N")
                        {
                            int count = 1;
                            while (ws.Cells[i + count, j].Value != null)
                            {
                                count++;
                            }
                            list.Add(i.ToString() + "+" + j.ToString() + "+" + (count - 1).ToString());
                        }
                    }
                }
            }
            return list;
        }
        public void Get_Image_Multi(string in_src, ref SortedDictionary<int, byte[]> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            try
            {

                FileInfo[] temp_lst = tar_d.GetFiles("*.jpg");
                if (temp_lst.Length > 0)
                {
                    for (int i = 0; i < temp_lst.Length; i++)
                    {
                        var sel_img = Bitmap.FromFile(temp_lst[i].FullName);
                        ImageConverter imgcon = new ImageConverter();
                        byte[] img_data = (byte[])imgcon.ConvertTo(sel_img, typeof(byte[]));
                        string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
                        if (f_na.Contains("-") == false && f_na.Contains("+") == false)
                        {

                            if (myCode.IsNumeric(f_na))
                            {
                                int f_inx = Convert.ToInt32(f_na);
                                lst_result.Add(f_inx, img_data);
                            }
                        }
                    }

                }
                else
                {
                    DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                    if (inter_type_lst.Length > 0)
                    {
                        foreach (var inter_lst in inter_type_lst)
                        {
                            Get_Image_Multi(inter_lst.FullName, ref lst_result);
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("Could not find a part of the path " + in_src, "Warning");
            }
        }
        public void Get_Image_Multi_2(string in_src, ref SortedDictionary<int, byte[]> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            try
            { 
                FileInfo[] temp_lst = tar_d.GetFiles("*.jpg");
                if (temp_lst.Length > 0)
                {
                    for (int i = 0; i < temp_lst.Length; i++)
                    {
                        var sel_img = Bitmap.FromFile(temp_lst[i].FullName);
                        ImageConverter imgcon = new ImageConverter();
                        byte[] img_data = (byte[])imgcon.ConvertTo(sel_img, typeof(byte[]));
                        string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' }).ToUpper();
                        if (f_na.Contains("S"))
                        {
                            f_na = f_na.Replace("S", string.Empty);
                            if (myCode.IsNumeric(f_na))
                            {
                                int f_inx = Convert.ToInt32(f_na);
                                lst_result.Add(f_inx, img_data);
                            }
                        }
                    } 
                }
                else
                {
                    DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                    if (inter_type_lst.Length > 0)
                    {
                        foreach (var inter_lst in inter_type_lst)
                        {
                            Get_Image_Multi_2(inter_lst.FullName, ref lst_result);
                        }
                    }
                } 
            }
            catch
            {
                MessageBox.Show("Could not find a part of the path " + in_src, "Warning");
            }
        }
        public void Get_Image_Multi_insert(string in_src, ref SortedDictionary<int, byte[]> lst_result, int log_inx)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            try
            {
                FileInfo[] temp_lst = tar_d.GetFiles("*.jpg");
                if (temp_lst.Length > 0)
                {
                    for (int i = 0; i < temp_lst.Length; i++)
                    {
                        var sel_img = Bitmap.FromFile(temp_lst[i].FullName);
                        ImageConverter imgcon = new ImageConverter();
                        byte[] img_data = (byte[])imgcon.ConvertTo(sel_img, typeof(byte[]));
                        string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
                        if (f_na.Contains("-") == false && f_na.Contains("+") == false)
                        {

                            if (myCode.IsNumeric(f_na))
                            {
                                int f_inx = Convert.ToInt32(f_na) + log_inx;
                                if (!lst_result.ContainsKey(f_inx))
                                {
                                    lst_result.Add(f_inx, img_data);
                                }

                            }
                        }
                    }
                }
                else
                {
                    DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                    if (inter_type_lst.Length > 0)
                    {
                        foreach (var inter_lst in inter_type_lst)
                        {
                            Get_Image_Multi_insert(inter_lst.FullName, ref lst_result, log_inx);
                        }
                    }
                }
            }
            catch
            {

            }

        }
        public void Get_Image_Multi_insert_2(string in_src, ref SortedDictionary<int, byte[]> lst_result, int log_inx)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            try
            {


                FileInfo[] temp_lst = tar_d.GetFiles("*.jpg");
                if (temp_lst.Length > 0)
                {
                    for (int i = 0; i < temp_lst.Length; i++)
                    {
                        var sel_img = Bitmap.FromFile(temp_lst[i].FullName);
                        ImageConverter imgcon = new ImageConverter();
                        byte[] img_data = (byte[])imgcon.ConvertTo(sel_img, typeof(byte[]));
                        string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });

                        if (f_na.Contains("-"))
                        {
                            f_na = f_na.Replace("-", string.Empty);
                            if (myCode.IsNumeric(f_na))
                            {
                                int f_inx = Convert.ToInt32(f_na) + log_inx;
                                if (!lst_result.ContainsKey(f_inx))
                                {
                                    lst_result.Add(f_inx, img_data);
                                }
                            }
                        }
                    }
                }
                else
                {
                    DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                    if (inter_type_lst.Length > 0)
                    {
                        foreach (var inter_lst in inter_type_lst)
                        {
                            Get_Image_Multi_insert_2(inter_lst.FullName, ref lst_result, log_inx);
                        }
                    }
                }
            }
            catch
            {
                //MessageBox.Show("Could not find a part of the path " + in_src, "Warning");
            }
        } 

        public void BVH_PTH_Data_Process(string src_path)
        {
            //try
            //{
            if (src_path != "")
            {
                Load_spec();
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { txtItemCode.Text, txtLotNo.Text, cb_process.SelectedItem.ToString() });
                //SortedDictionary<int, Dictionary<int, string>> logfile_result = new SortedDictionary<int, Dictionary<int, string>>();
                SortedDictionary<int, Dictionary<string, string>> logfile_result = new SortedDictionary<int, Dictionary<string, string>>();
                SortedDictionary<int, int> lst_Doday_count = new SortedDictionary<int, int>();
                Get_logfile_Multi_update20_2(src_path, ref logfile_result);
                Image_result = new SortedDictionary<int, byte[]>();
                Image_result_2 = new SortedDictionary<int, byte[]>();
                DirectoryInfo tar_d = new DirectoryInfo(src_path);
                DirectoryInfo[] arr_dic = tar_d.GetDirectories();

                foreach (DirectoryInfo d in arr_dic)
                {
                    if (d.Name.ToUpper().Replace(" ", "").Contains("DK"))
                    { 
                        Get_Image_Multi( Path.Combine( src_path, d.Name ), ref Image_result);
                        //Get_Image_Multi_2(Path.Combine(src_path, d.Name), ref Image_result_2);
                    }
                    if (d.Name.ToUpper().Replace(" ", "").Contains("DODAY"))
                    {
                        //Get_Image_Multi(Path.Combine(src_path, d.Name), ref Image_result);
                        Get_Image_Multi_2(Path.Combine(src_path, d.Name), ref Image_result_2);
                    }
                }
                DataTable dt_data = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_LOGFILE", filter_str).Clone();

                int log_inx = 0;
                foreach (var log in logfile_result)
                {
                    foreach (var log_val in log.Value)
                    {
                        dt_data.Rows.Add((log_inx + 1), txtItemCode.Text, txtLotNo.Text, cb_process.SelectedItem.ToString(), log.Key.ToString(), log_val.Key.ToString(), log_val.Value, src_path);
                        log_inx++;
                    }
                }

                DGV_logfile.DataSource = dt_data;
                insert_data_2(cb_process.SelectedItem.ToString().ToUpper(), Image_result, Image_result_2, dt_data);
                lbl_title.Text = "Data Analysis For " + cb_process.SelectedItem.ToString().ToUpper();

                DataTable dt_spec = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SPEC_" + cb_process.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));

                int count_sample = int.Parse(dt_spec.Rows[0]["Remark"].ToString().Split('+')[2]);
                if (DGV_data_process.Rows.Count < count_sample)
                {
                    btn_insert.Enabled = true;
                }
                else
                {
                    btn_insert.Enabled = false;
                }
                Check_spec();

            }
            //}
            //catch
            //{ 
            //} 
        }
        public void BVH_PTH_Data_Update_new()
        {
            if (edited_lst.Count > 0)
            { 
                DataTable main_tbl = TDMK_Code.Datatable_Filter(sqlcon, cb_process.Text, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                DataTable logfile_tbl = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                // DataTable BVH_masterlist = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_MASTERLIST", TDMK_Code.filter_str(new string[] { "ItemCode", "Process" }, new string[] { txtItemCode.Text, cb_process.Text }));
                DataTable BVH_masterlist = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_MASTERLIST", TDMK_Code.filter_str(new string[] { "ItemCode", "Process" }, new string[] { txtItemCode.Text, cb_process.Text }));
                if (BVH_masterlist.Rows.Count == 0)
                {

                    BVH_masterlist = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_MASTERLIST", TDMK_Code.filter_str(new string[] { "ItemCode", "Process" }, new string[] { "22B0500", cb_process.Text }));
                }
                List<string> ignored_col = new List<string>() { "ID", "ItemCode", "LotNo", "PCS_No", "Image_data", "Image_data_2", "Remark" };
                int id = logfile_tbl.AsEnumerable().Select(x => x.Field<int>("ID")).ToList().Max();
                List<string> item_arr = new List<string>() { "ID", "ItemCode", "LotNo", "Process", "PCS_No", "Region", "Zone", "Before_Data", "After_Data", "Time_Update", "Operator", "Depart", "Remark" };
                DateTime tim_up = DateTime.Now;
                int id_edit = TDMK_Code.SQL_MAX("VHX_Edit_History", "ID", sqlcon);
                int count = 0;
                foreach (var item in edited_lst)
                {
                    string pcs_no = DGV_data_process.Rows[item.Key].Cells["PCS_No"].Value.ToString();
                    string remark = DGV_data_process.Rows[item.Key].Cells["Remark"].Value.ToString();
                    string post_no;
                    DataView temp = BVH_masterlist.AsDataView();

                    List<object> item_arr_val = new List<object>();
                    foreach (var itemname in item.Value)
                    {
                        if (ignored_col.IndexOf(itemname) == -1)
                        {
                            temp.RowFilter = "ItemName = '" + itemname + "'";
                            int index = BVH_masterlist.Rows.IndexOf(temp[0].Row);
                            if (temp.Count > 0)
                            {
                                post_no = temp[0]["Post_No"].ToString();
                                string replace_data = DGV_data_process.Rows[item.Key].Cells[itemname].Value.ToString();
                                logfile_tbl.Rows.Add((id + count + 1), txtItemCode.Text, txtLotNo.Text, cb_process.Text, pcs_no, post_no, replace_data, remark);
                                string before_data = main_tbl.Rows[item.Key][itemname].ToString();
                                byte[] bef_data = System.Text.Encoding.UTF8.GetBytes(before_data);
                                byte[] aft_data = System.Text.Encoding.UTF8.GetBytes(replace_data);
                                item_arr_val = new List<object>() { (id_edit + count + 1), txtItemCode.Text, txtLotNo.Text, cb_process.Text, pcs_no, itemname, string.Empty, bef_data, aft_data, tim_up, txtUsername.Text, depart, remark };
                                exp_proc.Insert_Object_List("VHX_Edit_History", sqlcon, item_arr, item_arr_val);
                            }
                            count++;
                        }
                    }
                    try
                    {
                        byte[]  img_bef = (byte[])(main_tbl.Rows[item.Key]["Image_data"]);
                        byte[] img_aft = (byte[])(DGV_data_process.Rows[item.Key].Cells["Image_data"].Value);
                        item_arr_val = new List<object>() { (id_edit + count + 1), txtItemCode.Text, txtLotNo.Text, cb_process.Text, pcs_no, "Image_data", string.Empty, img_bef, img_aft, tim_up, txtUsername.Text, depart, remark };
                        exp_proc.Insert_Object_List("VHX_Edit_History", sqlcon, item_arr, item_arr_val);
                    }
                    catch
                    {
                    }

                    try
                    {
                        byte[] img_bef_2 = (byte[])(main_tbl.Rows[item.Key]["Image_data_2"]);
                        byte[] img_aft_2 = (byte[])(DGV_data_process.Rows[item.Key].Cells["Image_data_2"].Value);
                        item_arr_val = new List<object>() { (id_edit + count + 2), txtItemCode.Text, txtLotNo.Text, cb_process.Text, pcs_no, "Image_data", string.Empty, img_bef_2, img_aft_2, tim_up, txtUsername.Text, depart, remark };
                        exp_proc.Insert_Object_List("VHX_Edit_History", sqlcon, item_arr, item_arr_val);
                    }
                    catch
                    {

                    }
                }
                DGV_logfile.DataSource = logfile_tbl;
                save_data();

                edited_lst = new SortedDictionary<int, List<string>> { };

            }
        }

        public void BVH_PTH_Data_Update(string src_path)
        {

            if (src_path != "")
            {
                Load_spec();

                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process", "Remark" }, new string[] { txtItemCode.Text, txtLotNo.Text, cb_process.SelectedItem.ToString(), "" });
                //DGV_logfile.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_LOGFILE", filter_str);
                //DGV_data_process.DataSource = TDMK_Code.Datatable_Filter(sqlcon, cb_process.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Remark" }, new string[] { txtItemCode.Text, txtLotNo.Text, "" }));

                SortedDictionary<int, Dictionary<int, string>> logfile_result = new SortedDictionary<int, Dictionary<int, string>>();
                Imp_code.Get_logfile_Multi(src_path, ref logfile_result);
                Image_result = new SortedDictionary<int, byte[]>();
                Get_Image_Multi(src_path, ref Image_result);

                DataTable dt_data = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_LOGFILE", filter_str);
                string[] remark_all = dt_data.AsEnumerable().Select(x => x.Field<string>("Remark")).Distinct().ToArray();

                int i = 0;

                foreach (var log in logfile_result)
                {
                    while (i < dt_data.Rows.Count)
                    {

                        if (dt_data.Rows[i]["PCS_No"].ToString() == log.Key.ToString())
                        {

                            foreach (var log_val in log.Value)
                            {

                                DataRow dr = dt_data.NewRow();
                                int count_row = dt_data.Rows.Count;

                                dr[0] = count_row;
                                for (int j = 1; j < dt_data.Columns.Count - 1; j++)
                                {
                                    dr[j] = dt_data.Rows[i][j];
                                }
                                dr[7] = "NG_F" + remark_all.Length.ToString();
                                dt_data.Rows.Add(dr);

                                dt_data.Rows[i][5] = log_val.Key.ToString();
                                dt_data.Rows[i][6] = log_val.Value;
                                dt_data.Rows[i][7] = "";
                                i++;

                            }
                            break;
                        }
                        else
                        {
                            i++;
                        }
                    }



                }



                //dt_data = (DataTable)DGV_logfile.DataSource;
                DGV_logfile.DataSource = dt_data;

                insert_data_update(cb_process.SelectedItem.ToString().ToUpper(), Image_result, dt_data);
                lbl_title.Text = "Data Analysis For " + cb_process.SelectedItem.ToString().ToUpper();


                Check_spec();

            }

            //}
            //catch
            //{

            //}


        }
        public void Insert_logfile(string src_path, SortedDictionary<int, byte[]> Image_result, SortedDictionary<int, byte[]> Image_result_2)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { txtItemCode.Text, txtLotNo.Text, cb_process.SelectedItem.ToString() });
            SortedDictionary<int, Dictionary<int, string>> logfile_result = new SortedDictionary<int, Dictionary<int, string>>();
            //Imp_code.Get_logfile_Multi(src_path, ref logfile_result);
            SortedDictionary<int, int> lst_Doday_count = new SortedDictionary<int, int>();
            Get_logfile_Multi_update20_2(src_path, ref logfile_result,ref lst_Doday_count);
            DataTable dt_data = (DataTable)DGV_logfile.DataSource;

            int log_inx = DGV_data_process.Rows.Count;

            //Get_Image_Multi_insert(src_path, ref Image_result, log_inx);
            //Get_Image_Multi_insert_2(src_path, ref Image_result_2, log_inx);

            DirectoryInfo tar_d = new DirectoryInfo(src_path);
            DirectoryInfo[] arr_dic = tar_d.GetDirectories();

            foreach (DirectoryInfo d in arr_dic)
            {
                if (d.Name.ToUpper().Replace(" ", "").Contains("DODAY"))
                { 
                    Get_Image_Multi_insert(Path.Combine(src_path, d.Name), ref Image_result, log_inx);
                    Get_Image_Multi_insert_2(Path.Combine(src_path, d.Name), ref Image_result_2, log_inx); 
                }
            }

            int count_pcs_before = DGV_data_process.Rows.Count;
            foreach (var log in logfile_result)
            {
                foreach (var log_val in log.Value)
                {
                    dt_data.Rows.Add((log_inx + 1), txtItemCode.Text, txtLotNo.Text, cb_process.SelectedItem.ToString(), int.Parse(log.Key.ToString()) + count_pcs_before, log_val.Key.ToString(), log_val.Value, src_path);
                    log_inx++;
                }

            }

            DGV_logfile.DataSource = dt_data;
            insert_data_2(cb_process.SelectedItem.ToString().ToUpper(), Image_result, Image_result_2, dt_data);

            DataTable dt_spec = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SPEC_" + cb_process.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
            int count_sample = int.Parse(dt_spec.Rows[0]["Remark"].ToString().Split('+')[2]);
            if (DGV_data_process.Rows.Count < count_sample)
            {
                btn_insert.Enabled = true;
            }
            else
            {
                btn_insert.Enabled = false;
            }
            Check_spec();
            MessageBox.Show("Insert data completed", "Waring");


        }


        public void Load_spec_2()
        {
            string[] item = { "ItemCode" };
            string[] item_val = { txtItemCode.Text };
            DGV_spec.DataSource = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SPEC_" + cb_process.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(item, item_val));

            if (DGV_spec.Rows.Count == 0)
            {
                MessageBox.Show(new Form { TopMost = true }, "Please wait for setup spec", "Warning");
                string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                if (format_file != "")
                {
                    string[] spec = { "USL", "LSL" };

                    //myExcel.Worksheet ws = TDMK_Code.open_excel_file(format_file, "", "").Sheets["BVH&PTH"];

                    myExcel.Workbook wb = TDMK_Code.open_excel_file(format_file, "", "");
                    string _process = "BVH&PTH".Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
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
                        myExcel.Worksheet ws = wb.Sheets[mySheet];
                        //ws.Activate();
                        AutoCompleteStringCollection list = new AutoCompleteStringCollection();

                        //  string spec_ = setup_spec_BVH_PTH(ws, dt_spec.Columns.Count - 3,ref list).Split('+')[cb_process.SelectedIndex];
                        List<string[]> arr_process = new List<string[]> { arr_item_BVH_with, arr_item_BVH_without, arr_item_PTH };
                        string[] setup_spec = setup_spec_BVH_PTH_Update(ws, ref list).Split('+');
                        int tt = 0;
                        foreach (string lst in list)
                        {
                            string spec_ = setup_spec[tt];
                            tt++;
                            int no_proc = -1;
                            if (lst.Contains("B1"))
                                no_proc = 0;
                            if (lst.Contains("B2"))
                                no_proc = 1;
                            if (lst.Contains("P1"))
                                no_proc = 2;


                            if (no_proc != -1)
                            {
                                int t = 0;

                                DataTable dt_spec = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SPEC_" + cb_process.Items[no_proc].ToString().ToUpper(), TDMK_Code.filter_str(item, item_val));
                                foreach (string str in spec)
                                {
                                    DataRow dr = dt_spec.NewRow();
                                    dr["ID"] = TDMK_Code.SQL_MAX("SPEC_" + cb_process.Items[no_proc].ToString().ToUpper(), "ID", sqlcon) + 1 + t;
                                    dr["ItemCode"] = txtItemCode.Text;
                                    dr["Spec"] = str;

                                    for (int i = 3; i < dt_spec.Columns.Count - 1; i++)
                                    {
                                        dr[i] = spec_.Split('_')[i - 3].Split(';')[t];
                                    }

                                    string remark = lst;
                                    dr["Remark"] = remark;
                                    dt_spec.Rows.Add(dr);
                                    t++;
                                }
                                exp_proc.BatchBulkCopy(myVar.mysqlcon, dt_spec, "SPEC_" + cb_process.Items[no_proc].ToString().ToUpper());
                            }

                        }
                        wb.Close();
                        DGV_spec.DataSource = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SPEC_" + cb_process.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(item, item_val));
                        MessageBox.Show(new Form { TopMost = true }, "Setup spec complete", "Warning");
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Sheet BVH&PTH not found", "Warning");
                        wb.Close();
                    }
                }

                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Format of " + txtItemCode.Text + " not found", "Warning");
                }

            }
        }


        public void Load_spec()
        {
            DGV_data_view.DataSource = null;
            string[] item = { "ItemCode" };
            string[] item_val = { txtItemCode.Text };
            DGV_spec.DataSource = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SPEC_" + cb_process.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(item, item_val));

            if (DGV_spec.Rows.Count == 0)
            {
                foreach (string i in cb_process.Items)
                {
                    TDMK_Code.Delelte_FilteredItem_arr("SPEC_" + i.ToUpper(), sqlcon, TDMK_Code.filter_str(item, item_val));
                }

                MessageBox.Show(new Form { TopMost = true }, "Please wait for setup spec", "Warning");
                try
                {
                    string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                    if (format_file != "")
                    {
                        string[] spec = { "USL", "LSL" };

                        myExcel.Workbook wb = TDMK_Code.open_excel_file(format_file, "", "");
                        string _process = "BVH&PTH".Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
                        string mySheet = "";
                        foreach (myExcel.Worksheet tg_sht in wb.Worksheets)
                        {
                            string cur_sht_name = tg_sht.Name.ToUpper().Replace("-", "").Replace(" ", "");
                            if (cur_sht_name.Contains(_process))
                            {
                                mySheet = tg_sht.Name;
                                break;
                            }
                        }
                        if (mySheet != "")
                        {
                            myExcel.Worksheet ws = wb.Sheets[mySheet];
                            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
                            int no_process = 0;
                            List<string[]> arr_process = new List<string[]> { arr_item_BVH_with, arr_item_BVH_without, arr_item_PTH };
                            string[] setup_spec = setup_spec_BVH_PTH(ws, arr_process[no_process].Length - 5, ref list).Split('+');
                            int no_list = 0;
                            foreach (string spec_ in setup_spec)
                            {
                                if (setup_spec.Length == 3)
                                {
                                    if (no_process < 3)
                                    {
                                        int t = 0;

                                        DataTable dt_spec = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SPEC_" + cb_process.Items[no_process].ToString().ToUpper(), TDMK_Code.filter_str(item, item_val));
                                        foreach (string str in spec)
                                        {
                                            DataRow dr = dt_spec.NewRow();
                                            dr["ID"] = TDMK_Code.SQL_MAX("SPEC_" + cb_process.Items[no_process].ToString().ToUpper(), "ID", sqlcon) + 1 + t;
                                            dr["ItemCode"] = txtItemCode.Text;
                                            dr["Spec"] = str;

                                            for (int i = 3; i < dt_spec.Columns.Count - 1; i++)
                                            {
                                                dr[i] = spec_.Split('_')[i - 3].Split(';')[t];
                                            }

                                            string remark = list[no_list];
                                            dr["Remark"] = remark;
                                            dt_spec.Rows.Add(dr);
                                            t++;
                                        }
                                        no_list++;

                                        exp_proc.BatchBulkCopy(myVar.mysqlcon, dt_spec, "SPEC_" + cb_process.Items[no_process].ToString().ToUpper());
                                        no_process = no_process + 2;
                                    }
                                }
                                else
                                {

                                    if (no_process < 3 && spec_ != "")
                                    {
                                        int t = 0;

                                        DataTable dt_spec = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SPEC_" + cb_process.Items[no_process].ToString().ToUpper(), TDMK_Code.filter_str(item, item_val));
                                        foreach (string str in spec)
                                        {
                                            DataRow dr = dt_spec.NewRow();
                                            dr["ID"] = TDMK_Code.SQL_MAX("SPEC_" + cb_process.Items[no_process].ToString().ToUpper(), "ID", sqlcon) + 1 + t;
                                            dr["ItemCode"] = txtItemCode.Text;
                                            dr["Spec"] = str;

                                            for (int i = 3; i < dt_spec.Columns.Count - 1; i++)
                                            {
                                                dr[i] = spec_.Split('_')[i - 3].Split(';')[t];
                                            }

                                            string remark = list[no_process];
                                            dr["Remark"] = remark;
                                            dt_spec.Rows.Add(dr);
                                            t++;
                                        }

                                        exp_proc.BatchBulkCopy(myVar.mysqlcon, dt_spec, "SPEC_" + cb_process.Items[no_process].ToString().ToUpper());
                                        no_process++;
                                    }
                                }


                            }
                            wb.Close();
                            DGV_spec.DataSource = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SPEC_" + cb_process.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(item, item_val));
                            MessageBox.Show(new Form { TopMost = true }, "Setup spec complete", "Warning");
                        }

                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Sheet BVH&PTH not found", "Warning");
                            wb.Close();
                        }
                    }

                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Format of " + txtItemCode.Text + " not found", "Warning");
                    }

                }
                catch
                {

                }

            }
        }
        public void change_ID_DGV(DataGridView dgv)
        {
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                dgv.Rows[i].Cells["ID"].Value = i + 1;
            }

        }


        private void BVH_PTH_Update_Load(object sender, EventArgs e)
        {
            sqlcon = myVar.mysqlcon;

            cbProcess_MasterList.Items.Add("BVH_with_Bonding_Sheet");
            cbProcess_MasterList.Items.Add("BVH_without_Bonding_Sheet");
            cbProcess_MasterList.Items.Add("Plated_Through_Hole");



            arr_item_BVH_with[0] = "ID";
            arr_item_BVH_with[1] = "ItemCode";
            arr_item_BVH_with[2] = "LotNo";
            arr_item_BVH_with[3] = "PCS_No";
            arr_item_BVH_with[4] = "Plated_Cu_thickness_from_top_Base_Cu";
            arr_item_BVH_with[5] = "Plated_Cu_thickness_from_PI";
            arr_item_BVH_with[6] = "Plated_Cu_thickness_from_Adhesive";
            arr_item_BVH_with[7] = "Plated_Cu_thickness_from_bottom_Base_Cu";
            arr_item_BVH_with[8] = "Etched_down_base_Cu_thickness";
            arr_item_BVH_with[9] = "Minimum_plated_Cu_thickness_from_top_base_Cu";
            arr_item_BVH_with[10] = "Adhesive_etch_back";
            arr_item_BVH_with[11] = "Top_diameter";
            arr_item_BVH_with[12] = "Image_data";




            arr_item_BVH_without[0] = "ID";
            arr_item_BVH_without[1] = "ItemCode";
            arr_item_BVH_without[2] = "LotNo";
            arr_item_BVH_without[3] = "PCS_No";
            arr_item_BVH_without[4] = "Plated_Cu_thickness_from_top_Base_Cu";
            arr_item_BVH_without[5] = "Plated_Cu_thickness_from_PI";
            arr_item_BVH_without[6] = "Plated_Cu_thickness_from_Adhesive";
            arr_item_BVH_without[7] = "Plated_Cu_thickness_from_bottom_Base_Cu";
            arr_item_BVH_without[8] = "Etched_down_base_Cu_thickness";
            arr_item_BVH_without[9] = "Minimum_plated_Cu_thickness_from_top_base_Cu";
            arr_item_BVH_without[10] = "PI_etch_back";
            arr_item_BVH_without[11] = "Top_diameter";
            arr_item_BVH_without[12] = "Image_data";



            arr_item_PTH[0] = "ID";
            arr_item_PTH[1] = "ItemCode";
            arr_item_PTH[2] = "LotNo";
            arr_item_PTH[3] = "PCS_No";
            arr_item_PTH[4] = "Corner_thickness";
            arr_item_PTH[5] = "Side_wall_from_PI";
            arr_item_PTH[6] = "Side_wall_from_adhesive";
            arr_item_PTH[7] = "Side_wall_Nickel";
            arr_item_PTH[8] = "Min_Plate";
            arr_item_PTH[9] = "Etch_back";
            arr_item_PTH[10] = "Diameter";
            arr_item_PTH[11] = "Image_data";
            btn_insert.Enabled = false;
        }

        private void tsmExport_Click_1(object sender, EventArgs e)
        { 
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtOperator.Text != "")
            {
               // Export_BVH_PTH_data(sqlcon, txtItemCode.Text, txtLotNo.Text);
                string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                if (format_file != "")
                {
                    string export_path = F_exportEPPlus.get_reportpath(txtItemCode.Text, txtLotNo.Text, txtOperator.Text, "BVH-PTH", myVar.data_loc);
                    if (export_path != "")
                    {
                        F_exportEPPlus.Export_EPPlus(sqlcon, format_file, export_path, txtItemCode.Text, txtLotNo.Text, "BVH-PTH", null);
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Folder BVH-PTH not found!", "Warning");
                    }
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Please, fill in ItemCode / LotNo ", "Warning");
            } 
        } 
        public Boolean Check_spec_before_export(DataTable tbl_data, DataTable dt_spec)
        {
            bool chk = true;
            string[] item = { "ItemCode" };
            string[] item_val = { txtItemCode.Text};
            // DataTable dt_spec = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SPEC_" + cb_process.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(item, item_val)); 

            if (dt_spec.Rows.Count > 0)
            {
                for (int i = 0; i < tbl_data.Rows.Count; i++)
                {
                    for (int k = 4; k < 12; k++)
                    {
                        string USL = dt_spec.Rows[0][k - 1].ToString().Replace("/", string.Empty);
                        string LSL = dt_spec.Rows[1][k - 1].ToString().Replace("/", string.Empty);
                        if (tbl_data.Rows[i][k] != null)
                        {
                            if (!USL.Contains("NA"))
                            {
                                if (Double.TryParse(tbl_data.Rows[i][k].ToString(), out Double data))
                                {
                                    if (data > Double.Parse(USL))
                                    {
                                        chk = false;
                                        break;
                                    }
                                } 
                            }
                            if (!LSL.Contains("NA"))
                            { 
                                if (Double.TryParse(tbl_data.Rows[i][k].ToString(), out Double data))
                                {
                                    if (data < Double.Parse(LSL))
                                    {
                                        chk = false;
                                        break;
                                    }
                                } 
                            }
                        } 
                    }
                    if (!chk)
                        break;
                }
            }
            return chk; 
        }

        public void Check_spec()
        {
            for (int i = 0; i < DGV_data_process.Columns.Count; i++)
            {
                for (int j = 0; j < DGV_data_process.Rows.Count; j++)
                {
                    DGV_data_process.Rows[j].Cells[i].Style.BackColor = Color.White;
                }
            }  
            string[] item = { "ItemCode" };
            string[] item_val = { txtItemCode.Text };
            DataTable dt_spec = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SPEC_" + cb_process.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(item, item_val));

            if (dt_spec.Rows.Count > 0)
            {
                for (int i = 0; i < DGV_data_process.Rows.Count; i++)
                {
                    for (int k = 4; k < 12; k++)
                    {
                        string USL = dt_spec.Rows[0][k - 1].ToString().Replace("/", string.Empty);
                        string LSL = dt_spec.Rows[1][k - 1].ToString().Replace("/", string.Empty);
                        if (DGV_data_process.Rows[i].Cells[k].Value != null)
                        {
                            if (!USL.Contains("NA"))
                            {
                                if (Double.TryParse(DGV_data_process.Rows[i].Cells[k].Value.ToString(), out Double data))
                                {
                                    if (data > Double.Parse(USL))
                                    {
                                        DGV_data_process.Rows[i].Cells[k].Style.BackColor = Color.Red;
                                    }
                                }

                            }
                            if (!LSL.Contains("NA"))
                            {

                                if (Double.TryParse(DGV_data_process.Rows[i].Cells[k].Value.ToString(), out Double data))
                                {
                                    if (data < Double.Parse(LSL))
                                    {
                                        DGV_data_process.Rows[i].Cells[k].Style.BackColor = Color.Red;
                                    }
                                }

                            }
                        }

                    }
                }
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


                    curr_wrkbook.SaveAs(report_path);
                    break;
                }
            }

        }



        public void Export_BVH_PTH_data(SqlConnection tar_sqlcon, string ItemCode, string LotNo)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "")
            {

                string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);

                myExcel.Workbook curr_wrkbook = exp_proc.create_export_wrk(format_file, "BVH & PTH", txtItemCode.Text, txtLotNo.Text);

                List<string> tar_table = new List<string>() { "BVH_WITH_BONDING_SHEET", "BVH_WITHOUT_BONDING_SHEET", "PLATED_THROUGH_HOLE" };
                myExcel.Worksheet curr_wrksheet = curr_wrkbook.Sheets[1];

                List<DataTable> tar_dt = new List<DataTable>();
                for (int i = 0; i < tar_table.Count; i++)
                {
                    tar_dt.Add(TDMK_Code.Datatable_Filter(tar_sqlcon, tar_table[i], TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text })));
                }
                bool exp_en = false;
                foreach (var _dt in tar_dt)
                {
                    if (_dt.Rows.Count > 0)
                    {
                        exp_en = true;
                        break;
                    }
                }
                if (exp_en)
                {

                    foreach (var t in tar_table)
                    {

                        DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, t, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                        DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_" + t, TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
                        bool chk = true;
                        if (!Check_spec_before_export(dt, dt_spec))
                        {
                            chk = false;
                            if (MessageBox.Show("Data of " + t + " is out of spec. Do you want to export to checksheet?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                            {
                                chk = true;
                            }
                        }
                        if (chk)
                        {

                            if (dt_spec.Rows.Count > 0)
                            {
                                string list_val = dt_spec.Rows[0]["Remark"].ToString();
                                switch (t)
                                {
                                    case "BVH_WITH_BONDING_SHEET":
                                        DB_to_Excel(dt, curr_wrksheet, "BVH_WITH_BONDING_SHEET", 8, list_val);
                                        break;
                                    case "BVH_WITHOUT_BONDING_SHEET":
                                        DB_to_Excel(dt, curr_wrksheet, "BVH_WITHOUT_BONDING_SHEET", 8, list_val);
                                        break;
                                    case "PLATED_THROUGH_HOLE":
                                        DB_to_Excel(dt, curr_wrksheet, "PLATED_THROUGH_HOLE", 7, list_val);
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }


                    }

                    save_report_new(txtItemCode.Text, txtLotNo.Text, txtOperator.Text, curr_wrkbook, "BVH & PTH", myVar.data_loc);
                    MessageBox.Show(new Form { TopMost = true }, "Export to Checksheet completed!", "Warning");
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "No data. Can not export to Checksheet", "Warning");
                }

            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Please, fill in ItemCode / LotNo ", "Warning");
            }

        }
        //public void Export_BVH_PTH_data_Update(SqlConnection tar_sqlcon, string ItemCode, string LotNo)
        //{
        //    if (txtItemCode.Text != "" && txtLotNo.Text != "")
        //    {

        //        string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
        //        myExcel.Workbook curr_wrkbook = exp_proc.create_export_wrk(format_file, "BVH & PTH", txtItemCode.Text, txtLotNo.Text);

        //        List<string> tar_table = new List<string>() { "BVH_WITH_BONDING_SHEET", "BVH_WITHOUT_BONDING_SHEET", "PLATED_THROUGH_HOLE" };
        //        myExcel.Worksheet curr_wrksheet = curr_wrkbook.Sheets[1];


        //        foreach (var t in tar_table)
        //        {
        //            DataTable dt_all = TDMK_Code.Datatable_Filter(sqlcon, t, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo"}, new string[] { txtItemCode.Text, txtLotNo.Text }));
        //            List<DataTable> src_remark_tbl = new List<DataTable>();
        //            exp_proc.Get_ListTable(-1, dt_all, new string[] { "ItemCode", "LotNo", "Remark" }, ref src_remark_tbl, "PCS_No");
        //            bool chk_OK_time = false;
        //            foreach (var src_dt in src_remark_tbl)
        //            {
        //                string str_remark = src_dt.Rows[0]["Remark"].ToString();
        //                if (str_remark.Contains("-OK"))
        //                {
        //                    chk_OK_time = true;
        //                    DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_" + t, TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
        //                    if (dt_spec.Rows.Count > 0)
        //                    {
        //                        string list_val = dt_spec.Rows[0]["Remark"].ToString();
        //                        switch (t)
        //                        {
        //                            case "BVH_WITH_BONDING_SHEET":
        //                                DB_to_Excel(src_dt, curr_wrksheet, "BVH_WITH_BONDING_SHEET", 8, list_val);
        //                                break;
        //                            case "BVH_WITHOUT_BONDING_SHEET":
        //                                DB_to_Excel(src_dt, curr_wrksheet, "BVH_WITHOUT_BONDING_SHEET", 8, list_val);
        //                                break;
        //                            case "PLATED_THROUGH_HOLE":
        //                                DB_to_Excel(src_dt, curr_wrksheet, "PLATED_THROUGH_HOLE", 7, list_val);
        //                                break;
        //                            default:
        //                                break;
        //                        }
        //                    }
        //                    break;
        //                }

        //            }
        //            if (!chk_OK_time)
        //            {
        //                foreach (var src_dt in src_remark_tbl)
        //                {

        //                    string str_remark = src_dt.Rows[0]["Remark"].ToString();
        //                    if (str_remark != "")
        //                    {
        //                        string v_time = str_remark.Split('\\')[str_remark.Split('\\').Length - 2].Split('(').Last().Split(')')[0];
        //                        if (!myCode.IsNumeric(v_time))
        //                            v_time = "1";
        //                        if (v_time == chk_list_vtime.CheckedItems[0].ToString())
        //                        {
        //                            DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_" + t, TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
        //                            if (dt_spec.Rows.Count > 0)
        //                            {
        //                                string list_val = dt_spec.Rows[0]["Remark"].ToString();
        //                                switch (t)
        //                                {
        //                                    case "BVH_WITH_BONDING_SHEET":
        //                                        DB_to_Excel(src_dt, curr_wrksheet, "BVH_WITH_BONDING_SHEET", 8, list_val);
        //                                        break;
        //                                    case "BVH_WITHOUT_BONDING_SHEET":
        //                                        DB_to_Excel(src_dt, curr_wrksheet, "BVH_WITHOUT_BONDING_SHEET", 8, list_val);
        //                                        break;
        //                                    case "PLATED_THROUGH_HOLE":
        //                                        DB_to_Excel(src_dt, curr_wrksheet, "PLATED_THROUGH_HOLE", 7, list_val);
        //                                        break;
        //                                    default:
        //                                        break;
        //                                }
        //                            }


        //                            DataTable dt_copy = src_dt.Clone();
        //                            for(int j = 0; j< src_dt.Rows.Count; j++)
        //                            {
        //                                DataRow dr = dt_copy.NewRow();
        //                                for (int i = 0; i < dt_copy.Columns.Count-1; i++)
        //                                {
        //                                    dr[i] = src_dt.Rows[j][i];
        //                                }
        //                                dr["Remark"] = src_dt.Rows[j]["Remark"].ToString() + "-OK";
        //                                dt_copy.Rows.Add(dr);
        //                            }

        //                            exp_proc.BatchBulkCopy(sqlcon, dt_copy, t);
        //                            break;
        //                        }
        //                    }

        //                }
        //            }
        //        }
        //        MessageBox.Show(new Form { TopMost = true }, "Export to Checksheet completed!", "Warning");
        //        //}

        //    }
        //    else
        //    {
        //        MessageBox.Show(new Form { TopMost = true }, "Please, fill in ItemCode / LotNo ", "Warning");
        //    }

        //}
        public void update_item_val_filter(string tbl_name, SqlConnection my_conn, string filter_val, string updated_item_val)
        {

            using (SqlCommand command = my_conn.CreateCommand())
            {
                command.CommandText = "UPDATE " + tbl_name + " SET Remark = @Remark WHERE " + filter_val;
                command.Parameters.AddWithValue("@Remark", updated_item_val);

                my_conn.Open();
                command.ExecuteNonQuery();
                my_conn.Close();
            }

        }
        public List<string> data_NG()
        {
            List<string> list = new List<string> { };
            string[] item = { "ItemCode" };
            string[] item_val = { txtItemCode.Text };
            DataTable dt_spec = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SPEC_" + cb_process.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(item, item_val));
            DataTable tbl_data = TDMK_Code.Datatable_Filter(sqlcon, cb_process.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Remark" }, new string[] { txtItemCode.Text, txtLotNo.Text, "" }));
            if (dt_spec.Rows.Count > 0)
            {
                for (int i = 0; i < tbl_data.Rows.Count; i++)
                {
                    for (int k = 4; k < 12; k++)
                    {
                        string USL = dt_spec.Rows[0][k - 1].ToString().Replace("/", string.Empty);
                        string LSL = dt_spec.Rows[1][k - 1].ToString().Replace("/", string.Empty);
                        if (tbl_data.Rows[i][k] != null)
                        {
                            if (!USL.Contains("NA"))
                            {
                                if (Double.TryParse(tbl_data.Rows[i][k].ToString(), out Double data))
                                {
                                    if (data > Double.Parse(USL))
                                    {
                                        list.Add(tbl_data.Rows[i]["PCS_No"].ToString());
                                        break;
                                    }
                                }

                            }
                            else if (!LSL.Contains("NA"))
                            {

                                if (Double.TryParse(tbl_data.Rows[i][k].ToString(), out Double data))
                                {
                                    if (data < Double.Parse(LSL))
                                    {
                                        list.Add(tbl_data.Rows[i]["PCS_No"].ToString());
                                        break;
                                    }
                                }

                            }
                        }

                    }

                }
            }
            return list;
        }

        public List<string> Load_for_checklist(string tbl_name)
        {
            List<string> lst_time = new List<string> { };
            DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, tbl_name, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));

            string[] remark_all = dt.AsEnumerable().Select(x => x.Field<string>("Remark")).Distinct().ToArray();
            foreach (var str_remark in remark_all)
            {
                //string v_time = str_remark.Split('\\')[str_remark.Split('\\').Length - 2].Split('(').Last().Split(')')[0];
                DirectoryInfo temp_dir = new DirectoryInfo(str_remark);
                string temp_folder = temp_dir.Parent.Name;
                string v_time = temp_folder.Split('(').Last().Split(')')[0];
                if (!myCode.IsNumeric(v_time))
                {
                    lst_time.Add("1");
                }
                else
                {
                    lst_time.Add(v_time);
                }

            }
            return lst_time;
        }
        public void InsertPicture_Name(myExcel.Worksheet tar_wrksht, myExcel.Range tar_range, string picFile, int margin, string pic_name)
        {
            myExcel.Shape sel_picture = tar_wrksht.Shapes.AddPicture(picFile, MsoTriState.msoFalse, MsoTriState.msoTrue, tar_range.Left + margin, tar_range.Top + margin, tar_range.Width - 2 * margin, tar_range.Height - 2 * margin);
            sel_picture.LockAspectRatio = MsoTriState.msoTrue;
            sel_picture.Placement = myExcel.XlPlacement.xlMoveAndSize;
            sel_picture.Name = pic_name;


        }
        public void DB_to_Excel(DataTable src_DGV, myExcel.Worksheet curr_wrksheet, string process, int offset, string list_val)
        {
            string image_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = Path.Combine(image_folder, "1.jpg");


            int data_col = 8;

            int r = int.Parse(list_val.Split('+')[0]);
            int c = int.Parse(list_val.Split('+')[1]);
            int count_sample = int.Parse(list_val.Split('+')[2]);
            myExcel.Range curr_rgn = curr_wrksheet.Cells[r + 1, c + 2];
            myExcel.Range img_rgn = curr_rgn.Offset[0, offset];
            int qty = Math.Min(src_DGV.Rows.Count, count_sample);


            if (process == "PLATED_THROUGH_HOLE")
            {
                data_col = 7;
                curr_rgn = curr_wrksheet.Cells[r + 1, c + 1];
                img_rgn = curr_rgn.Offset[0, offset];
            }


            for (int m = 0; m < qty; m++)
            {
                int offs = int.Parse(src_DGV.Rows[m]["PCS_No"].ToString());

                for (int j = 0; j < data_col; j++)
                {
                    if (src_DGV.Rows[m][j + 4].ToString() != "")
                    {
                        curr_rgn.Offset[offs - 1, j].Value = src_DGV.Rows[m][j + 4].ToString();
                    }
                    else
                    {
                        curr_rgn.Offset[offs - 1, j].Value = "NA";
                    }
                }
            }

            for (int count = 0; count < qty; count++)
            {
                byte[] img_data = (byte[])src_DGV.Rows[count]["Image_data"];
                using (MemoryStream ms = new MemoryStream(img_data))
                {
                    Image image = Image.FromStream(ms);
                    image.Save(file_dic);
                    if (img_rgn.Value != null)
                    {
                        img_rgn.Delete(image);
                    }
                    InsertPicture_Name(curr_wrksheet, img_rgn.MergeArea, file_dic, 5, process + "_" + count.ToString());
                }
                img_rgn = img_rgn.Offset[1, 0];

            }


            //if(count_sample > src_DGV.Rows.Count)
            //{

            //    if (MessageBox.Show("Do you want insert logfile for " + process + "?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
            //    {
            //        check_save_more = true;
            //    }
            //    else
            //    {
            //        check_save_more = false;
            //    }
            //}
            //else
            //{
            //    check_save_more = false;
            //}
        }
        public void save_data_update()
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
            string tbl_process = cb_process.SelectedItem.ToString().ToUpper();
            List<string> tar_table_name = new List<string>() { "BVH_PTH_LOGFILE", tbl_process };
            List<DataTable> tar_table = new List<DataTable>() { (DataTable)DGV_logfile.DataSource, (DataTable)DGV_data_process.DataSource };

            int tbl_inx = 0;
            foreach (var t in tar_table_name)
            {
                TDMK_Code.Delelte_FilteredItem_arr(t, sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                int id = TDMK_Code.SQL_MAX(t, "ID", sqlcon);
                int r_inx = 0;
                foreach (DataRow dr in tar_table[tbl_inx].Rows)
                {
                    dr["ID"] = id + 1 + r_inx;
                    r_inx++;
                }
                exp_proc.BatchBulkCopy(sqlcon, tar_table[tbl_inx], t);
                tbl_inx++;

            }

            MessageBox.Show(new Form { TopMost = true }, "Save data complete", "Warning");

        }
        public void save_data()
        {


            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
            string tbl_process = cb_process.SelectedItem.ToString().ToUpper();
            List<string> tar_table_name = new List<string>() { "BVH_PTH_LOGFILE", tbl_process };
            List<DataTable> tar_table = new List<DataTable>() { (DataTable)DGV_logfile.DataSource, (DataTable)DGV_data_process.DataSource };

            int tbl_inx = 0;


            foreach (var t in tar_table_name)
            {
                /* start_lbl:*/
                TDMK_Code.Delelte_FilteredItem_arr(t, sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                int id = TDMK_Code.SQL_MAX(t, "ID", sqlcon);
                int r_inx = 0;
                foreach (DataRow dr in tar_table[tbl_inx].Rows)
                {
                    dr["ID"] = id + 1 + r_inx;
                    r_inx++;
                }
                exp_proc.BatchBulkCopy(sqlcon, tar_table[tbl_inx], t);
                tbl_inx++;

            }
            MessageBox.Show(new Form { TopMost = true }, "Save data completed", "Warning");
        }


        private void tsmSaveData_Click_1(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && cb_process.SelectedIndex != -1)
            {
                if (DGV_logfile.Rows.Count > 0 && DGV_data_process.Rows.Count > 0)
                {

                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                    string tbl_process = cb_process.SelectedItem.ToString().ToUpper();
                    List<string> tar_table_name = new List<string>() { "BVH_PTH_LOGFILE", tbl_process };
                    List<DataTable> tar_table = new List<DataTable>() { (DataTable)DGV_logfile.DataSource, (DataTable)DGV_data_process.DataSource };

                    int tbl_inx = 0;
                    DataTable dt_spec = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SPEC_" + cb_process.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
                    int count_sample = int.Parse(dt_spec.Rows[0]["Remark"].ToString().Split('+')[2]);
                    if (DGV_data_process.Rows.Count >= count_sample)
                    {
                        bool chk = true;
                        for (int i = 0; i < DGV_data_process.Columns.Count; i++)
                        {

                            for (int k = 0; k < DGV_data_process.Rows.Count; k++)
                            {
                                if (DGV_data_process.Rows[k].Cells[i].Style.BackColor == Color.Red)
                                {
                                    chk = false;
                                    goto save_lbl;
                                }
                                if((myCode.checkDBNull(DGV_data_process.Rows[k].Cells["Image_data"].Value) == "")|| (myCode.checkDBNull(DGV_data_process.Rows[k].Cells["Image_data_2"].Value) == ""))
                                {
                                    chk = false;
                                    goto save_lbl;
                                }    

                            }
                        }
                    save_lbl: if (chk)
                        {
                            foreach (var t in tar_table_name)
                            {
                                /* start_lbl:*/
                                if (t == "BVH_PTH_LOGFILE")
                                {
                                    TDMK_Code.Delelte_FilteredItem_arr(t, sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { txtItemCode.Text, txtLotNo.Text, cb_process.SelectedItem.ToString() }));
                                }
                                else
                                {
                                    TDMK_Code.Delelte_FilteredItem_arr(t, sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                                }

                                int id = TDMK_Code.SQL_MAX(t, "ID", sqlcon);
                                int r_inx = 0;
                                foreach (DataRow dr in tar_table[tbl_inx].Rows)
                                {
                                    dr["ID"] = id + 1 + r_inx;
                                    r_inx++;
                                }
                                exp_proc.BatchBulkCopy(sqlcon, tar_table[tbl_inx], t);
                                tbl_inx++;

                            }
                            MessageBox.Show(new Form { TopMost = true }, "Save data complete", "Warning");
                        }
                        else
                        {
                            if (MessageBox.Show(new Form { TopMost = true }, "Data is out of spec. Do you want to save to database ?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                            {
                                if (admin_mode)
                                {
                                    chk = true;
                                    goto save_lbl;

                                }
                                else
                                {
                                    MessageBox.Show("Please, login to save data", "Warning");
                                }

                            }

                        }

                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Not enough data. Cannot save to database", "Warning");
                    }


                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "No data", "Warning");
                }


            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Please, fill in ItemCode / LotNo / Operator/ Process", "Warning");
            }
        }


        private void tsmLoadData_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "")
            {
                if (cb_process.SelectedIndex != -1)
                {
                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                    string tbl_process = cb_process.SelectedItem.ToString().ToUpper();

                    lbl_title.Text = "Data Analysis For " + tbl_process;
                    DGV_logfile.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { txtItemCode.Text, txtLotNo.Text, cb_process.SelectedItem.ToString() }));
                    DataTable dt_before = TDMK_Code.Datatable_Filter(sqlcon, tbl_process, filter_str);
                    DGV_data_process.DataSource = dt_before;
                    if (DGV_data_process.Rows.Count > 0)
                    {
                        ((DataGridViewImageColumn)DGV_data_process.Columns["Image_data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                        ((DataGridViewImageColumn)DGV_data_process.Columns["Image_data"]).Width = 100;

                        ((DataGridViewImageColumn)DGV_data_process.Columns["Image_data_2"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                        ((DataGridViewImageColumn)DGV_data_process.Columns["Image_data_2"]).Width = 100;
                        foreach (DataGridViewRow dr in DGV_data_process.Rows)
                        {
                            dr.Height = 70;
                        }

                        Check_spec();
                        //  change_ID_DGV(DGV_logfile);
                        // change_ID_DGV(DGV_data_process);
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "No data", "Warning");
                    }


                    //DGV_data_process.Columns.RemoveAt(DGV_data_process.Columns.Count - 1);

                    //DataGridViewColumn ImageColumn = new DataGridViewImageColumn()
                    //{
                    //    Name = "Image_data",
                    //    HeaderText = "Image_data",
                    //    SortMode = DataGridViewColumnSortMode.NotSortable,
                    //    Width = 100,

                    //    ImageLayout = DataGridViewImageCellLayout.Zoom,
                    //    //AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader
                    //    // Resizable = DataGridViewTriState.False
                    //};
                    //DGV_data_process.Columns.Add(ImageColumn);


                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Select process", "Warning");
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Please, fill in ItemCode / LotNo / Operator ID", "Warning");
            }

        }


        private void cbProcess_MasterList_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btn_set_Click(object sender, EventArgs e)
        {

        }

        private void btn_saveDB_Click(object sender, EventArgs e)
        {

        }

        private void cbProcess_MasterList_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            DGV_MasterList.DataSource = null;
            //if (cbProcess_MasterList.SelectedIndex > -1 && txtItemCode.Text != "")
            //{
            //    string process = cbProcess_MasterList.SelectedItem.ToString().ToUpper();
            //    DataTable dt1 = new DataTable();
            //    dt1 = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_MASTERLIST", TDMK_Code.filter_str(new string[] { "ItemCode", "Process" }, new string[] { txtItemCode.Text, cbProcess_MasterList.SelectedItem.ToString() }));
            //    DGV_MasterList.DataSource = dt1;
            //    myCode.DGV_Auto_Resize(DGV_MasterList);

            //    foreach (DataGridViewColumn column in DGV_MasterList.Columns)
            //    {
            //        if ((column.Name != "Post_No") && (column.Name != "Remark"))
            //        {
            //            column.ReadOnly = true;
            //        }
            //    }
            //}
        }

        private void btn_set_Click_1(object sender, EventArgs e)
        {
            if (txt_masterlist.Text != "" && cbProcess_MasterList.SelectedIndex != -1)
            {

                string process = cbProcess_MasterList.SelectedItem.ToString().ToUpper();
                // DataTable dt1 = new DataTable();

                DataTable dt1 = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_MASTERLIST", TDMK_Code.filter_str(new string[] { "ItemCode", "Process" }, new string[] { txt_masterlist.Text, cbProcess_MasterList.SelectedItem.ToString() }));
                if (dt1.Rows.Count == 0)
                {

                    dt1 = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_MASTERLIST", TDMK_Code.filter_str(new string[] { "ItemCode", "Process" }, new string[] { "22B0500", cbProcess_MasterList.SelectedItem.ToString() }));
                }
                //dt1 = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_MASTERLIST", TDMK_Code.filter_str(new string[] { "ItemCode", "Process" }, new string[] { txt_masterlist.Text, cbProcess_MasterList.SelectedItem.ToString() }));
                if (dt1.Rows.Count > 0)
                {

                    DGV_MasterList.DataSource = dt1;
                    for (int i = 0; i < DGV_MasterList.Rows.Count; i++)
                    {
                        DGV_MasterList.Rows[i].Cells["ItemCode"].Value = txt_masterlist.Text;
                    }
                }
                else
                {
                    string[] ItemName = { "Location1", "Location2", "Location3", "Location4", "Etched_Cu_thickness", "Plated_Cu_thickness", "Adhesive_etch_back", "Top_diameter" };
                    string[] ItemName1 = { "Corner_thickness", "Side_wall_from_PI", "Side_wall_from_adhesive", "Side_wall_Nickel", "Min_Plate ", "Etch_back", "Diameter" };
                    DataTable dt = new DataTable();

                    dt.Columns.Add("ID");
                    dt.Columns.Add("ItemCode");
                    dt.Columns.Add("Process");
                    dt.Columns.Add("ItemName");
                    dt.Columns.Add("Post_No");
                    dt.Columns.Add("Remark");
                    if (process == "BVH_WITH_BONDING_SHEET")
                    {
                        for (int i = 0; i < 8; i++)
                        {
                            DataRow dr = dt.NewRow();

                            dr["ID"] = i + 1;
                            dr[1] = txt_masterlist.Text;
                            dr[2] = cbProcess_MasterList.SelectedItem;
                            dr[3] = arr_item_BVH_with[i + 4];
                            dr[4] = "";
                            dr[5] = "";
                            dt.Rows.Add(dr);
                        }
                    }
                    if (process == "BVH_WITHOUT_BONDING_SHEET")
                    {
                        for (int i = 0; i < 8; i++)
                        {
                            DataRow dr = dt.NewRow();

                            dr["ID"] = i + 1;
                            dr[1] = txt_masterlist.Text;
                            dr[2] = cbProcess_MasterList.SelectedItem;
                            dr[3] = arr_item_BVH_without[i + 4];
                            dr[4] = "";
                            dr[5] = "";
                            dt.Rows.Add(dr);
                        }
                    }
                    if (process == "PLATED_THROUGH_HOLE")
                    {
                        for (int i = 0; i < 7; i++)
                        {
                            DataRow dr = dt.NewRow();

                            dr["ID"] = i + 1;
                            dr[1] = txt_masterlist.Text;
                            dr[2] = cbProcess_MasterList.SelectedItem;
                            dr[3] = arr_item_PTH[i + 4];
                            dr[4] = "";
                            dr[5] = "";
                            dt.Rows.Add(dr);
                        }
                    }
                    DGV_MasterList.DataSource = dt;

                }
                myCode.DGV_Auto_Resize(DGV_MasterList);
                foreach (DataGridViewColumn column in DGV_MasterList.Columns)
                {
                    if (column.Name == "Post_No" || column.Name == "Remark")
                    {
                        column.ReadOnly = false;
                    }
                    else
                    {
                        column.ReadOnly = true;
                    }
                }

            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "ItemCode or Process has not been selected", "Warning!");
            }
        }

        private void btn_saveDB_Click_1(object sender, EventArgs e)
        {
            if (txt_masterlist.Text != "" && cbProcess_MasterList.SelectedIndex != -1)
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Process" }, new string[] { txt_masterlist.Text, cbProcess_MasterList.SelectedItem.ToString() });
                DataTable tar_table = (DataTable)DGV_MasterList.DataSource;

            start_lbl: DataTable cur_dt = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_MASTERLIST", filter_str);
                if (cur_dt.Rows.Count > 0)
                {

                    if (MessageBox.Show(new Form { TopMost = true }, "Table BVH_PTH_MASTERLIST : Data is existed. Do you want to overwrite?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        TDMK_Code.Delelte_FilteredItem_arr("BVH_PTH_MASTERLIST", sqlcon, filter_str);
                        goto start_lbl;

                    }

                }
                else
                {
                    int id = TDMK_Code.SQL_MAX("BVH_PTH_MASTERLIST", "ID", sqlcon);
                    int r_inx = 0;
                    foreach (DataRow dr in cur_dt.Rows)
                    {
                        dr["ID"] = id + 1 + r_inx;
                        r_inx++;
                    }
                    exp_proc.BatchBulkCopy(sqlcon, tar_table, "BVH_PTH_MASTERLIST");
                    MessageBox.Show(new Form { TopMost = true }, "Completed", "Warning!");

                }

            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "ItemCode or Process has not been selected", "Warning!");
            }

        }

        private void tableLayoutPanel30_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {

            if (txtitemcode_view.Text != "")
            {

                // TDMK_Code.fill_dataset_DGV("BVH_WITH_BONDING_SHEET", @"Select * from BVH_WITH_BONDING_SHEET WHERE ItemCode = '" + txtitemcode_view.Text + "' AND " + "LotNo = '" + txt_lotno_view.Text + "'" , DGV_BVH_with, sqlcon);
                // TDMK_Code.fill_dataset_DGV("BVH_WITHOUT_BONDING_SHEET", @"Select * from BVH_WITHOUT_BONDING_SHEET WHERE ItemCode = '" + txtitemcode_view.Text + "' AND " + "LotNo = '" + txt_lotno_view.Text + "'", DGV_BVH_without, sqlcon);
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtitemcode_view.Text, txt_lotno_view.Text });
                string tbl_process = cb_process_view.SelectedItem.ToString().ToUpper();


                DataTable dt_before = TDMK_Code.Datatable_Filter(sqlcon, tbl_process, filter_str);
                DGV_data_view.DataSource = dt_before;
                //   change_ID_DGV(DGV_data_view);
                if (DGV_data_view.Rows.Count > 0)
                {
                    ((DataGridViewImageColumn)DGV_data_view.Columns["Image_data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)DGV_data_view.Columns["Image_data"]).Width = 100;
                    ((DataGridViewImageColumn)DGV_data_view.Columns["Image_data_2"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)DGV_data_view.Columns["Image_data_2"]).Width = 100;
                    foreach (DataGridViewRow dr in DGV_data_view.Rows)
                    {
                        dr.Height = 70;
                    }

                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "No data", "Warning");
                }
            }

        }

        public String setup_spec_BVH_PTH(myExcel.Worksheet ws, int k, ref AutoCompleteStringCollection list)
        {
            //  AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            string info = "";

            for (int j = 1; j < 20; j++)
            {
                //if(j == 3 || j == 17) 
                // {
                for (int i = 1; i < 200; i++)
                {

                    if (myCode.checkDBNull(ws.Cells[i, j].Value) == "S/N")
                    {
                        int count = 1;
                        while (ws.Cells[i + count, j].Value != null)
                        {
                            count++;
                        }
                        list.Add(i.ToString() + "+" + j.ToString() + "+" + (count - 1).ToString());
                    }

                    if (myCode.checkDBNull(ws.Cells[i, j].Value) == "USL")
                    {
                        int a = 1;
                        if (myCode.checkDBNull(ws.Cells[i - 7, j + 1].Value).Replace(" ", string.Empty).ToUpper() == "Full filled via".Replace(" ", string.Empty).ToUpper())
                        {
                            a = 2;
                        }
                        for (int m = a; m < k + 2; m++)
                        {
                            string USL = myCode.checkDBNull(ws.Cells[i, j + m].Value);
                            string LSL = myCode.checkDBNull(ws.Cells[i + 1, j + m].Value);

                            if (USL != "")
                            {
                                //list.Add(USL + ";");
                                info += USL + ";";
                            }
                            else
                            {
                                //list.Add("NA" + ";");
                                info += "NA" + ";";
                            }

                            if (LSL != "")
                            {
                                //list.Add(LSL + ";");
                                info += LSL + ";";
                            }
                            else
                            {
                                //list.Add("NA" + ";");
                                info += "NA" + ";";
                            }
                            info += "_";

                        }
                        //list.Add( "+");
                        info += "+";
                    }


                    // }
                }

            }
            return info;
        }
        public string get_infor_each_process(myExcel.Worksheet ws, myExcel.Range rgn, string key_word, string offset_col, int k)
        {
            string inf = "";
            for (int i = 30; i < 200; i++)
            {
                for (int t = 0; t < 5; t++)
                {
                    if (myCode.checkDBNull(rgn.Offset[i, t].Value) == "S/N")
                    {
                        int count = 1;
                        while (rgn.Offset[i + count, t].Value != null)
                        {
                            if (myCode.IsNumeric(rgn.Offset[i + count, 1].Value.ToString()))
                                count++;
                        }
                        inf = key_word + "+" + i.ToString() + "+" + offset_col + "+" + (count - 1).ToString() + "/";


                        string info = "";
                        for (int offset = 0; offset < 15; offset++)
                        {
                            if (myCode.checkDBNull(ws.Cells[i - offset, 3].Value) == "USL")
                            {
                                for (int m = 2; m < k + 2; m++)
                                {
                                    string USL = myCode.checkDBNull(ws.Cells[i - offset, 3 + m].Value);
                                    string LSL = myCode.checkDBNull(ws.Cells[i - offset + 1, 3 + m].Value);

                                    if (USL != "")
                                    {
                                        info += USL + ";";
                                    }
                                    else
                                    {

                                        info += "NA" + ";";
                                    }

                                    if (LSL != "")
                                    {

                                        info += LSL + ";";
                                    }
                                    else
                                    {

                                        info += "NA" + ";";
                                    }
                                    info += "_";


                                }

                                info += "+";
                                break;

                            }
                        }
                        inf += info;

                        break;
                    }
                }



            }
            return inf;
        }

        public String setup_spec_BVH_PTH_Update(myExcel.Worksheet ws, ref AutoCompleteStringCollection list)
        {

            string info = "";
            string str_spec = "";


            myExcel.Range rgn_B = ws.Range["B1"];
            myExcel.Range rgn_P = ws.Range["O1"];

            for (int i = 1; i < 200; i++)
            {
                if (myCode.checkDBNull(rgn_B.Offset[i, 0].Value).ToUpper().Replace(" ", "").Contains("BVHWITHBONDINGSHEET"))
                {
                    info = get_infor_each_process(ws, rgn_B, "B1", "3", 8);
                    list.Add(info.Split('/')[0]);
                    str_spec = info.Split('/')[1] + "+";
                }
                else if (myCode.checkDBNull(rgn_B.Offset[i, 0].Value).ToUpper().Replace(" ", "").Contains("BVHWITHOUTBONDINGSHEET"))
                {
                    info = get_infor_each_process(ws, rgn_B, "B2", "3", 8);
                    list.Add(info.Split('/')[0]);
                    str_spec = info.Split('/')[1] + "+";
                }
                else if (myCode.checkDBNull(rgn_P.Offset[i, 0].Value).ToUpper().Replace(" ", "").Contains("PLATEDTHROUGHHOLE"))
                {
                    info = get_infor_each_process(ws, rgn_P, "P1", "17", 7);
                    list.Add(info.Split('/')[0]);
                    str_spec = info.Split('/')[1] + "+";
                }

            }

            return str_spec;
        }
        public String setup_spec_BVH_PTH_2(myExcel.Worksheet ws, int k, ref AutoCompleteStringCollection list)
        {
            //  AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            string info = "";

            for (int j = 1; j < 20; j++)
            {
                for (int i = 1; i < 200; i++)
                {

                    if (myCode.checkDBNull(ws.Cells[i, j].Value) == "S/N")
                    {
                        int count = 1;
                        while (ws.Cells[i + count, j].Value != null)
                        {
                            if (myCode.IsNumeric(ws.Cells[i + count, j].Value.ToString()))
                                count++;
                        }
                        list.Add(i.ToString() + "+" + j.ToString() + "+" + (count - 1).ToString());


                        for (int offset = 0; offset < 15; offset++)
                        {
                            if (myCode.checkDBNull(ws.Cells[i - offset, j].Value) == "USL")
                            {
                                for (int m = 2; m < k + 2; m++)
                                {
                                    string USL = myCode.checkDBNull(ws.Cells[i - offset, j + m].Value);
                                    string LSL = myCode.checkDBNull(ws.Cells[i - offset + 1, j + m].Value);

                                    if (USL != "")
                                    {

                                        info += USL + ";";
                                    }
                                    else
                                    {

                                        info += "NA" + ";";
                                    }

                                    if (LSL != "")
                                    {

                                        info += LSL + ";";
                                    }
                                    else
                                    {

                                        info += "NA" + ";";
                                    }
                                    info += "_";


                                }

                                info += "+";
                                break;

                            }

                        }


                    }



                }
            }
            return info;
        }
        public delegate void Get_path(String path);
        private void SetValue(String value)
        {
            this.txtLogfile.Text = value;
        }

        private void txtLogfile_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && cb_process.SelectedIndex != -1)
            {
                DataTable dt = TDMK_Code.Datatable_Filter(myVar.mysqlcon, cb_process.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show(new Form { TopMost = true }, "Data of " + cb_process.SelectedItem.ToString().ToUpper() + " is exist. Please update data", "Waring");
                    //if( MessageBox.Show(new Form { TopMost = true }, "Do you load data from database ?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    // {
                    //     string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                    //     string tbl_process = cb_process.SelectedItem.ToString().ToUpper();

                    //     lbl_title.Text = "Data For " + tbl_process;
                    //     DGV_logfile.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { txtItemCode.Text, txtLotNo.Text, cb_process.SelectedItem.ToString() }));
                    //     DataTable dt_before = TDMK_Code.Datatable_Filter(sqlcon, tbl_process, filter_str);
                    //     DGV_data_process.DataSource = dt_before;
                    //     if (DGV_data_process.Rows.Count > 0)
                    //     {
                    //         ((DataGridViewImageColumn)DGV_data_process.Columns["Image_data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    //         ((DataGridViewImageColumn)DGV_data_process.Columns["Image_data"]).Width = 100;
                    //         foreach (DataGridViewRow dr in DGV_data_process.Rows)
                    //         {
                    //             dr.Height = 70;
                    //         }

                    //         Check_spec();
                    //     }
                    // }

                }
                else
                {
                    FolderBrowserDialog f_open = new FolderBrowserDialog();
                    f_open.SelectedPath = System.Windows.Forms.Application.StartupPath;
                    if (f_open.ShowDialog() == DialogResult.OK)
                    {
                        txtLogfile.Text = f_open.SelectedPath;
                        if (txtLogfile.Text != "")
                        {

                            ImageDetail fr1 = new ImageDetail(SetValue);
                            fr1.startpath = txtLogfile.Text;
                            fr1.Show();


                        }
                        if (txtLogfile.Text != "")
                        {
                            btn_load.Enabled = true;
                            btn_insert.Enabled = false;
                        }


                    }
                }
            }



        }




        private void btn_setup_spec_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && cb_process.SelectedIndex != -1)
            {
                string[] item = { "ItemCode" };
                string[] item_val = { txtItemCode.Text };
                DGV_spec.DataSource = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SPEC_" + cb_process.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(item, item_val));
                if (DGV_spec.Rows.Count > 0)
                {
                    if (MessageBox.Show(new Form { TopMost = true }, "Spec is existed.Do you want to overwrite ? ", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        foreach (string process in cb_process.Items)
                        {
                            TDMK_Code.Delelte_FilteredItem_arr("SPEC_" + process.ToUpper(), myVar.mysqlcon, TDMK_Code.filter_str(item, item_val));
                        }

                        Load_spec();
                    }
                    else
                    {
                        return;
                    }
                }
                else
                {
                    Load_spec();
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "ItemCode or Process has not been selected", "Warning!");
            }



        }

        private void btn_insert_Click(object sender, EventArgs e)
        {
            //FolderBrowserDialog f_open = new FolderBrowserDialog();
            //f_open.SelectedPath = System.Windows.Forms.Application.StartupPath;
            //if (f_open.ShowDialog() == DialogResult.OK)
            //{
            //   Insert_logfile(f_open.SelectedPath, Image_result);
            //}



            if (txtLogfile.Text == "")
            {
                FolderBrowserDialog f_open = new FolderBrowserDialog();
                f_open.SelectedPath = System.Windows.Forms.Application.StartupPath;
                if (f_open.ShowDialog() == DialogResult.OK)
                {
                    txtLogfile.Text = f_open.SelectedPath;
                    if (txtLogfile.Text != "")
                    {

                        ImageDetail fr1 = new ImageDetail(SetValue);
                        fr1.startpath = txtLogfile.Text;
                        fr1.Show();
                    }
                }
            }
            else
            {
                ImageDetail fr1 = new ImageDetail(SetValue);
                fr1.startpath = txtLogfile.Text;
                fr1.Show();
            }
            insert_mode = true;
            btn_load.Enabled = true;

            // Insert_logfile(txtLogfile.Text, Image_result, Image_result_2);
            // btn_load.Enabled = false;

        }

        private void txtLogfile_TextChanged(object sender, EventArgs e)
        {
            txtLogfile.Text = txtLogfile.Text.Replace("\r", "").Replace("\n", "");
        }

        private void txtLogfile_KeyDown(object sender, KeyEventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && cb_process.SelectedIndex != -1)
            {
                DataTable dt = TDMK_Code.Datatable_Filter(myVar.mysqlcon, cb_process.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show(new Form { TopMost = true }, "Data of " + cb_process.SelectedItem.ToString().ToUpper() + " is exist. Please update data", "Waring");


                }
                else
                {
                    if (e.KeyCode == Keys.Enter)
                    {
                        if (txtLogfile.Text != "")
                        {

                            ImageDetail fr1 = new ImageDetail(SetValue);
                            fr1.startpath = txtLogfile.Text;
                            fr1.Show();
                            btn_load.Enabled = true;
                        }


                    }

                }

            }

        }

        private void DGV_spec_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            btn_Update_spec.Enabled = true;
        }

        private void btn_Update_spec_Click(object sender, EventArgs e)
        {
            if (cb_process.SelectedIndex != -1)
            {
                TDMK_Code.Delelte_FilteredItem_arr("SPEC_" + cb_process.SelectedItem.ToString().ToUpper(), sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
                exp_proc.BatchBulkCopy(sqlcon, (DataTable)DGV_spec.DataSource, "SPEC_" + cb_process.SelectedItem.ToString().ToUpper());
                DGV_spec.DataSource = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SPEC_" + cb_process.SelectedItem.ToString(), TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
                MessageBox.Show(new Form { TopMost = true }, "Update spec complete", "Warning");
                Check_spec();
                btn_Update_spec.Enabled = false;
            }

        }

        public void Get_logfile_Multi_update20_2(string in_src, ref SortedDictionary<int, Dictionary<int, string>> lst_result, ref SortedDictionary<int, int> DODAY_Count_lst)
        {
            SortedDictionary<int, Dictionary<int, string>> dic_dk = new SortedDictionary<int, Dictionary<int, string>> { };
            SortedDictionary<int, Dictionary<int, string>> dic_doday = new SortedDictionary<int, Dictionary<int, string>> { };


            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            DirectoryInfo[] arr_dic = tar_d.GetDirectories();

            foreach (DirectoryInfo d in arr_dic)
            {
                if (d.Name.ToUpper().Replace(" ", "").Contains("DK"))
                {
                    FileInfo[] temp_lst = d.GetFiles("*.csv");
                    if (temp_lst.Length > 0)
                    {
                        for (int i = 0; i < temp_lst.Length; i++)
                        {
                            Dictionary<int, string> temp = Imp_code.Get_LogFile_Data(temp_lst[i].FullName);
                            string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });

                            //f_na = f_na.Replace("+", "").Replace("-", "");
                            f_na = f_na.Replace(" ", "");
                            if (myCode.IsNumeric(f_na))
                            {
                                int f_inx = Convert.ToInt32(f_na);
                                if (!dic_dk.ContainsKey(f_inx))
                                {
                                    dic_dk.Add(f_inx, temp);
                                }
                            }
                        }
                    }
                }
                else if (d.Name.ToUpper().Replace(" ", "").Contains("DODAY"))
                {
                    FileInfo[] temp_lst = d.GetFiles("*.csv");
                    if (temp_lst.Length > 0)
                    {
                        for (int i = 0; i < temp_lst.Length; i++)
                        {
                            Dictionary<int, string> temp = Imp_code.Get_LogFile_Data(temp_lst[i].FullName);
                            string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
                            f_na = f_na.Replace(" ", "");
                            if (myCode.IsNumeric(f_na))
                            {
                                int f_inx = Convert.ToInt32(f_na);
                                if (!dic_doday.ContainsKey(f_inx))
                                {
                                    dic_doday.Add(f_inx, temp);
                                }
                            }
                        }
                    }
                }
            }
            DataTable dt_master = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_MASTERLIST", TDMK_Code.filter_str(new string[] { "ItemCode", "Process" }, new string[] { txtItemCode.Text, cb_process.SelectedItem.ToString() }));
            List<string> DoDay_lst= dt_master.AsEnumerable().Where(x=>x.Field<string>("Post_No").Replace(" ","").ToUpper().Contains("DODAY")).Select(x=>x.Field<string>("Post_No")).ToList();
            DoDay_lst = DoDay_lst.Select(x=>x.Split('_').LastOrDefault()).ToList();
            foreach (var item in dic_doday.Values)
            {
                var post_lst = item.Keys.ToList();
                List<int> remove_lst =post_lst.Where(x => DoDay_lst.IndexOf(x.ToString()) == -1).ToList();
               foreach(var k in remove_lst)
                {
                    item.Remove(k);
                }    
            }
          
            //SortedDictionary<int, Dictionary<int, string>> dic_sum = new SortedDictionary<int, Dictionary<int, string>> { };
            count_doday = dic_doday[1].Count;
            foreach (int smp in dic_dk.Keys)
            {
                Dictionary<int, string> dic1 = new Dictionary<int, string> { };
                int k = 1;
                foreach(string s in dic_doday[smp].Values)
                {
                    dic1.Add(k, s);
                    k++;
                }
                foreach (string s in dic_dk[smp].Values)
                {
                    dic1.Add(k, s);
                    k++;
                }
                
               lst_result.Add(smp, dic1);
            }
            foreach(var d in dic_doday)
            {
                if(DODAY_Count_lst.ContainsKey(d.Key))
                {
                    DODAY_Count_lst[d.Key] = d.Value.Count;
                }
                else
                {
                    DODAY_Count_lst.Add(d.Key,d.Value.Count);
                }
            }
             
             
        }
        public void Get_logfile_Multi_update20_2(string in_src, ref SortedDictionary<int, Dictionary<string, string>> lst_result)
        {
            SortedDictionary<int, Dictionary<int, string>> dic_dk = new SortedDictionary<int, Dictionary<int, string>> { };
            SortedDictionary<int, Dictionary<int, string>> dic_doday = new SortedDictionary<int, Dictionary<int, string>> { };


            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            DirectoryInfo[] arr_dic = tar_d.GetDirectories();

            foreach (DirectoryInfo d in arr_dic)
            {
                if (d.Name.ToUpper().Replace(" ", "").Contains("DK"))
                {
                    FileInfo[] temp_lst = d.GetFiles("*.csv");
                    if (temp_lst.Length > 0)
                    {
                        for (int i = 0; i < temp_lst.Length; i++)
                        {
                            Dictionary<int, string> temp = Imp_code.Get_LogFile_Data(temp_lst[i].FullName);
                            string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });

                            //f_na = f_na.Replace("+", "").Replace("-", "");
                            f_na = f_na.Replace(" ", "");
                            if (myCode.IsNumeric(f_na))
                            {
                                int f_inx = Convert.ToInt32(f_na);
                                if (!dic_dk.ContainsKey(f_inx))
                                {
                                    dic_dk.Add(f_inx, temp);
                                }
                            }
                        }
                    }
                }
                else if (d.Name.ToUpper().Replace(" ", "").Contains("DODAY"))
                {
                    FileInfo[] temp_lst = d.GetFiles("*.csv");
                    if (temp_lst.Length > 0)
                    {
                        for (int i = 0; i < temp_lst.Length; i++)
                        {
                            Dictionary<int, string> temp = Imp_code.Get_LogFile_Data(temp_lst[i].FullName);
                            string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
                            f_na = f_na.Replace(" ", "");
                            if (myCode.IsNumeric(f_na))
                            {
                                int f_inx = Convert.ToInt32(f_na);
                                if (!dic_doday.ContainsKey(f_inx))
                                {
                                    dic_doday.Add(f_inx, temp);
                                }
                            }
                        }
                    }
                }
            }
            //DataTable dt_master = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_MASTERLIST", TDMK_Code.filter_str(new string[] { "ItemCode", "Process" }, new string[] { txtItemCode.Text, cb_process.SelectedItem.ToString() }));
            //List<string> DoDay_lst = dt_master.AsEnumerable().Where(x => x.Field<string>("Post_No").Replace(" ", "").ToUpper().Contains("DODAY")).Select(x => x.Field<string>("Post_No")).ToList();
            //DoDay_lst = DoDay_lst.Select(x => x.Split('_').LastOrDefault()).ToList();
            //foreach (var item in dic_doday.Values)
            //{
            //    var post_lst = item.Keys.ToList();
            //    List<int> remove_lst = post_lst.Where(x => DoDay_lst.IndexOf(x.ToString()) == -1).ToList();
            //    foreach (var k in remove_lst)
            //    {
            //        item.Remove(k);
            //    }
            //}
            foreach(var doday in dic_doday)
            {
                Dictionary<string, string> dic_result = new Dictionary<string, string>();
                foreach(var doday_val in doday.Value )
                {
                    string _key = "DODAY_" + doday_val.Key.ToString();
                    if(dic_result.ContainsKey(_key))
                    {
                        dic_result[_key]=doday_val.Value;
                    }
                    else
                    {
                        dic_result.Add(_key, doday_val.Value);
                    }
                }
                if(lst_result.ContainsKey(doday.Key))
                {
                    lst_result[doday.Key] = dic_result;
                }
                else
                {
                    lst_result.Add(doday.Key, dic_result);
                }
            }
            foreach (var dk in dic_dk)
            {
                Dictionary<string, string> dic_result = new Dictionary<string, string>();
                foreach (var dk_val in dk.Value)
                {
                    string _key = "DK_" + dk_val.Key.ToString();
                    if (dic_result.ContainsKey(_key))
                    {
                        dic_result[_key]= dk_val.Value;
                    }
                    else
                    {
                        dic_result.Add(_key, dk_val.Value);
                    }
                }
                if (lst_result.ContainsKey(dk.Key))
                {
                    var temp_dic = lst_result[dk.Key];
                    foreach(var temp_dk in dic_result)
                    {
                        temp_dic.Add(temp_dk.Key,temp_dk.Value);
                    }
                }
                else
                {
                    lst_result.Add(dk.Key, dic_result);
                }
            }

            //SortedDictionary<int, Dictionary<int, string>> dic_sum = new SortedDictionary<int, Dictionary<int, string>> { };
            //count_doday = dic_doday[1].Count;
            //foreach (int smp in dic_dk.Keys)
            //{
            //    Dictionary<int, string> dic1 = new Dictionary<int, string> { };
            //    int k = 1;
            //    foreach (string s in dic_doday[smp].Values)
            //    {
            //        dic1.Add(k, s);
            //        k++;
            //    }
            //    foreach (string s in dic_dk[smp].Values)
            //    {
            //        dic1.Add(k, s);
            //        k++;
            //    }

            //    lst_result.Add(smp, dic1);
            //}


        }

        private void btn_load_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && cb_process.SelectedIndex != -1 && txtLogfile.Text != "")
            {
                count_doday = 0;
                DirectoryInfo temp_dir = new DirectoryInfo(txtLogfile.Text);
                string f_region = temp_dir.Name;
                switch (f_region.ToUpper())
                {
                    case "BVH BS":
                        f_region = "BVH With Bonding Sheet";
                        break;
                    case "BVH NBS":
                        f_region = "BVH Without Bonding Sheet";
                        break;
                }
                string f_name = temp_dir.Parent.Name;
                bool chk_name = false;
                bool check_BVH = f_region.ToUpper().Replace("_", string.Empty).Replace("-", string.Empty).Replace(" ", string.Empty) == cb_process.SelectedItem.ToString().ToUpper().Replace("_", string.Empty).Replace("-", string.Empty).Replace(" ", string.Empty) ? true : false;
                bool check_PTH = f_region.Contains("PTH") && (cb_process.SelectedItem.ToString() == "Plated_Through_Hole" || f_region.ToUpper().Replace("_", string.Empty).Replace("-", string.Empty).Replace(" ", string.Empty) == cb_process.SelectedItem.ToString().ToUpper().Replace("_", string.Empty).Replace("-", string.Empty).Replace(" ", string.Empty)) ? true : false;
                if (check_BVH || check_PTH)
                {
                    if (f_name.Split('-')[0].Trim() == txtItemCode.Text)
                    {
                        if (txtLotNo.Text.Contains("-") && f_name.Split('_')[0].Split('-').Length == 3)
                        {
                            if (Convert.ToDecimal(f_name.Split('-')[1]) == Convert.ToDecimal(txtLotNo.Text.ToString().Split('-')[0]) && Convert.ToDecimal(f_name.Split('-')[2].Split('_')[0]) == Convert.ToDecimal(txtLotNo.Text.ToString().Split('-')[1]))
                                chk_name = true;
                        }
                        else if (!txtLotNo.Text.Contains("-") && f_name.Split('_')[0].Split('-').Length == 2)
                        {
                            if (Convert.ToDecimal(f_name.Split('-')[1].Split('_')[0].Trim()) == Convert.ToDecimal(txtLotNo.Text.ToString()))
                            {
                                chk_name = true;
                            }
                        }
                    }
                }
                if (chk_name)
                {
                    if (update_mode)
                    {
                        Update_data(txtLogfile.Text);
                        btn_load.Enabled = false;
                        update_mode = false;
                    }
                    else if (insert_mode)
                    {
                        Insert_logfile(txtLogfile.Text, Image_result, Image_result_2);
                        insert_mode = false;
                        btn_load.Enabled = false;
                    }
                    else
                    {
                        BVH_PTH_Data_Process(txtLogfile.Text);
                        btn_load.Enabled = false;
                    }
                    myCode.Disable_Sort_DGV(DGV_data_process);
                }


                else
                {
                    if (temp_dir.Name.Contains(txtItemCode.Text))
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Select data for each process", "Warning");
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "ItemCode or LotNo not matched", "Warning");
                    }
                }

            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Please, fill in ItemCode / LotNo / Operator ID", "Warning");

            }
        }

        private void txtLogfile_Enter(object sender, EventArgs e)
        {
            //if (txtLogfile.Text != "")
            //{

            //    ImageDetail fr1 = new ImageDetail(SetValue);
            //    fr1.startpath = txtLogfile.Text;
            //    fr1.Show();


            //}
        }

        private void rbaddnew_CheckedChanged(object sender, EventArgs e)
        {
            //if (txtLotNo.Text != "" && txtItemCode.Text != "")
            //{
            //    if (rbaddnew.Checked)
            //    {
            //        rbupdate.Checked = false;
            //    }
            //}
        }

        private void rbupdate_CheckedChanged(object sender, EventArgs e)
        {
            //if(txtLotNo.Text != "" && txtItemCode.Text != "")
            //{
            //    if (rbupdate.Checked)
            //    {

            //        DataTable dt = TDMK_Code.Datatable_Filter(myVar.mysqlcon, cb_process.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
            //        if (dt.Rows.Count == 0)
            //        {
            //            MessageBox.Show("Data is not exist. Cannot update", "Warning");
            //            rbaddnew.Checked = true;
            //            rbupdate.Checked = false;

            //        }
            //        else
            //        {
            //            rbaddnew.Checked = false;
            //        }
            //    }

            //}

        }

        private void DGV_data_process_DataSourceChanged(object sender, EventArgs e)
        {
            edited_lst = new SortedDictionary<int, List<string>>();
        }

        private void DGV_data_process_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {

                //admin_mode = true;
                if (admin_mode)
                {
                    sel_cells_lst = new SortedDictionary<int, List<string>>();
                    List<string> ignored_col = new List<string>() { "ID", "ItemCode", "LotNo", "PCS_No", "Image_data", "Image_data_2", "Remark" };
                    foreach (DataGridViewCell cells in DGV_data_process.SelectedCells)
                    {
                        int r_inx = cells.RowIndex;
                        int c_inx = cells.ColumnIndex;
                        if (c_inx != -1 && r_inx != -1)
                        {
                            if (sel_cells_lst.Keys.ToList().IndexOf(r_inx) == -1)
                            {
                                foreach (DataGridViewColumn dc in DGV_data_process.Columns)
                                {
                                    string col_name = dc.Name;
                                    if (ignored_col.IndexOf(col_name) == -1)
                                    {
                                        if (sel_cells_lst.Keys.ToList().IndexOf(r_inx) != -1)
                                        {
                                            sel_cells_lst[r_inx].Add(col_name);
                                        }
                                        else
                                        {
                                            sel_cells_lst.Add(r_inx, new List<string>() { col_name });
                                        }
                                    }
                                }
                            }

                        }
                        //if(r_inx != -1 && c_inx != -1)
                        //{
                        //    string col_name = DGV_data_process.Columns[c_inx].Name;
                        //    if (ignored_col.IndexOf(col_name) == -1)
                        //    {
                        //        if (sel_cells_lst.Keys.ToList().IndexOf(r_inx) != -1)
                        //        {
                        //            sel_cells_lst[r_inx].Add(col_name);
                        //        }
                        //        else
                        //        {
                        //            sel_cells_lst.Add(r_inx, new List<string>() { col_name });
                        //        }
                        //    }
                        //}


                    }
                    if (sel_cells_lst.Count > 0)
                    {
                        FolderBrowserDialog f_open = new FolderBrowserDialog();
                        f_open.SelectedPath = System.Windows.Forms.Application.StartupPath;
                        SortedDictionary<int, Dictionary<int, string>> logfile_result = new SortedDictionary<int, Dictionary<int, string>>();
                        if (f_open.ShowDialog() == DialogResult.OK)
                        {

                            txtLogfile.Text = f_open.SelectedPath;
                            if (txtLogfile.Text != "")
                            {
                                ImageDetail fr1 = new ImageDetail(SetValue);
                                fr1.startpath = txtLogfile.Text;
                                fr1.Show();
                                update_mode = true;

                            }
                            if (txtLogfile.Text != "")
                            {
                                btn_load.Enabled = true;
                            } 
                        } 
                    }
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Please, login to edit data", "Warning");
                }
            }
        }
        public void Update_data(string src_path)
        {
            if (src_path != "")
            { 
                if (sel_cells_lst.Count > 0)
                { 
                    SortedDictionary<int, Dictionary<string, string>> logfile_result = new SortedDictionary<int, Dictionary<string, string>>();
                    Image_result = new SortedDictionary<int, byte[]>();
                    Image_result_2 = new SortedDictionary<int, byte[]>();
                    //Imp_code.Get_logfile_Multi(src_path, ref logfile_result);
                    SortedDictionary<int, int> lst_Doday_count = new SortedDictionary<int, int>();
                    Get_logfile_Multi_update20_2(src_path, ref logfile_result);

                    //Get_Image_Multi(src_path, ref Image_result);
                    //Get_Image_Multi_2(src_path, ref Image_result_2);

                    DirectoryInfo tar_d = new DirectoryInfo(src_path);
                    DirectoryInfo[] arr_dic = tar_d.GetDirectories();

                    //foreach (DirectoryInfo d in arr_dic)
                    //{
                    //    if (d.Name.ToUpper().Replace(" ", "").Contains("DODAY"))
                    //    {
                    //        Get_Image_Multi(Path.Combine(src_path, d.Name), ref Image_result);
                    //        Get_Image_Multi_2(Path.Combine(src_path, d.Name), ref Image_result_2); 
                    //    }
                    //}


                    foreach (DirectoryInfo d in arr_dic)
                    {
                        if (d.Name.ToUpper().Replace(" ", "").Contains("DK"))
                        {
                            Get_Image_Multi(Path.Combine(src_path, d.Name), ref Image_result);
                            //Get_Image_Multi_2(Path.Combine(src_path, d.Name), ref Image_result_2);
                        }
                        if (d.Name.ToUpper().Replace(" ", "").Contains("DODAY"))
                        {
                            //Get_Image_Multi(Path.Combine(src_path, d.Name), ref Image_result);
                            Get_Image_Multi_2(Path.Combine(src_path, d.Name), ref Image_result_2);
                        }
                    }



                    DataTable BVH_masterlist = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_MASTERLIST", TDMK_Code.filter_str(new string[] { "ItemCode", "Process" }, new string[] { txtItemCode.Text, cb_process.Text }));
                    if (BVH_masterlist.Rows.Count == 0)
                    {

                        BVH_masterlist = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_MASTERLIST", TDMK_Code.filter_str(new string[] { "ItemCode", "Process" }, new string[] { "22B0500", cb_process.Text }));
                    }
                    DataTable log_tbl = (DataTable)DGV_logfile.DataSource; 
                    
                   
                    List<string> pcs_empty = new List<string> { };
                    foreach (var item in sel_cells_lst)
                    {
                        string pcs_no = DGV_data_process.Rows[item.Key].Cells["PCS_No"].Value.ToString();
                        string remark = DGV_data_process.Rows[item.Key].Cells["Remark"].Value.ToString();
                        if (Image_result.ContainsKey(Convert.ToInt32(pcs_no)) && Image_result_2.ContainsKey(Convert.ToInt32(pcs_no)))
                        {
                            DGV_data_process.Rows[item.Key].Cells["Image_data"].Value = Image_result[Convert.ToInt32(pcs_no)];
                            DGV_data_process.Rows[item.Key].Cells["Image_data_2"].Value = Image_result_2[Convert.ToInt32(pcs_no)];
                            DGV_data_process.Rows[item.Key].Cells["Remark"].Value = src_path;
                            if (edited_lst.Keys.ToList().IndexOf(item.Key) == -1)
                            {
                                edited_lst.Add(item.Key, new List<string>() { "Image_data", "Image_data_2", "Remark" });
                            }
                            string post_no;
                            DataView temp = BVH_masterlist.AsDataView();
                            int count = 0;
                            for (int i = 4; i < DGV_data_process.ColumnCount - 3; i++)
                            {
                                DGV_data_process.Rows[item.Key].Cells[i].Value = "";
                            }

                            foreach (var itemname in item.Value)
                            {
                                temp.RowFilter = "ItemName = '" + itemname + "'";
                                int index = BVH_masterlist.Rows.IndexOf(temp[0].Row); 
                                if (temp.Count > 0)
                                {
                                    if (edited_lst[item.Key].IndexOf(itemname) == -1) 
                                    {
                                        edited_lst[item.Key].Add(itemname);
                                    }
                                    post_no = temp[0]["Post_No"].ToString().Replace(" ","").ToUpper();//.Split('_').LastOrDefault();
                                    if (post_no != "")
                                    {
                                        if (logfile_result.ContainsKey(Convert.ToInt32(pcs_no)))
                                        {
                                            if (logfile_result[Convert.ToInt32(pcs_no)].ContainsKey(post_no))
                                            {
                                                string replace_data = logfile_result[Convert.ToInt32(pcs_no)][post_no];
                                                DGV_data_process.Rows[item.Key].Cells[itemname].Value = replace_data;
                                            }
                                        }
                                    }


                                }
                                count++;
                            }
                        }
                        else
                        {
                            pcs_empty.Add(pcs_no);
                        }



                    }
                    if (pcs_empty.Count > 0)
                    {
                        string empty = "";
                        foreach (string item in pcs_empty)
                        {
                            empty += item + ";";
                        }
                        MessageBox.Show(new Form { TopMost = true }, "Logfile is not contains PCS: " + empty.TrimEnd(), "Warning");
                    }
                    Check_spec();

                    if (pcs_empty.Count == sel_cells_lst.Count)
                    {
                        MessageBox.Show(new Form { TopMost = true }, "No pcs are replaced", "Warning");
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Replace completed!", "Warning");
                        btn_Update.Enabled = true;
                    }


                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "No pcs are selected", "Warning");
                }

            }



        }

        private void btn_Update_Click(object sender, EventArgs e)
        {
            BVH_PTH_Data_Update_new();
            btn_Update.Enabled = false;

        }

        private void btnLogin_Click(object sender, EventArgs e)
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
                    MessageBox.Show(new Form { TopMost = true }, "Sai mật khẩu!");
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

        private void btnLoadData_Edit_Click(object sender, EventArgs e)
        {
            Search_Editted_Data();
        }

        private void cbProcess_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbProcess_Edit.SelectedIndex != -1 && txtItemCode_Edit.Text != "" && txtLotNo_Edit.Text != "")
            {
                Search_Editted_Data();
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
                DGV_Image_Edit.DataSource = dt_img;
                if (DGV_Image_Edit.Rows.Count > 0)
                {
                    ((DataGridViewImageColumn)DGV_Image_Edit.Columns["Before_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)DGV_Image_Edit.Columns["Before_Data"]).Width = 200;
                    ((DataGridViewImageColumn)DGV_Image_Edit.Columns["After_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)DGV_Image_Edit.Columns["After_Data"]).Width = 200;
                    foreach (DataGridViewRow dr in DGV_Image_Edit.Rows)
                    {
                        dr.Height = 150;
                    }

                }
            }
            else
            {
                DGV_DataEdit.DataSource = null;
                DGV_Image_Edit.DataSource = null;
                MessageBox.Show(new Form { TopMost = true }, "No data", "Warning");
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
            DataTable dt_img = src_dt.AsEnumerable().Where(x => x.Field<string>("Region") == "Image_data").CopyToDataTable();
            DataTable dt_data = src_dt.AsEnumerable().Where(x => x.Field<string>("Region") != "Image_data").CopyToDataTable();
            img_dt = dt_img.AsDataView().ToTable(false, new string[] { "PCS_No", "Before_Data", "After_Data", "Time_Update", "Operator", "Depart", "Remark" });
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
            data_dt = disp_dt.AsDataView().ToTable(false, new string[] { "PCS_No", "Region", "Before_Data", "After_Data", "Time_Update", "Operator", "Depart", "Remark" });
        }

        private void DGV_Image_Edit_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DGV_Image_Edit_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //int col_inx = e.ColumnIndex;
            int r_inx = e.RowIndex;
            //DataGridViewCell cur_cell = DGV_Image_Edit.CurrentCell;
            DataGridViewCell before_cell = DGV_Image_Edit.Rows[e.RowIndex].Cells["Before_Data"];
            DataGridViewCell after_cell = DGV_Image_Edit.Rows[e.RowIndex].Cells["After_Data"];

            byte[] before_data = (byte[])before_cell.Value;
            byte[] after_data = (byte[])after_cell.Value;

            if (myCode.checkDBNull(before_data) != "" && myCode.checkDBNull(after_data) != "")
            {

                Image_history fr1 = new Image_history();
                fr1.data_image_before_ = before_data;
                fr1.data_image_after_ = after_data;
                fr1.Show();
            }

            //if (myCode.checkDBNull(before_data) != "")
            //{
            //    using (MemoryStream ms = new MemoryStream(before_data))
            //    {
            //        pic_before.Image = Image.FromStream(ms);
            //    }
            //}
            //else
            //{
            //    pic_before.Image = null;
            //}

            //if (myCode.checkDBNull(after_data) != "")
            //{
            //    using (MemoryStream ms = new MemoryStream(after_data))
            //    {
            //        pic_after.Image = Image.FromStream(ms);
            //    }
            //}
            //else
            //{
            //    pic_after.Image = null;
            //}


        }

        private void pic_detail_view_DoubleClick(object sender, EventArgs e)
        {

        }

        private void DGV_data_view_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int col_inx = e.ColumnIndex;
            int r_inx = e.RowIndex;
            DataGridViewCell cur_cell = DGV_data_view.CurrentCell;
            if (DGV_data_view.Columns[col_inx].Name.Contains("Image"))
            {
                if (myCode.checkDBNull(cur_cell.Value) != "")
                {
                    byte[] data = (byte[])cur_cell.Value;
                    using (MemoryStream ms = new MemoryStream(data))
                    {
                        pic_detail_view.Image = Image.FromStream(ms);
                    }
                }
                else
                {
                    pic_detail_view.Image = null;
                }
                // byte[] data = (byte[])
                //if (data != null)
                //{
                //    using (MemoryStream ms = new MemoryStream(data))
                //    {
                //        pic_detail_view.Image = Image.FromStream(ms);
                //    }
                //}

            }
        }

        private void txtItemCode_Edit_TextChanged(object sender, EventArgs e)
        {
            DGV_DataEdit.DataSource = DGV_Image_Edit.DataSource = null;
            // pic_after.Image = pic_before.Image = null;
        }

        private void Cancel_Click(object sender, EventArgs e)
        {

        }

        private void txtItemCode_TextChanged(object sender, EventArgs e)
        {
            DGV_logfile.DataSource = DGV_data_process.DataSource = DGV_spec.DataSource = null;
            Image_result = null;
            Image_result_2 = null;

            edited_lst = null;
            sel_cells_lst = null;
            update_mode = false;
            insert_mode = false;
            lbl_title.Text = "Data Analysis";
        }

        private void txtLotNo_TextChanged(object sender, EventArgs e)
        {
            DGV_logfile.DataSource = DGV_data_process.DataSource = null;
            Image_result = null;
            Image_result_2 = null;

            edited_lst = null;
            sel_cells_lst = null;
            update_mode = false;
            insert_mode = false;
            lbl_title.Text = "Data Analysis";
        }

        private void txtLotNo_Edit_TextChanged(object sender, EventArgs e)
        {
            DGV_DataEdit.DataSource = DGV_Image_Edit.DataSource = null;
            // pic_after.Image = pic_before.Image = null;
        }

        private void txtLotNo_Validated(object sender, EventArgs e)
        {
            txtLotNo.Text = exp_proc.Lotno_Formated(txtLotNo.Text);
        }

        private void txt_lotno_view_Validated(object sender, EventArgs e)
        {
            txt_lotno_view.Text = exp_proc.Lotno_Formated(txt_lotno_view.Text);
        }

        private void txtLotNo_Edit_Validated(object sender, EventArgs e)
        {
            txtLotNo_Edit.Text = exp_proc.Lotno_Formated(txtLotNo_Edit.Text);
        }

        private void txt_masterlist_TextChanged(object sender, EventArgs e)
        {
            DGV_MasterList.DataSource = null;
        }

        //private void btn_Export_Click(object sender, EventArgs e)
        //{
        //    if (txtItemCode.Text != "" && txtLotNo.Text != "")
        //    {
        //        Export_BVH_PTH_data_Update(sqlcon, txtItemCode.Text, txtLotNo.Text);
        //        chk_list_vtime.Enabled = false;
        //        btn_Export.Enabled = false;
        //    }

        //}

        //private void btn_change_Click(object sender, EventArgs e)
        //{

        //    List<string> tar_table = new List<string>() { "BVH_WITH_BONDING_SHEET", "BVH_WITHOUT_BONDING_SHEET", "PLATED_THROUGH_HOLE" };

        //    foreach (var t in tar_table)
        //    {
        //        DataTable dt_all = TDMK_Code.Datatable_Filter(sqlcon, t, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
        //        List<DataTable> src_remark_tbl = new List<DataTable>();
        //        exp_proc.Get_ListTable(-1, dt_all, new string[] { "ItemCode", "LotNo", "Remark" }, ref src_remark_tbl, "PCS_No");
        //        bool chk_OK_time = false;
        //        foreach (var src_dt in src_remark_tbl)
        //        {
        //            string str_remark = src_dt.Rows[0]["Remark"].ToString();
        //            if (str_remark.Contains("-OK"))
        //            {
        //                DataTable dt_copy = src_dt.Clone();
        //                for (int j = 0; j < src_dt.Rows.Count; j++)
        //                {
        //                    DataRow dr = dt_copy.NewRow();
        //                    for (int i = 0; i < dt_copy.Columns.Count - 1; i++)
        //                    {z
        //                        dr[i] = src_dt.Rows[j][i];
        //                    }

        //                    dr["Remark"] = src_dt.Rows[j]["Remark"].ToString().Replace("-OK", string.Empty);
        //                    dt_copy.Rows.Add(dr);
        //                }
        //                exp_proc.BatchBulkCopy(sqlcon, dt_copy, t);
        //                break;
        //            }                 
        //        }
        //    }
        //    //chk_list_vtime.Enabled = true;
        //    MessageBox.Show("OK");
        //}

    }
}
