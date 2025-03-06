using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SEEV_DLL;
using TDMK_SQL;
using myExcel = Microsoft.Office.Interop.Excel;
using OK2SHIP;
using System.Data.SqlClient;
using System.Security.Cryptography;
using Microsoft.Office.Interop.Excel;
using System.Security.Policy;
using DataTable = System.Data.DataTable;

namespace VHX
{
    public partial class SolderMask : Form
    {
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SEI_Lib myCode = new SEI_Lib();
        myVar exp_proc = new myVar();
        Export_Image_Class diem = new Export_Image_Class();
        ExportEPPlus F_exportEPPlus = new ExportEPPlus();
        public bool admin_mode = false;

        public string depart = "QA";
        public SolderMask()
        {
            InitializeComponent();
        }

        private void SolderMask_Load(object sender, EventArgs e)
        {
            admin_mode = false;
        }
        public void SolderMask_Data_Process(string src_path)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && cb_region.SelectedIndex != -1)
            {
                Dictionary<string, SortedDictionary<int, byte[]>> dic_lst_result = new Dictionary<string, SortedDictionary<int, byte[]>>();
                Get_SolderMask_Image_Multi(src_path, ref dic_lst_result);
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                DataTable tar_dt = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SOLDERMASK_IMAGE", filter_str).Clone();
                int r_inx = 0;
                foreach (var lst_result in dic_lst_result)
                {
                    //string region = lst_result.Key;
                    string region = cb_region.SelectedItem.ToString();
                    foreach (var item in lst_result.Value)
                    {
                        string pcs = item.Key.ToString();
                        tar_dt.Rows.Add(r_inx + 1, txtItemCode.Text, txtLotNo.Text, region, pcs, item.Value, txtOperator.Text, DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"));
                        r_inx++;
                    }
                }
                DGV_Image_Graph.DataSource = tar_dt;

                ((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_Data"]).Width = 100;
                foreach (DataGridViewRow dr in DGV_Image_Graph.Rows)
                {
                    dr.Height = 70;
                }
            }
            else
            {
                MessageBox.Show("Please, fill in ItemCode / LotNo / Operator ID and select region", "Warning");
            }

        }
        public void Get_SolderMask_Image_Multi(string in_src, ref Dictionary<string, SortedDictionary<int, byte[]>> dic_lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            try
            {

                FileInfo[] temp_lst = tar_d.GetFiles("*.jpg");
                if (temp_lst.Length > 0)
                {
                    SortedDictionary<int, byte[]> lst_result = new SortedDictionary<int, byte[]>();
                    
                    for (int i = 0; i < temp_lst.Length; i++)
                    {
                        string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
                        if (!f_na.Contains("-"))
                        {
                            f_na = f_na.Replace("+", string.Empty);
                            if (myCode.IsNumeric(f_na))
                            {
                                var sel_img = Bitmap.FromFile(temp_lst[i].FullName);
                                ImageConverter imgcon = new ImageConverter();
                                byte[] img_data = (byte[])imgcon.ConvertTo(sel_img, typeof(byte[]));
                                int f_inx = Convert.ToInt32(f_na);
                                lst_result.Add(f_inx, img_data);
                            }

                        }
                    }
                    dic_lst_result.Add(tar_d.Name, lst_result);
                }
                else
                {
                    DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                    if (inter_type_lst.Length > 0)
                    {
                        foreach (var inter_lst in inter_type_lst)
                        {
                            Get_SolderMask_Image_Multi(inter_lst.FullName, ref dic_lst_result);
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("Could not find a part of the path " + in_src, "Warning");
            }
        }

        private void DGV_Image_Graph_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int col_inx = e.ColumnIndex;
            int r_inx = e.RowIndex;
            DataGridViewCell cur_cell = DGV_Image_Graph.CurrentCell;
            if (DGV_Image_Graph.Columns[col_inx].Name.Contains("Image"))
            {
                byte[] data = (byte[])cur_cell.Value;
                using (MemoryStream ms = new MemoryStream(data))
                {
                    picDetails.Image = Image.FromStream(ms);
                }
                string region = DGV_Image_Graph.Rows[r_inx].Cells["Region"].Value.ToString();
                string pcs = DGV_Image_Graph.Rows[r_inx].Cells["Pcs_No"].Value.ToString();
                lblImage_Graph.Text = "Image / Graph Details ---> Region: " + region + " / " + "Pcs_No: " + pcs;
            }
        }

        private void tsmLoadData_Click(object sender, EventArgs e)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
            DataTable cur_dt = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SOLDERMASK_IMAGE", filter_str);
            DGV_Image_Graph.DataSource = cur_dt;

            ((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            ((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_Data"]).Width = 100;
            foreach (DataGridViewRow dr in DGV_Image_Graph.Rows)
            {
                dr.Height = 70;
            }
            lblImage_Graph.Text = "Image / Graph Details";
        }

        public void Update_history()
        {

            DataTable image_before = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SOLDERMASK_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));

            List<string> item_arr = new List<string>() { "ID", "ItemCode", "LotNo", "Process", "PCS_No", "Region", "Zone", "Before_Data", "After_Data", "Time_Update", "Operator", "Depart", "Remark" };
            List<object> item_arr_val = new List<object>();
            DateTime tim_up = DateTime.Now;
            int id_edit = TDMK_Code.SQL_MAX("VHX_Edit_History", "ID", myVar.mysqlcon);
            int count = 0;
            DataTable logfile_after = (DataTable)DGV_Image_Graph.DataSource;
            DataTable image_after = (DataTable)DGV_Image_Graph.DataSource;

            int min_pcs = new int[] { image_before.Rows.Count, image_after.Rows.Count }.Min();
            for (int inx = 0; inx < min_pcs; inx++)
            {



                byte[] img_bef = (byte[])(image_before.Rows[inx]["Image_Data"]);
                // byte[] img_bef = null;

                byte[] img_aft = (byte[])(image_after.Rows[inx]["Image_Data"]);
                string pcs = image_after.Rows[inx]["Pcs_No"].ToString();
                string region = image_after.Rows[inx]["Region"].ToString();
                string remark1 = txtLogfile.Text;
                item_arr_val = new List<object>() { (id_edit + count + 1), txtItemCode.Text, txtLotNo.Text, "SolderMask", pcs, "Image_data", region, img_bef, img_aft, tim_up, txtUsername.Text, depart, remark1 };
                exp_proc.Insert_Object_List("VHX_Edit_History", myVar.mysqlcon, item_arr, item_arr_val);
                count++;
            }

        }

        private void tsmSaveData_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtOperator.Text != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
            start_lbl: DataTable cur_dt = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SOLDERMASK_IMAGE", filter_str);
                if (cur_dt.Rows.Count > 0)
                {
                    if (MessageBox.Show("Table " + "SOLDERMASK_IMAGE" + ": Data is existed. Do you want to replace?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        if (admin_mode)
                        {
                            Update_history();
                            TDMK_Code.Delelte_FilteredItem_arr("SOLDERMASK_IMAGE", myVar.mysqlcon, filter_str);
                            goto start_lbl;
                        }
                        else
                        {
                            MessageBox.Show("Please, login to save data", "Warning");
                        }

                    }
                }
                else
                {
                    if (DGV_Image_Graph.Rows.Count > 0)
                    {


                        int id = TDMK_Code.SQL_MAX("SOLDERMASK_IMAGE", "ID", myVar.mysqlcon);
                        int r_inx = 0;
                        foreach (DataRow dr in cur_dt.Rows)
                        {
                            dr["ID"] = id + 1 + r_inx;
                            r_inx++;
                        }
                        exp_proc.BatchBulkCopy(myVar.mysqlcon, (DataTable)DGV_Image_Graph.DataSource, "SOLDERMASK_IMAGE");
                        MessageBox.Show("Save data completed!", "Warning");
                    }
                    else
                    {
                        MessageBox.Show("No data!", "Warning");
                    }
                }
            }
            else
            {
                MessageBox.Show("Please, fill in ItemCode / LotNo / Operator ID", "Warning"); 
            }
        }

        private void txtLogfile_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                DataTable cur_dt = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SOLDERMASK_IMAGE", filter_str);
                //((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                //((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_Data"]).Width = 100;
                //foreach (DataGridViewRow dr in DGV_Image_Graph.Rows)
                //{
                //    dr.Height = 70;
                //}

                bool check_data = true;
                if (cur_dt.Rows.Count > 0)
                {
                    check_data = false;
                    DGV_Image_Graph.DataSource = cur_dt;
                    if (DGV_Image_Graph.Rows.Count > 0)
                    {
                        ((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                        ((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_Data"]).Width = 100;
                        foreach (DataGridViewRow dr in DGV_Image_Graph.Rows)
                        {
                            dr.Height = 70;
                        }
                    }
                    if (MessageBox.Show("Table SOLDERMASK_IMAGE : Data is existed. Do you want to update?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        check_data = true;
                    }

                    if (check_data)
                    {
                        if (admin_mode)
                        {
                            load_data_logfile_soldermask();
                        }
                        else
                        {
                            MessageBox.Show("Please, login to update data", "Warning");
                        }
                    }
                }
                else
                {
                    load_data_logfile_soldermask();
                }
            }
            else
            {
                MessageBox.Show("Please, fill in ItemCode / LotNo / Operator ID", "Warning");
            }
        }


        public void load_data_logfile_soldermask()
        {
            FolderBrowserDialog f_open = new FolderBrowserDialog();
            f_open.SelectedPath = System.Windows.Forms.Application.StartupPath;
            if (txtLogfile.Text != "")
            {
                f_open.SelectedPath = txtLogfile.Text;
            }
            if (f_open.ShowDialog() == DialogResult.OK)
            {
                txtLogfile.Text = f_open.SelectedPath;

                DirectoryInfo temp_dir = new DirectoryInfo(txtLogfile.Text);
                string f_name = temp_dir.Name;
                bool chk_name = false;


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


                if (chk_name)
                {

                    SolderMask_Data_Process(f_open.SelectedPath);
                }
                else
                {
                    MessageBox.Show("ItemCode or LotNo not matched", "Warning");
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

        private void tsmExport_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtOperator.Text != "")
            {
                string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                if (format_file != "")
                {
                    string export_path = F_exportEPPlus.get_reportpath(txtItemCode.Text, txtLotNo.Text, txtOperator.Text, "Solder Mask", myVar.data_loc);
                    if (export_path != "")
                    {
                        F_exportEPPlus.Export_EPPlus(myVar.mysqlcon, format_file, export_path, txtItemCode.Text, txtLotNo.Text, "SOLDERMASK", null);
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Folder Solder Mask not found!", "Warning");
                    }
                }

            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Please, fill in ItemCode / LotNo / Operator ", "Warning");
            }


        }

        public void Export_Solder_Image_1(SqlConnection sqlcon, myExcel.Worksheet ws, string Itemcode, string Lotno)
        {
            DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "SOLDERMASK_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { Itemcode, Lotno }));

            string[] xx = new string[4] { "BGA", "Hot Bar", "Connector", "Trace to trace" };
            bool check = false;
            foreach (DataRow dr in dt.Rows)
            {
                foreach (string s in xx)
                {
                    if (dr["Region"].ToString() == s)
                    {
                        check = true;
                    }
                }
                if (!check)
                {
                    dr["Region"] = "Connector";
                }
            }

            int row_val;
            foreach (string zone in xx)
            {
                for (int x = 1; x < 40; x++)
                {
                    if (ws.Cells[x, 1].Value != null && ws.Cells[x, 1].Value.ToString() == zone)
                    {
                        row_val = (int)x;
                        int number_image = 5;
                        string[] item = { "ItemCode", "LotNo", "Region" };
                        string[] item_val = { Itemcode, Lotno, zone };
                        //System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "SOLDERMASK_IMAGE", TDMK_Code.filter_str(item, item_val));
                        DataTable data_image = dt.Clone();
                        foreach (DataRow dr in dt.Rows)
                        {
                            if (dr["Region"].ToString() == zone)
                            {
                                DataRow dtrow = data_image.NewRow();
                                for (int i = 0; i < dt.Columns.Count; i++)
                                {
                                    dtrow[i] = dr[i];
                                }
                                data_image.Rows.Add(dtrow);
                            }

                        }
                        if (data_image.Rows.Count > 0)
                        {
                            if (data_image.Rows.Count < number_image)
                            {
                                number_image = data_image.Rows.Count;

                                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");

                                string file_dic = Path.Combine(file_folder, "1.jpg");

                                myExcel.Range curr_rgn = ws.Range["B" + row_val];
                                for (int m = 0; m < number_image; m++)
                                {
                                    byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                    MemoryStream ms = new MemoryStream(data);
                                    Image image = Image.FromStream(ms);
                                    image.Save(file_dic);
                                    diem.InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                                    curr_rgn = curr_rgn.Offset[0, 1];
                                }
                            }
                            else
                            {
                                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                                string file_dic = Path.Combine(file_folder, "1.jpg");

                                myExcel.Range curr_rgn = ws.Range["B" + row_val];
                                for (int m = 0; m < number_image; m++)
                                {
                                    byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                    MemoryStream ms = new MemoryStream(data);
                                    Image image = Image.FromStream(ms);
                                    image.Save(file_dic);
                                    diem.InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                                    curr_rgn = curr_rgn.Offset[0, 1];
                                }
                            }
                        }
                    }
                }

            }
        }
        public void Export_Solder_Image_Update(SqlConnection sqlcon, myExcel.Worksheet ws, string Itemcode, string Lotno)
        {

            int row_val;
            string zone = cb_region.SelectedItem.ToString();
            for (int x = 1; x < 40; x++)
            {
                if (myCode.checkDBNull(ws.Cells[x, 1].Value) == zone)
                {
                    row_val = (int)x;
                    int number_image = 5;
                    string[] item = { "ItemCode", "LotNo", "Region" };
                    string[] item_val = { Itemcode, Lotno, zone };
                    DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "SOLDERMASK_IMAGE", TDMK_Code.filter_str(item, item_val));


                    if (data_image.Rows.Count > 0)
                    {
                        if (data_image.Rows.Count < number_image)
                        {
                            number_image = data_image.Rows.Count;

                            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");

                            string file_dic = Path.Combine(file_folder, "1.jpg");

                            myExcel.Range curr_rgn = ws.Range["B" + row_val];
                            for (int m = 0; m < number_image; m++)
                            {
                                byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                MemoryStream ms = new MemoryStream(data);
                                Image image = Image.FromStream(ms);
                                image.Save(file_dic);
                                diem.InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                                curr_rgn = curr_rgn.Offset[0, 1];
                            }
                        }
                        else
                        {
                            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                            string file_dic = Path.Combine(file_folder, "1.jpg");

                            myExcel.Range curr_rgn = ws.Range["B" + row_val];
                            for (int m = 0; m < number_image; m++)
                            {
                                byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                MemoryStream ms = new MemoryStream(data);
                                Image image = Image.FromStream(ms);
                                image.Save(file_dic);
                                diem.InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                                curr_rgn = curr_rgn.Offset[0, 1];
                            }
                        }
                    }
                } 
            }

        }
        public void Export_LPI_data(myExcel.Worksheet ws, string ItemCode, string LotNo)
        {
            string connstr_IPQC = TDMK_Code.data_connection(myVar.server_name, "IPQC_Data", myVar.server_acc, myVar.server_pass).ConnectionString;
            SqlConnection sqlcon_IPQC = new SqlConnection(connstr_IPQC);
            DataTable LPI_table = TDMK_Code.Datatable_Filter(sqlcon_IPQC, "Printing_Process", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
            List<string> col_name = new List<string>() { "LPI_SM_Coverage", "LPI_Open_B1", "LPI_SM_L1" };
            DataTable LPI_data = LPI_table.AsDataView().ToTable(false, col_name.ToArray());
            int qty = Math.Min(LPI_data.Rows.Count, 5);
            List<string> tar_rgn = new List<string>() { "H13", "L13", "M13" };
            for (int i = 0; i < qty; i++)
            {
                for (int j = 0; j < tar_rgn.Count; j++)
                {
                    ws.Range[tar_rgn[j]].Offset[i, 0].Value = LPI_data.Rows[i][j].ToString();
                }
            }
        }
        private void txtLogfile_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (txtLogfile.Text != "")
                {
                    if (txtItemCode.Text != "" && txtLotNo.Text != "")
                    {
                        string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                        DataTable cur_dt = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SOLDERMASK_IMAGE", filter_str);
                     
                       

                        bool check_data = true;
                        if (cur_dt.Rows.Count > 0)
                        {
                            check_data = false;
                            DGV_Image_Graph.DataSource = cur_dt;
                            if (DGV_Image_Graph.Rows.Count > 0)
                            {
                                ((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                                ((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_Data"]).Width = 100;
                                foreach (DataGridViewRow dr in DGV_Image_Graph.Rows)
                                {
                                    dr.Height = 70;
                                }
                            }
                            if (MessageBox.Show("Table SOLDERMASK_IMAGE : Data is existed. Do you want to update?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                            {
                                check_data = true;
                            }

                            if (check_data)
                            {
                                if (admin_mode)
                                {
                                    DirectoryInfo temp_dir = new DirectoryInfo(txtLogfile.Text);
                                    string f_name = temp_dir.Name;
                                    bool chk_name = false;

                                    if (f_name.Split('-')[0].Trim() == txtItemCode.Text)
                                    {
                                        if (txtLotNo.Text.Contains("-") && f_name.Split('-').Length == 3)
                                        {
                                            if (Convert.ToDecimal(f_name.Split('-')[1]) == Convert.ToDecimal(txtLotNo.Text.ToString().Split('-')[0]) && Convert.ToDecimal(f_name.Split('-')[2].Split('_')[0]) == Convert.ToDecimal(txtLotNo.Text.ToString().Split('-')[1]))
                                                chk_name = true;
                                        }
                                        else if (!txtLotNo.Text.Contains("-") && f_name.Split('-').Length == 2)
                                        {
                                            if (Convert.ToDecimal(f_name.Split('-')[1].Split('_')[0].Trim()) == Convert.ToDecimal(txtLotNo.Text.ToString()))
                                            {
                                                chk_name = true;
                                            }
                                        }
                                    }

                                    if (chk_name)
                                    {

                                        SolderMask_Data_Process(txtLogfile.Text);
                                    }
                                    else
                                    {
                                        MessageBox.Show("ItemCode or LotNo not matched", "Warning");
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Please, login to update data", "Warning");
                                }
                            }
                        }
                        else
                        {
                            DirectoryInfo temp_dir = new DirectoryInfo(txtLogfile.Text);
                            string f_name = temp_dir.Name;
                            bool chk_name = false;

                            if (f_name.Split('-')[0].Trim() == txtItemCode.Text)
                            {
                                if (txtLotNo.Text.Contains("-") && f_name.Split('-').Length == 3)
                                {
                                    if (Convert.ToDecimal(f_name.Split('-')[1]) == Convert.ToDecimal(txtLotNo.Text.ToString().Split('-')[0]) && Convert.ToDecimal(f_name.Split('-')[2].Split('_')[0]) == Convert.ToDecimal(txtLotNo.Text.ToString().Split('-')[1]))
                                        chk_name = true;
                                }
                                else if (!txtLotNo.Text.Contains("-") && f_name.Split('-').Length == 2)
                                {
                                    if (Convert.ToDecimal(f_name.Split('-')[1].Split('_')[0].Trim()) == Convert.ToDecimal(txtLotNo.Text.ToString()))
                                    {
                                        chk_name = true;
                                    }
                                }
                            }

                            if (chk_name)
                            {

                                SolderMask_Data_Process(txtLogfile.Text);
                            }
                            else
                            {
                                MessageBox.Show("ItemCode or LotNo not matched", "Warning");
                            }
                        }

                    }
                    else
                    {
                        MessageBox.Show("Please, fill in ItemCode / LotNo / Operator ID", "Warning");
                    }

                }

            }

        } 

        private void txtLogfile_TextChanged(object sender, EventArgs e)
        {

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

        private void cbProcess_Edit_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void btn_load_edit_Click(object sender, EventArgs e)
        {

        }

        private void cbProcess_Edit_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void DGV_Image_Edit_CellContentDoubleClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btn_load_edit_Click_1(object sender, EventArgs e)
        {
            if (txtItemCode_Edit.Text != "" && txtLotNo_Edit.Text != "" )
            {
                Search_Editted_Data();
            }
        }
        public void Search_Editted_Data()
        {
            DataTable dt = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { txtItemCode_Edit.Text, txtLotNo_Edit.Text, "SolderMask" }));
            if (dt.Rows.Count > 0)
            {


                DataTable dt_img = new DataTable();
                Split_Data_Image(dt, ref dt_img);

                DGV_Image_Edit.DataSource = dt_img;
                if (DGV_Image_Edit.Rows.Count > 0)
                {
                    ((DataGridViewImageColumn)DGV_Image_Edit.Columns["Before_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)DGV_Image_Edit.Columns["Before_Data"]).Width = 100;
                    ((DataGridViewImageColumn)DGV_Image_Edit.Columns["After_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)DGV_Image_Edit.Columns["After_Data"]).Width = 100;
                    foreach (DataGridViewRow dr in DGV_Image_Edit.Rows)
                    {
                        dr.Height = 70;
                    }

                }
            }
            else
            {

                DGV_Image_Edit.DataSource = null;
                MessageBox.Show("No data", "Warning");
            }
            
        
        }
        public void Split_Data_Image(DataTable src_dt, ref DataTable img_dt)
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

                img_dt = dt_img.AsDataView().ToTable(false, new string[] { "PCS_No", "Before_Data", "After_Data", "Time_Update", "Operator", "Depart", "Remark" });
           
          
        }

        private void DGV_Image_Edit_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int r_inx = e.RowIndex;
            //DataGridViewCell cur_cell = DGV_Image_Edit.CurrentCell;
            DataGridViewCell before_cell = DGV_Image_Edit.Rows[e.RowIndex].Cells["Before_Data"];
            DataGridViewCell after_cell = DGV_Image_Edit.Rows[e.RowIndex].Cells["After_Data"];
            //if (DGV_Image.Columns[col_inx].Name.Contains("Image"))
            //{
            byte[] before_data = (byte[])before_cell.Value;
            byte[] after_data = (byte[])after_cell.Value;
            if (before_data != null && after_data != null)
            {
                using (MemoryStream ms = new MemoryStream(before_data))
                {
                    pic_before.Image = Image.FromStream(ms);
                }
                using (MemoryStream ms = new MemoryStream(after_data))
                {
                    pic_after.Image = Image.FromStream(ms);
                }
            }
        }

        private void txtItemCode_TextChanged(object sender, EventArgs e)
        {
            DGV_Image_Graph.DataSource = null;
            picDetails.Image = null;
        }

        private void txtLotNo_TextChanged(object sender, EventArgs e)
        {
            DGV_Image_Graph.DataSource = null;
            picDetails.Image = null;
        }

        private void txtItemCode_Edit_TextChanged(object sender, EventArgs e)
        {
            DGV_Image_Edit.DataSource = null;
            pic_after.Image = pic_before.Image = null;
        }

        private void txtLotNo_Edit_TextChanged(object sender, EventArgs e)
        {
            DGV_Image_Edit.DataSource = null;
            pic_after.Image = pic_before.Image = null;
        }

        private void txtLotNo_Validated(object sender, EventArgs e)
        {
            txtLotNo.Text = exp_proc.Lotno_Formated(txtLotNo.Text);
        }

        private void txtLotNo_Edit_Validated(object sender, EventArgs e)
        {
            txtLotNo_Edit.Text = exp_proc.Lotno_Formated(txtLotNo_Edit.Text);
        }
    }
}
