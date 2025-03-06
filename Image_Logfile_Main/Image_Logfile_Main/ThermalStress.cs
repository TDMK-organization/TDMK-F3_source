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
using System.Data.SqlClient;
using myExcel = Microsoft.Office.Interop.Excel;
using OK2SHIP;
using System.Diagnostics;
using System.Web;


namespace VHX
{
    public partial class ThermalStress : Form
    {
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SEI_Lib myCode = new SEI_Lib();
        myVar exp_proc = new myVar();
        Export_Image_Class diem = new Export_Image_Class();
        public bool admin_mode = false;
        public string depart = "QA";
        ExportEPPlus F_exportEPPlus = new ExportEPPlus();
        public ThermalStress()
        {
            InitializeComponent();
        }
        public int row_input;

        private void txtLogfile_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
            DataTable cur_dt = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "THERMAL_STRESS_IMAGE", filter_str);
            bool check_data = true;
            if (cur_dt.Rows.Count > 0)
            {
                check_data = false;
                DGV_Image_Graph.DataSource = cur_dt;
                ((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_Data"]).Width = 100;
                foreach (DataGridViewRow dr in DGV_Image_Graph.Rows)
                {
                    dr.Height = 70;
                }

                if (MessageBox.Show("Table THERMAL_STRESS_IMAGE : Data is existed. Do you want to update?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    check_data = true;
                }

                if (check_data)
                {
                    if (admin_mode)
                    {

                        FolderBrowserDialog f_open = new FolderBrowserDialog();
                        f_open.SelectedPath = Application.StartupPath;
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
                                ThermalStress_Data_Process(txtLogfile.Text);
                            }
                            else
                            {
                                MessageBox.Show("ItemCode or LotNo not matched", "Warning");
                            }
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
                FolderBrowserDialog f_open = new FolderBrowserDialog();
                f_open.SelectedPath = Application.StartupPath;
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
                        ThermalStress_Data_Process(txtLogfile.Text);
                    }
                    else
                    {
                        MessageBox.Show("ItemCode or LotNo not matched", "Warning");
                    }
                }


            }
        }
        public void ThermalStress_Data_Process(string src_path)
        {
            Dictionary<string, SortedDictionary<int, byte[]>> dic_lst_result = new Dictionary<string, SortedDictionary<int, byte[]>>();
            Get_ThermalStress_Image_Multi(src_path, ref dic_lst_result);
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
            DataTable tar_dt = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "THERMAL_STRESS_IMAGE", filter_str).Clone();
            int r_inx = 0;
            foreach (var lst_result in dic_lst_result)
            {
                string region = lst_result.Key;
                foreach (var item in lst_result.Value)
                {
                    string pcs = item.Key.ToString();
                    tar_dt.Rows.Add(r_inx + 1, txtItemCode.Text, txtLotNo.Text, region, pcs, item.Value, txtOperator.Text, DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), src_path);
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
         
        public void Get_ThermalStress_Image_Multi(string in_src, ref Dictionary<string, SortedDictionary<int, byte[]>> dic_lst_result)
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
                        string[] rmv = { "+", "-" };

                        f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                        if (myCode.IsNumeric(f_na))
                        {
                            var sel_img = Bitmap.FromFile(temp_lst[i].FullName);
                            ImageConverter imgcon = new ImageConverter();
                            byte[] img_data = (byte[])imgcon.ConvertTo(sel_img, typeof(byte[]));
                            int f_inx = Convert.ToInt32(f_na);
                            if (!lst_result.ContainsKey(f_inx))
                            {
                                lst_result.Add(f_inx, img_data);
                            }
                        }
                    }
                    dic_lst_result.Add(tar_d.Name.Replace(" ", "").ToUpper(), lst_result);
                }
                else
                {
                    DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                    if (inter_type_lst.Length > 0)
                    {
                        foreach (var inter_lst in inter_type_lst)
                        {
                            Get_ThermalStress_Image_Multi(inter_lst.FullName, ref dic_lst_result);
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("Could not find a part of the path " + in_src, "Warning");
            }

        }
        public void setup_qty_ThermalStress(string file_format, string ItemCode)
        {
            myExcel.Workbook wb = TDMK_Code.open_excel_file(file_format, "", "");
            string sheet_name = "";
            string tar_sheet = "CQRA - Thermal stress";
            for (int i = 0; i < wb.Sheets.Count; i++)
            {
                myExcel.Worksheet cur_sht = wb.Sheets[i + 1];
                string cur_sht_name = cur_sht.Name;
                if (tar_sheet.Replace("-", "").Replace(" ", "").ToUpper() == cur_sht_name.Replace("-", "").Replace(" ", "").ToUpper())
                {
                    sheet_name = cur_sht_name;
                    break;
                }
            }
            if (sheet_name != "")
            {
                int count_BVH = 0;
                int count_PTH = 0;

                for (int row = 2; row <= 50; row++)
                {
                    myExcel.Range cell = wb.Sheets[sheet_name].Cells[row, 1];
                    if (myCode.checkDBNull(cell.Value).ToUpper().Contains("BVH"))
                    {

                        count_BVH++;
                    }
                    if (myCode.checkDBNull(cell.Value).ToUpper().Contains("PTH"))
                    {
                        count_PTH++;
                    }
                }
                int ID = TDMK_Code.SQL_MAX("SETTING_PCS", "ID", myVar.mysqlcon) + 1;
                TDMK_Code.insert_val_arr("SETTING_PCS", myVar.mysqlcon, new string[] { "ID", "ItemCode", "Data_For", "Region_Zone", "Pcs_setting" }, new string[] { ID.ToString(), ItemCode, "CQRA - Thermal stress", "", "BVH:" + count_BVH.ToString() + ";PTH:" + count_PTH.ToString() });
                MessageBox.Show("Setup qty completed!", "Waring");

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
            DataTable cur_dt = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "THERMAL_STRESS_IMAGE", filter_str);
            DGV_Image_Graph.DataSource = cur_dt;
            if (DGV_Image_Graph.Rows.Count > 0)
            {
                ((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_Data"]).Width = 100;
                foreach (DataGridViewRow dr in DGV_Image_Graph.Rows)
                {
                    dr.Height = 70;
                }

                lblImage_Graph.Text = "Image / Graph Details";
            }
            else
            {
                MessageBox.Show("No data", "Warning");
            }
        }

        private void tsmSaveData_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtOperator.Text != "")
            {
                DataTable dt_setting = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SETTING_PCS", TDMK_Code.filter_str(new string[] { "ItemCode", "Data_For" }, new string[] { txtItemCode.Text, "CQRA - Thermal stress" }));
                if (dt_setting.Rows.Count > 0)
                {
                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                start_lbl: DataTable cur_dt = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "THERMAL_STRESS_IMAGE", filter_str);
                    if (cur_dt.Rows.Count > 0)
                    {
                        if (admin_mode)
                        {
                            Update_history();
                            TDMK_Code.Delelte_FilteredItem_arr("THERMAL_STRESS_IMAGE", myVar.mysqlcon, filter_str);
                            goto start_lbl;
                        }
                        else
                        {
                            MessageBox.Show("Please, login to save data", "Warning");
                        }
                    }
                    else
                    {
                        if (DGV_Image_Graph.Rows.Count > 0)
                        {
                            int id = TDMK_Code.SQL_MAX("THERMAL_STRESS_IMAGE", "ID", myVar.mysqlcon);
                            int r_inx = 0;
                            foreach (DataRow dr in cur_dt.Rows)
                            {
                                dr["ID"] = id + 1 + r_inx;
                                r_inx++;
                            }
                            exp_proc.BatchBulkCopy(myVar.mysqlcon, (DataTable)DGV_Image_Graph.DataSource, "THERMAL_STRESS_IMAGE");
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
                    MessageBox.Show("Please, setup qty before save data", "Warning");
                }
            }
            else
            {
                MessageBox.Show("Please, fill in ItemCode / LotNo / Operator ID", "Warning");
            }
        }
        public AutoCompleteStringCollection Get_infor_to_input_thermalstress(myExcel.Workbook wb)
        {
            myExcel.Worksheet ws = wb.Sheets[1];
            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 10; i < 50; i++)
            {
                if (ws.Cells[i, 1].Value != null)
                {
                    if (ws.Cells[i, 1].Value.Contains("PTH"))
                    {
                        list.Add(ws.Cells[i, 1].Value.ToString() + "+" + (i).ToString());
                    }
                    if (ws.Cells[i, 1].Value.Contains("BVH"))
                    {
                        list.Add(ws.Cells[i, 1].Value.ToString() + "+" + (i).ToString());
                    }
                }

            }
            return list;
        }

        private void tsmExport_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtOperator.Text != "")
            {
                 
                string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                if (format_file != "")
                {
                    string export_path = F_exportEPPlus.get_reportpath(txtItemCode.Text, txtLotNo.Text, txtOperator.Text, "CQRA - Thermal Stress", myVar.data_loc);
                    if (export_path != "")
                    {
                        F_exportEPPlus.Export_EPPlus(myVar.mysqlcon, format_file, export_path, txtItemCode.Text, txtLotNo.Text, "CQRA - THERMAL STRESS", null);
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Folder CQRA - Thermal stress not found!", "Warning");
                    }
                } 
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Please, fill in ItemCode / LotNo / Operator ID", "Warning");
            }



        }

        public void Export_Image_BVH(myExcel.Worksheet ws, System.Data.DataTable data_image, string row_val, int no_BVH)
        {
            int number_image = no_BVH * 5 + 5;
            if (data_image.Rows.Count < number_image)
            {
                number_image = data_image.Rows.Count - no_BVH * 5;

                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                string file_dic = Path.Combine(file_folder, "1.jpg");

                myExcel.Range curr_rgn = ws.Range["B" + row_val];
                for (int m = no_BVH * 5; m < number_image; m++)
                {
                    byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                    MemoryStream ms = new MemoryStream(data);
                    Image image = Image.FromStream(ms);
                    image.Save(file_dic);
                    diem.InsertPicture_Name(ws, curr_rgn, file_dic, 6);
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
        public void Export_Image_(myExcel.Worksheet ws, int number_image, System.Data.DataTable data_image, string row_val)
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
                    diem.InsertPicture_Name(ws, curr_rgn, file_dic, 6);
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

        public void Export_Image_new(myExcel.Worksheet ws, int number_image, System.Data.DataTable data_image, string row_val)
        {
            int count_sample = new int[] { number_image, data_image.Rows.Count }.Min();

            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = Path.Combine(file_folder, "1.jpg");

            myExcel.Range curr_rgn = ws.Range["B" + row_val];
            for (int m = 0; m < count_sample; m++)
            {
                byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                MemoryStream ms = new MemoryStream(data);
                Image image = Image.FromStream(ms);
                image.Save(file_dic);
                diem.InsertPicture_Name(ws, curr_rgn, file_dic, 6);
                curr_rgn = curr_rgn.Offset[0, 1];
            }
        }
        public void Export_Thermal_Stress_Image_old(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno, string Operator)
        {
            AutoCompleteStringCollection list = Get_infor_to_input_thermalstress(wb);
            myExcel.Worksheet ws = wb.Sheets[1];
            int no_BVH = 0;
            foreach (string x in list)
            {
                string region = x.Split('+')[0];
                region = region.Split(' ')[0];
                string row_val = x.Split('+')[1];
                int number_image = 5;
                string[] item = { "ItemCode", "LotNo", "Region" };
                string[] item_val = { Itemcode, Lotno, region };
                System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "THERMAL_STRESS_IMAGE", TDMK_Code.filter_str(item, item_val));
                if (region.Contains("BVH"))
                {
                    Export_Image_BVH(ws, data_image, row_val, no_BVH);
                    no_BVH++;
                }
                else
                {
                    if (data_image.Rows.Count > 0)
                    {
                        Export_Image_(ws, number_image, data_image, row_val);
                    }
                }
            }

            save_report_new(txtItemCode.Text, txtLotNo.Text, txtOperator.Text, wb, "CQRA - Thermal stress", myVar.data_loc);
            MessageBox.Show("Export to Checksheet completed!", "Warning");

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

        public void Export_Thermal_Stress_Image(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno, string Operator) //update14/3/2024
        {
            AutoCompleteStringCollection list = Get_infor_to_input_thermalstress(wb);
            myExcel.Worksheet ws = wb.Sheets[1];
            DataTable Data_tbl = (DataTable)DGV_Image_Graph.DataSource;

            List<DataTable> lst_tbl = new List<DataTable> { };
            Get_ListTable(-1, Data_tbl, new string[] {"Region"}, ref lst_tbl, "ItemCode");
            int i = 0;
            foreach (string x in list)
            {
                string region = x.Split('+')[0].ToUpper().Replace(" ", "").Replace("PHOTO", "").Replace("PICTURE", "");
                string row_val = x.Split('+')[1];
                //DataTable dt_setting = TDMK_Code.Datatable_Filter(sqlcon, "SETTING_PCS", TDMK_Code.filter_str(new string[] { "ItemCode", "Data_For" }, new string[] { txtItemCode.Text, "CQRA - Thermal stress" }));

                int number_image = 5;
                if(i < lst_tbl.Count)
                {
                    Export_Image_new(ws, number_image, lst_tbl[i], row_val);
                } 
            } 
            save_report_new(txtItemCode.Text, txtLotNo.Text, txtOperator.Text, wb, "CQRA - Thermal stress", myVar.data_loc);
            MessageBox.Show(new Form { TopMost = true }, "Export to checksheet completed!", "Warning");

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
                    try
                    {
                        if (File.Exists(report_path))
                            File.Delete(report_path);
                    }
                    catch
                    {

                    }
                    curr_wrkbook.SaveAs(report_path);
                    break;
                }
            }
        }
        public void Export_Image_old(myExcel.Worksheet ws, int number_image, System.Data.DataTable data_image, int row_input)
        {
            if (data_image.Rows.Count < number_image)
            {
                number_image = data_image.Rows.Count;

                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                string file_dic = Path.Combine(file_folder, "1.jpg");

                myExcel.Range curr_rgn = ws.Range["B" + row_input];
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

                myExcel.Range curr_rgn = ws.Range["B" + row_input];
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
        public void Update_history()
        {

            DataTable image_before = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "THERMAL_STRESS_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));

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
                string remark1 = image_after.Rows[inx]["Remark"].ToString();
                item_arr_val = new List<object>() { (id_edit + count + 1), txtItemCode.Text, txtLotNo.Text, "ThermalStress", pcs, "Image_data", region, img_bef, img_aft, tim_up, txtUsername.Text, depart, remark1 };
                exp_proc.Insert_Object_List("VHX_Edit_History", myVar.mysqlcon, item_arr, item_arr_val);
                count++;
            }


        }



        private void ThermalStress_Load(object sender, EventArgs e)
        {
            admin_mode = false;
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

        }

        private void rbQA_CheckedChanged_1(object sender, EventArgs e)
        {

        }

        private void rbDE_CheckedChanged_1(object sender, EventArgs e)
        {

        }

        private void rbNPI_CheckedChanged_1(object sender, EventArgs e)
        {

        }

        private void rbPro_CheckedChanged_1(object sender, EventArgs e)
        {

        }

        private void rbPE_CheckedChanged_1(object sender, EventArgs e)
        {

        }

        private void btn_load_edit_Click_1(object sender, EventArgs e)
        {

        }

        private void DGV_Image_Edit_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DGV_Image_Graph_MouseDown(object sender, MouseEventArgs e)
        {


        }

        private void btn_load_edit_Click(object sender, EventArgs e)
        {
            if (txtItemCode_Edit.Text != "" && txtLotNo_Edit.Text != "")
            {
                Search_Editted_Data();
            }
        }
        public void Search_Editted_Data()
        {
            DataTable dt = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { txtItemCode_Edit.Text, txtLotNo_Edit.Text, "ThermalStress" }));
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

            img_dt = dt_img.AsDataView().ToTable(false, new string[] { "PCS_No", "Zone", "Before_Data", "After_Data", "Time_Update", "Operator", "Depart", "Remark" });
        }

        private void DGV_Image_Edit_CellContentDoubleClick_1(object sender, DataGridViewCellEventArgs e)
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

        private void txtLogfile_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                DataTable cur_dt = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "THERMAL_STRESS_IMAGE", filter_str);
                bool check_data = true;
                if (cur_dt.Rows.Count > 0)
                {
                    check_data = false;
                    DGV_Image_Graph.DataSource = cur_dt;

                    ((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_Data"]).Width = 100;
                    foreach (DataGridViewRow dr in DGV_Image_Graph.Rows)
                    {
                        dr.Height = 70;
                    }

                    if (MessageBox.Show("Table THERMAL_STRESS_IMAGE : Data is existed. Do you want to Update?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        check_data = true;
                    }

                    if (check_data)
                    {
                        if (txtLogfile.Text != "")
                        {
                            if (admin_mode)
                            {
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
                                    ThermalStress_Data_Process(txtLogfile.Text);
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
                }
                else
                {
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
                        ThermalStress_Data_Process(txtLogfile.Text);
                    }
                    else
                    {
                        MessageBox.Show("ItemCode or LotNo not matched", "Warning");
                    }
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

        private void txtLogfile_TextChanged(object sender, EventArgs e)
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

        private void btn_change_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "")
            {

            lblsetup: DataTable dt_setting = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "SETTING_PCS", TDMK_Code.filter_str(new string[] { "ItemCode", "Data_For" }, new string[] { txtItemCode.Text, "CQRA - Thermal stress" }));
                if (dt_setting.Rows.Count == 0)
                {
                    MessageBox.Show(new Form { TopMost = true }, "Please wait for setup qty", "Warning");

                    string file_format = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                    if (file_format != "")
                    {
                        setup_qty_ThermalStress(file_format, txtItemCode.Text);

                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Format of " + txtItemCode.Text + " not found", "Warning");
                    }
                }
                else
                {
                    if (MessageBox.Show("Do you want to update qty?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        TDMK_Code.Delelte_FilteredItem_arr("SETTING_PCS", myVar.mysqlcon, TDMK_Code.filter_str(new string[] { "ItemCode", "Data_For" }, new string[] { txtItemCode.Text, "CQRA - Thermal stress" }));
                        goto lblsetup;

                    }
                    else
                    {
                        return;
                    }
                }
            }
            else
            {
                MessageBox.Show("Please, fill in ItemCode", "Warning");
            }




        }
    }
}
