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
using OK2SHIP;
using Microsoft.Office.Core;
using System.Security.Cryptography;
using Diem_Lib;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using OfficeOpenXml;
using OK2SHIP_SMT;
using Echeck_LogFile_Process;

//using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

namespace VHX
{
    public partial class StackUp_Main : Form
    {
        Impedance_data proc_data = new Impedance_data();
        myVar exp_proc = new myVar();
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SEI_Lib myCode = new SEI_Lib();
        OK2SHIP.Export_Image_Class img_proc = new OK2SHIP.Export_Image_Class();
        public Diem_Lib.Data_Log_file_Process stackup_logfile = new Diem_Lib.Data_Log_file_Process();
        public Funtion_setup_new F_setup_new = new Funtion_setup_new();
        public EPPlus_Lib TDMK_Code2 = new EPPlus_Lib();
        ExportEPPlus F_exportEPPlus = new ExportEPPlus();

        public string strcon = "";
        public SqlConnection sqlcon = null;
        SortedDictionary<string, List<string>> edited_lst = new SortedDictionary<string, List<string>>();
        public bool admin_mode = false;

        public List<string> sel_pcs_lst = new List<string> { };
        public string cur_zone = "";
        bool update_mode = false;
        bool add_new_mode = false;
        public string depart = "QA";
        ECheck_Process Echeck_Pro = new ECheck_Process();

        public StackUp_Main()
        {
            InitializeComponent();
        }
        public void Get_logfile_Multi(string in_src, ref SortedDictionary<int, Dictionary<int, string>> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            FileInfo[] temp_lst = tar_d.GetFiles("*.csv");
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    Dictionary<int, string> temp = Get_LogFile_Data(temp_lst[i].FullName);
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
                    int f_inx = Convert.ToInt32(f_na);
                    lst_result.Add(f_inx, temp);
                }
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_logfile_Multi(inter_lst.FullName, ref lst_result);
                    }
                }
            }
        }
        public Dictionary<int, string> Get_LogFile_Data(string in_src_file)
        {
            Dictionary<int, string> result_lst = new Dictionary<int, string>();
            string[] Lines = File.ReadAllLines(in_src_file);
            foreach (var line in Lines)
            {
                string str = new string(line.Where(s => s != '"').ToArray());
                string[] temp = str.Split(',');
                if (temp[0] != "")
                {
                    if (myCode.IsNumeric(temp[0]))
                    {
                        result_lst.Add(Convert.ToInt32(temp[0]), temp[2]);
                    }
                }
                else
                {
                    break;
                }
            }
            return result_lst;
        }
        public void Get_Impedance_logfile_Multi(string in_src, ref Dictionary<string, Dictionary<string, SortedDictionary<int, myVar.Impedance_data>>> dic_lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            FileInfo[] temp_lst = tar_d.GetFiles("*.xlsm");
            if (temp_lst.Length > 0)
            {
                Dictionary<string, SortedDictionary<int, myVar.Impedance_data>> lst_result = new Dictionary<string, SortedDictionary<int, myVar.Impedance_data>>();
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    SortedDictionary<int, myVar.Impedance_data> temp = Get_Impedance_LogFile_Data(temp_lst[i].FullName);
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
                    lst_result.Add(f_na, temp);
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
                        Get_Impedance_logfile_Multi(inter_lst.FullName, ref dic_lst_result);
                    }
                }
            }
        }
        public void Get_Impedance_logfile_Multi2(string in_src, ref Dictionary<string, SortedDictionary<int, myVar.Impedance_data>> dic_lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            FileInfo[] temp_lst = tar_d.GetFiles("*.xlsm");
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    SortedDictionary<int, myVar.Impedance_data> temp = Get_Impedance_LogFile_Data(temp_lst[i].FullName);
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
                    dic_lst_result.Add(f_na, temp);
                }
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_Impedance_logfile_Multi2(inter_lst.FullName, ref dic_lst_result);
                    }
                }
            }
        }
        public SortedDictionary<int, myVar.Impedance_data> Get_Impedance_LogFile_Data(string in_src_file)
        {
            SortedDictionary<int, myVar.Impedance_data> result_lst = new SortedDictionary<int, myVar.Impedance_data>();
            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            myExcel.Workbook wrkbk = xlsApp.Workbooks.Open(in_src_file); //TDMK_Code.open_excel_file(in_src_file, "", "");
            myExcel.Worksheet wrksht = wrkbk.Sheets["Result"];
            myExcel.Range id_rgn = wrksht.Range["A1"];
            myExcel.Range val_rgn = wrksht.Range["A7"];
            myExcel.Range sel_rgn = id_rgn;
            int col_inx = 0;
            string id_val = myCode.checkDBNull(sel_rgn.Value);
            while (id_val != "")
            {
                myVar.Impedance_data imp_graph = new myVar.Impedance_data();
                imp_graph.data_val = myCode.checkDBNull(val_rgn.Offset[0, col_inx].Value);
                if (Convert.ToInt32(id_val) <= 3)
                {
                    myExcel.Worksheet sel_sht = wrkbk.Sheets[id_val];
                    byte[] sel_img = exp_proc.get_image_excel(sel_sht);
                    imp_graph.grap_data = sel_img;
                }
                result_lst.Add(Convert.ToInt32(id_val), imp_graph);
                col_inx++;
                sel_rgn = id_rgn.Offset[0, col_inx];
                id_val = myCode.checkDBNull(sel_rgn.Value);
            }
            wrkbk.Close();
            //Marshal.ReleaseComObject(xlsApp);
            return result_lst;
        }
        public void Get_Impedance_VHX_Multi(string in_src, ref SortedDictionary<int, byte[]> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            FileInfo[] temp_lst = tar_d.GetFiles("*.jpg");
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    var sel_img = Bitmap.FromFile(temp_lst[i].FullName);
                    ImageConverter imgcon = new ImageConverter();
                    byte[] img_data = (byte[])imgcon.ConvertTo(sel_img, typeof(byte[]));
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
                    int f_inx = Convert.ToInt32(f_na);
                    lst_result.Add(f_inx, img_data);
                }
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_Impedance_VHX_Multi(inter_lst.FullName, ref lst_result);
                    }
                }
            }
        }

        private void txtLogfile_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            // Logfile_Data_Process();
            if (txtItemCode.Text != "" && txtLotNo.Text != "")
            {
                if (!update_mode)
                {
                    DataTable dt = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "STACKUP_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                    if (dt.Rows.Count > 0)
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Data of " + txtItemCode.Text + " _ " + txtLotNo.Text + " is exist. Please update data", "Waring");

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
                                add_new_mode = true;
                            }

                        }
                    }
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

                        }

                    }
                }

            }
        }

        public void Logfile_Data_Process()
        {
            FolderBrowserDialog f_open = new FolderBrowserDialog();
            //f_open.SelectedPath = Application.StartupPath;
            f_open.SelectedPath = txtLogfile.Text;
            if (f_open.ShowDialog() == DialogResult.OK)
            {
                txtLogfile.Text = f_open.SelectedPath;
                string fpath = f_open.SelectedPath;
                string f_name = fpath.Split('\\')[fpath.Split('\\').Length - 1];

                string[] lot_no = txtLotNo.Text.Split('-');


                if (f_name.Split('-')[0].Trim() == txtItemCode.Text && lot_no.Length == f_name.Split('_')[0].Split('-').Length - 1)
                {
                    if (lot_no.Length == 2 && Convert.ToDecimal(f_name.Split('-')[1].Trim()) == Convert.ToDecimal(lot_no[0].Trim()) && Convert.ToDecimal(f_name.Split('-')[2].Split('_')[0].Trim()) == Convert.ToDecimal(lot_no[1].Trim()))
                    {
                        StackUp_Data_Process(f_open.SelectedPath);

                    }
                    else if (lot_no.Length == 1 && Convert.ToDecimal(f_name.Split('-')[1].Split('_')[0].Trim()) == Convert.ToDecimal(lot_no[0].Trim()))
                    {

                        StackUp_Data_Process(f_open.SelectedPath);
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "ItemCode or LotNo not matched", "Warning");
                    }

                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "ItemCode or LotNo not matched", "Warning");
                }

            }
        }
        public void Logfile_Data_Process_Loadnew()
        {

            if (txtLogfile.Text != "")
            {
                DirectoryInfo temp_dir = new DirectoryInfo(txtLogfile.Text);

                string temp_folder = temp_dir.Parent.Name;
                string temp_zone = temp_dir.Name;
                string f_name = "";
                if (temp_zone.Contains(txtItemCode.Text))
                {
                    f_name = temp_zone;
                }
                else if (temp_folder.Contains(txtItemCode.Text))
                {
                    f_name = temp_folder;
                }


                // string f_name = temp_dir.Name;
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
                    if (f_name == temp_zone)
                    {
                        StackUp_Data_Process(txtLogfile.Text);
                    }
                    else if (f_name == temp_folder)
                    {
                        //string path_new = (txtLogfile.Text).Replace(temp_zone, string.Empty);

                        string path_new = Directory.GetParent(temp_dir.FullName).ToString();
                        StackUp_Data_Process(path_new);

                    }

                    // StackUp_Data_Process();
                }

                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "ItemCode or LotNo not matched", "Warning");
                }

            }

        }
        public void Logfile_Data_Process_Update()
        {

            if (txtLogfile.Text != "")
            {

                string fpath = txtLogfile.Text;
                DirectoryInfo temp_dir = new DirectoryInfo(fpath);
                string temp_folder = temp_dir.Parent.Name;
                string temp_zone = temp_dir.Name;
                string f_name = "";
                if (temp_zone.Contains(txtItemCode.Text))
                {
                    f_name = temp_zone;
                }
                else if (temp_folder.Contains(txtItemCode.Text))
                {
                    f_name = temp_folder;
                }


                string[] lot_no = txtLotNo.Text.Split('-');


                if (f_name.Split('-')[0].Trim() == txtItemCode.Text && lot_no.Length == f_name.Split('_')[0].Split('-').Length - 1)
                {
                    if (lot_no.Length == 2 && Convert.ToDecimal(f_name.Split('-')[1].Trim()) == Convert.ToDecimal(lot_no[0].Trim()) && Convert.ToDecimal(f_name.Split('-')[2].Split('_')[0].Trim()) == Convert.ToDecimal(lot_no[1].Trim()))
                    {
                        StackUp_Data_Process(txtLogfile.Text);

                    }
                    else if (lot_no.Length == 1 && Convert.ToDecimal(f_name.Split('-')[1].Split('_')[0].Trim()) == Convert.ToDecimal(lot_no[0].Trim()))
                    {

                        StackUp_Data_Process(txtLogfile.Text);
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "ItemCode or LotNo not matched", "Warning");
                    }

                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "ItemCode or LotNo not matched", "Warning");
                }

            }
        }
        public string fill_name(string text, int max_lenght)
        {
            int lenght = text.Length;
            if (lenght < max_lenght)
            {
                for (int i = 0; i < 5 - lenght; i++)
                {
                    text = "0" + text;
                }

            }
            return text;
        }
        public void StackUp_Data_Process(string src_path)
        {
            if (src_path != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                Dictionary<string, SortedDictionary<int, byte[]>> lstImage_result = new Dictionary<string, SortedDictionary<int, byte[]>>();
                Dictionary<string, Dictionary<string, SortedDictionary<int, string>>> lstData_result = new Dictionary<string, Dictionary<string, SortedDictionary<int, string>>>();
                Get_Stackuplogfile_Multi2(src_path, ref lstData_result);
                Get_Stackup_Image_Multi(src_path, ref lstImage_result);
                DataTable Stackup_Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_LOGFILE", filter_str).Clone();
                DataTable Stackup_image_tbl = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_IMAGE", filter_str).Clone();
                int r_inx = 0;
                foreach (var data in lstData_result)
                {
                    string region = "Zone_" + data.Key;
                    foreach (var pcs in data.Value)
                    {
                        string pcs_no = pcs.Key;
                        foreach (var pcs_data in pcs.Value)
                        {
                            Stackup_Data_tbl.Rows.Add(r_inx + 1, txtItemCode.Text, fill_name(txtLotNo.Text, 5), region, pcs_no, pcs_data.Value, DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), txtOperator.Text, "STACKUP", src_path);
                            r_inx++;
                        }
                    }
                }
                DGV_Stackup.DataSource = Stackup_Data_tbl;
                r_inx = 0;
                foreach (var region_img in lstImage_result)
                {
                    string region = "Zone_" + region_img.Key;
                    foreach (var img in region_img.Value)
                    {
                        Stackup_image_tbl.Rows.Add(r_inx + 1, txtItemCode.Text, fill_name(txtLotNo.Text, 5), region, img.Key, img.Value, txtOperator.Text, DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), src_path);
                        r_inx++;
                    }
                }
                DGV_Image_Graph.DataSource = Stackup_image_tbl;

                ((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_data"]).Width = 100;

                foreach (DataGridViewRow dr in DGV_Image_Graph.Rows)
                {
                    dr.Height = 70;
                }
                btn_load.Enabled = false;
                add_new_mode = false;
            }
        }
        public void StackUp_Data_Process_Update(string src_path, List<string> sel_pcs_lst, string cur_zone)
        {
            if (src_path != "")
            {
                try
                {

                    cur_zone = cur_zone.Replace("Zone_", string.Empty);
                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                    Dictionary<string, SortedDictionary<int, byte[]>> lstImage_result = new Dictionary<string, SortedDictionary<int, byte[]>>();
                    Dictionary<string, Dictionary<string, SortedDictionary<int, string>>> lstData_result = new Dictionary<string, Dictionary<string, SortedDictionary<int, string>>>();
                    Get_Stackuplogfile_Multi2_Update(src_path, cur_zone, ref lstData_result);
                    Get_Stackup_Image_Multi_Update(src_path, ref lstImage_result, cur_zone);
                    DataTable Stackup_Data_tbl_update = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_LOGFILE", filter_str).Clone();
                    DataTable Stackup_image_tbl_update = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_IMAGE", filter_str).Clone();
                    int r_inx = 0;
                    DateTime tim_up = DateTime.Now;
                    foreach (var data in lstData_result)
                    {
                        string region = "Zone_" + data.Key;
                        foreach (var pcs in data.Value)
                        {
                            string pcs_no = pcs.Key;
                            foreach (var pcs_data in pcs.Value)
                            {
                                Stackup_Data_tbl_update.Rows.Add(r_inx + 1, txtItemCode.Text, fill_name(txtLotNo.Text, 5), region, pcs_no, pcs_data.Value, DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), txtOperator.Text, "STACKUP", src_path);
                                r_inx++;
                            }
                        }
                    }
                    // DataTable dt = (DataTable)DGV_Stackup.DataSource;
                    //Update DGV_Zone_Data
                    DataView dv = Stackup_Data_tbl_update.AsDataView();
                    foreach (var item in sel_pcs_lst)
                    {
                        string filter = TDMK_Code.filter_str(new string[] { "Region", "PCS_No" }, new string[] { "Zone_" + cur_zone, item });
                        dv.RowFilter = filter;
                        int min_post = new int[] { dv.Count, DGV_Zone_Data.Rows.Count - 1 }.Min();
                        for (int inx = 0; inx < min_post; inx++)
                        {
                            DataRow dr = dv[inx].Row;
                            int r_x = Stackup_Data_tbl_update.Rows.IndexOf(dr);
                            DGV_Zone_Data.Rows[inx].Cells["PCS_No" + item].Value = Stackup_Data_tbl_update.Rows[r_x]["Data"];
                            // DGV_Zone_Data.Rows[inx].Cells["PCS_No" + item].Style.BackColor = Color.Red; 
                        }
                    }

                    //Update DGV_Stackup
                    DataTable dt_all = (DataTable)DGV_Stackup.DataSource;
                    DataView dv_all = dt_all.AsDataView();
                    foreach (var item in sel_pcs_lst)
                    {
                        string filter = TDMK_Code.filter_str(new string[] { "Region", "PCS_No" }, new string[] { "Zone_" + cur_zone, item });
                        dv_all.RowFilter = filter;
                        int min_post = new int[] { dv_all.Count, DGV_Zone_Data.Rows.Count - 1 }.Min();
                        for (int inx = 0; inx < min_post; inx++)
                        {
                            DataRow dr = dv_all[inx].Row;
                            int r_x = dt_all.Rows.IndexOf(dr);
                            DGV_Stackup.Rows[r_x].Cells["Data"].Value = DGV_Zone_Data.Rows[inx].Cells["PCS_No" + item].Value.ToString();
                            DGV_Stackup.Rows[r_x].Cells["Operator"].Value = txtOperator.Text;
                            DGV_Stackup.Rows[r_x].Cells["Time_Update"].Value = tim_up.ToString();
                            DGV_Stackup.Rows[r_x].Cells["Remark"].Value = txtLogfile.Text;

                        }
                    }

                    r_inx = 0;
                    foreach (var region_img in lstImage_result)
                    {
                        string region = "Zone_" + region_img.Key;
                        foreach (var img in region_img.Value)
                        {
                            Stackup_image_tbl_update.Rows.Add(r_inx + 1, txtItemCode.Text, fill_name(txtLotNo.Text, 5), region, img.Key, img.Value, DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), txtOperator.Text, src_path);
                            r_inx++;
                        }
                    }

                    //Update DGV_Image_Graph
                    int idx = 0;
                    DataTable dt_image = (DataTable)DGV_Image_Graph.DataSource;
                    DataView dv_image = dt_image.AsDataView();
                    foreach (var item in sel_pcs_lst)
                    {
                        string filter = TDMK_Code.filter_str(new string[] { "Region", "PCS_No" }, new string[] { "Zone_" + cur_zone, item });
                        dv_image.RowFilter = filter;

                        if (dv_image.Count > 0)
                        {
                            DataRow dr = dv_image[0].Row;
                            int r_x = dt_image.Rows.IndexOf(dr);
                            if (idx < Stackup_image_tbl_update.Rows.Count)
                            {
                                DGV_Image_Graph.Rows[r_x].Cells["Image_Data"].Value = Stackup_image_tbl_update.Rows[idx]["Image_Data"];
                                DGV_Image_Graph.Rows[r_x].Cells["Time_Update"].Value = tim_up.ToString();
                                DGV_Image_Graph.Rows[r_x].Cells["Operator"].Value = txtOperator.Text;
                                DGV_Image_Graph.Rows[r_x].Cells["Remark"].Value = txtLogfile.Text;

                                idx++;
                            }
                        }
                    }

                    if (edited_lst.Keys.ToList().IndexOf(cur_zone) == -1)
                    {
                        edited_lst.Add(cur_zone, sel_pcs_lst);
                    }
                    else if (edited_lst.Values.ToList().IndexOf(sel_pcs_lst) == -1)
                    {
                        edited_lst.Remove(cur_zone);
                        edited_lst.Add(cur_zone, sel_pcs_lst);
                    }
                    btn_load.Enabled = false;
                    update_mode = false;
                    MessageBox.Show(new Form { TopMost = true }, "Update data completed", "Warning");

                }
                catch
                {
                    MessageBox.Show(new Form { TopMost = true }, "Data is incorrect", "Warning");
                }

            }
        }
        public void Load_Spec_StackUp_To_DGV(string file_addr, DataGridView DGV, ref SortedDictionary<string, string> dic_spec_detail)
        {
            //SortedDictionary<string, string> dic_spec_detail = new SortedDictionary<string, string> { };
            myExcel.Workbook wb = TDMK_Code.open_excel_file(file_addr, "", "");
            string sheet_name = "";
            string tar_sheet = "Stack-up";
            for (int i = 0; i < wb.Sheets.Count; i++)
            {
                myExcel.Worksheet cur_sht = wb.Sheets[i + 1];
                string cur_sht_name = cur_sht.Name;
                //if (tar_sheet.Replace("-", "").ToUpper() == cur_sht_name.Replace("-", "").Replace(" ", "").ToUpper())
                //{
                if (cur_sht_name.Replace("-", "").Replace(" ", "").ToUpper().Contains("STACKUP"))
                {
                    sheet_name = cur_sht_name;
                    break;
                }
            }
            if (sheet_name != "")
            {
                myExcel.Worksheet ws = wb.Sheets[sheet_name];
                DGV.Rows.Add("USL", "USL");
                DGV.Rows.Add("LSL", "LSL");
                DGV.Rows.Add("Normal", "Normal");

                // Tìm cột USL
                int col_usl = 4;
                bool found = false;
                for (int row = 50; row <= 120; row++)
                {
                    for (int col = 3; col <= 7; col++)
                    {
                        myExcel.Range cell = ws.Cells[row, col];
                        if (cell.Value2 != null && (cell.Value2.ToString() == "USL"))
                        {
                            col_usl = col;
                            found = true;
                            break; // Thoát khỏi vòng lặp
                        }
                    }
                    if (found)
                    {
                        break; // Thoát khỏi vòng lặp ngoài cùng
                    }
                }

                for (int i = 5; i < 400; i++)
                {
                    if (ws.Cells[i, 1].value != null)
                    {

                        if (ws.Cells[i, 1].value.ToString() == "Zone" && ws.Cells[i + 1, 1].value != null)
                        {
                            string zone_ = ws.Cells[i + 1, 1].value.ToString();
                            string lst_spec_detail = "";
                            DGV.Columns.Add(ws.Cells[i + 1, 1].value.ToString(), ws.Cells[i + 1, 1].value.ToString());
                            for (int j = i + 1; j < i + 20; j++)
                            {
                                if (myCode.checkDBNull(ws.Cells[j, 6]) != "")
                                {
                                    if (myCode.checkDBNull(ws.Cells[j, 2].Value).ToUpper() == "ADHESIVE")
                                    {
                                        string usl = "";
                                        string lsl = "";
                                        string val_post = myCode.checkDBNull(ws.Cells[j, 6].Value);
                                        string val_pre = myCode.checkDBNull(ws.Cells[j, 5].Value);
                                        if (myCode.IsNumeric(val_pre))
                                        {
                                            usl = (Math.Round(double.Parse(val_post) * 1.1, 2)).ToString();
                                        }
                                        if (myCode.IsNumeric(val_post))
                                        {
                                            lsl = (Math.Round(double.Parse(val_post) * 0.9)).ToString();
                                        }

                                        lst_spec_detail += usl + "/" + lsl + ";";

                                    }
                                    else
                                    {
                                        lst_spec_detail += myCode.checkDBNull(ws.Cells[j, 6].Value) + ";";
                                    }

                                }
                                if (myCode.checkDBNull(ws.Cells[j + 1, 2].Value).Replace(" ", string.Empty).ToUpper() == "Total thickness".Replace(" ", string.Empty).ToUpper())
                                {
                                    break;
                                }
                            }
                            if (!dic_spec_detail.ContainsKey(zone_))
                            {
                                dic_spec_detail.Add(zone_, lst_spec_detail);
                            }


                            for (int j = i; j < i + 20; j++)
                            {

                                if (ws.Cells[j, col_usl].Value != null)
                                {
                                    if (ws.Cells[j, col_usl].Value.ToString() == "USL")
                                    {
                                        DGV.Rows[0].Cells[DGV.Columns.Count - 1].Value = ws.Cells[j, col_usl + 1].Value;
                                        DGV.Rows[1].Cells[DGV.Columns.Count - 1].Value = ws.Cells[j + 1, col_usl + 1].Value;
                                        if (ws.Cells[j - 2, 2].Value2 != null && ws.Cells[j - 2, 2].Value2.ToString() != "")
                                        {
                                            if (ws.Cells[j - 2, 2].Value2.ToString() == "Total thickness")
                                            {
                                                // DGV.Rows[2].Cells[DGV.Columns.Count - 1].Value = ((ws.Cells[j - 2, 6].Value2.ToString()).Split('+')[0]).Split('±')[0];
                                                DGV.Rows[2].Cells[DGV.Columns.Count - 1].Value = ws.Cells[j - 2, 6].Value2.ToString();
                                            }

                                        }
                                        if (ws.Cells[j - 3, 2].Value2 != null && ws.Cells[j - 3, 2].Value2.ToString() != "")
                                        {
                                            if (ws.Cells[j - 3, 2].Value2.ToString() == "Total thickness")
                                            {
                                                //DGV.Rows[2].Cells[DGV.Columns.Count - 1].Value = ((ws.Cells[j - 3, 6].Value2.ToString()).Split('+')[0]).Split('±')[0];
                                                DGV.Rows[2].Cells[DGV.Columns.Count - 1].Value = ws.Cells[j - 3, 6].Value2.ToString();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }


                int count_sample = 0;
                for (int row = 2; row <= 100; row++)
                {

                    myExcel.Range cell = wb.Sheets[sheet_name].Cells[row, 1];
                    if (cell.Value2 != null && (cell.Value2.ToString().ToUpper().Contains("ZONE")))
                    {
                        for (int col = 1; col < 20; col++)
                        {
                            string cell_val = myCode.checkDBNull(wb.Sheets[sheet_name].Cells[row, col].Value);
                            if (cell_val.Contains("#"))
                            {
                                count_sample++;
                            }
                            else if (cell_val == "")
                            {
                                break;
                            }
                        }
                        break;
                    }
                }
                int ID = TDMK_Code.SQL_MAX("SETTING_PCS", "ID", sqlcon) + 1;
                TDMK_Code.insert_val_arr("SETTING_PCS", sqlcon, new string[] { "ID", "ItemCode", "Data_For", "Region_Zone", "Pcs_setting" }, new string[] { ID.ToString(), txtItemCode.Text, "Stack-up", "", count_sample.ToString() });
                wb.Close();
                TDMK_Code.releaseObject(wb);
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Worksheet " + tar_sheet + " not found.\r\nPlease check format", "Warning");
            }


        }

        //public void change_ID_DGV(DataGridView dgv)
        //{
        //    DataTable dt = new DataTable();
        //    dt = (DataTable)dgv.DataSource;
        //    for(int i = 0; i < dt.)
        //}
        private void Impedance_main_Load(object sender, EventArgs e)
        {
            sqlcon = myVar.mysqlcon;
            admin_mode = false;
            //DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { txtItemCode_Edit.Text, txtLotNo_Edit.Text, "Stackup" }));
            //string[] arr_zone = dt.AsEnumerable().Select(x => x.Field<string>("Zone")).Distinct().ToArray();

            //foreach (string item in arr_zone)
            //{
            //    cbProcess_Edit.Items.Add(item);
            //}
            //cbProcess_Edit.SelectedIndex = 0;


        }
        private void DGV_Image_Graph_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int col_inx = e.ColumnIndex;
            int r_inx = e.RowIndex;
            DataGridViewCell cur_cell = DGV_Image_Graph.CurrentCell;
            if (DGV_Image_Graph.Columns[col_inx].Name.Contains("Image"))
            {
                if (myCode.checkDBNull(cur_cell.Value) != "")
                {
                    //byte[] data = (byte[])cur_cell.Value;
                    //using (MemoryStream ms = new MemoryStream(data))
                    //{
                    //    picDetails.Image = Image.FromStream(ms);
                    //}
                    string region = DGV_Image_Graph.Rows[r_inx].Cells["Region"].Value.ToString();
                    string pcs = DGV_Image_Graph.Rows[r_inx].Cells["Pcs_No"].Value.ToString();

                    View_detail_Image fr1 = new View_detail_Image();
                    fr1.data = (byte[])cur_cell.Value;
                    fr1.region_ = region;
                    fr1.pcs_ = pcs;
                    fr1.choose_sheet_ = "Stackup";
                    fr1.Show();

                }
                else
                {
                    //picDetails.Image = null;
                }

                //string region = DGV_Image_Graph.Rows[r_inx].Cells["Region"].Value.ToString();
                //string pcs = DGV_Image_Graph.Rows[r_inx].Cells["Pcs_No"].Value.ToString();
                //lblImage_Graph.Text = "Image / Graph Details ---> Region: " + region + " / " + "Pcs_No: " + pcs;
            }




        }
        private void tsmLoadData_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                List<string> tar_table_name = new List<string>() { "STACKUP_IMAGE", "STACKUP_LOGFILE" };
                List<DataGridView> tar_DGV = new List<DataGridView>() { DGV_Image_Graph, DGV_Stackup };
                for (int i = 0; i < tar_DGV.Count; i++)
                {
                    tar_DGV[i].DataSource = TDMK_Code.Datatable_Filter(sqlcon, tar_table_name[i], filter_str);
                    myCode.DGV_Auto_Resize(tar_DGV[i]);
                }
                lblZone.Text = "Zone Data Detail";
                DGV_Zone_Data.DataSource = null;



                ((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_data"]).Width = 100;

                foreach (DataGridViewRow dr in DGV_Image_Graph.Rows)
                {
                    dr.Height = 70;
                }

                if (DGV_Stackup.Rows.Count == 0)
                {
                    MessageBox.Show(new Form { TopMost = true }, "No data ", "Warning");
                }

                if (DGV_spec.Rows.Count > 0)
                {
                    DGV_spec.Rows.Clear();
                    DGV_spec.Columns.Clear();
                    DGV_spec.Columns.Add("Zone", "Zone");
                }



                AutoCompleteStringCollection list_zone = TDMK_Code.Load_Item_Filter_str(sqlcon, "STACKUP_SPEC", "Zone", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
                if (list_zone.Count > 0)
                {

                    DGV_spec.Rows.Add("USL", "USL");
                    DGV_spec.Rows.Add("LSL", "LSL");
                    DGV_spec.Rows.Add("Normal", "Normal");
                    DGV_spec.Rows.Add();
                    foreach (string zone in list_zone)
                    {
                        DGV_spec.Columns.Add(zone, zone);
                        AutoCompleteStringCollection usl = TDMK_Code.Load_Item_Filter_str(sqlcon, "STACKUP_SPEC", "USL", TDMK_Code.filter_str(new string[] { "ItemCode", "Zone" }, new string[] { txtItemCode.Text, zone }));
                        AutoCompleteStringCollection lsl = TDMK_Code.Load_Item_Filter_str(sqlcon, "STACKUP_SPEC", "LSL", TDMK_Code.filter_str(new string[] { "ItemCode", "Zone" }, new string[] { txtItemCode.Text, zone }));
                        AutoCompleteStringCollection normal = TDMK_Code.Load_Item_Filter_str(sqlcon, "STACKUP_SPEC", "Normal", TDMK_Code.filter_str(new string[] { "ItemCode", "Zone" }, new string[] { txtItemCode.Text, zone }));
                        DGV_spec.Rows[0].Cells[zone].Value = usl[0];
                        DGV_spec.Rows[1].Cells[zone].Value = lsl[0];
                        DGV_spec.Rows[2].Cells[zone].Value = normal[0];
                    }
                    DGV_spec.Rows.RemoveAt(DGV_spec.Rows.Count - 1);
                    //stackup_logfile.HightLigh_StackUp_LogFile(DGV_spec, DGV_Zone_Data);
                    DGV_spec.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                }
                else
                {

                    if (MessageBox.Show(new Form { TopMost = true }, "Please wait for setup spec", "Warning") == DialogResult.OK)
                    {
                        string file_format = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                        if (file_format != "")
                        {
                            SortedDictionary<string, string> dic_spec_detail = new SortedDictionary<string, string> { };
                            // stackup_logfile.Load_Spec_StackUp_To_DGV(file_format, DGV_spec);
                            Load_Spec_StackUp_To_DGV(file_format, DGV_spec, ref dic_spec_detail);
                            //stackup_logfile.HightLigh_StackUp_LogFile(DGV_spec, DGV_Zone_Data); 
                            Save_Spec_new(dic_spec_detail);
                            MessageBox.Show(new Form { TopMost = true }, "Setup spec completed", "Warning");



                        }
                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Format of " + txtItemCode.Text + " not found", "Warning");
                        }
                    }



                }

                if (DGV_spec.Rows.Count > 1)
                {
                    string[] arr_zone = ((DataTable)DGV_Stackup.DataSource).AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();
                    List<string> spec_col_lst = new List<string>();
                    foreach (DataGridViewColumn dgv_c in DGV_spec.Columns)
                    {
                        spec_col_lst.Add(dgv_c.Name);
                    }

                    foreach (string zone in arr_zone)
                    {
                        int col_inx = spec_col_lst.IndexOf(zone.Replace("Zone_", ""));

                        if (col_inx != -1)
                        {
                            double max = Convert.ToDouble(DGV_spec.Rows[0].Cells[col_inx].Value.ToString());
                            double min = Convert.ToDouble(DGV_spec.Rows[1].Cells[col_inx].Value.ToString());

                            DataTable dt_spec_zone = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "STACKUP_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Zone" }, new string[] { txtItemCode.Text, zone.Replace("Zone_", "") }));
                            if (dt_spec_zone.Rows.Count > 0)
                            {
                                string[] lst_spec_detail = dt_spec_zone.Rows[0]["Spec_detail"].ToString().Split(';');
                                DataView dv = ((DataTable)DGV_Stackup.DataSource).AsDataView();

                                string filter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { zone });
                                dv.RowFilter = filter;
                                string[] arr_pcs = dv.ToTable().AsEnumerable().Select(x => x.Field<string>("Pcs_No")).Distinct().ToArray();
                                foreach (string pcs in arr_pcs)
                                {
                                    DataView dv_filter = ((DataTable)DGV_Stackup.DataSource).AsDataView();
                                    dv_filter.RowFilter = TDMK_Code.filter_str(new string[] { "Region", "Pcs_No" }, new string[] { zone, pcs });
                                    check_spec_zone_pcs(lst_spec_detail, dv_filter, max, min, pcs);

                                }
                            }
                        }


                    }
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Please, fill in ItemCode / LotNo / Operator ", "Warning");
            }
        }

        public void check_spec_zone_pcs(string[] lst_spec_detail, DataView dv_detail, double USL, double LSL, string pcs)
        {

            int index = 0;
            double sum = 0;
            //if (pcs == "1")
            //{

            //}
            foreach (string sp in lst_spec_detail)
            {
                if (sp != "")
                {

                    if (index < dv_detail.Count)
                    {
                        DataRow row = dv_detail[index].Row;
                        int r_x = ((DataTable)DGV_Stackup.DataSource).Rows.IndexOf(row);
                        string val = DGV_Stackup.Rows[r_x].Cells["Data"].Value.ToString();
                        sum += double.Parse(val);

                        string nominal_val = sp;
                        double min = 0;
                        double max = 0;
                        if (nominal_val.Contains("+") || nominal_val.Contains("-"))
                        {
                            nominal_val = nominal_val.Replace("+", ";").Replace("-", ";").Replace("/", ";");
                            max = double.Parse(nominal_val.Split(';')[0]) + double.Parse(nominal_val.Split(';')[1]);
                            min = double.Parse(nominal_val.Split(';')[0]) - double.Parse(nominal_val.Split(';')[3]);

                        }
                        else if (nominal_val.Contains("±"))
                        {
                            max = double.Parse(nominal_val.Split('±')[0]) + double.Parse(nominal_val.Split('±')[1]);
                            min = double.Parse(nominal_val.Split('±')[0]) - double.Parse(nominal_val.Split('±')[1]);
                        }
                        else if (nominal_val.Contains("/"))
                        {
                            if (myCode.IsNumeric(nominal_val.Split('/')[1]))
                            {
                                min = double.Parse(nominal_val.Split('/')[1]);
                            }
                            if (myCode.IsNumeric(nominal_val.Split('/')[0]))
                            {
                                max = double.Parse(nominal_val.Split('/')[0]);
                            }

                        }
                        else
                        {
                            min = double.Parse(nominal_val) * 0.7;
                            max = double.Parse(nominal_val) * 1.3;
                        }


                        for (int j = 1; j < dv_detail.Count; j++)
                        {

                            if (double.Parse(val) > max || double.Parse(val) < min)
                            {
                                DGV_Stackup.Rows[r_x].Cells["Data"].Style.BackColor = Color.Red;
                            }
                            else
                            {
                                DGV_Stackup.Rows[r_x].Cells["Data"].Style.BackColor = Color.White;
                            }


                        }
                        index++;
                    }
                }
            }
            if (sum < LSL || sum > USL)
            {
                for (int i = 0; i < index; i++)
                {
                    DataRow row = dv_detail[i].Row;
                    int r_x = ((DataTable)DGV_Stackup.DataSource).Rows.IndexOf(row);
                    DGV_Stackup.Rows[r_x].Cells["Pcs_No"].Style.BackColor = Color.Red;
                }
            }
        }
        public bool Check_in_Spec_zone(System.Data.DataTable DGV_Data, System.Data.DataTable DGV_Spec, string zone)
        {
            bool result = true;
            if (DGV_Data.Rows.Count > 2)
            {
                if (DGV_Spec.Rows.Count > 1)
                {
                    List<string> spec_col_lst = new List<string>();
                    foreach (DataRow dr in DGV_Spec.Rows)
                    {
                        spec_col_lst.Add(dr[3].ToString());
                    }

                    int col_inx = spec_col_lst.IndexOf(zone.Replace("Zone_", ""));
                    if (col_inx != -1)
                    {

                        double max = Convert.ToDouble(DGV_Spec.Rows[col_inx]["USL"].ToString());
                        double min = Convert.ToDouble(DGV_Spec.Rows[col_inx]["LSL"].ToString());
                        double sum = 0;
                        for (int m = 0; m < DGV_Data.Rows.Count; m++)
                        {
                            //int m = DGV_Data.Rows.Count - 1;

                            if (DGV_Data.Rows[m]["Data"] != null)
                            {
                                if (stackup_logfile.IsNumber(DGV_Data.Rows[m]["Data"].ToString()))
                                {
                                    double x = Convert.ToDouble(DGV_Data.Rows[m]["Data"].ToString());
                                    //if (x < min || x > max)
                                    //{
                                    //    result = false;
                                    //    break;
                                    //}
                                    sum += x;
                                }
                            }

                        }
                        if (sum < min || sum > max)
                        {
                            result = false;
                            return result;
                        }


                    }
                    string pcs = DGV_Data.Rows[0]["Pcs_No"].ToString();
                    if (result && pcs == "1")
                    {
                        DataTable dt_spec_zone = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "STACKUP_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Zone" }, new string[] { txtItemCode.Text, zone.Replace("Zone_", "") }));
                        if (dt_spec_zone.Rows.Count > 0)
                        {
                            string[] lst_spec_detail = dt_spec_zone.Rows[0]["Spec_detail"].ToString().Split(';');
                            int idx = 0;
                            foreach (string spc in lst_spec_detail)
                            {
                                if (spc != "")
                                {
                                    string nominal_val = spc;
                                    double min = 0;
                                    double max = 0;
                                    if (nominal_val.Contains("+") || nominal_val.Contains("-"))
                                    {
                                        nominal_val = nominal_val.Replace("+", ";").Replace("-", ";").Replace("/", ";");
                                        max = double.Parse(nominal_val.Split(';')[0]) + double.Parse(nominal_val.Split(';')[1]);
                                        min = double.Parse(nominal_val.Split(';')[0]) - double.Parse(nominal_val.Split(';')[3]);

                                    }
                                    else if (nominal_val.Contains("±"))
                                    {
                                        max = double.Parse(nominal_val.Split('±')[0]) + double.Parse(nominal_val.Split('±')[1]);
                                        min = double.Parse(nominal_val.Split('±')[0]) - double.Parse(nominal_val.Split('±')[1]);
                                    }
                                    else if (nominal_val.Contains("/"))
                                    {
                                        if (myCode.IsNumeric(nominal_val.Split('/')[1]))
                                        {
                                            min = double.Parse(nominal_val.Split('/')[1]);
                                        }
                                        if (myCode.IsNumeric(nominal_val.Split('/')[0]))
                                        {
                                            max = double.Parse(nominal_val.Split('/')[0]);
                                        }
                                    }
                                    else
                                    {
                                        min = double.Parse(nominal_val) * 0.7;
                                        max = double.Parse(nominal_val) * 1.3;
                                    }

                                    double data = double.Parse(DGV_Data.Rows[idx]["Data"].ToString());
                                    if (data > max || data < min)
                                    {
                                        result = false;
                                        return result;
                                    }
                                    idx++;
                                }
                            }
                        }
                    }

                }
            }
            return result;
        }
        public bool Check_Stackup_data(System.Data.DataTable stakup_tbl, System.Data.DataTable spec_tbl)
        {
            bool result = true;
            if (stakup_tbl.Rows.Count > 0)
            {
                List<string> region_lst = stakup_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToList();
                foreach (var _region in region_lst)
                {
                    System.Data.DataTable temp = stakup_tbl.AsEnumerable().Where(x => x.Field<string>("Region") == _region).CopyToDataTable();
                    // DataTable pcs_tbl = new DataTable();
                    List<DataTable> src_sample = new List<DataTable>();
                    exp_proc.Get_ListTable(-1, temp, new string[] { "Pcs_No" }, ref src_sample, "Data");
                    foreach (DataTable dt_temp in src_sample)
                    {
                        if (!Check_in_Spec_zone(dt_temp, spec_tbl, _region))
                        {
                            result = false;
                            break;
                        }
                    }
                    if (!result)
                    {
                        break;
                    }
                }
            }
            return result;
        }
        private void tsmSaveData_Click(object sender, EventArgs e)
        {
            bool save_en = false;
            System.Data.DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_SPEC", "ItemCode ='" + txtItemCode.Text + "'");
            if (spec_dt.Rows.Count > 0)
            {
                if (Check_Stackup_data((DataTable)DGV_Stackup.DataSource, spec_dt))
                {
                    save_en = true;
                }

      save_lbl: if (save_en)
                {
                    if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtOperator.Text != "")
                    {
                        //if (admin_mode)
                        //{
                        
                        string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                        List<string> tar_table_name = new List<string>() { "STACKUP_IMAGE", "STACKUP_LOGFILE" };
                        List<DataTable> tar_table = new List<DataTable>() { (DataTable)DGV_Image_Graph.DataSource, (DataTable)DGV_Stackup.DataSource };
                        DataTable logfile_dt = (DataTable)DGV_Stackup.DataSource;
                        DataTable img_dt = (DataTable)DGV_Image_Graph.DataSource;
                        List<string>region_lst = logfile_dt.AsEnumerable().Select(x=>x.Field<string>("Region")).Distinct().ToList();
                        List<string> region_pcs_Not_Match = new List<string>();                        
                        foreach(string region in region_lst)
                        {
                            List<string> logfile_pcs_lst = logfile_dt.AsEnumerable().Where(x=>x.Field<string>("Region")==region).Select(x=>x.Field<string>("Pcs_No")).Distinct().ToList();
                            List<string> img_pcs_lst = img_dt.AsEnumerable().Where(x => x.Field<string>("Region") == region).Select(x => x.Field<string>("Pcs_No")).Distinct().ToList();
                            List<string> result_lst = Echeck_Pro.Get_intersec(new List<List<string>>() { logfile_pcs_lst, img_pcs_lst });   
                            if((result_lst.Count == 0 )||(result_lst.Count!=logfile_pcs_lst.Count))
                            {
                                region_pcs_Not_Match.Add(region);
                            }
                        }
                        if(region_pcs_Not_Match.Count>0)
                        {
                            string mess_cont = "Dữ liệu ảnh và logfile không khớp\r\nHãy kiểm tra lại các region sau:\r\n";
                            foreach(string r in region_pcs_Not_Match)
                            {
                                mess_cont+= r +"\r\n";
                            }
                            MessageBox.Show(new Form { TopMost = true }, mess_cont.Trim(), "Thông báo");
                            return;
                        }
                        int tbl_inx = 0;
                        bool save_history = false;
                        foreach (var t in tar_table_name)
                        {
                        start_lbl: DataTable cur_dt = TDMK_Code.Datatable_Filter(sqlcon, t, filter_str);
                            if (cur_dt.Rows.Count > 0)
                            {
                                if (MessageBox.Show(new Form { TopMost = true }, "Table " + t + ": Data is existed. Do you want to overwrite?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                                {
                                    if (!save_history)
                                    {
                                        Stackup_history();
                                        save_history = true;
                                    }

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


                        MessageBox.Show(new Form { TopMost = true }, "Save data completed", "Warning");
                        //}
                        //else
                        //{
                        //    MessageBox.Show(new Form { TopMost = true }, "Please, login to save data", "Warning");
                        //}

                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Please, fill in ItemCode / LotNo / Operator ", "Warning");
                    }
                }
                else
                {

                    if (MessageBox.Show(new Form { TopMost = true }, "Data is out of spec. Do you want to save to database?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        if(admin_mode)
                        {
                            save_en = true;
                            goto save_lbl;
                        }
                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Please, login to save out of spec data", "Warning");
                        }
                    }
                }
                //edited_lst = new SortedDictionary<int, List<string>>();
                sel_pcs_lst = new List<string> { };
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Please, Click LOAD SPEC to check data ", "Warning");
            }
        }
        private void txtLogfile_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //StackUp_Data_Process(txtLogfile.Text);
                // stackup_logfile.Load_StackUp_Data_To_DGV(txtLogfile.Text, DGV_Zone_Data);
                if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtLogfile.Text != "")
                {
                    if (!update_mode)
                    {
                        DataTable dt = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "STACKUP_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                        if (dt.Rows.Count > 0)
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Data of " + txtItemCode.Text + " _ " + txtLotNo.Text + " is exist. Please update data", "Waring");
                        }
                        else
                        {
                            if (txtLogfile.Text != "")
                            {
                                ImageDetail fr1 = new ImageDetail(SetValue);
                                fr1.startpath = txtLogfile.Text;
                                fr1.Show();

                                btn_load.Enabled = true;
                            }
                            add_new_mode = true;
                        }
                    }
                    else
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
        public AutoCompleteStringCollection Get_zone_infor_to_input_value(myExcel.Workbook wb)
        {
            myExcel.Worksheet ws = wb.Sheets[1];
            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 24; i < 400; i++)
            {
                if (ws.Cells[i, 1].Value != null)
                {
                    if (ws.Cells[i, 1].Value.ToString() == "Zone")
                    {
                        int k = 1;
                        int dem = 0;
                        while (ws.Cells[i + k, 5].Value != null)
                        {
                            dem = dem + 1;
                            k++;
                        }
                        list.Add(ws.Cells[i + 1, 1].Value.ToString() + "+" + (i + 1).ToString() + "+" + (dem - 1).ToString());
                    }
                }
            }
            return list;
        }
        public void Export_StackUp_Image_old(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            AutoCompleteStringCollection list = Get_zone_infor_to_input_image(wb);
            myExcel.Worksheet ws = wb.Sheets[1];
            int no_BVH = 0;
            foreach (string x in list)
            {
                string region = x.Split('+')[0];
                if (!x.Contains("H_X"))
                {
                    region = "Zone_" + region;
                }
                else
                {
                    region = "Zone_" + region.Split('_')[0];
                }
                string row_val = x.Split('+')[1];
                int number_image = 5;
                string[] item = { "ItemCode", "LotNo", "Region" };
                string[] item_val = { Itemcode, Lotno, region };

                System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_IMAGE", TDMK_Code.filter_str(item, item_val));
                if (region != "Zone_BVH")
                {
                    if (data_image.Rows.Count > 0)
                    {
                        Export_Image_(ws, number_image, data_image, row_val);
                    }
                }
                else
                {
                    Export_Image_BVH(ws, data_image, row_val, no_BVH);
                    no_BVH++;
                }
            }

        }
        public void Export_StackUp_Image(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            AutoCompleteStringCollection list = Get_zone_infor_to_input_image(wb);
            myExcel.Worksheet ws = wb.Sheets[1];
            DataTable dt_tbl = (DataTable)DGV_Image_Graph.DataSource;
            List<string> lst_region = dt_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).ToList();
            bool t = true;
            if (lst_region.Contains("Zone_BVH1") && lst_region.Contains("Zone_BVH2"))
            {
                t = false;
            }
            int no_BVH = 1;
            foreach (string x in list)
            {
                string region = x.Split('+')[0];
                string row_val = x.Split('+')[1];
                int number_image = 5;

                if (region.Contains("BVH") || region.Contains("PTH"))
                {
                    if (t)
                    {
                        region = "Zone_BVH";
                    }
                    else
                    {
                        region = "Zone_BVH" + no_BVH.ToString();
                        no_BVH++;
                    }
                }
                else if (region.Contains("PTH"))
                {
                    region = "Zone_PTH";
                }
                else
                {
                    region = "Zone_" + region;
                }

                string[] item = { "ItemCode", "LotNo", "Region" };
                string[] item_val = { Itemcode, Lotno, region };
                DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_IMAGE", TDMK_Code.filter_str(item, item_val));
                Export_Image_new(ws, number_image, data_image, row_val);

            }
        }

        public void Export_Image_new(myExcel.Worksheet ws, int number_image, System.Data.DataTable data_image, string row_val)
        {
            int count_sample = new int[] { number_image, data_image.Rows.Count }.Min();

            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = Path.Combine(file_folder, "1.jpg");

            myExcel.Range curr_rgn = ws.Range["F" + row_val];
            for (int m = 0; m < count_sample; m++)
            {
                byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                MemoryStream ms = new MemoryStream(data);
                Image image = Image.FromStream(ms);
                image.Save(file_dic);
                InsertPicture_Name(ws, curr_rgn, file_dic, 6);
                curr_rgn = curr_rgn.Offset[0, 1];
            }
        }
        public void Export_Image_(myExcel.Worksheet ws, int number_image, System.Data.DataTable data_image, string row_val)
        {
            if (data_image.Rows.Count < number_image)
            {
                number_image = data_image.Rows.Count;
                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                if (!Directory.Exists(file_folder))
                {
                    Directory.CreateDirectory(file_folder);
                }
                string file_dic = Path.Combine(file_folder, "1.jpg");
                myExcel.Range curr_rgn = ws.Range["F" + row_val];
                for (int m = 0; m < number_image; m++)
                {
                    byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                    MemoryStream ms = new MemoryStream(data);
                    Image image = Image.FromStream(ms);
                    image.Save(file_dic);
                    InsertPicture_Name(ws, curr_rgn, file_dic, 6);
                    curr_rgn = curr_rgn.Offset[0, 1];
                }
            }

            else
            {
                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                if (!Directory.Exists(file_folder))
                {
                    Directory.CreateDirectory(file_folder);
                }
                string file_dic = Path.Combine(file_folder, "1.jpg");
                myExcel.Range curr_rgn = ws.Range["F" + row_val];
                for (int m = 0; m < number_image; m++)
                {
                    byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                    MemoryStream ms = new MemoryStream(data);
                    Image image = Image.FromStream(ms);
                    image.Save(file_dic);
                    InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                    curr_rgn = curr_rgn.Offset[0, 1];
                }
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

                myExcel.Range curr_rgn = ws.Range["F" + row_val];
                for (int m = no_BVH * 5; m < number_image; m++)
                {
                    byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                    MemoryStream ms = new MemoryStream(data);
                    Image image = Image.FromStream(ms);
                    image.Save(file_dic);
                    InsertPicture_Name(ws, curr_rgn, file_dic, 6);
                    curr_rgn = curr_rgn.Offset[0, 1];
                }
            }

            else
            {
                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                string file_dic = Path.Combine(file_folder, "1.jpg");

                myExcel.Range curr_rgn = ws.Range["F" + row_val];
                for (int m = no_BVH * 5; m < number_image; m++)
                {
                    byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                    MemoryStream ms = new MemoryStream(data);
                    Image image = Image.FromStream(ms);
                    image.Save(file_dic);
                    InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                    curr_rgn = curr_rgn.Offset[0, 1];
                }
            }
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
            //tar_wrksht.Paste(tar_range, sel_picture);
            sel_picture.LockAspectRatio = MsoTriState.msoTrue;
            sel_picture.Placement = myExcel.XlPlacement.xlMoveAndSize;
            sel_picture.Name = pic_name;

        }
        public AutoCompleteStringCollection Get_zone_infor_to_input_image(myExcel.Workbook wb)
        {
            myExcel.Worksheet ws = wb.Sheets[1];
            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 20; i < 400; i++)
            {
                if (ws.Cells[i, 2].Value != null)
                {
                    if ((ws.Cells[i, 2].Value.Contains("Region/Zone")))
                    {
                        if (ws.Cells[i, 4].Value != null)
                        {
                            list.Add(ws.Cells[i, 4].Value.ToString() + "+" + (i + 1).ToString());
                        }
                        if (ws.Cells[i, 3].Value != null)
                        {
                            list.Add(ws.Cells[i, 3].Value.ToString() + "+" + (i + 1).ToString());
                        }

                    }
                    if (ws.Cells[i, 2].Value.Contains("PTH_X-section"))
                    {
                        list.Add(ws.Cells[i, 2].Value.ToString() + "+" + (i).ToString());
                    }
                    if (ws.Cells[i, 2].Value.Contains("BVH_X-section"))
                    {
                        list.Add(ws.Cells[i, 2].Value.ToString() + "+" + (i).ToString());
                    }
                }
                if (ws.Cells[i, 3].Value != null)
                {
                    if (ws.Cells[i, 3].Value.ToString() == "Region/Zone" && ws.Cells[i, 4].Value != "")
                    {
                        list.Add(ws.Cells[i, 4].Value.ToString() + "_" + (i + 1).ToString());
                    }
                }
            }
            return list;
        }
        private void tsmExport_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtOperator.Text != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                List<string> tar_table_name = new List<string>() { "STACKUP_IMAGE", "STACKUP_LOGFILE" };
                List<DataGridView> tar_DGV = new List<DataGridView>() { DGV_Image_Graph, DGV_Stackup };
                for (int i = 0; i < tar_DGV.Count; i++)
                {
                    tar_DGV[i].DataSource = TDMK_Code.Datatable_Filter(sqlcon, tar_table_name[i], filter_str);

                    myCode.DGV_Auto_Resize(tar_DGV[i]);
                }

                ((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)DGV_Image_Graph.Columns["Image_data"]).Width = 100;

                foreach (DataGridViewRow dr in DGV_Image_Graph.Rows)
                {
                    dr.Height = 70;
                }

                if (DGV_Image_Graph.Rows.Count > 0 && DGV_Stackup.Rows.Count > 0)
                {
                    string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                    if (format_file != "")
                    {
                        string export_path = F_exportEPPlus.get_reportpath(txtItemCode.Text, txtLotNo.Text, txtOperator.Text, "Stackup", myVar.data_loc);
                        if (export_path != "")
                        {
                            F_exportEPPlus.Export_EPPlus(sqlcon, format_file, export_path, txtItemCode.Text, txtLotNo.Text, "STACKUP", null);
                        }
                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Folder Stackup not found!", "Warning");
                        }
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Format not found!", "Warning");
                    }
                }
                else
                {
                    MessageBox.Show("No data", "Warning");
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Please, fill in ItemCode / LotNo / Operator ", "Warning");
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
        public void Get_Stackup_Image_Multi(string in_src, ref Dictionary<string, SortedDictionary<int, byte[]>> dic_lst_result)
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

                        if (f_na.Contains("-") == false && f_na.Contains("+") == false)
                        {

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
                            Get_Stackup_Image_Multi(inter_lst.FullName, ref dic_lst_result);
                        }
                    }
                }
            }
            catch
            {

            }
        }
        public void Get_Stackup_Image_Multi_Update(string in_src, ref Dictionary<string, SortedDictionary<int, byte[]>> dic_lst_result, string zone)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            try
            {
                FileInfo[] temp_lst = tar_d.GetFiles("*.jpg");
                if (temp_lst.Length > 0)
                {
                    if (tar_d.Name == zone)
                    {
                        SortedDictionary<int, byte[]> lst_result = new SortedDictionary<int, byte[]>();
                        for (int i = 0; i < temp_lst.Length; i++)
                        {
                            string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
                            if (f_na.Contains("-") == false && f_na.Contains("+") == false)
                            {
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
                        }
                        dic_lst_result.Add(tar_d.Name, lst_result);

                    }

                }
                else
                {
                    DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                    if (inter_type_lst.Length > 0)
                    {
                        foreach (var inter_lst in inter_type_lst)
                        {
                            Get_Stackup_Image_Multi_Update(inter_lst.FullName, ref dic_lst_result, zone);
                        }
                    }
                }
            }
            catch
            {

            }
        }

        public void Get_Stackup_Image_Multi_1(string in_src, ref Dictionary<string, SortedDictionary<int, byte[]>> dic_lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);

            FileInfo[] temp_lst = tar_d.GetFiles("*.jpg");
            if (temp_lst.Length > 0)
            {
                SortedDictionary<int, byte[]> lst_result = new SortedDictionary<int, byte[]>();
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
                    if (!f_na.Contains("-"))
                    {
                        var sel_img = Bitmap.FromFile(temp_lst[i].FullName);
                        ImageConverter imgcon = new ImageConverter();
                        byte[] img_data = (byte[])imgcon.ConvertTo(sel_img, typeof(byte[]));
                        int f_inx = Convert.ToInt32(f_na);
                        lst_result.Add(f_inx, img_data);
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
                        Get_Stackup_Image_Multi_1(inter_lst.FullName, ref dic_lst_result);
                    }
                }
            }
        }
        public void Get_Stackuplogfile_Multi(string in_src, ref SortedDictionary<string, SortedDictionary<int, string>> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            FileInfo[] temp_lst = tar_d.GetFiles("*.csv");
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    SortedDictionary<int, string> temp = Get_StackupLogFile_Data(temp_lst[i].FullName);
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });

                    int f_inx = Convert.ToInt32(f_na);
                    lst_result.Add(f_na, temp);
                }
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_Stackuplogfile_Multi(inter_lst.FullName, ref lst_result);
                    }
                }
            }
        }
        public void Get_Stackuplogfile_Multi2(string in_src, ref Dictionary<string, Dictionary<string, SortedDictionary<int, string>>> dic_lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            try
            {
                FileInfo[] temp_lst = tar_d.GetFiles("*.csv");
                if (temp_lst.Length > 0)
                {
                    Dictionary<string, SortedDictionary<int, string>> lst_result = new Dictionary<string, SortedDictionary<int, string>>();
                    for (int i = 0; i < temp_lst.Length; i++)
                    {

                        SortedDictionary<int, string> temp = Get_StackupLogFile_Data(temp_lst[i].FullName);
                        string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
                        if (!f_na.Contains("-") && !f_na.Contains("+"))
                        {
                            if (myCode.IsNumeric(f_na))
                            {
                                int f_inx = Convert.ToInt32(f_na);
                                lst_result.Add(f_na, temp);
                            }
                        }


                        //lst_result.Add(f_na, temp);
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
                            Get_Stackuplogfile_Multi2(inter_lst.FullName, ref dic_lst_result);
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show(new Form { TopMost = true }, "Could not find a part of the path " + in_src, "Warning");
            }
        }
        public void Get_Stackuplogfile_Multi2_Update(string in_src, string zone, ref Dictionary<string, Dictionary<string, SortedDictionary<int, string>>> dic_lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            try
            {
                FileInfo[] temp_lst = tar_d.GetFiles("*.csv");
                if (temp_lst.Length > 0)
                {
                    if (tar_d.Name == zone)
                    {
                        Dictionary<string, SortedDictionary<int, string>> lst_result = new Dictionary<string, SortedDictionary<int, string>>();
                        for (int i = 0; i < temp_lst.Length; i++)
                        {
                            SortedDictionary<int, string> temp = Get_StackupLogFile_Data(temp_lst[i].FullName);
                            string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
                            if (f_na.Contains("-") == false && f_na.Contains("+") == false)
                            {

                                if (myCode.IsNumeric(f_na))
                                {
                                    int f_inx = Convert.ToInt32(f_na);
                                    lst_result.Add(f_na, temp);

                                }
                            }


                        }
                        dic_lst_result.Add(tar_d.Name, lst_result);
                    }

                }
                else
                {
                    DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                    if (inter_type_lst.Length > 0)
                    {
                        foreach (var inter_lst in inter_type_lst)
                        {
                            Get_Stackuplogfile_Multi2_Update(inter_lst.FullName, zone, ref dic_lst_result);
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show(new Form { TopMost = true }, "Could not find a part of the path " + in_src, "Warning");
            }
        }
        public SortedDictionary<int, string> Get_StackupLogFile_Data(string in_src_file)
        {
            SortedDictionary<int, string> result_lst = new SortedDictionary<int, string>();
            string[] Lines = System.IO.File.ReadAllLines(in_src_file);
            foreach (var line in Lines)
            {
                string str = new string(line.Where(s => s != '"').ToArray());
                string[] temp = str.Split(',');
                if (temp[0] != "")
                {
                    if (myCode.IsNumeric(temp[0]))
                    {
                        result_lst.Add(Convert.ToInt32(temp[0]), temp[2]);
                    }
                }
                else
                {
                    break;
                }
            }
            return result_lst;
        }
        private void DGV_Stackup_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //int r_inx = e.RowIndex;
            //int c_inx = e.ColumnIndex;
            //DataGridViewCell Cur_cell = DGV_Stackup.CurrentCell;
            //if (DGV_Stackup.Columns[c_inx].Name == "Region")
            //{
            //    string region = Cur_cell.Value.ToString();
            //    lblZone.Text = "Zone Data Detail: " + region;
            //    DataTable src_dt = (DataTable)DGV_Stackup.DataSource;
            //    DataTable filter_tbl = src_dt.AsEnumerable().Where(x => x.Field<string>("Region") == region).CopyToDataTable();
            //    List<string> pcs_lst = new List<string>();
            //    exp_proc.Get_List_data(-1, filter_tbl, new string[] { "ItemCode", "LotNo" }, ref pcs_lst, "Pcs_No", true);
            //    DataTable detail_tbl = new DataTable();
            //    detail_tbl.Columns.Add("Nominal");

            //    foreach (var c in pcs_lst)
            //    {
            //        detail_tbl.Columns.Add("Pcs_No" + c);
            //    }

            //    DataTable dt_spec_zone = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "STACKUP_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Zone" }, new string[] { txtItemCode.Text, Cur_cell.Value.ToString().Replace("Zone_", "") }));
            //    if (dt_spec_zone.Rows.Count > 0)
            //    { 
            //        string[] lst_spec_detail = dt_spec_zone.Rows[0]["Spec_detail"].ToString().Split(';');
            //        int row = 0;
            //        foreach (string sp in lst_spec_detail)
            //        {
            //            if(sp != "")
            //            {
            //                detail_tbl.Rows.Add();
            //                detail_tbl.Rows[row][0] = sp;
            //                row++;
            //            } 
            //        }

            //        foreach (var c in pcs_lst)
            //        {
            //            List<string> pcs_val = filter_tbl.AsEnumerable().Where(x => x.Field<string>("Pcs_No") == c).Select(x => x.Field<string>("Data")).ToList();
            //            int idx = 0;
            //            foreach(DataRow row_ in detail_tbl.Rows)
            //            {
            //                //foreach (var p in pcs_val)
            //                //{
            //                if(idx < pcs_val.Count)
            //                {
            //                    row_["Pcs_No" + c] = pcs_val[idx];
            //                    idx++;
            //                } 
            //               //}
            //            } 

            //        } 
            //        DataRow dr = detail_tbl.NewRow();
            //        for (int i = 1; i < detail_tbl.Columns.Count; i++)
            //        {
            //            dr[0] = "Toltal";
            //            double sum_pcs = 0;
            //            foreach (DataRow r in detail_tbl.Rows)
            //            {
            //                sum_pcs += double.Parse(r[i].ToString());
            //            }
            //            dr[i] = sum_pcs;
            //        }
            //        detail_tbl.Rows.Add(dr);

            //        DGV_Zone_Data.DataSource = detail_tbl;
            //        DGV_Zone_Data.Rows[DGV_Zone_Data.RowCount - 1].Cells[0].Style.BackColor = Color.Green;
            //        for (int i = 1; i < DGV_Zone_Data.Columns.Count; i++)
            //        {
            //            DGV_Zone_Data.Rows[DGV_Zone_Data.RowCount - 1].Cells[i].Style.BackColor = Color.Yellow;
            //        }


            //        myCode.DGV_Auto_Resize(DGV_Zone_Data);
            //        HightLigh_StackUp_LogFile_One_Zone(DGV_spec, DGV_Zone_Data, Cur_cell.Value.ToString().Replace("Zone_", ""));
            //        DGV_Zone_Data.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //        DGV_Zone_Data.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            //    }
            //    else
            //    {
            //        MessageBox.Show("Please setup spec", "Warning");
            //    }
            //}
        }

        public bool IsNumber(string input)
        {
            double result;
            return double.TryParse(input, out result);
        }

        public void HightLigh_StackUp_LogFile_One_Zone(DataGridView DGV_Spec, DataGridView DGV_Data, string zone)
        {
            if (DGV_Spec.Rows.Count > 2)
            {
                if (DGV_Data.Rows.Count > 1)
                {
                    List<string> spec_col_lst = new List<string>();
                    foreach (DataGridViewColumn dgv_c in DGV_Spec.Columns)
                    {
                        spec_col_lst.Add(dgv_c.Name);
                    }
                    int col_inx = spec_col_lst.IndexOf(zone);
                    if (col_inx != -1)
                    {
                        for (int i = 0; i < DGV_Data.Columns.Count; i++)
                        {
                            double max = Convert.ToDouble(DGV_Spec.Rows[0].Cells[col_inx].Value.ToString());
                            double min = Convert.ToDouble(DGV_Spec.Rows[1].Cells[col_inx].Value.ToString());
                            //for (int m = 0; m < DGV_Data.Rows.Count; m++)
                            //{
                            int k = DGV_Data.Rows.Count - 1;
                            if (DGV_Data.Rows[k].Cells[i].Value != null)
                            {
                                if (IsNumber(DGV_Data.Rows[k].Cells[i].Value.ToString()))
                                {
                                    double x = Convert.ToDouble(DGV_Data.Rows[k].Cells[i].Value);
                                    if (x < min || x > max)
                                    {
                                        for (int m = 0; m < DGV_Data.Rows.Count - 1; m++)
                                        {
                                            DGV_Data.Rows[m].Cells[i].Style.BackColor = Color.White;
                                        }
                                        DGV_Data.Rows[k].Cells[i].Style.BackColor = Color.Red;
                                    }
                                    else
                                    {
                                        for (int m = 0; m < DGV_Data.Rows.Count - 1; m++)
                                        {
                                            //DGV_Data.Rows[m].Cells[i].Style.BackColor = Color.Red;
                                            DGV_Data.Rows[m].Cells[i].Style.BackColor = Color.White;
                                        }
                                    }
                                }

                            }
                            
                        }

                        for (int i = 0; i < DGV_Data.Rows.Count - 1; i++)
                        {
                            string nominal_val = DGV_Data.Rows[i].Cells[0].Value.ToString();
                            double min = 0;
                            double max = 0;
                            if (nominal_val.Contains("+") || nominal_val.Contains("-"))
                            {
                                nominal_val = nominal_val.Replace("+", ";").Replace("-", ";").Replace("/", ";");
                                max = double.Parse(nominal_val.Split(';')[0]) + double.Parse(nominal_val.Split(';')[1]);
                                min = double.Parse(nominal_val.Split(';')[0]) - double.Parse(nominal_val.Split(';')[3]);

                            }
                            else if (nominal_val.Contains("±"))
                            {
                                max = double.Parse(nominal_val.Split('±')[0]) + double.Parse(nominal_val.Split('±')[1]);
                                min = double.Parse(nominal_val.Split('±')[0]) - double.Parse(nominal_val.Split('±')[1]);
                            }
                            else if (nominal_val.Contains("/"))
                            {
                                if (myCode.IsNumeric(nominal_val.Split('/')[1]))
                                {
                                    min = double.Parse(nominal_val.Split('/')[1]);
                                }
                                if (myCode.IsNumeric(nominal_val.Split('/')[0]))
                                {
                                    max = double.Parse(nominal_val.Split('/')[0]);
                                }

                            }
                            else
                            {
                                min = double.Parse(nominal_val) * 0.7;
                                max = double.Parse(nominal_val) * 1.3;
                            }


                            for (int j = 1; j < DGV_Data.Columns.Count; j++)
                            {
                                //if (j == 1)
                                //{
                                //    if (myCode.checkDBNull(DGV_Data.Rows[i].Cells[j].Value.ToString()) != "")
                                //    {
                                //        double data = double.Parse(DGV_Data.Rows[i].Cells[j].Value.ToString());
                                //        if (data > max || data < min)
                                //        {
                                //            DGV_Data.Rows[i].Cells[j].Style.BackColor = Color.Red;
                                //        }
                                //        else
                                //        {
                                //            DGV_Data.Rows[i].Cells[j].Style.BackColor = Color.White;
                                //        }
                                //    }
                                //}
                                if (myCode.checkDBNull(DGV_Data.Rows[i].Cells[j].Value.ToString()) != "")
                                {
                                    double data = double.Parse(DGV_Data.Rows[i].Cells[j].Value.ToString());
                                    if (data > max || data < min)
                                    {
                                        DGV_Data.Rows[i].Cells[j].Style.BackColor = Color.Red;
                                    }
                                    else
                                    {
                                        DGV_Data.Rows[i].Cells[j].Style.BackColor = Color.White;
                                    }
                                }
                                else
                                {
                                    DGV_Data.Rows[i].Cells[j].Style.BackColor = Color.Red;
                                }    
                            }
                        }
                    }
                    else
                    {
                        for (int m = 0; m < DGV_Data.Rows.Count; m++)
                        {
                            for (int i = 0; i < DGV_Data.Columns.Count; i++)
                            {
                                if (DGV_Data.Rows[m].Cells[i].Value != null)
                                {
                                    DGV_Data.Rows[m].Cells[i].Style.BackColor = Color.Gray;
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                for (int m = 0; m < DGV_Data.Rows.Count; m++)
                {
                    for (int i = 0; i < DGV_Data.Columns.Count; i++)
                    {
                        if (DGV_Data.Rows[m].Cells[i].Value != null)
                        {
                            DGV_Data.Rows[m].Cells[i].Style.BackColor = Color.Gray;
                        }
                    }
                }
                MessageBox.Show(new Form { TopMost = true }, "Spec not found!");
            }

        }
        public void HightLigh_StackUp_LogFile_One_Zone_new(DataGridView DGV_Spec, DataGridView DGV_Data, string zone, DataTable dt_spec_zone)
        {
            if (DGV_Spec.Rows.Count > 2)
            {
                if (DGV_Data.Rows.Count > 1)
                {
                    List<string> spec_col_lst = new List<string>();
                    foreach (DataGridViewColumn dgv_c in DGV_Spec.Columns)
                    {
                        spec_col_lst.Add(dgv_c.Name);
                    }
                    int col_inx = spec_col_lst.IndexOf(zone);
                    if (col_inx != -1)
                    {
                        for (int i = 0; i < DGV_Data.Columns.Count; i++)
                        {
                            double max = Convert.ToDouble(DGV_Spec.Rows[0].Cells[col_inx].Value.ToString());
                            double min = Convert.ToDouble(DGV_Spec.Rows[1].Cells[col_inx].Value.ToString());
                            //for (int m = 0; m < DGV_Data.Rows.Count; m++)
                            //{
                            int k = DGV_Data.Rows.Count - 1;
                            if (DGV_Data.Rows[k].Cells[i].Value != null)
                            {
                                if (IsNumber(DGV_Data.Rows[k].Cells[i].Value.ToString()))
                                {
                                    double x = Convert.ToDouble(DGV_Data.Rows[k].Cells[i].Value);
                                    if (x < min || x > max)
                                    {
                                        for (int m = 0; m < DGV_Data.Rows.Count - 1; m++)
                                        {
                                            DGV_Data.Rows[m].Cells[i].Style.BackColor = Color.White;
                                        }
                                        DGV_Data.Rows[k].Cells[i].Style.BackColor = Color.Red;
                                    }
                                    else
                                    {
                                        for (int m = 0; m < DGV_Data.Rows.Count - 1; m++)
                                        {
                                            //DGV_Data.Rows[m].Cells[i].Style.BackColor = Color.Red;
                                            DGV_Data.Rows[m].Cells[i].Style.BackColor = Color.White;
                                        }
                                    }

                                }

                            }
                            //  }
                        }
                    }
                    else
                    {
                        for (int m = 0; m < DGV_Data.Rows.Count; m++)
                        {
                            for (int i = 0; i < DGV_Data.Columns.Count; i++)
                            {
                                if (DGV_Data.Rows[m].Cells[i].Value != null)
                                {
                                    DGV_Data.Rows[m].Cells[i].Style.BackColor = Color.Gray;
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                for (int m = 0; m < DGV_Data.Rows.Count; m++)
                {
                    for (int i = 0; i < DGV_Data.Columns.Count; i++)
                    {
                        if (DGV_Data.Rows[m].Cells[i].Value != null)
                        {
                            DGV_Data.Rows[m].Cells[i].Style.BackColor = Color.Gray;
                        }
                    }
                }
                MessageBox.Show(new Form { TopMost = true }, "Spec not found!");
            }

        }

        private void DGV_Zone_Data_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                DGV_Zone_Data.ContextMenuStrip = mnuZoneSpec;
                mnuZoneSpec.Show(e.Location);
            }
        }

        private void tsmSetupSpec_Click(object sender, EventArgs e)
        {
            Spec_Setup frmSpec = new Spec_Setup();
            frmSpec.Show();
        }

        public void setup_qty_stackup(string file_format, string ItemCode)
        {
            myExcel.Workbook wb = TDMK_Code.open_excel_file(file_format, "", "");
            string sheet_name = "";
            string tar_sheet = "Stack-up";
            for (int i = 0; i < wb.Sheets.Count; i++)
            {
                myExcel.Worksheet cur_sht = wb.Sheets[i + 1];
                string cur_sht_name = cur_sht.Name;
                if (tar_sheet.Replace("-", "").ToUpper() == cur_sht_name.Replace("-", "").Replace(" ", "").ToUpper())
                {
                    sheet_name = cur_sht_name;
                    break;
                }
            }
            if (sheet_name != "")
            {
                int count_sample = 0;
                for (int row = 2; row <= 100; row++)
                {

                    myExcel.Range cell = wb.Sheets[sheet_name].Cells[row, 1];
                    if (cell.Value2 != null && (cell.Value2.ToString().ToUpper().Contains("ZONE")))
                    {
                        for (int col = 1; col < 20; col++)
                        {
                            string cell_val = myCode.checkDBNull(wb.Sheets[sheet_name].Cells[row, col].Value);
                            if (cell_val.Contains("#"))
                            {
                                count_sample++;
                            }
                            else if (cell_val == "")
                            {
                                break;
                            }
                        }
                        break;
                    }
                }
                int ID = TDMK_Code.SQL_MAX("SETTING_PCS", "ID", sqlcon) + 1;
                TDMK_Code.insert_val_arr("SETTING_PCS", sqlcon, new string[] { "ID", "ItemCode", "Data_For", "Region_Zone", "Pcs_setting" }, new string[] { ID.ToString(), ItemCode, "Stack-up", "", count_sample.ToString() });

            }

        }

        private void btn_load_spec_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "")
            {
            //btn_savespec.Enabled = false;
            lbl_start_setup_spec:
                AutoCompleteStringCollection list_zone = TDMK_Code.Load_Item_Filter_str(sqlcon, "STACKUP_SPEC", "Zone", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
                if (DGV_spec.Rows.Count > 1 && list_zone.Count > 0)
                {
                    //DGV_spec.Rows.Clear();
                    //DGV_spec.Columns.Clear();
                    //DGV_spec.Columns.Add("Zone", "Zone"); 

                    if (MessageBox.Show(new Form { TopMost = true }, "Do you want to update spec?", "Data already exists!", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        TDMK_Code.Delelte_FilteredItem_arr("STACKUP_SPEC", sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
                        TDMK_Code.Delelte_FilteredItem_arr("SETTING_PCS", sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode", "Data_For" }, new string[] { txtItemCode.Text, "Stack-up" }));
                        DGV_spec.Rows.Clear();
                        DGV_spec.Columns.Clear();
                        DGV_spec.Columns.Add("Zone", "Zone");
                        DGV_Zone_Data.DataSource = null;
                        goto lbl_start_setup_spec;
                    }
                    else
                    {
                        return;
                    }

                }
                if (list_zone.Count > 0)
                {

                    DGV_spec.Rows.Add("USL", "USL");
                    DGV_spec.Rows.Add("LSL", "LSL");
                    DGV_spec.Rows.Add("Normal", "Normal");
                    DGV_spec.Rows.Add();
                    foreach (string zone in list_zone)
                    {
                        DGV_spec.Columns.Add(zone, zone);
                        AutoCompleteStringCollection usl = TDMK_Code.Load_Item_Filter_str(sqlcon, "STACKUP_SPEC", "USL", TDMK_Code.filter_str(new string[] { "ItemCode", "Zone" }, new string[] { txtItemCode.Text, zone }));
                        AutoCompleteStringCollection lsl = TDMK_Code.Load_Item_Filter_str(sqlcon, "STACKUP_SPEC", "LSL", TDMK_Code.filter_str(new string[] { "ItemCode", "Zone" }, new string[] { txtItemCode.Text, zone }));
                        AutoCompleteStringCollection normal = TDMK_Code.Load_Item_Filter_str(sqlcon, "STACKUP_SPEC", "Normal", TDMK_Code.filter_str(new string[] { "ItemCode", "Zone" }, new string[] { txtItemCode.Text, zone }));
                        DGV_spec.Rows[0].Cells[zone].Value = usl[0];
                        DGV_spec.Rows[1].Cells[zone].Value = lsl[0];
                        DGV_spec.Rows[2].Cells[zone].Value = normal[0];
                    }
                    DGV_spec.Rows.RemoveAt(DGV_spec.Rows.Count - 1);
                    //stackup_logfile.HightLigh_StackUp_LogFile(DGV_spec, DGV_Zone_Data);
                    DGV_spec.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                }
                else
                {

                    MessageBox.Show(new Form { TopMost = true }, "Please wait for setup spec", "Warning");

                    string file_format = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                    if (file_format != "")
                    {
                        //ExcelWorkbook wb = TDMK_Code2.open_excel_file(file_format);
                        //F_setup_new.setup_Stackup_new(wb, txtItemCode.Text, sqlcon);


                        SortedDictionary<string, string> dic_spec_detail = new SortedDictionary<string, string> { };
                        Load_Spec_StackUp_To_DGV(file_format, DGV_spec, ref dic_spec_detail);
                        Save_Spec_new(dic_spec_detail);


                        MessageBox.Show(new Form { TopMost = true }, "Setup spec completed", "Warning");
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Format of " + txtItemCode.Text + " not found", "Warning");
                    }
                }

                if (DGV_spec.Rows.Count > 1 && DGV_Stackup.Rows.Count > 1)
                {
                    string[] arr_zone = ((DataTable)DGV_Stackup.DataSource).AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();
                    List<string> spec_col_lst = new List<string>();
                    foreach (DataGridViewColumn dgv_c in DGV_spec.Columns)
                    {
                        spec_col_lst.Add(dgv_c.Name);
                    }

                    foreach (string zone in arr_zone)
                    {
                        int col_inx = spec_col_lst.IndexOf(zone.Replace("Zone_", ""));
                        if (col_inx != -1)
                        {
                            double max = Convert.ToDouble(DGV_spec.Rows[0].Cells[col_inx].Value.ToString());
                            double min = Convert.ToDouble(DGV_spec.Rows[1].Cells[col_inx].Value.ToString());

                            DataTable dt_spec_zone = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "STACKUP_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Zone" }, new string[] { txtItemCode.Text, zone.Replace("Zone_", "") }));
                            if (dt_spec_zone.Rows.Count > 0)
                            {
                                string[] lst_spec_detail = dt_spec_zone.Rows[0]["Spec_detail"].ToString().Split(';');
                                DataView dv = ((DataTable)DGV_Stackup.DataSource).AsDataView();

                                string filter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { zone });
                                dv.RowFilter = filter;
                                string[] arr_pcs = dv.ToTable().AsEnumerable().Select(x => x.Field<string>("Pcs_No")).Distinct().ToArray();
                                foreach (string pcs in arr_pcs)
                                {
                                    DataView dv_filter = ((DataTable)DGV_Stackup.DataSource).AsDataView();
                                    dv_filter.RowFilter = TDMK_Code.filter_str(new string[] { "Region", "Pcs_No" }, new string[] { zone, pcs });
                                    check_spec_zone_pcs(lst_spec_detail, dv_filter, max, min, pcs);
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Please, fill in ItemCode", "Warning");
            }
        }

        public void Save_Spec()
        {
            string[] item = new string[] { "ItemCode", "LotNo", "Zone", "USL", "LSL", "Normal", "ID" };
            string[] value = new string[7];
            value[0] = txtItemCode.Text;
            //value[1] = txtLotNo.Text;
            for (int i = 1; i < DGV_spec.Columns.Count; i++)
            {
                value[2] = DGV_spec.Columns[i].Name;
                value[3] = (DGV_spec.Rows[0].Cells[i].Value).ToString();
                value[4] = (DGV_spec.Rows[1].Cells[i].Value).ToString();
                value[5] = (DGV_spec.Rows[2].Cells[i].Value).ToString();
                value[6] = (TDMK_Code.SQL_MAX("STACKUP_SPEC", "ID", sqlcon) + 1).ToString();
                TDMK_Code.insert_val_arr("STACKUP_SPEC", sqlcon, item, value);
            }

        }
        public void Save_Spec_new(SortedDictionary<string, string> dic_spec_detail)
        {
            string[] item = new string[] { "ID", "ItemCode", "Zone", "USL", "LSL", "Normal", "Spec_detail" };
            string[] value = new string[7];
            value[1] = txtItemCode.Text;
            //value[1] = txtLotNo.Text;
            for (int i = 1; i < DGV_spec.Columns.Count; i++)
            {
                value[0] = (TDMK_Code.SQL_MAX("STACKUP_SPEC", "ID", sqlcon) + 1).ToString();
                value[2] = DGV_spec.Columns[i].Name;
                value[3] = (DGV_spec.Rows[0].Cells[i].Value).ToString();
                value[4] = (DGV_spec.Rows[1].Cells[i].Value).ToString();
                value[5] = (DGV_spec.Rows[2].Cells[i].Value).ToString();
                value[6] = dic_spec_detail[DGV_spec.Columns[i].Name];

                TDMK_Code.insert_val_arr("STACKUP_SPEC", sqlcon, item, value);
            }

        }
        private void btn_save_spec_Click(object sender, EventArgs e)
        {
            // sqlcon = stackup_logfile.ConnectDB();
            if (DGV_spec.Rows.Count < 2)
            {
                MessageBox.Show(new Form { TopMost = true }, "No data to save!");
            }
            else
            {
                AutoCompleteStringCollection list = TDMK_Code.Load_Item_Filter_str(sqlcon, "STACKUP_SPEC", "ID", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
                if (list.Count > 2)
                {
                    DialogResult result = MessageBox.Show(new Form { TopMost = true }, "Do you want to update spec?", "Data already exists!", MessageBoxButtons.OKCancel);
                    if (result == DialogResult.OK)
                    {
                        TDMK_Code.Delelte_FilteredItem_arr("STACKUP_SPEC", sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
                        Save_Spec();
                        MessageBox.Show(new Form { TopMost = true }, "Inserted!");
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Inserted!");
                    }
                }
                else
                {
                    Save_Spec();
                    MessageBox.Show(new Form { TopMost = true }, "Inserted!");
                }
                //btn_savespec.Enabled = false;
            }
        }

        private void DGV_spec_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //stackup_logfile.HightLigh_StackUp_LogFile(DGV_spec, DGV_Zone_Data);
            if (DGV_spec.Rows.Count > 0)
            {
                //  btn_savespec.Enabled = true;
            }

        }

        private void txtLogfile_TextChanged(object sender, EventArgs e)
        {

        }

        private void tblZoneTitle_Paint(object sender, PaintEventArgs e)
        {

        }

        private void DGV_Zone_Data_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            //if (e.Button == MouseButtons.Right)
            //{
            //    if (DGV_Zone_Data.SelectedCells.Count > 0)
            //    {
            //        SortedDictionary<int, List<string>> sel_cells_lst = new SortedDictionary<int, List<string>>();
            //        foreach (DataGridViewCell cells in DGV_Zone_Data.SelectedCells)
            //        {
            //            int r_inx = cells.RowIndex;
            //            int c_inx = cells.ColumnIndex;
            //            string col_name = DGV_Zone_Data.Columns[c_inx].Name;
            //            string _pcs_no = new string(col_name.Where(char.IsDigit).ToArray());
            //            //string _pcs_no = exp_proc.Extract_Number_from_String(col_name);
            //            int pcs_no = Convert.ToInt32(_pcs_no);
            //            if (sel_cells_lst.Keys.ToList().IndexOf(pcs_no) != -1)
            //            {
            //                sel_cells_lst[pcs_no].Add(r_inx.ToString());
            //            }
            //            else
            //            {
            //                sel_cells_lst.Add(pcs_no, new List<string>() { r_inx.ToString() });
            //            }
            //        }
            //        if (sel_cells_lst.Count > 0)
            //        {
            //            string cur_zone = lblZone.Text.Split(':')[1].Trim();
            //            DataTable dt = (DataTable)DGV_Stackup.DataSource;
            //            DataView dv = dt.AsDataView();
            //            foreach (var item in sel_cells_lst)
            //            {
            //                string filter = TDMK_Code.filter_str(new string[] { "Region", "PCS_No" }, new string[] { cur_zone, item.Key.ToString() });
            //                dv.RowFilter = filter;
            //                foreach (var p in item.Value)
            //                {
            //                    int inx = Convert.ToInt32(p);
            //                    DataRow dr = dv[inx].Row;
            //                    int r_inx = dt.Rows.IndexOf(dr);
            //                    DGV_Stackup.Rows[r_inx].Cells["Data"].Style.BackColor = Color.Red;
            //                    dt.Rows[r_inx]["Data"] = dt.Rows[r_inx]["Data"].ToString() + "_OK";
            //                }
            //            }
            //        }
            //    }
            // }

        }

        private void DGV_Zone_Data_DataSourceChanged(object sender, EventArgs e)
        {
            //edited_lst = new SortedDictionary<int, List<string>>();
        }
        private void SetValue(String value)
        {
            this.txtLogfile.Text = value;
        }

        private void DGV_Zone_Data_MouseDown_1(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (txtItemCode.Text != "" && txtLotNo.Text != "")
                {
                    DataTable logfile_before = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                    if (logfile_before.Rows.Count > 0)
                    {

                        // admin_mode = true;
                        if (admin_mode)
                        {
                            if (DGV_Zone_Data.SelectedCells.Count > 0)
                            {

                                sel_pcs_lst = new List<string> { };
                                foreach (DataGridViewCell cells in DGV_Zone_Data.SelectedCells)
                                {

                                    int c_inx = cells.ColumnIndex;
                                    string col_name = DGV_Zone_Data.Columns[c_inx].Name;
                                    string _pcs_no = new string(col_name.Where(char.IsDigit).ToArray());

                                    //  int pcs_no = Convert.ToInt32(_pcs_no);
                                    if (sel_pcs_lst.IndexOf(_pcs_no) == -1)
                                    {
                                        sel_pcs_lst.Add(_pcs_no);
                                    }

                                }
                                if (sel_pcs_lst.Count > 0)
                                {
                                    cur_zone = lblZone.Text.Split(':')[1].Trim();

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

                                        }
                                        if (txtLogfile.Text != "")
                                        {
                                            btn_load.Enabled = true;
                                            update_mode = true;
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
        public void Stackup_history()
        {
            DataTable logfile_before = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
            DataTable image_before = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));

            List<string> item_arr = new List<string>() { "ID", "ItemCode", "LotNo", "Process", "PCS_No", "Region", "Zone", "Before_Data", "After_Data", "Time_Update", "Operator", "Depart", "Remark" };
            List<object> item_arr_val = new List<object>();
            DateTime tim_up = DateTime.Now;
            int id_edit = TDMK_Code.SQL_MAX("VHX_Edit_History", "ID", sqlcon);
            int count = 0;
            DataTable logfile_after = (DataTable)DGV_Stackup.DataSource;
            if (edited_lst.Count > 0)
            {
                foreach (var zone in edited_lst)
                {
                    foreach (var pcs in zone.Value)
                    {

                        DataView dv = logfile_after.AsDataView();
                        string filter = TDMK_Code.filter_str(new string[] { "Region", "PCS_No" }, new string[] { "Zone_" + zone.Key, pcs });
                        dv.RowFilter = filter;
                        for (int inx = 0; inx < dv.Count; inx++)
                        {
                            DataRow dr = dv[inx].Row;
                            int r_x = logfile_after.Rows.IndexOf(dr);
                            string before_data = logfile_before.Rows[r_x]["Data"].ToString();
                            string after_data = logfile_after.Rows[r_x]["Data"].ToString();
                            string remark = logfile_after.Rows[r_x]["Remark"].ToString();
                            byte[] bef_data = System.Text.Encoding.UTF8.GetBytes(before_data);
                            byte[] aft_data = System.Text.Encoding.UTF8.GetBytes(after_data);

                            item_arr_val = new List<object>() { (id_edit + count + 1), txtItemCode.Text, txtLotNo.Text, "Stackup", pcs, "Logfile_data", "Zone_" + zone.Key, bef_data, aft_data, tim_up, txtUsername.Text, depart, remark };
                            exp_proc.Insert_Object_List("VHX_Edit_History", sqlcon, item_arr, item_arr_val);
                            count++;

                        }

                        DataTable image_after = (DataTable)DGV_Image_Graph.DataSource;
                        DataView dv_image = image_after.AsDataView();
                        dv_image.RowFilter = filter;

                        if (dv_image.Count > 0)
                        {
                            DataRow dr_ = dv_image[0].Row;
                            int r_x_ = image_after.Rows.IndexOf(dr_);

                            byte[] img_bef = (byte[])(image_before.Rows[r_x_]["Image_Data"]);
                            byte[] img_aft = (byte[])(image_after.Rows[r_x_]["Image_Data"]);
                            string remark1 = image_after.Rows[r_x_]["Remark"].ToString();
                            item_arr_val = new List<object>() { (id_edit + count + 1), txtItemCode.Text, txtLotNo.Text, "Stackup", pcs, "Image_data", "Zone_" + zone.Key, img_bef, img_aft, tim_up, txtUsername.Text, depart, remark1 };
                            exp_proc.Insert_Object_List("VHX_Edit_History", sqlcon, item_arr, item_arr_val);
                            count++;
                        }


                    }

                }

            }

        }

        private void btn_load_Click(object sender, EventArgs e)
        {
            if (txtLogfile.Text != "" && txtItemCode.Text != "" && txtLotNo.Text != "" && txtOperator.Text != "")
            {
                add_new_mode = true;
                if (add_new_mode)
                {
                    Logfile_Data_Process_Loadnew();
                    //add_new_mode = false;
                }
                //btn_load.Enabled = false;
                else if (update_mode)
                {
                    DirectoryInfo temp_dir = new DirectoryInfo(txtLogfile.Text);
                    string temp_folder = temp_dir.Parent.Name;
                    string temp_zone = temp_dir.Name;
                    string f_name = "";
                    if (temp_zone.Contains(txtItemCode.Text))
                    {
                        f_name = temp_zone;
                    }
                    else if (temp_folder.Contains(txtItemCode.Text))
                    {
                        f_name = temp_folder;
                    }

                    //string f_name = temp_dir.Name;
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

                        if (f_name == temp_zone)
                        {
                            StackUp_Data_Process_Update(txtLogfile.Text, sel_pcs_lst, cur_zone);
                        }
                        else if (f_name == temp_folder)
                        {
                            string path_new = Directory.GetParent(temp_dir.FullName).ToString();
                            StackUp_Data_Process_Update(path_new, sel_pcs_lst, cur_zone);

                        }

                    }
                    else
                    {
                        MessageBox.Show("ItemCode or LotNo not matched", "Warning");
                    }
                    //update_mode = false;
                    //  btn_load.Enabled = false;
                }

                if (DGV_spec.Rows.Count > 1 && DGV_Stackup.Rows.Count > 0)
                {
                    string[] arr_zone = ((DataTable)DGV_Stackup.DataSource).AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();
                    List<string> spec_col_lst = new List<string>();
                    foreach (DataGridViewColumn dgv_c in DGV_spec.Columns)
                    {
                        spec_col_lst.Add(dgv_c.Name);
                    }

                    foreach (string zone in arr_zone)
                    {
                        int col_inx = spec_col_lst.IndexOf(zone.Replace("Zone_", ""));


                        if (col_inx != -1)
                        {
                            double max = Convert.ToDouble(DGV_spec.Rows[0].Cells[col_inx].Value.ToString());
                            double min = Convert.ToDouble(DGV_spec.Rows[1].Cells[col_inx].Value.ToString());

                            DataTable dt_spec_zone = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "STACKUP_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Zone" }, new string[] { txtItemCode.Text, zone.Replace("Zone_", "") }));
                            if (dt_spec_zone.Rows.Count > 0)
                            {
                                string[] lst_spec_detail = dt_spec_zone.Rows[0]["Spec_detail"].ToString().Split(';');


                                DataView dv = ((DataTable)DGV_Stackup.DataSource).AsDataView();

                                string filter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { zone });
                                dv.RowFilter = filter;
                                string[] arr_pcs = dv.ToTable().AsEnumerable().Select(x => x.Field<string>("Pcs_No")).Distinct().ToArray();
                                foreach (string pcs in arr_pcs)
                                {
                                    DataView dv_filter = ((DataTable)DGV_Stackup.DataSource).AsDataView();
                                    dv_filter.RowFilter = TDMK_Code.filter_str(new string[] { "Region", "Pcs_No" }, new string[] { zone, pcs });
                                    check_spec_zone_pcs(lst_spec_detail, dv_filter, max, min, pcs);

                                }
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Please, fill in ItemCode / LotNo / Operator ", "Warning");
            }

        }



        private void txtItemCode_Edit_TextChanged(object sender, EventArgs e)
        {

        }

        private void cbProcess_Edit_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void DGV_Image_Edit_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DGV_Image_Edit_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btn_load_edit_Click(object sender, EventArgs e)
        {
            if (cbProcess_Edit.SelectedIndex != -1 && txtItemCode_Edit.Text != "" && txtLotNo_Edit.Text != "" || cbProcess_Edit.Text == "ALL")
            {
                Search_Editted_Data();
            }
        }
        public void Search_Editted_Data()
        {
            DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { txtItemCode_Edit.Text, txtLotNo_Edit.Text, "Stackup" }));
            if (cbProcess_Edit.SelectedIndex != 0)
            {
                dt = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process", "Zone" }, new string[] { txtItemCode_Edit.Text, txtLotNo_Edit.Text, "Stackup", cbProcess_Edit.SelectedItem.ToString() }));
            }
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
            DataTable dt_img = src_dt.AsEnumerable().Where(x => x.Field<string>("Region") == "Image_data").CopyToDataTable();
            DataTable dt_data = src_dt.AsEnumerable().Where(x => x.Field<string>("Region") != "Image_data").CopyToDataTable();
            img_dt = dt_img.AsDataView().ToTable(false, new string[] { "Zone", "PCS_No", "Before_Data", "After_Data", "Time_Update", "Operator", "Depart", "Remark" });
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
            data_dt = disp_dt.AsDataView().ToTable(false, new string[] { "Zone", "PCS_No", "Before_Data", "After_Data", "Time_Update", "Operator", "Depart", "Remark" });
        }

        private void cbProcess_Edit_SelectedIndexChanged_1(object sender, EventArgs e)
        {
        }

        private void cbProcess_Edit_Click(object sender, EventArgs e)
        {
            if (txtItemCode_Edit.Text != "" && txtLotNo_Edit.Text != "")
            {
                cbProcess_Edit.Items.Clear();
                cbProcess_Edit.Items.Add("ALL");
                DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { txtItemCode_Edit.Text, txtLotNo_Edit.Text, "Stackup" }));
                string[] arr_zone = dt.AsEnumerable().Select(x => x.Field<string>("Zone")).Distinct().ToArray();

                foreach (string item in arr_zone)
                {
                    cbProcess_Edit.Items.Add(item);
                }
                cbProcess_Edit.SelectedIndex = 0;
            }
            else
            {
                MessageBox.Show("Please, fill in ItemCode and LotNo", "Warning");
            }

        }

        private void label7_Click(object sender, EventArgs e)
        {

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

        private void DGV_Image_Edit_CellContentDoubleClick_1(object sender, DataGridViewCellEventArgs e)
        {
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

        private void txtItemCode_DoubleClick(object sender, EventArgs e)
        {
            txtItemCode.Text = "";
        }

        private void txtLotNo_DoubleClick(object sender, EventArgs e)
        {
            txtLotNo.Text = "";
        }

        private void txtItemCode_TextChanged(object sender, EventArgs e)
        {
            DGV_Stackup.DataSource = DGV_Zone_Data.DataSource = DGV_Image_Graph.DataSource = DGV_spec.DataSource = null;
            //picDetails.Image = null;
            update_mode = false;
            add_new_mode = false;

            if (DGV_spec.Rows.Count > 0)
            {
                DGV_spec.Rows.Clear();
                DGV_spec.Columns.Clear();
                DGV_spec.Columns.Add("Zone", "Zone");
            }
        }

        private void txtLotNo_TextChanged(object sender, EventArgs e)
        {
            DGV_Stackup.DataSource = DGV_Zone_Data.DataSource = DGV_Image_Graph.DataSource = null;
            //  picDetails.Image = null;
            update_mode = false;
            add_new_mode = false;

            if (DGV_spec.Rows.Count > 0)
            {
                DGV_spec.Rows.Clear();
                DGV_spec.Columns.Clear();
                DGV_spec.Columns.Add("Zone", "Zone");
            }
        }

        private void txtItemCode_Edit_TextChanged_1(object sender, EventArgs e)
        {
            DGV_DataEdit.DataSource = DGV_Image_Edit.DataSource = null;
            //pic_after.Image = pic_before.Image = null;
        }

        private void txtLotNo_Edit_TextChanged(object sender, EventArgs e)
        {
            DGV_DataEdit.DataSource = DGV_Image_Edit.DataSource = null;
            //pic_after.Image = pic_before.Image = null;
        }

        private void txtLotNo_Edit_Validated(object sender, EventArgs e)
        {
            txtLotNo_Edit.Text = exp_proc.Lotno_Formated(txtLotNo_Edit.Text);
        }

        private void txtLotNo_Validated(object sender, EventArgs e)
        {
            txtLotNo.Text = exp_proc.Lotno_Formated(txtLotNo.Text);
        }

        private void DGV_Stackup_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DGV_Stackup_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int r_inx = e.RowIndex;
            int c_inx = e.ColumnIndex;
            if (r_inx != -1 && c_inx != -1)
            {
                DataGridViewCell Cur_cell = DGV_Stackup.CurrentCell;
                if (DGV_Stackup.Columns[c_inx].Name == "Region")
                {
                    string region = Cur_cell.Value.ToString();
                    lblZone.Text = "Zone Data Detail: " + region;
                    DataTable src_dt = (DataTable)DGV_Stackup.DataSource;
                    DataTable filter_tbl = src_dt.AsEnumerable().Where(x => x.Field<string>("Region") == region).CopyToDataTable();
                    List<string> pcs_lst = new List<string>();
                    exp_proc.Get_List_data(-1, filter_tbl, new string[] { "ItemCode", "LotNo" }, ref pcs_lst, "Pcs_No", true);
                    DataTable detail_tbl = new DataTable();
                    detail_tbl.Columns.Add("Nominal");

                    foreach (var c in pcs_lst)
                    {
                        detail_tbl.Columns.Add("Pcs_No" + c);
                    }

                    DataTable dt_spec_zone = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "STACKUP_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Zone" }, new string[] { txtItemCode.Text, Cur_cell.Value.ToString().Replace("Zone_", "") }));
                    if (dt_spec_zone.Rows.Count > 0)
                    {
                        string[] lst_spec_detail = dt_spec_zone.Rows[0]["Spec_detail"].ToString().Split(';');
                        int row = 0;
                        foreach (string sp in lst_spec_detail)
                        {
                            if (sp != "")
                            {
                                detail_tbl.Rows.Add();
                                detail_tbl.Rows[row][0] = sp;
                                row++;
                            }
                        }

                        foreach (var c in pcs_lst)
                        {
                            List<string> pcs_val = filter_tbl.AsEnumerable().Where(x => x.Field<string>("Pcs_No") == c).Select(x => x.Field<string>("Data")).ToList();
                            int idx = 0;
                            foreach (DataRow row_ in detail_tbl.Rows)
                            {
                                //foreach (var p in pcs_val)
                                //{
                                if (idx < pcs_val.Count)
                                {
                                    row_["Pcs_No" + c] = pcs_val[idx];
                                    idx++;
                                }
                                //}
                            }
                        }
                        DataRow dr = detail_tbl.NewRow();
                        for (int i = 1; i < detail_tbl.Columns.Count; i++)
                        {
                            dr[0] = "Total";
                            double sum_pcs = 0;
                            foreach (DataRow r in detail_tbl.Rows)
                            {
                                if (myCode.checkDBNull(r[i].ToString()) != "")
                                {
                                    sum_pcs += double.Parse(r[i].ToString());
                                }
                            }
                            dr[i] = sum_pcs;
                        }
                        detail_tbl.Rows.Add(dr);
                        DGV_Zone_Data.DataSource = null;
                        DGV_Zone_Data.DataSource = detail_tbl;
                        for (int i = 1; i < DGV_Zone_Data.Columns.Count; i++)
                        {
                            DGV_Zone_Data.Rows[DGV_Zone_Data.RowCount - 1].Cells[i].Style.BackColor = Color.GreenYellow;
                            //if (i != 1)
                            //{
                            //    DGV_Zone_Data.Rows[0].Cells[i].Value = "";
                            //}
                            //DGV_Zone_Data.Rows[DGV_Zone_Data.RowCount - 1].Cells[i].Style.BackColor = Color.White;
                        }

                        myCode.DGV_Auto_Resize(DGV_Zone_Data);
                        HightLigh_StackUp_LogFile_One_Zone(DGV_spec, DGV_Zone_Data, Cur_cell.Value.ToString().Replace("Zone_", ""));
                        DGV_Zone_Data.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                        DGV_Zone_Data.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                    }
                    else
                    {
                        MessageBox.Show("Please setup", "Warning");
                    }
                }
            }
        }
    }
}

