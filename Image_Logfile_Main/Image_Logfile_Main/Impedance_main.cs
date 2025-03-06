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
using System.Threading;
using System.Diagnostics.Eventing.Reader;
using Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using DataTable = System.Data.DataTable;
using OK2SHIP_SMT;
using OfficeOpenXml;

namespace VHX
{
    public partial class Impedance_main : Form
    {
        Impedance_data proc_data = new Impedance_data();
        myVar exp_proc = new myVar();
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SEI_Lib myCode = new SEI_Lib();
        Export_Image_Class img_proc = new Export_Image_Class();
        public EPPlus_Lib TDMK_Code2 = new EPPlus_Lib();
        Funtion_setup_new F_setup_new = new Funtion_setup_new();
        ExportEPPlus F_exportEPPlus = new ExportEPPlus();

        public string strcon = "";
        public SqlConnection sqlcon = null;
        public Dictionary<string, SortedDictionary<int, myVar.Impedance_data>> Impedance_result = new Dictionary<string, SortedDictionary<int, myVar.Impedance_data>>();

        public int no_region = 1;
        bool check_load_DB = false;
        string curr_itemcode = "";
        string curr_LotNo = "";
        DataTable tbl_format = new DataTable();
        string location_export_graph = "";
        int id_logfile = 0;
        string selected_path = @"C:\Users\lyy\Desktop\Impedance";
        DataTable Graph_dt = new DataTable();
        DataTable Image_dt = new DataTable();

        public bool admin_mode = false;
        bool update_mode = false;
        bool update_mode_trw = false;
        int region_imp_update = 0;
        string zone_update = "";
        public string depart = "QA";
        public List<string> sel_sample_lst = new List<string> { };
        SortedDictionary<string, List<string>> edited_lst = new SortedDictionary<string, List<string>>();


        public Impedance_main()
        {
            InitializeComponent();
        }
        public void Get_file(string in_src, ref SortedDictionary<int, Dictionary<int, string>> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            FileInfo[] temp_lst = tar_d.GetFiles("*.csv");
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    Dictionary<int, string> temp = Get_LogFile_Data(temp_lst[i].FullName);
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });

                    f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                    int f_inx = Convert.ToInt32(f_na);
                    if (!lst_result.ContainsKey(f_inx))
                    {
                        lst_result.Add(f_inx, temp);
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
                        Get_logfile_Multi(inter_lst.FullName, ref lst_result);
                    }
                }
            }
        }
        public void Get_logfile_Multi(string in_src, ref SortedDictionary<int, Dictionary<int, string>> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            try
            {

                FileInfo[] temp_lst = tar_d.GetFiles("*.csv");
                if (temp_lst.Length > 0)
                {
                    for (int i = 0; i < temp_lst.Length; i++)
                    {
                        Dictionary<int, string> temp = Get_LogFile_Data(temp_lst[i].FullName);
                        string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });

                        f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                        if (myCode.IsNumeric(f_na))
                        {
                            int f_inx = Convert.ToInt32(f_na);
                            if (!lst_result.ContainsKey(f_inx))
                            {
                                lst_result.Add(f_inx, temp);
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
                            Get_logfile_Multi(inter_lst.FullName, ref lst_result);
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("Could not find a part of the path " + in_src, "Warning");
            }
        }

        public void Get_file_xlsm(string in_src, ref string f_path)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            try
            {
                FileInfo[] temp_lst = tar_d.GetFiles("*.xlsm");
                if (temp_lst.Length > 0)
                {
                    f_path = temp_lst[0].FullName;

                }
                else
                {
                    DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                    if (inter_type_lst.Length > 0)
                    {
                        foreach (var inter_lst in inter_type_lst)
                        {
                            Get_file_xlsm(inter_lst.FullName, ref f_path);
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("Error finding file", "Waring");
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
        public void Get_Impedance_logfile_Multi2_(string in_src, ref Dictionary<string, SortedDictionary<int, myVar.Impedance_data>> dic_lst_result)
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
                        Get_Impedance_logfile_Multi2_(inter_lst.FullName, ref dic_lst_result);
                    }
                }

            }
        }

        public void Get_Impedance_logfile_Multi2(string in_src, ref Dictionary<string, SortedDictionary<int, myVar.Impedance_data>> dic_lst_result)
        {
            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            myExcel.Workbook wrkbk = xlsApp.Workbooks.Open(in_src); //TDMK_Code.open_excel_file(in_src_file, "", "");
            myExcel.Worksheet wrksht = wrkbk.Sheets["Result"];
            int grp = 1;
            for (int i = 1; i < 5; i++)
            {
                SortedDictionary<int, myVar.Impedance_data> result_lst = new SortedDictionary<int, myVar.Impedance_data>();
                myExcel.Range id_rgn = wrksht.Range["A1"];
                myExcel.Range val_rgn = id_rgn.Offset[i, 0];
                if (myCode.checkDBNull(val_rgn.Value) != "")
                {
                    myExcel.Range sel_rgn = id_rgn;
                    int col_inx = 0;
                    string id_val = myCode.checkDBNull(sel_rgn.Value);

                    while (id_val != "")
                    {
                        myVar.Impedance_data imp_graph = new myVar.Impedance_data();
                        imp_graph.data_val = myCode.checkDBNull(val_rgn.Offset[0, col_inx].Value);
                        if (Convert.ToInt32(id_val) <= 3 && grp == 1)
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
                    dic_lst_result.Add(i.ToString(), result_lst);
                    grp++;
                }
                else
                {
                    break;
                }
            }
            wrkbk.Close();
        }
        public SortedDictionary<int, myVar.Impedance_data> Get_Impedance_LogFile_Data(string in_src_file)
        {
            SortedDictionary<int, myVar.Impedance_data> result_lst = new SortedDictionary<int, myVar.Impedance_data>();
            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            myExcel.Workbook wrkbk = xlsApp.Workbooks.Open(in_src_file); //TDMK_Code.open_excel_file(in_src_file, "", "");
            myExcel.Worksheet wrksht = wrkbk.Sheets["Result"];

            for (int i = 1; i < 5; i++)
            {
                myExcel.Range id_rgn = wrksht.Range["A1"];
                myExcel.Range val_rgn = id_rgn.Offset[i, 0];
                if (myCode.checkDBNull(val_rgn.Value) != null)
                {
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
                }
                else
                {
                    break;
                }
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
                    f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                    if (myCode.IsNumeric(f_na))
                    {
                        int f_inx = Convert.ToInt32(f_na);
                        if (!lst_result.ContainsKey(f_inx))
                        {
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
                        Get_Impedance_VHX_Multi(inter_lst.FullName, ref lst_result);
                    }
                }
            }

        }

        private void txtLogfile_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Logfile_Data_Impedance();

        }
        public void Logfile_Data_Tracewidth()
        {
            FolderBrowserDialog f_open = new FolderBrowserDialog();
            f_open.SelectedPath = selected_path;
            if (f_open.ShowDialog() == DialogResult.OK)
            {
                txtLogfile_Trw.Text = f_open.SelectedPath;

                //Tracewidth_Data_Process_2(f_open.SelectedPath);
                DirectoryInfo tar_parent = new DirectoryInfo(txtLogfile_Trw.Text);
                DirectoryInfo[] arr_dic_child = tar_parent.GetDirectories();

                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                DataTable VHX_log_tbl = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_VAL", filter_str).Clone();
                DataTable VHX_img_tbl = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_IMAGE", filter_str).Clone();
                int count_ = 1;
                if (arr_dic_child.Length > 0)
                {

                    foreach (DirectoryInfo tar_d in arr_dic_child)
                    {
                        Tracewidth_Data_Process_2(tar_d.FullName, ref VHX_log_tbl);
                        Image_Process(tar_d.FullName, ref VHX_img_tbl);
                        count_++;
                    }
                }
                DGV_VHX_Data.DataSource = VHX_log_tbl;
                DGV_Image_Tracewidth.DataSource = VHX_img_tbl;

                ((DataGridViewImageColumn)DGV_Image_Tracewidth.Columns["Image_Tracewidth"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)DGV_Image_Tracewidth.Columns["Image_Tracewidth"]).Width = 100;
                foreach (DataGridViewRow dr in DGV_Image_Tracewidth.Rows)
                {
                    dr.Height = 70;
                }

                selected_path = f_open.SelectedPath;
            }
        }
        //public void Tracewidth_Data_Process(string src_path)
        //{
        //    if (src_path != "")
        //    {
        //      string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
        //        //Dictionary<string, SortedDictionary<int, myVar.Impedance_data>> Impedance_result = new Dictionary<string, SortedDictionary<int, myVar.Impedance_data>>();
        //        //Get_Impedance_logfile_Multi2(src_path, ref Impedance_result);
        //        SortedDictionary<int, Dictionary<int, string>> logfile_result_VHX = new SortedDictionary<int, Dictionary<int, string>>();
        //        Get_logfile_Multi(src_path, ref logfile_result_VHX);              
        //        DataTable VHX_log_tbl = TDMK_Code.Datatable_Filter(sqlcon, "DATA_LOG_FILE", filter_str).Clone();

        //        int log_inx = 0;
        //        foreach (var log in logfile_result_VHX)
        //        {
        //            foreach (var log_val in log.Value)
        //            {
        //                VHX_log_tbl.Rows.Add((log_inx + 1), txtItemCode.Text, txtLotNo.Text, log_val.Key.ToString(), "Sample" + log.Key.ToString(), log_val.Value, DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), txtOperator.Text, "IMPEDANCE");
        //                log_inx++;
        //            }
        //        }             
        //        DGV_VHX_Data.DataSource = VHX_log_tbl;
        //        if (DGV_Impedance.Rows.Count > 0 && DGV_VHX_Data.Rows.Count > 0)
        //        {
        //            Image_Process(src_path, Impedance_result);
        //        }
        //    }
        //}
        public void Tracewidth_Data_Process_2(string src_path, ref DataTable VHX_log_tbl)
        {
            if (src_path != "")
            {
                //try
                //{


                SortedDictionary<int, Dictionary<int, string>> logfile_result_VHX = new SortedDictionary<int, Dictionary<int, string>>();
                Get_logfile_Multi(src_path, ref logfile_result_VHX);
                //DataTable VHX_log_tbl = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_VAL", filter_str).Clone();

                int log_inx = VHX_log_tbl.Rows.Count + 1;

                foreach (var log in logfile_result_VHX)
                {

                    if (log.Value.Count == 6)
                    {
                        int i = 1;
                        int no = 1;

                        foreach (var log_val in log.Value)
                        {
                            if (i == 2 || i == 5)
                            {
                                VHX_log_tbl.Rows.Add(log_inx, txtItemCode.Text, txtLotNo.Text, no.ToString(), "Sample" + log.Key.ToString(), log_val.Value, DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), txtOperator.Text, Path.GetFileName(src_path), src_path);
                                log_inx++;
                                no++;

                            }
                            i++;

                        }
                    }
                    else
                    {
                        int i = 0;
                        foreach (var log_val in log.Value)
                        {

                            if (i < 2)
                            {
                                VHX_log_tbl.Rows.Add(log_inx, txtItemCode.Text, txtLotNo.Text, log_val.Key.ToString(), "Sample" + log.Key.ToString(), log_val.Value, DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), txtOperator.Text, Path.GetFileName(src_path), src_path);
                                log_inx++;
                            }
                            i++;

                        }
                    }
                }

                //reset_DGV_color(DGV_VHX_Data);



                //}
                //catch
                //{
                //    MessageBox.Show("Could not find a part of the path " + src_path, "Warning");
                //}



            }


        }
        public void Tracewidth_Update(string src_path, string datafor)
        {
            if (src_path != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Data_For" }, new string[] { txtItemCode.Text, txtLotNo.Text, datafor });

                SortedDictionary<int, Dictionary<int, string>> logfile_result_VHX = new SortedDictionary<int, Dictionary<int, string>>();
                Get_logfile_Multi(src_path, ref logfile_result_VHX);
                DataTable VHX_log_tbl = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_VAL", filter_str).Clone();


                int log_inx = 0;

                foreach (var log in logfile_result_VHX)
                {

                    if (log.Value.Count == 6)
                    {
                        int i = 1;
                        int no = 1;

                        foreach (var log_val in log.Value)
                        {
                            if (i == 2 || i == 5)
                            {
                                VHX_log_tbl.Rows.Add((log_inx + 1), txtItemCode.Text, txtLotNo.Text, no.ToString(), "Sample" + log.Key.ToString(), log_val.Value, DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), txtOperator.Text, datafor, src_path);
                                log_inx++;
                                no++;

                            }
                            i++;

                        }
                    }
                    else
                    {
                        int i = 0;
                        foreach (var log_val in log.Value)
                        {

                            if (i < 2)
                            {
                                VHX_log_tbl.Rows.Add((log_inx + 1), txtItemCode.Text, txtLotNo.Text, log_val.Key.ToString(), "Sample" + log.Key.ToString(), log_val.Value, DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), txtOperator.Text, datafor, src_path);
                                log_inx++;


                            }
                            i++;

                        }
                    }

                }


                DGV_VHX_Data.DataSource = VHX_log_tbl;
                // Image_Process(src_path);

                //  btn_summary_trw.Enabled = true;
                //  btn_reset_trw.Enabled = true;
                //}
                //else
                //{
                //    MessageBox.Show("ItemCode or LotNo not matched", "Warning");
                //}
            }


        }

        public void Logfile_Data_Impedance()
        {
            //OpenFileDialog f_open = new OpenFileDialog();

            //f_open.Filter = "Excel(*.xlsm)|*.xlsm";
            //f_open.InitialDirectory = Application.StartupPath;
            //if (f_open.ShowDialog() == DialogResult.OK)
            //{
            //    if (f_open.FileName != "")
            //    {
            //        Impedance_Data_Process_0(f_open.FileName);

            //    }
            //}

            FolderBrowserDialog f_open = new FolderBrowserDialog();
            f_open.SelectedPath = selected_path;


            if (f_open.ShowDialog() == DialogResult.OK)
            {

                txtLogfile_Imp.Text = f_open.SelectedPath;
                Impedance_Data_Process_0(f_open.SelectedPath);
                //string f_name = fpath.Split('\\')[fpath.Split('\\').Length - 1];


                //string[] lot_no = txtLotNo.Text.Split('-');

                //if (f_name.Split('-')[0].Trim() == txtItemCode.Text && lot_no.Length == f_name.Split('-')[2].Split('_').Length)
                //{
                //    if (lot_no.Length == 2 && Convert.ToDecimal(f_name.Split('-')[1].Trim()) == Convert.ToDecimal(lot_no[0].Trim()) && Convert.ToDecimal(f_name.Split('-')[2].Split('_')[0].Trim()) == Convert.ToDecimal(lot_no[1].Trim()))
                //    {
                //        Impedance_Data_Process_0(f_open.SelectedPath);
                //    }
                //    else if (lot_no.Length == 1 && Convert.ToDecimal(f_name.Split('-')[1].Trim()) == Convert.ToDecimal(lot_no[0].Trim()))
                //    {

                //        Impedance_Data_Process_0(f_open.SelectedPath);
                //    }
                //    else
                //    {
                //        MessageBox.Show("ItemCode or LotNo not matched", "Warning");
                //    }
                //}
                //else
                //{
                //    MessageBox.Show("ItemCode or LotNo not matched", "Warning");
                //}

                selected_path = f_open.SelectedPath;

            }

        }
        public string fill_name(string text)
        {
            if (text.Split('-').Length == 1)
            {
                fill_name_2(text, 5);
            }
            else if (text.Split('-').Length == 2)
            {
                fill_name_2(text.Split('-')[0], 5);
                fill_name_2(text.Split('-')[1], 2);
            }

            return text;
        }
        public string fill_name_2(string text, int max_lenght)
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
        public void Impedance_Data_Process(string src_path)
        {
            if (src_path != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });

                /* Dictionary<string, SortedDictionary<int, myVar.Impedance_data>>*/
                Impedance_result = new Dictionary<string, SortedDictionary<int, myVar.Impedance_data>>();
                Get_Impedance_logfile_Multi2(src_path, ref Impedance_result);
                string sel_graph = Impedance_result.Keys.ToArray()[0];
                SortedDictionary<int, myVar.Impedance_data> graph_data = Impedance_result[sel_graph];

                DataTable Imp_data_dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_VAL", filter_str).Clone();

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
                    //cb_tracewidth.Items.Add("Impedance-" + item.Key);
                }
                DGV_Impedance.DataSource = Imp_data_dt;


                //if (DGV_Impedance.Rows.Count > 0 && DGV_VHX_Data.Rows.Count > 0)
                //{
                //    Image_Process(txttracewidth.Text, Impedance_result);
                //}

            }
        }

        public void Impedance_Data_Process_0(string src_path)
        {
            if (src_path != "" && cb_Impedance.SelectedIndex != -1)
            {
                //location_export_graph += no_region;
                string filename = Path.GetFileName(src_path);
                string zone = cb_Impedance.SelectedItem.ToString();
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Remark" }, new string[] { txtItemCode.Text, txtLotNo.Text, filename });

                //  lbl_restart: DataTable _Imp_data_dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_VAL_LOGFILE", filter_str);
                DataTable Imp_data_dt = new DataTable();
                //DataTable Graph_dt = new DataTable();
                DataTable cur_graph_dt = new DataTable();
                int cur_coupon = 0;
                no_region = 1;

                Imp_data_dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_VAL", filter_str).Clone();
                cur_graph_dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", filter_str).Clone();
                Impedance_result = new Dictionary<string, SortedDictionary<int, myVar.Impedance_data>>();
                string file_path = "";
                Get_file_xlsm(src_path, ref file_path);
                if (file_path != "")
                {
                    //try
                    //{
                    Get_Impedance_logfile_Multi2(file_path, ref Impedance_result);
                    string sel_graph = Impedance_result.Keys.ToArray()[0];
                    SortedDictionary<int, myVar.Impedance_data> graph_data = Impedance_result["1"];
                    int r_inx = 0;
                    foreach (var item in Impedance_result)
                    {
                        var cur_val = item.Value;
                        int pcs_no = 0;
                        foreach (var t in cur_val)
                        {
                            Imp_data_dt.Rows.Add(r_inx + 1, txtItemCode.Text, fill_name(txtLotNo.Text), t.Key, no_region + cur_coupon, t.Value.data_val, zone, src_path);
                            pcs_no++;
                            r_inx++;
                        }
                        no_region++;
                    }
                    DGV_Impedance.DataSource = Imp_data_dt;
                    // DGV_Graph.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                    r_inx = DGV_Graph.Rows.Count + 1;
                    //if (DGV_Graph.Rows.Count > 0)
                    //    Graph_dt = (DataTable)DGV_Graph.DataSource;

                    foreach (var t in graph_data)
                    {
                        if (t.Value.grap_data != null)
                        {
                            //Graph_dt.Rows.Add(r_inx, txtItemCode.Text, txtLotNo.Text, t.Key.ToString(), 1, t.Value.grap_data, txtOperator.Text, zone, src_path);
                            cur_graph_dt.Rows.Add(r_inx, txtItemCode.Text, txtLotNo.Text, t.Key.ToString(), 1, t.Value.grap_data, txtOperator.Text, zone, src_path);
                        }
                        else
                        {
                            break;
                        }
                        r_inx++;

                    }
                    //sort_data_impedance(DGV_Graph, "IMPEDANCE_GRAPH", Graph_dt);
                    DGV_Graph.DataSource = cur_graph_dt;

                    ((DataGridViewImageColumn)DGV_Graph.Columns["Image_Graph"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)DGV_Graph.Columns["Image_Graph"]).Width = 100;
                    foreach (DataGridViewRow dr in DGV_Graph.Rows)
                    {
                        dr.Height = 70;
                    }

                    btnSummary.Enabled = false;
                    //}
                    //catch
                    //{
                    //    MessageBox.Show("Unable to read data", "Warning");
                    //}

                    // Get_Impedance_logfile_Multi2(src_path, ref Impedance_result);

                    // btn_reset.Enabled = true;
                }
                else
                {
                    MessageBox.Show("File macro not found", "Warning");
                }


            }
            else
            {
                MessageBox.Show("Please, select Coupon/Patern", "Warning");
            }
        }


        public void Impedance_Data_Process_1(string src_path)
        {
            if (src_path != "")
            {
                location_export_graph += no_region;
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });

                Impedance_result = new Dictionary<string, SortedDictionary<int, myVar.Impedance_data>>();
                //string file_path = "";
                //Get_file_xlsm(src_path, ref file_path);
                //Get_Impedance_logfile_Multi2(file_path, ref Impedance_result);
                string sel_graph = Impedance_result.Keys.ToArray()[0];
                SortedDictionary<int, myVar.Impedance_data> graph_data = Impedance_result[sel_graph];

                DataTable Imp_data_dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_VAL", filter_str).Clone();
                if (DGV_Impedance.DataSource != null)
                {
                    Imp_data_dt = (DataTable)DGV_Impedance.DataSource;
                }
                int r_inx = 0;
                foreach (var item in Impedance_result)
                {
                    var cur_val = item.Value;
                    int pcs_no = 0;
                    foreach (var t in cur_val)
                    {
                        Imp_data_dt.Rows.Add(r_inx + 1, txtItemCode.Text, txtLotNo.Text, t.Key, no_region, t.Value.data_val, "Patern");
                        pcs_no++;
                        r_inx++;
                    }
                    no_region++;
                }
                DGV_Impedance.DataSource = Imp_data_dt;
                string filter_str_1 = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
                DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_FORMAT", filter_str_1);
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < DGV_Impedance.Rows.Count; i++)
                    {
                        int indx_imp = int.Parse(DGV_Impedance.Rows[i].Cells["Region"].Value.ToString());
                        if (indx_imp <= dt.Rows[0]["Impedance_USL"].ToString().Split(';').Length)
                        {
                            string imp_max = dt.Rows[0]["Impedance_USL"].ToString().Split(';')[indx_imp - 1];
                            string imp_min = dt.Rows[0]["Impedance_LSL"].ToString().Split(';')[indx_imp - 1];

                            Double max_target = Double.Parse(imp_max);
                            Double min_target = Double.Parse(imp_min);
                            Double data = Double.Parse(DGV_Impedance.Rows[i].Cells["Data"].Value.ToString());
                            if (data > max_target || data < min_target)
                            {
                                for (int j = 0; j < DGV_Impedance.Columns.Count; j++)
                                {
                                    DGV_Impedance.Rows[i].Cells[j].Style.BackColor = Color.Red;
                                }
                            }
                        }

                    }

                }
                MessageBox.Show(new Form { TopMost = true }, "Load data logfile completed", "Warning");
            }
        }

        public void Image_Process(string src_path, ref DataTable VHX_img_tbl)
        {
            if (src_path != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                SortedDictionary<int, byte[]> Image_result_VHX = new SortedDictionary<int, byte[]>();
                Get_Impedance_VHX_Multi(src_path, ref Image_result_VHX);

                int img_inx = VHX_img_tbl.Rows.Count + 1;
                //int qty = Math.Min(Image_result_VHX.Count, 3);
                int qty = 0;
                foreach (var log in Image_result_VHX)
                {
                    VHX_img_tbl.Rows.Add(img_inx, txtItemCode.Text, txtLotNo.Text, "Sample" + log.Key.ToString(), Path.GetFileName(src_path), log.Value, txtOperator.Text, DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), src_path);
                    img_inx++;
                    qty++;
                    if (qty == 3)
                    {
                        break;
                    }
                }


            }
        }

        public void Image_Process_2(string src_path, Dictionary<string, SortedDictionary<int, myVar.Impedance_data>> Impedance_result, string trw_select)
        {
            if (src_path != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });

                SortedDictionary<int, byte[]> Image_result_VHX = new SortedDictionary<int, byte[]>();
                Get_Impedance_VHX_Multi(src_path, ref Image_result_VHX);
                string sel_graph = Impedance_result.Keys.ToArray()[0];
                SortedDictionary<int, myVar.Impedance_data> graph_data = Impedance_result[sel_graph];
                DataTable VHX_img_tbl = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", filter_str).Clone();
                if (DGV_Graph.Rows.Count > 0)
                {
                    VHX_img_tbl = (DataTable)DGV_Graph.DataSource;
                }
                int img_inx = 0;
                int qty = Math.Min(Math.Min(Image_result_VHX.Count, graph_data.Count), 3);
                if (DGV_Graph.DataSource != null)
                {
                    VHX_img_tbl = (DataTable)DGV_Graph.DataSource;
                }

                foreach (var log in Image_result_VHX)
                {
                    VHX_img_tbl.Rows.Add(img_inx + 1, txtItemCode.Text, txtLotNo.Text, "Sample" + log.Key.ToString(), trw_select, log.Value, graph_data[img_inx + 1].grap_data, txtOperator.Text, DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"));
                    img_inx++;
                    if (img_inx >= qty)
                    {
                        break;
                    }
                }

                DGV_Graph.DataSource = VHX_img_tbl;


            }
        }



        public void Impedance_Data_Process_all(string src_path)
        {
            if (src_path != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                SortedDictionary<int, Dictionary<int, string>> logfile_result_VHX = new SortedDictionary<int, Dictionary<int, string>>();
                Get_logfile_Multi(src_path, ref logfile_result_VHX);
                SortedDictionary<int, byte[]> Image_result_VHX = new SortedDictionary<int, byte[]>();
                Get_Impedance_VHX_Multi(src_path, ref Image_result_VHX);
                Dictionary<string, SortedDictionary<int, myVar.Impedance_data>> Impedance_result = new Dictionary<string, SortedDictionary<int, myVar.Impedance_data>>();
                Get_Impedance_logfile_Multi2(src_path, ref Impedance_result);
                string sel_graph = Impedance_result.Keys.ToArray()[0];
                SortedDictionary<int, myVar.Impedance_data> graph_data = Impedance_result[sel_graph];
                DataTable VHX_log_tbl = TDMK_Code.Datatable_Filter(sqlcon, "DATA_LOG_FILE", filter_str).Clone();
                DataTable VHX_img_tbl = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", filter_str).Clone();
                DataTable Imp_data_dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_VAL", filter_str).Clone();
                /***************************** Insert to DATA_LOG_FILE Table of VHX *************************************************************************************/
                int log_inx = 0;
                foreach (var log in logfile_result_VHX)
                {
                    foreach (var log_val in log.Value)
                    {
                        VHX_log_tbl.Rows.Add((log_inx + 1), txtItemCode.Text, txtLotNo.Text, "Sample" + log.Key.ToString(), log_val.Key.ToString(), log_val.Value, DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), "", "IMPEDANCE");
                        log_inx++;
                    }
                }
                /***************************** Finished insert to DATA_LOG_FILE of VHX ****************************************************************************/
                /***************************** Insert to IMPEDANCE_IMAGE Table *************************************************************************************/
                int img_inx = 0;
                int qty = Math.Min(Math.Min(Image_result_VHX.Count, graph_data.Count), 3);
                foreach (var log in Image_result_VHX)
                {
                    VHX_img_tbl.Rows.Add(img_inx + 1, txtItemCode.Text, txtLotNo.Text, "Sample" + log.Key.ToString(), log.Key.ToString(), log.Value, graph_data[img_inx + 1].grap_data, "", DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"));
                    img_inx++;
                    if (img_inx >= qty)
                    {
                        break;
                    }
                }
                /***************************** Finished insert to IMPEDANCE_IMAGE Table *************************************************************************************/
                /***************************** Insert to IMPEDANCE_VAL Table  *************************************************************************************/
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
                /***************************** Finished insert to IMPEDANCE_VAL Table  *************************************************************************************/
                /**************************** Fill datatable to Datagridview ***********************************************************************************************/
                DGV_Impedance.DataSource = Imp_data_dt;
                DGV_VHX_Data.DataSource = VHX_log_tbl;
                DGV_Graph.DataSource = VHX_img_tbl;
                /**************************** Finished fill datatable to Datagridview ***********************************************************************************************/
            }
        }


        private void Impedance_main_Load(object sender, EventArgs e)
        {
            sqlcon = myVar.mysqlcon;
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { "22B0500", "00003" });
            Graph_dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", filter_str).Clone();
            Image_dt = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_IMAGE", filter_str).Clone();
            admin_mode = false;


        }

        private void DGV_Image_Graph_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int col_inx = e.ColumnIndex;
            int r_inx = e.RowIndex;
            DataGridViewCell cur_cell = DGV_Graph.CurrentCell;
            if (DGV_Graph.Columns[col_inx].Name.Contains("Image"))
            {
                //byte[] data = (byte[])cur_cell.Value;
                //using (MemoryStream ms = new MemoryStream(data))
                //{
                //    pic_imp.Image = Image.FromStream(ms);
                //}


                string region = DGV_Graph.Rows[r_inx].Cells["Region"].Value.ToString();
                string pcs = DGV_Graph.Rows[r_inx].Cells["Pcs_No"].Value.ToString();

                View_detail_Image fr1 = new View_detail_Image();
                fr1.data = (byte[])cur_cell.Value;
                fr1.region_ = region;
                fr1.pcs_ = pcs;
                fr1.choose_sheet_ = "Impedance";
                fr1.Show();
            }
        }

        private void tsmLoadData_Click(object sender, EventArgs e)
        {

        }
        public Boolean check_DGV_NG(DataGridView dgv)
        {
            bool chk = true;
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                if (dgv.Rows[i].Cells[0].Style.BackColor == Color.Red)
                {
                    chk = false;
                    break;
                }
            }
            return chk;
        }

        private void tsmSaveData_Click(object sender, EventArgs e)
        {
            //if (txtItemCode.Text != "" && txtLotNo.Text != "" && cb_Impedance.SelectedIndex != -1)
            //{

            //}
            //else
            //{
            //    MessageBox.Show("Please, fill in ItemCode / LotNo / Operator/ Process", "Warning");
            //}



        }

        private void txtLogfile_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void tsmExport_Click(object sender, EventArgs e)
        {

        }
        public void Export_Impedance_data(SqlConnection tar_sqlcon, string ItemCode, string LotNo)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            List<string> tar_table_name = new List<string>() { "IMPEDANCE_VAL", "TRACEWIDTH_VAL", "TRACEWIDTH_IMAGE", "IMPEDANCE_GRAPH" };
            //List<DataGridView> tar_DGV = new List<DataGridView>() { DGV_Image_Graph, DGV_VHX_Data, DGV_Impedance };
            List<DataTable> tar_dt = new List<DataTable>();
            for (int i = 0; i < tar_table_name.Count; i++)
            {
                tar_dt.Add(TDMK_Code.Datatable_Filter(tar_sqlcon, tar_table_name[i], filter_str));
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
                string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                if (format_file != "")
                {
                    myExcel.Workbook curr_wrkbook = exp_proc.create_export_wrk(format_file, "Impedance", txtItemCode.Text, txtLotNo.Text);
                    myExcel.Worksheet curr_wrksheet = curr_wrkbook.Sheets[1];
                    curr_wrksheet.Cells.Interior.Color = myExcel.XlRgbColor.rgbWhite;
                    for (int i = 0; i < tar_table_name.Count; i++)
                    {
                        DataTable sel_dt = tar_dt[i];
                        switch (tar_table_name[i])
                        {
                            case "IMPEDANCE_VAL":
                                Export_Impedance_Val(sel_dt, curr_wrksheet);
                                break;
                            case "TRACEWIDTH_VAL":
                                Export_VHX_data_2(sel_dt, curr_wrksheet);
                                break;
                            case "TRACEWIDTH_IMAGE":
                                Export_Image_VHX(sel_dt, curr_wrksheet);
                                break;
                            case "IMPEDANCE_GRAPH":
                                Export_Graph(sel_dt, curr_wrksheet);
                                break;
                        }
                    }
                    save_report_new(txtItemCode.Text, txtLotNo.Text, txtOperator.Text, curr_wrkbook, "Impedance", myVar.data_loc);
                    MessageBox.Show(new Form { TopMost = true }, "Export to Checksheet completed!", "Warning");
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Format of " + txtItemCode.Text + " not found", "Warning");
                }

            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "No data", "Warning");
            }



        }
        public void Export_data_impedance(DataTable in_tbl, myExcel.Worksheet curr_wrksheet, int index)
        {
            if (in_tbl.Rows.Count > 0)
            {
                string region = in_tbl.Rows[0]["Region"].ToString().Trim().Replace(" ", "");
                // myExcel.Range rgn = curr_wrksheet.Range["E30"];

                AutoCompleteStringCollection lst_rgn = Get_impedance_infor_to_input_value(curr_wrksheet);
                int cou_rgn = 1;
                myExcel.Range data_rgn;
                foreach (string rgn in lst_rgn)
                {
                    if (region == cou_rgn.ToString())
                    {
                        data_rgn = curr_wrksheet.Range[rgn];
                        int r_inx = 0;
                        foreach (DataRow dr in in_tbl.Rows)
                        {
                            data_rgn.Offset[r_inx, 0].Value = dr["Data"].ToString();
                            r_inx++;
                        }
                        break;
                    }
                    cou_rgn++;

                }

            }

        }
        public void Export_data_impedance_2(DataTable in_tbl, myExcel.Worksheet curr_wrksheet, int index)
        {
            if (in_tbl.Rows.Count > 0)
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
                DataTable dt_format = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", filter_str);

                string region = in_tbl.Rows[0]["Region"].ToString().Trim().Replace(" ", "");

                //AutoCompleteStringCollection lst_rgn = Get_impedance_infor_to_input_value(curr_wrksheet);
                int cou_rgn = 1;
                myExcel.Range data_rgn;
                for (int t = 0; t < dt_format.Rows.Count; t++)
                {
                    string rgn = dt_format.Rows[t]["Range"].ToString().Split('-')[0];
                    string count_sample = dt_format.Rows[t]["Range"].ToString().Split('-')[1];
                    if (region == cou_rgn.ToString())
                    {
                        data_rgn = curr_wrksheet.Range[rgn];
                        int r_inx = 0;
                        int min_sample = new int[] { int.Parse(count_sample), in_tbl.Rows.Count }.Min();
                        foreach (DataRow dr in in_tbl.Rows)
                        {
                            if (r_inx < min_sample)
                            {
                                data_rgn.Offset[r_inx, 0].Value = dr["Data"].ToString();
                                r_inx++;
                            }
                        }
                        break;
                    }
                    cou_rgn++;

                }


            }

        }
        public void Export_data_tracewidth(DataTable in_tbl, myExcel.Worksheet curr_wrksheet, int index)
        {
            if (in_tbl.Rows.Count > 0)
            {
                string region = in_tbl.Rows[0]["Region"].ToString().Trim().Replace(" ", "");
                // myExcel.Range rgn = curr_wrksheet.Range["E30"];
                AutoCompleteStringCollection lst_rgn = Get_tracewidth_infor_to_input_value(curr_wrksheet);
                int cou_rgn = 1;

                foreach (string item_rgn in lst_rgn)
                {
                    int r = int.Parse(item_rgn.Split('+')[0]);
                    int c = int.Parse(item_rgn.Split('+')[1]);

                    List<DataTable> src_sample = new List<DataTable>();
                    exp_proc.Get_ListTable(-1, in_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_sample, "Data");

                    myExcel.Range rgn = curr_wrksheet.Cells[r, c];
                    for (int k = 0; k < src_sample.Count; k++)
                    {
                        myExcel.Range data_rgn = rgn.Offset[2 * k, 0];
                        for (int j = 0; j < src_sample[k].Rows.Count; j++)
                        {
                            data_rgn.Offset[j, 0].Value = src_sample[k].Rows[j]["Data"].ToString();
                        }
                    }

                }

            }

        }
        public void Export_image_graph(DataTable in_tbl, myExcel.Worksheet curr_wrksheet, int index)
        {
            if (in_tbl.Rows.Count > 0)
            {
                string region = in_tbl.Rows[0]["Region"].ToString().Trim().Replace(" ", "");
                // myExcel.Range rgn = curr_wrksheet.Range["E30"];
                AutoCompleteStringCollection lst_rgn = Get_image_infor_to_input_value(curr_wrksheet);
                int cou_rgn = 1;
                int idx_imp = 1;

                foreach (string item_rgn in lst_rgn)
                {
                    int r = int.Parse(item_rgn.Split('+')[0]);
                    int c = int.Parse(item_rgn.Split('+')[1]);

                    DataTable src_dt = new DataTable();
                    foreach (DataRow dr in in_tbl.Rows)
                    {
                        if (dr["Remark"].ToString() == "Impedance-" + idx_imp.ToString())
                        {
                            for (int i = 0; i < in_tbl.Columns.Count; i++)
                            {
                                DataRow row = src_dt.NewRow();
                                row[i] = dr[i];
                                src_dt.Rows.Add(row);
                            }
                        }
                    }

                    // List<DataTable> src_sample = new List<DataTable>();
                    // exp_proc.Get_ListTable(-1, in_tbl, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_sample, "Data");

                    myExcel.Range rgn_img_data = curr_wrksheet.Cells[r, c];
                    myExcel.Range rgn_graph = rgn_img_data.Offset[rgn_img_data.Rows.Count + 2, 0];

                    exp_proc.Export_DatatableImage_Excel_2(src_dt, "Image_Data", curr_wrksheet, rgn_img_data, false);
                    exp_proc.Export_DatatableImage_Excel_2(src_dt, "Image_Graph", curr_wrksheet, rgn_graph, false);
                    idx_imp++;

                }

            }

        }

        public void Export_Impedance_Val(DataTable src_dt, myExcel.Worksheet curr_wrksheet)
        {
            List<DataTable> src_region_tbl = new List<DataTable>();
            exp_proc.Get_ListTable(-1, src_dt, new string[] { "ItemCode", "LotNo", "Region" }, ref src_region_tbl, "Data");
            int tbl_inx = 0;
            foreach (var dt in src_region_tbl)
            {
                //proc_data.Export_data_impedance2(dt, curr_wrksheet, tbl_inx);
                Export_data_impedance_2(dt, curr_wrksheet, tbl_inx);
                tbl_inx++;
            }
        }
        public AutoCompleteStringCollection Get_impedance_infor_to_input_value_2(myExcel.Worksheet ws)
        {

            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 24; i < 400; i++)
            {
                if (ws.Cells[i, 5].Value != null)
                {
                    if (ws.Cells[i, 5].Value.ToString().Replace(" ", "").ToUpper() == "Actual Impedance".Replace(" ", "").ToUpper())
                    {
                        if (ws.Cells[i + 1, 3].Value != null && ws.Cells[i + 1, 4].Value != null)
                        {
                            string max_imp = ws.Cells[i + 1, 3].Value.ToString();
                            string min_imp = ws.Cells[i + 1, 4].Value.ToString();

                            int count = 1;
                            while (ws.Cells[i + count, 2].Value != null)
                            {
                                count++;
                            }


                            list.Add("E" + (i + 1).ToString() + "-" + (count - 1).ToString() + "+" + max_imp + "+" + min_imp);
                            i = i + count;
                        }
                        //else
                        //{
                        //    int count = 1;
                        //    while (ws.Cells[i + count, 2].Value != null)
                        //    {
                        //        count++;
                        //    }

                        //    list.Add("E" + (i + 1).ToString() + "-" + (count - 1).ToString() + "+ 0 + 0");
                        //    i = i + count;
                        //}

                    }

                }
            }
            return list;
        }
        //public AutoCompleteStringCollection Get_impedance_infor_to_input_value(myExcel.Worksheet ws)
        //{

        //    AutoCompleteStringCollection list = new AutoCompleteStringCollection();
        //    for (int i = 24; i < 400; i++)
        //    {
        //        if (ws.Cells[i, 5].Value != null)
        //        {
        //            if (ws.Cells[i, 5].Value.ToString().Replace(" ", "").ToUpper() == "Actual Trace width".Replace(" ", "").ToUpper())
        //            {
        //                if (ws.Cells[i + 1, 3].Value != null && ws.Cells[i + 1, 4].Value != null)
        //                {
        //                    string max_imp = ws.Cells[i + 1, 3].Value.ToString();
        //                    string min_imp = ws.Cells[i + 1, 4].Value.ToString();
        //                    list.Add("E" + (i + 1).ToString() + "+" + max_imp + "+" + min_imp);
        //                }

        //            }
        //        }
        //    }
        //    return list;
        //}
        public AutoCompleteStringCollection Get_impedance_infor_to_input_value(myExcel.Worksheet ws)
        {

            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 24; i < 300; i++)
            {
                if (ws.Cells[i, 5].Value != null)
                {
                    if (ws.Cells[i, 5].Value.ToString().Replace(" ", "").ToUpper() == "Actual Impedance".Replace(" ", "").ToUpper())
                    {
                        list.Add("E" + (i + 1).ToString());
                    }
                }
            }
            return list;
        }


        public AutoCompleteStringCollection Get_tracewidth_infor_to_input_value(myExcel.Worksheet ws)
        {
            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 24; i < 300; i++)
            {
                for (int j = 5; j < 15; j++)
                {
                    if (ws.Cells[i, j].Value != null)
                    {
                        if (ws.Cells[i, j].Value.ToString().Replace(" ", "").ToUpper() == "Actual Trace width".Replace(" ", "").ToUpper())
                        {

                            int count = 1;
                            while (ws.Cells[i + count, j - 1].Value != null)
                            {
                                count++;
                            }
                            list.Add((i + 1).ToString() + "+" + j.ToString() + "+" + (count - 1).ToString());
                        }
                    }
                }
            }
            return list;
        }
        public AutoCompleteStringCollection Get_tracewidth_infor_to_input_value_2(myExcel.Worksheet ws)
        {
            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 24; i < 300; i++)
            {
                for (int j = 10; j < 15; j++)
                {

                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Replace(" ", "").ToUpper().Contains("Actual Trace width".Replace(" ", "").ToUpper()))
                    {
                        if (ws.Cells[i + 1, j - 1].Value != null && ws.Cells[i + 1, j - 2].Value != null)
                        {
                            int count = 1;
                            while (ws.Cells[i + count, j - 1].Value != null)
                            {
                                count++;
                            }

                            string max_tracewidth = ws.Cells[i + 1, j - 2].Value.ToString();
                            string min_tracewidth = ws.Cells[i + 1, j - 1].Value.ToString();
                            //  list.Add((i + 1).ToString() + ":" + j.ToString() + ":" + (count - 1).ToString() + "+" + max_tracewidth + "+" + min_tracewidth + "+" + myCode.checkDBNull(ws.Cells[i, j].Value).Split('(')[1].Split(')')[0]);
                            string region = "";
                            if (myCode.checkDBNull(ws.Cells[i, j].Value).Contains("(") && myCode.checkDBNull(ws.Cells[i, j].Value).Contains(")"))
                                region = myCode.checkDBNull(ws.Cells[i, j].Value).Split('(')[1].Split(')')[0];

                            list.Add((i + 1).ToString() + ":" + j.ToString() + ":" + (count - 1).ToString() + "+" + max_tracewidth + "+" + min_tracewidth + "+" + region);

                            i = i + count;
                        }
                        break;

                    }

                }
            }
            return list;
        }
        public AutoCompleteStringCollection Get_image_infor_to_input_value(myExcel.Worksheet ws)
        {
            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 24; i < 300; i++)
            {
                for (int j = 8; j < 18; j++)
                {
                    if (ws.Cells[i, j].Value != null)
                    {
                        if (ws.Cells[i, j].Value.ToString().Replace(" ", "").ToUpper() == "Sample 1".Replace(" ", "").ToUpper())
                        {
                            list.Add((i + 1).ToString() + "+" + j.ToString());
                        }
                    }
                }
            }
            return list;
        }
        public void Export_VHX_data(DataTable src_dt, myExcel.Worksheet curr_wrksheet)
        {

            AutoCompleteStringCollection lst_rgn = Get_tracewidth_infor_to_input_value(curr_wrksheet);

            int idx_imp = 1;

            foreach (string item_rgn in lst_rgn)
            {
                int r = int.Parse(item_rgn.Split('+')[0]);
                int c = int.Parse(item_rgn.Split('+')[1]);
                int count_sample = int.Parse(item_rgn.Split('+')[2]);

                DataTable s_dt = src_dt.Clone();
                foreach (DataRow dr in src_dt.Rows)
                {
                    if (dr["Data_For"].ToString() == "IMPEDANCE-" + idx_imp.ToString())
                    {
                        DataRow row = s_dt.NewRow();
                        for (int i = 0; i < src_dt.Columns.Count; i++)
                        {
                            row[i] = dr[i];
                        }
                        s_dt.Rows.Add(row);
                    }
                }
                List<DataTable> src_sample = new List<DataTable>();
                exp_proc.Get_ListTable(-1, s_dt, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_sample, "Data");
                count_sample = Math.Min(src_sample.Count, count_sample);
                myExcel.Range rgn = curr_wrksheet.Cells[r, c];
                for (int k = 0; k < count_sample; k++)
                {
                    myExcel.Range data_rgn = rgn.Offset[2 * k, 0];
                    for (int j = 0; j < src_sample[k].Rows.Count; j++)
                    {
                        data_rgn.Offset[j, 0].Value = src_sample[k].Rows[j]["Data"].ToString();
                    }
                }
                idx_imp++;

            }


        }

        public void Export_VHX_data_2(DataTable src_dt, myExcel.Worksheet curr_wrksheet)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
            DataTable dt_format = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", filter_str);
            //SortedDictionary<int, string> dic_trw = new SortedDictionary<int, string> { };
             

            int idx_imp = 1;
            for (int t = 0; t < dt_format.Rows.Count; t++)
            {
                string region = dt_format.Rows[t]["Region"].ToString();
                string item_rgn = dt_format.Rows[t]["Range"].ToString();
                int r = int.Parse(item_rgn.Split(':')[0]);
                int c = int.Parse(item_rgn.Split(':')[1]);
                int count_sample = int.Parse(item_rgn.Split(':')[2]);

                DataTable s_dt = src_dt.Clone();
                foreach (DataRow dr in src_dt.Rows)
                {
                    if (dr["Data_For"].ToString().Contains(region))
                    {
                        DataRow row = s_dt.NewRow();
                        for (int i = 0; i < src_dt.Columns.Count; i++)
                        {
                            row[i] = dr[i];
                        }
                        s_dt.Rows.Add(row);
                    }
                }
                List<DataTable> src_sample = new List<DataTable>();
                exp_proc.Get_ListTable(-1, s_dt, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_sample, "Data");
                int min_pcs = Math.Min(s_dt.Rows.Count, count_sample);

                myExcel.Range rgn = curr_wrksheet.Cells[r, c];
                for (int k = 0; 2 * k < min_pcs; k++)
                {
                    myExcel.Range data_rgn = rgn.Offset[2 * k, 0];
                    for (int j = 0; j < 2; j++)
                    {
                        data_rgn.Offset[j, 0].Value = src_sample[k].Rows[j]["Data"].ToString();
                    }
                }
                idx_imp++;

            }
        }
        public void Export_Image_VHX(DataTable src_dt, myExcel.Worksheet curr_wrksheet)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
            DataTable dt_format = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", filter_str);

            // AutoCompleteStringCollection lst_rgn = Get_image_infor_to_input_value(curr_wrksheet);

            int idx_imp = 1; 
            for (int t = 0; t < dt_format.Rows.Count; t++)
            {
                string item_rgn = dt_format.Rows[t]["Range"].ToString();
                string region = dt_format.Rows[t]["Region"].ToString();
                int r = int.Parse(item_rgn.Split(':')[0]) + 1;
                int c = int.Parse(item_rgn.Split(':')[1]) + 2;

                DataTable s_dt = src_dt.Clone();
                foreach (DataRow dr in src_dt.Rows)
                {
                    if (dr["Data_For"].ToString().Contains(region))
                    {
                        DataRow row = s_dt.NewRow();
                        for (int i = 0; i < src_dt.Columns.Count; i++)
                        {
                            row[i] = dr[i];
                        }
                        s_dt.Rows.Add(row);
                    }
                }
                myExcel.Range rgn_img_data = curr_wrksheet.Cells[r, c]; 
                exp_proc.Export_DatatableImage_Excel_2(s_dt, "Image_Tracewidth", curr_wrksheet, rgn_img_data, false);

                idx_imp++;
            }



        }
        public void Export_Graph(DataTable src_dt, myExcel.Worksheet curr_wrksheet)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
            DataTable dt_format = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", filter_str);
            DataTable dt_format_imp = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", filter_str);
            List<DataTable> src_region_tbl = new List<DataTable>();

            string[] region = { "Coupon", "Patern" };
            int no_trace = 0;
            foreach (string _regn in region)
            {
                DataTable graph_tbl = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Zone" }, new string[] { txtItemCode.Text, txtLotNo.Text, _regn }));
                string[] sel_val = graph_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();
                foreach (string item in sel_val)
                {
                    DataTable s_dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Zone", "Region" }, new string[] { txtItemCode.Text, txtLotNo.Text, _regn, item }));
                    if (no_trace < dt_format.Rows.Count && no_trace < dt_format_imp.Rows.Count)
                    {
                        string item_rgn = dt_format.Rows[no_trace]["Range"].ToString();
                        int r = int.Parse(item_rgn.Split(':')[0]) + 1;
                        int c = int.Parse(item_rgn.Split(':')[1]) + 2;

                        myExcel.Range rgn_img_data = curr_wrksheet.Cells[r, c];
                        myExcel.Range rgn_graph = rgn_img_data.Offset[rgn_img_data.Rows.Count + 2, 0];
                        exp_proc.Export_DatatableImage_Excel_2(s_dt, "Image_Graph", curr_wrksheet, rgn_graph, false);
                    }
                    int count = int.Parse(s_dt.Rows[0]["Region"].ToString());
                    no_trace += count;
                }

            }

            //  DataTable graph_tbl_coupon = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", TDMK_Code.filter_str(new string[] { "ItemCode" , "LotNo", "Zone"}, new string[] { txtItemCode.Text , txtLotNo.Text,  "Coupon" }));
            //  DataTable graph_tbl_patern = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Zone" }, new string[] { txtItemCode.Text, txtLotNo.Text, "Patern" }));
            ////  string[] sel_val = graph_tbl_coupon.AsEnumerable().Select(x => x.Field<string>("Remark")).Distinct().ToArray();
            //  //foreach (string item in sel_val)
            //  //{
            //  //    DataTable s_dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Zone", "Remark" }, new string[] { txtItemCode.Text, txtLotNo.Text, "" }));
            //  //}

            //  int idx_imp = 1;
            //  bool completed_coupon = false;

            //  for(int t = 0; t <  dt_format.Rows.Count; t++)
            //  {

            //      string item_rgn = dt_format.Rows[t]["Range"].ToString();
            //      int r = int.Parse(item_rgn.Split(':')[0]) + 1;
            //      int c = int.Parse(item_rgn.Split(':')[1]) + 2;

            //      myExcel.Range rgn_img_data = curr_wrksheet.Cells[r, c];
            //      myExcel.Range rgn_graph = rgn_img_data.Offset[rgn_img_data.Rows.Count + 2, 0];

            //      if (!completed_coupon)
            //      {
            //          if (graph_tbl_coupon.Rows.Count > 0)
            //          {
            //              string[] sel_val = graph_tbl_coupon.AsEnumerable().Select(x => x.Field<string>("Remark")).Distinct().ToArray();
            //              foreach (string item in sel_val)
            //              {
            //                  DataTable s_dt = graph_tbl_coupon.Clone();
            //                  for (int i = 0; i < graph_tbl_coupon.Rows.Count; i++)
            //                  {
            //                      if (graph_tbl_coupon.Rows[i]["Remark"].ToString() == item)
            //                      {
            //                          DataRow dr = s_dt.NewRow();
            //                          {
            //                              for (int m = 0; m < graph_tbl_coupon.Columns.Count; m++)
            //                              {
            //                                  dr[m] = graph_tbl_coupon.Rows[i][m];
            //                              }
            //                          }
            //                          idx_imp += int.Parse(graph_tbl_coupon.Rows[i]["Region"].ToString());
            //                          s_dt.Rows.Add(dr);
            //                      }
            //                  }
            //                  exp_proc.Export_DatatableImage_Excel_2(s_dt, "Image_Graph", curr_wrksheet, rgn_graph, false);



            //              }
            //              completed_coupon = true;


            //          }
            //      }
            //      else
            //      {
            //          if (graph_tbl_patern.Rows.Count > 0)
            //          {
            //              string[] sel_val = graph_tbl_patern.AsEnumerable().Select(x => x.Field<string>("Remark")).Distinct().ToArray();
            //              foreach (string item in sel_val)
            //              {
            //                  DataTable s_dt = graph_tbl_patern.Clone();
            //                  for (int i = 0; i < graph_tbl_patern.Rows.Count; i++)
            //                  {
            //                      if (graph_tbl_patern.Rows[i]["Remark"].ToString() == item)
            //                      {
            //                          DataRow dr = s_dt.NewRow();
            //                          {
            //                              for (int m = 0; c < graph_tbl_patern.Columns.Count; m++)
            //                              {
            //                                  dr[m] = graph_tbl_patern.Rows[i][m];
            //                              }
            //                          }
            //                          idx_imp += int.Parse(graph_tbl_patern.Rows[i]["Region"].ToString());
            //                      }
            //                  }


            //              }

            //          }



            //      }





            //      DataTable grp_dt = s_dt.Clone();
            //      foreach (DataRow dr in s_dt.Rows)
            //      {
            //          if (dr["Region"].ToString().Contains("Impedance " + idx_imp.ToString()))
            //          {
            //              DataRow row = grp_dt.NewRow();
            //              for (int i = 0; i < s_dt.Columns.Count; i++)
            //              {
            //                  row[i] = dr[i];
            //              }
            //              grp_dt.Rows.Add(row);
            //          }
            //      }



            //      idx_imp++;
            //  }
            //  foreach (string item_rgn in dt_format.Rows[0]["Range"].ToString().Split(';'))
            //  {

            //  }

        }
        private void txtLogfile_TextChanged(object sender, EventArgs e)
        {

        }

        private void txttracewidth_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void txttracewidth_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Logfile_Data_Tracewidth();
        }

        private void txttracewidth_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter)
            //{
            //    Tracewidth_Data_Process(txttracewidth.Text);

            //}
        }

        private void cb_tracewidth_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtItemCode.Text != curr_itemcode || txtLotNo.Text != curr_LotNo)
            {

                DGV_VHX_Data.DataSource = null;
                curr_itemcode = txtItemCode.Text;
                curr_LotNo = txtLotNo.Text;
                DGV_Impedance_Spec.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
                DGV_Tracewidth_spec.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
                // Image_dt = new DataTable();
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                if (!update_mode)
                {
                    Image_dt = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_IMAGE", filter_str).Clone();
                }



            }



            //if ((txtItemCode.Text != "") && (txtLotNo.Text != ""))
            //{
            //    if (cb_tracewidth.SelectedIndex > -1)
            //    {
            //        Logfile_Data_Tracewidth();
            //    }
            //}
            //else
            //{
            //    MessageBox.Show(new Form { TopMost = true }, "Please, input ItemCode / Lotno", "Warning");
            //}


        }
        private void cb_Impedance_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtItemCode.Text != curr_itemcode || txtLotNo.Text != curr_LotNo)
            {
                no_region = 1;
                DGV_Impedance.DataSource = null;
                curr_itemcode = txtItemCode.Text;
                curr_LotNo = txtLotNo.Text;
                DGV_Impedance_Spec.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
                DGV_Tracewidth_spec.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
                // Graph_dt = new DataTable();
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                Graph_dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", filter_str).Clone();
                //if (!update_mode && Graph_dt)
                //{
                //    DGV_Graph.DataSource = null;
                //}

            }
            //string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
            //DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", filter_str);
            //id_logfile = 1;
            //if (dt.Rows.Count > 0)
            //{
            //    if(DGV_Impedance_Spec.Rows.Count  == 0)
            //    {
            //        DGV_Impedance_Spec.DataSource = dt;
            //    }

            //    Logfile_Data_Impedance();
            //}
            //else
            //{
            //    MessageBox.Show(new Form { TopMost = true }, "Please, Click SETUP SPEC", "Warning");
            //}

        }
        public List<String> get_Location_excel(myExcel.Workbook curr_wrkbook)
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

        public void Load_spec_Imp(ref string max_imp, ref string min_imp, ref string max_tracwidth, ref string min_tracewidth)
        {

            string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
            myExcel.Workbook curr_wrkbook = TDMK_Code.open_excel_file(format_file, "", "");
            //myExcel.Workbook curr_wrkbook = exp_proc.create_export_wrk(format_file, "Impedance", txtItemCode.Text, txtLotNo.Text);
            string process_name = "Impedance";
            string _process = process_name.Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
            string mySheet = "";
            foreach (myExcel.Worksheet tg_sht in curr_wrkbook.Worksheets)
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
                myExcel.Worksheet ws = curr_wrkbook.Sheets[mySheet];

                for (int i = 24; i < 400; i++)
                {
                    if (ws.Cells[i, 3].Value != null)
                    {
                        if (ws.Cells[i, 3].Value.ToString().Contains("Spec Maximum"))
                        {
                            max_imp = ws.Cells[i + 1, 3].Value.ToString();
                            min_imp = ws.Cells[i + 1, 4].Value.ToString();
                            //int k = 1;
                            for (int k = 7; k < 17; k++)
                            {
                                if (ws.Cells[i, k].Value != null)
                                {
                                    if (ws.Cells[i, k].Value.ToString().Contains("Spec Maximum"))
                                    {
                                        max_tracwidth = ws.Cells[i + 1, k].Value.ToString();
                                        min_tracewidth = ws.Cells[i + 1, k + 1].Value.ToString();
                                        break;
                                    }
                                }
                            }

                            break;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Sheet Impedance not found", "Warning");
            }
        }

        private void txtItemCode_TextChanged(object sender, EventArgs e)
        {
            //DataTable imp_spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_FORMAT", "ItemCode ='" + txtItemCode.Text + "'");
            //if(imp_spec_dt.Rows.Count>0)
            //{
            //    string trace_rgn_val = myCode.checkDBNull( imp_spec_dt.Rows[0]["Tracewidth_range"]);
            //    string[] trace_arr = trace_rgn_val.Split(';');
            //    for (int i=0; i< trace_arr.Length;i++)
            //    {
            //        cb_tracewidth.Items.Add("Tracewidth-" + (i + 1).ToString());
            //    }
            //}
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

        //private void btn_setting_imp_Click(object sender, EventArgs e)
        //{
        //    reset_DGV_color(DGV_Impedance);
        //    reset_DGV_color(DGV_VHX_Data);
        //    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
        //    DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_IMPEDANCE", filter_str);
        //    if (dt.Rows.Count > 0)
        //    {
        //        txt_max_imp.Text = dt.Rows[0]["MAX_Impedance"].ToString();
        //        txt_min_imp.Text = dt.Rows[0]["MIN_Impedance"].ToString();

        //        txt_max_tracewidth.Text = dt.Rows[0]["MAX_Tracewidth"].ToString();
        //        txt_min_tracewidth.Text = dt.Rows[0]["MIN_Tracewidth"].ToString();
        //    }
        //    else
        //    {
        //        string max_imp = "";
        //        string min_imp = "";
        //        string max_tracewidth = "";
        //        string min_tracewidth = "";
        //        Load_spec_Imp(ref max_imp, ref min_imp, ref max_tracewidth, ref min_tracewidth);
        //        txt_max_imp.Text = max_imp;
        //        txt_min_imp.Text = min_imp;
        //        txt_max_tracewidth.Text = max_tracewidth;
        //        txt_min_tracewidth.Text = min_tracewidth;
        //    }

        //    if (txt_max_imp.Text != "" && txt_min_imp.Text != "")
        //    {
        //        check_OOS(DGV_Impedance, txt_max_imp.Text, txt_min_imp.Text);
        //        if (txt_max_tracewidth.Text != "" && txt_min_tracewidth.Text != "")
        //        {
        //            check_OOS(DGV_VHX_Data, txt_max_tracewidth.Text, txt_min_tracewidth.Text);
        //        }

        //    }
        //    else
        //    {
        //        MessageBox.Show("Setting Spec for Impedance", "Warning");
        //    }


        //}


        public void check_OOS(DataGridView dgv, string max, string min)
        {
            if (dgv.Rows.Count > 0)
            {
                Double max_target = Double.Parse(max);
                Double min_target = Double.Parse(min);
                for (int i = 0; i < dgv.Rows.Count; i++)
                {
                    Double data = Double.Parse(dgv.Rows[i].Cells["Data"].Value.ToString());
                    if (data > max_target || data < min_target)
                    {
                        for (int j = 0; j < dgv.Columns.Count; j++)
                        {
                            dgv.Rows[i].Cells[j].Style.BackColor = Color.Red;
                        }
                    }

                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, dgv.Name + " No data", "Warning");
            }
        }
        public void check_OOS_Impedance(DataGridView dgv_data, DataGridView dgv_spec)
        {
            if (dgv_data.Rows.Count > 0)
            {
                if (dgv_spec.Rows.Count > 0)
                {
                    for (int i = 0; i < dgv_data.Rows.Count; i++)
                    {
                        int region = int.Parse(dgv_data.Rows[i].Cells["Region"].Value.ToString());
                        if (region <= dgv_spec.Rows.Count)
                        {
                            Double max_target = Double.Parse(dgv_spec.Rows[region - 1].Cells["USL"].Value.ToString());
                            Double min_target = Double.Parse(dgv_spec.Rows[region - 1].Cells["LSL"].Value.ToString());
                            Double data = Double.Parse(dgv_data.Rows[i].Cells["Data"].Value.ToString());
                            if (data > max_target || data < min_target)
                            {
                                for (int j = 0; j < dgv_data.Columns.Count; j++)
                                {
                                    dgv_data.Rows[i].Cells[j].Style.BackColor = Color.Red;
                                }
                            }
                            else
                            {
                                for (int j = 0; j < dgv_data.Columns.Count; j++)
                                {
                                    dgv_data.Rows[i].Cells[j].Style.BackColor = Color.White;
                                }
                            }
                        }

                    }

                }
                //else
                //{
                //    MessageBox.Show(new Form { TopMost = true }, "Format does not contain Tracewidth", "Warning");
                //}

            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, dgv_data.Name + " No data", "Warning");
            }
        }
        public void check_OOS_Trw(DataGridView dgv_data, DataGridView dgv_spec)
        {
            if (dgv_data.Rows.Count > 0)
            {
                if (dgv_spec.Rows.Count > 0)
                {
                    List<string> lst_region = ((DataTable)dgv_spec.DataSource).AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToList();
                    for (int i = 0; i < dgv_data.Rows.Count; i++)
                    {
                        string reg = ((DataTable)dgv_data.DataSource).Rows[i]["Data_For"].ToString();
                        int region = 5;
                        int k = 0;
                        foreach (string item in lst_region)
                        {
                            if (reg.Contains(item))
                            {
                                region = k;
                                break;
                            }
                            k++;

                        }
                        if(region < dgv_spec.Rows.Count)
                        {
                            Double max_target = Double.Parse(dgv_spec.Rows[region].Cells["USL"].Value.ToString());
                            Double min_target = Double.Parse(dgv_spec.Rows[region].Cells["LSL"].Value.ToString());
                            Double data = Double.Parse(dgv_data.Rows[i].Cells["Data"].Value.ToString());
                            if (data > max_target || data < min_target)
                            {
                                for (int j = 0; j < dgv_data.Columns.Count; j++)
                                {
                                    dgv_data.Rows[i].Cells[j].Style.BackColor = Color.Red;
                                }
                            }
                            else
                            {
                                for (int j = 0; j < dgv_data.Columns.Count; j++)
                                {
                                    dgv_data.Rows[i].Cells[j].Style.BackColor = Color.White;
                                }
                            }
                        }    

                       // int region = int.Parse(dgv_data.Rows[i].Cells["Data_For"].Value.ToString().Replace("Tracewidth_", ""));
                        //string x = dgv_spec.Rows[region - 1].Cells["USL"].ToString();
                        

                    }
                }
                //else
                //{
                //    MessageBox.Show(new Form { TopMost = true }, "Format does not contain Tracewidth", "Warning");
                //}

            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, dgv_data.Name + " No data", "Warning");
            }
        }

        public void check_OOS_2(DataGridView dgv_data, DataGridView dgv_spec)
        {
            if (dgv_data.Rows.Count > 0)
            {

                for (int i = 0; i < dgv_data.Rows.Count; i++)
                {
                    int region = int.Parse(dgv_data.Rows[i].Cells["Region"].Value.ToString());
                    Double max_target = Double.Parse(dgv_spec.Rows[region - 1].Cells["USL"].Value.ToString());
                    Double min_target = Double.Parse(dgv_spec.Rows[region - 1].Cells["LSL"].Value.ToString());
                    Double data = Double.Parse(dgv_data.Rows[i].Cells["Data"].Value.ToString());
                    if (data > max_target || data < min_target)
                    {
                        for (int j = 0; j < dgv_data.Columns.Count; j++)
                        {
                            dgv_data.Rows[i].Cells[j].Style.BackColor = Color.Red;
                        }
                    }

                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, dgv_data.Name + " No data", "Warning");
            }
        }

        //private void btn_save_specimp_Click(object sender, EventArgs e)
        //{
        //    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });

        //    start_lbl: DataTable cur_dt = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_IMPEDANCE", filter_str);
        //    if (cur_dt.Rows.Count > 0)
        //    {
        //        if (MessageBox.Show("Do you want to overwrite?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
        //        {
        //            TDMK_Code.Delelte_FilteredItem_arr("SPEC_IMPEDANCE", sqlcon, filter_str);
        //            goto start_lbl;
        //        }
        //    }
        //    else
        //    {
        //        DataTable dt = cur_dt.Clone();
        //        int id = TDMK_Code.SQL_MAX("SPEC_IMPEDANCE", "ID", sqlcon);

        //        dt.Rows.Add(id, txtItemCode.Text, txt_max_imp.Text, txt_min_imp.Text, txt_max_tracewidth.Text, txt_min_tracewidth.Text);

        //        exp_proc.BatchBulkCopy(sqlcon, dt, "SPEC_IMPEDANCE");
        //        MessageBox.Show("Save data completed", "Warning");
        //    }

        //}


        //private void cb_tracewidth_Click(object sender, EventArgs e)
        //{
        //    if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtOperator.Text != "")
        //    {

        //        cb_tracewidth.Items.Clear();
        //        string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
        //        DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", filter_str);
        //        DataTable dt_imp = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", filter_str);
        //        if (dt_imp.Rows.Count > 0)
        //        {
        //            if (dt.Rows.Count > 0)
        //            {
        //                for (int i = 0; i < dt.Rows.Count; i++)
        //                {
        //                    cb_tracewidth.Items.Add("Tracewidth_" + (i + 1).ToString());
        //                }
        //                if (DGV_Tracewidth_spec.DataSource == null)
        //                    DGV_Tracewidth_spec.DataSource = dt;
        //            }
        //        }
        //        else
        //        {
        //            MessageBox.Show(new Form { TopMost = true }, "Please, Click SETUP SPEC", "Warning");
        //        }
        //    }

        //}

        private void tsmSpecSetup_Click(object sender, EventArgs e)
        {

        }

        private void txtItemCode_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                if (txtItemCode.Text != "")
                {
                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
                    DataTable dt_impedance_spec = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", filter_str);
                    DataTable dt_tracewidth_spec = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", filter_str);
                    if (dt_impedance_spec.Rows.Count == 0 && dt_tracewidth_spec.Rows.Count == 0)
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Spec not installed", "Warning");
                        btn_summary_trw.Enabled = false;
                        btnSummary.Enabled = false;
                    }
                    else
                    {
                        DGV_Impedance_Spec.DataSource = dt_impedance_spec;
                        DGV_Tracewidth_spec.DataSource = dt_tracewidth_spec;
                    }
                }
            } 
        }

        //private void btnSaveLogFile_Click(object sender, EventArgs e)
        //{
        //    string[] arr_tbl_name = { "IMPEDANCE_VAL_LOGFILE", "IMPEDANCE_GRAPH" };
        //    List<string> tar_table_name = new List<string>() { };
        //    List<DataGridView> lst_DGV = new List<DataGridView> { DGV_Impedance, DGV_Graph };
        //    List<DataTable> tar_table = new List<DataTable>() { };
        //    int tbl_inx = 0;
        //    int no_table = 0;
        //    foreach (DataGridView dgv in lst_DGV)
        //    {
        //        if (dgv.DataSource != null)
        //        {
        //            tar_table.Add((DataTable)dgv.DataSource);
        //            tar_table_name.Add(arr_tbl_name[no_table]);
        //        }
        //        no_table++;
        //    }
        //    int region = TDMK_Code.SQL_MAX("IMPEDANCE_VAL_LOGFILE" , "Region", sqlcon) + 1;
        //    foreach (var t in tar_table_name)
        //    {
        //        List<string> item = new List<string>() { "ItemCode", "LotNo", "Zone" , "Remark"};
        //        List<string> item_val = new List<string>() { txtItemCode.Text, txtLotNo.Text , myCode.checkDBNull(cb_Impedance.SelectedItem) , filename};


        //        string filter_str = TDMK_Code.filter_str(item.ToArray(), item_val.ToArray());
        //        start_lbl: DataTable cur_dt = TDMK_Code.Datatable_Filter(sqlcon, t, filter_str);
        //        if (cur_dt.Rows.Count > 0)
        //        {
        //            if (MessageBox.Show("Table " + t + ": Data is existed. Do you want to overwrite?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
        //            {
        //                TDMK_Code.Delelte_FilteredItem_arr(t, sqlcon, filter_str);
        //                goto start_lbl;
        //            }
        //        }
        //        else
        //        {
        //            int id = TDMK_Code.SQL_MAX(t, "ID", sqlcon);

        //            int r_inx = 0;
        //            foreach (DataRow dr in tar_table[tbl_inx].Rows)
        //            {
        //                dr["ID"] = id + 1 + r_inx; 
        //                //if(t == "IMPEDANCE_GRAPH")
        //                //     dr["Region"] = region;

        //                //if (t == "IMPEDANCE_VAL_LOGFILE")
        //                //    dr["Region"] = TDMK_Code.SQL_MAX(t, "Region", sqlcon) + int.Parse(dr["Region"].ToString());
        //                r_inx++;
        //            }
        //            exp_proc.BatchBulkCopy(sqlcon, tar_table[tbl_inx], t);
        //        }

        //    tbl_inx++;

        //    }
        //    MessageBox.Show("Save data completed", "Warning");
        //}
        public void resize_image_dgv()
        {
            ((DataGridViewImageColumn)DGV_Graph.Columns["Image_Graph"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            ((DataGridViewImageColumn)DGV_Graph.Columns["Image_Graph"]).Width = 100;
            foreach (DataGridViewRow dr in DGV_Graph.Rows)
            {
                dr.Height = 70;
            }

            ((DataGridViewImageColumn)DGV_Image_Tracewidth.Columns["Image_Tracewidth"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            ((DataGridViewImageColumn)DGV_Image_Tracewidth.Columns["Image_Tracewidth"]).Width = 100;
            foreach (DataGridViewRow dr in DGV_Image_Tracewidth.Rows)
            {
                dr.Height = 70;
            }
        }


        private void btnLoadLogFile_Click(object sender, EventArgs e)
        {

            if (txtItemCode.Text != "" && txtLotNo.Text != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                DGV_Impedance_Summary.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_VAL", filter_str);
                DGV_Graph.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", filter_str);
                DGV_Impedance_Spec.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));

                ((DataGridViewImageColumn)DGV_Graph.Columns["Image_Graph"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)DGV_Graph.Columns["Image_Graph"]).Width = 100;
                foreach (DataGridViewRow dr in DGV_Graph.Rows)
                {
                    dr.Height = 70;
                }


                DGV_Tracewidth_spec.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
                DGV_VHX_Data.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_VAL", filter_str);
                DGV_Image_Tracewidth.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_IMAGE", filter_str);

                ((DataGridViewImageColumn)DGV_Image_Tracewidth.Columns["Image_Tracewidth"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)DGV_Image_Tracewidth.Columns["Image_Tracewidth"]).Width = 100;
                foreach (DataGridViewRow dr in DGV_Image_Tracewidth.Rows)
                {
                    dr.Height = 70;
                }

                //pic_imp.Image = null;
                //pic_trw.Image = null;
                DGV_Impedance.DataSource = null;
              


                if (DGV_Impedance_Summary.Rows.Count > 0)
                {
                    check_OOS_Impedance(DGV_Impedance_Summary, DGV_Impedance_Spec);
                }
                if (DGV_VHX_Data.Rows.Count > 0)
                {
                    check_OOS_Trw(DGV_VHX_Data, DGV_Tracewidth_spec);
                }
            }
            else
            {
                MessageBox.Show("Please, fill in ItemCode / LotNo / Operator ID", "Warning");
            }

        }
        public void insert_data_trw(DataGridView src_dgv, DataTable tbl_data)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });

            DataTable dt = new DataTable();
            if (src_dgv.Rows.Count > 0)
            {
                dt = ((DataTable)src_dgv.DataSource);
            }
            else
            {
                dt = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_VAL", filter_str).Clone();
            }

            foreach (DataRow dr in tbl_data.Rows)
            {
                DataRow row = dt.NewRow();
                foreach (DataColumn dc in tbl_data.Columns)
                {
                    if (dc.ColumnName == "ID")
                    {

                        row["ID"] = dt.Rows.Count + 1;
                    }
                    else
                    {
                        row[dc.ColumnName] = dr[dc.ColumnName];
                    }

                }
                dt.Rows.Add(row);
            }
            src_dgv.DataSource = dt;


        }
        public void sort_data_impedance(DataGridView src_dgv_data, string tbl_name, DataTable dt_insert)
        {

            if (dt_insert.Rows.Count > 0)
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });

                DataTable src_Impedance_dt = dt_insert;
                DataTable Imp_sumary_dt = new DataTable();
                if (src_dgv_data.Rows.Count > 0)
                {
                    Imp_sumary_dt = ((DataTable)src_dgv_data.DataSource);
                }
                else
                {
                    Imp_sumary_dt = TDMK_Code.Datatable_Filter(sqlcon, tbl_name, filter_str).Clone();
                }


                DataTable coupon_dt = new DataTable();
                DataTable patern_dt = new DataTable();


                try
                {
                    coupon_dt = Imp_sumary_dt.AsEnumerable().Where(x => x.Field<string>("Zone") == "Coupon").CopyToDataTable();
                }
                catch
                {
                    // coupon_dt = Imp_sumary_dt.Clone();
                    coupon_dt = TDMK_Code.Datatable_Filter(sqlcon, tbl_name, filter_str).Clone();
                }
                try
                {
                    patern_dt = Imp_sumary_dt.AsEnumerable().Where(x => x.Field<string>("Zone") == "Patern").CopyToDataTable();
                }
                catch
                {
                    //  patern_dt = Imp_sumary_dt.Clone();
                    patern_dt = TDMK_Code.Datatable_Filter(sqlcon, tbl_name, filter_str).Clone();
                }


                int patern_pos = Convert.ToInt32(coupon_dt.AsEnumerable().Select(x => x.Field<string>("Region")).Max());


                if (src_Impedance_dt.Rows[0]["Zone"].ToString() == "Coupon")
                {
                    Imp_sumary_dt = coupon_dt;
                    foreach (DataRow dr in src_Impedance_dt.Rows)
                    {
                        DataRow coupon_dr = Imp_sumary_dt.NewRow();

                        foreach (DataColumn dc in Imp_sumary_dt.Columns)
                        {

                            if (dc.ColumnName != "Region")
                            {
                                if (dc.ColumnName == "ID")
                                {

                                    coupon_dr["ID"] = Imp_sumary_dt.Rows.Count + 1;
                                }
                                else
                                {
                                    coupon_dr[dc.ColumnName] = dr[dc.ColumnName];
                                }

                            }
                            else
                            {
                                coupon_dr[dc.ColumnName] = Convert.ToInt32(dr[dc.ColumnName]) + patern_pos;
                            }
                        }
                        Imp_sumary_dt.Rows.Add(coupon_dr);
                    }

                    foreach (DataRow dr in patern_dt.Rows)
                    {
                        // int id = Convert.ToInt32(Imp_sumary_dt.AsEnumerable().Select(x => x.Field<string>("ID")).Max());
                        DataRow partern_dr = Imp_sumary_dt.NewRow();
                        foreach (DataColumn dc in Imp_sumary_dt.Columns)
                        {
                            if (dc.ColumnName != "Region")
                            {
                                if (dc.ColumnName == "ID")
                                {
                                    partern_dr["ID"] = Imp_sumary_dt.Rows.Count + 1;
                                }
                                else
                                {
                                    partern_dr[dc.ColumnName] = dr[dc.ColumnName];
                                }
                            }
                            else
                            {
                                partern_dr[dc.ColumnName] = Convert.ToInt32(dr[dc.ColumnName]) + 1;
                            }
                        }
                        Imp_sumary_dt.Rows.Add(partern_dr);
                    }
                }
                else
                {
                    int region = Convert.ToInt32(Imp_sumary_dt.AsEnumerable().Select(x => x.Field<string>("Region")).Max());
                    foreach (DataRow dr in src_Impedance_dt.Rows)
                    {
                        // int id = Convert.ToInt32(Imp_sumary_dt.AsEnumerable().Select(x => x.Field<string>("ID")).Max());
                        DataRow coupon_dr = Imp_sumary_dt.NewRow();
                        foreach (DataColumn dc in Imp_sumary_dt.Columns)
                        {
                            if (dc.ColumnName != "Region")
                            {
                                if (dc.ColumnName == "ID")
                                {

                                    coupon_dr["ID"] = Imp_sumary_dt.Rows.Count + 1;
                                }
                                else
                                {
                                    coupon_dr[dc.ColumnName] = dr[dc.ColumnName];
                                }

                            }
                            else
                            {
                                // coupon_dr[dc.ColumnName] = Convert.ToInt32(dr[dc.ColumnName]) + patern_pos;
                                coupon_dr[dc.ColumnName] = region + int.Parse(dr["Region"].ToString());
                            }
                        }
                        Imp_sumary_dt.Rows.Add(coupon_dr);
                    }
                }
                src_dgv_data.DataSource = Imp_sumary_dt;
            }
        }
        public void sort_image_impedance(DataGridView src_dgv_data, string tbl_name, DataTable dt_insert)
        {
            if (dt_insert.Rows.Count > 0)
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });

                DataTable src_Impedance_dt = dt_insert;
                DataTable Imp_sumary_dt = Graph_dt;
                //if (src_dgv_data.Rows.Count > 0)
                //{
                //    Imp_sumary_dt = ((DataTable)src_dgv_data.DataSource);
                //}
                //else
                //{
                //    Imp_sumary_dt = TDMK_Code.Datatable_Filter(sqlcon, tbl_name, filter_str).Clone();
                //}

                DataTable coupon_dt = new DataTable();
                DataTable patern_dt = new DataTable();

                try
                {
                    coupon_dt = Imp_sumary_dt.AsEnumerable().Where(x => x.Field<string>("Zone") == "Coupon").CopyToDataTable();
                }
                catch
                {
                    // coupon_dt = Imp_sumary_dt.Clone();
                    coupon_dt = TDMK_Code.Datatable_Filter(sqlcon, tbl_name, filter_str).Clone();
                }
                try
                {
                    patern_dt = Imp_sumary_dt.AsEnumerable().Where(x => x.Field<string>("Zone") == "Patern").CopyToDataTable();
                }
                catch
                {
                    //  patern_dt = Imp_sumary_dt.Clone();
                    patern_dt = TDMK_Code.Datatable_Filter(sqlcon, tbl_name, filter_str).Clone();
                }


                int patern_pos = Convert.ToInt32(coupon_dt.AsEnumerable().Select(x => x.Field<string>("Region")).Max());


                if (src_Impedance_dt.Rows[0]["Zone"].ToString() == "Coupon")
                {
                    Imp_sumary_dt = coupon_dt;
                    foreach (DataRow dr in src_Impedance_dt.Rows)
                    {
                        DataRow coupon_dr = Imp_sumary_dt.NewRow();

                        foreach (DataColumn dc in Imp_sumary_dt.Columns)
                        {
                            if (dc.ColumnName != "Region")
                            {
                                if (dc.ColumnName == "ID")
                                {

                                    coupon_dr["ID"] = Imp_sumary_dt.Rows.Count + 1;
                                }
                                else
                                {
                                    coupon_dr[dc.ColumnName] = dr[dc.ColumnName];
                                }
                            }
                            else
                            {
                                coupon_dr[dc.ColumnName] = Convert.ToInt32(dr[dc.ColumnName]) + patern_pos;
                            }
                        }
                        Imp_sumary_dt.Rows.Add(coupon_dr);
                    }

                    foreach (DataRow dr in patern_dt.Rows)
                    {
                        // int id = Convert.ToInt32(Imp_sumary_dt.AsEnumerable().Select(x => x.Field<string>("ID")).Max());
                        DataRow partern_dr = Imp_sumary_dt.NewRow();
                        foreach (DataColumn dc in Imp_sumary_dt.Columns)
                        {
                            if (dc.ColumnName != "Region")
                            {
                                if (dc.ColumnName == "ID")
                                {

                                    partern_dr["ID"] = Imp_sumary_dt.Rows.Count + 1;
                                }
                                else
                                {

                                    partern_dr[dc.ColumnName] = dr[dc.ColumnName];
                                }
                            }
                            else
                            {
                                partern_dr[dc.ColumnName] = Convert.ToInt32(dr[dc.ColumnName]) + 1;
                            }
                        }
                        Imp_sumary_dt.Rows.Add(partern_dr);
                    }
                }
                else
                {
                    int region = Convert.ToInt32(Imp_sumary_dt.AsEnumerable().Select(x => x.Field<string>("Region")).Max());
                    foreach (DataRow dr in src_Impedance_dt.Rows)
                    {
                        // int id = Convert.ToInt32(Imp_sumary_dt.AsEnumerable().Select(x => x.Field<string>("ID")).Max());
                        DataRow coupon_dr = Imp_sumary_dt.NewRow();
                        foreach (DataColumn dc in Imp_sumary_dt.Columns)
                        {
                            if (dc.ColumnName != "Region")
                            {
                                if (dc.ColumnName == "ID")
                                {

                                    coupon_dr["ID"] = Imp_sumary_dt.Rows.Count + 1;
                                }
                                else
                                {
                                    coupon_dr[dc.ColumnName] = dr[dc.ColumnName];
                                }
                            }
                            else
                            {
                                // coupon_dr[dc.ColumnName] = Convert.ToInt32(dr[dc.ColumnName]) + patern_pos;
                                coupon_dr[dc.ColumnName] = region + 1;
                            }
                        }
                        Imp_sumary_dt.Rows.Add(coupon_dr);
                    }
                }
                src_dgv_data.DataSource = Imp_sumary_dt;
            }
        }


        public void Impedance_update(int region, DataTable tbl_data_update, DataTable tbl_graph_update, DataGridView dgv_summary, DataGridView dgv_image, string remark)
        {
            DataTable dt_imp_summary = (DataTable)dgv_summary.DataSource;
            DataView dv_data = dt_imp_summary.AsDataView();

            string filter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { region.ToString() });
            dv_data.RowFilter = filter;
            int min_pcs = new int[] { dv_data.Count, tbl_data_update.Rows.Count }.Min();
            for (int inx = 0; inx < min_pcs; inx++)
            {
                DataRow dr = dv_data[inx].Row;
                int r_x = dt_imp_summary.Rows.IndexOf(dr);
                dgv_summary.Rows[r_x].Cells["Data"].Value = tbl_data_update.Rows[inx]["Data"];
                dgv_summary.Rows[r_x].Cells["Zone"].Value = tbl_data_update.Rows[inx]["Zone"];
                dgv_summary.Rows[r_x].Cells["Remark"].Value = remark;
                // DGV_Zone_Data.Rows[inx].Cells["PCS_No" + item].Style.BackColor = Color.Red;

            }

            //  DataTable dt_graph_summary = (DataTable)dgv_image.DataSource;
            DataView dv_image = Graph_dt.AsDataView();

            dv_image.RowFilter = filter;
            min_pcs = new int[] { dv_image.Count, dgv_image.Rows.Count }.Min();
            for (int inx = 0; inx < min_pcs; inx++)
            {
                DataRow dr = dv_image[inx].Row;
                int r_x = Graph_dt.Rows.IndexOf(dr);
                Graph_dt.Rows[r_x]["Image_Graph"] = tbl_graph_update.Rows[inx]["Image_Graph"];
                Graph_dt.Rows[r_x]["Zone"] = tbl_graph_update.Rows[inx]["Zone"];
                Graph_dt.Rows[r_x]["Remark"] = remark;
                // DGV_Zone_Data.Rows[inx].Cells["PCS_No" + item].Style.BackColor = Color.Red;

            }

            dgv_image.DataSource = Graph_dt;

            if (edited_lst.Values.ToList().IndexOf(new List<string> { region_imp_update.ToString() }) == -1)
            {
                edited_lst.Add("Impedance_" + region_imp_update.ToString(), new List<string> { region_imp_update.ToString() });
            }


        }

        private void btnSummary_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtOperator.Text != "" && cb_Impedance.SelectedIndex != -1 && txtLogfile_Imp.Text != "")
            {

                DirectoryInfo temp_dir = new DirectoryInfo(txtLogfile_Imp.Text);

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
                    Graph_dt = (DataTable)DGV_Graph.DataSource;
                    Impedance_Data_Process_0(txtLogfile_Imp.Text);
                    if (DGV_Impedance.Rows.Count > 0)
                    {
                        DataTable dt_insert = (DataTable)DGV_Impedance.DataSource;
                        // Graph_dt = (DataTable)DGV_Graph.DataSource;
                        if (update_mode && region_imp_update != 0)
                        {
                            //Graph_dt = (DataTable)DGV_Graph.DataSource;
                            //  DataTable dt = Graph_dt;
                            DataTable dt_graph_update = (DataTable)DGV_Graph.DataSource;
                            Impedance_update(region_imp_update, dt_insert, dt_graph_update, DGV_Impedance_Summary, DGV_Graph, txtLogfile_Imp.Text);

                            ((DataGridViewImageColumn)DGV_Graph.Columns["Image_Graph"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                            ((DataGridViewImageColumn)DGV_Graph.Columns["Image_Graph"]).Width = 100;
                            foreach (DataGridViewRow dr in DGV_Graph.Rows)
                            {
                                dr.Height = 70;
                            }
                            MessageBox.Show("Replace data completed!", "Information");
                            update_mode = false;
                            // Graph_dt = new DataTable();
                            //   dt = (DataTable)DGV_Graph.DataSource;
                            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                            Graph_dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", filter_str).Clone();
                        }
                        else
                        {

                            sort_data_impedance(DGV_Impedance_Summary, "IMPEDANCE_VAL", dt_insert);
                            sort_image_impedance(DGV_Graph, "IMPEDANCE_GRAPH", (DataTable)DGV_Graph.DataSource);

                            ((DataGridViewImageColumn)DGV_Graph.Columns["Image_Graph"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                            ((DataGridViewImageColumn)DGV_Graph.Columns["Image_Graph"]).Width = 100;
                            foreach (DataGridViewRow dr in DGV_Graph.Rows)
                            {
                                dr.Height = 70;
                            }
                            Graph_dt = (DataTable)DGV_Graph.DataSource;
                        }



                        if (DGV_Impedance_Summary.Rows.Count > 0)
                        {
                            check_OOS_Impedance(DGV_Impedance_Summary, DGV_Impedance_Spec);
                        }

                        //btnLoadLogFile.Enabled = false;
                        // btnSummary.Enabled = false;
                    }
                }
                else
                {
                    MessageBox.Show("ItemCode or LotNo not matched", "Warning");
                }
            }
            else
            {
                MessageBox.Show("Nhập đầy đủ thông tin cho ItemCode / LotNo / Operator / Logfile", "Warning");
            }

        }

        private void DGV_Impedance_Summary_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btn_set_spec_Click(object sender, EventArgs e)
        {
            //tbl_format = new DataTable();
            //cb_tracewidth.Items.Clear();
            //string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
            //DataTable dt_format = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_FORMAT", filter_str);
            //start_lbl: if (dt_format.Rows.Count == 0)
            //{
            //    string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
            //    if (format_file != "")
            //    {
            //        myExcel.Workbook wb = TDMK_Code.open_excel_file(format_file, "", "");
            //        string _process = "Impedance".Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
            //        string mySheet = "";
            //        foreach (myExcel.Worksheet tg_sht in wb.Worksheets)
            //        {
            //            string cur_sht_name = tg_sht.Name.ToUpper().Replace("-", "").Replace(" ", "");
            //            if (cur_sht_name == _process)
            //            {
            //                mySheet = tg_sht.Name;
            //                break;
            //            }
            //        }
            //        if (mySheet != "")
            //        {
            //            myExcel.Worksheet curr_wrksheet = wb.Sheets[mySheet];
            //            curr_wrksheet.Activate();
            //            DataRow dr = dt_format.NewRow();
            //            dr["ItemCode"] = txtItemCode.Text;
            //            string Imp_range = "";
            //            string Imp_max = "";
            //            string Imp_min = "";
            //            foreach (string item in Get_impedance_infor_to_input_value_2(curr_wrksheet))
            //            {
            //                Imp_range += item.Split('+')[0] + ";";
            //                Imp_max += item.Split('+')[1] + ";";
            //                Imp_min += item.Split('+')[2] + ";";

            //            }
            //            dr["ID"] = TDMK_Code.SQL_MAX("IMPEDANCE_FORMAT", "ID", sqlcon) + 1;
            //            dr["Impedance_range"] = Imp_range.Remove(Imp_range.Length - 1);
            //            dr["Impedance_USL"] = Imp_max.Remove(Imp_max.Length - 1);
            //            dr["Impedance_LSL"] = Imp_min.Remove(Imp_min.Length - 1);
            //            string Trw_range = "";
            //            string Trw_max = "";
            //            string Trw_min = "";
            //            foreach (string item in Get_tracewidth_infor_to_input_value_2(curr_wrksheet))
            //            {
            //                Trw_range += item.Split('+')[0] + ";";
            //                Trw_max += item.Split('+')[1] + ";";
            //                Trw_min += item.Split('+')[2] + ";";
            //            }
            //            if (Trw_range != "")
            //            {
            //                dr["Tracewidth_range"] = Trw_range.Remove(Trw_range.Length - 1);
            //            }
            //            if (Trw_max != "")
            //            {
            //                dr["Tracewidth_USL"] = Trw_max.Remove(Trw_max.Length - 1);
            //            }
            //            if (Trw_min != "")
            //            {
            //                dr["Tracewidth_LSL"] = Trw_min.Remove(Trw_min.Length - 1);
            //            }
            //            dt_format.Rows.Add(dr);
            //            exp_proc.BatchBulkCopy(sqlcon, dt_format, "IMPEDANCE_FORMAT");
            //            DGV_Image_Graph.DataSource = dt_format;
            //            wb.Close();
            //            MessageBox.Show("Spec setup completed!", "Information");
            //        }
            //        else
            //        {
            //            MessageBox.Show("Sheet Impedance not found", "Warning");
            //            wb.Close();
            //        }
            //    }
            //    else
            //    {
            //        MessageBox.Show("Format of " + txtItemCode.Text + " not found", "Warning");
            //    }
            //}
            //else
            //{
            //    if (MessageBox.Show("Do you want to update Spec again?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
            //    {
            //        TDMK_Code.Delelte_FilteredItem_arr("IMPEDANCE_FORMAT", sqlcon, filter_str);
            //        dt_format = dt_format.Clone();
            //        goto start_lbl;
            //    }
            //    else
            //    {
            //        DGV_Image_Graph.DataSource = dt_format;
            //    }
            //}


            tbl_format = new DataTable();
            //  cb_tracewidth.Items.Clear();
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
            DataTable dt_format = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", filter_str);
        //  DataTable dt_spec_impedance = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", filter_str).Clone();
        start_lbl: if (dt_format.Rows.Count == 0)
            {
                string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                if (format_file != "")
                {
                    myExcel.Workbook wb = TDMK_Code.open_excel_file(format_file, "", "");
                    string _process = "Impedance".Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
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
                        //myExcel.Worksheet curr_wrksheet = wb.Sheets[mySheet];
                        //curr_wrksheet.Activate();
                        //DataRow dr = dt_format.NewRow();
                        //dr["ItemCode"] = txtItemCode.Text;
                        //string Imp_range = "";
                        //string Imp_max = "";
                        //string Imp_min = "";
                        //foreach (string item in Get_impedance_infor_to_input_value_2(curr_wrksheet))
                        //{
                        //    Imp_range += item.Split('+')[0] + ";";
                        //    Imp_max += item.Split('+')[1] + ";";
                        //    Imp_min += item.Split('+')[2] + ";";

                        //}
                        //dr["ID"] = TDMK_Code.SQL_MAX("IMPEDANCE_FORMAT", "ID", sqlcon) + 1;
                        //dr["Impedance_range"] = Imp_range.Remove(Imp_range.Length - 1);
                        //dr["Impedance_USL"] = Imp_max.Remove(Imp_max.Length - 1);
                        //dr["Impedance_LSL"] = Imp_min.Remove(Imp_min.Length - 1);
                        //string Trw_range = "";
                        //string Trw_max = "";
                        //string Trw_min = "";
                        //foreach (string item in Get_tracewidth_infor_to_input_value_2(curr_wrksheet))
                        //{
                        //    Trw_range += item.Split('+')[0] + ";";
                        //    Trw_max += item.Split('+')[1] + ";";
                        //    Trw_min += item.Split('+')[2] + ";";
                        //}
                        //if (Trw_range != "")
                        //{
                        //    dr["Tracewidth_range"] = Trw_range.Remove(Trw_range.Length - 1);
                        //}
                        //if (Trw_max != "")
                        //{
                        //    dr["Tracewidth_USL"] = Trw_max.Remove(Trw_max.Length - 1);
                        //}
                        //if (Trw_min != "")
                        //{
                        //    dr["Tracewidth_LSL"] = Trw_min.Remove(Trw_min.Length - 1);
                        //}
                        //dt_format.Rows.Add(dr);
                        //exp_proc.BatchBulkCopy(sqlcon, dt_format, "IMPEDANCE_FORMAT");
                        //DGV_Graph.DataSource = dt_format;
                        //wb.Close();
                        
                        

                        MessageBox.Show(new Form { TopMost = true }, "Spec setup completed!", "Information");
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Sheet Impedance not found", "Warning");
                        wb.Close();
                    }
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Format of " + txtItemCode.Text + " not found", "Warning");
                }
            }
            else
            {
                if (MessageBox.Show(new Form { TopMost = true }, "Do you want to update Spec again?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    TDMK_Code.Delelte_FilteredItem_arr("IMPEDANCE_FORMAT", sqlcon, filter_str);
                    dt_format = dt_format.Clone();
                    goto start_lbl;
                }
                else
                {
                    DGV_Graph.DataSource = dt_format;
                }
            }





        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void btn_setup_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "")
            {
                tbl_format = new DataTable();
                // cb_tracewidth.Items.Clear();
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
                // DataTable dt_format = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_FORMAT", filter_str);
                DataTable dt_impedance_spec = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", filter_str);
                DataTable dt_tracewidth_spec = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", filter_str);
            start_lbl: if (dt_impedance_spec.Rows.Count == 0 && dt_tracewidth_spec.Rows.Count == 0)
                {
                    string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                    if (format_file != "")
                    {
                        //ExcelWorkbook wb = TDMK_Code2.open_excel_file(format_file);
                        //F_setup_new.setup_Impedance_new(txtItemCode.Text, wb, sqlcon); 
 
                        myExcel.Workbook wb = TDMK_Code.open_excel_file(format_file, "", "");
                        string _process = "Impedance".Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
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
                            myExcel.Worksheet curr_wrksheet = wb.Sheets[mySheet];
                            curr_wrksheet.Activate();

                            int i = 0;
                            int id_imp = TDMK_Code.SQL_MAX("IMPEDANCE_SPEC", "ID", sqlcon);
                            AutoCompleteStringCollection lst_info_imp = Get_impedance_infor_to_input_value_2(curr_wrksheet);

                            foreach (string item in lst_info_imp)
                            {
                                DataRow dr = dt_impedance_spec.NewRow();
                                dr["ID"] = id_imp + i + 1;
                                dr["ItemCode"] = txtItemCode.Text;
                                dr["Region"] = "Impedance_" + (i + 1).ToString();
                                dr["Range"] = item.Split('+')[0];
                                dr["USL"] = item.Split('+')[1];
                                dr["LSL"] = item.Split('+')[2];
                                dt_impedance_spec.Rows.Add(dr);
                                i++;
                            }
                            i = 0;
                            int id_trace = TDMK_Code.SQL_MAX("TRACEWIDTH_SPEC", "ID", sqlcon);

                            AutoCompleteStringCollection lst_info_trw = Get_tracewidth_infor_to_input_value_2(curr_wrksheet);
                            foreach (string item in lst_info_trw)
                            {
                                DataRow dr = dt_tracewidth_spec.NewRow();
                                dr["ID"] = id_trace + i + 1;
                                dr["ItemCode"] = txtItemCode.Text;
                                dr["Region"] = item.Split('+')[3];
                                dr["Range"] = item.Split('+')[0];
                                dr["USL"] = item.Split('+')[1];
                                dr["LSL"] = item.Split('+')[2];
                                dt_tracewidth_spec.Rows.Add(dr);
                                i++;
                            }
                            exp_proc.BatchBulkCopy(sqlcon, dt_impedance_spec, "IMPEDANCE_SPEC");
                            exp_proc.BatchBulkCopy(sqlcon, dt_tracewidth_spec, "TRACEWIDTH_SPEC");
                            DGV_Impedance_Spec.DataSource = dt_impedance_spec;
                            DGV_Tracewidth_spec.DataSource = dt_tracewidth_spec;
                            wb.Close();
                            MessageBox.Show(new Form { TopMost = true }, "Spec setup completed!", "Information");

                        }
                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Sheet Impedance not found", "Warning");
                            wb.Close();
                        }
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Format of " + txtItemCode.Text + " not found", "Warning");
                    }
                }
                else
                {
                    DGV_Impedance_Spec.DataSource = dt_impedance_spec;
                    DGV_Tracewidth_spec.DataSource = dt_tracewidth_spec;

                    if (MessageBox.Show(new Form { TopMost = true }, "Do you want to update Spec again?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        TDMK_Code.Delelte_FilteredItem_arr("IMPEDANCE_SPEC", sqlcon, filter_str);
                        TDMK_Code.Delelte_FilteredItem_arr("TRACEWIDTH_SPEC", sqlcon, filter_str);
                        DGV_Impedance_Spec.DataSource = DGV_Tracewidth_spec.DataSource = null;
                        dt_impedance_spec = dt_impedance_spec.Clone();
                        dt_tracewidth_spec = dt_tracewidth_spec.Clone();
                        goto start_lbl;
                    } 
                }
            }
             

        }

        private void splitMain_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void DGV_Impedance_Spec_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void tableLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void btn_save_all_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtOperator.Text != "")
            {
                if (DGV_Impedance_Summary.DataSource == null && DGV_VHX_Data.DataSource == null)
                {
                    MessageBox.Show(new Form { TopMost = true }, "No data", "Warning");
                }
                else
                { 
                    DataTable dt_imp_sum = (DataTable)DGV_Impedance_Summary.DataSource;
                    DataTable dt_trw = (DataTable)DGV_VHX_Data.DataSource;

                    if (DGV_Impedance_Summary.Rows.Count != 0)
                    {
                        string[] count_Imp = dt_imp_sum.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();
                        int count_Imp_spec = DGV_Impedance_Spec.Rows.Count;

                        if (count_Imp.Length < count_Imp_spec)
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Not enough data Impedance. Cannot save to database", "Warning");
                        }
                        else
                        { 
                            string[] arr_tbl_name = { "IMPEDANCE_VAL", "IMPEDANCE_GRAPH" };
                            List<string> tar_table_name = new List<string>() { };
                            List<DataGridView> lst_DGV = new List<DataGridView> { DGV_Impedance_Summary, DGV_Graph };
                            List<DataTable> tar_table = new List<DataTable>() { };
                            int tbl_inx = 0;
                            int no_table = 0;
                            foreach (DataGridView dgv1 in lst_DGV)
                            {
                                if (dgv1.DataSource != null)
                                {
                                    tar_table.Add((DataTable)dgv1.DataSource);
                                    tar_table_name.Add(arr_tbl_name[no_table]);
                                }
                                no_table++;
                            }
                            bool chk = true;

                            DataGridView dgv = lst_DGV[0];
                            for (int k = 0; k < dgv.Rows.Count; k++)
                            {
                                if (dgv.Rows[k].Cells[1].Style.BackColor == Color.Red)
                                {
                                    //MessageBox.Show("Table " + dgv.Name + "could not save data", "Warning");
                                    chk = false;
                                    break;
                                } 
                            }

                            bool save_history = false; 
                            save_lbl: if (chk)
                            {
                                foreach (var t in tar_table_name)
                                {
                                    /* start_lbl:*/
                                    if (!save_history)
                                    {

                                        Update_history_update("Impedance");
                                        save_history = true;
                                    }
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

                                MessageBox.Show(new Form { TopMost = true }, "Save data Impedance completed", "Warning");
                                sel_sample_lst = new List<string> { };
                            }
                            else
                            {
                                if (MessageBox.Show(new Form { TopMost = true }, "Data of Impedance is out of spec. Do you want to save to database ?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                                {
                                    if (admin_mode)
                                    {
                                        chk = true;
                                        goto save_lbl;

                                    }
                                    else
                                    {
                                        MessageBox.Show("Please, login to save data", "Warning");
                                        return;
                                    } 
                                } 
                            }

                        }


                    }
                    if (DGV_VHX_Data.Rows.Count != 0)
                    {

                        string[] count_trw = dt_trw.AsEnumerable().Select(x => x.Field<string>("Data_For")).Distinct().ToArray();
                        List<string> lst_region_trw = ((DataTable)DGV_Tracewidth_spec.DataSource).AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToList();
                        int count_trw_spec = DGV_Tracewidth_spec.Rows.Count;

                        bool chk_region = lst_region_trw.All(x => count_trw.Any(s => s.Contains(x) && x != ""));
                        List<string> lst_not_exist = count_trw.Where(x => !lst_region_trw.Any(s => x.Contains(s) && s != "")).ToList();

                        if (count_trw.Length < count_trw_spec)
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Not enough data Tracewidth. Cannot save to database", "Warning");
                        }
                        else if (!chk_region)
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Kiểm tra lại thông tin các folder. Các vùng " + string.Join(",", lst_not_exist) + " không có trong format", "Warning");
                        }
                        else { 

                            string[] arr_tbl_name = { "TRACEWIDTH_VAL", "TRACEWIDTH_IMAGE" };
                            List<string> tar_table_name = new List<string>() { };
                            List<DataGridView> lst_DGV = new List<DataGridView> { DGV_VHX_Data, DGV_Image_Tracewidth };
                            List<DataTable> tar_table = new List<DataTable>() { };
                            int tbl_inx = 0;
                            int no_table = 0;
                            foreach (DataGridView dgv1 in lst_DGV)
                            {
                                if (dgv1.DataSource != null)
                                {
                                    tar_table.Add((DataTable)dgv1.DataSource);
                                    tar_table_name.Add(arr_tbl_name[no_table]);
                                }
                                no_table++;
                            }
                            bool chk = true;

                            DataGridView dgv = lst_DGV[0];
                            for (int k = 0; k < dgv.Rows.Count; k++)
                            {
                                if (dgv.Rows[k].Cells[1].Style.BackColor == Color.Red)
                                {

                                    chk = false;
                                    break;
                                }

                            }


                            bool save_history = false;

                        save_lbl: if (chk)
                            {
                                foreach (var t in tar_table_name)
                                {
                                    /* start_lbl:*/
                                    if (!save_history)
                                    {
                                        Update_history_update("Tracewidth");
                                        save_history = true;
                                    }
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

                                MessageBox.Show(new Form { TopMost = true }, "Save data Tracewidth completed", "Warning");
                                sel_sample_lst = new List<string> { };
                            }
                            else
                            {
                                if (MessageBox.Show(new Form { TopMost = true }, "Data of Tracewidth is out of spec. Do you want to save to database ?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                                {
                                    if (admin_mode)
                                    {
                                        chk = true;
                                        goto save_lbl;

                                    }
                                    else
                                    {
                                        MessageBox.Show("Please, login to save data", "Warning");
                                        return;
                                    }

                                }

                            }

                        }


                    }


                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Please, fill in ItemCode / LotNo / Operator/ Process", "Warning");
            }

        }
        public void Update_history_update(string process)
        {
            //DataTable logfile_before = TDMK_Code.Datatable_Filter(sqlcon, tbl_data_name, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
            //DataTable image_before = TDMK_Code.Datatable_Filter(sqlcon, tbl_image_name, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));

            List<string> item_arr = new List<string>() { "ID", "ItemCode", "LotNo", "Process", "PCS_No", "Region", "Zone", "Before_Data", "After_Data", "Time_Update", "Operator", "Depart", "Remark" };
            List<object> item_arr_val = new List<object>();
            DateTime tim_up = DateTime.Now;
            int id_edit = TDMK_Code.SQL_MAX("VHX_Edit_History", "ID", sqlcon);

            // DataTable logfile_after = (DataTable)dgv_data.DataSource;
            if (edited_lst.Count > 0)
            {
                foreach (var zone in edited_lst)
                {
                    if (zone.Key.Contains("Impedance") && process == "Impedance")
                    {
                        int count = 0;
                        DataTable logfile_after = (DataTable)DGV_Impedance_Summary.DataSource;
                        DataTable logfile_before = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_VAL", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                        DataTable image_before = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));

                        DataView dv = logfile_after.AsDataView();
                        string filter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { zone.Value[0] });
                        dv.RowFilter = filter;
                        for (int inx = 0; inx < dv.Count; inx++)
                        {
                            DataRow dr = dv[inx].Row;
                            int r_x = logfile_after.Rows.IndexOf(dr);
                            string before_data = logfile_before.Rows[r_x]["Data"].ToString();
                            string after_data = logfile_after.Rows[r_x]["Data"].ToString();
                            string remark = logfile_after.Rows[r_x]["Remark"].ToString();
                            string pcs = logfile_after.Rows[r_x]["Pcs_No"].ToString();
                            byte[] bef_data = System.Text.Encoding.UTF8.GetBytes(before_data);
                            byte[] aft_data = System.Text.Encoding.UTF8.GetBytes(after_data);

                            item_arr_val = new List<object>() { (id_edit + count + 1), txtItemCode.Text, txtLotNo.Text, "Impedance", pcs, "Logfile_Impedance_data", zone.Key, bef_data, aft_data, tim_up, txtUsername.Text, depart, remark };
                            exp_proc.Insert_Object_List("VHX_Edit_History", sqlcon, item_arr, item_arr_val);
                            count++;

                        }
                        DataTable image_after = (DataTable)DGV_Graph.DataSource;
                        DataView dv_image = image_after.AsDataView();
                        dv_image.RowFilter = filter;
                        for (int i = 0; i < dv_image.Count; i++)
                        {
                            DataRow dr_ = dv_image[i].Row;
                            int r_x_ = image_after.Rows.IndexOf(dr_);

                            byte[] img_bef = (byte[])(image_before.Rows[r_x_]["Image_Graph"]);
                            byte[] img_aft = (byte[])(image_after.Rows[r_x_]["Image_Graph"]);
                            string remark1 = image_after.Rows[r_x_]["Remark"].ToString();
                            string pcs1 = image_after.Rows[r_x_]["Pcs_No"].ToString();
                            item_arr_val = new List<object>() { (id_edit + count + 1), txtItemCode.Text, txtLotNo.Text, "Impedance", pcs1, "Image_data", zone.Key, img_bef, img_aft, tim_up, txtUsername.Text, depart, remark1 };
                            //  item_arr_val = new List<object>() { (id_edit + count + 1), txtItemCode.Text, txtLotNo.Text, "Stackup", pcs, "Image_data", "Zone_" + zone.Key, img_bef, img_aft, tim_up, txtOperator.Text, depart, remark1 };
                            exp_proc.Insert_Object_List("VHX_Edit_History", sqlcon, item_arr, item_arr_val);
                            count++;
                        }


                    }
                    else if (zone.Key == "Tracewidth" && process == "Tracewidth")
                    {
                        int count = 0;
                        foreach (var item in zone.Value)
                        {
                            DataTable logfile_after = (DataTable)DGV_VHX_Data.DataSource;
                            DataTable logfile_before = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_VAL", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                            DataTable image_before = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));

                            DataView dv = logfile_after.AsDataView();
                            string filter = TDMK_Code.filter_str(new string[] { "Pcs_No", "Data_For" }, new string[] { item.Split('+')[1], item.Split('+')[0] });

                            dv.RowFilter = filter;
                            for (int inx = 0; inx < dv.Count; inx++)
                            {
                                DataRow dr = dv[inx].Row;
                                int r_x = logfile_after.Rows.IndexOf(dr);
                                string before_data = logfile_before.Rows[r_x]["Data"].ToString();
                                string after_data = logfile_after.Rows[r_x]["Data"].ToString();
                                string remark = logfile_after.Rows[r_x]["Remark"].ToString();
                                string pcs = logfile_after.Rows[r_x]["Pcs_No"].ToString();
                                byte[] bef_data = System.Text.Encoding.UTF8.GetBytes(before_data);
                                byte[] aft_data = System.Text.Encoding.UTF8.GetBytes(after_data);

                                item_arr_val = new List<object>() { (id_edit + count + 1), txtItemCode.Text, txtLotNo.Text, "Tracewidth", pcs, "Logfile_Tracewidth_data", item.Split('+')[0], bef_data, aft_data, tim_up, txtUsername.Text, depart, remark };
                                exp_proc.Insert_Object_List("VHX_Edit_History", sqlcon, item_arr, item_arr_val);
                                count++;

                            }



                            DataTable image_after = (DataTable)DGV_Image_Tracewidth.DataSource;

                            DataView dv_image = image_after.AsDataView();

                            dv_image.RowFilter = filter;

                            for (int i = 0; i < dv_image.Count; i++)
                            {
                                DataRow dr_ = dv_image[i].Row;
                                int r_x_ = image_after.Rows.IndexOf(dr_);

                                byte[] img_bef = (byte[])(image_before.Rows[r_x_]["Image_Tracewidth"]);
                                byte[] img_aft = (byte[])(image_after.Rows[r_x_]["Image_Tracewidth"]);
                                string remark1 = image_after.Rows[r_x_]["Remark"].ToString();
                                string pcs1 = image_after.Rows[r_x_]["Pcs_No"].ToString();
                                item_arr_val = new List<object>() { (id_edit + count + 1), txtItemCode.Text, txtLotNo.Text, "Tracewidth", pcs1, "Image_data", item.Split('+')[0], img_bef, img_aft, tim_up, txtUsername.Text, depart, remark1 };
                                //  item_arr_val = new List<object>() { (id_edit + count + 1), txtItemCode.Text, txtLotNo.Text, "Stackup", pcs, "Image_data", "Zone_" + zone.Key, img_bef, img_aft, tim_up, txtOperator.Text, depart, remark1 };
                                exp_proc.Insert_Object_List("VHX_Edit_History", sqlcon, item_arr, item_arr_val);
                                count++;
                            }

                        }

                    }


                }

            }

        }
        public void Update_history()
        {
            //DataTable logfile_before = TDMK_Code.Datatable_Filter(sqlcon, tbl_data_name, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
            //DataTable image_before = TDMK_Code.Datatable_Filter(sqlcon, tbl_image_name, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));

            List<string> item_arr = new List<string>() { "ID", "ItemCode", "LotNo", "Process", "PCS_No", "Region", "Zone", "Before_Data", "After_Data", "Time_Update", "Operator", "Depart", "Remark" };
            List<object> item_arr_val = new List<object>();
            DateTime tim_up = DateTime.Now;
            int id_edit = TDMK_Code.SQL_MAX("VHX_Edit_History", "ID", sqlcon);

            // DataTable logfile_after = (DataTable)dgv_data.DataSource;
            if (edited_lst.Count > 0)
            {
                foreach (var zone in edited_lst)
                {
                    if (zone.Key.Contains("Impedance"))
                    {
                        int count = 0;
                        DataTable logfile_after = (DataTable)DGV_Impedance_Summary.DataSource;
                        DataTable logfile_before = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_VAL", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                        DataTable image_before = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));

                        DataView dv = logfile_after.AsDataView();
                        string filter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { zone.Value[0] });
                        dv.RowFilter = filter;
                        for (int inx = 0; inx < dv.Count; inx++)
                        {
                            DataRow dr = dv[inx].Row;
                            int r_x = logfile_after.Rows.IndexOf(dr);
                            string before_data = logfile_before.Rows[r_x]["Data"].ToString();
                            string after_data = logfile_after.Rows[r_x]["Data"].ToString();
                            string remark = logfile_after.Rows[r_x]["Remark"].ToString();
                            string pcs = logfile_after.Rows[r_x]["Pcs_No"].ToString();
                            byte[] bef_data = System.Text.Encoding.UTF8.GetBytes(before_data);
                            byte[] aft_data = System.Text.Encoding.UTF8.GetBytes(after_data);

                            item_arr_val = new List<object>() { (id_edit + count + 1), txtItemCode.Text, txtLotNo.Text, "Impedance", pcs, "Logfile_Impedance_data", zone.Key, bef_data, aft_data, tim_up, txtUsername.Text, depart, remark };
                            exp_proc.Insert_Object_List("VHX_Edit_History", sqlcon, item_arr, item_arr_val);
                            count++;

                        }
                        DataTable image_after = (DataTable)DGV_Graph.DataSource;
                        DataView dv_image = image_after.AsDataView();
                        dv_image.RowFilter = filter;
                        for (int i = 0; i < dv_image.Count; i++)
                        {
                            DataRow dr_ = dv_image[i].Row;
                            int r_x_ = image_after.Rows.IndexOf(dr_);

                            byte[] img_bef = (byte[])(image_before.Rows[r_x_]["Image_Graph"]);
                            byte[] img_aft = (byte[])(image_after.Rows[r_x_]["Image_Graph"]);
                            string remark1 = image_after.Rows[r_x_]["Remark"].ToString();
                            string pcs1 = image_after.Rows[r_x_]["Pcs_No"].ToString();
                            item_arr_val = new List<object>() { (id_edit + count + 1), txtItemCode.Text, txtLotNo.Text, "Impedance", pcs1, "Image_data", zone.Key, img_bef, img_aft, tim_up, txtUsername.Text, depart, remark1 };
                            //  item_arr_val = new List<object>() { (id_edit + count + 1), txtItemCode.Text, txtLotNo.Text, "Stackup", pcs, "Image_data", "Zone_" + zone.Key, img_bef, img_aft, tim_up, txtOperator.Text, depart, remark1 };
                            exp_proc.Insert_Object_List("VHX_Edit_History", sqlcon, item_arr, item_arr_val);
                            count++;
                        }


                    }
                    else if (zone.Key == "Tracewidth")
                    {
                        int count = 0;
                        foreach (var item in zone.Value)
                        {
                            DataTable logfile_after = (DataTable)DGV_VHX_Data.DataSource;
                            DataTable logfile_before = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_VAL", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                            DataTable image_before = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));

                            DataView dv = logfile_after.AsDataView();
                            string filter = TDMK_Code.filter_str(new string[] { "Pcs_No", "Data_For" }, new string[] { item.Split('+')[1], item.Split('+')[0] });

                            dv.RowFilter = filter;
                            for (int inx = 0; inx < dv.Count; inx++)
                            {
                                DataRow dr = dv[inx].Row;
                                int r_x = logfile_after.Rows.IndexOf(dr);
                                string before_data = logfile_before.Rows[r_x]["Data"].ToString();
                                string after_data = logfile_after.Rows[r_x]["Data"].ToString();
                                string remark = logfile_after.Rows[r_x]["Remark"].ToString();
                                string pcs = logfile_after.Rows[r_x]["Pcs_No"].ToString();
                                byte[] bef_data = System.Text.Encoding.UTF8.GetBytes(before_data);
                                byte[] aft_data = System.Text.Encoding.UTF8.GetBytes(after_data);

                                item_arr_val = new List<object>() { (id_edit + count + 1), txtItemCode.Text, txtLotNo.Text, "Tracewidth", pcs, "Logfile_Tracewidth_data", item.Split('+')[0], bef_data, aft_data, tim_up, txtUsername.Text, depart, remark };
                                exp_proc.Insert_Object_List("VHX_Edit_History", sqlcon, item_arr, item_arr_val);
                                count++;

                            }



                            DataTable image_after = (DataTable)DGV_Image_Tracewidth.DataSource;

                            DataView dv_image = image_after.AsDataView();

                            dv_image.RowFilter = filter;

                            for (int i = 0; i < dv_image.Count; i++)
                            {
                                DataRow dr_ = dv_image[i].Row;
                                int r_x_ = image_after.Rows.IndexOf(dr_);

                                byte[] img_bef = (byte[])(image_before.Rows[r_x_]["Image_Tracewidth"]);
                                byte[] img_aft = (byte[])(image_after.Rows[r_x_]["Image_Tracewidth"]);
                                string remark1 = image_after.Rows[r_x_]["Remark"].ToString();
                                string pcs1 = image_after.Rows[r_x_]["Pcs_No"].ToString();
                                item_arr_val = new List<object>() { (id_edit + count + 1), txtItemCode.Text, txtLotNo.Text, "Tracewidth", pcs1, "Image_data", item.Split('+')[0], img_bef, img_aft, tim_up, txtUsername.Text, depart, remark1 };
                                //  item_arr_val = new List<object>() { (id_edit + count + 1), txtItemCode.Text, txtLotNo.Text, "Stackup", pcs, "Image_data", "Zone_" + zone.Key, img_bef, img_aft, tim_up, txtOperator.Text, depart, remark1 };
                                exp_proc.Insert_Object_List("VHX_Edit_History", sqlcon, item_arr, item_arr_val);
                                count++;
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

        private void btn_export_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "")
            {
                // Export_Impedance_data(sqlcon, txtItemCode.Text, txtLotNo.Text);
                string format_file = exp_proc.find_format(myVar.data_loc, txtItemCode.Text);
                if (format_file != "")
                {
                    string export_path = F_exportEPPlus.get_reportpath(txtItemCode.Text, txtLotNo.Text, txtOperator.Text, "Impedance", myVar.data_loc);
                    if (export_path != "")
                    {
                        F_exportEPPlus.Export_EPPlus(sqlcon, format_file, export_path, txtItemCode.Text, txtLotNo.Text, "IMPEDANCE", null);
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Folder Impedance not found!", "Warning");
                    }
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Please, fill in ItemCode / LotNo ", "Warning");
            }
          
        }
        public void replace_tracewidth(DataGridView dgv)
        {
            //if (dgv.Rows.Count > 0)
            //{
            //    DataTable dt_summary = (DataTable)dgv.DataSource;
            //    DataView dv = dt_summary.AsDataView();
            //    string filter = TDMK_Code.filter_str(new string[] { "Data_For" }, new string[] { cb_tracewidth.SelectedItem.ToString() });

            //    dv.RowFilter = filter;
            //    int c = dv.Count;
            //    if (c > 0)
            //    {
            //        for (int inx = 0; inx < c; inx++)
            //        {
            //            DataRow dr = dv[0].Row;
            //            //int r_x = dt_summary.Rows.IndexOf(dr);
            //            dt_summary.Rows.Remove(dr);

            //        }
            //    }
            //    dgv.DataSource = dt_summary;

            //}
        }

        private void btn_summary_trw_Click(object sender, EventArgs e)
        {

            if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtLogfile_Trw.Text != "")
            {

                DirectoryInfo temp_dir = new DirectoryInfo(txtLogfile_Trw.Text);

                string f_name = temp_dir.Name;
                bool chk_name = false;

                if(!update_mode_trw)
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
                else
                {
                    chk_name = true;
                }
                 

                if (chk_name)
                {
                 
                    DGV_Tracewidth_spec.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));

                    if (update_mode_trw)
                    {
                        Image_dt = (DataTable)DGV_Image_Tracewidth.DataSource;
                        // DGV_VHX_Data.DataSource = DGV_Image_Tracewidth.DataSource = null;
                        Tracewidth_Update(txtLogfile_Trw.Text, sel_sample_lst[0].Split('+')[0]);
                        Tracewidth_Data_Process_Update(txtLogfile_Trw.Text, sel_sample_lst);

                        check_OOS_Trw(DGV_VHX_Data, DGV_Tracewidth_spec);
                        update_mode_trw = false;
                        // Image_dt = (DataTable)DGV_Image_Tracewidth.DataSource;
                    }
                    else
                    {
                        //if (cb_tracewidth.SelectedIndex != -1)
                        //{ 
                        //    Tracewidth_Data_Process_2(txtLogfile_Trw.Text);

                        //    if (DGV_VHX_Data.Rows.Count > 0)
                        //    {

                        //        replace_tracewidth(DGV_VHX_Data);
                        //        if (Image_dt.Rows.Count > 0)
                        //        {

                        //            DataView dv = Image_dt.AsDataView();
                        //            string filter = TDMK_Code.filter_str(new string[] { "Data_For" }, new string[] { cb_tracewidth.SelectedItem.ToString() });
                        //            dv.RowFilter = filter;
                        //            int c = dv.Count;
                        //            if (c > 0)
                        //            {
                        //                for (int inx = 0; inx < c; inx++)
                        //                {
                        //                    DataRow dr = dv[0].Row;
                        //                    Image_dt.Rows.Remove(dr);
                        //                }
                        //            }
                        //        }

                        //        insert_data_trw(DGV_VHX_Data, (DataTable)DGV_VHX_Data.DataSource);

                        //        DataTable dt = (DataTable)DGV_Image_Tracewidth.DataSource;
                        //        foreach (DataRow dr in dt.Rows)
                        //        {
                        //            DataRow row = Image_dt.NewRow();
                        //            foreach (DataColumn dc in Image_dt.Columns)
                        //            {
                        //                if (dc.ColumnName == "ID")
                        //                {

                        //                    row["ID"] = Image_dt.Rows.Count + 1;
                        //                }
                        //                else
                        //                {
                        //                    row[dc.ColumnName] = dr[dc.ColumnName];
                        //                }

                        //            }
                        //            Image_dt.Rows.Add(row);
                        //        }
                        //        DGV_Image_Tracewidth.DataSource = Image_dt;

                        //        if (DGV_VHX_Data.Rows.Count > 0)
                        //        {
                        //            check_OOS_Trw(DGV_VHX_Data, DGV_Tracewidth_spec);
                        //        }
                        //    }
                        //}


                        DirectoryInfo tar_parent = new DirectoryInfo(txtLogfile_Trw.Text);
                        DirectoryInfo[] arr_dic_child = tar_parent.GetDirectories();

                        string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                        DataTable VHX_log_tbl = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_VAL", filter_str).Clone();
                        DataTable VHX_img_tbl = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_IMAGE", filter_str).Clone();
                        int count_ = 1;
                        if (arr_dic_child.Length > 0)
                        {

                            foreach (DirectoryInfo tar_d in arr_dic_child)
                            {
                                Tracewidth_Data_Process_2(tar_d.FullName, ref VHX_log_tbl);
                                Image_Process(tar_d.FullName, ref VHX_img_tbl);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Folder lưu dữ liệu không đúng định dạng", "Thông báo");
                            return;
                        }    

                        DGV_VHX_Data.DataSource = VHX_log_tbl;
                        DGV_Image_Tracewidth.DataSource = VHX_img_tbl;

                        ((DataGridViewImageColumn)DGV_Image_Tracewidth.Columns["Image_Tracewidth"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                        ((DataGridViewImageColumn)DGV_Image_Tracewidth.Columns["Image_Tracewidth"]).Width = 100;
                        foreach (DataGridViewRow dr in DGV_Image_Tracewidth.Rows)
                        {
                            dr.Height = 70;
                        }

                    }

                    if (DGV_Image_Tracewidth.Rows.Count > 0)
                    {
                        ((DataGridViewImageColumn)DGV_Image_Tracewidth.Columns["Image_Tracewidth"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                        ((DataGridViewImageColumn)DGV_Image_Tracewidth.Columns["Image_Tracewidth"]).Width = 100;
                        foreach (DataGridViewRow dr in DGV_Image_Tracewidth.Rows)
                        {
                            dr.Height = 70;
                        }
                    }
                    if (DGV_VHX_Data.Rows.Count > 0)
                    {
                        check_OOS_Trw(DGV_VHX_Data, DGV_Tracewidth_spec);
                    }
                    // btn_summary_trw.Enabled = false;


                }
                else
                {

                    MessageBox.Show("ItemCode or LotNo not matched", "Warning");

                }
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });

                DGV_Tracewidth_spec.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
                DGV_VHX_Data.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "DATA_LOG_FILE", filter_str);
                DGV_Image_Tracewidth.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_IMAGE", filter_str);

            }

        }

        private void DGV_VHX_Data_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DGV_Image_Tracewidth_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int col_inx = e.ColumnIndex;
            int r_inx = e.RowIndex;
            DataGridViewCell cur_cell = DGV_Image_Tracewidth.CurrentCell;
            if (DGV_Image_Tracewidth.Columns[col_inx].Name.Contains("Image"))
            {
                //byte[] data = (byte[])cur_cell.Value;
                //using (MemoryStream ms = new MemoryStream(data))
                //{
                //    pic_trw.Image = Image.FromStream(ms);
                //}
                string region = DGV_Image_Tracewidth.Rows[r_inx].Cells["Data_For"].Value.ToString();
                string pcs = DGV_Image_Tracewidth.Rows[r_inx].Cells["Pcs_No"].Value.ToString();

                View_detail_Image fr1 = new View_detail_Image();
                fr1.data = (byte[])cur_cell.Value;
                fr1.region_ = region;
                fr1.pcs_ = pcs;
                fr1.choose_sheet_ = "Tracewidth";
                fr1.Show();
            }
        }

        private void txtItemCode_TextChanged_1(object sender, EventArgs e)
        {
            DGV_Graph.DataSource = DGV_Impedance.DataSource = DGV_VHX_Data.DataSource = DGV_VHX_Data.DataSource = DGV_Impedance_Summary.DataSource = DGV_Image_Tracewidth.DataSource = null;
            // pic_imp.Image = null;
            //pic_trw.Image = null;
            DGV_Impedance_Spec.DataSource = DGV_Tracewidth_spec.DataSource = null;



            update_mode = false;
            update_mode_trw = false;

            sel_sample_lst = new List<string> { };
            edited_lst = new SortedDictionary<string, List<string>>();

            txtLogfile_Imp.Text = "";
            txtLogfile_Trw.Text = "";

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_MouseDown(object sender, MouseEventArgs e)
        {
            //string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
            //DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", filter_str);
            //id_logfile = 1;
            //if (dt.Rows.Count > 0)
            //{
            //    if (DGV_Impedance_Spec.Rows.Count == 0)
            //    {
            //        DGV_Impedance_Spec.DataSource = dt;
            //    }

            //    Logfile_Data_Impedance();
            //}
            //else
            //{
            //    MessageBox.Show(new Form { TopMost = true }, "Please, Click SETUP SPEC", "Warning");
            //}
        }

        private void txtLogfile_Imp_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtLogfile_Imp_MouseDown(object sender, MouseEventArgs e)
        {
            //if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtOperator.Text != "")
            //{
            //    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
            //    DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", filter_str);
            //    id_logfile = 1;
            //    if (dt.Rows.Count > 0)
            //    {
            //        if (DGV_Impedance_Spec.Rows.Count == 0)
            //        {
            //            DGV_Impedance_Spec.DataSource = dt;
            //        }

            //        Logfile_Data_Impedance();
            //    }
            //    else
            //    {
            //        MessageBox.Show(new Form { TopMost = true }, "Please, Click SETUP SPEC", "Warning");
            //    }
            //}
            //else
            //{
            //    MessageBox.Show("Please, fill in ItemCode / LotNo / Operator ID", "Warning");
            //}

        }

        private void btn_reset_Click(object sender, EventArgs e)
        {
            DGV_Impedance.DataSource = null;
            DGV_Graph.DataSource = null;
            DGV_Impedance_Summary.DataSource = null;

            // Graph_dt = new DataTable();
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
            Graph_dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", filter_str).Clone();

        }

        private void txtLogfile_Imp_KeyDown(object sender, KeyEventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "")
            {
                if (e.KeyCode == Keys.Enter)
                {
                    if (!update_mode)
                    {
                        string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
                        DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", filter_str);
                        id_logfile = 1;
                        if (dt.Rows.Count > 0)
                        {
                            if (DGV_Impedance_Spec.Rows.Count == 0)
                            {
                                DGV_Impedance_Spec.DataSource = dt;
                            }
                            DataTable dt_ = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "IMPEDANCE_VAL", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                            string[] sel_region = dt_.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();

                            if (sel_region.Length == DGV_Impedance_Spec.Rows.Count)
                            {
                                MessageBox.Show("Data for IMPEDANCE is exist. Please click LOAD DATA and mouse down table Summary Data to update", "Waring");

                            }
                            else
                            {
                                if (DGV_Impedance_Summary.Rows.Count == 0)
                                {
                                    string filter_str1 = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                                    DGV_Impedance_Summary.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_VAL", filter_str1);
                                    DGV_Graph.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", filter_str1);

                                }
                                if (DGV_Graph.Rows.Count > 0)
                                {
                                    ((DataGridViewImageColumn)DGV_Graph.Columns["Image_Graph"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                                    ((DataGridViewImageColumn)DGV_Graph.Columns["Image_Graph"]).Width = 100;
                                    foreach (DataGridViewRow dr in DGV_Graph.Rows)
                                    {
                                        dr.Height = 70;
                                    }
                                    Graph_dt = (DataTable)DGV_Graph.DataSource;

                                }

                                btnSummary.Enabled = true;

                                //DataTable dt_ = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "IMPEDANCE_VAL", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                                //if (dt_.Rows.Count > 0)
                                //{
                                //    MessageBox.Show("Data for IMPEDANCE is exist. Please click LOAD DATA and mouse down table Summary Data to update", "Waring");

                                //}
                                //else
                                //{
                                //    btnSummary.Enabled = true;
                                //    // Logfile_Data_Impedance();
                                //    //Impedance_Data_Process_0(txtLogfile_Imp.Text);

                                //}
                            }
                        }
                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Please, Click SETUP SPEC", "Warning");
                        }

                    }
                    else
                    {
                        btnSummary.Enabled = true;
                    }
                }
            }
            else
            {
                MessageBox.Show("Please, fill in ItemCode / LotNo / Operator ID and select Coupon/Patern", "Warning");
            }

        }
        public void Impedance_update(int region, DataTable logfile_data, DataGridView src_DGV)
        {
            DataTable dt_imp_summary = (DataTable)src_DGV.DataSource;

            DataView dv = dt_imp_summary.AsDataView();
            string filter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { region.ToString() });
            dv.RowFilter = filter;
            int min_pcs_count = new int[] { dv.Count, logfile_data.Rows.Count }.Min();
            for (int inx = 0; inx < min_pcs_count; inx++)
            {
                DataRow dr = dv[inx].Row;
                int r_x = dt_imp_summary.Rows.IndexOf(dr);
                src_DGV.Rows[r_x].Cells["Data"].Value = logfile_data.Rows[inx]["Data"].ToString();
                src_DGV.Rows[r_x].Cells[""].Value = logfile_data.Rows[inx]["Data"].ToString();

                // DGV_Zone_Data.Rows[inx].Cells["PCS_No" + item].Style.BackColor = Color.Red;

            }



        }

        private void DGV_Impedance_Summary_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (txtItemCode.Text != "" && txtLotNo.Text != "")
                {
                    DataTable logfile_before = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_VAL", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                    if (logfile_before.Rows.Count > 0)
                    {
                        //  admin_mode = true;
                        if (admin_mode)
                        {

                            if (cb_Impedance.SelectedIndex != -1)
                            {

                                if (DGV_Impedance_Summary.SelectedCells.Count > 0)
                                {
                                    foreach (DataGridViewCell cells in DGV_Impedance_Summary.SelectedCells)
                                    {
                                        int r_inx = cells.RowIndex;
                                        //string col_name = DGV_Impedance_Summary.Columns[c_inx].Name;
                                        region_imp_update = int.Parse(DGV_Impedance_Summary.Rows[r_inx].Cells["Region"].Value.ToString());

                                        break;

                                    }

                                    if (MessageBox.Show(new Form { TopMost = true }, "Do you want to update data for Impedance " + region_imp_update.ToString() + " ?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                                    {
                                        MessageBox.Show("Please, \n 1. Select Coupon/Patern \n 2. Double click or paste the path in Logfile Location \n 3. Click the icon Load data", "Warning");
                                        update_mode = true;
                                        txtLogfile_Imp.Text = "";

                                        //foreach (DataGridViewCell cells in DGV_Impedance_Summary.SelectedCells)
                                        //{
                                        //    int r_inx = cells.RowIndex;
                                        //    //string col_name = DGV_Impedance_Summary.Columns[c_inx].Name;
                                        //    region_imp_update = int.Parse(DGV_Impedance_Summary.Rows[r_inx].Cells["Region"].Value.ToString());
                                        //    break;

                                        //}
                                        Graph_dt = (DataTable)DGV_Graph.DataSource;

                                    }

                                }

                            }
                            else
                            {
                                MessageBox.Show("Please, select Coupon/Patern", "Warning");
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

        private void textBox2_MouseDown(object sender, MouseEventArgs e)
        {

            //if (txtItemCode.Text != "" && txtLotNo.Text != "" && cb_tracewidth.SelectedIndex > -1)
            //{

            //        Logfile_Data_Tracewidth();

            //}
            //else
            //{
            //    MessageBox.Show(new Form { TopMost = true }, "Please, input ItemCode / Lotno / Data for", "Warning");
            //}
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
        public void Tracewidth_Data_Process_Update(string src_path, List<string> sel_pcs_lst)
        {
            if (src_path != "")
            {
                if (DGV_VHX_Data.DataSource != null)
                {

                    DataTable dt_summary = (DataTable)DGV_VHX_Data.DataSource;
                    DataView dv_summary = dt_summary.AsDataView();
                    DataView dv_data_update = ((DataTable)DGV_VHX_Data.DataSource).AsDataView();
                    DataView dv_image = Image_dt.AsDataView();


                    foreach (var item in sel_pcs_lst)
                    {
                        string filter = TDMK_Code.filter_str(new string[] { "Pcs_No", "Data_For" }, new string[] { item.Split('+')[1], item.Split('+')[0] });
                        dv_summary.RowFilter = filter;
                        dv_data_update.RowFilter = filter;


                        for (int inx = 0; inx < dv_data_update.Count; inx++)
                        {
                            DataRow dr = dv_summary[inx].Row;
                            int r_x = dt_summary.Rows.IndexOf(dr);
                            DGV_VHX_Data.Rows[r_x].Cells["Data"].Value = dv_data_update[inx]["Data"];
                            DGV_VHX_Data.Rows[r_x].Cells["Remark"].Value = dv_data_update[inx]["Remark"];

                        }

                        dv_image.RowFilter = filter;


                        for (int inx = 0; inx < dv_image.Count; inx++)
                        {
                            DataRow dr = dv_image[inx].Row;
                            int r_x = Image_dt.Rows.IndexOf(dr);

                            foreach (DataGridViewRow dr_i in DGV_Image_Tracewidth.Rows)
                            {
                                if (dr_i.Cells["Pcs_No"].Value.ToString() == item.Split('+')[1] && dr_i.Cells["Data_For"].Value.ToString() == item.Split('+')[0])
                                {
                                    Image_dt.Rows[r_x]["Image_Tracewidth"] = dr_i.Cells["Image_Tracewidth"].Value;
                                    Image_dt.Rows[r_x]["Remark"] = dr_i.Cells["Remark"].Value;
                                    break;
                                }
                            }



                        }


                    }

                    DGV_Image_Tracewidth.DataSource = Image_dt;

                    ((DataGridViewImageColumn)DGV_Image_Tracewidth.Columns["Image_Tracewidth"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)DGV_Image_Tracewidth.Columns["Image_Tracewidth"]).Width = 100;
                    foreach (DataGridViewRow dr in DGV_Image_Tracewidth.Rows)
                    {
                        dr.Height = 70;
                    }
                    MessageBox.Show("Replace data completed!", "Information");

                    if (edited_lst.Keys.ToList().IndexOf("Tracewidth") == -1)
                    {
                        edited_lst.Add("Tracewidth", sel_pcs_lst);
                    }
                    else if (edited_lst.Values.ToList().IndexOf(sel_pcs_lst) == -1)
                    {
                        edited_lst.Remove("Tracewidth");
                        edited_lst.Add("Tracewidth", sel_pcs_lst);
                    }
                }





            }
        }
        private void SetValue(String value)
        {
            this.txtLogfile_Trw.Text = value;
        }

        private void DGV_VHX_Data_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (txtItemCode.Text != "" && txtLotNo.Text != "")
                {
                    DataTable logfile_before = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_VAL", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                    if (logfile_before.Rows.Count > 0)
                    {
                        // admin_mode = true;
                        if (admin_mode)
                        {

                            SortedDictionary<int, List<string>> sel_cells_lst = new SortedDictionary<int, List<string>>();

                            //sel_sample_lst = new List<string> { };
                            foreach (DataGridViewCell cells in DGV_VHX_Data.SelectedCells)
                            {
                                int r_inx = cells.RowIndex;

                                string sel_trw = DGV_VHX_Data.Rows[r_inx].Cells["Data_For"].Value.ToString();
                                string sel_sample = DGV_VHX_Data.Rows[r_inx].Cells["Pcs_No"].Value.ToString();

                                if (sel_sample_lst.IndexOf(sel_trw + "+" + sel_sample) == -1)
                                {
                                    sel_sample_lst.Add(sel_trw + "+" + sel_sample);
                                } 

                            }
                            if (sel_sample_lst.Count > 0)
                            {
                                FolderBrowserDialog f_open = new FolderBrowserDialog();
                                f_open.SelectedPath = System.Windows.Forms.Application.StartupPath;

                                if (f_open.ShowDialog() == DialogResult.OK)
                                {

                                    txtLogfile_Trw.Text = f_open.SelectedPath;
                                    if (txtLogfile_Trw.Text != "")
                                    {
                                        ImageDetail fr1 = new ImageDetail(SetValue);
                                        fr1.startpath = txtLogfile_Trw.Text;
                                        fr1.Show();
                                        update_mode_trw = true;

                                    }
                                    if (txtLogfile_Trw.Text != "")
                                    {
                                        btn_summary_trw.Enabled = true;
                                    }
                                     
                                    //**********************************
                                    //txtLogfile_Trw.Text = f_open.SelectedPath;
                                    //Image_dt = (DataTable)DGV_Image_Tracewidth.DataSource;
                                    //Tracewidth_Update(txtLogfile_Trw.Text, sel_sample_lst[0].Split('+')[0]);
                                    //// Impedance_Data_Process_0(txtLogfile_Imp.Text);

                                    //update_mode_trw = true;
                                    //btn_summary_trw.Enabled = true;

                                    //**********************************
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

        private void cbProcess_Edit_SelectedIndexChanged_1(object sender, EventArgs e)
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
                        dr.Height = 140;
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
            try
            {
                DataTable dt_img = src_dt.AsEnumerable().Where(x => x.Field<string>("Region") == "Image_data").CopyToDataTable();
                img_dt = dt_img.AsDataView().ToTable(false, new string[] { "Zone", "PCS_No", "Before_Data", "After_Data", "Time_Update", "Operator", "Depart", "Remark" });
            }
            catch
            {

            }
            DataTable dt_data = src_dt.AsEnumerable().Where(x => x.Field<string>("Region") != "Image_data").CopyToDataTable();
            //img_dt = dt_img.AsDataView().ToTable(false, new string[] { "Time_Update", "PCS_No", "Before_Data", "After_Data" });
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

        private void btn_reset_trw_Click(object sender, EventArgs e)
        {
            DGV_VHX_Data.DataSource = null;
            DGV_Image_Tracewidth.DataSource = null;
            DGV_VHX_Data.DataSource = null;
            btn_reset.Enabled = false;
            // Graph_dt = new DataTable();
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
            Image_dt = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_IMAGE", filter_str).Clone();
            sel_sample_lst = new List<string> { };
        }

        private void txtLogfile_Trw_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtLogfile_Trw_MouseDoubleClick(object sender, MouseEventArgs e)
        {

            if (txtItemCode.Text != "" && txtLotNo.Text != "")
            {
                //if (!update_mode_trw)
                //{
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
                DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", filter_str);
                id_logfile = 1;
                if (dt.Rows.Count > 0)
                {
                    if (DGV_Tracewidth_spec.Rows.Count == 0)
                    {
                        DGV_Tracewidth_spec.DataSource = dt;
                    }

                    DataTable dt_ = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "TRACEWIDTH_VAL", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));

                    if (dt_.Rows.Count > 0)
                    {
                        MessageBox.Show("Data is exist. Please click LOAD DATA to update", "Waring");
                    }
                    else
                    {
                        FolderBrowserDialog f_open = new FolderBrowserDialog();
                        f_open.SelectedPath = selected_path;
                        if (f_open.ShowDialog() == DialogResult.OK)
                        {
                            txtLogfile_Trw.Text = f_open.SelectedPath;
                            selected_path = f_open.SelectedPath;
                            btn_summary_trw.Enabled = true;
                        }

                        //Logfile_Data_Tracewidth();
                    }
                }
                else
                {

                    MessageBox.Show(new Form { TopMost = true }, "Please, Click SETUP SPEC", "Warning");

                }
                //}
                //else
                //{

                //}




            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Please, input ItemCode / Lotno / Data for", "Warning");
            }


        }

        private void txtLogfile_Imp_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtOperator.Text != "" && cb_Impedance.SelectedIndex != -1)
            {
                if (!update_mode)
                {

                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
                    DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", filter_str);
                    id_logfile = 1;
                    if (dt.Rows.Count > 0)
                    {
                        if (DGV_Impedance_Spec.Rows.Count == 0)
                        {
                            DGV_Impedance_Spec.DataSource = dt;
                        }
                        DataTable dt_ = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "IMPEDANCE_VAL", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                        string[] sel_region = dt_.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();

                        if (sel_region.Length == DGV_Impedance_Spec.Rows.Count)
                        {
                            MessageBox.Show("Data for IMPEDANCE is exist. Please click LOAD DATA and mouse down table Summary Data to update", "Waring");

                        }
                        else
                        {
                            if (DGV_Impedance_Summary.Rows.Count == 0)
                            {
                                string filter_str1 = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                                DGV_Impedance_Summary.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_VAL", filter_str1);
                                DGV_Graph.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", filter_str1);

                            }
                            if (DGV_Graph.Rows.Count > 0)
                            {
                                ((DataGridViewImageColumn)DGV_Graph.Columns["Image_Graph"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                                ((DataGridViewImageColumn)DGV_Graph.Columns["Image_Graph"]).Width = 100;
                                foreach (DataGridViewRow dr in DGV_Graph.Rows)
                                {
                                    dr.Height = 70;
                                }
                                Graph_dt = (DataTable)DGV_Graph.DataSource;

                            }



                            FolderBrowserDialog f_open = new FolderBrowserDialog();
                            f_open.SelectedPath = selected_path;
                            if (f_open.ShowDialog() == DialogResult.OK)
                            {

                                txtLogfile_Imp.Text = f_open.SelectedPath;
                                //Impedance_Data_Process_0(f_open.SelectedPath);                           
                                selected_path = f_open.SelectedPath;
                                btnSummary.Enabled = true;

                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Please, Click SETUP SPEC", "Warning");
                    }
                }
                else
                {
                    FolderBrowserDialog f_open = new FolderBrowserDialog();
                    f_open.SelectedPath = selected_path;
                    if (f_open.ShowDialog() == DialogResult.OK)
                    {

                        txtLogfile_Imp.Text = f_open.SelectedPath;
                        selected_path = f_open.SelectedPath;
                        btnSummary.Enabled = true;

                    }
                }
            }
            else
            {
                MessageBox.Show("Please, fill in ItemCode / LotNo / Operator ID and Select Coupon/Patern", "Warning");
            }
        }

        private void txtLogfile_Trw_KeyDown(object sender, KeyEventArgs e)
        {

            if (txtItemCode.Text != "" && txtLotNo.Text != "")
            {
                if (e.KeyCode == Keys.Enter)
                {
                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
                    DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", filter_str);

                    if (dt.Rows.Count > 0)
                    {
                        if (DGV_Tracewidth_spec.Rows.Count == 0)
                        {
                            DGV_Tracewidth_spec.DataSource = dt;
                        }

                        DataTable dt_ = TDMK_Code.Datatable_Filter(myVar.mysqlcon, "TRACEWIDTH_VAL", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));

                        if (dt_.Rows.Count > 0)
                        {
                            MessageBox.Show("Data is exist. Please click LOAD DATA to update", "Waring");
                        }
                        else
                        {
                            //if (DGV_VHX_Data.Rows.Count == 0)
                            //{
                            //    string filter_str1 = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                            //    DGV_VHX_Data.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_VAL", filter_str1);
                            //    DGV_Image_Tracewidth.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_IMAGE", filter_str1);

                            //}

                            //if (DGV_Image_Tracewidth.Rows.Count > 0)
                            //{
                            //    ((DataGridViewImageColumn)DGV_Image_Tracewidth.Columns["Image_Tracewidth"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                            //    ((DataGridViewImageColumn)DGV_Image_Tracewidth.Columns["Image_Tracewidth"]).Width = 100;
                            //    foreach (DataGridViewRow dr in DGV_Image_Tracewidth.Rows)
                            //    {
                            //        dr.Height = 70;
                            //    }
                            //    Image_dt = (DataTable)DGV_Image_Tracewidth.DataSource;

                            //}

                            if (txtLogfile_Trw.Text != "")
                            {
                                ImageDetail fr1 = new ImageDetail(SetValue);
                                fr1.startpath = txtLogfile_Trw.Text;
                                fr1.Show();
                                btn_summary_trw.Enabled = true;
                            }

                        }


                    }
                    else
                    {

                        MessageBox.Show(new Form { TopMost = true }, "Please, Click SETUP SPEC", "Warning");

                    }
                }

            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Please, input ItemCode / Lotno / Data for", "Warning");
            }

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

            if (myCode.checkDBNull(before_data) != "" && myCode.checkDBNull(after_data) != "")
            {

                Image_history fr1 = new Image_history();
                fr1.data_image_before_ = before_data;
                fr1.data_image_after_ = after_data;
                fr1.Show();
            }
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

        private void txtLotNo_TextChanged(object sender, EventArgs e)
        {
            DGV_Graph.DataSource = DGV_Impedance.DataSource = DGV_VHX_Data.DataSource = DGV_VHX_Data.DataSource = DGV_Impedance_Summary.DataSource = DGV_Image_Tracewidth.DataSource = null;
            //pic_imp.Image = null;
            //pic_trw.Image = null;
            //cb_Impedance.Text = cb_tracewidth.Text = "";


            update_mode = false;
            update_mode_trw = false;

            sel_sample_lst = new List<string> { };
            edited_lst = new SortedDictionary<string, List<string>>();

            txtLogfile_Imp.Text = "";
            txtLogfile_Trw.Text = "";
        }

        private void btnSummary_KeyDown(object sender, KeyEventArgs e)
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

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void btn_loadspec1_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
                DataTable dt_impedance_spec = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", filter_str);
                DataTable dt_tracewidth_spec = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", filter_str);
                if (dt_impedance_spec.Rows.Count == 0 && dt_tracewidth_spec.Rows.Count == 0)
                {
                    MessageBox.Show(new Form { TopMost = true }, "Spec not installed", "Warning");
                }
                else
                {
                    DGV_Impedance_Spec.DataSource = dt_impedance_spec;
                    DGV_Tracewidth_spec.DataSource = dt_tracewidth_spec;
                }
            }
        }

        private void btn_loadspec2_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "")
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
                DataTable dt_impedance_spec = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", filter_str);
                DataTable dt_tracewidth_spec = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", filter_str);
                if (dt_impedance_spec.Rows.Count == 0 && dt_tracewidth_spec.Rows.Count == 0)
                {
                    MessageBox.Show(new Form { TopMost = true }, "Spec not installed", "Warning");
                }
                else
                {
                    DGV_Impedance_Spec.DataSource = dt_impedance_spec;
                    DGV_Tracewidth_spec.DataSource = dt_tracewidth_spec;
                }
            }
        }
    }
}
