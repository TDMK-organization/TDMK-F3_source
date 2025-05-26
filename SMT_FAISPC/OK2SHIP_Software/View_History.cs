using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
//using TDMK_SEEV_DLL;
using TDMK_SQL;
using System.Data.SqlClient;
using System.Security.Policy;
using System.Xml;
using System.Diagnostics.Eventing.Reader;
using static System.Windows.Forms.AxHost;
using System.Diagnostics;
using System.Collections;
using System.Linq.Expressions;
using OK2SHIP_Lib;

namespace OK2SHIP_Software
{
    public partial class View_History : Form
    {
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public SqlConnection sqlcon = null;
        TDMK_OK2SHIP myCode = new TDMK_OK2SHIP();

        public static string server_name = "";
        public static string server_acc = "";
        public static string server_pass = "";
        public static string sel_DB = "OK2SHIP_Period2";
        DataTable tbl_data = new DataTable();
        DataTable tbl_data_image = new DataTable();
        DataTable tbl_spec = new DataTable();   

        DataTable tbl_history = new DataTable();
        DataTable tbl_history_image = new DataTable();
        string process_s = "";
        Byte[] img_null = null;


        public string itemcodevh = "";
        public string lotnovh = "";
        public string sheet = "";


        public string itemcodevh_
        {
            get { return itemcodevh; }
            set { itemcodevh = value; }
        }

        public string lotnovh_
        {
            get { return lotnovh; }
            set { lotnovh = value; }
        }
        public string sheet_
        {
            get { return sheet; }
            set { sheet = value; }
        }


        public View_History()
        {
            InitializeComponent();
        }

        public void Get_all_zone_Impedance(string sheet, TreeNode main_node)
        {
            DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { itemcodevh, lotnovh, "Stackup" }));
            string[] arr_zone = dt.AsEnumerable().Select(x => x.Field<string>("Zone")).Distinct().ToArray();

            foreach (string item in arr_zone)
            {
                TreeNode Subnode = new TreeNode();
                Subnode.Text = item;
                main_node.Nodes.Add(Subnode);
            }


        }

        public void get_all_pcs(TreeNode mainnode)
        {
            
        }
        public bool IsTheSameCellValue(int column, int row, DataGridView dgv)
        {
            DataGridViewCell cell1 = dgv[column, row];
            DataGridViewCell cell2 = dgv[column, row - 1];
            if (cell1.Value == null || cell2.Value == null)
            {
                return false;
            }
            return cell1.Value.ToString() == cell2.Value.ToString();
        }

        private void View_History_Load(object sender, EventArgs e)
        {
            sqlcon = initial_data(sel_DB, true);
            string app_path = Application.StartupPath;
            string img_path = Path.Combine(app_path, "Img_null", "img_null.jpg");

            var sel_img = Bitmap.FromFile(img_path);
            ImageConverter imgcon = new ImageConverter();
            img_null = (byte[])imgcon.ConvertTo(sel_img, typeof(byte[]));

            dgv_edit_data.DataSource = DGV_data_image.DataSource = null;


            cb_zone.Items.Clear();
            cb_zone.Text = "ALL";
            cb_pcs.Items.Clear();
            cb_pcs.Text = "ALL";
            dgv_edit_data.Columns.Clear();
            dgv_edit_data.Rows.Clear();
            DGV_data_image.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            lbl_image.Text = "Image/Graph";
            DGV_data_image.Columns.Clear();
            DGV_data_image.Rows.Clear();

            if (itemcodevh != "" && lotnovh != "" && sheet != "")
            {
                if (sheet == "Stack-up")
                {
                    DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
                    if (dt_data_summary.Rows.Count > 0)
                    {

                        dgv_edit_data.AllowUserToAddRows = true;
                        DataTable dt_edit_history = new DataTable();
                        dt_edit_history.Columns.Add("ID");
                        dt_edit_history.Columns.Add("Zone");
                        dt_edit_history.Columns.Add("PCS");
                        dt_edit_history.Columns.Add("Before");
                        int max_count_edit = 0;
                        Dictionary<string, Dictionary<int, Dictionary<string, List<string>>>> table_edit = Get_table_edit_stackup(ref max_count_edit);
                        for (int i = 0; i < max_count_edit; i++)
                        {
                            dt_edit_history.Columns.Add("After_" + (i + 1).ToString());
                        }


                        Dictionary<int, Dictionary<string, List<string>>> table_zone_edit = new Dictionary<int, Dictionary<string, List<string>>> { };
                        List<int> count_edit = new List<int> { };
                        int id = 1;
                        foreach (var tbl_zone_edit in table_edit)
                        {
                            string zone = tbl_zone_edit.Key;
                            foreach (var tbl_pcs_edit in tbl_zone_edit.Value)
                            {
                                int r_pcs = tbl_pcs_edit.Value["Before"].Count;
                                int pcs = tbl_pcs_edit.Key;

                                for (int i = 0; i < r_pcs; i++)
                                {
                                    DataRow dr = dt_edit_history.NewRow();
                                    int idx = 3;

                                    dr[0] = id;
                                    dr[1] = zone;
                                    dr[2] = pcs;

                                    foreach (var lst_data_pcs in tbl_pcs_edit.Value)
                                    {
                                        dr[idx] = lst_data_pcs.Value[i];
                                        idx++;
                                    }

                                    dt_edit_history.Rows.Add(dr);
                                    id++;
                                }

                                DataRow dr_sum = dt_edit_history.NewRow();
                                dr_sum[0] = "Total";
                                dr_sum[1] = zone;
                                dr_sum[2] = pcs;
                                int in_col = 3;
                                foreach (var lst_data_pcs in tbl_pcs_edit.Value)
                                {
                                    double sum = 0;
                                    foreach (string data in lst_data_pcs.Value)
                                    {
                                        sum += double.Parse(data);
                                    }
                                    dr_sum[in_col] = sum;
                                    in_col++;
                                }
                                dt_edit_history.Rows.Add(dr_sum);

                            }
                        }
                        dgv_edit_data.DataSource = dt_edit_history;
                        for (int i = 0; i < dgv_edit_data.Rows.Count - 1; i++)
                        {
                            if (dgv_edit_data.Rows[i].Cells["ID"].Value.ToString() == "Total")
                            {
                                for (int j = 0; j < dgv_edit_data.ColumnCount; j++)
                                {
                                    dgv_edit_data.Rows[i].Cells[j].Style.BackColor = Color.GreenYellow;
                                }
                            }
                        }

                        tbl_data = dt_edit_history;
                        string[] arr_zone = tbl_data.AsEnumerable().Select(x => x.Field<string>("Zone")).Distinct().ToArray();
                        foreach (string item in arr_zone)
                        {
                            if(item != "")
                            {
                                cb_zone.Items.Add(item);
                            }
                            
                        }

                        check_spec_stackup_detail();
                        fill_table_image();
                    }
                    else
                    {
                        MessageBox.Show("No data", "Warning");
                    }

                }
                if (sheet == "BVH & PTH")
                {
                    dgv_edit_data.AllowUserToAddRows = true;

                    cb_zone.Items.Add("BVH_with_Bonding_Sheet");
                    cb_zone.Items.Add("BVH_without_Bonding_Sheet");
                    cb_zone.Items.Add("Plated_Through_Hole");

                }
                if (sheet == "Impedance")
                {
                    dgv_edit_data.AllowUserToAddRows = true;
                    cb_zone.Items.Add("Impedance");
                    cb_zone.Items.Add("Tracewidth");
                }

                if (sheet == "Solder Mask")
                {
                    DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "SOLDERMASK_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
                    if (dt_data_summary.Rows.Count > 0)
                    {

                        dgv_edit_data.AllowUserToAddRows = false;
                        fill_table_image_soldermask();
                        DataTable dt_inf = new DataTable();
                        dt_inf.Columns.Add("Edit");
                        dt_inf.Columns.Add("Depart");
                        dt_inf.Columns.Add("Operator");
                        dt_inf.Columns.Add("Logfile Location");
                        dt_inf.Columns.Add("Time_Update");
                        for (int i = 2; i < dgv_edit_data.Columns.Count; i++)
                        {

                            List<DataTable> tbl_time = new List<DataTable> { };
                            Get_ListTable_datetime(-1, tbl_history, new string[] { "Time_Update" }, ref tbl_time, "Remark");
                            string colname_select = dgv_edit_data.Columns[i].Name;
                            if (myCode.checkDBNull(dgv_edit_data.Rows[0].Cells[i].Value) != "")
                            {
                                if (colname_select.Contains("After_"))
                                {
                                    int time_edit = int.Parse(colname_select.Replace("After_", ""));
                                    if (time_edit <= tbl_time.Count)
                                    {
                                        DataTable tbl_select = tbl_time[time_edit - 1];
                                        string depart = tbl_select.Rows[0]["Depart"].ToString();
                                        string time_update = tbl_select.Rows[0]["Time_Update"].ToString();
                                        string Operator = tbl_select.Rows[0]["Operator"].ToString();
                                        string remark = tbl_select.Rows[0]["Remark"].ToString();
                                        DataRow dr = dt_inf.NewRow();
                                        dr[0] = colname_select;
                                        dr[1] = depart;
                                        dr[2] = Operator;
                                        dr[3] = remark;
                                        dr[4] = time_update;

                                        dt_inf.Rows.Add(dr);
                                    }
                                }
                            }

                        }

                        DGV_data_image.DataSource = dt_inf;
                        DGV_data_image.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                        lbl_image.Text = "Edit Information";
                    }
                    else
                    {
                        MessageBox.Show("No data", "Warning");
                    }
                }

                if (sheet == "CQRA - Thermal stress")
                {
                    DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "THERMAL_STRESS_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
                    if (dt_data_summary.Rows.Count > 0)
                    {


                        dgv_edit_data.AllowUserToAddRows = false;
                        cb_zone.Items.Add("BVH");
                        cb_zone.Items.Add("PTH");
                        fill_table_image_thermalstress();


                        DataTable dt_inf = new DataTable();
                        dt_inf.Columns.Add("Edit");
                        dt_inf.Columns.Add("Depart");
                        dt_inf.Columns.Add("Operator");
                        dt_inf.Columns.Add("Logfile Location");
                        dt_inf.Columns.Add("Time_Update");
                        for (int i = 4; i < dgv_edit_data.Columns.Count; i++)
                        {

                            List<DataTable> tbl_time = new List<DataTable> { };
                            Get_ListTable_datetime(-1, tbl_history, new string[] { "Time_Update" }, ref tbl_time, "Remark");
                            string colname_select = dgv_edit_data.Columns[i].Name;
                            if (myCode.checkDBNull(dgv_edit_data.Rows[0].Cells[i].Value) != "")
                            {
                                if (colname_select.Contains("After_"))
                                {
                                    int time_edit = int.Parse(colname_select.Replace("After_", ""));
                                    if (time_edit <= tbl_time.Count)
                                    {
                                        DataTable tbl_select = tbl_time[time_edit - 1];
                                        string depart = tbl_select.Rows[0]["Depart"].ToString();
                                        string time_update = tbl_select.Rows[0]["Time_Update"].ToString();
                                        string Operator = tbl_select.Rows[0]["Operator"].ToString();
                                        string remark = tbl_select.Rows[0]["Remark"].ToString();
                                        DataRow dr = dt_inf.NewRow();
                                        dr[0] = colname_select;
                                        dr[1] = depart;
                                        dr[2] = Operator;
                                        dr[3] = remark;
                                        dr[4] = time_update;

                                        dt_inf.Rows.Add(dr);
                                    }
                                }
                            }

                        }

                        DGV_data_image.DataSource = dt_inf;
                        DGV_data_image.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                        lbl_image.Text = "Edit Information";
                    }
                    else
                    {
                        MessageBox.Show("No data", "Warning");
                    }
                }

                if (sheet == "ACF")
                {
                    DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "ACF_FLATNESS", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
                    if (dt_data_summary.Rows.Count > 0)
                    {


                        dgv_edit_data.AllowUserToAddRows = true;
                        DataTable dt_edit_history = new DataTable();
                        dt_edit_history.Columns.Add("ID");
                        dt_edit_history.Columns.Add("Sample");
                        dt_edit_history.Columns.Add("Point");
                        dt_edit_history.Columns.Add("Before");
                        int max_count_edit = 0;
                        Dictionary<string, Dictionary<string, Dictionary<string, string>>> table_edit = Get_table_edit_ACFFlatness(ref max_count_edit);
                        for (int i = 0; i < max_count_edit; i++)
                        {
                            dt_edit_history.Columns.Add("After_" + (i + 1).ToString());
                        }


                        Dictionary<string, Dictionary<string, string>> table_zone_edit = new Dictionary<string, Dictionary<string, string>> { };
                        List<int> count_edit = new List<int> { };
                        int id = 1;
                        foreach (var tbl_zone_edit in table_edit)
                        {
                            string zone = tbl_zone_edit.Key;
                            foreach (var tbl_pcs_edit in tbl_zone_edit.Value)
                            {
                                string pcs = tbl_pcs_edit.Key;

                                DataRow dr = dt_edit_history.NewRow();
                                int idx = 3;

                                dr[0] = id;
                                dr[1] = zone;
                                dr[2] = pcs;

                                foreach (var lst_data_pcs in tbl_pcs_edit.Value)
                                {
                                    dr[idx] = lst_data_pcs.Value;
                                    idx++;
                                }

                                dt_edit_history.Rows.Add(dr);
                                id++;
                            }
                        }
                        dgv_edit_data.DataSource = dt_edit_history;
                        tbl_data = dt_edit_history;
                        string[] arr_zone = tbl_data.AsEnumerable().Select(x => x.Field<string>("Sample")).Distinct().ToArray();
                        foreach (string item in arr_zone)
                        {
                            cb_zone.Items.Add(item);
                        }
                    }
                    else
                    {
                        MessageBox.Show("No data", "Warning");
                    }
                }

                if (sheet == "CQRA - Moisture absorption")
                {
                    DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_MOISTURE_ABSORPTION", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
                    if (dt_data_summary.Rows.Count > 0)
                    {


                        dgv_edit_data.AllowUserToAddRows = true;
                        DataTable dt_edit_history = new DataTable();
                        dt_edit_history.Columns.Add("ID");
                        dt_edit_history.Columns.Add("PCS");
                        dt_edit_history.Columns.Add("Zone");
                        dt_edit_history.Columns.Add("Before");
                        int max_count_edit = 0;
                        Dictionary<string, Dictionary<string, Dictionary<string, string>>> table_edit = Get_table_edit_MOISTURE(ref max_count_edit);

                        for (int i = 0; i < max_count_edit; i++)
                        {
                            dt_edit_history.Columns.Add("After_" + (i + 1).ToString());
                        }


                        Dictionary<string, Dictionary<string, string>> table_zone_edit = new Dictionary<string, Dictionary<string, string>> { };
                        List<int> count_edit = new List<int> { };
                        int id = 1;
                        foreach (var tbl_zone_edit in table_edit)
                        {
                            string zone = tbl_zone_edit.Key;
                            foreach (var tbl_pcs_edit in tbl_zone_edit.Value)
                            {
                                string pcs = tbl_pcs_edit.Key;

                                DataRow dr = dt_edit_history.NewRow();
                                int idx = 3;
                                dr[0] = id;
                                dr[1] = zone;
                                dr[2] = pcs;

                                foreach (var lst_data_pcs in tbl_pcs_edit.Value)
                                {
                                    dr[idx] = lst_data_pcs.Value;
                                    idx++;
                                }

                                dt_edit_history.Rows.Add(dr);
                                id++;

                            }
                        }
                        dgv_edit_data.DataSource = dt_edit_history;
                        tbl_data = dt_edit_history;
                        //string[] arr_zone = tbl_data.AsEnumerable().Select(x => x.Field<string>("Pcs")).Distinct().ToArray();
                        //foreach (string item in arr_zone)
                        //{
                        //    cb_zone.Items.Add(item);
                        //}
                    }
                    else
                    {
                        MessageBox.Show("No data", "Warning");
                    }

                }

                if (sheet == "CQRA - bHast")
                {
                    dgv_edit_data.AllowUserToAddRows = true;
                    DataTable bHast_history = TDMK_Code.Datatable_Filter(sqlcon, "BHAST_EDIT_HISTORY", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
                    tbl_data = bHast_history;
                    string[] arr_remark = bHast_history.AsEnumerable().Select(x => x.Field<string>("Remark")).Distinct().ToArray();


                    cb_zone.Items.Clear();
                    cb_zone.Items.Add("Before");
                    for (int i = 0; i < arr_remark.Length - 1; i++)
                    {
                        cb_zone.Items.Add("After_" + (i + 1).ToString());
                        //cb_zone.Items.Add(arr_remark[i]);
                    }


                }
            }
 
        }

        public SqlConnection initial_data(string DB_name, bool sa_en)
        {
            SqlConnection _sqlcon_OK2SHIP;
            string app_path = Application.StartupPath;
            string config_file = Path.Combine(app_path, "Config", "config.txt");
            string[] my_config = myCode.read_config_arr(config_file);

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
            sqlcon = _sqlcon_OK2SHIP;
            return _sqlcon_OK2SHIP;
        }
        public Dictionary<int, Dictionary<string, List<byte[]>>> Get_table_edit_BVH_PTH_Image(string process, ref int max_count_edit)
        {
            DataTable dt_history = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { itemcodevh, lotnovh, process }));
            DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, process.ToUpper(), TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
            DataTable dt_data = new DataTable();
            DataTable dt_img = new DataTable();

            Split_Data_Image_BVH_PTH(dt_history, ref dt_data, ref dt_img); 

            List<DataTable> tbl_pcs = new List<DataTable>();
            Get_ListTable(-1, dt_data_summary, new string[] { "Pcs_No" }, ref tbl_pcs, "Remark");

            Dictionary<int, Dictionary<string, List<byte[]>>> dic_pcs = new Dictionary<int, Dictionary<string, List<byte[]>>> { };
            foreach (DataTable dt in tbl_pcs)
            {
                try
                {

                    Dictionary<string, List<byte[]>> src_edit = new Dictionary<string, List<byte[]>> { };
                    int pcs = int.Parse(dt.Rows[0]["Pcs_No"].ToString());

                    DataView dv = dt_img.AsDataView();
                    if (dv.Count > 0)
                    {
                        string filter = TDMK_Code.filter_str(new string[] { "PCS_No" }, new string[] { dt.Rows[0]["PCS_No"].ToString() });
                        dv.RowFilter = filter;
                    }

                    if (dv.Count == 0)
                    {
                         
                        List<byte[]> lst_image = new List<byte[]> { };

                        lst_image.Add((byte[])dt.Rows[0]["Image_data"]);
                        lst_image.Add((byte[])dt.Rows[0]["Image_data_2"]);
                        src_edit.Add("Before", lst_image);
                    }
                    else
                    {
                        List<DataTable> src_pcs_edit = new List<DataTable>();
                        Get_ListTable_datetime(-1, dv.ToTable(), new string[] { "Time_Update" }, ref src_pcs_edit, "Remark");
                        int i = 0;

                        foreach (DataTable dt_edit in src_pcs_edit)
                        {
                            if (i == 0)
                            {
                                //Dictionary<byte[], byte[]> dic_image_before = new Dictionary<byte[], byte[]> { };
                                List<byte[]> lst_image_before = new List<byte[]> { };

                                lst_image_before.Add((byte[])dt_edit.Rows[0]["Before_Data"]);
                                lst_image_before.Add((byte[])dt_edit.Rows[1]["Before_Data"]); 
                                src_edit.Add("Before", lst_image_before);


                                //Dictionary<byte[], byte[]> dic_image_after = new Dictionary<byte[], byte[]> { };
                                List<byte[]> lst_image_after = new List<byte[]> { };
                                lst_image_after.Add((byte[])dt_edit.Rows[0]["After_Data"]);
                                lst_image_after.Add((byte[])dt_edit.Rows[1]["After_Data"]);
                                src_edit.Add("After_1", lst_image_after);
                            }
                            else
                            {
                                // Dictionary<byte[], byte[]> dic_image_ = new Dictionary<byte[], byte[]> { };
                                List<byte[]> lst_image_ = new List<byte[]> { };
                                lst_image_.Add((byte[])dt_edit.Rows[0]["After_Data"]);
                                lst_image_.Add((byte[])dt_edit.Rows[1]["After_Data"]);
                                src_edit.Add("After_" + (i + 1).ToString(), lst_image_);
                            }

                            i++;

                        }
                        if (i > max_count_edit)
                        {
                            max_count_edit = i;
                        }
                    }



                    dic_pcs.Add(pcs, src_edit);

                }
                catch
                {

                }
            }

            return dic_pcs;

        }

        public Dictionary<string, Dictionary<string, Dictionary<string, byte[]>>> Get_table_edit_Impedance_Image(string process, string col_name_region, string col_name_image, string tbl_name, ref int max_count_edit)
        {
            DataTable dt_history = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { itemcodevh, lotnovh, process }));
            DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, tbl_name, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
            DataTable dt_data = new DataTable();
            DataTable dt_img = new DataTable();
            Split_Data_Image_Impedance(dt_history, ref dt_data, ref dt_img);

            List<DataTable> tbl_region = new List<DataTable> { };
            Get_ListTable(-1, dt_data_summary, new string[] { col_name_region }, ref tbl_region, "Remark");
            Dictionary<string, Dictionary<string, Dictionary<string, byte[]>>> dic_region = new Dictionary<string, Dictionary<string, Dictionary<string, byte[]>>> { };

            foreach (DataTable dt in tbl_region)
            {
                Dictionary<string, Dictionary<string, byte[]>> src_edit = new Dictionary<string, Dictionary<string, byte[]>> { };
                string region = dt.Rows[0][col_name_region].ToString();
                if (!region.Contains("Impedance_") && process == "Impedance")
                {
                    region = "Impedance_" + region;
                }

                DataView dv_region = dt_img.AsDataView();
                if (dv_region.Count > 0)
                {
                    string filter1 = TDMK_Code.filter_str(new string[] { "Zone" }, new string[] { region });
                    dv_region.RowFilter = filter1;
                }

                if (dv_region.Count > 0)
                {
                    string[] arr_pcs = dt.AsEnumerable().Select(x => x.Field<string>("PCS_No")).Distinct().ToArray();
                    foreach (string pcs in arr_pcs)
                    {
                        DataView dv = dt_img.AsDataView();
                        string filter = TDMK_Code.filter_str(new string[] { "Zone", "PCS_No" }, new string[] { region, pcs });
                        dv.RowFilter = filter;

                        if (dv.Count > 0)
                        {

                            DataTable dt_time_edit = dv.ToTable();
                            List<DataTable> src_pcs_edit = new List<DataTable>();
                            Get_ListTable_datetime(-1, dt_time_edit, new string[] { "Time_Update" }, ref src_pcs_edit, "PCS_No");
                            int i = 0;
                            Dictionary<string, byte[]> dic_pcs = new Dictionary<string, byte[]> { };
                            foreach (DataTable dt_edit in src_pcs_edit)
                            {
                                if (i == 0)
                                {

                                    dic_pcs.Add("Before", (byte[])dt_edit.Rows[0]["Before_Data"]);

                                    dic_pcs.Add("After", (byte[])dt_edit.Rows[0]["After_Data"]);

                                }
                                else
                                {
                                    dic_pcs.Add("After_" + (i + 1).ToString(), (byte[])dt_edit.Rows[0]["After_Data"]);
                                }
                                i++;
                            }
                            src_edit.Add(pcs, dic_pcs);


                            if (i > max_count_edit)
                            {
                                max_count_edit = i;
                            }
                        }


                    }

                }
                else
                {


                    foreach (DataRow dr in dt.Rows)
                    {
                        Dictionary<string, byte[]> dic_pcs_data = new Dictionary<string, byte[]> { };
                        dic_pcs_data.Add("Before", (byte[])dr[col_name_image]);
                        if (!src_edit.ContainsKey(dr["PCS_No"].ToString()))
                        {
                            src_edit.Add(dr["PCS_No"].ToString(), dic_pcs_data);
                        }

                    }


                }

                dic_region.Add(region, src_edit);

            }

            return dic_region;

        }
        public Dictionary<string, Dictionary<string, Dictionary<string, List<string>>>> Get_table_edit_Tracewidth_data(ref int max_count_edit)
        {
            DataTable dt_history = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { itemcodevh, lotnovh, "Tracewidth" }));
            tbl_history = dt_history;
            DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_VAL", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
            DataTable dt_data = new DataTable();
            DataTable dt_img = new DataTable();
            Split_Data_Image_Impedance(dt_history, ref dt_data, ref dt_img);

            List<DataTable> tbl_region = new List<DataTable> { };
            Get_ListTable(-1, dt_data_summary, new string[] { "Data_For" }, ref tbl_region, "Remark");
            Dictionary<string, Dictionary<string, Dictionary<string, List<string>>>> dic_region = new Dictionary<string, Dictionary<string, Dictionary<string, List<string>>>> { };

            foreach (DataTable dt in tbl_region)
            {
                Dictionary<string, Dictionary<string, List<string>>> src_edit = new Dictionary<string, Dictionary<string, List<string>>> { };
                string region = dt.Rows[0]["Data_For"].ToString();
                DataView dv_region = dt_data.AsDataView();
                if (dv_region.Count > 0)
                {
                    string filter1 = TDMK_Code.filter_str(new string[] { "Zone" }, new string[] { region });
                    dv_region.RowFilter = filter1;
                }


                if (dv_region.Count > 0)
                {
                    string[] arr_pcs = dt.AsEnumerable().Select(x => x.Field<string>("PCS_No")).Distinct().ToArray();
                    foreach (string pcs in arr_pcs)
                    {
                        DataView dv = dt_data.AsDataView();
                        string filter = TDMK_Code.filter_str(new string[] { "Zone", "PCS_No" }, new string[] { region, pcs });
                        dv.RowFilter = filter;

                        if (dv.Count > 0)
                        {

                            DataTable dt_time_edit = dv.ToTable();
                            List<DataTable> src_pcs_edit = new List<DataTable>();
                            Get_ListTable_datetime(-1, dt_time_edit, new string[] { "Time_Update" }, ref src_pcs_edit, "PCS_No");
                            int i = 0;
                            Dictionary<string, List<string>> dic_pcs = new Dictionary<string, List<string>> { };
                            foreach (DataTable dt_edit in src_pcs_edit)
                            {
                                if (i == 0)
                                {
                                    List<string> arr_sample_before = new List<string>();
                                    foreach (DataRow dr in dt_edit.Rows)
                                    {
                                        arr_sample_before.Add(dr["Before_Data"].ToString());
                                    }
                                    dic_pcs.Add("Before", arr_sample_before);

                                    List<string> arr_sample_after = new List<string>();
                                    foreach (DataRow dr in dt_edit.Rows)
                                    {
                                        arr_sample_after.Add(dr["After_Data"].ToString());
                                    }
                                    dic_pcs.Add("After", arr_sample_after);


                                }
                                else
                                {
                                    List<string> arr_sample_ = new List<string>();
                                    foreach (DataRow dr in dt_edit.Rows)
                                    {
                                        arr_sample_.Add(dr["After_Data"].ToString());
                                    }
                                    dic_pcs.Add("After_" + (i + 1).ToString(), arr_sample_);


                                }
                                i++;
                            }
                            src_edit.Add(pcs, dic_pcs);


                            if (i > max_count_edit)
                            {
                                max_count_edit = i;
                            }


                        }
                    }
                }
                else
                {

                    List<DataTable> tbl_sample = new List<DataTable> { };

                    Get_ListTable(-1, dt, new string[] { "Pcs_No" }, ref tbl_sample, "Data");
                    foreach (DataTable dt_edit in tbl_sample)
                    {
                        //if (dt_edit.Rows[0]["Pcs_No"].ToString() == pcs)
                        //{
                        Dictionary<string, List<string>> dic_pcs = new Dictionary<string, List<string>> { };

                        List<string> arr_sample_ = new List<string>();
                        foreach (DataRow dr in dt_edit.Rows)
                        {
                            arr_sample_.Add(dr["Data"].ToString());
                        }
                        dic_pcs.Add("Before", arr_sample_);
                        src_edit.Add(dt_edit.Rows[0]["Pcs_No"].ToString(), dic_pcs);

                        //  }

                    }


                }


                dic_region.Add(region, src_edit);

            }

            return dic_region;

        }
        public Dictionary<int, Dictionary<string, Dictionary<string, string>>> Get_table_edit_BVH_PTH(string process, ref int max_count_edit)
        {
            DataTable dt_history = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { itemcodevh, lotnovh, process }));
            tbl_history = dt_history;
            DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, process.ToUpper(), TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
            DataTable dt_data = new DataTable();
            DataTable dt_img = new DataTable();
            Split_Data_Image_BVH_PTH(dt_history, ref dt_data, ref dt_img);


            List<DataTable> tbl_pcs = new List<DataTable>();
            Get_ListTable(-1, dt_data_summary, new string[] { "Pcs_No" }, ref tbl_pcs, "Remark");
            Dictionary<int, Dictionary<string, Dictionary<string, string>>> dic_pcs = new Dictionary<int, Dictionary<string, Dictionary<string, string>>> { };


            foreach (DataTable dt in tbl_pcs)
            {
                Dictionary<string, Dictionary<string, string>> src_edit = new Dictionary<string, Dictionary<string, string>> { };
                int pcs = int.Parse(dt.Rows[0]["Pcs_No"].ToString());

                DataView dv = dt_data.AsDataView();
                if (dv.Count > 0)
                {
                    string filter = TDMK_Code.filter_str(new string[] { "PCS_No" }, new string[] { dt.Rows[0]["PCS_No"].ToString() });
                    dv.RowFilter = filter;
                }


                if (dv.Count == 0)
                {
                    Dictionary<string, string> dic_location = new Dictionary<string, string> { };
                    foreach (DataColumn dc in dt.Columns)
                    {
                        if (dc.ColumnName != "ID" && dc.ColumnName != "ItemCode" && dc.ColumnName != "LotNo" && dc.ColumnName != "PCS_No" && dc.ColumnName != "Image_data" && dc.ColumnName != "Image_data_2" && dc.ColumnName != "Remark")
                            dic_location.Add(dc.ColumnName, myCode.checkDBNull(dt.Rows[0][dc.ColumnName]));
                    }

                    src_edit.Add("Before", dic_location);
                }
                else
                {
                    List<DataTable> src_pcs_edit = new List<DataTable>();
                    Get_ListTable_datetime(-1, dv.ToTable(), new string[] { "Time_Update" }, ref src_pcs_edit, "After_Data");

                    int i = 0;

                    foreach (DataTable dt_edit in src_pcs_edit)
                    {
                        if (i == 0)
                        {
                            Dictionary<string, string> dic_location_before = new Dictionary<string, string> { };
                            foreach (DataRow dr in dt_edit.Rows)
                            {
                                dic_location_before.Add(dr["Region"].ToString(), myCode.checkDBNull(dr["Before_Data"]));
                            }
                            src_edit.Add("Before", dic_location_before);

                            //DataView dv_image = dt_img.AsDataView();
                            //string filter_image = TDMK_Code.filter_str(new string[] { "PCS_No", "Time_Update" }, new string[] { dt_edit.Rows[0]["PCS_No"].ToString(), dt_edit.Rows[0]["Time_Update"].ToString() });
                            //dv_image.RowFilter = filter_image;
                            //dic_image.Add((byte[])dv_image.ToTable().Rows[0]["Before_Data"], (byte[]) dv_image.ToTable().Rows[0]["Before_Data"]);


                            Dictionary<string, string> dic_location_after = new Dictionary<string, string> { };
                            foreach (DataRow dr in dt_edit.Rows)
                            {
                                dic_location_after.Add(dr["Region"].ToString(), myCode.checkDBNull(dr["After_Data"]));
                            }

                            src_edit.Add("After_1", dic_location_after);
                        }
                        else
                        {
                            Dictionary<string, string> dic_location_ = new Dictionary<string, string> { };
                            foreach (DataRow dr in dt_edit.Rows)
                            {
                                dic_location_.Add(dr["Region"].ToString(), myCode.checkDBNull(dr["After_Data"]));
                            }
                            src_edit.Add("After_" + (i + 1).ToString(), dic_location_);
                        }

                        i++;

                    }
                    if (i > max_count_edit)
                    {
                        max_count_edit = i;
                    }
                }



                dic_pcs.Add(pcs, src_edit);

            }

            return dic_pcs;
        }

        public Dictionary<string, Dictionary<string, Dictionary<string, string>>> Get_table_edit_Impedance_data(string process, ref int max_count_edit)
        {
            DataTable dt_history = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { itemcodevh, lotnovh, process }));
            tbl_history = dt_history;
            DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, process.ToUpper() + "_VAL", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
            DataTable dt_data = new DataTable();
            DataTable dt_img = new DataTable();
            Split_Data_Image_Impedance(dt_history, ref dt_data, ref dt_img);

            List<DataTable> tbl_region = new List<DataTable>();

            Dictionary<string, Dictionary<string, Dictionary<string, string>>> dic_pcs = new Dictionary<string, Dictionary<string, Dictionary<string, string>>> { };

            Get_ListTable(-1, dt_data_summary, new string[] { "Region" }, ref tbl_region, "Remark");

            foreach (DataTable dt in tbl_region)
            {
                Dictionary<string, Dictionary<string, string>> src_edit = new Dictionary<string, Dictionary<string, string>> { };
                string region = dt.Rows[0]["Region"].ToString();
                if (!region.Contains("Impedance_"))
                {
                    region = "Impedance_" + region;

                }
                DataView dv = dt_data.AsDataView();
                if (dv.Count > 0)
                {
                    string filter = TDMK_Code.filter_str(new string[] { "Zone" }, new string[] { region });
                    dv.RowFilter = filter;
                }


                if (dv.Count > 0)
                {

                    List<DataTable> src_pcs_edit = new List<DataTable>();
                    Get_ListTable_datetime(-1, dv.ToTable(), new string[] { "Time_Update" }, ref src_pcs_edit, "Remark");
                    int i = 0;

                    foreach (DataTable dt_edit in src_pcs_edit)
                    {
                        if (i == 0)
                        {
                            Dictionary<string, string> dic_pcs_before = new Dictionary<string, string> { };
                            foreach (DataRow dr in dt_edit.Rows)
                            {
                                dic_pcs_before.Add(dr["PCS_No"].ToString(), myCode.checkDBNull(dr["Before_Data"]));
                            }
                            src_edit.Add("Before", dic_pcs_before);



                            Dictionary<string, string> dic_pcs_after = new Dictionary<string, string> { };
                            foreach (DataRow dr in dt_edit.Rows)
                            {
                                dic_pcs_after.Add(dr["PCS_No"].ToString(), myCode.checkDBNull(dr["After_Data"]));
                            }

                            src_edit.Add("After_1", dic_pcs_after);
                        }
                        else
                        {
                            Dictionary<string, string> dic_pcs_ = new Dictionary<string, string> { };
                            foreach (DataRow dr in dt_edit.Rows)
                            {
                                dic_pcs_.Add(dr["PCS_No"].ToString(), myCode.checkDBNull(dr["After_Data"]));
                            }
                            src_edit.Add("After_" + (i + 1).ToString(), dic_pcs_);
                        }

                        i++;

                    }
                    if (i > max_count_edit)
                    {
                        max_count_edit = i;
                    }
                }
                else
                {
                    Dictionary<string, string> dic_pcs_data = new Dictionary<string, string> { };

                    foreach (DataRow dr in dt.Rows)
                    {
                        if (!dic_pcs_data.ContainsKey(dr["PCS_No"].ToString()))
                        {
                            dic_pcs_data.Add(dr["PCS_No"].ToString(), dr["Data"].ToString());
                        }

                    }
                    src_edit.Add("Before", dic_pcs_data);
                }



                dic_pcs.Add(region, src_edit);

            }

            return dic_pcs;

        }
        public Dictionary<string, Dictionary<string, Dictionary<string, string>>> Get_table_edit_MOISTURE(ref int max_count_edit)
        {
            DataTable dt_history = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { itemcodevh, lotnovh, "CQRA_MOISTURE_ABSORPTION" }));
            tbl_history = dt_history;
            DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_MOISTURE_ABSORPTION", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
            DataTable dt_data = new DataTable();
            DataTable dt_img = new DataTable();
            Split_Data_Image(dt_history, ref dt_data, ref dt_img);



            Dictionary<string, Dictionary<string, Dictionary<string, string>>> dic_all_zone = new Dictionary<string, Dictionary<string, Dictionary<string, string>>> { };
            max_count_edit = 0;
            if(dt_data_summary.Rows.Count > 0) 
            {
                for (int pcs = 1; pcs <= 5; pcs++)
                {

                    Dictionary<string, Dictionary<string, string>> dic_pcs = new Dictionary<string, Dictionary<string, string>> { };
                    string Pcs_No = pcs.ToString();
                    for (int k = 3; k <= 4; k++)
                    {
                        Dictionary<string, string> src_edit = new Dictionary<string, string> { };
                        string zone = dt_data_summary.Columns[k].ColumnName;
                        DataView dv = dt_data.AsDataView();

                        if (dv.Count > 0)
                        {
                            string filter = TDMK_Code.filter_str(new string[] { "Zone", "PCS_No" }, new string[] { zone, Pcs_No });
                            dv.RowFilter = filter;
                        }

                        if (dv.Count == 0)
                        {
                            src_edit.Add("Before", dt_data_summary.Rows[pcs - 1][zone].ToString());
                        }
                        else
                        {
                            List<DataTable> src_pcs_edit = new List<DataTable>();
                            Get_ListTable_datetime(-1, dv.ToTable(), new string[] { "Time_Update" }, ref src_pcs_edit, "After_Data");
                            int i = 0;

                            foreach (DataTable dt_edit in src_pcs_edit)
                            {
                                if (i == 0)
                                {
                                    src_edit.Add("Before", dt_edit.Rows[0]["Before_Data"].ToString());
                                    src_edit.Add("After_1", dt_edit.Rows[0]["After_Data"].ToString());
                                }
                                else
                                {

                                    src_edit.Add("After_" + (i + 1).ToString(), dt_edit.Rows[0]["After_Data"].ToString());


                                }

                                i++;

                            }
                            if (i > max_count_edit)
                            {
                                max_count_edit = i;
                            }
                        }



                        dic_pcs.Add(zone, src_edit);

                    }
                    dic_all_zone.Add(Pcs_No, dic_pcs);

                }
                
            }
            
          
            return dic_all_zone;

        }
        public Dictionary<string, Dictionary<string, Dictionary<string, string>>> Get_table_edit_ACFFlatness(ref int max_count_edit)
        {
            DataTable dt_history = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { itemcodevh, lotnovh, "ACF_Flatness" }));
          
            DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "ACF_FLATNESS", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
            DataTable dt_data = new DataTable();
            DataTable dt_img = new DataTable();
            Split_Data_Image(dt_history, ref dt_data, ref dt_img);
            tbl_history = dt_data;

            List<DataTable> src_sample = new List<DataTable>();
            Get_ListTable(-1, dt_data_summary, new string[] { "Sample" }, ref src_sample, "LotNo");
            Dictionary<string, Dictionary<string, Dictionary<string, string>>> dic_all_zone = new Dictionary<string, Dictionary<string, Dictionary<string, string>>> { };
            max_count_edit = 0;
            foreach (DataTable dt_zone in src_sample)
            {
                List<DataTable> tbl_pcs = new List<DataTable>();

                Dictionary<string, Dictionary<string, string>> dic_pcs = new Dictionary<string, Dictionary<string, string>> { };

                string zone = dt_zone.Rows[0]["Sample"].ToString();

                for (int point = 1; point <= 10; point++)
                {
                    Dictionary<string, string> src_edit = new Dictionary<string, string> { };
                    string pcs = "Point" + point.ToString();

                    DataView dv = dt_data.AsDataView();

                    if (dv.Count > 0)
                    {
                        string filter = TDMK_Code.filter_str(new string[] { "Zone", "PCS_No" }, new string[] { zone, "Point " + point.ToString() });
                        dv.RowFilter = filter;
                    }

                    if (dv.Count == 0)
                    {
                        src_edit.Add("Before", dt_zone.Rows[0][pcs].ToString());
                    }
                    else
                    {
                        List<DataTable> src_pcs_edit = new List<DataTable>();
                        Get_ListTable_datetime(-1, dv.ToTable(), new string[] { "Time_Update" }, ref src_pcs_edit, "After_Data");
                        int i = 0;

                        foreach (DataTable dt_edit in src_pcs_edit)
                        {
                            if (i == 0)
                            {
                                src_edit.Add("Before", dt_edit.Rows[0]["Before_Data"].ToString());
                                src_edit.Add("After_1", dt_edit.Rows[0]["After_Data"].ToString());
                            }
                            else
                            {

                                //foreach (DataRow dr in dt_edit.Rows)
                                //{
                                src_edit.Add("After_" + (i + 1).ToString(), dt_edit.Rows[0]["After_Data"].ToString());
                                // }

                            }

                            i++;

                        }
                        if (i > max_count_edit)
                        {
                            max_count_edit = i;
                        }
                    }



                    dic_pcs.Add(pcs, src_edit);

                }
                dic_all_zone.Add(zone, dic_pcs);

            }
            return dic_all_zone;

        }


        public Dictionary<string, Dictionary<int, Dictionary<string, List<string>>>> Get_table_edit_stackup(ref int max_count_edit)
        {
            DataTable dt_history = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { itemcodevh, lotnovh, "Stackup" }));
           
            DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
            DataTable dt_data = new DataTable();
            DataTable dt_img = new DataTable();
            Split_Data_Image(dt_history, ref dt_data, ref dt_img);
            tbl_history = dt_data;

            List<DataTable> src_sample = new List<DataTable>();
            Get_ListTable(-1, dt_data_summary, new string[] { "Region" }, ref src_sample, "Data");
            Dictionary<string, Dictionary<int, Dictionary<string, List<string>>>> dic_all_zone = new Dictionary<string, Dictionary<int, Dictionary<string, List<string>>>> { };
            max_count_edit = 0;
            foreach (DataTable dt_zone in src_sample)
            {
                List<DataTable> tbl_pcs = new List<DataTable>();
                Get_ListTable(-1, dt_zone, new string[] { "Pcs_No" }, ref tbl_pcs, "Data");
                Dictionary<int, Dictionary<string, List<string>>> dic_pcs = new Dictionary<int, Dictionary<string, List<string>>> { };

                string zone = dt_zone.Rows[0]["Region"].ToString();

                foreach (DataTable dt in tbl_pcs)
                {
                    Dictionary<string, List<string>> src_edit = new Dictionary<string, List<string>> { };
                    int pcs = int.Parse(dt.Rows[0]["Pcs_No"].ToString());

                    DataView dv = dt_data.AsDataView();

                    if (dv.Count > 0)
                    {
                        string filter = TDMK_Code.filter_str(new string[] { "Zone", "PCS_No" }, new string[] { zone, dt.Rows[0]["Pcs_No"].ToString() });
                        dv.RowFilter = filter;
                    }

                    if (dv.Count == 0)
                    {
                        List<string> lst_edit_pcs = new List<string>();
                        foreach (DataRow dr in dt.Rows)
                        {
                            lst_edit_pcs.Add(dr["Data"].ToString());
                        }
                        src_edit.Add("Before", lst_edit_pcs);
                    }
                    else
                    {
                        List<DataTable> src_pcs_edit = new List<DataTable>();
                        Get_ListTable_datetime(-1, dv.ToTable(), new string[] { "Time_Update" }, ref src_pcs_edit, "After_Data");
                        int i = 0;

                        foreach (DataTable dt_edit in src_pcs_edit)
                        {
                            if (i == 0)
                            {
                                List<string> lst_edit_before = new List<string>();
                                foreach (DataRow dr in dt_edit.Rows)
                                {
                                    lst_edit_before.Add(dr["Before_Data"].ToString());
                                }
                                src_edit.Add("Before", lst_edit_before);

                                List<string> lst_edit_after = new List<string>();
                                foreach (DataRow dr in dt_edit.Rows)
                                {
                                    lst_edit_after.Add(dr["After_Data"].ToString());
                                }
                                src_edit.Add("After_1", lst_edit_after);
                            }
                            else
                            {
                                List<string> lst_edit_ = new List<string>();
                                foreach (DataRow dr in dt_edit.Rows)
                                {
                                    lst_edit_.Add(dr["After_Data"].ToString());
                                }
                                src_edit.Add("After_" + (i + 1).ToString(), lst_edit_);
                            }

                            i++;

                        }
                        if (i > max_count_edit)
                        {
                            max_count_edit = i;
                        }
                    }



                    dic_pcs.Add(pcs, src_edit);

                }
                dic_all_zone.Add(zone, dic_pcs);

            }
            return dic_all_zone;

        }
        public Dictionary<string, Dictionary<int, Dictionary<string, Byte[]>>> Get_table_edit_stackup_image(ref int max_count_edit)
        {
            DataTable dt_history = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { itemcodevh, lotnovh, "Stackup" }));

            DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
            DataTable dt_data = new DataTable();
            DataTable dt_img = new DataTable();
            Split_Data_Image(dt_history, ref dt_data, ref dt_img);
            tbl_history = dt_data;
            //DataTable dt_img = src_dt.AsEnumerable().Where(x => x.Field<string>("Region") == "Image_data").CopyToDataTable();

            //List<DataTable> src_sample = new List<DataTable>();
            //Get_ListTable(-1, dt_data_summary, new string[] { "Region" }, ref src_sample, "PCS_No");
            string[] arr_region = dt_data_summary.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();


            Dictionary<string, Dictionary<int, Dictionary<string, Byte[]>>> dic_all_zone = new Dictionary<string, Dictionary<int, Dictionary<string, Byte[]>>> { };
            max_count_edit = 0;
            foreach (string _zone in arr_region)
            {
                // List<DataTable> tbl_pcs = new List<DataTable>();
                // Get_ListTable(-1, dt_zone, new string[] { "Pcs_No" }, ref tbl_pcs, "PCS_No");
                //DataTable dt_zone = dt_data_summary.AsEnumerable().Where(x => x.Field<string>("Region") == "_zone").CopyToDataTable();

                DataView dv1 = dt_data_summary.AsDataView();
                string filter1 = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { _zone });
                dv1.RowFilter = filter1;
                DataTable dt_zone = dv1.ToTable();



                string[] arr_pcs = dt_zone.AsEnumerable().Select(x => x.Field<string>("PCS_No")).Distinct().ToArray();


                Dictionary<int, Dictionary<string, Byte[]>> dic_pcs = new Dictionary<int, Dictionary<string, Byte[]>> { };

                //string zone = dt_zone.Rows[0]["Region"].ToString();

                foreach (string _pcs in arr_pcs)
                {
                    Dictionary<string, Byte[]> src_edit = new Dictionary<string, Byte[]> { };
                    // int pcs = int.Parse(dt.Rows[0]["Pcs_No"].ToString());

                    DataView dv = dt_img.AsDataView();
                    if (dv.Count > 0)
                    {
                        string filter = TDMK_Code.filter_str(new string[] { "Zone", "PCS_No" }, new string[] { _zone, _pcs });
                        dv.RowFilter = filter;
                    }

                    //DataTable dt_pcs = dt_zone.AsEnumerable().Where(x => x.Field<string>("Pcs_No") == "_pcs").CopyToDataTable();

                    DataView dv2 = dt_zone.AsDataView();
                    string filter2 = TDMK_Code.filter_str(new string[] { "Pcs_No" }, new string[] { _pcs });
                    dv2.RowFilter = filter2;
                    DataTable dt_pcs = dv2.ToTable();

                    if (dv.Count == 0)
                    {
                        src_edit.Add("Before", (Byte[])dt_pcs.Rows[0]["Image_Data"]);
                    }
                    else
                    {
                        // List<DataTable> src_pcs_edit = new List<DataTable>();
                        // Get_ListTable_datetime(-1, dv.ToTable(), new string[] { "Time_Update" }, ref src_pcs_edit, "PCS_No");
                        DataTable dt_all_time = dv.ToTable();

                        int i = 0;

                        foreach (DataRow dr in dt_all_time.Rows)
                        {

                            if (i == 0)
                            {

                                src_edit.Add("Before", (byte[])dt_all_time.Rows[0]["Before_Data"]);
                                src_edit.Add("After_1", (byte[])dt_all_time.Rows[0]["After_Data"]);
                            }
                            else
                            {
                                src_edit.Add("After_" + (i + 1).ToString(), (byte[])dt_all_time.Rows[i]["After_Data"]);
                            }

                            i++;

                        }
                        if (i > max_count_edit)
                        {
                            max_count_edit = i;
                        }
                    }



                    dic_pcs.Add(int.Parse(_pcs), src_edit);

                }
                dic_all_zone.Add(_zone, dic_pcs);

            }
            return dic_all_zone;

        }
        public Dictionary<string, Dictionary<string, Byte[]>> Get_table_edit_image_SolderMask(ref int max_count_edit)
        {
            DataTable dt_history = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { itemcodevh, lotnovh, "SolderMask" }));

            DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "SOLDERMASK_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
            DataTable dt_data = new DataTable();
            DataTable dt_img = new DataTable();
            Split_Data_Image(dt_history, ref dt_data, ref dt_img);
            tbl_history = dt_img;

            string[] arr_pcs = dt_data_summary.AsEnumerable().Select(x => x.Field<string>("Pcs_No")).Distinct().ToArray();


            Dictionary<string, Dictionary<string, Byte[]>> dic_all_pcs = new Dictionary<string, Dictionary<string, Byte[]>> { };
            max_count_edit = 0;
            foreach (string _pcs in arr_pcs)
            {


                Dictionary<string, Byte[]> dic_edit = new Dictionary<string, Byte[]> { };


                DataView dv = dt_img.AsDataView();
                string filter = TDMK_Code.filter_str(new string[] { "PCS_No" }, new string[] { _pcs });
                dv.RowFilter = filter;

                DataView dv2 = dt_data_summary.AsDataView();
                string filter2 = TDMK_Code.filter_str(new string[] { "Pcs_No" }, new string[] { _pcs });
                dv2.RowFilter = filter2;
                DataTable dt_pcs = dv2.ToTable();

                if (dv.Count == 0)
                {
                    dic_edit.Add("Before", (Byte[])dt_pcs.Rows[0]["Image_Data"]);
                }
                else
                {

                    DataTable dt_all_time = dv.ToTable();

                    int i = 0;

                    foreach (DataRow dr in dt_all_time.Rows)
                    {

                        if (i == 0)
                        {

                            dic_edit.Add("Before", (byte[])dt_all_time.Rows[0]["Before_Data"]);
                            dic_edit.Add("After_1", (byte[])dt_all_time.Rows[0]["After_Data"]);
                        }
                        else
                        {
                            dic_edit.Add("After_" + (i + 1).ToString(), (byte[])dt_all_time.Rows[i]["After_Data"]);
                        }

                        i++;

                    }
                    if (i > max_count_edit)
                    {
                        max_count_edit = i;
                    }
                }



                dic_all_pcs.Add(_pcs, dic_edit);


            }
            return dic_all_pcs;

        }
        public Dictionary<string, Dictionary<string, Dictionary<string, Byte[]>>> Get_table_edit_image_Thermal_stress(ref int max_count_edit)
        {
            DataTable dt_history = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { itemcodevh, lotnovh, "ThermalStress" }));

            DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "THERMAL_STRESS_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
            DataTable dt_data = new DataTable();
            DataTable dt_img = new DataTable();
            Split_Data_Image(dt_history, ref dt_data, ref dt_img);
            tbl_history = dt_img;
            Dictionary<string, Dictionary<string, Dictionary<string, Byte[]>>> dic_all_pcs = new Dictionary<string, Dictionary<string, Dictionary<string, byte[]>>> { };
            string[] arr_region = dt_data_summary.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();
            foreach (string region in arr_region)
            {
                Dictionary<string, Dictionary<string, Byte[]>> dic_region = new Dictionary<string, Dictionary<string, byte[]>> { };

                DataView dv_region = dt_img.AsDataView();
                string filter = TDMK_Code.filter_str(new string[] { "Zone" }, new string[] { region });
                dv_region.RowFilter = filter;

                string[] arr_pcs = dv_region.ToTable().AsEnumerable().Select(x => x.Field<string>("Pcs_No")).Distinct().ToArray();
                max_count_edit = 0;
                foreach (string _pcs in arr_pcs)
                {
                    Dictionary<string, Byte[]> dic_edit = new Dictionary<string, Byte[]> { };


                    DataView dv = dt_img.AsDataView();

                    dv.RowFilter = TDMK_Code.filter_str(new string[] { "PCS_No", "Zone" }, new string[] { _pcs , region});

                    DataView dv2 = dt_data_summary.AsDataView();
                    string filter2 = TDMK_Code.filter_str(new string[] { "Pcs_No", "Region" }, new string[] { _pcs, region });
                    dv2.RowFilter = filter2;
                    DataTable dt_pcs = dv2.ToTable();

                    if (dv.Count == 0)
                    {
                        dic_edit.Add("Before", (Byte[])dt_pcs.Rows[0]["Image_Data"]);
                    }
                    else
                    { 
                        DataTable dt_all_time = dv.ToTable();
                        int i = 0;

                        foreach (DataRow dr in dt_all_time.Rows)
                        { 
                            if (i == 0)
                            { 
                                dic_edit.Add("Before", (byte[])dt_all_time.Rows[0]["Before_Data"]);
                                dic_edit.Add("After_1", (byte[])dt_all_time.Rows[0]["After_Data"]);
                            }
                            else
                            {
                                dic_edit.Add("After_" + (i + 1).ToString(), (byte[])dt_all_time.Rows[i]["After_Data"]);
                            }

                            i++;

                        }
                        if (i > max_count_edit)
                        {
                            max_count_edit = i;
                        }
                    }
                    dic_region.Add(_pcs, dic_edit);


                }
                dic_all_pcs.Add(region, dic_region);

            }



            return dic_all_pcs;

        }
        public Dictionary<string, Dictionary<int, Dictionary<string, Byte[]>>> Get_table_edit_stackup_image_zone(string zone, ref int max_count_edit)
        {
            DataTable dt_history = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { itemcodevh, lotnovh, "Stackup" }));
            DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
            DataTable dt_data = new DataTable();
            DataTable dt_img = new DataTable();
            Split_Data_Image(dt_history, ref dt_data, ref dt_img);

            string[] arr_region = dt_data_summary.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();


            Dictionary<string, Dictionary<int, Dictionary<string, Byte[]>>> dic_all_zone = new Dictionary<string, Dictionary<int, Dictionary<string, Byte[]>>> { };
            max_count_edit = 0;
            foreach (string _zone in arr_region)
            {
                if (_zone == zone)
                {


                    DataView dv1 = dt_data_summary.AsDataView();
                    string filter1 = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { _zone });
                    dv1.RowFilter = filter1;
                    DataTable dt_zone = dv1.ToTable();



                    string[] arr_pcs = dt_zone.AsEnumerable().Select(x => x.Field<string>("PCS_No")).Distinct().ToArray();
                    Dictionary<int, Dictionary<string, Byte[]>> dic_pcs = new Dictionary<int, Dictionary<string, Byte[]>> { };

                    foreach (string _pcs in arr_pcs)
                    {
                        Dictionary<string, Byte[]> src_edit = new Dictionary<string, Byte[]> { };

                        DataView dv = dt_img.AsDataView();
                        if(dv.Count > 0)
                        {
                            string filter = TDMK_Code.filter_str(new string[] { "Zone", "PCS_No" }, new string[] { _zone, _pcs });
                            dv.RowFilter = filter;
                        }
                         

                        DataView dv2 = dt_zone.AsDataView();
                        string filter2 = TDMK_Code.filter_str(new string[] { "Pcs_No" }, new string[] { _pcs });
                        dv2.RowFilter = filter2;
                        DataTable dt_pcs = dv2.ToTable();

                        if (dv.Count == 0)
                        {
                            src_edit.Add("Before", (Byte[])dt_pcs.Rows[0]["Image_Data"]);
                        }
                        else
                        {

                            DataTable dt_all_time = dv.ToTable();

                            int i = 0;

                            foreach (DataRow dr in dt_all_time.Rows)
                            {

                                if (i == 0)
                                {

                                    src_edit.Add("Before", (byte[])dt_all_time.Rows[0]["Before_Data"]);
                                    src_edit.Add("After_1", (byte[])dt_all_time.Rows[0]["After_Data"]);
                                }
                                else
                                {
                                    src_edit.Add("After_" + (i + 1).ToString(), (byte[])dt_all_time.Rows[i]["After_Data"]);
                                }

                                i++;

                            }
                            if (i > max_count_edit)
                            {
                                max_count_edit = i;
                            }
                        }



                        dic_pcs.Add(int.Parse(_pcs), src_edit);

                    }
                    dic_all_zone.Add(_zone, dic_pcs);
                    break;
                }

            }
            return dic_all_zone;

        }
        public Dictionary<string, Dictionary<int, Dictionary<string, Byte[]>>> Get_table_edit_stackup_image_zone_pcs(string zone, string pcs, ref int max_count_edit)
        {
            DataTable dt_history = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { itemcodevh, lotnovh, "Stackup" }));
            DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
            DataTable dt_data = new DataTable();
            DataTable dt_img = new DataTable();
            Split_Data_Image(dt_history, ref dt_data, ref dt_img);

            string[] arr_region = dt_data_summary.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();


            Dictionary<string, Dictionary<int, Dictionary<string, Byte[]>>> dic_all_zone = new Dictionary<string, Dictionary<int, Dictionary<string, Byte[]>>> { };
            max_count_edit = 0;
            foreach (string _zone in arr_region)
            {
                if (_zone == zone)
                {


                    DataView dv1 = dt_data_summary.AsDataView();
                    string filter1 = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { _zone });
                    dv1.RowFilter = filter1;
                    DataTable dt_zone = dv1.ToTable();



                    string[] arr_pcs = dt_zone.AsEnumerable().Select(x => x.Field<string>("PCS_No")).Distinct().ToArray();
                    Dictionary<int, Dictionary<string, Byte[]>> dic_pcs = new Dictionary<int, Dictionary<string, Byte[]>> { };

                    foreach (string _pcs in arr_pcs)
                    {
                        if (_pcs == pcs)
                        {


                            Dictionary<string, Byte[]> src_edit = new Dictionary<string, Byte[]> { };

                            DataView dv = dt_img.AsDataView();
                            string filter = TDMK_Code.filter_str(new string[] { "Zone", "PCS_No" }, new string[] { _zone, _pcs });
                            dv.RowFilter = filter;

                            DataView dv2 = dt_zone.AsDataView();
                            string filter2 = TDMK_Code.filter_str(new string[] { "Pcs_No" }, new string[] { _pcs });
                            dv2.RowFilter = filter2;
                            DataTable dt_pcs = dv2.ToTable();

                            if (dv.Count == 0)
                            {
                                src_edit.Add("Before", (Byte[])dt_pcs.Rows[0]["Image_Data"]);
                            }
                            else
                            {

                                DataTable dt_all_time = dv.ToTable();

                                int i = 0;

                                foreach (DataRow dr in dt_all_time.Rows)
                                {

                                    if (i == 0)
                                    {

                                        src_edit.Add("Before", (byte[])dt_all_time.Rows[0]["Before_Data"]);
                                        src_edit.Add("After_1", (byte[])dt_all_time.Rows[0]["After_Data"]);
                                    }
                                    else
                                    {
                                        src_edit.Add("After_" + (i + 1).ToString(), (byte[])dt_all_time.Rows[i]["After_Data"]);
                                    }

                                    i++;

                                }
                                if (i > max_count_edit)
                                {
                                    max_count_edit = i;
                                }
                            }



                            dic_pcs.Add(int.Parse(_pcs), src_edit);
                            break;
                        }

                    }
                    dic_all_zone.Add(_zone, dic_pcs);
                    break;
                }

            }
            return dic_all_zone;

        }
        public void check_spec_Impedance(DataTable dt_data)
        {

            DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { itemcodevh }));
            tbl_spec = dt_spec;
            if (dt_spec.Rows.Count > 0)
            {
                for (int i = 0; i < dt_data.Rows.Count; i++)
                {
                    for (int j = 3; j < dt_data.Columns.Count; j++)
                    {
                        int region = int.Parse(dgv_edit_data.Rows[i].Cells["Region"].Value.ToString().Replace("Impedance_", ""));
                        if (region <= dt_spec.Rows.Count)
                        {
                            Double max_target = Double.Parse(dt_spec.Rows[region - 1]["USL"].ToString());
                            Double min_target = Double.Parse(dt_spec.Rows[region - 1]["LSL"].ToString());
                            if (myCode.checkDBNull(dt_data.Rows[i][j]) != "")
                            {
                                Double data = Double.Parse(dt_data.Rows[i][j].ToString());
                                if (data > max_target || data < min_target)
                                {
                                    dgv_edit_data.Rows[i].Cells[j].Style.BackColor = Color.Red;
                                }
                            }
                            else
                            {
                                dgv_edit_data.Rows[i].Cells[j].Style.BackColor = Color.White;

                            }
                        }
                    }

                }
            }
        }

        public void check_spec_Tracewidth(DataTable dt_data)
        {
            DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { itemcodevh }));
            tbl_spec = dt_spec;
            if (dt_spec.Rows.Count > 0)
            {
                for (int i = 0; i < dt_data.Rows.Count; i++)
                {
                    for (int j = 3; j < dt_data.Columns.Count; j++)
                    {
                        int region = int.Parse(dt_data.Rows[i]["Region"].ToString().Replace("Tracewidth_", ""));

                        Double max_target = Double.Parse(dt_spec.Rows[region - 1]["USL"].ToString());
                        Double min_target = Double.Parse(dt_spec.Rows[region - 1]["LSL"].ToString());
                        Double data = Double.Parse(dt_data.Rows[i][j].ToString());
                        if (data > max_target || data < min_target)
                        {
                            dgv_edit_data.Rows[i].Cells[j].Style.BackColor = Color.Red;
                        }

                        else
                        {
                            dgv_edit_data.Rows[i].Cells[j].Style.BackColor = Color.White;
                        }
                    }

                }
            }
        }

        private void cb_sheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            dgv_edit_data.DataSource = DGV_data_image.DataSource = null;


            cb_zone.Items.Clear();
            cb_zone.Text = "ALL";
            cb_pcs.Items.Clear();
            cb_pcs.Text = "ALL";
            dgv_edit_data.Columns.Clear();
            dgv_edit_data.Rows.Clear();
            DGV_data_image.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            lbl_image.Text = "Image/Graph";
            DGV_data_image.Columns.Clear();
            DGV_data_image.Rows.Clear();

            if (itemcodevh != "" && lotnovh != "")
            {
                if (sheet == "Stack-up")
                {
                    DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
                    if (dt_data_summary.Rows.Count > 0)
                    {
 
                    dgv_edit_data.AllowUserToAddRows = true;
                    DataTable dt_edit_history = new DataTable();
                    dt_edit_history.Columns.Add("ID");
                    dt_edit_history.Columns.Add("Zone");
                    dt_edit_history.Columns.Add("PCS");
                    dt_edit_history.Columns.Add("Before");
                    int max_count_edit = 0;
                    Dictionary<string, Dictionary<int, Dictionary<string, List<string>>>> table_edit = Get_table_edit_stackup(ref max_count_edit);
                    for (int i = 0; i < max_count_edit; i++)
                    {
                        dt_edit_history.Columns.Add("After_" + (i + 1).ToString());
                    }


                    Dictionary<int, Dictionary<string, List<string>>> table_zone_edit = new Dictionary<int, Dictionary<string, List<string>>> { };
                    List<int> count_edit = new List<int> { };
                    int id = 1;
                    foreach (var tbl_zone_edit in table_edit)
                    {
                        string zone = tbl_zone_edit.Key;
                        foreach (var tbl_pcs_edit in tbl_zone_edit.Value)
                        {
                            int r_pcs = tbl_pcs_edit.Value["Before"].Count;
                            int pcs = tbl_pcs_edit.Key;
                            
                            for (int i = 0; i < r_pcs; i++)
                            {
                                DataRow dr = dt_edit_history.NewRow();
                                int idx = 3;

                                dr[0] = id;
                                dr[1] = zone;
                                dr[2] = pcs;

                                foreach (var lst_data_pcs in tbl_pcs_edit.Value)
                                {
                                    dr[idx] = lst_data_pcs.Value[i];
                                    idx++;
                                }

                                dt_edit_history.Rows.Add(dr);
                                id++;
                            }

                                DataRow dr_sum = dt_edit_history.NewRow();
                                dr_sum[0] = "";
                                dr_sum[1] = "";
                                dr_sum[2] = "";
                                int in_col = 3;
                                foreach (var lst_data_pcs in tbl_pcs_edit.Value)
                                {
                                    double sum = 0;
                                    foreach (string data in lst_data_pcs.Value)
                                    {
                                        sum += double.Parse(data);
                                    }
                                    dr_sum[in_col] = sum;
                                    in_col++;
                                }
                                dt_edit_history.Rows.Add(dr_sum);

                            }
                    }
                    dgv_edit_data.DataSource = dt_edit_history;
                        for(int i = 0; i < dgv_edit_data.Rows.Count - 1; i++)
                        {
                            if (dgv_edit_data.Rows[i].Cells["Zone"].Value.ToString() == "")
                            {
                                for(int j = 0; i < dgv_edit_data.ColumnCount; j++)
                                {
                                    dgv_edit_data.Rows[i].Cells[j].Style.BackColor = Color.Yellow;
                                }
                            }
                        }

                    tbl_data = dt_edit_history;
                    string[] arr_zone = tbl_data.AsEnumerable().Select(x => x.Field<string>("Zone")).Distinct().ToArray();
                    foreach (string item in arr_zone)
                    {
                        cb_zone.Items.Add(item);
                    }

                    check_spec_stackup_detail();
                    fill_table_image();
                }
                else
                {
                    MessageBox.Show("No data", "Warning");
                }
                    
                }
                if (sheet == "BVH_PTH")
                {
                    dgv_edit_data.AllowUserToAddRows = true;

                    cb_zone.Items.Add("BVH_with_Bonding_Sheet");
                    cb_zone.Items.Add("BVH_without_Bonding_Sheet");
                    cb_zone.Items.Add("Plated_Through_Hole");

                }
                if (sheet == "Impedance")
                {
                    dgv_edit_data.AllowUserToAddRows = true;
                    cb_zone.Items.Add("Impedance");
                    cb_zone.Items.Add("Tracewidth");
                }

                if (sheet == "SOLDERMASK")
                {
                    DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "SOLDERMASK_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
                    if (dt_data_summary.Rows.Count > 0)
                    {

                        dgv_edit_data.AllowUserToAddRows = false;
                        fill_table_image_soldermask();
                        DataTable dt_inf = new DataTable();
                        dt_inf.Columns.Add("Edit");
                        dt_inf.Columns.Add("Depart");
                        dt_inf.Columns.Add("Operator");
                        dt_inf.Columns.Add("Logfile Location");
                        dt_inf.Columns.Add("Time_Update");
                        for (int i = 2; i < dgv_edit_data.Columns.Count; i++)
                        {

                            List<DataTable> tbl_time = new List<DataTable> { };
                            Get_ListTable_datetime(-1, tbl_history, new string[] { "Time_Update" }, ref tbl_time, "Remark");
                            string colname_select = dgv_edit_data.Columns[i].Name;
                            if (myCode.checkDBNull(dgv_edit_data.Rows[0].Cells[i].Value) != "")
                            {
                                if (colname_select.Contains("After_"))
                                {
                                    int time_edit = int.Parse(colname_select.Replace("After_", ""));
                                    if (time_edit <= tbl_time.Count)
                                    {
                                        DataTable tbl_select = tbl_time[time_edit - 1];
                                        string depart = tbl_select.Rows[0]["Depart"].ToString();
                                        string time_update = tbl_select.Rows[0]["Time_Update"].ToString();
                                        string Operator = tbl_select.Rows[0]["Operator"].ToString();
                                        string remark = tbl_select.Rows[0]["Remark"].ToString();
                                        DataRow dr = dt_inf.NewRow();
                                        dr[0] = colname_select;
                                        dr[1] = depart;
                                        dr[2] = Operator;
                                        dr[3] = remark;
                                        dr[4] = time_update;

                                        dt_inf.Rows.Add(dr);
                                    }
                                }
                            }

                        }

                        DGV_data_image.DataSource = dt_inf;
                        DGV_data_image.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                        lbl_image.Text = "Edit Information";
                    }
                    else
                    {
                        MessageBox.Show("No data", "Warning");
                    }
                }

                if (sheet == "THERMAL_STRESS")
                {
                    DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "THERMAL_STRESS_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
                    if (dt_data_summary.Rows.Count > 0)
                    {


                        dgv_edit_data.AllowUserToAddRows = false;
                        cb_zone.Items.Add("BVH");
                        cb_zone.Items.Add("PTH");
                        fill_table_image_thermalstress();


                        DataTable dt_inf = new DataTable();
                        dt_inf.Columns.Add("Edit");
                        dt_inf.Columns.Add("Depart");
                        dt_inf.Columns.Add("Operator");
                        dt_inf.Columns.Add("Logfile Location");
                        dt_inf.Columns.Add("Time_Update");
                        for (int i = 4; i < dgv_edit_data.Columns.Count; i++)
                        {

                            List<DataTable> tbl_time = new List<DataTable> { };
                            Get_ListTable_datetime(-1, tbl_history, new string[] { "Time_Update" }, ref tbl_time, "Remark");
                            string colname_select = dgv_edit_data.Columns[i].Name;
                            if (myCode.checkDBNull(dgv_edit_data.Rows[0].Cells[i].Value) != "")
                            {
                                if (colname_select.Contains("After_"))
                                {
                                    int time_edit = int.Parse(colname_select.Replace("After_", ""));
                                    if (time_edit <= tbl_time.Count)
                                    {
                                        DataTable tbl_select = tbl_time[time_edit - 1];
                                        string depart = tbl_select.Rows[0]["Depart"].ToString();
                                        string time_update = tbl_select.Rows[0]["Time_Update"].ToString();
                                        string Operator = tbl_select.Rows[0]["Operator"].ToString();
                                        string remark = tbl_select.Rows[0]["Remark"].ToString();
                                        DataRow dr = dt_inf.NewRow();
                                        dr[0] = colname_select;
                                        dr[1] = depart;
                                        dr[2] = Operator;
                                        dr[3] = remark;
                                        dr[4] = time_update;

                                        dt_inf.Rows.Add(dr);
                                    }
                                }
                            }

                        }

                        DGV_data_image.DataSource = dt_inf;
                        DGV_data_image.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                        lbl_image.Text = "Edit Information"; 
                    }
                    else
                    {
                        MessageBox.Show("No data", "Warning");
                    }
                }

                if (sheet == "ACF_FLATNESS")
                {
                    DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "ACF_FLATNESS", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
                    if (dt_data_summary.Rows.Count > 0)
                    {


                        dgv_edit_data.AllowUserToAddRows = true;
                        DataTable dt_edit_history = new DataTable();
                        dt_edit_history.Columns.Add("ID");
                        dt_edit_history.Columns.Add("Sample");
                        dt_edit_history.Columns.Add("Point");
                        dt_edit_history.Columns.Add("Before");
                        int max_count_edit = 0;
                        Dictionary<string, Dictionary<string, Dictionary<string, string>>> table_edit = Get_table_edit_ACFFlatness(ref max_count_edit);
                        for (int i = 0; i < max_count_edit; i++)
                        {
                            dt_edit_history.Columns.Add("After_" + (i + 1).ToString());
                        }


                        Dictionary<string, Dictionary<string, string>> table_zone_edit = new Dictionary<string, Dictionary<string, string>> { };
                        List<int> count_edit = new List<int> { };
                        int id = 1;
                        foreach (var tbl_zone_edit in table_edit)
                        {
                            string zone = tbl_zone_edit.Key;
                            foreach (var tbl_pcs_edit in tbl_zone_edit.Value)
                            {
                                string pcs = tbl_pcs_edit.Key;

                                DataRow dr = dt_edit_history.NewRow();
                                int idx = 3;

                                dr[0] = id;
                                dr[1] = zone;
                                dr[2] = pcs;

                                foreach (var lst_data_pcs in tbl_pcs_edit.Value)
                                {
                                    dr[idx] = lst_data_pcs.Value;
                                    idx++;
                                }

                                dt_edit_history.Rows.Add(dr);
                                id++;
                            }
                        }
                        dgv_edit_data.DataSource = dt_edit_history;
                        tbl_data = dt_edit_history;
                        string[] arr_zone = tbl_data.AsEnumerable().Select(x => x.Field<string>("Sample")).Distinct().ToArray();
                        foreach (string item in arr_zone)
                        {
                            cb_zone.Items.Add(item);
                        }
                    }
                    else
                    {
                        MessageBox.Show("No data", "Warning");
                    }
                }

                if (sheet == "MOISTURE ABSORPTION")
                {
                    DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_MOISTURE_ABSORPTION", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
                   if(dt_data_summary.Rows.Count > 0)
                    {

                   
                    dgv_edit_data.AllowUserToAddRows = true;
                    DataTable dt_edit_history = new DataTable();
                    dt_edit_history.Columns.Add("ID");
                    dt_edit_history.Columns.Add("PCS");
                    dt_edit_history.Columns.Add("Zone");
                    dt_edit_history.Columns.Add("Before");
                    int max_count_edit = 0;
                    Dictionary<string, Dictionary<string, Dictionary<string, string>>> table_edit = Get_table_edit_MOISTURE(ref max_count_edit);
                    
                        for (int i = 0; i < max_count_edit; i++)
                        {
                            dt_edit_history.Columns.Add("After_" + (i + 1).ToString());
                        }


                        Dictionary<string, Dictionary<string, string>> table_zone_edit = new Dictionary<string, Dictionary<string, string>> { };
                        List<int> count_edit = new List<int> { };
                        int id = 1;
                        foreach (var tbl_zone_edit in table_edit)
                        {
                            string zone = tbl_zone_edit.Key;
                            foreach (var tbl_pcs_edit in tbl_zone_edit.Value)
                            {
                                string pcs = tbl_pcs_edit.Key;

                                DataRow dr = dt_edit_history.NewRow();
                                int idx = 3;
                                dr[0] = id;
                                dr[1] = zone;
                                dr[2] = pcs;

                                foreach (var lst_data_pcs in tbl_pcs_edit.Value)
                                {
                                    dr[idx] = lst_data_pcs.Value;
                                    idx++;
                                }

                                dt_edit_history.Rows.Add(dr);
                                id++;

                            }
                        }
                        dgv_edit_data.DataSource = dt_edit_history;
                        tbl_data = dt_edit_history;
                        //string[] arr_zone = tbl_data.AsEnumerable().Select(x => x.Field<string>("Pcs")).Distinct().ToArray();
                        //foreach (string item in arr_zone)
                        //{
                        //    cb_zone.Items.Add(item);
                        //}
                    }
                    else
                    {
                        MessageBox.Show("No data", "Warning");
                    }
                   
                }

                if (sheet == "CQRA_BHAST")
                {
                    dgv_edit_data.AllowUserToAddRows = true;
                    DataTable bHast_history = TDMK_Code.Datatable_Filter(sqlcon, "BHAST_EDIT_HISTORY", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
                    tbl_data = bHast_history;
                    string[] arr_remark = bHast_history.AsEnumerable().Select(x => x.Field<string>("Remark")).Distinct().ToArray();

                    
                    cb_zone.Items.Clear();
                    cb_zone.Items.Add("Before");
                    for (int i = 0; i < arr_remark.Length - 1; i++)
                    {
                        cb_zone.Items.Add("After_" + (i + 1).ToString());
                        //cb_zone.Items.Add(arr_remark[i]);
                    }


                }
            }
        }
        public void filter_image_zone(string zone)
        {
            DGV_data_image.Columns.Clear();
            DGV_data_image.Rows.Clear();

            DGV_data_image.Columns.Add("ID", "ID");
            DGV_data_image.Columns.Add("Zone", "Zone");
            DGV_data_image.Columns.Add("Pcs", "Pcs");

            DataGridViewImageColumn dgv_col = new DataGridViewImageColumn();
            dgv_col.Name = "Before";
            dgv_col.HeaderText = "Before";
            DGV_data_image.Columns.Add(dgv_col);
            int max_count_edit = 0;
            Dictionary<string, Dictionary<int, Dictionary<string, Byte[]>>> table_edit = Get_table_edit_stackup_image_zone(zone, ref max_count_edit);
            for (int i = 0; i < max_count_edit; i++)
            {
                //dt_edit_history.Columns.Add("After_" + (i + 1).ToString());

                DataGridViewImageColumn dgv_col_after = new DataGridViewImageColumn();
                dgv_col_after.Name = "After_" + (i + 1).ToString();
                dgv_col_after.HeaderText = "After_" + (i + 1).ToString();
                DGV_data_image.Columns.Add(dgv_col_after);
            }


            int id = 1;
            foreach (var tbl_zone_edit in table_edit)
            {

                foreach (var tbl_pcs_edit in tbl_zone_edit.Value)
                {

                    int pcs = tbl_pcs_edit.Key;

                    //DataRow dr = dt_edit_history.NewRow();

                    DataGridViewRow dgvRow = new DataGridViewRow();
                    DGV_data_image.Rows.Add(dgvRow);

                    int idx = 3;

                    DGV_data_image.Rows[id - 1].Cells[0].Value = id;
                    DGV_data_image.Rows[id - 1].Cells[1].Value = zone;
                    DGV_data_image.Rows[id - 1].Cells[2].Value = pcs;

                    foreach (var lst_data_pcs in tbl_pcs_edit.Value)
                    {
                        
                        DGV_data_image.Rows[id - 1].Cells[idx].Value = myCode.byteArrayToImage(lst_data_pcs.Value);
                        idx++;
                    }



                    id++;


                }

            }
            if (DGV_data_image.Rows.Count > 0)
            {
                for (int i = 3; i < DGV_data_image.Columns.Count; i++)
                {
                    ((DataGridViewImageColumn)DGV_data_image.Columns[i]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)DGV_data_image.Columns[i]).Width = 100;

                    foreach (DataGridViewRow dr in DGV_data_image.Rows)
                    {
                        if (myCode.checkDBNull(dr.Cells[i].Value) == "")
                        {
                            dr.Cells[i].Value = myCode.byteArrayToImage(img_null);
                        }
                        dr.Height = 75;
                    }
                }

            }

        }
        public void filter_image_zone_pcs(string zone, string pcs)
        {
            DGV_data_image.Columns.Clear();
            DGV_data_image.Rows.Clear();

            DGV_data_image.Columns.Add("ID", "ID");
            DGV_data_image.Columns.Add("Zone", "Zone");
            DGV_data_image.Columns.Add("Pcs", "Pcs");

            DataGridViewImageColumn dgv_col = new DataGridViewImageColumn();
            dgv_col.Name = "Before";
            dgv_col.HeaderText = "Before";
            DGV_data_image.Columns.Add(dgv_col);
            int max_count_edit = 0;
            Dictionary<string, Dictionary<int, Dictionary<string, Byte[]>>> table_edit = Get_table_edit_stackup_image_zone_pcs(zone, pcs, ref max_count_edit);
            for (int i = 0; i < max_count_edit; i++)
            {
                //dt_edit_history.Columns.Add("After_" + (i + 1).ToString());

                DataGridViewImageColumn dgv_col_after = new DataGridViewImageColumn();
                dgv_col_after.Name = "After_" + (i + 1).ToString();
                dgv_col_after.HeaderText = "After_" + (i + 1).ToString();
                DGV_data_image.Columns.Add(dgv_col_after);
            }


            int id = 1;
            foreach (var tbl_zone_edit in table_edit)
            {

                foreach (var tbl_pcs_edit in tbl_zone_edit.Value)
                {

                    if (tbl_pcs_edit.Key.ToString() == pcs)
                    {


                        DataGridViewRow dgvRow = new DataGridViewRow();
                        DGV_data_image.Rows.Add(dgvRow);

                        int idx = 3;

                        DGV_data_image.Rows[id - 1].Cells[0].Value = id;
                        DGV_data_image.Rows[id - 1].Cells[1].Value = zone;
                        DGV_data_image.Rows[id - 1].Cells[2].Value = pcs;

                        foreach (var lst_data_pcs in tbl_pcs_edit.Value)
                        {
                            DGV_data_image.Rows[id - 1].Cells[idx].Value = myCode.byteArrayToImage(lst_data_pcs.Value);
                            idx++;
                        }



                        id++;
                        break;
                    }


                }

            }
            if (DGV_data_image.Rows.Count > 0)
            {
                for (int i = 3; i < DGV_data_image.Columns.Count; i++)
                {
                    ((DataGridViewImageColumn)DGV_data_image.Columns[i]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)DGV_data_image.Columns[i]).Width = 100;

                    foreach (DataGridViewRow dr in DGV_data_image.Rows)
                    {
                        if (myCode.checkDBNull(dr.Cells[i].Value) == "")
                        {
                            dr.Cells[i].Value = myCode.byteArrayToImage(img_null);
                        }
                        dr.Height = 75;
                    }
                }

            }

        }
        public void fill_table_image_soldermask()
        {

            dgv_edit_data.Columns.Clear();
            dgv_edit_data.Rows.Clear();

            // dgv_edit_data.Columns.Add("ID", "ID");
            dgv_edit_data.Columns.Add("PCS", "PCS");


            DataGridViewImageColumn dgv_col = new DataGridViewImageColumn();
            dgv_col.Name = "Before";
            dgv_col.HeaderText = "Before";
            dgv_edit_data.Columns.Add(dgv_col);
            int max_count_edit = 0;
            Dictionary<string, Dictionary<string, Byte[]>> table_edit = Get_table_edit_image_SolderMask(ref max_count_edit);
            for (int i = 0; i < max_count_edit; i++)
            {

                DataGridViewImageColumn dgv_col_after = new DataGridViewImageColumn();
                dgv_col_after.Name = "After_" + (i + 1).ToString();
                dgv_col_after.HeaderText = "After_" + (i + 1).ToString();
                dgv_edit_data.Columns.Add(dgv_col_after);
            }


            int id = 1;
            foreach (var tbl_pcs_edit in table_edit)
            {
                string pcs = tbl_pcs_edit.Key;

                DataGridViewRow dgvRow = new DataGridViewRow();
                dgv_edit_data.Rows.Add(dgvRow);

                int idx = 1;

                //dgv_edit_data.Rows[id - 1].Cells[0].Value = id;
                dgv_edit_data.Rows[id - 1].Cells[0].Value = pcs;


                foreach (var lst_data_pcs in tbl_pcs_edit.Value)
                {
                    dgv_edit_data.Rows[id - 1].Cells[idx].Value = myCode.byteArrayToImage(lst_data_pcs.Value);
                    idx++;
                }

                id++;


            }
            if (dgv_edit_data.Rows.Count > 0)
            {
                for (int i = 1; i < dgv_edit_data.Columns.Count; i++)
                {
                    ((DataGridViewImageColumn)dgv_edit_data.Columns[i]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)dgv_edit_data.Columns[i]).Width = 100;

                    foreach (DataGridViewRow dr in dgv_edit_data.Rows)
                    {
                        if (myCode.checkDBNull(dr.Cells[i].Value) == "")
                        {
                            dr.Cells[i].Value = myCode.byteArrayToImage(img_null);
                        }
                        dr.Height = 75;
                    }
                }

            }


        }
        public void fill_table_image_thermalstress()
        {

            dgv_edit_data.Columns.Clear();
            dgv_edit_data.Rows.Clear();

            dgv_edit_data.Columns.Add("ID", "ID");
            dgv_edit_data.Columns.Add("Zone", "Zone");
            dgv_edit_data.Columns.Add("PCS", "PCS");


            DataGridViewImageColumn dgv_col = new DataGridViewImageColumn();
            dgv_col.Name = "Before";
            dgv_col.HeaderText = "Before";
            dgv_edit_data.Columns.Add(dgv_col);
            int max_count_edit = 0;
            Dictionary<string, Dictionary<string, Dictionary<string, Byte[]>>> table_edit = Get_table_edit_image_Thermal_stress(ref max_count_edit);
            for (int i = 0; i < max_count_edit; i++)
            {

                DataGridViewImageColumn dgv_col_after = new DataGridViewImageColumn();
                dgv_col_after.Name = "After_" + (i + 1).ToString();
                dgv_col_after.HeaderText = "After_" + (i + 1).ToString();
                dgv_edit_data.Columns.Add(dgv_col_after);
            }


            int id = 1;
            foreach (var tbl_region_edit in table_edit)
            {
                string region = tbl_region_edit.Key;
                foreach (var tbl_pcs_edit in tbl_region_edit.Value)
                {
                    string pcs = tbl_pcs_edit.Key;

                    DataGridViewRow dgvRow = new DataGridViewRow();
                    dgv_edit_data.Rows.Add(dgvRow);

                    int idx = 3;

                    dgv_edit_data.Rows[id - 1].Cells[0].Value = id;
                    dgv_edit_data.Rows[id - 1].Cells[1].Value = region;
                    dgv_edit_data.Rows[id - 1].Cells[2].Value = pcs;

                    foreach (var lst_data_pcs in tbl_pcs_edit.Value)
                    {
                        dgv_edit_data.Rows[id - 1].Cells[idx].Value = myCode.byteArrayToImage(lst_data_pcs.Value);
                        idx++;
                    }

                    id++;
                }


            }
            if (dgv_edit_data.Rows.Count > 0)
            {
                for (int i = 3; i < dgv_edit_data.Columns.Count; i++)
                {
                    ((DataGridViewImageColumn)dgv_edit_data.Columns[i]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)dgv_edit_data.Columns[i]).Width = 100;

                    foreach (DataGridViewRow dr in dgv_edit_data.Rows)
                    {
                        if (myCode.checkDBNull(dr.Cells[i].Value) == "")
                        {
                            dr.Cells[i].Value = myCode.byteArrayToImage(img_null);
                        }
                        dr.Height = 75;
                    }
                }

            }

            // DGV_data_image.DataSource = dt_edit_history;
            //tbl_data = dt_edit_history;
        }
        public void fill_table_image()
        {
            // DataTable dt_edit_history = new DataTable();
            //dt_edit_history.Columns.Add("ID");
            //dt_edit_history.Columns.Add("Zone");
            //dt_edit_history.Columns.Add("Pcs");
            //dt_edit_history.Columns.Add("Before");

            DGV_data_image.Columns.Clear();
            DGV_data_image.Rows.Clear();

            DGV_data_image.Columns.Add("ID", "ID");
            DGV_data_image.Columns.Add("Zone", "Zone");
            DGV_data_image.Columns.Add("PCS", "PCS");

            DataGridViewImageColumn dgv_col = new DataGridViewImageColumn();
            dgv_col.Name = "Before";
            dgv_col.HeaderText = "Before";
            DGV_data_image.Columns.Add(dgv_col);
            int max_count_edit = 0;
            Dictionary<string, Dictionary<int, Dictionary<string, Byte[]>>> table_edit = Get_table_edit_stackup_image(ref max_count_edit);
            for (int i = 0; i < max_count_edit; i++)
            {
                //dt_edit_history.Columns.Add("After_" + (i + 1).ToString());

                DataGridViewImageColumn dgv_col_after = new DataGridViewImageColumn();
                dgv_col_after.Name = "After_" + (i + 1).ToString();
                dgv_col_after.HeaderText = "After_" + (i + 1).ToString();
                DGV_data_image.Columns.Add(dgv_col_after);
            }


            int id = 1;
            foreach (var tbl_zone_edit in table_edit)
            {
                string zone = tbl_zone_edit.Key;
                foreach (var tbl_pcs_edit in tbl_zone_edit.Value)
                {

                    int pcs = tbl_pcs_edit.Key;

                    DataGridViewRow dgvRow = new DataGridViewRow();
                    DGV_data_image.Rows.Add(dgvRow);

                    int idx = 3;

                    DGV_data_image.Rows[id - 1].Cells[0].Value = id;
                    DGV_data_image.Rows[id - 1].Cells[1].Value = zone;
                    DGV_data_image.Rows[id - 1].Cells[2].Value = pcs;

                    foreach (var lst_data_pcs in tbl_pcs_edit.Value)
                    {
                        DGV_data_image.Rows[id - 1].Cells[idx].Value = myCode.byteArrayToImage(lst_data_pcs.Value);
                        idx++;
                    }


                    // dt_edit_history.Rows.Add(dr);
                    id++;

                    // DGV_data_image.Rows.Add(dgvRow);
                }

            }
            if (DGV_data_image.Rows.Count > 0)
            {
                for (int i = 3; i < DGV_data_image.Columns.Count; i++)
                {

                    ((DataGridViewImageColumn)DGV_data_image.Columns[i]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)DGV_data_image.Columns[i]).Width = 100;

                    foreach (DataGridViewRow dr in DGV_data_image.Rows)
                    {
                        if (myCode.checkDBNull(dr.Cells[i].Value) == "")
                        {
                            dr.Cells[i].Value = myCode.byteArrayToImage(img_null);
                        }
                        dr.Height = 75;
                    }
                }

            }

            // DGV_data_image.DataSource = dt_edit_history;
            //tbl_data = dt_edit_history;
        }
        public void fill_table_image_BVH_PTH(string process)
        {

            DGV_data_image.Columns.Clear();
            DGV_data_image.Rows.Clear();

            DGV_data_image.Columns.Add("ID", "ID");
            DGV_data_image.Columns.Add("PCS", "PCS");


            DataGridViewImageColumn dgv_col = new DataGridViewImageColumn();
            dgv_col.Name = "Before";
            dgv_col.HeaderText = "Before";
            DGV_data_image.Columns.Add(dgv_col);
            int max_count_edit = 0;
            Dictionary<int, Dictionary<string, List<byte[]>>> table_edit = Get_table_edit_BVH_PTH_Image(process, ref max_count_edit);
            for (int i = 0; i < max_count_edit; i++)
            {

                DataGridViewImageColumn dgv_col_after = new DataGridViewImageColumn();
                dgv_col_after.Name = "After_" + (i + 1).ToString();
                dgv_col_after.HeaderText = "After_" + (i + 1).ToString();
                DGV_data_image.Columns.Add(dgv_col_after);
            }


            int id = 1;
            foreach (var tbl_pcs_edit in table_edit)
            {
                int pcs = tbl_pcs_edit.Key;

                //  DataGridViewRow dgvRow = new DataGridViewRow();
                DGV_data_image.Rows.Add(new DataGridViewRow());
                DGV_data_image.Rows.Add(new DataGridViewRow());

                int idx = 2;


                foreach (var lst_time_edit in tbl_pcs_edit.Value)
                {

                    DGV_data_image.Rows[id - 1].Cells[0].Value = id;
                    DGV_data_image.Rows[id - 1].Cells[1].Value = pcs;
                    DGV_data_image.Rows[id].Cells[0].Value = id + 1;
                    DGV_data_image.Rows[id].Cells[1].Value = pcs;
                    DGV_data_image.Rows[id - 1].Cells[idx].Value = myCode.byteArrayToImage(lst_time_edit.Value[0]);
                    DGV_data_image.Rows[id].Cells[idx].Value = myCode.byteArrayToImage(lst_time_edit.Value[1]);
                    idx++;

                }
                id = id + 2;


            }
            if (DGV_data_image.Rows.Count > 0)
            {
                for (int i = 2; i < DGV_data_image.Columns.Count; i++)
                {
                    ((DataGridViewImageColumn)DGV_data_image.Columns[i]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)DGV_data_image.Columns[i]).Width = 100;

                    foreach (DataGridViewRow dr in DGV_data_image.Rows)
                    {
                        if (myCode.checkDBNull(dr.Cells[i].Value) == "")
                        {
                            dr.Cells[i].Value = myCode.byteArrayToImage(img_null);
                        }
                        dr.Height = 75;
                    }
                }

            }

            // DGV_data_image.DataSource = dt_edit_history;
            //tbl_data = dt_edit_history;
        }
        public void fill_table_image_Impedance(string process, string col_name_region, string col_name_image, string tbl_name)
        {

            DGV_data_image.Columns.Clear();
            DGV_data_image.Rows.Clear();

            DGV_data_image.Columns.Add("ID", "ID");
            DGV_data_image.Columns.Add("Region", "Region");
            DGV_data_image.Columns.Add("PCS", "PCS");


            DataGridViewImageColumn dgv_col = new DataGridViewImageColumn();
            dgv_col.Name = "Before";
            dgv_col.HeaderText = "Before";
            DGV_data_image.Columns.Add(dgv_col);
            int max_count_edit = 0;
            Dictionary<string, Dictionary<string, Dictionary<string, byte[]>>> table_edit = Get_table_edit_Impedance_Image(process, col_name_region, col_name_image, tbl_name, ref max_count_edit);
            for (int i = 0; i < max_count_edit; i++)
            {

                DataGridViewImageColumn dgv_col_after = new DataGridViewImageColumn();
                dgv_col_after.Name = "After_" + (i + 1).ToString();
                dgv_col_after.HeaderText = "After_" + (i + 1).ToString();
                DGV_data_image.Columns.Add(dgv_col_after);
            }


            int id = 1;

            foreach (var tbl_zone_edit in table_edit)
            {
                string zone = tbl_zone_edit.Key;
                foreach (var tbl_pcs_edit in tbl_zone_edit.Value)
                {

                    string pcs = tbl_pcs_edit.Key;
                    DataGridViewRow dgvRow = new DataGridViewRow();
                    DGV_data_image.Rows.Add(dgvRow);

                    int idx = 3;

                    DGV_data_image.Rows[id - 1].Cells[0].Value = id;
                    DGV_data_image.Rows[id - 1].Cells[1].Value = zone;
                    DGV_data_image.Rows[id - 1].Cells[2].Value = pcs;

                    foreach (var lst_data_pcs in tbl_pcs_edit.Value)
                    {
                        DGV_data_image.Rows[id - 1].Cells[idx].Value = myCode.byteArrayToImage(lst_data_pcs.Value);
                        idx++;
                    }
                    id++;
                }

            }
            if (DGV_data_image.Rows.Count > 0)
            {
                for (int i = 3; i < DGV_data_image.Columns.Count; i++)
                {
                    ((DataGridViewImageColumn)DGV_data_image.Columns[i]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)DGV_data_image.Columns[i]).Width = 100;

                    foreach (DataGridViewRow dr in DGV_data_image.Rows)
                    {
                        if (myCode.checkDBNull(dr.Cells[i].Value) == "")
                        {
                            dr.Cells[i].Value = myCode.byteArrayToImage(img_null);
                        }
                        dr.Height = 75;
                    }
                }

            }


        }
        public void fill_table_image_Impedance_region(string process, string col_name_region, string col_name_image, string tbl_name, string zone)
        {

            DGV_data_image.Columns.Clear();
            DGV_data_image.Rows.Clear();

            DGV_data_image.Columns.Add("ID", "ID");
            DGV_data_image.Columns.Add("Region", "Region");
            DGV_data_image.Columns.Add("PCS", "PCS");


            DataGridViewImageColumn dgv_col = new DataGridViewImageColumn();
            dgv_col.Name = "Before";
            dgv_col.HeaderText = "Before";
            DGV_data_image.Columns.Add(dgv_col);
            int max_count_edit = 0;
            Dictionary<string, Dictionary<string, Dictionary<string, byte[]>>> table_edit = Get_table_edit_Impedance_Image(process, col_name_region, col_name_image, tbl_name, ref max_count_edit);
            for (int i = 0; i < max_count_edit; i++)
            {

                DataGridViewImageColumn dgv_col_after = new DataGridViewImageColumn();
                dgv_col_after.Name = "After_" + (i + 1).ToString();
                dgv_col_after.HeaderText = "After_" + (i + 1).ToString();
                DGV_data_image.Columns.Add(dgv_col_after);
            }


            int id = 1;

            foreach (var tbl_zone_edit in table_edit)
            {
                string zone_ = tbl_zone_edit.Key;
                if (zone_ == zone)
                {
                    foreach (var tbl_pcs_edit in tbl_zone_edit.Value)
                    {

                        string pcs = tbl_pcs_edit.Key;

                        DataGridViewRow dgvRow = new DataGridViewRow();
                        DGV_data_image.Rows.Add(dgvRow);

                        int idx = 3;

                        DGV_data_image.Rows[id - 1].Cells[0].Value = id;
                        DGV_data_image.Rows[id - 1].Cells[1].Value = zone;
                        DGV_data_image.Rows[id - 1].Cells[2].Value = pcs;

                        foreach (var lst_data_pcs in tbl_pcs_edit.Value)
                        {
                            DGV_data_image.Rows[id - 1].Cells[idx].Value = myCode.byteArrayToImage(lst_data_pcs.Value);
                            idx++;
                        }
                        id++;

                    }
                    break;
                }


            }
            if (DGV_data_image.Rows.Count > 0)
            {
                for (int i = 3; i < DGV_data_image.Columns.Count; i++)
                {
                    ((DataGridViewImageColumn)DGV_data_image.Columns[i]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)DGV_data_image.Columns[i]).Width = 100;

                    foreach (DataGridViewRow dr in DGV_data_image.Rows)
                    {
                        if (myCode.checkDBNull(dr.Cells[i].Value) == "")
                        {
                            dr.Cells[i].Value = myCode.byteArrayToImage(img_null);
                        }
                        dr.Height = 75;
                    }
                }

            }


        }
        public void fill_table_image_BVH_PTH_pcs(string process, int pcs_)
        {

            DGV_data_image.Columns.Clear();
            DGV_data_image.Rows.Clear();

            DGV_data_image.Columns.Add("ID", "ID");
            DGV_data_image.Columns.Add("PCS", "PCS");


            DataGridViewImageColumn dgv_col = new DataGridViewImageColumn();
            dgv_col.Name = "Before";
            dgv_col.HeaderText = "Before";
            DGV_data_image.Columns.Add(dgv_col);
            int max_count_edit = 0;
            Dictionary<int, Dictionary<string, List<byte[]>>> table_edit = Get_table_edit_BVH_PTH_Image(process, ref max_count_edit);
            for (int i = 0; i < max_count_edit; i++)
            {

                DataGridViewImageColumn dgv_col_after = new DataGridViewImageColumn();
                dgv_col_after.Name = "After_" + (i + 1).ToString();
                dgv_col_after.HeaderText = "After_" + (i + 1).ToString();
                DGV_data_image.Columns.Add(dgv_col_after);
            }


            int id = 1;
            foreach (var tbl_pcs_edit in table_edit)
            {
                int pcs = tbl_pcs_edit.Key;
                if (pcs == pcs_)
                {
                    // DataGridViewRow dgvRow = new DataGridViewRow();
                    DGV_data_image.Rows.Add(new DataGridViewRow());
                    DGV_data_image.Rows.Add(new DataGridViewRow());

                    int idx = 2;


                    foreach (var lst_time_edit in tbl_pcs_edit.Value)
                    {

                        DGV_data_image.Rows[id - 1].Cells[0].Value = id;
                        DGV_data_image.Rows[id - 1].Cells[1].Value = pcs;
                        DGV_data_image.Rows[id].Cells[0].Value = id + 1;
                        DGV_data_image.Rows[id].Cells[1].Value = pcs;

                        DGV_data_image.Rows[id - 1].Cells[idx].Value = myCode.byteArrayToImage(lst_time_edit.Value[0]);
                        DGV_data_image.Rows[id].Cells[idx].Value = myCode.byteArrayToImage(lst_time_edit.Value[1]);
                        idx++;

                    }
                    id++;

                    //DataGridViewRow dgvRow_2 = new DataGridViewRow();
                    //DGV_data_image.Rows.Add(dgvRow_2);

                    //idx = 2;

                    //DGV_data_image.Rows[id - 1].Cells[0].Value = id;
                    //DGV_data_image.Rows[id - 1].Cells[1].Value = pcs;
                    //foreach (var tbl_time_edit in tbl_pcs_edit.Value)
                    //{
                    //    DGV_data_image.Rows[id - 1].Cells[idx].Value = myCode.byteArrayToImage((tbl_time_edit.Value).Values.ToArray()[0]);
                    //    idx++;

                    //}
                    //id++;
                    break;
                }



            }
            if (DGV_data_image.Rows.Count > 0)
            {
                for (int i = 2; i < DGV_data_image.Columns.Count; i++)
                {
                    ((DataGridViewImageColumn)DGV_data_image.Columns[i]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)DGV_data_image.Columns[i]).Width = 100;

                    foreach (DataGridViewRow dr in DGV_data_image.Rows)
                    {
                        if (myCode.checkDBNull(dr.Cells[i].Value) == "")
                        {
                            dr.Cells[i].Value = myCode.byteArrayToImage(img_null);
                        }
                        dr.Height = 75;
                    }
                }

            }

            // DGV_data_image.DataSource = dt_edit_history;
            //tbl_data = dt_edit_history;
        }
        public void Get_ListTable_image(int col_inx, DataTable myDt, string[] src_arr, ref List<DataTable> src_lst_tbl, string tar_item)
        {
            DataRow[] temp_dr;
            if (col_inx < src_arr.Length - 1)
            {
                if (src_arr[col_inx + 1] != "")
                {
                    string[] sel_val = myDt.AsEnumerable().Select(x => x.Field<Byte[]>(src_arr[col_inx + 1]).ToString()).Distinct().ToArray();
                    if (sel_val.Length != 0)
                    {
                        foreach (string sv in sel_val)
                        {
                            if (sv != null)
                            {
                                DataTable curTbl = myDt.AsEnumerable().Where(r => r.Field<string>(src_arr[col_inx + 1]).ToString() == sv).CopyToDataTable();
                                Get_ListTable(col_inx + 1, curTbl, src_arr, ref src_lst_tbl, tar_item);
                            }
                        }
                    }
                    else
                    {
                        temp_dr = myDt.AsEnumerable().Where(x => x.Field<Byte[]>(tar_item) != null).ToArray();
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
        public void Get_ListTable_Image(int col_inx, DataTable myDt, string[] src_arr, ref List<DataTable> src_lst_tbl, string tar_item)
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
                                Get_ListTable_Image(col_inx + 1, curTbl, src_arr, ref src_lst_tbl, tar_item);
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
        public void Get_ListTable_datetime(int col_inx, DataTable myDt, string[] src_arr, ref List<DataTable> src_lst_tbl, string tar_item)
        {
            DataRow[] temp_dr;
            if (col_inx < src_arr.Length - 1)
            {
                if (src_arr[col_inx + 1] != "")
                {
                    string[] sel_val = myDt.AsEnumerable().Select(x => x.Field<DateTime>(src_arr[col_inx + 1]).ToString()).Distinct().ToArray();
                    if (sel_val.Length != 0)
                    {
                        foreach (string sv in sel_val)
                        {
                            if (sv != null)
                            {
                                DataTable curTbl = myDt.AsEnumerable().Where(r => r.Field<DateTime>(src_arr[col_inx + 1]).ToString() == sv).CopyToDataTable();
                                Get_ListTable_datetime(col_inx + 1, curTbl, src_arr, ref src_lst_tbl, tar_item);
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
            catch { }
            try
            {


                DataTable dt_data = src_dt.AsEnumerable().Where(x => x.Field<string>("Region") != "Image_data").CopyToDataTable();
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
            catch { };
        }

        public void Split_Data_Image_BVH_PTH(DataTable src_dt, ref DataTable data_dt, ref DataTable img_dt)
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
                img_dt = dt_img.AsDataView().ToTable(false, new string[] { "PCS_No", "Region", "Before_Data", "After_Data", "Time_Update", "Operator", "Depart", "Remark" });
            }
            catch { }
            try
            {


                DataTable dt_data = src_dt.AsEnumerable().Where(x => x.Field<string>("Region") != "Image_data").CopyToDataTable();
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
            catch { };
        }
        public void Split_Data_Image_Impedance(DataTable src_dt, ref DataTable data_dt, ref DataTable img_dt)
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
                img_dt = dt_img.AsDataView().ToTable(false, new string[] { "PCS_No", "Region", "Zone", "Before_Data", "After_Data", "Time_Update", "Operator", "Depart", "Remark" });
            }
            catch { }
            try
            {


                DataTable dt_data = src_dt.AsEnumerable().Where(x => x.Field<string>("Region") != "Image_data").CopyToDataTable();
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
                data_dt = disp_dt.AsDataView().ToTable(false, new string[] { "PCS_No", "Region", "Zone", "Before_Data", "After_Data", "Time_Update", "Operator", "Depart", "Remark" });
            }
            catch { };
        }

        private void cb_pcs_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_zone.SelectedIndex != -1 && cb_pcs.SelectedIndex != -1)
            {
                dgv_edit_data.Columns.Clear();
                DGV_data_image.Columns.Clear();

                string pcs = cb_pcs.SelectedItem.ToString();

                if (sheet == "Stack-up")
                {
                    DataView dv = tbl_data.AsDataView();
                    string filter = TDMK_Code.filter_str(new string[] { "Zone", "PCS" }, new string[] { cb_zone.SelectedItem.ToString(), pcs });
                    dv.RowFilter = filter;
                    dgv_edit_data.DataSource = dv.ToTable();

                    for (int i = 0; i < dgv_edit_data.Rows.Count - 1; i++)
                    {
                        if (dgv_edit_data.Rows[i].Cells["ID"].Value.ToString() == "Total")
                        {
                            for (int j = 0; j < dgv_edit_data.ColumnCount; j++)
                            {
                                dgv_edit_data.Rows[i].Cells[j].Style.BackColor = Color.GreenYellow;
                            }
                        }
                    }
                    check_spec_stackup_detail();
                    filter_image_zone_pcs(cb_zone.SelectedItem.ToString(), pcs);
                }
                if (sheet == "BVH & PTH")
                {

                    DataView dv = tbl_data.AsDataView();
                    string filter = TDMK_Code.filter_str(new string[] { "PCS" }, new string[] { pcs });
                    dv.RowFilter = filter;
                    dgv_edit_data.DataSource = dv;

                    fill_table_image_BVH_PTH_pcs(cb_zone.SelectedItem.ToString(), int.Parse(pcs));
                }
                if (sheet == "Impedance")
                {

                    DataView dv = tbl_data.AsDataView();
                    string filter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { pcs });
                    dv.RowFilter = filter;

                    dgv_edit_data.DataSource = dv;
                    if (cb_zone.SelectedItem.ToString() == "Impedance")
                    {
                        check_spec_Impedance(dv.ToTable());
                        fill_table_image_Impedance_region(cb_zone.SelectedItem.ToString(), "Region", "Image_Graph", "IMPEDANCE_GRAPH", pcs);
                    }
                    else
                    if (cb_zone.SelectedItem.ToString() == "Tracewidth")
                    {
                        check_spec_Impedance(dv.ToTable());
                        fill_table_image_Impedance_region(cb_zone.SelectedItem.ToString(), "Data_For", "Image_Tracewidth", "TRACEWIDTH_IMAGE", pcs);
                    }

                }
                if (sheet == "CQRA - Thermal stress")
                {
                    fill_table_image_thermalstress_zone_pcs(cb_zone.SelectedItem.ToString(), cb_pcs.SelectedItem.ToString());
                }

            }

        }

        private void tableLayoutPanel3_Paint(object sender, PaintEventArgs e)
        {

        }
        public void fill_table_image_thermalstress_zone(string zone)
        {

            dgv_edit_data.Columns.Clear();
            dgv_edit_data.Rows.Clear();

            dgv_edit_data.Columns.Add("ID", "ID");
            dgv_edit_data.Columns.Add("Zone", "Zone");
            dgv_edit_data.Columns.Add("Pcs", "Pcs");


            DataGridViewImageColumn dgv_col = new DataGridViewImageColumn();
            dgv_col.Name = "Before";
            dgv_col.HeaderText = "Before";
            dgv_edit_data.Columns.Add(dgv_col);
            int max_count_edit = 0;
            Dictionary<string, Dictionary<string, Dictionary<string, Byte[]>>> table_edit = Get_table_edit_image_Thermal_stress(ref max_count_edit);
            for (int i = 0; i < max_count_edit; i++)
            {

                DataGridViewImageColumn dgv_col_after = new DataGridViewImageColumn();
                dgv_col_after.Name = "After_" + (i + 1).ToString();
                dgv_col_after.HeaderText = "After_" + (i + 1).ToString();
                dgv_edit_data.Columns.Add(dgv_col_after);
            }


            int id = 1;
            foreach (var tbl_region_edit in table_edit)
            {
                string region = tbl_region_edit.Key;
                if (region == zone)
                {
                    foreach (var tbl_pcs_edit in tbl_region_edit.Value)
                    {
                        string pcs = tbl_pcs_edit.Key; 
                        DataGridViewRow dgvRow = new DataGridViewRow();
                        dgv_edit_data.Rows.Add(dgvRow);

                        int idx = 3; 
                        dgv_edit_data.Rows[id - 1].Cells[0].Value = id;
                        dgv_edit_data.Rows[id - 1].Cells[1].Value = region;
                        dgv_edit_data.Rows[id - 1].Cells[2].Value = pcs;

                        foreach (var lst_data_pcs in tbl_pcs_edit.Value)
                        {
                            dgv_edit_data.Rows[id - 1].Cells[idx].Value = myCode.byteArrayToImage(lst_data_pcs.Value);
                            idx++;
                        } 
                        id++;
                    }
                    break;
                }
               
            }
            if (dgv_edit_data.Rows.Count > 0)
            {
                for (int i = 3; i < dgv_edit_data.Columns.Count; i++)
                {
                    ((DataGridViewImageColumn)dgv_edit_data.Columns[i]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)dgv_edit_data.Columns[i]).Width = 100;

                    foreach (DataGridViewRow dr in dgv_edit_data.Rows)
                    {
                        dr.Height = 75;
                    }
                }

            }

            // DGV_data_image.DataSource = dt_edit_history;
            //tbl_data = dt_edit_history;
        }
        public void fill_table_image_thermalstress_zone_pcs(string zone, string _pcs)
        {

            dgv_edit_data.Columns.Clear();
            dgv_edit_data.Rows.Clear();

            dgv_edit_data.Columns.Add("ID", "ID");
            dgv_edit_data.Columns.Add("Zone", "Zone");
            dgv_edit_data.Columns.Add("Pcs", "Pcs");


            DataGridViewImageColumn dgv_col = new DataGridViewImageColumn();
            dgv_col.Name = "Before";
            dgv_col.HeaderText = "Before";
            dgv_edit_data.Columns.Add(dgv_col);
            int max_count_edit = 0;
            Dictionary<string, Dictionary<string, Dictionary<string, Byte[]>>> table_edit = Get_table_edit_image_Thermal_stress(ref max_count_edit);
            for (int i = 0; i < max_count_edit; i++)
            {

                DataGridViewImageColumn dgv_col_after = new DataGridViewImageColumn();
                dgv_col_after.Name = "After_" + (i + 1).ToString();
                dgv_col_after.HeaderText = "After_" + (i + 1).ToString();
                dgv_edit_data.Columns.Add(dgv_col_after);
            }


            int id = 1;
            foreach (var tbl_region_edit in table_edit)
            {
                string region = tbl_region_edit.Key;
                if (region == zone)
                {
                    foreach (var tbl_pcs_edit in tbl_region_edit.Value)
                    {
                        string pcs = tbl_pcs_edit.Key;
                        if (pcs == _pcs)
                        {

                            DataGridViewRow dgvRow = new DataGridViewRow();
                            dgv_edit_data.Rows.Add(dgvRow);

                            int idx = 3;
                            dgv_edit_data.Rows[id - 1].Cells[0].Value = id;
                            dgv_edit_data.Rows[id - 1].Cells[1].Value = region;
                            dgv_edit_data.Rows[id - 1].Cells[2].Value = pcs;

                            foreach (var lst_data_pcs in tbl_pcs_edit.Value)
                            {
                                dgv_edit_data.Rows[id - 1].Cells[idx].Value = myCode.byteArrayToImage(lst_data_pcs.Value);
                                idx++;
                            }
                            id++;
                            break;
                        }
                    }
                    break;
                }

            }
            if (dgv_edit_data.Rows.Count > 0)
            {
                for (int i = 3; i < dgv_edit_data.Columns.Count; i++)
                {
                    ((DataGridViewImageColumn)dgv_edit_data.Columns[i]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)dgv_edit_data.Columns[i]).Width = 100;

                    foreach (DataGridViewRow dr in dgv_edit_data.Rows)
                    {
                        dr.Height = 75;
                    }
                }

            }

            // DGV_data_image.DataSource = dt_edit_history;
            //tbl_data = dt_edit_history;
        }

        private void cb_zone_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_zone.SelectedIndex != -1)
            {
                cb_pcs.Text = "ALL";
                cb_pcs.Items.Clear();
                dgv_edit_data.Columns.Clear();
                if(sheet != "CQRA - Thermal stress")
                {
                    DGV_data_image.Columns.Clear();
                }
             

                if (sheet == "Stack-up")
                {

                    string zone = cb_zone.SelectedItem.ToString();
                    DataView dv = tbl_data.AsDataView();
                    string filter = TDMK_Code.filter_str(new string[] { "Zone" }, new string[] { zone });
                    dv.RowFilter = filter;

                    dgv_edit_data.DataSource = dv.ToTable();



                    for (int i = 0; i < dgv_edit_data.Rows.Count - 1; i++)
                    {
                        if (dgv_edit_data.Rows[i].Cells["ID"].Value.ToString() == "Total")
                        {
                            for (int j = 0; j < dgv_edit_data.ColumnCount; j++)
                            {
                                dgv_edit_data.Rows[i].Cells[j].Style.BackColor = Color.GreenYellow;
                            }
                        }
                    }

                    DataTable dt_zone = (DataTable)dgv_edit_data.DataSource;

                    //cb_pcs.Items.Clear();
                    string[] arr_pcs = dt_zone.AsEnumerable().Select(x => x.Field<string>("Pcs")).Distinct().ToArray();
                    cb_pcs.Items.Clear();
                    foreach (string item in arr_pcs)
                    {
                        cb_pcs.Items.Add(item);
                    }
                    cb_pcs.Text = "ALL";
                    check_spec_stackup_detail();
                    filter_image_zone(zone);

                }
                if (sheet == "BVH & PTH")
                {
                    DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, cb_zone.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
                    if (dt_data_summary.Rows.Count > 0)
                    {
                        DataTable dt_edit_history = new DataTable();
                        dt_edit_history.Columns.Add("ID");
                        dt_edit_history.Columns.Add("PCS");
                        dt_edit_history.Columns.Add("Region");
                        dt_edit_history.Columns.Add("Before");
                        int max_count_edit = 0;
                        Dictionary<int, Dictionary<string, Dictionary<string, string>>> table_edit = Get_table_edit_BVH_PTH(cb_zone.SelectedItem.ToString(), ref max_count_edit);
                        for (int i = 0; i < max_count_edit; i++)
                        {
                            dt_edit_history.Columns.Add("After_" + (i + 1).ToString());
                        }

                        List<int> count_edit = new List<int> { };
                        int id = 1;

                        foreach (var tbl_pcs_edit in table_edit)
                        {
                            int r_pcs = tbl_pcs_edit.Value["Before"].Count;
                            int pcs = tbl_pcs_edit.Key;


                            Dictionary<string, string> dic_region = new Dictionary<string, string>();
                            dic_region = tbl_pcs_edit.Value["Before"];
                            for (int i = 0; i < r_pcs; i++)
                            {
                                DataRow dr = dt_edit_history.NewRow();
                                int idx = 3;

                                dr[0] = id;
                                dr[1] = pcs;
                                dr[2] = dic_region.Keys.ToArray()[i];

                                foreach (var lst_data_pcs in tbl_pcs_edit.Value)
                                {
                                    dr[idx] = lst_data_pcs.Value[dic_region.Keys.ToArray()[i]];
                                    idx++;
                                }

                                dt_edit_history.Rows.Add(dr);
                                id++;
                            }

                        }

                        dgv_edit_data.DataSource = dt_edit_history;
                        tbl_data = dt_edit_history;
                        string[] arr_zone = tbl_data.AsEnumerable().Select(x => x.Field<string>("PCS")).Distinct().ToArray();
                        cb_pcs.Items.Clear();
                        foreach (string item in arr_zone)
                        {
                            cb_pcs.Items.Add(item);
                        }
                        Check_spec_BVH_PTH(cb_zone.SelectedItem.ToString());
                        fill_table_image_BVH_PTH(cb_zone.SelectedItem.ToString());
                    }
                    else
                    { 
                        MessageBox.Show("No data", "Warning");
                    }


                }

                if (sheet == "Impedance")
                { 
                    DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, cb_zone.SelectedItem.ToString().ToUpper() + "_VAL", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
                    if (dt_data_summary.Rows.Count > 0)
                    {

                        DataTable dt_edit_history = new DataTable();
                        dt_edit_history.Columns.Add("ID");
                        dt_edit_history.Columns.Add("Region");
                        dt_edit_history.Columns.Add("PCS");
                        dt_edit_history.Columns.Add("Before");
                        int max_count_edit = 0;

                        if (cb_zone.SelectedItem.ToString() == "Impedance")
                        {

                            Dictionary<string, Dictionary<string, Dictionary<string, string>>> table_edit = Get_table_edit_Impedance_data(cb_zone.SelectedItem.ToString(), ref max_count_edit);
                            for (int i = 0; i < max_count_edit; i++)
                            {
                                dt_edit_history.Columns.Add("After_" + (i + 1).ToString());
                            }
                            List<int> count_edit = new List<int> { };
                            int id = 1;


                            foreach (var tbl_region_edit in table_edit)
                            {
                                string region = tbl_region_edit.Key;


                                Dictionary<string, string> dic_region = new Dictionary<string, string>();
                                dic_region = tbl_region_edit.Value["Before"];
                                for (int i = 0; i < dic_region.Count; i++)
                                {
                                    DataRow dr = dt_edit_history.NewRow();
                                    int idx = 3;

                                    dr[0] = id;
                                    dr[1] = region;
                                    dr[2] = dic_region.Keys.ToArray()[i];

                                    foreach (var lst_data_pcs in tbl_region_edit.Value)
                                    {
                                        double a = Math.Round(Double.Parse(lst_data_pcs.Value[dic_region.Keys.ToArray()[i]]),3);
                                        dr[idx] = a;
                                        idx++;
                                    }

                                    dt_edit_history.Rows.Add(dr);
                                    id++;
                                }

                            }
                            dgv_edit_data.DataSource = dt_edit_history;
                            tbl_data = dt_edit_history;
                            check_spec_Impedance(tbl_data);
                            string[] arr_zone = tbl_data.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();
                            cb_pcs.Items.Clear();
                            foreach (string item in arr_zone)
                            {
                                cb_pcs.Items.Add(item);
                            }

                            fill_table_image_Impedance(cb_zone.SelectedItem.ToString(), "Region", "Image_Graph", "IMPEDANCE_GRAPH");
                        }

                        if (cb_zone.SelectedItem.ToString() == "Tracewidth")
                        {
                            Dictionary<string, Dictionary<string, Dictionary<string, List<string>>>> table_edit = Get_table_edit_Tracewidth_data(ref max_count_edit);
                            for (int i = 0; i < max_count_edit; i++)
                            {
                                dt_edit_history.Columns.Add("After_" + (i + 1).ToString());
                            }
                            List<int> count_edit = new List<int> { };
                            int id = 1;

                            foreach (var tbl_zone_edit in table_edit)
                            {
                                string zone = tbl_zone_edit.Key;
                                foreach (var tbl_pcs_edit in tbl_zone_edit.Value)
                                {
                                    for (int i = 0; i < 2; i++)
                                    {
                                        string pcs = tbl_pcs_edit.Key;

                                        DataRow dr = dt_edit_history.NewRow();

                                        int idx = 3;

                                        dr[0] = id;
                                        dr[1] = zone;
                                        dr[2] = pcs;

                                        foreach (var lst_data_pcs in tbl_pcs_edit.Value)
                                        {

                                            double a = Math.Round(Double.Parse(lst_data_pcs.Value[i]), 3);
                                            dr[idx] = a;

                                            //dr[idx] = lst_data_pcs.Value[i];
                                            idx++;
                                        }
                                        id++;
                                        dt_edit_history.Rows.Add(dr);
                                    }
                                }
                            }


                            dgv_edit_data.DataSource = dt_edit_history;
                            tbl_data = dt_edit_history;
                            check_spec_Tracewidth(tbl_data);
                            string[] arr_zone = tbl_data.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();
                            cb_pcs.Items.Clear();
                            foreach (string item in arr_zone)
                            {
                                cb_pcs.Items.Add(item);
                            }

                            fill_table_image_Impedance(cb_zone.SelectedItem.ToString(), "Data_For", "Image_Tracewidth", "TRACEWIDTH_IMAGE");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No data", "Warning");
                    }

                }

                if (sheet == "CQRA - Thermal stress")
                { 
                    string zone = cb_zone.SelectedItem.ToString(); 
                    fill_table_image_thermalstress_zone(zone);
                    //DataTable dt_zone = (DataTable)dgv_edit_data.DataSource;
                    DataTable dt_zone = myCode.DGV_To_Table(dgv_edit_data);
                    if (dt_zone.Rows.Count > 0)
                    {
                        string[] arr_pcs = dt_zone.AsEnumerable().Select(x => x.Field<string>("PCS")).Distinct().ToArray();
                        cb_pcs.Items.Clear();
                        //foreach (string item in arr_pcs)
                        //{
                        //    cb_pcs.Items.Add(item);
                        //}
                        cb_pcs.Text = "ALL";

                       
                    }
                }

                if (sheet == "CQRA - bHast")
                {
                    DataTable bHast_dt = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_BHAST_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcodevh, lotnovh }));
                    if (tbl_data.Rows.Count > 0)
                    {

                        string time_modified = cb_zone.SelectedItem.ToString();
                        List<DataTable> tbl_modf = new List<DataTable> { };
                        Get_ListTable(-1, tbl_data, new string[] { "Remark" }, ref tbl_modf, "Time");
                        if (time_modified == "Before")
                        {

                            foreach (DataTable dt in tbl_modf)
                            {
                                string remark = dt.Rows[0]["Remark"].ToString();
                                
                                if (remark.Contains("Modified: 1"))
                                {
                                    dgv_edit_data.DataSource = dt;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            foreach (DataTable dt in tbl_modf)
                            {
                                string remark = dt.Rows[0]["Remark"].ToString();

                                if (remark.Contains("Modified: " + (int.Parse(time_modified.Split('_')[1]) + 1).ToString()))
                                {
                                    dgv_edit_data.DataSource = dt;
                                    break;
                                }
                            }
                        }
                    }
                    else if (bHast_dt.Rows.Count > 0)
                    {
                        dgv_edit_data.DataSource = bHast_dt;
                    }
                    else
                    {
                        MessageBox.Show("No data", "Warning");
                    }

                    DataTable dt_data = (DataTable)dgv_edit_data.DataSource;
                    if(dgv_edit_data.Rows.Count > 0)
                    {
                        bHast_Graph f1 = new bHast_Graph(bHast_dt, itemcodevh, lotnovh);
                        f1.Show();
                    }
                   
                }

                if (sheet == "ACF")
                {

                    string sample = cb_zone.SelectedItem.ToString();
                    DataView dv = tbl_data.AsDataView();
                    string filter = TDMK_Code.filter_str(new string[] { "Sample" }, new string[] { sample });
                    dv.RowFilter = filter;

                    dgv_edit_data.DataSource = dv.ToTable();
                    DataTable dt_sample = (DataTable)dgv_edit_data.DataSource;

                    cb_pcs.Text = "ALL";
                    fill_table_sample_Flatness(sample);

                }
            }
        }
        public void fill_table_sample_Flatness(string sample)
        {
            dgv_edit_data.AllowUserToAddRows = true;
            DataTable dt_edit_history = new DataTable();
            dt_edit_history.Columns.Add("ID");
            dt_edit_history.Columns.Add("Sample");
            dt_edit_history.Columns.Add("Point");
            dt_edit_history.Columns.Add("Before");
            int max_count_edit = 0;
            Dictionary<string, Dictionary<string, Dictionary<string, string>>> table_edit = Get_table_edit_ACFFlatness(ref max_count_edit);
            for (int i = 0; i < max_count_edit; i++)
            {
                dt_edit_history.Columns.Add("After_" + (i + 1).ToString());
            }


            Dictionary<string, Dictionary<string, string>> table_zone_edit = new Dictionary<string, Dictionary<string, string>> { };
            List<int> count_edit = new List<int> { };
            int id = 1;
            foreach (var tbl_zone_edit in table_edit)
            {
                string zone = tbl_zone_edit.Key;
                if(zone == sample)
                {
                    foreach (var tbl_pcs_edit in tbl_zone_edit.Value)
                    {
                        string pcs = tbl_pcs_edit.Key;

                        DataRow dr = dt_edit_history.NewRow();
                        int idx = 3;

                        dr[0] = id;
                        dr[1] = zone;
                        dr[2] = pcs;

                        foreach (var lst_data_pcs in tbl_pcs_edit.Value)
                        {
                            dr[idx] = lst_data_pcs.Value;
                            idx++;
                        }

                        dt_edit_history.Rows.Add(dr);
                        id++;
                    }
                    break;
                }
                
            }
        }

        private void tableLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if(cb_zone_history.SelectedIndex != -1)
            //{
            //    if (sheet == "Stack-up")
            //    {

            //        string zone = cb_zone_history.SelectedItem.ToString();

            //        DataView dv = tbl_history.AsDataView();
            //        string filter = TDMK_Code.filter_str(new string[] { "Zone" }, new string[] { zone });
            //        dv.RowFilter = filter;

            //        DGV_history.DataSource = dv.ToTable();
            //        DataTable dt_zone = (DataTable)DGV_history.DataSource;

            //        cb_pcs_history.Items.Clear();
            //        string[] arr_pcs = dt_zone.AsEnumerable().Select(x => x.Field<string>("PCS_No")).Distinct().ToArray();
            //        foreach (string item in arr_pcs)
            //        {
            //            cb_pcs_history.Items.Add(item);
            //        }
            //        cb_pcs_history.Text = "ALL";
            //        DataView dv_image = tbl_history_image.AsDataView();

            //        dv_image.RowFilter = filter;
            //        DGV_image_history.DataSource = dv_image.ToTable();

            //        if (DGV_image_history.Rows.Count > 0)
            //        {
            //            ((DataGridViewImageColumn)DGV_image_history.Columns["Before_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            //            ((DataGridViewImageColumn)DGV_image_history.Columns["Before_Data"]).Width = 100;
            //            ((DataGridViewImageColumn)DGV_image_history.Columns["After_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            //            ((DataGridViewImageColumn)DGV_image_history.Columns["After_Data"]).Width = 100;
            //            foreach (DataGridViewRow dr in DGV_image_history.Rows)
            //            {
            //                dr.Height = 75;
            //            }

            //        }

            //    }
            //    if(sheet == "BVH_PTH")
            //    {



            //        DataTable dt_history = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { itemcodevh, lotnovh, cb_zone_history.SelectedItem.ToString() }));

            //        DataTable dt_data = new DataTable();
            //        DataTable dt_img = new DataTable();
            //        Split_Data_Image_BVH_PTH(dt_history, ref dt_data, ref dt_img);

            //        DataTable dt_new = new DataTable();
            //        dt_new.Columns.Add("Edit times");
            //        foreach (DataColumn dc in dt_data.Columns)
            //        {
            //            dt_new.Columns.Add(dc.ColumnName);

            //        }


            //        List<DataTable> src_sample = new List<DataTable>();
            //        Get_ListTable(-1, dt_data, new string[] { "Pcs_No" }, ref src_sample, "Before_Data");


            //        foreach (DataTable tbl in src_sample)
            //        {

            //            List<DataTable> src_zone_pcs = new List<DataTable>();
            //            Get_ListTable_datetime(-1, tbl, new string[] { "Time_Update" }, ref src_zone_pcs, "Before_Data");


            //            int edit_time = 1;
            //            foreach (DataTable dt in src_zone_pcs)
            //            {
            //                foreach (DataRow dr in dt.Rows)
            //                {
            //                    DataRow dtrow = dt_new.NewRow();
            //                    dtrow[0] = edit_time;
            //                    for (int i = 0; i < dt.Columns.Count; i++)
            //                    {
            //                        dtrow[i + 1] = dr[i];
            //                    }
            //                    dt_new.Rows.Add(dtrow);

            //                }
            //                edit_time++;
            //            }
            //        }


            //        DGV_history.DataSource = dt_new;
            //        tbl_history = dt_new;
            //        DGV_image_history.DataSource = dt_img;
            //        tbl_history_image = dt_img;
            //        if (DGV_image_history.Rows.Count > 0)
            //        {
            //            ((DataGridViewImageColumn)DGV_image_history.Columns["Before_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            //            ((DataGridViewImageColumn)DGV_image_history.Columns["Before_Data"]).Width = 100;
            //            ((DataGridViewImageColumn)DGV_image_history.Columns["After_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            //            ((DataGridViewImageColumn)DGV_image_history.Columns["After_Data"]).Width = 100;
            //            foreach (DataGridViewRow dr in DGV_image_history.Rows)
            //            {
            //                dr.Height = 75;
            //            }

            //        }

            //        cb_pcs_history.Items.Clear();
            //        string[] arr_pcs = dt_new.AsEnumerable().Select(x => x.Field<string>("PCS_No")).Distinct().ToArray();
            //        foreach (string item in arr_pcs)
            //        {
            //            cb_pcs_history.Items.Add(item);
            //        }



            //        //string[] arr_zone = tbl_data.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();
            //        //foreach (string item in arr_zone)
            //        //{
            //        //    cb_zone.Items.Add(item);
            //        //}

            //        //string[] arr_zone_history = tbl_history.AsEnumerable().Select(x => x.Field<string>("Zone")).Distinct().ToArray();
            //        //foreach (string item in arr_zone_history)
            //        //{
            //        //    cb_zone_history.Items.Add(item);
            //        //}
            //    }

            //    if (sheet == "Impedance")
            //    {

            //        DataTable dt_history = TDMK_Code.Datatable_Filter(sqlcon, "VHX_Edit_History", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { itemcodevh, lotnovh, cb_zone_history.SelectedItem.ToString() }));

            //        DataTable dt_data = new DataTable();
            //        DataTable dt_img = new DataTable();
            //        Split_Data_Image_Impedance(dt_history, ref dt_data, ref dt_img);

            //        DataTable dt_new = new DataTable();
            //        dt_new.Columns.Add("Edit times");
            //        foreach (DataColumn dc in dt_data.Columns)
            //        {
            //            dt_new.Columns.Add(dc.ColumnName);

            //        }


            //        List<DataTable> src_sample = new List<DataTable>();
            //        Get_ListTable(-1, dt_data, new string[] { "Zone" }, ref src_sample, "Before_Data");


            //        foreach (DataTable tbl in src_sample)
            //        {

            //            List<DataTable> src_zone_pcs = new List<DataTable>();
            //            Get_ListTable_datetime(-1, tbl, new string[] { "Time_Update" }, ref src_zone_pcs, "Before_Data");


            //            int edit_time = 1;
            //            foreach (DataTable dt in src_zone_pcs)
            //            {
            //                foreach (DataRow dr in dt.Rows)
            //                {
            //                    DataRow dtrow = dt_new.NewRow();
            //                    dtrow[0] = edit_time;
            //                    for (int i = 0; i < dt.Columns.Count; i++)
            //                    {
            //                        dtrow[i + 1] = dr[i];
            //                    }
            //                    dt_new.Rows.Add(dtrow);

            //                }
            //                edit_time++;
            //            }
            //        }


            //        DGV_history.DataSource = dt_new;
            //        tbl_history = dt_new;
            //        DGV_image_history.DataSource = dt_img;
            //        tbl_history_image = dt_img;
            //        if (DGV_image_history.Rows.Count > 0)
            //        {
            //            ((DataGridViewImageColumn)DGV_image_history.Columns["Before_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            //            ((DataGridViewImageColumn)DGV_image_history.Columns["Before_Data"]).Width = 100;
            //            ((DataGridViewImageColumn)DGV_image_history.Columns["After_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            //            ((DataGridViewImageColumn)DGV_image_history.Columns["After_Data"]).Width = 100;
            //            foreach (DataGridViewRow dr in DGV_image_history.Rows)
            //            {
            //                dr.Height = 75;
            //            }

            //        }
            //        cb_pcs_history.Items.Clear();
            //        string[] arr_zone = dt_new.AsEnumerable().Select(x => x.Field<string>("Zone")).Distinct().ToArray();
            //        foreach (string item in arr_zone)
            //        {
            //            cb_pcs_history.Items.Add(item);
            //        }

            //        if (cb_zone_history.SelectedItem.ToString() == "Impedance")
            //        {
            //            foreach (string item in arr_zone)
            //            {
            //                cb_pcs_history.Items.Add("Impedance_" + item);
            //            }
            //        }
            //        else if (cb_zone_history.SelectedItem.ToString() == "Tracewidth")
            //        {
            //            foreach (string item in arr_zone)
            //            {
            //                cb_pcs_history.Items.Add("Tracewidth_" + item);
            //            }
            //        }



            //    }
            //}

        }

        private void cb_pcs_history_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (cb_zone_history.SelectedIndex != -1 && cb_zone_history.SelectedIndex != -1)
            //{
            //    if (sheet == "Stack-up")
            //    {


            //        string pcs = cb_pcs_history.SelectedItem.ToString();

            //        DataView dv = tbl_history.AsDataView();
            //        string filter = TDMK_Code.filter_str(new string[] { "Zone", "Pcs_No" }, new string[] { cb_zone_history.SelectedItem.ToString(), pcs });
            //        dv.RowFilter = filter;
            //        DGV_history.DataSource = dv.ToTable();

            //        DataView dv_image = tbl_history_image.AsDataView();
            //        dv_image.RowFilter = filter;
            //        DGV_image_history.DataSource = dv_image.ToTable();

            //        if (DGV_image_history.Rows.Count > 0)
            //        {
            //            ((DataGridViewImageColumn)DGV_image_history.Columns["Before_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            //            ((DataGridViewImageColumn)DGV_image_history.Columns["Before_Data"]).Width = 100;
            //            ((DataGridViewImageColumn)DGV_image_history.Columns["After_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            //            ((DataGridViewImageColumn)DGV_image_history.Columns["After_Data"]).Width = 100;
            //            foreach (DataGridViewRow dr in DGV_image_history.Rows)
            //            {
            //                dr.Height = 75;
            //            }

            //        }
            //    }
            //    if (sheet == "BVH_PTH")
            //    {


            //        string pcs = cb_pcs_history.SelectedItem.ToString();

            //        DataView dv = tbl_history.AsDataView();
            //        string filter = TDMK_Code.filter_str(new string[] {"Pcs_No" }, new string[] { pcs});
            //        dv.RowFilter = filter;
            //        DGV_history.DataSource = dv.ToTable();

            //        DataView dv_image = tbl_history_image.AsDataView();
            //        dv_image.RowFilter = filter;
            //        DGV_image_history.DataSource = dv_image.ToTable();

            //        if (DGV_image_history.Rows.Count > 0)
            //        {
            //            ((DataGridViewImageColumn)DGV_image_history.Columns["Before_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            //            ((DataGridViewImageColumn)DGV_image_history.Columns["Before_Data"]).Width = 100;
            //            ((DataGridViewImageColumn)DGV_image_history.Columns["After_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            //            ((DataGridViewImageColumn)DGV_image_history.Columns["After_Data"]).Width = 100;
            //            foreach (DataGridViewRow dr in DGV_image_history.Rows)
            //            {
            //                dr.Height = 75;
            //            }

            //        }
            //    }


            //}
        }

        private void DGV_history_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
            if (e.RowIndex < 1 || e.ColumnIndex < 0)
                return;

            if (sheet == "Stack-up")
            {
                if (IsTheSameCellValue(0, e.RowIndex, DGV_data_image) && IsTheSameCellValue(1, e.RowIndex, DGV_data_image) && IsTheSameCellValue(2, e.RowIndex, DGV_data_image) && IsTheSameCellValue(5, e.RowIndex, DGV_data_image))
                {
                    e.AdvancedBorderStyle.Top = DataGridViewAdvancedCellBorderStyle.None;
                }
                else
                {
                    e.AdvancedBorderStyle.Top = DGV_data_image.AdvancedCellBorderStyle.Top;
                }
            }
            if (sheet == "BVH_PTH")
            {
                if (IsTheSameCellValue(0, e.RowIndex, DGV_data_image) && IsTheSameCellValue(1, e.RowIndex, DGV_data_image) && IsTheSameCellValue(5, e.RowIndex, DGV_data_image))
                {
                    e.AdvancedBorderStyle.Top = DataGridViewAdvancedCellBorderStyle.None;
                }
                else
                {
                    e.AdvancedBorderStyle.Top = DGV_data_image.AdvancedCellBorderStyle.Top;
                }
            }
            if (sheet == "Impedance")
            {
                if (IsTheSameCellValue(0, e.RowIndex, DGV_data_image) && IsTheSameCellValue(2, e.RowIndex, DGV_data_image) && IsTheSameCellValue(3, e.RowIndex, DGV_data_image) && IsTheSameCellValue(6, e.RowIndex, DGV_data_image))
                {
                    e.AdvancedBorderStyle.Top = DataGridViewAdvancedCellBorderStyle.None;
                }
                else
                {
                    e.AdvancedBorderStyle.Top = DGV_data_image.AdvancedCellBorderStyle.Top;
                }
            }


        }

        private void DGV_history_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex == 0)
                return;
            if (sheet == "Stack-up")
            {
                if (IsTheSameCellValue(0, e.RowIndex, DGV_data_image) && IsTheSameCellValue(1, e.RowIndex, DGV_data_image) && IsTheSameCellValue(2, e.RowIndex, DGV_data_image) && IsTheSameCellValue(5, e.RowIndex, DGV_data_image))
                {
                    if (e.ColumnIndex != 3 && e.ColumnIndex != 4)
                    {
                        e.Value = "";
                        e.FormattingApplied = true;
                    }

                }
            }
            if (sheet == "BVH_PTH")
            {
                if (IsTheSameCellValue(0, e.RowIndex, DGV_data_image) && IsTheSameCellValue(1, e.RowIndex, DGV_data_image) && IsTheSameCellValue(5, e.RowIndex, DGV_data_image))
                {
                    if (e.ColumnIndex != 3 && e.ColumnIndex != 4 && e.ColumnIndex != 2)
                    {
                        e.Value = "";
                        e.FormattingApplied = true;
                    }

                }
            }
            if (sheet == "Impedance")
            {
                if (IsTheSameCellValue(0, e.RowIndex, DGV_data_image) && IsTheSameCellValue(2, e.RowIndex, DGV_data_image) && IsTheSameCellValue(3, e.RowIndex, DGV_data_image) && IsTheSameCellValue(6, e.RowIndex, DGV_data_image))
                {
                    if (e.ColumnIndex != 4 && e.ColumnIndex != 5 && e.ColumnIndex != 1)
                    {
                        e.Value = "";
                        e.FormattingApplied = true;
                    } 
                }
            }

        }

        private void txtItemCode_TextChanged(object sender, EventArgs e)
        {
            DGV_data_image.Columns.Clear();
            dgv_edit_data.Columns.Clear();
            dgv_edit_data.DataSource = DGV_data_image.DataSource = null;
            cb_zone.Items.Clear();
            cb_zone.Text = "";
            cb_pcs.Items.Clear();
            cb_pcs.Text = "";
            sheet = "";


          
        }

        private void txtLotNo_TextChanged(object sender, EventArgs e)
        {
            DGV_data_image.Columns.Clear();
            dgv_edit_data.Columns.Clear();
            dgv_edit_data.DataSource = DGV_data_image.DataSource = null;
            cb_zone.Items.Clear();
            cb_zone.Text = "";
            cb_pcs.Items.Clear();
            cb_pcs.Text = "";
            sheet = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //    DataTable dt_edit_history = new DataTable();
            //    dt_edit_history.Columns.Add("ID");            
            //    dt_edit_history.Columns.Add("Zone");
            //    dt_edit_history.Columns.Add("Pcs");
            //    dt_edit_history.Columns.Add("Before");
            //    int max_count_edit = 0;
            //    Dictionary<string, Dictionary<int, Dictionary<string, List<string>>>> table_edit = Get_table_edit(ref max_count_edit);
            //    for(int i = 0; i < max_count_edit; i++)
            //    {
            //        dt_edit_history.Columns.Add("After_" + (i+1).ToString());
            //    }

            //    Dictionary<int, Dictionary<string, List<string>>> table_zone_edit = new Dictionary<int, Dictionary<string, List<string>>> { };
            //    //table_zone_edit = table_edit["Zone_A"];
            //    List<int> count_edit = new List<int> {};
            //    int id = 1;
            //    foreach (var tbl_zone_edit in table_edit)
            //    {
            //        string zone = tbl_zone_edit.Key;
            //        foreach (var tbl_pcs_edit in tbl_zone_edit.Value)
            //        {
            //            int r_pcs = tbl_pcs_edit.Value["Before"].Count;
            //            int pcs = tbl_pcs_edit.Key;
            //            for (int i = 0; i < r_pcs; i++)
            //            {
            //                DataRow dr = dt_edit_history.NewRow();
            //                int idx = 3;

            //                dr[0] = id;                       
            //                dr[1] = zone;
            //                dr[2] = pcs;

            //                foreach (var lst_data_pcs in tbl_pcs_edit.Value)
            //                {
            //                    dr[idx] = lst_data_pcs.Value[i];
            //                    idx++;
            //                }

            //                dt_edit_history.Rows.Add(dr);
            //                id++;
            //            }

            //        }
            //    }
            //    dgv_edit_data.DataSource = dt_edit_history;

        }

        private void dgv_edit_data_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (sheet != "Solder Mask" && sheet != "CQRA - Thermal stress")
            {
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                if (e.RowIndex < 1 || e.ColumnIndex < 0)
                    return;
                if (sheet == "Stack-up")
                {

                    if (IsTheSameCellValue(1, e.RowIndex, dgv_edit_data) && IsTheSameCellValue(2, e.RowIndex, dgv_edit_data))
                    {
                        e.AdvancedBorderStyle.Top = DataGridViewAdvancedCellBorderStyle.None;
                    }
                    else
                    {
                        e.AdvancedBorderStyle.Top = DGV_data_image.AdvancedCellBorderStyle.Top;
                    }
                }
                if (sheet == "BVH & PTH" || sheet == "Impedance" || sheet == "ACF" || sheet == "CQRA - Moisture absorption")
                {
                    if (IsTheSameCellValue(1, e.RowIndex, dgv_edit_data))
                    {
                        e.AdvancedBorderStyle.Top = DataGridViewAdvancedCellBorderStyle.None;
                    }
                    else
                    {
                        e.AdvancedBorderStyle.Top = DGV_data_image.AdvancedCellBorderStyle.Top;
                    }
                }
            } 
        }

        private void dgv_edit_data_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex == 0)
                return;
            if (sheet == "Stack-up")
            {
                if (IsTheSameCellValue(1, e.RowIndex, dgv_edit_data) && IsTheSameCellValue(2, e.RowIndex, dgv_edit_data))
                {
                    if (e.ColumnIndex == 1 && e.ColumnIndex == 2)
                    {
                        e.Value = "";
                        e.FormattingApplied = true;
                    }

                }

            }
            if (sheet == "BVH & PTH")
            {
                if (IsTheSameCellValue(1, e.RowIndex, dgv_edit_data))
                {
                    if (e.ColumnIndex == 1)
                    {
                        e.Value = "";
                        e.FormattingApplied = true;
                    }

                }
            }
            if (sheet == "ACF_FLATNESS" || sheet == "CQRA - Moisture absorption")
            {
                if (IsTheSameCellValue(1, e.RowIndex, dgv_edit_data))
                {
                    if (e.ColumnIndex == 1)
                    {
                        e.Value = "";
                        e.FormattingApplied = true;
                    }

                }
            }

        }


        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }


        private void dgv_edit_data_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {

        }

        private void dgv_edit_data_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

            int col_inx = e.ColumnIndex;
            int r_inx = e.RowIndex;
            int row_inx = e.RowIndex;
           
            DataGridViewCell cur_cell = dgv_edit_data.CurrentCell;
            if (sheet == "Solder Mask" || sheet == "CQRA - Thermal stress")
            {
                if (dgv_edit_data.Columns[col_inx].Name.Contains("Before") || dgv_edit_data.Columns[col_inx].Name.Contains("After"))
                {
                    if (myCode.checkDBNull(cur_cell.Value) != "")
                    {
                        string region = "";
                        if (dgv_edit_data.Columns.Contains("Zone"))
                        {
                            region = dgv_edit_data.Rows[r_inx].Cells["Zone"].Value.ToString();
                        }
                        else if (dgv_edit_data.Columns.Contains("Region"))
                        {
                            region = dgv_edit_data.Rows[r_inx].Cells["Region"].Value.ToString();
                        }

                        string pcs = dgv_edit_data.Rows[r_inx].Cells["PCS"].Value.ToString();

                        View_detail_Image fr1 = new View_detail_Image();
                        fr1.data = (Bitmap)cur_cell.Value;
                        fr1.region_ = region;
                        fr1.pcs_ = pcs;
                        //fr1.choose_sheet_ = "Stackup";

                        fr1.Show();

                    }
                }
            }
           
        }



        public void check_spec_stackup()
        {
            DataTable dt_data = (DataTable)dgv_edit_data.DataSource;
            DataView dv = dt_data.AsDataView();
            string[] arr_zone = dt_data.AsEnumerable().Select(x => x.Field<string>("Zone").ToString()).Distinct().ToArray();
            tbl_spec = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode"}, new string[] { itemcodevh}));
            foreach (var zone in arr_zone)
            {
                DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Zone" }, new string[] { itemcodevh, zone.Replace("Zone_", "") }));
               
                if (dt_spec.Rows.Count > 0)
                {
                    Double USL = Double.Parse(dt_spec.Rows[0]["USL"].ToString());
                    Double LSL = Double.Parse(dt_spec.Rows[0]["LSL"].ToString());

                    string filter = TDMK_Code.filter_str(new string[] { "Zone" }, new string[] { zone });
                    dv.RowFilter = filter;
                    List<DataTable> tbl_pcs = new List<DataTable> { };
                    Get_ListTable(-1, dv.ToTable(), new string[] { "Pcs" }, ref tbl_pcs, "ID");

                    foreach (DataTable dt in tbl_pcs)
                    {
                        for (int i = 3; i < dgv_edit_data.ColumnCount; i++)
                        {
                            if (myCode.checkDBNull(dt.Rows[0][i]) != "")
                            {
                                Double total_pcs = 0;
                                foreach (DataRow dr in dt.Rows)
                                {
                                    total_pcs += Double.Parse(dr[i].ToString());
                                }
                                DataView dv_pcs = dt_data.AsDataView();
                                dv_pcs.RowFilter = TDMK_Code.filter_str(new string[] { "Zone", "Pcs" }, new string[] { zone, dt.Rows[0]["Pcs"].ToString() });
                                if (total_pcs > USL || total_pcs < LSL)
                                {
                                    for (int inx = 0; inx < dv_pcs.Count; inx++)
                                    {
                                        DataRow dr = dv_pcs[inx].Row;
                                        int r_inx = dt_data.Rows.IndexOf(dr);
                                        dgv_edit_data.Rows[r_inx].Cells[i].Style.BackColor = Color.Red;

                                    }
                                }
                            }
                        }
                    }
                }
            
            }
        
        }
        public void check_spec_stackup_detail()
        {
            DataTable dt_data = (DataTable)dgv_edit_data.DataSource;
            DataView dv = dt_data.AsDataView();
            string[] arr_zone = dt_data.AsEnumerable().Select(x => x.Field<string>("Zone").ToString()).Distinct().ToArray();
            tbl_spec = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { itemcodevh }));
            foreach (var zone in arr_zone)
            {
                if(zone != "")
                {
                    DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Zone" }, new string[] { itemcodevh, zone.Replace("Zone_", "") }));

                    if (dt_spec.Rows.Count > 0)
                    {
                        Double USL = Double.Parse(dt_spec.Rows[0]["USL"].ToString());
                        Double LSL = Double.Parse(dt_spec.Rows[0]["LSL"].ToString());

                        string[] arr_spec_detail = dt_spec.Rows[0]["Spec_detail"].ToString().Split(';');

                        string filter = TDMK_Code.filter_str(new string[] { "Zone" }, new string[] { zone });
                        dv.RowFilter = filter;
                        List<DataTable> tbl_pcs = new List<DataTable> { };
                        Get_ListTable(-1, dv.ToTable(), new string[] { "Pcs" }, ref tbl_pcs, "ID");
                        SortedDictionary<int, List<double>> dic_spec = new SortedDictionary<int, List<double>> { };
                        int k = 1;
                        foreach (string nominal_value in arr_spec_detail)
                        {
                            if (nominal_value != "")
                            {
                                List<double> lst_spec = new List<double> { };
                                double min = 0;
                                double max = 0;
                                if (nominal_value.Contains("+") || nominal_value.Contains("-"))
                                {
                                    string nominal_val = nominal_value.Replace("+", ";").Replace("-", ";").Replace("/", ";");
                                    max = double.Parse(nominal_val.Split(';')[0]) + double.Parse(nominal_val.Split(';')[1]);
                                    min = double.Parse(nominal_val.Split(';')[0]) - double.Parse(nominal_val.Split(';')[3]);

                                }
                                else if (nominal_value.Contains("±"))
                                {
                                    max = double.Parse(nominal_value.Split('±')[0]) + double.Parse(nominal_value.Split('±')[1]);
                                    min = double.Parse(nominal_value.Split('±')[0]) - double.Parse(nominal_value.Split('±')[1]);
                                }
                                else
                                {
                                    min = double.Parse(nominal_value) * 0.9;
                                    max = double.Parse(nominal_value) * 1.1;
                                }

                                lst_spec.Add(max);
                                lst_spec.Add(min);
                                dic_spec.Add(k, lst_spec);
                            }

                            k++;


                        }
                        foreach (DataTable dt in tbl_pcs)
                        {
                            DataView dv_pcs = dt_data.AsDataView();
                            dv_pcs.RowFilter = TDMK_Code.filter_str(new string[] { "Zone", "Pcs" }, new string[] { zone, dt.Rows[0]["Pcs"].ToString() });

                            for (int i = 3; i < dgv_edit_data.ColumnCount; i++)
                            {
                                int s = 1;
                                if (myCode.checkDBNull(dt.Rows[0][i]) != "")
                                {
                                    //Double total_pcs = 0;
                                    if (s <= dic_spec.Count)
                                    {
                                        double max = dic_spec[s][0];
                                        double min = dic_spec[s][1];
                                        foreach (DataRow dr in dt.Rows)
                                        {

                                            double value = Double.Parse(dr[i].ToString());

                                            if (value > max || value < min)
                                            {

                                                DataRow row = dv_pcs[s - 1].Row;
                                                int r_inx = dt_data.Rows.IndexOf(row);
                                                dgv_edit_data.Rows[r_inx].Cells[i].Style.BackColor = Color.Red;

                                            }
                                            
                                            s++;


                                        }
                                    }



                                    Double total_pcs = 0;
                                    foreach (DataRow dr in dt.Rows)
                                    {
                                        total_pcs += Double.Parse(dr[i].ToString());
                                    }
                                   
                                    if (total_pcs > USL || total_pcs < LSL)
                                    {
                                       
                                            DataRow dr = dv_pcs[dv_pcs.Count - 1].Row;
                                            int r_inx = dt_data.Rows.IndexOf(dr) + 1;
                                            dgv_edit_data.Rows[r_inx].Cells[i].Style.BackColor = Color.Red;

                                    }


                                }
                            }
                        }
                    }
                }
               

            }

        }
        public void Check_spec_BVH_PTH(string process)
        {
            for (int i = 0; i < dgv_edit_data.Columns.Count; i++)
            {
                for (int j = 0; j <dgv_edit_data.Rows.Count; j++)
                {
                    dgv_edit_data.Rows[j].Cells[i].Style.BackColor = Color.White;
                }
            }

            string[] item = { "ItemCode" };
            string[] item_val = { itemcodevh };
            DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_" + process.ToUpper(), TDMK_Code.filter_str(item, item_val));
            tbl_spec = dt_spec;
            if (dt_spec.Rows.Count > 0)
            {
                for (int i = 0; i < dgv_edit_data.Rows.Count - 1; i++)
                {
                    for (int k = 3; k < dgv_edit_data.Columns.Count; k++)
                    {
                        string region = dgv_edit_data.Rows[i].Cells["Region"].Value.ToString();
                        string USL = dt_spec.Rows[0][region].ToString().Replace("/", string.Empty);
                        string LSL = dt_spec.Rows[1][region].ToString().Replace("/", string.Empty);
                        if (dgv_edit_data.Rows[i].Cells[k].Value != null)
                        {
                            if (!USL.Contains("NA"))
                            {
                                if (Double.TryParse(dgv_edit_data.Rows[i].Cells[k].Value.ToString(), out Double data))
                                {
                                    if (data > Double.Parse(USL))
                                    {
                                        dgv_edit_data.Rows[i].Cells[k].Style.BackColor = Color.Red;
                                    }
                                }

                            }
                            if (!LSL.Contains("NA"))
                            {

                                if (Double.TryParse(dgv_edit_data.Rows[i].Cells[k].Value.ToString(), out Double data))
                                {
                                    if (data < Double.Parse(LSL))
                                    {
                                        dgv_edit_data.Rows[i].Cells[k].Style.BackColor = Color.Red;
                                    }
                                }

                            }
                        }

                    }
                }
            }



        }

        private void dgv_edit_data_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            //int col_inx = e.ColumnIndex;
            //int row_inx = e.RowIndex;
            //if (row_inx != -1 && row_inx != dgv_edit_data.Rows.Count - 1 && col_inx != -1 && dgv_edit_data.Rows.Count > 0)
            //{
            //    string select_sheet = sheet;
            //    if (select_sheet == "STACKUP")
            //    {

            //        string pcs = dgv_edit_data.Rows[row_inx].Cells["PCS"].Value.ToString();
            //        string zone = dgv_edit_data.Rows[row_inx].Cells["Zone"].Value.ToString();

            //        DataView dv = tbl_history.AsDataView();
            //        if (dv.Count > 0)
            //        {

            //            string filter = TDMK_Code.filter_str(new string[] { "Zone", "PCS_No" }, new string[] { zone, pcs });
            //            dv.RowFilter = filter;
            //            DataTable dt_filter = dv.ToTable();
            //            List<DataTable> tbl_time = new List<DataTable> { };
            //            Get_ListTable_datetime(-1, dt_filter, new string[] { "Time_Update" }, ref tbl_time, "Remark");

            //            string USL = "";
            //            string LSL = "";
            //            DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Zone" }, new string[] { itemcodevh, zone.Replace("Zone_", "") }));

            //            if (dt_spec.Rows.Count > 0)
            //            {
            //                USL = dt_spec.Rows[0]["USL"].ToString();
            //                LSL = dt_spec.Rows[0]["LSL"].ToString();
            //            }

            //            string colname_select = dgv_edit_data.Columns[col_inx].Name;
            //            if (colname_select.Contains("After_"))
            //            {
            //                int time_edit = int.Parse(colname_select.Replace("After_", ""));
            //                if (time_edit <= tbl_time.Count)
            //                {
            //                    DataTable tbl_select = tbl_time[time_edit - 1];
            //                    string depart = tbl_select.Rows[0]["Depart"].ToString();
            //                    string time_update = tbl_select.Rows[0]["Time_Update"].ToString();
            //                    string Operator = tbl_select.Rows[0]["Operator"].ToString();
            //                    string remark = tbl_select.Rows[0]["Remark"].ToString();

            //                    dgv_edit_data.CurrentCell.ToolTipText = "Depart: " + depart + " \nOperator: " + Operator + " \nLogfile Location: " + remark + " \nTime_Update: " + time_update + " \nUSL: " + USL + " \nLSL: " + LSL;
            //                }
            //            }
            //        }
            //    }

            //    if (select_sheet == "BVH_PTH")
            //    {
            //        string pcs = dgv_edit_data.Rows[row_inx].Cells["PCS"].Value.ToString();
            //        string zone = dgv_edit_data.Rows[row_inx].Cells["Region"].Value.ToString();

            //        DataView dv = tbl_history.AsDataView();

            //        string filter = TDMK_Code.filter_str(new string[] { "Process", "Region", "PCS_No" }, new string[] { cb_zone.SelectedItem.ToString(), zone, pcs });
            //        dv.RowFilter = filter;
            //        DataTable dt_filter = dv.ToTable();
            //        List<DataTable> tbl_time = new List<DataTable> { };
            //        Get_ListTable_datetime(-1, dt_filter, new string[] { "Time_Update" }, ref tbl_time, "Remark");

            //        string USL = "";
            //        string LSL = "";
            //        DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_" + cb_zone.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { itemcodevh }));
            //        string region = dgv_edit_data.Rows[row_inx].Cells["Region"].Value.ToString();
            //        if (dt_spec.Rows.Count > 0)
            //        {
            //            USL = dt_spec.Rows[0][region].ToString().Replace("/", string.Empty);
            //            LSL = dt_spec.Rows[1][region].ToString().Replace("/", string.Empty);

            //        }

            //        string colname_select = dgv_edit_data.Columns[col_inx].Name;
            //        if (myCode.checkDBNull(dgv_edit_data.Rows[row_inx].Cells[col_inx].Value) != "")
            //        {
            //            if (colname_select.Contains("After_"))
            //            {
            //                int time_edit = int.Parse(colname_select.Replace("After_", ""));
            //                if (time_edit <= tbl_time.Count)
            //                {
            //                    DataTable tbl_select = tbl_time[time_edit - 1];
            //                    string depart = tbl_select.Rows[0]["Depart"].ToString();
            //                    string time_update = tbl_select.Rows[0]["Time_Update"].ToString();
            //                    string Operator = tbl_select.Rows[0]["Operator"].ToString();
            //                    string remark = tbl_select.Rows[0]["Remark"].ToString();

            //                    dgv_edit_data.CurrentCell.ToolTipText = "Depart: " + depart + " \nOperator: " + Operator + " \nLogfile Location: " + remark + " \nTime_Update: " + time_update + " \nUSL: " + USL + " \nLSL: " + LSL;
            //                }
            //            }
            //        }
            //    }

            //    if (select_sheet == "IMPEDANCE")
            //    {  
            //        string pcs = dgv_edit_data.Rows[row_inx].Cells["PCS"].Value.ToString();
            //        string zone = dgv_edit_data.Rows[row_inx].Cells["Region"].Value.ToString();
                   
            //        string colname_select = dgv_edit_data.Columns[col_inx].Name;
            //        if (myCode.checkDBNull(dgv_edit_data.Rows[row_inx].Cells[col_inx].Value) != "")
            //        {
            //            if (colname_select.Contains("After_"))
            //            {
            //                DataView dv = tbl_history.AsDataView();
            //                if (dv.Count > 0)
            //                {
            //                    string filter = TDMK_Code.filter_str(new string[] { "Zone", "PCS_No" }, new string[] { zone, pcs });
            //                    dv.RowFilter = filter;
            //                }


            //                DataTable dt_filter = dv.ToTable();
            //                List<DataTable> tbl_time = new List<DataTable> { };
            //                Get_ListTable_datetime(-1, dt_filter, new string[] { "Time_Update" }, ref tbl_time, "Remark");

            //                string USL = "";
            //                string LSL = "";
            //                DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, cb_zone.SelectedItem.ToString().ToUpper() + "_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { itemcodevh }));
            //                int region = int.Parse(zone[zone.Length - 1].ToString());
            //                if (dt_spec.Rows.Count > 0)
            //                {
            //                    USL = dt_spec.Rows[region - 1]["USL"].ToString();
            //                    LSL = dt_spec.Rows[region - 1]["LSL"].ToString();

            //                }


            //                int time_edit = int.Parse(colname_select.Replace("After_", ""));
            //                if (time_edit <= tbl_time.Count)
            //                {
            //                    DataTable tbl_select = tbl_time[time_edit - 1];
            //                    string depart = tbl_select.Rows[0]["Depart"].ToString();
            //                    string time_update = tbl_select.Rows[0]["Time_Update"].ToString();
            //                    string Operator = tbl_select.Rows[0]["Operator"].ToString();
            //                    string remark = tbl_select.Rows[0]["Remark"].ToString();

            //                    dgv_edit_data.CurrentCell.ToolTipText = "Depart: " + depart + " \nOperator: " + Operator + " \nLogfile Location: " + remark + " \nTime_Update: " + time_update + " \nUSL: " + USL + " \nLSL: " + LSL;
            //                }
            //            }
            //        }
            //    }

            //    if (select_sheet == "SOLDERMASK")
            //    {

            //        string pcs = dgv_edit_data.Rows[row_inx].Cells["PCS"].Value.ToString();
            //        string zone = dgv_edit_data.Rows[row_inx].Cells["Zone"].Value.ToString();

            //        DataView dv = tbl_history.AsDataView();
            //        string filter = TDMK_Code.filter_str(new string[] { "Zone", "PCS_No" }, new string[] { zone, pcs });
            //        dv.RowFilter = filter;
            //        DataTable dt_filter = dv.ToTable();
            //        List<DataTable> tbl_time = new List<DataTable> { };
            //        Get_ListTable_datetime(-1, dt_filter, new string[] { "Time_Update" }, ref tbl_time, "Remark");

                    

            //        string colname_select = dgv_edit_data.Columns[col_inx].Name;
            //        if (myCode.checkDBNull(dgv_edit_data.Rows[row_inx].Cells[col_inx].Value) != "")
            //        {
            //            if (colname_select.Contains("After_"))
            //            {
            //                int time_edit = int.Parse(colname_select.Replace("After_", ""));
            //                if (time_edit <= tbl_time.Count)
            //                {
            //                    DataTable tbl_select = tbl_time[time_edit - 1];
            //                    string depart = tbl_select.Rows[0]["Depart"].ToString();
            //                    string time_update = tbl_select.Rows[0]["Time_Update"].ToString();
            //                    string Operator = tbl_select.Rows[0]["Operator"].ToString();
            //                    string remark = tbl_select.Rows[0]["Remark"].ToString();

            //                    dgv_edit_data.CurrentCell.ToolTipText = "Depart: " + depart + " \nOperator: " + Operator + " \nLogfile Location: " + remark + " \nTime_Update: " + time_update;
            //                }
            //            }
            //        }
            //    }

            //    if (select_sheet == "ACF_FLATNESS")
            //    {
            //        string pcs = "Point " + dgv_edit_data.Rows[row_inx].Cells["Point"].Value.ToString().Replace("Point", string.Empty);
            //        string zone = dgv_edit_data.Rows[row_inx].Cells["Sample"].Value.ToString();

            //        DataView dv = tbl_history.AsDataView();
            //        string filter = TDMK_Code.filter_str(new string[] { "Zone", "PCS_No" }, new string[] { zone, pcs });
            //        dv.RowFilter = filter;
            //        DataTable dt_filter = dv.ToTable();
            //        List<DataTable> tbl_time = new List<DataTable> { };
            //        Get_ListTable_datetime(-1, dt_filter, new string[] { "Time_Update" }, ref tbl_time, "Remark");

            //        //string USL = "";
            //        //string LSL = "";
            //        //DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Zone" }, new string[] { itemcodevh, zone.Replace("Zone_", "") }));

            //        //if (dt_spec.Rows.Count > 0)
            //        //{
            //        //    USL = dt_spec.Rows[0]["USL"].ToString();
            //        //    LSL = dt_spec.Rows[0]["LSL"].ToString();
            //        //}

            //        string colname_select = dgv_edit_data.Columns[col_inx].Name;
            //        if (myCode.checkDBNull(dgv_edit_data.Rows[row_inx].Cells[col_inx].Value) != "")
            //        {
            //            if (colname_select.Contains("After_"))
            //            {
            //                int time_edit = int.Parse(colname_select.Replace("After_", ""));
            //                if (time_edit <= tbl_time.Count)
            //                {
            //                    DataTable tbl_select = tbl_time[time_edit - 1];
            //                    string depart = tbl_select.Rows[0]["Depart"].ToString();
            //                    string time_update = tbl_select.Rows[0]["Time_Update"].ToString();
            //                    string Operator = tbl_select.Rows[0]["Operator"].ToString();
            //                    dgv_edit_data.CurrentCell.ToolTipText = "Depart: " + depart + " \nOperator: " + Operator + " \nTime_Update: " + time_update;
            //                }
            //            }
            //        }
            //    }

            //    if (select_sheet == "MOISTURE ABSORPTION")
            //    {
            //        string pcs =  dgv_edit_data.Rows[row_inx].Cells["PCS"].Value.ToString();
            //        string zone = dgv_edit_data.Rows[row_inx].Cells["Zone"].Value.ToString();

            //        DataView dv = tbl_history.AsDataView();
            //        string filter = TDMK_Code.filter_str(new string[] { "Zone", "PCS_No" }, new string[] { zone, pcs });
            //        dv.RowFilter = filter;
            //        DataTable dt_filter = dv.ToTable();
            //        List<DataTable> tbl_time = new List<DataTable> { };
            //        Get_ListTable_datetime(-1, dt_filter, new string[] { "Time_Update" }, ref tbl_time, "Remark");

            //        //string USL = "";
            //        //string LSL = "";
            //        //DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Zone" }, new string[] { itemcodevh, zone.Replace("Zone_", "") }));

            //        //if (dt_spec.Rows.Count > 0)
            //        //{
            //        //    USL = dt_spec.Rows[0]["USL"].ToString();
            //        //    LSL = dt_spec.Rows[0]["LSL"].ToString();
            //        //}

            //        string colname_select = dgv_edit_data.Columns[col_inx].Name;
            //        if (myCode.checkDBNull(dgv_edit_data.Rows[row_inx].Cells[col_inx].Value) != "")
            //        {

            //            if (colname_select.Contains("After_"))
            //            {
            //                int time_edit = int.Parse(colname_select.Replace("After_", ""));
            //                if (time_edit <= tbl_time.Count)
            //                {
            //                    DataTable tbl_select = tbl_time[time_edit - 1];
            //                    string depart = tbl_select.Rows[0]["Depart"].ToString();
            //                    string time_update = tbl_select.Rows[0]["Time_Update"].ToString();
            //                    string Operator = tbl_select.Rows[0]["Operator"].ToString();
            //                    dgv_edit_data.CurrentCell.ToolTipText = "Depart: " + depart + " \nOperator: " + Operator + " \nTime_Update: " + time_update;
            //                }
            //            }
            //        }
            //    }
            //}


        }
            
        

        private void dgv_edit_data_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        public static byte[] imgToByteConverter(Image inImg)
        {
            ImageConverter imageConverter = new ImageConverter();
            return (byte[])imageConverter.ConvertTo(inImg, typeof(byte[]));
        }

        public DataTable DGVImage_To_Table(DataGridView dgv)
        {
            try
            {
                if (dgv.ColumnCount == 0)
                {
                    return null;
                }

                DataTable dataTable = new DataTable();
                foreach (DataGridViewColumn column in dgv.Columns)
                {
                    if (!(column.Name == string.Empty))
                    {
                        if (column.CellType.Name.Contains("Before") || column.CellType.Name.Contains("After"))
                        {
                            dataTable.Columns.Add(column.Name, Type.GetType("System.Byte[]"));
                        }
                        else
                        {
                            dataTable.Columns.Add(column.Name);
                        }

                        dataTable.Columns[column.Name].Caption = column.HeaderText;
                    }
                }

                if (dataTable.Columns.Count == 0)
                {
                    return null;
                }

                foreach (DataGridViewRow item in (IEnumerable)dgv.Rows)
                {
                    DataRow dataRow = dataTable.NewRow();
                    foreach (DataColumn column2 in dataTable.Columns)
                    {
                        if (column2.DataType == Type.GetType("System.Byte[]"))
                        {
                            dataRow[column2.ColumnName] = imgToByteConverter((Bitmap)item.Cells[column2.ColumnName].Value);
                        }
                        else
                        {
                            dataRow[column2.ColumnName] = item.Cells[column2.ColumnName].Value;
                        }
                    }

                    dataTable.Rows.Add(dataRow);
                }

                return dataTable;
            }
            catch
            {
                return null;
            }
        }


        private void DGV_data_image_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int col_inx = e.ColumnIndex;
            int r_inx = e.RowIndex;
            DataGridViewCell cur_cell = DGV_data_image.CurrentCell;
            if (DGV_data_image.Columns[col_inx].Name.Contains("Before") || DGV_data_image.Columns[col_inx].Name.Contains("After"))
            {
                if (myCode.checkDBNull(cur_cell.Value) != "")
                {
                    string region = "";
                    if (DGV_data_image.Columns.Contains("Zone"))
                    {
                        region = DGV_data_image.Rows[r_inx].Cells["Zone"].Value.ToString();
                    }
                    else if (DGV_data_image.Columns.Contains("Region"))
                    {
                        region = DGV_data_image.Rows[r_inx].Cells["Region"].Value.ToString();
                    }

                    string pcs = DGV_data_image.Rows[r_inx].Cells["PCS"].Value.ToString();

                    
                    
                    View_detail_Image fr1 = new View_detail_Image();
                    fr1.data = (Bitmap)cur_cell.Value;
                    fr1.region_ = region;
                    fr1.pcs_ = pcs;
                    fr1.Show();

                   
                     
                }
               
            }
        }

        private void dgv_edit_data_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int col_inx = e.ColumnIndex;
            int row_inx = e.RowIndex;
            if (row_inx != -1 && row_inx != dgv_edit_data.Rows.Count - 1 && col_inx != -1 && dgv_edit_data.Rows.Count > 0)
            {
                string select_sheet = sheet;
                if (select_sheet == "Stack-up")
                {
                    string pcs = dgv_edit_data.Rows[row_inx].Cells["PCS"].Value.ToString();
                    string zone = dgv_edit_data.Rows[row_inx].Cells["Zone"].Value.ToString();

                    DataView dv = tbl_history.AsDataView();
                    if (dv.Count > 0)
                    {
                        string filter = TDMK_Code.filter_str(new string[] { "Zone", "PCS_No" }, new string[] { zone, pcs });
                        dv.RowFilter = filter;
                        DataTable dt_filter = dv.ToTable();
                        List<DataTable> tbl_time = new List<DataTable> { };
                        Get_ListTable_datetime(-1, dt_filter, new string[] { "Time_Update" }, ref tbl_time, "Remark");

                        string USL = "";
                        string LSL = "";
                        // DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Zone" }, new string[] { itemcodevh, zone.Replace("Zone_", "") }));
                        DataView dv_spec = tbl_spec.AsDataView();
                        dv_spec.RowFilter = TDMK_Code.filter_str(new string[] { "Zone" }, new string[] { zone.Replace("Zone_", "") });
                        DataTable dt_spec = dv_spec.ToTable();

                        if (dt_spec.Rows.Count > 0)
                        {
                            USL = dt_spec.Rows[0]["USL"].ToString();
                            LSL = dt_spec.Rows[0]["LSL"].ToString();
                        }

                        string colname_select = dgv_edit_data.Columns[col_inx].Name;
                        if (colname_select.Contains("After_"))
                        {

                            int time_edit = int.Parse(colname_select.Replace("After_", ""));
                            if (time_edit <= tbl_time.Count)
                            {
                                DataTable tbl_select = tbl_time[time_edit - 1];
                                string depart = tbl_select.Rows[0]["Depart"].ToString();
                                string time_update = tbl_select.Rows[0]["Time_Update"].ToString();
                                string Operator = tbl_select.Rows[0]["Operator"].ToString();
                                string remark = tbl_select.Rows[0]["Remark"].ToString();
                                MessageBox.Show("*DEPART: " + depart + " \n*OPERATOR: " + Operator + " \n*LOGFILE LOCATION: " + remark + " \n*TIME_UPDATE: " + time_update + " \n*USL: " + USL + " \n*LSL: " + LSL, "Information");
                                //dgv_edit_data.CurrentCell.ToolTipText = "Depart: " + depart + " \nOperator: " + Operator + " \nLogfile Location: " + remark + " \nTime_Update: " + time_update + " \nUSL: " + USL + " \nLSL: " + LSL;
                            }
                        }
                        else if (colname_select.Contains("Before"))
                        {
                            if (dgv_edit_data.Rows[row_inx].Cells["ID"].Value.ToString() == "Total")
                            {
                                MessageBox.Show("*USL: " + USL + " \n*LSL: " + LSL, "Information");
                            }


                        }
                    }
                }

                if (select_sheet == "BVH & PTH")
                {
                    string pcs = dgv_edit_data.Rows[row_inx].Cells["PCS"].Value.ToString();
                    string zone = dgv_edit_data.Rows[row_inx].Cells["Region"].Value.ToString();

                    DataView dv = tbl_history.AsDataView();

                    string filter = TDMK_Code.filter_str(new string[] { "Process", "Region", "PCS_No" }, new string[] { cb_zone.SelectedItem.ToString(), zone, pcs });
                    dv.RowFilter = filter;
                    DataTable dt_filter = dv.ToTable();
                    List<DataTable> tbl_time = new List<DataTable> { };
                    Get_ListTable_datetime(-1, dt_filter, new string[] { "Time_Update" }, ref tbl_time, "Remark");

                    string USL = "";
                    string LSL = "";
                    //DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_" + cb_zone.SelectedItem.ToString().ToUpper(), TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { itemcodevh }));
                    DataTable dt_spec = tbl_spec;
                    string region = dgv_edit_data.Rows[row_inx].Cells["Region"].Value.ToString();
                    if (dt_spec.Rows.Count > 0)
                    {
                        USL = dt_spec.Rows[0][region].ToString().Replace("/", string.Empty);
                        LSL = dt_spec.Rows[1][region].ToString().Replace("/", string.Empty);

                    }

                    string colname_select = dgv_edit_data.Columns[col_inx].Name;
                    if (myCode.checkDBNull(dgv_edit_data.Rows[row_inx].Cells[col_inx].Value) != "")
                    {
                        if (colname_select.Contains("After_"))
                        {
                            int time_edit = int.Parse(colname_select.Replace("After_", ""));
                            if (time_edit <= tbl_time.Count)
                            {
                                DataTable tbl_select = tbl_time[time_edit - 1];
                                string depart = tbl_select.Rows[0]["Depart"].ToString();
                                string time_update = tbl_select.Rows[0]["Time_Update"].ToString();
                                string Operator = tbl_select.Rows[0]["Operator"].ToString();
                                string remark = tbl_select.Rows[0]["Remark"].ToString();
                                MessageBox.Show("*DEPART: " + depart + " \n*OPERATOR: " + Operator + " \n*LOGFILE LOCATION: " + remark + " \n*TIME_UPDATE: " + time_update + " \n*USL: " + USL + " \n*LSL: " + LSL, "Information");
                                // dgv_edit_data.CurrentCell.ToolTipText = "Depart: " + depart + " \nOperator: " + Operator + " \nLogfile Location: " + remark + " \nTime_Update: " + time_update + " \nUSL: " + USL + " \nLSL: " + LSL;
                            }
                        }
                        else if (colname_select.Contains("Before"))
                        {
                         
                                MessageBox.Show("*USL: " + USL + " \n*LSL: " + LSL, "Information");
                        
                        }
                    }
                } 
                if (select_sheet == "Impedance")
                {
                    string pcs = dgv_edit_data.Rows[row_inx].Cells["PCS"].Value.ToString();
                    string zone = dgv_edit_data.Rows[row_inx].Cells["Region"].Value.ToString();

                    string colname_select = dgv_edit_data.Columns[col_inx].Name;
                    if (myCode.checkDBNull(dgv_edit_data.Rows[row_inx].Cells[col_inx].Value) != "")
                    {

                        DataView dv = tbl_history.AsDataView();
                        if (dv.Count > 0)
                        {
                            string filter = TDMK_Code.filter_str(new string[] { "Zone", "PCS_No" }, new string[] { zone, pcs });
                            dv.RowFilter = filter;
                        }


                        DataTable dt_filter = dv.ToTable();
                        List<DataTable> tbl_time = new List<DataTable> { };
                        Get_ListTable_datetime(-1, dt_filter, new string[] { "Time_Update" }, ref tbl_time, "Remark");

                        string USL = "";
                        string LSL = "";
                        // DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, cb_zone.SelectedItem.ToString().ToUpper() + "_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { itemcodevh }));
                        DataTable dt_spec = tbl_spec;
                        int region = int.Parse(zone[zone.Length - 1].ToString());
                        if (dt_spec.Rows.Count > 0)
                        {
                            USL = dt_spec.Rows[region - 1]["USL"].ToString();
                            LSL = dt_spec.Rows[region - 1]["LSL"].ToString();
                        }
                        if (colname_select.Contains("After_"))
                        {
                            int time_edit = int.Parse(colname_select.Replace("After_", ""));
                            if (time_edit <= tbl_time.Count)
                            {
                                DataTable tbl_select = tbl_time[time_edit - 1];
                                string depart = tbl_select.Rows[0]["Depart"].ToString();
                                string time_update = tbl_select.Rows[0]["Time_Update"].ToString();
                                string Operator = tbl_select.Rows[0]["Operator"].ToString();
                                string remark = tbl_select.Rows[0]["Remark"].ToString();
                                MessageBox.Show("*DEPART: " + depart + " \n*OPERATOR: " + Operator + " \n*LOGFILE LOCATION: " + remark + " \n*TIME_UPDATE: " + time_update + " \n*USL: " + USL + " \n*LSL: " + LSL, "Information");
                                //dgv_edit_data.CurrentCell.ToolTipText = "Depart: " + depart + " \nOperator: " + Operator + " \nLogfile Location: " + remark + " \nTime_Update: " + time_update + " \nUSL: " + USL + " \nLSL: " + LSL;
                            }
                        }
                        else if (colname_select.Contains("Before"))
                        {

                            MessageBox.Show("*USL: " + USL + " \n*LSL: " + LSL, "Information");

                        }
                    }
                } 
                if (select_sheet == "ACF")
                {
                    string pcs = "Point " + dgv_edit_data.Rows[row_inx].Cells["Point"].Value.ToString().Replace("Point", string.Empty);
                    string zone = dgv_edit_data.Rows[row_inx].Cells["Sample"].Value.ToString();

                    DataView dv = tbl_history.AsDataView();
                    string filter = TDMK_Code.filter_str(new string[] { "Zone", "PCS_No" }, new string[] { zone, pcs });
                    dv.RowFilter = filter;
                    DataTable dt_filter = dv.ToTable();
                    List<DataTable> tbl_time = new List<DataTable> { };
                    Get_ListTable_datetime(-1, dt_filter, new string[] { "Time_Update" }, ref tbl_time, "Remark");

                    string colname_select = dgv_edit_data.Columns[col_inx].Name;
                    if (myCode.checkDBNull(dgv_edit_data.Rows[row_inx].Cells[col_inx].Value) != "")
                    {
                        if (colname_select.Contains("After_"))
                        {
                            int time_edit = int.Parse(colname_select.Replace("After_", ""));
                            if (time_edit <= tbl_time.Count)
                            {
                                DataTable tbl_select = tbl_time[time_edit - 1];
                                string depart = tbl_select.Rows[0]["Depart"].ToString();
                                string time_update = tbl_select.Rows[0]["Time_Update"].ToString();
                                string Operator = tbl_select.Rows[0]["Operator"].ToString();
                                //dgv_edit_data.CurrentCell.ToolTipText = "Depart: " + depart + " \nOperator: " + Operator + " \nTime_Update: " + time_update;
                                MessageBox.Show("*DEPART: " + depart + " \n*OPERATOR: " + Operator + " \n*TIME_UPDATE: " + time_update, "Information");
                            }
                        }
                    }
                }

                if (select_sheet == "CQRA - Moisture absorption")
                {
                    string pcs = dgv_edit_data.Rows[row_inx].Cells["PCS"].Value.ToString();
                    string zone = dgv_edit_data.Rows[row_inx].Cells["Zone"].Value.ToString();

                    DataView dv = tbl_history.AsDataView();
                    string filter = TDMK_Code.filter_str(new string[] { "Zone", "PCS_No" }, new string[] { zone, pcs });
                    dv.RowFilter = filter;
                    DataTable dt_filter = dv.ToTable();
                    List<DataTable> tbl_time = new List<DataTable> { };
                    Get_ListTable_datetime(-1, dt_filter, new string[] { "Time_Update" }, ref tbl_time, "Remark");

                    string colname_select = dgv_edit_data.Columns[col_inx].Name;
                    if (myCode.checkDBNull(dgv_edit_data.Rows[row_inx].Cells[col_inx].Value) != "")
                    {

                        if (colname_select.Contains("After_"))
                        {
                            int time_edit = int.Parse(colname_select.Replace("After_", ""));
                            if (time_edit <= tbl_time.Count)
                            {
                                DataTable tbl_select = tbl_time[time_edit - 1];
                                string depart = tbl_select.Rows[0]["Depart"].ToString();
                                string time_update = tbl_select.Rows[0]["Time_Update"].ToString();
                                string Operator = tbl_select.Rows[0]["Operator"].ToString();
                                //dgv_edit_data.CurrentCell.ToolTipText = "Depart: " + depart + " \nOperator: " + Operator + " \nTime_Update: " + time_update;
                                MessageBox.Show("Depart: " + depart + " \nOperator: " + Operator + " \nTime_Update: " + time_update, "Information");
                            }
                        }
                    }
                }

            }

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

        private void txtLotNo_Validated(object sender, EventArgs e)
        {
            lotnovh = Lotno_Formated(lotnovh);
        }
    }
}
