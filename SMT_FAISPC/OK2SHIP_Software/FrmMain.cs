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
using TDMK_SQL;
using System.Data.SqlClient;
using myExcel = Microsoft.Office.Interop.Excel;
using OK2SHIP_Measurements;
using System.Net;
using System.Globalization;
using Microsoft.Office.Core;
using System.Diagnostics;
using static System.Resources.ResXFileRef;
using static System.Net.WebRequestMethods;
using System.Xml.Linq;
using OK2SHIP_Lib;
//using Microsoft.Office.Interop.Excel;

namespace OK2SHIP_Software
{
    public partial class FrmMain : Form
    {
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        myVar myCode = new myVar();
        TDMK_OK2SHIP main_code = new TDMK_OK2SHIP();
        Export_VXH_ManualInput export_update = new Export_VXH_ManualInput();
        public SqlConnection sqlcon_OK2SHIP_Period2 = null;
        public static string sel_DB_OK2SHIP_Period2 = "OK2SHIP_Period2";
        public FrmMain()
        {
            InitializeComponent();
        }

        private void btnLoadData_Click(object sender, EventArgs e)
        {
            load_data_click_test();
            /************************************ Test Here *************************************************/
            //Find_data("22B0358-00003");

            /*************************************************************************************************/
        }
        public void load_data_click()
        {
            tvItemDetails.Nodes.Clear();
            AutoCompleteStringCollection FAI_LotNo_lst = TDMK_Code.Load_Item_Names_Filter(myVar.sqlcon_SMT, "FAI_Auto", "LotNo", "ItemCode = '" + txtItemCode.Text + "'");
            string[] IPQC_tbl = myCode.GetAllTables(myVar.sqlcon_IPQC);
            AutoCompleteStringCollection[] IPQC_LotNo_lst = new AutoCompleteStringCollection[IPQC_tbl.Length];
            AutoCompleteStringCollection ignored_lst = new AutoCompleteStringCollection() { "All_Items", "SpecList", "Spec_List", "Sequence" };
            int inx = 0;
            foreach (string tbl in IPQC_tbl)
            {
                if (!TDMK_Code.check_exist_list(tbl, ignored_lst))
                {
                    IPQC_LotNo_lst[inx] = TDMK_Code.Load_Item_Names_Filter(myVar.sqlcon_IPQC, tbl, "LotNo", "ItemCode = '" + txtItemCode.Text + "'");
                    foreach (string t in IPQC_LotNo_lst[inx])
                    {
                        if (!TDMK_Code.check_exist_list(t, FAI_LotNo_lst))
                        {
                            FAI_LotNo_lst.Add(t);
                        }
                    }
                }
                inx++;
            }
            string[] temp = new string[FAI_LotNo_lst.Count];
            FAI_LotNo_lst.CopyTo(temp, 0);
            Array.Sort(temp);
            tvItemDetails.Nodes.Add(txtItemCode.Text);
            string f_loc = Path.Combine(myVar.data_loc, "0.OK2SHIP report format");
            KeyValuePair<string, string> f_info = myCode.Get_Format_Info(txtItemCode.Text, f_loc);
            //AutoCompleteStringCollection Approve_LotNo_lst = TDMK_Code.Load_Item_Names_Filter(myVar.sqlcon_OK2SHIP, "Approve_Data", "LotNo", "ItemCode = '" + txtItemCode.Text + "'");
            AutoCompleteStringCollection Approve_LotNo_lst = TDMK_Code.Load_Item_Names_Filter(myVar.sqlcon_OK2SHIP, "Approve_Data", "LotNo", "ItemCode = '" + f_info.Key + "'");
            if (f_info.Key != null)
            {
                foreach (string t in temp)
                {
                    tvItemDetails.Nodes[0].Nodes.Add(t);
                    if (!TDMK_Code.check_exist_list(t, Approve_LotNo_lst))
                    {
                        string ID = (TDMK_Code.SQL_MAX("Approve_Data", "ID", myVar.sqlcon_OK2SHIP) + 1).ToString();
                        TDMK_Code.insert_val_arr("Approve_Data", myVar.sqlcon_OK2SHIP, new string[] { "ID", "Project_Name", "ItemCode", "LotNo" }, new string[] { ID, f_info.Value, f_info.Key, t });
                        //TDMK_Code.insert_val_arr("Approve_Data", myVar.sqlcon_OK2SHIP, new string[] { "ID", "ItemCode", "LotNo", "Approve" }, new string[] { ID, txtItemCode.Text, t, "No" });
                    }
                }
                DGV_Data.Columns.Clear();
                DataSet myDS = new DataSet();
                //TDMK_Code.fill_dataset_Filter(myDS, "Approve_Data", "ItemCode", txtItemCode.Text, myVar.sqlcon_OK2SHIP);
                TDMK_Code.fill_dataset_Filter(myDS, "Approve_Data", "ItemCode", f_info.Key, myVar.sqlcon_OK2SHIP);
                System.Data.DataTable mydt = new System.Data.DataTable();
                mydt = myDS.Tables[0];
                DGV_Data.DataSource = mydt;
                string[] progress_val = new string[mydt.Rows.Count];
                for (int i = 0; i < mydt.Rows.Count; i++)
                {
                    double sum = 0;
                    string _itemcode = mydt.Rows[i]["ItemCode"].ToString();
                    string _lotno = mydt.Rows[i]["LotNo"].ToString();
                    int excel_f = myCode.count_excel_data(_itemcode, _lotno);
                    int autodata_count = myCode.Count_AutoData(_itemcode, _lotno);
                    sum = ((double)(excel_f + autodata_count) / 42) * 100;
                    progress_val[i] = Convert.ToInt32(sum).ToString();
                }
                Sample.DataGridViewProgressColumn dgv_col = new Sample.DataGridViewProgressColumn();
                dgv_col.Name = "Progress";
                dgv_col.HeaderText = "Progress";
                dgv_col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                DGV_Data.Columns.Add(dgv_col);
                for (int i = 0; i < DGV_Data.RowCount; i++)
                {
                    DGV_Data.Rows[i].Cells["Progress"].Value = progress_val[i];
                }
                foreach (DataGridViewColumn dgv_c in DGV_Data.Columns)
                {
                    dgv_c.ReadOnly = true;
                }
                DGV_Data.AutoResizeColumns();
                main_code.Disable_Sort_DGV(DGV_Data);
                myVar.confirm_mode = true;
            }
            else
            {
                MessageBox.Show("Not found format of ItemCode: " + txtItemCode.Text, "Warning");
            }
        }
        public void load_data_click_test()
        {
            string ignoredSheet = Path.Combine(myVar.app_path, "IgnoredSheet.txt");
            string[] ignoredList_arr = myCode.read_config_arr2(ignoredSheet);
            tvItemDetails.Nodes.Clear();
            AutoCompleteStringCollection FAI_LotNo_lst = TDMK_Code.Load_Item_Names_Filter(myVar.sqlcon_SMT, "FAI_Auto", "LotNo", "ItemCode = '" + txtItemCode.Text + "'");
            string[] IPQC_tbl = myCode.GetAllTables(myVar.sqlcon_IPQC);
            AutoCompleteStringCollection[] IPQC_LotNo_lst = new AutoCompleteStringCollection[IPQC_tbl.Length];
            AutoCompleteStringCollection ignored_lst = new AutoCompleteStringCollection() { "All_Items", "SpecList", "Spec_List", "Sequence" };
            int inx = 0;
            foreach (string tbl in IPQC_tbl)
            {
                if (!TDMK_Code.check_exist_list(tbl, ignored_lst))
                {
                    IPQC_LotNo_lst[inx] = TDMK_Code.Load_Item_Names_Filter(myVar.sqlcon_IPQC, tbl, "LotNo", "ItemCode = '" + txtItemCode.Text + "'");
                    foreach (string t in IPQC_LotNo_lst[inx])
                    {
                        if (!TDMK_Code.check_exist_list(t, FAI_LotNo_lst))
                        {
                            FAI_LotNo_lst.Add(t);
                        }
                    }
                }
                inx++;
            }
            List<string> echeck_process = new List<string>() { "ELECTRICAL", "CQRA_REFLOW", "CQRA_HOT_OIL", "CQRA_THERMAL_CYCLING", "CQRA_HEAT_SOAK", "CQRA_THERMAL_SHOCK", "CQRA_BENDING", "CQRA_THERMAL_CYCLING_AND_BEND", "CQRA_HEAT_SOAK_AND_BEND", "CQRA_SURVIVAL_REFLOW", "CQRA_SURVIVAL_HOT_OIL" };
            List<string> main_lot = new List<string>();
            foreach(string p in echeck_process)
            {
                DataTable dt = TDMK_Code.Datatable_Filter(myVar.sqlcon_OK2SHIP_Period2, p, TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
                string[] cur_Lotno_lst = dt.AsEnumerable().Select(x => x.Field<string>("LotNo")).Distinct().ToArray();
                main_lot.AddRange(cur_Lotno_lst);
            }
            
            main_lot = main_lot.Distinct().ToList();
            foreach(var t in FAI_LotNo_lst)
            {
                if(main_lot.IndexOf(t.ToString())==-1)
                {
                    main_lot.Add(t.ToString());
                }
            }
            
            string[] temp = main_lot.ToArray();
            Array.Sort(temp);
            tvItemDetails.Nodes.Add(txtItemCode.Text);
            string f_loc = Path.Combine(myVar.data_loc, "0.OK2SHIP report format");
            KeyValuePair<string, string> f_info = myCode.Get_Format_Info(txtItemCode.Text, f_loc);
            AutoCompleteStringCollection Approve_LotNo_lst = TDMK_Code.Load_Item_Names_Filter(myVar.sqlcon_OK2SHIP, "Approve_Data", "LotNo", "ItemCode = '" + f_info.Key + "'");
            if (f_info.Key != null)
            {
                foreach (string t in temp)
                {
                    tvItemDetails.Nodes[0].Nodes.Add(t);
                    if (!TDMK_Code.check_exist_list(t, Approve_LotNo_lst))
                    {
                        string ID = (TDMK_Code.SQL_MAX("Approve_Data", "ID", myVar.sqlcon_OK2SHIP) + 1).ToString();
                        TDMK_Code.insert_val_arr("Approve_Data", myVar.sqlcon_OK2SHIP, new string[] { "ID", "Project_Name", "ItemCode", "LotNo" }, new string[] { ID, f_info.Value, f_info.Key, t });
                    }
                }
                DGV_Data.Columns.Clear();
                DataSet myDS = new DataSet();
                TDMK_Code.fill_dataset_Filter(myDS, "Approve_Data", "ItemCode", f_info.Key, myVar.sqlcon_OK2SHIP);
                System.Data.DataTable mydt = new System.Data.DataTable();
                mydt = myDS.Tables[0];
                DGV_Data.DataSource = mydt;
                string[] progress_val = new string[mydt.Rows.Count];
                int total_items = myCode.OK2SHIP_Item_list(txtItemCode.Text, "", "Items_Details", myVar.sqlcon_OK2SHIP).Count-3;
                for (int i = 0; i < mydt.Rows.Count; i++)
                {
                    double sum = 0;
                    string _itemcode = mydt.Rows[i]["ItemCode"].ToString();
                    string _lotno = mydt.Rows[i]["LotNo"].ToString();
                    Dictionary<string, bool> autoData = myCode.FAI_IPQC_Recycle_process(_itemcode, _lotno);
                    Dictionary<string, bool> process_lst = myCode.process_data_ready(_itemcode, _lotno, myVar.sqlcon_OK2SHIP_Period2, ignoredList_arr);
                    int excel_f = myCode.count_excel_data2(_itemcode, _lotno, process_lst.Keys.ToList());
                    sum = (autoData.Values.Count(x => x == true) + process_lst.Values.Count(x => x == true)+ excel_f) *100/ total_items;
                    progress_val[i] = Convert.ToInt32(sum).ToString();
                }
                Sample.DataGridViewProgressColumn dgv_col = new Sample.DataGridViewProgressColumn();
                dgv_col.Name = "Progress";
                dgv_col.HeaderText = "Progress";
                dgv_col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                DGV_Data.Columns.Add(dgv_col);
                for (int i = 0; i < DGV_Data.RowCount; i++)
                {
                    DGV_Data.Rows[i].Cells["Progress"].Value = progress_val[i];
                }
                foreach (DataGridViewColumn dgv_c in DGV_Data.Columns)
                {
                    dgv_c.ReadOnly = true;
                }
                DGV_Data.AutoResizeColumns();
                main_code.Disable_Sort_DGV(DGV_Data);
                myVar.confirm_mode = true;
            }
            else
            {
                MessageBox.Show("Not found format of ItemCode: " + txtItemCode.Text, "Warning");
            }
        }
        public void load_data(string tar_item_lot, TreeView tar_tvDetail)
        {
            string tar_folder = FrmLogin.app_path;// @"D:\Customer Projects\SEEV\TestAreas\OK2SHIP automation system";
            string[] folders = System.IO.Directory.GetDirectories(tar_folder);
            string[] main_list = new string[folders.Length];
            string[] tg_main_list = new string[folders.Length + 2];
            tg_main_list[0] = "ItemCode";
            tg_main_list[1] = "LotNo";
            //int i = 0;
            for (int i = 0; i < folders.Length; i++)
            {
                foreach (string c in folders)
                {
                    FileInfo f = new FileInfo(c);
                    if (f.Name.Split('.')[0] == (i + 1).ToString())
                    {
                        main_list[i] = f.Name;
                        tg_main_list[i + 2] = f.Name;
                        break;
                    }
                }
            }
            tar_tvDetail.Nodes.Clear();
            int node_inx = 0;
            foreach (string t in main_list)
            {
                string node_f = Path.Combine(tar_folder, t);
                string[] node_name = Directory.GetFiles(node_f, "*.xlsx").Where(s => s.Contains(tar_item_lot)).ToArray();
                tar_tvDetail.Nodes.Add(t);
                foreach (string c in node_name)
                {
                    FileInfo f = new FileInfo(c);
                    tar_tvDetail.Nodes[node_inx].Nodes.Add(f.Name);
                }
                node_inx++;
            }
        }
        public void Find_data(string[] tar_item_lot, string tar_itemcode)
        {
            string tar_folder = @"D:\Customer Projects\SEEV\TestAreas\OK2SHIP automation system";//FrmLogin.app_path;//
            string[] folders = System.IO.Directory.GetDirectories(tar_folder);
            string[] main_list = new string[folders.Length - 1];
            AutoCompleteStringCollection ignoredList = new AutoCompleteStringCollection() { "FAI", "CPK", "IPQC", "Recycle", "Materials", "Coverpage", "User Guidelines", "Declaration and Contents" };
            for (int i = 0; i < folders.Length; i++)
            {
                foreach (string c in folders)
                {
                    FileInfo f = new FileInfo(c);
                    if (f.Name.Split('.')[0] == (i + 1).ToString())
                    {
                        main_list[i] = f.Name;
                        break;
                    }
                }
            }
            int[] lot_sum = new int[tar_item_lot.Length];
            int lot_inx = 0;
            foreach (string t_lot in tar_item_lot)
            {
                string cur_item_lot = tar_itemcode + "-" + t_lot;
                foreach (string t in main_list)
                {
                    if (!TDMK_Code.check_exist_list(t, ignoredList))
                    {
                        string node_f = Path.Combine(tar_folder, t);
                        string[] node_name = Directory.GetFiles(node_f, "*.xlsx").Where(s => s.Contains(cur_item_lot)).ToArray();
                        if (node_name.Length > 0)
                        {
                            lot_sum[lot_inx]++;
                        }
                    }
                }

                lot_inx++;
            }

        }
        private void tvItems_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }

        private void btnTest_Click(object sender, EventArgs e)
        {

            DGV_Data.Columns.Clear();
            //get_ItemCode_Lotno(@"D:\Customer Projects\SEEV\TestAreas\OK2SHIP automation system", tvItemDetails);
            string connstr = TDMK_Code.data_connection(@"(local)\TDMK_SQLExpress", "SEI_DB", "sa", "vh031506").ConnectionString;
            SqlConnection sqlcon = new SqlConnection(connstr);
            DataSet myDS = new DataSet();
            TDMK_Code.fill_dataset_Filter(myDS, "OK2SHIP_Process", "ItemCode", txtItemCode.Text, sqlcon);
            //TDMK_Code.fill_dataset(myDS, "OK2SHIP_Process", sqlcon);
            System.Data.DataTable mydt = new System.Data.DataTable();
            mydt = myDS.Tables[0];

            DataTable temp_tbl = new DataTable();
            temp_tbl = mydt.AsDataView().ToTable(false, "ID", "ItemCode", "LotNo", "Approve");
            DGV_Data.DataSource = temp_tbl;
            string[] progress_val = new string[mydt.Rows.Count];
            for (int i = 0; i < mydt.Rows.Count; i++)
            {
                double sum = 0;
                double[] val = new double[43];
                for (int k = 0; k < 43; k++)
                {
                    val[k] = Convert.ToDouble(mydt.Rows[i][k + 3]);
                }
                sum = val.Average();
                progress_val[i] = Convert.ToInt32(sum).ToString();
            }
            Sample.DataGridViewProgressColumn dgv_col = new Sample.DataGridViewProgressColumn();
            dgv_col.Name = "Process";
            dgv_col.HeaderText = "Process";
            DGV_Data.Columns.Add(dgv_col);
            for (int i = 0; i < DGV_Data.RowCount; i++)
            {
                DGV_Data.Rows[i].Cells["Process"].Value = progress_val[i];
            }
            List<string> ItemCode_list = new List<string>();
            List<string> Lotno_list = new List<string>();
            foreach (TreeNode t in tvItemDetails.Nodes)
            {
                ItemCode_list.Add(t.Text);
                if (t.Nodes.Count > 0)
                {
                    foreach (TreeNode c in t.Nodes)
                    {
                        Lotno_list.Add(c.Text);
                    }
                }
            }
            //listBox3.DataSource = ItemCode_list;
            //listBox2.DataSource = Lotno_list;

        }
        public void Load_data_ItemCode(string src_itemcode, SqlConnection src_conn, DataGridView DGV_Process, TreeView tar_tv)
        {
            // DGV_detail.Columns.Clear();
            DGV_Process.Columns.Clear();
            tar_tv.Nodes.Clear();
            DataSet myDS = new DataSet();
            TDMK_Code.fill_dataset_Filter(myDS, "OK2SHIP_Process", "ItemCode", src_itemcode, src_conn);
            System.Data.DataTable mydt = new System.Data.DataTable();
            mydt = myDS.Tables[0];
            //DGV_detail.DataSource = mydt;
            DataTable temp_tbl = new DataTable();
            temp_tbl = mydt.AsDataView().ToTable(false, "ID", "ItemCode", "LotNo", "Approve");
            tar_tv.Nodes.Add(src_itemcode);
            List<string> lotno_lst = new List<string>();
            lotno_lst = mydt.AsEnumerable().Select(r => r.Field<string>("LotNo")).Distinct().ToList();
            foreach (string t in lotno_lst)
            {
                tar_tv.Nodes[0].Nodes.Add(t);
            }
            DGV_Process.DataSource = temp_tbl;
            string[] progress_val = new string[mydt.Rows.Count];
            for (int i = 0; i < mydt.Rows.Count; i++)
            {
                double sum = 0;
                double[] val = new double[43];
                for (int k = 0; k < 43; k++)
                {
                    val[k] = Convert.ToDouble(mydt.Rows[i][k + 3]);
                }
                sum = val.Average();
                progress_val[i] = Convert.ToInt32(sum).ToString();
            }
            Sample.DataGridViewProgressColumn dgv_col = new Sample.DataGridViewProgressColumn();
            dgv_col.Name = "Process";
            dgv_col.HeaderText = "Process";
            dgv_col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            DGV_Process.Columns.Add(dgv_col);
            for (int i = 0; i < DGV_Process.RowCount; i++)
            {
                DGV_Process.Rows[i].Cells["Process"].Value = progress_val[i];
            }
            DGV_Process.AutoResizeColumns();
            List<string> ItemCode_list = new List<string>();
            List<string> Lotno_list = new List<string>();
            foreach (TreeNode t in tvItemDetails.Nodes)
            {
                ItemCode_list.Add(t.Text);
                if (t.Nodes.Count > 0)
                {
                    foreach (TreeNode c in t.Nodes)
                    {
                        Lotno_list.Add(c.Text);
                    }
                }
            }

        }
        public void Create_table(string tbl_name)
        {
            string create_tbl = "CREATE TABLE IF NOT EXISTS [dbo].[" + tbl_name + "] (ID INT NOT NULL)";
            string[] mycol = { "Device", "Model" };
            string query = add_tbl_str("myTable", mycol);
            string connstr = TDMK_Code.data_connection(@"(local)\TDMK_SQLExpress", "SEI_DB", "sa", "vh031506").ConnectionString;
            SqlConnection con = new SqlConnection(connstr);
            if (con.State != ConnectionState.Open)
            {
                con.Open();
            }
            SqlCommand cmd = new SqlCommand(query, con);
            cmd.ExecuteNonQuery();
            con.Close();
        }
        public string add_tbl_str(string src_tbl_name, string[] col_name)
        {
            string query = "IF OBJECT_ID('dbo." + src_tbl_name + "', 'U') IS NULL ";
            query += "BEGIN ";
            query += "CREATE TABLE [dbo].[" + src_tbl_name + "](";
            query += "[ID] INT IDENTITY(1,1) NOT NULL CONSTRAINT pkID PRIMARY KEY,";
            foreach (string c in col_name)
            {
                query += "[" + c + "] NVARCHAR(MAX) NULL,";
            }
            query.TrimEnd(',');
            query += ")";
            query += " END";
            return query;
        }
        public void get_ItemCode_Lotno(string tar_location, TreeView tar_TV)
        {
            string tar_folder = tar_location;
            string[] node_name = Directory.GetFiles(tar_folder, "*.xlsx", SearchOption.AllDirectories);
            string[] main_list = new string[node_name.Length];
            int i = 0;
            foreach (string c in node_name)
            {
                FileInfo f = new FileInfo(c);
                main_list[i] = f.Name;
                i++;
            }
            main_list = main_list.Distinct().ToArray();
            int arr_num = main_list.Length;
            string[] tar_val = new string[arr_num];
            string[] item_code = new string[arr_num];
            string[] Lotno = new string[arr_num];
            int sel_inx = 0;
            for (int j = 0; j < arr_num; j++)
            {
                if (main_list[j].Contains("_"))
                {
                    tar_val[sel_inx] = main_list[j].Split('_')[0];
                    string[] temp = tar_val[sel_inx].Split('-');
                    item_code[sel_inx] = temp[0];
                    Lotno[sel_inx] = temp[1];
                    sel_inx++;
                }
            }
            Array.Resize(ref tar_val, sel_inx++);
            tar_val = tar_val.Distinct().ToArray();
            item_code = item_code.Distinct().ToArray();
            Lotno = Lotno.Distinct().ToArray();
            int node_inx = 0;
            foreach (string t in item_code)
            {
                if (t != null)
                {
                    tar_TV.Nodes.Add(t);
                    foreach (string c in tar_val)
                    {
                        if (c.Contains(t))
                        {
                            tar_TV.Nodes[node_inx].Nodes.Add(c.Split('-')[1]);
                        }
                    }
                    node_inx++;
                }
            }
        }

        private void tvItemDetails_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }

        private void tvItemDetails_AfterCheck(object sender, TreeViewEventArgs e)
        {
            //if(e.Node.Parent==null)
            //{
            //    foreach (TreeNode t in e.Node.Nodes)
            //    {
            //        t.Checked = true;
            //    }
            //}              
        }

        private void tvItems_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Text.Contains(".xlsx"))
            {
                string f_path = e.Node.FullPath;
                TDMK_Code.open_excel_file(FrmLogin.app_path, f_path, "");
            }
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
            //export_update.sqlcon_OK2SHIP_Period2 =  myCode.initial_data("";// initial_data_OK2SHIP_Period2(sel_DB_OK2SHIP_Period2, true);
        }
        public SqlConnection initial_data_OK2SHIP_Period2(string DB_name, bool sa_en)
        {
            SqlConnection _sqlcon_OK2SHIP;
            string app_path = Application.StartupPath;
            string config_file = Path.Combine(app_path, "Config", "config.txt");
            string[] my_config = myCode.read_config_arr(config_file);
            string server_name = "";
            string server_acc = "";
            string server_pass = "";


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
            sqlcon_OK2SHIP_Period2 = _sqlcon_OK2SHIP;
            return _sqlcon_OK2SHIP;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            create_Data();
        }
        public void create_Data()
        {
            List<string> lst_item = new List<string>();
            int col_count = lst_item.Count + 5;//48;
            string[] items = new string[col_count];
            string[] item_vals = new string[col_count];
            Random rnd = new Random();
            items[0] = "ID";
            items[1] = "ItemCode";
            items[2] = "LotNo";
            items[3] = "Approve";
            for (int i = 0; i < col_count - 5; i++)
            {
                items[4 + i] = "[" + lst_item[i].ToString() + "]";
            }

            //foreach (object t in listBox3.Items)
            //{
            //    item_vals[1] = t.ToString();
            //    foreach (object c in listBox2.Items)
            //    {
            //        item_vals[0] = (TDMK_Code.SQL_MAX("OK2SHIP_Process", "ID", FrmLogin.sqlcon) + 1).ToString();
            //        item_vals[2] = c.ToString();
            //        item_vals[3] = "Yes";
            //        for (int k = 0; k < col_count - 5; k++)
            //        {
            //            item_vals[4 + k] = rnd.Next(0, 100).ToString();
            //        }
            //        TDMK_Code.insert_val_arr("OK2SHIP_Process", FrmLogin.sqlcon, items, item_vals);

            //    }
            //}

            foreach (TreeNode t in tvItemDetails.Nodes)
            {
                item_vals[1] = t.Text;
                if (t.Nodes.Count > 0)
                {
                    foreach (TreeNode c in t.Nodes)
                    {
                        item_vals[0] = (TDMK_Code.SQL_MAX("OK2SHIP_Process", "ID", FrmLogin.sqlcon) + 1).ToString();
                        item_vals[2] = c.Text;
                        item_vals[3] = "Yes";
                        for (int k = 0; k < col_count - 5; k++)
                        {
                            item_vals[4 + k] = rnd.Next(0, 100).ToString();
                        }
                        TDMK_Code.insert_val_arr("OK2SHIP_Process", FrmLogin.sqlcon, items, item_vals);
                    }
                }
            }
            //TDMK_Code.fill_dataset_DGV("OK2SHIP_Process", TDMK_Code.SQL_CMD("OK2SHIP_Process"), DGV_Data, FrmLogin.sqlcon);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DGV_Data.Columns.Add("ID", "ID");
            DGV_Data.Columns.Add("ItemCode", "ItemCode");
            DGV_Data.Columns.Add("LotNo", "LotNo");
            DGV_Data.Columns.Add("Approve", "Approve");
            Sample.DataGridViewProgressColumn dgv_col = new Sample.DataGridViewProgressColumn();
            dgv_col.Name = "Process";
            dgv_col.HeaderText = "Process";
            dgv_col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            DGV_Data.Columns.Add(dgv_col);
            for (int i = 0; i < 3; i++)
            {
                DGV_Data.Rows.Add(i.ToString(), i.ToString(), i.ToString(), i.ToString(), i.ToString());
            }
            string tg = "45";
            DGV_Data.Rows[2].Cells["Process"].Value = tg;
            DGV_Data.AutoResizeColumns();

        }
        private void tvItemDetails_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            string f_name = "";
            DataSet myDS = new DataSet();
            DataTable mydt = new DataTable();
            if (e.Node.Checked)
            {
                if (e.Node.Parent != null)
                {

                    f_name = e.Node.Parent.Text + "-" + e.Node.Text;
                    //DGV_Data.Columns.Clear();
                    //TDMK_Code.fill_dataset_Filter_arr(myDS, FrmLogin.sqlcon, "OK2SHIP_Process", new string[] { "ItemCode", "LotNo" }, new string[] { e.Node.Parent.Text, e.Node.Text });
                    //mydt = myDS.Tables[0];
                    //DGV_Data.Columns.Add("Items", "Items");
                    //Sample.DataGridViewProgressColumn dgv_col = new Sample.DataGridViewProgressColumn();
                    //dgv_col.Name = "Process";
                    //dgv_col.HeaderText = "Process";
                    //dgv_col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    //DGV_Data.Columns.Add(dgv_col);
                    //for (int i=0;i<43;i++)
                    //{
                    //    string progress = mydt.Rows[0][i + 3].ToString();
                    //    string item_name = mydt.Columns[i + 3].ColumnName;
                    //    DGV_Data.Rows.Add(item_name, progress);
                    //}
                    //DGV_Data.AutoResizeColumns();
                    //DGV_Data.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
                    //Query_Data(e.Node.Parent.Text, e.Node.Text, tvItems, DGV_Data);
                    Details_Data frmDetails = new Details_Data(e.Node.Parent.Text, e.Node.Text);
                    frmDetails.Show();

                }
                else
                {
                    f_name = e.Node.Text;
                }
                //load_data(f_name,tvItems);
            }
        }
        private void DGV_temp_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridViewCell curr_cell;
            string curr_col_name;
            bool approve_en = true;
            if (e.Button == MouseButtons.Right)
            {
                curr_cell = DGV_Data.CurrentCell;
                curr_col_name = DGV_Data.Columns[curr_cell.ColumnIndex].Name.ToString();
                //string appr = DGV_Data.Rows[curr_cell.RowIndex].Cells["Approve"].Value.ToString();
                string QA_PIC_appr = DGV_Data.Rows[curr_cell.RowIndex].Cells["QA_PIC_Approve"].Value.ToString();
                string NPI_PE_appr = DGV_Data.Rows[curr_cell.RowIndex].Cells["NPI_PE_Manager"].Value.ToString();
                string QA_Manager_appr = DGV_Data.Rows[curr_cell.RowIndex].Cells["QA_Manager"].Value.ToString();
                string Customer_appr = DGV_Data.Rows[curr_cell.RowIndex].Cells["Customer_Approve"].Value.ToString();
                //string itemcode = DGV_Data.Rows[curr_cell.RowIndex].Cells["ItemCode"].Value.ToString().Substring(0,7);
                string itemcode = DGV_Data.Rows[curr_cell.RowIndex].Cells["ItemCode"].Value.ToString();
                string lotno = DGV_Data.Rows[curr_cell.RowIndex].Cells["LotNo"].Value.ToString();
                DataTable result_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_Declare, "Results", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcode, lotno }));
                string[] result_lst = result_tbl.AsEnumerable().Select(r => r.Field<string>("Results")).ToArray();
                if (result_lst.Length == 0)
                {
                    approve_en = false;
                }
                else
                {
                    foreach (string t in result_lst)
                    {
                        string tg = myCode.checkDBNull(t);
                        if ((tg == "") || (tg == "No data"))
                        {
                            approve_en = false;
                            break;
                        }
                    }
                }
                reset_jugde();
                switch (curr_col_name)
                {
                    case "QA_PIC_Approve":     //"Approve"
                        if (approve_en)
                        {
                            if (NPI_PE_appr == "")
                            {
                                DGV_Data.ContextMenuStrip = cmsResult;
                                tsmNA.Enabled = false;
                            }
                            else
                            {

                            }
                        }
                        else
                        {
                            MessageBox.Show("Not enough data", "Warning");
                        }
                        break;
                    case "NPI_PE_Manager":
                        if (QA_PIC_appr == "Pass")
                        {
                            if (QA_Manager_appr == "")
                            {
                                DGV_Data.ContextMenuStrip = cmsResult;
                                tsmNA.Enabled = false;
                            }
                        }
                        else
                        {
                            MessageBox.Show("Please, not yet approved by QA_PIC!", "Warning");
                        }
                        break;
                    case "QA_Manager":        // "Internal_Approve"
                        if ((QA_PIC_appr == "Pass") && (NPI_PE_appr == "Pass"))
                        {
                            if ((Customer_appr == "") || (Customer_appr == "NA"))
                            {
                                DGV_Data.ContextMenuStrip = cmsResult;
                                tsmNA.Enabled = false;
                            }
                        }
                        else
                        {
                            MessageBox.Show("Please, not yet approved by NPI_PE_Manager!", "Warning");
                        }
                        break;
                    case "Customer_Approve":        //"Customer_Approve"

                        if ((QA_PIC_appr == "Pass") && (NPI_PE_appr == "Pass"))
                        {
                            if (QA_Manager_appr == "Pass")
                            {
                                tsmNA.Enabled = true;
                            }
                            else
                            {
                                tsmPass.Enabled = false;
                                tsmFail.Enabled = false;
                                tsmBlank.Enabled = true;
                                tsmNA.Enabled = true;
                            }
                            DGV_Data.ContextMenuStrip = cmsResult;
                        }
                        else
                        {
                            MessageBox.Show("Please, not yet approved by QA_Manager!", "Warning");
                        }
                        break;
                    default:
                        DGV_Data.ContextMenuStrip = null;
                        break;
                }
            }
        }
        public void reset_jugde()
        {
            tsmPass.Enabled = true;
            tsmFail.Enabled = true;
            tsmBlank.Enabled = true;
            tsmNA.Enabled = false;
            DGV_Data.ContextMenuStrip = null;
        }
        public void Query_Data(string src_Itemcode, string src_Lotno, TreeView tar_tvDetail, DataGridView tar_DGV_process)
        {
            DataSet myDS = new DataSet();
            DataTable mydt = new DataTable();
            tar_DGV_process.Columns.Clear();
            TDMK_Code.fill_dataset_Filter_arr(myDS, FrmLogin.sqlcon, "OK2SHIP_Process", new string[] { "ItemCode", "LotNo" }, new string[] { src_Itemcode, src_Lotno });
            mydt = myDS.Tables[0];
            tar_DGV_process.Columns.Add("Items", "Items");
            Sample.DataGridViewProgressColumn dgv_col = new Sample.DataGridViewProgressColumn();
            dgv_col.Name = "Process";
            dgv_col.HeaderText = "Process";
            dgv_col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            tar_DGV_process.Columns.Add(dgv_col);
            for (int i = 0; i < 43; i++)
            {
                string progress = mydt.Rows[0][i + 3].ToString();
                string item_name = mydt.Columns[i + 3].ColumnName;
                tar_DGV_process.Rows.Add(item_name, progress);
            }
            tar_DGV_process.AutoResizeColumns();
            tar_DGV_process.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
        }

        public myExcel.Workbook create_export_wrk(string format_file, string process_name, string ItemCode, string Lotno)
        {
            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            xlsApp.DisplayAlerts = false;
            myExcel.Workbook wb = TDMK_Code.open_excel_file(format_file, "", "");
            string _process = process_name.Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
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

                myExcel.Workbook save_wb = TDMK_Code.Create_workbook();
                wb.Sheets[mySheet].Copy(After: save_wb.Sheets[1]);
                myExcel.Worksheet del_sht = save_wb.Sheets[1];
                del_sht.Delete();
                wb.Close();
                xlsApp.DisplayAlerts = true;
                return save_wb;
            }
            else
            {
                wb.Close();
                xlsApp.DisplayAlerts = true;
                return null;
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            string format_folder = @"0.OK2SHIP Report Format";
            string node_f = Path.Combine(myVar.data_loc, format_folder);
            string src_file_format = "";
            string report_name = "";
            bool export_en = false;
            for (int i = 0; i < DGV_Data.RowCount; i++)
            {
                /************************ Test Ignored ************************************************************************************************/
                //if (DGV_Data.Rows[i].Cells["QA_Manager"].Value.ToString() == "Pass") //&&(DGV_Data.Rows[i].Cells["Process"].Value.ToString()=="100")
                //{
                string tar_item_lot = DGV_Data.Rows[i].Cells["ItemCode"].Value.ToString();
                string[] node_name = Directory.GetFiles(node_f, "*.xlsm").Where(s => s.Contains(tar_item_lot)).ToArray();
                if (node_name.Length > 0)
                {
                    report_name = Path.GetFileNameWithoutExtension(node_name[0]);
                    src_file_format = Path.Combine(node_f, node_name[0]);
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Format of " + txtItemCode.Text + " not found", "Warning");
                }
                //}
                //else
                //{
                //    src_file_format = "";
                //}


               
                if (src_file_format != "")
                {
                    export_en = true;
                List<string> LotNo_list = new List<string>();
                List<string> ItemCode_list = new List<string>();
                foreach (TreeNode t in tvItemDetails.Nodes)
                {
                    if ((t.Checked) && (t.Nodes.Count > 0))
                    {
                        ItemCode_list.Add(t.Text);
                        foreach (TreeNode c in t.Nodes)
                        {
                            if (c.Checked)
                            {
                                LotNo_list.Add(c.Text);
                            }
                        }
                    }
                }
                if ((ItemCode_list.Count > 0) && (LotNo_list.Count > 0))
                {
                    string tar_folder = myVar.data_loc;
                    string report_folder = Path.Combine(myVar.app_path, "Report");
                    if (!System.IO.Directory.Exists(report_folder))
                        System.IO.Directory.CreateDirectory(report_folder);
                    foreach (string _itemcode in ItemCode_list)
                    {
                            foreach (string _lotno in LotNo_list)
                            {
                                string err_mess = "";
                                myExcel.Application xlsApp = TDMK_Code.StartExcel();
                                xlsApp.Visible = true;
                                //////try
                                //////{
                                /***********************  Export FAI Data **********************************************/

                                xlsApp.DisplayAlerts = false;
                                myExcel.Workbook exp_file;
                                exp_file = xlsApp.Workbooks.Add();
                                //////myCode.Export_FAI_Batch(_itemcode, _lotno, exp_file);
                                err_mess = "FAI export";
                                string export_path = Path.Combine(report_folder, report_name + "-" + _lotno + ".xlsm");
                                myExcel.Workbook report_saved = null;
                                if (System.IO.File.Exists(export_path))
                                {
                                    report_saved = TDMK_Code.open_excel_file(export_path, "", "");
                                }
                                else
                                {
                                    report_saved = TDMK_Code.open_excel_file(src_file_format, "", "");
                                    //report_saved.SaveAs(Path.Combine(report_folder, "NPI OK2Ship Report for Bare FPC_" + _itemcode + _lotno + ".xlsm"));

                                    report_saved.SaveAs(Path.Combine(report_folder, report_name + "-" + _lotno + ".xlsm"));
                                }

                                myCode.Export_To_FAI2(report_saved, _itemcode, _lotno, exp_file);
                                /*********************** Finish export FAI Data **********************************************/

                                /***********************  Export IPQC Data **********************************************/
                                err_mess = "FAI export";
                                myCode.Export_IPQC_Data2(report_saved, _itemcode, _lotno);
                                myCode.DES_Data_Factor(report_saved);
                                myCode.Fill_NA_Data(report_saved, "IPQC Data", "B15", 32);
                                /*********************** Finish export IPQC Data **********************************************/


                                /***********************  Export VHX **********************************************/
                                err_mess = "VHX export";

                                export_update.export_BVH_PTH(report_saved, _itemcode, _lotno);
                                export_update.export_Impedance(report_saved, _itemcode, _lotno);
                                export_update.export_SolderMask(report_saved, _itemcode, _lotno);
                                export_update.export_ThermalStress(report_saved, _itemcode, _lotno);
                                export_update.export_ACF(report_saved, _itemcode, _lotno);
                                export_update.Export_Bhast_Logfile(report_saved, _itemcode, _lotno);
                                export_update.export_Moisture(report_saved, _itemcode, _lotno);

                                /*********************** Finish export VHX **********************************************/



                                /***********************  Export Materials Data **********************************************/
                                //myCode.Export_Materials_Data(report_saved, _itemcode, _lotno);
                                /*********************** Finish export Materials Data **********************************************/

                                /***********************  Export Recycle_PGC_Copper Data **********************************************/
                                err_mess = "Recycle PGC & Copper export";
                                myCode.Export_Recycle_Data(report_saved, _itemcode, _lotno);
                                /*********************** Finish export Recycle_PGC_Copper Data **********************************************/

                                /***********************  Export Others from Excel Data ****************************************************/
                                //err_mess = "Excel file import";
                                ///// myCode.Export_fromExcel2(report_saved, _itemcode, _lotno, tar_folder);
                                /*********************** Finish Export Others from Excel Data **********************************************/

                                /***********************  Fill Build Info ******************************************************************/
                                err_mess = "Build Information data";
                                myCode.fill_data("User_Data", report_saved, _itemcode, _lotno);
                                myCode.Fill_Build_Info(report_saved);

                                /***********************  Finish Fill Build Info ******************************************************************/
                                /***********************  Export Declaration Data **********************************************/
                                err_mess = "Declaration export";
                                myCode.Export_Declaration_Data2(report_saved, _itemcode, _lotno);
                                myCode.fill_PushingDate(report_saved, _itemcode, _lotno);
                                myCode.Fill_BuildInfo_FAI(report_saved);
                                xlsApp.DisplayAlerts = true;
                                MessageBox.Show("Complete export data");
                            


                                    /*********************** Finish Export Declaration Data **********************************************/
                                ////}
                                ////catch (Exception ex)
                                ////{
                                ////    xlsApp.DisplayAlerts = true;
                                ////    myCode.Error_log(myVar.app_path, err_mess);
                                ////    MessageBox.Show("Error at process: " + err_mess, "Warning");
                                ////}
                            
                            }
                    }
                }
                else
                {
                    MessageBox.Show("Please, select ItemCode and LotNo", "Warning");
                }
            }
            }
            if (!export_en)
            {
                MessageBox.Show("Please, approve at the first!", "Warning");
            }





            /************************ Test Areas ************************************************************************************************/
            //string src_Itemcode = "22B0500";
            //string src_Lotno = "00003";
            //string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { src_Itemcode, src_Lotno });
            //string src_file = Path.Combine(myVar.app_path, "Temp_Folder", "NPI OK2Ship Report for Bare FPC_" + src_Itemcode + ".xlsm");
            //DataTable Result_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_Declare, "Results", filter_str);
            //DataSet temp_DS = new DataSet();
            //TDMK_Code.fill_dataset(temp_DS, "Results_Address", myVar.sqlcon_Declare);
            //DataTable results_addr = temp_DS.Tables[0];            
            //myExcel .Workbook mywrkbook = TDMK_Code.open_excel_file(src_file, "", "");
            //myExcel.Worksheet mywrksheet = mywrkbook.Sheets["Declaration and Contents"];
            //foreach (DataRow t in Result_tbl.Rows )
            //{
            //    string result_item = t.Field<string>("Items");                   
            //    string rgn_addr =  TDMK_Code.Get_item_val("Results_Address", "Items like '"+result_item+"'", "Address", myVar.sqlcon_Declare);
            //    mywrksheet .Range [rgn_addr].Value = t.Field<string>("Results");
            //}
            //string src_file = Path.Combine(myVar.app_path, "Temp_Folder", "NPI OK2Ship Report for Bare FPC_22B0500.xlsm");
            //myCode. Export_Declaration_Data(src_file, "22B0500", "00003");



            /************************************************************************************************************************************/


            //string src_file = Path.Combine(myVar.app_path, "Temp_Folder", "NPI OK2Ship Report for Bare FPC_22B050000003.xlsm");//"NPI OK2Ship Report for Bare FPC_22B050000003.xlsm"
            //string tar_folder = @"D:\Customer Projects\SEEV\TestAreas\OK2SHIP automation system";
            ///***********************  Export FAI Data **********************************************/
            //myCode.Export_FAI_Data(src_file, "22B0500", "00003");
            ///*********************** Finish export FAI Data **********************************************/

            ///***********************  Export IPQC Data **********************************************/
            //myCode.Export_IPQC_Data(src_file, "22B0500", "00003");
            ///*********************** Finish export IPQC Data **********************************************/

            ///***********************  Export Materials Data **********************************************/
            //myCode.Export_Materials_Data(src_file, "22B0500", "00003");
            ///*********************** Finish export Materials Data **********************************************/

            ///***********************  Export Recycle_PGC_Copper Data **********************************************/
            //myCode. Export_Recycle_Data(src_file, "22B0500", "00003");
            ///*********************** Finish export Recycle_PGC_Copper Data **********************************************/

            ///***********************  Export Others from Excel Data **********************************************/
            //myCode.Export_fromExcel(src_file, "22B0358", "00003", tar_folder);
            ///*********************** Finish Export Others from Excel Data **********************************************/

            ///***********************  Export Recycle_PGC_Copper Data **********************************************/
            //myCode. Export_Recycle_Data(src_file, "22B0500", "00003");
            ///*********************** Finish export Recycle_PGC_Copper Data **********************************************/

        }


        private void FrmMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            myVar.frm_mainscreen.Show();
        }
        public static void Update_Approve()
        {
            MessageBox.Show("OK");
        }

        private void tsmPass_Click(object sender, EventArgs e)
        {
            DGV_Set_Result(DGV_Data, "Pass");
        }
        public void DGV_Set_Result(DataGridView src_DGV, string result_val)
        {
            string before_val = myCode.checkDBNull(src_DGV.CurrentCell.Value);
            src_DGV.CurrentCell.Value = result_val;
            int col_inx = src_DGV.CurrentCell.ColumnIndex;
            int row_inx = src_DGV.CurrentCell.RowIndex;
            int id = Convert.ToInt32(src_DGV.Rows[row_inx].Cells["ID"].Value);
            if ((result_val == "Fail") || (result_val == ""))
            {
                for (int i = col_inx + 1; i < src_DGV.Columns.Count - 1; i++)
                {
                    src_DGV.Rows[row_inx].Cells[i].Value = "";
                }
            }
            List<string> items = new List<string>();
            List<string> items_val = new List<string>();
            for (int c_inx = 1; c_inx < src_DGV.Columns.Count - 1; c_inx++)
            {
                items.Add(src_DGV.Columns[c_inx].Name);
                items_val.Add(src_DGV.Rows[row_inx].Cells[c_inx].Value.ToString());
            }
            TDMK_Code.updatebyID_val_arr_Null("Approve_Data", myVar.sqlcon_OK2SHIP, id, items.ToArray(), items_val.ToArray());
            string itemcode = myCode.checkDBNull(src_DGV.Rows[row_inx].Cells["ItemCode"].Value);
            string lotno = myCode.checkDBNull(src_DGV.Rows[row_inx].Cells["LotNo"].Value);
            string depart = myVar.USerDepart;
            string itemEdit = src_DGV.Columns[col_inx].Name;
            int hist_id = TDMK_Code.SQL_MAX("Confirm_History", "ID", myVar.sqlcon_OK2SHIP) + 1;
            TDMK_Code.insert_val_arr2("Confirm_History", myVar.sqlcon_OK2SHIP, new string[] { "ID", "ItemCode", "LotNo", "UserID", "Depart", "Item_Edit", "Before_Edit", "After_Edit", "Time_Edit" }, new string[] { hist_id.ToString(), itemcode, lotno, myVar.UserID, myVar.USerDepart, itemEdit, before_val, result_val, DateTime.Now.ToString() });
        }
        private void tsmFail_Click(object sender, EventArgs e)
        {
            DGV_Set_Result(DGV_Data, "Fail");
        }
        private void tsmBlank_Click(object sender, EventArgs e)
        {
            DGV_Set_Result(DGV_Data, "");
        }
        private void tsmNA_Click(object sender, EventArgs e)
        {
            DGV_Set_Result(DGV_Data, "NA");
        }

        private void DGV_Data_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (MessageBox.Show("Do you want to export data ?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                FolderBrowserDialog f_open = new FolderBrowserDialog();
                f_open.SelectedPath = myVar.app_path;
                if (f_open.ShowDialog() == DialogResult.OK)
                {
                    string exp_fld = f_open.SelectedPath;
                    string cur_date = DateTime.Now.ToString("yyyyMMddHHmmss");
                    string f_name = Path.Combine(exp_fld, "OK2ShipReport_" + cur_date + ".csv");
                    ToFTP((DataTable)DGV_Data.DataSource, f_name, "ok2shipreport", "ok2shipreport00");
                }
            }

            //DataSet ds = new DataSet();
            //TDMK_Code.fill_dataset(ds, "Confirm_History", myVar.sqlcon_OK2SHIP);
            //DataTable src_dt = ds.Tables[0];
            //DataTable History_dt = Get_Historical_data(src_dt, "08:30:00", "10:20:00");
            //List<string> tar_list = History_dt.AsEnumerable().Select(x => x.Field<string>("ItemCode")).ToList();
            //DGV_Data.DataSource = get_approve_data(tar_list, myVar.sqlcon_OK2SHIP);
        }
        public void ToCSV(DataTable dtDataTable, string strFilePath)
        {
            Dictionary<int, List<string>> data_result = new Dictionary<int, List<string>>();
            int inx = 0;
            foreach (DataRow dr in dtDataTable.Rows)
            {
                string QA_PIC_appr = myCode.checkDBNull(dr["QA_PIC_Approve"]);
                string QA_Manager_appr = myCode.checkDBNull(dr["QA_Manager"]);
                string Customer_appr = myCode.checkDBNull(dr["Customer_Approve"]);
                string result = "";
                if (QA_PIC_appr == "Pass")
                {
                    if (Customer_appr == "NA")
                    {
                        result = QA_Manager_appr;
                    }
                    else
                    {
                        result = Customer_appr;
                    }
                }
                else
                {
                    result = QA_PIC_appr;
                }
                if (result != "")
                {
                    List<string> exp_result = new List<string>();
                    exp_result.Add(dr["ItemCode"].ToString());
                    exp_result.Add(dr["LotNo"].ToString());
                    exp_result.Add("2");
                    switch (result)
                    {
                        case "Pass":
                            exp_result.Add("9");
                            break;
                        case "Fail":
                            exp_result.Add("1");
                            break;
                    }
                    data_result.Add(inx, exp_result);
                    inx++;
                }
            }
            if (data_result.Count > 0)
            {
                StreamWriter sw = new StreamWriter(strFilePath, false);
                List<string> title_lst = new List<string>() { "item_c", "batchno", "process", "result" };
                string data1 = "";
                for (int i = 0; i < title_lst.Count; i++)
                {
                    data1 = data1 + title_lst[i] + ",";
                }
                data1 = data1.TrimEnd(',');
                sw.WriteLine(data1);
                for (int i = 0; i < data_result.Count; i++)
                {
                    List<string> _exp_result = data_result[i];
                    string data = "";
                    for (int j = 0; j < _exp_result.Count; j++)
                    {
                        if (j == 1)
                        {
                            data = data + "'" + _exp_result[j] + ",";
                        }
                        else
                        {
                            data = data + _exp_result[j] + ",";
                        }
                    }
                    data = data.TrimEnd(',');
                    sw.WriteLine(data);
                }
                sw.Close();
                MessageBox.Show("Export completed!", "Information");
            }
        }
        public void ToCSV2(DataTable dtDataTable, string strFilePath)
        {
            Dictionary<int, List<string>> data_result = new Dictionary<int, List<string>>();
            int inx = 0;
            foreach (DataRow dr in dtDataTable.Rows)
            {
                string QA_PIC_appr = myCode.checkDBNull(dr["QA_PIC_Approve"]);
                string QA_Manager_appr = myCode.checkDBNull(dr["QA_Manager"]);
                string Customer_appr = myCode.checkDBNull(dr["Customer_Approve"]);
                string result = "";
                if (QA_PIC_appr == "Pass")
                {
                    if (Customer_appr == "NA")
                    {
                        result = QA_Manager_appr;
                    }
                    else
                    {
                        result = Customer_appr;
                    }
                }
                else
                {
                    result = QA_PIC_appr;
                }
                if (result != "")
                {
                    List<string> exp_result = new List<string>();
                    exp_result.Add(dr["ItemCode"].ToString());
                    exp_result.Add(dr["LotNo"].ToString());
                    exp_result.Add("2");
                    switch (result)
                    {
                        case "Pass":
                            exp_result.Add("9");
                            break;
                        case "Fail":
                            exp_result.Add("1");
                            break;
                    }
                    data_result.Add(inx, exp_result);
                    inx++;
                }
            }
            if (data_result.Count > 0)
            {
                List<string> title_lst = new List<string>() { "item_c", "batchno", "process", "result" };
                myExcel.Workbook wrkbk = TDMK_Code.Create_workbook();
                myExcel.Worksheet wrksht = wrkbk.Worksheets[1];
                wrksht.Columns.NumberFormat = "@";
                myExcel.Range rgn = wrksht.Range["A1"];
                for (int i = 0; i < title_lst.Count; i++)
                {
                    rgn.Offset[0, i].Value = title_lst[i];
                }
                for (int i = 0; i < data_result.Count; i++)
                {
                    List<string> _exp_result = data_result[i];
                    for (int j = 0; j < _exp_result.Count; j++)
                    {
                        rgn.Offset[1 + i, j].Value = _exp_result[j];
                    }

                }
                wrksht.SaveAs(strFilePath, myExcel.XlFileFormat.xlCSV, myExcel.XlSaveAsAccessMode.xlShared);
                //wrkbk.SaveAs(strFilePath);
                wrkbk.Close();
                MessageBox.Show("Export completed!", "Information");
            }
        }
        public void ToCSV3(DataTable dtDataTable, string strFilePath)
        {
            Dictionary<int, List<string>> data_result = new Dictionary<int, List<string>>();
            int inx = 0;
            foreach (DataRow dr in dtDataTable.Rows)
            {
                string QA_PIC_appr = myCode.checkDBNull(dr["QA_PIC_Approve"]);
                string QA_Manager_appr = myCode.checkDBNull(dr["QA_Manager"]);
                string Customer_appr = myCode.checkDBNull(dr["Customer_Approve"]);
                string result = "";
                if (QA_PIC_appr == "Pass")
                {
                    if (Customer_appr == "NA")
                    {
                        result = QA_Manager_appr;
                    }
                    else
                    {
                        result = Customer_appr;
                    }
                }
                else
                {
                    result = QA_PIC_appr;
                }
                if (result != "")
                {
                    List<string> exp_result = new List<string>();
                    exp_result.Add(dr["ItemCode"].ToString());
                    exp_result.Add(dr["LotNo"].ToString());
                    exp_result.Add("2");
                    switch (result)
                    {
                        case "Pass":
                            exp_result.Add("9");
                            break;
                        case "Fail":
                            exp_result.Add("1");
                            break;
                    }
                    data_result.Add(inx, exp_result);
                    inx++;
                }
            }
            if (data_result.Count > 0)
            {
                //StreamWriter sw = new StreamWriter(strFilePath, false);
                using (var sw = System.IO.File.Create(strFilePath))
                {
                    List<string> title_lst = new List<string>() { "item_c", "batchno", "process", "result" };
                    //for (int i = 0; i < title_lst.Count; i++)
                    //{
                    //    sw.Write(title_lst[i]);
                    //    if (i < title_lst.Count - 1)
                    //    {
                    //        sw.Write(",");
                    //    }
                    //}

                    for (int i = 0; i < data_result.Count; i++)
                    {
                        List<string> _exp_result = data_result[i];
                        for (int j = 0; j < _exp_result.Count; j++)
                        {
                            if (!Convert.IsDBNull(_exp_result[j]))
                            {
                                var data = Encoding.ASCII.GetBytes(_exp_result[j] + ",");
                                sw.Write(data, 0, data.Length);
                            }
                        }
                    }
                    sw.Close();
                }
                MessageBox.Show("Export completed!", "Information");
            }
        }
        public void ToFTP(DataTable dtDataTable, string strFilePath, string ftpUsername, string ftpPassword)
        {
            Dictionary<int, List<string>> data_result = new Dictionary<int, List<string>>();
            int inx = 0;
            foreach (DataRow dr in dtDataTable.Rows)
            {
                string QA_PIC_appr = myCode.checkDBNull(dr["QA_PIC_Approve"]);
                string QA_Manager_appr = myCode.checkDBNull(dr["QA_Manager"]);
                string Customer_appr = myCode.checkDBNull(dr["Customer_Approve"]);
                string result = "";
                if (QA_PIC_appr == "Pass")
                {
                    if (Customer_appr == "NA")
                    {
                        result = QA_Manager_appr;
                    }
                    else
                    {
                        result = Customer_appr;
                    }
                }
                else
                {
                    result = QA_PIC_appr;
                }
                if (result != "")
                {
                    List<string> exp_result = new List<string>();
                    exp_result.Add(dr["ItemCode"].ToString());
                    exp_result.Add(dr["LotNo"].ToString());
                    exp_result.Add("2");
                    switch (result)
                    {
                        case "Pass":
                            exp_result.Add("9");
                            break;
                        case "Fail":
                            exp_result.Add("1");
                            break;
                    }
                    data_result.Add(inx, exp_result);
                    inx++;
                }
            }
            if (data_result.Count > 0)
            {
                StreamWriter sw = new StreamWriter(strFilePath, false);
                List<string> title_lst = new List<string>() { "item_c", "batchno", "process", "result" };
                string data1 = "";
                for (int i = 0; i < title_lst.Count; i++)
                {
                    data1 = data1 + title_lst[i] + ", ";
                }
                data1 = data1.Trim().TrimEnd(',');
                sw.WriteLine(data1);
                for (int i = 0; i < data_result.Count; i++)
                {
                    List<string> _exp_result = data_result[i];
                    string data = "";
                    for (int j = 0; j < _exp_result.Count; j++)
                    {
                        data = data + _exp_result[j] + ", ";
                    }
                    data = data.Trim().TrimEnd(',');
                    sw.WriteLine(data);
                }
                sw.Close();
                using (var client = new WebClient())
                {
                    client.Credentials = new NetworkCredential(ftpUsername, ftpPassword);
                    client.UploadFile("ftp://10.212.1.246/" + Path.GetFileName(strFilePath), WebRequestMethods.Ftp.UploadFile, strFilePath);
                }
                MessageBox.Show("Export completed!", "Information");
            }
        }

        private void DGV_Data_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            List<string> tar_list = new List<string>() { "QA_PIC_Approve", "NPI_PE_Manager", "QA_Manager", "Customer_Approve" };
            int col_inx = e.ColumnIndex;
            int row_inx = e.RowIndex;
            if ((col_inx != -1) && (row_inx != -1))
            {
                string sel_item = DGV_Data.Columns[col_inx].Name;
                if (tar_list.IndexOf(sel_item) != -1)
                {
                    string itemcode = DGV_Data.Rows[row_inx].Cells["ItemCode"].Value.ToString();
                    string lotno = DGV_Data.Rows[row_inx].Cells["LotNo"].Value.ToString();
                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Item_Edit" }, new string[] { itemcode, lotno, sel_item });
                    DataTable src_dt = TDMK_Code.Datatable_Filter(myVar.sqlcon_OK2SHIP, "Confirm_History", filter_str);
                    string tool_str = "";
                    foreach (DataRow dr in src_dt.Rows)
                    {
                        tool_str = tool_str + "User: " + dr["UserID"] + " / Depart: " + dr["Depart"] + " / Time_Edit: " + dr["Time_Edit"] + " / Confirm value: " + dr["After_Edit"] + "\r\n";
                    }
                    e.ToolTipText = tool_str.Trim();
                }
            }
        }

        public DataTable get_approve_data(List<string> ItemCode_list, SqlConnection sqlcon)
        {
            DataTable result = new DataTable();
            DataSet ds = new DataSet();
            TDMK_Code.fill_dataset(ds, "Approve_Data", sqlcon);
            DataTable dtDataTable = ds.Tables[0];
            var rowsToDelete = from r in dtDataTable.AsEnumerable() join c in ItemCode_list on r.Field<string>("ItemCode") equals c into g where g.Any() select r;
            result = rowsToDelete.CopyToDataTable();
            return result;
        }
        public DataTable Get_Historical_data(DataTable src_dt, string start_time, string stop_time)
        {
            DataView dview = src_dt.AsDataView();
            string from_date = DateTime.Now.ToShortDateString() + " " + start_time;//new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 8, 30, 0).ToString();
            string to_date = DateTime.Now.ToShortDateString() + " " + stop_time;//new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 17, 30, 0).ToString();
            dview.RowFilter = "Time_Edit >= #" + from_date + "# And Time_Edit <= #" + to_date + "#";
            return dview.ToTable();
        }



    }
} 