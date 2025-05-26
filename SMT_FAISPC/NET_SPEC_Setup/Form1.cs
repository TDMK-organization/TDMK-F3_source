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
using myExcel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.IO;
using IniLibs;

namespace NET_SPEC_Setup
{
    public partial class Form1 : Form
    {
        SEI_Lib myCode = new SEI_Lib();
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SqlConnection sqlcon = new SqlConnection();
        public string server_name;
        public string server_acc;
        public string server_pass;
        public string DB_name;
        public string app_path;
        public string Data_Location;
        public string format_folder;
        public string log_folder;
        //string connstr_OK2SHIP;
        public IniFile TDMK_init = new IniFile("Config.ini");
        public Form1()
        {
            InitializeComponent();
        }

        private void tsmLoad_data_Click(object sender, EventArgs e)
        {
            if(txtItemCode.Text!="")
            {
                //DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "NET_SPEC", "ItemCode = '" + txtItemCode.Text + "'");
                DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "NET_SPEC", TDMK_Code.filter_str( new string[] { "ItemCode","Remark"},new string[] { txtItemCode.Text, cbProcess.Text}));
                DGV_Data.DataSource = spec_dt;
                if(spec_dt.Rows.Count==0)
                {
                    if(MessageBox.Show("No data! \r\n Do you want to setup ?","Warning",MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        myExcel.Workbook spec_wrkbk = TDMK_Code.Create_workbook();
                        myExcel.Worksheet spec_wrksht = spec_wrkbk.Sheets[1];
                        int col_inx = 0;
                        foreach (DataColumn dc in spec_dt.Columns)
                        {
                            spec_wrksht.Range["A1"].Offset[0, col_inx].Value = dc.ColumnName;
                            col_inx++;
                        }
                        spec_wrkbk.Activate();
                    }
                }
            }
            else
            {
                MessageBox.Show("Please, enter ItemCode!", "Warning");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            sqlcon = initial_data("OK2SHIP_SMT",true);
            List<string> tar_process = new List<string>() {"FLEX_BENDING", "THERMAL_CYCLING_AND_BEND", "HEAT_SOAK_AND_BEND" };
            cbProcess.DataSource = tar_process;
        }
        public SqlConnection initial_data(string DB_name)
        {
            SqlConnection _sqlcon_OK2SHIP;
            string app_path = Application.StartupPath;
            string config_file = Path.Combine(app_path, "Config", "config.txt");
            string[] my_config = myCode.read_config_arr(config_file);
            string server_name = "";
            string server_acc = "";
            string server_pass = "";
            string Data_Location = "";
            string format_folder = "";
            string log_folder = "";
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
                if (c.Contains("Data_Location"))
                {
                    string[] temp = c.Split(':');
                    string tg = "";
                    if (temp.Length > 2)
                    {

                        for (int t = 1; t < temp.Length; t++)
                        {
                            tg = tg + temp[t] + ":";
                        }
                    }
                    Data_Location = tg.TrimEnd(':').Trim();
                }
                if (c.Contains("Format_Folder"))
                {
                    format_folder = c.Split('#')[1].Trim();
                }
                if (c.Contains("Log_folder"))
                {
                    log_folder = c.Split('#')[1].Trim();
                }
            }
            string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, DB_name, server_acc, server_pass).ConnectionString;
            _sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);
            return _sqlcon_OK2SHIP;
        }

        private void tsmImport_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            TDMK_Code.fill_dataset(ds, "NET_SPEC", sqlcon);
            DataTable spec_tbl = ds.Tables[0];
            DataTable cur_tbl = spec_tbl.Copy();
            cur_tbl.Rows.Clear();
            myExcel.Application xlapp = TDMK_Code.StartExcel();
            myExcel.Workbook cur_wrkbk = xlapp.ActiveWorkbook;
            if(cur_wrkbk!=null)
            {
                myExcel.Worksheet cur_wrksht = cur_wrkbk.Sheets[1];
                myExcel.Range cur_rgn = cur_wrksht.Range["E2"];
                int r_inx = 0;

                while (myCode.checkDBNull(cur_rgn.Offset[r_inx, 0].Value) != "")
                {
                    DataRow dr = cur_tbl.NewRow();
                    for (int i = 0; i < cur_tbl.Columns.Count; i++)
                    {
                        dr[i] = cur_rgn.Offset[r_inx, i - 4].Value;
                    }
                    cur_tbl.Rows.Add(dr);
                    r_inx++;
                }
                DGV_Data.DataSource = cur_tbl;
            }
            else
            {
                if( MessageBox.Show("Excel format file is closed!\r\nDo you want to re-open it?", "Warning",MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    myExcel.Workbook spec_wrkbk = TDMK_Code.Create_workbook();
                    myExcel.Worksheet spec_wrksht = spec_wrkbk.Sheets[1];
                    int col_inx = 0;
                    foreach (DataColumn dc in cur_tbl.Columns)
                    {
                        spec_wrksht.Range["A1"].Offset[0, col_inx].Value = dc.ColumnName;
                        col_inx++;
                    }
                    spec_wrkbk.Activate();
                }
            }
        }

        private void tsmExport_Click(object sender, EventArgs e)
        {
            if(DGV_Data.Rows.Count>0)
            {
                myExcel.Workbook exp_wrkbk = TDMK_Code.Create_workbook();
                TDMK_Code.Export_DGV_Excel3(DGV_Data, exp_wrkbk, true);
                myExcel.Worksheet exp_sht = exp_wrkbk.ActiveSheet;
                exp_sht.Range["A:A"].Delete();
            }
            else
            {
                MessageBox.Show("No data!", "Warning");
            }

        }
        private void tsmSave_Click(object sender, EventArgs e)
        {
            DataTable myDt = (DataTable)DGV_Data.DataSource;
            List<DataTable> mylst_dt = new List<DataTable>();
            AutoCompleteStringCollection lst_ItemCode = TDMK_Code.Load_Item_Names("NET_SPEC", "ItemCode", sqlcon);
            Get_ListTable(-1, myDt, new string[] { "ItemCode" }, ref mylst_dt, "Point+V");
            foreach(DataTable dt in mylst_dt)
            {
                bool pro_en = false;
                string itemcode = dt.Rows[0][1].ToString();
                if(TDMK_Code.check_exist_list(itemcode,lst_ItemCode))
                {
                    if(MessageBox.Show("ItemCode "+ itemcode +" is existed! \r\nDo you want to update?","Warning", MessageBoxButtons.YesNo)==DialogResult.Yes)
                    {
                        pro_en = true;
                    }
                    else
                    {
                        pro_en = false;
                    }
                }
                else
                {
                    pro_en = true;
                }
                if(pro_en)
                {
                    TDMK_Code.Delelte_FilteredItem_arr("NET_SPEC", sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { txtItemCode.Text, cbProcess.Text }));
                    int id = TDMK_Code.SQL_MAX("NET_SPEC", "ID", sqlcon);
                    int r_inx = 0;
                    foreach (DataRow dr in dt.Rows)
                    {
                        dr[0] = (id + r_inx + 1).ToString();
                        r_inx++;
                    }
                    BatchBulkCopy(sqlcon, dt, "NET_SPEC");
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
        public void Get_List_data(int col_inx, DataTable myDt, string[] src_arr, ref List<string> src_lst_data, string tar_item, bool distinct_en)
        {

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
                                Get_List_data(col_inx + 1, curTbl, src_arr, ref src_lst_data, tar_item, distinct_en);
                            }
                        }
                    }
                    else
                    {
                        if (distinct_en)
                        {
                            src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Distinct().ToArray());
                        }
                        else
                        {
                            src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).ToArray());
                        }
                    }
                }
            }
            else
            {
                if (distinct_en)
                {
                    src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Distinct().ToArray());
                }
                else
                {
                    src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).ToArray());
                }
            }
        }
        public void BatchBulkCopy(SqlConnection sqlcon_OK2SHIP, DataTable dataTable, string DestinationTbl)
        {
            DataTable dtInsertRows = dataTable;
            using (SqlBulkCopy sbc = new SqlBulkCopy(sqlcon_OK2SHIP))
            {
                if (sqlcon_OK2SHIP.State != ConnectionState.Open)
                {
                    sqlcon_OK2SHIP.Open();
                }
                sbc.DestinationTableName = DestinationTbl;
                foreach (DataColumn dc in dataTable.Columns)
                {
                    sbc.ColumnMappings.Add(dc.ColumnName, dc.ColumnName);
                }
                sbc.WriteToServer(dtInsertRows);
                sqlcon_OK2SHIP.Close();
            }
        }
        public SqlConnection initial_data(string DB_name, bool sa_en)
        {
            SqlConnection _sqlcon_OK2SHIP;
            if (!File.Exists("Config.ini"))
            {
                TDMK_init.Write("Server", "10.212.1.243", "SMT_Config");
                TDMK_init.Write("Account", "sa", "SMT_Config");
                TDMK_init.Write("Password", "seev@123;", "SMT_Config");
                TDMK_init.Write("Report_Location", @"\\10.212.1.200\rgo\OK2SHIP automation system", "SMT_Config");
                TDMK_init.Write("Data_Location", @"\\10.212.1.200\rgo\OK2SHIP automation system\TDMK_Data\TestAreas\OK2SHIP_Measurements\OMM", "SMT_Config");
                TDMK_init.Write("Format_Folder", @"\\10.212.1.200\rgo\OK2SHIP automation system\0.OK2SHIP report format", "SMT_Config");
                TDMK_init.Write("Log_folder", @"\\10.212.1.200\rgo\OK2SHIP automation system\TDMK_Data\TestAreas\OK2SHIP_Measurements", "SMT_Config");
            }
            server_name = TDMK_init.Read("Server", "SMT_Config");
            server_acc = TDMK_init.Read("Account", "SMT_Config");
            server_pass = TDMK_init.Read("Password", "SMT_Config");
            Data_Location = TDMK_init.Read("Data_Location", "SMT_Config");
            format_folder = TDMK_init.Read("Format_Folder", "SMT_Config");
            log_folder = TDMK_init.Read("Log_folder", "SMT_Config");
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
            return _sqlcon_OK2SHIP;
        }

        private void cbProcess_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(DGV_Data.Rows.Count!=0)
            {
                if(DGV_Data.DataSource!=null)
                {
                    DGV_Data.DataSource = null;
                }
                else
                {
                    DGV_Data.Rows.Clear();
                    DGV_Data.Columns.Clear();
                }
            }          
        }
    }
}
