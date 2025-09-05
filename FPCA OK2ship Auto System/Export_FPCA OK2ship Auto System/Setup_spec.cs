using Microsoft.Office.Interop.Excel;
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
using DataTable = System.Data.DataTable;
using static System.Net.Mime.MediaTypeNames;
using System.Diagnostics;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Security.Cryptography;
using System.Runtime.InteropServices;
using System.Diagnostics.Eventing.Reader;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OK2SHIP_SMT;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using Bending_Export;
using Export_FPCA_OK2ship_Auto_System.Services;
using Export_FPCA_OK2ship_Auto_System.Repositories;
using IniLibs;
using Export_FPCA_OK2ship_Auto_System.Libary;


namespace Export_FPCA_OK2ship_Auto_System
{
    public partial class Setup_Spec_SMT : Form
    {
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public TDMK_EPPLUS TDMK_Code2 = new TDMK_EPPLUS();
        // public EPPlus_Lib TDMK_Code2 = new EPPlus_Lib();
        public SEI_Lib myCode = new SEI_Lib();
        public SqlConnection sqlcon = null;
        public string strcon = "";
        string DB_name = "OK2SHIP_SMT";
        string data_loc = "";
        public SqlConnection sqlcon_F1 = null;
        public Bending_Export_Lib Bending_Exp = new Bending_Export_Lib();

        IniFile TDMK_init = new IniFile();

        public Setup_Spec_SMT()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // sqlcon = initial_data(DB_name, true);
            // sqlcon_F1 = initial_data_F1();
            string app_path = System.Windows.Forms.Application.StartupPath;
            sqlcon = Bending_Exp.initial_data("OK2SHIP_SMT", true);
            // data_loc = Bending_Exp.format_folder;
            //data_loc = System.Windows.Forms.Application.StartupPath;
            data_loc = Bending_Exp.find_config_path(app_path.Replace("\\FPCA OK2SHIP Auto System", ""), "SEEV Data");
            string config_path = Path.Combine(app_path.Replace("\\FPCA OK2SHIP Auto System", ""), "config.ini");
            TDMK_init = new IniFile(config_path);
            txtFormat.Text = TDMK_init.Read("Format_Folder", "SMT_Config") + $"\\SEEV Data";

            cbAll.Checked = true;
            for (int i = 0; i < cbl_sheet.Items.Count; i++)
            {
                cbl_sheet.SetItemChecked(i, true);
            }

            lst_Item.Items.Clear();
            DataTable dt_item_setup = Datatable_ALL(sqlcon, "SPEC_COMMENT_3");
            string[] arr_items = dt_item_setup.AsEnumerable().Select(x => x.Field<string>("ItemCode")).Distinct().ToArray();
            foreach (string item in arr_items)
            {
                lst_Item.Items.Add(item);
            }

            //fill_dgv_setup_detail(); 
        }

        public void fill_dgv_setup_detail()
        {
            DataTable dt = new DataTable();

            dt.Columns.Add("ID");
            dt.Columns.Add("ItemCode");
            dt.Columns.Add("Type");
            foreach (string sheet in cbl_sheet.Items)
            {
                dt.Columns.Add(sheet.Replace(" ", "_").Replace("(", "").Replace(")", "").ToUpper());
            }
            DataTable dt_item_setup = Datatable_ALL(sqlcon, "SPEC_COMMENT_3");
            // DataTable dt_spec_FAI = Datatable_ALL(sqlcon, "FAI_Spec");
            DataTable FAI_spec_dt = Datatable_ALL(sqlcon, "FAI_Spec");
            List<string> arr_items = FAI_spec_dt.AsEnumerable().Select(x => x.Field<string>("ItemCode")).Distinct().ToList();

            string[] _arr_items = dt_item_setup.AsEnumerable().Select(x => x.Field<string>("ItemCode")).Distinct().ToArray();
            arr_items.AddRange(_arr_items.Except(arr_items));
            string[] arr_remark = new string[] { "NPI", "MASS", "Other" };
            int id = 1;
            foreach (string item in arr_items)
            {
                foreach (string type in arr_remark)
                {
                    DataRow dr = dt.NewRow();
                    dr[0] = id;
                    dr[1] = item;
                    dr[2] = type;
                    List<string> lst_col_setup = dt_item_setup.AsEnumerable().Where(s => s.Field<string>("ItemCode") == item && s.Field<string>("Remark") == type).Select(x => x.Field<string>("Sheet")).ToList();

                    for (int i = 3; i < dt.Columns.Count; i++)
                    {
                        if (lst_col_setup.IndexOf(dt.Columns[i].ColumnName) != -1)
                        {
                            dr[i] = "x";
                        }
                        else if (dt.Columns[i].ColumnName == "FAI")

                        {
                            DataTable dt_spec_FAI = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Spec", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { item, type }));
                            if (dt_spec_FAI.Rows.Count > 0)
                            {
                                dr[i] = "x";
                            }
                        }
                    }

                    dt.Rows.Add(dr);
                    id++;
                }
            }
            dgv_setup_detail.DataSource = dt;

        }

        public void filter_ItemCode(string item)
        {
            DataTable dt = new DataTable();

            dt.Columns.Add("ID");
            dt.Columns.Add("ItemCode");
            dt.Columns.Add("Type");
            foreach (string sheet in cbl_sheet.Items)
            {
                dt.Columns.Add(sheet.Replace(" ", "_").Replace("(", "").Replace(")", "").ToUpper());
            }

            DataTable dt_item_setup = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { item }));
            DataTable FAI_Spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Spec", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { item }));
            if (dt_item_setup.Rows.Count != 0 || FAI_Spec_dt.Rows.Count != 0)
            {
                string[] arr_remark = new string[] { "NPI", "MASS", "Other" };
                int id = 1;

                foreach (string type in arr_remark)
                {
                    DataRow dr = dt.NewRow();
                    dr[0] = id;
                    dr[1] = item;
                    dr[2] = type;

                    List<string> lst_col_setup = dt_item_setup.AsEnumerable().Where(x => x.Field<string>("Remark") == type).Select(x => x.Field<string>("Sheet")).ToList();
                    if (FAI_Spec_dt.AsEnumerable().Any(x => x.Field<string>("Remark") == type))
                    {
                        lst_col_setup.Add("FAI");
                    }
                    for (int i = 3; i < dt.Columns.Count; i++)
                    {
                        if (lst_col_setup.IndexOf(dt.Columns[i].ColumnName) != -1)
                            dr[i] = "x";
                    }
                    dt.Rows.Add(dr);
                    id++;
                }
                dgv_setup_detail.DataSource = dt;
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "ItemCode  " + txt_ItemCode_filter.Text + "  chưa được cài đặt", "Thông báo");
            }

        }

        public void filter_ItemCode_not_msg(string item)
        {
            DataTable dt = new DataTable();

            dt.Columns.Add("ID");
            dt.Columns.Add("ItemCode");
            dt.Columns.Add("Type");
            foreach (string sheet in cbl_sheet.Items)
            {
                dt.Columns.Add(sheet.Replace(" ", "_").Replace("(", "").Replace(")", "").ToUpper());
            }

            DataTable dt_item_setup = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { item }));
            if (dt_item_setup.Rows.Count > 0)
            {
                string[] arr_remark = new string[] { "NPI", "MASS", "Other" };
                int id = 1;

                foreach (string type in arr_remark)
                {
                    DataRow dr = dt.NewRow();
                    dr[0] = id;
                    dr[1] = item;
                    dr[2] = type;

                    List<string> lst_col_setup = dt_item_setup.AsEnumerable().Where(x => x.Field<string>("Remark") == type).Select(x => x.Field<string>("Sheet")).ToList();
                    for (int i = 3; i < dt.Columns.Count; i++)
                    {
                        if (lst_col_setup.IndexOf(dt.Columns[i].ColumnName) != -1)
                            dr[i] = "x";
                    }
                    dt.Rows.Add(dr);
                    id++;
                }

                dgv_setup_detail.DataSource = dt;
            }
        }

        public DataTable Datatable_ALL(SqlConnection database_conn, string tbl_name)
        {
            DataSet dataSet = new DataSet();
            XElement xElement = new XElement(XName.Get("SQL", ""));
            xElement.Add("\n                SELECT *\n                FROM [");
            xElement.Add(tbl_name);
            xElement.Add("]\n               ");
            xElement.Add("ORDER BY ID;\n            ");
            string value = xElement.Value;
            if (database_conn.State != ConnectionState.Open)
            {
                database_conn.Open();
            }

            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(value, database_conn);
            sqlDataAdapter.Fill(dataSet, tbl_name);
            DataTable result = dataSet.Tables[tbl_name];
            database_conn.Close();
            return result;
        }

        public SqlConnection initial_data_F1()
        {
            SqlConnection _sqlcon_F1 = null;

            string server_name = "10.212.1.243";
            string server_acc = "sa";
            string server_pass = "seev@123;";
            string DB_name = "OK2SHIP_Period2";

            string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, DB_name, server_acc, server_pass).ConnectionString;
            _sqlcon_F1 = new SqlConnection(connstr_OK2SHIP);

            return _sqlcon_F1;
        }

        //public SqlConnection initial_data(string DB_name, bool sa_en)
        //{
        //    SqlConnection _sqlcon_OK2SHIP;
        //    string app_path = System.Windows.Forms.Application.StartupPath;
        //    string config_file = Path.Combine(app_path, "Config", "config.txt");
        //    string[] my_config = myCode.read_config_arr(config_file);
        //    string server_name = "";
        //    string server_acc = "";
        //    string server_pass = "";


        //    foreach (string c in my_config)
        //    {
        //        if (c.Contains("Server"))
        //        {
        //            server_name = c.Split(':')[1].Trim();
        //        }
        //        if (c.Contains("Account"))
        //        {
        //            server_acc = c.Split(':')[1].Trim();
        //        }
        //        if (c.Contains("Password"))
        //        {
        //            server_pass = c.Split(':')[1].Trim();
        //        }
        //        if (c.Contains("Data_Location"))
        //        {
        //            data_loc = c.Split('#')[1].Trim();
        //        }
        //    }
        //    if (sa_en)
        //    {
        //        string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, DB_name, server_acc, server_pass).ConnectionString;
        //        _sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);
        //    }
        //    else
        //    {
        //        string _strcon = TDMK_Code.data_connection2(server_name, DB_name).ConnectionString;
        //        _sqlcon_OK2SHIP = new SqlConnection(_strcon);
        //    }

        //    return _sqlcon_OK2SHIP;
        //}

        public void BatchBulkCopy(SqlConnection sqlcon, DataTable dataTable, string tablename)
        {
            DataTable dtInsertRows = dataTable;
            using (SqlBulkCopy sbc = new SqlBulkCopy(sqlcon))
            {
                if (sqlcon.State != ConnectionState.Open)
                {
                    sqlcon.Open();
                }
                sbc.DestinationTableName = tablename;
                foreach (DataColumn dc in dataTable.Columns)
                {
                    sbc.ColumnMappings.Add(dc.ColumnName, dc.ColumnName);
                }
                sbc.WriteToServer(dtInsertRows);
                sqlcon.Close();
            }
        }



        public Dictionary<string, List<string>> setup_spec_cross_section(myExcel.Worksheet ws)
        {
            //  AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            string info = "";
            string[] arr_region = { "Ngang", "Doc" };
            Dictionary<string, List<string>> dic_spec = new Dictionary<string, List<string>> { };

            int count_sample = 0;


            for (int i = 1; i < 100; i++)
            {
                for (int j = 1; j < 5; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Contains("Sample"))
                    {
                        while (myCode.checkDBNull(ws.Cells[i, j + count_sample].Value).Contains("Sample"))
                        {
                            count_sample++;
                        }
                        dic_spec.Add("count_sample", new List<string> { count_sample.ToString() });

                        break;
                    }
                }

                if (count_sample != 0)
                    break;

            }
            int region = 1;
            for (int i = 1; i < 100; i++)
            {
                for (int j = 1; j < 10; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Contains("Component"))
                    {
                        List<string> lst_spec = new List<string>();
                        lst_spec.Add(i.ToString() + ";" + j.ToString());
                        int r_offset = 2;
                        string a = myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value);
                        while (!myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value).Contain("Judgement"))
                        {
                            if (myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value) != "")
                            {
                                lst_spec.Add(myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value) + ";" + (i + r_offset).ToString() + ";" + (j - 1).ToString());

                            }
                            else
                            {
                                lst_spec.Add(myCode.checkDBNull((i + r_offset).ToString() + ";" + (j - 1).ToString()));
                            }
                            r_offset++;

                        }

                        dic_spec.Add(region.ToString(), lst_spec);

                        region++;
                        break;
                    }

                }
            }

            return dic_spec;


        }

        public Dictionary<string, List<string>> setup_spec_Peeltest(myExcel.Worksheet ws)
        {
            //  AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            string info = "";
            Dictionary<string, List<string>> dic_spec = new Dictionary<string, List<string>> { };

            int count_sample = 0;

            for (int i = 1; i < 100; i++)
            {
                for (int j = 1; j < 5; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Contains("Sample"))
                    {
                        List<string> lst_spec = new List<string>();
                        while (myCode.checkDBNull(ws.Cells[i, j + count_sample].Value).Contains("Sample"))
                        {
                            count_sample++;
                        }
                        lst_spec.Add(count_sample.ToString());

                        int r_offset = 1;
                        while (!myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value).Contains("Judgement"))
                        {
                            if (myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value) != "")
                            {
                                lst_spec.Add(myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value) + ";" + (i + r_offset).ToString() + ";" + (j - 1).ToString());

                            }
                            r_offset++;

                        }

                        dic_spec.Add(myCode.checkDBNull(ws.Cells[i + 1, j - 2].Value), lst_spec);

                        break;
                    }
                }


            }



            return dic_spec;


        }
        public string lst_spec_comment3_onproduct_old(myExcel.Worksheet ws)
        {
            //  AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            string lst_spec = "";
            Dictionary<string, List<string>> dic_spec = new Dictionary<string, List<string>> { };


            int count_sample = 0;

            for (int i = 1; i < 100; i++)
            {
                for (int j = 1; j < 5; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Contains("Sample"))
                    {

                        while (myCode.checkDBNull(ws.Cells[i, j + count_sample].Value).Contains("Sample"))
                        {
                            count_sample++;
                        }
                        lst_spec += count_sample.ToString() + ":";


                        for (int c_offset = 1; c_offset < 6; c_offset++)
                        {
                            if (myCode.checkDBNull(ws.Cells[i, j - c_offset].Value) != "")
                            {
                                lst_spec += myCode.checkDBNull(ws.Cells[i, j - c_offset].Value) + ";";
                                break;
                            }
                        }

                        for (int r_offset = 1; r_offset < 4; r_offset++)
                        {
                            if (myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value) != "")
                            {
                                lst_spec += myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value) + ";" + (i + r_offset).ToString() + ";" + (j - 1).ToString() + "+";

                            }
                        }

                        lst_spec += "_";
                    }
                }
            }
            return lst_spec;
        }


        public string convert_valcell_to_double(object val_cell, int digit)
        {
            string result = "";

            if (myCode.IsNumeric(myCode.checkDBNull(val_cell)))
            {
                result = Math.Round(double.Parse(myCode.checkDBNull(val_cell)), digit).ToString();
            }

            return result;
        }
        public string lst_spec_comment3_onproduct(ExcelWorksheet ws, string type)
        {
            //  AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            List<string> lst_spec = new List<string>();
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(ws, new[] { type.Equals("PSA") ? "PSA peeling" : "Liner peeling", "spec", "TAPE", "Sample" });
            int roRoot = ws.Cells[dic[type.Equals("PSA") ? "PSA peeling": "Liner peeling"].Split('-')[0]].End.Row;
            string[] listZ = new[] { "spec", "TAPE" };
            foreach (string key in listZ)
            {
                List<string> list = new List<string>();
                string[] sp = dic[key].Split('-');
                foreach (string item in sp)
                {
                    int rI = ws.Cells[item].End.Row;
                    if (rI > roRoot)
                    {
                        list.Add(item);
                    }
                }
                dic[key] = string.Join("-", list);
            }
            int countint = dic["Sample"].Split('-').Count() / dic["TAPE"].Split('-').Count();

            foreach (string item in dic["TAPE"].Split('-'))
            {
                int r = ws.Cells[item].End.Row;
                string address = "";
                int mMax = int.MaxValue;
                foreach (string jtem in dic["spec"].Split('-'))
                {
                    int rJ = ws.Cells[jtem].End.Row;
                    if (r < rJ)
                    {
                        if (rJ - r < mMax)
                        {
                            mMax = rJ - r;
                            address = jtem;
                        }
                    }
                }
                try
                {

                    while (true)
                    {
                        address = ExportProcess.AddColumn(address, 1);
                        string value = ws.Cells[address].Value.ToString();
                        string value2 = ws.Cells[ExportProcess.AddRow(address, -1)].Value.ToString();
                        if (string.IsNullOrEmpty(value))
                        {
                            break;
                        }
                        lst_spec.Add($"{ws.Cells[item].Value.ToString().Split('(', ')')[1]}|{value2}:{value}");

                    }
                }
                catch
                {

                }
            }
            return $"{countint}:" + string.Join("-", lst_spec);
        }
        public string lst_spec_comment3_IPQC_coupon(ExcelWorksheet ws)
        {
            string lst_spec = "";
            Dictionary<string, List<string>> dic_spec = new Dictionary<string, List<string>> { };
            int count_sample = 0;

            if (cb_Type.SelectedItem.ToString() == "NPI")
            {
                for (int i = 1; i < 100; i++)
                {
                    for (int j = 1; j < 5; j++)
                    {
                        if (myCode.checkDBNull(ws.Cells[i, j].Value).Contains("Sample"))
                        {
                            while (myCode.checkDBNull(ws.Cells[i, j + count_sample].Value).Contains("Sample"))
                            {
                                count_sample++;
                            }
                            lst_spec += count_sample.ToString() + ":";
                            lst_spec += "A" + "+";

                            for (int c_offset = 1; c_offset < 6; c_offset++)
                            {
                                if (myCode.checkDBNull(ws.Cells[i, j - c_offset].Value) != "")
                                {
                                    lst_spec += myCode.checkDBNull(ws.Cells[i, j - c_offset].Value) + ";";
                                    break;
                                }
                            }
                            for (int r_offset = 1; r_offset < 4; r_offset++)
                            {
                                if (myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value) != "")
                                {
                                    lst_spec += myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value) + ";" + (i + r_offset).ToString() + ";" + (j - 1).ToString() + "+";

                                }
                            }
                            if (myCode.checkDBNull(ws.Cells[i + 4, j - 1].Value) != "")
                            {
                                lst_spec += myCode.checkDBNull(ws.Cells[i + 4, j - 1].Value) + ";" + (i + 3).ToString() + ";" + (j - 1).ToString();
                            }
                            lst_spec += "_";

                            break;
                        }
                        //}
                    }
                }
            }
            else
            {
                for (int i = 1; i < 100; i++)
                {
                    for (int j = 1; j < 5; j++)
                    {
                        if (myCode.checkDBNull(ws.Cells[i, j].Value).Contains("Sample"))
                        {
                            while (myCode.checkDBNull(ws.Cells[i, j + count_sample].Value).Contains("Sample"))
                            {
                                count_sample++;
                            }
                            lst_spec += count_sample.ToString() + ":";
                            lst_spec += "B" + "+";

                            for (int c_offset = 1; c_offset < 6; c_offset++)
                            {
                                if (myCode.checkDBNull(ws.Cells[i, j - c_offset].Value) != "")
                                {
                                    lst_spec += myCode.checkDBNull(ws.Cells[i, j - c_offset].Value) + ";";
                                    break;
                                }
                            }

                            string[] txt_find = new string[4] { "Graph", "Picture", "Max", "Average" };
                            for (int r_offset = 1; r_offset < 6; r_offset++)
                            {
                                string text = myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value);
                                if (text != "")
                                {
                                    if (text.Contains("Graph") || text.Contains("Picture") || text.Contains("Max") || text.Contains("Average"))
                                        lst_spec += myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value) + ";" + (i + r_offset).ToString() + ";" + (j - 1).ToString() + "+";

                                }
                            }
                            string R = "";
                            string UCL = "";
                            string LCL = "";
                            for (int t = 6; t < 12; t++)
                            {
                                string val_cel = myCode.checkDBNull(ws.Cells[i + t, j + 1].Value).Replace(" ", "");
                                if (val_cel.Contains("R(gf)") || val_cel.Contains("R(N)") || val_cel.Replace(" ", "").Contains("R(kgf)"))
                                {
                                    R = convert_valcell_to_double(ws.Cells[i + t, j + 2].Value, 2);
                                }
                                if (val_cel.Contains("UCL"))
                                {
                                    UCL = convert_valcell_to_double(ws.Cells[i + t, j + 2].Value, 2);
                                }
                                if (val_cel.Contains("LCL"))
                                {
                                    LCL = convert_valcell_to_double(ws.Cells[i + t, j + 2].Value, 2);
                                }

                            }
                            lst_spec += R + ";" + UCL + ";" + LCL;
                            lst_spec += "_";

                        }
                    }
                }
            }


            return lst_spec;
        }
        public string lst_spec_comment3_Unmating_pulltest(myExcel.Worksheet ws)
        {
            //  AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            string lst_spec = "";
            Dictionary<string, List<string>> dic_spec = new Dictionary<string, List<string>> { };
            int count_sample = 0;

            for (int i = 1; i < 100; i++)
            {
                for (int j = 1; j < 5; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Contains("Sample"))
                    {

                        while (myCode.checkDBNull(ws.Cells[i, j + count_sample].Value).Contains("Sample"))
                        {
                            count_sample++;
                        }
                        lst_spec += count_sample.ToString() + ":";


                        for (int r_offset = 1; r_offset < 4; r_offset++)
                        {
                            if (myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value) != "")
                            {
                                lst_spec += myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value) + ";" + (i + r_offset).ToString() + ";" + (j - 1).ToString() + "+";

                            }
                        }

                        lst_spec += "_";
                    }
                }
            }
            return lst_spec;
        }

        public string lst_spec_comment3(myExcel.Worksheet ws, string sheet, string[] arr_ignore)
        {
            //  AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            string lst_spec = "";
            Dictionary<string, List<string>> dic_spec = new Dictionary<string, List<string>> { };


            int count_sample = 0;

            for (int i = 1; i < 100; i++)
            {
                for (int j = 1; j < 5; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Contains("Sample"))
                    {
                        while (myCode.checkDBNull(ws.Cells[i, j + count_sample].Value).Contains("Sample"))
                        {
                            count_sample++;
                        }
                        lst_spec += count_sample.ToString() + ":";


                        for (int r_offset = 1; r_offset < 3; r_offset++)
                        {
                            if (myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value) != "")
                            {
                                lst_spec += myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value) + ";" + (i + r_offset).ToString() + ";" + (j - 1).ToString() + "+";
                            }
                        }
                        if (arr_ignore.Contains(sheet.Replace(" ", "").ToUpper()))
                        {
                            if (myCode.checkDBNull(ws.Cells[i + 4, j - 1].Value) != "")
                            {
                                lst_spec += myCode.checkDBNull(ws.Cells[i + 4, j - 1].Value) + ";" + (i + 3).ToString() + ";" + (j - 1).ToString();
                            }
                        }
                        else
                        {
                            if (myCode.checkDBNull(ws.Cells[i + 3, j - 1].Value) != "")
                            {
                                lst_spec += myCode.checkDBNull(ws.Cells[i + 3, j - 1].Value) + ";" + (i + 3).ToString() + ";" + (j - 1).ToString();
                            }
                        }

                        break;
                    }
                }
            }
            return lst_spec;
        }

        public string lst_spec_peeltest(ExcelWorksheet ws)
        {
            #region Old Code
            ////  AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            //string lst_spec = "";
            //Dictionary<string, List<string>> dic_spec = new Dictionary<string, List<string>> { };
            //int count_sample = 0;

            //for (int i = 1; i < 100; i++)
            //{
            //    for (int j = 1; j < 5; j++)
            //    {
            //        if (myCode.checkDBNull(ws.Cells[i, j].Value).Contains("Sample"))
            //        {
            //            while (myCode.checkDBNull(ws.Cells[i, j + count_sample].Value).Contains("Sample"))
            //            {
            //                count_sample++;
            //            }
            //            lst_spec += count_sample.ToString() + ":";

            //            if (myCode.checkDBNull(ws.Cells[i + 1, j - 1].Value).Contains("ID") || myCode.checkDBNull(ws.Cells[i + 1, j - 2].Value).Contains("ID"))
            //            {
            //                lst_spec += "B" + "+";

            //                for (int r_offset = 2; r_offset < 5; r_offset++)
            //                {
            //                    if (myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value) != "")
            //                    {
            //                        lst_spec += myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value).Replace(" ", "") + ";" + (i + r_offset).ToString() + ";" + (j - 1).ToString() + "+";
            //                    }
            //                }

            //                if (!myCode.checkDBNull(ws.Cells[i + 3, j - 1].Value).Contains("≥"))
            //                {
            //                    for (int k = 4; k < 15; k++)
            //                    {
            //                        if (myCode.checkDBNull(ws.Cells[i + k, j - 1].Value).Contains("Judgement"))
            //                        {
            //                            lst_spec += myCode.checkDBNull(ws.Cells[i + k, j - 1].Value).Split('≥')[1].Split('N')[0].Replace(" ", "") + "+";
            //                            break;
            //                        }
            //                    }
            //                }
            //                else
            //                {
            //                    lst_spec += myCode.checkDBNull(ws.Cells[i + 3, j - 1].Value).Split('≥')[1].Split('N')[0].Replace(" ", "") + "+";
            //                }

            //                for (int t = 10; t < 16; t++)
            //                {
            //                    if (myCode.checkDBNull(ws.Cells[i + t, j + 1].Value).Replace(" ", "").ToUpper().Contains("R(N)"))
            //                    {
            //                        string R = Math.Round(double.Parse(myCode.checkDBNull(ws.Cells[i + t, j + 2].Value)), 2).ToString();
            //                        string UCL = Math.Round(double.Parse(myCode.checkDBNull(ws.Cells[i + t + 1, j + 2].Value)), 2).ToString();
            //                        string LCL = Math.Round(double.Parse(myCode.checkDBNull(ws.Cells[i + t + 2, j + 2].Value)), 2).ToString();

            //                        lst_spec += R + ";" + UCL + ";" + LCL;
            //                        break;
            //                    }
            //                }
            //                lst_spec += "_";

            //            }
            //            else
            //            {
            //                lst_spec += "A" + "+";

            //                for (int r_offset = 1; r_offset < 4; r_offset++)
            //                {
            //                    if (myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value) != "")
            //                    {
            //                        lst_spec += myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value).Replace(" ", "") + ";" + (i + r_offset).ToString() + ";" + (j - 1).ToString() + "+";

            //                    }
            //                }
            //                lst_spec += "_";

            //                break;
            //            }
            //        }
            //    }
            //}
            #endregion
            string str = "";
            if (ws.Name.ToUpper().Contains("PEEL"))
            {
                str = "Peeling";
            }
            else
            {
                str = "Pulling";

            }
            string findIndex = $"{str} Force";
            string valueIndex = "";
            string[] sampleList = new[] { "Sample", findIndex };
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(ws, sampleList);
            int count_sample = 0;
            if (dic.TryGetValue("Sample", out string value))
            {
                count_sample = value.Trim().Split('-').Count();
            }

            if (dic.TryGetValue(findIndex, out value))
            {
                valueIndex = ws.Cells[value].Text;
                try
                {
                    valueIndex = valueIndex.Split('(')[1].Replace(" ", "").TrimEnd(new[] { ')' });
                }
                catch
                {
                    valueIndex = "";
                }
                //count_sample = value.Trim().Split('-').Count();
            }


            return $"{count_sample}:{valueIndex}";
        }

        public string lst_spec_unmating(ExcelWorksheet ws)
        {
            //  AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            string lst_spec = "";
            Dictionary<string, List<string>> dic_spec = new Dictionary<string, List<string>> { };
            int count_sample = 0;

            for (int i = 1; i < 100; i++)
            {
                for (int j = 1; j < 5; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Contains("Sample"))
                    {
                        while (myCode.checkDBNull(ws.Cells[i, j + count_sample].Value).Contains("Sample"))
                        {
                            count_sample++;
                        }
                        lst_spec += count_sample.ToString() + ":";

                        lst_spec += "A" + "+";

                        for (int r_offset = 1; r_offset < 5; r_offset++)
                        {
                            if (myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value) != "")
                            {
                                lst_spec += myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value).Replace(" ", "") + ";" + (i + r_offset).ToString() + ";" + (j - 1).ToString() + "+";

                            }
                        }
                        lst_spec += "_";

                        break;

                    }
                }
            }
            return lst_spec;
        }
        public string lst_spec_sheartest(ExcelWorksheet ws)
        {
            //  AutoCompleteStringCollection list = new AutoCompleteStringCollection();

            Dictionary<string, List<string>> dic_spec = new Dictionary<string, List<string>> { };
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(ws, new[] { "Sample", "Shear Force" });
            int count_sample = dic["Sample"].Split('-').Count();

            return $"{count_sample}:{ws.Cells[dic["Shear Force"].Split('-')[1]].Text}";
        }

        public string lst_spec_cross_section_old(myExcel.Worksheet ws)
        {
            /*
             * * Type1: 
              Hottizontal: 20 
              Vertical: 32
             * Ngang: 2  (8)  
               (2 ảnh - 2 dữ liệu) - (2 ảnh - 2 dữ liệu)

             * Trụ ngang: 4 (12)
               (1 ảnh - 2 dữ liệu) - (1 ảnh - 2 dữ liệu) - (1 ảnh - 2 dữ liệu) - (1 ảnh - 2 dữ liệu)

             * Trụ: 2 (16)
               (1 ảnh - 2 dữ liệu - 1 ảnh - 4 dữ liệu )-  (1 ảnh - 2 dữ liệu - 1 ảnh - 4 dữ liệu )

             * Dọc: 2 (16)
               (1 ảnh - 2 dữ liệu - 1 ảnh - 4 dữ liệu )-  (1 ảnh - 2 dữ liệu - 1 ảnh - 4 dữ liệu )
             
              
             
               * * Type2: 
              Hottizontal: 26
              Vertical: 48
             * Ngang: 2  (8)  
               (2 ảnh - 2 dữ liệu) - (2 ảnh - 2 dữ liệu)

             * Trụ ngang 1: 4 (12)
               (1 ảnh - 2 dữ liệu) - (1 ảnh - 2 dữ liệu) - (1 ảnh - 2 dữ liệu) - (1 ảnh - 2 dữ liệu)

            * Trụ ngang 2: 2 (6)
               (1 ảnh - 2 dữ liệu) - (1 ảnh - 2 dữ liệu)  

             * Trụ 1: 2 (16)
               (1 ảnh - 2 dữ liệu - 1 ảnh - 4 dữ liệu )-  (1 ảnh - 2 dữ liệu - 1 ảnh - 4 dữ liệu )

            * Trụ 2: 2 (16)
               (1 ảnh - 2 dữ liệu - 1 ảnh - 4 dữ liệu )-  (1 ảnh - 2 dữ liệu - 1 ảnh - 4 dữ liệu )

             * Dọc: 2 (16)
               (1 ảnh - 2 dữ liệu - 1 ảnh - 4 dữ liệu )-  (1 ảnh - 2 dữ liệu - 1 ảnh - 4 dữ liệu )




               * * Type3: 
              Hottizontal: 32
              Vertical: 48
             * Ngang: 2  (8)  
               (2 ảnh - 2 dữ liệu) - (2 ảnh - 2 dữ liệu)

             * Trụ ngang 1: 4 (12)
               (1 ảnh - 2 dữ liệu) - (1 ảnh - 2 dữ liệu) - (1 ảnh - 2 dữ liệu) - (1 ảnh - 2 dữ liệu)

            * Trụ ngang 2: 4 (12)
                 (1 ảnh - 2 dữ liệu) - (1 ảnh - 2 dữ liệu) - (1 ảnh - 2 dữ liệu) - (1 ảnh - 2 dữ liệu) 

             * Trụ 1: 2 (16)
               (1 ảnh - 2 dữ liệu - 1 ảnh - 4 dữ liệu )-  (1 ảnh - 2 dữ liệu - 1 ảnh - 4 dữ liệu )

            * Trụ 2: 2 (16)
               (1 ảnh - 2 dữ liệu - 1 ảnh - 4 dữ liệu )-  (1 ảnh - 2 dữ liệu - 1 ảnh - 4 dữ liệu )

             * Dọc: 2 (16)
               (1 ảnh - 2 dữ liệu - 1 ảnh - 4 dữ liệu )-  (1 ảnh - 2 dữ liệu - 1 ảnh - 4 dữ liệu )
            
             */


            string lst_spec = "";
            Dictionary<string, List<string>> dic_spec = new Dictionary<string, List<string>> { };


            int count_sample = 0;

            for (int i = 1; i < 100; i++)
            {
                for (int j = 1; j < 8; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Contains("Sample"))
                    {

                        while (myCode.checkDBNull(ws.Cells[i, j + count_sample].Value).Contains("Sample"))
                        {
                            count_sample++;
                        }
                        string s = myCode.checkDBNull(ws.Cells[i + 3, j - 4].Value);
                        lst_spec += count_sample.ToString() + ":" + i.ToString() + ";" + j.ToString() + ";" + s + "+";
                        int col_begin;

                        for (int k = 1; k < 200; k++)
                        {
                            // string val = myCode.checkDBNull(ws.Cells[i + k, j - 5].Value).Replace(" ", "").ToUpper();
                            if (myCode.checkDBNull(ws.Cells[i + k, j - 5].Value).Replace(" ", "").ToUpper().Contains("Connector".ToUpper()))
                            {
                                myExcel.Range sel_rgn = ws.Cells[i + k, j - 5];
                                int a = sel_rgn.MergeArea.Rows.Count;
                                lst_spec += ws.Cells[i + k, j - 5].Value.ToString() + ";" + a.ToString() + ";" + (i + k).ToString() + "_";
                                k = k + a - 1;
                            }
                        }
                        break;

                    }
                }
            }
            return lst_spec;
        }
        public int count_mergcell_crosscut(ExcelRangeBase tar_rgn, ExcelWorksheet tar_wrksht)
        {
            int count = 0;

            int r_count = tar_rgn.End.Row;
            int col_count = tar_rgn.End.Column;
            var idx = tar_wrksht.GetMergeCellId(r_count, col_count);
            string mergedCellAddress = tar_rgn.Address;
            if (idx > 0)
            {
                mergedCellAddress = tar_wrksht.MergedCells[idx - 1];
            }
            count = tar_wrksht.Cells[mergedCellAddress].Rows;
            return count;
        }

        public string lst_spec_cross_section(ExcelWorksheet ws)
        {
            string str1 = "";
            Dictionary<string, List<string>> dictionary = new Dictionary<string, List<string>>();
            int num1 = 0;
            for (int index1 = 1; index1 < 100; ++index1)
            {
                for (int index2 = 1; index2 < 8; ++index2)
                {
                    if (myCode.checkDBNull(((ExcelRangeBase)ws.Cells[index1, index2]).Value).Contains("Sample"))
                    {
                        int num2 = 0;
                        while (myCode.checkDBNull(((ExcelRangeBase)ws.Cells[index1, index2 + num1]).Value).Contains("Sample"))
                            ++num1;
                        if (myCode.checkDBNull(((ExcelRangeBase)ws.Cells[index1 + 1, index2 - 1]).Value).Contains("ID") || myCode.checkDBNull(((ExcelRangeBase)ws.Cells[index1 + 1, index2 - 2]).Value).Contains("ID"))
                        {
                            string str2 = myCode.checkDBNull(((ExcelRangeBase)ws.Cells[index1 + 3, index2 - 1]).Value);
                            string str3 = str1 + num1.ToString() + ":" + index1.ToString() + ";" + index2.ToString() + ";" + str2 + "+";
                            for (int index3 = 1; index3 < 200; ++index3)
                            {
                                if (myCode.checkDBNull(((ExcelRangeBase)ws.Cells[index1 + index3, index2 - 2]).Value).Replace(" ", "").ToUpper().Contains("Connector".ToUpper()))
                                {
                                    int num3 = count_mergcell_crosscut((ExcelRangeBase)ws.Cells[index1 + index3, index2 - 2], ws);
                                    num2 += num3;
                                    str3 = str3 + ((ExcelRangeBase)ws.Cells[index1 + index3, index2 - 2]).Value.ToString() + ";" + num3.ToString() + ";" + (index1 + index3).ToString() + "_";
                                    index3 = index3 + num3 - 1;
                                }
                                if (myCode.checkDBNull(((ExcelRangeBase)ws.Cells[index1 + index3, index2 - 1]).Value).Replace(" ", "").ToUpper().Contains("Result".ToUpper()) || myCode.checkDBNull(((ExcelRangeBase)ws.Cells[index1 + index3, index2 - 2]).Value).Replace(" ", "").ToUpper().Contains("Result".ToUpper()))
                                {
                                    string str4 = "";
                                    for (int index4 = 0; index4 < 15; ++index4)
                                    {
                                        if (myCode.checkDBNull(((ExcelRangeBase)ws.Cells[index1 + index3 + 2 + index4, 3]).Value) != "")
                                            str4 = str4 + myCode.checkDBNull(((ExcelRangeBase)ws.Cells[index1 + index3 + 2 + index4, 3]).Value) + "@";
                                    }
                                    string str5 = str4 + "#";
                                    ExcelRangeBase cell = (ExcelRangeBase)ws.Cells[index1 + index3, index2 - 1];
                                    for (int index5 = 0; index5 < 5; ++index5)
                                    {
                                        for (int index6 = 0; index6 < 5; ++index6)
                                        {
                                            if (myCode.checkDBNull(cell.Offset(index5, index6).Value) == "R")
                                            {
                                                string str6 = str5 + "R: ";
                                                for (int index7 = 1; index7 < 15; ++index7)
                                                {
                                                    if (myCode.IsNumeric(myCode.checkDBNull(cell.Offset(index5 + index7, index6).Value)))
                                                        str6 = str6 + myCode.checkDBNull(cell.Offset(index5 + index7, index6).Value) + ";";
                                                }
                                                str5 = str6 + "^";
                                            }
                                            if (myCode.checkDBNull(cell.Offset(index5, index6).Value) == "UCL")
                                            {
                                                string str7 = str5 + "UCL: ";
                                                for (int index8 = 1; index8 < 15; ++index8)
                                                {
                                                    if (myCode.IsNumeric(myCode.checkDBNull(cell.Offset(index5 + index8, index6).Value)))
                                                        str7 = str7 + myCode.checkDBNull(cell.Offset(index5 + index8, index6).Value) + ";";
                                                }
                                                str5 = str7 + "^";
                                            }
                                            if (myCode.checkDBNull(cell.Offset(index5, index6).Value) == "LCL")
                                            {
                                                string str8 = str5 + "LCL: ";
                                                for (int index9 = 1; index9 < 15; ++index9)
                                                {
                                                    if (myCode.IsNumeric(myCode.checkDBNull(cell.Offset(index5 + index9, index6).Value)))
                                                        str8 = str8 + myCode.checkDBNull(cell.Offset(index5 + index9, index6).Value) + ";";
                                                }
                                                str5 = str8 + "^";
                                            }
                                            if (str5.Contains("R") && str5.Contains("UCL") && str5.Contains("LCL"))
                                                goto label_44;
                                        }
                                    }
                                label_44:
                                    str3 += str5;
                                    break;
                                }
                            }
                            string str9 = "";
                            ExcelRangeBase cell1 = (ExcelRangeBase)ws.Cells[index1 + 1, index2 + num1];
                            for (int index10 = 0; index10 < num2 + 2; ++index10)
                            {
                                if (myCode.checkDBNull(cell1.Offset(index10, 0).Value) != "")
                                    str9 = str9 + myCode.checkDBNull(cell1.Offset(index10, 0).Value) + ";";
                            }
                            str1 = str3 + "_" + str9;
                            break;
                        }
                        string str10 = myCode.checkDBNull(((ExcelRangeBase)ws.Cells[index1 + 3, index2 - 4]).Value);
                        str1 = str1 + num1.ToString() + ":" + index1.ToString() + ";" + index2.ToString() + ";" + str10 + "+";
                        for (int index11 = 1; index11 < 200; ++index11)
                        {
                            if (myCode.checkDBNull(((ExcelRangeBase)ws.Cells[index1 + index11, index2 - 5]).Value).Replace(" ", "").ToUpper().Contains("Connector".ToUpper()))
                            {
                                int num4 = count_mergcell_crosscut((ExcelRangeBase)ws.Cells[index1 + index11, index2 - 5], ws);
                                str1 = str1 + ((ExcelRangeBase)ws.Cells[index1 + index11, index2 - 5]).Value.ToString() + ";" + num4.ToString() + ";" + (index1 + index11).ToString() + "_";
                                index11 = index11 + num4 - 1;
                            }
                        }
                        break;
                    }
                }
            }
            return str1;
        }

        public string lst_spec_cross_section_old(ExcelWorksheet ws)
        {
            string lst_spec = "";
            Dictionary<string, List<string>> dic_spec = new Dictionary<string, List<string>> { };


            int count_sample = 0;

            for (int i = 1; i < 100; i++)
            {
                for (int j = 1; j < 8; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Contains("Sample"))
                    {

                        while (myCode.checkDBNull(ws.Cells[i, j + count_sample].Value).Contains("Sample"))
                        {
                            count_sample++;
                        }

                        if (myCode.checkDBNull(ws.Cells[i + 1, j - 1].Value).Contains("ID") || myCode.checkDBNull(ws.Cells[i + 1, j - 2].Value).Contains("ID"))
                        {
                            string s = myCode.checkDBNull(ws.Cells[i + 3, j - 1].Value);
                            lst_spec += count_sample.ToString() + ":" + i.ToString() + ";" + j.ToString() + ";" + s + "+";
                            for (int k = 1; k < 200; k++)
                            {
                                if (myCode.checkDBNull(ws.Cells[i + k, j - 2].Value).Replace(" ", "").ToUpper().Contains("Connector".ToUpper()))
                                {

                                    ExcelRangeBase sel_rgn = ws.Cells[i + k, j - 2];
                                    int a = count_mergcell_crosscut(sel_rgn, ws);
                                    //int a = sel_rgn.End.Row - (i + k) + 1;
                                    //int a = sel_rgn.MergeArea.Rows.Count;
                                    lst_spec += ws.Cells[i + k, j - 2].Value.ToString() + ";" + a.ToString() + ";" + (i + k).ToString() + "_";
                                    k = k + a - 1;
                                }
                                if (myCode.checkDBNull(ws.Cells[i + k, j - 1].Value).Replace(" ", "").ToUpper().Contains("Result".ToUpper()) || myCode.checkDBNull(ws.Cells[i + k, j - 2].Value).Replace(" ", "").ToUpper().Contains("Result".ToUpper()))
                                {
                                    string spec_mass = "";
                                    ExcelRangeBase rgn_result = ws.Cells[i + k, j - 1];
                                    for (int r = 0; r < 5; r++)
                                    {
                                        for (int c = 0; c < 5; c++)
                                        {
                                            if (myCode.checkDBNull(rgn_result.Offset(r, c).Value) == "R")
                                            {
                                                spec_mass += "R: ";
                                                for (int off_set = 1; off_set < 5; off_set++)
                                                {
                                                    spec_mass += myCode.checkDBNull(rgn_result.Offset(r + off_set, c).Value) + ";";
                                                }
                                                spec_mass += "^";

                                            }
                                            if (myCode.checkDBNull(rgn_result.Offset(r, c).Value) == "UCL")
                                            {
                                                spec_mass += "UCL: ";
                                                for (int off_set = 1; off_set < 5; off_set++)
                                                {
                                                    spec_mass += myCode.checkDBNull(rgn_result.Offset(r + off_set, c).Value) + ";";
                                                }
                                                spec_mass += "^";

                                            }
                                            if (myCode.checkDBNull(rgn_result.Offset(r, c).Value) == "LCL")
                                            {
                                                spec_mass += "LCL: ";
                                                for (int off_set = 1; off_set < 5; off_set++)
                                                {
                                                    spec_mass += myCode.checkDBNull(rgn_result.Offset(r + off_set, c).Value) + ";";
                                                }
                                                spec_mass += "^";

                                            }
                                            if (spec_mass.Contains("R") && spec_mass.Contains("UCL") && spec_mass.Contains("LCL"))
                                                goto lbl_exit;
                                        }
                                    }
                                lbl_exit:
                                    lst_spec += spec_mass;

                                    break;
                                }
                            }






                        }
                        else
                        {
                            string s = myCode.checkDBNull(ws.Cells[i + 3, j - 4].Value);
                            lst_spec += count_sample.ToString() + ":" + i.ToString() + ";" + j.ToString() + ";" + s + "+";
                            int col_begin;

                            for (int k = 1; k < 200; k++)
                            {
                                if (myCode.checkDBNull(ws.Cells[i + k, j - 5].Value).Replace(" ", "").ToUpper().Contains("Connector".ToUpper()))
                                {
                                    ExcelRangeBase sel_rgn = ws.Cells[i + k, j - 5];
                                    //  int a = sel_rgn.MergeArea.Rows.Count;

                                    int a = count_mergcell_crosscut(sel_rgn, ws);

                                    lst_spec += ws.Cells[i + k, j - 5].Value.ToString() + ";" + a.ToString() + ";" + (i + k).ToString() + "_";
                                    k = k + a - 1;
                                }
                            }
                        }
                        break;
                    }
                }
            }
            return lst_spec;
        }
        public string lst_spec_gap_connector(ExcelWorksheet ws)
        {
            string lst_spec = "";
            Dictionary<string, List<string>> dic_spec = new Dictionary<string, List<string>> { };
            List<string> export_possitions = new List<string> { "IOPIN", "LEFT", "RIGHT" };
            string poss = FindPositionsInColumnB(ws, export_possitions);

            int count_sample = 0;

            for (int i = 1; i < 100; i++)
            {
                for (int j = 1; j < 5; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Contains("Sample"))
                    {
                        while (myCode.checkDBNull(ws.Cells[i, j + count_sample].Value).Contains("Sample"))
                        {
                            count_sample++;
                        }
                        lst_spec += count_sample + ":" + j.ToString() + "+" + poss;

                        return lst_spec;
                    }
                }
            }
            return lst_spec;
        }
        public string lst_spec_ACF_old(myExcel.Worksheet ws)
        {
            string lst_spec = "";
            for (int i = 1; i < 40; i++)
            {
                for (int j = 1; j < 15; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Replace(" ", "").ToUpper().Contains("Result data machine".Replace(" ", "").ToUpper()))
                    {
                        string r_begin = i.ToString();

                        for (int k1 = 0; k1 < 5; k1++)
                        {
                            if (myCode.checkDBNull(ws.Cells[i + k1, j + 9].Value).Contains("Sample"))
                            {
                                int count_sample = 0;
                                while (myCode.checkDBNull(ws.Cells[i + k1, j + 9 + count_sample].Value).Contains("Sample"))
                                {
                                    count_sample++;
                                }
                                lst_spec += count_sample + ":" + (i + k1 + 1).ToString() + ";" + (j + 9).ToString() + ";" + r_begin + ";" + myCode.checkDBNull(ws.Cells[i, j].Value) + "_";
                                break;
                            }
                        }
                        break;
                    }

                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Replace(" ", "").ToUpper().Contains("Flatness".Replace(" ", "").ToUpper()))
                    {
                        return lst_spec;
                    }
                }
            }


            return lst_spec;
        }

        public string lst_spec_ACF(ExcelWorksheet ws)
        {
            string lst_spec = "";

            for (int i = 1; i < 40; i++)
            {
                for (int j = 1; j < 15; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Replace(" ", "").ToUpper().Contains("Wetting angle".Replace(" ", "").ToUpper()))
                    {
                        string r_begin = i.ToString();
                        string c_begin = j.ToString();

                        int count_sample = 0;
                        while (myCode.checkDBNull(ws.Cells[i, j + 9 + count_sample].Value).Contains("Sample"))
                        {
                            count_sample++;
                        }

                        lst_spec += count_sample + ":" + r_begin + ";" + c_begin + ";";

                        for (int k = 1; k < 5; k++)
                        {
                            lst_spec += myCode.checkDBNull(ws.Cells[i + k, j + 2].Value) + ";";
                        }

                        lst_spec += "_";
                        break;
                    }
                }
            }

            lst_spec += roughness_spec(ws);
            return lst_spec;


        }

        public string lst_spec_ACF_mass(ExcelWorksheet ws)
        {
            string lst_spec = "";

            for (int i = 1; i < 50; i++)
            {
                for (int j = 1; j < 50; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Contains("Sample"))
                    {
                        int count_sample = 0;
                        while (myCode.checkDBNull(ws.Cells[i, j + count_sample].Value).Contains("Sample"))
                        {
                            count_sample++;
                        }
                        lst_spec += count_sample.ToString() + ":";
                        goto lbl_spec;
                    }
                }
            }

        lbl_spec:
            for (int i = 1; i < 50; i++)
            {
                for (int j = 1; j < 20; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Replace(" ", "").ToUpper().Contains("R(N)"))
                    {
                        string R = "";
                        string UCL = "";
                        string LCL = "";

                        if (myCode.IsNumeric(myCode.checkDBNull(ws.Cells[i, j + 1].Value)))
                        {
                            R = Math.Round(double.Parse(myCode.checkDBNull(ws.Cells[i, j + 1].Value)), 2).ToString();
                        }
                        if (myCode.IsNumeric(myCode.checkDBNull(ws.Cells[i + 1, j + 1].Value)))
                        {
                            UCL = Math.Round(double.Parse(myCode.checkDBNull(ws.Cells[i + 1, j + 1].Value)), 2).ToString();
                        }

                        if (myCode.IsNumeric(myCode.checkDBNull(ws.Cells[i + 2, j + 1].Value)))
                        {
                            LCL = Math.Round(double.Parse(myCode.checkDBNull(ws.Cells[i + 2, j + 1].Value)), 2).ToString();

                        }

                        lst_spec += R + ";" + UCL + ";" + LCL + "^";
                        //return lst_spec;
                    }
                }
                if (lst_spec.Contains("^"))
                    break;
            }

            return lst_spec;
        }
        public string roughness_spec(ExcelWorksheet ws)
        {
            string spec = "";

            ExcelRangeBase curr_rgn_1 = ws.Cells["A1"];

            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 30; i < 100; i++)
            {
                if (myCode.checkDBNull(curr_rgn_1.Offset(i, 0).Value).ToString().Replace(" ", "").ToUpper() == "ACF pad location".Replace(" ", "").ToUpper())
                {
                    int k = 1;

                    while (curr_rgn_1.Offset(i + k, 1).Value.ToString() != "Roughness")
                    {
                        if (curr_rgn_1.Offset(i + k, 1).Value.ToString().Contains("Sa"))
                        {
                            string value = curr_rgn_1.Offset(i + k, 1).Value.ToString();
                            if (!spec.Contains("Sa"))
                            {
                                spec += value + "^";
                            }

                        }
                        if (curr_rgn_1.Offset(i + k, 1).Value.ToString().Contains("Sq"))
                        {
                            //string value = curr_rgn_1.Offset[i + k, 1].Value.ToString().Replace(" ", "").Replace("Sq(", "").Replace(")", "").Replace("um", "");
                            string value = curr_rgn_1.Offset(i + k, 1).Value.ToString();
                            if (!spec.Contains("Sq"))
                            {
                                spec += value + "^";
                            }

                        }
                        if (curr_rgn_1.Offset(i + k, 1).Value.ToString().Contains("Sdr"))
                        {

                            //string value = curr_rgn_1.Offset[i + k, 1].Value.ToString().Replace(" ", "").Replace("Sdr(", "").Replace(")", "").Replace("um", "");
                            string value = curr_rgn_1.Offset(i + k, 1).Value.ToString();
                            if (!spec.Contains("Sdr"))
                                spec += value;
                        }
                        k++;
                    }
                    break;
                }
            }


            return spec;
        }
        public Dictionary<string, List<string>> setup_spec_ACF(myExcel.Worksheet ws)
        {
            //  AutoCompleteStringCollection list = new AutoCompleteStringCollection();

            Dictionary<string, List<string>> dic_spec = new Dictionary<string, List<string>> { };

            int count_sample = 0;


            for (int i = 1; i < 100; i++)
            {
                for (int j = 1; j < 5; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Contains("Wetting angle"))
                    {
                        List<string> lst_spec = new List<string>();
                        while (myCode.checkDBNull(ws.Cells[i, j + 7 + count_sample].Value).Contains("Sample"))
                        {
                            count_sample++;
                        }
                        lst_spec.Add(count_sample.ToString());

                        int r_offset = 1;
                        while (!myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value).Contain("Judgement"))
                        {
                            if (myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value) != "")
                            {
                                lst_spec.Add(myCode.checkDBNull(ws.Cells[i + r_offset, j - 1].Value) + ";" + (i + r_offset).ToString() + ";" + (j - 1).ToString());
                            }
                            r_offset++;

                        }

                        dic_spec.Add(myCode.checkDBNull(ws.Cells[i + 1, j - 2].Value), lst_spec);

                        break;
                    }
                }


            }
            return dic_spec;


        }
        public string FindPositionsInColumnB(ExcelWorksheet ws, List<string> valuesToSearch)
        {

            //myExcel.Range range = ws.UsedRange;
            int rowCount = 150;
            int r_begin_region = 1;

            //List<string> positions = new List<string>();

            string position = "";


            foreach (string value in valuesToSearch)
            {
                int count_lk = 0;
            lbl_findnext:
                for (int row = r_begin_region; row <= rowCount; row++)
                {
                    //string cellValue = Convert.ToString((range.Cells[row, 1] as myExcel.Range).Value2);
                    for (int col = 1; col < 5; col++)
                    {
                        string cell_val = myCode.checkDBNull(ws.Cells[row, col].Value).Replace(" ", "").ToUpper();
                        if (cell_val.Contains(value))
                        {
                            bool check_contain = true;
                            if (value == "IOPIN" && (cell_val.Contains("RIGHT") || cell_val.Contains("LEFT")))
                            {
                                check_contain = false;
                            }
                            if (value == "RIGHT" && !cell_val.Contains("PIN"))
                            {
                                check_contain = false;
                            }

                            if (value == "LEFT" && !cell_val.Contains("PIN"))
                            {
                                check_contain = false;
                            }

                            if (check_contain)
                            {

                                for (int i = row; i < row + 10; i++)
                                {

                                    for (int col2 = col; col2 < 10; col2++)
                                    {
                                        if (myCode.checkDBNull(ws.Cells[i, col2].Value).Replace(" ", "").Contains("Sample"))
                                        {
                                            string spec_b1 = myCode.checkDBNull(ws.Cells[i + 4, col2 - 1].Value);
                                            int count_row = 0;

                                            while (myCode.checkDBNull(ws.Cells[i + count_row + 3, col2 - 1].Value) != "")
                                            {
                                                count_row++;

                                            }
                                            position += value + ";" + (i + 1).ToString() + ";" + count_row.ToString() + ";" + spec_b1 + "^";
                                            count_lk++;
                                            r_begin_region = i + count_row * 2 + 2;
                                            goto lbl_findnext;


                                        }
                                    }


                                }

                            }


                        }
                    }
                    if (count_lk == 2)
                        break;

                }

                position += "_";

            }
            return position;
        }


        public List<string> setup_spec_flexpending(myExcel.Worksheet ws)
        {
            //  AutoCompleteStringCollection list = new AutoCompleteStringCollection();

            List<string> result = new List<string>();
            int count_sample = 0;
            string location = "";

            for (int j = 1; j < 8; j++)
            {
                for (int i = 1; i < 200; i++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value) == "Sample")
                    {

                        while (myCode.checkDBNull(ws.Cells[i, j + count_sample * 3].Value).Contains("Sample"))
                        {
                            count_sample++;
                            location = i.ToString() + ";" + j.ToString();
                        }
                        break;
                    }

                }
            }
            result.Add(count_sample.ToString());
            result.Add(location);

            return result;
        }
        public string find_format_mass(string in_data_loc, string ItemCode)
        {

            string result = "";
            if (System.IO.Directory.Exists(in_data_loc))
            {
                string[] file_xlsm = Directory.GetFiles(in_data_loc, "*" + ItemCode + "*.xlsm");

                if (file_xlsm.Length > 0)
                {
                    result = file_xlsm[0];
                }
                else
                {
                    string[] file_xlsx = Directory.GetFiles(in_data_loc, "*" + ItemCode + "*.xlsx");
                    if (file_xlsx.Length > 0)
                    {
                        result = file_xlsx[0];

                    }
                    else
                    {
                        string[] file_xls = Directory.GetFiles(in_data_loc, "*" + ItemCode + "*.xls");
                        if (file_xls.Length > 0)
                        {
                            result = file_xls[0];

                        }
                    }

                }
            }
            result = result.Replace("~$", "");
            return result;
        }

        public void setup_format_commet3(string file_format, string ItemCode)
        {
            ExcelWorkbook wb = TDMK_EPPLUS.open_excel_file(file_format);
            string lst_sheet_notfound = "";

            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, cb_Type.SelectedItem.ToString() });
            DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", filter_str).Clone();
            int ID = TDMK_Code.SQL_MAX("SPEC_COMMENT_3", "ID", sqlcon) + 1;
            foreach (string sheet in cbl_sheet.CheckedItems)
            {
                string mySheet = "";

                if (sheet == "FAI")
                {
                    DataTable cur_dt = Load_Spec_fromFile(sqlcon, wb, ItemCode, new List<string> { "*.xlsx", "*.xlsm" }, cb_Type.SelectedItem.ToString());
                    Save_FAI_Spec(cur_dt, ItemCode, sqlcon, cb_Type.SelectedItem.ToString());
                }
                else
                {
                    foreach (ExcelWorksheet tg_sht in wb.Worksheets)
                    {
                        string cur_sht_name = tg_sht.Name.ToUpper().Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "");
                        if (cur_sht_name == sheet.ToUpper().Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", ""))
                        {
                            mySheet = tg_sht.Name;
                            break;
                        }
                    }
                    if (mySheet != "")
                    {
                        ExcelWorksheet ws = wb.Worksheets[mySheet];

                        string lst_spec = "";
                        switch (sheet.Replace(" ", "_").ToUpper())
                        {
                            case "PEEL_TEST":
                                lst_spec = lst_spec_peeltest(ws);
                                break;

                            case "(MATING)_PULL_TEST":
                                lst_spec = lst_spec_peeltest(ws);
                                break;
                            case "(IQC_UNMATING)_PULL_TEST":
                                lst_spec = lst_spec_unmating(ws);
                                break;

                            case "SHEAR_TEST":
                                lst_spec = lst_spec_sheartest(ws);
                                break;

                            case "IQC_LINER_PEELING_(COUPON)":
                                lst_spec = lst_spec_comment3_IPQC_coupon(ws);
                                break;

                            case "IQC_PSA_PEELING_(COUPON)":
                                lst_spec = lst_spec_comment3_IPQC_coupon(ws);
                                break;

                            case "LINER_PEEL_TEST_(ON_PRODUCT)":
                                lst_spec = lst_spec_comment3_onproduct(ws, "LINER");
                                break;

                            case "PSA_PEEL_TEST_(ON_PRODUCT)":
                                lst_spec = lst_spec_comment3_onproduct(ws, "PSA");
                                break;

                            case "CROSS_SECTION":
                                lst_spec = lst_spec_cross_section(ws);
                                break;

                            case "GAP_CONNECTOR":
                                lst_spec = lst_spec_gap_connector(ws);
                                break;

                            case "ACF":
                                lst_spec = lst_spec_ACF(ws);
                                break;
                            case "SEM_BSE_&_BINARIZATION":
                                lst_spec = SEMServices.setup_spec(ws);
                                break;
                            case "OQC_B2B_MATING-UNMATING":
                                lst_spec = OQCB2BMatingUnmatting.setup_spec(ws);
                                break;
                            case "THERMAL_CYCLING":
                            case "THERMAL_SHOCK":
                            case "HEAT_SOAK_AND_RECOVERY":
                                lst_spec = TCHSTSService.setup_spec(ws);
                                break;
                            case "ENVIRONMENT_EN-DURANCE":
                                lst_spec = EEDService.SetupSpec(ws);
                                break;
                            default:
                                //string str = sheet.Replace(" ", "_").ToUpper();
                                //Debugger.Break();
                                break;
                        }

                        string[] arr_spec = lst_spec.Split(':');
                        DataRow dr = dt_spec.NewRow();

                        dr[0] = ID;
                        dr[1] = ItemCode;
                        dr[2] = sheet.Replace(" ", "_").Replace("(", "").Replace(")", "").ToUpper();
                        dr[3] = arr_spec[0];
                        dr[4] = lst_spec.Replace(arr_spec[0] + ":", string.Empty);
                        dr[5] = cb_Type.SelectedItem.ToString();

                        dt_spec.Rows.Add(dr);
                        ID++;
                        // BatchBulkCopy(sqlcon, dt_spec, "SPEC_COMMENT_3");
                    }
                    else
                    {
                        lst_sheet_notfound += sheet + " ; ";
                    }
                }
            }

            //wb.Close();
            string msg = "";
            if (lst_sheet_notfound == "")
            {
                msg = "Cài đặt thành công";
                filter_ItemCode(ItemCode);

            }
            else
            {
                msg = "Không tìm thấy sheet " + lst_sheet_notfound.Remove(lst_sheet_notfound.Length - 2, 2);
            }
            MessageBox.Show(new Form { TopMost = true }, msg, "Thông báo");
            BatchBulkCopy(sqlcon, dt_spec, "SPEC_COMMENT_3");
            //fill_dgv_setup_detail();

        }
        public void setup_format_mass(string file_format, string ItemCode)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, cb_Type.SelectedItem.ToString() });
            DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", filter_str).Clone();
            int ID = TDMK_Code.SQL_MAX("SPEC_COMMENT_3", "ID", sqlcon) + 1;
            string sheets_notfound = "";
            foreach (string sheet in cbl_sheet.CheckedItems)
            {
                string format_path = find_format_mass(Path.Combine(file_format, sheet.ToString().Replace(" ", "_").Replace("(", "").Replace(")", "").ToUpper()), txtItemCode.Text);
                if (sheet == "FAI")
                {
                    if (format_path != "")
                    {
                        ExcelWorkbook wb = TDMK_EPPLUS.open_excel_file(format_path);
                        DataTable cur_dt = Load_Spec_fromFile(sqlcon, wb, ItemCode, new List<string> { "*.xlsx", "*.xlsm" }, cb_Type.SelectedItem.ToString());
                        Save_FAI_Spec(cur_dt, ItemCode, sqlcon, cb_Type.SelectedItem.ToString());
                    }
                    else
                    {
                        sheets_notfound += sheet + ";";
                    }
                }
                else
                {

                    if (format_path != "")
                    {

                        ExcelWorkbook wb = TDMK_EPPLUS.open_excel_file(format_path);

                        string[] arr_diff = { "SHEARTEST", "IQC Liner peeling (Coupon)".Replace(" ", "").ToUpper(), "IQC PSA peeling (Coupon)".Replace(" ", "").ToUpper() };
                        string[] arr_onproduct = { "Liner peel test (On product)".Replace(" ", "").ToUpper(), "PSA peel test (On product)".Replace(" ", "").ToUpper() };
                        string[] arr_comment_2 = { "Cross section".Replace(" ", "").ToUpper(), "GAP Connector".Replace(" ", "").ToUpper() };

                        if (wb.Worksheets.Count > 0)
                        {
                            ExcelWorksheet ws = wb.Worksheets[0];
                            string lst_spec = "";

                            switch (sheet.Replace(" ", "_").ToUpper())
                            {
                                case "PEEL_TEST":
                                    lst_spec = lst_spec_peeltest(ws);
                                    break;

                                case "(MATING)_PULL_TEST":
                                    lst_spec = lst_spec_peeltest(ws);
                                    break;
                                case "(IQC_UNMATING)_PULL_TEST":
                                    lst_spec = lst_spec_unmating(ws);
                                    break;

                                case "SHEAR_TEST":
                                    lst_spec = lst_spec_sheartest(ws);
                                    break;

                                case "IQC_LINER_PEELING_(COUPON)":
                                    lst_spec = lst_spec_comment3_IPQC_coupon(ws);
                                    break;

                                case "IQC_PSA_PEELING_(COUPON)":
                                    lst_spec = lst_spec_comment3_IPQC_coupon(ws);
                                    break;

                                case "LINER_PEEL_TEST_(ON_PRODUCT)":
                                    lst_spec = lst_spec_comment3_onproduct(ws, "LINER");
                                    break;

                                case "PSA_PEEL_TEST_(ON_PRODUCT)":
                                    lst_spec = lst_spec_comment3_onproduct(ws, "PSA");
                                    break;

                                case "CROSS_SECTION":
                                    lst_spec = lst_spec_cross_section(ws);
                                    break;

                                case "GAP_CONNECTOR":
                                    lst_spec = lst_spec_gap_connector(ws);
                                    break;

                                case "ACF":

                                    break;

                            }

                            string[] arr_spec = lst_spec.Split(':');
                            DataRow dr = dt_spec.NewRow();

                            dr[0] = ID;
                            dr[1] = ItemCode;
                            dr[2] = sheet.Replace(" ", "_").Replace("(", "").Replace(")", "").ToUpper();
                            dr[3] = arr_spec[0];
                            dr[4] = lst_spec.Replace(arr_spec[0] + ":", string.Empty);
                            dr[5] = cb_Type.SelectedItem.ToString();

                            dt_spec.Rows.Add(dr);
                            ID++;

                        }

                    }

                    else if (sheet == "ACF" && cb_Type.SelectedItem.ToString() == "MASS")
                    {
                        string format_path_ACF_mass = find_format_mass(Path.Combine(file_format, "ACF", "ACF_BONDING"), txtItemCode.Text);
                        if (format_path_ACF_mass != "")
                        {
                            ExcelWorkbook wb = TDMK_EPPLUS.open_excel_file(format_path_ACF_mass);
                            ExcelWorksheet ws = wb.Worksheets[0];
                            DataRow dr = dt_spec.NewRow();
                            string lst_spec = lst_spec_ACF_mass(ws);

                            dr[0] = ID;
                            dr[1] = ItemCode;
                            dr[2] = sheet.Replace(" ", "_").Replace("(", "").Replace(")", "").ToUpper();
                            dr[3] = lst_spec.Split(':')[0];
                            dr[4] = lst_spec.Split(':')[1];
                            dr[5] = cb_Type.SelectedItem.ToString();

                            dt_spec.Rows.Add(dr);
                        }
                        else
                        {
                            sheets_notfound += sheet + "_BONDING" + ";";
                        }

                    }
                    else
                    {
                        sheets_notfound += sheet + ";";
                    }
                }
            }
            BatchBulkCopy(sqlcon, dt_spec, "SPEC_COMMENT_3");

            sheets_notfound = sheets_notfound.TrimEnd(';');
            string msg = "";

            if (sheets_notfound == "")
            {
                msg = "Cài đặt thành công!";
            }
            else if (sheets_notfound.Split(';').Length == cbl_sheet.CheckedItems.Count)
            {
                msg += "Không tìm thấy format của " + sheets_notfound;
            }
            else
            {
                msg = "Cài đặt thành công!\n\n Không tìm thấy format của " + sheets_notfound;
            }

            MessageBox.Show(new Form { TopMost = true }, msg, "Thông báo");
            filter_ItemCode(ItemCode);
            // fill_dgv_setup_detail();

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog f_open = new OpenFileDialog();
            //f_open.Filter = "Excel(*.xlsx)|*.xlsm";
            f_open.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            f_open.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            if (f_open.ShowDialog() == DialogResult.OK)
            {
                if (f_open.FileName != "")
                {
                    txtFormat.Text = f_open.FileName;
                }

            }
            //    FolderBrowserDialog myfolder = new FolderBrowserDialog();
            //    if (myfolder.ShowDialog() == DialogResult.OK)
            //    {
            //        txtFormat.Text = myfolder.SelectedPath;
            //    }
        }
        public string find_format(string in_data_loc, string ItemCode)
        {
            //string result = "";
            //string[] files = Directory.GetFiles(Path.Combine(in_data_loc, "Format"), "*" + ItemCode + "*.xlsm");
            //if (files.Length > 0)
            //{
            //    result = files[0];
            //}
            //return result;
            string result = "";
            string[] file_xlsm = Directory.GetFiles(Path.Combine(in_data_loc, "Format"), "*" + ItemCode + "*.xlsm");
            if (file_xlsm.Length > 0)
            {
                result = file_xlsm[0];
            }
            else
            {
                string[] file_xlsx = Directory.GetFiles(Path.Combine(in_data_loc, "Format"), "*" + ItemCode + "*.xlsx");
                if (file_xlsx.Length > 0)
                {
                    result = file_xlsx[0];

                }
            }
            return result;
        }

        private void btnRUN_Click(object sender, EventArgs e)
        {
            if (txtFormat.Text != "" && txtItemCode.Text != "" && cb_Type.SelectedIndex != -1)
            {
                if (cbl_sheet.CheckedItems.Count == 0)
                {
                    MessageBox.Show(new Form { TopMost = true }, "Chưa có sheet nào được chọn", "Thông báo");
                }
                else
                {
                    if (cb_Type.SelectedItem.ToString() == "NPI")
                    {
                        string format_fpath = "";
                        FileAttributes attr = File.GetAttributes(txtFormat.Text);
                        if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
                        {
                            string[] file_format_lst_xlsm = Directory.GetFiles(txtFormat.Text, "*" + ".xlsm").Where(s => s.Contains(txtItemCode.Text)).ToArray();
                            if (file_format_lst_xlsm.Length > 0)
                            {
                                format_fpath = file_format_lst_xlsm[0];
                            }
                            else
                            {
                                string[] file_format_lst_xlsx = Directory.GetFiles(txtFormat.Text, "*" + ".xlsx").Where(s => s.Contains(txtItemCode.Text)).ToArray();
                                if (file_format_lst_xlsx.Length > 0)
                                {
                                    string sheet = cbl_sheet.CheckedItems[0].ToString().ToUpper().Replace(" ", "_").Replace("(", "").Replace(")", "");
                                    format_fpath = file_format_lst_xlsx[0];
                                    foreach (string f in file_format_lst_xlsx)
                                    {
                                        if (f.ToUpper().Replace(" ", "_").Replace("(", "").Replace(")", "").Contains(sheet))
                                        {
                                            format_fpath = f;
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            format_fpath = txtFormat.Text;
                        }

                        format_fpath = format_fpath.Replace("~$", "");

                        if (format_fpath != "" && (format_fpath.Contains(".xlsx") || format_fpath.Contains(".xlsm")))
                        {
                            if (lst_Item.Items.IndexOf(txtItemCode.Text) == -1)
                            {
                                setup_format_commet3(format_fpath, txtItemCode.Text);
                                lst_Item.Items.Clear();
                                DataTable dt_item_setup = Datatable_ALL(sqlcon, "SPEC_COMMENT_3");
                                string[] arr_items = dt_item_setup.AsEnumerable().Select(x => x.Field<string>("ItemCode")).Distinct().ToArray();
                                foreach (string item in arr_items)
                                {
                                    lst_Item.Items.Add(item);
                                }
                            }
                            else
                            {
                                string arr_sheet_setup = "";
                                foreach (string sheet in cbl_sheet.CheckedItems)
                                {
                                    DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, sheet.Replace(" ", "_").Replace("(", "").Replace(")", ""), cb_Type.SelectedItem.ToString() }));
                                    if (dt_spec.Rows.Count > 0)
                                    {
                                        arr_sheet_setup += sheet + ",";
                                    }
                                }

                                string msg = "";
                                if (arr_sheet_setup != "")
                                {
                                    if (arr_sheet_setup.Split(',').Length == cbl_sheet.Items.Count + 1)
                                    {
                                        msg = "Spec của tất cả các sheet đều đã được cài đặt. Bạn có muốn cập nhật không?";
                                    }
                                    else
                                    {
                                        arr_sheet_setup = arr_sheet_setup.Remove(arr_sheet_setup.LastIndexOf(','));
                                        msg = "Spec của " + arr_sheet_setup + " đã được cài đặt. Bạn có muốn cập nhật không?";
                                    }

                                    if (MessageBox.Show(new Form { TopMost = true }, msg, "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                                    {
                                        foreach (string sheet in cbl_sheet.CheckedItems)
                                        {
                                            TDMK_Code.Delelte_FilteredItem_arr("SPEC_COMMENT_3", sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, sheet.Replace(" ", "_").Replace("(", "").Replace(")", "").ToUpper(), cb_Type.SelectedItem.ToString() }));
                                        }

                                        setup_format_commet3(format_fpath, txtItemCode.Text);
                                        lst_Item.Items.Clear();
                                        DataTable dt_item_setup = Datatable_ALL(sqlcon, "SPEC_COMMENT_3");
                                        string[] arr_items = dt_item_setup.AsEnumerable().Select(x => x.Field<string>("ItemCode")).Distinct().ToArray();
                                        foreach (string item in arr_items)
                                        {
                                            lst_Item.Items.Add(item);
                                        }
                                    }
                                }
                                else
                                {
                                    setup_format_commet3(format_fpath, txtItemCode.Text);
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Không tìm thấy Format của " + txtItemCode.Text, "Thông báo");
                        }

                    }
                    else if (cb_Type.SelectedItem.ToString() == "MASS" /*|| cb_Type.SelectedItem.ToString() == "Other"*/)
                    {
                        string format_fpath = "";
                        FileAttributes attr = File.GetAttributes(txtFormat.Text);
                        if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
                        {
                            format_fpath = txtFormat.Text;
                            if (format_fpath != "")
                            {
                                if (lst_Item.Items.IndexOf(txtItemCode.Text) == -1)
                                {
                                    setup_format_mass(format_fpath, txtItemCode.Text);
                                    lst_Item.Items.Clear();
                                    DataTable dt_item_setup = Datatable_ALL(sqlcon, "SPEC_COMMENT_3");
                                    string[] arr_items = dt_item_setup.AsEnumerable().Select(x => x.Field<string>("ItemCode")).Distinct().ToArray();
                                    foreach (string item in arr_items)
                                    {
                                        lst_Item.Items.Add(item);
                                    }
                                }
                                else
                                {
                                    string arr_sheet_setup = "";
                                    foreach (string sheet in cbl_sheet.CheckedItems)
                                    {
                                        DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, sheet.Replace(" ", "_").Replace("(", "").Replace(")", ""), cb_Type.SelectedItem.ToString() }));
                                        if (dt_spec.Rows.Count > 0)
                                        {
                                            arr_sheet_setup += sheet + ",";
                                        }
                                    }
                                    string msg = "";
                                    if (arr_sheet_setup != "")
                                    {
                                        if (arr_sheet_setup.Split(',').Length == cbl_sheet.Items.Count)
                                        {
                                            msg = "Spec của tất cả các sheet đều đã được cài đặt. Bạn có muốn cập nhật không?";
                                        }
                                        else
                                        {
                                            msg = "Spec của " + arr_sheet_setup + " đã được cài đặt. Bạn có muốn cập nhật không?";
                                        }

                                        arr_sheet_setup = arr_sheet_setup.Remove(arr_sheet_setup.LastIndexOf(','));
                                        if (MessageBox.Show(new Form { TopMost = true }, msg, "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                                        {
                                            foreach (string sheet in cbl_sheet.CheckedItems)
                                            {
                                                TDMK_Code.Delelte_FilteredItem_arr("SPEC_COMMENT_3", sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, sheet.Replace(" ", "_").Replace("(", "").Replace(")", "").ToUpper(), cb_Type.SelectedItem.ToString() }));
                                            }

                                            //DataTable dt_1 = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text })); 
                                            setup_format_mass(format_fpath, txtItemCode.Text);
                                            lst_Item.Items.Clear();
                                            DataTable dt_item_setup = Datatable_ALL(sqlcon, "SPEC_COMMENT_3");
                                            string[] arr_items = dt_item_setup.AsEnumerable().Select(x => x.Field<string>("ItemCode")).Distinct().ToArray();
                                            foreach (string item in arr_items)
                                            {
                                                lst_Item.Items.Add(item);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        setup_format_mass(format_fpath, txtItemCode.Text);
                                    }

                                }
                            }
                        }
                        //else
                        //{
                        //    format_fpath = txtFormat.Text;
                        //}

                    }

                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Vui lòng điền đầy đủ thông tin!", "Thông báo");
            }

        }

        private void cbl_sheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (cb_Type.SelectedItem.ToString() != "NPI")
            //{
            //    txtFormat.Text = Path.Combine(data_loc, "Format", cb_Type.SelectedItem.ToString(), cbl_sheet.CheckedItems[0].ToString().Replace(" ", "_").Replace("(", "").Replace(")", "").ToUpper());
            //}

            //if (cbl_sheet.CheckedItems.Contains("ALL"))
            //{
            //    for (int i = 0; i < cbl_sheet.Items.Count; i++)
            //    {
            //        cbl_sheet.SetItemChecked(i, true);
            //    }
            //}
            //else
            //{
            //    // cbl_sheet.SetItemChecked(0, false);
            //    for (int i = 0; i < cbl_sheet.Items.Count; i++)
            //    {
            //        cbl_sheet.SetItemChecked(i, true);
            //    }
            //    for (int i = 1; i < cbl_sheet.Items.Count; i++)
            //    {
            //        if (cbl_sheet.CheckedItems.Contains(cbl_sheet.Items[i]))
            //        {
            //            cbl_sheet.SetItemChecked(i, true);
            //        }
            //    }
            //}
        }



        private void cbl_sheet_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            //for (int i = 0; i < cbl_sheet.Items.Count; i++)
            //{
            //    if (i != e.Index)
            //    {
            //        cbl_sheet.SetItemChecked(i, false);
            //    }
            //} 
        }


        private void cbAll_CheckedChanged(object sender, EventArgs e)
        {

            if (cbAll.Checked == true)
            {
                for (int i = 0; i < cbl_sheet.Items.Count; i++)
                {
                    cbl_sheet.SetItemChecked(i, true);
                }
            }
            else
            {
                for (int i = 0; i < cbl_sheet.Items.Count; i++)
                {
                    cbl_sheet.SetItemChecked(i, false);
                }
            }


        }

        private void Setup_Spec_SMT_FormClosed(object sender, FormClosedEventArgs e)
        {
            //MainScreen frm = new MainScreen();
            //frm.Show();
        }

        private void txtItemCode_Validated(object sender, EventArgs e)
        {
            txtItemCode.Text = txtItemCode.Text.Replace(" ", "").ToUpper();

        }

        private void cb_Type_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_Type.SelectedIndex != -1)
            {
                txtFormat.Text = Path.Combine(txtFormat.Text, "Format", cb_Type.SelectedItem.ToString());

                if (cb_Type.SelectedItem.ToString() == "MASS" || cb_Type.SelectedItem.ToString() == "Other")
                {
                    txtFormat.Enabled = false;
                }
                else
                {
                    txtFormat.Enabled = true;
                }


                //if (cb_Type.SelectedItem.ToString() == "Other")
                //{
                //    cbAll.Checked = false;
                //    cbAll.Enabled = false;
                //    txtFormat.Text = "";

                //}
                //else
                //{
                //    cbAll.Checked = true;
                //    cbAll.Enabled = true;
                //    txtFormat.Text = Path.Combine(data_loc, "Format");
                //}
            }
            txtFormat.Enabled = true;
        }

        private void btn_filter_Click(object sender, EventArgs e)
        {
            if (txt_ItemCode_filter.Text != "")
            {
                filter_ItemCode(txt_ItemCode_filter.Text);
            }
        }

        private void btn_all_Click(object sender, EventArgs e)
        {
            fill_dgv_setup_detail();
        }

        private void txt_ItemCode_filter_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btn_filter.PerformClick();
            }
        }

        private void txt_ItemCode_filter_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtItemCode_TextChanged(object sender, EventArgs e)
        {

        }

        private void txt_ItemCode_filter_Validated(object sender, EventArgs e)
        {

        }

        public bool check_columns_existed(DataTable src_tbl, string find_col_name)
        {
            bool result = false;
            foreach (DataColumn column in src_tbl.Columns)
            {
                string columnName = column.ColumnName;
                if (columnName == find_col_name)
                {
                    result = true;
                    break;
                }
            }

            return result;
        }

        public string Find_Cell_Addr(string search_key, string start_addr, ExcelWorksheet tar_wrksht, bool left_to_right)
        {
            string result = "";
            // Range range = ((_Worksheet)tar_wrksht).get_Range((object)start_addr, Type.Missing);
            ExcelRangeBase range = tar_wrksht.Cells[start_addr];
            for (int i = 0; i < 100; i++)
            {
                //Range range2 = range.get_Offset((object)i, (object)0);
                ExcelRangeBase range2 = range.Offset(i, 0);
                if (left_to_right)
                {
                    //range2 = range.get_Offset((object)0, (object)i);
                    range2 = range.Offset(0, i);
                }

                // string text = myCode.checkDBNull((dynamic)range2.get_Value(Type.Missing));
                string text = myCode.checkDBNull(range2.Value);
                if (text.Replace(" ", "").ToUpper() == search_key.Replace(" ", "").ToUpper())
                {
                    //result = range2.get_AddressLocal(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    result = range2.Address;
                    break;
                }
            }

            return result;
        }
        public string Find_Offset(string start_addr, ExcelWorksheet tar_wrksht, bool left_to_right, string search_key = "")
        {
            string result = "";
            //Range range = ((_Worksheet)tar_wrksht).get_Range((object)start_addr, Type.Missing);
            ExcelRangeBase range = tar_wrksht.Cells[start_addr];
            int num = 0;
            int num2 = 0;
            while (true)
            {
                // Range range2 = range.get_Offset((object)num, (object)0);
                ExcelRangeBase range2 = range.Offset(num, 0);
                if (left_to_right)
                {
                    //range2 = range.get_Offset((object)0, (object)num);
                    range2 = range.Offset(0, num);
                }

                //string text = myCode.checkDBNull((dynamic)range2.get_Value(Type.Missing));
                string text = myCode.checkDBNull(range2.Value);
                if (text == "")
                {
                    num2++;
                    if (num2 > 3)
                    {
                        break;
                    }
                }
                else
                {
                    num2 = 0;
                }

                if (search_key != "")
                {
                    if (text.Replace(" ", "").ToUpper() == search_key.Replace(" ", "").ToUpper())
                    {
                        //result = range2.get_AddressLocal(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                        result = range2.Address;
                        break;
                    }
                }
                else if (num2 == 0)
                {
                    //result = range2.get_AddressLocal(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    result = range2.Address;
                    break;
                }

                num++;
            }

            return result;
        }
        public void Get_ListTable(int col_inx, DataTable myDt, string[] src_arr, ref List<DataTable> src_lst_tbl, string tar_item)
        {
            if (col_inx < src_arr.Length - 1)
            {
                if (!(src_arr[col_inx + 1] != ""))
                {
                    return;
                }

                string[] array = (from x in myDt.AsEnumerable()
                                  select x.Field<string>(src_arr[col_inx + 1])).Distinct().ToArray();
                if (array.Length != 0)
                {
                    string[] array2 = array;
                    foreach (string sv in array2)
                    {
                        if (sv != null)
                        {
                            DataTable myDt2 = (from r in myDt.AsEnumerable()
                                               where r.Field<string>(src_arr[col_inx + 1]) == sv
                                               select r).CopyToDataTable();
                            Get_ListTable(col_inx + 1, myDt2, src_arr, ref src_lst_tbl, tar_item);
                        }
                    }
                }
                else
                {
                    DataRow[] array3 = (from x in myDt.AsEnumerable()
                                        where x.Field<string>(tar_item) != null
                                        select x).ToArray();
                    if (array3.Length != 0)
                    {
                        src_lst_tbl.Add(myDt);
                    }
                }
            }
            else
            {
                DataRow[] array3 = (from x in myDt.AsEnumerable()
                                    where x.Field<string>(tar_item) != null
                                    select x).ToArray();
                if (array3.Length != 0)
                {
                    src_lst_tbl.Add(myDt);
                }
            }
        }


        public DataTable Load_FAI_Spec_ToTable(SqlConnection sqlcon, string tar_ItemCode, string format_type)
        {
            DataTable dataTable = new DataTable();
            DataTable dataTable2 = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Spec", TDMK_Code.filter_str(new string[2] { "ItemCode", "Remark" }, new string[2] { tar_ItemCode, format_type }));
            List<string> search_key_lst = new List<string> { "FAI", "Dimension" };
            List<string> list = (from x in dataTable2.AsEnumerable()
                                 where search_key_lst.Any((string y) => x.Field<string>("SheetNo").Contains(y))
                                 select x.Field<string>("SheetNo")).Distinct().ToList();
            if (list.Count > 0)
            {
                List<DataTable> src_lst_tbl = new List<DataTable>();
                Get_ListTable(-1, dataTable2, new string[2] { "SheetNo", "FAI_No" }, ref src_lst_tbl, "Distribution");
                foreach (DataTable item2 in src_lst_tbl)
                {
                    if (item2.Rows.Count <= 0)
                    {
                        continue;
                    }

                    string item = myCode.checkDBNull(item2.Rows[0]["SheetNo"]);
                    if (list.IndexOf(item) == -1)
                    {
                        continue;
                    }

                    string text = myCode.checkDBNull(item2.Rows[0]["FAI_No"]);
                    if (!check_columns_existed(dataTable, text))
                    {
                        dataTable.Columns.Add(text);
                    }

                    if (dataTable.Rows.Count == 0)
                    {
                        for (int i = 0; i < 4; i++)
                        {
                            dataTable.Rows.Add();
                        }
                    }

                    dataTable.Rows[0][text] = item2.Rows[0]["NormDim"];
                    dataTable.Rows[1][text] = item2.Rows[0]["TolMax"];
                    dataTable.Rows[2][text] = item2.Rows[0]["TolMin"];
                    dataTable.Rows[3][text] = item2.Rows[0]["Distribution"];
                }
            }

            return dataTable;
        }
        public DataTable Load_Spec_fromFile_old(SqlConnection sqlcon, ExcelWorkbook workbook, string tar_ItemCode, string format_type = "NPI")
        {
            DataTable dataTable = new DataTable();
            AutoCompleteStringCollection autoCompleteStringCollection = new AutoCompleteStringCollection();
            string flt_str = TDMK_Code.filter_str(new string[2] { "ItemCode", "Remark" }, new string[2] { tar_ItemCode, format_type });
            autoCompleteStringCollection = TDMK_Code.Load_Item_Names_Filter(sqlcon, "FAI_Spec", "SheetNo", flt_str);
            List<string> list = new List<string>();

            List<string> source = new List<string> { "FAI", "SPC", "parentheses" };
            int count = workbook.Worksheets.Count;
            foreach (ExcelWorksheet tg in workbook.Worksheets)
            {
                if (!source.Any((string x) => tg.Name.Contains(x)))
                {
                    continue;
                }

                // tg.Activate();
                string cell = Find_Cell_Addr("Dim. No.", "B10", tg, left_to_right: false);
                string cell2 = Find_Cell_Addr("instrument", "B10", tg, left_to_right: false);
                string cell3 = Find_Offset(((_Worksheet)tg).get_Range((object)cell, Type.Missing).get_Offset((object)0, (object)1).get_AddressLocal(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing), tg, left_to_right: true);


                int num = ((_Worksheet)tg).get_Range((object)cell3, Type.Missing).Column - ((_Worksheet)tg).get_Range((object)cell2, Type.Missing).Column;
                Range range = ((_Worksheet)tg).get_Range((object)cell, Type.Missing).get_Offset((object)0, (object)num);
                Range range2 = ((_Worksheet)tg).get_Range((object)cell2, Type.Missing).get_Offset((object)0, (object)num);
                for (int i = 0; myCode.checkDBNull((dynamic)range.get_Offset((object)0, (object)i).get_Value(Type.Missing)) != ""; i++)
                {
                    string value = myCode.checkDBNull((dynamic)range.get_Offset((object)(-1), (object)i).get_Value(Type.Missing));
                    string text2 = myCode.checkDBNull((dynamic)range.get_Offset((object)0, (object)i).get_Value(Type.Missing)).Replace(" ", "");
                    string text3 = myCode.checkDBNull((dynamic)range.get_Offset((object)1, (object)i).get_Value(Type.Missing));
                    if (text3 == "")
                    {
                        text3 = "NA";
                    }

                    string value2 = myCode.checkDBNull((dynamic)range.get_Offset((object)6, (object)i).get_Value(Type.Missing));
                    string value3 = myCode.checkDBNull((dynamic)range.get_Offset((object)7, (object)i).get_Value(Type.Missing));
                    string text4 = text2 + "_" + text3;
                    string name = tg.Name;
                    string value4 = myCode.checkDBNull((dynamic)range2.get_Offset((object)0, (object)i).get_Value(Type.Missing));
                    if (check_columns_existed(dataTable, text4))
                    {
                        continue;
                    }

                    dataTable.Columns.Add(text4);
                    if (dataTable.Rows.Count == 0)
                    {
                        for (int j = 0; j < 6; j++)
                        {
                            dataTable.Rows.Add();
                        }
                    }

                    dataTable.Rows[0][text4] = text3;
                    dataTable.Rows[1][text4] = value2;
                    dataTable.Rows[2][text4] = value3;
                    dataTable.Rows[3][text4] = value;
                    dataTable.Rows[4][text4] = value4;
                    dataTable.Rows[5][text4] = tg.Name;
                }
            }

            return dataTable;
        }
        public DataTable Load_Spec_fromFile(SqlConnection sqlcon, ExcelWorkbook wb, string tar_ItemCode, List<string> extensions, string format_type = "NPI")
        {
            DataTable spec_dt = new DataTable();
            AutoCompleteStringCollection FAI_SheetNo_list = new AutoCompleteStringCollection();
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { tar_ItemCode, format_type });
            FAI_SheetNo_list = TDMK_Code.Load_Item_Names_Filter(sqlcon, "FAI_Spec", "SheetNo", filter_str);
            List<string> sht_keys = new List<string> { "FAI", "SPC", "parentheses" };
            int wrksheet_num = wb.Worksheets.Count;
            foreach (ExcelWorksheet tg in wb.Worksheets)
            {
                if (sht_keys.Any(x => tg.Name.Contains(x)))
                {
                    string dim_no_addr = Find_Cell_Addr("Dim. No.", "B10", tg, false);
                    string instrument_addr = Find_Cell_Addr("instrument", "B10", tg, false);
                    string FAI_data_addr = Find_Offset(tg.Cells[dim_no_addr].Offset(0, 1).Address, tg, true, "");
                    int off_set = tg.Cells[FAI_data_addr].End.Column - tg.Cells[instrument_addr].End.Column;
                    ExcelRangeBase sel_rgn = tg.Cells[dim_no_addr].Offset(0, off_set);// tg.Range["D19"];    //tg.Range["C17"];
                    ExcelRangeBase dev_rgn = tg.Cells[instrument_addr].Offset(0, off_set);// tg.Range["D23"];    //tg.Range["C21"]
                    int sel_inx = 0;
                    while (myCode.checkDBNull(sel_rgn.Offset(0, sel_inx).Value) != "")
                    {
                        double USL = 0;
                        double LSL = 0;
                        string t_checkside = myCode.checkDBNull(sel_rgn.Offset(-1, sel_inx).Value);
                        string t_FAIName = myCode.checkDBNull(sel_rgn.Offset(0, sel_inx).Value).Replace(" ", "");
                        string t_FAI_Setval = myCode.checkDBNull(sel_rgn.Offset(1, sel_inx).Value);
                        string t_FAI_UL = myCode.checkDBNull(sel_rgn.Offset(6, sel_inx).Value); ;// myCode.checkDBNull(sel_rgn.Offset[2, sel_inx].Value);
                        string t_FAI_LL = myCode.checkDBNull(sel_rgn.Offset(7, sel_inx).Value); ;// myCode.checkDBNull(sel_rgn.Offset[3, sel_inx].Value);
                        string t_FAI_Tol_plus = myCode.checkDBNull(sel_rgn.Offset(2, sel_inx).Value);
                        string t_FAI_Tol_minus = myCode.checkDBNull(sel_rgn.Offset(3, sel_inx).Value);
                        if ((t_FAI_Setval != "NA") && (t_FAI_Setval != ""))
                        {
                            double sv = Convert.ToDouble(t_FAI_Setval);
                            if (myCode.IsNumeric(t_FAI_Tol_plus))
                            {
                                USL = sv + Convert.ToDouble(t_FAI_Tol_plus);
                                t_FAI_UL = USL.ToString();
                            }
                            else
                            {
                                t_FAI_UL = t_FAI_Tol_plus;
                            }
                            if (myCode.IsNumeric(t_FAI_Tol_minus))
                            {
                                LSL = sv - Convert.ToDouble(t_FAI_Tol_minus);
                                t_FAI_LL = LSL.ToString();
                            }
                            else
                            {
                                t_FAI_LL = t_FAI_Tol_minus;
                            }
                        }
                        else
                        {
                            if (t_checkside == "SingleSide-USL")
                            {
                                //USL = Convert.ToDouble(t_FAI_Tol_plus);
                                t_FAI_UL = t_FAI_Tol_plus;// USL.ToString();
                            }
                            if (t_checkside == "SingleSide-LSL")
                            {
                                //LSL = Convert.ToDouble(t_FAI_Tol_minus);
                                t_FAI_LL = t_FAI_Tol_minus;// LSL.ToString();
                            }
                        }
                        string col_name = t_FAIName + "_" + t_FAI_Setval;
                        string t_FAI_sheetno = tg.Name;
                        string t_instrument = myCode.checkDBNull(dev_rgn.Offset(0, sel_inx).Value);
                        if (!check_columns_existed(spec_dt, col_name))
                        {
                            spec_dt.Columns.Add(col_name);
                            if (spec_dt.Rows.Count == 0)
                            {
                                for (int i = 0; i < 6; i++)
                                {
                                    spec_dt.Rows.Add();
                                }
                            }
                            spec_dt.Rows[0][col_name] = t_FAI_Setval; // Set val at Row =0
                            spec_dt.Rows[1][col_name] = t_FAI_UL; // UL at Row =1
                            spec_dt.Rows[2][col_name] = t_FAI_LL; // LL at Row =2
                            spec_dt.Rows[3][col_name] = t_checkside; // Check Side at Row =3
                            spec_dt.Rows[4][col_name] = t_instrument; // Instrument at Row =4
                            spec_dt.Rows[5][col_name] = tg.Name; // Instrument at Row =5
                        }
                        sel_inx++;
                    }
                }
            }
            return spec_dt;
        }
        public void Save_FAI_Spec(DataTable src_spec_dt, string tar_ItemCode, SqlConnection sqlcon, string type = "NPI")
        {
            TDMK_Code.Delelte_FilteredItem_arr("FAI_Spec", sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { tar_ItemCode, type }));
            string[] arr_item = new string[11]
            {
                "ID", "ItemCode", "LotNo", "NormDim", "TolMax", "TolMin", "Instrument", "FAI_No", "Distribution", "SheetNo",
                "Remark"
            };
            int num = TDMK_Code.SQL_MAX("FAI_Spec", "ID", sqlcon);
            foreach (DataColumn column in src_spec_dt.Columns)
            {
                string[] array = new string[11];
                array[1] = tar_ItemCode;
                array[10] = type;
                array[0] = num++.ToString();
                array[7] = column.ColumnName;
                array[3] = src_spec_dt.Rows[0][column].ToString();
                array[4] = src_spec_dt.Rows[1][column].ToString();
                array[5] = src_spec_dt.Rows[2][column].ToString();
                array[8] = src_spec_dt.Rows[3][column].ToString();
                array[6] = src_spec_dt.Rows[4][column].ToString();
                array[9] = src_spec_dt.Rows[5][column].ToString();
                TDMK_Code.insert_val_arr2("FAI_Spec", sqlcon, arr_item, array);
            }
        }



        //public bool Load_Spec2(string format_loc, string tar_ItemCode, string tar_LotNo, List<string> extensions, DataGridView tar_DGV_Spec, string format_type)
        //{
        //start_label: bool _result = false;
        //    //string format_type = "NPI";
        //    //if (rbMASS.Checked)
        //    //{
        //    //    format_type = "MASS";
        //    //}
        //    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { txtItemCode.Text, format_type });
        //    DataTable spec_dt = Load_FAI_Spec_ToTable(sqlcon, tar_ItemCode, format_type);
        //    if (spec_dt.Rows.Count == 0)
        //    {
        //        DataTable cur_dt = Load_Spec_fromFile(sqlcon, format_loc, tar_ItemCode, tar_LotNo, extensions, format_type);
        //        tar_DGV_Spec.DataSource = cur_dt;
        //        Save_FAI_Spec(cur_dt, tar_ItemCode, sqlcon, format_type);
        //    }
        //    //else
        //    //{
        //    //    if (MessageBox.Show(new Form { TopMost = true }, "Spec của ItemCode: " + txtItemCode.Text + " đã được cài đặt. Bạn muốn cập nhật?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
        //    //    {
        //    //        TDMK_Code.Delelte_FilteredItem_arr("FAI_Spec", sqlcon , filter_str);
        //    //        goto start_label;
        //    //    }
        //    //    else
        //    //    {
        //    //        DGV_Spec.DataSource = spec_dt;
        //    //        DGV_Spec.AutoResizeColumns();
        //    //        myCode.Disable_Sort_DGV(DGV_Spec);
        //    //    }
        //    //}
        //    return _result;
        //}
    }
}
