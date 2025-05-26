using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SQL;
using OK2SHIP_Lib;
using System.Data.SqlClient;
using System.IO;
using System.Security.Policy;

namespace OK2SHIP_Software
{
    public partial class Funtion_result : Form
    {
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        myVar myCode = new myVar();
        TDMK_OK2SHIP myCode2 = new TDMK_OK2SHIP();
        public SqlConnection sqlcon_ = null;
        public static string sel_DB = "OK2SHIP_Period2";
        public Funtion_result()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           DataTable dt = new DataTable();
            dt.Columns.Add("Zone");
            dt.Columns.Add("Result");


            SortedDictionary<string, string> result_all_zone = Stackup_result( txtItemCode.Text, txtLotNo.Text, sqlcon_);
            // SortedDictionary<string, string> result_all_zone = BVH_PTH_result(txtItemCode.Text, txtLotNo.Text, sqlcon_);
            //SortedDictionary<string, List<string>> result_all = Impedance_result(txtItemCode.Text, txtLotNo.Text, sqlcon_);
            //SortedDictionary<string, string> result_all_zone = BHAST_result(txtItemCode.Text, txtLotNo.Text, sqlcon_);
            foreach (var zone_result in result_all_zone)
            {
                dt.Rows.Add(zone_result.Key, zone_result.Value);

            }
            //foreach (var zone_result in result_all)
            //{
            //    string result_region = "";
            //    foreach(string item in zone_result.Value)
            //    {
            //        result_region += item + "__";
            //    }
            //     dt.Rows.Add(zone_result.Key, result_region);

            //}
            dataGridView1.DataSource = dt;
        }
        public string result(string process)
        {
            string slt = "";
            switch (process)
            {
                case "Stackup":
                    SortedDictionary<string, string> result_all_zone = Stackup_result(txtItemCode.Text, txtLotNo.Text, sqlcon_);
                    foreach (var zone_result in result_all_zone)
                    {
                        slt += zone_result.Key + ":" + zone_result.Value + ";";

                    }
                    break;
            }


            return slt;
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
        public SortedDictionary<string, string> Stackup_result( string ItemCode, string LotNo, SqlConnection sqlcon)
        {
            DataTable dt_data_summary = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
            //DataTable tbl_spec = TDMK_Code.Datatable_Filter(myVar.sqlcon_OK2SHIP, "STACKUP_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { ItemCode }));
            string[] arr_zone = dt_data_summary.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();
            SortedDictionary<string, string> result_all_zone = new SortedDictionary<string, string> {};
            foreach (string zone in arr_zone)
            {
                DataTable dt_zone = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo , zone}));
                DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Zone" }, new string[] { ItemCode, zone.Replace("Zone_", "") }));
                if (dt_spec.Rows.Count > 0)
                {
                    Double USL = Double.Parse(dt_spec.Rows[0]["USL"].ToString());
                    Double LSL = Double.Parse(dt_spec.Rows[0]["LSL"].ToString());

                    string[] arr_spec_detail = dt_spec.Rows[0]["Spec_detail"].ToString().Split(';');

                    string filter = TDMK_Code.filter_str(new string[] { "Zone" }, new string[] { zone });
                    DataView dv_zone = dt_zone.AsDataView();
                    List<DataTable> tbl_pcs = new List<DataTable> { };
                    Get_ListTable(-1, dv_zone.ToTable(), new string[] { "Pcs_No" }, ref tbl_pcs, "Data");
                    SortedDictionary<int, List<double>> dic_spec = new SortedDictionary<int, List<double>> { };
                    bool check = true;
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
                    int count_pcs_NG = 0;
                  
                    foreach (DataTable dt in tbl_pcs)
                    {
                        int s = 1; 
                        Double total_pcs = 0;

                        foreach (DataRow dr in dt.Rows)
                        {

                            if (s <= dic_spec.Count)
                            {
                                double value = Double.Parse(dr["Data"].ToString());
                                total_pcs += value;
                                double max = dic_spec[s][0];
                                double min = dic_spec[s][1];

                                if (value > max || value < min)
                                {
                                    check = false;
                                    count_pcs_NG++;
                                    break;

                                }
                                s++;
                            }
                        }
                        if (check)
                        {
                            if (total_pcs > USL || total_pcs < LSL)
                            {
                                check = false;
                                count_pcs_NG++;
                                break;

                            }
                          
                        }

                    }


                    string result_pcs = count_pcs_NG.ToString() + "F/" + tbl_pcs.Count.ToString();
                    result_all_zone.Add(zone, result_pcs);
                }
            }
            return result_all_zone;
            

        }
        public SortedDictionary<string, string> BVH_PTH_result( string ItemCode, string LotNo, SqlConnection sqlcon)
        {

            string[] arr_zone = { "BVH_WITH_BONDING_SHEET", "BVH_WITHOUT_BONDING_SHEET", "PLATED_THROUGH_HOLE" };
            SortedDictionary<string, string> result_all_zone = new SortedDictionary<string, string> { };
            foreach (string process in arr_zone)
            {
                DataTable dt_process = TDMK_Code.Datatable_Filter(sqlcon, process, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo}));
                DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_" + process.ToUpper(), TDMK_Code.filter_str(new string[] { "ItemCode"}, new string[] { ItemCode }));
                if (dt_spec.Rows.Count > 0)
                {
                    int count_pcs_NG = 0;
                    for (int i = 0; i < dt_process.Rows.Count; i++)
                    {
                        bool check = true;
                        for (int k = 4; k < 12; k++)
                        {
                            string USL = dt_spec.Rows[0][k - 1].ToString().Replace("/", string.Empty);
                            string LSL = dt_spec.Rows[1][k - 1].ToString().Replace("/", string.Empty);
                            if ( myCode.checkDBNull( dt_process.Rows[i][k]) != "")
                            {
                                if (!USL.Contains("NA"))
                                {
                                    if (Double.TryParse(dt_process.Rows[i][k].ToString(), out Double data))
                                    {
                                        if (data > Double.Parse(USL))
                                        {
                                            check = false;
                                            count_pcs_NG++;
                                            break;
                                        }
                                    }

                                }
                                if (!LSL.Contains("NA"))
                                { 
                                    if (Double.TryParse(dt_process.Rows[i][k].ToString(), out Double data))
                                    {
                                        if (data < Double.Parse(LSL))
                                        {
                                            check = false;
                                            count_pcs_NG++;
                                            break;
                                        }
                                    }

                                }
                            }
                        }
                    
                    }
                    string result_pcs = count_pcs_NG.ToString() + "/" + dt_process.Rows.Count.ToString();
                    result_all_zone.Add(process, result_pcs);
                }
            }
            return result_all_zone;


        }
        public SortedDictionary<string, List<string>> Impedance_result(string ItemCode, string LotNo, SqlConnection sqlcon)
        { 
            string[] arr_zone = { "IMPEDANCE", "TRACEWIDTH" };
            SortedDictionary<string, List<string>> result_all_zone = new SortedDictionary<string, List<string>> { };
            foreach (string process in arr_zone)
            {
                DataTable dt_process = TDMK_Code.Datatable_Filter(sqlcon, process + "_VAL", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, process + "_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { ItemCode }));
                List<string> lst_result_region = new List<string> { };
                if(process == "IMPEDANCE")
                {
                    if (dt_spec.Rows.Count > 0)
                    {
                        List<DataTable> tbl_region = new List<DataTable>() { };

                        Get_ListTable(-1, dt_process, new string[] { "Region" }, ref tbl_region, "Data");
                        foreach (DataTable dt in tbl_region)
                        {
                            int count_pcs_NG = 0;
                            for (int i = 0; i < dt.Rows.Count; i++)
                            { 
                                bool check = true;
                                int region = int.Parse(dt_process.Rows[i]["Region"].ToString());

                                if (region <= dt_spec.Rows.Count)
                                {
                                    Double max_target = Double.Parse(dt_spec.Rows[region - 1]["USL"].ToString());
                                    Double min_target = Double.Parse(dt_spec.Rows[region - 1]["LSL"].ToString());
                                    Double data = Double.Parse(dt_process.Rows[i]["Data"].ToString());
                                    if (data > max_target || data < min_target)
                                    {
                                        check = false;
                                        count_pcs_NG++; 
                                    } 
                                }  
                            }

                            string result_pcs = count_pcs_NG.ToString() + "/" + dt.Rows.Count.ToString();
                            lst_result_region.Add(result_pcs);
                        }
                    }
                }
                else if(process == "TRACEWIDTH")
                {
                    if (dt_spec.Rows.Count > 0)
                    {
                        List<DataTable> tbl_region = new List<DataTable>() { };

                        Get_ListTable(-1, dt_process, new string[] { "Region" }, ref tbl_region, "Data");
                        foreach (DataTable dt in tbl_region)
                        {
                            int count_pcs_NG = 0;
                            for (int i = 0; i < dt.Rows.Count; i++)
                            { 
                                bool check = true;
                                int region = int.Parse(dt.Rows[i]["Data_For"].ToString().Replace("Tracewidth_", ""));

                                if (region <= dt_spec.Rows.Count)
                                {
                                    Double max_target = Double.Parse(dt_spec.Rows[region - 1]["USL"].ToString());
                                    Double min_target = Double.Parse(dt_spec.Rows[region - 1]["LSL"].ToString());
                                    Double data = Double.Parse(dt_process.Rows[i]["Data"].ToString());
                                    if (data > max_target || data < min_target)
                                    {
                                        check = false;
                                        count_pcs_NG++;

                                    } 
                                }
                            }

                            string result_pcs = count_pcs_NG.ToString() + "/" + dt.Rows.Count.ToString();
                            lst_result_region.Add(result_pcs);
                        }
                    }
                    
                }
               
                result_all_zone.Add(process, lst_result_region); 
            }
            return result_all_zone;


        }

        public SortedDictionary<string, string> BHAST_result(string ItemCode, string LotNo, SqlConnection sqlcon)
        {
            SortedDictionary<string, string> result_all_zone = new SortedDictionary<string, string> { };
            DataTable dt_process = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_BHAST_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));

           
            for (int i = 4; i < dt_process.Columns.Count; i++)
            {
                bool check = true;
                foreach (DataRow dr in dt_process.Rows)
                {
                    if (myCode.checkDBNull(dr[i]) != "")
                    {
                        // Double data = Double.Parse(myCode.checkDBNull(dr[i]));
                        if (Double.TryParse(myCode.checkDBNull(dr[i]), out Double data))
                        {
                            if (data > Math.Pow(10, 4))
                            {
                                check = false;
                                break;
                            }
                        }
                    }
                }
                if (!check)
                {
                    result_all_zone.Add(dt_process.Columns[i].ColumnName, "False");
                } 
            }

            return result_all_zone;
        }
              
           
        public SqlConnection initial_data(string DB_name, bool sa_en)
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
            sqlcon_ = _sqlcon_OK2SHIP;
            return _sqlcon_OK2SHIP;
        }

        private void Test_funtion_Load(object sender, EventArgs e)
        {
            sqlcon_ = initial_data(sel_DB, true);
        }
    }
}
