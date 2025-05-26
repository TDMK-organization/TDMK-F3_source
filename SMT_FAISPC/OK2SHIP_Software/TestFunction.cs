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
using myExcel = Microsoft.Office.Interop.Excel;
using OK2SHIP_Measurements;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using OK2SHIP_Lib;

namespace OK2SHIP_Software
{
    public partial class TestFunction : Form
    {
        myVar myCode = new myVar();
        TDMK_OK2SHIP myCode2 = new TDMK_OK2SHIP();
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        TDMK_OK2SHIP.FAI_Spec[] testFAI_spec = new TDMK_OK2SHIP.FAI_Spec[1000];
        myExcel.Workbook exp_file;
        int process_id = 0;
        string sel_tbl = "";

        public TestFunction()
        {
            InitializeComponent();
        }

        private void TestFunction_Load(object sender, EventArgs e)
        {
            myVar.app_path = Application.StartupPath;
            myCode.initial_data();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //string ItemCode = "22B0473";// "22B0358";//"22B0886"; //22B0358-00003
            //string LotNo = "00004";// "00003";// "00001";
            //string src_file = @"D:\Customer Projects\SEEV\TestAreas\OK2SHIP Software\OK2SHIP automation system\0.OK2SHIP report format\Ok2ship Request SEEV D84 F-CAM SC-P Flex_P2(C4.0) DOE 22B0473.xlsm";
            //string test_file = @"D:\Customer Projects\SEEV\TestAreas\OK2SHIP Software\TestReport.xlsx";
            ////DirectoryInfo temp = new DirectoryInfo(@"D:\Customer Projects\SEEV\Giai Doan 2\TDMK Source Code\TestAreas\OK2SHIP Software\OK2SHIP automation system\0.OK2SHIP report format");
            ////FileInfo[] f_lst = temp.GetFiles();
            ////foreach (var f in f_lst)
            ////{
            ////    if(f.FullName.Contains(txtItemCode.Text))
            ////    {
            ////        txtTest.Text = f.Name;
            ////    }
            ////}
            //string f_loc = @"D:\Customer Projects\SEEV\Giai Doan 2\TDMK Source Code\TestAreas\OK2SHIP Software\OK2SHIP automation system\0.OK2SHIP report format";
            //KeyValuePair<string, string> result = myCode.Get_Format_Info(txtItemCode.Text, f_loc);
            //txtTemp.Text = result.Key;
            //txtTest.Text = result.Value;


            //DataSet ds = new DataSet();
            //TDMK_Code.fill_dataset(ds, "Confirm_History", myVar.sqlcon_OK2SHIP);
            //DataTable src_dt = ds.Tables[0];
            //DataTable History_dt = Get_Historical_data(src_dt, "08:30:00", "10:20:00");
            //if(History_dt.Rows.Count>0)
            //{
            //    List<string> tar_list = History_dt.AsEnumerable().Select(x => x.Field<string>("ItemCode")).ToList();
            //    DGV_Data.DataSource = get_approve_data(tar_list, myVar.sqlcon_OK2SHIP);
            //}
            //else
            //{
            //    MessageBox.Show("No data");
            //}

            //DataTable dt = TDMK_Code.Datatable_Filter(myVar.sqlcon_OK2SHIP, "Items_Details", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
            //DGV_Data.DataSource = dt;
            //List<string> main_items = dt.AsEnumerable().Select(x => x.Field<string>("Items")).Distinct().ToList();
            //List<string> tar_items = new List<string>();
            //char[] trim_char = new char[] { ' ', '-','_' };
            //foreach(string item in main_items)
            //{
            //    string temp = item;
            //    if(item.Contains("FAI")||item.Contains("SPC"))
            //    {
            //        temp = "FAI";
            //    }
            //    //temp = myString(temp, trim_char);
            //    if (tar_items.IndexOf(temp) ==-1)
            //    {
            //        tar_items.Add(temp);
            //    }
            //}
            //listBox1.DataSource = OK2SHIP_Item_list(txtItemCode.Text,txtLotNo.Text,"Results",myVar.sqlcon_Declare);

            //string file = @"D:\Customer Projects\SEEV\SMT Project\Source Code\TestAreas\Format\FAI\OQC Dimension-8DRWD-719059-00129-1-10-2023.xlsx";
            //string file = @"D:\Customer Projects\SEEV\SMT Project\Source Code\TestAreas\OK2SHIP Automation System\Report Format\FAI\OQC Dimension 2GVCD-71798T.xlsx";
            string file = @"D:\Customer Projects\SEEV\SMT Project\Source Code\TestAreas\OK2SHIP Automation System\Report Format\FAI\SEEV-D48 R-cam (AV-H)_C5.0 _821-04933-01 7S0011.xlsm";
            myExcel.Workbook wrkbk = TDMK_Code.open_excel_file(file, "", "");
            myExcel.Worksheet wrksht = wrkbk.Sheets["FAI 1"];
            //myExcel.Worksheet wrksht = wrkbk.Sheets[1];
            string dim_no_addr = Find_Cell_Addr("Dim. No.", "B10", wrksht,false);
            string instrument_addr = Find_Cell_Addr("instrument", "B10", wrksht,false);
            //string sample_addr = Find_Cell_Addr("Sample No", "B10", wrksht,false);
            string data_addr = Find_Offset(wrksht.Range[dim_no_addr].Offset[0,1].AddressLocal, wrksht, true, "");
            int off_set = wrksht.Range[data_addr].Column - wrksht.Range[dim_no_addr].Column;
            //string sample_qty = Find_Offset(sample_addr, wrksht, false, "Judgement");
            Dictionary<string, int> fai_loc = Find_FAI_addr_qty(dim_no_addr, wrksht, false);
            listBox1.Items.Add(wrksht.Range[dim_no_addr].Offset[0,off_set].AddressLocal);
            listBox1.Items.Add(wrksht.Range[instrument_addr].Offset[0, off_set].AddressLocal);
            //listBox1.Items.Add(sample_addr);
            listBox1.Items.Add(data_addr);
            if(fai_loc.Count>0)
            {
                string addr = fai_loc.Keys.ToList()[0];
                listBox1.Items.Add(addr);
                listBox1.Items.Add(fai_loc[addr].ToString());
            }
            
            //listBox1.Items.Add(sample_qty);
            //int qty = wrksht.Range[sample_qty].Row - wrksht.Range[sample_addr].Row-1;
            //listBox1.Items.Add(qty.ToString());

        }
        public string Find_Cell_Addr(string search_key, string start_addr, myExcel.Worksheet tar_wrksht, bool left_to_right)
        {
            string result = "";
            myExcel.Range cycle_rgn = tar_wrksht.Range[start_addr];
            for (int i = 0; i < 100; i++)
            {
                myExcel.Range sel_rgn = cycle_rgn.Offset[i, 0];
                if(left_to_right)
                {
                    sel_rgn = cycle_rgn.Offset[0, i];
                }
                string sel_rgn_val = myCode.checkDBNull(sel_rgn.Value);
                if (sel_rgn_val.Replace(" ", "").ToUpper() == search_key.Replace(" ", "").ToUpper())
                {
                    result = sel_rgn.AddressLocal;
                    break;
                }
            }
            return result;
        }
        public string Find_Offset(string start_addr, myExcel.Worksheet tar_wrksht,bool left_to_right, string search_key = "")
        {
            string result = "";
            myExcel.Range cycle_rgn = tar_wrksht.Range[start_addr];
            int c_offset = 0;
            int empty_count = 0;
            while (true)
            {
                myExcel.Range sel_rgn = cycle_rgn.Offset[c_offset,0];
                if (left_to_right)
                {
                    sel_rgn = cycle_rgn.Offset[0, c_offset];
                }
                string sel_rgn_val = myCode.checkDBNull(sel_rgn.Value);
                if (sel_rgn_val == "")
                {
                    empty_count++;
                    if (empty_count > 3)
                    {
                        break;
                    }
                }
                else
                {
                    empty_count = 0;
                }
                if (search_key!="")
                {
                    if (sel_rgn_val.Replace(" ", "").ToUpper() == search_key.Replace(" ", "").ToUpper())
                    {
                        result = sel_rgn.AddressLocal;
                        break;
                    }
                }
                else
                {
                    if(empty_count==0)
                    {
                        result = sel_rgn.AddressLocal;
                        break;
                    }    
                }
                c_offset++;
            }
            return result;
        }
        public Dictionary<string,int> Find_FAI_addr_qty(string start_addr, myExcel.Worksheet tar_wrksht, bool left_to_right)
        {
            Dictionary<string, int> result = new Dictionary<string, int>();
            myExcel.Range cycle_rgn = tar_wrksht.Range[start_addr];
            int offset = 0;
            int empty_count = 0;
            string data_addr = "";
            int qty = 0;
            while (true)
            {
                myExcel.Range sel_rgn = cycle_rgn.Offset[offset, 0];
                if (left_to_right)
                {
                    sel_rgn = cycle_rgn.Offset[0, offset];
                }
                string sel_rgn_val = myCode.checkDBNull(sel_rgn.Value);
                if (sel_rgn_val == "")
                {
                    empty_count++;
                    if (empty_count > 3)
                    {
                        break;
                    }
                }
                else
                {
                    if(myCode2.IsNumeric(sel_rgn_val))
                    {
                        if(data_addr=="")
                        {
                            data_addr = sel_rgn.AddressLocal;
                        }
                        qty++;
                    }
                    empty_count = 0;
                }
                offset++;
            }
            if(data_addr!="")
            {
                result.Add(data_addr, qty);
            }
            return result;
        }
        public string remove_special_char(string src_str, List<char> remove_char)
        {
            return new string(src_str.Where(x => remove_char.IndexOf(x) == -1).ToArray());
        }
        public string myString(string src_str, char[] replace_char)
        {
            string result = ""; //replace_char.Aggregate(src_str,(c1,c2)=>c1.Replace(c2,'\0'));
            char[] src_char = src_str.ToArray();
            foreach(char t in src_char)
            {
                if (replace_char.ToList().IndexOf(t)==-1)
                {
                    result = result + t.ToString();
                }
            }
            return result;
        }
        public List<string> OK2SHIP_Item_list(string ItemCode, string LotNo, string tar_table, SqlConnection sqlcon)
        {
            List<string> result = new List<string>();
            DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, tar_table, TDMK_Code.filter_str(new string[] { "ItemCode","LotNo" }, new string[] { ItemCode, LotNo}));
            DGV_Data.DataSource = dt;
            List<string> main_items = dt.AsEnumerable().Select(x => x.Field<string>("Items")).Distinct().ToList();
            foreach (string item in main_items)
            {
                string temp = item;
                if (item.Contains("FAI") || item.Contains("SPC"))
                {
                    temp = "FAI";
                }
                if (result.IndexOf(temp) == -1)
                {
                    result.Add(temp);
                }
            }
            return result;
        }

        public int count_excel_data(string ItemCode, string LotNo)
        {
            int ignor_sum = 0;
            int excel_file_sum = 0;
            string ignoredSheet = Path.Combine(myVar.app_path, "IgnoredSheet.txt");
            List<string> ignoredList = new List<string>() { "FAI", "SPC", "CPK", "IPQC", "Recycle", "Materials", "Coverpage", "User Guidelines", "Declaration and Contents" };
            string[] ignoredList_arr = myCode. read_config_arr2(ignoredSheet);
            ignoredList.AddRange(ignoredList_arr);
            listBox3.DataSource = ignoredList;
            excel_file_sum += ignoredList_arr.Length;
            string tar_item_lot = ItemCode + "-" + LotNo;
            string tar_folder = myVar.data_loc;
            string[] folders = System.IO.Directory.GetDirectories(tar_folder);
            string[] main_list = new string[folders.Length];
            int f_inx = 0;
            foreach (string c in folders)
            {
                FileInfo f = new FileInfo(c);
                string[] temp = f.Name.Split('.');
                if (myCode2.IsNumeric(temp[0]) && (temp[0]!="0"))
                {
                    main_list[f_inx] = f.Name;
                    f_inx++;
                }
            }
            Array.Resize(ref main_list, f_inx);
            //listBox1.DataSource = main_list.ToList();
            foreach (string t in main_list)
            {
                listBox1.Items.Add(t);
                if (TDMK_Code.check_exist_list_index(t, ignoredList) == -1)
                {
                    string node_f = Path.Combine(tar_folder, t);
                    string[] node_name = Directory.GetFiles(node_f, "*.xlsx").Where(s => s.Contains(tar_item_lot)).ToArray();
                    if (node_name.Length > 0)
                    {
                        listBox2.Items.Add(node_name[0]);
                        excel_file_sum++;
                    }
                    
                }
                else
                {
                    listBox2.Items.Add("Ignored_Items");
                    ignor_sum++;
                }
            }
            txtTest.Text = ignor_sum.ToString();
            
            return excel_file_sum;
        }
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {


        }

        private void button2_Click(object sender, EventArgs e)
        {

            //string LotNo = "00004";
            //string f_name = Path.GetFileNameWithoutExtension(txtTest.Text);
            //string ItemCode = f_name.Substring(f_name.Length - 12);
            //f_name = f_name.Replace("Ok2ship Request SEEV", "").Replace("Flex","#");
            //txtTemp.Text= f_name.Split('#')[0].Trim();
            char[] trim_char = new char[] { ' ', '-', '_' };
            List<string> tar_process = new List<string>() { "ELECTRICAL", "CQRA_REFLOW", "CQRA_HOT_OIL", "CQRA_THERMAL_CYCLING", "CQRA_HEAT_SOAK", "CQRA_THERMAL_SHOCK", "CQRA_BENDING", "CQRA_THERMAL_CYCLING_AND_BEND", "CQRA_HEAT_SOAK_AND_BEND", "CQRA_SURVIVAL_REFLOW", "CQRA_SURVIVAL_HOT_OIL" };
            DataTable dt = (DataTable)DGV_Data.DataSource;
            var query = from p in dt.AsEnumerable().Where(x => tar_process.Any(t => myString(t, trim_char).ToUpper() == myString(x.Field<string>("Items"), trim_char).ToUpper())) select p;
            List<string> mylist = new List<string>();
            foreach (var q in query)
            {
                int inx = dt.Rows.IndexOf(q);
                DGV_Data.Rows[inx].Cells["ID"].Style.BackColor = Color.Red;
                var myProcess = from p in tar_process where myString(p, trim_char).ToUpper() == myString(q.Field<string>("Items"), trim_char).ToUpper() select p;
                string sel_process = myProcess.First().ToString();
                dt.Rows[inx]["Results"] = myCode.OK2SHIP_Process_Summary(txtItemCode.Text, txtLotNo.Text, sel_process, myVar.sqlcon_OK2SHIP_Period2);
            }
            listBox2.DataSource = mylist;
            //int r_inx = 0;
            //foreach(DataRow dr in dt.Rows)
            //{
            //    string cell_val = myString(dr["Items"].ToString(), trim_char).ToUpper();
            //    foreach(string p in tar_process)
            //    {
            //        if(myString(p,trim_char).ToUpper()==cell_val)
            //        {
            //            DGV_Data.Rows[r_inx].Cells["ID"].Style.BackColor = Color.Red;
            //        }
            //    }
            //    r_inx++;
            //}


        }
        
        public void Fill_NA_Data(myExcel.Workbook tar_wrkbook, string wrksheet_name, string tar_data_rgn, int row_num)
        {
            myExcel.Worksheet cur_wrksht = tar_wrkbook.Sheets[wrksheet_name];
            myExcel.Range data_rgn = cur_wrksht.Range[tar_data_rgn];
            int col_inx = 0;
            while (myCode.checkDBNull(data_rgn.Offset[-2, col_inx].Value) != "")
            {

                for (int i = 0; i < row_num; i++)
                {
                    string rgn_val = myCode.checkDBNull(data_rgn.Offset[i, col_inx].Value);
                    if (rgn_val == "")
                    {
                        data_rgn.Offset[i, col_inx].Value = "N/A";
                    }
                }

                col_inx++;
            }
        }
    
        private void button3_Click(object sender, EventArgs e)
        {
            string src_string = txtItemCode.Text;
            string chk_string = txtLotNo.Text;
            if(src_string.Contains(chk_string))
            {
                MessageBox.Show("Contain", "Info");
            }
            else
            {
                MessageBox.Show("NOT include", "Info");
            }
        }
        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void RB_Etching_Process_CheckedChanged(object sender, EventArgs e)
        {
            sel_tbl = myCode.get_target_Table(RB_Etching_Process);
        }
        public int get_process_it(RadioButton src_RB, string[] process_lst)
        {
            return TDMK_Code.check_exist_list_index2(src_RB.Text, process_lst.ToList());
        }

        private void RB_CopperPlating_CheckedChanged(object sender, EventArgs e)
        {
            sel_tbl = myCode.get_target_Table(RB_CopperPlating);
        }

        private void RB_Printing_Process_CheckedChanged(object sender, EventArgs e)
        {
            sel_tbl = myCode.get_target_Table(RB_Printing_Process);
        }

        private void RB_UV_Process_CheckedChanged(object sender, EventArgs e)
        {
            sel_tbl = myCode.get_target_Table(RB_UV_Process);
        }

        private void RB_CoverLay_Process_CheckedChanged(object sender, EventArgs e)
        {
            sel_tbl = myCode.get_target_Table(RB_CoverLay_Process);
        }

        private void RBRoughness_CheckedChanged(object sender, EventArgs e)
        {
            sel_tbl = "Roughness";
        }

        private void RBAU_NI_CheckedChanged(object sender, EventArgs e)
        {
            sel_tbl = "AU_NI";
        }

        public DataTable get_approve_data(List<string> ItemCode_list, SqlConnection sqlcon)
        {
            DataTable result = new DataTable();
            DataSet ds = new DataSet();
            TDMK_Code.fill_dataset(ds, "Approve_Data", sqlcon);
            DataTable dtDataTable = ds.Tables[0];
            var rowsToDelete = from r in dtDataTable.AsEnumerable() join c in ItemCode_list on r.Field<string>("ItemCode") equals c into g where g.Any() select r;
            if(rowsToDelete.Count() > 0)
            {
                result = rowsToDelete.CopyToDataTable();
            }
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
