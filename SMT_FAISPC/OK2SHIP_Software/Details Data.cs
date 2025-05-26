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
using OK2SHIP_Software;
using OK2SHIP_Measurements;
using OK2SHIP_Lib;
using TDMK_SQL;

namespace OK2SHIP_Software
{
    public partial class Details_Data : Form
    {
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        myVar myCode = new myVar();
        TDMK_OK2SHIP myCode2 = new TDMK_OK2SHIP();
        public string ItemCode { get; set; }
        public string Lotno { get; set; }
        public string item_confirm;
        public Details_Data()
        {
            InitializeComponent();
        }
        public Details_Data(string _itemcode, string _lotno)
        {
            ItemCode = _itemcode;
            Lotno = _lotno;
            InitializeComponent();
        }

        private void Details_Data_Load(object sender, EventArgs e)
        {
            //ItemCode = "22B0500";
            //Lotno = "00003";

            char[] trim_char = new char[] { ' ', '\r', '\n' };
            string ignoredSheet = Path.Combine(myVar.app_path, "IgnoredSheet.txt");
            string[] ignoredList_arr = myCode.read_config_arr2(ignoredSheet);
            List<string> ignoredList = new List<string>() { "FAI", "SPC", "CPK", "IPQC", "Recycle", "Materials", "Coverpage", "User Guidelines", "Declaration and Contents" };
            ignoredList.AddRange(ignoredList_arr);
            lblDetail.Text = "OK2SHIP Report Data Contents " + ItemCode + "-" + Lotno;
            load_data(ItemCode + "-" + Lotno, tvItem_Details);
            DataTable result = myCode.Declaration_Results(ItemCode, Lotno);
            //bool confirm_en = true;
            if (result.Rows.Count == 0)
            {
                MessageBox.Show("No data", "Warning");
            }
            else
            {
                DGV_Results.DataSource = result;                        
                foreach (TreeNode tv in tvItem_Details.Nodes)
                {
                    string item_val = tv.Text.Trim(trim_char).ToUpper();
                    if (tv.Nodes.Count == 0)
                    {
                        foreach (DataGridViewRow dr in DGV_Results.Rows)
                        {
                            string dr_val = myCode.checkDBNull(dr.Cells["Items"].Value).Trim(trim_char).ToUpper();
                            if (TDMK_Code.check_exist_list_index(dr_val, ignoredList) == -1)
                            {
                                if (item_val.Contains(dr_val))
                                {
                                    dr.Cells["Results"].Style.BackColor = Color.Red;
                                }
                            }
                        }
                    }
                }
                foreach (DataGridViewRow dr in DGV_Results.Rows)
                {
                    string dr_val = myCode.checkDBNull(dr.Cells["Results"].Value);
                    switch (dr_val)
                    {
                        case "No data":
                            dr.Cells["Results"].Style.BackColor = Color.GreenYellow;
                            break;
                        case "Fail":
                            dr.Cells["Results"].Style.BackColor = Color.Red;
                            dr.Cells["Results"].ReadOnly = true;
                            break;
                        default:
                            if(dr_val.Contains("F/"))
                            {
                                string num = myCode2.Get_Number_String(dr_val.Replace("F", ""), '/').Trim();
                                if(num!="0")
                                {
                                    dr.Cells["Results"].Style.BackColor = Color.Red;
                                    dr.Cells["Results"].ReadOnly = true;
                                }
                            }
                            break;
                    }
                    //if ((dr_val == "No data") || (dr_val == "Fail"))
                    //{
                    //    dr.Cells["Results"].Style.BackColor = Color.Red;
                    //}
                }
                foreach (DataGridViewColumn dc in DGV_Results.Columns)
                {
                    if(dc.Name!="Results")
                    {
                        dc.ReadOnly = true;
                    }
                }
                DGV_Results.AutoResizeColumns();
                myCode2.Disable_Sort_DGV(DGV_Results);
                
            }
            //btnConfirm.Enabled = confirm_en;

        }
     
        public void Query_Data(string src_Itemcode, string src_Lotno, TreeView tar_tvDetail, DataGridView tar_DGV_process)
        {
            DataSet myDS = new DataSet();
            DataTable mydt = new DataTable();
            tar_DGV_process.Columns.Clear();
            //f_name = e.Node.Parent.Text + "-" + e.Node.Text;
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
            load_data(ItemCode+"-"+Lotno, tvItem_Details);
        }
        public void load_data(string tar_item_lot, TreeView tar_tvDetail)
        {
            string tar_folder = myVar.data_loc;// @"D:\Customer Projects\SEEV\TestAreas\OK2SHIP automation system";
            string[] folders = System.IO.Directory.GetDirectories(tar_folder);
            string[] main_list = new string[folders.Length-2];
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

        private void tvItem_Details_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Text.Contains(".xlsx"))
            {
                string f_path = e.Node.FullPath;
                TDMK_Code.open_excel_file(myVar.data_loc, f_path, "");
            }
            if (e.Node .Text .Contains ("FAI"))
            {
                FAI_Data tar_FAI = new FAI_Data(ItemCode, Lotno);
                tar_FAI.Show();
            }
            if (e.Node.Text.Contains("IPQC"))
            {
                IPQC_Data tar_IPQC = new IPQC_Data(ItemCode, Lotno, myVar.sqlcon_IPQC);
                tar_IPQC.Show();
            }
            //string stt = e.Node.Text.Split('.')[0];

            if (!e.Node.Text.Contains("IPQC") && !e.Node.Text.Contains(".xlsx") && !e.Node.Text.Contains("FAI"))
            {
                View_History fr1 = new View_History();
                fr1.itemcodevh_ = ItemCode;
                fr1.lotnovh_ = Lotno_Formated(Lotno);
                fr1.sheet_ = e.Node.Text.Split('.')[1];
                fr1.Show();
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

        private void btnTest_Click(object sender, EventArgs e)
        {
            create_random_data();
        }
        public void create_random_data()
        {
            int col_count = tvItem_Details.Nodes.Count + 5;//48;
            string[] items = new string[col_count];
            string[] item_vals = new string[col_count];
            Random rnd = new Random();
            items[0] = "ID";
            items[1] = "ItemCode";
            items[2] = "LotNo";
            items[3] = "MDate";
            for (int i = 0; i < col_count - 5; i++)
            {
                items[4 + i] = "[" + tvItem_Details.Nodes[i].Text + "]";
            }
            item_vals[1] = ItemCode;
            item_vals[2] = Lotno;
            item_vals[3] = DateTime.Now.ToString();
            item_vals[0] = (TDMK_Code.SQL_MAX("OK2SHIP_Data", "ID", FrmLogin.sqlcon) + 1).ToString();
            for (int k = 0; k < col_count - 5; k++)
            {
                item_vals[4 + k] = rnd.Next(0, 10).ToString() + "/F";
            }                
            TDMK_Code.insert_val_arr("OK2SHIP_Data", FrmLogin.sqlcon, items, item_vals);
        }

        private void btnConfirm_Click(object sender, EventArgs e)
        {
            //if (txtResult .Text !="")
            //{
            //    AutoCompleteStringCollection tg = TDMK_Code.Load_Item_Names_Filter(myVar.sqlcon_Declare, "Results1", "ID", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, Lotno }));
            //    if(tg.Count ==0)
            //    {
            //        int ID = TDMK_Code.SQL_MAX("Results1", "ID", myVar.sqlcon_Declare) + 1;
            //        TDMK_Code .insert_val_arr ("Results1",myVar .sqlcon_Declare , new string[] {"ID", "ItemCode", "LotNo", "["+ item_confirm+"]",""}, new string[] {ID.ToString() , ItemCode, Lotno, txtResult.Text ,"" });
            //    }
            //    else
            //    {
            //        string ID=  TDMK_Code.Get_item_val("Results1", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo"}, new string[] { ItemCode, Lotno}),"ID",myVar .sqlcon_Declare );
            //        TDMK_Code.updatebyID_val_arr("Results1", myVar.sqlcon_Declare, Convert.ToInt32 (ID), new string[] { "ID", "ItemCode", "LotNo", "[" + item_confirm + "]", "" }, new string[] { ID.ToString(), ItemCode, Lotno, txtResult.Text, "" });
            //    }
            //}
            if (DGV_Results .Rows.Count >0)
            {
                for(int i = 0; i< DGV_Results .RowCount;i++)
                {
                    string _result = checkDBNull( DGV_Results.Rows[i].Cells["Results"].Value);
                    string ID = checkDBNull(DGV_Results.Rows[i].Cells["ID"].Value);
                    TDMK_Code.update_item_val_filter("Results", myVar.sqlcon_Declare, "ID = " + ID, "Results", _result);                    
                }
                //string flt_cmd = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, Lotno });
                //DGV_Results.DataSource = TDMK_Code.Datatable_Filter(myVar.sqlcon_Declare, "Results", flt_cmd);
                MessageBox.Show("Update completed!");
            }
        }
        public string checkDBNull(object src_str)
        {
            string _result;
            if (src_str != null)
            {
                _result = src_str.ToString();
            }
            else
            {
                _result = "";
            }
            return _result;
        }

        private void tvItem_Details_AfterSelect(object sender, TreeViewEventArgs e)
        {
            //TreeNode cur_item = tvItem_Details.SelectedNode;
            //if (cur_item.Parent!=null )
            //{
            //    btnConfirm.Enabled = false;
            //}
            //else
            //{                
            //    //lblInfo.Text = "Please, confirm result for item: " + tvItem_Details.SelectedNode.Text;
            //    item_confirm = tvItem_Details.SelectedNode.Text;
            //    btnConfirm.Enabled = true;
            //}
            
        }
    }
}
