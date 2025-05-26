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
using OK2SHIP_Measurements;
using myExcel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;

using OK2SHIP_Lib;

namespace OK2SHIP_Software
{
    public partial class FormatSetup : Form
    {
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public TDMK_OK2SHIP myCode = new TDMK_OK2SHIP();
        myVar myCode2 = new myVar();
       // TDMK_OK2SHIP.FAI_Spec[] testFAI_spec = new TDMK_OK2SHIP.FAI_Spec[1000];
        List<TDMK_OK2SHIP.FAI_Spec> _testFAI_spec = new List<TDMK_OK2SHIP.FAI_Spec>();
        public string program_name;
        //public SqlConnection sqlcon_SMT;
        public FormatSetup()
        {
            InitializeComponent();
        }

        private void RB_FAI_CheckedChanged(object sender, EventArgs e)
        {
            if (RB_FAI.Checked)
            {
                program_name = "FAI";
                lblInfo.Text = "FAI/SPC Specification List";
                DGV_Spec.DataSource = null;
                DGV_Spec.Columns.Clear();
            }
        }
        private void RB_IPQC_CheckedChanged(object sender, EventArgs e)
        {
            if (RB_IPQC.Checked)
            {
                program_name = "IPQC";
                lblInfo.Text = "IPQC Specification List";
                DGV_Spec.DataSource = null;
                DGV_Spec.Columns.Clear();
            }
        }

        private void RB_Format_CheckedChanged(object sender, EventArgs e)
        {
            if (RB_Format.Checked)
            {
                program_name = "Format_Structure";
                DGV_Spec.DataSource = null;
                DGV_Spec.Columns.Clear();
            }
        }
        public bool Load_SheetName(string format_loc, string tar_ItemCode, string file_extension)
        {
            bool _result = false;
            string tar_format_file;// = Path.Combine(format_loc, tar_format + file_extension);    
            List<string> lst_sheet = new List<string>();
            string[] file_format_lst = Directory.GetFiles(format_loc, "*" + file_extension).Where(s => s.Contains(tar_ItemCode)).ToArray();
            string[] items = new string[] { "ID", "ItemCode", "Items" };
            string[] items_val = new string[3];
            items_val[1] = txtItemCode.Text;
            DataTable dt = TDMK_Code.Datatable_Filter(myVar.sqlcon_OK2SHIP , "Items_Details", "ItemCode = '" + tar_ItemCode + "'");
            if (dt.Rows.Count == 0)
            {
                if (file_format_lst.Length > 0)
                {
                    tar_format_file = file_format_lst[0];
                    myExcel.Workbook tar_wkbook = TDMK_Code.open_excel_file(tar_format_file, "", "");
                    int wrksheet_num = tar_wkbook.Worksheets.Count;
                    foreach (myExcel.Worksheet tg in tar_wkbook.Worksheets)
                    {
                        items_val[0] = (TDMK_Code.SQL_MAX("Items_Details", "ID", myVar.sqlcon_OK2SHIP) + 1).ToString();
                        items_val[2] = tg.Name;
                        TDMK_Code.insert_val_arr("Items_Details", myVar.sqlcon_OK2SHIP, items, items_val);
                    }
                    MessageBox.Show(new Form { TopMost = true },"Completed setup OK2SHIP Components results!", "Warning");
                    dt = TDMK_Code.Datatable_Filter(myVar.sqlcon_OK2SHIP, "Items_Details", "ItemCode = '" + tar_ItemCode + "'");
                    lstSheet.DataSource = dt.AsEnumerable().Select(r => r.Field<string>("Items")).ToList();
                    _result = true;
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true },"File format not found!", "Warning");
                    _result = false;
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },"OK2SHIP Components result already installed!", "Warning");
                lstSheet.DataSource = dt.AsEnumerable().Select(r => r.Field<string>("Items")).ToList();
            }
            return _result;
        }
        public bool Load_SheetName2(string format_loc, string tar_ItemCode, string file_extension)
        {
            bool _result = false;
            string tar_format_file;// = Path.Combine(format_loc, tar_format + file_extension);    
            List<string> lst_sheet = new List<string>();
            string[] file_format_lst = Directory.GetFiles(format_loc, "*" + file_extension).Where(s => s.Contains(tar_ItemCode)).ToArray();
            string[] items = new string[] { "ID", "ItemCode", "Items", "Address" };
            string[] items_val = new string[4];
            items_val[1] = txtItemCode.Text;
        start_label: DataTable dt = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "Items_Details", "ItemCode = '" + tar_ItemCode + "'");
            if (dt.Rows.Count == 0)
            {
                if (file_format_lst.Length > 0)
                {
                    tar_format_file = file_format_lst[0];
                    myExcel.Workbook tar_wkbook = TDMK_Code.open_excel_file(tar_format_file, "", "");
                    myExcel.Worksheet declare_sht = tar_wkbook.Sheets["Table of Contents"];
                    myExcel.Range item_rgn = declare_sht.Range["B17"];
                    myExcel.Range addr_rgn = declare_sht.Range["I17"];
                    //int wrksheet_num = tar_wkbook.Worksheets.Count;
                    int wrksht_inx = 0;
                    foreach (myExcel.Worksheet tg in tar_wkbook.Worksheets)
                    {
                        items_val[3] = addr_rgn.Offset[wrksht_inx, 0].AddressLocal;
                        items_val[0] = (TDMK_Code.SQL_MAX("Items_Details", "ID", myVar.sqlcon_SMT) + 1).ToString();
                        items_val[2] = tg.Name;
                        TDMK_Code.insert_val_arr("Items_Details", myVar.sqlcon_SMT, items, items_val);
                        wrksht_inx++;
                    }                    
                    dt = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "Items_Details", "ItemCode = '" + tar_ItemCode + "'");
                    lstSheet.DataSource = dt.AsEnumerable().Select(r => r.Field<string>("Items")).ToList();
                    MessageBox.Show(new Form { TopMost = true },"Completed setup OK2SHIP Components results!", "Warning");
                    _result = true;
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true },"File format not found!", "Warning");
                    _result = false;
                }
            }
            else
            {
                if (MessageBox.Show(new Form { TopMost = true },"OK2SHIP Components result already installed. Do you want to update it?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    TDMK_Code.Delelte_FilteredItem_arr("Items_Details", myVar.sqlcon_SMT, "ItemCode = '" + txtItemCode.Text + "'");
                    TDMK_Code.Delelte_FilteredItem_arr("Results", myVar.sqlcon_SMT, "ItemCode = '" + txtItemCode.Text + "'");
                    goto start_label;
                }
                else
                {
                    lstSheet.DataSource = dt.AsEnumerable().Select(r => r.Field<string>("Items")).ToList();
                }
            }
            return _result;
        }
        public void Setup_IPQC_Spec(string f_loc, string tar_ItemCode, SqlConnection tar_sqlcon)
        {
            DataSet myDS = new DataSet();
            myCode.TDMK_Code.fill_dataset(myDS, "All_Items", tar_sqlcon);
            DataTable Spec_addr_tbl = myDS.Tables[0];
            string[] process_arr = new string[] { "Etching_Process", "UV_Process", "Copper_Plating_Process", "Cover_Lay_Process", "Printing_Process", "AU Plating Process" };
            List<string>[] itemLst_arr = new List<string>[6];
            List<string>[] ColName_List_arr= new List<string>[6];
            List<string>[] Process_name_lst_arr= new List<string>[6];
            
            for(int i=0;i<6;i++)
            {
                itemLst_arr[i]= Spec_addr_tbl.AsEnumerable().Where(r=>r.Field<string>("Process_Name").Equals(process_arr[i])).Select(r => r.Field<string>("Report_Name")).ToList();
                ColName_List_arr[i] = Spec_addr_tbl.AsEnumerable().Where(r => r.Field<string>("Process_Name").Equals(process_arr[i])).Select(r => r.Field<string>("Col_Name")).ToList();
                Process_name_lst_arr[i] = Spec_addr_tbl.AsEnumerable().Where(r => r.Field<string>("Process_Name").Equals(process_arr[i])).Select(r => r.Field<string>("Process_Name")).ToList();
            }


            //List<string> itemLst = Spec_addr_tbl.AsEnumerable().Select(r => r.Field<string>("Report_Name")).ToList();
            //List<string> ColName_List = Spec_addr_tbl.AsEnumerable().Select(r => r.Field<string>("Col_Name")).ToList();
            //List<string> Process_name_lst = Spec_addr_tbl.AsEnumerable().Select(r => r.Field<string>("Process_Name")).ToList();
            List<string> cur_itemLst = new List<string>();
            List<string> cur_ColName_List = new List<string>();
            List<string> cur_Process_name_lst = new List<string>();
            char[] trim_chr = { ':', ' ', '\r', '\n' };
            string[] items = new string[] { "ID", "ItemCode", "Item", "Spec", "Address", "Remark", "Col_Name" };
            int inx = 0;
            int arr_inx = 0;
        start_label: AutoCompleteStringCollection ItemCode_lst = TDMK_Code.Load_Item_Names_Filter(tar_sqlcon, "SpecList", "ItemCode", "ItemCode = '" + tar_ItemCode + "'");
            if (ItemCode_lst.Count == 0)
            {
                string[] file_format_lst = Directory.GetFiles(f_loc, "*.xlsm").Where(s => s.Contains(tar_ItemCode)).ToArray();
                if (file_format_lst.Length > 0)
                {
                    string f_name = file_format_lst[0];
                    myExcel.Workbook cur_wrk = myCode.TDMK_Code.open_excel_file(f_name, "", "");
                    myExcel.Worksheet cur_wrksht = cur_wrk.Sheets["IPQC Data"];
                    myExcel.Range item_rgn = cur_wrksht.Range["B13"];
                    myExcel.Range spec_rgn = cur_wrksht.Range["B14"];
                    string[] layer_id = new string[] { "L1", "L2", "L3", "L4" };
                    while (myCode.checkDBNull(item_rgn.Offset[0, inx].Value) != "")
                    {
                        string rgn_val = myCode.checkDBNull(item_rgn.Offset[0, inx].Value);
                        string spec_val = myCode.checkDBNull(spec_rgn.Offset[0, inx].Value);
                        if(spec_val!="")
                        {
                            spec_val = spec_val.Replace("\u2265", ">=");
                            spec_val = spec_val.Replace("\u2264", "<=");
                            if (item_rgn.Offset[-1, inx].MergeCells)
                            {
                                myExcel.Range temp_rgn = item_rgn.Offset[-1, inx].MergeArea.Cells[1, 1];
                                string temp_rgn_val = myCode.checkDBNull(temp_rgn.Value);
                                int pro_inx = 0;
                                switch (temp_rgn_val.Trim(trim_chr).Replace(" ", "").ToUpper())
                                {
                                    case "DESDATA":
                                        pro_inx = 0;// "Etching_Process";
                                        for (int i = 0; i < 4; i++)
                                        {
                                            if (rgn_val.Contains(layer_id[i]))
                                            {
                                                cur_itemLst = Spec_addr_tbl.AsEnumerable().Where(r => r.Field<string>("Item_Name").Contains(layer_id[i]) && r.Field<string>("Process_Name").Equals(process_arr[pro_inx])).Select(r => r.Field<string>("Report_Name")).ToList();
                                                cur_ColName_List = Spec_addr_tbl.AsEnumerable().Where(r => r.Field<string>("Item_Name").Contains(layer_id[i]) && r.Field<string>("Process_Name").Equals(process_arr[pro_inx])).Select(r => r.Field<string>("Col_Name")).ToList();
                                                cur_Process_name_lst = Spec_addr_tbl.AsEnumerable().Where(r => r.Field<string>("Item_Name").Contains(layer_id[i]) && r.Field<string>("Process_Name").Equals(process_arr[pro_inx])).Select(r => r.Field<string>("Process_Name")).ToList();
                                                break;
                                            }
                                        }
                                        break;
                                    case "INNERLAYERVIA[LASER]":
                                        pro_inx = 1;// "UV_Process";
                                        cur_itemLst = Spec_addr_tbl.AsEnumerable().Where(r => r.Field<string>("Col_Name").Contains("INNER") && r.Field<string>("Process_Name").Equals(process_arr[pro_inx])).Select(r => r.Field<string>("Report_Name")).ToList();
                                        cur_ColName_List = Spec_addr_tbl.AsEnumerable().Where(r => r.Field<string>("Col_Name").Contains("INNER") && r.Field<string>("Process_Name").Equals(process_arr[pro_inx])).Select(r => r.Field<string>("Col_Name")).ToList();
                                        cur_Process_name_lst = Spec_addr_tbl.AsEnumerable().Where(r => r.Field<string>("Col_Name").Contains("INNER") && r.Field<string>("Process_Name").Equals(process_arr[pro_inx])).Select(r => r.Field<string>("Process_Name")).ToList();
                                        break;
                                    case "OUTERLAYERVIA[LASER]":
                                        pro_inx = 1;// "UV_Process";
                                        cur_itemLst = Spec_addr_tbl.AsEnumerable().Where(r => r.Field<string>("Col_Name").Contains("OUTER") && r.Field<string>("Process_Name").Equals(process_arr[pro_inx])).Select(r => r.Field<string>("Report_Name")).ToList();
                                        cur_ColName_List = Spec_addr_tbl.AsEnumerable().Where(r => r.Field<string>("Col_Name").Contains("OUTER") && r.Field<string>("Process_Name").Equals(process_arr[pro_inx])).Select(r => r.Field<string>("Col_Name")).ToList();
                                        cur_Process_name_lst = Spec_addr_tbl.AsEnumerable().Where(r => r.Field<string>("Col_Name").Contains("OUTER") && r.Field<string>("Process_Name").Equals(process_arr[pro_inx])).Select(r => r.Field<string>("Process_Name")).ToList();
                                        break;
                                    case "INNERLAYERVIA[PLATING]":
                                        pro_inx = 2;// "Copper_PLating_Process";
                                        cur_itemLst = Spec_addr_tbl.AsEnumerable().Where(r => r.Field<string>("Col_Name").Contains("INNER") && r.Field<string>("Process_Name").Equals(process_arr[pro_inx])).Select(r => r.Field<string>("Report_Name")).ToList();
                                        cur_ColName_List = Spec_addr_tbl.AsEnumerable().Where(r => r.Field<string>("Col_Name").Contains("INNER") && r.Field<string>("Process_Name").Equals(process_arr[pro_inx])).Select(r => r.Field<string>("Col_Name")).ToList();
                                        cur_Process_name_lst = Spec_addr_tbl.AsEnumerable().Where(r => r.Field<string>("Col_Name").Contains("INNER") && r.Field<string>("Process_Name").Equals(process_arr[pro_inx])).Select(r => r.Field<string>("Process_Name")).ToList();
                                        break;
                                    case "OUTERLAYERVIA[PLATING]":
                                        pro_inx = 2;// "Copper_Plating_Process";
                                        cur_itemLst = Spec_addr_tbl.AsEnumerable().Where(r => r.Field<string>("Col_Name").Contains("OUTER") && r.Field<string>("Process_Name").Equals(process_arr[pro_inx])).Select(r => r.Field<string>("Report_Name")).ToList();
                                        cur_ColName_List = Spec_addr_tbl.AsEnumerable().Where(r => r.Field<string>("Col_Name").Contains("OUTER") && r.Field<string>("Process_Name").Equals(process_arr[pro_inx])).Select(r => r.Field<string>("Col_Name")).ToList();
                                        cur_Process_name_lst = Spec_addr_tbl.AsEnumerable().Where(r => r.Field<string>("Col_Name").Contains("OUTER") && r.Field<string>("Process_Name").Equals(process_arr[pro_inx])).Select(r => r.Field<string>("Process_Name")).ToList();
                                        break;
                                    case "CVL":
                                        pro_inx = 3;// "Cover_Lay_Process";
                                        cur_itemLst = itemLst_arr[pro_inx];
                                        cur_ColName_List = ColName_List_arr[pro_inx];
                                        cur_Process_name_lst = Process_name_lst_arr[pro_inx];
                                        break;
                                    case "LPI":
                                        pro_inx = 4;// "Printing_Process";
                                        cur_itemLst = itemLst_arr[pro_inx];
                                        cur_ColName_List = ColName_List_arr[pro_inx];
                                        cur_Process_name_lst = Process_name_lst_arr[pro_inx];
                                        break;
                                    case "ENIG":
                                        pro_inx = 5;// "AU Plating Process";
                                        cur_itemLst = itemLst_arr[pro_inx];
                                        cur_ColName_List = ColName_List_arr[pro_inx];
                                        cur_Process_name_lst = Process_name_lst_arr[pro_inx];
                                        break;
                                }
                            }
                            int sel_inx = TDMK_Code.check_exist_list_index(rgn_val, cur_itemLst);
                            if (sel_inx != -1)
                            {
                                string[] item_vals = new string[7];
                                item_vals[1] = tar_ItemCode;
                                item_vals[0] = (TDMK_Code.SQL_MAX("SpecList", "ID", tar_sqlcon) + 1).ToString();
                                item_vals[2] = cur_itemLst[sel_inx];
                                item_vals[3] = spec_val;
                                item_vals[4] = spec_rgn.Offset[1, inx].AddressLocal.ToString();
                                item_vals[5] = cur_Process_name_lst[sel_inx];
                                item_vals[6] = cur_ColName_List[sel_inx];
                                TDMK_Code.insert_val_arr("SpecList", tar_sqlcon, items, item_vals);
                                cur_itemLst.RemoveAt(sel_inx);
                                cur_ColName_List.RemoveAt(sel_inx);
                                cur_Process_name_lst.RemoveAt(sel_inx);
                                arr_inx++;
                            }
                            inx++;
                        }
                        else
                        {
                            MessageBox.Show(new Form { TopMost = true },"Invalid format structure", "Warning");
                            cur_wrk.Close();
                            return;
                        }

                    }
                    MessageBox.Show(new Form { TopMost = true },"Complete setup IPQC Spec for ItemCode: " + tar_ItemCode, "Warning");
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true },"Format not found!", "Warning");
                }
            }
            else
            {
                if (MessageBox.Show(new Form { TopMost = true },"ItemCode " + tar_ItemCode + " is created. Do you want to update it?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    TDMK_Code.Delelte_FilteredItem_arr("SpecList", myVar.sqlcon_IPQC, "ItemCode = '" + txtItemCode.Text + "'");
                    goto start_label;
                }
            }
        }

        private void btnRUN_Click(object sender, EventArgs e)
        {
           
            string format_type = "MASS";
            if(rbNPI.Checked)
            {
                format_type = "NPI";
            }
            string file_loc = Path.Combine(txtFormat.Text, format_type);
            DGV_Spec.Columns.Clear();
            DGV_Spec.DataSource = null;
            switch (program_name)
            {
                case "FAI":
                    //DGV_Spec.DataSource= myCode.Load_Spec_fromFile(myVar.sqlcon_SMT,file_loc, txtItemCode.Text, "", new List<string>() { "*.xlsx", "*.xlsm" },format_type);
                    Load_Spec2(file_loc, txtItemCode.Text, "", new List<string>() { "*.xlsx", "*.xlsm" }, DGV_Spec);
                    break;
                //case "IPQC":
                //    Setup_IPQC_Spec(file_loc, txtItemCode.Text,myVar. sqlcon_IPQC);
                //    DGV_Spec.DataSource= TDMK_Code.Datatable_Filter(myVar.sqlcon_IPQC, "SpecList", "ItemCode = '" + txtItemCode.Text + "'");
                //    foreach (DataGridViewColumn c in DGV_Spec.Columns)
                //    {
                //        if (c.Name != "Spec")
                //        {
                //            foreach (DataGridViewRow r in DGV_Spec.Rows)
                //            {
                //                r.Cells[c.Index].ReadOnly = true;
                //            }
                //        }
                //    }
                //    DGV_Spec.AutoResizeColumns();
                //    myCode.Disable_Sort_DGV(DGV_Spec);
                //    break;
                case "Format_Structure":
                    Load_SheetName2(file_loc, txtItemCode.Text, ".xlsm");
                    break;
                case "Remove Links":
                    Clean_Format(file_loc,txtItemCode.Text,"xlsm");
                    break;
            }
        }

        public void Clean_Format(string file_loc, string tar_ItemCode, string f_ext)
        {
            string[] file_format_lst = Directory.GetFiles(file_loc, "*" + f_ext).Where(s => s.Contains(tar_ItemCode)).ToArray();
            string tar_format_file;
            if (file_format_lst.Length > 0)
            {
                tar_format_file = file_format_lst[0];
                myExcel.Workbook src_wrkbook = TDMK_Code.open_excel_file(tar_format_file, "", "");
                myCode2.break_links(src_wrkbook);
                var ranges = src_wrkbook.Names;
                int leftoveritems;
                leftoveritems = ranges.Count;
                int error_count = 0;
                while (leftoveritems - error_count > 0)
                {
                    int i = 1;
                    try
                    {
                        while (i <= leftoveritems)
                        {
                            var currentName = ranges.Item(i, Type.Missing, Type.Missing);
                            if (currentName.Name != "SheetNames")
                            {
                                currentName.Delete();
                            }
                            else
                            {
                                error_count++;
                            }
                            i++;
                        }
                    }
                    catch (Exception ex)
                    {
                        error_count++;
                    }
                    ranges = src_wrkbook.Names;
                    leftoveritems = ranges.Count;
                }
                MessageBox.Show(new Form { TopMost = true },"Finished");
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },"Format not found!", "Warning");
            }
        }
        public bool Load_Spec(string format_loc, string tar_ItemCode, string tar_LotNo, List<string> extensions, DataGridView tar_DGV_Spec)
        {
        start_label: bool _result = false;
            string tar_format_file;// = Path.Combine(format_loc, tar_format + file_extension);         
            AutoCompleteStringCollection FAI_SheetNo_list = new AutoCompleteStringCollection();
            string format_type = "NPI";
            if(rbMASS.Checked)
            {
                format_type = "MASS";
            }
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { txtItemCode.Text, format_type });
            //FAI_SheetNo_list = TDMK_Code.Load_Item_Names_Filter(myVar.sqlcon_SMT, "FAI_Spec", "SheetNo", "ItemCode = '" + txtItemCode.Text + "'");
            FAI_SheetNo_list = TDMK_Code.Load_Item_Names_Filter(myVar.sqlcon_SMT, "FAI_Spec", "SheetNo",filter_str);
            if (FAI_SheetNo_list.Count == 0)
            {
                string[] spec_item_str = new string[14];
                spec_item_str[0] = "NormDim";
                spec_item_str[1] = "Tol_Max";
                spec_item_str[2] = "Tol_Min";
                spec_item_str[3] = "Distribution";
                spec_item_str[4] = "Instrument";
                spec_item_str[5] = "STDEV";
                spec_item_str[6] = "Mean";
                spec_item_str[7] = "MAX";
                spec_item_str[8] = "MIN";
                spec_item_str[9] = "Cp";
                spec_item_str[10] = "Cpkl";
                spec_item_str[11] = "Cpku";
                spec_item_str[12] = "Cpk";
                spec_item_str[13] = "Cpkm";

                List<string> file_format_lst = new List<string>();// = Directory.GetFiles(format_loc, "*" + file_extension).Where(s => s.Contains(tar_ItemCode)).ToArray();
                //List<string> extensions = new List<string>() { "*.xlsx", "*.xlsm" };
                DirectoryInfo directory = new DirectoryInfo(format_loc);
                var files = extensions.SelectMany(e => directory.EnumerateFiles(e, SearchOption.AllDirectories)).Where(x => x.FullName.Contains(tar_ItemCode));
                foreach (var item in files)
                {
                    file_format_lst.Add(item.FullName);
                }
                if (file_format_lst.Count > 0)
                {
                    tar_format_file = file_format_lst[0];
                    myExcel.Workbook tar_wkbook = TDMK_Code.open_excel_file(tar_format_file, "", "");
                    int wrksheet_num = tar_wkbook.Worksheets.Count;
                    int spec_col_inx = 0;
                    foreach (myExcel.Worksheet tg in tar_wkbook.Worksheets)
                    {
                        try
                        {
                            if ((tg.Name.Contains("FAI")) || (tg.Name.Contains("SPC")) || (tg.Name.Contains("CPK")) || (tg.Name.Contains("parentheses")))
                            {
                                string dim_no_addr = myCode2.Find_Cell_Addr("Dim. No.", "B10", tg, false);
                                string instrument_addr = myCode2.Find_Cell_Addr("instrument", "B10", tg, false);
                                string FAI_data_addr = myCode2.Find_Offset(tg.Range[dim_no_addr].Offset[0, 1].AddressLocal, tg, true, "");
                                int off_set = tg.Range[FAI_data_addr].Column - tg.Range[instrument_addr].Column;
                                myExcel.Range sel_rgn = tg.Range[dim_no_addr].Offset[0, off_set];// tg.Range["D19"];    //tg.Range["C17"];
                                myExcel.Range dev_rgn = tg.Range[instrument_addr].Offset[0, off_set];// tg.Range["D23"];    //tg.Range["C21"]
                                int sel_inx = 0;
                                while (myCode.checkDBNull(sel_rgn.Offset[0, sel_inx].Value) != "")
                                {
                                    string t_checkside = myCode.checkDBNull(sel_rgn.Offset[-1, sel_inx].Value);
                                    string t_FAIName = myCode.checkDBNull(sel_rgn.Offset[0, sel_inx].Value).Replace(" ", "");
                                    string t_FAI_Setval = myCode.checkDBNull(sel_rgn.Offset[1, sel_inx].Value);
                                    if (t_FAI_Setval == "")
                                    {
                                        t_FAI_Setval = "0";
                                    }
                                    string t_FAI_UL = myCode.checkDBNull(sel_rgn.Offset[2, sel_inx].Value);
                                    string t_FAI_LL = myCode.checkDBNull(sel_rgn.Offset[3, sel_inx].Value);
                                    string dgv_col_name = t_FAIName + "_" + t_FAI_Setval;
                                    string t_FAI_sheetno = tg.Name;
                                    string t_instrument = myCode.checkDBNull(dev_rgn.Offset[0, sel_inx].Value);
                                    _testFAI_spec.Add(new TDMK_OK2SHIP.FAI_Spec(dgv_col_name, t_checkside, t_FAI_Setval, t_FAI_UL, t_FAI_LL, t_FAI_sheetno, t_instrument));
                                    if (!myCode.check_columns_existed(myCode.DGV_To_Table(tar_DGV_Spec), dgv_col_name))
                                    {
                                        tar_DGV_Spec.Columns.Add(dgv_col_name, t_FAIName);

                                        if (tar_DGV_Spec.Rows.Count == 0)
                                        {
                                            for (int i = 0; i < 5; i++)
                                            {
                                                tar_DGV_Spec.Rows.Add();
                                                tar_DGV_Spec.Rows[i].HeaderCell.Value = spec_item_str[i];
                                            }
                                        }
                                        tar_DGV_Spec.Rows[0].Cells[dgv_col_name].Value = t_FAI_Setval; // Set val at Row =0
                                        tar_DGV_Spec.Rows[1].Cells[dgv_col_name].Value = t_FAI_UL; // UL at Row =1
                                        tar_DGV_Spec.Rows[2].Cells[dgv_col_name].Value = t_FAI_LL; // LL at Row =2
                                        tar_DGV_Spec.Rows[3].Cells[dgv_col_name].Value = t_checkside; // Check Side at Row =3
                                        tar_DGV_Spec.Rows[4].Cells[dgv_col_name].Value = t_instrument; // Instrument at Row =4
                                        spec_col_inx++;
                                    }
                                    sel_inx++;
                                }
                            }
                        }
                        catch
                        {
                            continue;
                        }
                        
                    }
                    tar_DGV_Spec.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                    myCode.Disable_Sort_DGV(tar_DGV_Spec);
                    tar_DGV_Spec.AutoResizeColumns();
                    tar_wkbook.Close();
                    string data_type = "NPI";
                    if(rbMASS.Checked)
                    {
                        data_type = "MASS";
                    }
                    myCode.save_FAI_spec(_testFAI_spec.ToArray(), txtItemCode.Text, myVar.sqlcon_SMT, data_type);
                    MessageBox.Show(new Form { TopMost = true },"Completed set up Spec  for ItemCode: " + txtItemCode.Text + "!");
                    _result = true;
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true },"File format not found");
                    _result = false;
                }
            }
            else
            {
                if (MessageBox.Show(new Form { TopMost = true },"Spec of ItemCode: " + txtItemCode.Text + " is existed. Do you want to update it?", "Information", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    //TDMK_Code.Delelte_FilteredItem_arr("FAI_Spec", myVar.sqlcon_SMT, "ItemCode ='" + txtItemCode.Text + "'");
                    TDMK_Code.Delelte_FilteredItem_arr("FAI_Spec", myVar.sqlcon_SMT, filter_str);
                    goto start_label;
                }
                else
                {
                    //DGV_Spec.DataSource = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "FAI_Spec", "ItemCode = '" + txtItemCode.Text + "'");
                    DGV_Spec.DataSource = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "FAI_Spec", filter_str);
                    DGV_Spec.AutoResizeColumns();
                    myCode.Disable_Sort_DGV(DGV_Spec);
                }
            }
            return _result;
        }
        public bool Load_Spec2(string format_loc, string tar_ItemCode, string tar_LotNo, List<string> extensions, DataGridView tar_DGV_Spec)
        {
            start_label: bool _result = false;
            string format_type = "NPI";
            if (rbMASS.Checked)
            {
                format_type = "MASS";
            }
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { txtItemCode.Text, format_type });
            DataTable spec_dt = myCode.Load_FAI_Spec_ToTable(myVar.sqlcon_SMT, tar_ItemCode, format_type);
            if(spec_dt.Rows.Count==0)
            {
                DataTable cur_dt = myCode.Load_Spec_fromFile(myVar.sqlcon_SMT, format_loc, tar_ItemCode, tar_LotNo, extensions, format_type);
                tar_DGV_Spec.DataSource = cur_dt;
                myCode.Save_FAI_Spec(cur_dt, tar_ItemCode, myVar.sqlcon_SMT, format_type);
            }
            else
            {
                if (MessageBox.Show(new Form { TopMost = true }, "Spec của ItemCode: " + txtItemCode.Text + " đã được cài đặt. Bạn muốn cập nhật?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    TDMK_Code.Delelte_FilteredItem_arr("FAI_Spec", myVar.sqlcon_SMT, filter_str);
                    goto start_label;
                }
                else
                {
                    DGV_Spec.DataSource = spec_dt;
                    DGV_Spec.AutoResizeColumns();
                    myCode.Disable_Sort_DGV(DGV_Spec);
                }
            }
            return _result;
        }
        public DataTable Load_Spec_fromFile(string format_loc, string tar_ItemCode, string tar_LotNo, List<string> extensions)
        {
            //start_label: bool _result = false;
            DataTable spec_dt = new DataTable();
            string tar_format_file;// = Path.Combine(format_loc, tar_format + file_extension);         
            AutoCompleteStringCollection FAI_SheetNo_list = new AutoCompleteStringCollection();
            string format_type = "NPI";
            if (rbMASS.Checked)
            {
                format_type = "MASS";
            }
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { txtItemCode.Text, format_type });
            FAI_SheetNo_list = TDMK_Code.Load_Item_Names_Filter(myVar.sqlcon_SMT, "FAI_Spec", "SheetNo", filter_str);
            if (FAI_SheetNo_list.Count == 0)
            {
                List<string> file_format_lst = new List<string>();// = Directory.GetFiles(format_loc, "*" + file_extension).Where(s => s.Contains(tar_ItemCode)).ToArray();
                DirectoryInfo directory = new DirectoryInfo(format_loc);
                var files = extensions.SelectMany(e => directory.EnumerateFiles(e, SearchOption.AllDirectories)).Where(x => x.FullName.Contains(tar_ItemCode));
                foreach (var item in files)
                {
                    file_format_lst.Add(item.FullName);
                }
                if (file_format_lst.Count > 0)
                {
                    tar_format_file = file_format_lst[0];
                    myExcel.Workbook tar_wkbook = TDMK_Code.open_excel_file(tar_format_file, "", "");
                    int wrksheet_num = tar_wkbook.Worksheets.Count;
                    foreach (myExcel.Worksheet tg in tar_wkbook.Worksheets)
                    {
                        try
                        {
                            if ((tg.Name.Contains("FAI")) || (tg.Name.Contains("SPC")) || (tg.Name.Contains("CPK")) || (tg.Name.Contains("parentheses")))
                            {
                                string dim_no_addr = myCode2.Find_Cell_Addr("Dim. No.", "B10", tg, false);
                                string instrument_addr = myCode2.Find_Cell_Addr("instrument", "B10", tg, false);
                                string FAI_data_addr = myCode2.Find_Offset(tg.Range[dim_no_addr].Offset[0, 1].AddressLocal, tg, true, "");
                                int off_set = tg.Range[FAI_data_addr].Column - tg.Range[instrument_addr].Column;
                                myExcel.Range sel_rgn = tg.Range[dim_no_addr].Offset[0, off_set];// tg.Range["D19"];    //tg.Range["C17"];
                                myExcel.Range dev_rgn = tg.Range[instrument_addr].Offset[0, off_set];// tg.Range["D23"];    //tg.Range["C21"]
                                int sel_inx = 0;
                                while (myCode.checkDBNull(sel_rgn.Offset[0, sel_inx].Value) != "")
                                {
                                    string t_checkside = myCode.checkDBNull(sel_rgn.Offset[-1, sel_inx].Value);
                                    string t_FAIName = myCode.checkDBNull(sel_rgn.Offset[0, sel_inx].Value).Replace(" ", "");
                                    string t_FAI_Setval = myCode.checkDBNull(sel_rgn.Offset[1, sel_inx].Value);
                                    if (t_FAI_Setval == "")
                                    {
                                        t_FAI_Setval = "NA";
                                    }
                                    string t_FAI_UL = myCode.checkDBNull(sel_rgn.Offset[2, sel_inx].Value);
                                    string t_FAI_LL = myCode.checkDBNull(sel_rgn.Offset[3, sel_inx].Value);
                                    string col_name = t_FAIName + "_" + t_FAI_Setval;
                                    string t_FAI_sheetno = tg.Name;
                                    string t_instrument = myCode.checkDBNull(dev_rgn.Offset[0, sel_inx].Value);
                                    string USL = myCode.checkDBNull(sel_rgn.Offset[6, sel_inx].Value);
                                    string LSL = myCode.checkDBNull(sel_rgn.Offset[7, sel_inx].Value);
                                    if (!myCode.check_columns_existed(spec_dt,col_name))
                                    {
                                        spec_dt.Columns.Add(col_name);
                                        if(spec_dt.Rows.Count==0)
                                        {
                                            for(int i=0;i<7;i++)
                                            {
                                                spec_dt.Rows.Add();
                                            }
                                        }
                                        spec_dt.Rows[0][col_name] = t_FAI_Setval; // Set val at Row =0
                                        spec_dt.Rows[1][col_name] = t_FAI_UL; // UL at Row =1
                                        spec_dt.Rows[2][col_name] = t_FAI_LL; // LL at Row =2
                                        spec_dt.Rows[3][col_name] = t_checkside; // Check Side at Row =3
                                        spec_dt.Rows[4][col_name] = t_instrument; // Instrument at Row =4
                                        spec_dt.Rows[5][col_name] = USL; // Check Side at Row =3
                                        spec_dt.Rows[6][col_name] = LSL; // Instrument at Row =4
                                    }
                                }
                            }
                        }
                        catch
                        {
                            continue;
                        }

                    }
                    tar_wkbook.Close();
                    MessageBox.Show(new Form { TopMost = true }, "Completed set up Spec  for ItemCode: " + txtItemCode.Text + "!");
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "File format not found");
                }
            }
            //else
            //{
            //    if (MessageBox.Show(new Form { TopMost = true }, "Spec of ItemCode: " + txtItemCode.Text + " is existed. Do you want to update it?", "Information", MessageBoxButtons.YesNo) == DialogResult.Yes)
            //    {
            //        TDMK_Code.Delelte_FilteredItem_arr("FAI_Spec", myVar.sqlcon_SMT, filter_str);
            //        goto start_label;
            //    }
            //    else
            //    {
            //        DGV_Spec.DataSource = myCode.Load_Spec_Table(myVar.sqlcon_SMT, tar_ItemCode);//TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "FAI_Spec", filter_str);
            //        DGV_Spec.AutoResizeColumns();
            //        myCode.Disable_Sort_DGV(DGV_Spec);
            //    }
            //}
            return spec_dt;
        }


        private void btnBrowse_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog myfolder = new FolderBrowserDialog();
            if (myfolder.ShowDialog() == DialogResult.OK)
            {
                txtFormat.Text = myfolder.SelectedPath;
            }
        }

        private void FormatSetup_Load(object sender, EventArgs e)
        {
            RB_FAI.Checked = true;
            program_name = "FAI";
            myVar.sqlcon_SMT= myCode2.initial_data("OK2SHIP_SMT", true);
            txtFormat.Text = myVar.format_loc;
        }

        private void RBBreakLinks_CheckedChanged(object sender, EventArgs e)
        {
            if(RBBreakLinks.Checked)
            {
                program_name = "Remove Links";
                DGV_Spec.DataSource = null;
                DGV_Spec.Columns.Clear();
            }
        }

        private void DGV_Spec_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        public string Find_Cell_Addr(string search_key, string start_addr, myExcel.Worksheet tar_wrksht)
        {
            string result = "";
            myExcel.Range cycle_rgn = tar_wrksht.Range[start_addr];
            for (int i = 0; i < 100; i++)
            {
                myExcel.Range sel_rgn = cycle_rgn.Offset[i, 0];
                string sel_rgn_val = myCode.checkDBNull(sel_rgn.Value);
                if (sel_rgn_val.Replace(" ", "").ToUpper() == search_key.Replace(" ", "").ToUpper())
                {
                    result = sel_rgn.AddressLocal;
                    break;
                }
            }
            return result;
        }
    }
}
