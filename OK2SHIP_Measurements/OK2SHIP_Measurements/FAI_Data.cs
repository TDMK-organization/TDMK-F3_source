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
using System.Data.SqlClient;
using OK2SHIP_Lib;

namespace OK2SHIP_Measurements
{
    public partial class FAI_Data : Form
    {
        public string ItemCode { get; set; }
        public string Lotno { get; set; }
        public string format_type { get; set; }
        public string shift { get; set; }
        public bool FAI_mode { get; set; }
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        TDMK_OK2SHIP myCode = new TDMK_OK2SHIP();
        //List<TDMK_OK2SHIP.FAI_Spec> testFAI_spec = new List<TDMK_OK2SHIP.FAI_Spec>();
        //TDMK_OK2SHIP.FAI_Spec[] testFAI_spec = new TDMK_OK2SHIP.FAI_Spec[1000];
        public FAI_Data()
        {
            InitializeComponent();
        }
        public FAI_Data(string _itemcode, string _lotno, string _format_type, string _shift, bool _FAI_mode=true)
        {
            ItemCode = _itemcode;
            Lotno = _lotno;
            format_type = _format_type;
            shift = _shift;
            FAI_mode = _FAI_mode;
            InitializeComponent();
        }

        private void FAI_Data_Load(object sender, EventArgs e)
        {
            if(FAI_mode)
            {
                if (myCode.Load_Spec(myVar.sqlcon_SMT, "", ItemCode, Lotno, ".xlsm", DGV_Spec, "", "")) //format_type.ToUpper()
                {
                    Load_FAI_DGV(DGV_Data, myVar.sqlcon_SMT, "FAI_Auto", "FAI_No", "FAI_Data", ItemCode, Lotno, "");//format_type.ToUpper()
                    myCode.Load_CPK_DGV_Histogram(DGV_Data, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
                    myCode.Check_Data_inSpec2(DGV_Data, DGV_Spec);
                    for (int r = 0; r < DGV_Data.Rows.Count; r++)
                    {
                        DGV_Data.Rows[r].HeaderCell.Value = (r + 1).ToString();
                    }
                    DGV_Data.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                }
                else
                {
                    MessageBox.Show("Không tìm thấy Spec", "Thông báo");
                }
            }
            else
            {
                if(myCode.Load_PTH_Diameter_Spec(ItemCode,myVar.sqlcon_SMT,"SPEC_PTH_Diameter",ref DGV_Spec))
                {
                    myCode.Load_PTH_Diameter_DGV(DGV_Data, myVar.sqlcon_SMT, "PTH_Diameter", ItemCode, Lotno);
                    myCode.Load_CPK_DGV_Histogram(DGV_Data, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
                    myCode.Check_Data_inSpec2(DGV_Data, DGV_Spec);
                    for (int r = 0; r < DGV_Data.Rows.Count; r++)
                    {
                        DGV_Data.Rows[r].HeaderCell.Value = (r + 1).ToString();
                    }
                    DGV_Data.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                }
            }

        }

        private void DGV_Data_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                List<string> col_list = new List<string>();
                foreach (DataGridViewCell c in DGV_Data.SelectedCells)
                {
                    int k = c.ColumnIndex;
                    if (!col_list.Contains(DGV_Data.Columns[k].Name))
                    {
                        col_list.Add(DGV_Data.Columns[k].Name);
                    }
                }

                DataTable test_tbl = myCode.DGV_To_Table(DGV_Data);
                int arr_num = col_list.Count;
                TDMK_OK2SHIP.Calculate_CPK[] myCalc_CPK = new TDMK_OK2SHIP.Calculate_CPK[arr_num];
                TDMK_OK2SHIP.FAI_Histogram_Data[] myHistogram_data = new TDMK_OK2SHIP.FAI_Histogram_Data[arr_num];
                int col_list_inx = 0;
                foreach (string t in col_list)
                {
                    TDMK_OK2SHIP.FAI_Spec sel_FAI_test;
                    foreach (TDMK_OK2SHIP.FAI_Spec ref_spec in myCode.testFAI_spec)
                    {
                        if (ref_spec.FAI_Name == t)
                        {
                            sel_FAI_test = ref_spec;
                            string[] temp_FAI_data = test_tbl.AsEnumerable().Select(r => r.Field<string>(t)).ToArray();
                            Double[] data_arr = new double[temp_FAI_data.Length];
                            int inx = 0;
                            foreach (string c in temp_FAI_data)
                            {
                                if (myCode.checkDBNull(c) != "")
                                {
                                    data_arr[inx] = Convert.ToDouble(c);
                                    inx++;
                                }
                            }
                            Array.Resize(ref data_arr, inx);
                            if(inx >0)
                            {
                                myCode.Calcul_CPK_FAI2(ref_spec, data_arr, ref myCalc_CPK[col_list_inx], ref myHistogram_data[col_list_inx]);
                                FrmHistogram chart_form = new FrmHistogram();
                                chart_form.src_bin_data = myHistogram_data[col_list_inx]._Bin_data;// bin_data;
                                chart_form.src_freq_bin_data = myHistogram_data[col_list_inx]._Freq_bin_data;//freq_bin_data;
                                chart_form.src_FAI_Data = myHistogram_data[col_list_inx]._FAI_Data;
                                chart_form.FAI_Spec_val = ref_spec;
                                chart_form.src_CPK_result = myCalc_CPK[col_list_inx];
                                chart_form.src_modified_NormDist = myHistogram_data[col_list_inx]._Modified_NormDist_data; //modified_NormDist;
                                chart_form.Show();
                            }
                            break;
                        }
                    }
                    col_list_inx++;
                }
            }
            catch
            {

            }
        }
        public void Load_FAI_DGV(DataGridView tar_DGV, SqlConnection tar_sqlcon, string tar_tbl, string FAI_No_name, string FAI_Data_Col_name, string tar_ItemCode, string tar_LotNo, string format_type)
        {
            DataTable FAI_spec_dt = TDMK_Code.Datatable_Filter(tar_sqlcon, "FAI_Spec", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { tar_ItemCode, format_type }));
            DataTable FAI_data_dt = TDMK_Code.Datatable_Filter(tar_sqlcon, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode","LotNo", "Remark" }, new string[] { tar_ItemCode,tar_LotNo, format_type }));
            List<string>temp_lst =  FAI_spec_dt.AsEnumerable().Select(x=>x.Field<string>("FAI_No")).Distinct().ToList();
            List<string> main_FAI = new List<string>();
            char[] split_chars = new char[] { '_','/','\\' };
            foreach(string t in temp_lst)
            {
                string setval = t.Split(split_chars).LastOrDefault().Trim();
                string fai_name = t.Split(split_chars).FirstOrDefault().Trim();
                string sel_FAI = fai_name + "_" + setval;
                if(main_FAI.IndexOf(sel_FAI) == -1)
                {
                    main_FAI.Add(sel_FAI);
                }
            }
            DataTable tbl_data = new DataTable();
            foreach (string c in main_FAI)
            {
                List<string> FAI_vals = FAI_data_dt.AsEnumerable().Where(x=>x.Field<string>(FAI_No_name)==c).Select(x=>x.Field<string>(FAI_Data_Col_name)).ToList();

                if (!myCode.check_columns_existed(tbl_data, c))
                {
                    tbl_data.Columns.Add(c);
                }
                int tbl_row_count = tbl_data.Rows.Count;
                if (tbl_row_count < FAI_vals.Count)
                {
                    for (int i = 0; i < FAI_vals.Count - tbl_row_count; i++)
                    {
                        tbl_data.Rows.Add();
                    }
                }
                int inx = 0;
                foreach (string t in FAI_vals)
                {
                    tbl_data.Rows[inx][c] = t;
                    inx++;
                }
            }
            tar_DGV.DataSource = tbl_data;
            foreach (DataGridViewColumn t in tar_DGV.Columns)
            {
                t.HeaderText = t.HeaderText.Split('_')[0];
            }
            myCode.Disable_Sort_DGV(tar_DGV);
        }
        public void Get_List_data2(int col_inx, DataTable myDt, string[] src_arr, ref List<string> src_lst_data, string tar_item, bool distinct_en)
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
                                Get_List_data2(col_inx + 1, curTbl, src_arr, ref src_lst_data, tar_item, distinct_en);
                            }
                        }
                    }
                    else
                    {
                        if (distinct_en)
                        {
                            //src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Distinct().ToArray());
                            src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Where(r => r != null).Distinct().ToArray());
                        }
                        else
                        {
                            //src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).ToArray());
                            src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Where(r => r != null).ToArray());
                        }
                    }
                }
            }
            else
            {
                if (distinct_en)
                {
                    //src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Distinct().ToArray());
                    src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Where(r => r != null).Distinct().ToArray());
                }
                else
                {
                    //src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).ToArray());
                    src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Where(r => r != null).ToArray());
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
        public void Load_PTH_Diameter_DGV(DataGridView tar_DGV, SqlConnection tar_sqlcon, string tar_tbl, string tar_ItemCode, string tar_LotNo)
        {
            DataTable PTH_Diameter_dt = TDMK_Code.Datatable_Filter(tar_sqlcon, tar_tbl, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { tar_ItemCode, tar_LotNo }));
            DataTable result = new DataTable();
            List<DataTable> normdim_dt_lst= new List<DataTable>();
            myCode.Get_ListTable(-1, PTH_Diameter_dt, new string[] { "Normdim" }, ref normdim_dt_lst, "Data");
            foreach(DataTable dt in normdim_dt_lst)
            {
                string setval = myCode.checkDBNull(dt.Rows[0]["Normdim"]);
                string col_name = "PTH-Diameter_" + setval;
                if(!result.Columns.Contains(col_name))
                {
                    result.Columns.Add(col_name);
                    if(result.Rows.Count<dt.Rows.Count)
                    {
                        int row_add = dt.Rows.Count-result.Rows.Count;
                        for(int i=0;i<row_add; i++)
                        {
                            result.Rows.Add();
                        }
                    }
                    for(int j=0;j<dt.Rows.Count;j++)
                    {
                        result.Rows[j][col_name] = dt.Rows[j]["Data"];
                    }
                }    
            }
            tar_DGV.DataSource = result;
            foreach (DataGridViewColumn t in tar_DGV.Columns)
            {
                t.HeaderText = t.HeaderText.Split('_')[0];
            }
            myCode.Disable_Sort_DGV(tar_DGV);
        }
    }
}
