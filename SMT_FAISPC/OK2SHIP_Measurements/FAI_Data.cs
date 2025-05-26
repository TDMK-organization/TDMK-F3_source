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

namespace OK2SHIP_Software
{
    public partial class FAI_Data : Form
    {
        public string ItemCode { get; set; }
        public string Lotno { get; set; }
        public string format_type { get; set; }
        public string shift { get; set; }

        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        TDMK_OK2SHIP myCode = new TDMK_OK2SHIP();
        //List<TDMK_OK2SHIP.FAI_Spec> testFAI_spec = new List<TDMK_OK2SHIP.FAI_Spec>();
        //TDMK_OK2SHIP.FAI_Spec[] testFAI_spec = new TDMK_OK2SHIP.FAI_Spec[1000];
        public FAI_Data()
        {
            InitializeComponent();
        }
        public FAI_Data(string _itemcode, string _lotno, string _format_type, string _shift)
        {
            ItemCode = _itemcode;
            Lotno = _lotno;
            format_type = _format_type;
            shift = _shift;
            InitializeComponent();
        }

        private void FAI_Data_Load(object sender, EventArgs e)
        {
            if(myCode.Load_Spec(myVar.sqlcon_SMT, "", ItemCode, Lotno, ".xlsm", DGV_Spec, "",format_type))
            {
                myCode.Load_FAI_DGV2(DGV_Data, myVar.sqlcon_SMT, "FAI_Auto", "FAI_No", "FAI_Data", ItemCode, Lotno, format_type,shift);
                //DGV_Data.DataSource = myCode.Load_FAI_ToTable(myVar.sqlcon_SMT, ItemCode, Lotno, format_type);
                myCode.Load_CPK_DGV_Histogram(DGV_Data, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
                myCode.Check_Data_inSpec2(DGV_Data, DGV_Spec);
                for(int r=0;r<DGV_Data.Rows.Count;r++)
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
    }
}
