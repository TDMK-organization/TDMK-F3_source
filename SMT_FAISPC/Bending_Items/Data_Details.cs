using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Bending_Items;
using VHX;
using OK2SHIP;
using TDMK_SEEV_DLL;
using TDMK_SQL;
using ZedGraph;

namespace OK2SHIP
{
    public partial class Data_Details : Form
    {
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public SEI_Lib myCode = new SEI_Lib();
        public ECheck_Process proc_data = new ECheck_Process();
        public DataTable src_data { get; set; }
        public string ItemCode { get; set; }
        public string process_name { get; set; }
        public SqlConnection sqlcon { get; set; }
        DataTable _spec_dt;
        public Data_Details()
        {
            InitializeComponent();
        }
        public Data_Details(DataTable _src_data, string _ItemCode, string _process_name, SqlConnection _sqlcon)
        {
            InitializeComponent();
            src_data = _src_data;
            ItemCode = _ItemCode;
            process_name = _process_name;
            sqlcon = _sqlcon;
        }
        private void Electrical_Details_Load(object sender, EventArgs e)
        {
            DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon, "NET_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { ItemCode, process_name }));
            _spec_dt = spec_dt.AsDataView().ToTable(false, new string[] { "Point+V", "Point-V", "LSL", "USL" });
            proc_data.Electrical_CPK_Details(src_data, DGV_CPK_NET,sqlcon, DGV_NG_NET, ItemCode, process_name);            
        }

        private void DGV_CPK_NET_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {            
            int r_inx = e.RowIndex;
            DataRow dr = src_data.Rows[r_inx];
            List<string> cur_Net_lst = new List<string>();
            foreach (DataColumn dc in src_data.Columns)
            {
                cur_Net_lst.Add(dr[dc].ToString());
            }
            IEnumerable<double> cur_net_val = Calcu_process.ConvertToDouble(cur_Net_lst);
            SEI_Lib.FAI_Histogram_Data temp_hist = new SEI_Lib.FAI_Histogram_Data();
            SEI_Lib.Calculate_CPK temp_CPK = new SEI_Lib.Calculate_CPK();

            string USL = (cur_net_val.Max()*1.1).ToString();// _spec_dt.Rows[r_inx][3].ToString();
            string LSL = (cur_net_val.Min() * 0.9).ToString(); ;// _spec_dt.Rows[r_inx][2].ToString();



            string NET_name = "NET No_" + (r_inx + 1).ToString();
            proc_data.Calcul_CPK(cur_net_val.ToArray(), ref temp_CPK, ref temp_hist, USL, LSL, 3, 50);
            Histogram_Draw(zedHisto, temp_hist._Bin_data, temp_hist._Freq_bin_data, temp_hist._Modified_NormDist_data,  USL, LSL, NET_name);
            Chart_Draw(zedChart, cur_net_val.ToArray(), USL, LSL, NET_name);
            txtSTDEV.Text = temp_CPK.stdev.ToString("#0.00#");
            txtMean.Text = temp_CPK.mean.ToString("#0.00#");
            txtMax.Text = temp_CPK.max.ToString("#0.00#");
            txtMin.Text = temp_CPK.min.ToString("#0.00#");
            txtCPK.Text = temp_CPK.CPK.ToString("#0.00#");
            txtUL.Text = USL;
            txtLL.Text = LSL;
        }
        public void Histogram_Draw(ZedGraph.ZedGraphControl tar_Graph, double[] src_bin, int[] src_freq_bin, double[] src_modified, string USL, string LSL, string NET_name)
        {
            GraphPane myPane = tar_Graph.GraphPane;
            myPane.CurveList.Clear();
            PointPairList NormDistList = new PointPairList();
            PointPairList FAIPointList = new PointPairList();
            PointPairList USL_list = new PointPairList();
            PointPairList LSL_list = new PointPairList();

            for (int i = 0; i < src_freq_bin.Length; i++)
            {
                double x = src_bin[i];
                double y = src_freq_bin[i];
                double y2 = src_modified[i];
                FAIPointList.Add(x, y);
                NormDistList.Add(x, y2);
            }
            double freq_bin_Min = src_freq_bin.Min();
            double freq_bin_Max = src_freq_bin.Max();
            if (myCode.IsNumeric_Val(USL) != "")
            {
                double UL = Convert.ToDouble(USL);
                USL_list.Add(UL, freq_bin_Min);
                USL_list.Add(UL, freq_bin_Max);
                LineItem UL_curve = myPane.AddCurve("USL", USL_list, Color.Red, SymbolType.None);
            }
            if (myCode.IsNumeric_Val(LSL) != "")
            {
                double LL = Convert.ToDouble(LSL);
                LSL_list.Add(LL, freq_bin_Min);
                LSL_list.Add(LL, freq_bin_Max);
                LineItem LL_curve = myPane.AddCurve("LSL", LSL_list, Color.Blue, SymbolType.None);
            }
            BarItem FAI_Curve = myPane.AddBar(NET_name, FAIPointList, Color.Gray);
            LineItem NormDist_Curve = myPane.AddCurve("NormDist", NormDistList, Color.DarkBlue, SymbolType.Square);
            NormDist_Curve.Line.Width = 3;
            tar_Graph.IsShowPointValues = true;
            tar_Graph.PointValueEvent += new ZedGraphControl.PointValueHandler(MyPointValueHandler);
            myPane.XAxis.Scale.Min = src_bin.Min();
            myPane.XAxis.Scale.Max = src_bin.Max();
            myPane.XAxis.Scale.MajorStep = (src_bin.Max() - src_bin.Min()) / 50;
            myPane.Title.Text = NET_name + "_Histogram";
            tar_Graph.AxisChange();
            tar_Graph.Invalidate();
        }

        public void Chart_Draw(ZedGraph.ZedGraphControl tar_Graph, double[] src_FAI, string USL, string LSL, string NET_name)
        {
            GraphPane mychart = tar_Graph.GraphPane;
            mychart.CurveList.Clear();
            PointPairList FAIDataList = new PointPairList();
            PointPairList UL_Chart = new PointPairList();
            PointPairList LL_Chart = new PointPairList();
            if (myCode.IsNumeric_Val(USL) != "")
            {
                double UL = Convert.ToDouble(USL);
                for (int i = 0; i < src_FAI.Length; i++)
                {
                    UL_Chart.Add(i, UL);
                }
                LineItem UL_line = mychart.AddCurve("UL", UL_Chart, Color.Blue, SymbolType.None);
            }
            if (myCode.IsNumeric_Val(LSL) != "")
            {
                double LL = Convert.ToDouble(LSL);
                for (int i = 0; i < src_FAI.Length; i++)
                {
                    LL_Chart.Add(i, LL);
                }
                LineItem LL_line = mychart.AddCurve("LL", LL_Chart, Color.Red, SymbolType.None);
            }
            for (int i = 0; i < src_FAI.Length; i++)
            {
                FAIDataList.Add(i, src_FAI[i]);
                //UL_Chart.Add(i, UL);
                //LL_Chart.Add(i, LL);
            }
            LineItem FAIData_line = mychart.AddCurve("NET Data", FAIDataList, Color.Black);
            tar_Graph.IsShowPointValues = true;
            tar_Graph.PointValueEvent += new ZedGraphControl.PointValueHandler(MyPointValueHandler);
            mychart.Title.Text = NET_name + "_Trend";
            tar_Graph.AxisChange();
            // Make sure the Graph gets redrawn
            tar_Graph.Invalidate();
        }
        private string MyPointValueHandler(ZedGraphControl control, GraphPane pane, CurveItem curve, int iPt)
        {
            // Get the PointPair that is under the mouse
            PointPair pt = curve[iPt];

            return curve.Label.Text + " is " + pt.Y.ToString("f3") + " at " + pt.X.ToString("f3");
        }
    }
}
