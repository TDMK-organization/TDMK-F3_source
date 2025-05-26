using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZedGraph;
using OK2SHIP_Measurements;
using OK2SHIP_Lib;

namespace OK2SHIP_Measurements
{
    public partial class FrmHistogram : Form
    {
        public TDMK_OK2SHIP myCode = new TDMK_OK2SHIP();
        public FrmHistogram()
        {
            InitializeComponent();
        }
        public TDMK_OK2SHIP.FAI_Spec FAI_Spec_val { get; set; }
        public double[] src_bin_data { get; set; }
        public int[] src_freq_bin_data { get; set; }
        public double[] src_FAI_Data { get; set; }   
        public double [] src_modified_NormDist { get; set; }
        public TDMK_OK2SHIP.Calculate_CPK src_CPK_result { get; set; }
        private void Form2_Load(object sender, EventArgs e)
        {
            Draw_Histogram(zgHistogramFAI1, src_bin_data, src_freq_bin_data, src_modified_NormDist, FAI_Spec_val);
            Draw_Chart(zgChartFAI1, src_FAI_Data, FAI_Spec_val);
            Display_Result(src_CPK_result, FAI_Spec_val);
        }
        private string MyPointValueHandler(ZedGraphControl control, GraphPane pane, CurveItem curve, int iPt)
        {
            // Get the PointPair that is under the mouse
            PointPair pt = curve[iPt];

            return curve.Label.Text + " is " + pt.Y.ToString("f3") + " at " + pt.X.ToString("f3");
        }
        public void Draw_Histogram (ZedGraph.ZedGraphControl tar_Graph, double [] src_bin, int [] src_freq_bin, double [] src_modified, TDMK_OK2SHIP .FAI_Spec src_FAI_spec)
        {
            GraphPane myPane =tar_Graph.GraphPane;
            PointPairList NormDistList = new PointPairList();
            PointPairList FAIPointList = new PointPairList();
            PointPairList USL_list = new PointPairList();
            PointPairList LSL_list = new PointPairList();

            for (int i = 0; i < 50; i++)
            {
                double x = src_bin[i];
                double y = src_freq_bin[i];
                double y2 = src_modified[i];
                FAIPointList.Add(x, y);
                NormDistList.Add(x, y2);

            }
            double freq_bin_Min = src_freq_bin.Min();
            double freq_bin_Max = src_freq_bin.Max();
            double UL = 0;// = Convert.ToDouble(src_FAI_spec.SetVal) + Convert.ToDouble(src_FAI_spec.UL);
            double LL = 0;// = Convert.ToDouble(src_FAI_spec.SetVal) - Convert.ToDouble(src_FAI_spec.LL);
            LineItem UL_curve;//= myPane.AddCurve("USL", USL_list, Color.Red, SymbolType.None);
            LineItem LL_curve;// = myPane.AddCurve("LSL", LSL_list, Color.Blue, SymbolType.None);
            if (myCode.IsNumeric(src_FAI_spec.UL))
            {
                UL = Convert.ToDouble(src_FAI_spec.SetVal) + Convert.ToDouble(src_FAI_spec.UL);
                USL_list.Add(UL, freq_bin_Min);
                USL_list.Add(UL, freq_bin_Max);
                UL_curve = myPane.AddCurve("USL", USL_list, Color.Red, SymbolType.None);
            }
            if(myCode.IsNumeric(src_FAI_spec.LL))
            {
                LL = Convert.ToDouble(src_FAI_spec.SetVal) - Convert.ToDouble(src_FAI_spec.LL);
                LSL_list.Add(LL, freq_bin_Min);
                LSL_list.Add(LL, freq_bin_Max);
                LL_curve = myPane.AddCurve("LSL", LSL_list, Color.Blue, SymbolType.None);
            }
            BarItem FAI_Curve = myPane.AddBar(src_FAI_spec.FAI_Name, FAIPointList, Color.Gray);
            LineItem NormDist_Curve = myPane.AddCurve("NormDist", NormDistList, Color.DarkBlue, SymbolType.Square);
            NormDist_Curve.Line.Width = 3;
            tar_Graph.IsShowPointValues = true;
            tar_Graph.PointValueEvent += new ZedGraphControl.PointValueHandler(MyPointValueHandler);
            myPane.XAxis.Scale.Min = src_bin.Min();
            myPane.XAxis.Scale.Max = src_bin.Max();
            myPane.XAxis.Scale.MajorStep = (src_bin.Max() - src_bin.Min()) / 50;
            myPane.Title.Text = src_FAI_spec.FAI_Name + "_Histogram";
            tar_Graph.AxisChange();
            tar_Graph.Invalidate();
        }

        public void Draw_Chart(ZedGraph.ZedGraphControl tar_Graph, double[] src_FAI, TDMK_OK2SHIP.FAI_Spec src_FAI_spec)
        {
            GraphPane mychart = tar_Graph.GraphPane;
            PointPairList FAIDataList = new PointPairList();
            PointPairList UL_Chart = new PointPairList();
            PointPairList LL_Chart = new PointPairList();
            double UL = 0;// Convert.ToDouble(src_FAI_spec.SetVal)  + Convert.ToDouble(src_FAI_spec.UL);
            double LL = 0;// Convert.ToDouble(src_FAI_spec.SetVal) - Convert.ToDouble(src_FAI_spec.LL);
            bool UL_en = false;
            bool LL_en = false;
            if(myCode.IsNumeric(src_FAI_spec.UL))
            {
                UL = Convert.ToDouble(src_FAI_spec.SetVal) + Convert.ToDouble(src_FAI_spec.UL);
                UL_en = true;
            }
            else
            {
                UL_en = false;
            }
            if(myCode.IsNumeric(src_FAI_spec.LL))
            {
                LL = Convert.ToDouble(src_FAI_spec.SetVal) - Convert.ToDouble(src_FAI_spec.LL);
                LL_en = true;
            }
            else
            {
                LL_en = false;
            }
            for (int i = 0; i < src_FAI.Length; i++)
            {
                FAIDataList.Add(i, src_FAI[i]);
                if(UL_en)
                {
                    UL_Chart.Add(i, UL);
                }
                if(LL_en)
                {
                    LL_Chart.Add(i, LL);
                }                                
            }
            LineItem FAIData_line = mychart.AddCurve("FAI Data", FAIDataList, Color.Black);
            LineItem UL_line;// = mychart.AddCurve("UL", UL_Chart, Color.Blue, SymbolType.None);
            LineItem LL_line;// = mychart.AddCurve("LL", LL_Chart, Color.Red, SymbolType.None);
            if (UL_en)
            {
                UL_line = mychart.AddCurve("UL", UL_Chart, Color.Blue, SymbolType.None);
            }
            if(LL_en)
            {
                LL_line = mychart.AddCurve("LL", LL_Chart, Color.Red, SymbolType.None);
            }            
            tar_Graph.IsShowPointValues = true;
            tar_Graph.PointValueEvent += new ZedGraphControl.PointValueHandler(MyPointValueHandler);
            mychart.Title.Text = src_FAI_spec.FAI_Name + "_Trend";
            tar_Graph.AxisChange();
            // Make sure the Graph gets redrawn
            tar_Graph.Invalidate();
        }
        public void Display_Result(TDMK_OK2SHIP.Calculate_CPK src_CPK_result, TDMK_OK2SHIP.FAI_Spec src_FAI_spec)
        {
            double USL = 0;// Convert.ToDouble(src_FAI_spec.SetVal) + Convert.ToDouble(src_FAI_spec.UL);
            double LSL = 0;// Convert.ToDouble(src_FAI_spec.SetVal) - Convert.ToDouble(src_FAI_spec.LL);
            if(myCode.IsNumeric(src_FAI_spec.UL))
            {
                USL = Convert.ToDouble(src_FAI_spec.SetVal) + Convert.ToDouble(src_FAI_spec.UL);
                txtUL.Text = USL.ToString();//src_FAI_spec.UL;
            }
            else
            {
                txtUL.Text = "NA";
            }
            if(myCode.IsNumeric(src_FAI_spec.LL))
            {
                LSL = Convert.ToDouble(src_FAI_spec.SetVal) - Convert.ToDouble(src_FAI_spec.LL);
                txtLL.Text = LSL.ToString(); //src_FAI_spec.LL;
            }
            else
            {
                txtLL.Text = "NA";
            }
            lblFAI.Text = "CPK Result: " + src_FAI_spec.FAI_Name;
            txtSideCheck.Text = src_FAI_spec.check_side;
            txtNorminalDim.Text = src_FAI_spec.SetVal;
            //txtUL.Text = USL.ToString();//src_FAI_spec.UL;
            //txtLL.Text = LSL.ToString(); //src_FAI_spec.LL;

            txtSTDEV.Text = src_CPK_result.stdev.ToString ();
            txtMean.Text = src_CPK_result.mean.ToString();
            txtMax.Text = src_CPK_result.max.ToString();
            txtMin.Text = src_CPK_result.min.ToString();

            txtCP.Text = myCode.checkDBNull(src_CPK_result.CP);//.ToString ();
            txtCPKL.Text = myCode.checkDBNull(src_CPK_result.CPKL);//.ToString();
            txtCPKU.Text = myCode.checkDBNull(src_CPK_result.CPKU);//.ToString();
            txtCPK.Text = myCode.checkDBNull(src_CPK_result.CPK);//.ToString();
            txtCPKM.Text = myCode.checkDBNull(src_CPK_result.CPKM);//.ToString();

            this.Text = src_FAI_spec.FAI_Name + "_Histogram";
            tabChart.TabPages[0].Text = src_FAI_spec.FAI_Name + "_Histogram";
            tabChart.TabPages[1].Text = src_FAI_spec.FAI_Name + "_Calculate CPK";

        }
    }
}
