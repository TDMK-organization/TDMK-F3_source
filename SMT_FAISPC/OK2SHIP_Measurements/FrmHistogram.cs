using OK2SHIP_Lib;
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
using OK2SHIP_Lib;

namespace OK2SHIP_Measurements
{
    public partial class FrmHistogram : Form
    {
        TDMK_OK2SHIP myLib = new TDMK_OK2SHIP();
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
        public TDMK_OK2SHIP myCode = new TDMK_OK2SHIP();
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
            string curve_name = curve.Label.Text;
            if(curve_name.Contains("NormDist"))
            {
                return curve_name + " is " + pt.Y.ToString("f3") + " at areas of " + pt.X.ToString("f3");
            }
            else
            {
                if(!curve_name.Contains("FAI"))
                {
                    curve_name = "FAI-" + curve_name.Split('_')[0];
                }
                return curve_name + ": total pcs is " + pt.Y.ToString() + " at areas of " + pt.X.ToString("f3");
            }    
        }
        private string FAIPointValueHandler(ZedGraphControl control, GraphPane pane, CurveItem curve, int iPt)
        {
            // Get the PointPair that is under the mouse
            PointPair pt = curve[iPt];
            string curve_name = pane.Title.Text.Split('_')[0];
            if(!curve_name.Contains("FAI"))
            {
                curve_name = "FAI-" + curve_name;
            }    
            return curve_name + ": No" + pt.X.ToString() + " is " + pt.Y.ToString("f3");
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
            if(myCode.IsNumeric_Val(src_FAI_spec.UL)!="")
            {
                double UL =Convert.ToDouble(src_FAI_spec.UL);//Convert.ToDouble(src_FAI_spec.SetVal) + Convert.ToDouble(src_FAI_spec.UL);
                USL_list.Add(UL, freq_bin_Min);
                USL_list.Add(UL, freq_bin_Max);
                LineItem UL_curve = myPane.AddCurve("USL", USL_list, Color.Red, SymbolType.None);
            }
            if(myCode.IsNumeric_Val(src_FAI_spec.LL)!="")
            {
                double LL = Convert.ToDouble(src_FAI_spec.LL);//Convert.ToDouble(src_FAI_spec.SetVal) - Convert.ToDouble(src_FAI_spec.LL);
                LSL_list.Add(LL, freq_bin_Min);
                LSL_list.Add(LL, freq_bin_Max);
                LineItem LL_curve = myPane.AddCurve("LSL", LSL_list, Color.Blue, SymbolType.None);
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
            if (myCode.IsNumeric_Val(src_FAI_spec.UL) != "")
            {
                double UL = Convert.ToDouble(src_FAI_spec.UL);//Convert.ToDouble(src_FAI_spec.SetVal) + Convert.ToDouble(src_FAI_spec.UL);
                for (int i = 0; i < src_FAI.Length; i++)
                {
                    UL_Chart.Add(i+1, UL);
                }
                LineItem UL_line = mychart.AddCurve("UL", UL_Chart, Color.Blue, SymbolType.None);
            }
            if(myCode.IsNumeric_Val(src_FAI_spec.LL)!="")
            {
                double LL = Convert.ToDouble(src_FAI_spec.LL);//Convert.ToDouble(src_FAI_spec.SetVal) - Convert.ToDouble(src_FAI_spec.LL);
                for (int i = 0; i < src_FAI.Length; i++)
                {
                    LL_Chart.Add(i+1, LL);
                }
                LineItem LL_line = mychart.AddCurve("LL", LL_Chart, Color.Red, SymbolType.None);
            }               
            for (int i = 0; i < src_FAI.Length; i++)
            {
                FAIDataList.Add(i+1, src_FAI[i]);
                //UL_Chart.Add(i, UL);
                //LL_Chart.Add(i, LL);
            }
            LineItem FAIData_line = mychart.AddCurve("FAI Data", FAIDataList, Color.Black);                 
            tar_Graph.IsShowPointValues = true;
            tar_Graph.PointValueEvent += new ZedGraphControl.PointValueHandler(FAIPointValueHandler);
            mychart.Title.Text = src_FAI_spec.FAI_Name + "_Trend";
            tar_Graph.AxisChange();
            // Make sure the Graph gets redrawn
            tar_Graph.Invalidate();
        }
        public void Display_Result(TDMK_OK2SHIP.Calculate_CPK src_CPK_result, TDMK_OK2SHIP.FAI_Spec src_FAI_spec)
        {
            if(myCode.IsNumeric_Val(src_FAI_spec.UL)!="")
            {
                double USL = Convert.ToDouble(src_FAI_spec.UL); //Convert.ToDouble(src_FAI_spec.SetVal) + Convert.ToDouble(src_FAI_spec.UL);
                txtUL.Text = USL.ToString();//src_FAI_spec.UL;
            }
            else
            {
                txtUL.Text = "NA";
            }
            if(myCode.IsNumeric_Val(src_FAI_spec.LL)!="")
            {
                double LSL = Convert.ToDouble(src_FAI_spec.LL);//Convert.ToDouble(src_FAI_spec.SetVal) - Convert.ToDouble(src_FAI_spec.LL);
                txtLL.Text = LSL.ToString(); //src_FAI_spec.LL;
            }
            else
            {
                txtLL.Text = "NA";
            }
            
            lblFAI.Text = "CPK Result: " + src_FAI_spec.FAI_Name;
            txtSideCheck.Text = src_FAI_spec.check_side;
            txtNorminalDim.Text = src_FAI_spec.SetVal;      
            txtSTDEV.Text = src_CPK_result.stdev.ToString ("#0.00#");
            txtMean.Text = src_CPK_result.mean.ToString("#0.00#");
            txtMax.Text = src_CPK_result.max.ToString("#0.00#");
            txtMin.Text = src_CPK_result.min.ToString("#0.00#");
            txtCP.Text = src_CPK_result.CP;
            txtCPKL.Text = src_CPK_result.CPKL;
            txtCPKU.Text = src_CPK_result.CPKU;
            txtCPK.Text = src_CPK_result.CPK.ToString("#0.00#");
            if(src_CPK_result.CPK<1.67)
            {
                txtCPK.BackColor = Color.Yellow;
                txtCPK.ForeColor = Color.Red;
            }
            txtCPKM.Text = src_CPK_result.CPKM.ToString();
            this.Text = src_FAI_spec.FAI_Name + "_Histogram";
            tabChart.TabPages[0].Text = src_FAI_spec.FAI_Name + "_Histogram";
            tabChart.TabPages[1].Text = src_FAI_spec.FAI_Name + "_Calculate CPK";

        }
    }
}
