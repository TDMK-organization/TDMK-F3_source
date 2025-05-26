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

namespace OK2SHIP_Software
{
    public partial class IPQC_Chart : Form
    {
        public double[] _data { get; set; }
        public TDMK_OK2SHIP.IPQC_Spec _ipqc_spec { get; set; }
        public IPQC_Chart()
        {
            InitializeComponent();
        }
        public IPQC_Chart( double[] src_data, TDMK_OK2SHIP.IPQC_Spec src_spec)
        {
            InitializeComponent();
            _data = src_data;
            _ipqc_spec = src_spec;
        }
        private void IPQC_Chart_Load(object sender, EventArgs e)
        {
            Draw_Chart(zgChartIPQC, _data, _ipqc_spec);
        }
        public void Draw_Chart(ZedGraph.ZedGraphControl tar_Graph, double[] src_IPQC, TDMK_OK2SHIP.IPQC_Spec src_IPQC_spec)
        {
            GraphPane mychart = tar_Graph.GraphPane;
            PointPairList IPQC_DataList = new PointPairList();
            PointPairList UL_Chart = new PointPairList();
            PointPairList LL_Chart = new PointPairList();
            bool UL_en = false;
            bool LL_en = false;
            double UL=0;
            double LL=0;

            if (src_IPQC_spec.UL !="")
            {
                UL_en = true;
                UL = Convert.ToDouble(src_IPQC_spec.UL);
            }
            if (src_IPQC_spec.LL != "")
            {
                LL_en = true;
                LL = Convert.ToDouble(src_IPQC_spec.LL);
            }          
            
            for (int i = 0; i < src_IPQC.Length; i++)
            {
                IPQC_DataList.Add(i, src_IPQC[i]);
                if (UL_en)
                {
                    UL_Chart.Add(i, UL);
                }
                if (LL_en)
                {
                    LL_Chart.Add(i, LL);
                }
            }
            LineItem FAIData_line = mychart.AddCurve("IPQC Data", IPQC_DataList, Color.Black);
            if (UL_en)
            {
                LineItem UL_line = mychart.AddCurve("UL", UL_Chart, Color.Blue, SymbolType.None);
            }
            if (LL_en)
            {
                LineItem LL_line = mychart.AddCurve("LL", LL_Chart, Color.Red, SymbolType.None);
            }
            tar_Graph.IsShowPointValues = true;
            tar_Graph.PointValueEvent += new ZedGraphControl.PointValueHandler(MyPointValueHandler);
            mychart.Title.Text = src_IPQC_spec.IPQC_Name + "_Trend";
            tar_Graph.AxisChange();
            // Make sure the Graph gets redrawn
            tar_Graph.Invalidate();
        }
        private string MyPointValueHandler(ZedGraphControl control, GraphPane pane, CurveItem curve, int iPt)
        {
            // Get the PointPair that is under the mouse
            PointPair pt = curve[iPt];

            return curve.Label.Text + " is " + pt.Y.ToString("f3") + " at " + (pt.X+1).ToString("f3");
        }
    }
}
