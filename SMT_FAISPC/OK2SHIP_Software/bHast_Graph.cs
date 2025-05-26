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


namespace OK2SHIP_Software
{
    public partial class bHast_Graph : Form
    {
        TDMK_OK2SHIP myCode = new TDMK_OK2SHIP();
        public DataTable graph_dt { get; set; }
        public string ItemCode { get; set; }
        public string LotNo { get; set; }
        public bHast_Graph()
        {
            InitializeComponent();
        }
        public bHast_Graph(DataTable src_dt,string itemcode, string lotno)
        {
            InitializeComponent();
            graph_dt = src_dt;
            ItemCode = itemcode;
            LotNo = lotno;
        }

        private void bHast_Graph_Load(object sender, EventArgs e)
        {
            if (graph_dt!=null)
            {
                DrawGraph_bHast2(graph_dt, zedGraphControl1, ItemCode, LotNo);
            }    
        }
        public void DrawGraph_bHast(DataTable DGV, ZedGraphControl zedGraphControl1)
        {
           // Tạo đối tượng GraphPane
            GraphPane myPane = zedGraphControl1.GraphPane;
            myPane.CurveList.Clear();
            // Đặt tiêu đề cho biểu đồ
            myPane.Title.Text = "Pismo C4.0";
            // Đặt tên cho trục x và y
            myPane.XAxis.Title.Text = "Time (hours)";
            myPane.YAxis.Title.Text = "Resistance (.10^4 Ω)";
            // Lấy dữ liệu từ datagrid view
            double[] time = new double[DGV.Rows.Count - 1]; // Tạo mảng trống
            double[] ch01 = new double[DGV.Rows.Count - 1];
            double[] ch02 = new double[DGV.Rows.Count - 1];
            double[] ch03 = new double[DGV.Rows.Count - 1];
            double[] ch04 = new double[DGV.Rows.Count - 1];
            double[] ch05 = new double[DGV.Rows.Count - 1];
            double[] ch06 = new double[DGV.Rows.Count - 1];
            double[] ch07 = new double[DGV.Rows.Count - 1];
            double[] ch08 = new double[DGV.Rows.Count - 1];

            for (int i = 0; i < DGV.Rows.Count - 1; i++)
            {
                time[i] = Convert.ToDouble(DGV.Rows[i]["Time"]);
                if (DGV.Rows[i]["Ch01"].ToString() != "<1E+03")
                {
                    ch01[i] = (Convert.ToDouble(DGV.Rows[i]["Ch01"])) / 10000;
                }

                if (DGV.Rows[i]["Ch02"].ToString() != "<1E+03")
                {
                    ch02[i] = (Convert.ToDouble(DGV.Rows[i]["Ch02"])) / 10000;
                }
                if (DGV.Rows[i]["Ch03"].ToString() != "<1E+03")
                {
                    ch03[i] = (Convert.ToDouble(DGV.Rows[i]["Ch03"])) / 10000;
                }
                if (DGV.Rows[i]["Ch04"].ToString() != "<1E+03")
                {
                    ch04[i] = (Convert.ToDouble(DGV.Rows[i]["Ch04"])) / 10000;
                }
                if (DGV.Rows[i]["Ch05"].ToString() != "<1E+03")
                {
                    ch05[i] = (Convert.ToDouble(DGV.Rows[i]["Ch05"])) / 10000;
                }
                if (DGV.Rows[i]["Ch06"].ToString() != "<1E+03")
                {
                    ch06[i] = (Convert.ToDouble(DGV.Rows[i]["Ch06"])) / 10000;
                }
                if (DGV.Rows[i]["Ch07"].ToString() != "<1E+03")
                {
                    ch07[i] = (Convert.ToDouble(DGV.Rows[i]["Ch07"])) / 10000;
                }
                if (DGV.Rows[i]["Ch08"].ToString() != "<1E+03")
                {
                    ch08[i] = (Convert.ToDouble(DGV.Rows[i]["Ch08"])) / 10000;
                }
            }
            // Vẽ đường cho các kênh
            LineItem curve1 = myPane.AddCurve("Ch01", time, ch01, Color.Blue, SymbolType.None);
            LineItem curve2 = myPane.AddCurve("Ch02", time, ch02, Color.Red, SymbolType.None);
            LineItem curve3 = myPane.AddCurve("Ch03", time, ch03, Color.Green, SymbolType.None);
            LineItem curve4 = myPane.AddCurve("Ch04", time, ch04, Color.Black, SymbolType.None);
            LineItem curve5 = myPane.AddCurve("Ch05", time, ch05, Color.Yellow, SymbolType.None);
            LineItem curve6 = myPane.AddCurve("Ch06", time, ch06, Color.Cyan, SymbolType.None);
            LineItem curve7 = myPane.AddCurve("Ch07", time, ch07, Color.Brown, SymbolType.None);
            LineItem curve8 = myPane.AddCurve("Ch08", time, ch08, Color.Orange, SymbolType.None);
            // Hiển thị biểu đồ
            zedGraphControl1.AxisChange();
        }
        public void DrawGraph_bHast2(DataTable DGV, ZedGraphControl zedGraphControl1, string tar_ItemCode, string tar_LotNo)
        {
            // Tạo đối tượng GraphPane
            GraphPane myPane = zedGraphControl1.GraphPane;
            myPane.CurveList.Clear();
            // Đặt tiêu đề cho biểu đồ
            myPane.Title.Text = "bHast Graph: "+ tar_ItemCode + "-" + tar_LotNo;
            // Đặt tên cho trục x và y
            myPane.XAxis.Title.Text = "Time (hours)";
            myPane.YAxis.Title.Text = "Resistance (.10^4 Ω)";
            List<Color> CH_color_lst = new List<Color>() { Color.Blue, Color.Red, Color.YellowGreen, Color.Brown, Color.Green, Color.Cyan, Color.Purple, Color.Orange };
            List<string> data_col = new List<string>();
            foreach(DataColumn dc in DGV.Columns)
            {
                if(dc.ColumnName.Contains("Time")||dc.ColumnName.Contains("Ch"))
                {
                    data_col.Add(dc.ColumnName);
                }
            }
            DataTable data_tbl = DGV.AsDataView().ToTable(false, data_col.ToArray());
            List<string> time_data = new List<string>();
            
            Dictionary<string, List<string>> CH_data_lst = new Dictionary<string, List<string>>();
            foreach(var d in data_col)
            {
                List<string> temp_data = new List<string>();
                temp_data = data_tbl.AsEnumerable().Select(x => x.Field<string>(d)).ToList();
                CH_data_lst.Add(d, temp_data);
            }
            int inx = 0;
            foreach(var item in CH_data_lst)
            {
                if(item.Key.Contains("Time"))
                {
                    time_data = item.Value;
                }
                else
                {
                    List<double> x_val = new List<double>();
                    List<double> y_val = new List<double>();
                    for (int i = 0; i < item.Value.Count; i++)
                    {
                        if (myCode.IsNumeric(item.Value[i]))
                        {
                            double y = Convert.ToDouble(item.Value[i]);
                            y_val.Add(y / 10000);
                            x_val.Add(Convert.ToDouble(time_data[i]));
                        }
                    }
                    if(inx<8)
                    {
                        myPane.AddCurve(item.Key, x_val.ToArray(), y_val.ToArray(), CH_color_lst[inx], SymbolType.None);
                    }                  
                    inx++;
                }
            }
            // Hiển thị biểu đồ
            zedGraphControl1.AxisChange();
        }
    }
}
