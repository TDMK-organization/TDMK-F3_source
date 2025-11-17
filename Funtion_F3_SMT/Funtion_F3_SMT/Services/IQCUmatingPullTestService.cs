using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OK2SHIP_SMT.Services
{
    public class IQCUmatingPullTestService
    {
        public static DataTable getStructor()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("ItemCode", typeof(string));
            dt.Columns.Add("LotNo", typeof(string));
            dt.Columns.Add("Sheet", typeof(string));
            dt.Columns.Add("Region", typeof(string));
            dt.Columns.Add("Sample", typeof(string));
            dt.Columns.Add("Image", typeof(Image));
            dt.Columns.Add("Graph", typeof(Image));
            dt.Columns.Add("Data", typeof(string));
            dt.Columns.Add("Operator", typeof(string));
            dt.Columns.Add("Time_Update", typeof(string));
            dt.Columns.Add("Remark", typeof(string));
            return dt;
        }
    }
}
