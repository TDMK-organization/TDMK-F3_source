using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OK2SHIP_SMT.Services
{
    public class CrossSectionService
    {
        public DataTable getStructorTable()
        {
            DataTable table = new DataTable();
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("ItemCode", typeof(string));
            table.Columns.Add("LotNo", typeof(string));
            table.Columns.Add("Sheet", typeof(string));
            table.Columns.Add("Region", typeof(string));
            table.Columns.Add("Sample", typeof(string));
            table.Columns.Add("Image1", typeof(Image));
            table.Columns.Add("Image2", typeof(Image));
            table.Columns.Add("Data", typeof(string));
            table.Columns.Add("Operator", typeof(string));
            table.Columns.Add("Time_Update", typeof(string));
            table.Columns.Add("Remark", typeof(string));
            return table;
        }
    }
}
