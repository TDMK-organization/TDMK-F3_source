using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OK2SHIP_SMT.Services
{
    public class ShearTestServices
    {
        public static DataTable getStructor()
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("ID", typeof(int));
            dataTable.Columns.Add("ItemCode", typeof(string));
            dataTable.Columns.Add("LotNo", typeof(string));
            dataTable.Columns.Add("Sheet", typeof(string));
            dataTable.Columns.Add("Region", typeof(string));
            dataTable.Columns.Add("Sample", typeof(string));
            dataTable.Columns.Add("Image", typeof(Image));
            dataTable.Columns.Add("Graph", typeof(Image));
            dataTable.Columns.Add("Data", typeof(string));
            dataTable.Columns.Add("Mode 1: Solder joint crack", typeof(string));
            dataTable.Columns.Add("Mode 2: Pad lift", typeof(string));
            dataTable.Columns.Add("Mode 3: Solder joint lift", typeof(string));
            dataTable.Columns.Add("Mode 4: Intermetallic break", typeof(string));
            dataTable.Columns.Add("Mode 5: Component damage", typeof(string));
            dataTable.Columns.Add("Mode 6: Component detached", typeof(string));
            dataTable.Columns.Add("Mode 7: Flex torn", typeof(string));
            dataTable.Columns.Add("Operator", typeof(string));
            dataTable.Columns.Add("Time_Update", typeof(string));
            dataTable.Columns.Add("Remark", typeof(string));
            return dataTable;
        }
    }
    }
