using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OK2SHIP_SMT.Services
{
    public class PeelTestWOSUSService
    {
        private DBContext _dBContext = new DBContext();
        public bool checkPeelTest(string itemCode, string lotNo)
        {
            DataTable dataTable = _dBContext.LoadDataTable("PEEL_TEST", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            if (dataTable.Rows.Count <= 0)
            {
                return false;
            }
            foreach (DataRow row in dataTable.Rows)
            {
                int num = int.Parse(row["Mode 1: Solder joint crack"].ToString().Split('%')[0]);
                if (num > 50)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
