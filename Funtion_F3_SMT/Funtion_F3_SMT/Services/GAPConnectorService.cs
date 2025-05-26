using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OK2SHIP_SMT.Services
{
    public class GAPConnectorService
    {
        public static List<string> Get_ProductID(string location, int numPcs, string itemCode)
        {
            string locationfile = Directory.GetFiles(location).Where(file => file.ToLower().EndsWith(".csv")).ToList().FirstOrDefault(item => Path.GetFileName(item).Contains(itemCode));
            DataTable dt = FileFolderRepository.ConvertCsvToDataTable(locationfile);
            List<string> productIDs = dt.AsEnumerable().Select(row => row.Field<string>("ProductID")).Take(numPcs).ToList();
      
            return productIDs;
        }
    }
}
