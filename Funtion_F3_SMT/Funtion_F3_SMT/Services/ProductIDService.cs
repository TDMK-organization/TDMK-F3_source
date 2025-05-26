using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OK2SHIP_SMT.Services
{
    public class ProductIDService
    {

        private string _NAMETABLE = "PRODUCT_ID";
        public Dictionary<string, string> _listFile { get; set; } = new Dictionary<string, string>();

        public ProductIDService(string itemCode, string lotno, string location, string[] listCompareCommon, string[] listCompare)
        {
            List<string> list = FileFolderRepository.GetFileByExtension(location, ".csv").ToList();
            foreach (string item in list)
            {
                bool prime = true;
                string fileName = FileFolderRepository.GetFileNameWithoutExtension(item).ToUpper();
                foreach (string com in listCompareCommon)
                {
                    if (!fileName.Contains(com.ToUpper()) || !ContainItemCode(itemCode, lotno, fileName))
                    {
                        prime = false;
                        break;
                    }
                }
                if (prime)
                {
                    foreach (string rieng in listCompare)
                    {
                        if (item.Contains(rieng) && !_listFile.TryGetValue(rieng, out string value))
                        {
                            _listFile.Add(rieng, item);
                            break;
                        }
                    }
                }
            }
        }

        public bool ContainItemCode(string itemCode, string lotNo, string fileName)
        {
            string part = fileName.Split('_')[0].Replace("00000B", "_");
            string itemCodez = part.Split('_')[0];
            string lotNoz = part.Split('_')[1];
            if (itemCode.Equals(itemCodez) && lotNoz.Equals(lotNoz))
            {
                return true;
            }
            return false;
        }
        public List<string> getListProductID(string location)
        {
            DataTable data = FileFolderRepository.ConvertCsvToDataTable(location);
            List<string> list = new List<string>();
            foreach (DataRow row in data.Rows)
            {
                list.Add(row["ProductID"].ToString());
            }
            return list;
        }

        public DataTable ReadFile(string location, string itemCode, string lotNo)
        {
            string fileName = FileFolderRepository.GetFileName(location);

            if (ContainItemCode(itemCode, lotNo, location))
            {
                throw new Exception("Không trùng itemcode Lotno");
            }
            DataTable data = FileFolderRepository.ConvertCsvToDataTable(location);
            return data;
        }

        public int Save(DataTable dataDGV, bool prime)
        {
            DBContext _dbContext = new DBContext();
            if (dataDGV.Rows.Count <= 0)
            {
                throw new Exception("Không có data");
            }

            string number = dataDGV.Rows[0]["IndicationNumber"].ToString().Replace("00000B", "_");
            string itemCode = number.Split('_')[0];
            string lotNo = number.Split('_')[1];

            DataTable dataTable = new DataTable();
            if (prime && dataTable.Rows.Count > 0)
            {
                dataTable = _dbContext.LoadDataTable(_NAMETABLE, new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
                throw new Exception("1234 - HHHH");
            }
            else
            {
                dataTable = _dbContext.GetTableStructure(_NAMETABLE);
            }
            DataRow row = dataTable.NewRow();
            row["ItemCode"] = itemCode;
            row["LotNo"] = lotNo;
            string listRes = "";
            foreach (DataRow rowz in dataDGV.Rows)
            {
                listRes += $"{rowz["ProductID"]}_";
            }
            row["ProductIDList"] = listRes.TrimEnd('_');
            dataTable.Rows.Add(row);
            return _dbContext.BuckDataTable(dataTable, _NAMETABLE, new[] { "ItemCode", "LotNo" }, null, "ID");
        }
        public static List<string> getProductId(string _itemcode, string _lotno)
        {
            DBContext _dbContext = new DBContext();
            string _NAMETABLE = "PRODUCT_ID";
            string itemCode = _itemcode.Trim();
            string lotNo = _lotno.Trim();
            DataTable dataTable = _dbContext.LoadDataTable(_NAMETABLE, new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            if (dataTable.Rows.Count <= 0)
            {
                return new List<string>();
            }
            string json = dataTable.Rows[0]["ProductIDList"].ToString();

            return json.Split('_').ToList();
        }
        public DataTable Load(string _itemcode, string _lotno)
        {
            DBContext _dbContext = new DBContext();
            string itemCode = _itemcode.Trim();
            string lotNo = _lotno.Trim();
            DataTable dataTable = _dbContext.LoadDataTable(_NAMETABLE, new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            if (dataTable.Rows.Count <= 0)
            {
                throw new Exception("Không có dữ liệu");
            }
            return dataTable;
        }
    }
}
