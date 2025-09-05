using Export_FPCA_OK2ship_Auto_System.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Export_FPCA_OK2ship_Auto_System.Services
{
    public class PIDService
    {

        private string _NAMETABLE = "PRODUCT_ID";
        public Dictionary<string, string> _listFile { get; set; } = new Dictionary<string, string>();
        private DBContext _dbContext = new DBContext();
       
        public static int InsertProductID(string itemCode, string lotNo, string process, string content)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            DBContext _dbContext = new DBContext();
            DataTable dataTable = _dbContext.GetTableStructure("PRODUCT_ID");
            DataRow row = dataTable.NewRow();
            row["ID"] = _dbContext.GetID("PRODUCT_ID") + 1;
            row["ItemCode"] = itemCode;
            row["LotNo"] = lotNo;
            row["ProductIDList"] = content;
            row["Process"] = process;
            dataTable.Rows.Add(row);
            return _dbContext.BuckDataTable(dataTable, "PRODUCT_ID", new[] { "ItemCode", "LotNo", "Process" });
        }
        public static string ConverterProductID(DataTable dataTable, string idCol = "ID")
        {
            if (!dataTable.Columns.Contains("ProductID"))
            {
                throw new Exception("Datatable không có product ID");
            }
            List<string> list = new List<string>();
            foreach (DataRow row in dataTable.Rows)
            {
                string item = $"{row["ID"]} - {row["ProductID"]}";
                list.Add(item);
            }
            return String.Join(",", list); ;

        }
        //public static void ConvertterProudctID(DataTable dataTable, string itemCode, string lotNo, string process, )
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

        public static void FillProductID(DataTable dt_analysis, string itemCode, string lotNo, string process)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            DBContext _db = new DBContext();
            DataTable dataTable = _db.LoadDataTable("PRODUCT_ID", new[] { "ItemCode", "LotNo", "Process" }, new[] { itemCode, lotNo, process });
            if (dataTable.Rows.Count <= 0)
            {
                throw new Exception("Không tồn tại product ID của sheet này!");
            }
            Dictionary<string, string> dic = new Dictionary<string, string>();
            string[] content = dataTable.Rows[0]["ProductIDList"].ToString().Split(',');
            foreach (string item in content)
            {
                string id = item.Split('-')[0].Trim();
                string productID = item.Split('-')[1].Trim();
                dic.Add(id, productID);
            }
            dt_analysis.Columns.Add("ProductID");
            foreach (DataRow row in dt_analysis.Rows)
            {
                if (dic.TryGetValue(row["ID"].ToString(), out string productID))
                {
                    row["ProductID"] = productID;
                }
            }
        }
    }
}
