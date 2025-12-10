using OfficeOpenXml;
using OfficeOpenXml.Style;
using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.Services
{
    public class TCHSTSService
    {
        #region Properties

        private readonly string NAME_TABLE_SQL = "TC_HS_TS";
        private DBContext _dbContext;
        #endregion
        #region Constructor
        public TCHSTSService()
        {
            _dbContext = new DBContext();
        }
        #endregion
        #region Methods
        public DataTable ReadFile(string location, string itemCode, string lotNo, string type, string PIDLocation)
        {
            if (string.IsNullOrEmpty(location.Trim()))
            {
                throw new Exception("Location không tồn tại");
            }
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            bool prime = false;
            PIDLocation = PIDLocation.Trim();
            DataTable dataTable = new DataTable();
            dataTable = _dbContext.GetTableStructure(NAME_TABLE_SQL);
            if (!string.IsNullOrEmpty(PIDLocation))
            {
                ProductIDService productID = new ProductIDService(itemCode, lotNo, PIDLocation, new[] { "Copy" }, new[] { "Thermal cycling", "Heat Soak", "Thermal shock" });
                Debugger.Break();
                prime = productID._listFile.Count > 0;
            }
            string[] dict = FileFolderRepository.GetSubFolders(location);
            foreach (var item in dict)
            {
                string nameItem = FileFolderRepository.GetFolderName(item);
                if (nameItem.Split('.').Count() > 1)
                {

                    string zone = "";
                    string folderName = nameItem.Split('.')[1].Trim();
                    switch (folderName)
                    {
                        case "TS":
                            zone = "TS";
                            break;
                        case "HS":
                            zone = "HS";
                            break;
                        case "TC":
                            zone = "TC";
                            break;
                        default:
                            Debugger.Break();
                            break;
                    }
                    if (!string.IsNullOrEmpty(zone))
                    {
                        string[] files = FileFolderRepository.GetSubFolders(item);
                        foreach (var folder in files)
                        {
                            if (ValidateService.compareItemCodeLotNo(FileFolderRepository.GetFolderName(folder), $"{itemCode}-{lotNo}"))
                            {
                                string[] fileLocation = FileFolderRepository.GetFileByExtension(folder, "csv");
                                if (fileLocation.Count() > 0)
                                {
                                    DataRow dr = dataTable.NewRow();
                                    dr["ID"] = dataTable.Rows.Count + 1;
                                    dr["ItemCode"] = itemCode;
                                    dr["LotNo"] = lotNo;
                                    dr["Type"] = switchType(zone);
                                    DataTable newDt = new DataTable();
                                    DataTable dt = SolveFile(fileLocation[0], zone, out newDt);
                                    if (prime)
                                    {

                                    }
                                    dr["DataLog"] = $"{ConverterService.DataTableToJson(newDt)}@{ConverterService.DataTableToJson(dt)}";
                                    dataTable.Rows.Add(dr);
                                }
                            }

                        }
                    }
                }
            }

            return dataTable;
        }
        private static string switchType(string nameType)
        {
            nameType = nameType.Trim();
            switch (nameType)
            {
                case "All":
                    return "*";
                case "TS":
                    return "Thermal Shock";
                case "HS":
                    return "Heat Soak and Recovery";
                case "TC":
                    return "Thermal Cycling";
                case "Thermal Shock":
                    return "TS";
                case "Heat Soak and Recovery":
                    return "HS";
                case "Thermal Cycling":
                    return "TC";
                default:
                    return "";
            }
        }
        private string[] HEADERROWTS = new[] { "Before", "After 100 cycles", "After 200 cycles" };
        private string[] HEADERROWHS = new[] { "Before", "After 100 hours", "After 200 hours", "After 300 hours", "After 400 hours", "After 500 hours" };
        private string[] HEADERROWTC = new[] { "Before", "After 100 cycles", "After 200 cycles", "After 300 cycles", "After 400 cycles", "After 500 cycles" };
        private DataTable SolveFile(string location, string zone, out DataTable tableResult)
        {
            ///
            string[] headerRow;
            switch (zone)
            {
                case "TS":
                    headerRow = HEADERROWTS;
                    break;
                case "HS":
                    headerRow = HEADERROWHS;
                    break;
                case "TC":
                    headerRow = HEADERROWTC;
                    break;
                default:
                    throw new Exception("Lỗi Zone");
            }
            DataTable logfile = FileFolderRepository.ConvertCsvToDataTable(location);
            DataTable ze = new DataTable();
            bool primez = false;
            foreach (DataColumn column in logfile.Columns)
            {
                string value = logfile.Rows[0][column].ToString();
                if (!string.IsNullOrEmpty(value))
                {
                    ze.Columns.Add(column.ColumnName);
                    if (!primez)
                    {
                        primez = true;
                        ze.Rows.Add(ze.NewRow());
                        ze.Rows.Add(ze.NewRow());
                        ze.Rows.Add(ze.NewRow());
                    }
                    ze.Rows[0][column.ColumnName] = logfile.Rows[0][column.ColumnName];
                    ze.Rows[1][column.ColumnName] = logfile.Rows[1][column.ColumnName];
                    ze.Rows[2][column.ColumnName] = logfile.Rows[2][column.ColumnName];

                }
            }
            tableResult = ze;
            logfile.Columns.RemoveAt(10);
            logfile.Columns.RemoveAt(9);
            logfile.Columns.RemoveAt(8);
            logfile.Columns.RemoveAt(7);
            logfile.Columns.RemoveAt(6);
            logfile.Columns.RemoveAt(5);
            logfile.Columns.RemoveAt(4);
            logfile.Columns.RemoveAt(0);
            int colCount = logfile.Columns.Count;
            //DataRow row = logfile
            logfile.Columns.Add("Comestic");
            logfile.Columns.Add("ID");
            logfile.Columns.Add("Function test");
            logfile.Columns.Add("Content");
            // Tạo một danh sách để lưu trữ các hàng cần xóa
            List<DataRow> rowsToRemove = new List<DataRow>();
            string format = "yyyyMMdd HH:mm:ss";
            string date = logfile.Rows[5]["Time"].ToString();
            if (!DateTime.TryParseExact(date, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime datePoint))
            {
                throw new Exception("Error convert Datetime");
            }
            int prime = 0;
            int i = 1;
            int countint = 0;
            // Duyệt qua từng hàng trong DataTable
            foreach (DataRow row in logfile.Rows)
            {
                // Kiểm tra xem cột chỉ định có chứa dấu "_" hay không
                if (row["UUT"] != DBNull.Value && row["UUT"].ToString().Contains("_"))
                {
                    // Thêm hàng vào danh sách cần xóa
                    rowsToRemove.Add(row);
                }
                else
                {
                    if (countint++ >= 3)
                    {
                        row["ID"] = i++;
                    }
                    row["Function test"] = FunctionTest(logfile.Rows[1], logfile.Rows[2], row, colCount) ? "PASS" : "FAIL";
                    string stzt = row["Time"].ToString();
                    if (DateTime.TryParseExact(stzt, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dayz))
                    {
                        int today = Math.Abs(datePoint.Day - dayz.Day);
                        if (today >= 2)
                        {
                            prime++;
                            prime = prime >= headerRow.Count() ? prime-- : prime++;
                            datePoint = dayz;
                        }
                        row["Content"] = headerRow[prime];
                    }
                }
            }
            // Xóa các hàng đã được đánh dấu
            foreach (DataRow rowToRemove in rowsToRemove)
            {
                logfile.Rows.Remove(rowToRemove);
            }

            logfile.Rows.Remove(logfile.Rows[2]);
            logfile.Rows.Remove(logfile.Rows[1]);
            logfile.Rows.Remove(logfile.Rows[0]);
            DataTable dataTable = logfile.Clone();
            Dictionary<string, List<DataRow>> dic = new Dictionary<string, List<DataRow>>();
            foreach (DataRow row in logfile.Rows)
            {
                DataRow rowZ = dataTable.NewRow();
                foreach (DataColumn item in dataTable.Columns)
                {
                    rowZ[item.ColumnName] = row[item.ColumnName];
                }
                rowZ["UUT"] = rowZ["UUT"].ToString().ToUpper();
                string id = rowZ["UUT"].ToString();
                if (dic.TryGetValue(id, out List<DataRow> rowz))
                {
                    rowz.Add(rowZ);
                }
                else
                {
                    List<DataRow> list = new List<DataRow>();
                    list.Add(rowZ);
                    dic.Add(id, list);
                }
            }
            int iZ = 1;
            foreach (string key in dic.Keys)
            {
                foreach (DataRow row in dic[key])
                {
                    row["ID"] = iZ;
                    dataTable.Rows.Add(row);
                }
                iZ++;
            }
            return dataTable;
        }
        private bool FunctionTest(DataRow rowSampleMax, DataRow rowSampleMin, DataRow rowValue, int n)
        {
            try
            {

                int i = 3;

                while (i < n)
                {
                    if (rowSampleMax[i].ToString().Contains('M'))
                    {
                        return true;
                    }
                    double maxValue = double.Parse(rowSampleMax[i].ToString());
                    double minValue = double.Parse(rowSampleMin[i].ToString());
                    double value = double.Parse(rowValue[i].ToString());
                    if (value > maxValue || value < minValue)
                    {
                        return false;
                    }
                    i++;
                }
            }
            catch
            {
                return false;
            }
            return true;
        }

        public int Save(DataTable dataSource, bool prime)
        {
            if (dataSource.Rows.Count == 0)
            {
                throw new Exception("Không có dữ liệu để lưu");
            }
            //1. Check data had in database (itemcode, lotno)
            // 1.1 Get ItemCode lotno in datatable
            string itemCode = dataSource.Rows[0]["ItemCode"].ToString();
            string lotno = dataSource.Rows[0]["LotNo"].ToString();

            // 1.2 Load data from database
            DataTable dt = _dbContext.LoadDataTable(NAME_TABLE_SQL, new string[] { "ItemCode", "LotNo" }, new string[] { itemCode, lotno }, new string[] { "ItemCode", "LotNo" });

            // 1.3 Nếu dữ liệu đã tồn tại và không được phép lưu thì trả về thông báo để hỏi người dùng

            if (!prime && dt.Rows.Count > 0)
            {
                throw new Exception("2267 - Dữ liệu đã tồn tại");
            }

            //2. Save data to database if prime = true
            //2.1 Lưu và thay thế
            foreach (DataRow row in dataSource.Rows)
            {
                row["Type"] = switchType(row["Type"].ToString());
            }
            int count = _dbContext.BuckDataTable(dataSource, NAME_TABLE_SQL, new string[] { "ItemCode", "LotNo" }, null, "Id");
            return count;
        }

        public DataTable Load(string itemCode, string lotNo, string type)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            type = switchType(type.Trim());
            DataTable dt = new DataTable();
            if (!type.Equals("*"))
            {
                dt = _dbContext.LoadDataTable(NAME_TABLE_SQL, new[] { "ItemCode", "LotNo", "Type" }, new[] { itemCode, lotNo, type });
            }
            else
            {
                dt = _dbContext.LoadDataTable(NAME_TABLE_SQL, new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            }
            int i = 1;
            foreach (DataRow item in dt.Rows)
            {
                item["ID"] = i++;
                item["Type"] = switchType(item["Type"].ToString());
            }
            return dt;
        }

        public string Export(string itemCode, string lotNo, string type)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            type = switchType(type.Trim());
            DataTable dt = new DataTable();
            if (type.Equals("*"))
            {

                dt = _dbContext.LoadDataTable(NAME_TABLE_SQL, new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            }
            else
            {
                dt = _dbContext.LoadDataTable(NAME_TABLE_SQL, new[] { "ItemCode", "LotNo", "Type" }, new[] { itemCode, lotNo, type });
            }
            if (dt.Rows.Count == 0)
            {
                throw new Exception("Không có dữ liệu");
            }
            string processMulti = "";
            ExportProcess exportProcess = new ExportProcess();
            using (ExcelPackage ex = exportProcess.FindFormatProcess(NAME_TABLE_SQL, itemCode, lotNo))
            {
                ExcelWorksheet workSheet = null;
                //exportProcess.FindSheet(ex, "Thermal Shock");

                foreach (DataRow item in dt.Rows)
                {
                    string process = item["Type"].ToString().Trim();
                    string[] headerCol = null;
                    switch (process)
                    {
                        case "TS":
                            process = "Thermal Shock";
                            headerCol = HEADERROWTS;
                            break;
                        case "TC":
                            process = "Thermal Cycling";
                            headerCol = HEADERROWTC;
                            break;
                        case "HS":
                            process = "Heat Soak and Recovery";
                            headerCol = HEADERROWHS;
                            break;
                        default:
                            throw new Exception("Lỗi type");
                    }
                    workSheet = exportProcess.FindSheet(ex, process);
                    processMulti += $"{process}:";

                    // find col
                    IDictionary<string, string> dic = ExportProcess.FindAddressByText(workSheet, headerCol.Concat(new[] { "Flex SN" }).ToArray());
                    string json = item["DataLog"].ToString().Split('@')[1];
                    DataTable dataTable = ConverterService.JsonToDataTable(json);
                    int pcs = 45, bf = 0, at1 = 0, at2 = 0, at3 = 0, at4 = 0, at5 = 0;
                    foreach (DataRow rowz in dataTable.Rows)
                    {

                        int colPlus = 0;
                        string procz = rowz["Content"].ToString();
                        switch (procz)
                        {
                            case "Before":
                                colPlus = bf++;
                                break;
                            case "After 100 hours":
                            case "After 100 cycles":
                                colPlus = at1++;
                                break;
                            case "After 200 hours":
                            case "After 200 cycles":
                                colPlus = at2++;
                                break;
                            case "After 300 hours":
                            case "After 300 cycles":
                                colPlus = at3++;
                                break;
                            case "After 400 hours":
                            case "After 400 cycles":
                                colPlus = at4++;
                                break;
                            case "After 500 hours":
                            case "After 500 cycles":
                                colPlus = at5++;
                                break;
                        }
                        if (colPlus < pcs)
                        {
                            if (dic.TryGetValue("Flex SN", out string snz))
                            {
                                snz = ExportProcess.AddColumn(snz, colPlus + 1);
                                if (string.IsNullOrEmpty(rowz["UUT"].ToString()))
                                {
                                    throw new Exception("Kiểm tra lại data không có Flex SN");
                                }
                                if (string.IsNullOrEmpty(workSheet.Cells[snz].Text))
                                {
                                    workSheet.Cells[snz].Value = rowz["UUT"];
                                }
                                else
                                {
                                    if (!rowz["UUT"].ToString().Contains(workSheet.Cells[snz].Text))
                                    {
                                        throw new Exception("không trùng flex sn");
                                    }
                                }
                            }
                            if (dic.TryGetValue(procz, out string address))
                            {
                                //Comestic
                                address = ExportProcess.AddColumn(address, colPlus + 2);
                                workSheet.Cells[address].Value = rowz["Comestic"];
                                //if (rowz["Comestic"] != null)
                                //{
                                //    string z = rowz["Comestic"].ToString();
                                //    if (rowz["Comestic"].ToString().Contains("OK"))
                                //    {
                                //        Color col = System.Drawing.ColorTranslator.FromHtml("#B7DEE8");
                                //        workSheet.Cells[address].Style.Fill.BackgroundColor.SetColor(Color.CornflowerBlue);
                                //    }
                                //    if (rowz["Comestic"].ToString().Contains("OK"))
                                //    {
                                //        Color col = System.Drawing.ColorTranslator.FromHtml("#B7DEE8");
                                //        workSheet.Cells[address].Style.Fill.BackgroundColor.SetColor(Color.CornflowerBlue);
                                //    }
                                //    if (string.IsNullOrEmpty(rowz["Comestic"].ToString()))
                                //    {
                                //        Color col = System.Drawing.ColorTranslator.FromHtml("#B7DEE8");
                                //        workSheet.Cells[address].Style.Fill.BackgroundColor.SetColor(Color.CornflowerBlue);
                                //    }
                                //}
                                //else
                                //{
                                //    Color col = System.Drawing.ColorTranslator.FromHtml("#B7DEE8");
                                //    workSheet.Cells[address].Style.Fill.BackgroundColor.SetColor(Color.CornflowerBlue);
                                //}
                                //Function test
                                address = ExportProcess.AddRow(address, 1);
                                workSheet.Cells[address].Value = rowz["Function test"];
                                //if (rowz["Function test"] != null)
                                //{
                                //    if (rowz["Function test"].ToString().ToLower().Contains("pass"))
                                //    {
                                //        workSheet.Cells[address].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                //    }
                                //    if (rowz["Function test"].ToString().ToLower().Contains("fail"))
                                //    {
                                //        workSheet.Cells[address].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                //    }
                                //}
                                //else
                                //{
                                //    workSheet.Cells[address].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                //}
                            }
                        }


                    }
                }
                exportProcess.SaveExcelWorksheet(ex, processMulti.TrimEnd(':'), $"{itemCode.Trim()}-{lotNo.Trim()}");

            }
            return "Export succesfully!";
        }

        public void ExportLogFile(string itemCode, string lotNo, string type)
        {

        }
        #endregion
    }
}

