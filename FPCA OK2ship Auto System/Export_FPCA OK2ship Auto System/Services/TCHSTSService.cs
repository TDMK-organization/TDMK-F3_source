using Export_FPCA_OK2ship_Auto_System.Repositories;
using Export_FPCA_OK2ship_Auto_System.Services.TDMK_services;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Export_FPCA_OK2ship_Auto_System.Services
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
        public DataTable ReadFile(string location, string itemCode, string lotNo, string type)
        {
            if (string.IsNullOrEmpty(location.Trim()))
            {
                throw new Exception("Location không tồn tại");
            }
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            DataTable dataTable = new DataTable();
            dataTable = _dbContext.GetTableStructure(NAME_TABLE_SQL);

            string[] dict = FileFolderRepository.GetSubFolders(location);
            foreach (var item in dict)
            {
                string nameItem = FileFolderRepository.GetFolderName(item);
                if (nameItem.Split('.').Count() > 1)
                {

                    string zone;
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
                            throw new Exception($"Lỗi FolderName");
                    }
                    string[] files = FileFolderRepository.GetSubFolders(item);
                    foreach (var folder in files)
                    {
                        if (TDMK_ValidateService.compareItemCodeLotNo(FileFolderRepository.GetFolderName(folder), $"{itemCode}-{lotNo}"))
                        {
                            string[] fileLocation = FileFolderRepository.GetFileByExtension(folder, "csv");
                            if (fileLocation.Count() > 0)
                            {
                                DataRow dr = dataTable.NewRow();
                                dr["ID"] = dataTable.Rows.Count + 1;
                                dr["ItemCode"] = itemCode;
                                dr["LotNo"] = lotNo;
                                dr["Type"] = zone;
                                dr["DataLog"] = TDMK_ConverterService.DataTableToJson(SolveFile(fileLocation[0], zone));
                                dataTable.Rows.Add(dr);
                            }
                        }

                    }
                }
            }
            return dataTable;
        }
        private string[] HEADERROWTS = new[] { "Before", "After 100 hours", "After 200 hours" };
        private string[] HEADERROWHS = new[] { "Before", "After 100 hours", "After 200 hours", "After 300 hours", "After 400 hours", "After 500 hours" };
        private string[] HEADERROWTC = new[] { "Before", "After 100 cycles", "After 200 cycles", "After 300 cycles", "After 400 cycles", "After 500 cycles" };
        private DataTable SolveFile(string location, string zone)
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

            return logfile;
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
            int count = _dbContext.BuckDataTable(dataSource, NAME_TABLE_SQL, new string[] { "ItemCode", "LotNo" }, null, "Id");
            return count;
        }

        public DataTable Load(string itemCode, string lotNo)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            DataTable dt = _dbContext.LoadDataTable(NAME_TABLE_SQL, new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            int i = 1;
            foreach (DataRow item in dt.Rows)
            {
                item["ID"] = i++;
            }
            return dt;
        }

        public string Export(string itemCode, string lotNo)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            string messsage = "";
            DataTable dt = _dbContext.LoadDataTable(NAME_TABLE_SQL, new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
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
                    string json = item["DataLog"].ToString();
                    DataTable dataTable = TDMK_ConverterService.JsonToDataTable(json);
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
                            if (dic.TryGetValue(procz, out string address))
                            {
                                //Comestic
                                address = ExportProcess.AddColumn(address, colPlus + 2);
                                workSheet.Cells[address].Value = rowz["Comestic"];

                                //Function test
                                address = ExportProcess.AddRow(address, 1);
                                workSheet.Cells[address].Value = rowz["Function test"];

                            }
                        }


                    }
                }
                exportProcess.SaveExcelWorksheet(ex, processMulti.TrimEnd(':'), $"{itemCode.Trim()}-{lotNo.Trim()}");

            }
            return "Export succesfully!";
        }

        public static string setup_spec(ExcelWorksheet ws)
        {
            int pcs = 0;
            IList<string> valz = new List<string>();
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(ws, new[] { "#", "Content" });
            if (dic.TryGetValue("#", out string value))
            {
                pcs = value.Split('-').Count();
            }
            if (dic.TryGetValue("Content", out string address))
            {
                address = ExportProcess.AddRow(address, 3);
                string text = ws.Cells[address].Text;
                while (!string.IsNullOrEmpty(text))
                {
                    valz.Add(text);
                    address = ExportProcess.AddRow(address, 2);
                    text = ws.Cells[address].Text;
                }

            }
            return $"{pcs}:{string.Join(":", valz)}";
        }
        #endregion
    }
}

