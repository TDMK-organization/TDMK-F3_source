using OfficeOpenXml;
using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.Services
{
    public class BarCodeVertification
    {
        public bool BarCodeCheckRule(string barcode, string format, ref string[] result)
        {
            return true;
        }
        public DataTable ReadProcess(string locationFolder, string itemCode, string lotNo)
        {
            DataTable dataTable = new DataTable();

            try
            {
                string[] colView = { "DateTime", "Module", "Overall Grade" };
                DataTable dataTables = FileFolderRepository.ConvertCsvToDataTable(locationFolder, 100).DefaultView.ToTable(false, colView);
                dataTable = dataTables.AsEnumerable()
                                      .Where(row => row.Field<string>("Overall Grade") == "A")
                                      .Take(100)
                                      .CopyToDataTable();
                dataTable.Columns.Add("ID");
                //convert DateTime
                dataTable.Columns.Add("Datetime", typeof(String));
                foreach (DataRow row in dataTable.Rows)
                {
                    if (DateTime.TryParse(row["DateTime"].ToString(), out DateTime dateTimeValue))
                    {
                        row["Datetime"] = checkDOM(row["Module"].ToString()).ToString("MMM-dd");
                    }
                    else
                    {
                        // Handle conversion failure (e.g., assign DateTime.MinValue)
                        row["Datetime"] = DateTime.MinValue;
                    }
                }

                // Remove old column
                dataTable.Columns.Remove("DateTime");
                //
                dataTable.Columns.Add("ItemCode");
                dataTable.Columns.Add("LotNo");
                int id = 1;
                foreach (DataRow item in dataTable.Rows)
                {
                    item["ItemCode"] = itemCode;
                    item["LotNo"] = lotNo;
                    item["ID"] = id++;
                }
                //Debugger.Break();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dataTable;
        }

        public void SaveProcess(DataTable dataTable, bool prime, string itemCode, string lotNo)
        {
            string json = ConverterService.DataTableToJson(dataTable);

            using (DBContext db = new DBContext())
            {
                DataTable saveDt = db.GetTableStructure("BAR_CODE_VERIFICATION");
                DataRow dr = saveDt.NewRow();
                dr["ItemCode"] = itemCode;
                dr["LotNo"] = lotNo;
                dr["ListSN"] = json;
                saveDt.Rows.Add(dr);
                if (prime)
                {
                    db.BuckDataTable(saveDt, "BAR_CODE_VERIFICATION", new[] { "ItemCode", "LotNo" }, null, "Id");
                }
                else
                {
                    db.SaveDataTable(saveDt, "BAR_CODE_VERIFICATION", null, "Id");
                }
            }

        }
        public string GetCheckSumData(string location)
        {
            string extension = Path.GetExtension(location);
            if (!extension.Contains("xlsx"))
            {
                Debugger.Break();
                return "";
            }
            ExportProcess exportProcess = new ExportProcess();
            using (ExcelPackage ex = exportProcess.OpenFileExcel(location))
            {
                using (ExcelWorksheet workSheet = ex.Workbook.Worksheets[0])
                {
                    string[] strings = { "EEEEEEE code", "Factory code", "Item name" };
                    IDictionary<string, string> addressDic = ExportProcess.FindAddressByText(workSheet, strings);
                    string factoryCode = workSheet.Cells[workSheet.Cells[addressDic["Factory code"]].End.Row, workSheet.Cells[addressDic["Factory code"]].End.Column + 2].Text.ToString();
                    string ECode = workSheet.Cells[workSheet.Cells[addressDic["EEEEEEE code"]].End.Row + 1, workSheet.Cells[addressDic["EEEEEEE code"]].End.Column].Text.ToString();
                    string addressItem = addressDic["Item name"].Split('-')[1];
                    string itemName = workSheet.Cells[workSheet.Cells[addressItem].End.Row + 1, workSheet.Cells[addressItem].End.Column].Text.ToString();
                    return $"{itemName}-{factoryCode}-{ECode}";
                }
            }
        }
        public string Export(string itemCode, string lotNo)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            ExportProcess exportProcess = new ExportProcess();
            using (ExcelPackage ex = exportProcess.FindFormatProcess("Bar Code Verification", itemCode, lotNo))
            {
                using (ExcelWorksheet workSheet = exportProcess.FindSheet(ex, "Bar Code Verification"))
                {
                    //string address = exportProcess.FindAddressByText(workSheet, "Bar Code Verification");
                    string[] colHeader = { "No", "SN", "Follow Bar code Rule?", "Grade" };
                    string[] colHeader1 = { "Code rule" };
                    string[] colHeader2 = { "Nhập ngày sản xuất vào đây", "Coppy tu log file vào đây", "Factory (PPP) check", "Day of manufacturing (DOM)", "Laser machine (Sssss) check", "Serian No  (sSSSS) check ", "Item name (EEEEEEE) Check" };
                    IDictionary<string, string> addressHeader = ExportProcess.FindAddressByText(workSheet, colHeader1.Concat(colHeader2).ToArray());
                    IDictionary<string, string> addressHeader2 = ExportProcess.FindAddressByText(workSheet, colHeader.ToArray(), true);
                    addressHeader = DictionaryService.MergeDictionaries(addressHeader, addressHeader2);
                    DBContext db = new DBContext();
                    DataTable dz = db.LoadDataTable("BAR_CODE_VERIFICATION", new string[] { "ItemCode", "LotNo" }, new string[] { itemCode, lotNo });
                    DataTable dt = ConverterService.JsonToDataTable(dz.Rows[0]["ListSN"].ToString());
                    int i = 0;
                    //Take Code rule
                    bool primeCodeRule = addressHeader.TryGetValue("Code rule", out string add);
                    IList<string> listCodeRule = new List<string>();
                    if (primeCodeRule)
                    {
                        int startS = 0;
                        string code = dt.Rows[0]["Module"].ToString();
                        workSheet.Cells[ExportProcess.AddRow(add, -1)].Value = code;
                        while (workSheet.Cells[ExportProcess.AddRow(add, 1)].Value != null)
                        {
                            add = ExportProcess.AddRow(add, 1);
                            string codeRule = workSheet.Cells[add].Value == null ? null : workSheet.Cells[add].Value.ToString();
                            listCodeRule.Add(codeRule);
                            workSheet.Cells[ExportProcess.AddColumn(add, 2)].Value = code.Substring(startS, codeRule.Length);
                            workSheet.Cells[ExportProcess.AddColumn(add, 3)].Value = ValidateService.isDigitAndChar(code.Substring(startS, codeRule.Length)) ? "OK" : "NG";
                            startS += codeRule.Length;
                        }
                    }
                    TableOfContentService tbd = new TableOfContentService();

                    DataTable tbc_DT = tbd.getDataTableByItemCode(itemCode);
                    if (tbc_DT.Rows.Count <= 0)
                    {
                        throw new Exception("Chưa có dữ liệu table of content!");
                    }
                    string EEEEECode = tbc_DT.Rows[0]["EEEECode"].ToString().TrimEnd('\n').TrimEnd('\r');
                    string FactoryCode = tbc_DT.Rows[0]["FactoryCode"].ToString();

                    foreach (DataRow item in dt.Rows)
                    {
                        i++;
                        ///Table 2

                        if (addressHeader.TryGetValue("Nhập ngày sản xuất vào đây", out string addressz))
                        {
                            addressz = ExportProcess.AddRow(addressz, i + 1);
                            workSheet.Cells[addressz].Value = item["DateTime"];
                        }
                        if (addressHeader.TryGetValue("Coppy tu log file vào đây", out addressz))
                        {
                            addressz = ExportProcess.AddRow(addressz, i + 1);
                            workSheet.Cells[addressz].Value = item["Module"];
                        }
                        string[] bit = JudgeSN(item["Module"].ToString(), listCodeRule);
                        //if (bit.Length == listCodeRule.Count)
                        //{
                        string module = item["Module"].ToString();
                        workSheet.Cells[ExportProcess.AddColumn(addressz, 1)].Value = checkPPP(item["Module"].ToString(), FactoryCode) ? "OK" : "NG";
                        workSheet.Cells[ExportProcess.AddColumn(addressz, 2)].Value = checkDOM(module) != DateTime.MinValue ? "OK" : "NG";
                        workSheet.Cells[ExportProcess.AddColumn(addressz, 3)].Value = checkSssss(item["Module"].ToString()) ? "OK" : "NG";
                        workSheet.Cells[ExportProcess.AddColumn(addressz, 4)].Value = checksSSSS(item["Module"].ToString()) ? "OK" : "NG";
                        workSheet.Cells[ExportProcess.AddColumn(addressz, 5)].Value = checkEEEE(item["Module"].ToString(), EEEEECode) ? "OK" : "NG";
                        //}


                        /// Table 1
                        foreach (var str in colHeader)
                        {
                            if (addressHeader.TryGetValue(str, out string address))
                            {
                                address = ExportProcess.AddRow(address, i);
                                if (str.Contains("No"))
                                {
                                    workSheet.Cells[address].Value = i;
                                }
                                else if (str.Contains("Follow Bar code Rule?"))
                                {
                                    workSheet.Cells[address].Value = bit.Length == listCodeRule.Count ? "Yes" : "No";
                                }
                                else if (str.Contains("SN"))
                                {
                                    workSheet.Cells[address].Value = item["Module"];
                                }
                                else if (str.Contains("Grade"))
                                {
                                    workSheet.Cells[address].Value = item["Overall Grade"];
                                }
                                else
                                {
                                    workSheet.Cells[address].Value = item[str];
                                }
                            }
                        }



                    }
                    exportProcess.SaveExcelWorksheet(ex, "Bar Code Verification", $"{itemCode} - {lotNo}");

                }
            }
            return "Export Complete";
        }
        private string[] JudgeSN(string SN, IList<string> code)
        {
            List<string> res = new List<string>();
            int startS = 0;
            foreach (var item in code)
            {
                int length = item.Length;
                string sub = SN.Substring(startS, length);
                if (ValidateService.isDigitAndChar(sub))
                {
                    res.Add(sub);
                }
                startS += length;
            }
            return res.ToArray();
        }

        public DataTable LoadProcess(string itemCode, string lotNo)
        {
            DataTable dt = new DataTable();
            DBContext db = new DBContext();
            dt = db.LoadDataTable("BAR_CODE_VERIFICATION", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            if (dt.Rows.Count <= 0)
            {
                throw new Exception("Không có dữ liệu của itemcode lotno này");
            }
            DataTable dz = ConverterService.JsonToDataTable(dt.Rows[0]["ListSN"].ToString());

            return dz;
        }

        public DataTable CheckSum(DataTable dataTable)
        {

            if(dataTable == null)
            {
                return dataTable;
            }
            if (dataTable.Rows.Count <= 0)
            {
                throw new Exception("Không có dữ liệu của itemcode này");
            }
            string itemCode = dataTable.Rows[0]["ItemCode"].ToString();

            // get code
            TableOfContentService tbd = new TableOfContentService();
            DataTable tbc_DT = tbd.getDataTableByItemCode(itemCode);
            if (tbc_DT.Rows.Count <= 0)
            {
                throw new Exception("Chưa có dữ liệu table of content!");
            }
            string EEEEECode = tbc_DT.Rows[0]["EEEECode"].ToString().TrimEnd('\n').TrimEnd('\r');
            string FactoryCode = tbc_DT.Rows[0]["FactoryCode"].ToString();

            DataTable dtz = new DataTable();
            dtz.Columns.Add("Datetime");
            dtz.Columns.Add("Module");
            dtz.Columns.Add("PPP");
            dtz.Columns.Add("DOM");
            dtz.Columns.Add("Sssss");
            dtz.Columns.Add("sSSSS");
            dtz.Columns.Add("EEEEEEE");

            foreach (DataRow item in dataTable.Rows)
            {

                DataRow row = dtz.NewRow();
                row["DOM"] = "NG";


                if (DateTime.TryParse(item["Datetime"].ToString(), out DateTime date))
                {
                    row["Datetime"] = item["Datetime"];
                    row["DOM"] = checkDOM(item["Module"].ToString()) != DateTime.MinValue ? "OK" : "NG";
                }


                row["Module"] = item["Module"];
                row["PPP"] = checkPPP(item["Module"].ToString(), FactoryCode) ? "OK" : "NG";
                row["Sssss"] = checkSssss(item["Module"].ToString()) ? "OK" : "NG";
                row["sSSSS"] = checksSSSS(item["Module"].ToString()) ? "OK" : "NG";
                row["EEEEEEE"] = checkEEEE(item["Module"].ToString(), EEEEECode) ? "OK" : "NG";

                dtz.Rows.Add(row);
            }

            return dtz;

        }
        private static bool checkSssss(string Module)
        {
            try
            {
                Module = Module.Trim().Substring(6, 1).Trim();
            }
            catch
            {
                return false;
            }
            if (Module.Length != 1)
            {
                return false;
            }
            char S = Module.ToUpper()[0];
            if (S <= 'Z' && S >= 'A' || S <= '9' && S >= '0')
            {
                return true;
            }
            return false;
        }
        private static bool checksSSSS(string Module)
        {
            try
            {
                Module = Module.Trim().Substring(7, 4).Trim();
            }
            catch
            {
                return false;
            }
            if (Module.Length != 4)
            {
                return false;
            }
            int i = 0;
            while (i < Module.Length)
            {
                char S = Module.ToUpper()[i];
                if (S <= 'Z' && S >= 'A')
                {
                }
                else if (S <= '9' && S >= '0')
                {
                }
                else
                {
                    return false;
                }
                i++;
            }
            return true;
        }
        private static bool checkEEEE(string Module, string code)
        {
            try
            {
                Module = Module.Trim().Substring(11).Trim();
            }
            catch
            {
                return false;
            }
            code = code.Trim();
            if (Module.Length != 7)
            {
                return false;
            }

            return code.Equals(Module);
        }
        private static DateTime checkDOM(string Module)
        {
            string DOM_DB = "0123456789ABCDEFGHJKLMNPQRSTUVWXYZ";
            try
            {
                Module = Module.Trim().Substring(3, 3).Trim();
            }
            catch
            {
                return DateTime.MinValue;
            }

            if (Module.Length != 3)
            {
                return DateTime.MinValue;
            }
            int D = DOM_DB.IndexOf(Module[0]) * 34 * 34;
            int O = DOM_DB.IndexOf(Module[1]) * 34;
            int M = DOM_DB.IndexOf(Module[2]);
            int dom_Value = D + O + M;
            DateTime dateTime = intToDate(dom_Value);

            return dateTime;
        }
        public static DateTime intToDate(int n)
        {
            // 1/1/1970 là mốc thời gian gốc (Unix epoch)
            DateTime goc = new DateTime(1970, 1, 1);

            // Thêm n ngày vào mốc thời gian gốc
            DateTime ketQua = goc.AddDays(n);

            return ketQua;
        }
        private static bool checkPPP(string Module, string code)
        {
            Module = Module.Trim().Substring(0, 3).Trim();
            code = code.Trim();
            if (Module.Length != 3)
            {
                return false;
            }

            return code.Equals(Module);
        }

        public int UpdateCodeByItemName(string itemName, string factoryCode, string eCode)
        {
            itemName = itemName.Trim();
            factoryCode = factoryCode.Trim();
            eCode = eCode.Trim();

            DBContext _db = new DBContext();
            DataTable dt = _db.LoadDataTable("TABLE_OF_CONTENT_SETTING", new[] { "ItemName" }, new[] { itemName }, new[] { "ItemName", "EEEECode", "FactoryCode" });
            if(dt.Rows.Count <= 0)
            {
                throw new Exception($"Chưa cài đặt table of content itemName {itemName} trên hệ thống!");
            }
            foreach (DataRow row in dt.Rows)
            {
                row["FactoryCode"] = factoryCode;
                row["EEEECode"] = eCode;
            }
            return _db.Update("TABLE_OF_CONTENT_SETTING", dt, new[] { "ItemName" });
        }
    }
}
