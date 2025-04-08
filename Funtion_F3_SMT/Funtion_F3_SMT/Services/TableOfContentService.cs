using Newtonsoft.Json;
using OfficeOpenXml;
using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.Services
{

    public class TableOfContentService
    {
        private DBContext _dBContext = new DBContext();
        private const string _NAMETABLE = "TABLE_OF_CONTENT";
        public TableOfContentService()
        {

        }

        public bool checkItemCodeLotNo(string itemcode, string lotno)
        {
            if (string.IsNullOrEmpty(itemcode) || string.IsNullOrEmpty(lotno))
            {
                return false;
            }
            return true;
        }
        public DataTable getAllTOC(string[] filter, string[] value, string[] selectedCol = null, int pageNumber = 1, int pageSize = 20, string[] orderBy = null)
        {
            DataTable dataTable = new DataTable();
            if (selectedCol == null)
            {
                selectedCol = new string[] { "ID", "ItemCode", "LotNo", "BuildDate", "SendDate", "DeliveryQuatity", "ShippingTo", "ContentTable" };
            }
            dataTable = _dBContext.LoadDataTableOfPath(_NAMETABLE, filter, value, selectedCol, pageNumber, pageSize, orderBy);
            return dataTable;
        }

        public DataTable getTableOfContent(string itemCode, string lotNo)
        {
            DataTable dataTable = new DataTable();
            dataTable = _dBContext.LoadDataTableOfPath(_NAMETABLE, new string[] { "ItemCode", "LotNo" }, new string[] { itemCode, lotNo }, new string[] { "Content" });

            return dataTable;
        }
        public DataTable MakeTable(string date = "")
        {
            DataTable dataTable = new DataTable();
            string[] textCol = new string[] { "Submission Items", "Test", "Reference Document #", "# of Rejects", "Total Qty", "Target date Request", "Target date Submission", "Status" };
            string[] nameCol = new string[] { "subItem", "test", "refDoc", "rejects", "total", "request", "sub", "status" };
            for (int i = 0; i < textCol.Length; i++)
            {
                dataTable.Columns.Add(textCol[i], typeof(string));
            }

            string[] textRow = new string[] { "Declaration", "Deviation summary", "Flex Assembly", "FAI", "CPK", "Mic peeling force test", "Connector shear off test", "Peeling test for BGA component", "Shear Force test for BGA Component", "B2B CONNECTOR PEELING TEST ", "Clip Peeling test ", "E75 IO Mating/Unmating test", "Thermal Cycling", "Heat Soak and Recovery", "Thermal stress", "Hotbar loop test", "Thermal shock", "Flex bending test ", "Bending after thermal cycling", "Bending after heat soak ", "Click Ratio", "Wetting contact angle", "ACF flatness", "ACF peel test", "Surface roughness", "Bar code checking (Grade and Rule)", "Package Drop Test /Vibration Result", "Process flow" };
            string[] textRow1 = new string[] { "Declaration", "Deviation", "< Yeild Bridge >", "<Measument>", "<Measument>", "<OQC Test>", "<OQC Test>", "<OQC Test>", "<OQC Test>", "<OQC Test>", "<OQC Test>", "< ORT Test - Back End >", "< ORT Test - Back End >", "< ORT Test - Back End >", "< ORT Test - Back End >", "< ORT Test - Back End >", "< ORT Test - Back End >", "< ORT Test - Back End >", "< ORT Test - Back End >", "< ORT Test - Back End >", "Tact switch", "ACF", "ACF", "ACF", "ACF", "Bar code", "< Pakaging >", "Process Flow" };
            string[] textRow2 = new string[] { "N/A", "N/A", "N/A", "MCO", "MCO", "Flex CPP", "Flex CPP", "Flex CPP", "Flex CPP", "Flex CPP", "CPP", "NA", "080-03910", "080-03910", "080-03910", "080-03910", "080-03910", "080-03910", "080-03910", "080-03910", "MCO", "080-03911", "080-03911", "080-03911", "080-03911", "080-03910", "080-03920", "CPP" };
            string[] textRow3 = new string[] { "Prior to ship", "Prior to ship", "Prior to ship", "Prior to ship", "Prior to ship", "Prior to ship", "Prior to ship", "Prior to ship", "Prior to ship", "Prior to ship", "Prior to ship", "30days max after shipping", "30days max after shipping", "30days max after shipping", "30days max after shipping", "5days max after shipping", "30days max after shipping", "Prior to ship", "30days max after shipping", "30days max after shipping", "Prior to ship", "Prior to ship", "Prior to ship", "Prior to ship", "Prior to ship", "Prior to ship", "Prior to ship", "Prior to ship" };
            string[] textRow4 = new string[] { "", "", "NA", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "NA", "NA" };
            string[] textRow5 = new string[] { "", "", "NA", "32", "32", "10", "10", "10", "10", "10", "10", "45", "45", "45", "5", "10", "45", "20", "20", "20", "32", "5", "5", "5", "32", "5", "NA", "NA" };
            string[] naList = new[] { "Click Ratio", "E75 IO Mating/Unmating test", "Shear Force test for BGA Component", "Mic peeling force test" };
            string[] naList2 = new[] { "Thermal stress","Hotbar loop test" };
            for (int i = 0; i < textRow.Length; i++)
            {
                var row = dataTable.NewRow();
                row["Test"] = textRow[i];
                row["Submission Items"] = textRow1[i];
                row["Reference Document #"] = textRow2[i];
                row["# of Rejects"] = textRow4[i];
                row["Total Qty"] = textRow5[i];
                row["Target date Request"] = textRow3[i];
                if (naList.Contains(row["Test"].ToString().Trim()))
                {
                    row["Target date Submission"] = "NA";
                }
                if (naList2.Contains(row["Test"].ToString().Trim()))
                {
                    row["Target date Submission"] = "N/A(No Hot Bar)";
                }
                if(row["Target date Request"].Equals("Prior to ship") && string.IsNullOrEmpty(row["Target date Submission"].ToString()))
                {
                    row["Target date Submission"] = date;
                }
                dataTable.Rows.Add(row);
            }
            return dataTable;
        }

        public int SaveData(DataTable table, DataTable content, bool prime)
        {
            var row = table.Rows[0];
            string itemCode = row["ItemCode"].ToString();
            string lotNo = row["LotNo"].ToString();
            if (!checkItemCodeLotNo(itemCode, lotNo))
            {
                throw new Exception("Hãy nhập ItemCode và lotNo");
            }
            if (!prime)
            {
                DataTable dt = _dBContext.LoadDataTable(_NAMETABLE, new string[] { "ItemCode", "LotNo" }, new string[] { itemCode, lotNo }, new string[] { "ItemCode" });
                if (dt.Rows.Count > 0)
                {
                    throw new Exception("1412 - ItemCode LotNo đã tồn tại bạn có muốn ghi đè!");
                }

            }
            string json = ConverterService.DataTableToJson(content);
            row["ContentTable"] = json;
            int changeRow = _dBContext.BuckDataTable(table, _NAMETABLE, new string[] { "ItemCode", "LotNo" }, null, "ID");
            return changeRow;
        }

        public DataTable getDataTable()
        {
            return _dBContext.GetTableStructure(_NAMETABLE);

        }
        public DateTime? ConvertDate(string str)
        {
            string format = "M/d/yyyy h:mm:ss tt";
            if (DateTime.TryParseExact(str, format, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime parsedDateTime))
            {
                return parsedDateTime;
            }
            else
            {
                return null;
            }
        }
        public void Export(string itemCode, string lotno)
        {
            itemCode = itemCode.Trim();
            lotno = lotno.Trim();
            DataTable dt = _dBContext.LoadDataTable(_NAMETABLE, new string[] { "ItemCode", "LotNo" }, new string[] { itemCode, lotno });
            DataTable dt2 = _dBContext.LoadDataTable("TABLE_OF_CONTENT_SETTING", new string[] { "ItemCode" }, new string[] { itemCode });
            if (dt.Rows.Count <= 0)
            {
                throw new Exception($"{itemCode} - {lotno} chưa tồn tại dữ liệu");
            }
            ExportProcess exportProcess = new ExportProcess();
            using (ExcelPackage ex = exportProcess.FindFormatProcess("Table of Contents", itemCode, lotno))
            {
                if (ex == null)
                {
                    throw new Exception($"{itemCode} - {lotno} chưa tồn tại format dữ liệu");
                }
                using (ExcelWorksheet workSheet = exportProcess.FindSheet(ex, "Table of Contents"))
                {
                    string[] path1 = new string[] { "Program Name:", "MCO & Revision:", "ODB++ & Revision", "Build Config:", "Lot #:", "Ok2Build date:", "Delivery Qty:", "Ok2send date:", "Shipping to", "Shipping form:", "X-out Rate:", "EEEE Code" };
                    string[] path2 = new string[] { "Submission", "Declaration", "Deviation summary", "Flex Assembly", "FAI", "CPK", "Mic peeling force test", "Connector shear off test", "Peeling test for BGA component", "Shear Force test for BGA Component", "B2B CONNECTOR PEELING TEST ", "Clip Peeling test ", "E75 IO Mating/Unmating test", "Thermal Cycling", "Heat Soak and Recovery", "Thermal stress", "Hotbar loop test", "Thermal shock", "Flex bending test ", "Bending after thermal cycling", "Bending after heat soak ", "Click Ratio", "Wetting contact angle", "ACF flatness", "ACF peel test", "Surface roughness", "Bar code checking (Grade and Rule)", "Package Drop Test /Vibration Result", "Process flow" };
                    IDictionary<string, string> addressHeader = ExportProcess.FindAddressByText(workSheet, path1.Concat(path2).ToArray(), true);
                    var row = dt.Rows[0];
                    var row2 = dt2.Rows[0];
                    if (addressHeader.TryGetValue("Program Name:", out string address))
                    {
                        address = ExportProcess.AddColumn(address, 1);
                        workSheet.Cells[address].Value = row2["ProgramName"];
                    }

                    if (addressHeader.TryGetValue("MCO & Revision:", out address))
                    {
                        address = ExportProcess.AddColumn(address, 1);
                        workSheet.Cells[address].Value = row2["MCORevision"];
                    }

                    if (addressHeader.TryGetValue("ODB++ & Revision", out address))
                    {
                        address = ExportProcess.AddColumn(address, 1);
                        workSheet.Cells[address].Value = row2["ODBRevision"].ToString().Trim();
                    }
                    if (addressHeader.TryGetValue("Build Config:", out address))
                    {
                        address = ExportProcess.AddColumn(address, 1);
                        workSheet.Cells[address].Value = row2["Build"].ToString().Trim();
                    }
                    if (addressHeader.TryGetValue("Lot #:", out address))
                    {
                        address = ExportProcess.AddColumn(address, 1);
                        workSheet.Cells[address].Value = row["ItemCode"].ToString().Trim() + "-" + row["LotNo"].ToString().Trim();
                    }
                    if (addressHeader.TryGetValue("Ok2Build date:", out address))
                    {
                        address = ExportProcess.AddColumn(address, 1);
                        DateTime? time = ConvertDate(row["BuildDate"].ToString().Trim());
                        if (time != null)
                        {
                            workSheet.Cells[address].Value = ((DateTime)time).ToString("d-MMM");
                        }
                    }
                    if (addressHeader.TryGetValue("Ok2send date:", out address))
                    {
                        address = ExportProcess.AddColumn(address, 1);
                        DateTime? time = ConvertDate(row["SendDate"].ToString().Trim());
                        if (time != null)
                        {
                            workSheet.Cells[address].Value = ((DateTime)time).ToString("d-MMM");
                        }
                    }

                    if (addressHeader.TryGetValue("Delivery Qty:", out address))
                    {
                        address = ExportProcess.AddColumn(address, 1);
                        workSheet.Cells[address].Value = row["DeliveryQuatity"].ToString().Trim();
                    }
                    if (addressHeader.TryGetValue("Shipping to", out address))
                    {
                        address = ExportProcess.AddColumn(address, 1);
                        workSheet.Cells[address].Value = row["ShippingTo"].ToString().Trim();
                    }
                    if (addressHeader.TryGetValue("Shipping form:", out address))
                    {
                        address = ExportProcess.AddColumn(address, 1);
                        workSheet.Cells[address].Value = row2["ShippingFrom"].ToString().Trim();
                    }
                    if (addressHeader.TryGetValue("X-out Rate:", out address))
                    {
                        address = ExportProcess.AddColumn(address, 1);
                        workSheet.Cells[address].Value = row2["XOUTRate"].ToString().Trim();
                    }
                    if (addressHeader.TryGetValue("EEEE Code", out address))
                    {
                        address = ExportProcess.AddColumn(address, 1);
                        //workSheet.Cells[address].Value = row2["EEEE-Code"].ToString().Trim();
                    }
                    // Gắn vào table
                    // 1. Lấy dữ liệu content
                    string content = row["ContentTable"].ToString();
                    // 2. convert to dataTable
                    DataTable dt_Content = ConverterService.JsonToDataTable(content);

                    if (addressHeader.TryGetValue("Submission", out string addressCol))
                    {
                        string addressRow;
                        foreach (DataRow item in dt_Content.Rows)
                        {
                            try
                            {
                                //Debugger.Break();
                                string str_search = item["Test"].ToString();
                                if (addressHeader.TryGetValue(str_search, out addressRow))
                                {
                                    if (str_search.Equals("Declaration"))
                                    {
                                        address = workSheet.Cells[workSheet.Cells[addressRow.Split('-')[1]].Start.Row, workSheet.Cells[addressCol].Start.Column].Address;
                                    }
                                    else
                                    {
                                        address = workSheet.Cells[workSheet.Cells[addressRow].Start.Row, workSheet.Cells[addressCol].Start.Column].Address;
                                    }
                                    workSheet.Cells[address].Value = item["Target date Submission"];
                                    workSheet.Cells[ExportProcess.AddColumn(address, 1)].Value = item["Status"];

                                }
                            }
                            catch (Exception exz)
                            {
                                Debugger.Break();
                            }

                        }
                    }
                    exportProcess.SaveExcelWorksheet(ex, "Table of Contents", $"{itemCode}-{lotno}");

                }

            }
            //throw new Exception("Export Thành công!");
        }

        public DataTable GetDataFormFile(string location)
        {
            location = location.Trim();
            if (!File.Exists(location))
            {
                throw new Exception("Đường dẫn File không tồn tại");
            }
            Dictionary<string, string> keyValuePairs = new Dictionary<string, string>();
            DataTable dt = _dBContext.GetTableStructure($"{_NAMETABLE}_SETTING");
            int id = 1;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(location))
            {
                using (ExcelWorksheet workSheet = package.Workbook.Worksheets[0])
                {
                    string[] header2 = new[] { "Build" };
                    string[] header = new[] { "Item", "Program Name", "Item Code", "MCO & Revision", "ODB++ & Revision", "Build" };
                    IDictionary<string, string> dic = ExportProcess.FindAddressByText(workSheet, header.Concat(header2).ToArray(), false);
                    dic.Add("Item Code", dic["Item"].Split('-')[1]);
                    dic["Item"] = dic["Item"].Split('-')[0];
                    dic["Build"] = dic["Build"].Split('-')[1];
                    string addressRow = ExportProcess.AddRow(dic["Item"], 1);

                    while (!string.IsNullOrEmpty(workSheet.Cells[addressRow].Text.Trim()))
                    {

                        DataRow dr = dt.NewRow();
                        string address = DictionaryService.GetValueOrNull(dic, "Item");
                        dr["ItemName"] = workSheet.Cells[workSheet.Cells[addressRow].Start.Row, workSheet.Cells[address].Start.Column].Text.Split(' ')[0];
                        address = DictionaryService.GetValueOrNull(dic, "Program Name");
                        dr["ProgramName"] = workSheet.Cells[workSheet.Cells[addressRow].Start.Row, workSheet.Cells[address].Start.Column].Text;
                        address = DictionaryService.GetValueOrNull(dic, "MCO & Revision");
                        dr["MCORevision"] = workSheet.Cells[workSheet.Cells[addressRow].Start.Row, workSheet.Cells[address].Start.Column].Text;
                        address = DictionaryService.GetValueOrNull(dic, "ODB++ & Revision");
                        dr["ODBRevision"] = workSheet.Cells[workSheet.Cells[addressRow].Start.Row, workSheet.Cells[address].Start.Column].Text;
                        address = DictionaryService.GetValueOrNull(dic, "Build");
                        dr["Build"] = workSheet.Cells[workSheet.Cells[addressRow].Start.Row, workSheet.Cells[address].Start.Column].Text;
                        dr["ShippingFrom"] = "SEEV";
                        dr["XOUTRate"] = "-";

                        address = DictionaryService.GetValueOrNull(dic, "Item Code");
                        string[] itemCodeStr = workSheet.Cells[workSheet.Cells[addressRow].Start.Row, workSheet.Cells[address].Start.Column].Text.Replace("\n", "-").Replace(" ", "").Split('-');
                        foreach (var itemCode in itemCodeStr)
                        {
                            DataRow drs = ExportProcess.CloneDataRow(dr);
                            drs["ItemCode"] = itemCode.Trim();
                            drs["Id"] = id++;
                            dt.Rows.Add(drs);
                            // đảm bảo không có itemcode giống nhau trong cùng một sheet
                            keyValuePairs.Add(itemCode, "1");
                        }

                        addressRow = ExportProcess.AddRow(addressRow, 1);
                    }

                }
            }
            return dt;
        }


        public int SaveDataItemName(DataTable dataTable, bool prime)
        {
            if (!prime)
            {
                List<string> list = dataTable.AsEnumerable().Select(row => row.Field<string>("ItemCode")).ToList();
                // Kiểm tra xem có ItemCode nào đã tồn tại hay chưa
                string[] str = _dBContext.checkListIsExist(list.ToArray(), _NAMETABLE + "_SETTING", "ItemCode");
                // Nếu có trả về thông báo cho người dùng xác nhận
                if (str.Count() > 0)
                {
                    throw new Exception($"1234 - các item đã tồn tại: {string.Join("','", str)}");
                }
            }
            else
            {
                List<string> list = dataTable.AsEnumerable().Select(row => row.Field<string>("ItemName")).ToList();
                int res = _dBContext.DeleteData(_NAMETABLE + "_SETTING", "ItemName", list.ToArray());
            }

            // Nếu không thì lưu vào datatable
            return _dBContext.SaveDataTable(dataTable, _NAMETABLE + "_SETTING");
        }

        public List<string> GetListItemCodeByItemName(string itemName, DataTable dtz = null)
        {
            if (dtz != null)
            {
                return dtz.AsEnumerable().Where(rowz => rowz.Field<string>("ItemName") == itemName).Select(row => row.Field<string>("ItemCode")).ToList();
            }
            itemName = itemName.Trim();
            DataTable dt = _dBContext.LoadDataTable(_NAMETABLE + "_SETTING", new[] { "ItemName" }, new[] { itemName }, new[] { "ItemCode" });
            return dt.AsEnumerable().Select(row => row.Field<string>("ItemCode")).ToList();
        }
        public DataTable getDataTableByItemName(string itemName, DataTable dtz = null)
        {
            DataTable dt = new DataTable();
            itemName = itemName.Trim();
            if (string.IsNullOrEmpty(itemName))
            {
                throw new Exception("Hãy nhập itemName");
            }
            if (dtz != null)
            {

                foreach (DataRow row in dtz.Rows)
                {
                    if (row.Field<string>("itemName") == itemName)
                    {
                        DataRow newRow = ExportProcess.CloneDataRow(row);
                        dt.Rows.Add(newRow);
                    }
                }

            }
            else
            {

                dt = _dBContext.LoadDataTable(_NAMETABLE + "_SETTING", new[] { "ItemName" }, new[] { itemName });
            }
            int id = 1;
            foreach (DataRow row in dt.Rows)
            {
                row["Id"] = id++;
            }
            return dt;
        }
        public DataTable loadTOCSItemCode(string ItemCode)
        {
            DataTable dt = _dBContext.LoadDataTable(_NAMETABLE + "_SETTING", new[] { "ItemCode" }, new[] { ItemCode });
            return dt;
        }
    }
}
