using OfficeOpenXml;
using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Security.Authentication;
using System.Text;
using System.Threading.Tasks;

namespace OK2SHIP_SMT.Services
{
    public class PackagingService
    {
        private DBContext _dbContext = new DBContext();
        private string _NAMETABLE = "PACKAGING";
        public DataTable CreateDataImage()
        {
            return _dbContext.GetTableStructure(_NAMETABLE);
        }
        public static DataTable createDataVote()
        {
            string[] name = new[] { "1. Any Tray deformation",
                                    "2. Any flex damage ",
                                    "3. Any Flex out of tray/package ",
                                    "4. Any component/glue damage, crack..",
                                    "5. Any Flex fail of function test",
                                    "6. Any liner drop, line partially peel" };
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Name");
            dataTable.Columns.Add("Result");
            foreach (string item in name)
            {
                DataRow row = dataTable.NewRow();
                row["Name"] = item;
                dataTable.Rows.Add(row);
            }
            return dataTable;
        }
        public DataTable getStructor()
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("ID", typeof(int));
            dataTable.Columns.Add("ShippingTo", typeof(string));
            dataTable.Columns.Add("TrayCode", typeof(string));
            dataTable.Columns.Add("ItemCode", typeof(string));
            dataTable.Columns.Add("FlexTop", typeof(Image));
            dataTable.Columns.Add("FlexBottom", typeof(Image));
            dataTable.Columns.Add("Tray", typeof(Image));
            dataTable.Columns.Add("TrayAL", typeof(Image));
            dataTable.Columns.Add("ALBag", typeof(Image));
            dataTable.Columns.Add("CartonBox", typeof(Image));
            dataTable.Columns.Add("Data", typeof(string));
            dataTable.Columns.Add("LocationIMG", typeof(string));

            return dataTable;
        }
        public int Save(string itemCode, Dictionary<string, KeyValuePair<Dictionary<string, Image>, string>> keyValuePairs, bool prime)
        {
            itemCode = itemCode.Trim();
            if (string.IsNullOrEmpty(itemCode))
            {
                throw new ArgumentException("Item code cannot be null or empty.");
            }
            if (keyValuePairs == null || keyValuePairs.Count == 0)
            {
                throw new ArgumentException("KeyValuePairs cannot be null or empty.");
            }
            if (UserSession.Instance.IsLoggedIn == false)
            {
                throw new AuthenticationException("User is not logged in.");
            }
            DataTable dataTable = new DataTable();
            if (prime == false)
            {
                dataTable = _dbContext.LoadDataTable($"{_NAMETABLE}_LOGFILE", new[] { "ItemCode" }, new[] { itemCode });
                if (dataTable.Rows.Count > 0)
                {
                    throw new Exception($"1234 - Item code {itemCode} already exists in the database.");
                }
            }
            dataTable = getStructor();
            foreach (string item in keyValuePairs.Keys)
            {
                string shippingTo = item.Split('_')[0].Trim();
                string trayCode = item.Split('_')[1].Trim();
                DataRow row = dataTable.NewRow();
                row["ItemCode"] = itemCode;
                row["ShippingTo"] = shippingTo;
                row["TrayCode"] = trayCode;
                row["Data"] = keyValuePairs[item].Value;

                foreach (string item1 in keyValuePairs[item].Key.Keys)
                {
                    string key = "";
                    switch (item1)
                    {
                        case "1":
                            key = "FlexTop";
                            break;
                        case "2":
                            key = "FlexBottom";
                            break;
                        case "3":
                            key = "Tray";
                            break;
                        case "4":
                            key = "TrayAL";
                            break;
                        case "5":
                            key = "ALBag";
                            break;
                        case "6":
                            key = "CartonBox";
                            break;
                    }

                    row[key] = keyValuePairs[item].Key[item1];
                }
                dataTable.Rows.Add(row);
            }
            _dbContext.DeleteData($"{_NAMETABLE}_LOGFILE", "ItemCode", new[] { itemCode });
            NasRepository nas = new NasRepository();
            string location = nas.HandleImageDataTable(dataTable, $"{_NAMETABLE}", itemCode, "");

            foreach (DataRow row in dataTable.Rows)
            {
                row["LocationIMG"] = location;
            }
            return _dbContext.BuckDataTable(dataTable, $"{_NAMETABLE}_LOGFILE", new[] { "ItemCode" }, null, "ID");
        }
        public DataTable Load(string itemCode)
        {
            itemCode = itemCode.Trim();
            if (string.IsNullOrEmpty(itemCode))
            {
                throw new ArgumentException("Item code cannot be null or empty.");
            }

            DataTable dataTable = _dbContext.LoadDataTable("PACKAGING_LOGFILE", new[] { "ItemCode" }, new[] { itemCode });
            if(dataTable.Rows.Count <= 0)
            {
                return dataTable;
            }
            NasRepository nas = new NasRepository();
            nas.MergeDataTable(dataTable, $"{_NAMETABLE}", itemCode, "", dataTable.Rows[0]["LocationIMG"].ToString());
            return dataTable;
        }
        public Dictionary<string, KeyValuePair<Dictionary<string, Image>, string>> LoadDictionary(string itemCode)
        {
            Dictionary<string, KeyValuePair<Dictionary<string, Image>, string>> result = new Dictionary<string, KeyValuePair<Dictionary<string, Image>, string>>();
            DataTable dataTable = Load(itemCode);
            if (dataTable.Rows.Count == 0)
            {
                throw new Exception($"Item code {itemCode} not found in the database.");
            }
            foreach (DataRow row in dataTable.Rows)
            {
                string shippingTo = row["ShippingTo"].ToString().Trim();
                string trayCode = row["TrayCode"].ToString().Trim();
                string data = row["Data"].ToString();
                Dictionary<string, Image> images = new Dictionary<string, Image>();
                try
                {
                    images.Add("1", (Image)row["FlexTop"]);
                }
                catch { }
                try
                {

                    images.Add("2", (Image)row["FlexBottom"]);
                }
                catch { }
                try
                {
                    images.Add("3", (Image)row["Tray"]);
                }
                catch
                {
                }
                try
                {
                    images.Add("4", (Image)row["TrayAL"]);

                }
                catch
                {

                }
                try
                {

                    images.Add("5", (Image)row["ALBag"]);
                }
                catch
                {

                }
                try
                {

                    images.Add("6", (Image)row["CartonBox"]);
                }
                catch
                {

                }
                result.Add($"{shippingTo}_{trayCode}", new KeyValuePair<Dictionary<string, Image>, string>(images, data));
            }
            return result;
        }



        public void ExportToExcel(string itemCode, bool prime)
        {
            DataTable dataTable = Load(itemCode);
            if (dataTable.Rows.Count == 0)
            {
                throw new Exception($"Item code {itemCode} not found in the database.");
            }
            ExportProcess exportProcess = new ExportProcess();
            using (ExcelPackage package = exportProcess.FindFormatProcess("PAKAGING", itemCode, ""))
            {
                using (ExcelWorksheet worksheet = exportProcess.FindSheet(package, "Packaging"))
                {
                    string[] name = new[] { "Packing Ship", "Picture", "6. Any liner drop" };
                    IDictionary<string, string> dic = ExportProcess.FindAddressByText(worksheet, name);

                    dic["Picture"] = dic["Picture"].Split('-').FirstOrDefault();
                    string startAddress = $"{dic["Packing Ship"]}:{worksheet.Cells[worksheet.Cells[dic["6. Any liner drop"]].End.Row, worksheet.Cells[dic["Picture"]].End.Column + 1]}";
                    string PasteAddress = dic["Packing Ship"];
                    for (int i = 0; i < dataTable.Rows.Count - 1; i++)
                    {
                        PasteAddress = ExportProcess.AddColumn(PasteAddress, 3);
                        ExportProcess.CopyColumn(worksheet, worksheet.Cells[startAddress], PasteAddress);
                    }
                    name = new[] { "Packing Ship", "Picture", "1. Any Tray deformation" };
                    dic = ExportProcess.FindAddressByText(worksheet, name, true);
                    dic.Add("Packing Ship" , ExportProcess.FindAddressByText(worksheet, new[] {"Packing Ship"}).First().Value);
                    for (int i = 0; i < dic["Picture"].Split('-').Count(); i++)
                    {

                        DataRow row = dataTable.Rows[i];
                        string address = dic["Packing Ship"].Split('-')[i];
                        string value = worksheet.Cells[address].Text;
                        value = value.Replace("{tray code}", row["TrayCode"].ToString().Trim());
                        value = value.Replace("{packing ship}", row["ShippingTo"].ToString().Trim());
                        worksheet.Cells[address].Value = value;
                        address = dic["Picture"].Split('-')[i];
                        while (true)
                        {
                            address = ExportProcess.AddColumn(ExportProcess.AddRow(address, 1), -1);
                            value = worksheet.Cells[address].Text;
                            if (string.IsNullOrEmpty(value))
                            {
                                break;
                            }

                            address = ExportProcess.AddColumn(address, 1);
                            if (value.Contains("Flex") && value.Contains("Top") && value.Contains("side"))
                            {
                                ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[address], TDMK_ImageConverter.ImageToByteArray((Image)row["FlexTop"], ImageFormat.Jpeg), $"FlexTopImage{i}");
                            }
                            else if (value.Contains("Flex") && value.Contains("Bottom") && value.Contains("side"))
                            {
                                ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[address], TDMK_ImageConverter.ImageToByteArray((Image)row["FlexBottom"], ImageFormat.Jpeg), $"FlexBottomImage{i}");
                            }
                            else if (value.Contains("Tray") && !value.Contains("AL"))
                            {
                                ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[address], TDMK_ImageConverter.ImageToByteArray((Image)row["Tray"], ImageFormat.Jpeg), $"Tray{i}");
                            }
                            else if (value.Contains("Tray") && value.Contains("AL") && value.Contains("bag"))
                            {
                                ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[address], TDMK_ImageConverter.ImageToByteArray((Image)row["TrayAL"], ImageFormat.Jpeg), $"TrayAL{i}");
                                ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[address], TDMK_ImageConverter.ImageToByteArray((Image)row["ALBag"], ImageFormat.Jpeg), $"ALBAG{i}");
                            }
                            else if (value.Contains("Carton Box"))
                            {
                                ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[address], TDMK_ImageConverter.ImageToByteArray((Image)row["CartonBox"], ImageFormat.Jpeg), $"CartonBox{i}");
                            }
                            else { break; }


                        }
                        DataTable JsonZ = ConverterService.JsonToDataTable(row["Data"].ToString());
                        int iz = 0;
                        address = ExportProcess.AddColumn(dic["1. Any Tray deformation"].Split('-')[i], 1);
                        while (true)
                        {
                            var zs = worksheet.Cells[address].Value;
                            if (zs == null || string.IsNullOrEmpty(zs.ToString()))
                            {
                                break;
                            }
                            worksheet.Cells[address].Value = JsonZ.Rows[iz]["Result"];
                            address = ExportProcess.AddRow(address, 1);
                        }

                    }
                    exportProcess.SaveExcelWorksheet(package, "Packaging", $"{itemCode}-packaging");
                }
            }
        }

        public Dictionary<string, KeyValuePair<Dictionary<string, Image>, string>> GetData(string location, string itemCode)
        {
            Dictionary<string, KeyValuePair<Dictionary<string, Image>, string>> _diz = new Dictionary<string, KeyValuePair<Dictionary<string, Image>, string>>();
            itemCode = itemCode.Trim();
            string foldername = FileFolderRepository.GetFolderName(location);
            if (!foldername.Equals(itemCode))
            {
                throw new Exception("Không match itemcode");
            }
            string[] st = FileFolderRepository.GetSubFolders(location);
            foreach (string item in st)
            {
                string z = FileFolderRepository.GetFolderName(item).Split('_')[0].Trim();
                string z2 = FileFolderRepository.GetFolderName(item).Split('_')[1].Trim();
                IList<KeyValuePair<Image, string>> list = FileFolderRepository.ListAllPictureInAFolder(item);
                Dictionary<string, Image> dic = new Dictionary<string, Image>();
                foreach (var image in list)
                {
                    if (dic.Keys.Count() > 5)
                    {
                        break;
                    }
                    dic.Add((dic.Keys.Count() + 1).ToString(), image.Key);
                }
                string zq = "[{\"Name\":\"1. Any Tray deformation\",\"Result\":\"\"},{\"Name\":\"2. Any flex damage \",\"Result\":\"\"},{\"Name\":\"3. Any Flex out of tray/package \",\"Result\":\"\"},{\"Name\":\"4. Any component/glue damage, crack..\",\"Result\":\"\"},{\"Name\":\"5. Any Flex fail of function test\",\"Result\":\"\"},{\"Name\":\"6. Any liner drop, line partially peel\",\"Result\":\"\"}]";
                KeyValuePair<Dictionary<string, Image>, string> values = new KeyValuePair<Dictionary<string, Image>, string>(dic, zq);
                _diz.Add($"{z}_{z2}", values);
            }
            return _diz;
        }
    }
}
