using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OK2SHIP_SMT.Repositories;

namespace OK2SHIP_SMT.Services
{
    public class IQCPeelingTest
    {
        public string _ITEMCODE_E = "";
        public Dictionary<string, DataTable> _DATA = new Dictionary<string, DataTable>();
        private DBContext _dbContext = new DBContext();

        public IQCPeelingTest()
        {
        }

        public string GetData(string location, string itemNameUser, string partNameUser, string lotNoUser,
            string invoiceUser)
        {
            location = location.Trim();
            // 1. Kiểm tra user
            if (!UserSession.Instance.IsLoggedIn)
            {
                throw new Exception("Login Please!");
            }

            // 2. Kiểm tra location xác thực
            if (string.IsNullOrEmpty(location))
            {
                throw new Exception("Enter your location!");
            }

            string msg = GetInfor(location, out string itemName, out string partName, out string lotNo,
                out string Invoice, out string tape, out string makerName);
            if (!(itemNameUser.Trim() == itemName && partNameUser.Trim() == partName && lotNoUser.Trim() == lotNo &&
                  invoiceUser.Trim() == Invoice))
            {
                throw new Exception("ItemCode lotNo not match!");
            }


            string[] listFolder = FileFolderRepository.GetSubFolders(location);

            foreach (string file in listFolder)
            {
                string itemFolderName = new DirectoryInfo(file).Name;
                string category = itemFolderName.Split('-')[0].Trim().ToUpper(); // psa liner\
                //LINER
                if (category.Equals("LINER"))
                {
                    HandleFolderData(category, file, tape, makerName);
                }

                //PSA
                if (category.Equals("PSA"))
                {
                    //psa-a, psa-b
                    string[] files = FileFolderRepository.GetSubFolders(file);
                    if (files.Length > 0)
                    {
                        foreach (string fileItem in files)
                        {
                            category = new DirectoryInfo(fileItem).Name.ToUpper().Replace(" ", "");
                            if (category.Split('-')[0].Trim().ToUpper() == "PSA")
                            {
                                HandleFolderData(category, fileItem, tape, makerName);
                            }
                        }
                    }
                }
            }

            return "Get Data succesfully!";
        }

        private DataTable getDataTable_data()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Picture", typeof(byte[]));
            dt.Columns.Add("Graph", typeof(byte[]));
            dt.Columns.Add("Average", typeof(string));
            dt.Columns.Add("Judgement", typeof(string));
            return dt;
        }

        private KeyValuePair<byte[], string> HandleFile(string location)
        {
            string value = "";
            KeyValuePair<byte[], string> res = new KeyValuePair<byte[], string>();
            using (ExcelPackage package = ExportProcess.openPackage(location))
            {
                using (ExcelWorksheet worksheet = package.Workbook.Worksheets[0])
                {
                    IDictionary<string, string> dic = ExportProcess.FindAddressByText(worksheet, new[] { "Average" });
                    try
                    {
                        if (dic.TryGetValue("Average", out string address))
                        {
                            int i = 1;

                            while (i <= 20)
                            {
                                string addressZ = worksheet.Cells[worksheet.Cells[address].End.Row,
                                        worksheet.Cells[address].End.Column + i]
                                    .Address;
                                var z = worksheet.Cells[addressZ].Value;
                                if (z != null)
                                {
                                    value = z.ToString().Trim();
                                    if (!string.IsNullOrEmpty(value))
                                    {
                                        if (double.TryParse(value, out double ressz))
                                        {
                                            break;
                                        }
                                    }
                                }

                                i++;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                    finally
                    {
                        byte[] image = (worksheet.Drawings["Picture 1"] as ExcelPicture).Image.ImageBytes;
                        res = new KeyValuePair<byte[], string>(image, value);
                    }
                }
            }

            return res;
        }

        private void HandleFolderData(string category, string location, string tape, string makerName)
        {
            DataTable dt = getDataTable_data();
            string[] list = FileFolderRepository.GetFileByExtension(location, "xlsx");
            list = list.OrderBy(filePath =>
            {
                // Lấy tên file kèm phần mở rộng (ví dụ: "1.xlsx")
                string fileName = Path.GetFileNameWithoutExtension(filePath).Trim();

                // Chuyển đổi tên file thành số nguyên để sắp xếp chính xác
                int.TryParse(fileName, out int fileNumber);
                return fileNumber;
            }).ToArray();
            List<KeyValuePair<Image, string>> images = FileFolderRepository.ListAllPictureInAFolder(location);
            int i = 0;
            while (i < list.Length && i < images.Count)
            {
                DataRow row = dt.NewRow();
                row["ID"] = i;
                row["Picture"] = TDMK_ImageConverter.ImageToByteArray(images[i].Key, ImageFormat.Jpeg);

                KeyValuePair<byte[], string> data = HandleFile(list[i]);
                row["Graph"] = data.Key;
                row["Average"] = data.Value;

                dt.Rows.Add(row);
                i++;
            }

            _DATA.Add($"{category}_{makerName}_TAPE {tape}", dt);
            //Debugger.Break();
        }

        public string GetInfor(string location, out string itemName, out string partName, out string lotNo,
            out string Invoice, out string tape, out string makername)
        {
            location = location.Trim();
            string msg = "";
            itemName = "";
            partName = "";
            lotNo = "";
            Invoice = "";
            tape = "";
            makername = "";
            if (string.IsNullOrEmpty(location))
            {
                return "Enter your location!";
            }

            string folderName = new DirectoryInfo(location).Name;

            string[] splits = folderName.Split('_');
            itemName = splits[0];
            partName = splits[1];
            lotNo = splits[2];
            Invoice = splits[3].Split('-')[0];
            tape = partName.Split('-')[1];
            makername = partName.Split('-')[3];

            return msg;
        }

        private string SQL_NAME = "IQC_PEELING_TEST_NAS";

        public void Save(string itemName, string partName, string lotNo, string invoice, bool skipDB = false)
        {
            DataTable dt = new DataTable();

            if (skipDB == false)
            {
                dt = _dbContext.LoadDataTable(SQL_NAME, new[] { "ItemName", "PartName", "LotNo", "Invoice" },
                    new[] { itemName, partName, lotNo, invoice });
                if (dt.Rows.Count > 0)
                {
                    throw new DataException("Dữ liệu đã tồn tại có tiếp tục");
                }
            }
            else
            {
                dt = _dbContext.GetTableStructure(SQL_NAME);
            }

            List<string> list = new List<string>();
            List<string> listAddress = new List<string>();
            foreach (string item in _DATA.Keys)
            {
                NasRepository _nasRepository = new NasRepository();
                string location = _nasRepository.HandleImageDataTable(_DATA[item], SQL_NAME, $"{itemName}-{partName}",
                    $"{lotNo}-{invoice}\\{item}");
                string dataItem = ConverterService.DataTableToJson(_DATA[item]);
                list.Add($"{item}\u200D{dataItem}");
                listAddress.Add($"{item}\u200D{location}");
            }

            string data = String.Join("\u200B", list);
            string locationData = String.Join("\u200B", listAddress);
            DataRow row = dt.NewRow();

            row["ItemName"] = itemName;
            row["PartName"] = partName;
            row["LotNo"] = lotNo;
            row["Invoice"] = invoice;
            row["Data"] = data;
            row["LocationImage"] = locationData;
            dt.Rows.Add(row);

            int res = _dbContext.BuckDataTable(dt, SQL_NAME, new[] { "ItemName", "PartName", "LotNo", "Invoice" }, null,
                "ID");
            _DATA.Clear();
            if (res > 0)
            {
                throw new Exception("Success");
            }
        }

        public void Load(string itemName, string partName, string lotNo, string invoice)
        {
            _DATA.Clear();
            if (string.IsNullOrEmpty(itemName) || string.IsNullOrEmpty(lotNo) || string.IsNullOrEmpty(invoice) ||
                string.IsNullOrEmpty(partName))
            {
                throw new Exception("Please enter information!");
            }

            DataTable dt = _dbContext.LoadDataTable(SQL_NAME, new[] { "ItemName", "PartName", "LotNo", "Invoice" },
                new[] { itemName, partName, lotNo, invoice });
            if (dt.Rows.Count <= 0)
            {
                throw new Exception("Data is not valid");
            }

            _ITEMCODE_E = dt.Rows[0]["ItemCodeExport"].ToString();
            string data = dt.Rows[0]["Data"].ToString();
            string[] splits = data.Split('\u200B');
            foreach (string item in splits)
            {
                string[] value = item.Split('\u200D');
                DataTable dtZ = ConverterService.JsonToDataTable(value[1]);
                _DATA.Add(value[0], dtZ);
            }

            NasRepository _nasRepository = new NasRepository();
            string locationIMG = dt.Rows[0]["LocationImage"].ToString();
            string[] splitI = locationIMG.Split('\u200B');
            foreach (string item in splitI)
            {
                string[] value = item.Split('\u200D');
                if (_DATA.TryGetValue(value[0], out DataTable dtZ))
                {
                    _nasRepository.MergeDataTable(dtZ, SQL_NAME, $"{itemName}-{partName}",
                        $"{lotNo}-{invoice}\\{item}",
                        value[1]);
                }
            }
        }

        public void Export(string itemName, string partName, string lotNo, string invoice)
        {
            Load(itemName, partName, lotNo, invoice);
            ExportProcess process = new ExportProcess();
            using (ExcelPackage package = process.FindFormatWithItemCode(_ITEMCODE_E))
            {
                using (ExcelWorksheet ws_1 = process.FindSheet(package, "IQC PSA peeling (Coupon)"))
                using (ExcelWorksheet ws_2 = process.FindSheet(package, "IQC Liner peeling (Coupon)"))
                {
                    foreach (string key in _DATA.Keys)
                    {
                        if (key.Contains("LINER"))
                        {
                            ExportLinerPSA(ws_2, key);
                        }

                        if (key.Contains("PSA-B"))
                        {
                            ExportLinerPSA(ws_1, key);
                        }
                    }


                    process.SaveExcelWorksheet(package, "IQC PSA peeling (Coupon):IQC Liner peeling (Coupon)",
                        $"{itemName.Trim()}-{lotNo.Trim()}");
                }
            }
        }

        private void ExportLinerPSA(ExcelWorksheet workSheet, string type)
        {
            string keyWord = "";
            if (type.ToUpper().Contains("PSA"))
            {
                keyWord = "psa_name supplier_Tape";
            }

            if (type.ToUpper().Contains("LINER"))
            {
                keyWord = "Liner_name supplier_Tape";
            }

            IDictionary<string, string> dic =
                ExportProcess.FindAddressByText(workSheet, new[] { "Sample 5", "Average Force", keyWord });

            if (dic.TryGetValue(keyWord, out string value))
            {
                WriteData(workSheet, value, type, _DATA[type]);
            }
        }

        private void WriteData(ExcelWorksheet worksheet, string address, string key, DataTable value)
        {
            worksheet.Cells[address].Value = key;
            int counting = 5, i = 0;
            while (counting > i)
            {
                address = ExportProcess.AddColumn(address, 1);
                if (worksheet.Cells[address].Value != null)
                {
                    string addressZ = ExportProcess.AddRow(address, 1);
                    ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[addressZ],
                        (byte[])value.Rows[i]["Picture"], $"Picture_{key}_{i}");
                    addressZ = ExportProcess.AddRow(address, 2);
                    ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[addressZ],
                        (byte[])value.Rows[i]["Graph"], $"Graph_{key}_{i}");
                    addressZ = ExportProcess.AddRow(address, 3);
                    if (double.TryParse(value.Rows[i]["Average"].ToString(), out double average))
                    {
                        worksheet.Cells[addressZ].Value = average;
                    }

                    addressZ = ExportProcess.AddRow(address, 4);
                    worksheet.Cells[addressZ].Value = value.Rows[i]["Judgement"];
                    i++;
                }
            }
        }

        public void SaveItem(string itemName, string partName, string lotNo, string invoice, string sitemcode)
        {
            DataTable dt = _dbContext.LoadDataTable(SQL_NAME, new[] { "ItemName", "PartName", "LotNo", "Invoice" },
                new[] { itemName, partName, lotNo, invoice });

            dt.Rows[0]["ItemCodeExport"] = sitemcode;
            int res = _dbContext.BuckDataTable(dt, SQL_NAME, new[] { "ItemName", "PartName", "LotNo", "Invoice" }, null, "ID");
            if (res > 0)
            {
                throw new Exception("Success");
            }
        }
    }
}