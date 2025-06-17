using Export_FPCA_OK2ship_Auto_System.Libary;
using Export_FPCA_OK2ship_Auto_System.Repositories;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Remoting.Contexts;
using System.Runtime.Remoting.Messaging;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

namespace Export_FPCA_OK2ship_Auto_System.Services
{
    public class OQCB2BMatingUnmatting
    {
        private DBContext _context = new DBContext();

        public OQCB2BMatingUnmatting()
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

        }

        public DataTable ProcessRead(string location, string itemCode, string lotNo)
        {
            DataTable dataTable = new DataTable();
            string[] files = Directory.GetDirectories(location);
            string fileT1 = files.FirstOrDefault(x => new DirectoryInfo(x).Name.Equals("T1"));
            string[] fileCSVT1 = FileFolderRepository.ListAllFileInFolder(fileT1, ".xlsx").ToArray();

            Dictionary<string, IList<KeyValuePair<Image, string>>> listPicture = FileFolderRepository.ListAllPictureInFolder(location);
            dataTable = _context.GetTableStructure("OQC_B2B_Mating_Unmating");
            int i = 0;
            Dictionary<int, int> PrimeCheck = new Dictionary<int, int>();
            foreach (var item in listPicture)
            {
                string folderName = FileFolderRepository.GetFolderName(item.Key);
                if (folderName.Equals("T0") || folderName.Equals("T1") || folderName.Equals("T30"))
                {
                    foreach (var jtem in item.Value)
                    {
                        string nameFile = jtem.Value;
                        int id = getId(nameFile);
                        if (id <= 0)
                        {
                            throw new Exception("Lỗi đổi tên file ảnh thành id");
                        }
                        int value;

                        // nếu đã tồn tại row
                        if (PrimeCheck.TryGetValue(id, out value))
                        {
                            //chọn mục ảnh
                            if (nameFile.Contains("+"))
                            {
                                dataTable.Rows[value][folderName + "U"] = TDMK_ImageConverter.ImageToByteArray(jtem.Key, ImageFormat.Jpeg);
                            }
                            else
                            {
                                dataTable.Rows[value][folderName] = TDMK_ImageConverter.ImageToByteArray(jtem.Key, ImageFormat.Jpeg);
                            }
                        }
                        // nếu chưa tồn tại row
                        else
                        {
                            // cho vào mảng kiểm tra
                            PrimeCheck.Add(id, i);
                            // fill dữ liệu vào row
                            DataRow row = dataTable.NewRow();
                            row["ID"] = id;
                            row["ItemCode"] = itemCode;
                            row["lotNo"] = lotNo;
                            //chọn mục ảnh
                            if (nameFile.Contains("+"))
                            {
                                row[folderName + "U"] = TDMK_ImageConverter.ImageToByteArray(jtem.Key, ImageFormat.Jpeg);
                            }
                            else
                            {
                                row[folderName] = TDMK_ImageConverter.ImageToByteArray(jtem.Key, ImageFormat.Jpeg);
                            }
                            // thêm row vao datatable
                            row["Judgement"] = "NG";
                            dataTable.Rows.Add(row);
                            // Debugger.Break();
                            i++;
                        }
                    }
                }
            }
            if (fileCSVT1.Count() >= 0)
            {
                foreach (string item in fileCSVT1)
                {
                    string iStr = Path.GetFileNameWithoutExtension(item);
                    if (int.TryParse(iStr, out int num))
                    {
                        if (PrimeCheck.TryGetValue(num, out int index))
                        {
                            Image image;
                            Double force = ReadFileGraph(item, out image);
                            dataTable.Rows[index]["Graph"] = TDMK_ImageConverter.ImageToByteArray(image, ImageFormat.Jpeg);
                            dataTable.Rows[index]["Force"] = force;

                        }
                    }
                }
            }
            int iz = 1;
            foreach (DataRow item in dataTable.Rows)
            {
                item["Id"] = iz++;
            }
            //Debugger.Break();
            return dataTable;
        }

        private double ReadFileGraph(string locationFile, out Image image)
        {
            FileInfo file = new FileInfo(locationFile);
            image = null;

            using (ExcelPackage package = new ExcelPackage(file))
            {
                string[] colHeader = new string[] { "Max" };
                using (ExcelWorksheet worksheet = package.Workbook.Worksheets[0])
                {
                    ExcelDrawings Images = worksheet.Drawings;
                    using (var drawing = worksheet.Drawings[0])
                    {
                        if (drawing is ExcelPicture picture)
                        {
                            byte[] z = picture.Image.ImageBytes;
                            image = TDMK_ImageConverter.ByteArrayToImage(z);
                        }
                    }
                    IDictionary<string, string> dic = ExportProcess.FindAddressByText(worksheet, new string[] { "Max" }, true);
                    if (dic.TryGetValue("Max", out string address))
                    {
                        TDMK_EPPLUS _lub = new TDMK_EPPLUS();
                        int cell = _lub.Get_Last_Cells_addr(address, worksheet, true);

                        string num = (string)worksheet.Cells[worksheet.Cells[address].Start.Row, cell].Value;
                        if (double.TryParse(num, out double force))
                        {
                            return force;
                        }
                    }
                }
            }
            return 0;
        }

        public void SaveProcess(DataTable dataTables, bool prime = false)
        {
            string err = "";
            if (!prime)
            {
                int valuesT0 = dataTables.AsEnumerable()
                             .Select(row => row.Field<byte[]>("T0"))
                             .ToList().Count(item => item != null);
                int valuesT1 = dataTables.AsEnumerable()
                             .Select(row => row.Field<byte[]>("T1"))
                             .ToList().Count(item => item != null);
                int valuesT30 = dataTables.AsEnumerable()
                             .Select(row => row.Field<byte[]>("T30"))
                             .ToList().Count(item => item != null);
                int valuesT0U = dataTables.AsEnumerable()
                             .Select(row => row.Field<byte[]>("T0U"))
                             .ToList().Count(item => item != null);
                int valuesT1U = dataTables.AsEnumerable()
                             .Select(row => row.Field<byte[]>("T1U"))
                             .ToList().Count(item => item != null);
                int valuesT30U = dataTables.AsEnumerable()
                             .Select(row => row.Field<byte[]>("T30U"))
                             .ToList().Count(item => item != null);
                if (valuesT0 < 10 || valuesT0U < 10)
                {
                    err += "Không đủ dữ liệu T0\n";
                }
                if (valuesT1 < 10 || valuesT1U < 10)
                {
                    err += "Không đủ dữ liệu T1\n";
                }
                if (valuesT30 < 10 || valuesT30U < 10)
                {
                    err += "Không đủ dữ liệu T30\n";
                }
                if (!string.IsNullOrEmpty(err))
                {
                    throw new Exception(err);
                }

                if (valuesT0 > 10 || valuesT0U > 10)
                {
                    err += "Thừa đủ dữ liệu T0\n";
                }
                if (valuesT1 > 10 || valuesT1U > 10)
                {
                    err += "Thừa đủ dữ liệu T1\n";
                }
                if (valuesT30 > 10 || valuesT30U > 10)
                {
                    err += "Thừa đủ dữ liệu T30\n";
                }
                if (!string.IsNullOrEmpty(err))
                {
                    throw new Exception("1404-" + err);
                }
            }

            DBContext db = new DBContext();
            dataTables = dataTables.AsEnumerable()
                        .OrderBy(row => row.Field<int>("id")) // Sắp xếp theo ID tăng dần
                        .Take(10) // Lấy 10 phần tử đầu tiên
                        .CopyToDataTable();

            db.BuckDataTable(dataTables, "OQC_B2B_Mating_Unmating", new string[] { "ItemCode", "lotNo" }, null, "ID");
        }

        public DataTable LoadProcess(string itemCode, string lotNo)
        {
            DataTable dataTable = new DataTable();
            dataTable = _context.LoadDataTable("OQC_B2B_Mating_Unmating", new string[] { "ItemCode", "lotNo" }, new string[] { itemCode, lotNo });
            int i = 1;
            foreach (DataRow item in dataTable.Rows)
            {
                item["ID"] = i++;
                item["Judgement"] = checkARow(item) ? "OK" : "NG";
            }

            return dataTable;
        }
        public static bool checkARow(DataRow row)
        {
            if (double.TryParse(row["Force"].ToString(), out double num))
            {
                if (num < 5)
                {
                    return false;
                }
            }
            if (string.IsNullOrEmpty(row["FailureMode"].ToString()))
            {
                return false;
            }
            string[] str = new[] { "T0", "T1", "T30", "T0U", "T1U", "T30U", "Graph" };
            foreach (var item in str)
            {
                if (row[item] == null || row[item] == DBNull.Value)
                {
                    return false;
                }
            }
            return true;
        }

        private int getId(string fileName)
        {
            if (!char.IsDigit(fileName[0]))
            {
                return -1;
            }
            string pattern = @"^\d+";

            Regex regex = new Regex(pattern);
            Match match = regex.Match(fileName);

            if (match.Success)
            {
                return int.Parse(match.Value);
            }
            return -1;
        }
        public static void Export(SqlConnection sql, ExcelWorksheet workSheet, string itemcode, string lotno)
        {
            DBContext _context = new DBContext(sql);
            string addressNum = "";

            //string address = exportProcess.FindAddressByText(workSheet, "Bar Code Verification");
            string[] colHeader = { "Flex SN", "Min Force (N)", "Max Force (N)", "Average Force (N)", "Sample 1", "Sample 2", "Sample 3", "Sample 4", "Sample 5", "Sample 6", "Sample 7", "Sample 8", "Sample 9", "Sample 10" };
            string[] rowHeader = { "Picture T0", "Picture T30", "Unmating force", "Picture T1", "Graph unmating at T1", "Failure mode", "Judgement" };
            IDictionary<string, string> addressHeader = TDMK_DictionaryService.MergeDictionaries<string, string>(ExportProcess.FindAddressByText(workSheet, colHeader.ToArray(), true), ExportProcess.FindAddressByText(workSheet, rowHeader.ToArray()));

            DataTable dataTable = _context.LoadDataTable("OQC_B2B_Mating_Unmating", new string[] { "ItemCode", "lotNo" }, new string[] { itemcode, lotno });
            int i = 1;
            foreach (DataRow row in dataTable.Rows)
            {
                string col = addressHeader[$"Sample {i}"];

                //// cho vào T0
                byte[] Image = (byte[])row["T0"];
                //adress T0
                if (addressHeader.TryGetValue("Flex SN", out string ValueZZ))
                {
                    string addZ = workSheet.Cells[workSheet.Cells[ValueZZ].Start.Row, workSheet.Cells[col].Start.Column].Address;
                    workSheet.Cells[addZ].Value = row["ProductID"];
                }
                string address = addressHeader[$"Picture T0"];
                if (address.Contains("-"))
                {
                    address = address.Split('-')[0];
                    address = workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address;
                    ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"PictureT0Sample{i}");
                }
                else
                {
                    address = workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address;
                    ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"PictureT0Sample{i}");
                }

                //// Insert Image in T0U
                Image = (byte[])row["T0U"];
                address = ExportProcess.AddRow(workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address, 1);
                ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"PictureT0USample{i}");
                //// Insert Image in 30
                Image = (byte[])row["T30"];
                address = addressHeader[$"Picture T30"];
                if (address.Contains("-"))
                {
                    address = address.Split('-')[0];
                }
                address = workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address;
                ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"PictureT30Sample{i}");
                //// Insert Image in 30U
                Image = (byte[])row["T30U"];
                address = ExportProcess.AddRow(workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address, 1);
                ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"PictureT30USample{i}");
                //// Insert Image in 30
                Image = (byte[])row["T1"];
                address = addressHeader[$"Picture T1"];
                if (address.Contains("-"))
                {
                    address = address.Split('-')[0];
                }
                address = workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address;
                ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"PictureT1Sample{i}");
                //// Insert Image in 30U
                Image = (byte[])row["T1U"];
                address = ExportProcess.AddRow(workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address, 1);
                ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"PictureT1USample{i}");
                //// Insert Image in 30
                Image = (byte[])row["Graph"];
                address = addressHeader[$"Graph unmating at T1"];
                address = workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address;
                ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"GraphT1Sample{i}");
                //// Insert Image in 30
                if (addressHeader.TryGetValue("Unmating force", out address))
                {
                    address = workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address;
                    double num = (double)row["Force"];
                    workSheet.Cells[address].Value = num;
                    addressNum += $"{address}, ";

                }
                /// write judgment
                address = addressHeader[$"Judgement"];
                address = workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address;
                workSheet.Cells[address].Value = row["Judgement"];

                // cho failure mode
                string failureMode = workSheet.Cells[workSheet.Cells[addressHeader[$"Failure mode"]].Start.Row, workSheet.Cells[col].Start.Column].Address;
                workSheet.Cells[failureMode].Value = row["FailureMode"];
                i++;
            }
            addressNum = addressNum.Trim().TrimEnd(',');
            if (addressHeader.TryGetValue("Min Force (N)", out string addressz))
            {
                addressz = ExportProcess.AddColumn(addressz, 1);
                int colZ = workSheet.Cells[addressHeader["Unmating force"]].End.Row - workSheet.Cells[addressz].End.Row;
                int ro = workSheet.Cells[addressHeader["Sample 10"]].End.Column - workSheet.Cells[addressz].End.Column;
                workSheet.Cells[addressz].FormulaR1C1 = $"=MIN(R[{colZ}]C:R[{colZ}]C[{ro}])";
            }
            if (addressHeader.TryGetValue("Max Force (N)", out addressz))
            {
                addressz = ExportProcess.AddColumn(addressHeader["Max Force (N)"], 1);
                int colZ = workSheet.Cells[addressHeader["Unmating force"]].End.Row - workSheet.Cells[addressz].End.Row;
                int ro = workSheet.Cells[addressHeader["Sample 10"]].End.Column - workSheet.Cells[addressz].End.Column;
                workSheet.Cells[addressz].FormulaR1C1 = $"=MAX(R[{colZ}]C:R[{colZ}]C[{ro}])";
            }
            if (addressHeader.TryGetValue("Max Force (N)", out addressz))
            {
                addressz = ExportProcess.AddColumn(addressHeader["Average Force (N)"], 1);
                int colZ = workSheet.Cells[addressHeader["Unmating force"]].End.Row - workSheet.Cells[addressz].End.Row;
                int ro = workSheet.Cells[addressHeader["Sample 10"]].End.Column - workSheet.Cells[addressz].End.Column;
                workSheet.Cells[addressz].FormulaR1C1 = $"=AVERAGE(R[{colZ}]C:R[{colZ}]C[{ro}])";
            }
        }



        public string Export(string itemcode, string lotno)
        {
            string addressNum = "";
            ExportProcess exportProcess = new ExportProcess();
            using (ExcelPackage ex = exportProcess.FindFormatProcess("OQC B2B Mating-Unmating", itemcode, lotno))
            {
                using (ExcelWorksheet workSheet = exportProcess.FindSheet(ex, "OQC B2B Mating-Unmating"))
                {
                    //string address = exportProcess.FindAddressByText(workSheet, "Bar Code Verification");
                    string[] colHeader = { "Min Force (N)", "Max Force (N)", "Average Force (N)", "Sample 1", "Sample 2", "Sample 3", "Sample 4", "Sample 5", "Sample 6", "Sample 7", "Sample 8", "Sample 9", "Sample 10" };
                    string[] rowHeader = { "Picture T0", "Picture T30", " Unmating force at T1", "Unmating picture at T1", "Graph unmating at T1", "Failure mode", "Judgement " };
                    IDictionary<string, string> addressHeader = TDMK_DictionaryService.MergeDictionaries<string, string>(ExportProcess.FindAddressByText(workSheet, colHeader.ToArray(), true), ExportProcess.FindAddressByText(workSheet, rowHeader.ToArray()));

                    DataTable dataTable = _context.LoadDataTable("OQC_B2B_Mating_Unmating", new string[] { "ItemCode", "lotNo" }, new string[] { itemcode, lotno });
                    int i = 1;
                    foreach (DataRow row in dataTable.Rows)
                    {
                        string col = addressHeader[$"Sample {i}"];

                        //// cho vào T0
                        byte[] Image = (byte[])row["T0"];
                        //adress T0
                        string address = addressHeader[$"Picture T0"];
                        if (address.Contains("-"))
                        {
                            address = address.Split('-')[0];
                            address = workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address;
                            ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"PictureT0Sample{i}");
                        }
                        else
                        {
                            address = workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address;
                            ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"PictureT0Sample{i}");
                        }

                        //// Insert Image in T0U
                        Image = (byte[])row["T0U"];
                        address = ExportProcess.AddRow(workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address, 1);
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"PictureT0USample{i}");
                        //// Insert Image in 30
                        Image = (byte[])row["T30"];
                        address = addressHeader[$"Picture T30"];
                        if (address.Contains("-"))
                        {
                            address = address.Split('-')[0];
                        }
                        address = workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address;
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"PictureT30Sample{i}");
                        //// Insert Image in 30U
                        Image = (byte[])row["T30U"];
                        address = ExportProcess.AddRow(workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address, 1);
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"PictureT30USample{i}");
                        //// Insert Image in 30
                        Image = (byte[])row["T1"];
                        address = addressHeader[$"Unmating picture at T1"];
                        if (address.Contains("-"))
                        {
                            address = address.Split('-')[0];
                        }
                        address = workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address;
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"PictureT1Sample{i}");
                        //// Insert Image in 30U
                        Image = (byte[])row["T1U"];
                        address = ExportProcess.AddRow(workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address, 1);
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"PictureT1USample{i}");
                        //// Insert Image in 30
                        Image = (byte[])row["Graph"];
                        address = addressHeader[$"Graph unmating at T1"];
                        address = workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address;
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"GraphT1Sample{i}");
                        //// Insert Image in 30
                        if (addressHeader.TryGetValue("Unmating force at T1", out address))
                        {
                            address = workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address;
                            double num = (double)row["Force"];
                            workSheet.Cells[address].Value = num;
                            addressNum += $"{address}, ";

                        }
                        /// write judgment
                        address = addressHeader[$"Judgement"];
                        address = workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address;
                        workSheet.Cells[address].Value = row["Judgement"].Equals("OK") ? "Pass" : row["Judgement"];

                        // cho failure mode
                        string failureMode = workSheet.Cells[workSheet.Cells[addressHeader[$"Failure mode"]].Start.Row, workSheet.Cells[col].Start.Column].Address;
                        workSheet.Cells[failureMode].Value = row["FailureMode"];
                        i++;
                    }
                    addressNum = addressNum.Trim().TrimEnd(',');
                    string addressz = ExportProcess.AddColumn(addressHeader["Min Force (N)"], 1);
                    workSheet.Cells[addressz].Formula = $"MAX({addressNum})";
                    addressz = ExportProcess.AddColumn(addressHeader["Max Force (N)"], 1);
                    workSheet.Cells[addressz].Formula = $"MIN({addressNum})";
                    addressz = ExportProcess.AddColumn(addressHeader["Average Force (N)"], 1);
                    workSheet.Cells[addressz].Formula = $"AVERAGE({addressNum})";

                    exportProcess.SaveExcelWorksheet(ex, "OQC B2B Mating-Unmating", $"{itemcode.Trim()}-{lotno.Trim()}");

                }
            }
            return "done";
        }
        public static string setup_spec(ExcelWorksheet worksheet)
        {
            int pcs = 0;
            string value = "";
            string[] addressCol = new[] { "Unmating force at T1", "Sample" };
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(worksheet, addressCol, false);
            if (dic.TryGetValue("Sample", out string address))
            {
                pcs = address.Split('-').Count();
            }
            try
            {

                if (dic.TryGetValue("Unmating force at T1", out address))
                {
                    value = worksheet.Cells[address].Text.Split('(')[1].Trim().TrimEnd(')').Trim();
                }
            }
            catch (Exception ex)
            {
            }
            return $"{pcs} : {value}";
        }
    }
}
