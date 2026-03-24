using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OK2SHIP_SMT.Libary;
using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

namespace OK2SHIP_SMT.Services
{
    public class OQCB2BMatingUnmatting : IDisposable
    {
        private DBContext _context = new DBContext();

        public OQCB2BMatingUnmatting()
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;

        }
        public DataTable getDataTableStructor()
        {
            DataTable table = new DataTable();
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("T0", typeof(Image));
            table.Columns.Add("T1", typeof(Image));
            table.Columns.Add("T30", typeof(Image));
            table.Columns.Add("T0U", typeof(Image));
            table.Columns.Add("T1U", typeof(Image));
            table.Columns.Add("T30U", typeof(Image));
            table.Columns.Add("Force", typeof(double));
            table.Columns.Add("FailureMode", typeof(string));
            table.Columns.Add("Judgement", typeof(string));
            table.Columns.Add("Graph", typeof(Image));
            table.Columns.Add("ProductID", typeof(string));
            return table;
        }
        public DataTable ProcessRead(string location, string itemCode, string lotNo, string productIDLocation)
        {
            //location = location.Replace("\\UMT", "");
            DataTable dataTable = new DataTable();
            string[] files = Directory.GetDirectories(location);
            string fileT1 = files.FirstOrDefault(x => new DirectoryInfo(x).Name.Equals("T1") || new DirectoryInfo(x).Name.Equals("L1"));
            string[] fileCSVT1 = FileFolderRepository.ListAllFileInFolder(fileT1, ".xlsx").ToArray();

            Dictionary<string, IList<KeyValuePair<Image, string>>> listPictureZ = FileFolderRepository.ListAllPictureInFolder(location);
            Dictionary<string, IList<KeyValuePair<Image, string>>> listPicture = new Dictionary<string, IList<KeyValuePair<Image, string>>>();
            ///sortting
            foreach (string key in listPictureZ.Keys)
            {
                var regex = new Regex(@"(\d+)(?:\.\w+)?$");

                var z = listPictureZ[key].Select(pair =>
                {
                    var match = regex.Match(pair.Value);
                    if (match.Success && int.TryParse(match.Groups[1].Value, out int number))
                    {
                        return new { Pair = pair, SortValue = number, Success = true };
                    }
                    else
                    {
                        return new { Pair = pair, SortValue = int.MaxValue, Success = false };
                    }
                })
                .OrderBy(x => x.SortValue)

               .Select((x, index) =>
               {
                   // 1. Lấy Image (Key) gốc
                   var originalImage = x.Pair.Key;

                   // 2. Tính số thứ tự (bắt đầu từ 1)
                   int fileNumber = index + 1;

                   // 3. Định dạng tên file mới (Ví dụ: "image_1.jpg")
                   string fileExtension = ".jpg";

                   // Bạn có thể tùy chỉnh tiền tố và đuôi file ở đây
                   string newFileName = $"{fileNumber}{fileExtension}";

                   // 4. Trả về KeyValuePair<Image, string> mới với tên file đã được đổi
                   return new KeyValuePair<Image, string>(originalImage, newFileName);
               })
    .ToList();
                listPicture[key] = z;
            }
            dataTable = getDataTableStructor();
            int i = 0, firstID0 = -1, firstID1 = -1, firstID30 = -1;
            int cfirstID0 = 0, cfirstID1 = 0, cfirstID30 = 0;
            int clastID0 = 0, clastID1 = 0, clastID30 = 0;

            Dictionary<int, int> PrimeCheck = new Dictionary<int, int>();

            foreach (var item in listPicture)
            {
                string folderName = FileFolderRepository.GetFolderName(item.Key);
                if (folderName.Equals("T0") || folderName.Equals("T1") || folderName.Equals("T30") || folderName.Equals("L0") || folderName.Equals("L1") || folderName.Equals("L30"))
                {
                    int t = int.Parse(folderName.Substring(1));
                    if (folderName.Contains('L'))
                    {
                        folderName = folderName.Replace("L", "T");
                    }
                    foreach (var jtem in item.Value)
                    {
                        string nameFile = jtem.Value;
                        int id = getId(nameFile);
                        int firstID = 0, cFirstID = 0;
                        switch (t)
                        {
                            case 1:
                                if (firstID1 == -1)
                                {
                                    firstID1 = id;
                                    clastID0 = id;

                                }
                                if (id - clastID0 > 1)
                                {
                                    cfirstID0 += id - clastID1;
                                }
                                clastID0 = id;
                                firstID = firstID1;
                                cFirstID = cfirstID1;
                                break;
                            case 0:
                                if (firstID0 == -1)
                                {
                                    firstID0 = id;
                                }
                                firstID = firstID0;
                                cFirstID = cfirstID0;
                                break;
                            case 30:
                                if (firstID30 == -1)
                                {
                                    firstID30 = id;
                                }
                                firstID = firstID30;
                                cFirstID = cfirstID30;
                                break;
                        }

                        id = id - firstID;
                        // 1 - 10 10 - 20
                        bool primePlus = ((int)id / 10) % 2 == 1;
                        id = id % 10 + 1;
                        if (id <= 0)
                        {
                            throw new Exception("Lỗi đổi tên file ảnh thành id");
                        }
                        int value;

                        // nếu đã tồn tại row
                        if (PrimeCheck.TryGetValue(id, out value))
                        {
                            //chọn mục ảnh
                            if (primePlus)
                            {
                                dataTable.Rows[value][folderName + "U"] = jtem.Key;
                            }
                            else
                            {
                                dataTable.Rows[value][folderName] = jtem.Key;
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
                            //chọn mục ảnh
                            if (primePlus)
                            {
                                row[folderName + "U"] = jtem.Key;
                            }
                            else
                            {
                                row[folderName] = jtem.Key;
                            }
                            // thêm row vao datatable
                            row["Judgement"] = "NG";
                            dataTable.Rows.Add(row);
                            //Debugger.Break();
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
                            dataTable.Rows[index]["Graph"] = image;
                            dataTable.Rows[index]["Force"] = force;

                        }
                    }
                }
            }
            int iz = 1;
            DataView dv = dataTable.DefaultView;
            dv.Sort = "ID ASC";
            DataTable sortedDt = dv.ToTable();
            productIDLocation = productIDLocation.Trim();
            try
            {

                if (!string.IsNullOrEmpty(productIDLocation))
                {
                    ProductIDService productID = new ProductIDService(itemCode, lotNo, productIDLocation, new[] { "OQC", "B2B", "Matting", "Un-Matting" }, new[] { itemCode });
                    IList<string> list = productID.getListProductID(productID._listFile[itemCode]);
                    int iZ = 0;
                    foreach (DataRow row in sortedDt.Rows)
                    {
                        if (iZ < list.Count)
                        {
                            row["ProductID"] = list[iZ];
                            iZ++;
                        }
                    }

                }
            }
            catch
            {
                MessageBox.Show("Kiểm tra lại product ID");
            }
            //Debugger.Break();
            return sortedDt;
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

        public void SaveProcess(DataTable dataTables, string itemCode, string lotNo, bool prime = false)
        {
            string err = "";
            //if (!prime)
            //{
            //    int valuesT0 = dataTables.AsEnumerable()
            //                 .Select(row => row.Field<Image>("T0"))
            //                 .ToList().Count(item => item != null);
            //    int valuesT1 = dataTables.AsEnumerable()
            //                 .Select(row => row.Field<byte[]>("T1"))
            //                 .ToList().Count(item => item != null);
            //    int valuesT30 = dataTables.AsEnumerable()
            //                 .Select(row => row.Field<byte[]>("T30"))
            //                 .ToList().Count(item => item != null);
            //    int valuesT0U = dataTables.AsEnumerable()
            //                 .Select(row => row.Field<byte[]>("T0U"))
            //                 .ToList().Count(item => item != null);
            //    int valuesT1U = dataTables.AsEnumerable()
            //                 .Select(row => row.Field<byte[]>("T1U"))
            //                 .ToList().Count(item => item != null);
            //    int valuesT30U = dataTables.AsEnumerable()
            //                 .Select(row => row.Field<byte[]>("T30U"))
            //                 .ToList().Count(item => item != null);
            //    if (valuesT0 < 10 || valuesT0U < 10)
            //    {
            //        err += "Không đủ dữ liệu T0\n";
            //    }
            //    if (valuesT1 < 10 || valuesT1U < 10)
            //    {
            //        err += "Không đủ dữ liệu T1\n";
            //    }
            //    if (valuesT30 < 10 || valuesT30U < 10)
            //    {
            //        err += "Không đủ dữ liệu T30\n";
            //    }
            //    if (!string.IsNullOrEmpty(err))
            //    {
            //        throw new Exception(err);
            //    }

            //    if (valuesT0 > 10 || valuesT0U > 10)
            //    {
            //        err += "Thừa đủ dữ liệu T0\n";
            //    }
            //    if (valuesT1 > 10 || valuesT1U > 10)
            //    {
            //        err += "Thừa đủ dữ liệu T1\n";
            //    }
            //    if (valuesT30 > 10 || valuesT30U > 10)
            //    {
            //        err += "Thừa đủ dữ liệu T30\n";
            //    }
            //    if (!string.IsNullOrEmpty(err))
            //    {
            //        throw new Exception("1404-" + err);
            //    }
            //}

            DBContext db = new DBContext();
            dataTables = dataTables.AsEnumerable()
                        .OrderBy(row => row.Field<int>("id")) // Sắp xếp theo ID tăng dần
                        .Take(10) // Lấy 10 phần tử đầu tiên
                        .CopyToDataTable();
            NasRepository _nas = new NasRepository();
            string location = _nas.HandleImageDataTable(dataTables, "OQC_B2B_Mating_Unmating_Nas", itemCode, lotNo);
            DataTable dt = _context.GetTableStructure("OQC_B2B_Mating_Unmating_Nas");
            DataRow rowZ = dt.NewRow();
            rowZ["ItemCode"] = itemCode;
            rowZ["LotNo"] = lotNo;
            rowZ["Data"] = ConverterService.DataTableToJson(dataTables);
            rowZ["LoactionImg"] = location;
            dt.Rows.Add(rowZ);
            db.BuckDataTable(dt, "OQC_B2B_Mating_Unmating_Nas", new string[] { "ItemCode", "lotNo" }, null, "ID");
        }

        public DataTable LoadProcess(string itemCode, string lotNo, bool legacy = false)
        {
            DataTable dataTable = new DataTable();
            if (legacy)
            {

                dataTable = _context.LoadDataTable("OQC_B2B_Mating_Unmating", new string[] { "ItemCode", "lotNo" }, new string[] { itemCode, lotNo });
                int i = 1;
                foreach (DataRow item in dataTable.Rows)
                {
                    item["ID"] = i++;
                    item["Judgement"] = checkARow(item) ? "OK" : "NG";
                }

            }
            else
            {
                dataTable = _context.LoadDataTable("OQC_B2B_Mating_Unmating_NAS", new string[] { "ItemCode", "lotNo" }, new string[] { itemCode, lotNo });
                if (dataTable.Rows.Count <= 0)
                {
                    throw new Exception("No data");
                }
                string json = dataTable.Rows[0]["Data"].ToString();
                string location = dataTable.Rows[0]["LoactionImg"].ToString();
                dataTable = ConverterService.JsonToDataTable(json);
                NasRepository _nas = new NasRepository();
                _nas.MergeDataTable(dataTable, "OQC_B2B_Mating_Unmating_NAS", itemCode, lotNo, location);
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

        public string Export(string itemcode, string lotno, bool legacy = false)
        {
            string addressNum = "";
            ExportProcess exportProcess = new ExportProcess();
            using (ExcelPackage ex = exportProcess.FindFormatProcess("OQC B2B Mating-Unmating", itemcode, lotno))
            {
                using (ExcelWorksheet workSheet = exportProcess.FindSheet(ex, "OQC B2B Mating-Unmating"))
                {
                    //string address = exportProcess.FindAddressByText(workSheet, "Bar Code Verification");
                    string[] colHeader = { "Flex SN", "Min Force (N)", "Max Force (N)", "Average Force (N)", "Sample 1", "Sample 2", "Sample 3", "Sample 4", "Sample 5", "Sample 6", "Sample 7", "Sample 8", "Sample 9", "Sample 10" };
                    string[] rowHeader = { "Picture T0", "Picture T30", "Unmating force", "Picture T1", "Graph unmating at T1", "Failure mode", "Judgement" };
                    IDictionary<string, string> addressHeader = DictionaryService.MergeDictionaries<string, string>(ExportProcess.FindAddressByText(workSheet, colHeader.ToArray(), true), ExportProcess.FindAddressByText(workSheet, rowHeader.ToArray()));

                    DataTable dataTable = LoadProcess(itemcode, lotno, legacy);
                    int i = 1;
                    foreach (DataRow row in dataTable.Rows)
                    {
                        string col = addressHeader[$"Sample {i}"];

                        //// cho vào T0
                        ///
                        byte[] Image;
                        try
                        {
                            Image = TDMK_ImageConverter.ImageToByteArray((Image)row["T0"], ImageFormat.Jpeg);
                        }
                        catch
                        {
                            Image = (byte[])row["T0"];
                        }
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
                        try
                        {
                            Image = TDMK_ImageConverter.ImageToByteArray((Image)row["T0U"], ImageFormat.Jpeg);
                        }
                        catch
                        {
                            Image = (byte[])row["T0U"];
                        }
                        address = ExportProcess.AddRow(workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address, 1);
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"PictureT0USample{i}");
                        //// Insert Image in 30
                        try
                        {

                            Image = TDMK_ImageConverter.ImageToByteArray((Image)row["T30"], ImageFormat.Jpeg);
                        }
                        catch
                        {
                            Image = (byte[])row["T30"];

                        }
                        address = addressHeader[$"Picture T30"];
                        if (address.Contains("-"))
                        {
                            address = address.Split('-')[0];
                        }
                        address = workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address;
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"PictureT30Sample{i}");
                        //// Insert Image in 30U
                        try
                        {

                            Image = TDMK_ImageConverter.ImageToByteArray((Image)row["T30U"], ImageFormat.Jpeg);
                        }
                        catch
                        {
                            Image = (byte[])row["T30U"];
                        }
                        address = ExportProcess.AddRow(workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address, 1);
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"PictureT30USample{i}");
                        //// Insert Image in 30
                        try
                        {

                            Image = TDMK_ImageConverter.ImageToByteArray((Image)row["T1"], ImageFormat.Jpeg);
                        }
                        catch
                        {
                            Image = (byte[])row["T1"];
                        }

                        address = addressHeader[$"Picture T1"];
                        if (address.Contains("-"))
                        {
                            address = address.Split('-')[0];
                        }
                        address = workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address;
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"PictureT1Sample{i}");
                        //// Insert Image in 30U
                        try
                        {

                            Image = TDMK_ImageConverter.ImageToByteArray((Image)row["T1U"], ImageFormat.Jpeg);
                        }
                        catch
                        {
                            Image = (byte[])row["T1U"];
                        }

                        address = ExportProcess.AddRow(workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address, 1);
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"PictureT1USample{i}");
                        //// Insert Image in 30
                        ///

                        try
                        {

                            Image = TDMK_ImageConverter.ImageToByteArray((Image)row["Graph"], ImageFormat.Jpeg);
                        }
                        catch
                        {
                            Image = (byte[])row["Graph"];
                        }

                        address = addressHeader[$"Graph unmating at T1"];
                        address = workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address;
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], Image, $"GraphT1Sample{i}");
                        //// Insert Image in 30
                        if (addressHeader.TryGetValue("Unmating force", out address))
                        {
                            address = workSheet.Cells[workSheet.Cells[address].Start.Row, workSheet.Cells[col].Start.Column].Address;
                            double num = double.Parse(row["Force"].ToString());
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

                    exportProcess.SaveExcelWorksheet(ex, "OQC B2B Mating-Unmating", $"{itemcode.Trim()}-{lotno.Trim()}");

                }
            }
            return "done";
        }

        public void Dispose()
        {
            GC.Collect();
        }
    }
}
