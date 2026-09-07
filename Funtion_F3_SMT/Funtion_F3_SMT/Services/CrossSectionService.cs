using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OfficeOpenXml;
using OK2SHIP_SMT.Repositories;

namespace OK2SHIP_SMT.Services
{
    public class CrossSectionService
    {
        public static DataTable SelectData(DataTable dataTable)
        {
            // Kiểm tra an toàn nếu dữ liệu rỗng hoặc không chứa cột "select"
            if (dataTable == null || !dataTable.Columns.Contains("select"))
            {
                return dataTable?.Copy();
            }

            // 1. Clone cấu trúc của DataTable cũ (chỉ copy các cột, không copy dữ liệu)
            DataTable newData = dataTable.Clone();

            // 2. Lọc các dòng có điều kiện cột "select" là true
            // Sử dụng Convert.ToBoolean để tránh lỗi nếu dữ liệu trong DB lưu dưới dạng 1/0, chuỗi "True", hoặc boolean
            var filteredRows = dataTable.AsEnumerable()
                .Where(row => row["select"] != DBNull.Value &&
                              Convert.ToBoolean(row["select"]) == true);

            // 3. Import các dòng đã lọc vào DataTable mới
            foreach (DataRow row in filteredRows)
            {
                newData.ImportRow(row);
            }

            // 4. Xóa cột "select" ở DataTable mới trước khi trả về
            newData.Columns.Remove("select");

            return newData;
        }

        public static Dictionary<string, string> loadSpec(string itemCode)
        {
            DBContext _dbContext = new DBContext();
            DataTable spec = _dbContext.LoadDataTable("SPEC_COMMENT_3", new[] { "ItemCode", "Sheet" },
                new[] { itemCode, "CROSS_SECTION" });
            if (spec.Rows.Count <= 0)
            {
                throw new Exception("Spec is not valid");
            }
            string specz = spec.Rows[0]["Location"].ToString();
            List<string> list = specz.Split('|').ToList();
            Dictionary<string, string> dic = new Dictionary<string, string>();
            foreach (string item in list)
            {
                try
                {

                    string key = item.Split(';')[0].Trim();
                    dic.Add(key.ToUpper(), item);
                }
                catch
                {
                    
                }
                
            }
            return dic;
        }

        public static List<string> check_SPEC(DataTable dt, string itemCode)
        {
            Dictionary<string, string> spec = loadSpec(itemCode);
            List<string> list = new List<string>();
            foreach (DataRow row in dt.Rows)
            {
                Debugger.Break();
                string region = row["Region"].ToString();
                if (spec.TryGetValue(region.ToUpper(), out string valueSpec))
                {
                    string[] split = valueSpec.Split(';');
                }
            }
            return list;
        }

        public static string _DBNAME = "CROSS_SECTION_NAS";


        public static DataTable load(string itemcode, string lotno, string sheet)
        {
            DBContext _context = new DBContext();
            DataTable dt = _context.LoadDataTable(_DBNAME, new[] { "ItemCode", "LotNo", "Sheet" },
                new[] { itemcode, lotno, sheet });
            return dt;
        }

        public static DataTable HandleData(DataTable dt)
        {
            string locationImage = dt.Rows[0]["LocationImg"].ToString();
            string json = dt.Rows[0]["Data"].ToString();
            string itemCode = dt.Rows[0]["ItemCode"].ToString().Trim();
            string lotNo = dt.Rows[0]["LotNo"].ToString().Trim();
            DataTable res = ConverterService.JsonToDataTable(json);
            NasRepository nas = new NasRepository();
            nas.MergeDataTable(res, _DBNAME, itemCode, lotNo, locationImage);
            return res;
        }

        public static void Save(DataTable dataTable, string itemcode, string lotno, string sheet, DataTable dtSave)
        {
            if (dataTable.Rows.Count <= 0)
            {
                throw new Exception("Data không tồn tại");
            }

            DBContext _context = new DBContext();
            dtSave.Clear();
            DataRow row = dtSave.NewRow();
            row["ItemCode"] = itemcode;
            row["LotNo"] = lotno;
            row["Sheet"] = sheet;
            NasRepository nas = new NasRepository();
            string location = nas.HandleImageDataTable(dataTable, _DBNAME, itemcode, lotno);
            row["Data"] = ConverterService.DataTableToJson(dataTable);
            row["LocationImg"] = location;
            dtSave.Rows.Add(row);
            _context.BuckDataTable(dtSave, _DBNAME, new[] { "ItemCode", "LotNo", "Sheet" }, null, "ID");
        }

        public static void ExportClip(string itemcode, string lotno)
        {
            DataTable dt_data = HandleData(load(itemcode, lotno, "Clip"));
            List<string> distinctRegions = dt_data.AsEnumerable()
                .Select(row => row.Field<string>("Region"))
                .Distinct()
                .ToList();
            ExportProcess process = new ExportProcess();
            using (ExcelPackage package = process.FindFormatWithItemCode(itemcode))
            using (ExcelWorksheet worksheet = process.FindSheet(package, "Cross Section Clip"))
            {
                ExportWSClip(worksheet, dt_data);

                process.SaveExcelWorksheet(package, "Cross Section Clip",
                    $"{itemcode}-{lotno}-crossection_clip");
            }
        }

        private static void ExportWSClip(ExcelWorksheet worksheet, DataTable dtData)
        {
            int maxRow = worksheet.Dimension != null ? worksheet.Dimension.End.Row : 0;

            IDictionary<string, string> dic =
                ExportProcess.FindAddressByText(worksheet,
                    new[] { "Clip Horizontal", "Clip Vertical" });
            foreach (string key in dic.Keys)
            {
                string category = "";
                if (key.ToLower().Contains("horizontal"))
                {
                    category = "DOC";
                }

                if (key.ToLower().Contains("vertical"))
                {
                    category = "NGANG";
                }

                List<DataRow> filteredRows = dtData.AsEnumerable()
                    .Where(row => row.Field<string>("Region").Equals(category))
                    .ToList();
                ExportInOne(worksheet, category, ExportProcess.AddColumn(dic[key], 1), filteredRows, maxRow);
            }
        }


        public static void Export(string itemcode, string lotno)
        {
            DataTable dt_data = HandleData(load(itemcode, lotno, "Shield b2b"));
            List<string> distinctRegions = dt_data.AsEnumerable()
                .Select(row => row.Field<string>("Region"))
                .Distinct()
                .ToList().Select(x => x != null ? x.Replace("PIN", "").Trim() : x)
                .ToList();

            ExportProcess process = new ExportProcess();
            using (ExcelPackage package = process.FindFormatWithItemCode(itemcode))
            using (ExcelWorksheet worksheet1 = process.FindSheet(package, "Cross Section (1.1)"))
            using (ExcelWorksheet worksheet2 = process.FindSheet(package, "Cross Section (1.2)"))
            {
                ExportWS(worksheet1, dt_data, distinctRegions);
                ExportWS(worksheet2, dt_data, distinctRegions);

                process.SaveExcelWorksheet(package, "Cross Section (1.1):Cross Section (1.2)",
                    $"{itemcode}-{lotno}-crossection");
            }
        }

        public static void ExportWS(ExcelWorksheet ws, DataTable dt_data, List<string> distinctRegions)
        {
            IDictionary<string, string> dicAddress =
                ExportProcess.FindAddressByText(ws, distinctRegions.ToArray());
            int maxRow = ws.Dimension != null ? ws.Dimension.End.Row : 0;
            foreach (string key in dicAddress.Keys)
            {
                string z = key;
                if (key.Contains("NGANG") && !key.Contains("SHIELD"))
                {
                    z = $"PIN {key}";
                }

                List<DataRow> filteredRows = dt_data.AsEnumerable()
                    .Where(row => row.Field<string>("Region").Equals(z))
                    .ToList();
                ExportInOne(ws, key, dicAddress[key], filteredRows, maxRow);
            }
        }

        public static void ExportInOne(ExcelWorksheet ws, string category, string address, List<DataRow> filteredRows,
            int maxrow)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            string name = category;
            string addressI = address;
            int i = 0, percent = 0, mma = 0, z = 0;
            int currentRow = int.Parse(Regex.Match(addressI, @"\d+").Value);
            // mainIMG, %percenti, numz 
            while (name.ToUpper() != "FLEX SN" && currentRow <= maxrow)
            {
                currentRow++;
                if (name.Contains(category))
                {
                    dic.Add("mainIMG", addressI);
                }

                else if (name.Contains("%"))
                {
                    dic.Add($"%{percent++}", addressI);
                }

                else if (name.ToUpper().Contains("MIN") || name.ToUpper().Contains("MAX") ||
                         name.ToUpper().Contains("AVERAGE"))
                {
                    dic.Add($"mma{z}_{mma++}", addressI);
                }
                else if (name.Contains("(") && name.Contains(")"))
                {
                    dic.Add($"text{z}_{i++}", addressI);
                }
                else
                {
                    i = 0;
                    mma = 0;
                    dic.Add($"image{z++}", addressI);
                }

                addressI = ExportProcess.AddRow(addressI, 1);
                name = ws.Cells[addressI].Text;
            }

            foreach (DataRow row in filteredRows)
            {
                string nameImage = row["NameImage"].ToString().Split('.')[0];
                string[] split = nameImage.Split('-');
                if (int.TryParse(row["Sample"].ToString(), out int number))
                {
                    if (split.Length > 1)
                    {
                        if (int.TryParse(split[1], out int num))
                        {
                            if (dic.TryGetValue($"image{num - 1}", out string add))
                            {
                                try
                                {
                                    ExportProcess.InsertImageToCell(ws,
                                        ws.Cells[ExportProcess.AddColumn(add, number + 3)], (byte[])row["Image"],
                                        $"{row["NameImage"]}_{number}{Guid.NewGuid()}");
                                }
                                catch (Exception ex)
                                {
                                    Debugger.Break();
                                }
                            }

                            int ix = 0;
                            while (dic.TryGetValue($"%{ix}", out string addPercent))
                            {
                                string targetCell = ExportProcess.AddColumn(addPercent, number + 3);

                                ws.Cells[targetCell].Value = 1;

                                ws.Cells[targetCell].Style.Numberformat.Format = "0%";
                                ix++;
                            }

                            ix = 0;


                            string[] splitData = row["Data"].ToString().Split(';');
                            while (dic.TryGetValue($"text{num}_{ix}", out string addPercent))
                            {
                                if (ix >= splitData.Length)
                                {
                                    break;
                                }

// Đặt biến cho địa chỉ ô để code dễ đọc hơn
                                string targetCell = ExportProcess.AddColumn(addPercent, number + 3);

// Thử chuyển đổi chuỗi thành số thập phân (double)
                                if (double.TryParse(splitData[ix], out double parsedNumber))
                                {
                                    // Làm tròn 2 chữ số và gán vào ô
                                    ws.Cells[targetCell].Value = Math.Round(parsedNumber, 2);

                                    // (Tùy chọn) Định dạng ô Excel để luôn hiển thị đủ 2 chữ số (VD: hiển thị 15.50 thay vì 15.5)
                                    ws.Cells[targetCell].Style.Numberformat.Format = "0.00";
                                }
                                else
                                {
                                    // Nếu dữ liệu không phải là số (ví dụ: bị rỗng, hoặc chữ "N/A"), cứ ghi chuỗi gốc vào
                                    ws.Cells[targetCell].Value = "N/A";
                                }

                                ix++;
                            }

                            while (dic.TryGetValue($"mma{num}_{ix}", out string addPercent))
                            {
                                int ireal;
                                switch (ix)
                                {
                                    case 0:
                                        ireal = 4;
                                        break;
                                    case 1:
                                        ireal = 0;
                                        break;

                                    case 2:
                                        ireal = 2;
                                        break;

                                    case 3:
                                        ireal = 5;
                                        break;

                                    case 4:
                                        ireal = 1;
                                        break;

                                    case 5:
                                        ireal = 3;
                                        break;
                                    default:
                                        ireal = ix;
                                        break;
                                }

                                if (ireal < splitData.Length)
                                {
                                    string targetCell = ExportProcess.AddColumn(addPercent, number + 3);

                                    // Thử chuyển đổi chuỗi thành số thập phân, dùng InvariantCulture để tránh lỗi dấu phẩy/chấm
                                    if (double.TryParse(splitData[ireal], System.Globalization.NumberStyles.Any,
                                            System.Globalization.CultureInfo.InvariantCulture,
                                            out double parsedNumberReal))
                                    {
                                        // Gán giá trị đã làm tròn 2 chữ số
                                        ws.Cells[targetCell].Value = Math.Round(parsedNumberReal, 2);

                                        // Định dạng ô trên Excel để hiển thị chuẩn 2 số thập phân (VD: 10.00)
                                        ws.Cells[targetCell].Style.Numberformat.Format = "0.00";
                                    }
                                    else
                                    {
                                        // Gán lại chuỗi gốc nếu không thể chuyển thành số
                                        ws.Cells[targetCell].Value = "N/A";
                                    }
                                }

                                ix++;
                            }
                        }
                    }
                    else
                    {
                        if (dic.TryGetValue("mainIMG", out string add))
                        {
                            try
                            {
                                ExportProcess.InsertImageToCell(ws, ws.Cells[ExportProcess.AddColumn(add, number + 3)],
                                    (byte[])row["Image"], $"{row["NameImage"]}_{number}");
                            }
                            catch (Exception ex)
                            {
                                Debugger.Break();
                            }
                        }
                    }
                }
            }
        }

        public static DataTable ReadLogfile(string in_src, string infor, bool b2b = true)
        {
            DataTable dt = getStructor();
            DataColumn selectCol = dt.Columns.Add("Select", typeof(bool));
            selectCol.DefaultValue = true;
            string[] arraySub = FileFolderRepository.GetSubFolders(in_src);
            foreach (string fileLocation in arraySub)
            {
                string folderName = FileFolderRepository.GetFolderName(fileLocation);
                switch (folderName)
                {
                    case "Crosscut Section B2B":
                        if (b2b)
                        {
                            solveB2B(fileLocation, dt);
                        }

                        break;
                    case "Crosscut Section CLIP":
                        if (!b2b)
                        {
                            solveClip(fileLocation, dt);
                        }

                        break;
                    default:
                        break;
                }
            }

            return dt;
        }

        private static void solveClip(string fileLocation, DataTable dt)
        {
            string[] arraySubFolder = FileFolderRepository.GetSubFolders(fileLocation + "\\CLIP");
            foreach (string subfolder in arraySubFolder)
            {
                GetPopulatedData(subfolder, dt);
            }
        }

        public static void solveB2B(string fileLocation, DataTable dt)
        {
            string[] arraySubFolder = FileFolderRepository.GetSubFolders(fileLocation);
            foreach (string subfolder in arraySubFolder)
            {
                GetPopulatedData(subfolder, dt);
            }
        }

        public static void GetPopulatedData(string fileLocation, DataTable dt)
        {
            string folderName = FileFolderRepository.GetFolderName(fileLocation);
            string[] listFile = Directory.GetFiles(fileLocation);

            // Lấy danh sách toàn bộ file ảnh trong thư mục
            var jpgFiles = listFile.Where(f => f.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase)).ToList();

            // 1. Xác định các Sample ID (số thứ tự 1, 2, 3...) từ các tiền tố gốc (21, 23, 24...)
            var uniquePrefixes = jpgFiles
                .Select(f => Path.GetFileName(f).Split('-', '.')[0]) // Lấy "21" từ "21-1.jpg"
                .Where(p => int.TryParse(p, out _))
                .Distinct()
                .OrderBy(p => int.Parse(p))
                .ToList();

            // Tạo Dictionary ánh xạ: Tiền tố gốc -> Sample ID (vd: "21" -> "1", "23" -> "2")
            Dictionary<string, string> sampleDict = uniquePrefixes
                .Select((prefix, index) => new { prefix, sampleId = (index + 1).ToString() })
                .ToDictionary(x => x.prefix, x => x.sampleId);

            // 2. Duyệt qua từng file ảnh để tạo dòng dữ liệu tương ứng
            foreach (var jpgPath in jpgFiles)
            {
                DataRow row = dt.NewRow();
                row["ID"] = dt.Rows.Count + 1;
                string fileName = Path.GetFileName(jpgPath); // vd: 21-1.jpg
                string baseName = Path.GetFileNameWithoutExtension(jpgPath); // vd: 21-1
                string prefix = fileName.Split('-', '.')[0]; // vd: 21

                // Tìm đường dẫn file csv tương ứng (vd: tìm 21-1.csv)
                string expectedCsvName = $"{baseName}.csv";
                string csvPath = listFile.FirstOrDefault(f =>
                    Path.GetFileName(f).Equals(expectedCsvName, StringComparison.OrdinalIgnoreCase));

                // Gán dữ liệu vào các cột
                row["NameImage"] = fileName;
                row["Image"] = File.ReadAllBytes(jpgPath); // Chuyển ảnh thành byte[]
                row["Region"] = folderName;
                if (csvPath != null)
                {
                    row["Data"] = getData(csvPath); // Lưu địa chỉ file csv
                }

                if (sampleDict.ContainsKey(prefix))
                {
                    row["Sample"] = sampleDict[prefix]; // Gán số thứ tự Sample (1, 2, 3...)
                }

                dt.Rows.Add(row);
            }
        }

        public static string getData(string address)
        {
            try
            {
                List<string> resultList = new List<string>();

                // Đọc toàn bộ các dòng trong file CSV
                string[] lines = File.ReadAllLines(address);
                bool prime = false;
                int no = 0, res = 0;
                foreach (string line in lines)
                {
                    // Tách dòng thành các cột (dùng ',' cho file CSV chuẩn)
                    string[] columns = line.Split(',');

                    // Đảm bảo dòng có ít nhất 3 cột (No., Measure, Result) để tránh lỗi OutOfBounds
                    if (columns.Length >= 3)
                    {
                        if (!prime)
                        {
                            string linez = line.ToUpper();
                            if (line.ToUpper().Contains("NO.") && line.ToUpper().Contains("RESULT"))
                            {
                                string[] z = line.Split(',');
                                for (int i = 0; i < z.Length; i++)
                                {
                                    string item = z[i];
                                    if (item.ToUpper().Contains("NO."))
                                    {
                                        no = i;
                                    }

                                    if (item.ToUpper().Contains("RESULT"))
                                    {
                                        res = i;
                                    }
                                }

                                prime = true;
                            }
                        }
                        else
                        {
                            string[] z = line.Split(',');

                            if (int.TryParse(z[no].Replace("\"", ""), out int num))
                            {
                                resultList.Add(z[res]);
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                }

                return string.Join(";", resultList);
            }
            catch
            {
                return "";
            }
        }

        public DataTable getStructorTable()
        {
            DataTable table = new DataTable();
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("ItemCode", typeof(string));
            table.Columns.Add("LotNo", typeof(string));
            table.Columns.Add("Sheet", typeof(string));
            table.Columns.Add("Region", typeof(string));
            table.Columns.Add("Sample", typeof(string));
            table.Columns.Add("Image1", typeof(Image));
            table.Columns.Add("Image2", typeof(Image));
            table.Columns.Add("Data", typeof(string));
            table.Columns.Add("Operator", typeof(string));
            table.Columns.Add("Time_Update", typeof(string));
            table.Columns.Add("Remark", typeof(string));
            return table;
        }

        public static DataTable getStructor()
        {
            DataTable table = new DataTable();
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("ProductID", typeof(string));
            table.Columns.Add("Region", typeof(string));
            table.Columns.Add("NameImage", typeof(string));
            table.Columns.Add("Sample", typeof(string));
            table.Columns.Add("Image", typeof(byte[]));
            table.Columns.Add("Data", typeof(string));
            return table;
        }

        public static DataTable getStructorTableByte(string s = "CROSS")
        {
            DataTable table = new DataTable();
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("ProductID", typeof(string));
            table.Columns.Add("ItemCode", typeof(string));
            table.Columns.Add("LotNo", typeof(string));
            table.Columns.Add("Sheet", typeof(string));
            table.Columns.Add("Region", typeof(string));
            table.Columns.Add("Sample", typeof(string));
            if (s == "GAP")
            {
                table.Columns.Add("Image", typeof(byte[]));
            }

            table.Columns.Add("Image1", typeof(byte[]));
            table.Columns.Add("Image2", typeof(byte[]));
            table.Columns.Add("Data", typeof(string));
            table.Columns.Add("Operator", typeof(string));
            table.Columns.Add("Time_Update", typeof(string));
            table.Columns.Add("Remark", typeof(string));
            return table;
        }
    }
}