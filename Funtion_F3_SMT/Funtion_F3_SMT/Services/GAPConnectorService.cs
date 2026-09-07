using OfficeOpenXml;
using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace OK2SHIP_SMT.Services
{
    public class GAPConnectorService
    {
        private DBContext _dbContext = new DBContext();
        public static void Load(string itemcode, string lotno)
        {
            DBContext _context = new DBContext();
           // DataTable dt =  _context.LoadDataTable("GAP_CONNECTOR_NAS",  new[]{"ItemCode", "LotNo"})
        }
        public static void Export(string itemcode, string lotNo)
        {
            
        }
        public void Export(ExcelWorksheet workSheet, DataTable dataTable, DataTable spec)
        {
            int SpecNum = int.Parse(spec.Rows[0]["Count_Sample"].ToString());
            string[] healder = new[]
                { "IO PIN", "GROUNDING PIN_LEFT", "GROUNDING PIN_RIGHT", "Sample 1", "Average", "Min", "Max" };
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(workSheet, healder);

            Dictionary<string, string> _heal = new Dictionary<string, string>();
            foreach (string item in new[] { "IO PIN", "GROUNDING PIN_LEFT", "GROUNDING PIN_RIGHT" })
            {
                if (dic.TryGetValue("Sample 1", out string addressS))
                {
                    string addRes = "";
                    int min = int.MaxValue;
                    foreach (string jtem in addressS.Split('-'))
                    {
                        int distance = ExportProcess.DistanceRow(dic[item], jtem);
                        if (distance < min && distance > 0)
                        {
                            addRes = jtem;
                            min = distance;
                        }
                    }

                    _heal.Add(dic[item], addRes);
                }
            }

            foreach (string item in new[] { "IO PIN", "GROUNDING PIN_LEFT", "GROUNDING PIN_RIGHT" })
            {
                DataTable originalTable = new DataTable();
                switch (item)
                {
                    case "IO PIN":
                        originalTable = dataTable.AsEnumerable()
                            .Where(row => row.Field<string>("region") == "GAP DOC")
                            .CopyToDataTable();
                        break;
                    case "GROUNDING PIN_LEFT":
                        originalTable = dataTable.AsEnumerable()
                            .Where(row => row.Field<string>("region") == "GAP TRU")
                            .OrderBy(row =>
                            {
                                string sampleValue = row.Field<string>("Sample");
                                // Kiểm tra và chuyển đổi. Nếu chuyển đổi thất bại hoặc null/empty, coi là 0 (hoặc giá trị thấp nhất)
                                if (int.TryParse(sampleValue, out int result))
                                {
                                    return result;
                                }

                                return int.MinValue;
                            })
                            .Take(5)
                            .CopyToDataTable();
                        break;
                    case "GROUNDING PIN_RIGHT":
                        originalTable = dataTable.AsEnumerable()
                            .Where(row => row.Field<string>("region") == "GAP TRU")
                            .OrderBy(row =>
                            {
                                string sampleValue = row.Field<string>("Sample");
                                // Kiểm tra và chuyển đổi. Nếu chuyển đổi thất bại hoặc null/empty, coi là 0 (hoặc giá trị thấp nhất)
                                if (int.TryParse(sampleValue, out int result))
                                {
                                    return result;
                                }

                                return int.MinValue;
                            })
                            .Skip(5)
                            .Take(5)
                            .CopyToDataTable();
                        break;
                }

                bool prime = false;
                string address = _heal[dic[item]];
                for (int i = 0; i < SpecNum; i++)
                {
                    string addressP = ExportProcess.AddColumn(address, i);
                    // Kiểm tra và ghi product ID
                    if (prime || workSheet.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressP, -1), -1)].Text
                            .Contains("Flex"))
                    {
                        prime = true;
                        if (originalTable.Columns.Contains("ProductID"))
                        {
                            workSheet.Cells[ExportProcess.AddRow(addressP, -1)].Value =
                                originalTable.Rows[i]["ProductID"];
                        }
                    }
                    else
                    {
                        addressP = ExportProcess.AddRow(addressP, -1);
                    }

                    addressP = ExportProcess.AddRow(addressP, 1);
                    bool imageMode = !originalTable.Columns["Image"].GetType().ToString().Contains("byte");
                    if (!imageMode)
                    {
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[addressP],
                            TDMK_ImageConverter.ImageToByteArray((Image)originalTable.Rows[i]["Image"],
                                ImageFormat.Jpeg), $"{Guid.NewGuid()}");
                    }
                    else
                    {
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[addressP],
                            (byte[])originalTable.Rows[i]["Image"], $"{Guid.NewGuid()}");
                    }

                    addressP = ExportProcess.AddRow(addressP, 1);
                    if (!imageMode)
                    {
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[addressP],
                            TDMK_ImageConverter.ImageToByteArray((Image)originalTable.Rows[i]["Image1"],
                                ImageFormat.Jpeg), $"{Guid.NewGuid()}");
                    }
                    else
                    {
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[addressP],
                            (byte[])originalTable.Rows[i]["Image1"], $"{Guid.NewGuid()}");
                    }

                    string[] data1 = originalTable.Rows[i]["Data"].ToString().Split('/')[0].Trim().TrimEnd(';')
                        .Split(';');
                    try
                    {
                        for (int j = 0; j < data1.Count(); j++)
                        {
                            addressP = ExportProcess.AddRow(addressP, 1);
                            workSheet.Cells[addressP].Value = double.Parse(data1[j]);
                            workSheet.Cells[addressP].Style.Numberformat.Format = "#,##0.00";
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Lỗi khi ghi data {ex.Message}");
                    }

                    addressP = ExportProcess.AddRow(addressP, 1);
                    if (!imageMode)
                    {
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[addressP],
                            TDMK_ImageConverter.ImageToByteArray((Image)originalTable.Rows[i]["Image2"],
                                ImageFormat.Jpeg), $"{Guid.NewGuid()}");
                    }
                    else
                    {
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[addressP],
                            (byte[])originalTable.Rows[i]["Image2"], $"{Guid.NewGuid()}");
                    }

                    string[] data2 = originalTable.Rows[i]["Data"].ToString().Split('/')[1].Trim().TrimEnd(';')
                        .Split(';');
                    try
                    {
                        for (int j = 0; j < data2.Count(); j++)
                        {
                            addressP = ExportProcess.AddRow(addressP, 1);
                            workSheet.Cells[addressP].Value = double.Parse(data2[j]);
                            workSheet.Cells[addressP].Style.Numberformat.Format = "#,##0.00";
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Lỗi khi ghi data {ex.Message}");
                    }

                    addressP = ExportProcess.AddRow(addressP, 1);
                    workSheet.Cells[addressP].Value = "OK";
                }
            }

            foreach (string key in new[] { "Average", "Min", "Max" })
            {
                foreach (string adress in dic[key].Split('-'))
                {
                    for (int i = 0; i < 4; i++)
                    {
                        string address = ExportProcess.AddRow(adress, i + 1);
                        workSheet.Cells[address].Value = 0.1;
                        switch (key)
                        {
                            case "Average":
                                workSheet.Cells[address].FormulaR1C1 = $"=AVERAGE(RC[-5]:RC[-1])";
                                break;
                            case "Min":
                                workSheet.Cells[address].FormulaR1C1 = $"=min(RC[-6]:RC[-2])";
                                break;
                            case "Max":
                                workSheet.Cells[address].FormulaR1C1 = $"=max(RC[-7]:RC[-3])";
                                break;
                            default:
                                Debugger.Break();
                                break;
                        }
                    }
                }
            }
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

        public static DataTable ReadLogfile(string in_src, string infor)
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
                    case "GAP Connector":
                        solveB2B(fileLocation, dt);
                        break;

                    default:
                        break;
                }
            }

            return dt;
        }
        
        public static List<string> check_SPEC(DataTable dt, string itemCode)
        {
            List<string> list = new List<string>();
        
            return list;
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

        public static DataTable getConstructor()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("ItemCode", typeof(string));
            dt.Columns.Add("LotNo", typeof(string));
            dt.Columns.Add("Sheet", typeof(string));
            dt.Columns.Add("Region", typeof(string));
            dt.Columns.Add("Sample", typeof(string));
            dt.Columns.Add("Image", typeof(Image));
            dt.Columns.Add("Image1", typeof(Image));
            dt.Columns.Add("Image2", typeof(Image));
            dt.Columns.Add("Data", typeof(string));
            dt.Columns.Add("Operator", typeof(string));
            dt.Columns.Add("Time_Update", typeof(string));
            dt.Columns.Add("Remark", typeof(string));
            return dt;
        }

        public static Dictionary<string, List<string>> Get_ProductID(string location, string itemCode, string lotNo)
        {
            if (string.IsNullOrEmpty(location))
            {
                return new Dictionary<string, List<string>>();
            }

            Dictionary<string, List<string>> list = new Dictionary<string, List<string>>();
            ProductIDService service =
                new ProductIDService(itemCode, lotNo, location, new[] { "GAP" }, new[] { itemCode });
            List<string> listZ = service.getListProductID(service._listFile[itemCode]);
            if (listZ.Count >= 10)
            {
                List<string> doc = new List<string>();
                for (int i = 0; i < 5; i++)
                {
                    doc.Add(listZ[i]);
                }

                List<string> tru = new List<string>();
                for (int i = 0; i < 10; i++)
                {
                    tru.Add(listZ[i]);
                }

                list.Add("DOC", doc);
                list.Add("TRU", tru);
            }

            return list;
        }
    }
}