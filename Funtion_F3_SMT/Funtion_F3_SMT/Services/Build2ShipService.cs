using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using OfficeOpenXml;
using OK2SHIP_SMT.Repositories;

namespace OK2SHIP_SMT.Services
{
    public class Build2ShipService
    {
        public Dictionary<string, string> _DATA = new Dictionary<string, string>();
        public string location;

        List<string> sheetNames = new List<string>
        {
            "Cross Section",
            "GAP Connector",
            "Peel Test",
            "Pull Test",
            "(IQC Unmating) Pull Test",
            "Shear test",
            "OQC B2B Mating-Unmating",
            "IQC Liner peeling",
            "IQC PSA peeling",
            "Liner peel test (On product)",
            "PSA peel test (On product)",
            "Air bubble btw Liner-PSA",
            "Air bubble btw PSA-FPC",
            "SEM Binarization"
        };

        public void GetStatus(string location)
        {
            ExportProcess exportProcess = new ExportProcess();
            using (ExcelPackage package = ExportProcess.openPackage(location))
            {
                foreach (ExcelWorksheet ws in package.Workbook.Worksheets)
                {
                    string name = ws.Name;
                    foreach (string sheetname in sheetNames)
                    {
                        if (ws.Name.ToUpper().Replace(" ", "") == sheetname.ToUpper().Replace(" ", ""))
                        {
                            _DATA.Add(name, "Sẵn sàng");
                        }
                    }
                }
            }
        }

        public void SaveSheet(string addressExcel, string itemcode, string lotno, string category)
        {
            // Cấu hình LicenseContext cho EPPlus (Bắt buộc từ phiên bản EPPlus 5+)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            ExportProcess exportProcess = new ExportProcess();

            // Mở file Excel gốc
            using (ExcelPackage sourcePackage = ExportProcess.openPackage(addressExcel))
            {
                // Tạo một ExcelPackage mới để chứa sheet được trích xuất
                using (ExcelPackage newPackage = new ExcelPackage())
                {
                    bool found = false;

                    using (ExcelWorksheet ws = exportProcess.FindSheet(sourcePackage, category))
                    {
                        // Copy sheet sang package mới

                        switch (category)
                        {
                            case "Peel Test":
                            case "Pull Test":
                            case "Shear test":
                                HandleSheet(ws, category);
                                break;
                            default:
                                break;
                        }

                        newPackage.Workbook.Worksheets.Add(ws.Name, ws);
                        found = true;
                    }

                    if (found)
                    {
                        NasRepository nas = new NasRepository();

                        string addressSave =
                            $"{nas._nasAddress}\\SEEV Data\\Report\\NPI\\BUILD_2_SHIP\\{itemcode}_{lotno}";
                        if (!Directory.Exists(addressSave))
                        {
                            Directory.CreateDirectory(addressSave);
                        }

                        string addressFileSave =
                            $"{addressSave}\\{category}.{FileFolderRepository.GetFileName(addressExcel).Split('.')[1]}";
                        FileInfo fileInfo = new FileInfo(addressFileSave);
                        newPackage.SaveAs(fileInfo);
                    }
                    else
                    {
                        throw new FileNotFoundException("Không tìm thấy sheet phù hợp với điều kiện đã cho.");
                    }
                }
            }
        }

        private void HandleSheet(ExcelWorksheet ws, string category)
        {
            int sampleTo = 10, sampleFrom = 1;
            IDictionary<string, string> _dic = ExportProcess.FindAddressByText(ws, new[] { "Sample" });
            string[] splitSample = _dic["Sample"].Split('-');
            foreach (string s in splitSample)
            {
                try
                {
                    string sampleName = ws.Cells[s].Value.ToString();
                    int sample = int.Parse(sampleName.Replace("Sample", "").Trim());
                    if (sample <= sampleTo && sample >= sampleFrom)
                    {
                        int row = ws.Cells[s].Start.Row;
                        int col = ws.Cells[s].Start.Column;

                        // 2. Xác định tọa độ 2 ô bên dưới cột này
                        // Ví dụ: Ô chứa "Sample" ở dòng `row`, thì 2 ô dưới là `row + 1` và `row + 2` (tương ứng với Picture và Graph)
                        int targetRow1 = row + 1;
                        int targetRow2 = row + 2;

                        // // Lấy tọa độ dạng góc của 2 ô này để so sánh với vị trí của hình ảnh
                        // var cell1 = ws.Cells[targetRow1, col];
                        // var cell2 = ws.Cells[targetRow2, col];

                        // 3. Duyệt qua tất cả các hình ảnh (Drawing) có trong Worksheet để xóa nếu nằm trong vùng tọa độ này
                        // EPPlus quản lý hình ảnh trong ws.Drawings
                        for (int i = ws.Drawings.Count - 1; i >= 0; i--)
                        {
                            var drawing = ws.Drawings[i];

                            // Kiểm tra nếu hình ảnh nằm bắt đầu từ cột hiện tại và nằm trong khoảng 2 hàng bên dưới
                            if (drawing.From.Column == col - 1 && (drawing.From.Row >= targetRow1 - 1 &&
                                                                   drawing.From.Row <= targetRow2 - 1))
                            {
                                ws.Drawings.Remove(i);
                            }
                        }
                    }
                }
                catch
                {
                }
            }
        }

        public void Upload(string logfileAdd, string category, string itemCode, string lotNo)
        {
            if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo))
            {
                throw new Exception("Nhập itemcode lotno");
            }
            category = category.Split('_')[1];
            SaveSheet(logfileAdd, itemCode, lotNo, category);
        }

        public void Load(string itemCode, string lotNo)
        {
            NasRepository nas = new NasRepository();
            string addressSave = $"{nas._nasAddress}\\SEEV Data\\Report\\NPI\\BUILD_2_SHIP\\{itemCode}_{lotNo}";
            string[] array = Directory.GetFiles(addressSave);
            _DATA.Clear();
            foreach (string s in array)
            {
                string nameFile = FileFolderRepository.GetFileNameWithoutExtension(s);
                _DATA.Add(nameFile, "OK");
            }
        }
    }
}