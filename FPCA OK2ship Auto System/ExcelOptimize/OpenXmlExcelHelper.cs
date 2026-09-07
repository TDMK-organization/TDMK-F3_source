using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelOptimize
{
    public class OpenXmlExcelHelper
    {
        public static async Task MergeLargeExcelFilesAsync(string masterPath, string[] smallFilePaths, string outputFilePath, IProgress<int> progress, CancellationToken cancellationToken = default)
        {
            // 1. Tạo bản sao từ file Master gốc sang file Output để bảo vệ file gốc
            if (File.Exists(outputFilePath))
            {
                File.Delete(outputFilePath);
            }
            File.Copy(masterPath, outputFilePath, true);

            // 2. Đẩy toàn bộ quá trình xuống Background Thread để WinForms không bị đơ
            await Task.Run(() =>
            {
                int totalFiles = smallFilePaths.Length;
                int processedFiles = 0;

                foreach (string smallFilePath in smallFilePaths)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    try
                    {
                        // Bước A: Mở nhanh file lẻ ở chế độ Read-Only để lấy danh sách tên các Sheet cần di chuyển
                        var sheetNames = new List<string>();
                        using (SpreadsheetDocument sourceDoc = SpreadsheetDocument.Open(smallFilePath, false))
                        {
                            WorkbookPart sourceWorkbookPart = sourceDoc.WorkbookPart;
                            if (sourceWorkbookPart?.Workbook?.Sheets != null)
                            {
                                foreach (Sheet sheet in sourceWorkbookPart.Workbook.Sheets.Elements<Sheet>())
                                {
                                    if (sheet.Name != null && sheet.Name.HasValue)
                                    {
                                        sheetNames.Add(sheet.Name.Value);
                                    }
                                }
                            }
                        }

                        // Bước B: Duyệt qua từng tên Sheet vừa lấy được và gọi hàm MoveSheetToMaster
                        foreach (string sheetName in sheetNames)
                        {
                            // Gọi hàm độc lập để nhấc Sheet vào file đích
                            MoveSheetToMaster(smallFilePath, sheetName, outputFilePath, sheetName);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Ghi nhận lỗi của riêng file đó để không làm sập toàn bộ tiến trình gộp 40 file
                        Console.WriteLine($"[Lỗi] Không thể xử lý file {smallFilePath}: {ex.Message}");
                    }

                    // 3. Cập nhật tiến độ xử lý về giao diện UI WinForms
                    processedFiles++;
                    if (progress != null)
                    {
                        int percentComplete = (int)((double)processedFiles / totalFiles * 100);
                        progress.Report(percentComplete);
                    }
                }
            }, cancellationToken);
        }

        public static void MoveSheetToMaster(string sourcePath, string sheetNameInSource, string masterPath, string newSheetNameInMaster)
        {
            // Mở file Master ở chế độ Read/Write
            using (SpreadsheetDocument masterDoc = SpreadsheetDocument.Open(masterPath, true))
            {
                WorkbookPart masterWorkbookPart = masterDoc.WorkbookPart;

                // Mở file lẻ ở chế độ Read-Only để đảm bảo an toàn dữ liệu
                using (SpreadsheetDocument sourceDoc = SpreadsheetDocument.Open(sourcePath, false))
                {
                    WorkbookPart sourceWorkbookPart = sourceDoc.WorkbookPart;

                    // 1. Tìm Sheet nguồn dựa vào tên hiển thị
                    Sheet sourceSheet = sourceWorkbookPart.Workbook.Sheets
                        .Elements<Sheet>()
                        .FirstOrDefault(s => s.Name == sheetNameInSource);

                    if (sourceSheet == null)
                        throw new Exception($"Không tìm thấy Sheet '{sheetNameInSource}' trong file nguồn.");

                    // Lấy WorksheetPart tương ứng từ ID của Sheet
                    WorksheetPart sourceWorksheetPart = (WorksheetPart)sourceWorkbookPart.GetPartById(sourceSheet.Id);

                    // 2. Tiến hành Import Part xuyên Document
                    WorksheetPart importedWorksheetPart = masterWorkbookPart.AddPart(sourceWorksheetPart);

                    // --- KHU VỰC XỬ LÝ NÂNG CAO (SẼ VIẾT THÊM Ở DƯỚI) ---
                    // TODO: Đi sâu vào importedWorksheetPart để đồng bộ SharedStrings và Styles
                    // TODO: Đi sâu vào để import các ImagePart và DrawingsPart đính kèm
                    // --------------------------------------------------

                    // 3. Tạo một định danh duy nhất (Relationship ID) trong file Master cho Sheet mới
                    string masterPartId = masterWorkbookPart.GetIdOfPart(importedWorksheetPart);

                    // 4. Đăng ký Sheet vào danh sách quản lý của Workbook Master
                    Sheets masterSheets = masterWorkbookPart.Workbook.Sheets;
                    if (masterSheets == null)
                    {
                        masterSheets = new Sheets();
                        masterWorkbookPart.Workbook.AppendChild(masterSheets);
                    }

                    string finalSheetName = newSheetNameInMaster;

                    // KIỂM TRA VÀ XÓA SHEET CŨ NẾU TỒN TẠI TRƯỚC KHI THÊM MỚI
                    Sheet existingSheet = masterSheets.Elements<Sheet>().FirstOrDefault(s => s.Name == finalSheetName);
                    if (existingSheet != null)
                    {
                        string oldPartId = existingSheet.Id?.Value;

                        // A. Xóa thẻ hiển thị (Tab) khỏi danh sách Workbook
                        existingSheet.Remove();

                        // B. Xóa luồng dữ liệu vật lý (WorksheetPart) để giải phóng dung lượng file Excel
                        if (!string.IsNullOrEmpty(oldPartId))
                        {
                            masterWorkbookPart.DeletePart(oldPartId);
                        }
                    }

                    // Tính toán SheetId tiếp theo (phải là số nguyên duy nhất)
                    uint nextSheetId = masterSheets.Elements<Sheet>().Any()
                        ? masterSheets.Elements<Sheet>().Max(s => s.SheetId.Value) + 1
                        : 1;

                    // Thêm đối tượng Sheet mới vào cây cấu trúc
                    Sheet newMasterSheet = new Sheet()
                    {
                        Id = masterPartId,
                        SheetId = nextSheetId,
                        Name = finalSheetName
                    };
                    masterSheets.AppendChild(newMasterSheet);
                }

                // Lưu lại mọi thay đổi cấu trúc trên file vật lý Master
                masterWorkbookPart.Workbook.Save();
            }
        }
    }
}