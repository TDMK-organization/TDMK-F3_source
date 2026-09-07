using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
namespace XMLHelper
{
    public class OpenXmlExcelHelper
    {
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
                    // Hàm ImportPart sẽ tự động nhân bản phần lõi XML của Sheet sang cấu trúc file Master
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

                    // Tính toán SheetId tiếp theo (phải là số nguyên duy nhất)
                    uint nextSheetId = 1;
                    if (masterSheets.Elements<Sheet>().Any())
                    {
                        nextSheetId = masterSheets.Elements<Sheet>().Max(s => s.SheetId.Value) + 1;
                    }

                    // Kiểm tra trùng tên Sheet trong Master
                    string finalSheetName = newSheetNameInMaster;
                    int counter = 1;
                    while (masterSheets.Elements<Sheet>().Any(s => s.Name == finalSheetName))
                    {
                        finalSheetName = $"{newSheetNameInMaster}_{counter}";
                        counter++;
                    }

                    // Thêm đối tượng Sheet mới vào cây cấu trúc
                    Sheet newMasterSheet = new Sheet()
                    {
                        Id = masterPartId,
                        SheetId = nextSheetId,
                        Name = finalSheetName
                    };
                    masterSheets.AppendChild(newMasterSheet);
                }

                // Lưu lại mọi thay đổi cấu trúc cấu trúc trên file vật lý Master
                masterWorkbookPart.Workbook.Save();
            }
        }
    }
}

