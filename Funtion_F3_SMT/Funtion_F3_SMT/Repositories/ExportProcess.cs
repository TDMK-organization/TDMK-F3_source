using IniLibs;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Reflection;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;
using OfficeOpenXml.Drawing;
using OK2SHIP_SMT.Services;
using System.Data;
using OK2SHIP_SMT.Libary;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace OK2SHIP_SMT.Repositories
{
    class ExportProcess
    {
        public string msZg = "";
        private string FORMAT_LOACTION = null;
        private string EXPORT_LOACTION = null;
        IniFile TDMK_init = new IniFile();
        public string _EXTENSION = "";

        public ExportProcess(string type = "NPI")
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            //config file
            string app_path =
                System.Windows.Forms.Application.StartupPath.Replace(@"\FPCA OK2SHIP Auto System\VHX-IMADA", "");
            string config_path = Path.Combine(app_path, "config.ini");
            TDMK_init = new IniFile(config_path);

            FORMAT_LOACTION = TDMK_init.Read("Format_Folder", "SMT_Config") + $"\\SEEV Data\\Format\\{type}";
            EXPORT_LOACTION = TDMK_init.Read("Report_Location", "SMT_Config") + $"\\SEEV Data\\Report\\{type}";
            msZg =
                $"Đường dẫn config: {config_path} \n Đường dẫn format: {FORMAT_LOACTION} \n Đường dẫn export: {EXPORT_LOACTION}";
            if (FileFolderRepository.checkLocationIsValid(FORMAT_LOACTION) &&
                FileFolderRepository.checkLocationIsValid(EXPORT_LOACTION))
            {
            }
            else
            {
                throw new Exception("Đường dẫn không hợp lệ, kiểm tra lại config.ini");
            }
        }
        public static void GetRealDimension(ExcelWorksheet sheet, out int realEndRow, out int realEndColumn, int limitRow = 5000, int limitCol = 100)
        {
            realEndRow = 1;
            realEndColumn = 1;

            if (sheet.Dimension == null) return;

            // Khoanh vùng quét tối đa để chống kẹt CPU nếu Dimension lên tới cả triệu dòng
            int maxRow = Math.Min(sheet.Dimension.End.Row, limitRow);
            int maxCol = Math.Min(sheet.Dimension.End.Column, limitCol);

            // 1. Quét lùi từ dưới lên để tìm Dòng Thật
            for (int r = maxRow; r >= 1; r--)
            {
                bool hasData = false;
                for (int c = 1; c <= maxCol; c++)
                {
                    var cellValue = sheet.Cells[r, c].Value;
                    if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                    {
                        hasData = true;
                        break;
                    }
                }
                if (hasData)
                {
                    realEndRow = r;
                    break; // Trúng đích, thoát ngay
                }
            }

            // 2. Quét lùi từ phải qua trái để tìm Cột Thật
            for (int c = maxCol; c >= 1; c--)
            {
                bool hasData = false;
                for (int r = 1; r <= realEndRow; r++)
                {
                    var cellValue = sheet.Cells[r, c].Value;
                    if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                    {
                        hasData = true;
                        break;
                    }
                }
                if (hasData)
                {
                    realEndColumn = c;
                    break; // Trúng đích, thoát ngay
                }
            }
        }
        public static void CleanPhantomDimension(ExcelWorksheet sheet)
        {
            if (sheet.Dimension == null) return;

            // 1. Tìm kích thước thật sự (chỉ chứa chữ/dữ liệu)
            GetRealDimension(sheet, out int realEndRow, out int realEndColumn);

            int phantomEndRow = sheet.Dimension.End.Row;
            int phantomEndCol = sheet.Dimension.End.Column;

            // 2. Chặt đứt toàn bộ dòng ảo phía dưới
            if (phantomEndRow > realEndRow)
            {
                // Xóa từ dòng (realEndRow + 1) với số lượng dòng thừa
                sheet.DeleteRow(realEndRow + 1, phantomEndRow - realEndRow);
            }

            // 3. Chặt đứt toàn bộ cột ảo phía bên phải
            if (phantomEndCol > realEndColumn)
            {
                sheet.DeleteColumn(realEndColumn + 1, phantomEndCol - realEndColumn);
            }
        }
        public static ExcelPackage openPackage(string location)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            return new ExcelPackage(location);
        }

        public static byte[] ConvertDataRowToByte(DataRow row, string columnName)
        {
            byte[] image;
            try
            {
                image = TDMK_ImageConverter.ImageToByteArray((Image)row[columnName], ImageFormat.Jpeg);
            }
            catch
            {
                image = (byte[])row[columnName];
            }

            return image;
        }

        public static void CopyColumn(ExcelWorksheet worksheet, ExcelRangeBase rangeStart, string address)
        {
            worksheet.Cells[rangeStart.Address].Copy(worksheet.Cells[address]);
            worksheet.Cells[rangeStart.Address].CopyStyles(worksheet.Cells[address]);
            for (int i = 0; i < rangeStart.Columns; i++)
            {
                for (int j = 0; j < rangeStart.Rows; j++)
                {
                    string add = AddRow(AddColumn(rangeStart.Address.Split(':')[0], i), j);
                    int colu = worksheet.Cells[add].Start.Column;
                    int coluz = worksheet.Cells[address].Start.Column + i;
                    worksheet.Column(coluz).Width = worksheet.Column(colu).Width;
                }
            }
        }

        private static List<int> GetIntegersFromStringRegex(string input)
        {
            List<int> integers = new List<int>();
            MatchCollection matches = Regex.Matches(input, @"\d+"); // @"\d+" matches one or more digits

            foreach (Match match in matches)
            {
                if (int.TryParse(match.Value, out int number))
                {
                    integers.Add(number);
                }
            }

            return integers;
        }

        public static int DistanceRow(string address1, string address2)
        {
            int s1 = GetIntegersFromStringRegex(address1)[0];
            int s2 = GetIntegersFromStringRegex(address2)[0];
            return (s2 - s1);
        }

        public void CopyAndInsert(ExcelWorksheet workSheet, string addressRange, ref string addressStart,
            bool insert = false)
        {
            //
            workSheet.InsertRow(workSheet.Cells[addressStart].End.Row, 1);
            addressStart = AddRow(addressStart, 1);

            int rowNum = 1;
            int colNum = 1;
            if (addressRange.Contains(":"))
            {
                string[] addressRangeList = addressRange.Split(':');
                int startRow = workSheet.Cells[addressRangeList[0]].End.Row;
                int endRow = workSheet.Cells[addressRangeList[1]].End.Row;
                int startCol = workSheet.Cells[addressRangeList[0]].End.Column;
                int endCol = workSheet.Cells[addressRangeList[1]].End.Column;
                rowNum = endRow - startRow + 1;
                colNum = endCol - startCol + 1;
            }

            if (insert)
            {
                workSheet.InsertRow(workSheet.Cells[addressStart].End.Row, rowNum);
            }

            string addressInsert = AddColumn(addressStart, colNum + 2);
            addressInsert = AddRow(addressInsert, rowNum + 2);

            //parste value
            addressInsert = $"{addressStart}:{addressInsert}";
            workSheet.Cells[addressRange].Copy(workSheet.Cells[addressInsert]);
            //workSheet.Cells[addressRange].CopyStyles(workSheet.Cells[addressInsert]);

            //resize
            int startRowz = workSheet.Cells[addressStart].End.Row;
            int colRowz = workSheet.Cells[addressRange].Start.Row;
            for (int i = 0; i <= rowNum; i++)
            {
                workSheet.Row(startRowz + i).Height = workSheet.Row(colRowz + i).Height;
                workSheet.Row(startRowz + i).StyleID = workSheet.Row(colRowz + i).StyleID;
                workSheet.Row(startRowz + i).Style.Font.Bold = workSheet.Row(colRowz + i).Style.Font.Bold;
                workSheet.Row(startRowz + i).Style.Font.Size = workSheet.Row(colRowz + i).Style.Font.Size;
                workSheet.Row(startRowz + i).Style.Font.Family = workSheet.Row(colRowz + i).Style.Font.Family;
                workSheet.Row(startRowz + i).Style.HorizontalAlignment =
                    workSheet.Row(colRowz + i).Style.HorizontalAlignment;
                workSheet.Row(startRowz + i).Style.VerticalAlignment =
                    workSheet.Row(colRowz + i).Style.VerticalAlignment;
            }

            try
            {
                int startR = workSheet.Cells[addressRange.Split(':')[0]].Start.Row;
                int startC = workSheet.Cells[addressRange.Split(':')[0]].Start.Column;
                int endR = workSheet.Cells[addressRange.Split(':')[1]].Start.Row;
                int endC = workSheet.Cells[addressRange.Split(':')[1]].Start.Column;
                int addStartR = workSheet.Cells[addressStart].Start.Row;
                int addStartC = workSheet.Cells[addressStart].Start.Column;
                for (int r = 0; r <= endR - startR + 1; r++)
                {
                    for (int c = 0; c <= endC - startC + 1; c++)
                    {
                        ExcelRange CellFrom = workSheet.Cells[startR + r, startC + c];
                        ExcelRange CellTo = workSheet.Cells[addStartR + r, addStartC + c];
                        CopyStyleOfACell(CellFrom, CellTo);
                    }
                }
            }
            catch
            {
            }

            //workSheet.Cells["A40"].StyleID = workSheet.Cells["A20"].StyleID;
            addressStart = AddRow(addressStart, rowNum);
        }

        private void CopyStyleOfACell(ExcelRange cellFrom, ExcelRange cellTo)
        {
            try
            {
                cellFrom.CopyStyles(cellTo);
            }
            catch
            {
            }
        }

        public static void CopyRowStyle(ExcelWorksheet workSheet, int rowForm, int rowTo)
        {
            workSheet.Row(rowTo).Height = workSheet.Row(rowForm).Height;
            //workSheet.Row(rowTo).StyleID = workSheet.Row(rowForm).StyleID;
            //workSheet.Row(rowTo).Style.Font.Bold = workSheet.Row(rowForm).Style.Font.Bold;
            //workSheet.Row(rowTo).Style.HorizontalAlignment = workSheet.Row(rowForm).Style.HorizontalAlignment;
            //workSheet.Row(rowTo).Style.VerticalAlignment = workSheet.Row(rowForm).Style.VerticalAlignment;
            //workSheet.Row(rowTo).Style.Border.Top.Style = workSheet.Row(rowForm).Style.Border.Top.Style;
            //workSheet.Row(rowTo).Style.Border.Bottom.Style = workSheet.Row(rowForm).Style.Border.Bottom.Style;
            //workSheet.Row(rowTo).Style.Border.Left.Style = workSheet.Row(rowForm).Style.Border.Top.Style;
            //workSheet.Row(rowTo).Style.Border.Right.Style = workSheet.Row(rowForm).Style.Border.Right.Style;
        }

        public static string getRangeBaseAddressByCellAddress(ExcelWorksheet worksheet, string address)
        {
            ExcelAddress addressEx = worksheet.Cells[address];
            string z = worksheet.MergedCells[addressEx.Start.Row, addressEx.Start.Column];
            return worksheet.MergedCells[addressEx.Start.Row, addressEx.Start.Column];
        }

        public ExcelPackage OpenFileExcel(string location)
        {
            return new ExcelPackage(location);
        }

        public ExcelPackage FindFormatWithItemCode(string itemCode)
        {
            itemCode = itemCode.Trim();
            if (Directory.Exists(FORMAT_LOACTION))
            {
                string[] files = Directory.GetFiles(FORMAT_LOACTION);
                foreach (var item in files)
                {
                    string fileName = Path.GetFileNameWithoutExtension(item);

                    if (item.Contains(itemCode))
                    {
                        _EXTENSION = GetFileExtension(item);
                        return new ExcelPackage($"{item}");
                    }
                }
            }

            throw new Exception("Không tìm thấy format File!");
        }

        public static DataRow CloneDataRow(DataRow sourceRow)
        {
            DataTable table = sourceRow.Table;
            DataRow newRow = table.NewRow();
            newRow.ItemArray = sourceRow.ItemArray;
            return newRow;
        }

        public static string GetFileExtension(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                return string.Empty; // or throw an exception
            }

            return Path.GetExtension(filePath);
        }

        public ExcelPackage FindFormatProcess(string process, string itemcode, string lotno)
        {
            switch (process)
            {
                case "OQC B2B Mating-Unmating":
                    return FindFormatWithItemCode(itemcode);
                case "Bar Code Verification":
                    return FindFormatWithItemCode(itemcode);
                case "SEM BSE & Binarization":
                    return FindFormatWithItemCode(itemcode);
                case "Table of Contents":
                    return FindFormatWithItemCode(itemcode);
                default:
                    return FindFormatWithItemCode(itemcode);
            }
        }

        public static string ConvertAddressRangeBase(ExcelWorksheet ws, string rangeAddress, string addressSelect)
        {
            int rStart = 0, cStart = 0;
            int rEnd = 0, cEnd = 0;
            string startAddress = "", endAddress = "";
            int i = 0;
            if (rangeAddress.Contains(":"))
            {
                throw new Exception("Chưa triển khai convert rangbse");
            }
            else if (rangeAddress.Contains("+"))
            {
                string[] address = rangeAddress.Split('+');
                // tìm kiếm địa chỉ đầu và cuối
                while (string.IsNullOrEmpty(startAddress) || string.IsNullOrEmpty(endAddress))
                {
                    if (string.IsNullOrEmpty(startAddress))
                    {
                        try
                        {
                            startAddress = address[i].Trim();
                        }
                        catch
                        {
                        }
                    }

                    if (string.IsNullOrEmpty(endAddress))
                    {
                        try
                        {
                            endAddress = address[address.Length - i].Trim();
                        }
                        catch
                        {
                        }
                    }

                    i++;
                }

                rStart = ws.Cells[startAddress].Start.Row - ws.Cells[addressSelect].Start.Row;
                cStart = ws.Cells[startAddress].Start.Column - ws.Cells[addressSelect].Start.Column;
                rEnd = ws.Cells[endAddress].Start.Row - ws.Cells[addressSelect].Start.Row;
                cEnd = ws.Cells[endAddress].Start.Column - ws.Cells[addressSelect].Start.Column;
                return $"R[{rStart}]C[{cStart}]:R[{rEnd}]C[{cEnd}]";
            }
            else
            {
                rStart = ws.Cells[rangeAddress].Start.Row - ws.Cells[addressSelect].Start.Row;
                cStart = ws.Cells[rangeAddress].Start.Column - ws.Cells[addressSelect].Start.Column;
                return $"R[{rStart}]C[{cStart}]";
            }

            throw new Exception("Eroo");
        }

        public string SaveExcelWorksheet(ExcelPackage excelPackage, string sheetName, string nameFile,
            string type = "NPI", bool open = true)
        {
            string[] sheetNames = sheetName.Replace(" ", "_").ToUpper().Split(':');
            DateTime nowDate = DateTime.Now;
            IList<ExcelWorksheet> worksheets = new List<ExcelWorksheet>();
            foreach (var item in excelPackage.Workbook.Worksheets)
            {
                if (!sheetNames.Contains(item.Name.Trim().Replace(" ", "_").ToUpper()))
                {
                    worksheets.Add(item);
                }
            }

            foreach (ExcelWorksheet item in worksheets)
            {
                excelPackage.Workbook.Worksheets.Delete(item.Name);
            }

            //excelPackage.Workbook.Worksheets.Delete("Rev History");
            int month = nowDate.Month;

            // Lưu package vào địa chỉ được chỉ định
            string[] str = nameFile.Split('.');
            string folderName = $"{EXPORT_LOACTION}\\{sheetNames[0]}\\THANG {month}\\";

            if (!Directory.Exists(folderName))
            {
                try
                {
                    // Tạo thư mục nếu nó không tồn tại
                    Directory.CreateDirectory(folderName);
                }
                catch (Exception ex)
                {
                    // Xử lý lỗi nếu không thể tạo thư mục
                    throw new Exception($"Lỗi: Không thể tạo thư mục '{folderName}'. Lỗi: {ex.Message}");
                }
            }

            FileInfo file = new FileInfo($"{folderName}{nameFile}{_EXTENSION}");
            excelPackage.SaveAs(file);
            if (open)
            {
                Process.Start(file.FullName);
            }

            return file.FullName;
        }

        public static void InsertImageToCell(ExcelWorksheet wsSheet1, ExcelRangeBase tar_rgn, byte[] img_data,
            string pic_name)
        {
            using (tar_rgn)
            {
                int row = tar_rgn.End.Row;
                int column = tar_rgn.End.Column;
                int mergeCellId = wsSheet1.GetMergeCellId(row, column);
                string address = tar_rgn.Address;
                if (mergeCellId > 0)
                {
                    address = wsSheet1.MergedCells[mergeCellId - 1];
                }

                int rows = wsSheet1.Cells[address].Rows;
                int columns = wsSheet1.Cells[address].Columns;
                try
                {
                    wsSheet1.Drawings.Remove(pic_name);
                }
                catch
                {
                }

                int row2 = tar_rgn.Start.Row;
                int column2 = tar_rgn.Start.Column;
                double num = 0.0;
                double num2 = 0.0;
                for (int i = 0; i < columns; i++)
                {
                    num += wsSheet1.Column(column2 + i).Width;
                }

                for (int j = 0; j < rows; j++)
                {
                    num2 += wsSheet1.Row(row2 + j).Height;
                }

                MemoryStream pictureStream = new MemoryStream(img_data);
                ExcelPicture excelPicture = wsSheet1.Drawings.AddPicture(pic_name, pictureStream);
                ExcelWorkbook workbook = wsSheet1.Workbook;
                decimal maxFontWidth = (decimal)workbook.MaxFontWidth;
                int num3 = (int)(num2 / 0.75);
                int num4 = (int)decimal.Truncate((256m * (decimal)num + decimal.Truncate(128m / maxFontWidth)) / 256m *
                                                 maxFontWidth);
                int num5 = (int)(0.05 * (double)Math.Min(num3, num4));
                excelPicture.SetPosition(row2 - 1, num5, column2 - 1, num5);
                excelPicture.SetSize(num4 - 2 * num5, num3 - 2 * num5);
                //excelPicture.SetSize(50);
            }
        }

        public ExcelWorksheet FindSheet(ExcelPackage excelPackage, string sheetName)
        {
            foreach (ExcelWorksheet sheet in excelPackage.Workbook.Worksheets)
            {
                if (sheet.Name.Trim().Replace(" ", "_").ToUpper().Equals(sheetName.Replace(" ", "_").ToUpper()))
                {
                    return sheet;
                }
            }

            return null;
            throw new Exception("Không tìm thấy sheet kiểm tra lại tên sheet");
        }

        ///
        ///Not Fix
        public static string FindAddressByText(ExcelWorksheet excelWorksheet, string text, string address = null)
        {
            for (int i = 0; i < excelWorksheet.Dimension.Start.Column; i++)
            {
                for (int j = 0; j < excelWorksheet.Dimension.Start.Row; j++)
                {
                    if (excelWorksheet.Cells[i, j].Text == text)
                    {
                        return excelWorksheet.Cells[i, j].Address;
                    }
                }
            }

            return null;
        }

        public static string AddColumn(string address, int column)
        {
            ExcelCellAddress cell = new ExcelCellAddress(address);
            return new ExcelCellAddress(cell.Row, cell.Column + column).Address;
        }

        public static string AddRow(string address, int row)
        {
            ExcelCellAddress cell = new ExcelCellAddress(address);
            return new ExcelCellAddress(cell.Row + row, cell.Column).Address;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="colHeader"></param>
        /// <param name="eq"></param>
        /// <returns>text with value is list of address of column</returns>
        public static IDictionary<string, string> FindAddressByText(
            ExcelWorksheet workSheet,
            string[] colHeader,
            bool eq = false,
            string formAddress = "",
            string endAddress = "",
            int maxRows = 100, // Thêm tham số giới hạn số hàng (mặc định 100)
            int maxColumns = 60) // Thêm tham số giới hạn số cột (mặc định 60)
        {
            // 1. Xác định tọa độ bắt đầu và kết thúc mà KHÔNG dùng Dimension
            int startRow = 1;
            int startColumn = 1;
            int endRow = maxRows;
            int endColumn = maxColumns;

            if (!string.IsNullOrEmpty(formAddress))
            {
                var startCell = workSheet.Cells[formAddress];
                startRow = startCell.Start.Row;
                startColumn = startCell.Start.Column;
            }

            if (!string.IsNullOrEmpty(endAddress))
            {
                var endCell = workSheet.Cells[endAddress];
                endRow = endCell.End.Row;
                endColumn = endCell.End.Column;
            }

            // 2. TIỀN XỬ LÝ (TỐI ƯU CPU): 
            // Format các từ khóa tìm kiếm (ToUpper, Trim) đúng 1 lần ở ngoài cùng, 
            // thay vì phải lặp lại hàng vạn lần bên trong vòng lặp.
            var searchTargets = colHeader
                .Where(h => !string.IsNullOrEmpty(h))
                .Select(h => new
                {
                    Original = h.Trim().Replace("\n", ""),
                    Upper = h.Trim().Replace("\n", "").ToUpper()
                })
                .ToList();

            IDictionary<string, string> addressHeader = new Dictionary<string, string>();

            // 3. Quét qua vùng đã được khoanh vùng (Giới hạn tối đa: 100 hàng x 60 cột)
            for (int r = startRow; r <= endRow; r++)
            {
                for (int c = startColumn; c <= endColumn; c++)
                {
                    // Lấy text của ô hiện tại
                    string cellValue = workSheet.Cells[r, c].Text;

                    // Nếu ô trống, bỏ qua ngay lập tức để tiết kiệm chi phí xử lý chuỗi
                    if (string.IsNullOrEmpty(cellValue))
                        continue;

                    // Xử lý giá trị ô để so sánh
                    cellValue = cellValue.Trim().Replace("\n", "").ToUpper();

                    foreach (var target in searchTargets)
                    {
                        bool isMatch = eq
                            ? cellValue.Equals(target.Upper)
                            : cellValue.Contains(target.Upper);

                        if (isMatch)
                        {
                            if (addressHeader.TryGetValue(target.Original, out string existingAddress))
                            {
                                // Cập nhật giá trị nếu đã tồn tại
                                addressHeader[target.Original] = existingAddress + "-" + workSheet.Cells[r, c].Address;
                            }
                            else
                            {
                                // Thêm mới
                                addressHeader.Add(target.Original, workSheet.Cells[r, c].Address);
                            }
                        }
                    }
                }
            }

            return addressHeader;
        }

        public static void AddBorderToImage(ExcelWorksheet worksheet, string imageName, Color color, int size = 1,
            eLineStyle styleLine = eLineStyle.Solid)
        {
            if (worksheet == null)
            {
                throw new Exception($"Sheet not found.");
            }

            ExcelPicture picture = null;

            foreach (var shape in worksheet.Drawings)
            {
                if (shape is ExcelPicture pic && pic.Name == imageName)
                {
                    picture = pic;
                    break;
                }
            }

            if (picture == null)
            {
                throw new Exception($"Image '{imageName}' not found in sheet .");
            }

            picture.Border.LineStyle = eLineStyle.Solid;
            picture.Border.Width = size;
            picture.Border.Fill.Color = color;
        }

        // Hàm để copy một hàng từ vị trí này sang vị trí khác trong cùng một worksheet
        public static void CopyRow(ExcelWorksheet worksheet, int sourceRow, int destinationRow)
        {
            if (worksheet == null)
            {
                throw new ArgumentNullException("worksheet");
            }

            // Lấy vùng chứa dữ liệu của dòng nguồn
            ExcelRange sourceRowRange = worksheet.Cells[sourceRow, 1, sourceRow, worksheet.Dimension.End.Column];
            //Copy dòng
            worksheet.InsertRow(destinationRow, 1);
            ExcelRange destinationRowRange =
                worksheet.Cells[destinationRow, 1, destinationRow, worksheet.Dimension.End.Column];
            destinationRowRange.Value = sourceRowRange.Value;

            // Bạn có thể muốn copy định dạng, kiểu dữ liệu, v.v.
            CopyRowFormat(worksheet, sourceRow, destinationRow);
        }


        static void CopyRowFormat(ExcelWorksheet worksheet, int sourceRow, int destinationRow)
        {
            if (worksheet == null)
            {
                throw new ArgumentNullException("worksheet");
            }

            ExcelRange sourceRowRange = worksheet.Cells[sourceRow, 1, sourceRow, worksheet.Dimension.End.Column];
            ExcelRange destinationRowRange =
                worksheet.Cells[destinationRow, 1, destinationRow, worksheet.Dimension.End.Column];

            destinationRowRange.StyleID = sourceRowRange.StyleID;
            destinationRowRange.Merge = sourceRowRange.Merge;
            destinationRowRange.Hyperlink = sourceRowRange.Hyperlink;
            destinationRowRange.Formula = sourceRowRange.Formula;
            destinationRowRange.FormulaR1C1 = sourceRowRange.FormulaR1C1;
            destinationRowRange.Value = sourceRowRange.Value;
            destinationRowRange.RichText.Text = sourceRowRange.RichText.Text;
            destinationRowRange.Style.Font.Bold = sourceRowRange.Style.Font.Bold;
            destinationRowRange.Style.Font.Italic = sourceRowRange.Style.Font.Italic;
            destinationRowRange.Style.Font.UnderLine = sourceRowRange.Style.Font.UnderLine;
            destinationRowRange.Style.Fill.PatternType = sourceRowRange.Style.Fill.PatternType;
            destinationRowRange.Style.HorizontalAlignment = sourceRowRange.Style.HorizontalAlignment;
            destinationRowRange.Style.VerticalAlignment = sourceRowRange.Style.VerticalAlignment;
            destinationRowRange.Style.WrapText = sourceRowRange.Style.WrapText;
            destinationRowRange.Style.ShrinkToFit = sourceRowRange.Style.ShrinkToFit;
            destinationRowRange.Style.Numberformat.Format = sourceRowRange.Style.Numberformat.Format;
            destinationRowRange.Style.Border.Top.Style = sourceRowRange.Style.Border.Top.Style;
            destinationRowRange.Style.Border.Bottom.Style = sourceRowRange.Style.Border.Bottom.Style;
            destinationRowRange.Style.Border.Left.Style = sourceRowRange.Style.Border.Left.Style;
            destinationRowRange.Style.Border.Right.Style = sourceRowRange.Style.Border.Right.Style;
        }
    }
}