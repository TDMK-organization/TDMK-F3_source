using IniLibs;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using TDMK_EPPLUS_7;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Reflection;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using OfficeOpenXml.Drawing;
using OK2SHIP_SMT.Services;
using System.Data;
using OfficeOpenXml;
namespace OK2SHIP_SMT.Repositories
{
    class ExportProcess
    {

        private string FORMAT_LOACTION = null;
        private string EXPORT_LOACTION = null;
        IniFile TDMK_init = new IniFile();
        private string _EXTENSION = "";
        public static TDMK_EPPLUS7_lib _EPPLUS_7;
        public ExportProcess(string type = "NPI")
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //config file
            string app_path = System.Windows.Forms.Application.StartupPath.Replace("\\VHX-IMADA", "");
            string config_path = Path.Combine(app_path, "config.ini");
            TDMK_init = new IniFile(config_path);
            FORMAT_LOACTION = TDMK_init.Read("Format_Folder", "SMT_Config") + $"\\SEEV Data\\Format\\{type}";
            EXPORT_LOACTION = TDMK_init.Read("Report_Location", "SMT_Config") + $"\\SEEV Data\\Format\\{type}";
        }

        public ExcelPackage FindFormatBarCode()
        {
            return new ExcelPackage($"{FORMAT_LOACTION}\\SEEV-OK2Ship_Format change.xlsm");
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
            return null;
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
                    return null;
            }
        }

        public void SaveExcelWorksheet(ExcelPackage excelPackage, string sheetName, string nameFile, string type = "NPI")
        {

            DateTime nowDate = DateTime.Now;
            foreach (var item in excelPackage.Workbook.Worksheets)
            {
                if (!item.Name.Trim().Equals(sheetName))
                {
                    excelPackage.Workbook.Worksheets.Delete(item);
                }
            }
            //excelPackage.Workbook.Worksheets.Delete("Rev History");
            int month = nowDate.Month;

            // Lưu package vào địa chỉ được chỉ định
            string[] str = nameFile.Split('.');
            string folderName = $"{EXPORT_LOACTION}\\{sheetName}\\THANG {month}\\";

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
            Process.Start(file.FullName);


        }

        public void InsertImageToCell(ExcelWorksheet wsSheet1, ExcelRangeBase tar_rgn, byte[] img_data, string pic_name)
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
                decimal maxFontWidth = workbook.MaxFontWidth;
                int num3 = (int)(num2 / 0.75);
                int num4 = (int)decimal.Truncate((256m * (decimal)num + decimal.Truncate(128m / maxFontWidth)) / 256m * maxFontWidth);
                int num5 = (int)(0.05 * (double)Math.Min(num3, num4));
                excelPicture.SetPosition(row2 - 1, num5, column2 - 1, num5);
                excelPicture.SetSize(num4 - 2 * num5, num3 - 2 * num5);

            }
        }

        public ExcelWorksheet FindSheet(ExcelPackage excelPackage, string sheetName)
        {
            foreach (ExcelWorksheet sheet in excelPackage.Workbook.Worksheets)
            {
                if (sheet.Name.Trim().Equals(sheetName))
                {
                    return sheet;
                }
            }
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
        public static IDictionary<string, string> FindAddressByText(ExcelWorksheet workSheet, string[] colHeader, bool eq = false)
        {
            IList<string> colHeaderz = colHeader.ToList();
            IDictionary<string, string> addressHeader = new Dictionary<string, string>();
            for (int i = 1; i <= workSheet.Dimension.Columns + 1; i++)
            {
                for (int j = 1; j <= workSheet.Dimension.Rows + 1; j++)
                {
                    string cellValue = workSheet.Cells[j, i].Text.Trim().Replace("\n", "");
                    if (!string.IsNullOrEmpty(cellValue.ToString()))
                    {
                        foreach (string str in colHeaderz)
                        {
                            string strz = str.Trim().Replace("\n", "");
                            bool prime = false;
                            if (eq)
                            {
                                prime = cellValue.Equals(str);
                            }
                            else
                            {
                                prime = cellValue.Contains(str);
                            }
                            if (prime)
                            {
                                if (addressHeader.TryGetValue(strz, out string value))
                                {
                                    // update value
                                    addressHeader[strz] = addressHeader[strz] + "-" + workSheet.Cells[j, i].Address;
                                }
                                else
                                {
                                    addressHeader.Add(strz, workSheet.Cells[j, i].Address);
                                }
                            }
                        }
                    }
                    if (addressHeader.Count() == colHeader.Count())
                    {
                        return addressHeader;
                    }

                }
            }
            return addressHeader;
        }

        public static void AddBorderToImage(ExcelWorksheet worksheet, string imageName, Color color, int size = 1, eLineStyle styleLine = eLineStyle.Solid)
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
    }
}
