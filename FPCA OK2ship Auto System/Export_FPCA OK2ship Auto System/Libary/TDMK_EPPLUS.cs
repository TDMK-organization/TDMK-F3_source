using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Export_FPCA_OK2ship_Auto_System.Libary
{
    public class TDMK_EPPLUS
    {

        public struct MergeAreas_Info
        {
            public int col_qty;
            public int row_qty;
        }
        public static ExcelWorkbook open_excel_file(string file_name)
        {
            ExcelWorkbook result = null;
            FileInfo excel_file = new FileInfo(file_name);
            if (File.Exists(file_name))
            {
                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                ExcelPackage myexcel = new ExcelPackage(excel_file);
                result = myexcel.Workbook;
            }
            return result;
        }
        public static Image get_pic(ExcelWorksheet wrksht, string pic_name)
        {
            Image result = null;
            var pics = wrksht.Drawings.Where(x => x.DrawingType == eDrawingType.Picture).Select(x => x.As.Picture);
            foreach (var pic in pics)
            {
                var name = pic.Name;
                if (name == pic_name)
                {
                    var cur_img_data = pic.Image.ImageBytes;
                    result = byteArrayToImage(cur_img_data);
                    break;
                }
            }
            return result;
        }
        public byte[] get_pic_data(ExcelWorksheet wrksht, string pic_name)
        {
            byte[] result = null;
            var pics = wrksht.Drawings.Where(x => x.DrawingType == eDrawingType.Picture).Select(x => x.As.Picture);
            foreach (var pic in pics)
            {
                var name = pic.Name;
                if (name == pic_name)
                {
                    result = pic.Image.ImageBytes;
                    break;
                }
            }
            return result;
        }
        public static Image byteArrayToImage(byte[] byteArrayIn)
        {
            using (MemoryStream mStream = new MemoryStream(byteArrayIn))
            {
                return Image.FromStream(mStream);
            }
        }
        public string Find_Cell_Addr(string search_key, string start_addr, ExcelWorksheet tar_wrksht, bool left_to_right, bool en_contain = false)
        {
            List<char> reject_chars = new List<char>() { ' ', '-', '_', '\r', '\n' };
            string result = "";
            ExcelRangeBase cycle_rgn = tar_wrksht.Cells[start_addr];
            for (int i = 0; i < 100; i++)
            {
                ExcelRangeBase sel_rgn = cycle_rgn.Offset(i, 0);
                if (left_to_right)
                {
                    sel_rgn = cycle_rgn.Offset(0, i);
                }
                string sel_rgn_val = checkDBNull(sel_rgn.Value);
                if (en_contain)
                {
                    if (remove_special_chars(sel_rgn_val, reject_chars).ToUpper().Contains(remove_special_chars(search_key, reject_chars).ToUpper()))
                    {
                        result = sel_rgn.Address;
                        break;
                    }
                }
                else
                {
                    if (remove_special_chars(sel_rgn_val, reject_chars).ToUpper() == remove_special_chars(search_key, reject_chars).ToUpper())
                    {
                        result = sel_rgn.Address;
                        break;
                    }
                }
            }
            return result;
        }
        public string Find_Cell_Location(string search_key, ExcelWorksheet tar_wrksht, bool en_contain = false)
        {
            ExcelAddressBase wrk_sht_dimension = tar_wrksht.Dimension;
            int rows = wrk_sht_dimension.Rows;
            int cols = wrk_sht_dimension.Columns;
            List<char> reject_chars = new List<char>() { ' ', '-', '_', '\r', '\n', '&' };
            string result = "";
            string _search_key = remove_special_chars(search_key, reject_chars).ToUpper();
            for (int r_inx = 1; r_inx < rows; r_inx++)
            {
                for (int c_inx = 1; c_inx < cols; c_inx++)
                {
                    ExcelRangeBase cur_cell = tar_wrksht.Cells[r_inx, c_inx];
                    string cur_cell_val = remove_special_chars(checkDBNull(cur_cell.Value), reject_chars).ToUpper();
                    if (en_contain)
                    {
                        if (cur_cell_val.Contains(_search_key))
                        {
                            result = cur_cell.Address;
                            return result;
                        }
                    }
                    else
                    {
                        if (cur_cell_val == _search_key)
                        {
                            result = cur_cell.Address;
                            return result;
                        }
                    }

                }
            }
            return result;
        }
        public string FindByOffset(string start_addr, ExcelWorksheet tar_wrksht, bool left_to_right, string search_key = "")
        {
            List<char> reject_chars = new List<char>() { ' ', '-', '_', '\r', '\n' };
            string result = "";
            ExcelRangeBase cycle_rgn = tar_wrksht.Cells[start_addr];
            int offset = 0;
            int empty_count = 0;
            while (true)
            {
                ExcelRangeBase sel_rgn = cycle_rgn.Offset(offset, 0);
                if (left_to_right)
                {
                    sel_rgn = cycle_rgn.Offset(0, offset);
                }
                string sel_rgn_val = checkDBNull(sel_rgn.Value);
                if (sel_rgn_val == "")
                {
                    empty_count++;
                    if (empty_count > 3)
                    {
                        break;
                    }
                }
                else
                {
                    empty_count = 0;
                }
                if (search_key != "")
                {
                    if (remove_special_chars(sel_rgn_val, reject_chars).ToUpper() == remove_special_chars(search_key, reject_chars).ToUpper())
                    {
                        result = sel_rgn.Address;
                        break;
                    }
                }
                else
                {
                    if (empty_count == 0)
                    {
                        result = sel_rgn.Address;
                        break;
                    }
                }
                offset++;
            }
            return result;
        }
        public string checkDBNull(object src_str)
        {
            string _result;
            if (src_str != null)
            {
                _result = src_str.ToString();
            }
            else
            {
                _result = "";
            }
            return _result;
        }
        public string remove_special_chars(string src_string, List<char> reject_chars)
        {
            string result = "";
            result = new string(src_string.ToArray().Where(x => reject_chars.IndexOf(x) == -1).ToArray());
            return result;
        }
        public void insert_pic(ExcelWorksheet wsSheet1, ExcelRange tar_rgn, Image src_pic, string pic_name)
        {
            using (tar_rgn)
            {
                int r_count = tar_rgn.End.Row;
                int col_count = tar_rgn.End.Column;
                var idx = wsSheet1.GetMergeCellId(r_count, col_count);
                string mergedCellAddress = tar_rgn.Address;
                if (idx > 0)
                {
                    mergedCellAddress = wsSheet1.MergedCells[idx - 1];
                }
                int r = wsSheet1.Cells[mergedCellAddress].Rows;
                int c = wsSheet1.Cells[mergedCellAddress].Columns;
                try
                {
                    wsSheet1.Drawings.Remove(pic_name);
                }
                catch
                {

                }
                int _rowIndex = tar_rgn.Start.Row;
                int _colIndex = tar_rgn.Start.Column;
                double col_w = 0;// wsSheet1.Column(_colIndex).Width;
                double row_h = 0;// wsSheet1.Row(_rowIndex).Height;
                for (int i = 0; i < c; i++)
                {
                    col_w += wsSheet1.Column(_colIndex + i).Width;
                }
                for (int i = 0; i < r; i++)
                {
                    row_h += wsSheet1.Row(_rowIndex + i).Height;
                }
                byte[] img_data = imgToByteConverter(src_pic);
                var steam = new MemoryStream(img_data);
                ExcelPicture pic = wsSheet1.Drawings.AddPicture(pic_name, steam);//img
                ExcelWorkbook wrkbk = wsSheet1.Workbook;
                decimal mdw = (decimal)wrkbk.MaxFontWidth;
                int pixelHeight = (int)(row_h / 0.75);
                int pixelWidth = (int)decimal.Truncate(((256 * (decimal)col_w + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
                int offset = (int)(0.05 * Math.Min(pixelHeight, pixelWidth));
                pic.SetPosition(_rowIndex - 1, offset, _colIndex - 1, offset);
                pic.SetSize(pixelWidth - 2 * offset, pixelHeight - 2 * offset);
            }
        }
        public void insert_pic(ExcelWorksheet wsSheet1, ExcelRange tar_rgn, byte[] img_data, string pic_name)
        {
            using (tar_rgn)
            {
                int r_count = tar_rgn.End.Row;
                int col_count = tar_rgn.End.Column;
                var idx = wsSheet1.GetMergeCellId(r_count, col_count);
                string mergedCellAddress = tar_rgn.Address;
                if (idx > 0)
                {
                    mergedCellAddress = wsSheet1.MergedCells[idx - 1];
                }
                int r = wsSheet1.Cells[mergedCellAddress].Rows;
                int c = wsSheet1.Cells[mergedCellAddress].Columns;
                try
                {
                    wsSheet1.Drawings.Remove(pic_name);
                }
                catch
                {

                }
                int _rowIndex = tar_rgn.Start.Row;
                int _colIndex = tar_rgn.Start.Column;
                double col_w = 0;// wsSheet1.Column(_colIndex).Width;
                double row_h = 0;// wsSheet1.Row(_rowIndex).Height;
                for (int i = 0; i < c; i++)
                {
                    col_w += wsSheet1.Column(_colIndex + i).Width;
                }
                for (int i = 0; i < r; i++)
                {
                    row_h += wsSheet1.Row(_rowIndex + i).Height;
                }
                var steam = new MemoryStream(img_data);
                ExcelPicture pic = wsSheet1.Drawings.AddPicture(pic_name, steam);//img
                ExcelWorkbook wrkbk = wsSheet1.Workbook;
                decimal mdw = (decimal)wrkbk.MaxFontWidth;
                int pixelHeight = (int)(row_h / 0.75);
                int pixelWidth = (int)decimal.Truncate(((256 * (decimal)col_w + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
                int offset = (int)(0.05 * Math.Min(pixelHeight, pixelWidth));
                pic.SetPosition(_rowIndex - 1, offset, _colIndex - 1, offset);
                pic.SetSize(pixelWidth - 2 * offset, pixelHeight - 2 * offset);
            }
        }
        public byte[] imgToByteConverter(Image inImg)
        {
            ImageConverter imgCon = new ImageConverter();
            return (byte[])imgCon.ConvertTo(inImg, typeof(byte[]));
        }
        public string Find_Start_Addr(string start_addr, ExcelWorksheet tar_wrksht, bool left_to_right, string search_key = "")
        {
            List<char> reject_chars = new List<char>() { ' ', '-', '_', '\r', '\n' };
            string result = "";
            string tar_rgn_addr = start_addr;
            ExcelRangeBase cycle_rgn = tar_wrksht.Cells[start_addr];
            int c_offset = 0;
            int empty_count = 0;
            while (true)
            {
                ExcelRangeBase sel_rgn = cycle_rgn.Offset(c_offset, 0);
                if (left_to_right)
                {
                    sel_rgn = cycle_rgn.Offset(0, c_offset);
                }
                string sel_rgn_val = checkDBNull(sel_rgn.Value);
                if (sel_rgn_val == "")
                {
                    empty_count++;
                    if (empty_count > 3)
                    {
                        break;
                    }
                }
                else
                {
                    empty_count = 0;
                }
                if (search_key != "")
                {
                    if (remove_special_chars(sel_rgn_val, reject_chars).ToUpper() == remove_special_chars(search_key, reject_chars).ToUpper())
                    {
                        tar_rgn_addr = sel_rgn.Address;
                        break;
                    }
                }
                else
                {
                    if (empty_count == 0)
                    {
                        tar_rgn_addr = sel_rgn.Address;
                        break;
                    }
                }
                c_offset++;
            }
            ExcelRange tar_rgn = tar_wrksht.Cells[tar_rgn_addr];
            int r_count = tar_rgn.End.Row;
            int col_count = tar_rgn.End.Column;
            var idx = tar_wrksht.GetMergeCellId(r_count, col_count);
            string mergedCellAddress = tar_rgn.Address;
            if (idx > 0)
            {
                mergedCellAddress = tar_wrksht.MergedCells[idx - 1];
            }
            int r = tar_wrksht.Cells[mergedCellAddress].Rows;
            int c = tar_wrksht.Cells[mergedCellAddress].Columns;
            if (left_to_right)
            {
                result = tar_rgn.Offset(0, c).Address;
            }
            else
            {
                result = tar_rgn.Offset(r, 0).Address;
            }
            return result;
        }
        public string get_offset_addr(string start_addr, ExcelWorksheet tar_wrksht, int offset_val, bool left_to_right)
        {
            string result = start_addr;
            ExcelRangeBase tar_rgn = tar_wrksht.Cells[start_addr];
            int inx = 0;
            while (inx < offset_val)
            {
                int r_count = tar_rgn.End.Row;
                int col_count = tar_rgn.End.Column;
                var idx = tar_wrksht.GetMergeCellId(r_count, col_count);
                string mergedCellAddress = tar_rgn.Address;
                if (idx > 0)
                {
                    mergedCellAddress = tar_wrksht.MergedCells[idx - 1];
                }
                int r = tar_wrksht.Cells[mergedCellAddress].Rows;
                int c = tar_wrksht.Cells[mergedCellAddress].Columns;
                if (left_to_right)
                {
                    tar_rgn = tar_rgn.Offset(0, c);
                }
                else
                {
                    tar_rgn = tar_rgn.Offset(r, 0);
                }
                inx++;
            }
            result = tar_rgn.Address;
            return result;
        }
        public ExcelPackage open_excel(string file_name)
        {
            ExcelPackage myexcel = null;
            FileInfo excel_file = new FileInfo(file_name);
            if (File.Exists(file_name))
            {
                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                myexcel = new ExcelPackage(excel_file);
            }
            return myexcel;
        }
        public int Get_Last_Cells_addr(string start_addr, ExcelWorksheet tar_wrksht, bool left_to_right)
        {
            int result = 0;
            ExcelRangeBase sel_rgn = tar_wrksht.Cells[start_addr];
            string sel_rgn_val = checkDBNull(sel_rgn);
            while (sel_rgn_val != "")
            {
                sel_rgn = tar_wrksht.Cells[get_offset_addr(sel_rgn.Address, tar_wrksht, 1, left_to_right)];
                sel_rgn_val = checkDBNull(sel_rgn.Value);
            }
            if (left_to_right)
            {
                sel_rgn = sel_rgn.Offset(0, -1);
                result = sel_rgn.End.Column;
            }
            else
            {
                sel_rgn = sel_rgn.Offset(-1, 0);
                result = sel_rgn.End.Row;
            }
            return result;
        }
        public ExcelRangeBase Get_Used_Range(string start_addr, ExcelWorksheet wrksht)
        {
            int from_addr_row = wrksht.Cells[start_addr].Start.Row;
            int from_addr_col = wrksht.Cells[start_addr].Start.Column;
            int to_addr_row = Get_Last_Cells_addr(start_addr, wrksht, false);
            int to_adrr_col = Get_Last_Cells_addr(start_addr, wrksht, true);
            ExcelRangeBase used_rgn = wrksht.Cells[from_addr_row, from_addr_col, to_addr_row, to_adrr_col];
            return used_rgn;
        }
        public void Draw_border(ExcelRangeBase used_rgn)
        {
            used_rgn.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            used_rgn.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            used_rgn.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            used_rgn.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            used_rgn.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            used_rgn.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }
        public void Fill_Range_Color(ExcelRangeBase used_rgn, Color sel_color)
        {
            used_rgn.Style.Fill.PatternType = ExcelFillStyle.Solid;
            used_rgn.Style.Fill.BackgroundColor.SetColor(sel_color);
        }
        public string remove_special_char(string src_str, List<char> remove_char)
        {
            return new string(src_str.Where(x => remove_char.IndexOf(x) == -1).ToArray());
        }
        public void Marking_Worksheet(ExcelWorkbook src_wrkbk, string sheet_name, Color sel_color)
        {
            List<char> remove_char = new List<char>() { '-', '_', '\r', '\n', '&' };
            foreach (var sht in src_wrkbk.Worksheets)
            {
                string sht_name = remove_special_char(sht.Name, remove_char).ToUpper();
                if (sht_name == remove_special_char(sheet_name, remove_char).ToUpper())
                {
                    sht.TabColor = sel_color;
                    break;
                }
            }
        }
        public void Marking_Worksheet(ExcelWorksheet src_wrksht, Color sel_color)
        {
            src_wrksht.TabColor = sel_color;
        }
        public MergeAreas_Info Get_Cells_Info(ExcelWorksheet src_wrksht, ExcelRangeBase src_rgn)
        {
            MergeAreas_Info result = new MergeAreas_Info();
            int r_count = src_rgn.End.Row;
            int col_count = src_rgn.End.Column;
            var idx = src_wrksht.GetMergeCellId(r_count, col_count);
            string mergedCellAddress = src_rgn.Address;
            if (idx > 0)
            {
                mergedCellAddress = src_wrksht.MergedCells[idx - 1];
            }
            result.col_qty = src_wrksht.Cells[mergedCellAddress].Columns;
            result.row_qty = src_wrksht.Cells[mergedCellAddress].Rows;
            return result;
        }
    }
}
