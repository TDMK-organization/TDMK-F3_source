
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Drawing;
using OK2SHIP_Lib;
using OK2SHIP_SMT;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SEEV_DLL;
using TDMK_SQL;
using VHX;
using myExcel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using static TDMK_EPPLUS_7.TDMK_EPPLUS7_lib;
using System.Security.Policy;
using System.Drawing.Imaging;
using Echeck_LogFile_Process;
using Microsoft.VisualBasic;
using OfficeOpenXml.Style;

namespace OK2SHIP
{
    public class ExportEPPlus
    {

        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public SEI_Lib myCode = new SEI_Lib();
        public EPPlus_Lib TDMK_Code2 = new EPPlus_Lib();


        public void Export_Chemical_resist_Image(SqlConnection sqlcon, ExcelWorksheet ws, string Itemcode, string Lotno)
        {
            int dem = 0;
            for (int i = 1; i < 30; i++)
            {
                for (int j = 1; j < 15; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).ToString().Replace(" ", "").ToUpper().Contains("Photo after test".Replace(" ", "").ToUpper()))
                    {
                        for (int k = 0; k < 15; k++)
                        {
                            if (myCode.IsNumeric(myCode.checkDBNull(ws.Cells[i + k, j - 4].Value)))
                            {
                                dem = dem + 1;
                            }
                        }
                        string[] list = { "Picture before", "Picture after" };
                        ExcelRangeBase curr_rgn = ws.Cells[i + 1, j - 1];
                        ExcelRangeBase rgn_judge = ws.Cells[i + 1, j - 2];
                        foreach (string zone in list)
                        {
                            string[] item = { "ItemCode", "LotNo", "Region" };
                            string[] item_val = { Itemcode, Lotno, zone };
                            System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_CHEMICAL_RESISTANCE_IMAGE", TDMK_Code.filter_str(item, item_val));

                            int number_image = new int[] { dem, data_image.Rows.Count }.Min();
                            for (int m = 0; m < number_image; m++)
                            {
                                byte[] img_byte = (byte[])(data_image.Rows[m]["Image_Data"]);
                                InsertPicture_Name(ws, curr_rgn, img_byte, curr_rgn.Address);
                                int col_offset = curr_rgn.Columns;
                                int row_off = curr_rgn.Rows;

                                curr_rgn = curr_rgn.Offset(row_off, 0);
                                rgn_judge.Offset(m, 0).Value = data_image.Rows[m]["Remark"];

                            }
                            curr_rgn = ws.Cells[i + 1, j];

                        }

                        return;

                    }
                }
            }


        }

        public void Export_Flux_Resist_Image(SqlConnection sqlcon, ExcelWorksheet ws, string Itemcode, string Lotno)
        {
            int start_with_FR4 = 24;
            int start_without_FR4 = 110;

            if (ws != null)
            {
                for (int i = 1; i < 200; i++)
                {
                    if (ws.Cells[i, 1].Value != null)
                    {
                        if (ws.Cells[i, 1].Value.ToString().Contains("test with FR4 cover"))
                        {
                            start_with_FR4 = i;
                        }
                        if (ws.Cells[i, 1].Value.ToString().Contains("test without FR4 cover"))
                        {
                            start_without_FR4 = i;
                        }
                    }
                }
                // Xử lý start_without_FR4 trước để đưa ra các vị trí cần input ảnh
                AutoCompleteStringCollection list_withoutFR4 = new AutoCompleteStringCollection();
                int start_top_side_withoutFR4 = 29;
                int start_bottom_side_withoutFR4 = 69;
                for (int j = Math.Min( start_with_FR4, start_without_FR4); j < Math.Max(start_with_FR4, start_without_FR4); j++)
                {

                    if (ws.Cells[j, 1].Value != null)
                    {
                        if (ws.Cells[j, 1].Value.ToString() == "Top side")
                        {
                            start_top_side_withoutFR4 = j;
                        }
                        if (ws.Cells[j, 1].Value.ToString() == "Bottom side")
                        {
                            start_bottom_side_withoutFR4 = j;
                        }
                    }
                }
                //
                for (int k = start_top_side_withoutFR4; k < start_bottom_side_withoutFR4; k++)
                {
                    if (ws.Cells[k, 3].Value != null)
                    {
                        if (ws.Cells[k, 3].Value.ToString() == "Applicable")
                        {
                            list_withoutFR4.Add("Without_FR4_Top side_" + ws.Cells[k, 2].Value.ToString() + "+" + (k + 2).ToString());
                        }
                    }
                }
                for (int k = start_bottom_side_withoutFR4; k < start_with_FR4; k++)
                {
                    if (ws.Cells[k, 3].Value != null)
                    {
                        if (ws.Cells[k, 3].Value.ToString() == "Applicable")
                        {
                            list_withoutFR4.Add("Without_FR4_Bottom side_" + ws.Cells[k, 2].Value.ToString() + "+" + (k + 2).ToString());
                        }
                    }
                }
                // Input ảnh vào withoutFR4
                foreach (string l in list_withoutFR4)
                {
                    string zone = l.Split('+')[0];
                    string row_input = l.Split('+')[1];
                    int number_image = 5;
                    string[] item = { "ItemCode", "LotNo", "Region" };
                    string[] item_val = { Itemcode, Lotno, zone };
                    System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_FLUX_RESIST_IMAGE", TDMK_Code.filter_str(item, item_val));
                    if (data_image.Rows.Count > 0)
                    {
                        number_image = Math.Min(data_image.Rows.Count, 5);

                        ExcelRangeBase curr_rgn = ws.Cells["H" + row_input];
                        ExcelRangeBase flux_data_rgn = curr_rgn.Offset(1, -4);
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] img_byte = (byte[])(data_image.Rows[m]["Image_Data"]);
                            InsertPicture_Name(ws, curr_rgn, img_byte, curr_rgn.Address);
                            int col_offset = curr_rgn.Columns;
                            int row_off = curr_rgn.Rows;

                            string flux = myCode.checkDBNull(data_image.Rows[m]["Flux_Data"]);
                            if (TDMK_Code.IsNumeric(flux))
                            {
                                double flux_val = Convert.ToDouble(flux) *1000;
                                flux_data_rgn.Offset(m, 0).Value = flux_val;
                                if (myCode.check_in_limit2("90", "50", flux_val.ToString()) == Color.White)
                                {
                                    flux_data_rgn.Offset(m, 2).Value = "Pass";
                                }
                                else
                                {
                                    flux_data_rgn.Offset(m, 2).Value = "Fail";
                                }
                            }
                            else
                            {
                                flux_data_rgn.Offset(m, 0).Value = flux;
                            }

                            curr_rgn = curr_rgn.Offset(0, col_offset);
                        }
                    }
                }

                //
                // Xử lý start_with_FR4 
                AutoCompleteStringCollection list_withFR4 = new AutoCompleteStringCollection();
                int start_top_side_withFR4 = 115;
                int start_bottom_side_withFR4 = 155;
                for (int j = start_with_FR4; j < 270; j++)
                {
                    if (ws.Cells[j, 1].Value != null)
                    {
                        if (ws.Cells[j, 1].Value.ToString() == "Top side")
                        {
                            start_top_side_withFR4 = j;
                        }
                        if (ws.Cells[j, 1].Value.ToString() == "Bottom side")
                        {
                            start_bottom_side_withFR4 = j;
                        }
                    }
                }
                //
                for (int k = start_top_side_withFR4; k < start_bottom_side_withFR4; k++)
                {
                    if (ws.Cells[k, 3].Value != null)
                    {
                        if (ws.Cells[k, 3].Value.ToString() == "Applicable")
                        {
                            list_withFR4.Add("With_FR4_Top side_" + ws.Cells[k, 2].Value.ToString() + "+" + (k + 2).ToString());
                        }
                    }
                }
                for (int k = start_bottom_side_withFR4; k < 270; k++)
                {
                    if (ws.Cells[k, 3].Value != null)
                    {
                        if (ws.Cells[k, 3].Value.ToString() == "Applicable")
                        {
                            list_withFR4.Add("With_FR4_Bottom side_" + ws.Cells[k, 2].Value.ToString() + "+" + (k + 2).ToString());
                        }
                    }
                }

                // Input ảnh vào withFR4
                foreach (string l in list_withFR4)
                {
                    string zone = l.Split('+')[0];
                    string row_input = l.Split('+')[1];
                    int number_image = 5;
                    string[] item = { "ItemCode", "LotNo", "Region" };
                    string[] item_val = { Itemcode, Lotno, zone };
                    System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_FLUX_RESIST_IMAGE", TDMK_Code.filter_str(item, item_val));
                    if (data_image.Rows.Count > 0)
                    {
                        number_image = Math.Min(data_image.Rows.Count, 5);

                        ExcelRangeBase curr_rgn = ws.Cells["H" + row_input];
                        ExcelRangeBase flux_data_rgn = curr_rgn.Offset(1, -4);
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])(data_image.Rows[m]["Image_Data"]);
                            InsertPicture_Name(ws, curr_rgn, data, curr_rgn.Address);
                            int col_offset = curr_rgn.Columns;
                            int row_off = curr_rgn.Rows;


                            string flux = myCode.checkDBNull(data_image.Rows[m]["Flux_Data"]);
                            if (TDMK_Code.IsNumeric(flux))
                            {
                                double flux_val = Convert.ToDouble(flux)*1000;
                                flux_data_rgn.Offset(m, 0).Value = flux_val;
                                if (myCode.check_in_limit2("90", "50", flux_val.ToString()) == Color.White)
                                {
                                    flux_data_rgn.Offset(m, 2).Value = "Pass";
                                }
                                else
                                {
                                    flux_data_rgn.Offset(m, 2).Value = "Fail";
                                }
                            }
                            else
                            {
                                flux_data_rgn.Offset(m, 0).Value = flux;
                            }
                            curr_rgn = curr_rgn.Offset(0, col_offset);
                        }
                    }
                }
            }


        }


        public int FindValueInColumn(ExcelWorksheet worksheet, int col_search, string searchValue, int startRow, int endRow)
        {
            for (int row = startRow; row <= endRow; row++)
            {
                string cellValue = worksheet.Cells[row, col_search].Value?.ToString();
                if (cellValue != null && cellValue.Trim().Equals(searchValue, StringComparison.OrdinalIgnoreCase))
                {
                    return row;
                }
            }

            return -1;
        }


        public ExcelRangeBase Find_cell(ExcelRangeBase start_cell, string key_word)
        {
            ExcelRangeBase cell = null;
            for (int i = 1; i < 15; i++)
            {
                if (myCode.checkDBNull(start_cell.Offset(0, i).Value.ToString().Replace(" ", "").ToUpper()).Contains(key_word.Replace(" ", "").ToUpper()))
                {
                    cell = start_cell.Offset(0, i);
                    break;
                }

            }

            return cell;

        }

        public void Export_ACF_ALL_CAMERA(SqlConnection sqlcon, ExcelWorksheet ws, string Itemcode, string Lotno)
        { 
            int count = 0;
            string[] lst_region = new string[] { "TRAI", "GIUA", "PHAI" };
            ExcelRangeBase curr_rgn = null;
            string[] item = { "ItemCode", "LotNo", "Region" };

            for (int i = 50; i < 100; i++)
            {
                for (int j = 1; j < 15; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).ToString().Replace(" ", "").ToUpper().Contains("Force-displacement profile".Replace(" ", "").ToUpper()))
                    {
                        string zone = lst_region[count];

                        string[] item_val = { Itemcode, Lotno, zone };
                        System.Data.DataTable data_image_force = TDMK_Code.Datatable_Filter(sqlcon, "ACF_GRAPH_FORCE_IMAGE", TDMK_Code.filter_str(item, item_val));
                        System.Data.DataTable data_image_tape_test = TDMK_Code.Datatable_Filter(sqlcon, "ACF_AFTER_TAPE_TEST_IMAGE", TDMK_Code.filter_str(item, item_val));
                        System.Data.DataTable data_image_glass_coupon = TDMK_Code.Datatable_Filter(sqlcon, "ACF_GLASS_COUPON_IMAGE", TDMK_Code.filter_str(item, item_val));

                        int number_image = new int[] { 5, data_image_force.Rows.Count, data_image_glass_coupon.Rows.Count, data_image_tape_test.Rows.Count }.Min();
                        // string glass_coupon_addr = Find_Cell_Addr("photos of glass coupon after test", ws.Cells[i, j].Address, ws, true);
                        // string after_tape_test_addr = Find_Cell_Addr("photos ACF pad of flex after test", ws.Cells[i, j].Address, ws, true);
                        ExcelRangeBase glass_coupon_rgn = Find_cell(ws.Cells[i, j], "photos of glass coupon after test").Offset(1, 0);
                        ExcelRangeBase after_tape_test_rgn = Find_cell(ws.Cells[i, j], "photos ACF pad of flex after test").Offset(1, 0);

                        curr_rgn = ws.Cells[i + 1, j];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image_force.Rows[m]["Image_Graph"];
                            byte[] data_tape_test = (byte[])data_image_tape_test.Rows[m]["Image_Data"];
                            byte[] data_glass_coupon = (byte[])data_image_glass_coupon.Rows[m]["Image_Data"];
                            string value = data_image_force.Rows[m]["Force_Data"].ToString();

                            // export graph
                            InsertPicture_Name(ws, curr_rgn, data, curr_rgn.Address);

                            // export glass coupon
                            if (glass_coupon_rgn != null)
                            {
                                InsertPicture_Name(ws, glass_coupon_rgn.Offset(m, 0), data_glass_coupon, glass_coupon_rgn.Offset(m, 0).Address);
                            }

                            // export tape test 
                            if (after_tape_test_rgn != null)
                            {
                                InsertPicture_Name(ws, after_tape_test_rgn.Offset(m, 0), data_tape_test, after_tape_test_rgn.Offset(m, 0).Address);
                            }

                            try
                            {
                                curr_rgn.Offset(0, 1).Value = Convert.ToDouble(value);
                            }
                            catch
                            {
                                curr_rgn.Offset(0, 1).Value = value;
                            }
                            curr_rgn = curr_rgn.Offset(1, 0);
                        }

                        count++;
                    }
                    if (count == 3)
                        goto lbl_export_next;

                }

            }
        lbl_export_next:
            // export ECF cleaning

            int row_input = FindValueInColumn(ws, 1, "Photo before cleaning", 30, 110);

            if (row_input > 1)
            {
                string[] item_val = { Itemcode, Lotno, "Photo before cleaning" };
                int input_image = 10;

                System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "ACF_CLEANING_IMAGE", TDMK_Code.filter_str(item, item_val));

                if (data_image.Rows.Count < 10)
                {
                    input_image = data_image.Rows.Count;
                }

                curr_rgn = ws.Cells["B" + row_input];
                for (int i = 0; i < input_image; i++)
                {
                    byte[] data = (byte[])data_image.Rows[i]["Image_Data"];

                    InsertPicture_Name(ws, curr_rgn, data, curr_rgn.Address);
                    //InsertPicture_Name(ws, curr_rgn.Offset(1, 0), data, curr_rgn.Offset(1, 0).Address);
                    curr_rgn = curr_rgn.Offset(0, 1);
                }


                //Export "Photo after ACF peel test"
                /*curr_rgn = curr_rgn.Offset(2, -input_image);
                System.Data.DataTable data_image_trai = TDMK_Code.Datatable_Filter(sqlcon, "ACF_AFTER_TAPE_TEST_IMAGE", TDMK_Code.filter_str(item, new string[] { Itemcode, Lotno, "TRAI" }));
                System.Data.DataTable data_image_giua = TDMK_Code.Datatable_Filter(sqlcon, "ACF_AFTER_TAPE_TEST_IMAGE", TDMK_Code.filter_str(item, new string[] { Itemcode, Lotno, "GIUA" }));
                System.Data.DataTable data_image_phai = TDMK_Code.Datatable_Filter(sqlcon, "ACF_AFTER_TAPE_TEST_IMAGE", TDMK_Code.filter_str(item, new string[] { Itemcode, Lotno, "PHAI" }));

                if (count == 3)
                {
                    for (int i = 0; i < 10; i++)
                    {
                        byte[] data = null;

                        if (i == 0 || i == 1 || i == 2)
                        {
                            if (i < data_image_trai.Rows.Count)
                            {
                                data = (byte[])data_image_trai.Rows[i]["Image_Data"];
                            }
                        }
                        else if (i == 3 || i == 4 || i == 5 || i == 6)
                        {
                            if (i - 3 < data_image_giua.Rows.Count)
                            {
                                data = (byte[])data_image_giua.Rows[i - 3]["Image_Data"];
                            }

                        }
                        else if (i == 7 || i == 8 || i == 9)
                        {
                            if (i - 7 < data_image_phai.Rows.Count)
                            {
                                data = (byte[])data_image_phai.Rows[i - 7]["Image_Data"];
                            }
                        }

                        if (data != null)
                        {
                            InsertPicture_Name(ws, curr_rgn, data, curr_rgn.Address);
                            curr_rgn = curr_rgn.Offset(0, 1);
                        }
                    }
                }
                else if (count == 1)
                {
                    for (int i = 0; i < 10; i++)
                    {
                        if (i < data_image_trai.Rows.Count)
                        {
                            byte[] data = (byte[])data_image_trai.Rows[i]["Image_Data"];
                            if (data != null)
                            {

                                InsertPicture_Name(ws, curr_rgn, data, curr_rgn.Address);
                                curr_rgn = curr_rgn.Offset(0, 1);

                            }
                        }

                    }
                }*/
            }
        }

        public void InsertPicture_Name(ExcelWorksheet wsSheet1, ExcelRangeBase tar_rgn, byte[] img_data, string pic_name)
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
            double col_w = 0;
            double row_h = 0;
            for (int i = 0; i < c; i++)
            {
                col_w += wsSheet1.Column(_colIndex + i).Width;
            }
            for (int i = 0; i < r; i++)
            {
                row_h += wsSheet1.Row(_rowIndex + i).Height;
            }

            var _img = BitmapToBytes(img_data, ImageFormat.Jpeg);
            var img = Export_Image_Class.imgToByteConverter(_img);

            //using (MemoryStream ms = new MemoryStream(img_data))
            //{
            using (MemoryStream ms = new MemoryStream(img))
            {
                ExcelPicture pic = wsSheet1.Drawings.AddPicture(pic_name, ms);//img 
                ExcelWorkbook wrkbk = wsSheet1.Workbook;
                decimal mdw = wrkbk.MaxFontWidth;
                int pixelHeight = (int)(row_h / 0.75);
                int pixelWidth = (int)decimal.Truncate(((256 * (decimal)col_w + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
                int offset = (int)(0.05 * Math.Min(pixelHeight, pixelWidth));
                pic.SetPosition(_rowIndex - 1, offset, _colIndex - 1, offset);
                pic.SetSize(pixelWidth - 2 * offset, pixelHeight - 2 * offset);

            }
        }

        public void Export_CQRA_bHast_Image(SqlConnection sqlcon, ExcelWorksheet ws, string Itemcode, string Lotno)
        {
            int row_input = 0;
            if (ws != null)
            {
                string zone = "Picture";
                int dem = 0;
                for (int i = 1; i < 30; i++)
                {
                    if (ws.Cells[i, 1].Value != null)
                    {
                        if (ws.Cells[i, 1].Value.ToString().Contains("Picture"))
                        {
                            row_input = i;
                        }
                    }
                }
                if (row_input > 0)
                {
                    for (int j = 0; j < 15; j++)
                    {
                        if (ws.Cells[row_input - 1, 2 + j].Value != null)
                        {
                            dem = dem + 1;
                        }

                    }

                    int number_image = 5;
                    string[] item = { "ItemCode", "LotNo", "Region" };
                    string[] item_val = { Itemcode, Lotno, zone };
                    System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_BHAST_IMAGE", TDMK_Code.filter_str(item, item_val));
                    if (data_image.Rows.Count > 0)
                    {
                        if (data_image.Rows.Count < number_image)
                        {
                            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                            string file_dic = Path.Combine(file_folder, "1.jpg");

                            ExcelRangeBase curr_rgn = ws.Cells["B" + row_input + 1];
                            for (int m = 0; m < data_image.Rows.Count; m++)
                            {
                                byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                InsertPicture_Name(ws, curr_rgn, data, curr_rgn.Address);
                                curr_rgn = curr_rgn.Offset(0, 1);
                            }
                        }
                        else
                        {
                            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                            string file_dic = Path.Combine(file_folder, "1.jpg");

                            ExcelRangeBase curr_rgn = ws.Cells["B" + row_input + 1];
                            for (int m = 0; m < number_image; m++)
                            {
                                byte[] data = (byte[])data_image.Rows[m]["Image_Data"];

                                InsertPicture_Name(ws, curr_rgn, data, curr_rgn.Address);
                                curr_rgn = curr_rgn.Offset(0, 1);
                            }
                        }
                    }
                }

                else
                {
                    export_formatmoi_13_12("CQRA_BHAST_IMAGE", ws, Itemcode, Lotno, zone, sqlcon);
                }
            }


        }
        public void Export_Bhast_Logfile(ExcelWorkbook export_wrkbook, string itemcode, string lotno, SqlConnection sqlcon)
        {
            string[] item = { "ItemCode", "LotNo" };
            string[] item_val = { itemcode, lotno };
            string[] list_ch = { "Ch01", "Ch02", "Ch03", "Ch04", "Ch05", "Ch06", "Ch07", "Ch08" };
            int start_col = 1;
            int start_row = 1;

            ExcelWorksheet ws = null;
            ExcelWorkbook curr_wrkbook = export_wrkbook;
            if (curr_wrkbook != null)
            {
                string _process = "CQRA - bHast".Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
                foreach (ExcelWorksheet tg_sht in curr_wrkbook.Worksheets)
                {
                    string cur_sht_name = tg_sht.Name.ToUpper().Replace("-", "").Replace(" ", "");
                    if (cur_sht_name == _process)
                    {
                        ws = tg_sht;
                        break;
                    }
                }

                if (ws != null)
                {
                    for (int i = 15; i < 35; i++)
                    {
                        for (int j = 4; j < 20; j++)
                        {
                            if (ws.Cells[i, j].Value != null)
                            {
                                if (ws.Cells[i, j].Value.ToString().ToUpper() == "CH01")
                                {
                                    start_row = i;
                                    start_col = j; break;
                                }
                            }
                        }
                    }
                    foreach (string chanel in list_ch)
                    {
                        try
                        {
                            AutoCompleteStringCollection list = TDMK_Code.Load_Item_Filter_str(sqlcon, "CQRA_BHAST_LOGFILE", chanel, TDMK_Code.filter_str(item, item_val));

                            string number = chanel.Replace("Ch0", "");
                            int col_offset = Int32.Parse(number);

                            for (int y = 1; y <= list.Count; y++)
                            {
                                if (TDMK_Code.IsNumeric(list[y - 1]))
                                {
                                    ws.Cells[start_row + y, start_col + col_offset - 1].Value = Convert.ToDouble(list[y - 1]);
                                    ws.Cells[start_row + y, start_col + col_offset - 1].Style.Numberformat.Format = "0.00E+00";
                                }
                                else
                                {
                                    ws.Cells[start_row + y, start_col + col_offset - 1].Value = list[y - 1];
                                }
                            }
                        }
                        catch
                        {

                        }

                    }

                }

            }
        }
        public void export_formatmoi_13_12(string table_name, ExcelWorksheet ws, string Itemcode, string Lotno, string zone, SqlConnection sqlcon)
        {
            int dem = 0;
            for (int i = 1; i < 30; i++)
            {
                for (int j = 1; j < 15; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).ToString().ToUpper().Contains("SCREENSHOT"))
                    {
                        for (int k = 0; k < 15; k++)
                        {
                            if (myCode.IsNumeric(myCode.checkDBNull(ws.Cells[i + k, j - 4].Value)))
                            {
                                dem = dem + 1;
                            }
                        }

                        string[] item = { "ItemCode", "LotNo", "Region" };
                        string[] item_val = { Itemcode, Lotno, zone };
                        System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, table_name, TDMK_Code.filter_str(item, item_val));
                        int number_image = new int[] { dem, data_image.Rows.Count }.Min();

                        ExcelRangeBase curr_rgn = ws.Cells[i + 1, j];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                            InsertPicture_Name(ws, curr_rgn, data, curr_rgn.Address);
                            curr_rgn = curr_rgn.Offset(1, 0);
                        }

                        return;
                    }
                }
            }
        }

        public void Export_High_Speed_Ball_Shear_image(SqlConnection sqlcon, ExcelWorksheet ws, string Itemcode, string Lotno)
        {
            string[] item = { "ItemCode", "LotNo" };
            string[] item_val = { Itemcode, Lotno };
            int number_image = 10;
            System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_HIGH_SPEED_BALL_SHEAR_IMAGE", TDMK_Code.filter_str(item, item_val));
            if (data_image.Rows.Count > 0)
            {

                int qty = Math.Min(data_image.Rows.Count, number_image);
                int row_val;
                for (int x = 1; x < 40; x++)
                {
                    if (ws.Cells[x, 1].Value != null && ws.Cells[x, 1].Value.ToString() == "Photo")
                    {
                        row_val = (int)x;
                        ExcelRangeBase curr_rgn = ws.Cells["B" + row_val];
                        for (int m = 0; m < qty; m++)
                        {
                            int col_off = (int)(m / 2) + 5 * (m % 2);
                            byte[] photo_data = (byte[])data_image.Rows[m]["Image_Photo"];
                            byte[] graph_data = (byte[])data_image.Rows[m]["Image_Graph"];
                            List<byte[]> image_data = new List<byte[]> { photo_data, graph_data };
                            for (int i = 0; i < image_data.Count; i++)
                            {
                                InsertPicture_Name(ws, curr_rgn.Offset(i,col_off), image_data[i], curr_rgn.Offset(i, col_off).Address);
                            }
                            //curr_rgn.Offset[image_data.Count, col_off].Value = data_image.Rows[m]["Force_Data"];

                            if (TDMK_Code.IsNumeric(myCode.checkDBNull(data_image.Rows[m]["Force_Data"])))
                            {
                                double force_val = Convert.ToDouble(data_image.Rows[m]["Force_Data"]);
                                curr_rgn.Offset(image_data.Count, col_off).Value = force_val;
                                if (force_val >= 250)
                                {
                                    curr_rgn.Offset(image_data.Count + 2, col_off).Value = "Pass";
                                }
                                else
                                {
                                    curr_rgn.Offset(image_data.Count + 2, col_off).Value = "Fail";
                                }
                            }

                        }
                        break;
                    }
                }
            }
        }

        public void Export_CQRA_Solderability_Image(SqlConnection sqlcon, ExcelWorksheet ws, string Itemcode, string Lotno)
        {
            int dem = 0;
            for (int i = 1; i < 30; i++)
            {
                for (int j = 1; j < 15; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).ToString().Replace(" ", "").ToUpper().Contains("Unit S/N".Replace(" ", "").ToUpper()))
                    {
                        for (int k = 0; k < 15; k++)
                        {
                            if (myCode.IsNumeric(myCode.checkDBNull(ws.Cells[i + k, j].Value)))
                            {
                                dem = dem + 1;
                            }
                        }

                        ExcelRangeBase curr_rgn = ws.Cells[i + 1, j + 2];

                        string[] item = { "ItemCode", "LotNo" };
                        string[] item_val = { Itemcode, Lotno };
                        System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_SOLDERABILITY_IMAGE", TDMK_Code.filter_str(item, item_val));

                        int number_image = new int[] { dem, data_image.Rows.Count }.Min();
                        ExcelRangeBase rgn_judge = ws.Cells[i + 1, j + 3];

                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                            InsertPicture_Name(ws, curr_rgn, data, curr_rgn.Address);
                            rgn_judge.Offset(m, 0).Value = data_image.Rows[m]["Remark"];
                            curr_rgn = curr_rgn.Offset(1, 0);
                        }
                        return;

                    }
                }
            }


        }

        public void Export_CQRA_Dielectric_withstanding(SqlConnection sqlcon, ExcelWorksheet ws, string Itemcode, string Lotno)
        {
            int dem = 0;

            for (int i = 1; i < 30; i++)
            {
                for (int j = 1; j < 15; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).ToString().Replace(" ", "").ToUpper().Contains("Photo after test".Replace(" ", "").ToUpper()))
                    {
                        for (int k = 0; k < 15; k++)
                        {
                            if (myCode.IsNumeric(myCode.checkDBNull(ws.Cells[i + k, j - 3].Value)))
                            {
                                dem = dem + 1;
                            }
                        }

                        string[] item = { "ItemCode", "LotNo" };
                        string[] item_val = { Itemcode, Lotno };
                        System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_DIELECTRIC_WITHSTANDING_IMAGE", TDMK_Code.filter_str(item, item_val));

                        int number_image = new int[] { dem, data_image.Rows.Count }.Min();
                        ExcelRangeBase curr_rgn = ws.Cells[i + 1, j];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data1 = (byte[])data_image.Rows[m]["Image_After"];
                            InsertPicture_Name(ws, curr_rgn, data1, curr_rgn.Address);

                            byte[] data2 = (byte[])data_image.Rows[m]["Image_Before"];
                            InsertPicture_Name(ws, curr_rgn.Offset(0, -1), data2, curr_rgn.Offset(0, -1).Address);

                            curr_rgn = curr_rgn.Offset(1, 0);
                        }

                        return;

                    }
                }
            }
        }
        public void Export_OQC_Test_Image(SqlConnection sqlcon, ExcelWorksheet ws, string Itemcode, string Lotno)
        {
            string[] arr = { "Solder Mask", "PI & SUS", "Gold", "EMI" };

            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 1; i < 50; i++)
            {
                if (ws.Cells[i, 1].Value != null)
                {
                    foreach (string x in arr)
                    {
                        if (ws.Cells[i, 1].Value.ToString().Contains(x))
                        {
                            list.Add((ws.Cells[i, 1].Value.ToString().Split('/')[0]).TrimEnd() + "+" + (i + 2).ToString());
                        }
                    }

                }
            }

            foreach (string xx in list)
            {
                string zone = xx.Split('+')[0];
                string row_input = xx.Split('+')[1];
                int number_image = 5;
                string[] item = { "ItemCode", "LotNo", "Region" };
                string[] item_val = { Itemcode, Lotno, zone };
                System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "OQC_TEST_IMAGE", TDMK_Code.filter_str(item, item_val));
                if (data_image.Rows.Count > 0)
                {
                    if (data_image.Rows.Count < number_image)
                    {
                        number_image = data_image.Rows.Count;
                        ExcelRangeBase curr_rgn = ws.Cells["C" + row_input];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                            InsertPicture_Name(ws, curr_rgn, data, curr_rgn.Address);
                            curr_rgn = curr_rgn.Offset(0, 1);
                        }
                    }
                    else
                    {
                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        string file_dic = Path.Combine(file_folder, "1.jpg");

                        ExcelRangeBase curr_rgn = ws.Cells["C" + row_input];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                            InsertPicture_Name(ws, curr_rgn, data, curr_rgn.Address);
                            curr_rgn = curr_rgn.Offset(0, 1);
                        }
                    }
                }
            }

        }
        public void Export_ACFFlatness(SqlConnection sqlcon, ExcelWorksheet ws, string ItemCode, string LotNo)
        {
            System.Data.DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, "ACF_FLATNESS", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
            if (src_dt.Rows.Count > 0)
            {
                ExcelRangeBase curr_rgn = ws.Cells["A1"];
                for (int j = 0; j < 5; j++)
                {
                    curr_rgn = curr_rgn.Offset(0, j);
                    for (int i = 0; i < 100; i++)
                    { 
                        if (myCode.checkDBNull(curr_rgn.Offset(i, 0).Value).Contains("Point"))
                        {
                            curr_rgn = curr_rgn.Offset(i, 0);
                            goto lbl_export;
                        }

                    }

                }

            lbl_export:
                for (int i = 0; i < src_dt.Rows.Count; i++)
                {
                    for (int j = 0; j < 10; j++)
                    {
                        string val = src_dt.Rows[i][4 + j].ToString();
                        if (myCode.IsNumeric(val))
                        {
                            curr_rgn.Offset(i + 1, j).Value = double.Parse(val) * 1000;
                        }

                    }
                }

            }


        }

        public string ACF_pad_location_newformat(ExcelWorksheet ws)
        {
            // List<int> lst_info = new List<int> { };
            string get_info = "";
            ExcelRangeBase curr_rgn_1 = ws.Cells["A1"];

            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 30; i < 200; i++)
            {
                if (curr_rgn_1.Offset(i, 0).Value != null)
                {
                    if (curr_rgn_1.Offset(i, 0).Value.ToString().Replace(" ", "").ToUpper() == "ACF pads Surface Roughness".Replace(" ", "").ToUpper())
                    {
                        int k = 1;
                        int count_sample = 0;
                        get_info = i.ToString();
                        for (int m = 1; m < 20; m++)
                        {
                            if (myCode.checkDBNull(curr_rgn_1.Offset(i + m + count_sample, 0).Value).Replace(" ", "").ToUpper().Contains("Sample".Replace(" ", "").ToUpper()))
                            {
                                get_info = (i + m).ToString();

                                while (myCode.checkDBNull(curr_rgn_1.Offset(i + m + count_sample, 0).Value).Replace(" ", "").ToUpper().Contains("Sample".Replace(" ", "").ToUpper()))
                                {
                                    count_sample++;
                                }
                                break;
                            }
                        }
                        while (myCode.checkDBNull(curr_rgn_1.Offset(i + 2, k).Value).Replace(" ", string.Empty).ToUpper().Contains("surface".Replace(" ", string.Empty).ToUpper()))
                        {
                            if (myCode.checkDBNull(curr_rgn_1.Offset(i + 2, k).Value).Contains("Sa"))
                            {

                                get_info += "+ sa";
                            }
                            if (myCode.checkDBNull(curr_rgn_1.Offset(i + 2, k).Value).Contains("Sq"))
                            {

                                get_info += "+ sq";
                            }
                            if (myCode.checkDBNull(curr_rgn_1.Offset(i + 2, k).Value).Contains("Sdr"))
                            {
                                get_info += "+ sdr";
                            }
                            k++;
                        }
                        get_info += "+" + count_sample.ToString();
                        break;
                    }
                }
            }
            return get_info;
        }


        public string str_replace(string str_, List<string> lst_repl)
        {
            foreach (string i in lst_repl)
            {
                str_ = str_.Replace(i, "");
            }
            return str_;
        }

        public string get_reportpath(string ItemCode, string LotNo, string id, string sheet, string fpath)
        {
            string report_path = "";
            DirectoryInfo tar_parent = new DirectoryInfo(fpath);
            DirectoryInfo[] arr_dic_child = tar_parent.GetDirectories();
            foreach (DirectoryInfo folder_sheet in arr_dic_child)
            {
                string f_name = folder_sheet.Name;
                List<string> lst_repl = new List<string> { "_", "-", " ", "&" };
                if (str_replace(f_name, lst_repl).ToUpper().Contains(str_replace(sheet, lst_repl).ToUpper()))
                {
                    report_path = Path.Combine(fpath, f_name, ItemCode + "-" + LotNo + "_" + id + ".xlsx");
                    break;
                }
            } 
            return report_path;

        } 

        public List<List<string>> lst_roughness(System.Data.DataTable dt)
        {
            List<List<string>> lst_roughness = new List<List<string>>() { };
            string[] arr_name = { "L1_Roughness_Sa", "L1_Roughness_Sq", "L1_Roughness_Sdr", "L2_Roughness_Sa", "L2_Roughness_Sq", "L2_Roughness_Sdr", "L3_Roughness_Sa", "L3_Roughness_Sq", "L3_Roughness_Sdr" };
            foreach (var item in arr_name)
            {
                lst_roughness.Add(dt.AsEnumerable().Select(x => x.Field<string>(item)).ToList());
            } 
            return lst_roughness;
        }
        public void Export_ACF_Roughness_newformat(string itemcode, string lotno, ExcelWorksheet ws, SqlConnection sqlcon_IPQC)
        {
            if (sqlcon_IPQC != null)
            {
                System.Data.DataTable dt1 = TDMK_Code.Datatable_Filter(sqlcon_IPQC, "Roughness", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcode, lotno }));
                if (dt1.Rows.Count > 0)
                {
                    List<List<string>> lst_rghness = lst_roughness(dt1); 
                    ExcelRangeBase curr_rgn_1 = ws.Cells["A1"];

                    List<int> vitri_sa = new List<int> { };
                    List<int> vitri_sq = new List<int> { };
                    List<int> vitri_sdr = new List<int> { };
                    string[] arr_info = ACF_pad_location_newformat(ws).Split('+');
                    string vitri_ACF_pad = arr_info[0];

                    if (myCode.IsNumeric(arr_info[0]))
                    {
                        int row = int.Parse(arr_info[0]);
                        curr_rgn_1 = curr_rgn_1.Offset(row, 0);
                        int r = 1;
                        for (int i = 0; i < 9; i++)
                        {
                            if (r < arr_info.Length - 1)
                            {
                                if (arr_info[r].Contains("sa"))
                                {
                                    vitri_sa.Add(i);
                                }
                                if (arr_info[r].Contains("sq"))
                                {
                                    vitri_sq.Add(i);
                                }
                                if (arr_info[r].Contains("sdr"))
                                {
                                    vitri_sdr.Add(i);
                                }
                            } 
                            r++;
                        }

                        int m = 0;
                        int k;

                        int count_pt = dt1.Rows.Count;
                        while (m < count_pt)
                        {
                            int c = 0;
                            foreach (int j in vitri_sa)
                            {
                                try
                                {
                                    curr_rgn_1.Offset(m, j + 1).Value = Convert.ToDouble(lst_rghness[c * 3][m]);
                                }
                                catch
                                {
                                    curr_rgn_1.Offset(m, j + 1).Value = lst_rghness[c * 3][m];
                                }
                                
                                c++;

                            }
                            c = 0;
                            foreach (int j in vitri_sq)
                            {
                                try
                                {
                                    curr_rgn_1.Offset(m, j + 1).Value = Convert.ToDouble(lst_rghness[1 + c * 3][m]);
                                }
                                catch
                                {
                                    curr_rgn_1.Offset(m, j + 1).Value = lst_rghness[1 + c * 3][m];
                                }
                                
                                c++;
                            }
                            c = 0;
                            foreach (int j in vitri_sdr)
                            {
                                try
                                {
                                    curr_rgn_1.Offset(m, j + 1).Value = Convert.ToDouble( lst_rghness[2 + c * 3][m]);
                                }
                                catch
                                {
                                    curr_rgn_1.Offset(m, j + 1).Value = lst_rghness[2 + c * 3][m];
                                }
                                
                                c++;
                            }
                            m++;


                        }
                    }

                }
            }
        }

        public Image BitmapToBytes(byte[] image, ImageFormat pFormat)
        {
            var imageObject = new Bitmap(new MemoryStream(image));

            var stream = new MemoryStream();
            imageObject.Save(stream, pFormat);

            return new Bitmap(stream);
        }

        public void Export_EPPlus(SqlConnection sqlcon, string file_format, string export_path, string itemcode, string lotno, string sheet, SqlConnection sqlcon_IPQC)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelPackage xlPackage = open_excel(file_format);
            ExcelWorkbook wb_format = xlPackage.Workbook;

            ExcelWorkbook report_saved = null;
            ExcelPackage sourcePackage = null;

        lbl_export:
            if (System.IO.File.Exists(export_path))
            {
                if (MessageBox.Show(new Form { TopMost = true }, "Báo cáo đã tồn tại. Bạn có muốn thay thế không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    try
                    {
                        System.IO.File.Delete(export_path);
                        goto lbl_export;
                    }
                    catch
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Báo cáo đang được mở", "Thông báo");
                    }
                }
            }
            else
            {

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                sourcePackage = new ExcelPackage();
                report_saved = sourcePackage.Workbook;
                List<string> lst_repl = new List<string> { "_", "-", " ", "&", "(", ")" };
                string _process = str_replace(sheet, lst_repl).ToUpper();
                string mySheet = "";
                foreach (ExcelWorksheet tg_sht in wb_format.Worksheets)
                {
                    // string cur_sht_name = tg_sht.Name.ToUpper().Replace("_", "").Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "");
                    string cur_sht_name = str_replace(tg_sht.Name.ToUpper(), lst_repl).ToUpper();
                    if (cur_sht_name.Contains(_process)||_process.Contains(cur_sht_name))
                    {
                        mySheet = tg_sht.Name;
                        break;
                    }
                }

                if (mySheet != "")
                {
                    ExcelWorksheet worksheet0 = wb_format.Worksheets[mySheet];
                    report_saved.Worksheets.Add(mySheet, worksheet0);
                    
                    ExcelWorksheet ws = report_saved.Worksheets[0];
                  //  ws.ConditionalFormatting.RemoveAll();

                    switch (sheet)
                    {
                        case "CQRA_BHAST":
                            Export_CQRA_bHast_Image(sqlcon, ws, itemcode, lotno);
                            break;
                        case "CQRA_SOLDERABILITY":
                            Export_CQRA_Solderability_Image(sqlcon, ws, itemcode, lotno);
                            break;
                        case "CQRA_DIELECTRIC_WITHSTANDING":
                            Export_CQRA_Dielectric_withstanding(sqlcon, ws, itemcode, lotno);
                            break;
                        case "CQRA_HIGH_SPEED_BALL_SHEAR":
                            Export_High_Speed_Ball_Shear_image(sqlcon, ws, itemcode, lotno);
                            break;
                        case "OQC_TEST":
                            Export_OQC_Test_Image(sqlcon, ws, itemcode, lotno);
                            break;
                        case "CQRA_CHEMICAL_RESISTANCE":
                            Export_Chemical_resist_Image(sqlcon, ws, itemcode, lotno);
                            break;
                        case "CQRA_IR_VIA_TO_VIA":
                            export_formatmoi_13_12("CQRA_IR_VIA_TO_VIA_IMAGE", ws, itemcode, lotno, "Picture", sqlcon);
                            break;
                        case "CQRA_IR_TRACE_TO_TRACE":
                            export_formatmoi_13_12("CQRA_IR_TRACE_TO_TRACE_IMAGE", ws, itemcode, lotno, "Picture", sqlcon);
                            break;
                        case "CQRA_IR_LAYER_TO_LAYER":
                            export_formatmoi_13_12("CQRA_IR_LAYER_TO_LAYER_IMAGE", ws, itemcode, lotno, "Picture", sqlcon);
                            break;
                        case "CQRA_FLUX_RESIST":
                            Export_Flux_Resist_Image(sqlcon, ws, itemcode, lotno);
                            break; 
                        case "OQC ACF":
                            Export_ACFFlatness(sqlcon, ws, itemcode, lotno);
                            Export_ACF_ALL_CAMERA(sqlcon, ws, itemcode, lotno);
                            Export_ACF_Roughness_newformat(itemcode, lotno, ws, sqlcon_IPQC);
                            break; 

                        case "STACKUP":
                            Export_stackup(sqlcon, ws, itemcode, lotno);
                            break; 

                        case "BVH-PTH":
                            Export_BVH_PTH(sqlcon, ws, itemcode, lotno);
                            break; 

                        case "IMPEDANCE":
                            Export_Impedance(sqlcon, ws, itemcode, lotno);
                            break;

                        case "SOLDERMASK":
                            Export_SolderMask(sqlcon, ws, itemcode, lotno);
                            break;

                        case "CQRA - THERMAL STRESS":
                            Export_Thermal_Stress_Image(sqlcon, ws, itemcode, lotno);
                            break; 
                    } 

                    foreach (var cell in ws.Cells)
                    {
                        //if (cell.Value == null && cell.Style.Fill.BackgroundColor.Equals(Color.Red))
                        //{
                            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            cell.Style.Fill.BackgroundColor.SetColor(Color.Transparent);
                       // } 
                    } 
                    if (sourcePackage != null)
                    {
                        sourcePackage.SaveAs(new FileInfo(export_path));
                    } 
                    try
                    {  
                        //var excelApp = new myExcel.Application();
                        //excelApp.Visible = true;
                        //excelApp.Workbooks.Open(export_path);
                        ProcessStartInfo pi = new ProcessStartInfo(export_path);
                        Process.Start(pi);
                    }
                    catch
                    {
                        MessageBox.Show(new Form { TopMost = true }, "File opening error !", "Warning");
                    }
                    MessageBox.Show(new Form { TopMost = true }, "Export to checksheet complete!", "Warning");
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Sheet" + _process + " not found !", "Warning");
                }

            }
        }


        public ExcelPackage open_excel(string file_name)
        {
            ExcelPackage myexcel = null;
            FileInfo excel_file = new FileInfo(file_name);
            if (File.Exists(file_name))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                myexcel = new ExcelPackage(excel_file);
            }
            return myexcel;
        }

        /***********************  Export Stackup Function **********************************************/
        public void Export_stackup(SqlConnection sqlcon, ExcelWorksheet ws, string ItemCode, string LotNo)
        {
            AutoCompleteStringCollection list_val = Get_zone_infor_to_input_value(ws);
            foreach (string list in list_val)
            {
                string zone = list.Split('+')[0]; // A
                int row = int.Parse(list.Split('+')[1]); // 77
                int number_val = int.Parse(list.Split('+')[2]); //8
                for (int i = 1; i < 6; i++)
                {
                    string[] item = { "ItemCode", "LotNo", "Region", "Data_For", "Pcs_No" };
                    string[] item_val = { ItemCode, LotNo, "Zone_" + zone, "STACKUP", i.ToString() };
                    System.Data.DataTable data_table = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_LOGFILE", TDMK_Code.filter_str(item, item_val));
                    DataGridView DGV_tem = new DataGridView();
                    DGV_tem.DataSource = data_table.Columns[6];
                    if (data_table.Rows.Count > 0)
                    { 
                        int min_pcs = (int)Math.Min(data_table.Rows.Count, number_val); 
                        for (int j = 0; j < min_pcs; j++)
                        {
                            string cur_val = myCode.checkDBNull(data_table.Rows[j]["Data"]);
                            if (TDMK_Code.IsNumeric(cur_val))
                            {
                                ws.Cells[row + j, i + 6].Value = Convert.ToDecimal(cur_val);
                                ws.Cells[row + j, i + 6].Style.Numberformat.Format = "#0.000";
                            }
                        }

                    }
                }

            }
            Export_StackUp_Image(sqlcon, ws, ItemCode, LotNo);


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
        public AutoCompleteStringCollection Get_zone_infor_to_input_value(ExcelWorksheet ws)
        {
            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 24; i < 400; i++)
            {
                if (ws.Cells[i, 1].Value != null)
                {
                    if (ws.Cells[i, 1].Value.ToString() == "Zone")
                    {
                        int k = 1;
                        int dem = 0;
                        while (ws.Cells[i + k, 5].Value != null)
                        {
                            dem = dem + 1;
                            k++;
                        }
                        list.Add(ws.Cells[i + 1, 1].Value.ToString() + "+" + (i + 1).ToString() + "+" + (dem - 1).ToString());
                    }
                }
            }
            return list;
        }
        public AutoCompleteStringCollection Get_zone_infor_to_input_image(ExcelWorksheet ws)
        {
            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 20; i < 400; i++)
            {
                if (ws.Cells[i, 2].Value != null)
                {
                    if ((ws.Cells[i, 2].Value.ToString().Contains("Region/Zone")))
                    {
                        if (ws.Cells[i, 4].Value != null)
                        {
                            list.Add(ws.Cells[i, 4].Value.ToString() + "+" + (i + 1).ToString());
                        }
                        if (ws.Cells[i, 3].Value != null)
                        {
                            list.Add(ws.Cells[i, 3].Value.ToString() + "+" + (i + 1).ToString());
                        }

                    }
                    if (ws.Cells[i, 2].Value.ToString().Contains("PTH_X-section"))
                    {
                        list.Add(ws.Cells[i, 2].Value.ToString() + "+" + (i).ToString());
                    }
                    if (ws.Cells[i, 2].Value.ToString().Contains("BVH_X-section"))
                    {
                        list.Add(ws.Cells[i, 2].Value.ToString() + "+" + (i).ToString());
                    }
                }
                if (ws.Cells[i, 3].Value != null)
                {
                    if (ws.Cells[i, 3].Value.ToString() == "Region/Zone" && ws.Cells[i, 4].Value != "")
                    {
                        list.Add(ws.Cells[i, 4].Value.ToString() + "_" + (i + 1).ToString());
                    }
                }
            }
            return list;
        }
        public void Export_Image_new(ExcelWorksheet ws, int number_image, System.Data.DataTable data_image, string row_val)
        {
            int count_sample = new int[] { number_image, data_image.Rows.Count }.Min();
            ExcelRangeBase curr_rgn = ws.Cells["F" + row_val];
            for (int m = 0; m < count_sample; m++)
            {
                byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                int col = 1;
                InsertPicture_Name2(ws, curr_rgn, data, curr_rgn.Address, ref col);
                curr_rgn = curr_rgn.Offset(0, col);
            }
        }
        public void Export_StackUp_Image(SqlConnection sqlcon, ExcelWorksheet ws, string Itemcode, string Lotno)
        {
            AutoCompleteStringCollection list = Get_zone_infor_to_input_image(ws);
            System.Data.DataTable dt_tbl = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { Itemcode, Lotno }));
            List<string> lst_region = dt_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).ToList();
            bool t = true;
            if (lst_region.Contains("Zone_BVH1") && lst_region.Contains("Zone_BVH2"))
            {
                t = false;
            }
            int no_BVH = 1;
            foreach (string x in list)
            {
                string region = x.Split('+')[0];
                string row_val = x.Split('+')[1];
                int number_image = 5;

                if (region.Contains("BVH") || region.Contains("PTH"))
                {
                    if (t)
                    {
                        region = "Zone_BVH";
                    }
                    else
                    {
                        region = "Zone_BVH" + no_BVH.ToString();
                        no_BVH++;
                    }
                }
                else if (region.Contains("PTH"))
                {
                    region = "Zone_PTH";
                }
                else
                {
                    region = "Zone_" + region;
                }

                string[] item = { "ItemCode", "LotNo", "Region" };
                string[] item_val = { Itemcode, Lotno, region };
                System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_IMAGE", TDMK_Code.filter_str(item, item_val));
                Export_Image_new(ws, number_image, data_image, row_val);

            }
        }

        public void InsertPicture_Name2(ExcelWorksheet wsSheet1, ExcelRangeBase tar_rgn, byte[] img_data, string pic_name, ref int col)
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
                //byte[] img_data = imgToByteConverter(src_pic);
                var _img = BitmapToBytes(img_data, ImageFormat.Jpeg);
                var img = Export_Image_Class.imgToByteConverter(_img);

                //var steam = new MemoryStream(img_data);
                var stream = new MemoryStream(img);
                ExcelPicture pic = wsSheet1.Drawings.AddPicture(pic_name, stream);//img
                ExcelWorkbook wrkbk = wsSheet1.Workbook;
                decimal mdw = wrkbk.MaxFontWidth;
                int pixelHeight = (int)(row_h / 0.75);
                int pixelWidth = (int)decimal.Truncate(((256 * (decimal)col_w + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
                int offset = (int)(0.05 * Math.Min(pixelHeight, pixelWidth));
                pic.SetPosition(_rowIndex - 1, offset, _colIndex - 1, offset);
                pic.SetSize(pixelWidth - 2 * offset, pixelHeight - 2 * offset);
                col = c;
            }
        }


        /***********************  Export BVH-PTH Function **********************************************/

        public void Export_BVH_PTH(SqlConnection sqlcon, ExcelWorksheet ws, string ItemCode, string LotNo)
        {

            List<string> tar_table = new List<string>() { "BVH_WITH_BONDING_SHEET", "BVH_WITHOUT_BONDING_SHEET", "PLATED_THROUGH_HOLE" };
            List<DataTable> tar_dt = new List<DataTable>();
            for (int i = 0; i < tar_table.Count; i++)
            {
                tar_dt.Add(TDMK_Code.Datatable_Filter(sqlcon, tar_table[i], TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo })));
            }
            bool exp_en = false;
            foreach (var _dt in tar_dt)
            {
                if (_dt.Rows.Count > 0)
                {
                    exp_en = true;
                    break;
                }
            }
            if (exp_en)
            {
                foreach (var t in tar_table)
                {
                    DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, t, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                    DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_" + t, TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { ItemCode }));

                    if (dt_spec.Rows.Count > 0)
                    {
                        string list_val = dt_spec.Rows[0]["Remark"].ToString();
                        switch (t)
                        {
                            case "BVH_WITH_BONDING_SHEET":
                                DB_to_Excel(dt, ws, "BVH_WITH_BONDING_SHEET", 8, list_val);
                                break;
                            case "BVH_WITHOUT_BONDING_SHEET":
                                DB_to_Excel(dt, ws, "BVH_WITHOUT_BONDING_SHEET", 8, list_val);
                                break;
                            case "PLATED_THROUGH_HOLE":
                                DB_to_Excel(dt, ws, "PLATED_THROUGH_HOLE", 7, list_val);
                                break;
                            default:
                                break;
                        }
                    }
                }
            }

        }
        public void DB_to_Excel(System.Data.DataTable src_DGV, ExcelWorksheet curr_wrksheet, string process, int offset, string list_val)
        {
            //string image_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            //string file_dic = Path.Combine(image_folder, "1.jpg");

            int data_col = 8;

            int r = int.Parse(list_val.Split('+')[0]);
            int c = int.Parse(list_val.Split('+')[1]);
            int count_sample = int.Parse(list_val.Split('+')[2]);
            ExcelRangeBase curr_rgn = curr_wrksheet.Cells[r + 1, c + 2];
            ExcelRangeBase img_rgn = curr_rgn.Offset(0, offset);
            int qty = Math.Min(src_DGV.Rows.Count, count_sample);
            if (process == "PLATED_THROUGH_HOLE")
            {
                data_col = 7;
                curr_rgn = curr_wrksheet.Cells[r + 1, c + 1];
                img_rgn = curr_rgn.Offset(0, offset);
            }
            for (int m = 0; m < qty; m++)
            {
                int offs = int.Parse(src_DGV.Rows[m]["PCS_No"].ToString());

                for (int j = 0; j < data_col; j++)
                {
                    string cur_val = myCode.checkDBNull(src_DGV.Rows[m][j + 4]);
                    if (cur_val != "")
                    {
                        curr_rgn.Offset(offs - 1, j).Value = Convert.ToDecimal(cur_val);
                        curr_rgn.Offset(offs - 1, j).Style.Numberformat.Format = "#0.000";
                    }
                    else
                    {
                        curr_rgn.Offset(offs - 1, j).Value = "NA";
                    }
                }
            }

            for (int count = 0; count < qty; count++)
            {
                if(myCode.checkDBNull(src_DGV.Rows[count]["Image_data"])!="")
                {
                    byte[] img_data = (byte[])src_DGV.Rows[count]["Image_data"];
                    InsertPicture_Name_BVHPTH(curr_wrksheet, img_rgn, img_data, img_rgn.Address);
                }
                img_rgn = img_rgn.Offset(1, 0);
            }
        }
        public void InsertPicture_Name_BVHPTH(ExcelWorksheet wsSheet1, ExcelRangeBase tar_rgn, byte[] img_data, string pic_name)
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
                //byte[] img_data = imgToByteConverter(src_pic);
                using (var steam = new MemoryStream(img_data))
                {
                    ExcelPicture pic = wsSheet1.Drawings.AddPicture(pic_name, steam);//img
                    ExcelWorkbook wrkbk = wsSheet1.Workbook;
                    decimal mdw = wrkbk.MaxFontWidth;
                    int pixelHeight = (int)(row_h / 0.75);
                    int pixelWidth = (int)decimal.Truncate(((256 * (decimal)col_w + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
                    int offset = (int)(0.05 * Math.Min(pixelHeight, pixelWidth));
                    pic.SetPosition(_rowIndex - 1, offset, _colIndex - 1, offset);
                    pic.SetSize(pixelWidth - 2 * offset, pixelHeight - 2 * offset);
                }
                // var steam = new MemoryStream(img_data);
            }
        }


        /***********************  Export Impedance Function **********************************************/
        public void Export_Impedance(SqlConnection sqlcon, ExcelWorksheet ws, string ItemCode, string LotNo)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            List<string> tar_table_name = new List<string>() { "IMPEDANCE_VAL", "TRACEWIDTH_VAL", "TRACEWIDTH_IMAGE", "IMPEDANCE_GRAPH" };

            List<DataTable> tar_dt = new List<DataTable>();
            for (int i = 0; i < tar_table_name.Count; i++)
            {
                tar_dt.Add(TDMK_Code.Datatable_Filter(sqlcon, tar_table_name[i], filter_str));
            }
            bool exp_en = false;
            foreach (var _dt in tar_dt)
            {
                if (_dt.Rows.Count > 0)
                {
                    exp_en = true;
                    break;
                }
            }
            if (exp_en)
            {
                ExcelWorksheet curr_wrksheet = ws;
                for (int i = 0; i < tar_table_name.Count; i++)
                {
                    DataTable sel_dt = tar_dt[i];
                    switch (tar_table_name[i])
                    {
                        case "IMPEDANCE_VAL":
                            Export_Impedance_Val(sqlcon, sel_dt, curr_wrksheet, ItemCode);
                            break;
                        case "TRACEWIDTH_VAL":
                            Export_VHX_data_2(sqlcon, sel_dt, curr_wrksheet, ItemCode);
                            break;
                        case "TRACEWIDTH_IMAGE":
                            Export_Image_VHX(sqlcon, sel_dt, curr_wrksheet, ItemCode);
                            break;
                        case "IMPEDANCE_GRAPH":
                            Export_Graph(sqlcon, curr_wrksheet, ItemCode, LotNo);
                            break;
                    }
                }

            }
        }
        public void Get_ListTable(int col_inx, DataTable myDt, string[] src_arr, ref List<DataTable> src_lst_tbl, string tar_item)
        {
            DataRow[] temp_dr;
            if (col_inx < src_arr.Length - 1)
            {
                if (src_arr[col_inx + 1] != "")
                {
                    string[] sel_val = myDt.AsEnumerable().Select(x => x.Field<string>(src_arr[col_inx + 1])).Distinct().ToArray();
                    if (sel_val.Length != 0)
                    {
                        foreach (string sv in sel_val)
                        {
                            if (sv != null)
                            {
                                DataTable curTbl = myDt.AsEnumerable().Where(r => r.Field<string>(src_arr[col_inx + 1]) == sv).CopyToDataTable();
                                Get_ListTable(col_inx + 1, curTbl, src_arr, ref src_lst_tbl, tar_item);
                            }
                        }
                    }
                    else
                    {
                        temp_dr = myDt.AsEnumerable().Where(x => x.Field<string>(tar_item) != null).ToArray();
                        if (temp_dr.Length > 0)
                        {
                            src_lst_tbl.Add(myDt);
                        }
                    }
                }
            }
            else
            {
                temp_dr = myDt.AsEnumerable().Where(x => x.Field<string>(tar_item) != null).ToArray();
                if (temp_dr.Length > 0)
                {
                    src_lst_tbl.Add(myDt);
                }
            }
        }
        public void Export_Impedance_Val(SqlConnection sqlcon, DataTable src_dt, ExcelWorksheet curr_wrksheet, string ItemCode)
        {
            List<DataTable> src_region_tbl = new List<DataTable>();
            Get_ListTable(-1, src_dt, new string[] { "ItemCode", "LotNo", "Region" }, ref src_region_tbl, "Data");
            foreach (var in_tbl in src_region_tbl)
            {
                if (in_tbl.Rows.Count > 0)
                {
                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { ItemCode });
                    DataTable dt_format = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", filter_str);
                    string region = in_tbl.Rows[0]["Region"].ToString().Trim().Replace(" ", "");
                    //AutoCompleteStringCollection lst_rgn = Get_impedance_infor_to_input_value(curr_wrksheet);
                    int cou_rgn = 1;
                    ExcelRangeBase data_rgn;
                    for (int t = 0; t < dt_format.Rows.Count; t++)
                    {
                        string rgn = dt_format.Rows[t]["Range"].ToString().Split('-')[0];
                        string count_sample = dt_format.Rows[t]["Range"].ToString().Split('-')[1];
                        if (region == cou_rgn.ToString())
                        {
                            data_rgn = curr_wrksheet.Cells[rgn];
                            int r_inx = 0;
                            int min_sample = new int[] { int.Parse(count_sample), in_tbl.Rows.Count }.Min();
                            foreach (DataRow dr in in_tbl.Rows)
                            {
                                if (r_inx < min_sample)
                                {
                                    if (TDMK_Code.IsNumeric(myCode.checkDBNull(dr["Data"])))
                                    {
                                        data_rgn.Offset(r_inx, 0).Value = Convert.ToDecimal(dr["Data"]);
                                        data_rgn.Offset(r_inx, 0).Style.Numberformat.Format = "#0.00";
                                    }
                                    r_inx++;
                                }
                            }
                            break;
                        }
                        cou_rgn++;
                    }
                }
            }
        }
        public void Export_VHX_data_2(SqlConnection sqlcon, DataTable src_dt, ExcelWorksheet curr_wrksheet, string ItemCode)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { ItemCode });
            DataTable dt_format = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", filter_str);
            //SortedDictionary<int, string> dic_trw = new SortedDictionary<int, string> { };


            int idx_imp = 1;
            for (int t = 0; t < dt_format.Rows.Count; t++)
            {
                string region = dt_format.Rows[t]["Region"].ToString();
                string item_rgn = dt_format.Rows[t]["Range"].ToString();
                int r = int.Parse(item_rgn.Split(':')[0]);
                int c = int.Parse(item_rgn.Split(':')[1]);
                int count_sample = int.Parse(item_rgn.Split(':')[2]);

                DataTable s_dt = src_dt.Clone();
                foreach (DataRow dr in src_dt.Rows)
                {
                    if (dr["Data_For"].ToString().Contains(region))
                    {
                        DataRow row = s_dt.NewRow();
                        for (int i = 0; i < src_dt.Columns.Count; i++)
                        {
                            row[i] = dr[i];
                        }
                        s_dt.Rows.Add(row);
                    }
                }
                List<DataTable> src_sample = new List<DataTable>();
                Get_ListTable(-1, s_dt, new string[] { "ItemCode", "LotNo", "Pcs_No" }, ref src_sample, "Data");
                int min_pcs = Math.Min(s_dt.Rows.Count, count_sample);
                ExcelRangeBase rgn = curr_wrksheet.Cells[r, c];
                int inx = 0;
                int inx_off = 0;
                while(inx_off < min_pcs)
                {
                    int off_set = src_sample[inx].Rows.Count;
                    ExcelRangeBase data_rgn = rgn.Offset(inx_off, 0);
                    for (int j = 0; j < off_set; j++)
                    {
                        string cur_val = myCode.checkDBNull(src_sample[inx].Rows[j]["Data"]);
                        if (TDMK_Code.IsNumeric(cur_val))
                        {
                            data_rgn.Offset(j, 0).Value = Convert.ToDecimal(cur_val);
                            data_rgn.Offset(j, 0).Style.Numberformat.Format = "#0.000";
                        }
                    }
                    inx_off += off_set;
                    inx++;
                }


                //for (int k = 0; 2 * k < min_pcs; k++)
                //{
                //    int r_off = src_sample[k].Rows.Count;
                //    ExcelRangeBase data_rgn = rgn.Offset(2 * k, 0);
                //    for (int j = 0; j < 2; j++)
                //    {
                //        string cur_val = myCode.checkDBNull(src_sample[k].Rows[j]["Data"]);
                //        if (TDMK_Code.IsNumeric(cur_val))
                //        {
                //            data_rgn.Offset(j, 0).Value = Convert.ToDecimal(cur_val);
                //            data_rgn.Offset(j, 0).Style.Numberformat.Format = "#0.000";
                //        }
                //    }
                //}
                idx_imp++;

            }
        }
        public void Export_Image_VHX(SqlConnection sqlcon, DataTable src_dt, ExcelWorksheet curr_wrksheet, string ItemCode)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { ItemCode });
            DataTable dt_format = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", filter_str);

            // AutoCompleteStringCollection lst_rgn = Get_image_infor_to_input_value(curr_wrksheet);

            int idx_imp = 1;
            for (int t = 0; t < dt_format.Rows.Count; t++)
            {
                string item_rgn = dt_format.Rows[t]["Range"].ToString();
                string region = dt_format.Rows[t]["Region"].ToString();
                int r = int.Parse(item_rgn.Split(':')[0]) + 1;
                int c = int.Parse(item_rgn.Split(':')[1]) + 2;

                DataTable s_dt = src_dt.Clone();
                foreach (DataRow dr in src_dt.Rows)
                {
                    if (dr["Data_For"].ToString().Contains(region))
                    {
                        DataRow row = s_dt.NewRow();
                        for (int i = 0; i < src_dt.Columns.Count; i++)
                        {
                            row[i] = dr[i];
                        }
                        s_dt.Rows.Add(row);
                    }
                }
                ExcelRangeBase rgn_img_data = curr_wrksheet.Cells[r, c];
                Export_DatatableImage_Excel_2(s_dt, "Image_Tracewidth", curr_wrksheet, rgn_img_data, false);

                idx_imp++;
            }



        }

        public void Export_DatatableImage_Excel_2(DataTable dt, string Image_Col_name, ExcelWorksheet tar_wrksht, ExcelRangeBase sel_rgn, bool row_offset)
        {
            int r_inx = 0;
            foreach (DataRow dr in dt.Rows)
            {
                byte[] img_byte = (byte[])(dr[Image_Col_name]);
                int col = 1;
                InsertPicture_Name2(tar_wrksht, sel_rgn, img_byte, sel_rgn.Address, ref col);
                int col_offset = sel_rgn.Columns;
                int row_off = sel_rgn.Rows;
                if (row_offset)
                {
                    sel_rgn = sel_rgn.Offset(row_off, 0);
                }
                else
                {
                    sel_rgn = sel_rgn.Offset(0, col);
                }
                r_inx++;
            }
        }
        public void Export_Graph(SqlConnection sqlcon, ExcelWorksheet curr_wrksheet, string ItemCode, string LotNo)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { ItemCode });
            DataTable dt_format = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", filter_str);
            DataTable dt_format_imp = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", filter_str);
            List<DataTable> src_region_tbl = new List<DataTable>();

            string[] region = { "Coupon", "Patern" };
            int no_trace = 0;
            foreach (string _regn in region)
            {
                DataTable graph_tbl = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Zone" }, new string[] { ItemCode, LotNo, _regn }));
                string[] sel_val = graph_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();
                foreach (string item in sel_val)
                {
                    DataTable s_dt = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Zone", "Region" }, new string[] { ItemCode, LotNo, _regn, item }));
                    if (no_trace < dt_format.Rows.Count && no_trace < dt_format_imp.Rows.Count)
                    {
                        string item_rgn = dt_format.Rows[no_trace]["Range"].ToString();
                        int r = int.Parse(item_rgn.Split(':')[0]) + 1;
                        int c = int.Parse(item_rgn.Split(':')[1]) + 2;
                        ExcelRangeBase rgn_img_data = curr_wrksheet.Cells[r, c];
                        int r_offset = Get_Cells_Info(curr_wrksheet, rgn_img_data).row_qty;
                        int _r_off = Get_Cells_Info(curr_wrksheet, rgn_img_data.Offset(r_offset, 0)).row_qty;
                        ExcelRangeBase rgn_graph = rgn_img_data.Offset(r_offset + _r_off, 0);
                        Export_DatatableImage_Excel_2(s_dt, "Image_Graph", curr_wrksheet, rgn_graph, false);
                    }
                    int count = int.Parse(s_dt.Rows[0]["Region"].ToString());
                    no_trace += count;
                }
            }
        }

        public MergeAreas_Info Get_Cells_Info(ExcelWorksheet src_wrksht, ExcelRangeBase src_rgn)
        {
            MergeAreas_Info result = default(MergeAreas_Info);
            int row = src_rgn.End.Row;
            int column = src_rgn.End.Column;
            int mergeCellId = src_wrksht.GetMergeCellId(row, column);
            string address = src_rgn.Address;
            if (mergeCellId > 0)
            {
                address = src_wrksht.MergedCells[mergeCellId - 1];
            }

            result.col_qty = src_wrksht.Cells[address].Columns;
            result.row_qty = src_wrksht.Cells[address].Rows;
            return result;
        }


        /***********************  Export SolderMask Function **********************************************/
        public void Export_SolderMask(SqlConnection sqlcon, ExcelWorksheet ws, string ItemCode, string LotNo)
        {
            Export_Solder_Image_1(sqlcon, ws, ItemCode, LotNo);
            Export_LPI_data(ws, ItemCode, LotNo);

        }
        public void Export_LPI_data(ExcelWorksheet ws, string ItemCode, string LotNo)
        {
            string connstr_IPQC = TDMK_Code.data_connection(myVar.server_name, "IPQC_Data", myVar.server_acc, myVar.server_pass).ConnectionString;
            SqlConnection sqlcon_IPQC = new SqlConnection(connstr_IPQC);
            DataTable LPI_table = TDMK_Code.Datatable_Filter(sqlcon_IPQC, "Printing_Process", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
            List<string> col_name = new List<string>() { "LPI_SM_Coverage", "LPI_Open_B1", "LPI_SM_L1" };
            DataTable LPI_data = LPI_table.AsDataView().ToTable(false, col_name.ToArray());
            int qty = Math.Min(LPI_data.Rows.Count, 5);
            List<string> tar_rgn = new List<string>() { "H13", "L13", "M13" };
            for (int i = 0; i < qty; i++)
            {
                for (int j = 0; j < tar_rgn.Count; j++)
                {
                    string cur_val = myCode.checkDBNull(LPI_data.Rows[i][j]);
                    if (TDMK_Code.IsNumeric(cur_val))
                    {
                        ws.Cells[tar_rgn[j]].Offset(i, 0).Value = Convert.ToDecimal(cur_val);
                        ws.Cells[tar_rgn[j]].Offset(i, 0).Style.Numberformat.Format = "#0.000";
                    }
                }
            }
        }
        public void Export_Solder_Image_1(SqlConnection sqlcon, ExcelWorksheet ws, string Itemcode, string Lotno)
        {
            DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, "SOLDERMASK_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { Itemcode, Lotno }));
            string[] xx = new string[4] { "BGA", "Hot Bar", "Connector", "Trace to trace" };
            bool check = false;
            foreach (DataRow dr in dt.Rows)
            {
                foreach (string s in xx)
                {
                    if (dr["Region"].ToString() == s)
                    {
                        check = true;
                    }
                }
                if (!check)
                {
                    dr["Region"] = "Connector";
                }
            }

            int row_val;
            foreach (string zone in xx)
            {
                for (int x = 1; x < 40; x++)
                {
                    if (ws.Cells[x, 1].Value != null && ws.Cells[x, 1].Value.ToString() == zone)
                    {
                        row_val = (int)x;
                        int number_image = 5;
                        string[] item = { "ItemCode", "LotNo", "Region" };
                        string[] item_val = { Itemcode, Lotno, zone };
                        DataTable data_image = dt.Clone();
                        foreach (DataRow dr in dt.Rows)
                        {
                            if (dr["Region"].ToString() == zone)
                            {
                                DataRow dtrow = data_image.NewRow();
                                for (int i = 0; i < dt.Columns.Count; i++)
                                {
                                    dtrow[i] = dr[i];
                                }
                                data_image.Rows.Add(dtrow);
                            }

                        }
                        if (data_image.Rows.Count > 0)
                        {
                            if (data_image.Rows.Count < number_image)
                            {
                                number_image = data_image.Rows.Count;
                                ExcelRangeBase curr_rgn = ws.Cells["B" + row_val];
                                for (int m = 0; m < number_image; m++)
                                {
                                    byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                    int col = 1;
                                    InsertPicture_Name2(ws, curr_rgn, data, curr_rgn.Address, ref col);
                                    curr_rgn = curr_rgn.Offset(0, col);
                                }
                            }
                            else
                            {
                                ExcelRangeBase curr_rgn = ws.Cells["B" + row_val];
                                for (int m = 0; m < number_image; m++)
                                {
                                    byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                    int col = 1;
                                    InsertPicture_Name2(ws, curr_rgn, data, curr_rgn.Address, ref col);
                                    curr_rgn = curr_rgn.Offset(0, col);
                                }
                            }
                        }
                    }
                }

            }
        }

        public string remove_special_chars(string src_string, List<char> reject_chars)
        {
            string text = "";
            return new string((from x in src_string.ToArray()
                               where reject_chars.IndexOf(x) == -1
                               select x).ToArray());
        }

        public string Find_Start_Addr(string start_addr, ExcelWorksheet tar_wrksht, bool left_to_right, string search_key = "")
        {
            List<char> reject_chars = new List<char> { ' ', '-', '_', '\r', '\n' };
            string text = "";
            string address = start_addr;
            ExcelRangeBase excelRangeBase = tar_wrksht.Cells[start_addr];
            int num = 0;
            int num2 = 0;
            while (true)
            {
                ExcelRangeBase excelRangeBase2 = excelRangeBase.Offset(num, 0);
                if (left_to_right)
                {
                    excelRangeBase2 = excelRangeBase.Offset(0, num);
                }

                string text2 = myCode.checkDBNull(excelRangeBase2.Value);
                if (text2 == "")
                {
                    num2++;
                    if (num2 > 3)
                    {
                        break;
                    }
                }
                else
                {
                    num2 = 0;
                }

                if (search_key != "")
                {
                    if (remove_special_chars(text2, reject_chars).ToUpper() == remove_special_chars(search_key, reject_chars).ToUpper())
                    {
                        address = excelRangeBase2.Address;
                        break;
                    }
                }
                else if (num2 == 0)
                {
                    address = excelRangeBase2.Address;
                    break;
                }

                num++;
            }

            ExcelRange excelRange = tar_wrksht.Cells[address];
            int row = excelRange.End.Row;
            int column = excelRange.End.Column;
            int mergeCellId = tar_wrksht.GetMergeCellId(row, column);
            string address2 = excelRange.Address;
            if (mergeCellId > 0)
            {
                address2 = tar_wrksht.MergedCells[mergeCellId - 1];
            }

            int rows = tar_wrksht.Cells[address2].Rows;
            int columns = tar_wrksht.Cells[address2].Columns;
            if (left_to_right)
            {
                return excelRange.Offset(0, columns).Address;
            }

            return excelRange.Offset(rows, 0).Address;
        }





        /***********************  Export ThermalStress Function **********************************************/


        public void Export_Thermal_Stress_Image(SqlConnection sqlcon, ExcelWorksheet ws, string Itemcode, string Lotno)
        {
            AutoCompleteStringCollection list = Get_infor_to_input_thermalstress(ws);

            string[] item = { "ItemCode", "LotNo" };
            string[] item_val = { Itemcode, Lotno };
            DataTable Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, "THERMAL_STRESS_IMAGE", TDMK_Code.filter_str(item, item_val));


            List<DataTable> lst_tbl = new List<DataTable> { };
            Get_ListTable(-1, Data_tbl, new string[] { "Region" }, ref lst_tbl, "ItemCode");
            int i = 0;
            foreach (string x in list)
            {
                string region = x.Split('+')[0].ToUpper().Replace(" ", "").Replace("PHOTO", "").Replace("PICTURE", "");
                string row_val = x.Split('+')[1];

                int number_image = 5;
                if (i < lst_tbl.Count)
                {
                    Export_Image_new_ThermalStress(ws, number_image, lst_tbl[i], row_val);
                }
                i++;
            }

        }

        public void Export_Image_new_ThermalStress(ExcelWorksheet ws, int number_image, DataTable data_image, string row_val)
        {
            int count_sample = new int[] { number_image, data_image.Rows.Count }.Min();

            ExcelRangeBase curr_rgn = ws.Cells["B" + row_val];
            for (int m = 0; m < count_sample; m++)
            {
                byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                int col = 1;
                InsertPicture_Name2(ws, curr_rgn, data, curr_rgn.Address, ref col);
                curr_rgn = curr_rgn.Offset(0, col);
            }
        }

        public AutoCompleteStringCollection Get_infor_to_input_thermalstress(ExcelWorksheet ws)
        {
            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 10; i < 50; i++)
            {
                if (ws.Cells[i, 1].Value != null)
                {
                    if (ws.Cells[i, 1].Value.ToString().Contains("PTH"))
                    {
                        list.Add(ws.Cells[i, 1].Value.ToString() + "+" + (i).ToString());
                    }
                    if (ws.Cells[i, 1].Value.ToString().Contains("BVH"))
                    {
                        list.Add(ws.Cells[i, 1].Value.ToString() + "+" + (i).ToString());
                    }
                }

            }
            return list;
        }

    }
}
