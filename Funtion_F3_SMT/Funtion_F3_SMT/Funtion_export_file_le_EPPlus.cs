using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Style;
using OK2SHIP_SMT;
using OK2SHIP_SMT.Repositories;
using OK2SHIP_SMT.Services;
using Patagames.Ocr.Enums;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Deployment.Application;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using TDMK_Image;
using TDMK_SEEV_DLL;
using TDMK_SQL;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TreeView;
using DataTable = System.Data.DataTable;
using Image = System.Drawing.Image;
using myExcel = Microsoft.Office.Interop.Excel;

namespace Funtion_F3_SMT
{

    public class Funtion_export_file_le_EPPlus
    {
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public SEI_Lib myCode = new SEI_Lib();
        //public EPPlus_Lib TDMK_EPPLUS = new EPPlus_Lib();

        //export cross_section
        public void export_excel_cross_section(ExcelWorksheet ws, string itemcode, string lotno, DataTable Data_tbl, DataTable dt_spec, DataGridView dgv_data, string type, string leader, int count_sample)
        {
            string regionZ = "";
            int iz = 1;
            foreach (DataRow row in Data_tbl.Rows)
            {
                if (row["Region"].ToString().Equals(regionZ))
                {

                }
                else
                {
                    regionZ = row["Region"].ToString();
                    iz = 1;
                }
                row["Sample"] = iz++;
            }
            List<string> list = new List<string>();
            int counting = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
            for (int i = 0; i < counting; i++)
            {
                list.Add($"Sample {i + 1}");
            }
            list.Add("Flex SN");
            string address = "";
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(ws, list.ToArray(), true);
            foreach (DataRow row in Data_tbl.Rows)
            {
                string dif = "";
                string region = row["Region"].ToString();
                int samplePcs = int.Parse(row["Sample"].ToString()) % counting;
                int step = (int.Parse(row["Sample"].ToString()) - 1) / counting;
                if (samplePcs == 0)
                {
                    samplePcs = counting;
                }
                if (region.Contains("TRAI"))
                {
                    region = region.Replace("TRAI", "T");
                }
                if (region.Contains("PHAI"))
                {
                    region = region.Replace("PHAI", "P");
                }
                switch (region)
                {
                    case "NGANG":
                        if (dic.TryGetValue("Flex SN", out address))
                        {
                            if (dic.TryGetValue($"Sample {samplePcs}", out string addSample))
                            {
                                int vari = 0;
                                if (step <= 0)
                                {
                                    vari = 0;
                                }
                                if (step >= 1)
                                {
                                    vari = 1;
                                }
                                address = ws.Cells[ws.Cells[address.Split('-')[vari]].End.Row, ws.Cells[addSample].End.Column].Address;
                                ws.Cells[address].Value = row["ProductID"];
                                address = ExportProcess.AddRow(address, 2 - vari);
                                ExportProcess.InsertImageToCell(ws, ws.Cells[address], (byte[])row[$"Image1"], $"NGANG1{Guid.NewGuid()}");
                                address = ExportProcess.AddRow(address, 1);
                                ExportProcess.InsertImageToCell(ws, ws.Cells[address], (byte[])row[$"Image2"], $"NGANG2{Guid.NewGuid()}");
                                try
                                {
                                    string[] data = row["Data"].ToString().Replace(" ", "").TrimEnd(';').Split(';');
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[0]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[1]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                }
                                catch
                                {
                                    MessageBox.Show($"Có lỗi tại row {row["ID"]}");
                                }
                            }
                        }
                        break;
                    case "TRUNGANG_P":
                    case "TRUNGANG_T":
                        dif = region.Equals("TRUNGANG_P") ? "2 - 5" : "3 - 4";

                        if (dic.TryGetValue("Flex SN", out address))
                        {
                            int z = int.Parse(dif.Split('-')[step]);
                            address = address.Split('-')[z];
                            if (dic.TryGetValue($"Sample {samplePcs}", out string addSample))
                            {
                                address = ws.Cells[ws.Cells[address].End.Row, ws.Cells[addSample].End.Column].Address;
                                ws.Cells[address].Value = row["ProductID"];
                                address = ExportProcess.AddRow(address, 1);
                                ExportProcess.InsertImageToCell(ws, ws.Cells[address], (byte[])row["Image1"], $"{Guid.NewGuid()}");
                                try
                                {
                                    string[] data = row["Data"].ToString().Replace(" ", "").TrimEnd(';').Split(';');
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[0]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[1]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                }
                                catch
                                {
                                    MessageBox.Show($"Có lỗi tại row {row["ID"]}");
                                }

                            }
                        }

                        break;
                    case "TRU_T":
                    case "TRU_P":
                    case "DOC_T":
                    case "DOC_P":
                        switch (region)
                        {
                            case "TRU_T":
                                step = 6;
                                break;
                            case "TRU_P":
                                step = 7;
                                break;
                            case "DOC_T":
                                step = 8;
                                break;
                            case "DOC_P":
                                step = 9;
                                break;
                        }
                        if (dic.TryGetValue("Flex SN", out address))
                        {
                            address = address.Split('-')[step];
                            if (dic.TryGetValue($"Sample {samplePcs}", out string addSample))
                            {
                                address = ws.Cells[ws.Cells[address].End.Row, ws.Cells[addSample].End.Column].Address;
                                if (Data_tbl.Columns.Contains("ProductID"))
                                {
                                    ws.Cells[address].Value = row["ProductID"].ToString();
                                }
                                address = ExportProcess.AddRow(address, 1);
                                ExportProcess.InsertImageToCell(ws, ws.Cells[address], (byte[])row["Image1"], $"{Guid.NewGuid()}");
                                address = ExportProcess.AddRow(address, 1);
                                ExportProcess.InsertImageToCell(ws, ws.Cells[address], (byte[])row["Image2"], $"{Guid.NewGuid()}");
                                try
                                {
                                    string[] data = row["Data"].ToString().Replace(" ", "").Split('/');
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[0].TrimEnd(';').Split(';')[0]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[0].TrimEnd(';').Split(';')[1]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[1].TrimEnd(';').Split(';')[1]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[1].TrimEnd(';').Split(';')[4]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[1].TrimEnd(';').Split(';')[2]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[1].TrimEnd(';').Split(';')[5]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[1].TrimEnd(';').Split(';')[0]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[1].TrimEnd(';').Split(';')[3]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                }
                                catch
                                {
                                    MessageBox.Show($"Có lỗi tại row {row["ID"]}");
                                }
                            }
                        }
                        break;
                    default:
                        break;
                }
            }
        }


        public SortedDictionary<int, string> dic_judge_crosscut(int count_sample, DataTable dt_data, DataGridView dgv_data, ref bool judge_all)
        {
            List<DataTable> lst_Table = new List<DataTable> { };
            Get_ListTable(-1, dt_data, new string[] { "Region" }, ref lst_Table, "Data");
            SortedDictionary<int, string> dic_judge = new SortedDictionary<int, string> { };

            foreach (DataTable dt in lst_Table)
            {
                if (count_sample < dt.Rows.Count)
                {
                    string reggion = dt.Rows[0]["Region"].ToString();
                    for (int i = 0; i < count_sample; i++)
                    {
                        if (reggion.Contains("NGANG"))
                        {
                            int ID1 = int.Parse(dt.Rows[i]["ID"].ToString());
                            int ID2 = int.Parse(dt.Rows[i + count_sample]["ID"].ToString());
                            if (dgv_data.Rows[ID1 - 1].Cells["Data"].Style.BackColor == Color.Red || dgv_data.Rows[ID2 - 1].Cells["Data"].Style.BackColor == Color.Red)
                            {
                                if (!dic_judge.ContainsKey(i + 1))
                                {
                                    dic_judge.Add(i + 1, "Fail");
                                    judge_all = false;
                                }
                            }
                        }
                        else
                        {
                            int ID = int.Parse(dt.Rows[i]["ID"].ToString());
                            if (dgv_data.Rows[ID - 1].Cells["Data"].Style.BackColor == Color.Red)
                            {
                                if (!dic_judge.ContainsKey(i + 1))
                                {
                                    dic_judge.Add(i + 1, "Fail");
                                    judge_all = false;
                                }

                            }

                        }
                    }
                }
            }

            return dic_judge;
        }

        public ExcelRangeBase find_cell(ExcelWorksheet ws, string find_item)
        {
            ExcelRangeBase cur_cell = null;
            if (find_item == "IPQC")
                find_item = "SMT";

            for (int j = 1; j < 14; j++)
            {
                for (int i = 1; i < 14; i++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).ToUpper().Replace(" ", "").Contains(find_item))
                    {
                        cur_cell = ws.Cells[i, j];
                        return cur_cell;

                    }
                }
            }

            return cur_cell;

        }
        public void complete_sheet_other(ExcelWorksheet ws, int c_setup, int r_ofset_del, string sheet, DataTable dt_spec)
        {

            int c_format = int.Parse(dt_spec.Rows[0]["Count_sample"].ToString());
            if (c_setup < c_format)
            {
                if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST" || sheet == "SHEAR_TEST")
                {
                    foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                    {
                        if (spec_region != "")
                        {
                            string[] str_location = spec_region.Split('+');
                            ExcelRangeBase cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[1]), int.Parse(str_location[1].Split(';')[2]) + 1];

                            for (int j = 0; j < c_format - c_setup; j++)
                            {
                                ExcelRangeBase cell_begin = cell_Pic.Offset(-2, c_setup + j);
                                for (int i = 0; i < r_ofset_del; i++)
                                {
                                    cell_begin.Offset(i, 0).Value = "";
                                    //  cell_begin.Offset(j, 0).Style.Border.BorderAround(ExcelBorderStyle.None);
                                    cell_begin.Offset(i, 0).Style.Border.Top.Style = ExcelBorderStyle.None;
                                    cell_begin.Offset(i, 0).Style.Border.Bottom.Style = ExcelBorderStyle.None;
                                    cell_begin.Offset(i, 0).Style.Border.Left.Style = ExcelBorderStyle.None;
                                    cell_begin.Offset(i, 0).Style.Border.Right.Style = ExcelBorderStyle.None;
                                }
                            }
                        }
                    }
                }
                else if (sheet == "CROSS_SECTION")
                {
                    int row_pic = int.Parse(dt_spec.Rows[0]["Location"].ToString().Split('+')[1].Split('_')[0].Split(';')[2]);
                    int col_pic = int.Parse(dt_spec.Rows[0]["Location"].ToString().Split('+')[0].Split(';')[1]);
                    ExcelRangeBase cell_pic = ws.Cells[row_pic, col_pic];

                    for (int j = 0; j < c_format - c_setup; j++)
                    {
                        ExcelRangeBase cell_begin = cell_pic.Offset(-2, c_setup + j);
                        for (int i = 0; i < r_ofset_del; i++)
                        {
                            cell_begin.Offset(i, 0).Value = "";
                            //  cell_begin.Offset(j, 0).Style.Border.BorderAround(ExcelBorderStyle.None);
                            cell_begin.Offset(i, 0).Style.Border.Top.Style = ExcelBorderStyle.None;
                            cell_begin.Offset(i, 0).Style.Border.Bottom.Style = ExcelBorderStyle.None;
                            cell_begin.Offset(i, 0).Style.Border.Left.Style = ExcelBorderStyle.None;
                            cell_begin.Offset(i, 0).Style.Border.Right.Style = ExcelBorderStyle.None;

                        }
                    }

                }
                else if (sheet.Contains("ON_PRODUCT") || sheet.Contains("COUPON"))
                {
                    foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                    {
                        if (spec_region != "")
                        {
                            string[] str_location = spec_region.Split('+');
                            ExcelRangeBase cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[2]), int.Parse(str_location[1].Split(';')[3]) + 1];
                            for (int j = 0; j < c_format - c_setup; j++)
                            {
                                ExcelRangeBase cell_begin = cell_Pic.Offset(-2, c_setup + j);
                                for (int i = 0; i < r_ofset_del; i++)
                                {
                                    cell_begin.Offset(i, 0).Value = "";
                                    // cell_begin.Offset(j, 0).Style.Border.BorderAround(ExcelBorderStyle.None);
                                    cell_begin.Offset(i, 0).Style.Border.Top.Style = ExcelBorderStyle.None;
                                    cell_begin.Offset(i, 0).Style.Border.Bottom.Style = ExcelBorderStyle.None;
                                    cell_begin.Offset(i, 0).Style.Border.Left.Style = ExcelBorderStyle.None;
                                    cell_begin.Offset(i, 0).Style.Border.Right.Style = ExcelBorderStyle.None;
                                }
                            }
                        }
                    }
                }
                else
                {
                    foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                    {
                        if (spec_region != "")
                        {
                            string[] str_location = spec_region.Split('+');
                            ExcelRangeBase cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[1]), int.Parse(str_location[1].Split(';')[2]) + 1];

                            for (int j = 0; j < c_format - c_setup; j++)
                            {
                                ExcelRangeBase cell_begin = cell_Pic.Offset(-2, c_setup + j);
                                for (int i = 0; i < r_ofset_del; i++)
                                {
                                    cell_begin.Offset(i, 0).Value = "";
                                    //cell_begin.Offset(j, 0).Style.Border.BorderAround(ExcelBorderStyle.None);
                                    cell_begin.Offset(i, 0).Style.Border.Top.Style = ExcelBorderStyle.None;
                                    cell_begin.Offset(i, 0).Style.Border.Bottom.Style = ExcelBorderStyle.None;
                                    cell_begin.Offset(i, 0).Style.Border.Left.Style = ExcelBorderStyle.None;
                                    cell_begin.Offset(i, 0).Style.Border.Right.Style = ExcelBorderStyle.None;
                                }
                            }
                        }
                    }
                }
            }

        }


        public void insert_columns_other(ExcelWorksheet ws, int c_setup, int row_insert, string type, DataTable dt_spec, string sheet)
        {
            int c_format = int.Parse(dt_spec.Rows[0]["Count_sample"].ToString());
            if (c_setup > c_format)
            {
                ExcelRangeBase cell_end = null;
                int k_offset = -2;
                if (type == "NPI")
                {
                    k_offset = -1;
                }
                if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST" || sheet == "SHEAR_TEST")
                {
                    foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                    {
                        if (spec_region != "")
                        {
                            string[] str_location = spec_region.Split('+');
                            ExcelRangeBase cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[1]), int.Parse(str_location[1].Split(';')[2]) + 1];
                            if (cell_end == null)
                                cell_end = cell_Pic.Offset(row_insert, c_setup - 1);
                            double width_col = ws.Columns[int.Parse(str_location[1].Split(';')[2]) + 1].Width;
                            for (int i = 0; i < c_setup - c_format; i++)
                            {
                                ExcelRangeBase cell_begin = cell_Pic.Offset(k_offset, c_format + i);
                                int c_begin = int.Parse(str_location[1].Split(';')[2]) + 1;
                                ws.Columns[c_begin + c_format + i].Width = width_col + 5;

                                float size_text = cell_Pic.Offset(k_offset, c_format - 1).Style.Font.Size;
                                cell_begin.Value = "Sample " + (c_format + i + 1).ToString();
                                for (int j = 0; j < 13; j++)
                                {
                                    cell_begin.Offset(j, 0).Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    cell_begin.Offset(j, 0).Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                    cell_begin.Offset(j, 0).Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                    cell_begin.Offset(j, 0).Style.Border.Right.Style = ExcelBorderStyle.Thin;

                                    cell_begin.Offset(j, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    cell_begin.Offset(j, 0).Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                    cell_begin.Offset(j, 0).Style.Font.Size = size_text;
                                    //cell_begin.Offset(j, 0).AutoFitColumns();
                                }
                            }
                        }

                    }
                    if (cell_end != null)
                    {
                        ws.PrinterSettings.PrintArea = ws.Cells["A1:" + cell_end.Address];
                    }
                }
                else if (sheet == "CROSS_SECTION")
                {
                    int row_pic = int.Parse(dt_spec.Rows[0]["Location"].ToString().Split('+')[1].Split('_')[0].Split(';')[2]);
                    int col_pic = int.Parse(dt_spec.Rows[0]["Location"].ToString().Split('+')[0].Split(';')[1]);
                    ExcelRangeBase cell_pic = ws.Cells[row_pic, col_pic];
                    if (cell_end == null)
                        cell_end = cell_pic.Offset(row_insert, c_setup - 1);

                    double width_col = ws.Columns[col_pic].Width;
                    for (int i = 0; i < c_setup - c_format; i++)
                    {
                        ExcelRangeBase cell_begin = cell_pic.Offset(k_offset, c_format + i);

                        ws.Columns[col_pic + c_format + i].Width = width_col;
                        ExcelRangeBase rgn_from = ws.Cells[cell_begin.Offset(0, -1).Address + ":" + cell_begin.Offset(row_insert, -1).Address];
                        ExcelRangeBase rgn_to = ws.Cells[cell_begin.Address + ":" + cell_begin.Offset(row_insert, 0).Address];
                        rgn_from.Copy(rgn_to);
                        cell_begin.Value = "Sample " + (c_format + i + 1).ToString();
                    }
                    if (cell_end != null)
                    {
                        ws.PrinterSettings.PrintArea = ws.Cells["A1:" + cell_end.Address];
                    }


                }
                else if (sheet.Contains("ON_PRODUCT"))
                {
                    foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                    {
                        if (spec_region != "")
                        {
                            string[] str_location = spec_region.Split('+');
                            ExcelRangeBase cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[2]), int.Parse(str_location[1].Split(';')[3]) + 1];
                            float size_text = cell_Pic.Offset(k_offset, c_format - 1).Style.Font.Size;
                            if (cell_end == null)
                                cell_end = cell_Pic.Offset(row_insert, c_setup - 1);
                            double width_col = ws.Columns[int.Parse(str_location[1].Split(';')[3]) + 1].Width;
                            for (int i = 0; i < c_setup - c_format; i++)
                            {
                                ExcelRangeBase cell_begin = cell_Pic.Offset(k_offset, c_format + i);
                                //string col_name = cell_begin.Address.Split('$')[2];
                                //int col = cell_begin.End.Column;
                                int c_begin = int.Parse(str_location[1].Split(';')[3]) + 1;

                                ws.Columns[c_begin + c_format + i].Width = width_col;

                                cell_begin.Value = "Sample " + (c_format + i + 1).ToString();
                                for (int j = 0; j < 8; j++)
                                {
                                    cell_begin.Offset(j, 0).Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    cell_begin.Offset(j, 0).Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                    cell_begin.Offset(j, 0).Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                    cell_begin.Offset(j, 0).Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                    cell_begin.Offset(j, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    cell_begin.Offset(j, 0).Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                    cell_begin.Offset(j, 0).Style.Font.Size = size_text;
                                }

                            }
                        }
                    }
                    if (cell_end != null)
                    {
                        ws.PrinterSettings.PrintArea = ws.Cells["A1:" + cell_end.Address];
                    }

                }
                else
                {
                    foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                    {
                        if (spec_region != "")
                        {
                            string[] str_location = spec_region.Split('+');
                            ExcelRangeBase cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[1]), int.Parse(str_location[1].Split(';')[2]) + 1];
                            if (cell_end == null)
                                cell_end = cell_Pic.Offset(row_insert, c_setup - 1);
                            double width_col = ws.Columns[int.Parse(str_location[1].Split(';')[2]) + 1].Width;
                            for (int i = 0; i < c_setup - c_format; i++)
                            {
                                ExcelRangeBase cell_begin = cell_Pic.Offset(k_offset, c_format + i);

                                int c_begin = int.Parse(str_location[1].Split(';')[2]) + 1;
                                ws.Columns[c_begin + c_format + i].Width = width_col;
                                float size_text = cell_Pic.Offset(k_offset, c_format - 1).Style.Font.Size;
                                cell_begin.Value = "Sample " + (c_format + i + 1).ToString();
                                for (int j = 0; j < 8; j++)
                                {
                                    cell_begin.Offset(j, 0).Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    cell_begin.Offset(j, 0).Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                    cell_begin.Offset(j, 0).Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                    cell_begin.Offset(j, 0).Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                    cell_begin.Offset(j, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    cell_begin.Offset(j, 0).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    cell_begin.Offset(j, 0).Style.Font.Size = size_text;

                                }

                            }
                        }
                    }
                    if (cell_end != null)
                    {
                        ws.PrinterSettings.PrintArea = ws.Cells["A1:" + cell_end.Address];
                    }

                }
            }
        }
        public Boolean Check_spec_crossection(DataTable Data_tbl, string ItemCode, DataTable dt_spec)
        {
            bool chk = true;
            if (dt_spec.Rows.Count > 0)
            {

                string spec = dt_spec.Rows[0]["Location"].ToString().Split('+')[0].Split(';')[2];
                if (spec.Contains("<"))
                {
                    spec = spec.Split('<')[1].Replace(" ", string.Empty).Replace("µm", string.Empty).Replace(")", string.Empty);
                    if (TDMK_Code.IsNumeric(spec))
                    {
                        for (int i = 0; i < Data_tbl.Rows.Count; i++)
                        {
                            string val = Data_tbl.Rows[i]["Data"].ToString();
                            if (val != "")
                            {
                                if (Double.Parse(val.Split(';')[0]) >= Double.Parse(spec))
                                {
                                    return false;
                                }

                                if (val.Contains("/"))
                                {
                                    foreach (string data in val.Replace(" ", "").Split('/')[1].Split(';'))
                                    {
                                        if (TDMK_Code.IsNumeric(data))
                                        {
                                            if (Double.Parse(data) > 4.5 || Double.Parse(data) < 1.5)
                                            {

                                                return false;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

            }
            return chk;
        }

        public void Export_DatatableImage_Excel_tape(DataTable dt, string Image_Col_name, ExcelWorksheet tar_wrksht, ExcelRangeBase sel_rgn, bool row_offset, int count_sample)
        {
            string file_folder = System.IO.Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = System.IO.Path.Combine(file_folder, DateTime.Now.ToString("ddMMMyy_HHmmss") + ".jpg");
            int r_inx = 0;
            //ExcelRangeBase sel_rgn = tar_wrksht.Range[tar_rgn];
            ExcelRangeBase rgn_begin = sel_rgn;

            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (count < count_sample)
                {
                    byte[] img_byte = (byte[])(dr[Image_Col_name]);

                    //Image img = byteArrayToImage(img_byte);
                    //Image test_img = resizeImage(0.2, img);

                    //byte[] img_data = imgToByteConverter(test_img);
                    //InsertPicture_Name(tar_wrksht, sel_rgn, img_data, sel_rgn.Address);

                    int col_offset = sel_rgn.Columns;
                    int row_off = sel_rgn.Rows;
                    if (row_offset)
                    {
                        sel_rgn = sel_rgn.Offset(row_off, 0);
                    }
                    else
                    {
                        sel_rgn = sel_rgn.Offset(0, col_offset);
                    }

                    r_inx++;
                }
                count++;
            }
        }

        public void resize_image(byte[] img_byte, ExcelWorksheet tar_wrksht, ExcelRangeBase sel_rgn)
        {
            double size_x = 1500;
            Image img = byteArrayToImage(img_byte);
            double ratio = size_x / img.PhysicalDimension.Width;
            if (ratio < 1)
            {
                Image test_img = resizeImage(ratio, img);
                byte[] img_data = imgToByteConverter(test_img);
                InsertPicture_Name(tar_wrksht, sel_rgn, img_data, sel_rgn.Address);
            }
            else
            {
                InsertPicture_Name(tar_wrksht, sel_rgn, img_byte, sel_rgn.Address);
            }

        }


        public void Export_DatatableImage_Excel(DataTable dt, string Image_Col_name, ExcelWorksheet tar_wrksht, ExcelRangeBase sel_rgn, bool row_offset, int count_sample)
        {
            string file_folder = System.IO.Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            //string file_dic = System.IO.Path.Combine(file_folder, DateTime.Now.ToString("ddMMMyy_HHmmss") + ".jpg");
            int r_inx = 0;
            //ExcelRangeBase sel_rgn = tar_wrksht.Range[tar_rgn];
            ExcelRangeBase rgn_begin = sel_rgn;

            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (count < count_sample)
                {
                    byte[] _img_byte = (byte[])(dr[Image_Col_name]);
                    byte[] img_byte = ImageCompress.CompressImage(_img_byte);

                    InsertPicture_Name(tar_wrksht, sel_rgn, img_byte, sel_rgn.Address);
                    int col_offset = sel_rgn.Columns;
                    int row_off = sel_rgn.Rows;
                    if (row_offset)
                    {
                        sel_rgn = sel_rgn.Offset(row_off, 0);
                    }
                    else
                    {
                        sel_rgn = sel_rgn.Offset(0, col_offset);
                    }
                    r_inx++;
                }
                count++;
            }
        }
        public Image byteArrayToImage(byte[] byteArrayIn)
        {
            using (MemoryStream mStream = new MemoryStream(byteArrayIn))
            {
                return Image.FromStream(mStream);
            }
        }

        public void fill_val(ExcelRangeBase cell, string val)
        {
            if (myCode.IsNumeric(val))
            {
                cell.Value = double.Parse(val);
                cell.Style.Numberformat.Format = "0.00";
            }
        }

        public void export_Hottizontal_NPI(string reg, ExcelWorksheet ws, int row_begin, int col_begin, int no_lk, List<string> lst_region, int count_row, DataTable Data_tbl, DataTable dt_spec)
        {
            int count_sample = int.Parse(dt_spec.Rows[0]["Count_sample"].ToString());
            ExcelRangeBase cell_Pic1 = ws.Cells[row_begin, col_begin];
            ExcelRangeBase cell_Pic2 = ws.Cells[row_begin + 1, col_begin];
            ExcelRangeBase cell_data = ws.Cells[row_begin + 2, col_begin];
            string filter_reg = "";

            string txt_find = "";
            if (reg.ToUpper().Contains("MALE"))
            {
                txt_find = no_lk.ToString() + "_";
            }

            foreach (string i in lst_region)
            {
                if (i.ToUpper().Contains("NGANG") && !i.ToUpper().Contains("TRU") && i.ToUpper().Contains(txt_find))
                {
                    filter_reg = i;
                    break;
                }
            }
            if (filter_reg != "")
            {
                DataView dv = Data_tbl.AsDataView();
                dv.RowFilter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { filter_reg });
                DataTable tbl_region = dv.ToTable();

                if (tbl_region.Rows.Count > 0)
                {
                    Export_DatatableImage_Excel_ngang(tbl_region, "Image1", ws, cell_Pic1, false, count_sample);
                    Export_DatatableImage_Excel_ngang(tbl_region, "Image2", ws, cell_Pic2, false, count_sample);

                    List<string> lst_data = tbl_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                    int k = 0;
                    for (int i = 0; i < tbl_region.Rows.Count; i++)
                    {
                        if (i < 2 * count_sample)
                        {
                            //cell_data.Offset(0, k).Value = double.Parse(lst_data[i].Split(';')[0]);
                            //cell_data.Offset(0, k).Style.Numberformat.Format = "0.00";
                            //cell_data.Offset(1, k).Value = double.Parse(lst_data[i].Split(';')[1]);
                            //cell_data.Offset(1, k).Style.Numberformat.Format = "0.00";
                            fill_val(cell_data.Offset(0, k), lst_data[i].Split(';')[0]);
                            fill_val(cell_data.Offset(1, k), lst_data[i].Split(';')[1]);
                            k++;

                            if (i == count_sample - 1)
                            {
                                cell_data = cell_data.Offset(4, 0);
                                k = 0;
                            }
                        }
                        if (i == 2 * count_sample - 1)
                        {
                            break;
                        }
                    }
                }
            }

            switch (count_row)
            {
                case 20:

                    row_begin = row_begin + 8;
                    Dictionary<string, int> _dic = new Dictionary<string, int> { };
                    _dic.Add("_T", 9);
                    _dic.Add("_P", 3);
                    foreach (var item in _dic)
                    {
                        ExcelRangeBase cell_Pic1_2 = ws.Cells[row_begin, col_begin];
                        ExcelRangeBase cell_data_2 = ws.Cells[row_begin + 1, col_begin];

                        string filter_reg_2 = "";
                        // string txt_find_2 = "";
                        if (reg.ToUpper().Contains("MALE"))
                        {
                            txt_find = "MALE";

                        }
                        else if (reg.ToUpper().Contains("FEMALE"))
                        {
                            txt_find = "FEMALE";
                        }

                        foreach (string i in lst_region)
                        {
                            if (i.Replace(" ", "").ToUpper().Contains("TRUNGANG") && i.Replace(" ", "").ToUpper().Contains(item.Key) && i.ToUpper().Contains(txt_find))
                            {
                                filter_reg_2 = i;
                                break;
                            }
                        }
                        if (filter_reg_2 != "")
                        {
                            //DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo, filter_reg_2 }));

                            DataView dv = Data_tbl.AsDataView();
                            dv.RowFilter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { filter_reg_2 });
                            DataTable tbl_region = dv.ToTable();

                            if (Data_tbl.Rows.Count > 0)
                            {
                                // int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());

                                Export_DatatableImage_Excel_trungang(tbl_region, "Image1", ws, cell_Pic1_2, false, count_sample, item.Value);
                                //Export_DatatableImage_Excel_trungang(tbl_region, "Image2", ws, cell_Pic2, false, count_sample, item.Value);

                                List<string> lst_data = tbl_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                                int k = 0;
                                for (int i = 0; i < tbl_region.Rows.Count; i++)
                                {
                                    if (i < 2 * count_sample)
                                    {
                                        //cell_data_2.Offset(0, k).Value = double.Parse(lst_data[i].Split(';')[0]);
                                        //cell_data_2.Offset(0, k).Style.Numberformat.Format = "0.00";
                                        //cell_data_2.Offset(1, k).Value = double.Parse(lst_data[i].Split(';')[1]);
                                        //cell_data_2.Offset(1, k).Style.Numberformat.Format = "0.00";
                                        fill_val(cell_data_2.Offset(0, k), lst_data[i].Split(';')[0]);
                                        fill_val(cell_data_2.Offset(1, k), lst_data[i].Split(';')[1]);
                                        k++;
                                        if (i == count_sample - 1)
                                        {
                                            cell_data_2 = cell_data_2.Offset(item.Value, 0);
                                            k = 0;
                                        }
                                    }
                                    if (i == 2 * count_sample - 1)
                                    {
                                        break;
                                    }
                                }
                            }

                        }
                        row_begin = row_begin + 3;
                    }
                    break;

                case 26:
                    row_begin = row_begin + 8;
                    Dictionary<string, int> _dic1 = new Dictionary<string, int> { };
                    _dic1.Add("1_T", 9);
                    _dic1.Add("1_P", 3);
                    _dic1.Add("2", 3);

                    foreach (var item in _dic1)
                    {
                        ExcelRangeBase cell_Pic1_2 = ws.Cells[row_begin, col_begin];
                        ExcelRangeBase cell_data_2 = ws.Cells[row_begin + 1, col_begin];

                        string filter_reg_2 = "";
                        if (reg.ToUpper().Contains("MALE"))
                        {
                            txt_find = "MALE";

                        }
                        else if (reg.ToUpper().Contains("FEMALE"))
                        {
                            txt_find = "FEMALE";
                        }

                        foreach (string i in lst_region)
                        {
                            if (i.Replace(" ", "").ToUpper().Contains("TRUNGANG") && i.Replace(" ", "").ToUpper().Contains(item.Key) && i.ToUpper().Contains(txt_find))
                            {
                                filter_reg_2 = i;
                                break;
                            }
                        }
                        if (filter_reg_2 != "")
                        {
                            //DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo, filter_reg_2 }));
                            DataView dv = Data_tbl.AsDataView();
                            dv.RowFilter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { filter_reg_2 });
                            DataTable tbl_region = dv.ToTable();

                            if (tbl_region.Rows.Count > 0)
                            {
                                //int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
                                Export_DatatableImage_Excel_trungang(tbl_region, "Image1", ws, cell_Pic1_2, false, count_sample, item.Value);

                                List<string> lst_data = tbl_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                                int k = 0;
                                for (int i = 0; i < tbl_region.Rows.Count; i++)
                                {
                                    if (i < 2 * count_sample)
                                    {
                                        //cell_data_2.Offset(0, k).Value = double.Parse(lst_data[i].Split(';')[0]);
                                        //cell_data_2.Offset(0, k).Style.Numberformat.Format = "0.00";
                                        //cell_data_2.Offset(1, k).Value = double.Parse(lst_data[i].Split(';')[1]);
                                        //cell_data_2.Offset(1, k).Style.Numberformat.Format = "0.00";
                                        fill_val(cell_data_2.Offset(0, k), lst_data[i].Split(';')[0]);
                                        fill_val(cell_data_2.Offset(1, k), lst_data[i].Split(';')[1]);
                                        k++;
                                        if (i == count_sample - 1)
                                        {
                                            cell_data_2 = cell_data_2.Offset(item.Value, 0);
                                            k = 0;
                                        }
                                    }
                                    if (i == 2 * count_sample - 1)
                                    {
                                        break;
                                    }
                                }
                            }

                        }
                        if (item.Key == "1_T")
                        {
                            row_begin = row_begin + 3;
                        }
                        else if (item.Key == "1_P")
                        {
                            row_begin = row_begin + 9;
                        }

                    }
                    break;

                case 32:
                    row_begin = row_begin + 8;
                    Dictionary<string, int> _dic2 = new Dictionary<string, int> { };
                    _dic2.Add("1_T", 9);
                    _dic2.Add("1_P", 3);
                    _dic2.Add("2_T", 9);
                    _dic2.Add("2_P", 3);
                    foreach (var item in _dic2)
                    {
                        ExcelRangeBase cell_Pic1_2 = ws.Cells[row_begin, col_begin];
                        ExcelRangeBase cell_data_2 = ws.Cells[row_begin + 1, col_begin];

                        string filter_reg_2 = "";
                        if (reg.ToUpper().Contains("MALE"))
                        {
                            txt_find = "MALE";

                        }
                        else if (reg.ToUpper().Contains("FEMALE"))
                        {
                            txt_find = "FEMALE";
                        }
                        foreach (string i in lst_region)
                        {
                            if (i.Replace(" ", "").ToUpper().Contains("TRUNGANG") && i.Replace(" ", "").ToUpper().Contains(item.Key) && i.ToUpper().Contains(txt_find))
                            {
                                filter_reg_2 = i;
                                break;
                            }
                        }
                        if (filter_reg_2 != "")
                        {
                            //DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo, filter_reg_2 }));
                            DataView dv = Data_tbl.AsDataView();
                            dv.RowFilter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { filter_reg_2 });
                            DataTable tbl_region = dv.ToTable();

                            if (tbl_region.Rows.Count > 0)
                            {
                                //int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
                                Export_DatatableImage_Excel_trungang(tbl_region, "Image1", ws, cell_Pic1_2, false, count_sample, item.Value);

                                List<string> lst_data = tbl_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                                int k = 0;
                                for (int i = 0; i < tbl_region.Rows.Count; i++)
                                {
                                    if (i < 2 * count_sample)
                                    {
                                        //cell_data_2.Offset(0, k).Value = double.Parse(lst_data[i].Split(';')[0]);
                                        //cell_data_2.Offset(0, k).Style.Numberformat.Format = "0.00";
                                        //cell_data_2.Offset(1, k).Value = double.Parse(lst_data[i].Split(';')[1]);
                                        //cell_data_2.Offset(1, k).Style.Numberformat.Format = "0.00";
                                        fill_val(cell_data_2.Offset(0, k), lst_data[i].Split(';')[0]);
                                        fill_val(cell_data_2.Offset(1, k), lst_data[i].Split(';')[1]);
                                        k++;
                                        if (i == count_sample - 1)
                                        {
                                            cell_data_2 = cell_data_2.Offset(item.Value, 0);
                                            k = 0;
                                        }
                                    }
                                    if (i == 2 * count_sample - 1)
                                    {
                                        break;
                                    }
                                }
                            }

                        }
                        if (item.Key.Contains("T"))
                        {
                            row_begin = row_begin + 3;
                        }
                        else
                        {
                            row_begin = row_begin + 9;
                        }

                    }
                    break;


                default:
                    break;

            }
        }

        public void Export_DatatableImage_Excel_ngang_mass(DataTable dt, string Image_Col_name, ExcelWorksheet tar_wrksht, ExcelRangeBase sel_rgn, bool row_offset, int count_sample)
        {
            int r_inx = 0;

            ExcelRangeBase rgn_begin = sel_rgn;

            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (count < 2 * count_sample)
                {
                    byte[] _img_byte = (byte[])(dr[Image_Col_name]);
                    byte[] img_byte = ImageCompress.CompressImage(_img_byte);
                    //Image img = byteArrayToImage(img_byte);
                    //Image test_img = resizeImage(0.2, img);
                    //byte[] img_data = imgToByteConverter(test_img);

                    InsertPicture_Name(tar_wrksht, sel_rgn, img_byte, sel_rgn.Address);
                    //InsertPicture_Name(tar_wrksht, sel_rgn, img_data, sel_rgn.Address);
                    int col_offset = sel_rgn.Columns;
                    int row_off = sel_rgn.Rows;
                    if (row_offset)
                    {
                        sel_rgn = sel_rgn.Offset(row_off, 0);
                    }
                    else
                    {
                        sel_rgn = sel_rgn.Offset(0, col_offset);
                    }
                    r_inx++;
                    count++;

                    if (count == count_sample)
                    {
                        sel_rgn = sel_rgn.Offset(3, -count_sample);

                    }
                }
                else
                {
                    break;
                }
            }
        }

        public void export_Hottizontal_MASS(string reg, ExcelWorksheet ws, int row_begin, int col_begin, int no_lk, List<string> lst_region, int count_row, DataTable Data_tbl, DataTable dt_spec, int count_sample)
        {
            ExcelRangeBase cell_Pic = ws.Cells[row_begin, col_begin];
            ExcelRangeBase cell_data = ws.Cells[row_begin + 1, col_begin];
            string filter_reg = "";
            string txt_find = "";
            if (reg.ToUpper().Contains("MALE"))
            {
                txt_find = no_lk.ToString() + "_";
            }


            foreach (string i in lst_region)
            {
                if (i.ToUpper().Contains("NGANG") && !i.ToUpper().Contains("TRU") && i.ToUpper().Contains(txt_find))
                {
                    filter_reg = i;
                    break;
                }
            }
            if (filter_reg != "")
            {
                DataView dv = Data_tbl.AsDataView();
                dv.RowFilter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { filter_reg });
                DataTable tbl_region = dv.ToTable();

                if (tbl_region.Rows.Count > 0)
                {
                    Export_DatatableImage_Excel_ngang_mass(tbl_region, "Image1", ws, cell_Pic, false, count_sample);
                    //Export_DatatableImage_Excel_ngang_mass(tbl_region, "Image2", ws, cell_Pic, false, count_sample);

                    List<string> lst_data = tbl_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                    int k = 0;
                    for (int i = 0; i < tbl_region.Rows.Count; i++)
                    {
                        if (i < 2 * count_sample)
                        {
                            if (myCode.IsNumeric(lst_data[i].Split(';')[0]))
                            {
                                //cell_data.Offset(0, k).Value = Math.Round(double.Parse(lst_data[i].Split(';')[0]), 2);
                                //cell_data.Offset(0, k).Style.Numberformat.Format = "0.00";
                                fill_val(cell_data.Offset(0, k), lst_data[i].Split(';')[0]);


                            }
                            if (myCode.IsNumeric(lst_data[i].Split(';')[1]))
                            {
                                //cell_data.Offset(1, k).Value = Math.Round(double.Parse(lst_data[i].Split(';')[1]), 2);
                                //cell_data.Offset(1, k).Style.Numberformat.Format = "0.00";
                                fill_val(cell_data.Offset(1, k), lst_data[i].Split(';')[1]);
                            }
                            k++;
                            if (i == count_sample - 1)
                            {
                                cell_data = cell_data.Offset(3, 0);
                                k = 0;
                            }
                        }
                        if (i == 2 * count_sample - 1)
                        {
                            break;
                        }
                    }
                }
            }
        }

        public void Export_DatatableImage_Excel_trungang(DataTable dt, string Image_Col_name, ExcelWorksheet tar_wrksht, ExcelRangeBase sel_rgn, bool row_offset, int count_sample, int r_offset)
        {
            int r_inx = 0;

            ExcelRangeBase rgn_begin = sel_rgn;

            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (count < 2 * count_sample)
                {
                    byte[] _img_byte = (byte[])(dr[Image_Col_name]);
                    byte[] img_byte = ImageCompress.CompressImage(_img_byte);
                    ////Image img = byteArrayToImage(img_byte);
                    ////Image test_img = resizeImage(0.2, img);
                    ////byte[] img_data = imgToByteConverter(test_img);

                    InsertPicture_Name(tar_wrksht, sel_rgn, img_byte, sel_rgn.Address);
                    //resize_image(img_byte, tar_wrksht, sel_rgn);

                    int col_offset = sel_rgn.Columns;
                    int row_off = sel_rgn.Rows;
                    if (row_offset)
                    {
                        sel_rgn = sel_rgn.Offset(row_off, 0);
                    }
                    else
                    {
                        sel_rgn = sel_rgn.Offset(0, col_offset);
                    }
                    r_inx++;
                    count++;

                    if (count == count_sample)
                    {
                        sel_rgn = sel_rgn.Offset(r_offset, -count_sample);
                        //count = 0;
                    }
                }
            }


        }
        //public void InsertPicture_Name_old(ExcelWorksheet tar_wrksht,ExcelRangeBase tar_range, string picFile, int margin)
        //{
        //    float left;
        //    float top;
        //    float width;
        //    float height;
        //    string pic_name = System.IO.Path.GetFileName(picFile);

        //    if (tar_range.Merge)
        //    {
        //        // ExcelRangeBase refer_range = tar_range.MergeArea;//Range["G16"];
        //        ExcelRangeBase refer_range = tar_range;

        //        left = (float)(tar_range.Left) + margin;
        //        top = (float)(tar_range.Top) + margin;
        //        width = (float)(refer_range.Width) - 2 * margin;
        //        height = (float)(refer_range.Height) - 2 * margin;
        //        //width = (float)(refer_range.Width * 2.8) + 2 * margin;
        //        //height = (float)(refer_range.Height * 8.9) + 2 * margin;
        //    }
        //    else
        //    {
        //        left = (float)tar_range.Left + margin;
        //        top = (float)tar_range.Top + margin;
        //        width = (float)tar_range.Width - 2 * margin;
        //        height = (float)tar_range.Height - 2 * margin;
        //    }
        //    myExcel.Shape sel_picture = tar_wrksht.Shapes.AddPicture(picFile, MsoTriState.msoFalse, MsoTriState.msoTrue, left, top, width, height);
        //    //tar_wrksht.Paste(tar_range, sel_picture);
        //    sel_picture.LockAspectRatio = MsoTriState.msoTrue;
        //    sel_picture.Placement = myExcel.XlPlacement.xlMoveAndSize;
        //    sel_picture.Name = pic_name; 
        //}
        public byte[] imgToByteConverter(Image inImg)
        {
            ImageConverter imgCon = new ImageConverter();
            return (byte[])imgCon.ConvertTo(inImg, typeof(byte[]));
        }

        public void InsertPicture_Name(ExcelWorksheet wsSheet1, ExcelRangeBase tar_rgn, byte[] img_data, string pic_name)
        {
            //using (tar_rgn)
            //{
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

            // byte[] img_data = imgToByteConverter(src_pic);
            using (MemoryStream ms = new MemoryStream(img_data))
            {
                ExcelPicture pic = wsSheet1.Drawings.AddPicture(pic_name, ms);//img 
                ExcelWorkbook wrkbk = wsSheet1.Workbook;
                decimal mdw = (decimal)wrkbk.MaxFontWidth;
                int pixelHeight = (int)(row_h / 0.75);
                int pixelWidth = (int)decimal.Truncate(((256 * (decimal)col_w + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
                int offset = (int)(0.05 * Math.Min(pixelHeight, pixelWidth));
                pic.SetPosition(_rowIndex - 1, offset, _colIndex - 1, offset);
                pic.SetSize(pixelWidth - 2 * offset, pixelHeight - 2 * offset);

            }

        }

        public Image BitmapToBytes(byte[] image, ImageFormat pFormat)
        {
            var imageObject = new Bitmap(new MemoryStream(image));

            var stream = new MemoryStream();
            imageObject.Save(stream, pFormat);

            return new Bitmap(stream);
        }

        public void InsertPicture_Name_old(ExcelWorksheet wsSheet1, ExcelRangeBase tar_rgn, byte[] img_data, string pic_name)
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
            var img = imgToByteConverter(_img);

            //using (MemoryStream ms = new MemoryStream(img_data))
            //{
            using (MemoryStream ms = new MemoryStream(img))
            {
                ExcelPicture pic = wsSheet1.Drawings.AddPicture(pic_name, ms);//img 
                ExcelWorkbook wrkbk = wsSheet1.Workbook;
                decimal mdw = (decimal)wrkbk.MaxFontWidth;
                int pixelHeight = (int)(row_h / 0.75);
                int pixelWidth = (int)decimal.Truncate(((256 * (decimal)col_w + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
                int offset = (int)(0.05 * Math.Min(pixelHeight, pixelWidth));
                pic.SetPosition(_rowIndex - 1, offset, _colIndex - 1, offset);
                pic.SetSize(pixelWidth - 2 * offset, pixelHeight - 2 * offset);

            }
        }

        public void InsertPicture_Name_Graph(ExcelWorksheet wsSheet1, ExcelRangeBase tar_rgn, byte[] img_data, string pic_name)
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

            using (MemoryStream ms = new MemoryStream(img_data))
            {
                ExcelPicture pic = wsSheet1.Drawings.AddPicture(pic_name, ms);//img
                ExcelWorkbook wrkbk = wsSheet1.Workbook;
                decimal mdw = (decimal)wrkbk.MaxFontWidth;
                int pixelHeight = (int)(row_h / 0.75);
                int pixelWidth = (int)decimal.Truncate(((256 * (decimal)col_w + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
                int offset = (int)(0.05 * Math.Min(pixelHeight, pixelWidth));
                pic.SetPosition(_rowIndex - 1, offset, _colIndex - 1, offset);
                pic.SetSize(pixelWidth - 2 * offset, pixelHeight - 2 * offset);

            }
        }
        public void Export_DatatableImage_Excel_ngang(DataTable dt, string Image_Col_name, ExcelWorksheet tar_wrksht, ExcelRangeBase sel_rgn, bool row_offset, int count_sample)
        {
            int r_inx = 0;
            ExcelRangeBase rgn_begin = sel_rgn;

            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (count < 2 * count_sample)
                {
                    byte[] _img_byte = (byte[])(dr[Image_Col_name]);
                    byte[] img_byte = ImageCompress.CompressImage(_img_byte);
                    InsertPicture_Name(tar_wrksht, sel_rgn, img_byte, sel_rgn.Address);


                    ////Image img = byteArrayToImage(img_byte);
                    ////Image test_img = resizeImage(0.2, img);
                    ////byte[] img_data = imgToByteConverter(test_img); 
                    // resize_image(img_byte, tar_wrksht, sel_rgn);

                    int col_offset = sel_rgn.Columns;
                    int row_off = sel_rgn.Rows;
                    if (row_offset)
                    {
                        sel_rgn = sel_rgn.Offset(row_off, 0);
                    }
                    else
                    {
                        sel_rgn = sel_rgn.Offset(0, col_offset);
                    }
                    r_inx++;
                    count++;



                    if (count == count_sample)
                    {
                        sel_rgn = sel_rgn.Offset(4, -count_sample);
                        //count = 0;
                    }
                }
                else
                {
                    break;
                }
            }
        }


        // export GAP connector
        public void export_excel_gap_connector(ExcelWorksheet ws, DataTable Data_tbl, DataTable dt_spec)
        {
            GAPConnectorService gapService = new GAPConnectorService();

            gapService.Export(ws, Data_tbl, dt_spec);
            #region OLDCODE
            //int col_begin = int.Parse(dt_spec.Rows[0]["Location"].ToString().Split('+')[0]);
            //string[] region = dt_spec.Rows[0]["Location"].ToString().Split('+')[1].Split('_');
            //int count_sample = int.Parse(dt_spec.Rows[0]["Count_sample"].ToString());

            //List<string> lst_region_db = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToList();
            //SortedDictionary<int, string> dic_judgement = Check_spec_GAP(Data_tbl, dt_spec, "GAP_CONNECTOR");

            //foreach (string reg in region)
            //{
            //    if (reg != "")
            //    {
            //        if (reg.Split(';')[0] == "IOPIN")
            //        {
            //            int code = 1;
            //            foreach (string pos in reg.Split('^'))
            //            {
            //                if (pos != "")
            //                {
            //                    int row_begin = int.Parse(pos.Split(';')[1]);
            //                    int count_r_offset = int.Parse(pos.Split(';')[2]);
            //                    ExcelRangeBase cell_Pic = ws.Cells[row_begin, col_begin];
            //                    ExcelRangeBase cell_Pic1 = ws.Cells[row_begin + 1, col_begin];
            //                    ExcelRangeBase cell_Pic2 = ws.Cells[row_begin + count_r_offset + 6, col_begin];
            //                    ExcelRangeBase cell_data = ws.Cells[row_begin + 2, col_begin];
            //                    string find_text = "";

            //                    if (reg.Split('^').Length == 3)
            //                    {
            //                        find_text = code.ToString();
            //                    }
            //                    string filter_reg = "";
            //                    foreach (string i in lst_region_db)
            //                    {
            //                        if (i.ToUpper().Contains("DOC") && i.Contains(find_text))
            //                        {
            //                            filter_reg = i;
            //                            break;
            //                        }
            //                    }
            //                    // DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo, filter_reg }));

            //                    if (filter_reg != "")
            //                    {
            //                        DataTable tbl_region = Data_tbl.AsEnumerable().Where(s => s.Field<string>("Region") == filter_reg).CopyToDataTable();
            //                        if (tbl_region.Rows.Count > 0)
            //                        {
            //                            // int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString()); 
            //                            Export_DatatableImage_Excel(tbl_region, "Image", ws, cell_Pic, false, count_sample);
            //                            Export_DatatableImage_Excel(tbl_region, "Image1", ws, cell_Pic1, false, count_sample);
            //                            Export_DatatableImage_Excel(tbl_region, "Image2", ws, cell_Pic2, false, count_sample);

            //                            List<string> lst_data = tbl_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
            //                            //List<double> lst_1 = new List<double> { };
            //                            //List<double> lst_2 = new List<double> { };
            //                            int offset = 0;

            //                            bool chk_allsample = true;
            //                            for (int i = 0; i < count_sample; i++)
            //                            {
            //                                if (i < tbl_region.Rows.Count)
            //                                {

            //                                    int min_r = new int[] { count_r_offset, lst_data[i].Split('/')[0].Split(';').Length }.Min();
            //                                    offset = min_r;
            //                                    //lst_1.Add(double.Parse(lst_data[i].Split('/')[0].Split(';')[min_r - 1]));
            //                                    //lst_2.Add(double.Parse(lst_data[i].Split('/')[1].Split(';')[min_r - 1]));
            //                                    for (int k = 0; k < min_r; k++)
            //                                    {
            //                                        cell_data.Offset(k, i).Value = double.Parse(lst_data[i].Split('/')[0].Split(';')[k]);
            //                                        cell_data.Offset(k, i).Style.Numberformat.Format = "0.00";

            //                                        cell_data.Offset(k + count_r_offset + 1, i).Value = double.Parse(lst_data[i].Split('/')[1].Split(';')[k]);
            //                                        cell_data.Offset(k + count_r_offset + 1, i).Style.Numberformat.Format = "0.00";
            //                                    }


            //                                    int ID = int.Parse(tbl_region.Rows[i]["ID"].ToString());

            //                                    if (dic_judgement.ContainsKey(ID - 1))
            //                                    {
            //                                        if (dic_judgement[ID - 1] == "FAIL")
            //                                        {
            //                                            cell_data.Offset(min_r + count_r_offset + 1, i).Value = "NG";
            //                                            chk_allsample = false;

            //                                        }
            //                                        else
            //                                        {
            //                                            cell_data.Offset(min_r + count_r_offset + 1, i).Value = "OK";
            //                                        }
            //                                    }

            //                                }

            //                            }


            //                            if (chk_allsample)
            //                            {
            //                                for (int i = 0; i < 3; i++)
            //                                {
            //                                    cell_data.Offset(offset + count_r_offset + 1, count_sample + i).Value = "OK";
            //                                }
            //                            }
            //                            else
            //                            {
            //                                for (int i = 0; i < 3; i++)
            //                                {
            //                                    cell_data.Offset(offset + count_r_offset + 1, count_sample + i).Value = "NG";
            //                                }
            //                            }

            //                            //cell_data.Offset(offset + count_r_offset + 1, count_sample).Value = "OK";
            //                            //if (lst_1.Max() - lst_1.Min() > 30 || lst_2.Max() - lst_2.Min() > 30)
            //                            //{
            //                            //    cell_data.Offset(offset + count_r_offset + 1, count_sample + 1).Value = "NG";
            //                            //    cell_data.Offset(offset + count_r_offset + 1, count_sample + 2).Value = "NG";
            //                            //}
            //                            //else
            //                            //{
            //                            //    cell_data.Offset(offset + count_r_offset + 1, count_sample + 1).Value = "OK";
            //                            //    cell_data.Offset(offset + count_r_offset + 1, count_sample + 2).Value = "OK";
            //                            //}


            //                        }
            //                    }

            //                }
            //                code++;
            //            }
            //        }
            //        else if (reg.Split(';')[0] == "LEFT" || reg.Split(';')[0] == "RIGHT")
            //        {
            //            int code = 1;
            //            foreach (string pos in reg.Split('^'))
            //            {
            //                if (pos != "")
            //                {
            //                    int row_begin = int.Parse(pos.Split(';')[1]);
            //                    int count_r_offset = int.Parse(pos.Split(';')[2]);
            //                    ExcelRangeBase cell_Pic = ws.Cells[row_begin, col_begin];
            //                    ExcelRangeBase cell_Pic1 = ws.Cells[row_begin + 1, col_begin];
            //                    ExcelRangeBase cell_Pic2 = ws.Cells[row_begin + count_r_offset + 8, col_begin];
            //                    ExcelRangeBase cell_data = ws.Cells[row_begin + 2, col_begin];

            //                    string find_text = "";

            //                    if (reg.Split('^').Length == 3)
            //                    {
            //                        find_text = code.ToString();
            //                    }
            //                    string filter_reg = "";
            //                    foreach (string i in lst_region_db)
            //                    {
            //                        if (i.ToUpper().Contains("TRU") && i.Contains(find_text))
            //                        {
            //                            filter_reg = i;
            //                            break;
            //                        }
            //                    }
            //                    //DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo, filter_reg }));
            //                    if (filter_reg != "")
            //                    {
            //                        DataTable tbl_region = Data_tbl.AsEnumerable().Where(s => s.Field<string>("Region") == filter_reg).CopyToDataTable();
            //                        if (tbl_region.Rows.Count > 0)
            //                        {
            //                            //int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString()); 
            //                            if (reg.Contains("RIGHT"))
            //                            {
            //                                for (int r = 0; r < count_sample; r++)
            //                                {
            //                                    tbl_region.Rows.Remove(tbl_region.Rows[0]);
            //                                }
            //                            }

            //                            Export_DatatableImage_Excel(tbl_region, "Image", ws, cell_Pic, false, count_sample);
            //                            Export_DatatableImage_Excel(tbl_region, "Image1", ws, cell_Pic1, false, count_sample);
            //                            Export_DatatableImage_Excel(tbl_region, "Image2", ws, cell_Pic2, false, count_sample);


            //                            List<string> lst_data = tbl_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
            //                            List<double> lst_1 = new List<double> { };
            //                            List<double> lst_2 = new List<double> { };
            //                            int offset = 0;
            //                            for (int i = 0; i < count_sample; i++)
            //                            {
            //                                if (i < tbl_region.Rows.Count)
            //                                {
            //                                    try
            //                                    {
            //                                        int a = lst_data[i].Split('/')[0].Split(';').Length;
            //                                        int b = lst_data[i].Split('/')[1].Split(';').Length;
            //                                        string z = lst_data[i].Split('/')[1];
            //                                        int min_r = new int[] { count_r_offset, a, b }.Min();

            //                                        offset = min_r;
            //                                        lst_1.Add(double.Parse(lst_data[i].Split('/')[0].Split(';')[min_r]));
            //                                        lst_2.Add(double.Parse(lst_data[i].Split('/')[1].Split(';')[min_r]));


            //                                        for (int k = 0; k < min_r; k++)
            //                                        {
            //                                            cell_data.Offset(k, i).Value = double.Parse(lst_data[i].Split('/')[0].Split(';')[k]);
            //                                            cell_data.Offset(k, i).Style.Numberformat.Format = "0.00";
            //                                            cell_data.Offset(k + count_r_offset + 1, i).Value = double.Parse(lst_data[i].Split('/')[1].Split(';')[k]);
            //                                            cell_data.Offset(k + count_r_offset + 1, i).Style.Numberformat.Format = "0.00";
            //                                        }


            //                                        int ID = int.Parse(tbl_region.Rows[i]["ID"].ToString());

            //                                        if (dic_judgement.ContainsKey(ID - 1))
            //                                        {
            //                                            if (dic_judgement[ID - 1] == "FAIL")
            //                                            {
            //                                                cell_data.Offset(min_r + count_r_offset + 1, i).Value = "NG";

            //                                            }
            //                                            else
            //                                            {
            //                                                cell_data.Offset(min_r + count_r_offset + 1, i).Value = "OK";
            //                                            }
            //                                        }
            //                                    }
            //                                    catch
            //                                    {
            //                                        throw new Exception("Lỗi dữ liệu đẩy vào hãy kiểm tra lại logfile!");

            //                                    }
            //                                }
            //                            }
            //                            cell_data.Offset(offset + count_r_offset + 1, count_sample).Value = "OK";

            //                            if (lst_1.Max() - lst_1.Min() > 30 || lst_2.Max() - lst_2.Min() > 30)
            //                            {
            //                                cell_data.Offset(offset + count_r_offset + 1, count_sample + 1).Value = "NG";
            //                                cell_data.Offset(offset + count_r_offset + 1, count_sample + 2).Value = "NG";
            //                            }
            //                            else
            //                            {
            //                                cell_data.Offset(offset + count_r_offset + 1, count_sample + 1).Value = "OK";
            //                                cell_data.Offset(offset + count_r_offset + 1, count_sample + 2).Value = "OK";
            //                            }
            //                        }
            //                    }
            //                }
            //                code++;
            //            }
            //        }
            //    }
            //}
            #endregion
        }
        public void export_excel_onproduct(ExcelWorksheet ws, string itemcode, string lotno, string sheet, DataTable Data_all, DataTable dt_spec, DataGridView dgv_data, string type, string leader, int count_sample)
        {
            bool judge_all = true;

            List<string> lst_region = Data_all.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToList();
            string[] region = dt_spec.Rows[0]["Location"].ToString().Split('_');

            if (type.Contains("MASS") || type == "Other")
            {
                string str_infor = Data_all.Rows[0]["Sheet"].ToString().Replace(sheet, "").Replace("/", "");
                if (str_infor.Contains("_"))
                {
                    string itemname = str_infor.Split('_')[0];
                    string line = str_infor.Split('_')[1];
                    string ca = str_infor.Split('_')[2];
                    string date = str_infor.Split('_')[3];
                    string worker = str_infor.Split('_')[4];
                    List<string> lst_infor = new List<string> { itemname, itemcode, lotno, line, ca, date, worker };
                    export_info_mass(ws, lst_infor, sheet, leader);
                }
                for (int i = 1; i < 20; i++)
                {
                    for (int j = 1; j < 4; j++)
                    {
                        if (myCode.checkDBNull(ws.Cells[i, j].Value).ToUpper().Replace(" ", "").Contains("ITEM-LOT"))
                        {
                            ws.Cells[i, j].Offset(0, 1).Value = itemcode + "-" + lotno;
                            break;

                        }
                    }
                }
            }

            int r_end = 0;
            if (region.Length == 2)
            {
                r_end = 25;
            }
            else if (region.Length == 3)
            {
                r_end = 40;
            }
            else if (region.Length == 4)
            {
                r_end = 52;
            }
            insert_columns_other(ws, count_sample, r_end, type, dt_spec, sheet);

            foreach (string reg in region)
            {
                if (reg != "")
                {
                    string[] str_location = reg.Split('+');
                    string cpn = "";

                    foreach (string r in lst_region)
                    {
                        string key1 = r.Split('_')[0].Replace(" ", "").ToUpper();
                        string key2 = r.Split('_')[1].Replace(" ", "").ToUpper();

                        if (reg.Replace("Component", "Tape").Replace(" ", "").ToUpper().Contains(key1) && reg.Replace("Component", "Tape").Replace(" ", "").ToUpper().Contains(key2))
                        {
                            cpn = r;
                            break;
                        }
                    }
                    ExcelRangeBase cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[2]), int.Parse(str_location[1].Split(';')[3]) + 1];
                    ExcelRangeBase cell_graph = ws.Cells[int.Parse(str_location[2].Split(';')[1]), int.Parse(str_location[1].Split(';')[3]) + 1];
                    ExcelRangeBase cell_data = ws.Cells[int.Parse(str_location[3].Split(';')[1]), int.Parse(str_location[1].Split(';')[3]) + 1];

                    if (cpn != "")
                    {
                        DataTable Data_tbl = Data_all.AsEnumerable().Where(s => s.Field<string>("Region") == cpn).CopyToDataTable();
                        if (Data_tbl.Rows.Count > 0)
                        {

                            Export_DatatableImage_Excel(Data_tbl, "Image", ws, cell_Pic, false, count_sample);
                            Export_DatatableImage_Excel_Graph(Data_tbl, "Graph", ws, cell_graph, false, count_sample);

                            List<string> lst_data = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                            List<string> lst_judge = new List<string> { };
                            for (int i = 0; i < count_sample; i++)
                            {
                                if (i < Data_tbl.Rows.Count)
                                {
                                    if (myCode.IsNumeric(lst_data[i].Split('_')[0].Replace("Max:", "")))
                                    {
                                        cell_data.Offset(0, i).Value = Math.Round(double.Parse(lst_data[i].Split('_')[0].Replace("Max:", "")), 2);
                                        cell_data.Offset(0, i).Style.Numberformat.Format = "0.00";
                                    }
                                    if (myCode.IsNumeric(lst_data[i].Split('_')[1].Replace("Average:", "")))
                                    {
                                        cell_data.Offset(1, i).Value = Math.Round(double.Parse(lst_data[i].Split('_')[1].Replace("Average:", "")), 2);
                                        cell_data.Offset(0, i).Style.Numberformat.Format = "0.00";
                                    }
                                }
                            }

                            for (int i = 1; i < 13; i++)
                            {
                                if (myCode.checkDBNull(cell_data.Offset(i, -1).Value).Replace(" ", "").ToUpper().Contains("MINFORCE"))
                                {
                                    ExcelRangeBase cell_min = cell_data.Offset(i, 0);
                                    ExcelRangeBase cell_max = cell_data.Offset(i + 1, 0);
                                    ExcelRangeBase cell_ave = cell_data.Offset(i + 2, 0);
                                    int c_offset = count_sample;
                                    int r_offset = i - 1;
                                    cell_min.FormulaR1C1 = "=MIN(R[-" + r_offset.ToString() + "]C[0]:R[-" + r_offset.ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                    cell_max.FormulaR1C1 = "=MAX(R[-" + (r_offset + 1).ToString() + "]C[0]:R[-" + (r_offset + 1).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                    cell_ave.FormulaR1C1 = "=AVERAGE(R[-" + (r_offset + 2).ToString() + "]C[0]:R[-" + (r_offset + 2).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                    cell_min.Style.Numberformat.Format = "0.00";
                                    cell_max.Style.Numberformat.Format = "0.00";
                                    cell_ave.Style.Numberformat.Format = "0.00";

                                    if (type == "NPI")
                                    {
                                        ExcelRangeBase cell_min1 = cell_data.Offset(i, 2);
                                        ExcelRangeBase cell_max1 = cell_data.Offset(i + 1, 2);
                                        ExcelRangeBase cell_ave1 = cell_data.Offset(i + 2, 2);

                                        cell_min1.FormulaR1C1 = "=MIN(R[-" + (i - 1).ToString() + "]C[-2]:R[-" + (i - 1).ToString() + "]C[" + (c_offset - 3).ToString() + "])";
                                        cell_max1.FormulaR1C1 = "=MAX(R[-" + i.ToString() + "]C[-2]:R[-" + i.ToString() + "]C[" + (c_offset - 3).ToString() + "])";
                                        cell_ave1.FormulaR1C1 = "=AVERAGE(R[-" + (i + 1).ToString() + "]C[-2]:R[-" + (i + 1).ToString() + "]C[" + (c_offset - 3).ToString() + "])";

                                        cell_min1.Style.Numberformat.Format = "0.00";
                                        cell_max1.Style.Numberformat.Format = "0.00";
                                        cell_ave1.Style.Numberformat.Format = "0.00";
                                    }

                                    break;
                                }
                            }


                            for (int i = 0; i < count_sample; i++)
                            {
                                if (i < Data_tbl.Rows.Count)
                                {
                                    int ID = int.Parse(Data_tbl.Rows[i]["ID"].ToString());
                                    if (dgv_data.Rows[ID - 1].Cells["Data"].Style.BackColor == Color.Red)
                                    {
                                        lst_judge.Add("Fail");
                                        judge_all = false;
                                    }
                                    else
                                    {
                                        lst_judge.Add("Pass");
                                    }
                                }
                            }

                            for (int t = 1; t < 5; t++)
                            {
                                if (myCode.checkDBNull(cell_data.Offset(t, -1).Value).ToUpper().Contains("JUDGEMENT") && !myCode.checkDBNull(cell_data.Offset(t, -1).Value).ToUpper().Replace(" ", "").Contains("FAILUREMODE"))
                                {
                                    for (int i = 0; i < lst_judge.Count; i++)
                                    {
                                        cell_data.Offset(t, i).Value = lst_judge[i];
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            if (type.Contains("MASS"))
            {
                ExcelRangeBase cur_rgn_judge = find_cell(ws, "JUDGEMENT");
                if (cur_rgn_judge != null)
                {
                    if (judge_all)
                    {
                        cur_rgn_judge.Offset(1, 0).Value = "OK";
                        ws.Cells[1, 1].Value = "OK";
                    }
                    else
                    {
                        cur_rgn_judge.Offset(1, 0).Value = "NG";
                        ws.Cells[1, 1].Value = "NG";
                    }
                }
            }

            int r_offset_del = 8;
            complete_sheet_other(ws, count_sample, r_offset_del, sheet, dt_spec);

        }


        //export onproduct
        public void export_excel_onproduct_NPI(ExcelWorksheet ws, DataTable Data_all, DataTable dt_spec, string sheet, ref bool export_ok)
        {

            int count_sample = int.Parse(dt_spec.Rows[0]["Count_sample"].ToString());
            if (dt_spec.Rows.Count > 0)
            {
                List<string> lst_region = Data_all.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToList();
                string[] region = dt_spec.Rows[0]["Location"].ToString().Split('_');

                SortedDictionary<int, string> dic_judgement = check_spec_onproduct(Data_all, dt_spec, sheet);

                bool check_all = true;
                foreach (var result in dic_judgement.Values)
                {
                    if (result == "FAIL")
                    {
                        check_all = false;
                        break;
                    }
                }
                if (check_all)
                {
                    foreach (string reg in region)
                    {
                        if (reg != "")
                        {
                            string[] str_location = reg.Split('+');
                            string cpn = "";

                            foreach (string r in lst_region)
                            {
                                string key1 = r.Split('_')[0].Replace(" ", "").ToUpper();
                                string key2 = r.Split('_')[1].Replace(" ", "").ToUpper();

                                if (reg.Replace("Component", "Tape").Replace(" ", "").ToUpper().Contains(key1) && reg.Replace("Component", "Tape").Replace(" ", "").ToUpper().Contains(key2))
                                {
                                    cpn = r;
                                    break;
                                }
                            }
                            ExcelRangeBase cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[2]), int.Parse(str_location[1].Split(';')[3]) + 1];
                            ExcelRangeBase cell_graph = ws.Cells[int.Parse(str_location[2].Split(';')[1]), int.Parse(str_location[1].Split(';')[3]) + 1];
                            ExcelRangeBase cell_data = ws.Cells[int.Parse(str_location[3].Split(';')[1]), int.Parse(str_location[1].Split(';')[3]) + 1];

                            if (cpn != "")
                            {
                                DataTable Data_tbl = Data_all.AsEnumerable().Where(s => s.Field<string>("Region") == cpn).CopyToDataTable();
                                if (Data_tbl.Rows.Count > 0)
                                {

                                    Export_DatatableImage_Excel_tape(Data_tbl, "Image", ws, cell_Pic, false, count_sample);
                                    Export_DatatableImage_Excel_Graph(Data_tbl, "Graph", ws, cell_graph, false, count_sample);

                                    List<string> lst_data = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                                    List<string> lst_judge = new List<string> { };
                                    for (int i = 0; i < count_sample; i++)
                                    {
                                        if (i < Data_tbl.Rows.Count)
                                        {
                                            cell_data.Offset(0, i).Value = double.Parse(lst_data[i].Split('_')[0].Replace("Max:", ""));
                                            cell_data.Offset(0, i).Style.Numberformat.Format = "0.00";
                                            cell_data.Offset(1, i).Value = double.Parse(lst_data[i].Split('_')[1].Replace("Average:", ""));
                                            cell_data.Offset(1, i).Style.Numberformat.Format = "0.00";


                                        }
                                    }

                                    for (int i = 1; i < 13; i++)
                                    {
                                        if (myCode.checkDBNull(cell_data.Offset(i, -1).Value).Replace(" ", "").ToUpper().Contains("MINFORCE"))
                                        {
                                            ExcelRangeBase cell_min = cell_data.Offset(i, 0);
                                            ExcelRangeBase cell_max = cell_data.Offset(i + 1, 0);
                                            ExcelRangeBase cell_ave = cell_data.Offset(i + 2, 0);

                                            int c_offset = count_sample;
                                            cell_min.FormulaR1C1 = "=MIN(R[-" + i.ToString() + "]C[0]:R[-" + i.ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                            cell_max.FormulaR1C1 = "=MAX(R[-" + (i + 1).ToString() + "]C[0]:R[-" + (i + 1).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                            cell_ave.FormulaR1C1 = "=AVERAGE(R[-" + (i + 2).ToString() + "]C[0]:R[-" + (i + 2).ToString() + "]C[" + (c_offset - 1).ToString() + "])";

                                            ExcelRangeBase cell_min1 = cell_data.Offset(i, 2);
                                            ExcelRangeBase cell_max1 = cell_data.Offset(i + 1, 2);
                                            ExcelRangeBase cell_ave1 = cell_data.Offset(i + 2, 2);

                                            cell_min1.FormulaR1C1 = "=MIN(R[-" + (i - 1).ToString() + "]C[-2]:R[-" + (i - 1).ToString() + "]C[" + (c_offset - 3).ToString() + "])";
                                            cell_max1.FormulaR1C1 = "=MAX(R[-" + i.ToString() + "]C[-2]:R[-" + i.ToString() + "]C[" + (c_offset - 3).ToString() + "])";
                                            cell_ave1.FormulaR1C1 = "=AVERAGE(R[-" + (i + 1).ToString() + "]C[-2]:R[-" + (i + 1).ToString() + "]C[" + (c_offset - 3).ToString() + "])";

                                            break;
                                        }
                                    }

                                    for (int i = 0; i < count_sample; i++)
                                    {
                                        if (i < Data_tbl.Rows.Count)
                                        {
                                            int ID = int.Parse(Data_tbl.Rows[i]["ID"].ToString());
                                            if (dic_judgement.ContainsKey(ID - 1))
                                            {
                                                if (dic_judgement[ID - 1] == "FAIL")
                                                {
                                                    lst_judge.Add("Fail");

                                                }
                                                else
                                                {
                                                    lst_judge.Add("Pass");
                                                }
                                            }

                                        }
                                    }

                                    for (int t = 1; t < 5; t++)
                                    {
                                        if (myCode.checkDBNull(cell_data.Offset(t, -1).Value).ToUpper().Contains("JUDGEMENT") && !myCode.checkDBNull(cell_data.Offset(t, -1).Value).ToUpper().Replace(" ", "").Contains("FAILUREMODE"))
                                        {
                                            for (int i = 0; i < lst_judge.Count; i++)
                                            {
                                                cell_data.Offset(t, i).Value = lst_judge[i];

                                            }
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {

                    export_ok = false;
                    MessageBox.Show(new Form { TopMost = true }, sheet + " chứa dữ liệu NG. Không thể xuất dữ liệu vào báo cáo", "Warning");
                }

            }
            // }
        }
        public void Export_DatatableImage_Excel_Graph(DataTable dt, string Image_Col_name, ExcelWorksheet tar_wrksht, ExcelRangeBase sel_rgn, bool row_offset, int count_sample)
        {
            string file_folder = System.IO.Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = System.IO.Path.Combine(file_folder, DateTime.Now.ToString("ddMMMyy_HHmmss") + ".jpg");
            int r_inx = 0;
            //ExcelRangeBase sel_rgn = tar_wrksht.Range[tar_rgn];
            ExcelRangeBase rgn_begin = sel_rgn;
            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (count < count_sample)
                {
                    byte[] _img_byte = (byte[])(dr[Image_Col_name]);
                    byte[] img_byte = ImageCompress.CompressImage(_img_byte);
                    InsertPicture_Name(tar_wrksht, sel_rgn, img_byte, sel_rgn.Address);

                    int col_offset = sel_rgn.Columns;
                    int row_off = sel_rgn.Rows;
                    if (row_offset)
                    {
                        sel_rgn = sel_rgn.Offset(row_off, 0);
                    }
                    else
                    {
                        sel_rgn = sel_rgn.Offset(0, col_offset);
                    }
                    r_inx++;
                }
                count++;
            }

        }

        //export peel,pull, shear

        public void export_excel_peel_pull_shear_NPI(ExcelWorksheet ws, string sheet, DataTable Data_tbl, DataTable dt_spec, ref bool export_ok)
        {

            int reg = 0;
            List<DataTable> lst_Table = new List<DataTable> { };
            Get_ListTable(-1, Data_tbl, new string[] { "Region" }, ref lst_Table, "Data");
            string[] region_data = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();

            Dictionary<string, string> dic_mode = new Dictionary<string, string> { };
            dic_mode.Add("Mode 1#", "Mode 1: Solder joint crack");
            dic_mode.Add("Mode 2#", "Mode 2: Pad lift");
            dic_mode.Add("Mode 3#", "Mode 3: Solder joint lift");
            dic_mode.Add("Mode 4#", "Mode 4: Intermetallic break");
            dic_mode.Add("Mode 5#", "Mode 5: Component damage");
            dic_mode.Add("Mode 6#", "Mode 6: Component detached");
            dic_mode.Add("Mode 7#", "Mode 7: Flex torn");


            SortedDictionary<int, string> dic_judgement = new SortedDictionary<int, string> { };
            if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST")
            {
                dic_judgement = Check_spec_Peel_Pull(sheet, Data_tbl, dt_spec);
            }
            else if (sheet == "SHEAR_TEST")
            {
                dic_judgement = Check_spec_ShearTest_new(Data_tbl, dt_spec);
            }
            bool check_all = true;
            foreach (var result in dic_judgement.Values)
            {
                if (result == "FAIL")
                {
                    check_all = false;
                    break;
                }
            }

            if (check_all)
            {
                int count_sample = int.Parse(dt_spec.Rows[0]["Count_sample"].ToString());
                foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                {
                    if (spec_region != "")
                    {
                        if (reg < lst_Table.Count)
                        {
                            DataTable dt_region = lst_Table[reg];

                            string[] str_location = spec_region.Split('+');
                            ExcelRangeBase cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[1]), int.Parse(str_location[1].Split(';')[2]) + 1];
                            ExcelRangeBase cell_graph = ws.Cells[int.Parse(str_location[2].Split(';')[1]), int.Parse(str_location[1].Split(';')[2]) + 1];
                            ExcelRangeBase cell_data = ws.Cells[int.Parse(str_location[3].Split(';')[1]), int.Parse(str_location[1].Split(';')[2]) + 1];
                            //int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());

                            Export_DatatableImage_Excel(dt_region, "Image", ws, cell_Pic, false, count_sample);
                            Export_DatatableImage_Excel_Graph(dt_region, "Graph", ws, cell_graph, false, count_sample);

                            List<string> lst_data = dt_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                            List<string> lst_judge = new List<string> { };


                            for (int i = 0; i < count_sample; i++)
                            {
                                if (i < dt_region.Rows.Count)
                                {

                                    //cell_data.Offset(0, i).Value = lst_data[i];
                                    cell_data.Offset(0, i).Value = Math.Round(double.Parse(lst_data[i]), 2);
                                    cell_data.Offset(0, i).Style.Numberformat.Format = "0.00";
                                    // cell_data.Calculate();
                                    int ID = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                    if (dic_judgement.ContainsKey(ID - 1))
                                    {
                                        if (dic_judgement[ID - 1] == "FAIL")
                                        {
                                            lst_judge.Add("Fail");

                                        }
                                        else
                                        {
                                            lst_judge.Add("Pass");
                                        }
                                    }
                                }
                            }

                            for (int i = 1; i < 13; i++)
                            {
                                if (myCode.checkDBNull(cell_data.Offset(i, -1).Value).Replace(" ", "").ToUpper().Contains("MINFORCE"))
                                {
                                    ExcelRangeBase cell_min = cell_data.Offset(i, 0);
                                    ExcelRangeBase cell_max = cell_data.Offset(i + 1, 0);
                                    ExcelRangeBase cell_ave = cell_data.Offset(i + 2, 0);
                                    int c_offset = count_sample;
                                    cell_min.FormulaR1C1 = "=MIN(R[-" + i.ToString() + "]C[0]:R[-" + i.ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                    cell_max.FormulaR1C1 = "=MAX(R[-" + (i + 1).ToString() + "]C[0]:R[-" + (i + 1).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                    cell_ave.FormulaR1C1 = "=AVERAGE(R[-" + (i + 2).ToString() + "]C[0]:R[-" + (i + 2).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                    //ws.Workbook.CalcMode = ExcelCalcMode.Automatic;  
                                    //cell_min.Formula = "=MIN(" + cell_data.Address + ":" + cell_data.Offset(0, c_offset - 1).Address + ")";


                                    ////  cell_min.Style.Numberformat.Format = "0.00";
                                    //cell_max.Formula = "=MAX(" + cell_data.Address + ":" + cell_data.Offset(0, c_offset - 1).Address + ")";
                                    //cell_ave.Formula = "=AVERAGE(" + cell_data.Address + ":" + cell_data.Offset(0, c_offset - 1).Address + ")";
                                    break;
                                }
                            }

                            int k1 = 1;
                            int k2 = 7;
                            string type = dt_spec.Rows[0]["Location"].ToString().Split('+')[0];

                            if (sheet == "SHEAR_TEST" && type == "A")
                            {
                                for (int i = 0; i < count_sample; i++)
                                {
                                    if (i < dt_region.Rows.Count)
                                    {
                                        cell_data.Offset(1, i).Value = double.Parse(lst_data[i]) * 9.81;
                                        cell_data.Offset(1, i).Style.Numberformat.Format = "0.00";
                                    }
                                }

                                k1 = 2;
                                k2 = 8;

                                for (int i = k1; i <= k2; i++)
                                {
                                    List<string> lst_mode = dt_region.AsEnumerable().Select(x => x.Field<string>(dic_mode["Mode " + (i - 1).ToString() + "#"])).ToList();
                                    for (int j = 0; j < count_sample; j++)
                                    {
                                        if (j < lst_mode.Count)
                                        {
                                            if (lst_mode[j].Contains("("))
                                            {
                                                cell_data.Offset(i, j).FormulaR1C1 = "=" + lst_mode[j].Split('(')[1].Split(')')[0];

                                            }
                                            else
                                            {
                                                cell_data.Offset(i, j).Value = lst_mode[j];
                                            }
                                            cell_data.Offset(i, j).Style.Numberformat.Format = "0%";
                                        }
                                    }
                                }
                                for (int i = 0; i < count_sample; i++)
                                {
                                    cell_data.Offset(9, i).Value = lst_judge[i];
                                }

                            }
                            else
                            {
                                for (int i = k1; i <= k2; i++)
                                {
                                    List<string> lst_mode = dt_region.AsEnumerable().Select(x => x.Field<string>(dic_mode["Mode " + i.ToString() + "#"])).ToList();
                                    for (int j = 0; j < count_sample; j++)
                                    {
                                        if (j < lst_mode.Count)
                                        {
                                            if (lst_mode[j].Contains("("))
                                            {
                                                cell_data.Offset(i, j).FormulaR1C1 = "=" + lst_mode[j].Split('(')[1].Split(')')[0];

                                            }
                                            else
                                            {
                                                cell_data.Offset(i, j).Value = lst_mode[j];

                                                // cell_data.Offset(i, j).FormulaR1C1 = "='" + lst_mode[j];

                                            }
                                            cell_data.Offset(i, j).Style.Numberformat.Format = "0%";
                                        }
                                    }
                                }

                                for (int i = 0; i < Math.Min(count_sample, lst_judge.Count); i++)
                                {
                                    cell_data.Offset(8, i).Value = lst_judge[i];
                                }
                            }


                        }
                        reg++;
                    }

                }
            }
            else
            {
                export_ok = false;
                MessageBox.Show(new Form { TopMost = true }, sheet + " chứa dữ liệu NG. Không thể xuất dữ liệu vào báo cáo", "Warning");
            }
        }

        public string get_number_spec2(string str_in)
        {
            List<char> lst_csplit = new List<char>() { '>', '<', '≥', '≤' };
            string v = "";
            foreach (char c in lst_csplit)
            {
                if (str_in.Contains(c.ToString()))
                {
                    v = str_in.Split(c)[1];
                    break;
                }

            }
            return v;
        }

        public string formula_judgement_peelpull(ExcelRangeBase cur_rgn_judge, DataTable dt_spec)
        {
            string str_judge = "=IF(OR(D4=\"\"),\"Miss\",(IF(OR(D14<$F$25,D14>$F$24,(MAX($D$14:$F$14)-MIN($D$14:$F$14))>$F$23),\"Fail\",\"Pass\")))";


            return str_judge;

        }


        public void export_excel_peel_pull_shear(ExcelWorksheet ws, string itemcode, string lotno, string sheet, DataTable Data_tbl, DataTable dt_spec, DataGridView dgv_data, string type, string leader, int count_sample, bool prime = false)
        {

            string testName = "Force!!!!!!!!!!!!!!!", forceName = "Froezzzz", caching = "crack";
            if (ws.Name.Contains("Peel"))
            {
                testName = "Peeling Force";
                forceName = "Judgement peeling force";
            }
            if (ws.Name.Contains("Pull"))
            {
                testName = "Pulling Force";
                forceName = "Judgement pulling force";
            }
            if (ws.Name.Contains("Shear"))
            {
                testName = "Shear Force (Kgf)";
                forceName = "Judgement Shear force";
                caching = "shear";
            }


            string[] nameAddress = new[] { "Flex SN", testName, forceName, "Sample", "STDEV", "CPK", "Min Force (N)", "Max Force (N)", "Average Force (N)", "Picture", "Graph", "Final judgement", "Judgement failure mode", $"Solder joint {caching}", "Pad lift", "Solder joint lift", "Intermetallic break", "Component damage", "Component detached", "Flex torn" };
            //Find Dictionary have name column {sample 1} and address
            IDictionary<string, string> addressDic = ExportProcess.FindAddressByText(ws, nameAddress);

            string addressForce = "";
            if (addressDic.TryGetValue("Sample", out string address))
            {
                if (address.Split('-').Count() < Data_tbl.Rows.Count)
                {
                    int z = Data_tbl.Rows.Count - address.Split('-').Count();
                    string a = address.Split('-')[address.Split('-').Count() - 1];
                    for (int c = 0; c < z; c++)
                    {
                        string aStart = ExportProcess.AddColumn(ExportProcess.AddRow(a, -1), c);
                        string aEnd = ExportProcess.AddColumn(ExportProcess.AddRow(a, 19), c);
                        ExportProcess.CopyColumn(ws, ws.Cells[$"{aStart}:{aEnd}"], ExportProcess.AddRow(ExportProcess.AddColumn(a, c + 1), -1));
                        ws.Cells[ExportProcess.AddColumn(a, c + 1)].Value = $"Sample {count_sample + c + 1}";
                        addressDic["Sample"] = $"{addressDic["Sample"]}-{ExportProcess.AddColumn(a, c + 1)}";
                    }
                    address = addressDic["Sample"];
                }
                if (count_sample != address.Split('-').Count())
                {
                    if (MessageBox.Show("Số pcs của format không phù hợp bạn có muốn tiếp tục", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.No)
                    {
                        return;
                    }

                }

            }

            try
            {
                int rowIndex = -1;
                IList<double> numbers = new List<double>();
                foreach (string addressCol in address.Split('-'))
                {
                    rowIndex++;
                    int column = ws.Cells[addressCol].End.Column;
                    ///
                    DataRow item = Data_tbl.Rows[rowIndex];
                    if (addressDic.TryGetValue("Flex SN", out string adz))
                    {
                        adz = ws.Cells[ws.Cells[adz].Start.Row, ws.Cells[addressCol].Start.Column].Address;
                        if (Data_tbl.Columns.Contains("ProductID"))
                        {
                            ws.Cells[adz].Value = item["ProductID"];
                        }
                    }
                    if (addressDic.TryGetValue("Picture", out string addressRow))
                    {
                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        if (item["Image"] != DBNull.Value && item["Image"] is byte[])
                        {
                            InsertPicture_Name(ws, ws.Cells[address], (byte[])item["Image"], $"{rowIndex} - picture");
                        }
                    }
                    if (addressDic.TryGetValue("Graph", out addressRow))
                    {
                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        if (item["Graph"] != DBNull.Value && item["Graph"] is byte[])
                        {
                            InsertPicture_Name(ws, ws.Cells[address], (byte[])item["Graph"], $"{rowIndex} - Graph");
                        }
                    }
                    if (addressDic.TryGetValue(testName, out addressRow))
                    {
                        addressRow = addressRow.Split('-')[0];
                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        int z = testName.Contains("Shear") ? 1 : 0;
                        addressForce += $" {ExportProcess.AddRow(address, z)} + ";
                        ws.Cells[address].Value = double.Parse(item["Data"].ToString().Trim());
                        if (testName.Contains("Shear"))
                        {
                            ws.Cells[ExportProcess.AddRow(address, 1)].FormulaR1C1 = $" =R[-1]C*9.8";
                        }
                    }


                    string primeMode = "";
                    double maxMode = double.MinValue;
                    int pin = 0;
                    try
                    {
                        string pinZ = item["Mode 1: Solder joint crack"].ToString().Split('%')[1].Split('/')[1].Replace(")", "");
                        pin = int.Parse(pinZ);
                    }
                    catch
                    {
                       
                    }
                    if (addressDic.TryGetValue($"Solder joint {caching}", out addressRow))
                    {
                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        string[] content = item["Mode 1: Solder joint crack"].ToString().Split('%');
                        if (content.Count() > 1 && content[1].Contains("("))
                        {
                            ws.Cells[address].Formula = $"={content[1]}";

                            if (double.TryParse(content[1].Split('/')[0].Replace("(", ""), out double number))
                            {
                                if (number > maxMode)
                                {
                                    maxMode = number;
                                    primeMode = "Mode #1";
                                }
                            }
                        }
                        else
                        {
                            ws.Cells[address].Formula = $"=0/{pin}";
                        }
                    }
                    if (addressDic.TryGetValue("Pad lift", out addressRow))
                    {
                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        string[] content = item["Mode 2: Pad lift"].ToString().Split('%');
                        if (content.Count() > 1 && content[1].Contains("("))
                        {
                            ws.Cells[address].Formula = $"={content[1]}";

                            if (double.TryParse(content[1].Split('/')[0].Replace("(", ""), out double number))
                            {
                                if (number > maxMode)
                                {
                                    maxMode = number;
                                    primeMode = "Mode #2";
                                }
                            }
                        }
                        else
                        {
                            ws.Cells[address].Formula = $"=0/{pin}";
                        }
                    }
                    if (addressDic.TryGetValue("Solder joint lift", out addressRow))
                    {

                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        string[] content = item["Mode 3: Solder joint lift"].ToString().Split('%');
                        if (content.Count() > 1 && content[1].Contains("("))
                        {
                            ws.Cells[address].Formula = $"={content[1]}";

                            if (double.TryParse(content[1].Split('/')[0].Replace("(", ""), out double number))
                            {
                                if (number > maxMode)
                                {
                                    maxMode = number;
                                    primeMode = "Mode #3";
                                }
                            }
                        }
                        else
                        {
                            ws.Cells[address].Formula = $"=0/{pin}";
                        }
                    }
                    if (addressDic.TryGetValue("Intermetallic break", out addressRow))
                    {

                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        string[] content = item["Mode 4: Intermetallic break"].ToString().Split('%');
                        if (content.Count() > 1 && content[1].Contains("("))
                        {
                            ws.Cells[address].Formula = $"={content[1]}";

                            if (double.TryParse(content[1].Split('/')[0].Replace("(", ""), out double number))
                            {
                                if (number > maxMode)
                                {
                                    maxMode = number;
                                    primeMode = "Mode #4";
                                }
                            }
                        }
                        else
                        {
                            ws.Cells[address].Formula = $"=0/{pin}";
                        }
                    }
                    if (addressDic.TryGetValue("Component damage", out addressRow))
                    {
                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        string[] content = item["Mode 5: Component damage"].ToString().Split('%');
                        if (content.Count() > 1 && content[1].Contains("("))
                        {
                            ws.Cells[address].Formula = $"={content[1]}";

                            if (double.TryParse(content[1].Split('/')[0].Replace("(", ""), out double number))
                            {
                                if (number > maxMode)
                                {
                                    maxMode = number;
                                    primeMode = "Mode #5";
                                }
                            }
                        }
                        else
                        {
                            ws.Cells[address].Formula = $"=0/{pin}";
                        }
                    }
                    if (addressDic.TryGetValue("Component detached", out addressRow))
                    {

                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        string[] content = item["Mode 6: Component detached"].ToString().Split('%');
                        if (content.Count() > 1 && content[1].Contains("("))
                        {
                            ws.Cells[address].Formula = $"={content[1]}";

                            if (double.TryParse(content[1].Split('/')[0].Replace("(", ""), out double number))
                            {
                                if (number > maxMode)
                                {
                                    maxMode = number;
                                    primeMode = "Mode #6";
                                }
                            }
                        }
                        else
                        {
                            ws.Cells[address].Formula = $"=0/{pin}";
                        }
                    }
                    if (addressDic.TryGetValue("Flex torn", out addressRow))
                    {
                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        string[] content = item["Mode 7: Flex torn"].ToString().Split('%');
                        if (content.Count() > 1 && content[1].Contains("("))
                        {
                            ws.Cells[address].Formula = $"={content[1]}";

                            if (double.TryParse(content[1].Split('/')[0].Replace("(", ""), out double number))
                            {
                                if (number > maxMode)
                                {
                                    maxMode = number;
                                    primeMode = "Mode #7";
                                }
                            }
                        }
                        else
                        {
                            ws.Cells[address].Formula = $"=0/{pin}";
                        }
                    }
                    //Debugger.Break();
                    double valueForce = double.MinValue;

                    if (addressDic.TryGetValue(forceName, out addressRow))
                    {
                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;

                        if (double.TryParse(item["Data"].ToString(), out valueForce))
                        {
                            if (dt_spec.Rows.Count > 0)
                            {
                                DataRow row = dt_spec.Rows[0];
                                string specLow = row["Location"].ToString().Replace("N", "").Replace(">", "").Trim();
                                if (double.TryParse(specLow, out double num))
                                {
                                    if (addressDic.TryGetValue(testName, out string addressRowZ))
                                    {

                                        int r = ws.Cells[address].Start.Row - ws.Cells[addressRowZ].Start.Row;
                                        ws.Cells[address].FormulaR1C1 = $"=IF(R[{r * -1}]C>{num}, \"Pass\",\"Fail\" )";
                                    }

                                }
                            }
                            else
                            {
                                ws.Cells[address].Value = valueForce > 5 ? "Pass" : "NG";
                            }
                        }
                        numbers.Add(valueForce);
                    }
                    if (addressDic.TryGetValue("Judgement failure mode", out addressRow))
                    {
                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        //ws.Cells[address].Value = primeMode;
                        if (addressDic.TryGetValue("Flex torn", out string addMode7))
                        {
                            string addressMode7 = ws.Cells[ws.Cells[addMode7].End.Row, column].Address;
                            if (addressDic.TryGetValue($"Solder joint {caching}", out string addMode1))
                            {
                                string addressMode1 = ws.Cells[ws.Cells[addMode1].End.Row, column].Address;

                                addMode1 = ExportProcess.AddColumn(addMode1, -1);
                                addMode7 = ExportProcess.AddColumn(addMode7, -1);
                                int r1 = ws.Cells[addMode1].End.Row - ws.Cells[address].End.Row;
                                int c1 = ws.Cells[addMode1].End.Column - ws.Cells[address].End.Column;
                                int r7 = ws.Cells[addMode7].End.Row - ws.Cells[address].End.Row;
                                int c7 = ws.Cells[addMode7].End.Column - ws.Cells[address].End.Column;


                                int row1 = ws.Cells[addressMode1].End.Row - ws.Cells[address].End.Row;
                                int col1 = ws.Cells[addressMode1].End.Column - ws.Cells[address].End.Column;
                                int row7 = ws.Cells[addressMode7].End.Row - ws.Cells[address].End.Row;
                                int col7 = ws.Cells[addressMode7].End.Column - ws.Cells[address].End.Column;

                                ws.Cells[address].FormulaR1C1 = $"=INDEX(R[{r1}]C[{c1}]:R[{r7}]C[{c7}],MATCH(MAX(R[{row1}]C[{col1}]:R[{row7}]C[{col7}]),R[{row1}]C[{col1}]:R[{row7}]C[{col7}],0))";
                            }
                        }
                    }
                    if (addressDic.TryGetValue("Final judgement", out addressRow))
                    {
                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        if (addressDic.TryGetValue("Intermetallic break", out string addressRowZ))
                        {
                            int r = ws.Cells[addressRowZ].End.Row - ws.Cells[addressRow].End.Row;
                            ws.Cells[address].FormulaR1C1 = $"=IF(R[{r}]C > 0, \"SEM/EDX analysis 1 piece\", \"Pass\")";
                        }

                        //if (valueForce > 5 && (primeMode.Contains("5") || primeMode.Contains("2")))
                        //{
                        //    ws.Cells[address].Value = "Pass";
                        //}
                        //else
                        //{
                        //    ws.Cells[address].Value = "Fail";

                        //}

                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            if (addressDic.TryGetValue("Max Force (N)", out string addressRowz))
            {
                address = ws.Cells[ws.Cells[addressRowz].End.Row, ws.Cells[addressRowz].End.Column + 2].Address;
                ws.Cells[address].FormulaR1C1 = $"=MAX({ExportProcess.ConvertAddressRangeBase(ws, addressForce, address)})";
            }
            if (addressDic.TryGetValue("Min Force (N)", out addressRowz))
            {
                address = ws.Cells[ws.Cells[addressRowz].End.Row, ws.Cells[addressRowz].End.Column + 2].Address;
                ws.Cells[address].FormulaR1C1 = $"=MIN({ExportProcess.ConvertAddressRangeBase(ws, addressForce, address)})";
            }
            if (addressDic.TryGetValue("Average Force (N)", out string addressAverage))
            {
                addressAverage = ws.Cells[ws.Cells[addressAverage].End.Row, ws.Cells[addressAverage].End.Column + 2].Address;
                ws.Cells[addressAverage].FormulaR1C1 = $"=average({ExportProcess.ConvertAddressRangeBase(ws, addressForce, addressAverage)})";
            }

            if (addressDic.TryGetValue("STDEV", out string addressSTDEV))
            {
                addressSTDEV = ws.Cells[ws.Cells[addressSTDEV].End.Row, ws.Cells[addressRowz].End.Column + 2].Address;
                ws.Cells[addressSTDEV].FormulaR1C1 = $"=STDEV({ExportProcess.ConvertAddressRangeBase(ws, addressForce, addressSTDEV)})";
                if (addressDic.TryGetValue("CPK", out addressRowz))
                {
                    address = ws.Cells[ws.Cells[addressRowz].End.Row, ws.Cells[addressRowz].End.Column + 2].Address;
                    ws.Cells[address].FormulaR1C1 = $"=({ExportProcess.ConvertAddressRangeBase(ws, addressAverage, address)}-{ConverterService.GetNumberFromString(dt_spec.Rows[0]["Location"].ToString())})/(3*{ExportProcess.ConvertAddressRangeBase(ws, addressSTDEV, address)})";
                }
            }

            if (prime)
            {
                ws.Name = "Peel Test without SUS";
            }
            #region OldCode
            //bool judge_all = true;
            //// Type mass and other
            //if (type.Contains("MASS") || type == "Other")
            //{
            //    ExcelRangeBase cur_rgn_type = find_cell(ws, type.Replace("MASS(", "").Replace(")", ""));
            //    if (cur_rgn_type != null)
            //        cur_rgn_type.Value = " v" + cur_rgn_type.Value.ToString();

            //    string str_infor = Data_tbl.Rows[0]["Sheet"].ToString().Replace(sheet, "").Replace("/", "");
            //    if (str_infor.Contains("_"))
            //    {
            //        string itemname = str_infor.Split('_')[0];
            //        string line = str_infor.Split('_')[1];
            //        string ca = str_infor.Split('_')[2];
            //        string date = str_infor.Split('_')[3];
            //        string worker = str_infor.Split('_')[4];
            //        List<string> lst_infor = new List<string> { itemname, itemcode, lotno, line, ca, date, worker };
            //        export_info_mass(ws, lst_infor, sheet, leader);
            //    }

            //}

            //int reg = 0;
            //List<DataTable> lst_Table = new List<DataTable> { };
            //Get_ListTable(-1, Data_tbl, new string[] { "Region" }, ref lst_Table, "Data");
            //string[] region_data = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();

            //Dictionary<string, string> dic_mode = new Dictionary<string, string> { };
            //dic_mode.Add("Mode 1#", "Mode 1: Solder joint crack");
            //dic_mode.Add("Mode 2#", "Mode 2: Pad lift");
            //dic_mode.Add("Mode 3#", "Mode 3: Solder joint lift");
            //dic_mode.Add("Mode 4#", "Mode 4: Intermetallic break");
            //dic_mode.Add("Mode 5#", "Mode 5: Component damage");
            //dic_mode.Add("Mode 6#", "Mode 6: Component detached");
            //dic_mode.Add("Mode 7#", "Mode 7: Flex torn");

            //int sochan = 0;
            //List<string> arr_str = Data_tbl.AsEnumerable().Select<DataRow, string>((System.Func<DataRow, string>)(x => x.Field<string>(dic_mode["Mode 1#"]))).ToList<string>();
            //foreach (string i in arr_str)
            //{
            //    if (i.Contains("("))
            //    {
            //        if (myCode.IsNumeric(i.Split('/')[1].Split(')')[0]))
            //            sochan = int.Parse(i.Split('/')[1].Split(')')[0]);

            //        break;
            //    }
            //} 

            //int r_end;
            //if (lst_Table.Count == 1)
            //{
            //    r_end = 22;
            //}
            //else
            //{
            //    r_end = 42;
            //}
            //insert_columns_other(ws, count_sample, r_end, type, dt_spec, sheet);
            //foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
            //{
            //    if (spec_region != "")
            //    {
            //        if (reg < lst_Table.Count)
            //        {
            //            DataTable dt_region = lst_Table[reg];

            //            string[] str_location = spec_region.Split('+');
            //            ExcelRangeBase cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[1]), int.Parse(str_location[1].Split(';')[2]) + 1];
            //            ExcelRangeBase cell_graph = ws.Cells[int.Parse(str_location[2].Split(';')[1]), int.Parse(str_location[1].Split(';')[2]) + 1];
            //            ExcelRangeBase cell_data = ws.Cells[int.Parse(str_location[3].Split(';')[1]), int.Parse(str_location[1].Split(';')[2]) + 1];

            //            // string spec = get_number_spec2(spec_region.Split('+')[3].Split(';')[0].Split('(')[1].Split(')')[0].Replace(" ", string.Empty).Replace("N", ""));

            //            Export_DatatableImage_Excel(dt_region, "Image", ws, cell_Pic, false, count_sample);
            //            Export_DatatableImage_Excel_Graph(dt_region, "Graph", ws, cell_graph, false, count_sample);

            //            List<string> lst_data = dt_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
            //            List<string> lst_judge = new List<string> { };

            //            for (int i = 0; i < count_sample; i++)
            //            {
            //                if (i < dt_region.Rows.Count)
            //                {
            //                    //cell_data.Offset[0, i].Value = lst_data[i];
            //                    if (myCode.IsNumeric(lst_data[i]))
            //                    {
            //                        cell_data.Offset(0, i).Value = Math.Round(double.Parse(lst_data[i]), 2);
            //                        cell_data.Offset(0, i).Style.Numberformat.Format = "0.00";
            //                    }

            //                    int ID = int.Parse(dt_region.Rows[i]["ID"].ToString());
            //                    if (dgv_data.Rows[ID - 1].Cells["Data"].Style.BackColor == Color.Red)
            //                    {
            //                        lst_judge.Add("Fail");
            //                        judge_all = false;
            //                    }
            //                    else
            //                    {
            //                        lst_judge.Add("Pass");
            //                    }
            //                }
            //            }
            //            for (int i = 1; i < 13; i++)
            //            {
            //                if (myCode.checkDBNull(cell_data.Offset(i, -1).Value).Replace(" ", "").ToUpper().Contains("MINFORCE"))
            //                {
            //                    ExcelRangeBase cell_min = cell_data.Offset(i, 0);
            //                    ExcelRangeBase cell_max = cell_data.Offset(i + 1, 0);
            //                    ExcelRangeBase cell_ave = cell_data.Offset(i + 2, 0);
            //                    int c_offset = count_sample;
            //                    cell_min.FormulaR1C1 = "=MIN(R[-" + i.ToString() + "]C[0]:R[-" + i.ToString() + "]C[" + (c_offset - 1).ToString() + "])";
            //                    cell_max.FormulaR1C1 = "=MAX(R[-" + (i + 1).ToString() + "]C[0]:R[-" + (i + 1).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
            //                    cell_ave.FormulaR1C1 = "=AVERAGE(R[-" + (i + 2).ToString() + "]C[0]:R[-" + (i + 2).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
            //                    cell_min.Offset(0, i).Style.Numberformat.Format = "0.00";
            //                    cell_max.Offset(0, i).Style.Numberformat.Format = "0.00";
            //                    cell_ave.Offset(0, i).Style.Numberformat.Format = "0.00";

            //                    break;
            //                }
            //            }

            //            int k1 = 1;
            //            int k2 = 7;
            //            string typeAB = dt_spec.Rows[0]["Location"].ToString().Split('+')[0];


            //            if (sheet == "SHEAR_TEST" && typeAB == "A")
            //            {
            //                for (int i = 0; i < count_sample; i++)
            //                {
            //                    if (i < dt_region.Rows.Count)
            //                    {
            //                        cell_data.Offset(1, i).Value = double.Parse(lst_data[i]) * 9.81;
            //                        cell_data.Offset(1, i).Style.Numberformat.Format = "0.00";
            //                    }
            //                }

            //                k1 = 2;
            //                k2 = 8;

            //                for (int i = k1; i <= k2; i++)
            //                {
            //                    List<string> lst_mode = dt_region.AsEnumerable().Select(x => x.Field<string>(dic_mode["Mode " + (i - 1).ToString() + "#"])).ToList();
            //                    for (int j = 0; j < count_sample; j++)
            //                    {
            //                        if (j < lst_mode.Count)
            //                        {
            //                            if (lst_mode[j].Contains("(") && sochan != 0)
            //                            {
            //                                cell_data.Offset(i, j).FormulaR1C1 = "=" + lst_mode[j].Split('(')[1].Split(')')[0];

            //                            }
            //                            else if (lst_mode[j] == "0.00%" && sochan != 0)
            //                            {
            //                                cell_data.Offset(i, j).FormulaR1C1 = "=0/" + sochan.ToString();
            //                            }
            //                            else if (lst_mode[j] == "100%" && sochan != 0)
            //                            {
            //                                cell_data.Offset(i, j).FormulaR1C1 = "=" + sochan + "/" + sochan.ToString();
            //                            }
            //                            else
            //                            {
            //                                cell_data.Offset(i, j).Value = lst_mode[j];
            //                            }
            //                            cell_data.Offset(i, j).Style.Numberformat.Format = "0%";
            //                        }
            //                    }
            //                }
            //                for (int i = 0; i < count_sample; i++)
            //                {
            //                    cell_data.Offset(9, i).Value = lst_judge[i];
            //                    //if (myCode.IsNumeric(spec))
            //                    //{
            //                    //    cell_data.Offset(9, i).FormulaR1C1 = "=IF(" + cell_data.Offset(1, i).Address + "<" + spec + ",\"Fail\",\"Pass\")";
            //                    //}
            //                    //else
            //                    //{
            //                    //    cell_data.Offset(9, i).Value = "Pass";
            //                    // }
            //                }

            //            }
            //            else
            //            {
            //                for (int i = k1; i <= k2; i++)
            //                {
            //                    List<string> lst_mode = dt_region.AsEnumerable().Select(x => x.Field<string>(dic_mode["Mode " + i.ToString() + "#"])).ToList();
            //                    for (int j = 0; j < count_sample; j++)
            //                    {
            //                        if (j < lst_mode.Count)
            //                        {
            //                            if (lst_mode[j].Contains("(") && sochan != 0)
            //                            {
            //                                cell_data.Offset(i, j).FormulaR1C1 = "=" + lst_mode[j].Split('(')[1].Split(')')[0];

            //                            }
            //                            else if (lst_mode[j] == "0.00%" && sochan != 0)
            //                            {
            //                                cell_data.Offset(i, j).FormulaR1C1 = "=0/" + sochan.ToString();
            //                            }
            //                            else if (lst_mode[j] == "100%" && sochan != 0)
            //                            {
            //                                cell_data.Offset(i, j).FormulaR1C1 = "=" + sochan + "/" + sochan.ToString();
            //                            }
            //                            else
            //                            {
            //                                cell_data.Offset(i, j).Value = lst_mode[j];
            //                            }
            //                            cell_data.Offset(i, j).Style.Numberformat.Format = "0%";

            //                        }
            //                    }
            //                }

            //                //string addr_R = cell_data.Offset(9, 2).Address;
            //                //string addr_UCL = cell_data.Offset(10, 2).Address;
            //                //string addr_LCL = cell_data.Offset(11, 2).Address;
            //                //string addr_celldata_begin = cell_data.Address;
            //                //string addr_celldata_end = cell_data.Offset(0, Math.Min(count_sample, lst_judge.Count) - 1).Address;


            //                for (int i = 0; i < Math.Min(count_sample, lst_judge.Count); i++)
            //                {
            //                    cell_data.Offset(8, i).Value = lst_judge[i];
            //                    //string addr_cell = cell_data.Offset(8, i).Address;
            //                    //string addr_currcelldata = cell_data.Offset(0 ,i).Address;
            //                    //string s = "";
            //                    //string a = "=IF(OR(" + addr_currcelldata + "=\"\"),\"Miss\",(IF(OR(AND(NOT(" + addr_LCL + " =\"\"), " + addr_currcelldata  + "<" + addr_LCL + "), AND(NOT(" + addr_UCL + " =\"\"), " + addr_currcelldata + ">" + addr_UCL + "),AND(NOT(" + addr_R + " =\"\")," + "(MAX(" + addr_celldata_begin + ":" + addr_celldata_end + ")-MIN(" + addr_celldata_begin + ":" + addr_celldata_end + "))" + ">" + addr_R + ")),\"Fail\",\"Pass\")))";
            //                    //cell_data.Offset(8, i).FormulaR1C1 = a;
            //                }
            //            }
            //        }
            //        reg++;
            //    }

            //}
            //int r_offset_del = 14;
            //complete_sheet_other(ws, count_sample, r_offset_del, sheet, dt_spec);


            //if (type.Contains("MASS"))
            //{
            //    ExcelRangeBase cur_rgn_judge = find_cell(ws, "JUDGEMENT");
            //    if (cur_rgn_judge != null)
            //    {
            //        if (judge_all)
            //        {
            //            cur_rgn_judge.Offset(1, 0).Value = "OK";
            //            ws.Cells[1, 1].Value = "OK";
            //        }
            //        else
            //        {
            //            cur_rgn_judge.Offset(1, 0).Value = "NG";
            //            ws.Cells[1, 1].Value = "NG";
            //        }
            //    }
            //}
            #endregion
        }
        public static double StandardDeviation(IEnumerable<double> values)
        {
            double average = values.Average();
            double sumOfSquaresOfDifferences = values.Sum(val => Math.Pow(val - average, 2));
            return Math.Round(Math.Sqrt(sumOfSquaresOfDifferences / (values.Count() - 1)), 2); // Sử dụng (n-1) cho mẫu (sample)
        }
        public static double GetDecimalFromString(string input)
        {
            string pattern = @"\d+\.\d+"; // Biểu thức chính quy tìm số thập phân
            Match match = Regex.Match(input, pattern);

            if (match.Success)
            {
                return double.Parse(match.Value);
            }
            else
            {
                return 0; // Hoặc ném ngoại lệ nếu không tìm thấy số thập phân
            }
        }
        public Image resizeImage(int newWidth, int newHeight, string stPhotoPath)
        {
            Image imgPhoto = Image.FromFile(stPhotoPath);

            int sourceWidth = imgPhoto.Width;
            int sourceHeight = imgPhoto.Height;

            //Consider vertical pics
            if (sourceWidth < sourceHeight)
            {
                int buff = newWidth;

                newWidth = newHeight;
                newHeight = buff;
            }

            int sourceX = 0, sourceY = 0, destX = 0, destY = 0;
            float nPercent = 0, nPercentW = 0, nPercentH = 0;

            nPercentW = ((float)newWidth / (float)sourceWidth);
            nPercentH = ((float)newHeight / (float)sourceHeight);
            if (nPercentH < nPercentW)
            {
                nPercent = nPercentH;
                destX = System.Convert.ToInt16((newWidth -
                          (sourceWidth * nPercent)) / 2);
            }
            else
            {
                nPercent = nPercentW;
                destY = System.Convert.ToInt16((newHeight -
                          (sourceHeight * nPercent)) / 2);
            }

            int destWidth = (int)(sourceWidth * nPercent);
            int destHeight = (int)(sourceHeight * nPercent);


            Bitmap bmPhoto = new Bitmap(newWidth, newHeight,
                          PixelFormat.Format24bppRgb);

            bmPhoto.SetResolution(imgPhoto.HorizontalResolution,
                         imgPhoto.VerticalResolution);

            Graphics grPhoto = Graphics.FromImage(bmPhoto);
            grPhoto.Clear(Color.Black);
            grPhoto.InterpolationMode =
                System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

            grPhoto.DrawImage(imgPhoto,
                new System.Drawing.Rectangle(destX, destY, destWidth, destHeight),
                new System.Drawing.Rectangle(sourceX, sourceY, sourceWidth, sourceHeight),
                GraphicsUnit.Pixel);

            grPhoto.Dispose();
            imgPhoto.Dispose();
            return bmPhoto;
        }
        public Image resizeImage(double nPercent, string stPhotoPath)
        {
            Image imgPhoto = Image.FromFile(stPhotoPath);
            int sourceWidth = imgPhoto.Width;
            int sourceHeight = imgPhoto.Height;
            int sourceX = 0, sourceY = 0, destX = 0, destY = 0;
            int destWidth = (int)(sourceWidth * nPercent);
            int destHeight = (int)(sourceHeight * nPercent);
            Bitmap bmPhoto = new Bitmap(destWidth, destHeight,
                          PixelFormat.Format24bppRgb);

            bmPhoto.SetResolution(imgPhoto.HorizontalResolution,
                         imgPhoto.VerticalResolution);

            Graphics grPhoto = Graphics.FromImage(bmPhoto);
            grPhoto.Clear(Color.Black);
            grPhoto.InterpolationMode = InterpolationMode.HighQualityBicubic;

            grPhoto.DrawImage(imgPhoto,
                new System.Drawing.Rectangle(destX, destY, destWidth, destHeight),
                new System.Drawing.Rectangle(sourceX, sourceY, sourceWidth, sourceHeight),
                GraphicsUnit.Pixel);
            grPhoto.Dispose();
            imgPhoto.Dispose();
            return bmPhoto;
        }
        public Image resizeImage(double nPercent, Image src_img)
        {
            int sourceWidth = src_img.Width;
            int sourceHeight = src_img.Height;
            int sourceX = 0, sourceY = 0, destX = 0, destY = 0;
            int destWidth = (int)(sourceWidth * nPercent);
            int destHeight = (int)(sourceHeight * nPercent);
            Bitmap bmPhoto = new Bitmap(destWidth, destHeight, PixelFormat.Format24bppRgb);
            bmPhoto.SetResolution(src_img.HorizontalResolution, src_img.VerticalResolution);
            Graphics grPhoto = Graphics.FromImage(bmPhoto);
            grPhoto.Clear(Color.Black);
            grPhoto.InterpolationMode = InterpolationMode.HighQualityBicubic;
            grPhoto.DrawImage(src_img, new System.Drawing.Rectangle(destX, destY, destWidth, destHeight), new System.Drawing.Rectangle(sourceX, sourceY, sourceWidth, sourceHeight), GraphicsUnit.Pixel);
            grPhoto.Dispose();
            src_img.Dispose();
            return bmPhoto;
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


        // export coupon

        public void export_excel_coupon_NPI(ExcelWorksheet ws, DataTable Data_all, DataTable dt_spec, string mysheet, ref bool export_ok)
        {
            SortedDictionary<int, string> dic_judgement = check_spec_coupon(Data_all, dt_spec, mysheet);
            bool check_all = true;
            foreach (var result in dic_judgement.Values)
            {
                if (result == "FAIL")
                {
                    check_all = false;
                    break;
                }
            }

            if (check_all)
            {
                int count_sample = int.Parse(dt_spec.Rows[0]["Count_sample"].ToString());
                if (dt_spec.Rows.Count > 0)
                {
                    List<string> lst_region = Data_all.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToList();
                    string[] region = dt_spec.Rows[0]["Location"].ToString().Split('_');

                    foreach (string reg in region)
                    {
                        if (reg != "")
                        {
                            string[] str_location = reg.Split('+');
                            string cpn = "";

                            foreach (string r in lst_region)
                            {
                                bool chk = true;
                                foreach (string key in r.Split('_'))
                                {
                                    if (!reg.Replace(" ", "").ToUpper().Contains(key.Replace(" ", "").ToUpper()))
                                    {
                                        chk = false;
                                        break;

                                    }
                                }

                                if (chk)
                                {
                                    cpn = r;
                                    break;
                                }
                            }


                            if (cpn != "")
                            {
                                DataTable dt_region = Data_all.AsEnumerable().Where(s => s.Field<string>("Region") == cpn).CopyToDataTable();
                                if (dt_region.Rows.Count > 0)
                                {

                                    List<string> lst_data = dt_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                                    List<string> lst_judge = new List<string> { };


                                    ExcelRangeBase cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[2]), int.Parse(str_location[1].Split(';')[3]) + 1];
                                    ExcelRangeBase cell_graph = ws.Cells[int.Parse(str_location[2].Split(';')[1]), int.Parse(str_location[1].Split(';')[3]) + 1];
                                    ExcelRangeBase cell_data = ws.Cells[int.Parse(str_location[3].Split(';')[1]), int.Parse(str_location[1].Split(';')[3]) + 1];
                                    //  ExcelRange cell_data_test = ws.Cells[int.Parse(str_location[3].Split(';')[1]), int.Parse(str_location[1].Split(';')[3]) + 1];
                                    Export_DatatableImage_Excel_tape(dt_region, "Image", ws, cell_Pic, false, count_sample);
                                    Export_DatatableImage_Excel_Graph(dt_region, "Graph", ws, cell_graph, false, count_sample);



                                    for (int i = 0; i < count_sample; i++)
                                    {
                                        if (i < dt_region.Rows.Count)
                                        {
                                            //cell_data_test.Offset(0, i).Te = lst_data[i];
                                            //cell_data.Offset(0, i).Value = lst_data[i];
                                            cell_data.Offset(0, i).Value = double.Parse(lst_data[i]);
                                            cell_data.Offset(0, i).Style.Numberformat.Format = "0.00";


                                        }
                                    }

                                    for (int i = 1; i < 13; i++)
                                    {
                                        if (myCode.checkDBNull(cell_data.Offset(i, -1).Value).Replace(" ", "").ToUpper().Contains("MINFORCE"))
                                        {
                                            ExcelRangeBase cell_min = cell_data.Offset(i, 0);
                                            ExcelRangeBase cell_max = cell_data.Offset(i + 1, 0);
                                            ExcelRangeBase cell_ave = cell_data.Offset(i + 2, 0);

                                            int c_offset = count_sample;
                                            cell_min.FormulaR1C1 = "=MIN(R[-" + i.ToString() + "]C[0]:R[-" + i.ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                            cell_max.FormulaR1C1 = "=MAX(R[-" + (i + 1).ToString() + "]C[0]:R[-" + (i + 1).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                            cell_ave.FormulaR1C1 = "=AVERAGE(R[-" + (i + 2).ToString() + "]C[0]:R[-" + (i + 2).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                            break;
                                        }
                                    }
                                    for (int i = 0; i < count_sample; i++)
                                    {
                                        if (i < dt_region.Rows.Count)
                                        {
                                            int ID = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                            if (dic_judgement.ContainsKey(ID - 1))
                                            {
                                                if (dic_judgement[ID - 1] == "FAIL")
                                                {
                                                    lst_judge.Add("Fail");

                                                }
                                                else
                                                {
                                                    lst_judge.Add("Pass");
                                                }
                                            }

                                        }
                                    }

                                    for (int t = 1; t < 5; t++)
                                    {
                                        if (myCode.checkDBNull(cell_data.Offset(t, -1).Value).ToUpper().Contains("JUDGEMENT") && !myCode.checkDBNull(cell_data.Offset(t, -1).Value).ToUpper().Replace(" ", "").Contains("FAILUREMODE"))
                                        {
                                            for (int i = 0; i < lst_judge.Count; i++)
                                            {
                                                cell_data.Offset(t, i).Value = lst_judge[i];

                                            }
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                export_ok = false;
                MessageBox.Show(new Form { TopMost = true }, mysheet + " chứa dữ liệu NG. Không thể xuất dữ liệu vào báo cáo", "Warning");
            }

        }

        public void export_excel_coupon(ExcelWorksheet ws, string itemcode, string lotno, string sheet, DataTable Data_all, DataTable dt_spec, DataGridView dgv_data, string type, string leader, int count_sample)
        {
            bool judge_all = true;
            List<string> lst_region = Data_all.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToList();
            string[] region = dt_spec.Rows[0]["Location"].ToString().Split('_');

            if (type.Contains("MASS") || type == "Other")
            {
                ExcelRangeBase cur_rgn_type = find_cell(ws, type.Replace("MASS(", "").Replace(")", ""));
                if (cur_rgn_type != null)
                    cur_rgn_type.Value = " v" + cur_rgn_type.Value.ToString();


                string str_infor = Data_all.Rows[0]["Sheet"].ToString().Replace(sheet, "").Replace("/", "");
                if (str_infor.Contains("_"))
                {
                    string itemname = str_infor.Split('_')[0];
                    string line = str_infor.Split('_')[1];
                    string ca = str_infor.Split('_')[2];
                    string date = str_infor.Split('_')[3];
                    string worker = str_infor.Split('_')[4];
                    List<string> lst_infor = new List<string> { itemname, itemcode, lotno, line, ca, date, worker };
                    export_info_mass(ws, lst_infor, sheet, leader);
                }

                for (int i = 1; i < 20; i++)
                {
                    for (int j = 1; j < 4; j++)
                    {
                        if (myCode.checkDBNull(ws.Cells[i, j].Value).ToUpper().Replace(" ", "").Contains("ITEM-LOT"))
                        {
                            ws.Cells[i, j].Offset(0, 1).Value = itemcode + "-" + lotno;
                            break;

                        }
                    }
                }


            }

            int r_end = 0;
            if (region.Length == 2)
            {
                r_end = 25;
            }
            else if (region.Length == 3)
            {
                r_end = 40;
            }
            else if (region.Length == 4)
            {
                r_end = 52;
            }
            insert_columns_other(ws, count_sample, r_end, type, dt_spec, sheet);

            foreach (string reg in region)
            {
                if (reg != "")
                {
                    string[] str_location = reg.Split('+');
                    string cpn = "";

                    foreach (string r in lst_region)
                    {

                        bool chk = true;
                        foreach (string key in r.Split('_'))
                        {
                            if (!reg.Replace(" ", "").ToUpper().Contains(key.Replace(" ", "").ToUpper()))
                            {
                                chk = false;
                                break;
                            }
                        }

                        if (chk)
                        {
                            cpn = r;
                            break;
                        }

                    }


                    if (cpn != "")
                    {
                        DataTable dt_region = Data_all.AsEnumerable().Where(s => s.Field<string>("Region") == cpn).CopyToDataTable();
                        if (dt_region.Rows.Count > 0)
                        {
                            List<string> lst_data = dt_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                            List<string> lst_judge = new List<string> { };


                            ExcelRangeBase cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[2]), int.Parse(str_location[1].Split(';')[3]) + 1];
                            ExcelRangeBase cell_graph = ws.Cells[int.Parse(str_location[2].Split(';')[1]), int.Parse(str_location[1].Split(';')[3]) + 1];
                            ExcelRangeBase cell_data = ws.Cells[int.Parse(str_location[3].Split(';')[1]), int.Parse(str_location[1].Split(';')[3]) + 1];

                            Export_DatatableImage_Excel(dt_region, "Image", ws, cell_Pic, false, count_sample);
                            Export_DatatableImage_Excel_Graph(dt_region, "Graph", ws, cell_graph, false, count_sample);


                            if (type == "NPI")
                            {
                                for (int i = 0; i < count_sample; i++)
                                {
                                    if (i < dt_region.Rows.Count)
                                    {
                                        if (myCode.IsNumeric(lst_data[i]))
                                        {
                                            cell_data.Offset(0, i).Value = Math.Round(double.Parse(lst_data[i]), 2);
                                            cell_data.Offset(0, i).Style.Numberformat.Format = "0.00";
                                        }

                                    }
                                }
                            }
                            else
                            {
                                for (int i = 0; i < count_sample; i++)
                                {
                                    if (i < dt_region.Rows.Count)
                                    {
                                        if (myCode.IsNumeric(lst_data[i].Split('_')[0].Replace("Max:", "")))
                                        {
                                            cell_data.Offset(0, i).Value = Math.Round(double.Parse(lst_data[i].Split('_')[0].Replace("Max:", "")), 2);
                                            cell_data.Offset(0, i).Style.Numberformat.Format = "0.00";
                                        }
                                        if (myCode.IsNumeric(lst_data[i].Split('_')[1].Replace("Average:", "")))
                                        {
                                            cell_data.Offset(1, i).Value = Math.Round(double.Parse(lst_data[i].Split('_')[1].Replace("Average:", "")), 2);
                                            cell_data.Offset(1, i).Style.Numberformat.Format = "0.00";
                                        }

                                    }
                                }
                            }

                            for (int i = 1; i < 13; i++)
                            {
                                if (myCode.checkDBNull(cell_data.Offset(i, -1).Value).Replace(" ", "").ToUpper().Contains("MINFORCE"))
                                {
                                    ExcelRangeBase cell_min = cell_data.Offset(i, 0);
                                    ExcelRangeBase cell_max = cell_data.Offset(i + 1, 0);
                                    ExcelRangeBase cell_ave = cell_data.Offset(i + 2, 0);

                                    int c_offset = count_sample;
                                    int r_offset = i - 1;
                                    cell_min.FormulaR1C1 = "=MIN(R[-" + r_offset.ToString() + "]C[0]:R[-" + r_offset.ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                    cell_max.FormulaR1C1 = "=MAX(R[-" + (r_offset + 1).ToString() + "]C[0]:R[-" + (r_offset + 1).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                    cell_ave.FormulaR1C1 = "=AVERAGE(R[-" + (r_offset + 2).ToString() + "]C[0]:R[-" + (r_offset + 2).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                    break;
                                }
                            }


                            //if (type.Contains("MASS"))
                            //{
                            for (int i = 0; i < count_sample; i++)
                            {
                                if (i < dt_region.Rows.Count)
                                {
                                    int ID = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                    if (dgv_data.Rows[ID - 1].Cells["Data"].Style.BackColor == Color.Red)
                                    {
                                        lst_judge.Add("Fail");
                                        judge_all = false;
                                    }
                                    else
                                    {
                                        lst_judge.Add("Pass");
                                    }
                                }
                            }

                            //for (int i = 0; i < count_sample; i++)
                            //{
                            //    cell_data.Offset(3, i).Value = lst_judge[i];
                            //}

                            for (int t = 1; t < 5; t++)
                            {
                                if (myCode.checkDBNull(cell_data.Offset(t, -1).Value).ToUpper().Contains("JUDGEMENT") && !myCode.checkDBNull(cell_data.Offset(t, -1).Value).ToUpper().Replace(" ", "").Contains("FAILUREMODE"))
                                {
                                    for (int i = 0; i < lst_judge.Count; i++)
                                    {
                                        cell_data.Offset(t, i).Value = lst_judge[i];
                                    }
                                    break;
                                }
                            }
                            //}

                        }
                    }
                }
            }


            int r_offset_del = 8;
            complete_sheet_other(ws, count_sample, r_offset_del, sheet, dt_spec);

            if (type.Contains("MASS"))
            {
                ExcelRangeBase cur_rgn_judge = find_cell(ws, "JUDGEMENT");
                if (cur_rgn_judge != null)
                {
                    if (judge_all)
                    {
                        cur_rgn_judge.Offset(1, 0).Value = "OK";
                        ws.Cells[1, 1].Value = "OK";
                    }
                    else
                    {
                        cur_rgn_judge.Offset(1, 0).Value = "NG";
                        ws.Cells[1, 1].Value = "NG";
                    }
                }
            }


        }


        // export unmating 
        public void export_excel_unmating_NPI(ExcelWorksheet ws, DataTable Data_tbl, DataTable dt_spec, ref bool export_ok)
        {

            int count_sample = int.Parse(dt_spec.Rows[0]["Count_sample"].ToString());
            int reg = 0;
            List<DataTable> lst_Table = new List<DataTable> { };
            Get_ListTable(-1, Data_tbl, new string[] { "Region" }, ref lst_Table, "Data");
            string[] region_data = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();

            SortedDictionary<int, string> dic_judgement = Check_spec_unmatingpull(Data_tbl, dt_spec);
            bool check_all = true;
            foreach (var result in dic_judgement.Values)
            {
                if (result == "FAIL")
                {
                    check_all = false;
                    break;
                }
            }
            if (check_all)
            {
                foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                {
                    if (spec_region != "")
                    {
                        if (reg < lst_Table.Count)
                        {

                            DataTable dt_region = lst_Table[reg];
                            string[] str_location = spec_region.Split('+');
                            ExcelRangeBase cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[1]), int.Parse(str_location[1].Split(';')[2]) + 1];
                            ExcelRangeBase cell_graph = ws.Cells[int.Parse(str_location[2].Split(';')[1]), int.Parse(str_location[1].Split(';')[2]) + 1];
                            ExcelRangeBase cell_data = ws.Cells[int.Parse(str_location[3].Split(';')[1]), int.Parse(str_location[1].Split(';')[2]) + 1];
                            // int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());

                            Export_DatatableImage_Excel_unmating(dt_region, "Image", ws, cell_Pic, false, count_sample);
                            Export_DatatableImage_Excel_Graph(dt_region, "Graph", ws, cell_graph, false, count_sample);

                            List<string> lst_data = dt_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                            List<string> lst_judge = new List<string> { };

                            string type = dt_spec.Rows[0]["Location"].ToString().Split('+')[0];
                            if (type == "A")
                            {
                                for (int i = 0; i < count_sample; i++)
                                {
                                    if (i < dt_region.Rows.Count)
                                    {
                                        if (myCode.IsNumeric(lst_data[i]))
                                        {
                                            cell_data.Offset(0, i).Value = double.Parse(lst_data[i]);
                                            cell_data.Offset(0, i).Style.Numberformat.Format = "0.00";
                                        }
                                    }
                                }
                            }
                            else if (type == "B")
                            {
                                for (int i = 0; i < count_sample; i++)
                                {
                                    if (i < dt_region.Rows.Count)
                                    {
                                        if (myCode.IsNumeric(lst_data[i]))
                                        {
                                            cell_data.Offset(1, i).Value = lst_data[i];
                                            cell_data.Offset(1, i).Style.Numberformat.Format = "0.00";
                                        }
                                    }
                                }
                            }

                            for (int i = 1; i < 13; i++)
                            {
                                if (myCode.checkDBNull(cell_data.Offset(i, -1).Value).Replace(" ", "").ToUpper().Contains("MINFORCE"))
                                {
                                    ExcelRangeBase cell_min = cell_data.Offset(i, 0);
                                    ExcelRangeBase cell_max = cell_data.Offset(i + 1, 0);
                                    ExcelRangeBase cell_ave = cell_data.Offset(i + 2, 0);
                                    int c_offset = count_sample;
                                    cell_min.FormulaR1C1 = "=MIN(R[-" + i.ToString() + "]C[0]:R[-" + i.ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                    cell_max.FormulaR1C1 = "=MAX(R[-" + (i + 1).ToString() + "]C[0]:R[-" + (i + 1).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                    cell_ave.FormulaR1C1 = "=AVERAGE(R[-" + (i + 2).ToString() + "]C[0]:R[-" + (i + 2).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                    break;
                                }
                            }


                            for (int i = 0; i < count_sample; i++)
                            {
                                if (i < dt_region.Rows.Count)
                                {
                                    cell_data.Offset(0, i).Value = lst_data[i];
                                    int ID = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                    if (dic_judgement.ContainsKey(ID - 1))
                                    {
                                        if (dic_judgement[ID - 1] == "FAIL")
                                        {
                                            lst_judge.Add("Fail");

                                        }
                                        else
                                        {
                                            lst_judge.Add("Pass");
                                        }
                                    }

                                }
                            }


                            for (int i = 0; i < count_sample; i++)
                            {
                                if (i < lst_judge.Count)
                                    cell_data.Offset(2, i).Value = lst_judge[i];
                            }

                        }
                        reg++;
                    }
                }
            }
            else
            {
                export_ok = false;
                MessageBox.Show(new Form { TopMost = true }, "IQC_UNMATING_PULL_TEST chứa dữ liệu NG. Không thể xuất dữ liệu vào báo cáo", "Warning");
            }
        }

        public string find_config_path(string src_string, string f_name)
        {
            DirectoryInfo di = new DirectoryInfo(src_string);
            var folder_lst = di.GetDirectories().Select(x => x.FullName).ToList();
            var temp2 = folder_lst.Where(x => new DirectoryInfo(x).Name == f_name).ToList();
            if (temp2.Count != 0)
            {
                return temp2.FirstOrDefault();
            }
            else
            {
                if (di.Parent != null)
                {
                    return find_config_path(di.Parent.FullName, f_name);
                }
                else
                {
                    return string.Empty;
                }
            }
        }


        public void export_excel_unmating(ExcelWorksheet ws, string itemcode, string lotno, string sheet, DataTable Data_tbl, DataTable dt_spec, DataGridView dgv_data, string type, string leader, int count_sample)
        {

            bool judge_all = true;
            if (type.Contains("MASS") || type == "Other")
            {
                ExcelRangeBase cur_rgn_type = find_cell(ws, type.Replace("MASS(", "").Replace(")", ""));
                if (cur_rgn_type != null)
                    cur_rgn_type.Value = " v" + cur_rgn_type.Value.ToString();

                string str_infor = Data_tbl.Rows[0]["Sheet"].ToString().Replace(sheet, "").Replace("/", "");
                if (str_infor.Contains("_"))
                {
                    string itemname = str_infor.Split('_')[0];
                    string line = str_infor.Split('_')[1];
                    string ca = str_infor.Split('_')[2];
                    string date = str_infor.Split('_')[3];
                    string worker = str_infor.Split('_')[4];
                    List<string> lst_infor = new List<string> { itemname, itemcode, lotno, line, ca, date, worker };
                    export_info_mass(ws, lst_infor, sheet, leader);
                }
            }
            int count_reg = dt_spec.Rows[0]["Location"].ToString().Split('_').Length - 1;
            int r_end = 10 * count_reg;
            insert_columns_other(ws, count_sample, r_end, type, dt_spec, sheet);

            int reg = 0;
            List<DataTable> lst_Table = new List<DataTable> { };
            Get_ListTable(-1, Data_tbl, new string[] { "Region" }, ref lst_Table, "Data");
            string[] region_data = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();
            foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
            {
                if (spec_region != "")
                {
                    if (reg < lst_Table.Count)
                    {

                        DataTable dt_region = lst_Table[reg];
                        string[] str_location = spec_region.Split('+');
                        ExcelRangeBase cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[1]), int.Parse(str_location[1].Split(';')[2]) + 1];
                        ExcelRangeBase cell_graph = ws.Cells[int.Parse(str_location[2].Split(';')[1]), int.Parse(str_location[1].Split(';')[2]) + 1];
                        ExcelRangeBase cell_data = ws.Cells[int.Parse(str_location[3].Split(';')[1]), int.Parse(str_location[1].Split(';')[2]) + 1];

                        Export_DatatableImage_Excel_unmating(dt_region, "Image", ws, cell_Pic, false, count_sample);
                        Export_DatatableImage_Excel_Graph(dt_region, "Graph", ws, cell_graph, false, count_sample);

                        List<string> lst_data = dt_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                        List<string> lst_judge = new List<string> { };

                        string typeAB = dt_spec.Rows[0]["Location"].ToString().Split('+')[0];
                        if (typeAB == "A")
                        {
                            for (int i = 0; i < count_sample; i++)
                            {
                                if (i < dt_region.Rows.Count)
                                {
                                    if (myCode.IsNumeric(lst_data[i]))
                                    {
                                        cell_data.Offset(0, i).Value = Math.Round(double.Parse(lst_data[i]), 2);
                                        cell_data.Offset(0, i).Style.Numberformat.Format = "0.00";
                                    }

                                }
                            }
                        }
                        else if (typeAB == "B")
                        {
                            for (int i = 0; i < count_sample; i++)
                            {
                                if (i < dt_region.Rows.Count)
                                {
                                    //  cell_data.Offset[1, i].Value = lst_data[i];
                                    cell_data.Offset(1, i).Value = Math.Round(double.Parse(lst_data[i]), 2);
                                    cell_data.Offset(1, i).Style.Numberformat.Format = "0.00";
                                }
                            }
                        }

                        double specZ = 0;
                        try
                        {
                            string z = dt_spec.Rows[0]["Location"].ToString().Split('(')[1].Split(')')[0].TrimEnd('N').TrimStart('≥');
                            specZ = double.Parse(z);
                        }
                        catch
                        {

                        }
                        for (int i = 1; i < 13; i++)
                        {
                            if (myCode.checkDBNull(cell_data.Offset(i, -1).Value).Replace(" ", "").ToUpper().Contains("MINFORCE"))
                            {
                                ExcelRangeBase cell_min = cell_data.Offset(i, 0);
                                ExcelRangeBase cell_max = cell_data.Offset(i + 1, 0);
                                ExcelRangeBase cell_ave = cell_data.Offset(i + 2, 0);
                                ExcelRangeBase cell_stdev = cell_data.Offset(i + 3, 0);
                                ExcelRangeBase cell_cpk = cell_data.Offset(i + 4, 0);
                                int c_offset = count_sample;
                                int r_offset = i;
                                cell_min.FormulaR1C1 = "=MIN(R[-" + r_offset.ToString() + "]C[0]:R[-" + r_offset.ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                cell_max.FormulaR1C1 = "=MAX(R[-" + (r_offset + 1).ToString() + "]C[0]:R[-" + (r_offset + 1).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                cell_ave.FormulaR1C1 = "=AVERAGE(R[-" + (r_offset + 2).ToString() + "]C[0]:R[-" + (r_offset + 2).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                cell_stdev.FormulaR1C1 = "=STDEV(R[-" + (r_offset + 3).ToString() + "]C[0]:R[-" + (r_offset + 3).ToString() + "]C[" + (c_offset - 3).ToString() + "])";
                                cell_cpk.FormulaR1C1 = $"=(R[{-2}]C[0]-{specZ})/(3*R[{-1}]C[0])";
                                cell_min.Style.Numberformat.Format = "0.00";
                                cell_max.Style.Numberformat.Format = "0.00";
                                cell_ave.Style.Numberformat.Format = "0.00";
                                cell_stdev.Style.Numberformat.Format = "0.00";
                                cell_cpk.Style.Numberformat.Format = "0.00";
                                break;
                            }
                        }

                        for (int i = 0; i < count_sample; i++)
                        {
                            if (i < dt_region.Rows.Count)
                            {
                                //   cell_data.Offset(0, i).Value = lst_data[i];
                                int ID = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                if (dgv_data.Rows[ID - 1].Cells["Data"].Style.BackColor == Color.Red)
                                {
                                    lst_judge.Add("Fail");
                                    judge_all = false;
                                }
                                else
                                {
                                    lst_judge.Add("Pass");
                                }
                            }
                        }
                        for (int i = 0; i < count_sample; i++)
                        {
                            if (i < lst_judge.Count)
                                cell_data.Offset(2, i).Value = lst_judge[i];
                        }
                    }
                    reg++;
                }
            }


            int r_offset_del = 8;
            complete_sheet_other(ws, count_sample, r_offset_del, sheet, dt_spec);
            if (type.Contains("MASS"))
            {
                ExcelRangeBase cur_rgn_judge = find_cell(ws, "JUDGEMENT");
                if (cur_rgn_judge != null)
                {
                    if (judge_all)
                    {
                        cur_rgn_judge.Offset(1, 0).Value = "OK";
                        ws.Cells[1, 1].Value = "OK";
                    }
                    else
                    {
                        cur_rgn_judge.Offset(1, 0).Value = "NG";
                        ws.Cells[1, 1].Value = "NG";
                    }
                }
            }
        }

        public void Export_DatatableImage_Excel_unmating(DataTable dt, string Image_Col_name, ExcelWorksheet tar_wrksht, ExcelRangeBase sel_rgn, bool row_offset, int count_sample)
        {
            int r_inx = 0;

            ExcelRangeBase rgn_begin = sel_rgn;

            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (count < 2 * count_sample)
                {
                    byte[] _img_byte = (byte[])(dr[Image_Col_name]);
                    byte[] img_byte = ImageCompress.CompressImage(_img_byte);
                    InsertPicture_Name(tar_wrksht, sel_rgn, img_byte, sel_rgn.Address);

                    int col_offset = sel_rgn.Columns;
                    int row_off = sel_rgn.Rows;
                    if (row_offset)
                    {
                        sel_rgn = sel_rgn.Offset(row_off, 0);
                    }
                    else
                    {
                        sel_rgn = sel_rgn.Offset(0, col_offset);
                    }
                    r_inx++;
                    count++;

                    if (count == count_sample)
                    {
                        sel_rgn = sel_rgn.Offset(1, -count_sample);

                    }
                }
                else
                {
                    break;
                }
            }


        }



        // export ACF 

        public void export_ACF_Wetting(ExcelWorksheet ws, string ItemCode, string LotNo, SqlConnection sqlcon)
        {
            try
            {
                ACFService service = new ACFService();
                service.ExportWCA(ws, ItemCode, LotNo);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        public void Export_ACF_Peel_NPI(ExcelWorksheet ws, string ItemCode, string LotNo, SqlConnection sqlcon)
        {
            try
            {

                new ACFService().ExportBoding(ws, ItemCode, LotNo);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //public ExcelRangeBase find_cell(ExcelWorksheet ws, string find_item)
        //{
        //    ExcelRangeBase cur_cell = null;

        //    for (int j = 1; j < 15; j++)
        //    {
        //        for (int i = 1; i < 15; i++)
        //        {
        //            if (myCode.checkDBNull(ws.Cells[i, j].Value).ToUpper().Replace(" ", "").Contains(find_item))
        //            {
        //                cur_cell = ws.Cells[i, j];
        //                return cur_cell;

        //            }
        //        }
        //    }

        //    return cur_cell;

        //}
        public void Export_ACF_Peel_MASS(ExcelWorksheet ws, DataTable tbl_Data, DataGridView dgv_peel)
        {
            bool judgeall_peel = true;

            ExcelRangeBase sel_rgn = ws.Cells[1, 1];
            string find_addr = Find_addr_peel("Before tape test", sel_rgn);
            if (find_addr == "")
                find_addr = Find_addr_peel("Glass coupon", sel_rgn);
            if (find_addr != "")
            {
                ExcelRangeBase rgn_ = ws.Cells[find_addr];
                ExcelRangeBase cur_rgn = ws.Cells[find_addr].Offset(0, 1);

                if (cur_rgn != null)
                {
                    Export_DatatableImage_Excel_ACF(tbl_Data, "Image_Before", ws, cur_rgn, false);
                    Export_DatatableImage_Excel_ACF(tbl_Data, "Image_After", ws, cur_rgn.Offset(1, 0), false);
                    Export_DatatableImage_Excel_Graph_ACF(tbl_Data, "Graph", ws, cur_rgn.Offset(2, 0), false);
                    List<string> lst_data = tbl_Data.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                    for (int i = 0; i < tbl_Data.Rows.Count; i++)
                    {
                        // cur_rgn.Offset(3, i).Value = lst_data[i];
                        if (myCode.IsNumeric(lst_data[i]))
                        {
                            cur_rgn.Offset(3, i).Value = Math.Round(double.Parse(lst_data[i]), 2);
                            cur_rgn.Offset(3, i).Style.Numberformat.Format = "0.00";

                        }
                    }

                    int count_sample = tbl_Data.Rows.Count;

                    List<string> lst_judge = new List<string> { };
                    for (int i = 0; i < count_sample; i++)
                    {

                        if (dgv_peel.Rows[i].Cells["Data"].Style.BackColor == Color.Red)
                        {
                            lst_judge.Add("Fail");
                            judgeall_peel = false;

                        }
                        else
                        {
                            lst_judge.Add("Pass");
                        }
                    }


                    for (int i = 1; i < 13; i++)
                    {
                        if (myCode.checkDBNull(cur_rgn.Offset(i, -1).Value).Replace(" ", "").ToUpper().Contains("MINFORCE"))
                        {
                            if (count_sample == 5 || count_sample == 32)
                            {
                                ExcelRangeBase cell_min = cur_rgn.Offset(i, 0);
                                ExcelRangeBase cell_max = cur_rgn.Offset(i + 1, 0);
                                ExcelRangeBase cell_ave = cur_rgn.Offset(i + 2, 0);

                                int c_offset = count_sample;
                                cell_min.FormulaR1C1 = "=MIN(R[-" + (i - 3).ToString() + "]C[0]:R[-" + (i - 3).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                cell_max.FormulaR1C1 = "=MAX(R[-" + (i - 2).ToString() + "]C[0]:R[-" + (i - 2).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                cell_ave.FormulaR1C1 = "=AVERAGE(R[-" + (i - 1).ToString() + "]C[0]:R[-" + (i - 1).ToString() + "]C[" + (c_offset - 1).ToString() + "])";

                            }
                            else if (count_sample == 15)
                            {
                                for (int k = 0; k < 3; k++)
                                {
                                    if (k == 0)
                                    {
                                        ExcelRangeBase cell_min = cur_rgn.Offset(i, 5 * k);
                                        ExcelRangeBase cell_max = cur_rgn.Offset(i + 1, 5 * k);
                                        ExcelRangeBase cell_ave = cur_rgn.Offset(i + 2, 5 * k);

                                        int c_offset = 5;
                                        cell_min.FormulaR1C1 = "=MIN(R[-" + (i - 3).ToString() + "]C[0]:R[-" + (i - 3).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                        cell_max.FormulaR1C1 = "=MAX(R[-" + (i - 2).ToString() + "]C[0]:R[-" + (i - 2).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                        cell_ave.FormulaR1C1 = "=AVERAGE(R[-" + (i - 1).ToString() + "]C[0]:R[-" + (i - 1).ToString() + "]C[" + (c_offset - 1).ToString() + "])";

                                    }
                                    else
                                    {
                                        ExcelRangeBase cell_min = cur_rgn.Offset(i, 5 * k + 1);
                                        ExcelRangeBase cell_max = cur_rgn.Offset(i + 1, 5 * k + 1);
                                        ExcelRangeBase cell_ave = cur_rgn.Offset(i + 2, 5 * k + 1);

                                        int c_offset = 5;
                                        cell_min.FormulaR1C1 = "=MIN(R[-" + (i - 3).ToString() + "]C[-1]:R[-" + (i - 3).ToString() + "]C[" + (c_offset - 2).ToString() + "])";
                                        cell_max.FormulaR1C1 = "=MAX(R[-" + (i - 2).ToString() + "]C[-1]:R[-" + (i - 2).ToString() + "]C[" + (c_offset - 2).ToString() + "])";
                                        cell_ave.FormulaR1C1 = "=AVERAGE(R[-" + (i - 1).ToString() + "]C[-1]:R[-" + (i - 1).ToString() + "]C[" + (c_offset - 2).ToString() + "])";

                                    }
                                }
                            }
                            break;
                        }
                    }

                    for (int smp = 0; smp < count_sample; smp++)
                    {
                        cur_rgn.Offset(5, smp).Value = lst_judge[smp];
                    }
                    ExcelRangeBase cur_rgn_judge = find_cell(ws, "JUDGEMENT");
                    if (cur_rgn_judge != null)
                    {
                        if (judgeall_peel)
                        {
                            cur_rgn_judge.Offset(1, 0).Value = "OK";
                            ws.Cells[1, 1].Value = "OK";
                        }
                        else
                        {
                            cur_rgn_judge.Offset(1, 0).Value = "NG";
                            ws.Cells[1, 1].Value = "NG";
                        }
                    }
                }
            }

        }

        public void Export_ACFFlatness(ExcelWorksheet ws, string ItemCode, string LotNo, SqlConnection sqlcon)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, "ACF_FLATNESS", filter_str);


            ProductIDService.FillProductID(src_dt, ItemCode, LotNo, "ACF_FLATNESS");
            if (src_dt.Columns.Contains("ProductID"))
            {
                IDictionary<string, string> dic = ExportProcess.FindAddressByText(ws, new[] { "Flex ID" }, true);
                if (dic.TryGetValue("Flex ID", out string value))
                {
                    int i = 1;
                    foreach (DataRow row in src_dt.Rows)
                    {
                        ws.Cells[ExportProcess.AddRow(value, i++ + 1)].Value = row["ProductID"];
                    }
                }
            }
            if (src_dt.Rows.Count > 0)
            {
                ExcelRangeBase curr_rgn = ws.Cells[1, 1];
                for (int i = 0; i < 100; i++)
                {
                    if (curr_rgn.Offset(i, 0).Value != null)
                    {
                        if (curr_rgn.Offset(i, 0).Value.ToString().Replace(" ", "").ToUpper().Contains("ACF Flatness".Replace(" ", "").ToUpper()) && !curr_rgn.Offset(i, 0).Value.ToString().Replace(" ", "").ToUpper().Contains("Measurement".Replace(" ", "").ToUpper()))
                        {
                            curr_rgn = curr_rgn.Offset(i, 1);

                            for (int c = 1; c < 5; c++)
                            {
                                if (myCode.checkDBNull(curr_rgn.Offset(2, c).Offset(-1, 0).Value).Contains("Point"))
                                {
                                    curr_rgn = curr_rgn.Offset(2, c);
                                    goto lbl_export;
                                }
                                else if (myCode.checkDBNull(curr_rgn.Offset(1, c).Offset(-1, 0).Value).Contains("Point"))
                                {
                                    curr_rgn = curr_rgn.Offset(1, c);
                                    goto lbl_export;
                                }
                            }
                        }
                    }
                }
            lbl_export:
                for (int i = 0; i < 5; i++)
                {
                    for (int j = 0; j < 11; j++)
                    {
                        curr_rgn.Offset(i, j).Value = double.Parse(src_dt.Rows[i][4 + j].ToString());
                        curr_rgn.Offset(i, j).Style.Numberformat.Format = "0.00";
                    }
                    curr_rgn.Offset(i, 11).Value = src_dt.Rows[i][4 + 11].ToString();
                }
            }
        }
        public void Export_ACF_Roughness(ExcelWorksheet ws, string ItemCode, string LotNo, SqlConnection sqlcon)
        {
            new ACFService().ExportRoughness(ItemCode, LotNo, ws);
        }

        public Boolean check_roughness_data(DataTable Data_tbl)
        {
            string spec_roughness = "Sa(<0.5um)^Sq(<0.75um)^Sdr(>0.0075)";


            bool chk = true;
            string[] arr_spec = spec_roughness.Split('^');
            double sa = 0;
            double sq = 0;
            double sdr = 0;
            foreach (string spec in arr_spec)
            {
                if (spec.Contains("Sa"))
                {
                    sa = double.Parse(spec.Replace(" ", "").Replace("Sa(", "").Replace(")", "").Replace("um", "").Replace("<", ""));
                }
                if (spec.Contains("sq"))
                {
                    sq = double.Parse(spec.Replace(" ", "").Replace("Sq(", "").Replace(")", "").Replace("um", "").Replace("<", ""));
                }
                if (spec.Contains("Sdr"))
                {
                    sdr = double.Parse(spec.Replace(" ", "").Replace("Sdr(", "").Replace(")", "").Replace("um", "").Replace(">", ""));
                }
            }
            for (int i = 0; i < Data_tbl.Rows.Count; i++)
            {
                for (int j = 0; j < Data_tbl.Columns.Count; j++)
                {
                    if (Data_tbl.Columns[j].ColumnName.Contains("Sa") && sa != 0)
                    {

                        if (myCode.IsNumeric(myCode.checkDBNull(Data_tbl.Rows[i][j])))
                        {
                            double data = double.Parse(myCode.checkDBNull(Data_tbl.Rows[i][j]));
                            if (data >= sa)
                            {
                                return false;
                            }


                        }
                    }

                    if (Data_tbl.Columns[j].ColumnName.Contains("Sq") && sq != 0)
                    {

                        if (myCode.IsNumeric(myCode.checkDBNull(Data_tbl.Rows[i][j])))
                        {
                            double data = double.Parse(myCode.checkDBNull(Data_tbl.Rows[i][j]));
                            if (data >= sq)
                            {
                                return false;
                            }

                        }
                    }

                    if (Data_tbl.Columns[j].ColumnName.Contains("Sdr") && sdr != 0)
                    {

                        if (myCode.IsNumeric(myCode.checkDBNull(Data_tbl.Rows[i][j])))
                        {
                            double data = double.Parse(myCode.checkDBNull(Data_tbl.Rows[i][j]));
                            if (data <= sdr)
                            {
                                return false;
                            }


                        }
                    }
                }
            }

            return chk;
        }


        public string ACF_pad_location(ExcelWorksheet ws)
        {

            string get_info = "";
            //ExcelRangeBase curr_rgn_1 = ws.Range["A1"];
            ExcelRangeBase curr_rgn_1 = ws.Cells[1, 1];
            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 10; i < 80; i++)
            {
                if (ws.Cells[1 + i, 1].Value != null)
                {
                    if (ws.Cells[1 + i, 1].Value.ToString().Replace(" ", "").ToUpper() == "ACF pad location".Replace(" ", "").ToUpper())
                    {
                        int k = 1;
                        get_info = i.ToString();
                        while (ws.Cells[1 + i + k, 2].Value.ToString() != "Roughness")
                        {
                            if (ws.Cells[1 + i + k, 2].Value.ToString().Contains("Sa"))
                            {

                                get_info += "+ sa";
                            }
                            if (ws.Cells[1 + i + k, 2].Value.ToString().Contains("Sq"))
                            {

                                get_info += "+ sq";
                            }
                            if (ws.Cells[1 + i + k, 2].Value.ToString().Contains("Sdr"))
                            {

                                get_info += "+ sdr";
                            }
                            k++;
                        }


                        get_info += "+" + k.ToString();
                        break;
                    }
                }
            }
            return get_info;


        }
        public List<List<string>> lst_roughness(DataTable dt)
        {
            List<List<string>> lst_roughness = new List<List<string>>() { };
            string[] arr_name = { "L1_Roughness_Sa", "L1_Roughness_Sq", "L1_Roughness_Sdr", "L2_Roughness_Sa", "L2_Roughness_Sq", "L2_Roughness_Sdr", "L3_Roughness_Sa", "L3_Roughness_Sq", "L3_Roughness_Sdr" };
            foreach (var item in arr_name)
            {
                List<string> lst_data = dt.AsEnumerable().Select(x => x.Field<string>(item)).ToList();
                lst_roughness.Add(lst_data);
            }

            return lst_roughness;
        }

        public void insert_row_excel(ExcelWorksheet ws, int rcopy_begin, int rcopy_end)
        {
            //ExcelRangeBase copyRange = ws.Rows[rcopy_begin + ":" + rcopy_end];
            //copyRange.Insert(myExcel.XlInsertShiftDirection.xlShiftDown);
            //ExcelRangeBase dest = ws.Rows[rcopy_begin + ":" + rcopy_end];
            //copyRange.Copy(dest);
        }
        public double CalculateStandardDeviation(IEnumerable<double> values)
        {
            double result = 0.0;
            if (values.Any())
            {
                double avg = values.Average();
                double num = values.Sum((double d) => Math.Pow(d - avg, 2.0));
                result = Math.Sqrt(num / (double)(values.Count() - 1));
            }

            return result;
        }
        public string Find_addr_peel(string in_item, ExcelRangeBase start_rgn)
        {
            //string result = start_rgn.AddressLocal.ToString();
            string result = "";
            int r_inx = 0;
            while (r_inx < 60)
            {
                if (myCode.checkDBNull(start_rgn.Offset(r_inx, 0).Value).Replace(" ", "").ToUpper().Contains(in_item.Replace(" ", "").ToUpper()))
                {
                    result = start_rgn.Offset(r_inx, 0).Address;
                    break;
                }
                r_inx++;
            }
            return result;
        }


        public Boolean check_cpk_roughness(SqlConnection sqlcon, string ItemCode, string LotNo)
        {
            bool chk = true;
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, "Roughness", filter_str);


            if (src_dt.Rows.Count > 0)
            {
                SortedDictionary<string, List<double>> dic_all = new SortedDictionary<string, List<double>> { };
                string[] arr_name = { "L1_Roughness_Sa", "L1_Roughness_Sq", "L1_Roughness_Sdr", "L2_Roughness_Sa", "L2_Roughness_Sq", "L2_Roughness_Sdr", "L3_Roughness_Sa", "L3_Roughness_Sq", "L3_Roughness_Sdr" };
                foreach (var item in arr_name)
                {
                    List<string> lst_data = src_dt.AsEnumerable().Select(x => x.Field<string>(item)).ToList();
                    List<double> lst_double = new List<double> { };

                    foreach (string data in lst_data)
                    {
                        if (myCode.IsNumeric(data))
                        {
                            lst_double.Add(double.Parse(data));
                        }

                    }
                    dic_all.Add(item, lst_double);
                }

                //  List<List<string>> lst_rghness = lst_roughness(src_dt, ref dic_all);

                foreach (var item in dic_all)
                {
                    List<double> lst = item.Value;
                    if (lst.Count > 0)
                    {
                        double max = lst.Max();
                        double min = lst.Min();
                        double Mean = lst.Average();
                        double STDEV = CalculateStandardDeviation(lst);

                        string vitri = item.Key;
                        double cpk = 0;
                        if (vitri.Contains("Sa"))
                        {
                            cpk = (0.5 - Mean) / (3 * STDEV);
                        }
                        else if (vitri.Contains("Sq"))
                        {
                            cpk = (0.75 - Mean) / (3 * STDEV);
                        }
                        else if (vitri.Contains("Sdr"))
                        {
                            cpk = (Mean - 0.0075) / (3 * STDEV);
                        }


                        if (cpk < 1.33)
                        {
                            chk = false;
                            break;
                        }
                    }


                }

            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Chưa có dữ liệu Roughness", "Thông báo");
                chk = false;
            }
            return chk;
        }

        public void Export_DatatableImage_Excel_ACF(DataTable dt, string Image_Col_name, ExcelWorksheet tar_wrksht, ExcelRangeBase sel_rgn, bool row_offset)
        {
            string file_folder = System.IO.Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = System.IO.Path.Combine(file_folder, DateTime.Now.ToString("ddMMMyy_HHmmss") + ".jpg");
            int r_inx = 0;
            //ExcelRangeBase sel_rgn = tar_wrksht.Range[tar_rgn];
            ExcelRangeBase rgn_begin = sel_rgn;

            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                byte[] _img_byte = (byte[])(dr[Image_Col_name]);
                byte[] img_byte = ImageCompress.CompressImage(_img_byte);
                InsertPicture_Name(tar_wrksht, sel_rgn, img_byte, sel_rgn.Address);

                int col_offset = sel_rgn.Columns;
                int row_off = sel_rgn.Rows;
                if (row_offset)
                {
                    sel_rgn = sel_rgn.Offset(row_off, 0);
                }
                else
                {
                    sel_rgn = sel_rgn.Offset(0, col_offset);
                }
                r_inx++;
                count++;
            }
        }

        public void Export_DatatableImage_Excel_Graph_ACF(DataTable dt, string Image_Col_name, ExcelWorksheet tar_wrksht, ExcelRangeBase sel_rgn, bool row_offset)
        {
            string file_folder = System.IO.Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            //  string file_dic = System.IO.Path.Combine(file_folder, DateTime.Now.ToString("ddMMMyy_HHmmss") + ".jpg");
            int r_inx = 0;

            ExcelRangeBase rgn_begin = sel_rgn;
            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                byte[] _img_byte = (byte[])(dr[Image_Col_name]);
                byte[] img_byte = ImageCompress.CompressImage(_img_byte);
                InsertPicture_Name(tar_wrksht, sel_rgn, img_byte, sel_rgn.Address);

                int col_offset = sel_rgn.Columns;
                int row_off = sel_rgn.Rows;
                if (row_offset)
                {
                    sel_rgn = sel_rgn.Offset(row_off, 0);
                }
                else
                {
                    sel_rgn = sel_rgn.Offset(0, col_offset);
                }
                r_inx++;

                count++;
            }



        }



        //judgement

        public SortedDictionary<int, string> Check_spec_Peel_Pull(string mysheet, DataTable tbl_data, DataTable dt_spec)
        {
            SortedDictionary<int, string> dic_judgement = new SortedDictionary<int, string> { };

            for (int i = 0; i < tbl_data.Rows.Count; i++)
            {
                dic_judgement.Add(i, "PASS");
            }
            if (dt_spec.Rows.Count > 0)
            {
                int reg = 0;
                List<DataTable> lst_Table = new List<DataTable> { };
                Get_ListTable(-1, tbl_data, new string[] { "Region" }, ref lst_Table, "Data");
                foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                {
                    if (reg < lst_Table.Count)
                    {
                        DataTable dt_region = lst_Table[reg];
                        string type = spec_region.Split('+')[0];
                        if (type == "A")
                        {
                            string spec = spec_region.Split('+')[3].Split(';')[0].Split('(')[1].Split(')')[0].Replace(" ", string.Empty).Replace("N", string.Empty).Replace("≥", string.Empty).Replace(">", string.Empty).Replace("<", string.Empty);

                            if (TDMK_Code.IsNumeric(spec))
                            {

                                for (int i = 0; i < dt_region.Rows.Count; i++)
                                {
                                    string val = dt_region.Rows[i]["Data"].ToString();
                                    if (myCode.IsNumeric(val))
                                    {
                                        int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                        if (Double.Parse(val) < Double.Parse(spec))
                                        {
                                            //dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                            dic_judgement[r - 1] = "FAIL";
                                        }

                                    }

                                }
                            }
                        }
                        else if (type == "B")
                        {
                            string spec = spec_region.Split('+')[4];
                            List<double> lst_data = new List<double> { };
                            if (TDMK_Code.IsNumeric(spec))
                            {
                                double UCL = double.Parse(spec_region.Split('+')[5].Split(';')[1]);
                                double LCL = double.Parse(spec_region.Split('+')[5].Split(';')[2]);

                                SortedDictionary<int, string> dic_data = new SortedDictionary<int, string> { };
                                for (int i = 0; i < dt_region.Rows.Count; i++)
                                {
                                    string val = dt_region.Rows[i]["Data"].ToString();
                                    if (myCode.IsNumeric(val))
                                    {
                                        int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                        dic_data.Add(r, val);
                                        lst_data.Add(double.Parse(val));
                                        if (Double.Parse(val) < Double.Parse(spec) || Double.Parse(val) > UCL || Double.Parse(val) < LCL)
                                        {
                                            //dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                            dic_judgement[r - 1] = "FAIL";

                                        }

                                    }
                                }
                                double R = double.Parse(spec_region.Split('+')[5].Split(';')[0]);

                                double tb = lst_data.ToArray().Average();

                                foreach (int r1 in dic_data.Keys)
                                {
                                    foreach (int r2 in dic_data.Keys)
                                    {
                                        if (r2 < r1)
                                        {
                                            double sub_data = Math.Abs(double.Parse(dic_data[r1]) - double.Parse(dic_data[r2]));

                                            if (sub_data > R && tb != 0)
                                            {

                                                double a1 = Math.Abs(double.Parse(dic_data[r1]) - tb);
                                                double a2 = Math.Abs(double.Parse(dic_data[r2]) - tb);
                                                if (a1 > a2)
                                                {
                                                    //dgv.Rows[r1 - 1].Cells["Data"].Style.BackColor = Color.Red;
                                                    dic_judgement[r1 - 1] = "FAIL";
                                                }
                                                else
                                                {
                                                    dic_judgement[r2 - 1] = "FAIL";
                                                    //dgv.Rows[r2 - 1].Cells["Data"].Style.BackColor = Color.Red;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        reg++;
                    }
                }

            }


            Check_Alldata(tbl_data, mysheet, ref dic_judgement);
            return dic_judgement;
        }

        public SortedDictionary<int, string> Check_spec_ShearTest_new(DataTable tbl_data, DataTable dt_spec)
        {
            SortedDictionary<int, string> dic_judgement = new SortedDictionary<int, string> { };

            for (int i = 0; i < tbl_data.Rows.Count; i++)
            {
                dic_judgement.Add(i, "PASS");
            }
            int reg = 0;
            List<DataTable> lst_Table = new List<DataTable> { };
            Get_ListTable(-1, tbl_data, new string[] { "Region" }, ref lst_Table, "Data");
            foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
            {
                if (reg < lst_Table.Count)
                {
                    DataTable dt_region = lst_Table[reg];
                    string type = spec_region.Split('+')[0];
                    if (type == "A")
                    {
                        string spec = spec_region.Split('+')[3].Split(';')[0].Split('(')[1].Split(')')[0].Replace(" ", string.Empty).Replace("N", string.Empty).Replace("≥", string.Empty).Replace(">", string.Empty).Replace("<", string.Empty);

                        for (int i = 0; i < dt_region.Rows.Count; i++)
                        {
                            string val = dt_region.Rows[i]["Data"].ToString();
                            if (myCode.IsNumeric(val))
                            {
                                int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                if (Double.Parse(spec) >= Double.Parse(val) * 9.81)
                                {
                                    //dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                    dic_judgement[r - 1] = "FAIL";

                                }
                            }

                        }
                    }
                    //else if (type == "B")
                    //{

                    //    double UCL = double.Parse(spec_region.Split('+')[4].Split(';')[1]);
                    //    double LCL = double.Parse(spec_region.Split('+')[4].Split(';')[2]);


                    //    SortedDictionary<int, string> dic_data = new SortedDictionary<int, string> { };
                    //    List<double> lst_data = new List<double> { };
                    //    for (int i = 0; i < dt_region.Rows.Count; i++)
                    //    {
                    //        string val = dt_region.Rows[i]["Data"].ToString();
                    //        if (myCode.IsNumeric(val))
                    //        {
                    //            int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
                    //            dic_data.Add(r, val);
                    //            lst_data.Add(double.Parse(val));
                    //            if (Double.Parse(val) > UCL || Double.Parse(val) < LCL)
                    //            {
                    //                //dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                    //                dic_judgement[r - 1] = "FAIL";

                    //            }
                    //        }

                    //    }
                    //    double R = double.Parse(spec_region.Split('+')[4].Split(';')[0]);

                    //    double tb = lst_data.Average();

                    //    foreach (int r1 in dic_data.Keys)
                    //    {
                    //        foreach (int r2 in dic_data.Keys)
                    //        {
                    //            if (r2 < r1)
                    //            {
                    //                double sub_data = Math.Abs(double.Parse(dic_data[r1]) - double.Parse(dic_data[r2]));
                    //                if (sub_data > R && tb != 0)
                    //                {

                    //                    double a1 = Math.Abs(double.Parse(dic_data[r1]) - tb);
                    //                    double a2 = Math.Abs(double.Parse(dic_data[r2]) - tb);
                    //                    if (a1 > a2)
                    //                    {
                    //                        // dgv.Rows[r1 - 1].Cells["Data"].Style.BackColor = Color.Red;
                    //                        dic_judgement[r1 - 1] = "FAIL";
                    //                    }
                    //                    else
                    //                    {
                    //                        // dgv.Rows[r2 - 1].Cells["Data"].Style.BackColor = Color.Red;
                    //                        dic_judgement[r2 - 1] = "FAIL";
                    //                    }
                    //                }
                    //            }
                    //        }
                    //    }


                    //}
                    reg++;
                }
            }

            Check_Alldata(tbl_data, "SHEAR_TEST", ref dic_judgement);
            return dic_judgement;
        }

        public SortedDictionary<int, string> check_spec_onproduct(DataTable tbl_data, DataTable dt_spec, string mysheet)
        {
            SortedDictionary<int, string> dic_judgement = new SortedDictionary<int, string> { };
            for (int i = 0; i < tbl_data.Rows.Count; i++)
            {
                dic_judgement.Add(i, "PASS");
            }

            int reg = 0;
            List<DataTable> lst_Table = new List<DataTable> { };
            Get_ListTable(-1, tbl_data, new string[] { "Region" }, ref lst_Table, "Data");
            foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
            {
                if (reg < lst_Table.Count)
                {
                    //DataTable dt_region = lst_Table[reg];
                    DataTable Data_all = tbl_data;
                    List<string> lst_region = Data_all.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToList();

                    string cpn = "";
                    foreach (string ar in lst_region)
                    {
                        if (spec_region.Replace(" ", "").ToUpper().Replace("COMPONENT", "TAPE").Contains(ar.Split('_')[0].Replace(" ", "").ToUpper()) && spec_region.Replace(" ", "").ToUpper().Replace("COMPONENT", "TAPE").Contains(ar.Split('_')[1].Replace(" ", "").ToUpper()))
                        {
                            cpn = ar;
                            break;
                        }

                    }

                    DataView dv = Data_all.AsDataView();
                    string str_filter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { cpn });
                    dv.RowFilter = str_filter;
                    DataTable dt_region = dv.ToTable();


                    string type = spec_region.Split('+')[0];
                    if (type == "A")
                    {
                        string spec_max = spec_region.Split('+')[3].Split(';')[0].Replace("(gf)", "");
                        string ll_max = "";
                        string ul_max = "";
                        string ll_ave = "";
                        string ul_ave = "";

                        if (spec_max.Contains("(") && spec_max.Contains(")"))
                        {
                            spec_max = spec_max.Split('(')[1].Split(')')[0].Replace(" ", ")").Replace("gf", "");
                            if (spec_max.Contains("-"))
                            {
                                ll_max = spec_max.Split('-')[0];
                                ul_max = spec_max.Split('-')[1];
                            }

                        }

                        string spec_average = spec_region.Split('+')[4].Split(';')[0].Replace("(gf)", "");
                        if (spec_max.Contains("(") && spec_max.Contains(")"))
                        {
                            spec_average = spec_average.Split('(')[1].Split(')')[0].Replace(" ", ")").Replace("gf", "");
                            if (spec_average.Contains("-"))
                            {
                                ll_ave = spec_average.Split('-')[0];
                                ul_ave = spec_average.Split('-')[1];
                            }

                        }



                        if (myCode.IsNumeric(ul_max) && myCode.IsNumeric(ll_max))
                        {
                            for (int i = 0; i < dt_region.Rows.Count; i++)
                            {
                                string val = dt_region.Rows[i]["Data"].ToString().Split('_')[0].Replace("Max:", "");
                                if (myCode.IsNumeric(val))
                                {
                                    int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                    if (Double.Parse(val) < Double.Parse(ll_max) || Double.Parse(val) > Double.Parse(ul_max))
                                    {
                                        // dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                        dic_judgement[r - 1] = "FAIL";
                                    }
                                }

                            }
                        }

                        if (myCode.IsNumeric(ul_ave) && myCode.IsNumeric(ll_ave))
                        {
                            for (int i = 0; i < dt_region.Rows.Count; i++)
                            {
                                string val = dt_region.Rows[i]["Data"].ToString().Split('_')[1].Replace("Average:", "");
                                if (myCode.IsNumeric(val))
                                {
                                    int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                    if (Double.Parse(val) < Double.Parse(ll_ave) || Double.Parse(val) > Double.Parse(ul_ave))
                                    {
                                        //dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                        dic_judgement[r - 1] = "FAIL";
                                    }
                                }
                            }
                        }
                    }
                    else if (type == "B")
                    {
                        double UCL = 0;
                        double LCL = 0;
                        try
                        {
                            if (spec_region.Split('+')[6].Split(';')[1] != "")
                            {
                                UCL = double.Parse(spec_region.Split('+')[6].Split(';')[1]);
                            }
                            if (spec_region.Split('+')[6].Split(';')[2] != "")
                            {
                                LCL = double.Parse(spec_region.Split('+')[6].Split(';')[2]);
                            }
                        }
                        catch
                        {
                            if (spec_region.Split('+')[5].Split(';')[1] != "")
                            {
                                UCL = double.Parse(spec_region.Split('+')[5].Split(';')[1]);
                            }
                            if (spec_region.Split('+')[5].Split(';')[2] != "")
                            {
                                LCL = double.Parse(spec_region.Split('+')[5].Split(';')[2]);
                            }
                        }


                        SortedDictionary<int, string> dic_data = new SortedDictionary<int, string> { };
                        List<double> lst_data = new List<double> { };
                        for (int i = 0; i < dt_region.Rows.Count; i++)
                        {
                            string val = dt_region.Rows[i]["Data"].ToString().Split('_')[1].Replace("Average:", "");
                            if (myCode.IsNumeric(val))
                            {
                                int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                dic_data.Add(r, val);
                                lst_data.Add(double.Parse(val));

                                if (Double.Parse(val) > UCL && UCL != 0)
                                {
                                    //dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                    dic_judgement[r - 1] = "FAIL";

                                }
                                if (Double.Parse(val) < LCL && LCL != 0)
                                {
                                    //dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                    dic_judgement[r - 1] = "FAIL";

                                }
                            }
                        }


                        string r_spec = "";
                        try
                        {
                            r_spec = spec_region.Split('+')[6].Split(';')[0];
                        }
                        catch
                        {
                            r_spec = spec_region.Split('+')[5].Split(';')[0];
                        }


                        if (r_spec != "")
                        {
                            double R = double.Parse(r_spec);
                            double tb = lst_data.ToArray().Average();

                            foreach (int r1 in dic_data.Keys)
                            {
                                foreach (int r2 in dic_data.Keys)
                                {
                                    if (r2 < r1)
                                    {
                                        double sub_data = Math.Abs(double.Parse(dic_data[r1]) - double.Parse(dic_data[r2]));

                                        if (sub_data > R && tb != 0)
                                        {
                                            double a1 = Math.Abs(double.Parse(dic_data[r1]) - tb);
                                            double a2 = Math.Abs(double.Parse(dic_data[r2]) - tb);
                                            if (a1 > a2)
                                            {
                                                //dgv.Rows[r1 - 1].Cells["Data"].Style.BackColor = Color.Red;
                                                dic_judgement[r1 - 1] = "FAIL";
                                            }
                                            else
                                            {
                                                // dgv.Rows[r2 - 1].Cells["Data"].Style.BackColor = Color.Red;
                                                dic_judgement[r2 - 1] = "FAIL";
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    // } 
                    reg++;
                }

            }

            Check_Alldata(tbl_data, mysheet, ref dic_judgement);
            return dic_judgement;
        }
        public SortedDictionary<int, string> Check_spec_unmatingpull(DataTable tbl_data, DataTable dt_spec)
        {
            SortedDictionary<int, string> dic_judgement = new SortedDictionary<int, string> { };
            for (int i = 0; i < tbl_data.Rows.Count; i++)
            {
                dic_judgement.Add(i, "PASS");
            }

            string spec = dt_spec.Rows[0]["Location"].ToString().Split('+')[3].Split(';')[0].Split('(')[1].Split(')')[0].Replace(" ", string.Empty).Replace("N", string.Empty).Replace("≥", string.Empty).Replace(">", string.Empty).Replace("<", string.Empty);
            if (TDMK_Code.IsNumeric(spec))
            {
                for (int i = 0; i < tbl_data.Rows.Count; i++)
                {
                    string val = tbl_data.Rows[i]["Data"].ToString();
                    if (myCode.IsNumeric(val))
                    {
                        if (Double.Parse(val) < Double.Parse(spec))
                        {
                            dic_judgement[i] = "FAIL";

                        }
                    }

                }
            }

            Check_Alldata(tbl_data, "IQC_UNMATING_PULL_TEST", ref dic_judgement);
            return dic_judgement;

        }

        public SortedDictionary<int, string> check_spec_coupon(DataTable tbl_data, DataTable dt_spec, string mysheet)
        {
            SortedDictionary<int, string> dic_judgement = new SortedDictionary<int, string> { };
            for (int i = 0; i < tbl_data.Rows.Count; i++)
            {
                dic_judgement.Add(i, "PASS");
            }

            int reg = 0;
            List<DataTable> lst_Table = new List<DataTable> { };
            Get_ListTable(-1, tbl_data, new string[] { "Region" }, ref lst_Table, "Data");
            foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
            {
                if (reg < lst_Table.Count)
                {
                    DataTable dt_region = lst_Table[reg];
                    string type = spec_region.Split('+')[0];
                    if (type == "A")
                    {
                        string spec = "";

                        if (mysheet == "IQC_LINER_PEELING_COUPON")
                        {
                            if (spec_region.Split('+')[3].Contains("Judgement") && spec_region.Split('+')[3].Contains("/") && spec_region.Split('+')[3].Contains("("))
                            {
                                spec = spec_region.Split('+')[3].Split(';')[0].Split('(')[1].Split(')')[0].Replace(" ", string.Empty).Split('/')[0].Replace("gf", "");
                            }
                            else if (spec_region.Split('+')[4].Contains("Judgement") && spec_region.Split('+')[4].Contains("/") && spec_region.Split('+')[4].Contains("("))
                            {
                                spec = spec_region.Split('+')[4].Split(';')[0].Split('(')[1].Split(')')[0].Replace(" ", string.Empty).Split('/')[0].Replace("gf", "");
                            }


                            if (spec.Contains("-"))
                            {
                                string ll = spec.Split('-')[0];
                                string ul = spec.Split('-')[1];

                                if (myCode.IsNumeric(ul) && myCode.IsNumeric(ll))
                                {
                                    for (int i = 0; i < dt_region.Rows.Count; i++)
                                    {
                                        string val = dt_region.Rows[i]["Data"].ToString();
                                        if (myCode.IsNumeric(val))
                                        {
                                            int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                            if (Double.Parse(val) < Double.Parse(ll) || Double.Parse(val) > Double.Parse(ul))
                                            {
                                                //dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                                dic_judgement[r - 1] = "FAIL";
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if (mysheet == "IQC_PSA_PEELING_COUPON")
                        {
                            //spec = spec_region.Split('+')[3].Split(';')[0].Split('(')[1].Split('N')[0].Replace(" ", string.Empty).Replace("≥", string.Empty).Replace(">", string.Empty).Replace("<", string.Empty);
                            if (spec_region.Split('+')[3].Contains("Judgement") && spec_region.Split('+')[3].Contains("/") && spec_region.Split('+')[3].Contains("("))
                            {
                                spec = spec_region.Split('+')[3].Split(';')[0].Split('(')[1].Split('N')[0].Replace(" ", string.Empty).Replace("≥", string.Empty).Replace(">", string.Empty).Replace("<", string.Empty);
                            }
                            else if (spec_region.Split('+')[4].Contains("Judgement") && spec_region.Split('+')[4].Contains("/") && spec_region.Split('+')[4].Contains("("))
                            {
                                spec = spec_region.Split('+')[3].Split(';')[0].Split('(')[1].Split('N')[0].Replace(" ", string.Empty).Replace("≥", string.Empty).Replace(">", string.Empty).Replace("<", string.Empty);
                            }



                            if (TDMK_Code.IsNumeric(spec))
                            {

                                for (int i = 0; i < dt_region.Rows.Count; i++)
                                {
                                    string val = dt_region.Rows[i]["Data"].ToString();
                                    if (myCode.IsNumeric(val))
                                    {
                                        int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                        if (Double.Parse(val) < Double.Parse(spec))
                                        {
                                            //dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                            dic_judgement[r - 1] = "FAIL";
                                        }
                                    }

                                }
                            }
                        }
                    }
                    else if (type == "B")
                    {
                        double UCL = 0;
                        double LCL = 0;
                        try
                        {
                            if (spec_region.Split('+')[6].Split(';')[1] != "")
                            {
                                UCL = double.Parse(spec_region.Split('+')[6].Split(';')[1]);
                            }
                            if (spec_region.Split('+')[6].Split(';')[2] != "")
                            {
                                LCL = double.Parse(spec_region.Split('+')[6].Split(';')[2]);
                            }
                        }
                        catch
                        {
                            if (spec_region.Split('+')[5].Split(';')[1] != "")
                            {
                                UCL = double.Parse(spec_region.Split('+')[5].Split(';')[1]);
                            }
                            if (spec_region.Split('+')[5].Split(';')[2] != "")
                            {
                                LCL = double.Parse(spec_region.Split('+')[5].Split(';')[2]);
                            }
                        }




                        SortedDictionary<int, string> dic_data = new SortedDictionary<int, string> { };

                        List<double> lst_data = new List<double> { };
                        for (int i = 0; i < dt_region.Rows.Count; i++)
                        {
                            string val = dt_region.Rows[i]["Data"].ToString().Split('_')[1].Replace("Average:", "");
                            if (myCode.IsNumeric(val))
                            {
                                int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                dic_data.Add(r, val);
                                lst_data.Add(double.Parse(val));

                                if (Double.Parse(val) > UCL && UCL != 0)
                                {
                                    //dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                    dic_judgement[r - 1] = "FAIL";

                                }
                                if (Double.Parse(val) < LCL && LCL != 0)
                                {
                                    //dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                    dic_judgement[r - 1] = "FAIL";

                                }
                            }
                        }

                        string r_spec = "";

                        try
                        {
                            r_spec = spec_region.Split('+')[6].Split(';')[0];
                        }
                        catch
                        {
                            r_spec = spec_region.Split('+')[5].Split(';')[0];
                        }


                        if (r_spec != "")
                        {
                            double R = double.Parse(r_spec);
                            double tb = lst_data.ToArray().Average();

                            foreach (int r1 in dic_data.Keys)
                            {
                                foreach (int r2 in dic_data.Keys)
                                {
                                    if (r2 < r1)
                                    {
                                        double sub_data = Math.Abs(double.Parse(dic_data[r1]) - double.Parse(dic_data[r2]));

                                        if (sub_data > R && tb != 0)
                                        {

                                            double a1 = Math.Abs(double.Parse(dic_data[r1]) - tb);
                                            double a2 = Math.Abs(double.Parse(dic_data[r2]) - tb);
                                            if (a1 > a2)
                                            {
                                                //dgv.Rows[r1 - 1].Cells["Data"].Style.BackColor = Color.Red;
                                                dic_judgement[r1 - 1] = "FAIL";

                                            }
                                            else
                                            {

                                                //dgv.Rows[r2 - 1].Cells["Data"].Style.BackColor = Color.Red;
                                                dic_judgement[r2 - 1] = "FAIL";
                                            }
                                        }
                                    }
                                }
                            }
                        }

                    }

                    reg++;
                }



            }


            Check_Alldata(tbl_data, mysheet, ref dic_judgement);
            return dic_judgement;
        }


        public void Check_Alldata(DataTable tbl_data, string sheet, ref SortedDictionary<int, string> dic_judgement)
        {
            if (sheet != "CROSS_SECTION" && sheet != "GAP_CONNECTOR")
            {
                List<DataTable> lst_Table = new List<DataTable> { };
                Get_ListTable(-1, tbl_data, new string[] { "Region" }, ref lst_Table, "Data");
                foreach (DataTable dt_region in lst_Table)
                {
                    List<double> lst_data = new List<double> { };
                    SortedDictionary<int, string> dic_data = new SortedDictionary<int, string> { };

                    double tb = 0;
                    if (sheet == "LINER_PEEL_TEST_ON_PRODUCT" || sheet == "PSA_PEEL_TEST_ON_PRODUCT")
                    {
                        for (int i = 0; i < dt_region.Rows.Count; i++)
                        {
                            string val = myCode.checkDBNull(dt_region.Rows[i]["Data"]).Split('_')[1].Replace("Average:", "");
                            if (myCode.IsNumeric(val))
                            {
                                dic_data.Add(i, val);
                                lst_data.Add(double.Parse(val));
                            }
                        }
                    }
                    else
                    {
                        for (int i = 0; i < dt_region.Rows.Count; i++)
                        {
                            string val = myCode.checkDBNull(dt_region.Rows[i]["Data"]);
                            if (val != "")
                            {
                                dic_data.Add(i, val);
                                if (myCode.IsNumeric(val))
                                {
                                    lst_data.Add(double.Parse(val));
                                }
                            }
                        }

                    }

                    if (lst_data.Count > 0)
                    {
                        tb = lst_data.ToArray().Average();
                    }


                    double R = 0;
                    switch (sheet)
                    {
                        case "PEEL_TEST":
                            R = 10;
                            break;

                        case "MATING_PULL_TEST":
                            R = 15;
                            break;

                        case "SHEAR_TEST":
                            R = 5;
                            break;

                        case "IQC_UNMATING_PULL_TEST":
                            R = 10;
                            break;

                        case "IQC_LINER_PEELING_COUPON":
                            R = 0;
                            break;

                        case "IQC_PSA_PEELING_COUPON":
                            R = 0;
                            break;

                        case "LINER_PEEL_TEST_ON_PRODUCT":
                            R = 20;
                            break;

                        case "PSA_PEEL_TEST_ON_PRODUCT":
                            R = 20;
                            break;

                        default:
                            break;
                    }

                    if (R != 0)
                    {
                        foreach (int r1 in dic_data.Keys)
                        {
                            foreach (int r2 in dic_data.Keys)
                            {
                                if (r2 > r1)
                                {
                                    double sub_data = 0;
                                    sub_data = Math.Abs(double.Parse(dic_data[r1]) - double.Parse(dic_data[r2]));


                                    if (sub_data > R && tb != 0)
                                    {

                                        double a1 = Math.Abs(double.Parse(dic_data[r1]) - tb);
                                        double a2 = Math.Abs(double.Parse(dic_data[r2]) - tb);
                                        if (a1 > a2)
                                        {
                                            int ID = int.Parse(dt_region.Rows[r1]["ID"].ToString());
                                            //dgv.Rows[ID - 1].Cells["Data"].Style.BackColor = Color.Red;
                                            dic_judgement[ID - 1] = "FAIL";
                                        }
                                        else
                                        {
                                            int ID = int.Parse(dt_region.Rows[r2]["ID"].ToString());
                                            dic_judgement[ID - 1] = "FAIL";
                                            // dgv.Rows[ID - 1].Cells["Data"].Style.BackColor = Color.Red;
                                        }
                                    }

                                }
                            }
                        }
                    }
                }

            }
        }

        public SortedDictionary<int, string> Check_spec_GAP(DataTable tbl_data, DataTable dt_spec, string mysheet)
        {
            SortedDictionary<int, string> dic_judgement = new SortedDictionary<int, string> { };
            for (int i = 0; i < tbl_data.Rows.Count; i++)
            {
                dic_judgement.Add(i, "PASS");
            }

            for (int i = 0; i < tbl_data.Rows.Count; i++)
            {
                string val = myCode.checkDBNull(tbl_data.Rows[i]["Data"]);
                if (myCode.IsNumeric(val.Split('/')[0].TrimEnd(' ', ';').Split(';').Last()) && myCode.IsNumeric(val.Split('/')[1].TrimEnd(' ', ';').Split(';').Last()))
                {
                    double data1 = double.Parse(val.Split('/')[0].TrimEnd(' ', ';').Split(';').Last());
                    double data2 = double.Parse(val.Split('/')[1].TrimEnd(' ', ';').Split(';').Last());

                    double sub_data = Math.Abs(data1 - data2);
                    if (sub_data > 30)
                    {
                        dic_judgement[i] = "FAIL";
                        break;
                    }
                }
            }

            string[] region = dt_spec.Rows[0]["Location"].ToString().Split('+')[1].Split('_');
            DataTable Data_tbl = tbl_data;

            List<string> lst_region_db = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).ToList();
            foreach (string reg in region)
            {
                if (reg != "")
                {
                    int code = 1;
                    foreach (string pos in reg.Split('^'))
                    {
                        if (pos != "")
                        {
                            int row_begin = int.Parse(pos.Split(';')[1]);
                            int count_r_offset = int.Parse(pos.Split(';')[2]);
                            string spec_b = pos.Split(';')[3];
                            if (spec_b.Contains("<") && spec_b.Contains("µm"))
                            {
                                if (myCode.IsNumeric(spec_b.Split('<')[1].Replace(" ", "").Replace("µm", "").Replace(")", "")))
                                {
                                    double spec = double.Parse(spec_b.Split('<')[1].Replace(" ", "").Replace("µm", "").Replace(")", ""));

                                    string find_text = "";
                                    if (reg.Split('^').Length == 3)
                                    {
                                        find_text = code.ToString();
                                    }
                                    string filter_reg = "";

                                    if (pos.Split(';')[0] == "IOPIN")
                                    {
                                        foreach (string i in lst_region_db)
                                        {
                                            if (i.ToUpper().Contains("DOC") && i.Contains(find_text))
                                            {
                                                filter_reg = i;
                                                break;
                                            }
                                        }
                                    }
                                    else if (pos.Split(';')[0] == "RIGHT" || pos.Split(';')[0] == "LEFT")
                                    {
                                        foreach (string i in lst_region_db)
                                        {
                                            if (i.ToUpper().Contains("TRU") && i.Contains(find_text))
                                            {
                                                filter_reg = i;
                                                break;
                                            }
                                        }
                                    }

                                    DataView dv = Data_tbl.AsDataView();
                                    dv.RowFilter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { filter_reg });
                                    int min = new int[] { dv.Count, tbl_data.Rows.Count - 1 }.Min();
                                    for (int inx = 0; inx < min; inx++)
                                    {
                                        DataRow dr = dv[inx].Row;
                                        int r_x = Data_tbl.Rows.IndexOf(dr);

                                        string val = myCode.checkDBNull(Data_tbl.Rows[r_x]["Data"]);
                                        double data1 = double.Parse(val.Split('/')[0].Split(';')[1]);
                                        double data2 = double.Parse(val.Split('/')[1].Split(';')[1]);

                                        if (data1 > spec || data2 > spec)
                                        {
                                            dic_judgement[r_x] = "FAIL";
                                        }

                                    }
                                }
                            }
                        }
                        code++;
                    }

                }

            }

            return dic_judgement;

        }

        /*************************************************************************************************************************/
        public string find_format_mass(string in_data_loc, string ItemCode)
        {
            string result = "";
            if (System.IO.Directory.Exists(in_data_loc))
            {
                string[] file_xlsm = Directory.GetFiles(in_data_loc, "*" + ItemCode + "*.xlsm");

                if (file_xlsm.Length > 0)
                {
                    result = file_xlsm[0];
                }
                else
                {
                    string[] file_xlsx = Directory.GetFiles(in_data_loc, "*" + ItemCode + "*.xlsx ");
                    if (file_xlsx.Length > 0)
                    {
                        result = file_xlsx[0];
                    }
                }
            }
            result = result.Replace("~$", "");
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



        public ExcelWorksheet Copy(ExcelWorkbook workbook, string existingWorksheetName, string newWorksheetName)
        {
            ExcelWorksheet worksheet = workbook.Worksheets.Copy(existingWorksheetName, newWorksheetName);
            return worksheet;
        }


        public ExcelPackage Create_report_NPI(string file_format, string report_folder, string itemcode, string lotno, string process_name)
        {
            ExcelWorksheet ws = null;

            string export_path = System.IO.Path.Combine(report_folder, itemcode + "-" + lotno + ".xlsx");

            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage xlPackage = open_excel(export_path);
            ExcelWorkbook wb_format = xlPackage.Workbook;


            ExcelWorkbook report_saved = null;
            ExcelPackage sourcePackage = null;

        lbl_export:
            if (System.IO.File.Exists(export_path))
            {
                if (MessageBox.Show(new Form { TopMost = true }, "Báo cáo đã tồn tại. Bạn có muốn thay thế không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    System.IO.File.Delete(export_path);
                    goto lbl_export;
                }
            }
            else
            {
                if (!System.IO.Directory.Exists(report_folder))
                    System.IO.Directory.CreateDirectory(report_folder);

                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                sourcePackage = new ExcelPackage(new FileInfo(file_format));
                report_saved = sourcePackage.Workbook;
                string _process = process_name.Replace("_", "").Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "").ToUpper();
                string mySheet = "";
                foreach (ExcelWorksheet tg_sht in wb_format.Worksheets)
                {
                    string cur_sht_name = tg_sht.Name.ToUpper().Replace("_", "").Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "");
                    if (_process == cur_sht_name)
                    {
                        mySheet = tg_sht.Name;
                        break;
                    }
                }

                if (mySheet != "")
                {
                    ExcelWorksheet worksheet0 = wb_format.Worksheets[mySheet];
                    report_saved.Worksheets.Add(mySheet, worksheet0);
                    report_saved.Worksheets.Delete(0);

                    ws = report_saved.Worksheets[0];
                }

            }
            return sourcePackage;

        }





        public void export_NPI_ACF(string file_format, string report_folder, string itemcode, string lotno, string process_name, SqlConnection sqlcon)
        {
            string export_path = System.IO.Path.Combine(report_folder, itemcode + "-" + lotno + ".xlsx");

            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage xlPackage = open_excel(file_format);
            ExcelWorkbook wb_format = xlPackage.Workbook;

            ExcelWorkbook report_saved = null;
            ExcelPackage sourcePackage = null;

        lbl_export:
            if (System.IO.File.Exists(export_path))
            {
                if (MessageBox.Show(new Form { TopMost = true }, "Báo cáo đã tồn tại. Bạn có muốn thay thế không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    System.IO.File.Delete(export_path);
                    goto lbl_export;
                }
            }
            else
            {
                if (!System.IO.Directory.Exists(report_folder))
                    System.IO.Directory.CreateDirectory(report_folder);

                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                sourcePackage = new ExcelPackage();
                report_saved = sourcePackage.Workbook;



                string _process = process_name.Replace("_", "").Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "").ToUpper();
                string mySheet = "";
                foreach (ExcelWorksheet tg_sht in wb_format.Worksheets)
                {
                    string cur_sht_name = tg_sht.Name.ToUpper().Replace("_", "").Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "");
                    if (_process == cur_sht_name)
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
                    export_ACF_Wetting(ws, itemcode, lotno, sqlcon);
                    Export_ACF_Peel_NPI(ws, itemcode, lotno, sqlcon);
                    try
                    {

                        Export_ACFFlatness(ws, itemcode, lotno, sqlcon);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    try
                    {
                        Export_ACF_Roughness(ws, itemcode, lotno, sqlcon);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }

            }

            if (sourcePackage != null)
            {
                sourcePackage.SaveAs(new FileInfo(export_path));
            }

            //myExcel.Workbook wb = TDMK_Code.open_excel_file(export_path, "", "");

            try
            {
                var excelApp = new myExcel.Application();
                excelApp.Visible = true;
                excelApp.Workbooks.Open(export_path);
            }
            catch
            {
                MessageBox.Show(new Form { TopMost = true }, "Lỗi khi mở file báo cáo", "Thông báo");
            }
            MessageBox.Show(new Form { TopMost = true }, "Xuất báo cáo thành công", "Thông báo");

        }
        public void export_info_peel(ExcelWorksheet ws, List<string> lst_infor, string leader)
        {

            string[] arr_item = new string[5] { "ITEMNAME:", "ITEMCODE:", "LOTNO:", "DATE:", "PROCESS:" };
            Dictionary<string, string> dic_item = new Dictionary<string, string>();

            int i = 0;
            foreach (string item in arr_item)
            {
                dic_item.Add(item, lst_infor[i]);
                i++;

            }

            foreach (var val in dic_item)
            {
                ExcelRangeBase cur_rgn = find_cell(ws, val.Key);
                if (cur_rgn != null)
                {
                    if (val.Key == "LOTNO:")
                    {
                        //cur_rgn.Offset(0, 1).FormulaR1C1 = "'" + val.Value;
                        cur_rgn.Offset(0, 1).Value = val.Value;
                        cur_rgn.Offset(0, 1).Style.Numberformat.Format = "#";


                    }
                    else
                    {
                        cur_rgn.Offset(0, 1).Value = val.Value;
                    }

                }
            }
            // string leader = txtOperator.Text;
            ExcelRangeBase cur_rgn_worker = find_cell(ws, "WORKER");
            if (cur_rgn_worker != null)
                cur_rgn_worker.Offset(1, 0).Value = lst_infor[5];

            ExcelRangeBase cur_rgn_leader = find_cell(ws, "LEADER");
            if (cur_rgn_leader != null)
                cur_rgn_leader.Offset(1, 0).Value = leader;

        }

        public void export_info_mass(ExcelWorksheet ws, List<string> lst_infor, string sheet, string leader)
        {
            string[] arr_item = new string[6] { "ITEMNAME:", "ITEMCODE:", "LOTNO:", "LINE:", "SHIFT/CA:", "DATE:" };
            Dictionary<string, string> dic_item = new Dictionary<string, string>();

            int i = 0;
            foreach (string item in arr_item)
            {
                dic_item.Add(item, lst_infor[i]);
                i++;

            }

            if (!sheet.Contains("COUPON") && !sheet.Contains("PRODUCT"))
            {
                foreach (var val in dic_item)
                {
                    ExcelRangeBase cur_rgn = find_cell(ws, val.Key);
                    if (cur_rgn != null)
                    {
                        if (val.Key == "LOTNO:")
                        {
                            //cur_rgn.Offset(0, 1).FormulaR1C1 = "'" + val.Value;
                            cur_rgn.Offset(0, 1).Value = val.Value;
                            cur_rgn.Offset(0, 1).Style.Numberformat.Format = "#";
                        }
                        else
                        {
                            cur_rgn.Offset(0, 1).Value = val.Value;
                        }
                    }
                }
            }

            ExcelRangeBase cur_rgn_worker = find_cell(ws, "WORKER");
            if (cur_rgn_worker != null)
                cur_rgn_worker.Offset(1, 0).Value = lst_infor[6];

            ExcelRangeBase cur_rgn_leader = find_cell(ws, "LEADER");
            if (cur_rgn_leader != null)
                cur_rgn_leader.Offset(1, 0).Value = leader;
        }

        public void export_MASS_ACFpeel(string file_format, string report_folder, string itemcode, string lotno, DataGridView dgv_peel, string type, string leader)
        {
            //string export_path = System.IO.Path.Combine(report_folder, itemcode + "-" + lotno + ".xlsx");
            string remark = dgv_peel.Rows[0].Cells["Remark"].Value.ToString();
            // string export_path = System.IO.Path.Combine(report_folder, type.Replace("MASS", "").Replace("(", "").Replace(")", "") + "-" + itemcode + "-" + lotno + "-" + remark + ".xlsx");
            string Itemname = remark.Split('_')[0];
            string export_path = System.IO.Path.Combine(report_folder, "ACF Bonding " + Itemname + "-" + itemcode + "-" + lotno + ".xlsx");

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
                if (!System.IO.Directory.Exists(report_folder))
                    System.IO.Directory.CreateDirectory(report_folder);

                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                sourcePackage = new ExcelPackage(new FileInfo(file_format));
                report_saved = sourcePackage.Workbook;

                ExcelWorksheet ws = report_saved.Worksheets[0];
                DataTable dt_peel = (DataTable)dgv_peel.DataSource;
                if (dt_peel != null)
                {
                    string[] infor = dt_peel.Rows[0]["Remark"].ToString().Split('_');
                    string itemname = infor[0];
                    string date = infor[1];
                    string worker = infor[2];
                    string process = type.Split('(')[1].Split(')')[0];

                    List<string> lst_infor = new List<string> { itemname, itemcode, lotno, date, process, worker };
                    export_info_peel(ws, lst_infor, leader);
                    Export_ACF_Peel_MASS(ws, dt_peel, dgv_peel);

                    if (sourcePackage != null)
                    {
                        sourcePackage.SaveAs(new FileInfo(export_path));
                    }
                    //myExcel.Workbook wb = TDMK_Code.open_excel_file(export_path, "", ""); 

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
                        MessageBox.Show(new Form { TopMost = true }, "Lỗi khi mở file báo cáo", "Thông báo");
                    }

                    MessageBox.Show(new Form { TopMost = true }, "Xuất báo cáo thành công", "Thông báo");
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Không có dữ liệu", "Thông báo");
                }
            }
        }

        public bool _PRIME_WITHOUT_SUS = false;
        public void export_new_viewdata3(string file_format, string report_folder, string export_path, string itemcode, string lotno, string type, DataTable Data_all, DataTable dt_spec, string sheet, DataGridView dgv_data, string leader, int count_sample, string itemCodeNVL, string lotNoNVL)
        {
            ExcelWorkbook wb_format = null;
            if (type == "NPI")
            {
                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                ExcelPackage xlPackage = open_excel(file_format);
                wb_format = xlPackage.Workbook;
            }

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
                if (!System.IO.Directory.Exists(report_folder))
                    System.IO.Directory.CreateDirectory(report_folder);

                if (type == "NPI")
                {
                    string _process = sheet.Replace("_", "").Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "").ToUpper();
                    string mySheet = "";
                    foreach (ExcelWorksheet tg_sht in wb_format.Worksheets)
                    {
                        string cur_sht_name = tg_sht.Name.ToUpper().Replace("_", "").Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "");
                        if (_process == cur_sht_name)
                        {
                            mySheet = tg_sht.Name;
                            break;
                        }
                    }

                    if (mySheet != "")
                    {
                        ExcelPackage.LicenseContext = LicenseContext.Commercial;
                        sourcePackage = new ExcelPackage();
                        report_saved = sourcePackage.Workbook;

                        ExcelWorksheet worksheet0 = wb_format.Worksheets[mySheet];
                        report_saved.Worksheets.Add(mySheet, worksheet0);

                    }
                }
                else if (type.Contains("MASS"))
                {
                    ExcelPackage.LicenseContext = LicenseContext.Commercial;
                    sourcePackage = new ExcelPackage(new FileInfo(file_format));
                    report_saved = sourcePackage.Workbook;
                }


                if (report_saved != null)
                {
                    ExcelWorksheet ws = report_saved.Worksheets[0];
                    ws.ConditionalFormatting.RemoveAll();

                    if (sheet == "GAP_CONNECTOR")
                    {
                        export_excel_gap_connector(ws, Data_all, dt_spec);
                    }
                    else if (sheet == "CROSS_SECTION")
                    {
                        export_excel_cross_section(ws, itemcode, lotno, Data_all, dt_spec, dgv_data, type, leader, count_sample);
                    }
                    else if (sheet.Contains("ON_PRODUCT"))
                    {
                        export_excel_onproduct(ws, itemcode, lotno, sheet, Data_all, dt_spec, dgv_data, type, leader, count_sample);
                    }
                    else if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST" || sheet == "SHEAR_TEST")
                    {
                        export_excel_peel_pull_shear(ws, itemcode, lotno, sheet, Data_all, dt_spec, dgv_data, type, leader, count_sample, _PRIME_WITHOUT_SUS);
                    }
                    else if (sheet == "IQC_UNMATING_PULL_TEST")
                    {
                        export_excel_unmating(ws, itemcode, lotno, sheet, Data_all, dt_spec, dgv_data, type, leader, count_sample);
                    }
                    else
                    {
                        export_excel_coupon(ws, itemcode, lotno, sheet, Data_all, dt_spec, dgv_data, type, leader, count_sample);
                        string filename = FileFolderRepository.GetFileName(export_path);
                        export_path = export_path.Replace(filename, "");
                        string exten = filename.Split('.')[1];


                        string key = $"{lotNoNVL}";
                        int pos = filename.IndexOf(key) + key.Length;
                        filename = filename.Substring(0, pos);


                        export_path += $"{filename}.{exten}";
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
                        MessageBox.Show(new Form { TopMost = true }, "Lỗi khi mở file báo cáo", "Thông báo");
                    }

                    MessageBox.Show(new Form { TopMost = true }, "Xuất báo cáo thành công", "Thông báo");
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Tên sheet chưa đúng", "Thông báo");
                }
            }
        }

    }
}

