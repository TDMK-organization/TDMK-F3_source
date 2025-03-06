using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SEEV_DLL;
using TDMK_SQL;
using myExcel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using System.IO;
using Image = System.Drawing.Image;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Data.SqlClient;
using OfficeOpenXml.FormulaParsing;
using OK2SHIP_SMT;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using static System.Net.Mime.MediaTypeNames;
using TDMK_Image;

namespace Export_FPCA_OK2ship_Auto_System
{

    public class Funtion_export_FPCA
    {
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public SEI_Lib myCode = new SEI_Lib();
        public EPPlus_Lib TDMK_Code2 = new EPPlus_Lib();



        // export cross_section
        public void export_excel_cross_section(string ItemCode, string LotNo, ExcelWorksheet ws, DataTable Data_tbl, DataTable dt_spec, ref bool export_ok)
        {

            if (Check_spec_crossection(Data_tbl, ItemCode, dt_spec))
            {
                string location_spec = dt_spec.Rows[0]["Location"].ToString();
                int col_begin = int.Parse(location_spec.Split('+')[0].Split(';')[1]);
                string[] region = location_spec.Split('+')[1].Split('_');

                List<string> lst_region = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToList();
                int no_lk = 1;
                int count_sample = int.Parse(dt_spec.Rows[0]["Count_sample"].ToString());
                foreach (string reg in region)
                {
                    if (!reg.Contains("^") && reg != "")
                    {
                        int count_row = int.Parse(reg.Split(';')[1]);
                        int row_begin = int.Parse(reg.Split(';')[2]);
                        if (reg.ToUpper().Contains("Horizontal".ToUpper()))
                        {
                            export_Hottizontal_NPI(reg, ws, row_begin, col_begin, no_lk, lst_region, count_row, Data_tbl, dt_spec);

                        }

                        if (reg.ToUpper().Contains("Vertical".ToUpper()))
                        {

                            List<string> lst_key = new List<string> { };
                            if (count_row == 32)
                            {
                                lst_key = new List<string> { "TRU_T", "TRU_P", "DOC_T", "DOC_P" };
                            }
                            else if (count_row == 48)
                            {
                                lst_key = new List<string> { "TRU1_T", "TRU1_P", "TRU2_T", "TRU2_P", "DOC_T", "DOC_P" };
                            }

                            foreach (string item in lst_key)
                            {
                                ExcelRangeBase cell_Pic1_2 = ws.Cells[row_begin, col_begin];
                                ExcelRangeBase cell_Pic2_2 = ws.Cells[row_begin + 3, col_begin];
                                ExcelRangeBase cell_data_2 = ws.Cells[row_begin + 1, col_begin];

                                string filter_reg_2 = "";

                                string txt_find = "";
                                if (reg.ToUpper().Contains("MALE"))
                                {
                                    txt_find = no_lk.ToString() + "_";
                                }

                                foreach (string i in lst_region)
                                {
                                    if (i.Replace(" ", "").ToUpper().Contains(item) && i.ToUpper().Contains(txt_find))
                                    {
                                        filter_reg_2 = i;
                                        break;
                                    }
                                }
                                if (filter_reg_2 != "")
                                {

                                    DataView dv = Data_tbl.AsDataView();
                                    dv.RowFilter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { filter_reg_2 });
                                    DataTable tbl_region = dv.ToTable();

                                    if (tbl_region.Rows.Count > 0)
                                    {
                                        //int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
                                        Export_DatatableImage_Excel(tbl_region, "Image1", ws, cell_Pic1_2, false, count_sample);
                                        Export_DatatableImage_Excel(tbl_region, "Image2", ws, cell_Pic2_2, false, count_sample);

                                        List<string> lst_data = tbl_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();

                                        for (int i = 0; i < tbl_region.Rows.Count; i++)
                                        {
                                            if (i < count_sample)
                                            {
                                                cell_data_2.Offset(0, i).Value = double.Parse(lst_data[i].Split('/')[0].Split(';')[0]);
                                                cell_data_2.Offset(0, i).Style.Numberformat.Format = "0.00";
                                                cell_data_2.Offset(1, i).Value = double.Parse(lst_data[i].Split('/')[0].Split(';')[1]);
                                                cell_data_2.Offset(1, i).Style.Numberformat.Format = "0.00";

                                                cell_data_2.Offset(3, i).Value = double.Parse(lst_data[i].Split('/')[1].Split(';')[0]);
                                                cell_data_2.Offset(3, i).Style.Numberformat.Format = "0.00";
                                                cell_data_2.Offset(4, i).Value = double.Parse(lst_data[i].Split('/')[1].Split(';')[2]);
                                                cell_data_2.Offset(4, i).Style.Numberformat.Format = "0.00";
                                                cell_data_2.Offset(5, i).Value = double.Parse(lst_data[i].Split('/')[1].Split(';')[1]);
                                                cell_data_2.Offset(5, i).Style.Numberformat.Format = "0.00";
                                                cell_data_2.Offset(6, i).Value = double.Parse(lst_data[i].Split('/')[1].Split(';')[3]);
                                                cell_data_2.Offset(6, i).Style.Numberformat.Format = "0.00";
                                            }
                                        }
                                    }
                                }
                                row_begin = row_begin + 8;
                            }
                            no_lk++;
                        }
                    }
                }


                // SortedDictionary<int, string> dic_judge = dic_judge_crosscut(count_sample, Data_tbl); 
                //for (int j = 1; j < 5; j++)
                //{
                //    for (int i = 1; i < 120; i++)
                //    {
                //        if (myCode.checkDBNull(ws.Cells[i, j].Value).ToUpper().Replace(" ", "") == "RESULT")
                //        {
                //            for (int k = 0; k < count_sample; k++)
                //            {
                //                if (dic_judge.ContainsKey(k + 1))
                //                {
                //                    ws.Cells[i, j + 2 + k] = "Fail";
                //                }
                //                else
                //                {
                //                    ws.Cells[i, j + 2 + k] = "Pass";
                //                }
                //            }
                //        return;
                //        }
                //    }
                //}
            }
            else
            {
                export_ok = false;
                MessageBox.Show(new Form { TopMost = true }, "CROSS_SECTION chứa dữ liệu NG. Không thể xuất dữ liệu vào báo cáo", "Warning");
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

        public void Export_DatatableImage_Excel_tape(DataTable dt, string Image_Col_name, ExcelWorksheet tar_wrksht, ExcelRangeBase sel_rgn, bool row_offset, int count_sample)
        {
            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = Path.Combine(file_folder, DateTime.Now.ToString("ddMMMyy_HHmmss") + ".jpg");
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
            //try
            //{
            //    File.Delete(file_dic);
            //}
            //catch
            //{ 
            //}

        }


        public void Export_DatatableImage_Excel(DataTable dt, string Image_Col_name, ExcelWorksheet tar_wrksht, ExcelRangeBase sel_rgn, bool row_offset, int count_sample)
        {
            
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

                    //int col_offset = sel_rgn.Columns.Count;
                    //int row_off = sel_rgn.Rows.Count;
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
                            cell_data.Offset(0, k).Value = double.Parse(lst_data[i].Split(';')[0]);
                            cell_data.Offset(0, k).Style.Numberformat.Format = "0.00";
                            cell_data.Offset(1, k).Value = double.Parse(lst_data[i].Split(';')[1]);
                            cell_data.Offset(1, k).Style.Numberformat.Format = "0.00";
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
                                        cell_data_2.Offset(0, k).Value = double.Parse(lst_data[i].Split(';')[0]);
                                        cell_data_2.Offset(0, k).Style.Numberformat.Format = "0.00";
                                        cell_data_2.Offset(1, k).Value = double.Parse(lst_data[i].Split(';')[1]);
                                        cell_data_2.Offset(1, k).Style.Numberformat.Format = "0.00";
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
                                        cell_data_2.Offset(0, k).Value = double.Parse(lst_data[i].Split(';')[0]);
                                        cell_data_2.Offset(0, k).Style.Numberformat.Format = "0.00";
                                        cell_data_2.Offset(1, k).Value = double.Parse(lst_data[i].Split(';')[1]);
                                        cell_data_2.Offset(1, k).Style.Numberformat.Format = "0.00";
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
                                        cell_data_2.Offset(0, k).Value = double.Parse(lst_data[i].Split(';')[0]);
                                        cell_data_2.Offset(0, k).Style.Numberformat.Format = "0.00";
                                        cell_data_2.Offset(1, k).Value = double.Parse(lst_data[i].Split(';')[1]);
                                        cell_data_2.Offset(1, k).Style.Numberformat.Format = "0.00";
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

        public void Export_DatatableImage_Excel_trungang(DataTable dt, string Image_Col_name, ExcelWorksheet tar_wrksht, ExcelRangeBase sel_rgn, bool row_offset, int count_sample, int r_offset)
        { 
            int r_inx = 0;
            //ExcelRangeBase sel_rgn = tar_wrksht.Range[tar_rgn];
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
        //    string pic_name = Path.GetFileName(picFile);

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
                decimal mdw = wrkbk.MaxFontWidth;
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
        //export GAP connector
        public void export_excel_gap_connector(ExcelWorksheet ws, DataTable Data_tbl, DataTable dt_spec, ref bool export_ok)
        {

            int col_begin = int.Parse(dt_spec.Rows[0]["Location"].ToString().Split('+')[0]);
            string[] region = dt_spec.Rows[0]["Location"].ToString().Split('+')[1].Split('_');
            int count_sample = int.Parse(dt_spec.Rows[0]["Count_sample"].ToString());

            List<string> lst_region_db = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToList();

            SortedDictionary<int, string> dic_judgement = Check_spec_GAP(Data_tbl, dt_spec, "GAP_CONNECTOR");
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
                        if (reg.Split(';')[0] == "IOPIN")
                        {
                            int code = 1;
                            foreach (string pos in reg.Split('^'))
                            {
                                if (pos != "")
                                {
                                    int row_begin = int.Parse(pos.Split(';')[1]);
                                    int count_r_offset = int.Parse(pos.Split(';')[2]);
                                    ExcelRangeBase cell_Pic = ws.Cells[row_begin, col_begin];
                                    ExcelRangeBase cell_Pic1 = ws.Cells[row_begin + 1, col_begin];
                                    ExcelRangeBase cell_Pic2 = ws.Cells[row_begin + count_r_offset + 2, col_begin];
                                    ExcelRangeBase cell_data = ws.Cells[row_begin + 2, col_begin];
                                    string find_text = "";

                                    if (reg.Split('^').Length == 3)
                                    {
                                        find_text = code.ToString();
                                    }
                                    string filter_reg = "";
                                    foreach (string i in lst_region_db)
                                    {
                                        if (i.ToUpper().Contains("DOC") && i.Contains(find_text))
                                        {
                                            filter_reg = i;
                                            break;
                                        }
                                    }
                                    // DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo, filter_reg }));

                                    if (filter_reg != "")
                                    {
                                        DataTable tbl_region = Data_tbl.AsEnumerable().Where(s => s.Field<string>("Region") == filter_reg).CopyToDataTable();
                                        if (tbl_region.Rows.Count > 0)
                                        {
                                            // int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());

                                            Export_DatatableImage_Excel(tbl_region, "Image", ws, cell_Pic, false, count_sample);
                                            Export_DatatableImage_Excel(tbl_region, "Image1", ws, cell_Pic1, false, count_sample);
                                            Export_DatatableImage_Excel(tbl_region, "Image2", ws, cell_Pic2, false, count_sample);

                                            List<string> lst_data = tbl_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                                            //List<double> lst_1 = new List<double> { };
                                            //List<double> lst_2 = new List<double> { };
                                            int offset = 0;
                                            for (int i = 0; i < count_sample; i++)
                                            {
                                                if (i < tbl_region.Rows.Count)
                                                {

                                                    int min_r = new int[] { count_r_offset, lst_data[i].Split('/')[0].Split(';').Length }.Min();
                                                    offset = min_r;
                                                    //lst_1.Add(double.Parse(lst_data[i].Split('/')[0].Split(';')[min_r - 1]));
                                                    //lst_2.Add(double.Parse(lst_data[i].Split('/')[1].Split(';')[min_r - 1]));
                                                    for (int k = 0; k < min_r; k++)
                                                    {
                                                        cell_data.Offset(k, i).Value = double.Parse(lst_data[i].Split('/')[0].Split(';')[k]);
                                                        cell_data.Offset(k, i).Style.Numberformat.Format = "0.00";

                                                        cell_data.Offset(k + count_r_offset + 1, i).Value = double.Parse(lst_data[i].Split('/')[1].Split(';')[k]);
                                                        cell_data.Offset(k + count_r_offset + 1, i).Style.Numberformat.Format = "0.00";
                                                    } 
                                                    int ID = int.Parse(tbl_region.Rows[i]["ID"].ToString());

                                                    if (dic_judgement.ContainsKey(ID - 1))
                                                    {
                                                        if (dic_judgement[ID - 1] == "FAIL")
                                                        {
                                                            cell_data.Offset(min_r + count_r_offset + 1, i).Value = "NG";

                                                        }
                                                        else
                                                        {
                                                            cell_data.Offset(min_r + count_r_offset + 1, i).Value = "OK";
                                                        }
                                                    }

                                                } 
                                            }

                                            //cell_data.Offset(offset + count_r_offset + 1, count_sample).Value = "OK";

                                            //if (lst_1.Max() - lst_1.Min() > 30 || lst_2.Max() - lst_2.Min() > 30)
                                            //{
                                            //    cell_data.Offset(offset + count_r_offset + 1, count_sample + 1).Value = "NG";
                                            //    cell_data.Offset(offset + count_r_offset + 1, count_sample + 2).Value = "NG";
                                            //}
                                            //else
                                            //{
                                            //    cell_data.Offset(offset + count_r_offset + 1, count_sample + 1).Value = "OK";
                                            //    cell_data.Offset(offset + count_r_offset + 1, count_sample + 2).Value = "OK";
                                            //} 
                                        }
                                    }

                                }
                                code++;
                            }
                        }
                        else if (reg.Split(';')[0] == "LEFT" || reg.Split(';')[0] == "RIGHT")
                        {
                            int code = 1;
                            foreach (string pos in reg.Split('^'))
                            {
                                if (pos != "")
                                {
                                    int row_begin = int.Parse(pos.Split(';')[1]);
                                    int count_r_offset = int.Parse(pos.Split(';')[2]);
                                    ExcelRangeBase cell_Pic = ws.Cells[row_begin, col_begin];
                                    ExcelRangeBase cell_Pic1 = ws.Cells[row_begin + 1, col_begin];
                                    ExcelRangeBase cell_Pic2 = ws.Cells[row_begin + count_r_offset + 2, col_begin];
                                    ExcelRangeBase cell_data = ws.Cells[row_begin + 2, col_begin];

                                    string find_text = "";

                                    if (reg.Split('^').Length == 3)
                                    {
                                        find_text = code.ToString();
                                    }
                                    string filter_reg = "";
                                    foreach (string i in lst_region_db)
                                    {
                                        if (i.ToUpper().Contains("TRU") && i.Contains(find_text))
                                        {
                                            filter_reg = i;
                                            break;
                                        }
                                    }
                                    //DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo, filter_reg }));
                                    if (filter_reg != "")
                                    {
                                        DataTable tbl_region = Data_tbl.AsEnumerable().Where(s => s.Field<string>("Region") == filter_reg).CopyToDataTable();
                                        if (tbl_region.Rows.Count > 0)
                                        {
                                            //int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString()); 
                                            if (reg.Contains("RIGHT"))
                                            {
                                                for (int r = 0; r < count_sample; r++)
                                                {
                                                    tbl_region.Rows.Remove(tbl_region.Rows[0]);
                                                }
                                            }

                                            Export_DatatableImage_Excel(tbl_region, "Image", ws, cell_Pic, false, count_sample);
                                            Export_DatatableImage_Excel(tbl_region, "Image1", ws, cell_Pic1, false, count_sample);
                                            Export_DatatableImage_Excel(tbl_region, "Image2", ws, cell_Pic2, false, count_sample);


                                            List<string> lst_data = tbl_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                                            List<double> lst_1 = new List<double> { };
                                            List<double> lst_2 = new List<double> { };
                                            int offset = 0;
                                            for (int i = 0; i < count_sample; i++)
                                            {
                                                if (i < tbl_region.Rows.Count)
                                                {
                                                    int min_r = new int[] { count_r_offset, lst_data[i].Split('/')[0].Split(';').Length }.Min();
                                                    offset = min_r;
                                                    lst_1.Add(double.Parse(lst_data[i].Split('/')[0].Split(';')[min_r - 1]));
                                                    lst_2.Add(double.Parse(lst_data[i].Split('/')[1].Split(';')[min_r - 1]));

                                                    for (int k = 0; k < min_r; k++)
                                                    {
                                                        cell_data.Offset(k, i).Value = double.Parse(lst_data[i].Split('/')[0].Split(';')[k]);
                                                        cell_data.Offset(k, i).Style.Numberformat.Format = "0.00";
                                                        cell_data.Offset(k + count_r_offset + 1, i).Value = double.Parse(lst_data[i].Split('/')[1].Split(';')[k]);
                                                        cell_data.Offset(k + count_r_offset + 1, i).Style.Numberformat.Format = "0.00";
                                                    }


                                                    int ID = int.Parse(tbl_region.Rows[i]["ID"].ToString());

                                                    if (dic_judgement.ContainsKey(ID - 1))
                                                    {
                                                        if (dic_judgement[ID - 1] == "FAIL")
                                                        {
                                                            cell_data.Offset(min_r + count_r_offset + 1, i).Value = "NG";

                                                        }
                                                        else
                                                        {
                                                            cell_data.Offset(min_r + count_r_offset + 1, i).Value = "OK";
                                                        }
                                                    }

                                                }
                                            }
                                            cell_data.Offset(offset + count_r_offset + 1, count_sample).Value = "OK";

                                            if (lst_1.Max() - lst_1.Min() > 30 || lst_2.Max() - lst_2.Min() > 30)
                                            {
                                                cell_data.Offset(offset + count_r_offset + 1, count_sample + 1).Value = "NG";
                                                cell_data.Offset(offset + count_r_offset + 1, count_sample + 2).Value = "NG";
                                            }
                                            else
                                            {
                                                cell_data.Offset(offset + count_r_offset + 1, count_sample + 1).Value = "OK";
                                                cell_data.Offset(offset + count_r_offset + 1, count_sample + 2).Value = "OK";
                                            }
                                        }
                                    }
                                }
                                code++;
                            }
                        }
                    }
                }
            }
            else
            {
                export_ok = false;
                MessageBox.Show(new Form { TopMost = true }, "GAP_CONNECTOR chứa dữ liệu NG. Không thể xuất dữ liệu vào báo cáo", "Warning");
            }



        }  

        // export onproduct 
        public void export_excel_onproduct(ExcelWorksheet ws, DataTable Data_all, DataTable dt_spec, string sheet, ref bool export_ok)
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

        public void export_excel_peel_pull_shear(ExcelWorksheet ws, string sheet, DataTable Data_tbl, DataTable dt_spec, ref bool export_ok)
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
        public void export_excel_coupon(ExcelWorksheet ws, DataTable Data_all, DataTable dt_spec, string mysheet, ref bool export_ok)
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

                                            // ExcelRange cell_min = ws.Cells[int.Parse(str_location[3].Split(';')[1]) + i, int.Parse(str_location[1].Split(';')[3]) + 1];

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


        // export unmating 
        public void export_excel_unmating(ExcelWorksheet ws, DataTable Data_tbl, DataTable dt_spec, ref bool export_ok)
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
                                    //cell_data.Offset(0, i).Value = lst_data[i];
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

            DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet","Remark" }, new string[] { ItemCode, "ACF","NPI" }));
            if (dt_spec.Rows.Count > 0)
            {
                int count_sample = int.Parse(dt_spec.Rows[0]["Count_sample"].ToString());
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet" }, new string[] { ItemCode, "ACF" });

                DataTable Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, "ACF_WETTING", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                if (Data_tbl.Rows.Count > 0)
                {
                    List<string> lst_machine = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Machine")).Distinct().ToList();
                    int count_machine_DB = lst_machine.Count();
                    int count_machine_format = dt_spec.Rows[0]["Location"].ToString().Split('_').Length - 1;

                    int r_begin = int.Parse(dt_spec.Rows[0]["Location"].ToString().Split('_')[0].Split(';')[0]);
                    int c_begin = int.Parse(dt_spec.Rows[0]["Location"].ToString().Split('_')[0].Split(';')[1]);

                    // Range line = (Range)ws.Row[r_begin];
                    ws.InsertRow(r_begin, 1);
                    if (count_machine_DB > count_machine_format)
                    {
                        insert_row_excel(ws, r_begin, r_begin + 6);
                        ws.InsertRow(r_begin, 7);
                    }
                    foreach (string machine in lst_machine)
                    {

                        Dictionary<int, List<Double>> dic_list = new Dictionary<int, List<double>> { };
                        List<double> lst_after_plasma_ACF = new List<double> { };
                        List<double> lst_after_plasma_GND = new List<double> { };
                        List<double> lst_Before_Packing_ACF = new List<double> { };
                        List<double> lst_Before_Packing_GND = new List<double> { };


                        ws.Cells[r_begin, 1].Value = "Result data machine " + machine;

                        DataTable tbl_machine = TDMK_Code.Datatable_Filter(sqlcon, "ACF_WETTING", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Machine" }, new string[] { ItemCode, LotNo, machine }));
                        int col_data = c_begin + 9;
                        for (int i = 0; i < tbl_machine.Rows.Count; i = i + 2)
                        {
                            if (col_data < col_data + count_sample)
                            {

                                if (myCode.checkDBNull(tbl_machine.Rows[i]["After_Plasma"]) != "")
                                {
                                    ws.Cells[r_begin + 2, col_data].Value = double.Parse(tbl_machine.Rows[i]["After_Plasma"].ToString());
                                    ws.Cells[r_begin + 2, col_data].Style.Numberformat.Format = "0.00";
                                    lst_after_plasma_ACF.Add(double.Parse(myCode.checkDBNull(myCode.checkDBNull(tbl_machine.Rows[i]["After_Plasma"]))));
                                }

                                if (myCode.checkDBNull(tbl_machine.Rows[i + 1]["After_Plasma"]) != "")
                                {
                                    ws.Cells[r_begin + 3, col_data].Value = double.Parse(tbl_machine.Rows[i + 1]["After_Plasma"].ToString());
                                    ws.Cells[r_begin + 3, col_data].Style.Numberformat.Format = "0.00";
                                    lst_after_plasma_GND.Add(double.Parse(myCode.checkDBNull(tbl_machine.Rows[i + 1]["After_Plasma"])));
                                }

                                if (myCode.checkDBNull(tbl_machine.Rows[i]["Before_Packing"]) != "")
                                {
                                    ws.Cells[r_begin + 4, col_data].Value = double.Parse(tbl_machine.Rows[i]["Before_Packing"].ToString());
                                    ws.Cells[r_begin + 4, col_data].Style.Numberformat.Format = "0.00";
                                    lst_Before_Packing_ACF.Add(double.Parse(myCode.checkDBNull(tbl_machine.Rows[i]["Before_Packing"])));
                                }

                                if (myCode.checkDBNull(tbl_machine.Rows[i + 1]["Before_Packing"]) != "")
                                {
                                    ws.Cells[r_begin + 5, col_data].Value = double.Parse(tbl_machine.Rows[i + 1]["Before_Packing"].ToString());
                                    ws.Cells[r_begin + 5, col_data].Style.Numberformat.Format = "0.00";
                                    lst_Before_Packing_GND.Add(double.Parse(myCode.checkDBNull(tbl_machine.Rows[i + 1]["Before_Packing"])));
                                }

                                col_data++;
                            }
                        }
                        dic_list.Add(2, lst_after_plasma_ACF);
                        dic_list.Add(3, lst_after_plasma_GND);
                        dic_list.Add(4, lst_Before_Packing_ACF);
                        dic_list.Add(5, lst_Before_Packing_GND);


                        foreach (var lst in dic_list)
                        {
                            string max = lst.Value.Max().ToString();
                            string min = lst.Value.Min().ToString();
                            string Average = Math.Round(lst.Value.Average(), 2).ToString();
                            double STDEV = CalculateStandardDeviation(lst.Value);

                            string _spec = ws.Cells[r_begin + lst.Key, c_begin + 2].Value.ToString();
                            double Mean = double.Parse(_spec.Replace("<", "").Replace(">", "").Replace("°", "").Replace(" ", ""));
                            double cpk = Math.Round((Mean - lst.Value.Average()) / (3 * STDEV), 2);

                            ws.Cells[r_begin + lst.Key, c_begin + 3].Value = min;
                            ws.Cells[r_begin + lst.Key, c_begin + 4].Value = max;
                            ws.Cells[r_begin + lst.Key, c_begin + 5].Value = Average;
                            ws.Cells[r_begin + lst.Key, c_begin + 6].Value = STDEV;
                            ws.Cells[r_begin + lst.Key, c_begin + 7].Value = cpk;
                            if (cpk > 1.33)
                            {
                                ws.Cells[r_begin + lst.Key, c_begin + 8].Value = "OK";
                            }
                            else
                            {
                                ws.Cells[r_begin + lst.Key, c_begin + 8].Value = "NG";
                            }

                        }

                        r_begin = r_begin + 7;
                    }
                }
            }


        }

        public void Export_ACF_Peel(ExcelWorksheet ws, string ItemCode, string LotNo, SqlConnection sqlcon)
        {
            DataTable dt_analysis = TDMK_Code.Datatable_Filter(sqlcon, "ACF_BONDING", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
            DataTable tbl_Data = new DataTable();
            if (dt_analysis.Rows.Count > 0)
            {
                string[] arr_val = dt_analysis.AsEnumerable().Select(x => x.Field<string>("Remark")).Distinct().ToArray();
                string filter = "";
                foreach (string item in arr_val)
                {
                    if (item.Contains("NPI"))
                    {
                        filter = item;
                        tbl_Data = dt_analysis.AsEnumerable().Where(r => r.Field<string>("Remark") == filter).CopyToDataTable();
                        break;
                    }
                }
            }
            if (tbl_Data.Rows.Count > 0)
            {
                ExcelRangeBase sel_rgn = ws.Cells[1, 1];
                string find_addr = Find_addr_peel("Before tape test", sel_rgn);
                if (find_addr == "")
                    find_addr = Find_addr_peel("Glass coupon", sel_rgn);
                //string find_addr = Find_addr_peel("ACF Bonding", sel_rgn);
                if (find_addr != "")
                {

                    ExcelRangeBase rgn_ = ws.Cells[find_addr];
                    // ExcelRangeBase cur_rgn = ws.Cells[find_addr].Offset(1, 1);
                    ExcelRangeBase cur_rgn = ws.Cells[find_addr].Offset(0, 1);
                    //if (myCode.checkDBNull(rgn_.Offset(1, 0).Value).Contains("Product"))
                    //{
                    //    //cur_rgn = cur_rgn.Offset(1, 0);
                    //    cur_rgn = cur_rgn.Offset(0, 1);
                    //}



                    if (cur_rgn != null)
                    {
                        Export_DatatableImage_Excel_ACF(tbl_Data, "Image_Before", ws, cur_rgn, false);
                        Export_DatatableImage_Excel_ACF(tbl_Data, "Image_After", ws, cur_rgn.Offset(1, 0), false);
                        Export_DatatableImage_Excel_Graph_ACF(tbl_Data, "Graph", ws, cur_rgn.Offset(2, 0), false);
                        List<string> lst_data = tbl_Data.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                        for (int i = 0; i < tbl_Data.Rows.Count; i++)
                        {
                            cur_rgn.Offset(3, i).Value = lst_data[i];
                        }


                        int count_sample = tbl_Data.Rows.Count;

                        //List<string> lst_judge = new List<string> { };
                        //for (int i = 0; i < count_sample; i++)
                        //{

                        //    if (dgv_peel.Rows[i].Cells["Data"].Style.BackColor == Color.Red)
                        //    {
                        //        lst_judge.Add("Fail");
                        //        judgeall_peel = false;

                        //    }
                        //    else
                        //    {
                        //        lst_judge.Add("Pass");
                        //    }
                        //}


                        for (int i = 1; i < 13; i++)
                        {
                            if (myCode.checkDBNull(cur_rgn.Offset(i, -1).Value).Replace(" ", "").ToUpper().Contains("MINFORCE"))
                            {
                                ExcelRangeBase cell_min = cur_rgn.Offset(i, 0);
                                ExcelRangeBase cell_max = cur_rgn.Offset(i + 1, 0);
                                ExcelRangeBase cell_ave = cur_rgn.Offset(i + 2, 0);
                                int c_offset = count_sample;
                                cell_min.FormulaR1C1 = "=MIN(R[-" + (i - 3).ToString() + "]C[0]:R[-" + (i - 3).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                cell_max.FormulaR1C1 = "=MAX(R[-" + (i - 2).ToString() + "]C[0]:R[-" + (i - 2).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                cell_ave.FormulaR1C1 = "=AVERAGE(R[-" + (i - 1).ToString() + "]C[0]:R[-" + (i - 1).ToString() + "]C[" + (c_offset - 1).ToString() + "])";
                                break;
                            }
                        }


                    }

                }
            }


        }

        public void Export_ACFFlatness(ExcelWorksheet ws, string ItemCode, string LotNo, SqlConnection sqlcon)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, "ACF_FLATNESS", filter_str);


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
                                //if (myCode.checkDBNull(curr_rgn.Offset(1, c).Offset(-1, 0).Value).Contains("Point"))
                                //{
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
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, "Roughness", filter_str);


            if (src_dt.Rows.Count > 0)
            {

                List<List<string>> lst_rghness = lst_roughness(src_dt);
                ExcelRangeBase curr_rgn_1 = ws.Cells[1, 1];


                List<int> vitri_sa = new List<int> { };
                List<int> vitri_sq = new List<int> { };
                List<int> vitri_sdr = new List<int> { };
                string[] arr_info = ACF_pad_location(ws).Split('+');
                string vitri_ACF_pad = arr_info[0];

                int row = int.Parse(arr_info[0]);
                int r_count = int.Parse(arr_info[arr_info.Length - 1]);
                curr_rgn_1 = curr_rgn_1.Offset(row, 0);

                for (int i = 1; i < arr_info.Length; i++)
                {
                    if (arr_info[i].Contains("sa"))
                    {
                        vitri_sa.Add(i);
                    }
                    if (arr_info[i].Contains("sq"))
                    {
                        vitri_sq.Add(i);
                    }
                    if (arr_info[i].Contains("sdr"))
                    {
                        vitri_sdr.Add(i);
                    }
                }

                int m = 0;
                int k;
                while (m < 32)
                {
                    for (int c = 0; c < 4; c++)
                    {
                        for (int i = 2; i <= 9; i++)
                        {
                            k = 0;
                            foreach (int j in vitri_sa)
                            {
                                if (myCode.IsNumeric(lst_rghness[k][m]))
                                {
                                    curr_rgn_1.Offset(j, i).Value = double.Parse(lst_rghness[k][m]);
                                    curr_rgn_1.Offset(j, i).Style.Numberformat.Format = "0.00";
                                }
                                else
                                {
                                    curr_rgn_1.Offset(j, i).Value = lst_rghness[k][m];
                                }

                                k += 3;
                            }
                            k = 1;
                            foreach (int j in vitri_sq)
                            {
                                if (myCode.IsNumeric(lst_rghness[k][m]))
                                {
                                    curr_rgn_1.Offset(j, i).Value = double.Parse(lst_rghness[k][m]);
                                    curr_rgn_1.Offset(j, i).Style.Numberformat.Format = "0.00";
                                }
                                else
                                {
                                    curr_rgn_1.Offset(j, i).Value = lst_rghness[k][m];

                                }
                                //curr_rgn_1.Offset(j, i).Value = lst_rghness[k][m];
                                k += 3;
                            }
                            k = 2;
                            foreach (int j in vitri_sdr)
                            {
                                if (myCode.IsNumeric(lst_rghness[k][m]))
                                {
                                    curr_rgn_1.Offset(j, i).Value = double.Parse(lst_rghness[k][m]);
                                    curr_rgn_1.Offset(j, i).Style.Numberformat.Format = "0.00";
                                }
                                else
                                {
                                    curr_rgn_1.Offset(j, i).Value = lst_rghness[k][m];
                                }
                                //curr_rgn_1.Offset(j, i).Value = lst_rghness[k][m];
                                k += 3;
                            }
                            m++;
                        }

                        curr_rgn_1 = curr_rgn_1.Offset(r_count, 0);

                    }
                }





            }

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
                                            //dgv.Rows[r_x].Cells["Data"].Style.BackColor = Color.Red;
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
    }

}

