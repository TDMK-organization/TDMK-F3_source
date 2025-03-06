using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AForge.Video;
using AForge.Video.DirectShow;
using TDMK_SQL;
using System.Data.SqlClient;
using System.Diagnostics.Tracing;
using System.Runtime.InteropServices;
using System.IO;
using TDMK_SEEV_DLL;
using System.Collections;
using DataTable = System.Data.DataTable;
using myExcel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using ACF_Process;
using System.Diagnostics;
using System.Security.Policy;
//using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;


//using TDMK_SEEV_DLL;
namespace OK2SHIP
{
    public class Export_Image_Class
    {
        // public TDMK_Class TDMK_Code = new TDMK_Class();
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SEI_Lib myCode = new SEI_Lib();
        ACF_Process_Lib ACF_Proc = new ACF_Process_Lib();
        SqlConnection sqlcon;
        public myExcel.Range sel_rgn;
        public System.Data.DataTable dt;
        public static byte[] imgToByteConverter(Image inImg)
        {
            ImageConverter imgCon = new ImageConverter();
            return (byte[])imgCon.ConvertTo(inImg, typeof(byte[]));
        }

        public Image byteArrayToImage(byte[] byteArrayIn)
        {
            using (MemoryStream mStream = new MemoryStream(byteArrayIn))
            {
                return Image.FromStream(mStream);
            }
        }

        public void ConnectionDB()
        {
            string app_path = System.Windows.Forms.Application.StartupPath;
            string file_config = app_path + @"\" + "config.txt";
            string server = File.ReadLines(file_config).Skip(0).First().Remove(0, 7);
            string user = File.ReadLines(file_config).Skip(1).First().Remove(0, 8);
            string pass = File.ReadLines(file_config).Skip(2).First().Remove(0, 9);
            string database = File.ReadLines(file_config).Skip(3).First().Remove(0, 8);
            string str_sqlcon = TDMK_Code.data_connection(server, database, user, pass).ConnectionString;
            sqlcon = new SqlConnection(str_sqlcon);
        }
        public void Move_Image_To_Excel(string source_folder_image, string des_excel_file_address, string sheet_name_target, string mode_move, int number_image, string start_range, string typy_of_image)
        {
            string des_folder = Path.GetDirectoryName(des_excel_file_address);
            string des_excel_filename = Path.GetFileName(des_excel_file_address);
            myExcel.Workbook wb2 = TDMK_Code.open_excel_file(des_folder, des_excel_filename, "");
            myExcel.Worksheet ws2 = wb2.Sheets[sheet_name_target];
            myExcel.Range curr_rgn = ws2.Range[start_range];
            // Xuất ảnh
            for (int j = 1; j <= number_image; j++)
            {
                string file_name = j + "." + typy_of_image;
                InsertPicture_Name(ws2, curr_rgn, Path.Combine(source_folder_image, file_name), 5);
                if (mode_move == "ngang")
                {
                    curr_rgn = curr_rgn.Offset[0, 1];
                }
                else
                {
                    curr_rgn = curr_rgn.Offset[1, 0];
                }
            }
        }
        public void InsertPicture_Name(myExcel.Worksheet tar_wrksht, myExcel.Range tar_range, string picFile, int margin)
        {
            float left;
            float top;
            float width;
            float height;
            string pic_name = Path.GetFileName(picFile);

            if (tar_range.MergeCells)
            {
                myExcel.Range refer_range = tar_range.MergeArea;//Range["G16"];
                left = (float)(tar_range.Left) + margin;
                top = (float)(tar_range.Top) + margin;
                width = (float)(refer_range.Width) - 2 * margin;
                height = (float)(refer_range.Height) - 2 * margin;
                //width = (float)(refer_range.Width * 2.8) + 2 * margin;
                //height = (float)(refer_range.Height * 8.9) + 2 * margin;
            }
            else
            {
                left = (float)tar_range.Left + margin;
                top = (float)tar_range.Top + margin;
                width = (float)tar_range.Width - 2 * margin;
                height = (float)tar_range.Height - 2 * margin;
            }
            myExcel.Shape sel_picture = tar_wrksht.Shapes.AddPicture(picFile, MsoTriState.msoFalse, MsoTriState.msoTrue, left, top, width, height);
            //tar_wrksht.Paste(tar_range, sel_picture);
            sel_picture.LockAspectRatio = MsoTriState.msoTrue;
            sel_picture.Placement = myExcel.XlPlacement.xlMoveAndSize;
            sel_picture.Name = pic_name;

        }
        public void Clear_DGV(DataGridView DGV)
        {
            if (DGV.DataSource == null)
            {
                DGV.Rows.Clear();
            }
            else
            {
                while (DGV.Rows.Count > 0)
                {
                    DGV.Rows.RemoveAt(0);
                }
            }
        }
        public System.Data.DataTable GetContentAsDataTable(DataGridView dgv)
        {
            try
            {
                if (dgv.ColumnCount == 0) return null;
                System.Data.DataTable dtSource = new System.Data.DataTable();
                foreach (DataGridViewColumn col in dgv.Columns)
                {
                    if (col.Name == string.Empty) continue;
                    if (col.CellType.Name.Contains("Image"))
                    {
                        dtSource.Columns.Add(col.Name, System.Type.GetType("System.Byte[]"));
                    }
                    else
                    {
                        dtSource.Columns.Add(col.Name);
                    }
                    dtSource.Columns[col.Name].Caption = col.HeaderText;
                }
                if (dtSource.Columns.Count == 0) return null;
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    DataRow drNewRow = dtSource.NewRow();
                    foreach (DataColumn col in dtSource.Columns)
                    {
                        if (col.DataType == System.Type.GetType("System.Byte[]"))
                        {
                            drNewRow[col.ColumnName] = imgToByteConverter((Bitmap)row.Cells[col.ColumnName].Value);
                        }
                        else
                        {
                            drNewRow[col.ColumnName] = row.Cells[col.ColumnName].Value;
                        }
                    }
                    dtSource.Rows.Add(drNewRow);
                }
                return dtSource;
            }
            catch
            {
                return null;
            }
        }
        public void DGV_Image_To_Excel(System.Data.DataTable dt, string Image_Col_name, myExcel.Worksheet tar_wrksht, string tar_rgn, string mode_move, int number_image)
        {
            int r_inx = 0;
            int x = number_image;
            DataRow dr;
            for (int i = 0; i < x; i++)
            {
                dr = dt.Rows[i];
                //byte[] img_byte = Encoding.ASCII.GetBytes(dr[Image_Col_name].ToString());

                byte[] img_byte = (byte[])(dr[Image_Col_name]);

                Image temp = byteArrayToImage(img_byte);
                Clipboard.SetDataObject(temp, false);

                if (mode_move == "ngang")
                {
                    sel_rgn = tar_wrksht.Range[tar_rgn].Offset[0, r_inx];
                }
                else
                {
                    sel_rgn = tar_wrksht.Range[tar_rgn].Offset[r_inx, 0];
                }
                int shape_count = tar_wrksht.Shapes.Count;
                tar_wrksht.Paste(sel_rgn, temp);
                myExcel.Shape sel_pic = tar_wrksht.Shapes.Item(shape_count + 1);
                //sel_pic.Name = dr["Image_Name"].ToString();
                sel_pic.LockAspectRatio = MsoTriState.msoFalse;
                float pic_ratio = sel_pic.Width / sel_pic.Height;
                float rgn_ratio = (float)sel_rgn.Width / (float)sel_rgn.RowHeight;
                var rate = pic_ratio / rgn_ratio;
                float margin = (float)(sel_rgn.RowHeight * 0.05);
                if (rate > 1)
                {
                    sel_pic.Width = (float)sel_rgn.Width - 2 * margin;
                    sel_pic.Height = (float)(sel_rgn.Width / rgn_ratio) - 2 * margin;
                }
                else
                {
                    sel_pic.Height = (float)(sel_rgn.Height) - 2 * margin;
                    sel_pic.Width = (float)(sel_rgn.Height * rgn_ratio) - 2 * margin;
                }
                sel_pic.Top = (float)sel_rgn.Top + margin;
                sel_pic.Left = (float)sel_rgn.Left + margin;
                sel_pic.LockAspectRatio = MsoTriState.msoTrue;
                r_inx++;
            }
        }


        public System.Data.DataTable Load_image_from_db_with_region(string Itemcode, string Lotno, string region, string table_name, SqlConnection sqlcon, DataGridView DGV)
        {
            try
            {
                string[] item = { "ItemCode", "LotNo", "Region" };
                string[] item_val = { Itemcode, Lotno, region };
                dt = TDMK_Code.Datatable_Filter(sqlcon, table_name, TDMK_Code.filter_str(item, item_val));
            }
            catch (Exception)
            {
                MessageBox.Show("Không tìm thấy dữ liệu!", "Chú ý!");
            }

            return dt;
        }
        public void Load_image_from_db(string Itemcode, string Lotno, string table_name, SqlConnection sqlcon, DataGridView DGV)
        {
            DataGridView DGV_Image = DGV;
            if (DGV_Image.DataSource == null)
            {
                DGV_Image.Rows.Clear();
                DGV_Image.Columns.Clear();
            }
            //string[] item = { "ItemCode", "LotNo" };
            //string[] item_val = { tb_itemcode.Text, tb_lotno.Text };
            try
            {
                string[] item = { "ItemCode", "LotNo" };
                string[] item_val = { Itemcode, Lotno };
                DGV_Image.DataSource = TDMK_Code.Datatable_Filter(sqlcon, table_name, TDMK_Code.filter_str(item, item_val));
            }
            catch (Exception)
            {
                MessageBox.Show("Không tìm thấy dữ liệu!", "Chú ý!");
            }
        }

        public void Load_data_from_DB(SqlConnection sqlcon, DataGridView DGV, string Itemcode, string Lotno, string tablename)
        {
            DataGridView DGV_Image = DGV;
            if (DGV_Image.DataSource == null)
            {
                DGV_Image.Rows.Clear();
                DGV_Image.Columns.Clear();
            }
            //string[] item = { "ItemCode", "LotNo" };
            //string[] item_val = { tb_itemcode.Text, tb_lotno.Text };
            try
            {
                string[] item = { "ItemCode", "LotNo" };
                string[] item_val = { Itemcode, Lotno };
                DGV_Image.DataSource = TDMK_Code.Datatable_Filter(sqlcon, tablename, TDMK_Code.filter_str(item, item_val));
            }
            catch (Exception)
            {
                MessageBox.Show("Check Connection!", "Chú ý!");
            }
            if (DGV.Rows.Count < 2)
            {
                MessageBox.Show("No data!");
            }
        }
        // High speed ball shear
        //public TDMK_Class tdmk_code = new TDMK_Class();

        myExcel.Application xlApp = null;
        myExcel.Workbooks workbooks = null;
        myExcel.Workbook workbook = null;
        Hashtable sheets;
        public string xlFilePath;

        public void OpenExcel(string xlFilePath)
        {
            xlApp = new myExcel.Application();
            workbooks = xlApp.Workbooks;
            workbook = workbooks.Open(xlFilePath);
            sheets = new Hashtable();
            int count = 1;
            // Storing worksheet names in Hashtable.
            foreach (myExcel.Worksheet sheet in workbook.Sheets)
            {
                sheets[count] = sheet.Name;
                count++;
            }
        }
        public void CloseExcel(string xlFilePath)
        {
            workbook.Close(false, xlFilePath, null); // Close the connection to workbook
            Marshal.FinalReleaseComObject(workbook); // Release unmanaged object references.
            workbook = null;

            workbooks.Close();
            Marshal.FinalReleaseComObject(workbooks);
            workbooks = null;

            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);
            xlApp = null;
        }

        public void Export_ACF_Image(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            myExcel.Worksheet ws = wb.Sheets["ACF"];
            List<string> search_item = new List<string>() { "Photo before cleaning", "Photo after cleaning", "Photo after OQC testing", "Before tape test", "After tape test", "Tape test picture" };//
            myExcel.Range sel_rgn = ws.Range["A16"];
            foreach (var item in search_item)
            {
                string tar_table = "";
                string tar_col_data = "Image_Data";
                if (item.Contains("cleaning") || (item.Contains("OQC")))
                {
                    tar_table = "ACF_CLEANING_IMAGE";
                }
                else
                {
                    if (item.Contains("picture"))
                    {
                        tar_table = "ACF_GRAPH_FORCE_IMAGE";
                        tar_col_data = "Image_Graph";
                    }
                    else
                    {
                        tar_table = "ACF_" + item.Trim().Replace(" ", "_").ToUpper() + "_IMAGE";
                    }
                }
                myExcel.Range data_rgn = ws.Range[Find_addr(item, sel_rgn)].Offset[0, 1];
                DataTable data_tbl = TDMK_Code.Datatable_Filter(sqlcon, tar_table, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { Itemcode, Lotno }));
                List<DataTable> src_lst_tbl = new List<DataTable>();
                Get_ListTable(-1, data_tbl, new string[] { "ItemCode", "LotNo", "Region" }, ref src_lst_tbl, tar_col_data);
                foreach (var dt in src_lst_tbl)
                {
                    if (dt.Rows.Count > 0)
                    {
                        string region = dt.Rows[0]["Region"].ToString();
                        int off_set = 0;
                        switch (region)
                        {
                            case "Picture":
                                off_set = 0;
                                break;
                            case "TRAI":
                                off_set = 0;
                                break;
                            case "GIUA":
                                off_set = 1;
                                break;
                            case "PHAI":
                                off_set = 2;
                                break;
                            default:
                                off_set = 0;
                                break;
                        }
                        int qty = Math.Min(dt.Rows.Count, 5);
                        for (int i = 0; i < qty; i++)
                        {
                            myExcel.Range cur_rgn = data_rgn.Offset[0, i + 5 * off_set];
                            byte[] data = (byte[])dt.Rows[i][tar_col_data];
                            using (MemoryStream ms = new MemoryStream(data))
                            {
                                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                                string file_dic = Path.Combine(file_folder, "1.jpg");
                                Image image = Image.FromStream(ms);
                                image.Save(file_dic);
                                InsertPicture_Name(ws, cur_rgn, file_dic, 10);
                            }
                            if (tar_col_data.Contains("Graph"))
                            {
                                myExcel.Range force_rgn = data_rgn.Offset[1, i + 5 * off_set];
                                force_rgn.Value = dt.Rows[i]["Force_Data"];
                            }
                        }
                    }
                }
                sel_rgn = data_rgn.Offset[0, -1];
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
                        temp_dr = myDt.AsEnumerable().Where(x => x.Field<byte[]>(tar_item) != null).ToArray();
                        if (temp_dr.Length > 0)
                        {
                            src_lst_tbl.Add(myDt);
                        }
                    }
                }
            }
            else
            {
                temp_dr = myDt.AsEnumerable().Where(x => x.Field<byte[]>(tar_item) != null).ToArray();
                if (temp_dr.Length > 0)
                {
                    src_lst_tbl.Add(myDt);
                }
            }
        }
        public string Find_addr(string in_item, myExcel.Range start_rgn)
        {
            string result = start_rgn.AddressLocal.ToString();
            int r_inx = 0;
            while (r_inx < 30)
            {
                if (myCode.checkDBNull(start_rgn.Offset[r_inx, 0].Value) == in_item)
                {
                    result = start_rgn.Offset[r_inx, 0].AddressLocal;
                    break;
                }
                r_inx++;
            }
            return result;
        }
        public void Export_StackUp_Image(SqlConnection sqlcon, myExcel.Workbook wb, string excel_file_addr, string Itemcode, string Lotno)
        {
            AutoCompleteStringCollection list = Get_zone_infor_to_input_image(excel_file_addr, "Stack-up");
            //for (int i = 0; i < list.Count; i++)
            //{
            //    //listBox1.Items.Add(list[i]);
            //}
            wb = TDMK_Code.open_excel_file(Path.GetDirectoryName(excel_file_addr), Path.GetFileName(excel_file_addr), "");
            myExcel.Worksheet ws = wb.Sheets["Stack-up"];
            foreach (string x in list)
            {
                string region = x.Split('+')[0];
                if (x.Contains("H_X"))
                {
                    region = region;
                }
                else
                    region = "Zone_" + region;

                string row_val = x.Split('+')[1];
                int number_image = 5;
                string[] item = { "ItemCode", "LotNo", "Region" };
                string[] item_val = { Itemcode, Lotno, region };
                System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "STACKUP_IMAGE", TDMK_Code.filter_str(item, item_val));

                if (data_image.Rows.Count > 0)
                {
                    if (data_image.Rows.Count < number_image)
                    {
                        number_image = data_image.Rows.Count;

                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        string file_dic = Path.Combine(file_folder, "1.jpg");

                        myExcel.Range curr_rgn = ws.Range["F" + row_val];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                            MemoryStream ms = new MemoryStream(data);
                            Image image = Image.FromStream(ms);
                            image.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn, file_dic, 6);
                            curr_rgn = curr_rgn.Offset[0, 1];
                        }
                    }

                    else
                    {
                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        string file_dic = Path.Combine(file_folder, "1.jpg");

                        myExcel.Range curr_rgn = ws.Range["F" + row_val];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                            MemoryStream ms = new MemoryStream(data);
                            Image image = Image.FromStream(ms);
                            image.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                            curr_rgn = curr_rgn.Offset[0, 1];
                        }
                    }
                }
            }
        }

        public void Export_Solder_Image(SqlConnection sqlcon, myExcel.Worksheet ws, string Itemcode, string Lotno)
        {
            string[] xx = new string[4] { "BGA", "Hot Bar", "Connector", "Trace to trace" };
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
                        System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "SOLDERMASK_IMAGE", TDMK_Code.filter_str(item, item_val));
                        if (data_image.Rows.Count > 0)
                        {
                            if (data_image.Rows.Count < number_image)
                            {
                                number_image = data_image.Rows.Count;

                                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                                string file_dic = Path.Combine(file_folder, "1.jpg");

                                myExcel.Range curr_rgn = ws.Range["B" + row_val];
                                for (int m = 0; m < number_image; m++)
                                {
                                    byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                    MemoryStream ms = new MemoryStream(data);
                                    Image image = Image.FromStream(ms);
                                    image.Save(file_dic);
                                    InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                                    curr_rgn = curr_rgn.Offset[0, 1];
                                }
                            }
                            else
                            {
                                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                                string file_dic = Path.Combine(file_folder, "1.jpg");

                                myExcel.Range curr_rgn = ws.Range["B" + row_val];
                                for (int m = 0; m < number_image; m++)
                                {
                                    byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                    MemoryStream ms = new MemoryStream(data);
                                    Image image = Image.FromStream(ms);
                                    image.Save(file_dic);
                                    InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                                    curr_rgn = curr_rgn.Offset[0, 1];
                                }
                            }
                        }
                    }
                }

            }
        }
        public void Export_High_Speed_Ball_Shear_image(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            myExcel.Worksheet ws;
            try
            {
                ws = wb.Sheets["CQRA - High speed ball shear"];
            }
            catch (Exception e)
            {
                ws = wb.Sheets["CQRA - High speed ball shear "];
            }
            string[] xx = new string[2] { "Photo", "Graph" };
            int row_val;
            foreach (string zone in xx)
            {
                for (int x = 1; x < 40; x++)
                {
                    if (ws.Cells[x, 1].Value != null && ws.Cells[x, 1].Value.ToString() == zone)
                    {
                        row_val = (int)x;
                        int number_image = 10;
                        string[] item = { "ItemCode", "LotNo", "Region" };
                        string[] item_val = { Itemcode, Lotno, zone };
                        System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_HIGH_SPEED_BALL_SHEAR_IMAGE", TDMK_Code.filter_str(item, item_val));
                        if (data_image.Rows.Count > 0)
                        {
                            if (data_image.Rows.Count < number_image)
                            {
                                number_image = data_image.Rows.Count;

                                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                                string file_dic = Path.Combine(file_folder, "1.jpg");

                                myExcel.Range curr_rgn = ws.Range["B" + row_val];
                                for (int m = 0; m < number_image; m++)
                                {
                                    byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                    MemoryStream ms = new MemoryStream(data);
                                    Image image = Image.FromStream(ms);
                                    image.Save(file_dic);
                                    InsertPicture_Name(ws, curr_rgn, file_dic, 6);
                                    curr_rgn = curr_rgn.Offset[0, 1];
                                }
                            }
                            else
                            {
                                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                                string file_dic = Path.Combine(file_folder, "1.jpg");

                                myExcel.Range curr_rgn = ws.Range["B" + row_val];
                                for (int m = 0; m < number_image; m++)
                                {
                                    byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                    MemoryStream ms = new MemoryStream(data);
                                    Image image = Image.FromStream(ms);
                                    image.Save(file_dic);
                                    InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                                    curr_rgn = curr_rgn.Offset[0, 1];
                                }
                            }
                        }
                    }
                }

            }

        }
        public void Export_High_Speed_Ball_Shear_image_new(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            string[] item = { "ItemCode", "LotNo" };
            string[] item_val = { Itemcode, Lotno };
            int number_image = 10;
            System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_HIGH_SPEED_BALL_SHEAR_IMAGE", TDMK_Code.filter_str(item, item_val));
            if (data_image.Rows.Count > 0)
            {
                myExcel.Worksheet ws = wb.Sheets[1];
                int qty = Math.Min(data_image.Rows.Count, number_image);
                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                string file_dic = Path.Combine(file_folder, "1.jpg");
                int row_val;
                for (int x = 1; x < 40; x++)
                {
                    if (ws.Cells[x, 1].Value != null && ws.Cells[x, 1].Value.ToString() == "Photo")
                    {
                        row_val = (int)x;
                        myExcel.Range curr_rgn = ws.Range["B" + row_val];
                        for (int m = 0; m < qty; m++)
                        {
                            int col_off = (int)(m / 2) + 5 * (m % 2);
                            byte[] photo_data = (byte[])data_image.Rows[m]["Image_Photo"];
                            byte[] graph_data = (byte[])data_image.Rows[m]["Image_Graph"];
                            List<byte[]> image_data = new List<byte[]> { photo_data, graph_data };
                            for (int i = 0; i < image_data.Count; i++)
                            {
                                using (MemoryStream ms = new MemoryStream(image_data[i]))
                                {
                                    Image image = Image.FromStream(ms);
                                    image.Save(file_dic);
                                    InsertPicture_Name(ws, curr_rgn.Offset[i, col_off], file_dic, 6);
                                }
                            }
                            //curr_rgn.Offset[image_data.Count, col_off].Value = data_image.Rows[m]["Force_Data"];

                            if (TDMK_Code.IsNumeric(myCode.checkDBNull(data_image.Rows[m]["Force_Data"])))
                            {
                                double force_val = Convert.ToDouble(data_image.Rows[m]["Force_Data"]) * 1000;
                                curr_rgn.Offset[image_data.Count, col_off].Value = force_val;
                                if (force_val >= 250)
                                {
                                    curr_rgn.Offset[image_data.Count + 2, col_off].Value = "Pass";
                                }
                                else
                                {
                                    curr_rgn.Offset[image_data.Count + 2, col_off].Value = "Fail";
                                }
                            }

                        }
                        break;
                    }
                }
            }
        }
        /// add function
        public int row_input;
        public void Export_CQRA_Solderability_Image(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            myExcel.Worksheet ws;
            try
            {
                ws = wb.Sheets["CQRA - Solderability"];
            }
            catch (Exception e)
            {
                ws = wb.Sheets["CQRA - Solderability "];
            }
            string zone = "Picture";
            int dem = 0;
            for (int i = 1; i < 20; i++)
            {
                if (ws.Cells[i, 1].Value != null)
                {
                    if (ws.Cells[i, 1].Value.ToString().Contains("Picture"))
                    {
                        row_input = i;
                    }
                }
            }
            for (int j = 0; j < 15; j++)
            {
                if (ws.Cells[row_input - 1, 2 + j].Value != null)
                {
                    dem = dem + 1;
                }
            }
            int number_image = dem;
            string[] item = { "ItemCode", "LotNo", "Region" };
            string[] item_val = { Itemcode, Lotno, zone };
            System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_SOLDERABILITY_IMAGE", TDMK_Code.filter_str(item, item_val));
            if (data_image.Rows.Count > 0)
            {
                if (data_image.Rows.Count < number_image)
                {
                    number_image = data_image.Rows.Count;

                    string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                    string file_dic = Path.Combine(file_folder, "1.jpg");

                    myExcel.Range curr_rgn = ws.Range["B" + row_input];
                    for (int m = 0; m < number_image; m++)
                    {
                        byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                        MemoryStream ms = new MemoryStream(data);
                        Image image = Image.FromStream(ms);
                        image.Save(file_dic);
                        InsertPicture_Name(ws, curr_rgn, file_dic, 6);
                        curr_rgn = curr_rgn.Offset[0, 1];
                    }
                }
                else
                {
                    string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                    string file_dic = Path.Combine(file_folder, "1.jpg");

                    myExcel.Range curr_rgn = ws.Range["B" + row_input];
                    for (int m = 0; m < number_image; m++)
                    {
                        byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                        MemoryStream ms = new MemoryStream(data);
                        Image image = Image.FromStream(ms);
                        image.Save(file_dic);
                        InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                        curr_rgn = curr_rgn.Offset[0, 1];
                    }
                }
            }

        }
        public void Export_Thermal_Stress_Image(SqlConnection sqlcon, myExcel.Worksheet ws, string Itemcode, string Lotno)
        {
            // Get zone to input data
            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int k = 1; k < 60; k++)
            {
                if (ws.Cells[k, 1].Value != null)
                {
                    if (ws.Cells[k, 1].Value.ToString().Contains("hoto") || ws.Cells[k, 1].Value.ToString().Contains("icture"))
                    {
                        list.Add(ws.Cells[k, 1].Value.ToString() + "+" + k.ToString());
                    }
                }
            }
            foreach (string x in list)
            {
                string zone = x.Split('+')[0];
                int row_input = Int32.Parse(x.Split('+')[1]);
                int dem = 0;
                for (int j = 0; j < 15; j++)
                {
                    if (ws.Cells[row_input - 1, 2 + j].Value != null)
                    {
                        dem = dem + 1;
                    }
                }
                int number_image = dem;
                string[] item = { "ItemCode", "LotNo", "Region" };
                string[] item_val = { Itemcode, Lotno, zone };
                System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "THERMAL_STRESS_IMAGE", TDMK_Code.filter_str(item, item_val));
                if (data_image.Rows.Count > 0)
                {
                    if (data_image.Rows.Count < number_image)
                    {
                        number_image = data_image.Rows.Count;

                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        string file_dic = Path.Combine(file_folder, "1.jpg");

                        myExcel.Range curr_rgn = ws.Range["B" + row_input];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                            MemoryStream ms = new MemoryStream(data);
                            Image image = Image.FromStream(ms);
                            image.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn, file_dic, 6);
                            curr_rgn = curr_rgn.Offset[0, 1];
                        }
                    }
                    else
                    {
                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        string file_dic = Path.Combine(file_folder, "1.jpg");

                        myExcel.Range curr_rgn = ws.Range["B" + row_input];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                            MemoryStream ms = new MemoryStream(data);
                            Image image = Image.FromStream(ms);
                            image.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                            curr_rgn = curr_rgn.Offset[0, 1];
                        }
                    }
                }
            }

        }
        public void Export_IR_via_to_via_Image(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            // myExcel.Worksheet ws = wb.Sheets["CQRA - IR via to via"];

            myExcel.Worksheet ws = wb.Sheets[1]; 
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
                    int number_image = dem;
                    string[] item = { "ItemCode", "LotNo", "Region" };
                    string[] item_val = { Itemcode, Lotno, zone };
                    System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_IR_VIA_TO_VIA_IMAGE", TDMK_Code.filter_str(item, item_val));
                    if (data_image.Rows.Count > 0)
                    {
                        if (data_image.Rows.Count < number_image)
                        {
                            number_image = data_image.Rows.Count;

                            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                            string file_dic = Path.Combine(file_folder, "1.jpg");

                            myExcel.Range curr_rgn = ws.Range["B" + row_input];
                            for (int m = 0; m < number_image; m++)
                            {
                                byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                MemoryStream ms = new MemoryStream(data);
                                Image image = Image.FromStream(ms);
                                image.Save(file_dic);
                                InsertPicture_Name(ws, curr_rgn, file_dic, 6);
                                curr_rgn = curr_rgn.Offset[0, 1];
                            }
                        }
                        else
                        {
                            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                            string file_dic = Path.Combine(file_folder, "1.jpg");

                            myExcel.Range curr_rgn = ws.Range["B" + row_input];
                            for (int m = 0; m < number_image; m++)
                            {
                                byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                MemoryStream ms = new MemoryStream(data);
                                Image image = Image.FromStream(ms);
                                image.Save(file_dic);
                                InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                                curr_rgn = curr_rgn.Offset[0, 1];
                            }
                        }
                    }
                   // MessageBox.Show(new Form { TopMost = true }, "Hoàn thành!");
                }
                else
                {
                    export_formatmoi_13_12("CQRA_IR_VIA_TO_VIA_IMAGE", ws, Itemcode, Lotno, zone, sqlcon);

                }


            }





        }
        public void Export_IR_trace_to_trace_Image(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {

            myExcel.Worksheet ws = wb.Sheets[1];
             
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
                    int number_image = dem;
                    string[] item = { "ItemCode", "LotNo", "Region" };
                    string[] item_val = { Itemcode, Lotno, zone };
                    System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_IR_TRACE_TO_TRACE_IMAGE", TDMK_Code.filter_str(item, item_val));
                    if (data_image.Rows.Count > 0)
                    {
                        if (data_image.Rows.Count < number_image)
                        {
                            number_image = data_image.Rows.Count;

                            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                            string file_dic = Path.Combine(file_folder, "1.jpg");

                            myExcel.Range curr_rgn = ws.Range["B" + row_input];
                            for (int m = 0; m < number_image; m++)
                            {
                                byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                MemoryStream ms = new MemoryStream(data);
                                Image image = Image.FromStream(ms);
                                image.Save(file_dic);
                                InsertPicture_Name(ws, curr_rgn, file_dic, 6);
                                curr_rgn = curr_rgn.Offset[0, 1];
                            }
                        }
                        else
                        {
                            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                            string file_dic = Path.Combine(file_folder, "1.jpg");

                            myExcel.Range curr_rgn = ws.Range["B" + row_input];
                            for (int m = 0; m < number_image; m++)
                            {
                                byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                MemoryStream ms = new MemoryStream(data);
                                Image image = Image.FromStream(ms);
                                image.Save(file_dic);
                                InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                                curr_rgn = curr_rgn.Offset[0, 1];
                            }
                        }
                    }
                }
                else
                {
                    export_formatmoi_13_12("CQRA_IR_TRACE_TO_TRACE_IMAGE", ws, Itemcode, Lotno, zone, sqlcon);
                }

            }
        }


        public void export_formatmoi_13_12(string table_name, myExcel.Worksheet ws, string Itemcode, string Lotno, string zone, SqlConnection sqlcon)
        {
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

                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        string file_dic = Path.Combine(file_folder, "1.jpg");

                        myExcel.Range curr_rgn = ws.Cells[i + 1, j];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                            MemoryStream ms = new MemoryStream(data);
                            Image image = Image.FromStream(ms);
                            image.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn, file_dic, 6);
                            curr_rgn = curr_rgn.Offset[1, 0];
                        }
                     //   MessageBox.Show(new Form { TopMost = true }, "Hoàn thành!");

                        return;

                    }
                }
            }
        }


        public void Export_IR_layer_to_layer_Image(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {

            myExcel.Worksheet ws = wb.Sheets[1]; 
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
                    int number_image = dem;
                    string[] item = { "ItemCode", "LotNo", "Region" };
                    string[] item_val = { Itemcode, Lotno, zone };
                    System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_IR_LAYER_TO_LAYER_IMAGE", TDMK_Code.filter_str(item, item_val));
                    if (data_image.Rows.Count > 0)
                    {
                        if (data_image.Rows.Count < number_image)
                        {
                            number_image = data_image.Rows.Count;

                            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                            string file_dic = Path.Combine(file_folder, "1.jpg");

                            myExcel.Range curr_rgn = ws.Range["B" + row_input];
                            for (int m = 0; m < number_image; m++)
                            {
                                byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                MemoryStream ms = new MemoryStream(data);
                                Image image = Image.FromStream(ms);
                                image.Save(file_dic);
                                InsertPicture_Name(ws, curr_rgn, file_dic, 6);
                                curr_rgn = curr_rgn.Offset[0, 1];
                            }
                        }
                        else
                        {
                            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                            string file_dic = Path.Combine(file_folder, "1.jpg");

                            myExcel.Range curr_rgn = ws.Range["B" + row_input];
                            for (int m = 0; m < number_image; m++)
                            {
                                byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                MemoryStream ms = new MemoryStream(data);
                                Image image = Image.FromStream(ms);
                                image.Save(file_dic);
                                InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                                curr_rgn = curr_rgn.Offset[0, 1];
                            }
                        }
                    }
                }
                else
                {
                    export_formatmoi_13_12("CQRA_IR_LAYER_TO_LAYER_IMAGE", ws, Itemcode, Lotno, zone, sqlcon);
                }
            }
        }

        public myExcel.Workbook create_export_wrk(string format_file, string process_name)
        {
            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            xlsApp.DisplayAlerts = false;
            myExcel.Workbook wb = TDMK_Code.open_excel_file(format_file, "", "");
            if (wb != null)
            {
                string _process = process_name.Replace("_", "").Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "").ToUpper();
                string mySheet = "";
                foreach (myExcel.Worksheet tg_sht in wb.Worksheets)
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
                    myExcel.Workbook save_wb = TDMK_Code.Create_workbook();
                    wb.Sheets[mySheet].Copy(After: save_wb.Sheets[1]);
                    myExcel.Worksheet del_sht = save_wb.Sheets[1];
                    del_sht.Delete();
                    wb.Close();
                    xlsApp.DisplayAlerts = true;
                    return save_wb;
                }
                else
                {
                    wb.Close();
                    xlsApp.DisplayAlerts = true;
                    return null;
                }
            }
            else
            {
                return null;
            }
        }

        public void Export_Chemical_resist_Image(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {

            string[] list = { "Picture before", "Picture after" };
            myExcel.Worksheet ws = null;
            
            foreach (myExcel.Worksheet tg_sht in wb.Worksheets)
            {
                string cur_sht_name = tg_sht.Name.ToUpper().Replace("_", "").Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "");
                if (cur_sht_name == "CQRA - Chemical resistance".Replace("-", "").Replace(" ", "").ToUpper() || cur_sht_name == "CQRA - Chemical resist".Replace("-", "").Replace(" ", "").ToUpper())
                {
                    ws = wb.Sheets[tg_sht.Name];
                    break;
                }
            }


            if (ws != null)
            {
                foreach (string zone in list)
                {
                    int dem = 0;
                    for (int i = 1; i < 30; i++)
                    {
                        if (myCode.checkDBNull(ws.Cells[i, 1].Value).ToString().ToUpper().Replace(" ", "").Contains(zone.ToUpper().Replace(" ", "")))
                        {
                            row_input = i;
                        }
                    }
                    if (row_input > 0)
                    {
                        int number_image = 5;
                        string[] item = { "ItemCode", "LotNo", "Region" };
                        string[] item_val = { Itemcode, Lotno, zone };
                        System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_CHEMICAL_RESISTANCE_IMAGE", TDMK_Code.filter_str(item, item_val));
                        if (data_image.Rows.Count > 0)
                        {

                            if (data_image.Rows.Count < number_image)
                            {
                                number_image = data_image.Rows.Count;

                                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                                string file_dic = Path.Combine(file_folder, "1.jpg");

                                myExcel.Range curr_rgn = ws.Range["B" + row_input];
                                for (int m = 0; m < number_image; m++)
                                {
                                    byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                    MemoryStream ms = new MemoryStream(data);
                                    Image image = Image.FromStream(ms);
                                    image.Save(file_dic);
                                    InsertPicture_Name(ws, curr_rgn, file_dic, 6);
                                    curr_rgn = curr_rgn.Offset[0, 1];
                                }
                            }
                            else
                            {
                                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                                string file_dic = Path.Combine(file_folder, "1.jpg");

                                myExcel.Range curr_rgn = ws.Range["B" + row_input];
                                for (int m = 0; m < number_image; m++)
                                {
                                    byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                    MemoryStream ms = new MemoryStream(data);
                                    Image image = Image.FromStream(ms);
                                    image.Save(file_dic);
                                    InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                                    curr_rgn = curr_rgn.Offset[0, 1];
                                }
                            }
                        }
                       
                    }

                   

                }

            }

        }
        public byte[] File_image_to_byte(string file)
        {
            Byte[] data = new Byte[0];
            Image myImg = Image.FromFile(file);
            ImageConverter imgCon = new ImageConverter();
            data = (byte[])imgCon.ConvertTo(myImg, typeof(byte[]));
            return data;
        }
        public void Save_High_speed_Image_LogFile_toDB(DataTable src_dt, string Itemcode, string Lot, string folder, SqlConnection sqlcon)
        {
            if (src_dt.Rows.Count > 0)
            {
                //string tablename = "CQRA_HIGH_SPEED_BALL_SHEAR_IMAGE";
                //string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { Itemcode, Lot });
                string[] file_image = Directory.GetFiles(folder, "*.JPG");
                string[] file_data = Directory.GetFiles(folder, "*.csv");
                Dictionary<string, byte[]> image_data_lst = new Dictionary<string, byte[]>();
                Dictionary<string, string> data_val_lst = new Dictionary<string, string>();
                foreach (string file in file_image)
                {
                    Byte[] data = new Byte[0];
                    Image myImg = Image.FromFile(file);
                    ImageConverter imgCon = new ImageConverter();
                    data = (byte[])imgCon.ConvertTo(myImg, typeof(byte[]));
                    string pcs_no = (Path.GetFileName(file)).Split('.')[0];
                    if (image_data_lst.Keys.ToList().IndexOf(pcs_no) == -1)
                    {
                        image_data_lst.Add(pcs_no, data);
                    }
                }
                foreach (string f_data in file_data)
                {
                    myExcel.Workbook tar_wrk = TDMK_Code.open_excel_file(f_data, "", "");
                    myExcel.Worksheet tar_sht = tar_wrk.Sheets[1];
                    myExcel.Range id_rgn = tar_sht.Range["E6"];
                    myExcel.Range data_rgn = tar_sht.Range["G6"];
                    int r_inx = 0;
                    while (myCode.checkDBNull(id_rgn.Offset[r_inx, 0].Value) != "")
                    {
                        string inx = Convert.ToString(id_rgn.Offset[r_inx, 0].Value);
                        List<string> temp = data_val_lst.Keys.ToList();
                        if (data_val_lst.Keys.ToList().IndexOf(inx) == -1)
                        {
                            data_val_lst.Add(inx, myCode.checkDBNull(data_rgn.Offset[r_inx, 0].Value));
                        }
                        r_inx++;
                    }
                    tar_wrk.Close();
                }
                List<string> pcs_lst = src_dt.AsEnumerable().Select(x => x.Field<string>("Pcs_No")).ToList();
                foreach (var pcs in pcs_lst)
                {
                    int r_inx = image_data_lst.Keys.ToList().IndexOf(pcs);
                    if (r_inx != -1)
                    {
                        src_dt.Rows[pcs_lst.IndexOf(pcs)]["Image_Graph"] = image_data_lst[pcs];
                    }
                    int data_inx = data_val_lst.Keys.ToList().IndexOf(pcs);
                    if (data_inx != -1)
                    {
                        src_dt.Rows[pcs_lst.IndexOf(pcs)]["Force_Data"] = data_val_lst[pcs];
                    }
                }
            }
            else
            {
                MessageBox.Show("Please, capture image for Photo region at first", "Warning");
            }
        }

        public void BatchBulkCopy(SqlConnection sqlcon, DataTable dataTable, string DestinationTbl)
        {
            DataTable dtInsertRows = dataTable;
            using (SqlBulkCopy sbc = new SqlBulkCopy(sqlcon))
            {
                if (sqlcon.State != ConnectionState.Open)
                {
                    sqlcon.Open();
                }
                sbc.DestinationTableName = DestinationTbl;
                foreach (DataColumn dc in dataTable.Columns)
                {
                    sbc.ColumnMappings.Add(dc.ColumnName, dc.ColumnName);
                }
                sbc.WriteToServer(dtInsertRows);
                sqlcon.Close();
            }
        }

        public void Save_Image_LogFile_toDB(string Itemcode, string Lot, string folder, SqlConnection sqlcon, string tablename, string Region)
        {
            //string tablename = "CQRA_HIGH_SPEED_BALL_SHEAR_IMAGE";
            string[] file_image;
            file_image = Directory.GetFiles(folder, "*.JPG");
            if (file_image.Length == null)
            {
                file_image = Directory.GetFiles(folder, "*.PNG");
            }

            foreach (string file in file_image)
            {
                Byte[] data = new Byte[0];
                Image myImg = Image.FromFile(file);
                // Image myImg = (Bitmap)myImg
                ImageConverter imgCon = new ImageConverter();
                data = (byte[])imgCon.ConvertTo(myImg, typeof(byte[]));

                int ID = TDMK_Code.SQL_MAX(tablename, "ID", sqlcon) + 1;
                string itemcode = Itemcode;
                string lotno = Lot;
                string pcs_no = (Path.GetFileName(file)).Split('.')[0];
                string time = DateTime.Now.ToString();
                string op = "Auto";

                string query = "Insert into " + tablename + "(ID,ItemCode,LotNo,Region,Image_Data,Pcs_No,Time_Update,Operator ) values (@ID, @ItemCode, @LotNo,@Region,@Image_Data,@Pcs_No,@Update_Time,@Operator)";

                using (SqlCommand cmd = new SqlCommand(query))
                {
                    cmd.Connection = sqlcon;
                    cmd.Parameters.AddWithValue("@ID", ID);
                    cmd.Parameters.AddWithValue("@ItemCode", itemcode);
                    cmd.Parameters.AddWithValue("@LotNo", lotno);
                    cmd.Parameters.AddWithValue("@Region", Region);
                    cmd.Parameters.AddWithValue("@Image_Data", data);
                    cmd.Parameters.AddWithValue("@Pcs_No", pcs_no);
                    cmd.Parameters.AddWithValue("@Update_Time", time);
                    cmd.Parameters.AddWithValue("@Operator", op);
                    sqlcon.Open();
                    cmd.ExecuteNonQuery();
                    sqlcon.Close();

                }

            }
        }

        public void Save_High_speed_Data_LogFile_toDB(string Itemcode, string Lot, string folder, SqlConnection sqlcon)
        {
            string tablename = "CQRA_HIGH_SPEED_BALL_SHEAR_LOGFILE";
            string[] file_image = Directory.GetFiles(folder, "*.csv");
            foreach (string file in file_image)
            {
                myExcel.Workbook wb = TDMK_Code.open_excel_file(Path.GetDirectoryName(file), Path.GetFileName(file), "");
                myExcel.Worksheet ws = wb.Sheets[1];
                for (int i = 3; i < 30; i++)
                {
                    if (ws.Cells[i, 1].Value != null)
                    {
                        if (ws.Cells[i, 1].Value.ToString() == "TEST")
                        {
                            string data = ws.Cells[i, 6].Value.ToString();
                            int ID = TDMK_Code.SQL_MAX(tablename, "ID", sqlcon) + 1;
                            string itemcode = Itemcode;
                            string lotno = Lot; ;
                            string pcs_no = (i - 5).ToString();
                            string time = DateTime.Now.ToString();
                            string op = "Auto";
                            string Region = "Shear force";

                            string query = "Insert into " + tablename + "(ID,ItemCode,LotNo,Region,Data,Pcs_No,Time_Update,Operator ) values (@ID, @ItemCode, @LotNo,@Region,@Data,@Pcs_No,@Update_Time,@Operator)";

                            using (SqlCommand cmd = new SqlCommand(query))
                            {
                                cmd.Connection = sqlcon;
                                cmd.Parameters.AddWithValue("@ID", ID);
                                cmd.Parameters.AddWithValue("@ItemCode", itemcode);
                                cmd.Parameters.AddWithValue("@LotNo", lotno);
                                cmd.Parameters.AddWithValue("@Region", Region);
                                cmd.Parameters.AddWithValue("@Data", data);
                                cmd.Parameters.AddWithValue("@Pcs_No", pcs_no);
                                cmd.Parameters.AddWithValue("@Update_Time", time);
                                cmd.Parameters.AddWithValue("@Operator", op);
                                sqlcon.Open();
                                cmd.ExecuteNonQuery();
                                sqlcon.Close();
                            }
                        }
                    }

                }


            }
        }

        public void Export_Data_High_SpeedBallShear_Data(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            myExcel.Worksheet ws = wb.Sheets["CQRA - High speed ball shear"];
            for (int i = 3; i < 30; i++)
            {
                if (ws.Cells[i, 1].Value != null)
                {
                    if (ws.Cells[i, 1].Value.ToString().Contains("Shear force"))
                    {
                        int number_image = 10;
                        string[] item = { "ItemCode", "LotNo", "Region" };
                        string[] item_val = { Itemcode, Lotno, "Shear force" };
                        System.Data.DataTable data = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_HIGH_SPEED_BALL_SHEAR_LOGFILE", TDMK_Code.filter_str(item, item_val));
                        if (data.Rows.Count > 0)
                        {

                            if (data.Rows.Count < number_image)
                            {
                                number_image = data.Rows.Count;

                                // myExcel.Range curr_rgn = ws.Range["B" + row_input];
                                for (int m = 0; m < number_image; m++)
                                {
                                    string val_data = data.Rows[m]["Data"].ToString();
                                    ws.Cells[i, m + 2].Value = val_data;
                                    //curr_rgn = curr_rgn.Offset[0, 1];
                                }
                            }
                            else
                            {
                                //myExcel.Range curr_rgn = ws.Range["B" + row_input];
                                for (int m = 0; m < number_image; m++)
                                {
                                    string val_data = data.Rows[m]["Data"].ToString();
                                    ws.Cells[i, m + 2].Value = val_data;
                                    //curr_rgn = curr_rgn.Offset[0, 1];
                                }
                            }
                        }
                    }
                }
            }
        }
        public void Export_CQRA_Hot_bar_loop_test_Image(SqlConnection sqlcon, myExcel.Workbook wb, string excel_file_addr, string Itemcode, string Lotno)
        {
            myExcel.Worksheet ws;
            try
            {
                ws = wb.Sheets["CQRA - Hot bar loop test"];
            }
            catch (Exception e)
            {
                ws = wb.Sheets["CQRA - Hot bar loop test "];
            }
            int dem = 0;
            string zone = "Picture";
            for (int i = 1; i < 30; i++)
            {
                if (ws.Cells[i, 1].Value != null)
                {
                    if (ws.Cells[i, 1].Value.ToString().Contains(zone))
                    {
                        row_input = i;
                    }
                }
            }
            for (int j = 0; j < 15; j++)
            {
                if (ws.Cells[row_input - 1, 2 + j].Value != null)
                {
                    dem = dem + 1;
                }
            }
            int number_image = dem;
            string[] item = { "ItemCode", "LotNo", "Region" };
            string[] item_val = { Itemcode, Lotno, zone };
            System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "HOT_BAR_LOOP_TEST_IMAGE", TDMK_Code.filter_str(item, item_val));
            if (data_image.Rows.Count > 0)
            {
                if (data_image.Rows.Count < number_image)
                {
                    number_image = data_image.Rows.Count;

                    string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                    string file_dic = Path.Combine(file_folder, "1.jpg");

                    myExcel.Range curr_rgn = ws.Range["B" + row_input];
                    for (int m = 0; m < number_image; m++)
                    {
                        byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                        MemoryStream ms = new MemoryStream(data);
                        Image image = Image.FromStream(ms);
                        image.Save(file_dic);
                        InsertPicture_Name(ws, curr_rgn, file_dic, 6);
                        curr_rgn = curr_rgn.Offset[0, 1];
                    }
                }
                else
                {
                    string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                    string file_dic = Path.Combine(file_folder, "1.jpg");

                    myExcel.Range curr_rgn = ws.Range["B" + row_input];
                    for (int m = 0; m < number_image; m++)
                    {
                        byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                        MemoryStream ms = new MemoryStream(data);
                        Image image = Image.FromStream(ms);
                        image.Save(file_dic);
                        InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                        curr_rgn = curr_rgn.Offset[0, 1];
                    }
                }
            }

        }

        // end add
        public int dem;
        public int start;
        public int end;
        public void Export_CQRA_Dielectric_withstanding_old(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            myExcel.Worksheet ws;
            try
            {
                ws = wb.Sheets["CQRA - Dielectric withstanding"];
            }
            catch (Exception e)
            {
                ws = wb.Sheets["CQRA - Dielectric withstanding "];
            }

            string[] arr = { "After+H", "Before+C" };

            for (int i = 1; i < 30; i++)
            {
                if (ws.Cells[i, 1].Value != null)
                {
                    if (ws.Cells[i, 1].Value.ToString() == "1")
                    {
                        start = i;
                    }
                    if (ws.Cells[i, 1].Value.ToString() == "5")
                    {
                        end = i;
                    }
                }
            }

            int row_input = start;
            int number_image = end - start + 1;

            // MessageBox.Show(number_image.ToString());
            foreach (string xx in arr)
            {
                string zone = xx.Split('+')[0];
                string range = xx.Split('+')[1];  // ra B hoặc I
                string[] item = { "ItemCode", "LotNo", "Region" };
                string[] item_val = { Itemcode, Lotno, zone };
                System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_DIELECTRIC_WITHSTANDING_IMAGE", TDMK_Code.filter_str(item, item_val));
                if (data_image.Rows.Count > 0)
                {
                    // export data
                    try
                    {
                        if (zone == "After")
                        {
                            myExcel.Range curr_rgn = ws.Range["G18"];
                            for (int m = 0; m < 5; m++)
                            {
                                string x = data_image.Rows[m]["Remark"].ToString();
                                if (x != "")
                                {
                                    curr_rgn.Value = x;
                                }
                                curr_rgn = curr_rgn.Offset[1, 0];
                            }
                        }

                    }
                    catch (Exception)
                    {

                    }


                    if (data_image.Rows.Count < number_image)
                    {
                        number_image = data_image.Rows.Count;

                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        string file_dic = Path.Combine(file_folder, "1.jpg");

                        myExcel.Range curr_rgn = ws.Range[range + row_input];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                            MemoryStream ms = new MemoryStream(data);
                            Image image = Image.FromStream(ms);
                            image.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn, file_dic, 6);
                            curr_rgn = curr_rgn.Offset[1, 0];
                        }

                    }
                    else
                    {
                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        string file_dic = Path.Combine(file_folder, "1.jpg");

                        myExcel.Range curr_rgn = ws.Range[range + row_input];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                            MemoryStream ms = new MemoryStream(data);
                            Image image = Image.FromStream(ms);
                            image.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn, file_dic, 6);
                            curr_rgn = curr_rgn.Offset[1, 0];
                        }
                    }
                }
            }

        }
        public void Export_CQRA_Dielectric_withstanding_old_2(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_DIELECTRIC_WITHSTANDING_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { Itemcode, Lotno }));
            if (src_dt.Rows.Count > 0)
            {
                int start = 0;
                int end = 0;
                myExcel.Worksheet ws = wb.Sheets[1];
                for (int i = 1; i < 30; i++)
                {
                    if (ws.Cells[i, 1].Value != null)
                    {
                        if (ws.Cells[i, 1].Value.ToString() == "1")
                        {
                            start = i;
                        }
                        if (ws.Cells[i, 1].Value.ToString() == "5")
                        {
                            end = i;
                        }
                    }
                }
                int number_image = end - start + 1;
                myExcel.Range before_rgn = ws.Range["E16"];
                myExcel.Range after_rgn = ws.Range["F16"];
                myExcel.Range resistance_rgn = ws.Range["C16"];
                int qty = Math.Min(number_image, src_dt.Rows.Count);
                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                if (!Directory.Exists(file_folder))
                {
                    Directory.CreateDirectory(file_folder);
                }
                string file_dic = Path.Combine(file_folder, "1.jpg");
                for (int m = 0; m < qty; m++)
                {
                    byte[] bef_data = (byte[])src_dt.Rows[m]["Image_Before"];
                    using (MemoryStream ms = new MemoryStream(bef_data))
                    {
                        Image image = Image.FromStream(ms);
                        image.Save(file_dic);
                        InsertPicture_Name(ws, before_rgn.Offset[m, 0], file_dic, 6);
                    }
                    byte[] aft_data = (byte[])src_dt.Rows[m]["Image_After"];
                    using (MemoryStream ms = new MemoryStream(aft_data))
                    {
                        Image image = Image.FromStream(ms);
                        image.Save(file_dic);
                        InsertPicture_Name(ws, after_rgn.Offset[m, 0], file_dic, 6);
                    }
                    resistance_rgn.Offset[m, 0].Value = src_dt.Rows[m]["Resistance_Data"].ToString();
                }
            }
        }

        public void export_formatmoi_Dielectric_withstanding_13_12(string table_name, myExcel.Worksheet ws, string Itemcode, string Lotno, string zone, SqlConnection sqlcon)
        {
            for (int i = 1; i < 30; i++)
            {
                for (int j = 1; j < 15; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).ToString().ToUpper().Contains("Photo after testT".Replace(" ", "").ToUpper()))
                    {
                        for (int k = 0; k < 15; k++)
                        {
                            if (myCode.IsNumeric(myCode.checkDBNull(ws.Cells[i + k, j - 4].Value)))
                            {
                                dem = dem + 1;
                            }
                        }

                        string[] item = { "ItemCode", "LotNo" };
                        string[] item_val = { Itemcode, Lotno };
                        System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, table_name, TDMK_Code.filter_str(item, item_val));

                        int number_image = new int[] { dem, data_image.Rows.Count }.Min();

                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        string file_dic = Path.Combine(file_folder, "1.jpg");

                        myExcel.Range curr_rgn = ws.Cells[i + 1, j];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data1 = (byte[])data_image.Rows[m]["Image_Before"];
                            MemoryStream ms1 = new MemoryStream(data1);
                            Image image1 = Image.FromStream(ms1);
                            image1.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn, file_dic, 6);



                            byte[] data2 = (byte[])data_image.Rows[m]["Image_After"];
                            MemoryStream ms2 = new MemoryStream(data2);
                            Image image2 = Image.FromStream(ms2);
                            image2.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn.Offset[-1, 0], file_dic, 6);
                            curr_rgn = curr_rgn.Offset[1, 0];
                        }
                        MessageBox.Show(new Form { TopMost = true }, "Hoàn thành!");

                        return;

                    }
                }
            }
        }

        public void Export_CQRA_Dielectric_withstanding(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            myExcel.Worksheet ws = wb.Sheets[1];
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
                        System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_DIELECTRIC_WITHSTANDING_IMAGE" , TDMK_Code.filter_str(item, item_val));

                        int number_image = new int[] { dem, data_image.Rows.Count }.Min();

                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        string file_dic = Path.Combine(file_folder, "1.jpg");

                        myExcel.Range curr_rgn = ws.Cells[i + 1, j];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data1 = (byte[])data_image.Rows[m]["Image_After"];
                            MemoryStream ms1 = new MemoryStream(data1);
                            Image image1 = Image.FromStream(ms1);
                            image1.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn, file_dic, 6);



                            byte[] data2 = (byte[])data_image.Rows[m]["Image_Before"];
                            MemoryStream ms2 = new MemoryStream(data2);
                            Image image2 = Image.FromStream(ms2);
                            image2.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn.Offset[0, -1], file_dic, 6);
                            curr_rgn = curr_rgn.Offset[1, 0];
                        }
                        MessageBox.Show(new Form { TopMost = true }, "Hoàn thành!");

                        return;

                    }
                }
            }
        }
        // CQRA_FLUX_RESIST_IMAGE

        public void Export_Flux_Resist_Image(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            int start_with_FR4 = 24;
            int start_without_FR4 = 110;
           
            myExcel.Worksheet ws = wb.Sheets[1];
             
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
                for (int j = start_with_FR4; j < start_without_FR4; j++)
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
                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        string file_dic = Path.Combine(file_folder, "1.jpg");
                        myExcel.Range curr_rgn = ws.Range["H" + row_input];
                        myExcel.Range flux_data_rgn = curr_rgn.Offset[1, -4];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                            MemoryStream ms = new MemoryStream(data);
                            Image image = Image.FromStream(ms);
                            image.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn, file_dic, 20);
                            string flux = myCode.checkDBNull(data_image.Rows[m]["Flux_Data"]);
                            if (TDMK_Code.IsNumeric(flux))
                            {
                                double flux_val = Convert.ToDouble(flux);
                                flux_data_rgn.Offset[m, 0].Value = flux_val;
                                if (myCode.check_in_limit2("90", "50", flux)==Color.White)
                                {
                                    flux_data_rgn.Offset[m, 2].Value = "Pass";
                                }
                                else
                                {
                                    flux_data_rgn.Offset[m, 2].Value = "Fail";
                                }
                            }
                            else
                            {
                                flux_data_rgn.Offset[m, 0].Value = flux;
                            }
                            curr_rgn = curr_rgn.Offset[0, 1];
                        }


                        //if (data_image.Rows.Count < number_image)
                        //{
                        //    number_image = data_image.Rows.Count;

                        //    string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        //    string file_dic = Path.Combine(file_folder, "1.jpg");

                        //    myExcel.Range curr_rgn = ws.Range["H" + row_input];
                        //    for (int m = 0; m < number_image; m++)
                        //    {
                        //        byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                        //        MemoryStream ms = new MemoryStream(data);
                        //        Image image = Image.FromStream(ms);
                        //        image.Save(file_dic);
                        //        InsertPicture_Name(ws, curr_rgn, file_dic, 20);
                        //        curr_rgn = curr_rgn.Offset[0, 1];
                        //    }
                        //}
                        //else
                        //{
                        //    string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        //    string file_dic = Path.Combine(file_folder, "1.jpg");

                        //    myExcel.Range curr_rgn = ws.Range["H" + row_input];
                        //    for (int m = 0; m < number_image; m++)
                        //    {
                        //        byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                        //        MemoryStream ms = new MemoryStream(data);
                        //        Image image = Image.FromStream(ms);
                        //        image.Save(file_dic);
                        //        InsertPicture_Name(ws, curr_rgn, file_dic, 20);
                        //        curr_rgn = curr_rgn.Offset[0, 1];
                        //    }
                        //}
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
                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        string file_dic = Path.Combine(file_folder, "1.jpg");
                        myExcel.Range curr_rgn = ws.Range["H" + row_input];
                        myExcel.Range flux_data_rgn = curr_rgn.Offset[1, -4];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                            MemoryStream ms = new MemoryStream(data);
                            Image image = Image.FromStream(ms);
                            image.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn, file_dic, 20);
                            string flux = myCode.checkDBNull(data_image.Rows[m]["Flux_Data"]);
                            if (TDMK_Code.IsNumeric(flux))
                            {
                                double flux_val = Convert.ToDouble(flux);
                                flux_data_rgn.Offset[m, 0].Value = flux_val;
                                if (myCode.check_in_limit2("90", "50", flux) == Color.White)
                                {
                                    flux_data_rgn.Offset[m, 2].Value = "Pass";
                                }
                                else
                                {
                                    flux_data_rgn.Offset[m, 2].Value = "Fail";
                                }
                            }
                            else
                            {
                                flux_data_rgn.Offset[m, 0].Value = flux;
                            }
                            curr_rgn = curr_rgn.Offset[0, 1];
                        }

                        //if (data_image.Rows.Count < number_image)
                        //{
                        //    number_image = data_image.Rows.Count;

                        //    string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        //    string file_dic = Path.Combine(file_folder, "1.jpg");

                        //    myExcel.Range curr_rgn = ws.Range["H" + row_input];
                        //    for (int m = 0; m < number_image; m++)
                        //    {
                        //        byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                        //        MemoryStream ms = new MemoryStream(data);
                        //        Image image = Image.FromStream(ms);
                        //        image.Save(file_dic);
                        //        InsertPicture_Name(ws, curr_rgn, file_dic, 20);
                        //        curr_rgn = curr_rgn.Offset[0, 1];
                        //    }
                        //}
                        //else
                        //{
                        //    string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        //    string file_dic = Path.Combine(file_folder, "1.jpg");

                        //    myExcel.Range curr_rgn = ws.Range["H" + row_input];
                        //    for (int m = 0; m < number_image; m++)
                        //    {
                        //        byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                        //        MemoryStream ms = new MemoryStream(data);
                        //        Image image = Image.FromStream(ms);
                        //        image.Save(file_dic);
                        //        InsertPicture_Name(ws, curr_rgn, file_dic, 20);
                        //        curr_rgn = curr_rgn.Offset[0, 1];
                        //    }
                        //}
                    }
                }
            }


        }
        public void Export_OQC_Test_Image(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            string[] arr = { "Solder Mask", "PI & SUS", "Gold", "EMI" };
            
            myExcel.Worksheet ws = null;
            foreach (myExcel.Worksheet tg_sht in wb.Worksheets)
            {
                string cur_sht_name = tg_sht.Name.ToUpper().Replace("_", "").Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "");
                if (cur_sht_name == "OQC test".Replace("-", "").Replace(" ", "").ToUpper())
                {
                    ws = wb.Sheets[tg_sht.Name];
                    break;
                }
            }

            if (ws != null)
            {


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

                            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                            string file_dic = Path.Combine(file_folder, "1.jpg");

                            myExcel.Range curr_rgn = ws.Range["C" + row_input];
                            for (int m = 0; m < number_image; m++)
                            {
                                byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                MemoryStream ms = new MemoryStream(data);
                                Image image = Image.FromStream(ms);
                                image.Save(file_dic);
                                InsertPicture_Name(ws, curr_rgn, file_dic, 12);
                                curr_rgn = curr_rgn.Offset[0, 1];
                            }
                        }
                        else
                        {
                            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                            string file_dic = Path.Combine(file_folder, "1.jpg");

                            myExcel.Range curr_rgn = ws.Range["C" + row_input];
                            for (int m = 0; m < number_image; m++)
                            {
                                byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                MemoryStream ms = new MemoryStream(data);
                                Image image = Image.FromStream(ms);
                                image.Save(file_dic);
                                InsertPicture_Name(ws, curr_rgn, file_dic, 12);
                                curr_rgn = curr_rgn.Offset[0, 1];
                            }
                        }
                    }
                }
            }
        }
        public void Export_CQRA_bHast_Image(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
             
            myExcel.Worksheet ws = wb.Sheets[1]; 

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
                if(row_input > 0)
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

                            myExcel.Range curr_rgn = ws.Range["B" + row_input + 1];
                            for (int m = 0; m < data_image.Rows.Count; m++)
                            {
                                byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                MemoryStream ms = new MemoryStream(data);
                                Image image = Image.FromStream(ms);
                                image.Save(file_dic);
                                InsertPicture_Name(ws, curr_rgn, file_dic, 6);
                                curr_rgn = curr_rgn.Offset[0, 1];
                            }
                        }
                        else
                        {
                            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                            string file_dic = Path.Combine(file_folder, "1.jpg");

                            myExcel.Range curr_rgn = ws.Range["B" + row_input + 1];
                            for (int m = 0; m < number_image; m++)
                            {
                                byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                MemoryStream ms = new MemoryStream(data);
                                Image image = Image.FromStream(ms);
                                image.Save(file_dic);
                                InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                                curr_rgn = curr_rgn.Offset[0, 1];
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
        public void Export_Impedance_Image(SqlConnection sqlcon, myExcel.Workbook wb, string excel_file_addr, string Itemcode, string Lotno)
        {

            myExcel.Worksheet ws;
            try
            {
                ws = wb.Sheets["Impedance"];
            }
            catch (Exception e)
            {
                ws = wb.Sheets["Impedance "];
            }

            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = Path.Combine(file_folder, "1.jpg");
            for (int x = 10; x < 80; x++)
            {
                if (ws.Cells[x, 15].Value != null && ws.Cells[x, 15].Value.ToString() == "Picture")
                {
                    myExcel.Range refer_range = ws.Range["O" + x];
                    //MessageBox
                    myExcel.Range range_11 = refer_range.Offset[2, 0];
                    myExcel.Range range_12 = refer_range.Offset[4, 0];
                    myExcel.Range range_21 = refer_range.Offset[2, 1];
                    myExcel.Range range_22 = refer_range.Offset[4, 0];
                    myExcel.Range range_31 = refer_range.Offset[2, 2];
                    myExcel.Range range_32 = refer_range.Offset[4, 2];

                    string[] regions = { "Picture Sample 1", "Picture Sample 2", "Picture Sample 3" };
                    int number_image = 2;
                    string[] item = { "ItemCode", "LotNo", "Region" };

                    foreach (string region in regions)
                    {
                        string[] item_val = { Itemcode, Lotno, region };
                        System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_IMAGE", TDMK_Code.filter_str(item, item_val));
                        if (data_image != null)
                        {

                            if (region == "Picture Sample 1")
                            {
                                byte[] data1 = (byte[])data_image.Rows[0]["Image_Data"];
                                MemoryStream ms1 = new MemoryStream(data1);
                                Image image1 = Image.FromStream(ms1);
                                image1.Save(file_dic);
                                InsertPicture_Name(ws, ws.Range["O42"], file_dic, 10);

                                byte[] data2 = (byte[])data_image.Rows[1]["Image_Data"];
                                MemoryStream ms2 = new MemoryStream(data2);
                                Image image2 = Image.FromStream(ms2);
                                image2.Save(file_dic);
                                InsertPicture_Name(ws, ws.Range["O60"], file_dic, 10);
                            }
                            if (region == "Picture Sample 2")
                            {
                                byte[] data1 = (byte[])data_image.Rows[0]["Image_Data"];
                                MemoryStream ms1 = new MemoryStream(data1);
                                Image image1 = Image.FromStream(ms1);
                                image1.Save(file_dic);
                                InsertPicture_Name(ws, ws.Range["R42"], file_dic, 10);

                                byte[] data2 = (byte[])data_image.Rows[1]["Image_Data"];
                                MemoryStream ms2 = new MemoryStream(data2);
                                Image image2 = Image.FromStream(ms2);
                                image2.Save(file_dic);
                                InsertPicture_Name(ws, ws.Range["R60"], file_dic, 10);
                            }
                            if (region == "Picture Sample 2")
                            {
                                byte[] data1 = (byte[])data_image.Rows[0]["Image_Data"];
                                MemoryStream ms1 = new MemoryStream(data1);
                                Image image1 = Image.FromStream(ms1);
                                image1.Save(file_dic);
                                InsertPicture_Name(ws, ws.Range["U42"], file_dic, 10);

                                byte[] data2 = (byte[])data_image.Rows[1]["Image_Data"];
                                MemoryStream ms2 = new MemoryStream(data2);
                                Image image2 = Image.FromStream(ms2);
                                image2.Save(file_dic);
                                InsertPicture_Name(ws, ws.Range["U60"], file_dic, 10);
                            }

                        }
                    }


                }
            }
        }
        public AutoCompleteStringCollection Get_zone_infor_to_input_image(string workbook_file_address, string work_sheet)
        {
            string folder = Path.GetDirectoryName(workbook_file_address);
            string file_name = Path.GetFileName(workbook_file_address);
            myExcel.Workbook wb = TDMK_Code.open_excel_file(folder, file_name, "");
            myExcel.Worksheet ws = wb.Sheets[work_sheet];

            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 20; i < 400; i++)
            {
                if (ws.Cells[i, 2].Value != null)
                {
                    if ((ws.Cells[i, 2].Value.Contains("Region/Zone")))
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
                    if (ws.Cells[i, 2].Value.Contains("PTH_X-section"))
                    {
                        list.Add(ws.Cells[i, 2].Value.ToString() + "+" + (i).ToString());
                    }
                    if (ws.Cells[i, 2].Value.Contains("BVH_X-section"))
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
        public AutoCompleteStringCollection Get_zone_infor_to_input_value(string workbook_file_address, string work_sheet)
        {
            string folder = Path.GetDirectoryName(workbook_file_address);
            string file_name = Path.GetFileName(workbook_file_address);
            myExcel.Workbook wb = TDMK_Code.open_excel_file(folder, file_name, "");
            myExcel.Worksheet ws = wb.Sheets[work_sheet];

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
                        //listBox1.Items.Add(ws.Cells[i + 1, 1].Value.ToString() + "_" + (i + 1).ToString() + "_" + (dem - 1).ToString());
                    }
                }
            }
            return list;
        }

        public AutoCompleteStringCollection Get_ACF_infor_to_input_value(myExcel.Worksheet ws)
        {
            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 2; i < 100; i++)
            {
                if (ws.Cells[i, 1].Value != null)
                {

                    if (ws.Cells[i, 1].Value.Contains("Photo before cleaning"))
                        list.Add(ws.Cells[i, 1].Value.ToString() + "+" + "B" + (i) + "+10");
                    if (ws.Cells[i, 1].Value.Contains("Photo after cleaning"))
                        list.Add(ws.Cells[i, 1].Value.ToString() + "+" + "B" + (i) + "+10");
                    if (ws.Cells[i, 1].Value.Contains("Photo after OQC testing"))
                        list.Add(ws.Cells[i, 1].Value.ToString() + "+" + "B" + (i) + "+10");
                    if (ws.Cells[i, 1].Value.Contains("Before tape test"))
                        list.Add(ws.Cells[i, 1].Value.ToString() + "+" + "B" + (i) + "+5");
                    if (ws.Cells[i, 1].Value.Contains("After tape test"))
                        list.Add(ws.Cells[i, 1].Value.ToString() + "+" + "B" + (i) + "+5");
                    //if (ws.Cells[i, 1].Value.Contains("Tape test picture"))
                    //    list.Add(ws.Cells[i, 1].Value.ToString() + "+" + "B" + (i) + "+5" );
                }
            }
            return list;
        }
        public void Export_DatatableImage_Excel(System.Data.DataTable dt, string Image_Col_name, myExcel.Worksheet tar_wrksht, string tar_rgn, bool row_offset)
        {
            int num = 0;
            myExcel.Range range = ((myExcel._Worksheet)tar_wrksht).get_Range((object)tar_rgn, Type.Missing);
            foreach (DataRow row2 in dt.Rows)
            {
                myExcel.Range mergeArea = range.MergeArea;
                byte[] byteArrayIn = (byte[])row2[Image_Col_name];
                Image image = byteArrayToImage(byteArrayIn);
                Clipboard.SetDataObject(image, copy: false);
                int count = tar_wrksht.Shapes.Count;
                tar_wrksht.Paste(mergeArea, image);
                myExcel.Shape shape = tar_wrksht.Shapes.Item(count + 1);
                shape.LockAspectRatio = MsoTriState.msoFalse;
                float num2 = 0f;
                float num3 = 0f;
                myExcel.Range range2 = (dynamic)mergeArea.Cells[1, 1];
                int row = range2.Row;
                for (int i = 0; i < mergeArea.Rows.Count; i++)
                {
                    myExcel.Range range3 = (dynamic)tar_wrksht.Rows[row + i, Type.Missing];
                    num2 += (float)(dynamic)range3.Height;
                }

                int column = range2.Column;
                for (int j = 0; j < mergeArea.Columns.Count; j++)
                {
                    myExcel.Range range4 = (dynamic)tar_wrksht.Columns[column + j, Type.Missing];
                    num3 += (float)(dynamic)range4.Width;
                }

                float num4 = shape.Width / shape.Height;
                float num5 = num3 / num2;
                float num6 = num4 / num5;
                float num7 = ((!(num5 > 1f)) ? ((float)((double)num3 * 0.05)) : ((float)((double)num2 * 0.05)));
                if (num6 > 1f)
                {
                    shape.Width = num3 - 2f * num7;
                    shape.Height = num3 / num5 - 2f * num7;
                }
                else
                {
                    shape.Height = num2 - 2f * num7;
                    shape.Width = num2 * num5 - 2f * num7;
                }

                shape.Top = (float)(dynamic)range.Top + num7;
                shape.Left = (float)(dynamic)range.Left + num7;
                shape.LockAspectRatio = MsoTriState.msoTrue;
                int count2 = range.Columns.Count;
                int count3 = range.Rows.Count;
                range = ((!row_offset) ? range.get_Offset((object)0, (object)count2) : range.get_Offset((object)count3, (object)0));
                num++;
            }
        }

        public void Move_Value_To_Excel(string source_file_excel_address, string sheet_name_target, int col_get_data, int start_row_get_data, string des_excel_file_address, string mode_move, int start_col_to_input, int start_row_to_input, int number_val)
        {
            string excel_file_name = Path.GetFileName(source_file_excel_address);
            string source_folder = Path.GetDirectoryName(source_file_excel_address);
            string des_excel_filename = Path.GetFileName(des_excel_file_address);
            string des_folder = Path.GetDirectoryName(des_excel_file_address);
            // Thread.Sleep(100);
            OpenExcel(source_file_excel_address);
            myExcel.Worksheet ws = workbook.Sheets[1];
            myExcel.Workbook wb2 = TDMK_Code.open_excel_file(des_folder, des_excel_filename, "");
            myExcel.Worksheet ws2 = wb2.Sheets[sheet_name_target];
            // Xuất dữ liệu
            if (mode_move == "ngang")
            {
                int x = start_col_to_input;
                for (int i = start_row_get_data; i < number_val + start_row_get_data; i++)
                {
                    if (ws.Cells[i, col_get_data].Value != null)
                    {
                        ws2.Cells[start_row_to_input, x].Value = ws.Cells[i, col_get_data].Value.ToString();
                        x = x + 1;
                    }
                }
                CloseExcel(source_file_excel_address);
            }
            if (mode_move == "doc")
            {
                int x = start_row_to_input;
                for (int i = start_row_get_data; i < number_val + start_row_get_data; i++)
                {
                    if (ws.Cells[i, col_get_data].Value != null)
                    {
                        ws2.Cells[x, start_col_to_input].Value = ws.Cells[i, col_get_data].Value.ToString();
                        x = x + 1;
                    }
                }
                CloseExcel(source_file_excel_address);
            }
            MessageBox.Show("Export dữ liệu hoàn tất!");
        }
        public DataTable Load_ACF_Graph_Force(string in_src_file, string ItemCode, string LotNo, SqlConnection sqlcon)
        {
            Dictionary<string, Dictionary<int, ACF_Process_Lib.ACF_Tape_Test>> lst_result = new Dictionary<string, Dictionary<int, ACF_Process_Lib.ACF_Tape_Test>>();
            ACF_Proc.Get_TapeTest_logfile(in_src_file, ref lst_result);
            DataTable sel_dt = TDMK_Code.Datatable_Filter(sqlcon, "ACF_GRAPH_FORCE_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo })).Clone();
            int r_inx = 0;
            List<string> lst_region = new List<string>() { "TRAI", "PHAI", "GIUA" };
            foreach (var _result in lst_result)
            {
                string region_name = _result.Key.ToUpper();
                if(lst_region.IndexOf(region_name) != -1)
                {
                    foreach (var val in _result.Value)
                    {
                        sel_dt.Rows.Add();
                        sel_dt.Rows[r_inx]["Image_Graph"] = val.Value.picture;
                        sel_dt.Rows[r_inx]["Force_Data"] = val.Value.force;
                        sel_dt.Rows[r_inx]["ID"] = r_inx + 1;
                        sel_dt.Rows[r_inx]["ItemCode"] = ItemCode;
                        sel_dt.Rows[r_inx]["LotNo"] = LotNo;
                        sel_dt.Rows[r_inx]["Region"] = _result.Key;
                        sel_dt.Rows[r_inx]["Pcs_No"] = val.Key;
                        sel_dt.Rows[r_inx]["Time_Update"] = DateTime.Now.ToString("dd-MMM-yy HH:mm:ss");
                        sel_dt.Rows[r_inx]["Operator"] = "Auto";
                        r_inx++;
                    }
                }
                else
                {
                    MessageBox.Show("Tên folder: " + region_name + " không đúng định dạng (TRAI/PHAI/GIUA)", "Thông báo");
                }
            }
            return sel_dt;
        }
        public string find_format(string in_data_loc, string ItemCode)
        {
            string result = "";
            string[] files = Directory.GetFiles(Path.Combine(in_data_loc, "0.OK2SHIP report format"), "*" + ItemCode + "*.xlsm");
            result = files[0];
            return result;
        }
        public string Find_Export_Path(string format_file, string tar_process)
        {
            string result = "";
            DirectoryInfo d = new DirectoryInfo(format_file);
            DirectoryInfo root_d = d.Parent.Parent;
            DirectoryInfo[] d_arr = root_d.GetDirectories();
            string[] myLst = d_arr.Where(x => !x.Name.Contains("TDMK")).Where(x => x.Name.ToUpper().Split('.')[1].Replace("-", "").Replace(" ", "") == tar_process.Replace("_", "")).Select(x => x.Name).ToArray();
            if (myLst.Length > 0)
            {
                result = myLst[0];
            }
            return result;
        }
        public void Export_ACF_Cleaning_Edit_22_9(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            string[] search_col = { "Photo before cleaning", "Photo after cleaning", "Photo after OQC testing", "Photo after ACF peel test" };
            myExcel.Worksheet ws;
            try
            {
                ws = wb.Sheets["ACF"];
            }
            catch (Exception e)
            {
                ws = wb.Sheets["OQC ACF"];
            }
            foreach (string zone in search_col)
            {
                int max_image = 10;
                int input_image = 10;
                int row_input = FindValueInColumn(ws, 1, zone, 5, 110);

                if (row_input > 1)
                {
                    string[] item = { "ItemCode", "LotNo", "Region" };
                    string[] item_val = { Itemcode, Lotno, zone };
                    System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "ACF_CLEANING_IMAGE", TDMK_Code.filter_str(item, item_val));

                    {
                        if (data_image.Rows.Count < 10)
                        {
                            input_image = data_image.Rows.Count;
                        }
                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        string file_dic = Path.Combine(file_folder, "1.jpg");

                        myExcel.Range curr_rgn = ws.Range["B" + row_input];
                        for (int i = 0; i < input_image; i++)
                        {
                            byte[] data = (byte[])data_image.Rows[i]["Image_Data"];
                            MemoryStream ms = new MemoryStream(data);
                            Image image = Image.FromStream(ms);
                            image.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                            curr_rgn = curr_rgn.Offset[0, 1];
                        }
                    }
                }
            }
        }
        public int FindValueInColumn(myExcel.Worksheet worksheet, int col_search, string searchValue, int startRow, int endRow)
        {
            for (int row = startRow; row <= endRow; row++)
            {
                string cellValue = worksheet.Cells[row, col_search].Value?.ToString();
                if (cellValue != null && cellValue.Trim().Equals(searchValue, StringComparison.OrdinalIgnoreCase))
                {
                    return row;
                }
            }

            // Nếu không tìm thấy, trả về -1 để thể hiện không tìm thấy
            return -1;
        }

    }
}
