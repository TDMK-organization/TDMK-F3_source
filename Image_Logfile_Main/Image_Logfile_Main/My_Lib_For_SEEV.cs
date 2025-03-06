using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using TDMK_SEEV_DLL;
using TDMK_SQL;
//using TDMK_DLL;
using System.Data.SqlClient;
using System.Diagnostics.Tracing;
using System.Runtime.InteropServices;
using System.IO;
using myExcel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Linq.Expressions;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.AxHost;
using System.Windows.Forms.VisualStyles;
using System.Collections;
using excel = Microsoft.Office.Interop.Excel;
using System.Security.Policy;
using System.Security.Cryptography;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using ZedGraph;
//using OfficeOpenXml;
using System.Runtime.Remoting.Channels;

using System.Runtime.InteropServices.ComTypes;

//using TDMK_SEEV_DLL;
namespace Diem_Lib
{
    
    public class Export_Image_Class
    {
        
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        //SEI_Lib myCode = new SEI_Lib();
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
                excel.Range refer_range = tar_range.MergeArea;//Range["G16"];
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
                string[] item_val = { Itemcode, Lotno};
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
       // public TDMK_Class tdmk_code = new TDMK_Class();

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

        public void Export_ACF_Image(SqlConnection sqlcon, myExcel.Workbook wb, string excel_file_addr, string Itemcode, string Lotno)
        {
            myExcel.Worksheet ws = wb.Sheets["ACF"];
            AutoCompleteStringCollection list = Get_ACF_infor_to_input_value(excel_file_addr, "ACF");
            for (int i = 0; i < list.Count; i++)
            {
                //listBox1.Items.Add(list[i]);
            }
            foreach (string x in list)
            {
                try
                {
                    string region = x.Split('+')[0]; // region
                    string row_val = x.Split('+')[1]; // B16
                    int number_image = int.Parse(x.Split('+')[2]); // 10 or 5
                    string[] item = { "ItemCode", "LotNo", "Region" };
                    string[] item_val = { Itemcode, Lotno, region };
                    System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "ACF_IMAGE", TDMK_Code.filter_str(item, item_val));
                    DataGridView DGV = new DataGridView();
                    DGV.DataSource = data_image;
                    if (data_image.Rows.Count > 0)
                    {
                        if (data_image.Rows.Count < number_image)
                        {
                            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                            string file_dic = Path.Combine(file_folder, "1.jpg");
                            myExcel.Range curr_rgn = ws.Range[row_val];
                            for (int m = 0; m < data_image.Rows.Count; m++)
                            {
                                byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                MemoryStream ms = new MemoryStream(data);
                                Image image = Image.FromStream(ms);
                                image.Save(file_dic);

                                InsertPicture_Name(ws, curr_rgn, file_dic, 10);
                                curr_rgn = curr_rgn.Offset[0, 1];
                            }
                        }
                        else
                        {
                            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                            string file_dic = Path.Combine(file_folder, "1.jpg");
                            myExcel.Range curr_rgn = ws.Range[row_val];
                            for (int m = 0; m < number_image; m++)
                            {
                                byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                                MemoryStream ms = new MemoryStream(data);
                                Image image = Image.FromStream(ms);
                                image.Save(file_dic);

                                InsertPicture_Name(ws, curr_rgn, file_dic, 10);
                                curr_rgn = curr_rgn.Offset[0, 1];
                            }
                        }

                    }

                }
                catch (Exception)
                {

                }
            }
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
                if (x.Contains("H_X") == false)
                {
                    region = "Zone_" + region;
                } 
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

        public void Export_Solder_Image(SqlConnection sqlcon, myExcel.Workbook wb, string excel_file_addr, string Itemcode, string Lotno)
        {            
            myExcel.Worksheet ws;
            try
            {
                ws = wb.Sheets["Solder Mask"];
            }
            catch (Exception)
            {
                ws = wb.Sheets["Solder Mask "];
            }
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
        public void Export_High_Speed_Ball_Shear_image(SqlConnection sqlcon, myExcel.Workbook wb, string excel_file_addr, string Itemcode, string Lotno)
        {
            myExcel.Worksheet ws;
            try
            {
                ws = wb.Sheets["CQRA - High speed ball shear"];
            }
            catch (Exception)
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
        /// add function
        public int row_input;
        public void Export_CQRA_Solderability_Image(SqlConnection sqlcon, myExcel.Workbook wb, string excel_file_addr, string Itemcode, string Lotno)
        {           
            myExcel.Worksheet ws;
            try
            {
                ws = wb.Sheets["CQRA - Solderability"];
            }
            catch (Exception)
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
            System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "SOLDERABILITY_IMAGE", TDMK_Code.filter_str(item, item_val));
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
        public void Export_Thermal_Stress_Image(SqlConnection sqlcon, myExcel.Workbook wb, string excel_file_addr, string Itemcode, string Lotno)
        {
            // Get zone to input data
            AutoCompleteStringCollection list = new AutoCompleteStringCollection();            
            myExcel.Worksheet ws;
            try
            {
                ws = wb.Sheets["CQRA - Thermal stress"];
            }
            catch (Exception)
            {
                ws = wb.Sheets["CQRA - Thermal stress "];
            }
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
        public void Export_IR_via_to_via_Image(SqlConnection sqlcon, myExcel.Workbook wb, string excel_file_addr, string Itemcode, string Lotno)
        {
            myExcel.Worksheet ws = wb.Sheets["CQRA - IR via to via"];
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
            System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "IR_VIA_TO_VIA_IMAGE", TDMK_Code.filter_str(item, item_val));
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
        public void Export_IR_trace_to_trace_Image(SqlConnection sqlcon, myExcel.Workbook wb, string excel_file_addr, string Itemcode, string Lotno)
        {
            myExcel.Worksheet ws = wb.Sheets["CQRA - IR trace to trace"];
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
            System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "IR_TRACE_TO_TRACE_IMAGE", TDMK_Code.filter_str(item, item_val));
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
        public void Export_IR_layer_to_layer_Image(SqlConnection sqlcon, myExcel.Workbook wb, string excel_file_addr, string Itemcode, string Lotno)
        {
           
            myExcel.Worksheet ws;
            try
            {
                ws = wb.Sheets["CQRA - IR layer to layer"];
            }
            catch (Exception)
            {
                ws = wb.Sheets["CQRA - IR layer to layer "];
            }
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
            System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "IR_LAYER_TO_LAYER_IMAGE", TDMK_Code.filter_str(item, item_val));
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
        
        public void Export_Chemical_resist_Image(SqlConnection sqlcon, myExcel.Workbook wb, string excel_file_addr, string Itemcode, string Lotno)
        {

            string[] list = { "Picture before", "Picture after" };
            myExcel.Worksheet ws;
            try
            {
                ws = wb.Sheets["CQRA - Chemical resistance"];
            }
            catch (Exception)
            {
                ws = wb.Sheets["CQRA - Chemical resist "];
            }          

            foreach (string zone in list)
            {
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

                int number_image = 5;
                string[] item = { "ItemCode", "LotNo", "Region" };
                string[] item_val = { Itemcode, Lotno, zone };
                System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CHEMICAL_RESIST_IMAGE", TDMK_Code.filter_str(item, item_val));
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
        public void Save_High_speed_Image_LogFile_toDB(string Itemcode, string Lot, string folder,SqlConnection sqlcon)
        {
            string tablename = "CQRA_HIGH_SPEED_BALL_SHEAR_IMAGE";
            string[] file_image = Directory.GetFiles(folder, "*.PNG");
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
                string Region = "Graph";

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
  
        public void Save_Image_LogFile_toDB(string Itemcode, string Lot, string folder, SqlConnection sqlcon,string tablename,string Region)
        {
            //string tablename = "CQRA_HIGH_SPEED_BALL_SHEAR_IMAGE";
            string[] file_image;
            file_image = Directory.GetFiles(folder, "*.JPG");
            if(file_image.Length == 0)
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

        public void Save_High_speed_Data_LogFile_toDB(string Itemcode, string Lot, string folder,SqlConnection sqlcon)
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

        public void Export_Data_High_SpeedBallShear_Data(SqlConnection sqlcon, myExcel.Workbook wb, string excel_file_addr, string Itemcode, string Lotno)
        {
            myExcel.Worksheet ws = wb.Sheets["CQRA - High speed ball shear"];
            for(int i = 3; i < 30;i++)
            {
                if (ws.Cells[i,1].Value != null)
                {
                    if (ws.Cells[i,1].Value.ToString().Contains("Shear force"))
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
            catch (Exception)
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
        public void Export_CQRA_Dielectric_withstanding(SqlConnection sqlcon, myExcel.Workbook wb, string excel_file_addr, string Itemcode, string Lotno)
        {
            myExcel.Worksheet ws;
            try
            {
                ws = wb.Sheets["CQRA - Dielectric withstanding"];
            }
            catch (Exception)
            {
                ws = wb.Sheets["CQRA - Dielectric withstanding "];
            }
            //AutoCompleteStringCollection list_col = new AutoCompleteStringCollection();
            //for (int xx =3;xx<20;xx++)
            //{
            //    for(int yy = 3;yy<20;yy++)
            //    {
            //        if (ws.Cells[xx,yy].Value!= null)
            //        {
            //            if(ws.Cells[xx, yy].Value.ToString() == "Picture")
            //            {
            //                myExcel.Range sel_rgn = ws.Cells[xx, yy];
            //                string cell = sel_rgn.AddressLocal.ToString();
            //                list_col.Add(cell.Split('$')[0]);
            //            }    
            //        }
                  
            //    }    
            //}    
            string[] arr = { "After+H" , "Before+C" };          
           
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
                        if(zone == "After")
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
                    catch(Exception )
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
        // CQRA_FLUX_RESIST_IMAGE
       
        public void Export_Flux_Resist_Image(SqlConnection sqlcon, myExcel.Workbook wb, string excel_file_addr, string Itemcode, string Lotno)
        {
            int start_with_FR4 = 24;
            int start_without_FR4 = 110;
            myExcel.Worksheet ws;
            try
            {
                ws = wb.Sheets["CQRA - Flux resist"];
            }
            catch (Exception)
            {
                ws = wb.Sheets["CQRA - Flux resist "];
            }
           
            for(int i = 1; i< 200;i++)
            {
                if (ws.Cells[i,1].Value !=  null)
                {
                    if (ws.Cells[i,1].Value.ToString().Contains("test with FR4 cover"))
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
            foreach(string l in list_withoutFR4)
            {
                string zone = l.Split('+')[0];
                string row_input = l.Split('+')[1];
                int number_image = 5;
                string[] item = { "ItemCode", "LotNo", "Region" };
                string[] item_val = { Itemcode, Lotno, zone };
                System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_FLUX_RESIST_IMAGE", TDMK_Code.filter_str(item, item_val));
                if (data_image.Rows.Count > 0)
                {
                    if (data_image.Rows.Count < number_image)
                    {
                        number_image = data_image.Rows.Count;

                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        string file_dic = Path.Combine(file_folder, "1.jpg");

                        myExcel.Range curr_rgn = ws.Range["H" + row_input];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                            MemoryStream ms = new MemoryStream(data);
                            Image image = Image.FromStream(ms);
                            image.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn, file_dic, 20);
                            curr_rgn = curr_rgn.Offset[0, 1];
                        }
                    }
                    else
                    {
                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        string file_dic = Path.Combine(file_folder, "1.jpg");

                        myExcel.Range curr_rgn = ws.Range["H" + row_input];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                            MemoryStream ms = new MemoryStream(data);
                            Image image = Image.FromStream(ms);
                            image.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn, file_dic, 20);
                            curr_rgn = curr_rgn.Offset[0, 1];
                        }
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
                    if (data_image.Rows.Count < number_image)
                    {
                        number_image = data_image.Rows.Count;

                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        string file_dic = Path.Combine(file_folder, "1.jpg");

                        myExcel.Range curr_rgn = ws.Range["H" + row_input];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                            MemoryStream ms = new MemoryStream(data);
                            Image image = Image.FromStream(ms);
                            image.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn, file_dic, 20);
                            curr_rgn = curr_rgn.Offset[0, 1];
                        }
                    }
                    else
                    {
                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        string file_dic = Path.Combine(file_folder, "1.jpg");

                        myExcel.Range curr_rgn = ws.Range["H" + row_input];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
                            MemoryStream ms = new MemoryStream(data);
                            Image image = Image.FromStream(ms);
                            image.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn, file_dic, 20);
                            curr_rgn = curr_rgn.Offset[0, 1];
                        }
                    }
                }
            }
        }
        public void Export_OQC_Test_Image(SqlConnection sqlcon, myExcel.Workbook wb, string excel_file_addr, string Itemcode, string Lotno)
        {
            string[] arr = { "Solder Mask", "PI & SUS", "Gold", "EMI" };
            myExcel.Worksheet ws;
            try
            {
                ws = wb.Sheets["OQC test"];
            }
            catch (Exception)
            {
                ws = wb.Sheets["OQC test "];
            }
         
            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for(int i = 1;i< 50;i++)
            {
                if (ws.Cells[i,1].Value != null)
                {
                    foreach(string x in arr)
                    {
                        if (ws.Cells[i, 1].Value.ToString().Contains(x))
                        {
                            list.Add((ws.Cells[i, 1].Value.ToString().Split('/')[0]).TrimEnd() + "+" + (i+2).ToString());
                        }    
                    }    
                    
                }    
            } 
           
            foreach(string xx in list)
            {
                string zone = xx.Split('+')[0];
                string row_input = xx.Split('+')[1];
                int number_image = 5;
                string[] item = { "ItemCode", "LotNo", "Region" };
                string[] item_val = { Itemcode, Lotno, zone};
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
                            InsertPicture_Name(ws, curr_rgn, file_dic,12);
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
        public void Export_CQRA_bHast_Image(SqlConnection sqlcon, myExcel.Workbook wb, string excel_file_addr, string Itemcode, string Lotno)
        {
            myExcel.Worksheet ws;
            try
            {
                ws = wb.Sheets["CQRA - bHast"];
            }
            catch (Exception)
            {
                ws = wb.Sheets["CQRA - bHast "];
            }
          
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
            System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_BHAST_IMAGE", TDMK_Code.filter_str(item, item_val));
            if (data_image.Rows.Count > 0)
            {
                if (data_image.Rows.Count < number_image)
                {
                    
                    string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                    string file_dic = Path.Combine(file_folder, "1.jpg");

                    myExcel.Range curr_rgn = ws.Range["B" + row_input];
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

        /*

        string[] item_val = { Itemcode, Lotno, region };
        System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_IMAGE", TDMK_Code.filter_str(item, item_val));
                            if (data_image.Rows.Count > 0)
                            {
                                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
        string file_dic = Path.Combine(file_folder, "1.jpg");

        myExcel.Range curr_rgn = refer_range.Offset[2, 0];
                                if (data_image.Rows.Count<2)
                                {
                                    
                                    for (int m = 0; m<data_image.Rows.Count; m++)
                                    {
                                        byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
        MemoryStream ms = new MemoryStream(data);
        Image image = Image.FromStream(ms);
        image.Save(file_dic);
                                        InsertPicture_Name(ws, curr_rgn, file_dic, 6);
        curr_rgn = curr_rgn.Offset[4, 0];
                                    }
}
                                else
{

    for (int m = 0; m < 2; m++)
    {
        byte[] data = (byte[])data_image.Rows[m]["Image_Data"];
        MemoryStream ms = new MemoryStream(data);
        Image image = Image.FromStream(ms);
        image.Save(file_dic);
        InsertPicture_Name(ws, curr_rgn, file_dic, 6);
        curr_rgn = curr_rgn.Offset[4, 0];
    }
}
                            }
                            

        */
        public void Export_Impedance_Image(SqlConnection sqlcon, myExcel.Workbook wb, string excel_file_addr, string Itemcode, string Lotno)
        {
          
            myExcel.Worksheet ws ;
            try
            {
                 ws = wb.Sheets["Impedance"];
            }
            catch(Exception)
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
                    
                    string[] item = { "ItemCode", "LotNo", "Region" };

                    foreach (string region in regions)
                    {
                        string[] item_val = { Itemcode, Lotno, region };
                        System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_IMAGE", TDMK_Code.filter_str(item, item_val));
                        if(data_image !=  null)
                        {
                            
                            if(region == "Picture Sample 1")
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
                        if(ws.Cells[i, 4].Value != null)
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

        public AutoCompleteStringCollection Get_ACF_infor_to_input_value(string workbook_file_address, string work_sheet)
        {
            string folder = Path.GetDirectoryName(workbook_file_address);
            string file_name = Path.GetFileName(workbook_file_address);
            myExcel.Workbook wb = TDMK_Code.open_excel_file(folder, file_name, "");
            myExcel.Worksheet ws = wb.Sheets[work_sheet];

            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
           
            for (int i = 2; i < 100; i++)
            {
                if (ws.Cells[i, 1].Value != null)
                {
                    
                        if(ws.Cells[i, 1].Value.Contains("Photo before cleaning"))
                            list.Add(ws.Cells[i , 1].Value.ToString() + "+" + "B" + (i) + "+10");
                        if (ws.Cells[i , 1].Value.Contains("Photo after cleaning"))
                            list.Add(ws.Cells[i , 1].Value.ToString() + "+" + "B" + (i ) + "+10");
                        if (ws.Cells[i , 1].Value.Contains("Photo after OQC testing"))
                            list.Add(ws.Cells[i , 1].Value.ToString() + "+" + "B" + (i ) + "+10");
                        if (ws.Cells[i, 1].Value.Contains("Before tape test"))
                            list.Add(ws.Cells[i, 1].Value.ToString() + "+" + "B" + (i) + "+5");
                        if (ws.Cells[i, 1].Value.Contains("After tape test"))
                            list.Add(ws.Cells[i, 1].Value.ToString() + "+" + "B" + (i) + "+5");
                        if (ws.Cells[i, 1].Value.Contains("Tape test picture"))
                            list.Add(ws.Cells[i, 1].Value.ToString() + "+" + "B" + (i) + "+5" );
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
  
  
    }
    public class Data_Log_file_Process
    {
       // public ExcelPackage excel = new ExcelPackage();
        public TDMK_SQL_Lib TDMK_Code =  new TDMK_SQL_Lib();
        public SqlConnection sql_con;
       // SEI_Lib myCode = new SEI_Lib();
        public bool IsNumber(string input)
        {
            double result;
            return double.TryParse(input, out result);
        }
        public string Browser()
        {
            string file_selected = "";
            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "All Files (*.*)|*.*";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = true;

            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                file_selected = choofdlog.FileName;
                // string[] arrAllFiles = choofdlog.FileNames; //used when Multiselect = true
                return file_selected;
            }
            return file_selected;
        }

        public void Load_StackUp_Data_To_DGV(string folder_addr, DataGridView DGV)
        {
            if(folder_addr!="")
            {
                System.Data.DataTable dt = new System.Data.DataTable();
                string[] zone_folder = Directory.GetDirectories(folder_addr);
                foreach (string folder in zone_folder)
                {
                    DirectoryInfo zone_ = new DirectoryInfo(folder);
                    string zone_name = zone_.Name;
                    string[] files = Directory.GetFiles(folder, "*.csv");
                    foreach (string file in files)
                    {
                        string file_name = Path.GetFileName(file);
                        int row = 0;
                        dt.Columns.Add("Zone_" + zone_name + "_" + file_name.Replace(".csv", ""));
                        List<string> list = Get_List_StackupLogFile_Data(file);
                        for (int i = 0; i < list.Count; i++)
                        {
                            if (dt.Rows.Count <= row)
                            {
                                dt.Rows.Add();
                            }
                            dt.Rows[row]["Zone_" + zone_name + "_" + file_name.Replace(".csv", "")] = list[i];
                            row++;
                        }
                    }
                }
                DGV.DataSource = dt;
                DGV.AutoResizeColumns();
                DGV.ReadOnly = true;
            }
        }
        string FindNearestNonEmpty_CellAbove(excel.Worksheet worksheet, excel.Range cell)
        {
            int row = cell.Row - 1; // Tìm từ dòng phía trên
            while (row >= 1)
            {
                excel.Range aboveCell = worksheet.Cells[row, cell.Column];
                if (aboveCell.Value2 != null && aboveCell.Value2.ToString() != "")
                {
                    string aboveCellValue = aboveCell.Value2.ToString();
                    string[] parts = aboveCellValue.Split('+'); // Tách chuỗi thành các phần tử con dựa trên dấu '+'
                    return parts[0].Trim(); // Lấy phần tử đầu tiên của mảng kết quả và loại bỏ khoảng trắng thừa
                }
                row--;
            }
            return null; // Không tìm thấy cell nào
        }
        public void Load_Spec_StackUp_To_DGV(string file_addr,DataGridView DGV)
        {            

                excel.Workbook wb = TDMK_Code.open_excel_file(file_addr, "", "");
                string sheet_name = "";
                string tar_sheet = "Stack-up";
                for (int i = 0; i < wb.Sheets.Count; i++)
                {
                    excel.Worksheet cur_sht = wb.Sheets[i + 1];
                    string cur_sht_name = cur_sht.Name;
                    if (tar_sheet.Replace("-", "").ToUpper() == cur_sht_name.Replace("-", "").Replace(" ", "").ToUpper())
                    {
                        sheet_name = cur_sht_name;
                        break;
                    }
                }
                if (sheet_name != "")
                {
                    excel.Worksheet ws = wb.Sheets[sheet_name];

                    DGV.Rows.Add("USL", "USL");
                    DGV.Rows.Add("LSL", "LSL");
                    DGV.Rows.Add("Normal", "Normal");

                    // Tìm cột USL
                    int col_usl = 4;
                    bool found = false;
                    for (int row = 50; row <= 120; row++)
                    {
                        for (int col = 3; col <= 7; col++)
                        {
                            excel.Range cell = ws.Cells[row, col];
                            if (cell.Value2 != null && (cell.Value2.ToString() == "USL"))
                            {
                                col_usl = col;
                                found = true;
                                break; // Thoát khỏi vòng lặp
                            }
                        }
                        if (found)
                        {
                            break; // Thoát khỏi vòng lặp ngoài cùng
                        }
                    }

                    for (int i = 5; i < 400; i++)
                    {
                        if (ws.Cells[i, 1].value != null)
                        {
                            if (ws.Cells[i, 1].value.ToString() == "Zone" && ws.Cells[i + 1, 1].value != null)
                            {
                                DGV.Columns.Add(ws.Cells[i + 1, 1].value.ToString(), ws.Cells[i + 1, 1].value.ToString());
                                for (int j = i; j < i + 20; j++)
                                {
                                    if (ws.Cells[j, col_usl].Value != null)
                                    {
                                        if (ws.Cells[j, col_usl].Value.ToString() == "USL")
                                        {
                                            DGV.Rows[0].Cells[DGV.Columns.Count - 1].Value = ws.Cells[j, col_usl + 1].Value;
                                            DGV.Rows[1].Cells[DGV.Columns.Count - 1].Value = ws.Cells[j + 1, col_usl + 1].Value;
                                            if (ws.Cells[j - 2, 2].Value2 != null && ws.Cells[j - 2, 2].Value2.ToString() != "")
                                            {
                                                if (ws.Cells[j - 2, 2].Value2.ToString() == "Total thickness")
                                                {
                                                    DGV.Rows[2].Cells[DGV.Columns.Count - 1].Value = ((ws.Cells[j - 2, 6].Value2.ToString()).Split('+')[0]).Split('±')[0];
                                                }

                                            }
                                            if (ws.Cells[j - 3, 2].Value2 != null && ws.Cells[j - 3, 2].Value2.ToString() != "")
                                            {
                                                if (ws.Cells[j - 3, 2].Value2.ToString() == "Total thickness")
                                                {
                                                    DGV.Rows[2].Cells[DGV.Columns.Count - 1].Value = ((ws.Cells[j - 3, 6].Value2.ToString()).Split('+')[0]).Split('±')[0];
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    wb.Close();
                    TDMK_Code.releaseObject(wb);
                }
                else
                {
                    MessageBox.Show("Worksheet " + tar_sheet + " not found.\r\nPlease check format", "Warning");
                }
           

        }
        public SqlConnection ConnectDB()
        {
            string app_path = System.Windows.Forms.Application.StartupPath;
            string file_config = Path.Combine(app_path, "config.txt");
            string server = File.ReadLines(file_config).Skip(0).First().Remove(0, 7);
            string user = File.ReadLines(file_config).Skip(1).First().Remove(0, 8);
            string pass = File.ReadLines(file_config).Skip(2).First().Remove(0, 9);
            string database = File.ReadLines(file_config).Skip(3).First().Remove(0, 8);

            string str_sqlcon = TDMK_Code.data_connection(server, database, user, pass).ConnectionString;
            sql_con = new SqlConnection(str_sqlcon);
            return sql_con;
        }    
        public void Save_StacKup_DGV_toDB(string item_code,string lot_no,string person,DataGridView DGV_data,SqlConnection sql_con)
        {
           
            string[] item = { "ID", "ItemCode", "Lotno", "Region", "Pcs_No", "Data","Operator" };

            if (DGV_data.Rows.Count > 1)
            {
                for (int i = 0; i < DGV_data.Columns.Count; i++)
                {
                    string col_name = DGV_data.Columns[i].Name;
                    string zone = col_name.Split('_')[0] + "_" + col_name.Split('_')[1];
                    string pcs = col_name.Split('_')[2];

                    for (int j = 0; j < DGV_data.Rows.Count; j++)
                    {
                        if (DGV_data.Rows[j].Cells[i].Value != null)
                        {
                            if (DGV_data.Rows[j].Cells[i].Value.ToString() != "")
                            {
                                string[] item_1 = { "ItemCode", "Lotno" };
                                string[] item_val_1 = { item_code, lot_no };
                                AutoCompleteStringCollection list = TDMK_Code.Load_Item_Filter_str(sql_con, "STACKUP_LOGFILE","Data" ,TDMK_Code.filter_str(item_1, item_val_1));
                               
                                string[] item_val = { (TDMK_Code.SQL_MAX("STACKUP_LOGFILE", "ID", sql_con) + 1).ToString(), item_code, lot_no, zone, pcs, DGV_data.Rows[j].Cells[i].Value.ToString(), person };

                                TDMK_Code.insert_val_arr("STACKUP_LOGFILE", sql_con, item, item_val);
                               
                            }
                        }
                    }
                }
                MessageBox.Show("Saved all to DB!");
            }
            else
                MessageBox.Show("No data!", "Warning");
        }
        public void HightLigh_StackUp_LogFile(DataGridView DGV_Spec, DataGridView DGV_Data)
        {
            List<string> spec_col_lst = new List<string>();
            foreach(DataGridViewColumn dgv_c in DGV_Spec.Columns)
            {
                spec_col_lst.Add(dgv_c.Name);
            }
            for(int i = 0; i < DGV_Data.Columns.Count; i++)
            {
                string zone = (DGV_Data.Columns[i].Name.Split('_')[1]).Replace("Zone","").Replace("_","").Trim(); 
                // Zone_A_1 ; ...[1] = A
                int col_inx = spec_col_lst.IndexOf(zone);
                if(col_inx!=-1)
                {
                    double max = Convert.ToDouble(DGV_Spec.Rows[0].Cells[col_inx].Value.ToString());
                    double min = Convert.ToDouble(DGV_Spec.Rows[1].Cells[col_inx].Value.ToString());
                    for (int m = 0; m < DGV_Data.Rows.Count - 1; m++)
                    {
                        if (DGV_Data.Rows[m].Cells[i].Value != null)
                        {
                            if (IsNumber(DGV_Data.Rows[m].Cells[i].Value.ToString()))
                            {
                                double x = Convert.ToDouble(DGV_Data.Rows[m].Cells[i].Value);
                                if (x < min || x > max)
                                {
                                    DGV_Data.Rows[m].Cells[i].Style.BackColor = Color.Red;
                                }
                                else
                                    DGV_Data.Rows[m].Cells[i].Style.BackColor = Color.White;
                            }
                        }
                    }
                }
                else
                {
                    for (int m = 0; m < DGV_Data.Rows.Count; m++)
                    {
                        if (DGV_Data.Rows[m].Cells[i].Value != null)
                        {
                            DGV_Data.Rows[m].Cells[i].Style.BackColor = Color.Gray;
                        }
                    }
                }  
            }
        }
        public void HightLigh_StackUp_LogFile_One_Zone(DataGridView DGV_Spec, DataGridView DGV_Data,string zone)
        {
           if(DGV_Spec.Rows.Count > 2)
           {
                if (DGV_Data.Rows.Count > 1)
                {
                    List<string> spec_col_lst = new List<string>();
                    foreach (DataGridViewColumn dgv_c in DGV_Spec.Columns)
                    {
                        spec_col_lst.Add(dgv_c.Name);
                    }
                    int col_inx = spec_col_lst.IndexOf(zone);
                    if (col_inx != -1)
                    {
                        for (int i = 0; i < DGV_Data.Columns.Count; i++)
                        {
                            double max = Convert.ToDouble(DGV_Spec.Rows[0].Cells[col_inx].Value.ToString());
                            double min = Convert.ToDouble(DGV_Spec.Rows[1].Cells[col_inx].Value.ToString());
                            for (int m = 0; m < DGV_Data.Rows.Count; m++)
                            {
                                if (DGV_Data.Rows[m].Cells[i].Value != null)
                                {
                                    if (IsNumber(DGV_Data.Rows[m].Cells[i].Value.ToString()))
                                    {
                                        double x = Convert.ToDouble(DGV_Data.Rows[m].Cells[i].Value);
                                        if (x < min || x > max)
                                        {
                                            DGV_Data.Rows[m].Cells[i].Style.BackColor = Color.Red;
                                        }
                                        else
                                            DGV_Data.Rows[m].Cells[i].Style.BackColor = Color.White;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        for (int m = 0; m < DGV_Data.Rows.Count; m++)
                        {
                            for (int i = 0; i < DGV_Data.Columns.Count; i++)
                            {
                                if (DGV_Data.Rows[m].Cells[i].Value != null)
                                {
                                    DGV_Data.Rows[m].Cells[i].Style.BackColor = Color.Gray;
                                }
                            }
                        }
                    }
                }
           }
           else
           {
                for (int m = 0; m < DGV_Data.Rows.Count; m++)
                {
                    for (int i = 0; i < DGV_Data.Columns.Count; i++)
                    {
                        if (DGV_Data.Rows[m].Cells[i].Value != null)
                        {
                            DGV_Data.Rows[m].Cells[i].Style.BackColor = Color.Gray;
                        }
                    }
                }
                MessageBox.Show("Spec not found!");
           }    
            
        }

        public int Find_Row_IndexOf_Value(string filePath, string value_search, int col_index)
        {
            int lineNumber = 0;
            using (StreamReader reader = new StreamReader(filePath))
            {
                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    lineNumber++;

                    string[] cells = line.Split(',');

                    if (cells.Length > 0 && cells[0].Trim() == value_search)
                    {
                        return lineNumber;
                    }
                }
            }

            return -1; // Trả về -1 nếu không tìm thấy giá trị
        }
        /// For Bhast Sheet
        public static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum;
        }

        public SortedDictionary<string, string> Get_bHast_LogFile_Data(string in_src_file,int col_get)
        {
            int line_start = Find_Row_IndexOf_Value(in_src_file, "Total Time[h]", 0);
           
            SortedDictionary<string, string> result_lst = new SortedDictionary<string, string>();
            string[] Lines = System.IO.File.ReadAllLines(in_src_file).Skip(line_start).ToArray();
            
            foreach (var line in Lines)
            {
              //  MessageBox.Show(line.ToString());
                string str = new string(line.Where(s => s != '"').ToArray());
                string[] temp = str.Split(',');
                if (temp[0] != "")
                {
                    if (IsNumeric(temp[0]))
                    {
                        result_lst.Add(Convert.ToString(temp[0]), temp[col_get -1]); // lấy cột đàu và cột cần lấy data
                    }
                }
                else
                {
                    break;
                }
            }
            return result_lst;
        }
        public List<string> GetData_In_Excel(string filePath, int column, int startRow, int endRow)
        {
            List<string> values = new List<string>();

            try
            {
                using (StreamReader reader = new StreamReader(filePath))
                {
                    int currentLine = 0;
                    while (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine();
                        if (currentLine >= startRow && currentLine <= endRow)
                        {
                            string[] rowData = line.Split(',');
                            if (column < rowData.Length)
                            {
                                values.Add(rowData[column]);
                            }
                            else
                            {
                                MessageBox.Show($"Column {column} does not exist in the CSV file.");
                                return null;
                            }
                        }
                        currentLine++;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show($"Error reading the CSV file: {e.Message}");
            }

            return values;
        }
        public List<string> Get_List_Bhast_LogFile_Data(string in_src_file,int col_get)
        {
            SortedDictionary<string, string> dict = Get_bHast_LogFile_Data( in_src_file, col_get);
            List<string> list = dict.ToList().Select(x => x.Value).ToList();
            return list;
        }
        public System.Data.DataTable Get_bHast_Data(string file_addr)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            //excel.Workbook wb = TDMK_Code.open_excel_file(file_addr, "", "");
            //excel.Worksheet ws = wb.Sheets[1];
            dt.Columns.Add("ID");
            dt.Columns.Add("ItemCode");
            dt.Columns.Add("LotNo");
            dt.Columns.Add("Time");
            dt.Columns.Add("Ch01");
            dt.Columns.Add("Ch02");
            dt.Columns.Add("Ch03");
            dt.Columns.Add("Ch04");
            dt.Columns.Add("Ch05");
            dt.Columns.Add("Ch06");
            dt.Columns.Add("Ch07");
            dt.Columns.Add("Ch08");

            string filename = Path.GetFileName(file_addr);
            string[] item_lots = filename.Split('+');

            foreach (string x in item_lots)
            {
                
                string Item_Lot = x.Replace(".CSV", "").Replace(".csv", "");
                int chanel_Start = Convert.ToInt32((Item_Lot.Split('(')[1].Replace(")", "")).Split('-')[0]);//.Remove(')');
                
                int chanel_end = Convert.ToInt32((Item_Lot.Split('(')[1].Replace(")", "")).Split('-')[1]);//.Remove(')');
               
                string itemlot = Item_Lot.Split('(')[0];
                string item = itemlot.Split('-')[0];
                string lot = itemlot.Split('-')[1];
                for(int i = 0;i< 241;i++)
                {
                    dt.Rows.Add();
                    dt.Rows[i]["ID"] = dt.Rows.Count;
                    dt.Rows[i]["ItemCode"] = item;
                    dt.Rows[i]["LotNo"] = lot;
                }
                // Cột thời gian
                int start_row = Find_Row_IndexOf_Value(file_addr, "Total Time[h]",0);
                List<string> list_val_time = GetData_In_Excel(file_addr, 0,start_row,start_row + 240);
                for (int j = 0; j < list_val_time.Count; j++)
                {
                    
                        dt.Rows[j]["Time"] = list_val_time[j];
                }
                for (int i = 0; i < 8; i++)
                {
                    List<string> list_val = GetData_In_Excel(file_addr, chanel_Start +i, start_row, start_row + 240);
                    for (int j = 0; j < list_val.Count; j++)
                    {
                        if (IsNumeric(list_val[j]) == false)
                        {
                            dt.Rows[j][i + 4] = "<1E+03";
                        }
                        else
                            dt.Rows[j][i + 4] = list_val[j];
                    }
                }


                // start_row = 41
                //for (int i = start_row + 1; i < start_row + 242; i++)
                // {


                // dt.Rows[dt.Rows.Count - 1]["Time"] = ws.Cells[i, 1].Value;


                ////Ch1
                //if (IsNumeric(ws.Cells[i, start_col].Value) == false)
                //{
                //    dt.Rows[dt.Rows.Count - 1]["Ch01"] = "<1E+03";

                //}
                //else
                //    dt.Rows[dt.Rows.Count - 1]["Ch01"] = ws.Cells[i, start_col].Value;
                ////Ch02
                //if (IsNumeric(ws.Cells[i, start_col + 1].Value) == false)
                //{
                //    dt.Rows[dt.Rows.Count - 1]["Ch02"] = "<1E+03";
                //}
                //else
                //    dt.Rows[dt.Rows.Count - 1]["Ch02"] = ws.Cells[i, start_col + 1].Value;

                ////Ch03
                //if (IsNumeric(ws.Cells[i, start_col + 2].Value) == false)
                //{
                //    dt.Rows[dt.Rows.Count - 1]["Ch03"] = "<1E+03";
                //}
                //else
                //    dt.Rows[dt.Rows.Count - 1]["Ch03"] = ws.Cells[i, start_col + 2].Value;


                ////Ch04
                //if (IsNumeric(ws.Cells[i, start_col + 3].Value) == false)
                //{
                //    dt.Rows[dt.Rows.Count - 1]["Ch04"] = "<1E+03";
                //}
                //else
                //    dt.Rows[dt.Rows.Count - 1]["Ch04"] = ws.Cells[i, start_col + 3].Value;

                ////Ch05
                //if (IsNumeric(ws.Cells[i, start_col + 4].Value) == false)
                //{
                //    dt.Rows[dt.Rows.Count - 1]["Ch05"] = "<1E+03";
                //}
                //else
                //    dt.Rows[dt.Rows.Count - 1]["Ch05"] = ws.Cells[i, start_col + 4].Value;

                ////Ch06
                //if (IsNumeric(ws.Cells[i, start_col + 5].Value) == false)
                //{
                //    dt.Rows[dt.Rows.Count - 1]["Ch06"] = "<1E+03";
                //}
                //else
                //    dt.Rows[dt.Rows.Count - 1]["Ch06"] = ws.Cells[i, start_col + 5].Value;

                ////Ch07
                //if (IsNumeric(ws.Cells[i, start_col + 6].Value) == false)
                //{
                //    dt.Rows[dt.Rows.Count - 1]["Ch07"] = "<1E+03";
                //}
                //else
                //    dt.Rows[dt.Rows.Count - 1]["Ch07"] = ws.Cells[i, start_col + 6].Value;

                ////Ch08
                //if (IsNumeric(ws.Cells[i, start_col + 7].Value) == false)
                //{
                //    dt.Rows[dt.Rows.Count - 1]["Ch08"] = "<1E+03";
                //}
                //else
                //    dt.Rows[dt.Rows.Count - 1]["Ch08"] = ws.Cells[i, start_col + 7].Value;
                //}

            }
            //DGV.DataSource = dt;
            // wb.Close();
            return dt;
        }

        // export
       
        public void BatchBulkCopy(SqlConnection sqlcon_OK2SHIP, System.Data.DataTable dataTable, string DestinationTbl)
        {
            System.Data.DataTable dtInsertRows = dataTable;
            using (SqlBulkCopy sbc = new SqlBulkCopy(sqlcon_OK2SHIP))
            {
                if (sqlcon_OK2SHIP.State != ConnectionState.Open)
                {
                    sqlcon_OK2SHIP.Open();
                }
                sbc.DestinationTableName = DestinationTbl;
                foreach (DataColumn dc in dataTable.Columns)
                {
                    sbc.ColumnMappings.Add(dc.ColumnName, dc.ColumnName);
                }
                sbc.WriteToServer(dtInsertRows);
                sqlcon_OK2SHIP.Close();
            }
        }
        public void Save_Bhast_Logfile(string logfile_addr, SqlConnection sqlcon)
        {
            sqlcon = ConnectDB();
            System.Data.DataTable data = Get_bHast_Data(logfile_addr);
            string[] item = { "ID", "ItemCode", "LotNo", "Time", "Ch01", "Ch02", "Ch03", "Ch04", "Ch05", "Ch06", "Ch07", "Ch08" };
            string[] item_val = new string[12];
            for (int i = 0; i < data.Rows.Count; i++)
            {
                item_val[0] = (TDMK_Code.SQL_MAX("Bhast_Logfile".ToUpper(), "ID", sqlcon) + 1).ToString();
                item_val[1] = data.Rows[i]["ItemCode"].ToString();
                item_val[2] = data.Rows[i]["LotNo"].ToString();
                item_val[3] = data.Rows[i]["Time"].ToString();
                item_val[4] = data.Rows[i]["Ch01"].ToString();
                item_val[5] = data.Rows[i]["Ch02"].ToString();
                item_val[6] = data.Rows[i]["Ch03"].ToString();
                item_val[7] = data.Rows[i]["Ch04"].ToString();
                item_val[8] = data.Rows[i]["Ch05"].ToString();
                item_val[9] = data.Rows[i]["Ch06"].ToString();
                item_val[10] = data.Rows[i]["Ch07"].ToString();
                item_val[11] = data.Rows[i]["Ch08"].ToString();
                TDMK_Code.insert_val_arr("Bhast_Logfile".ToUpper(), sqlcon, item, item_val);
            }
        }
        public void DrawGraph_bHast(DataGridView DGV,ZedGraphControl zedGraphControl1)
        {

            // Tạo đối tượng GraphPane
            GraphPane myPane = zedGraphControl1.GraphPane;
            myPane.CurveList.Clear();
            // Đặt tiêu đề cho biểu đồ
            myPane.Title.Text = "Pismo C4.0";

            // Đặt tên cho trục x và y
            myPane.XAxis.Title.Text = "Time (hours)";
            myPane.YAxis.Title.Text = "Resistance ( Ω)";

            // Lấy dữ liệu từ datagrid view
            double[] time = new double[DGV.RowCount - 1]; // Tạo mảng trống
            double[] ch01 = new double[DGV.RowCount - 1];
            double[] ch02 = new double[DGV.RowCount - 1];
            double[] ch03 = new double[DGV.RowCount - 1];
            double[] ch04 = new double[DGV.RowCount - 1];
            double[] ch05 = new double[DGV.RowCount - 1];
            double[] ch06 = new double[DGV.RowCount - 1];
            double[] ch07 = new double[DGV.RowCount - 1];
            double[] ch08 = new double[DGV.RowCount - 1];

            for (int i = 0; i < DGV.RowCount - 1; i++)
            {
                time[i] = Convert.ToDouble(DGV.Rows[i].Cells["Time"].Value);
                if (DGV.Rows[i].Cells["Ch01"].Value.ToString() != "<1E+03")
                {
                    ch01[i] = (Convert.ToDouble(DGV.Rows[i].Cells["Ch01"].Value)) ;
                }

                if (DGV.Rows[i].Cells["Ch02"].Value.ToString() != "<1E+03")
                {
                    ch02[i] = (Convert.ToDouble(DGV.Rows[i].Cells["Ch02"].Value)) ;
                }
                if (DGV.Rows[i].Cells["Ch03"].Value.ToString() != "<1E+03")
                {
                    ch03[i] = (Convert.ToDouble(DGV.Rows[i].Cells["Ch03"].Value));
                }
                if (DGV.Rows[i].Cells["Ch04"].Value.ToString() != "<1E+03")
                {
                    ch04[i] = (Convert.ToDouble(DGV.Rows[i].Cells["Ch04"].Value)) ;
                }
                if (DGV.Rows[i].Cells["Ch05"].Value.ToString() != "<1E+03")
                {
                    ch05[i] = (Convert.ToDouble(DGV.Rows[i].Cells["Ch05"].Value)) ;
                }
                if (DGV.Rows[i].Cells["Ch06"].Value.ToString() != "<1E+03")
                {
                    ch06[i] = (Convert.ToDouble(DGV.Rows[i].Cells["Ch06"].Value)) ;
                }
                if (DGV.Rows[i].Cells["Ch07"].Value.ToString() != "<1E+03")
                {
                    ch07[i] = (Convert.ToDouble(DGV.Rows[i].Cells["Ch07"].Value)) ;
                }
                if (DGV.Rows[i].Cells["Ch08"].Value.ToString() != "<1E+03")
                {
                    ch08[i] = (Convert.ToDouble(DGV.Rows[i].Cells["Ch08"].Value));
                }


            }

            // Vẽ đường cho các kênh
            LineItem curve1 = myPane.AddCurve("Ch01", time, ch01, Color.Blue, SymbolType.None);
            LineItem curve2 = myPane.AddCurve("Ch02", time, ch02, Color.Red, SymbolType.None);
            LineItem curve3 = myPane.AddCurve("Ch03", time, ch03, Color.Green, SymbolType.None);
            LineItem curve4 = myPane.AddCurve("Ch04", time, ch04, Color.Black, SymbolType.None);
            LineItem curve5 = myPane.AddCurve("Ch05", time, ch05, Color.Yellow, SymbolType.None);
            LineItem curve6 = myPane.AddCurve("Ch06", time, ch06, Color.Cyan, SymbolType.None);
            LineItem curve7 = myPane.AddCurve("Ch07", time, ch07, Color.Brown, SymbolType.None);
            LineItem curve8 = myPane.AddCurve("Ch08", time, ch08, Color.Orange, SymbolType.None);

            // Hiển thị biểu đồ
            zedGraphControl1.AxisChange();
        }
        public System.Data.DataTable DGVToDataTable(DataGridView DGV)
        {
            // Chuyển đổi dữ liệu từ DataGridView sang DataTable
            System.Data.DataTable dataTableFromDataGridView = new System.Data.DataTable() ;
            if (DGV.Rows.Count > 2)
            {
                dataTableFromDataGridView = ((System.Data.DataTable)DGV.DataSource).Copy();
                
            }
            return dataTableFromDataGridView;
        }
        public string Get_String_Value(string input, string searchValue)
        {
            // Đầu vào : 718830 - 03(5 - 12) + 71827G - 563(1 - 4) + 528797 - 813(16 - 23).CSV", và giá trị :  718830 - 03
            // Đầu ra : 718830 - 03(5 - 12)
            int startIndex = input.IndexOf(searchValue);
            if (startIndex == -1)
            {
                return "";
            }

            int endIndex = input.IndexOf(")", startIndex);
            if (endIndex == -1)
            {
                return "";
            }

            return input.Substring(startIndex, endIndex - startIndex + 1);
        }
        public static List<string> AutoCompleteCollection_To_List(AutoCompleteStringCollection autoCompleteCollection)
        {
            List<string> resultList = new List<string>();

            foreach (string item in autoCompleteCollection)
            {
                resultList.Add(item);
            }

            return resultList;
        }
        public async Task Export_MultipleColumns_Async(string fileName, List<List<string>> data, int startRow, int startColumn,excel.Worksheet ws)
        {
            await Task.Run(() => {
                FileInfo fileInfo = new FileInfo(fileName);

                if (!fileInfo.Exists)
                {
                    throw new FileNotFoundException("File does not exist. Please provide a valid file.");
                }

                int currentRow = startRow;
                    int currentColumn = startColumn;

                    foreach (List<string> columnData in data)
                    {
                        currentRow = startRow;
                        foreach (string item in columnData)
                        {
                        ws.Cells[currentColumn][currentRow+1].Value = IsNumeric(item) ? item.ToString() : "<1E+03";
                            currentRow++;
                        }
                        currentColumn++;
                    }
            });
        }
        public async void Export_Bhast_Logfile(string itemcode, string lotno, string excel_addr, SqlConnection sqlcon)
        {
           
            string[] item = { "ItemCode", "LotNo" };
            string[] item_val = { itemcode, lotno };
            string[] list_ch = { "Ch01", "Ch02", "Ch03", "Ch04", "Ch05", "Ch06", "Ch07", "Ch08" };
            excel.Application excelApp = new excel.Application();
            excelApp.DisplayAlerts = false; // tat canh bao            
            excel.Workbook wb = excelApp.Workbooks.Open(excel_addr);

            // Hiển thị ứng dụng Excel
            excelApp.Visible = true;
           // excel.Workbook wb = TDMK_Code.open_excel_file(excel_addr, "", "");
            excel.Worksheet ws = wb.Sheets["CQRA - bHast"];
           
            int start_col = 1;
            int start_row = 1;
            // tìm vị trí Ch01, vị trí khác offset
            for (int i = 15; i < 35; i++)
            {
                for (int j = 4; j < 20; j++)
                {
                    if (ws.Cells[i, j].Value != null)
                    {
                        if (ws.Cells[i, j].Value.ToString() == "Ch01")
                        {
                            start_row = i;
                            start_col = j; break;
                        }
                    }
                }
            }
            foreach (string chanel in list_ch)
            {
                //AutoCompleteStringCollection list = TDMK_Code.Load_Item_Filter_str(sqlcon, "CQRA_Bhast_Logfile".ToUpper(), chanel, TDMK_Code.filter_str(item, item_val));
                //if (list.Count > 1)
                //{
                //    string number = chanel.Replace("Ch0", "");
                //    int col_offset = Int32.Parse(number);

                //    for (int y = 1; y <= list.Count; y++)
                //    {
                //        ws.Cells[start_row + y, start_col + col_offset - 1].Value = IsNumeric(list[y - 1]) ? list[y - 1].ToString() : "<1E +03";
                //    }
                //}
                //else
                //    MessageBox.Show("Chưa có dữ liệu trong DataBase!");
            }

            ///// export with mutils task
            List<List<string>> data = new List<List<string>>
                {
                AutoCompleteCollection_To_List(TDMK_Code.Load_Item_Filter_str(sqlcon, "CQRA_Bhast_Logfile".ToUpper(), "Ch01", TDMK_Code.filter_str(item, item_val))),
                AutoCompleteCollection_To_List(TDMK_Code.Load_Item_Filter_str(sqlcon, "CQRA_Bhast_Logfile".ToUpper(), "Ch02", TDMK_Code.filter_str(item, item_val))),
                AutoCompleteCollection_To_List(TDMK_Code.Load_Item_Filter_str(sqlcon, "CQRA_Bhast_Logfile".ToUpper(), "Ch03", TDMK_Code.filter_str(item, item_val))),
                AutoCompleteCollection_To_List(TDMK_Code.Load_Item_Filter_str(sqlcon, "CQRA_Bhast_Logfile".ToUpper(), "Ch04", TDMK_Code.filter_str(item, item_val))),
                AutoCompleteCollection_To_List(TDMK_Code.Load_Item_Filter_str(sqlcon, "CQRA_Bhast_Logfile".ToUpper(), "Ch05", TDMK_Code.filter_str(item, item_val))),
                AutoCompleteCollection_To_List(TDMK_Code.Load_Item_Filter_str(sqlcon, "CQRA_Bhast_Logfile".ToUpper(), "Ch06", TDMK_Code.filter_str(item, item_val))),
                AutoCompleteCollection_To_List(TDMK_Code.Load_Item_Filter_str(sqlcon, "CQRA_Bhast_Logfile".ToUpper(), "Ch07", TDMK_Code.filter_str(item, item_val))),
                AutoCompleteCollection_To_List(TDMK_Code.Load_Item_Filter_str(sqlcon, "CQRA_Bhast_Logfile".ToUpper(), "Ch08", TDMK_Code.filter_str(item, item_val))),

                };
            try
            {
              
                    await Export_MultipleColumns_Async(excel_addr, data, start_row, start_col, ws);
                    MessageBox.Show("Data exported successfully.");
                
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting data: {ex.Message}");
            }
            excelApp.DisplayAlerts = true;
            excelApp.Quit();

        }
        public void Get_Stackup_Image_Multi(string in_src, ref Dictionary<string, SortedDictionary<int, byte[]>> dic_lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            FileInfo[] temp_lst = tar_d.GetFiles("*.jpg");
            if (temp_lst.Length > 0)
            {
                SortedDictionary<int, byte[]> lst_result = new SortedDictionary<int, byte[]>();
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
                    if (!f_na.Contains("-"))
                    {
                        var sel_img = Bitmap.FromFile(temp_lst[i].FullName);
                        ImageConverter imgcon = new ImageConverter();
                        byte[] img_data = (byte[])imgcon.ConvertTo(sel_img, typeof(byte[]));
                        int f_inx = Convert.ToInt32(f_na);
                        lst_result.Add(f_inx, img_data);
                    }
                }
                dic_lst_result.Add(tar_d.Name, lst_result);
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_Stackup_Image_Multi(inter_lst.FullName, ref dic_lst_result);
                    }
                }
            }
        }
        public void Get_Stackuplogfile_Multi(string in_src, ref SortedDictionary<string, SortedDictionary<int, string>> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            FileInfo[] temp_lst = tar_d.GetFiles("*.csv");
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    SortedDictionary<int, string> temp = Get_StackupLogFile_Data(temp_lst[i].FullName);
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
                    int f_inx = Convert.ToInt32(f_na);
                    lst_result.Add(f_na, temp);
                }
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_Stackuplogfile_Multi(inter_lst.FullName, ref lst_result);
                    }
                }
            }
        }

        public void Get_Stackuplogfile_Multi2(string in_src, ref Dictionary<string, Dictionary<string, SortedDictionary<int, string>>> dic_lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            FileInfo[] temp_lst = tar_d.GetFiles("*.csv");
            if (temp_lst.Length > 0)
            {
                Dictionary<string, SortedDictionary<int, string>> lst_result = new Dictionary<string, SortedDictionary<int, string>>();
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    SortedDictionary<int, string> temp = Get_StackupLogFile_Data(temp_lst[i].FullName);
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
                    lst_result.Add(f_na, temp);
                }
                dic_lst_result.Add(tar_d.Name, lst_result);
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_Stackuplogfile_Multi2(inter_lst.FullName, ref dic_lst_result);
                    }
                }
            }
        }
        public SortedDictionary<int, string> Get_StackupLogFile_Data(string in_src_file)
        {
            SortedDictionary<int, string> result_lst = new SortedDictionary<int, string>();
            string[] Lines = System.IO.File.ReadAllLines(in_src_file);
            foreach (var line in Lines)
            {
                string str = new string(line.Where(s => s != '"').ToArray());
                string[] temp = str.Split(',');
                if (temp[0] != "")
                {
                    if (IsNumeric(temp[0]))
                    {
                        result_lst.Add(Convert.ToInt32(temp[0]), temp[2]);
                    }
                }
                else
                {
                    break;
                }
            }
            return result_lst;
        }

        public List<string> Get_List_StackupLogFile_Data(string file_csv)
        {
           
            SortedDictionary<int, string> dict = Get_StackupLogFile_Data(file_csv);
            List<string> list = dict.ToList().Select(x => x.Value).ToList();
            return list;
        }
        public bool Check_Stackup_data(System.Data.DataTable stakup_tbl, System.Data.DataTable spec_tbl)
        {
            bool result = true;
            List<string> region_lst = stakup_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToList();
            foreach(var _region in region_lst)
            {
                System.Data.DataTable temp = stakup_tbl.AsEnumerable().Where(x => x.Field<string>("Region") == _region).CopyToDataTable();
                if (!Check_in_Spec_zone(temp,spec_tbl,_region))
                {
                    result = false;
                    break;
                }    
            }
            return result;
        }
        public bool Check_in_Spec_zone(System.Data.DataTable DGV_Spec, System.Data.DataTable DGV_Data, string zone)
        {
            bool result = true;
            if (DGV_Spec.Rows.Count > 2)
            {
                if (DGV_Data.Rows.Count > 1)
                {
                    List<string> spec_col_lst = new List<string>();
                    foreach (DataColumn dgv_c in DGV_Spec.Columns)
                    {
                        spec_col_lst.Add(dgv_c.ColumnName);
                    }
                    int col_inx = spec_col_lst.IndexOf(zone);
                    if (col_inx != -1)
                    {
                        for (int i = 0; i < DGV_Data.Columns.Count; i++)
                        {
                            double max = Convert.ToDouble(DGV_Spec.Rows[0][col_inx].ToString());
                            double min = Convert.ToDouble(DGV_Spec.Rows[1][col_inx].ToString());
                            for (int m = 0; m < DGV_Data.Rows.Count; m++)
                            {
                                if (DGV_Data.Rows[m][i] != null)
                                {
                                    if (IsNumber(DGV_Data.Rows[m].ToString()))
                                    {
                                        double x = Convert.ToDouble(DGV_Data.Rows[m][i]);
                                        if (x < min || x > max)
                                        {
                                            result = false;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        result = false;
                    }
                }
            }
            return result;
        }
    }
}
