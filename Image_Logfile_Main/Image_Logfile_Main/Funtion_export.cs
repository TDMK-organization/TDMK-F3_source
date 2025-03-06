using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TDMK_SEEV_DLL;
using TDMK_SQL;
using myExcel = Microsoft.Office.Interop.Excel;
using Image = System.Drawing.Image;
using Microsoft.Office.Core;
using System.Security.Policy;
using System.Windows.Forms;
using System.Security.Cryptography;
using System.Web.UI.Design.WebControls;
using VHX;
using OK2SHIP_Lib;
using DataTable = System.Data.DataTable;

namespace OK2SHIP
{
    public class Funtion_export
    {
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SEI_Lib myCode = new SEI_Lib();
        myVar exp_proc = new myVar();
        TDMK_OK2SHIP TDMK_OK2SHIP = new TDMK_OK2SHIP();
        public void Export_CQRA_bHast_Image(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            int row_input = 0;
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

        public void Export_CQRA_Dielectric_withstanding(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            int dem = 0;
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
                        System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_DIELECTRIC_WITHSTANDING_IMAGE", TDMK_Code.filter_str(item, item_val));

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

        public void export_CQRA_Solderability_Image(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            int dem = 0;

            myExcel.Worksheet ws = wb.Sheets[1];

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

                        myExcel.Range curr_rgn = ws.Cells[i + 1, j + 2];

                        string[] item = { "ItemCode", "LotNo" };
                        string[] item_val = { Itemcode, Lotno };
                        System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_SOLDERABILITY_IMAGE", TDMK_Code.filter_str(item, item_val));

                        int number_image = new int[] { dem, data_image.Rows.Count }.Min();

                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        string file_dic = Path.Combine(file_folder, "1.jpg");
                        myExcel.Range rgn_judge = ws.Cells[i + 1, j + 3];

                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data1 = (byte[])data_image.Rows[m]["Image_Data"];
                            MemoryStream ms1 = new MemoryStream(data1);
                            Image image1 = Image.FromStream(ms1);
                            image1.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn, file_dic, 6);
                            rgn_judge.Offset[m, 0].Value = data_image.Rows[m]["Remark"];
                            curr_rgn = curr_rgn.Offset[1, 0];
                        } 
                        return;

                    }
                }
            }


        }



        public void Export_Chemical_resist_Image(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            int dem = 0;

            myExcel.Worksheet ws = wb.Sheets[1];

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
                        myExcel.Range curr_rgn = ws.Cells[i + 1, j - 1];
                        myExcel.Range rgn_judge = ws.Cells[i + 1, j - 2];
                        foreach (string zone in list)
                        {
                            string[] item = { "ItemCode", "LotNo", "Region" };
                            string[] item_val = { Itemcode, Lotno, zone };
                            System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_CHEMICAL_RESISTANCE_IMAGE", TDMK_Code.filter_str(item, item_val));

                            int number_image = new int[] { dem, data_image.Rows.Count }.Min();

                            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                            string file_dic = Path.Combine(file_folder, "1.jpg");


                            for (int m = 0; m < number_image; m++)
                            {
                                byte[] data1 = (byte[])data_image.Rows[m]["Image_Data"];
                                MemoryStream ms1 = new MemoryStream(data1);
                                Image image1 = Image.FromStream(ms1);
                                image1.Save(file_dic);
                                InsertPicture_Name(ws, curr_rgn, file_dic, 6);

                                curr_rgn = curr_rgn.Offset[1, 0];
                                rgn_judge.Offset[m, 0].Value = data_image.Rows[m]["Remark"];
                            }
                            curr_rgn = ws.Cells[i + 1, j];

                        } 

                        return;

                    }
                }
            }


        }

        public void Export_ACF_Cleaning(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            string[] search_col = { "Photo before cleaning", "Photo after cleaning", "Photo after OQC testing", "Photo after ACF peel test" };
            myExcel.Worksheet ws = wb.Sheets[1];
            int input_image = 10;
            int row_input = FindValueInColumn(ws, 1, "Photo before cleaning", 30, 110);

            if (row_input > 1)
            {
                string[] item = { "ItemCode", "LotNo", "Region" };
                string[] item_val = { Itemcode, Lotno, "Photo before cleaning" };
                System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "ACF_CLEANING_IMAGE", TDMK_Code.filter_str(item, item_val));

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
                    InsertPicture_Name(ws, curr_rgn.Offset[1, 0], file_dic, 5);
                    curr_rgn = curr_rgn.Offset[0, 1];
                }

                curr_rgn = curr_rgn.Offset[1, -input_image];
                System.Data.DataTable data_image_trai = TDMK_Code.Datatable_Filter(sqlcon, "ACF_AFTER_TAPE_TEST", TDMK_Code.filter_str(item, new string[] { Itemcode, Lotno, "TRAI" }));
                System.Data.DataTable data_image_giua = TDMK_Code.Datatable_Filter(sqlcon, "ACF_AFTER_TAPE_TEST", TDMK_Code.filter_str(item, new string[] { Itemcode, Lotno, "GIUA" }));
                System.Data.DataTable data_image_phai = TDMK_Code.Datatable_Filter(sqlcon, "ACF_AFTER_TAPE_TEST", TDMK_Code.filter_str(item, new string[] { Itemcode, Lotno, "PHAI" }));

                for (int i = 0; i < input_image; i++)
                {
                    byte[] data = null;
                    if (i == 0 || i == 1 || i == 2)
                    {
                        data = (byte[])data_image_trai.Rows[i]["Image_Data"];
                    }
                    else if (i == 3 || i == 4 || i == 5 || i == 6)
                    {
                        data = (byte[])data_image_giua.Rows[i]["Image_Data"];
                    }
                    else if (i == 7 || i == 8 || i == 9)
                    {
                        data = (byte[])data_image_phai.Rows[i]["Image_Data"];
                    }


                    if (data != null)
                    {
                        MemoryStream ms = new MemoryStream(data);
                        Image image = Image.FromStream(ms);
                        image.Save(file_dic);
                        InsertPicture_Name(ws, curr_rgn, file_dic, 5);

                        curr_rgn = curr_rgn.Offset[0, 1];
                    }

                }

            }
        }


        public void export_ACF_ALL_CAMERA(myExcel.Workbook wb, string Itemcode, string Lotno, SqlConnection sqlcon)
        {
            myExcel.Worksheet ws = wb.Sheets[1];
            int count = 0;
            string[] lst_region = new string[] { "TRAI", "GIUA", "PHAI" };
            myExcel.Range curr_rgn = null;
            string[] item = { "ItemCode", "LotNo", "Region" };
            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = Path.Combine(file_folder, "1.jpg");

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

                        int number_image = new int[] { 5, data_image_force.Rows.Count, data_image_glass_coupon.Rows.Count }.Min();
                        string glass_coupon_addr = TDMK_OK2SHIP.Find_Cell_Addr("photos of glass coupon after test", ws.Cells[i, j].Address, ws, true);
                        string after_tape_test_addr = TDMK_OK2SHIP.Find_Cell_Addr("photos ACF pad of flex after test", ws.Cells[i, j].Address, ws, true);
                        myExcel.Range glass_coupon_rgn = ws.Range[glass_coupon_addr].Offset[1, 0];
                        myExcel.Range after_tape_test_rgn = ws.Range[after_tape_test_addr].Offset[1, 0];

                        curr_rgn = ws.Cells[i + 1, j];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image_force.Rows[m]["Image_Graph"];
                            byte[] data_tape_test = (byte[])data_image_tape_test.Rows[m]["Image_Data"];
                            byte[] data_glass_coupon = (byte[])data_image_glass_coupon.Rows[m]["Image_Data"];
                            string value = data_image_force.Rows[m]["Force_Data"].ToString();

                            // export graph
                            MemoryStream ms = new MemoryStream(data);
                            Image image = Image.FromStream(ms);
                            image.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn, file_dic, 5);

                            // export glass coupon
                            MemoryStream ms3 = new MemoryStream(data_glass_coupon);
                            Image image3 = Image.FromStream(ms3);
                            image3.Save(file_dic);
                            //InsertPicture_Name(ws, curr_rgn.Offset[0, 3], file_dic, 5);
                            InsertPicture_Name(ws, glass_coupon_rgn.Offset[m,0], file_dic, 5);


                            // export tape test
                            MemoryStream ms2 = new MemoryStream(data_tape_test);
                            Image image2 = Image.FromStream(ms2);
                            image2.Save(file_dic);
                            //InsertPicture_Name(ws, curr_rgn.Offset[0, 4], file_dic, 5);
                            InsertPicture_Name(ws, after_tape_test_rgn.Offset[m,0], file_dic, 5);
                            curr_rgn.Offset[0, 1].Value = value;
                            curr_rgn = curr_rgn.Offset[1, 0];
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

                curr_rgn = ws.Range["B" + row_input];
                for (int i = 0; i < input_image; i++)
                {
                    byte[] data = (byte[])data_image.Rows[i]["Image_Data"];
                    MemoryStream ms = new MemoryStream(data);
                    Image image = Image.FromStream(ms);
                    image.Save(file_dic);
                    InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                    InsertPicture_Name(ws, curr_rgn.Offset[1, 0], file_dic, 5);
                    curr_rgn = curr_rgn.Offset[0, 1];
                }


                //Export "Photo after ACF peel test"
                curr_rgn = curr_rgn.Offset[2, -input_image];
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
                            if(i < data_image_trai.Rows.Count)
                            {
                                data = (byte[])data_image_trai.Rows[i]["Image_Data"];
                            } 
                        }
                        else if (i == 3 || i == 4 || i == 5 || i == 6)
                        {
                            if(i - 3 < data_image_giua.Rows.Count)
                            {
                                data = (byte[])data_image_giua.Rows[i - 3]["Image_Data"];
                            }
                            
                        }
                        else if (i == 7 || i == 8 || i == 9)
                        {
                            if(i - 7 < data_image_phai.Rows.Count)
                            {
                                data = (byte[])data_image_phai.Rows[i - 7]["Image_Data"];
                            } 
                        }

                        if (data != null)
                        {
                            MemoryStream ms = new MemoryStream(data);
                            Image image = Image.FromStream(ms);
                            image.Save(file_dic);
                            InsertPicture_Name(ws, curr_rgn, file_dic, 5);
                            curr_rgn = curr_rgn.Offset[0, 1];
                        }
                    }
                }
                else if (count == 1)
                {
                    for (int i = 0; i < 10; i++)
                    {
                        if(i < data_image_trai.Rows.Count)
                        {
                            byte[] data = (byte[])data_image_trai.Rows[i]["Image_Data"];

                            //string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                            //string file_dic = Path.Combine(file_folder, "1.jpg");

                            if (data != null)
                            {
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

        //public void export_ACF_ALL_CAMERA_old(myExcel.Workbook wb, string Itemcode, string Lotno, SqlConnection sqlcon)
        //{
        //    myExcel.Worksheet ws = wb.Sheets[1];
        //    int count = 0;
        //    string[] lst_region = new string[] { "TRAI", "GIUA", "PHAI" };
        //    myExcel.Range curr_rgn = null;
        //    string[] item = { "ItemCode", "LotNo", "Region" };
        //    string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
        //    string file_dic = Path.Combine(file_folder, "1.jpg");

        //    for (int i = 50; i < 100; i++)
        //    {
        //        for (int j = 1; j < 15; j++)
        //        {
        //            if (myCode.checkDBNull(ws.Cells[i, j].Value).ToString().Replace(" ", "").ToUpper().Contains("Force-displacement profile".Replace(" ", "").ToUpper()))
        //            {
        //                string zone = lst_region[count];

        //                string[] item_val = { Itemcode, Lotno, zone };
        //                System.Data.DataTable data_image_force = TDMK_Code.Datatable_Filter(sqlcon, "ACF_GRAPH_FORCE_IMAGE", TDMK_Code.filter_str(item, item_val));
        //                System.Data.DataTable data_image_tape_test = TDMK_Code.Datatable_Filter(sqlcon, "ACF_AFTER_TAPE_TEST_IMAGE", TDMK_Code.filter_str(item, item_val));
        //                //  System.Data.DataTable data_image_glass_coupon = TDMK_Code.Datatable_Filter(sqlcon, "ACF_GLASS_COUPON_IMAGE", TDMK_Code.filter_str(item, item_val));

        //                // int number_image = new int[] { 5, data_image_force.Rows.Count , data_image_tape_test.Rows.Count }.Min();


        //                curr_rgn = ws.Cells[i + 1, j];
        //                for (int m = 0; m < 5; m++)
        //                {
        //                    // export graph
        //                    if (data_image_force.Rows.Count > m)
        //                    {
        //                        byte[] data = (byte[])data_image_force.Rows[m]["Image_Graph"];
        //                        string value = data_image_force.Rows[m]["Force_Data"].ToString();
        //                        MemoryStream ms = new MemoryStream(data);
        //                        Image image = Image.FromStream(ms);
        //                        image.Save(file_dic);
        //                        InsertPicture_Name(ws, curr_rgn, file_dic, 5);
        //                        curr_rgn.Offset[0, 1].Value = value;
        //                    }

        //                    // export glass coupon
        //                    if (data_image_tape_test.Rows.Count > m + 5)
        //                    {
        //                        byte[] data_glass_coupon = (byte[])data_image_tape_test.Rows[m + 5]["Image_Data"];
        //                        MemoryStream ms3 = new MemoryStream(data_glass_coupon);
        //                        Image image3 = Image.FromStream(ms3);
        //                        image3.Save(file_dic);
        //                        InsertPicture_Name(ws, curr_rgn.Offset[0, 3], file_dic, 5);
        //                    }



        //                    // export tape test
        //                    if (data_image_tape_test.Rows.Count > m)
        //                    {
        //                        byte[] data_tape_test = (byte[])data_image_tape_test.Rows[m]["Image_Data"];
        //                        MemoryStream ms2 = new MemoryStream(data_tape_test);
        //                        Image image2 = Image.FromStream(ms2);
        //                        image2.Save(file_dic);
        //                        InsertPicture_Name(ws, curr_rgn.Offset[0, 4], file_dic, 5);
        //                    }


        //                    curr_rgn = curr_rgn.Offset[1, 0];
        //                }

        //                count++;
        //            }
        //            if (count == 3)
        //                goto lbl_export_next;

        //        }

        //    }
        //lbl_export_next:
        //    // export ECF cleaning

        //    int row_input = FindValueInColumn(ws, 1, "Photo before cleaning", 30, 110);

        //    if (row_input > 1)
        //    {
        //        string[] item_val = { Itemcode, Lotno, "Photo before cleaning" };
        //        int input_image = 10;

        //        System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "ACF_CLEANING_IMAGE", TDMK_Code.filter_str(item, item_val));

        //        if (data_image.Rows.Count < 10)
        //        {
        //            input_image = data_image.Rows.Count;
        //        }

        //        curr_rgn = ws.Range["B" + row_input];
        //        for (int i = 0; i < input_image; i++)
        //        {
        //            byte[] data = (byte[])data_image.Rows[i]["Image_Data"];
        //            MemoryStream ms = new MemoryStream(data);
        //            Image image = Image.FromStream(ms);
        //            image.Save(file_dic);
        //            InsertPicture_Name(ws, curr_rgn, file_dic, 5);
        //            InsertPicture_Name(ws, curr_rgn.Offset[1, 0], file_dic, 5);
        //            curr_rgn = curr_rgn.Offset[0, 1];
        //        }


        //        //Export "Photo after ACF peel test"
        //        curr_rgn = curr_rgn.Offset[2, -input_image];
        //        System.Data.DataTable data_image_trai = TDMK_Code.Datatable_Filter(sqlcon, "ACF_AFTER_TAPE_TEST_IMAGE", TDMK_Code.filter_str(item, new string[] { Itemcode, Lotno, "TRAI" }));
        //        System.Data.DataTable data_image_giua = TDMK_Code.Datatable_Filter(sqlcon, "ACF_AFTER_TAPE_TEST_IMAGE", TDMK_Code.filter_str(item, new string[] { Itemcode, Lotno, "GIUA" }));
        //        System.Data.DataTable data_image_phai = TDMK_Code.Datatable_Filter(sqlcon, "ACF_AFTER_TAPE_TEST_IMAGE", TDMK_Code.filter_str(item, new string[] { Itemcode, Lotno, "PHAI" }));

        //        if (count == 1)
        //        {
        //            for (int i = 0; i < 10; i++)
        //            {
        //                byte[] data = null;

        //                if (i == 0 || i == 1 || i == 2)
        //                {
        //                    data = (byte[])data_image_trai.Rows[i]["Image_Data"];
        //                }
        //                else if (i == 3 || i == 4 || i == 5 || i == 6)
        //                {
        //                    data = (byte[])data_image_giua.Rows[i - 3]["Image_Data"];
        //                }
        //                else if (i == 7 || i == 8 || i == 9)
        //                {
        //                    data = (byte[])data_image_phai.Rows[i - 7]["Image_Data"];
        //                }

        //                if (data != null)
        //                {
        //                    MemoryStream ms = new MemoryStream(data);
        //                    Image image = Image.FromStream(ms);
        //                    image.Save(file_dic);
        //                    InsertPicture_Name(ws, curr_rgn, file_dic, 5);
        //                    curr_rgn = curr_rgn.Offset[0, 1];
        //                }
        //            }
        //        }
        //        else if (count == 3)
        //        {
        //            for (int i = 0; i < 10; i++)
        //            {
        //                byte[] data = (byte[])data_image_trai.Rows[i]["Image_Data"];

        //                //string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
        //                //string file_dic = Path.Combine(file_folder, "1.jpg");

        //                if (data != null)
        //                {
        //                    MemoryStream ms = new MemoryStream(data);
        //                    Image image = Image.FromStream(ms);
        //                    image.Save(file_dic);
        //                    InsertPicture_Name(ws, curr_rgn, file_dic, 5);
        //                    curr_rgn = curr_rgn.Offset[0, 1];

        //                }
        //            }
        //        }
        //    }
        //}

        public void Export_IR_via_to_via_Image(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            // myExcel.Worksheet ws = wb.Sheets["CQRA - IR via to via"];
            int row_input = 0;
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
                    MessageBox.Show(new Form { TopMost = true }, "Hoàn thành!");
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

        public void Export_IR_layer_to_layer_Image(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno)
        {
            int row_input = 0;
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
           
            return -1;
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

        public void export_formatmoi_13_12(string table_name, myExcel.Worksheet ws, string Itemcode, string Lotno, string zone, SqlConnection sqlcon)
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
                        MessageBox.Show(new Form { TopMost = true }, "Hoàn thành!");

                        return;
                    }
                }
            }
        }
    }
}
