using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using myExcel = Microsoft.Office.Interop.Excel;
using TDMK_SQL;

namespace OK2SHIP_Software
{
    internal class Export_VXH_ManualInput
    {
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        myVar myCode = new myVar();
        
        public SqlConnection sqlcon_OK2SHIP_Period2 = null;
        public static string sel_DB_OK2SHIP_Period2 = "OK2SHIP_Period2";

        public myExcel.Workbook create_export_wrk(string format_file, string process_name)
        {
            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            xlsApp.DisplayAlerts = false;
            myExcel.Workbook wb = TDMK_Code.open_excel_file(format_file, "", "");
            string _process = process_name.Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
            string mySheet = "";
            foreach (myExcel.Worksheet tg_sht in wb.Worksheets)
            {
                string cur_sht_name = tg_sht.Name.ToUpper().Replace("-", "").Replace(" ", "");
                if (cur_sht_name == _process)
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




        /***********************  Export Stackup Function **********************************************/
        public myExcel.Workbook export_stackup(myExcel.Workbook export_wrkbook, string ItemCode, string LotNo)
        {
           
            myExcel.Worksheet ws = export_wrkbook.Sheets[1];


            if (ws != null)
            {
                ws.Cells.Interior.Color = myExcel.XlRgbColor.rgbWhite;
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
                        System.Data.DataTable data_table = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP_Period2, "STACKUP_LOGFILE", TDMK_Code.filter_str(item, item_val));
                        DataGridView DGV_tem = new DataGridView();
                        DGV_tem.DataSource = data_table.Columns[6];
                        if (data_table.Rows.Count > 0)
                        {
                            if (data_table.Rows.Count < number_val)
                            {
                                for (int j = 0; j < data_table.Rows.Count; j++)
                                {
                                    ws.Cells[row + j, i + 6].Value = (data_table.Rows[j]["Data"]);
                                }
                            }
                            else
                            {
                                for (int j = 0; j < number_val; j++)
                                {
                                    ws.Cells[row + j, i + 6].Value = (data_table.Rows[j]["Data"]);
                                }
                            }
                        }
                    }

                }
                Export_StackUp_Image(myVar.sqlcon_OK2SHIP, ws, ItemCode, LotNo);
            }
            return export_wrkbook;

        }
        public AutoCompleteStringCollection Get_zone_infor_to_input_value(myExcel.Worksheet ws)
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
        public AutoCompleteStringCollection Get_zone_infor_to_input_image(myExcel.Worksheet ws)
        {

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
        public void Export_StackUp_Image(SqlConnection sqlcon, myExcel.Worksheet ws, string Itemcode, string Lotno)
        {
            AutoCompleteStringCollection list = Get_zone_infor_to_input_image(ws);

            int no_BVH = 0;
            foreach (string x in list)
            {
                string region = x.Split('+')[0];
                if (!x.Contains("H_X"))
                {
                    region = "Zone_" + region;
                }
                else
                {
                    region = "Zone_" + region.Split('_')[0];
                }
                string row_val = x.Split('+')[1];
                int number_image = 5;
                string[] item = { "ItemCode", "LotNo", "Region" };
                string[] item_val = { Itemcode, Lotno, region };

                System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP_Period2, "STACKUP_IMAGE", TDMK_Code.filter_str(item, item_val));
                if (region != "Zone_BVH")
                {
                    if (data_image.Rows.Count > 0)
                    {
                        Export_Image_(ws, number_image, data_image, row_val);
                    }

                }
                else
                {
                    Export_Image_BVH(ws, data_image, row_val, no_BVH);
                    no_BVH++;
                }

            }

        }
        public void Export_Image_BVH(myExcel.Worksheet ws, System.Data.DataTable data_image, string row_val, int no_BVH)
        {
            int number_image = no_BVH * 5 + 5;
            if (data_image.Rows.Count < number_image)
            {
                number_image = data_image.Rows.Count - no_BVH * 5;

                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                string file_dic = Path.Combine(file_folder, "1.jpg");

                myExcel.Range curr_rgn = ws.Range["F" + row_val];
                for (int m = no_BVH * 5; m < number_image; m++)
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
                for (int m = no_BVH * 5; m < number_image; m++)
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
        //public void InsertPicture_Name(myExcel.Worksheet tar_wrksht, myExcel.Range tar_range, string picFile, int margin)
        //{
        //    float left;
        //    float top;
        //    float width;
        //    float height;
        //    string pic_name = Path.GetFileName(picFile);

        //    if (tar_range.MergeCells)
        //    {
        //        myExcel.Range refer_range = tar_range.MergeArea;
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
        //    // sel_picture.Placement = myExcel.XlPlacement.xlFreeFloating;
        //    sel_picture.Name = pic_name;

        //}

        public void InsertPicture_Name(myExcel.Worksheet tar_wrksht, myExcel.Range tar_range, string picFile, int margin)
        {
            tar_wrksht.Activate();
            string pic_name = Path.GetFileName(picFile);
            myExcel.Range sel_rgn = tar_range;
            myExcel.Range mer_sel_rgn = sel_rgn.MergeArea;
            float h = 0;
            float w = 0;
            myExcel.Range start_rgn = mer_sel_rgn.Cells[1, 1];

            int r = start_rgn.Row;
            for (int i = 0; i < mer_sel_rgn.Rows.Count; i++)
            {

                myExcel.Range temp1 = tar_wrksht.Rows[r + i];
                h = h + (float)temp1.Height;
            }
            int c_inx = start_rgn.Column;
            for (int i = 0; i < mer_sel_rgn.Columns.Count; i++)
            {
                myExcel.Range temp1 = tar_wrksht.Columns[c_inx + i];
                w += (float)temp1.Width;
            }

            myExcel.Shape sel_picture = tar_wrksht.Shapes.AddPicture(picFile, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, start_rgn.Left + margin, start_rgn.Top + margin, w - 2 * margin, h - 2 * margin);

            sel_picture.Placement = myExcel.XlPlacement.xlMoveAndSize;
            sel_picture.Name = pic_name;
        }
        public void Export_Image_(myExcel.Worksheet ws, int number_image, System.Data.DataTable data_image, string row_val)
        {
            if (data_image.Rows.Count < number_image)
            {
                number_image = data_image.Rows.Count;
                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                if (!Directory.Exists(file_folder))
                {
                    Directory.CreateDirectory(file_folder);
                }
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
                if (!Directory.Exists(file_folder))
                {
                    Directory.CreateDirectory(file_folder);
                }
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






        /***********************  Export BVH-PTH Function **********************************************/
        public void export_BVH_PTH(myExcel.Workbook export_wrkbook, string ItemCode, string LotNo)
        {

            myExcel.Workbook curr_wrkbook = export_wrkbook;
            myExcel.Worksheet ws = null;
            if (curr_wrkbook != null)
            {
                string _process = "BVH & PTH".Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
                foreach (myExcel.Worksheet tg_sht in curr_wrkbook.Worksheets)
                {
                    string cur_sht_name = tg_sht.Name.ToUpper().Replace("-", "").Replace(" ", "");
                    if (cur_sht_name == _process)
                    {
                        ws = tg_sht;
                        break;
                    }
                }


                List<string> tar_table = new List<string>() { "BVH_WITH_BONDING_SHEET", "BVH_WITHOUT_BONDING_SHEET", "PLATED_THROUGH_HOLE" };
                myExcel.Worksheet curr_wrksheet = ws;

                List<DataTable> tar_dt = new List<DataTable>();
                for (int i = 0; i < tar_table.Count; i++)
                {
                    tar_dt.Add(TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP_Period2, tar_table[i], TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo })));
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

                        DataTable dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP_Period2, t, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                        DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP_Period2, "SPEC_" + t, TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { ItemCode }));
                        bool chk = true;
                        if (!Check_spec_before_export(dt, dt_spec, ItemCode))
                        {
                            chk = false;
                            if (MessageBox.Show("Data of " + t + " is out of spec. Do you want to export to checksheet?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                            {
                                chk = true;
                            }
                        }
                        if (chk)
                        {

                            if (dt_spec.Rows.Count > 0)
                            {
                                string list_val = dt_spec.Rows[0]["Remark"].ToString();
                                switch (t)
                                {
                                    case "BVH_WITH_BONDING_SHEET":
                                        DB_to_Excel(dt, curr_wrksheet, "BVH_WITH_BONDING_SHEET", 8, list_val);
                                        break;
                                    case "BVH_WITHOUT_BONDING_SHEET":
                                        DB_to_Excel(dt, curr_wrksheet, "BVH_WITHOUT_BONDING_SHEET", 8, list_val);
                                        break;
                                    case "PLATED_THROUGH_HOLE":
                                        DB_to_Excel(dt, curr_wrksheet, "PLATED_THROUGH_HOLE", 7, list_val);
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }


                    }

                }

            }

        }
        public Boolean Check_spec_before_export(DataTable tbl_data, DataTable dt_spec, string ItemCode)
        {
            bool chk = true;
            string[] item = { "ItemCode" };
            string[] item_val = { ItemCode };

            if (dt_spec.Rows.Count > 0)
            {
                for (int i = 0; i < tbl_data.Rows.Count; i++)
                {
                    for (int k = 4; k < 12; k++)
                    {
                        string USL = dt_spec.Rows[0][k - 1].ToString().Replace("/", string.Empty);
                        string LSL = dt_spec.Rows[1][k - 1].ToString().Replace("/", string.Empty);
                        if (tbl_data.Rows[i][k] != null)
                        {
                            if (!USL.Contains("NA"))
                            {
                                if (Double.TryParse(tbl_data.Rows[i][k].ToString(), out Double data))
                                {
                                    if (data > Double.Parse(USL))
                                    {
                                        chk = false;
                                        break;
                                    }
                                }

                            }
                            if (!LSL.Contains("NA"))
                            {

                                if (Double.TryParse(tbl_data.Rows[i][k].ToString(), out Double data))
                                {
                                    if (data < Double.Parse(LSL))
                                    {
                                        chk = false;
                                        break;
                                    }
                                }

                            }
                        }

                    }
                    if (!chk)
                        break;
                }
            }
            return chk;



        }
        public void DB_to_Excel(DataTable src_DGV, myExcel.Worksheet curr_wrksheet, string process, int offset, string list_val)
        {
            string image_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = Path.Combine(image_folder, "1.jpg");


            int data_col = 8;

            int r = int.Parse(list_val.Split('+')[0]);
            int c = int.Parse(list_val.Split('+')[1]);
            int count_sample = int.Parse(list_val.Split('+')[2]);
            myExcel.Range curr_rgn = curr_wrksheet.Cells[r + 1, c + 2];
            myExcel.Range img_rgn = curr_rgn.Offset[0, offset];
            int qty = Math.Min(src_DGV.Rows.Count, count_sample);


            if (process == "PLATED_THROUGH_HOLE")
            {
                data_col = 7;
                curr_rgn = curr_wrksheet.Cells[r + 1, c + 1];
                img_rgn = curr_rgn.Offset[0, offset];
            }


            for (int m = 0; m < qty; m++)
            {
                int offs = int.Parse(src_DGV.Rows[m]["PCS_No"].ToString());

                for (int j = 0; j < data_col; j++)
                {
                    if (myCode.checkDBNull(src_DGV.Rows[m][j + 4]) != "")
                    {
                        curr_rgn.Offset[offs - 1, j].Value = src_DGV.Rows[m][j + 4].ToString();
                    }
                    else
                    {
                        curr_rgn.Offset[offs - 1, j].Value = "NA";
                    }
                }
            }

            for (int count = 0; count < qty; count++)
            {
                byte[] img_data = (byte[])src_DGV.Rows[count]["Image_data"];
                using (MemoryStream ms = new MemoryStream(img_data))
                {
                    Image image = Image.FromStream(ms);
                    image.Save(file_dic);
                    if (img_rgn.Value != null)
                    {
                        img_rgn.Delete(image);
                    }
                    InsertPicture_Name_BVHPTH(curr_wrksheet, img_rgn.MergeArea, file_dic, 5, process + "_" + count.ToString());
                }
                img_rgn = img_rgn.Offset[1, 0];

            }
        }
        public void InsertPicture_Name_BVHPTH(myExcel.Worksheet tar_wrksht, myExcel.Range tar_range, string picFile, int margin, string pic_name)
        {
            tar_wrksht.Activate();
            myExcel.Shape sel_picture = tar_wrksht.Shapes.AddPicture(picFile, MsoTriState.msoFalse, MsoTriState.msoTrue, tar_range.Left + margin, tar_range.Top + margin, tar_range.Width - 2 * margin, tar_range.Height - 2 * margin);
            sel_picture.LockAspectRatio = MsoTriState.msoTrue;
            
            sel_picture.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
            sel_picture.Placement = myExcel.XlPlacement.xlMoveAndSize;
            // sel_picture.Placement = myExcel.XlPlacement.xlFreeFloating;
            sel_picture.Name = pic_name;


        }







        /***********************  Export Impedance Function **********************************************/
        public void export_Impedance(myExcel.Workbook export_wrkbook, string ItemCode, string LotNo)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo });
            List<string> tar_table_name = new List<string>() { "IMPEDANCE_VAL", "TRACEWIDTH_VAL", "TRACEWIDTH_IMAGE", "IMPEDANCE_GRAPH" };

            List<DataTable> tar_dt = new List<DataTable>();
            for (int i = 0; i < tar_table_name.Count; i++)
            {
                tar_dt.Add(TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP_Period2, tar_table_name[i], filter_str));
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

                myExcel.Worksheet ws = null;
                myExcel.Workbook curr_wrkbook = export_wrkbook;
                string _process = "Impedance".Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
                if (curr_wrkbook != null)
                {
                    foreach (myExcel.Worksheet tg_sht in curr_wrkbook.Worksheets)
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
                        myExcel.Worksheet curr_wrksheet = ws;
                        curr_wrksheet.Cells.Interior.Color = myExcel.XlRgbColor.rgbWhite;
                        for (int i = 0; i < tar_table_name.Count; i++)
                        {
                            DataTable sel_dt = tar_dt[i];
                            switch (tar_table_name[i])
                            {
                                case "IMPEDANCE_VAL":
                                    Export_Impedance_Val(sel_dt, curr_wrksheet, ItemCode);
                                    break;
                                case "TRACEWIDTH_VAL":
                                    Export_VHX_data_2(sel_dt, curr_wrksheet, ItemCode);
                                    break;
                                case "TRACEWIDTH_IMAGE":
                                    Export_Image_VHX(sel_dt, curr_wrksheet, ItemCode);
                                    break;
                                case "IMPEDANCE_GRAPH":
                                    Export_Graph(curr_wrksheet, ItemCode, LotNo);
                                    break;
                            }
                        }

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
        public void Export_Impedance_Val(DataTable src_dt, myExcel.Worksheet curr_wrksheet, string ItemCode)
        {
            List<DataTable> src_region_tbl = new List<DataTable>();
            Get_ListTable(-1, src_dt, new string[] { "ItemCode", "LotNo", "Region" }, ref src_region_tbl, "Data");

            foreach (var in_tbl in src_region_tbl)
            {
                if (in_tbl.Rows.Count > 0)
                {
                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { ItemCode });
                    DataTable dt_format = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP_Period2, "IMPEDANCE_SPEC", filter_str);

                    string region = in_tbl.Rows[0]["Region"].ToString().Trim().Replace(" ", "");

                    //AutoCompleteStringCollection lst_rgn = Get_impedance_infor_to_input_value(curr_wrksheet);
                    int cou_rgn = 1;
                    myExcel.Range data_rgn;
                    for (int t = 0; t < dt_format.Rows.Count; t++)
                    {
                        string rgn = dt_format.Rows[t]["Range"].ToString().Split('-')[0];
                        string count_sample = dt_format.Rows[t]["Range"].ToString().Split('-')[1];
                        if (region == cou_rgn.ToString())
                        {
                            data_rgn = curr_wrksheet.Range[rgn];
                            int r_inx = 0;
                            int min_sample = new int[] { int.Parse(count_sample), in_tbl.Rows.Count }.Min();
                            foreach (DataRow dr in in_tbl.Rows)
                            {
                                if (r_inx < min_sample)
                                {
                                    data_rgn.Offset[r_inx, 0].Value = dr["Data"].ToString();
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
        public void Export_VHX_data_2(DataTable src_dt, myExcel.Worksheet curr_wrksheet, string ItemCode)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { ItemCode });
            DataTable dt_format = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP_Period2, "TRACEWIDTH_SPEC", filter_str);


            int idx_imp = 1;
            for (int t = 0; t < dt_format.Rows.Count; t++)
            {
                string item_rgn = dt_format.Rows[t]["Range"].ToString();
                int r = int.Parse(item_rgn.Split(':')[0]);
                int c = int.Parse(item_rgn.Split(':')[1]);
                int count_sample = int.Parse(item_rgn.Split(':')[2]);

                DataTable s_dt = src_dt.Clone();
                foreach (DataRow dr in src_dt.Rows)
                {
                    if (dr["Data_For"].ToString() == "Tracewidth_" + idx_imp.ToString())
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

                myExcel.Range rgn = curr_wrksheet.Cells[r, c];
                for (int k = 0; 2 * k < min_pcs; k++)
                {
                    myExcel.Range data_rgn = rgn.Offset[2 * k, 0];
                    for (int j = 0; j < 2; j++)
                    {
                        data_rgn.Offset[j, 0].Value = src_sample[k].Rows[j]["Data"].ToString();
                    }
                }
                idx_imp++;

            } 
        }
        public void Export_Image_VHX(DataTable src_dt, myExcel.Worksheet curr_wrksheet, string ItemCode)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { ItemCode });
            DataTable dt_format = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP_Period2, "TRACEWIDTH_SPEC", filter_str);

            // AutoCompleteStringCollection lst_rgn = Get_image_infor_to_input_value(curr_wrksheet);

            int idx_imp = 1;
            int no_graph = 0;

            for (int t = 0; t < dt_format.Rows.Count; t++)
            {
                string item_rgn = dt_format.Rows[t]["Range"].ToString();
                int r = int.Parse(item_rgn.Split(':')[0]) + 1;
                int c = int.Parse(item_rgn.Split(':')[1]) + 2;

                DataTable s_dt = src_dt.Clone();
                foreach (DataRow dr in src_dt.Rows)
                {
                    if (dr["Data_For"].ToString().Contains("Tracewidth_" + idx_imp.ToString()))
                    {
                        DataRow row = s_dt.NewRow();
                        for (int i = 0; i < src_dt.Columns.Count; i++)
                        {
                            row[i] = dr[i];
                        }
                        s_dt.Rows.Add(row);
                    }
                }
                myExcel.Range rgn_img_data = curr_wrksheet.Cells[r, c];

                Export_DatatableImage_Excel_2(s_dt, "Image_Tracewidth", curr_wrksheet, rgn_img_data, false);




                idx_imp++;
            }



        }
        public void Export_DatatableImage_Excel_2(DataTable dt, string Image_Col_name, myExcel.Worksheet tar_wrksht, myExcel.Range sel_rgn, bool row_offset)
        {
            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = Path.Combine(file_folder, DateTime.Now.ToString("ddMMMyy_HHmmss") + ".jpg");
            int r_inx = 0;
            // myExcel.Range sel_rgn = tar_wrksht.Range[tar_rgn];
            foreach (DataRow dr in dt.Rows)
            {
                byte[] img_byte = (byte[])(dr[Image_Col_name]);
                using (MemoryStream ms = new MemoryStream(img_byte))
                {
                    Image temp = Image.FromStream(ms);
                    temp.Save(file_dic);
                }
                InsertPicture_Name(tar_wrksht, sel_rgn, file_dic, 5);
                int col_offset = sel_rgn.Columns.Count;
                int row_off = sel_rgn.Rows.Count;
                if (row_offset)
                {
                    sel_rgn = sel_rgn.Offset[row_off, 0];
                }
                else
                {
                    sel_rgn = sel_rgn.Offset[0, col_offset];
                }
                r_inx++;
            }
            try
            {
                System.IO.File.Delete(file_dic);
            }
            catch
            {

            }

        }
        public void Export_Graph(myExcel.Worksheet curr_wrksheet, string ItemCode, string LotNo)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { ItemCode });
            DataTable dt_format = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP_Period2, "TRACEWIDTH_SPEC", filter_str);
            DataTable dt_format_imp = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP_Period2, "IMPEDANCE_SPEC", filter_str);
            List<DataTable> src_region_tbl = new List<DataTable>();

            string[] region = { "Coupon", "Patern" };
            int no_trace = 0;
            foreach (string _regn in region)
            {
                DataTable graph_tbl = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP_Period2, "IMPEDANCE_GRAPH", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Zone" }, new string[] { ItemCode, LotNo, _regn }));
                string[] sel_val = graph_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();
                foreach (string item in sel_val)
                {
                    DataTable s_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP_Period2, "IMPEDANCE_GRAPH", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Zone", "Region" }, new string[] { ItemCode, LotNo, _regn, item }));
                    if (no_trace < dt_format.Rows.Count && no_trace < dt_format_imp.Rows.Count)
                    {
                        string item_rgn = dt_format.Rows[no_trace]["Range"].ToString();
                        int r = int.Parse(item_rgn.Split(':')[0]) + 1;
                        int c = int.Parse(item_rgn.Split(':')[1]) + 2;

                        myExcel.Range rgn_img_data = curr_wrksheet.Cells[r, c];
                        myExcel.Range rgn_graph = rgn_img_data.Offset[rgn_img_data.Rows.Count + 2, 0];
                        Export_DatatableImage_Excel_2(s_dt, "Image_Graph", curr_wrksheet, rgn_graph, false);
                    }
                    int count = int.Parse(s_dt.Rows[0]["Region"].ToString());
                    no_trace += count;
                }

            }
        }







        /***********************  Export SolderMask Function **********************************************/
        public void export_SolderMask(myExcel.Workbook export_wrkbook, string ItemCode, string LotNo)
        {

            myExcel.Workbook curr_wrkbook = export_wrkbook;
            myExcel.Worksheet ws = null;
            if (curr_wrkbook != null)
            {
                string _process = "SolderMask".Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
                foreach (myExcel.Worksheet tg_sht in curr_wrkbook.Worksheets)
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
                    Export_Solder_Image_1(sqlcon_OK2SHIP_Period2, ws, ItemCode, LotNo);
                    Export_LPI_data(ws, ItemCode, LotNo);

                }

            }

        }
        public void Export_LPI_data(myExcel.Worksheet ws, string ItemCode, string LotNo)
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
                    ws.Range[tar_rgn[j]].Offset[i, 0].Value = LPI_data.Rows[i][j].ToString();
                }
            }
        }
        public void Export_Solder_Image_1(SqlConnection sqlcon, myExcel.Worksheet ws, string Itemcode, string Lotno)
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
                        //System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "SOLDERMASK_IMAGE", TDMK_Code.filter_str(item, item_val));
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





        /***********************  Export ThermalStress Function **********************************************/
        public void export_ThermalStress(myExcel.Workbook export_wrkbook, string ItemCode, string LotNo)
        {
            myExcel.Workbook curr_wrkbook = export_wrkbook;
            myExcel.Worksheet ws = null;
            if (curr_wrkbook != null)
            {
                string _process = "CQRA - Thermal stress".Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
                foreach (myExcel.Worksheet tg_sht in curr_wrkbook.Worksheets)
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
                    Export_Thermal_Stress_Image(sqlcon_OK2SHIP_Period2, ws, ItemCode, LotNo);
                }

            }
        }
        public void Export_Thermal_Stress_Image(SqlConnection sqlcon, myExcel.Worksheet ws, string Itemcode, string Lotno)
        {
            AutoCompleteStringCollection list = Get_infor_to_input_thermalstress(ws);

            // myExcel.Worksheet ws = wb.Sheets[1];
            int no_BVH = 0;
            foreach (string x in list)
            {
                string region = x.Split('+')[0];
                region = region.Split(' ')[0];
                string row_val = x.Split('+')[1];
                int number_image = 5;
                string[] item = { "ItemCode", "LotNo", "Region" };
                string[] item_val = { Itemcode, Lotno, region };
                System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, "THERMAL_STRESS_IMAGE", TDMK_Code.filter_str(item, item_val));
                if (region.Contains("BVH"))
                {
                    Export_Image_BVH_Thermalstress(ws, data_image, row_val, no_BVH);
                    no_BVH++;
                }
                else
                {
                    if (data_image.Rows.Count > 0)
                    {
                        Export_Image_ThermalStress(ws, number_image, data_image, row_val);
                    }
                }
            }


        }
        public AutoCompleteStringCollection Get_infor_to_input_thermalstress(myExcel.Worksheet ws)
        {
            // myExcel.Worksheet ws = wb.Sheets[1];
            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 10; i < 50; i++)
            {
                if (ws.Cells[i, 1].Value != null)
                {

                    if (ws.Cells[i, 1].Value.Contains("PTH"))
                    {
                        list.Add(ws.Cells[i, 1].Value.ToString() + "+" + (i).ToString());
                    }
                    if (ws.Cells[i, 1].Value.Contains("BVH"))
                    {
                        list.Add(ws.Cells[i, 1].Value.ToString() + "+" + (i).ToString());
                    }
                }

            }
            return list;
        }
        public void Export_Image_BVH_Thermalstress(myExcel.Worksheet ws, System.Data.DataTable data_image, string row_val, int no_BVH)
        {
            int number_image = no_BVH * 5 + 5;
            if (data_image.Rows.Count < number_image)
            {
                number_image = data_image.Rows.Count - no_BVH * 5;

                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                string file_dic = Path.Combine(file_folder, "1.jpg");

                myExcel.Range curr_rgn = ws.Range["B" + row_val];
                for (int m = no_BVH * 5; m < number_image; m++)
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
        public void Export_Image_ThermalStress(myExcel.Worksheet ws, int number_image, System.Data.DataTable data_image, string row_val)
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



        /*********************** Export ACF Function **********************************************/
        public void export_ACF(myExcel.Workbook export_wrkbook, string ItemCode, string LotNo)
        {
            myExcel.Workbook curr_wrkbook = export_wrkbook;
            myExcel.Worksheet ws = null;
            if (curr_wrkbook != null)
            {
                
                string _process = "ACF".Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
                foreach (myExcel.Worksheet tg_sht in curr_wrkbook.Worksheets)
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

                    Export_ACF_Image(sqlcon_OK2SHIP_Period2, ws, ItemCode, LotNo);
                    Export_ACFFlatness(sqlcon_OK2SHIP_Period2, ws, ItemCode, LotNo);
                    Export_ACF_Roughness(ItemCode, LotNo, ws);
                 
                }

            }
        }
        public void Get_ListTable_Image(int col_inx, DataTable myDt, string[] src_arr, ref List<DataTable> src_lst_tbl, string tar_item)
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
                                Get_ListTable_Image(col_inx + 1, curTbl, src_arr, ref src_lst_tbl, tar_item);
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
        public void Export_ACF_Image(SqlConnection sqlcon, myExcel.Worksheet ws, string Itemcode, string Lotno)
        {
            //myExcel.Worksheet ws = wb.Sheets["ACF"];
            List<string> search_item = new List<string>() { "Photo before cleaning", "Photo after cleaning", "Photo after OQC testing", "Before tape test", "After tape test", "Tape test picture" };//
            myExcel.Range sel_rgn = ws.Range["A16"];

            foreach (var item in search_item)
            {
                int num_insert = 5;
                string tar_table = "";
                string tar_col_data = "Image_Data";
                if (item.Contains("cleaning") || (item.Contains("OQC")))
                {
                    tar_table = "ACF_CLEANING_IMAGE";
                    num_insert = 10;
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
                Get_ListTable_Image(-1, data_tbl, new string[] { "ItemCode", "LotNo", "Region" }, ref src_lst_tbl, tar_col_data);
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
                        int qty = Math.Min(dt.Rows.Count, num_insert);
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
        public void Export_ACFFlatness(SqlConnection sqlcon, myExcel.Worksheet ws, string ItemCode, string LotNo)
        {
            if (ItemCode != "" && LotNo != "")
            {
                DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, "ACF_FLATNESS", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                if (src_dt.Rows.Count > 0)
                {
                    myExcel.Range curr_rgn = ws.Range["A1"];
                    for (int i = 0; i < 100; i++)
                    {
                        if (curr_rgn.Offset[i, 0].Value != null)
                        {
                            if (curr_rgn.Offset[i, 0].Value.ToString().Replace(" ", "").ToUpper() == "ACF Flatness Measurement".Replace(" ", "").ToUpper())
                            {
                                curr_rgn = curr_rgn.Offset[i, 0];
                                break;
                            }
                        }
                    }
                    for (int i = 0; i < 5; i++)
                    {
                        for (int j = 0; j < 10; j++)
                        {
                            curr_rgn.Offset[4 + i, 2 + j].Value = src_dt.Rows[i][4 + j].ToString();
                        }
                    }

                }

            }
        }
        public void Export_ACF_Roughness(string itemcode, string lotno, myExcel.Worksheet ws)
        {
            
            DataTable dt1 = new DataTable();
            dt1 = TDMK_Code.Datatable_Filter(myVar.sqlcon_IPQC, "Roughness", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemcode, lotno }));
            if (dt1.Rows.Count > 0)
            {
                List<List<string>> lst_rghness = lst_roughness(dt1);

                myExcel.Range curr_rgn_1 = ws.Range["A1"];

                List<int> vitri_sa = new List<int> { };
                List<int> vitri_sq = new List<int> { };
                List<int> vitri_sdr = new List<int> { };
                string[] arr_info = ACF_pad_location(ws).Split('+');
                string vitri_ACF_pad = arr_info[0];

                int row = int.Parse(arr_info[0]);
                int r_count = int.Parse(arr_info[arr_info.Length - 1]);
                curr_rgn_1 = curr_rgn_1.Offset[row, 0];

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
                                curr_rgn_1.Offset[j, i].Value = lst_rghness[k][m];
                                k += 3;
                            }
                            k = 1;
                            foreach (int j in vitri_sq)
                            {
                                curr_rgn_1.Offset[j, i].Value = lst_rghness[k][m];
                                k += 3;
                            }
                            k = 2;
                            foreach (int j in vitri_sdr)
                            {
                                curr_rgn_1.Offset[j, i].Value = lst_rghness[k][m];
                                k += 3;
                            }
                            m++;
                        }

                        curr_rgn_1 = curr_rgn_1.Offset[r_count, 0];

                    }
                }
             
            }
        }
        public List<List<string>> lst_roughness(DataTable dt)
        {
            List<List<string>> lst_roughness = new List<List<string>>() { };
            string[] arr_name = { "L1_Roughness_Sa", "L1_Roughness_Sq", "L1_Roughness_Sdr", "L2_Roughness_Sa", "L2_Roughness_Sq", "L2_Roughness_Sdr", "L3_Roughness_Sa", "L3_Roughness_Sq", "L3_Roughness_Sdr" };
            foreach (var item in arr_name)
            {
                lst_roughness.Add(dt.AsEnumerable().Select(x => x.Field<string>(item)).ToList());
            }

            return lst_roughness;
        }
        public string ACF_pad_location(myExcel.Worksheet ws)
        {
            // List<int> lst_info = new List<int> { };
            string get_info = "";
            myExcel.Range curr_rgn_1 = ws.Range["A1"];

            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 30; i < 100; i++)
            {
                if (curr_rgn_1.Offset[i, 0].Value != null)
                {
                    if (curr_rgn_1.Offset[i, 0].Value.ToString().Replace(" ", "").ToUpper() == "ACF pad location".Replace(" ", "").ToUpper())
                    {
                        int k = 1;
                        get_info = i.ToString();
                        while (curr_rgn_1.Offset[i + k, 1].Value.ToString() != "Roughness")
                        {
                            if (curr_rgn_1.Offset[i + k, 1].Value.ToString().Contains("Sa"))
                            {

                                get_info += "+ sa";
                            }
                            if (curr_rgn_1.Offset[i + k, 1].Value.ToString().Contains("Sq"))
                            {

                                get_info += "+ sq";
                            }
                            if (curr_rgn_1.Offset[i + k, 1].Value.ToString().Contains("Sdr"))
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




        /*********************** Export Moisture Function **********************************************/
        public void export_Moisture(myExcel.Workbook export_wrkbook, string ItemCode, string LotNo)
        {
            if (ItemCode != "" && LotNo != "")
            {
                myExcel.Workbook curr_wrkbook = export_wrkbook;
                myExcel.Worksheet ws = null;
                DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon_OK2SHIP_Period2, "CQRA_MOISTURE_ABSORPTION", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));

                if(src_dt.Rows.Count > 0)
                {
                    if (curr_wrkbook != null)
                    {
                        string _process = "CQRA - Moisture absorption".Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
                        foreach (myExcel.Worksheet tg_sht in curr_wrkbook.Worksheets)
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
                            myExcel.Range curr_rgn = ws.Range["B17"];
                            for (int i = 0; i < 5; i++)
                            {
                                curr_rgn.Offset[i, 0].Value = src_dt.Rows[i]["Weight_before"].ToString();
                                curr_rgn.Offset[i, 1].Value = src_dt.Rows[i]["Weight_after"].ToString();
                            }

                        }

                    }
                }

               
            
            }
        }






        /*********************** Export Bhast Funcion **********************************************/
        public void Export_Bhast_Logfile(myExcel.Workbook export_wrkbook, string itemcode, string lotno)
        {
            string[] item = { "ItemCode", "LotNo" };
            string[] item_val = { itemcode, lotno };
            string[] list_ch = { "Ch01", "Ch02", "Ch03", "Ch04", "Ch05", "Ch06", "Ch07", "Ch08" };
            int start_col = 1;
            int start_row = 1;
           
            myExcel.Worksheet ws = null;
            myExcel.Workbook curr_wrkbook = export_wrkbook;
                if (curr_wrkbook != null)
                {
                    string _process = "CQRA - bHast".Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper();
                    foreach (myExcel.Worksheet tg_sht in curr_wrkbook.Worksheets)
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
                            try
                            {
                                AutoCompleteStringCollection list = TDMK_Code.Load_Item_Filter_str(sqlcon_OK2SHIP_Period2, "CQRA_BHAST_LOGFILE", chanel, TDMK_Code.filter_str(item, item_val));

                                string number = chanel.Replace("Ch0", "");
                                int col_offset = Int32.Parse(number);

                                for (int y = 1; y <= list.Count; y++)
                                {
                                    ws.Cells[start_row + y, start_col + col_offset - 1].Value = list[y - 1];
                                }
                            }
                            catch
                            {

                            }

                        }
                         
                    }
                   
                } 
        }
    }
}
