using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SEEV_DLL;
using TDMK_SQL;

using myExcel = Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Reflection;

namespace VHX
{
    public class Impedance_data
    {
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SEI_Lib myCode = new SEI_Lib();

        public myExcel.Workbook new_workbook(string f_path)
        {
            string format_path = Path.Combine(Application.StartupPath, "Software Type2(copy).xlsx");
            // string targetPath = Application.StartupPath + @"\Export";
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook curr_wrkbook = app.Workbooks.Open(format_path, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            myExcel.Worksheet curr_wrksheet = curr_wrkbook.Sheets["Impedance"];

            var newbook = app.Workbooks.Add(1);
            curr_wrksheet.Copy(newbook.Sheets[1]);

            newbook.SaveAs(f_path);
            newbook.Close();

            // curr_wrkbook.Close();
            return curr_wrkbook;
        }
        public void Export_data(string ItemCode, string LotNo, string Process, int imp, FileInfo[] File_Info)
        {
            if (ItemCode != "" && LotNo != "")
            {
                //string app_path = @"E:\SEEV_GĐ2\Test Areas\Format\Impedance";
                //string app_path = Path.Combine(Application.StartupPath, "Test Areas", "Format", Process);
                //string f_name_main = ItemCode + ".xlsx";
                //try
                //{
                //    myExcel.Workbook curr_wrkbook_main = TDMK_Code.open_excel_file(app_path, f_name_main, "");
                //    string f_name = Path.Combine(Application.StartupPath, "Test Areas", "Export", Process, ItemCode + "_" + LotNo + ".xlsx");

                //    //int imp = comboBox2.SelectedIndex;
                myExcel.Workbook curr_wrkbook;
                string format_path = Path.Combine(Application.StartupPath, "Software Type2(copy).xlsx");
                //string export_path = Path.Combine(Application.StartupPath, "Export");
                string targetPath = Path.Combine(Application.StartupPath, "Impedance", "Export");
                string f_name = ItemCode + "_" + LotNo + "_" + ".xlsx";

                if (!System.IO.File.Exists(Path.Combine(targetPath, f_name)))
                {
                    curr_wrkbook = new_workbook(Path.Combine(targetPath, f_name));
                }
                else
                {
                    curr_wrkbook = TDMK_Code.open_excel_file(targetPath, f_name, "");
                }


                if (imp == 0)
                    {
                        foreach (FileInfo fileinfo in File_Info)
                        {
                            myExcel.Workbook curr_wrkbook_1 = TDMK_Code.open_excel_file(fileinfo.FullName, "", "");
                            Insert_Impedance(curr_wrkbook, curr_wrkbook_1, imp + 1);
                            imp++;
                            curr_wrkbook_1.Close();
                        }
                       // curr_wrkbook_main.SaveAs(f_name);
                    }
                    else
                    {
                        FileInfo fileinfo = File_Info[imp - 1];
                        myExcel.Workbook curr_wrkbook_1 = TDMK_Code.open_excel_file(fileinfo.FullName, "", "");
                        Insert_Impedance(curr_wrkbook, curr_wrkbook_1, imp);
                        curr_wrkbook_1.Close();
                        //curr_wrkbook_main.SaveAs(f_name);
                    }
                    MessageBox.Show("Hoàn thành", "Thông báo");
                
            //    }
            //    catch
            //    {
            //        MessageBox.Show("Không tìm thấy file Format", "Thông báo");

            //    }

            }
            else
            {
                MessageBox.Show("Nhập đầy đủ thông tin", "Thông báo");
            }

        }
        public void Export_data_impedance(string ItemCode, string LotNo, int imp, SqlConnection sqlcon, string table_name, myExcel.Worksheet curr_wrksheet)
        {
            if (ItemCode != "" && LotNo != "")
            {
                myExcel.Range curr_rgn = curr_wrksheet.Range["E40"];

                if (imp == 0)
                {
                    for (int i = 1; i < 5; i++)
                    {
                        myExcel.Range curr_rgn_main = curr_rgn.Offset[(i - 1) * 53, 0];
                        DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, table_name, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo, i.ToString() }));
                        if (dt.Rows.Count > 0)
                        {
                            for (int m = 0; m < dt.Rows.Count; m++)
                            {
                                int ofs = int.Parse(dt.Rows[m]["Pcs_no"].ToString());
                                curr_rgn_main.Offset[ofs, 0].Value = dt.Rows[m]["Data"].ToString();

                            }
                        }
                    }
                }
                else
                {
                    myExcel.Range curr_rgn_main = curr_rgn.Offset[(imp - 1) * 53, 0];
                    DataTable dt = TDMK_Code.Datatable_Filter(sqlcon, table_name, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo, imp.ToString() }));
                    if (dt.Rows.Count > 0)
                    {
                        for (int m = 0; m < dt.Rows.Count; m++)
                        {
                            int ofs = int.Parse(dt.Rows[m]["Pcs_no"].ToString());
                            curr_rgn_main.Offset[ofs, 0].Value = dt.Rows[m]["Data"].ToString();

                        }

                    }
                }
                MessageBox.Show("Hoàn thành", "Thông báo");
                //    }
                //    catch
                //    {
                //        MessageBox.Show("Không tìm thấy file Format", "Thông báo");

                //    }

            }
            else
            {
                MessageBox.Show("Nhập đầy đủ thông tin", "Thông báo");
            }

        }
        public void Export_data_impedance2(DataTable in_tbl, myExcel.Worksheet curr_wrksheet, int index)
        {
            if(in_tbl.Rows.Count>0)
            {
                string region = in_tbl.Rows[0]["Region"].ToString().Trim().Replace(" ","");
                myExcel.Range data_rgn;
                switch (region)
                {
                    case "Impedance-1":
                        data_rgn = curr_wrksheet.Range["E41"];
                        break;
                    case "Impedance-2":
                        data_rgn = curr_wrksheet.Range["E94"];
                        break;
                    case "Impedance-3":
                        data_rgn = curr_wrksheet.Range["E147"];
                        break;
                    case "Impedance-4":
                        data_rgn = curr_wrksheet.Range["E200"];
                        break;
                    default:
                        data_rgn = curr_wrksheet.Range["E41"].Offset[53*index,0];
                        break;
                }
                int r_inx = 0;
                foreach(DataRow dr in in_tbl.Rows)
                {
                    data_rgn.Offset[r_inx,0].Value = dr["Data"].ToString();
                    r_inx++;
                }
            }

        }

        public void loadtoDB_imp(myExcel.Workbook curr_wrkbook_1, int imp, string ItemCode, string LotNo, SqlConnection sqlcon)
        {
            try
            {

                myExcel.Worksheet curr_wrksheet_1 = curr_wrkbook_1.Sheets["Result"];

                myExcel.Range curr_rgn_1 = curr_wrksheet_1.Range["A7"];
                myExcel.Range rgn_pcs = curr_wrksheet_1.Range["A1"];

                Dictionary<string, int> myDict = new Dictionary<string, int>();
                for (int i = 0; i < 32; i++)
                {
                    myDict.Add(curr_rgn_1.Offset[0, i].Value.ToString(), int.Parse(rgn_pcs.Offset[0, i].Value.ToString()));
                }
                int j = 0;
                string[] data = new string[32];
                int indx = 0;
                foreach (KeyValuePair<string, int> author in myDict.OrderBy(key => key.Value))
                {
                    data[indx] = author.Key;
                    indx++;
                }

                for (int i = 0; i < 32; i++)
                {
                    int ID = TDMK_Code.SQL_MAX("IMPEDANCE_VAL", "ID", sqlcon) + 1;

                    string query = "insert into " + "IMPEDANCE_VAL" + " values (@ID, @ItemCode, @LotNo, @Pcs_No, @Region, @Data, @Remark)";
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = sqlcon;
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Parameters.AddWithValue("@ItemCode", ItemCode);
                        cmd.Parameters.AddWithValue("@LotNo", LotNo);
                        cmd.Parameters.AddWithValue("@Pcs_No", i + 1);
                        cmd.Parameters.AddWithValue("@Region", imp);
                        cmd.Parameters.AddWithValue("@Data", data[i]);
                        cmd.Parameters.AddWithValue("@Remark", "");
                        sqlcon.Open();
                        cmd.ExecuteNonQuery();
                        sqlcon.Close();

                    }
                }


                //   MessageBox.Show("Hoàn thành", "Thông báo");

            }
            catch
            {

            }
        }
        public void loadtoDGV_imp(myExcel.Workbook curr_wrkbook_1, int imp, DataTable dt, string ItemCode, string LotNo)
        {
            myExcel.Worksheet curr_wrksheet_1 = curr_wrkbook_1.Sheets["Result"];
            myExcel.Range curr_rgn_1 = curr_wrksheet_1.Range["A7"];
            myExcel.Range rgn_pcs = curr_wrksheet_1.Range["A1"];
            Dictionary<string, int> myDict = new Dictionary<string, int>();
            for (int i = 0; i < 32; i++)
            {
                myDict.Add(curr_rgn_1.Offset[0, i].Value.ToString(), int.Parse(rgn_pcs.Offset[0, i].Value.ToString()));
            }
            int j = 0;
            string[] data = new string[32];
            int indx = 0;
            foreach (KeyValuePair<string, int> author in myDict.OrderBy(key => key.Value))
            {
                data[indx] = author.Key;
                indx++;
            }

            for (int i = 0; i < 32; i++)
            {
                DataRow dr = dt.NewRow();
                dr[0] = dt.Rows.Count + 1;
                dr[1] = ItemCode;
                dr[2] = LotNo;
                dr[3] = i + 1;
                dr[4] = imp;
                dr[5] = data[i];
                dr[6] = "";

                dt.Rows.Add(dr);


            }


            //   MessageBox.Show("Hoàn thành", "Thông báo");


        }
        public void Insert_Impedance(myExcel.Workbook curr_wrkbook_main, myExcel.Workbook curr_wrkbook_1, int imp)
        {
            myExcel.Worksheet curr_wrksheet_1 = curr_wrkbook_1.Sheets["Result"];
            myExcel.Range curr_rgn_1 = curr_wrksheet_1.Range["A7"];
            myExcel.Range rgn_pcs = curr_wrksheet_1.Range["A1"];

            // myExcel.Workbook curr_wrkbook_main = TDMK_Code.open_excel_file(app_path, f_name_main, "");
            myExcel.Worksheet curr_wrksheet_main = curr_wrkbook_main.Sheets["Impedance"];
            myExcel.Range curr_rgn = curr_wrksheet_main.Range["E41"];
            myExcel.Range curr_rgn_main = curr_rgn.Offset[(imp - 1) * 53, 0];

            Dictionary<string, int> myDict = new Dictionary<string, int>();
            for (int i = 0; i < 32; i++)
            {
                myDict.Add(curr_rgn_1.Offset[0, i].Value.ToString(), int.Parse(rgn_pcs.Offset[0, i].Value.ToString()));
            }
            int j = 0;
            foreach (KeyValuePair<string, int> author in myDict.OrderBy(key => key.Value))
            {
                // Console.WriteLine("Key: {0}, Value: {1}", author.Key, author.Value);

                curr_rgn_main.Offset[j, 0].Value = author.Key;
                j++;
            }

        }

        public SqlConnection initial_data(string DB_name, bool sa_en)
        {
            SqlConnection _sqlcon_OK2SHIP;
            string app_path = Application.StartupPath;
            string config_file = Path.Combine(app_path, "Config", "config.txt");
            string[] my_config = myCode.read_config_arr(config_file);
            string server_name = "";
            string server_acc = "";
            string server_pass = "";
            string Data_Location = "";
            string format_folder = "";
            string log_folder = "";
            foreach (string c in my_config)
            {
                if (c.Contains("Server"))
                {
                    server_name = c.Split(':')[1].Trim();
                }
                if (c.Contains("Account"))
                {
                    server_acc = c.Split(':')[1].Trim();
                }
                if (c.Contains("Password"))
                {
                    server_pass = c.Split(':')[1].Trim();
                }
                if (c.Contains("Data_Location"))
                {
                    string[] temp = c.Split(':');
                    string tg = "";
                    if (temp.Length > 2)
                    {

                        for (int t = 1; t < temp.Length; t++)
                        {
                            tg = tg + temp[t] + ":";
                        }
                    }
                    Data_Location = tg.TrimEnd(':').Trim();
                }
                if (c.Contains("Format_Folder"))
                {
                    format_folder = c.Split('#')[1].Trim();
                }
                if (c.Contains("Log_folder"))
                {
                    log_folder = c.Split('#')[1].Trim();
                }
            }
            if (sa_en)
            {
                string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, DB_name, server_acc, server_pass).ConnectionString;
                _sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);
            }
            else
            {
                string _strcon = data_connection2(server_name, DB_name).ConnectionString;
                _sqlcon_OK2SHIP = new SqlConnection(_strcon);
            }

            return _sqlcon_OK2SHIP;
        }
        public SqlConnectionStringBuilder data_connection2(string _servername, string _databasename)
        {
            return new SqlConnectionStringBuilder
            {
                DataSource = _servername,
                InitialCatalog = _databasename,
                IntegratedSecurity = true
            };
        }

        public void Load_DB_to_DGV(string ItemCode, string LotNo, DataGridView DGV1, SqlConnection sqlcon, string table_name, int imp)
        {
            if (imp == 0)
            {
                DGV1.DataSource = TDMK_Code.Datatable_Filter(sqlcon, table_name, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                //TDMK_Code.fill_dataset_DGV(table_name, @"Select * from " + table_name + " WHERE ItemCode = '" + ItemCode + "'" + " AND LotNo = '" + LotNo + "'", DGV1, sqlcon);

            }
            else
            {
                // TDMK_Code.fill_dataset_DGV(table_name, @"Select * from " + table_name + " WHERE ItemCode = '" + ItemCode + "'" + " AND LotNo = '" + LotNo + "'" + " AND Region = '" + imp + "'", DGV1, sqlcon);
                DGV1.DataSource = TDMK_Code.Datatable_Filter(sqlcon, table_name, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo, imp.ToString() }));

            }
            if (DGV1.Rows.Count > 0)
                MessageBox.Show("Hoàn thành", "Thông báo");
            else
                MessageBox.Show("Không có dữ liệu", "Thông báo");
        }
        public void Load_to_DB(string ItemCode, string LotNo, string table_name, SqlConnection sqlcon, int imp, FileInfo[] File_Inf)
        {
            DataTable dt = new DataTable();
            dt = TDMK_Code.Datatable_Filter(sqlcon, table_name, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
            if (dt.Rows.Count > 0)
            {
                var result = MessageBox.Show("Ghi đè dữ liệu", "Thông báo", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    TDMK_Code.Delelte_FilteredItem_arr(table_name, sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                }

            }
            // int imp = comboBox2.SelectedIndex;
            if (imp == 0)
            {
                int region = 1;
                foreach (FileInfo fileinfo in File_Inf)
                {
                    myExcel.Workbook curr_wrkbook_1 = TDMK_Code.open_excel_file(fileinfo.FullName, "", "");
                    loadtoDB_imp(curr_wrkbook_1, imp + 1, ItemCode, LotNo, sqlcon);
                    region++;
                    curr_wrkbook_1.Close();
                    imp++;
                }
            }
            else
            {
                FileInfo fileinfo = File_Inf[imp - 1];
                myExcel.Workbook curr_wrkbook_1 = TDMK_Code.open_excel_file(fileinfo.FullName, "", "");

                loadtoDB_imp(curr_wrkbook_1, imp, ItemCode, LotNo, sqlcon);
                curr_wrkbook_1.Close();
            }


            MessageBox.Show("Hoàn thành", "Thông báo");
        }
        public void Load_to_DB_impedance(string ItemCode, string LotNo, string table_name, SqlConnection sqlcon, int imp, FileInfo[] File_Inf, DataGridView dgv)
        {
            DataTable dt = new DataTable();
            dt = TDMK_Code.Datatable_Filter(sqlcon, table_name, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
            if (dt.Rows.Count > 0)
            {
                var result = MessageBox.Show("Ghi đè dữ liệu", "Thông báo", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    TDMK_Code.Delelte_FilteredItem_arr(table_name, sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                    save_to_db(table_name, sqlcon, dgv);
                }
            }
            else
            {
                save_to_db(table_name, sqlcon, dgv);
            } 
        }
        public void save_to_db(string table_name, SqlConnection sqlcon, DataGridView DGV)
        {


            string[] arr_item = { "ID", "ItemCode", "LotNo", "Pcs_No", "Region", "Data", "Remark" };
            string[] arr_item_val = new string[7];
            for (int i = 0; i < DGV.Rows.Count; i++)
            {
                arr_item_val[0] = (TDMK_Code.SQL_MAX(table_name, "ID", sqlcon) + 1).ToString();
                for (int j = 1; j < DGV.Columns.Count; j++)
                {
                    arr_item_val[j] = DGV.Rows[i].Cells[j].Value.ToString();
                }
                TDMK_Code.insert_val_arr(table_name, sqlcon, arr_item, arr_item_val);
            }


            MessageBox.Show("Hoàn thành", "Thông báo");
        }

        public void Load_DGV(string Itemcode, string LotNo, DataGridView DGV1, int imp, FileInfo[] File_Info)
        {
                DataTable dt = new DataTable();
                dt.Columns.Add("ID");
                dt.Columns.Add("ItemCode");
                dt.Columns.Add("LotNo");
                dt.Columns.Add("Pcs_No");
                dt.Columns.Add("Region");
                dt.Columns.Add("Data");
                dt.Columns.Add("Remark");
                if (imp == 0)
                {
                    int i = 1;
                    foreach (FileInfo fileinfo in File_Info)
                    {
                        myExcel.Workbook curr_wrkbook_1 = TDMK_Code.open_excel_file(fileinfo.FullName, "", "");
                        loadtoDGV_imp(curr_wrkbook_1, i, dt, Itemcode, LotNo);
                        i++;
                        curr_wrkbook_1.Close();
                    }
                }
                else
                {
                    FileInfo fileinfo = File_Info[imp - 1];
                    myExcel.Workbook curr_wrkbook_1 = TDMK_Code.open_excel_file(fileinfo.FullName, "", "");
                    loadtoDGV_imp(curr_wrkbook_1, imp, dt, Itemcode, LotNo);
                    curr_wrkbook_1.Close();
                }
                DGV1.DataSource = dt;
                MessageBox.Show("Hoàn thành", "Thông báo");
        }
    }
    

}
