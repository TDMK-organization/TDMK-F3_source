using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using myExcel = Microsoft.Office.Interop.Excel;
using TDMK_SQL;
using System.IO;
using System.Data.SqlClient;
using TDMK_SEEV_DLL;
using IniLibs;
using System.Diagnostics;
using ZedGraph;
using System.Net.NetworkInformation;

namespace Create_Table_DB
{
    public partial class Form1 : Form
    {
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SqlConnection sqlcon_SMT;
        SEI_Lib SEEV_Code = new SEI_Lib();
        public static SqlConnection sel_sqlcon;
        public string server_name;
        public string server_acc;
        public string server_pass;
        public string DB_name;
        public string app_path;
        public string Data_Location;
        public string format_folder;
        public string log_folder;
        //string connstr_OK2SHIP;
        public IniFile TDMK_init = new IniFile("Config.ini");
        public Form1()
        {
            InitializeComponent();
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            if (txtFile.Text != "")
            {
                myExcel.Workbook myWrkbk = TDMK_Code.open_excel_file(txtFile.Text, "", "");
                foreach (myExcel.Worksheet sht in myWrkbk.Sheets)
                {
                    myExcel.Range tar_rgn = sht.Range["B1"];
                    List<string> tbl_lst = new List<string>();
                    List<string> col_lst = new List<string>();
                    int col_inx = 0;
                    int row_inx = 0;
                    while (SEEV_Code.checkDBNull(tar_rgn.Offset[0, col_inx].Value) != "")
                    {
                        if (tar_rgn.Offset[0, col_inx].Value != "Comment")
                        {
                            string temp_tbl = tar_rgn.Offset[0, col_inx].Value;
                            if (sht.Name.ToUpper().Contains("LOGFILE"))
                            {
                                temp_tbl += "_LOGFILE";
                            }
                            tbl_lst.Add(temp_tbl.Replace(" ", "").Replace("-", "_").ToUpper().Trim());
                        }
                        col_inx++;
                    }
                    while (SEEV_Code.checkDBNull(tar_rgn.Offset[2 + row_inx, -1].Value) != "")
                    {
                        col_lst.Add(tar_rgn.Offset[2 + row_inx, -1].Value);
                        row_inx++;
                    }
                    foreach (string tbl in tbl_lst)
                    {
                        TDMK_Code.Create_tables_SQL(tbl, col_lst.ToArray(), sqlcon_SMT.ConnectionString);
                    }

                }
                MessageBox.Show("Finished!");
            }
            else
            {
                MessageBox.Show("Chọn file cấu trúc dữ liệu....", "Thông báo");
            }

        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            string app_path = Application.StartupPath;
            OpenFileDialog f_open = new OpenFileDialog();
            f_open.InitialDirectory = app_path;
            if (f_open.ShowDialog() == DialogResult.OK)
            {
                if (f_open.FileName != "")
                {
                    txtFile.Text = f_open.FileName;
                }
            }
        }
        public string[] read_config_arr(string src_config_file)
        {
            string[] _result = new string[10];
            StreamReader reader;
            int inx = 0;
            reader = new StreamReader(src_config_file);
            while ((!reader.EndOfStream) && (inx < 10))
            {
                string temp = reader.ReadLine();
                if (temp != "")
                {
                    _result[inx] = temp;
                    inx++;
                }
            }
            Array.Resize<string>(ref _result, inx);
            reader.Close();
            reader.Dispose();
            return _result;
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            sqlcon_SMT = initial_data("OK2SHIP_Period2", true);
            //sqlcon_SMT = initial_data("SEI_FAI", true);
            cbType.SelectedIndex = 0;
        }
        public string add_tbl_str(string src_tbl_name, string[] col_name)
        {
            //string _result = "";
            string query = "";
            query = "IF OBJECT_ID('dbo." + src_tbl_name + "', 'U') IS NULL ";
            query += "BEGIN ";
            query += "CREATE TABLE [dbo].[" + src_tbl_name + "](";
            query += "[ID] INT NOT NULL,";
            foreach (string c in col_name)
            {
                if (c.Contains("Image"))
                {
                    query += "[" + c + "] VARBINARY(MAX) NULL,";
                }
                else
                {
                    query += "[" + c + "] NVARCHAR(MAX) NULL,";
                }
            }
            query.TrimEnd(',');
            query += ")";
            query += " END";
            return query;
        }
        public void Create_tables_SQL(string tbl_name, string[] col_name, string constring)
        {
            string create_table = add_tbl_str(tbl_name, col_name);
            SqlConnection con = new SqlConnection(constring);
            if (con.State != ConnectionState.Open)
            {
                con.Open();
            }
            SqlCommand cmd = new SqlCommand(create_table, con);
            cmd.ExecuteNonQuery();
            con.Close();
            MessageBox.Show("Success!");
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (txtFile.Text != "")
            {
                myExcel.Workbook myWrkbk = TDMK_Code.open_excel_file(txtFile.Text, "", "");
                foreach (myExcel.Worksheet sht in myWrkbk.Sheets)
                {
                    myExcel.Range tar_rgn = sht.Range["B1"];
                    List<string> tbl_lst = new List<string>();
                    int col_inx = 0;
                    while (SEEV_Code.checkDBNull(tar_rgn.Offset[0, col_inx].Value) != "")
                    {
                        if (tar_rgn.Offset[0, col_inx].Value != "Comment")
                        {
                            string temp_tbl = tar_rgn.Offset[0, col_inx].Value;
                            tbl_lst.Add(temp_tbl.Replace(" ", "").Replace("-", "_").ToUpper().Trim());
                        }
                        col_inx++;
                    }
                    foreach (string tbl in tbl_lst)
                    {
                        delete_tables_SQL(tbl, sqlcon_SMT);
                    }
                }
                MessageBox.Show("Hoàn thành!");
                lstTable.DataSource = GetAllTables(sqlcon_SMT);
            }
            else
            {
                if (lstTable.Items.Count > 0)
                {
                    string tbl = lstTable.SelectedItem.ToString();
                    delete_tables_SQL(tbl, sqlcon_SMT);
                    MessageBox.Show("Hoàn thành!");
                    lstTable.DataSource = GetAllTables(sqlcon_SMT);
                }
            }
        }
        public void delete_tables_SQL(string tbl_name, SqlConnection con)
        {
            try
            {
                string create_table = "DROP TABLE " + tbl_name;
                //SqlConnection con = new SqlConnection(constring);
                if (con.State != ConnectionState.Open)
                {
                    con.Open();
                }
                SqlCommand cmd = new SqlCommand(create_table, con);
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Delete table " + tbl_name + " : success!");
            }
            catch
            {
                MessageBox.Show("Table " + tbl_name + " : is not existed!");
            }

        }
        public string Table_Add_Column_str(string src_tbl_name, string[] col_name)
        {
            //string _result = "";
            string query = "";
            query += "ALTER TABLE [dbo].[" + src_tbl_name + "] ADD ";
            foreach (string c in col_name)
            {
                if (c.Contains("Image"))
                {
                    query += "[" + c + "] VARBINARY(MAX) NULL,";
                }
                else
                {
                    query += "[" + c + "] NVARCHAR(MAX) NULL,";
                }
            }
            query = query.TrimEnd(',');
            query += ";";
            return query;
        }

        public SqlConnection initial_data(string DB_name, bool sa_en)
        {
            SqlConnection _sqlcon_OK2SHIP;
            if (!File.Exists("Config.ini"))
            {
                TDMK_init.Write("Server", "10.212.1.243", "SMT_Config");
                TDMK_init.Write("Account", "sa", "SMT_Config");
                TDMK_init.Write("Password", "seev@123;", "SMT_Config");
                TDMK_init.Write("Report_Location", @"\\10.212.1.200\rgo\OK2SHIP automation system", "SMT_Config");
                TDMK_init.Write("Data_Location", @"\\10.212.1.200\rgo\OK2SHIP automation system\TDMK_Data\TestAreas\OK2SHIP_Measurements\OMM", "SMT_Config");
                TDMK_init.Write("Format_Folder", @"\\10.212.1.200\rgo\OK2SHIP automation system\0.OK2SHIP report format", "SMT_Config");
                TDMK_init.Write("Log_folder", @"\\10.212.1.200\rgo\OK2SHIP automation system\TDMK_Data\TestAreas\OK2SHIP_Measurements", "SMT_Config");
            }
            server_name = TDMK_init.Read("Server", "SMT_Config");
            server_acc = TDMK_init.Read("Account", "SMT_Config");
            server_pass = TDMK_init.Read("Password", "SMT_Config");
            Data_Location = TDMK_init.Read("Data_Location", "SMT_Config");
            format_folder = TDMK_init.Read("Format_Folder", "SMT_Config");
            log_folder = TDMK_init.Read("Log_folder", "SMT_Config");
            if (sa_en)
            {
                string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, DB_name, server_acc, server_pass).ConnectionString;
                _sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);
            }
            else
            {
                string _strcon = TDMK_Code.data_connection2(server_name, DB_name).ConnectionString;
                _sqlcon_OK2SHIP = new SqlConnection(_strcon);
            }
            return _sqlcon_OK2SHIP;
        }

        private void btnGetAllTable_Click(object sender, EventArgs e)
        {
            lstTable.DataSource = GetAllTables(sqlcon_SMT);
        }
        public List<string> GetAllTables(SqlConnection connection)
        {
            List<string> result = new List<string>();
            if (connection.State != ConnectionState.Open)
            {
                connection.Open();
            }
            SqlCommand cmd = new SqlCommand("SELECT name FROM sys.Tables", connection);
            System.Data.SqlClient.SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
                result.Add(reader["name"].ToString());
            connection.Close();
            return result;
        }

        private void btnDeleteAll_Click(object sender, EventArgs e)
        {
            if (lstTable.Items.Count > 0)
            {
                foreach (string tbl in lstTable.Items)
                {
                    delete_tables_SQL(tbl, sqlcon_SMT);
                }
            }
        }

        private void btnDeleteData_Click(object sender, EventArgs e)
        {
            if (lstTable.SelectedIndex != -1)
            {
                string tar_tbl = lstTable.SelectedItem.ToString();
                TDMK_Code.Delelte_FilteredItem_arr(tar_tbl, sqlcon_SMT, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
                MessageBox.Show("Xóa dữ liệu ItemCode / LotNo : " + txtItemCode.Text + " / " + txtLotNo.Text + " thành công!", "Thông báo");
            }
            else
            {
                MessageBox.Show("Chọn bảng dữ liệu cần xóa", "Thông báo");
            }
        }

        private void btnDelAll_Click(object sender, EventArgs e)
        {
            if (lstTable.Items.Count > 0)
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
                foreach (var t in lstTable.Items)
                {
                    TDMK_Code.Delelte_FilteredItem_arr(t.ToString(), sqlcon_SMT, filter_str);
                }
                MessageBox.Show("Hoàn thành", "Thông báo");
            }
        }

        private void btnDelete_Col_Click(object sender, EventArgs e)
        {
            if (lstTable.SelectedIndex != -1)
            {
                string col = txtColName.Text;
                List<string> tar_col_lst = getColumnsName(lstTable.SelectedItem.ToString(), sqlcon_SMT);
                if (tar_col_lst.IndexOf(col) != -1)
                {
                    SQL_Command_Exe(delete_table_column(lstTable.SelectedItem.ToString(), txtColName.Text), sqlcon_SMT);
                    lstTable_SelectedIndexChanged(null, null);
                }
                else
                {
                    MessageBox.Show("Cột dữ liệu " + txtColName.Text + " không tồn tại", "Thông báo");
                }
            }
        }

        private void btnCreate_Col_Click(object sender, EventArgs e)
        {
            if (lstTable.SelectedIndex != -1)
            {
                string col = txtColName.Text;
                string datatype = cbType.SelectedItem.ToString();
                if (col != "" && datatype != "")
                {
                    List<string> tar_col_lst = getColumnsName(lstTable.SelectedItem.ToString(), sqlcon_SMT);
                    if (tar_col_lst.IndexOf(col) == -1)
                    {
                        string col_add_str = Add_column_str(lstTable.SelectedItem.ToString(), col, datatype);
                        SQL_Command_Exe(col_add_str, sqlcon_SMT);
                        lstTable_SelectedIndexChanged(null, null);
                    }
                    else
                    {
                        MessageBox.Show("Cột dữ liệu " + txtColName.Text + " đã có, không thể tạo thêm", "Thông báo");
                    }
                }
                else
                {
                    MessageBox.Show("Nhập đủ thông tin cho Column Name / Column Type", "Thông báo");
                }
            }
        }
        public string Add_column_str(string src_tbl_name, string col_name, string col_type)
        {
            string query = "";
            query += "ALTER TABLE [dbo].[" + src_tbl_name + "] ADD ";
            query += "[" + col_name + "] " + col_type + " NULL,";
            query = query.TrimEnd(',');
            query += ";";
            return query;
        }
        public string delete_table_column(string src_tbl, string col_name)
        {
            string query = "";
            query += "ALTER TABLE [dbo].[" + src_tbl + "] DROP COLUMN ";
            query += "[" + col_name + "];";
            return query;
        }
        public void SQL_Command_Exe(string command, SqlConnection sqlcon)
        {
            try
            {
                if (sqlcon.State != ConnectionState.Open)
                {
                    sqlcon.Open();
                }
                SqlCommand cmd = new SqlCommand(command, sqlcon);
                cmd.ExecuteNonQuery();
                sqlcon.Close();
                MessageBox.Show("Success!");
            }
            catch
            {
                MessageBox.Show("Xảy ra lỗi khi thực hiện lệnh. Kiểm tra lại các thông tin", "Thông báo");
            }

        }
        public List<string> getColumnsName(string src_tbl, SqlConnection sqlcon)
        {
            List<string> listacolumnas = new List<string>();
            using (SqlCommand command = sqlcon.CreateCommand())
            {
                command.CommandText = "select c.name from sys.columns c inner join sys.tables t on t.object_id = c.object_id and t.name = '" + src_tbl + "' and t.type = 'U'";
                if (sqlcon.State != ConnectionState.Open)
                {
                    sqlcon.Open();
                }
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        listacolumnas.Add(reader.GetString(0));
                    }
                }
                sqlcon.Close();
            }
            return listacolumnas;
        }

        private void lstTable_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstTable.SelectedIndex > -1)
            {
                string tbl_name = lstTable.SelectedItem.ToString();
                lstTable_Struct.DataSource = getColumnsName(tbl_name, sqlcon_SMT);
            }
        }

        private void lstTable_Struct_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstTable_Struct.SelectedIndex != -1)
            {
                txtColName.Text = lstTable_Struct.SelectedItem.ToString();
            }
        }
        void DB_SelectedChanged(object sender, EventArgs e)
        {
            foreach (RadioButton rb in GB_Database.Controls)
            {
                if (rb.Checked)
                {
                    string tar_db = rb.Text;
                    sqlcon_SMT = initial_data(tar_db, true);
                    break;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (lstTable.Items.Count > 0)
            {
                CloneProcess("422435");
            }
            else
            {
                MessageBox.Show("Không có bảng trong database !");
            }
        }
        private Dictionary<string, DataTable> DB_List = new Dictionary<string, DataTable>();
        private void CloneProcess(string itemCode)
        {
            string lotno = "00168";
            DB_List = new Dictionary<string, DataTable>();
            foreach (var item in lstTable.Items)
            {
                DataTable newTable = getTable(item.ToString(), itemCode, lotno);
                //DataTable newTable = new DataTable();
                DB_List.Add(item.ToString(), newTable);
            }
            WriteDictionaryToFile(DB_List, "C:\\Users\\servi\\Desktop\\New Text Document.txt");
            MessageBox.Show("Thanh cong");
        }
        /// <summary>
        /// Lấy tất cả dữ liệu trong bảng
        /// </summary>
        /// <param name="nameTable"></param>
        /// <returns></returns>
        private DataTable getTable(string nameTable, string itemCode, string lotno)
        {
            try
            {
                string fillter = $"ItemCode = '{itemCode}' and LotNo Like '%{lotno}%'";
                return TDMK_Code.Datatable_Filter(sqlcon_SMT, nameTable, fillter);
            }
            catch
            {
                //MessageBox.Show($"Table {nameTable} không có lotno");
                return new DataTable();
            }
        }
        public static Dictionary<string, DataTable> ReadDictionaryFromFile(string filePath)
        {
            Dictionary<string, DataTable> dataTables = new Dictionary<string, DataTable>();

            try
            {
                using (StreamReader reader = new StreamReader(filePath, Encoding.UTF8))
                {
                    string line;
                    DataTable currentTable = null;
                    string tableName = null;

                    while ((line = reader.ReadLine()) != null)
                    {
                        if (line.StartsWith("Table:"))
                        {
                            tableName = line.Substring(7).Trim(); // Lấy tên bảng
                            currentTable = new DataTable(tableName);
                            dataTables.Add(tableName, currentTable);
                        }
                        else if (currentTable != null)
                        {
                            if (line.Contains('\t')) // Kiểm tra xem có phải là dòng dữ liệu hay không
                            {
                                string[] values = line.Split('\t');

                                // Nếu là dòng tiêu đề cột
                                if (currentTable.Columns.Count == 0)
                                {
                                    foreach (string value in values)
                                    {
                                        currentTable.Columns.Add(value.Trim()); // Thêm cột
                                    }
                                }
                                else // Nếu là dòng dữ liệu
                                {
                                    DataRow row = currentTable.NewRow();
                                    for (int i = 0; i < values.Length && i < currentTable.Columns.Count; i++)
                                    {
                                        row[i] = values[i].Trim(); // Gán giá trị cho cột
                                    }
                                    currentTable.Rows.Add(row);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi đọc file: {ex.Message}");
            }

            return dataTables;
        }
        public static void WriteDictionaryToFile(Dictionary<string, DataTable> dataTables, string filePath)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(filePath, false, Encoding.UTF8))
                {
                    foreach (var kvp in dataTables)
                    {
                        string tableName = kvp.Key;
                        DataTable table = kvp.Value;

                        writer.WriteLine($"Table: {tableName}"); // Ghi tên bảng

                        // Ghi tiêu đề cột
                        foreach (DataColumn column in table.Columns)
                        {
                            writer.Write($"{column.ColumnName}\t"); // Sử dụng tab để phân tách
                        }
                        writer.WriteLine();

                        // Ghi dữ liệu
                        foreach (DataRow row in table.Rows)
                        {
                            foreach (object item in row.ItemArray)
                            {
                                writer.Write($"{item}\t"); // Sử dụng tab để phân tách
                            }
                            writer.WriteLine();
                        }

                        writer.WriteLine(); // Dòng trống giữa các bảng
                    }
                }

                Console.WriteLine($"Dữ liệu đã được ghi vào file: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi ghi file: {ex.Message}");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DB_List = ReadDictionaryFromFile("C:\\Users\\servi\\Desktop\\New Text Document.txt");
            foreach (var i in DB_List)
            {
                if (i.Value.Rows.Count > 0)
                {
                    var ItemCode = i.Value.Rows[0]["ItemCode"].ToString();
                    var LotNo = i.Value.Rows[0]["LotNo"].ToString();
                    var table = getTable(i.Key, ItemCode, LotNo);
                    if (table.Rows.Count > 0)
                    {
                        TDMK_Code.Delelte_FilteredItem_arr(i.Key, sqlcon_SMT, $"LotNo = '{LotNo}' and ItemCode = '{ItemCode}'");
                    }
                    InsertDataTable(i.Value, sqlcon_SMT, i.Key);

                }
            }
            MessageBox.Show("Insert thành côgn");
        }

        public void InsertDataTable(DataTable list_dt, SqlConnection connection, string tableName)
        {
            foreach (DataRow item in list_dt.Rows)
            {
                try
                {

                    string[] val = item.ItemArray.Select(z => z.ToString()).ToArray(); ;
                    var fillter = list_dt.Columns.Cast<DataColumn>().Select(col => col.ColumnName).ToArray();
                    TDMK_Code.insert_val_arr2(tableName, connection, fillter, val);
                }
                catch
                {

                }
            }
        }
    }


}
