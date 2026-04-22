using CsvHelper;
using CsvHelper.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using SkiaSharp;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace TDMK_Helper
{
    public class TDMK_Helpers
    {
        public static string imageFolder { get; set; }
        public static List<string> listSQLConnection_TDMK { get; set; }

        #region Initialization and Configuration
        public TDMK_Helpers()
        {
            listSQLConnection_TDMK = new List<string>()
            {
                initial_SQL_Command("OK2SHIP_Period2", true),
                initial_SQL_Command("OK2SHIP_Q2_2025", true), // bỏ tạm cho F1 chạy
                initial_SQL_Command("OK2SHIP_Q3_2025", true),
                //initial_SQL_Command("OK2SHIP_Q1_2026", true),
            };
            string app_path = Application.StartupPath;
            string config_file = Path.Combine(app_path, "Config", "config.txt");
            string[] my_config = read_config_arr(config_file);
            foreach (string c in my_config)
            {
                if (c.Contains("Saved_Image_Location"))
                {
                    imageFolder = c.Split('#')[1].Trim();
                }
            }
        }
        private static string[] read_config_arr(string src_config_file)
        {
            string[] array = new string[1000];
            int num = 0;
            StreamReader streamReader = new StreamReader(src_config_file);
            while (!streamReader.EndOfStream && num < 1000)
            {
                string text = streamReader.ReadLine();
                if (text != "")
                {
                    array[num] = text;
                    num++;
                }
            }

            Array.Resize(ref array, num);
            streamReader.Close();
            streamReader.Dispose();
            return array;
        }

        private static string initial_SQL_Command(string DB_name, bool sa_en)
        {
            string app_path = Application.StartupPath;
            string config_file = Path.Combine(app_path, "Config", "config.txt");
            string[] my_config = read_config_arr(config_file);
            string server_name = "", server_acc = "", server_pass = "", data_loc = "";

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
                    data_loc = c.Split('#')[1].Trim();
                }
            }
            string result = "";
            if (sa_en)
            {
                result = new SqlConnectionStringBuilder
                {
                    DataSource = server_name,
                    InitialCatalog = DB_name,
                    UserID = server_acc,
                    Password = server_pass,
                    ConnectTimeout = 0
                }.ToString();
            }
            else
            {
                result = new SqlConnectionStringBuilder
                {
                    DataSource = server_name,
                    InitialCatalog = DB_name,
                    IntegratedSecurity = true,
                    ConnectTimeout = 0
                }.ToString();
            }
            return result;
        }
        #endregion

        #region Database Connection Management
        public static SqlConnection GetSqlConnectionByIndex(List<string> listSQLConnection, int position = 0)
        {
            SqlConnection conn = new SqlConnection(listSQLConnection[position]);
            return conn;
        }

        public static SqlConnection GetSqlConnectionPeriod2()
        {
            SqlConnection conn = new SqlConnection(listSQLConnection_TDMK[0]);
            return conn;
        }

        public static SqlConnection GetSqlConnectionLastDB()
        {
            SqlConnection conn = new SqlConnection(listSQLConnection_TDMK[listSQLConnection_TDMK.Count - 1]);
            return conn;
        }
        public static SqlConnection GetSqlConnectionDeclaration()
        {
            SqlConnection conn = new SqlConnection(initial_SQL_Command("Declaration", true));
            return conn;
        }

        public static void CloseSQLConnection(SqlConnection conn)
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
        #endregion

        #region Database Operations
        public static async Task<DataTable> GetDataTableWithMultipleSQL(string filter_str, string tableName)
        {
            using (var cts = new CancellationTokenSource())
            {
                var tasks = new Task<DataTable>[listSQLConnection_TDMK.Count];
                DataTable schemaTable = null;

                for (int i = 0; i < listSQLConnection_TDMK.Count; i++)
                {
                    var connStr = listSQLConnection_TDMK[i];
                    var dbIndex = i + 1;
                    tasks[i] = ExecuteQueryAsync(connStr, filter_str, tableName, cts.Token, dbIndex);
                }

                var completedTasks = new List<Task<DataTable>>(tasks);

                while (completedTasks.Count > 0)
                {
                    try
                    {
                        var completedTask = await Task.WhenAny(completedTasks);
                        var result = await completedTask;

                        if (result != null)
                        {
                            if (schemaTable == null)
                            {
                                schemaTable = result.Clone();
                            }

                            if (result.Rows.Count > 0)
                            {
                                cts.Cancel();
                                return result;
                            }

                            completedTasks.Remove(completedTask);
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[ERROR] Database connection error: {ex.Message}");
                    }
                }

                return schemaTable ?? new DataTable();
            }
        }

        public static String ReadFileTxt(String filePath)
        {
            String content = String.Empty;
            try
            {
                using (StreamReader sr = new StreamReader(filePath))
                {
                    content = sr.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                TDMK_Message.MessageBoxTDMK_Error($"Error reading file: {ex.Message}");
            }
            return content;
        }

        private static async Task<DataTable> ExecuteQueryAsync(string connectionString, string filter, string tableName, CancellationToken cancellationToken, int dbIndex = 0)
        {
            try
            {
                using (var connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync(cancellationToken);

                    var cacheKey = GenerateCacheKey(tableName, filter);
                    if (Cache.Get(cacheKey) is DataTable cachedData && cachedData.Rows.Count > 0)
                    {
                        return cachedData;
                    }

                    var sql = BuildSqlQuery(tableName, filter);

                    using (var command = new SqlCommand(sql, connection))
                    {
                        command.CommandTimeout = 0;

                        using (var adapter = new SqlDataAdapter(command))
                        {
                            var dataTable = new DataTable();
                            cancellationToken.ThrowIfCancellationRequested();
                            adapter.Fill(dataTable);

                            //if (dataTable.Rows.Count > 0)
                            //{
                            //    Cache.Set(cacheKey, dataTable, 60 * 15);
                            //}

                            return dataTable;
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                return new DataTable();
            }
            catch (Exception ex)
            {
                TDMK_Message.MessageBoxTDMK_Error($"Database error: {ex.Message}");
                return new DataTable();
            }
        }

        public static DataTable GetDataTableWithCaChe(SqlConnection conn, string filter_str, string tableName, bool filterById = true)
        {
            var cacheKey = GenerateCacheKey(tableName, filter_str);

            //if (Cache.Get(cacheKey) is DataTable cachedDt && cachedDt.Rows.Count > 0)
            //{
            //    return cachedDt;
            //}

            var dt = GetDataTableDirect(conn, tableName, filter_str, filterById);

            //if (dt != null && dt.Rows.Count > 0)
            //{
            //    Cache.Set(cacheKey, dt, 60 * 15);
            //}

            return dt ?? new DataTable();
        }

        private static DataTable GetDataTableDirect(SqlConnection connection, string tableName, string filterCondition, bool filterById = true)
        {
            var sql = BuildSqlQuery(tableName, filterCondition, filterById);

            using (var command = new SqlCommand(sql, connection))
            {
                command.CommandTimeout = 0;
                using (var adapter = new SqlDataAdapter(command))
                {
                    var dataTable = new DataTable();
                    adapter.Fill(dataTable);
                    return dataTable;
                }
            }
        }

        private static string BuildSqlQuery(string tableName, string filterCondition, bool filterById = true)
        {
            var sql = new StringBuilder(256);
            sql.Append("SELECT * FROM [").Append(tableName).Append("]");

            if (!string.IsNullOrEmpty(filterCondition))
            {
                sql.Append(" WHERE ").Append(filterCondition);
            }
            if (filterById)
                sql.Append(" ORDER BY ID");
            return sql.ToString();
        }

        //just for check evidance
        public static async Task<DataTable> GetListEvidence(string filter_str)
        {
            string ConnectionString = initial_SQL_Command("Declaration", true);
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                string command = "SELECT ItemCode, LotNo FROM dbo.EVIDENCE_DATA";
                if (!string.IsNullOrEmpty(filter_str))
                {
                    command += " WHERE " + filter_str;
                }
                //command += " EXCEPT SELECT DISTINCT ItemCode, LotNo FROM dbo.HISTORY_EVIDENCE;";
                SqlCommand cmd = new SqlCommand(command, conn)
                {
                    CommandTimeout = 0,
                };

                await conn.OpenAsync();

                //var cacheKey = GenerateCacheKey("EVIDENCE", filter_str);
                //if (Cache.Get(cacheKey) is DataTable cachedData && cachedData.Rows.Count > 0)
                //{
                //    return cachedData;
                //}

                using (SqlDataReader reader = await cmd.ExecuteReaderAsync())
                {
                    DataTable dTable = null;
                    bool hasLoadedSchema = false;
                    var tempTable = new DataTable();
                    tempTable.Columns.AddRange(new DataColumn[]{
                        new DataColumn("ItemCode", typeof(string)),
                        new DataColumn("LotNo", typeof(string)),
                        new DataColumn("Type", typeof(string)),
                        //new DataColumn("Evidence_Image", typeof(byte[])),
                        new DataColumn("Result", typeof(string)),
                    } );

                    while (await reader.ReadAsync())
                    {
                        var row = tempTable.NewRow();
                        row["ItemCode"] = reader.GetString(0);
                        row["LotNo"] = reader.GetString(1);
                        row["Type"] = "";
                        //row["Evidence_Image"] = null;
                        row["Result"] = "FAIL";
                        tempTable.Rows.Add(row);
                    }

                    if (!hasLoadedSchema)
                    {
                        dTable = tempTable.Clone();
                        hasLoadedSchema = true;
                    }

                    if (tempTable.Rows.Count > 0)
                    {
                        // nếu không có điều kiện truy ván thì lấy 200 dòng cuối cùng trong bảng
                        if (string.IsNullOrEmpty(filter_str))
                        {
                            var reversed = tempTable.AsEnumerable()
                            .Reverse();
                            //.Take(200);

                            dTable = reversed.CopyToDataTable();
                        }
                        else
                        {
                            dTable = tempTable;
                        }
                    }

                    //if (dTable != null && dTable.Rows.Count > 0)
                    //{

                    //    Cache.Set(cacheKey, dTable, 60 * 8);
                    //}

                    if (dTable != null && dTable.Rows.Count > 0 && dTable.Columns.Contains("ID"))
                    {
                        DataView dataView = dTable.DefaultView;
                        dataView.Sort = "ID ASC";
                        dTable = dataView.ToTable();
                        //Cache.Set(cacheKey, dTable, 480);
                    }

                    return dTable ?? new DataTable();
                }
            }
        }

        public static async Task<DataTable> GetDataWithStoreProducere(string tableName, string filter_str, bool searchCacheData = false)
        {
            string ConnectionString = listSQLConnection_TDMK[0];
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                SqlCommand cmd = new SqlCommand("sp_GetCQRAData", conn)
                {
                    CommandTimeout = 0,
                    CommandType = CommandType.StoredProcedure
                };
                cmd.Parameters.AddWithValue("@TableName", tableName);
                cmd.Parameters.AddWithValue("@FilterStr", filter_str);

                await conn.OpenAsync();

                //var cacheKey = GenerateCacheKey(tableName, filter_str);
                //if (Cache.Get(cacheKey) is DataTable cachedData && cachedData.Rows.Count > 0 && searchCacheData)
                //{
                //    return cachedData;
                //}

                using (SqlDataReader reader = await cmd.ExecuteReaderAsync())
                {
                    DataTable dTable = null;
                    bool hasLoadedSchema = false;

                    do
                    {
                        if (reader.IsClosed) break;
                        var tempTable = new DataTable();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            tempTable.Columns.Add(reader.GetName(i), reader.GetFieldType(i));
                        }

                        while (await reader.ReadAsync())
                        {
                            var row = tempTable.NewRow();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                row[i] = reader.IsDBNull(i) ? DBNull.Value : reader.GetValue(i);
                            }
                            tempTable.Rows.Add(row);
                        }

                        if (!hasLoadedSchema)
                        {
                            dTable = tempTable.Clone();
                            hasLoadedSchema = true;
                        }

                        if (tempTable.Rows.Count > 0)
                        {
                            dTable = tempTable;
                            break;
                        }

                    } while (!reader.IsClosed && await reader.NextResultAsync());

                    if (dTable != null && dTable.Rows.Count > 0)
                    {
                        DataView dataView = dTable.DefaultView;
                        dataView.Sort = "ID ASC";
                        dTable = dataView.ToTable();
                        //if (searchCacheData)
                            //Cache.Set(cacheKey, dTable, 60 * 8);
                    }

                    return dTable ?? new DataTable();
                }
            }
        }
        public static void BatchBulkCopy(SqlConnection sqlcon_OK2SHIP, DataTable dataTable, string DestinationTbl)
        {
            using (SqlBulkCopy sbc = new SqlBulkCopy(sqlcon_OK2SHIP))
            {
                if (sqlcon_OK2SHIP.State != ConnectionState.Open)
                {
                    sqlcon_OK2SHIP.Open();
                }
                if (dataTable.Columns.Contains("ID")) dataTable.Columns.Remove("ID");
                if (dataTable.Columns.Contains("status")) dataTable.Columns.Remove("status");

                sbc.DestinationTableName = DestinationTbl;
                foreach (DataColumn dc in dataTable.Columns)
                {
                    sbc.ColumnMappings.Add(dc.ColumnName, dc.ColumnName);
                }
                sbc.WriteToServer(dataTable);
                if (sqlcon_OK2SHIP.State == ConnectionState.Open)
                {
                    sqlcon_OK2SHIP.Close();
                }
            }
        }

        public async static Task DeleteDataFromDatabase(string process_name, string filter_str)
        {
            DataTable dtable = ConvertTableJsonToTable(await GetDataWithStoreProducere(process_name, filter_str));

            if (dtable.Rows.Count > 0)
            {
                foreach (DataColumn col in dtable.Columns)
                {
                    if (col.ColumnName.ToLower().Contains("image"))
                    {
                        DeleteAllImageBeforeDeleteData(dtable, col.ColumnName);
                    }
                }
            }
            foreach (string command in listSQLConnection_TDMK)
            {
                using (SqlConnection conn = new SqlConnection(command))
                {
                    if (conn.State != ConnectionState.Open)
                    {
                        conn.Open();
                    }
                    Delelte_FilteredItem_arr(process_name, conn, filter_str);
                }
            }
            RemoveDataFromCache(filter_str, process_name);
        }

        public static void Delelte_FilteredItem_arr(string tbl_name, SqlConnection database_conn, string filter_string)
        {
            var sql = new StringBuilder(256);
            sql.Append("DELETE FROM [").Append(tbl_name).Append("]");

            if (!string.IsNullOrEmpty(filter_string))
            {
                sql.Append(" WHERE ").Append(filter_string);
            }

            using (SqlCommand sqlCommand = new SqlCommand(sql.ToString(), database_conn))
            {
                if (database_conn.State != ConnectionState.Open)
                {
                    database_conn.Open();
                }
                sqlCommand.ExecuteNonQuery();
                database_conn.Close();
            }
        }
        #endregion

        #region Cache Management
        private static string GenerateCacheKey(string tableName, string filter)
        {
            var input = $"{tableName}|{filter ?? ""}";
            using (var md5 = MD5.Create())
            {
                var hash = md5.ComputeHash(Encoding.UTF8.GetBytes(input));
                return Convert.ToBase64String(hash);
            }
        }

        public static void RemoveDataFromCache(string filter_str, string tableName)
        {
            var cacheKey = GenerateCacheKey(tableName, filter_str);
            if (Cache.Contains(cacheKey))
                Cache.Remove(cacheKey);
        }

        public static void ClearAllCacheWhenSaveData()
        {
            Cache.RemoveAll();
        }
        #endregion

        #region DataTable Conversions
        public static DataTable ConvertDataGridViewToDataTable(DataGridView dgv)
        {
            DataTable dt = new DataTable();

            foreach (DataGridViewColumn col in dgv.Columns)
            {
                if (!col.Visible) continue;

                Type columnType;
                if (col is DataGridViewImageColumn || (col is DataGridViewButtonColumn))
                {
                    columnType = typeof(byte[]);
                }
                else if (col is DataGridViewComboBoxColumn)
                {
                    columnType = typeof(string);
                }
                else
                {
                    columnType = col.ValueType ?? typeof(string);
                }

                dt.Columns.Add(col.Name, columnType);
            }

            foreach (DataGridViewRow dgvRow in dgv.Rows)
            {
                if (dgvRow.IsNewRow) continue;

                DataRow dtRow = dt.NewRow();

                foreach (DataGridViewColumn col in dgv.Columns)
                {
                    if (!col.Visible) continue;

                    object cellValue = dgvRow.Cells[col.Name].Value;

                    if (cellValue is System.Drawing.Image img)
                    {
                        dtRow[col.Name] = ConvertImageToByteArray(new Bitmap(img));
                    }
                    else if (cellValue is string strValue && strValue == "")
                    {
                        dtRow[col.Name] = DBNull.Value;
                    }
                    else if (cellValue != null && cellValue != DBNull.Value)
                    {
                        dtRow[col.Name] = cellValue;
                    }
                    else
                    {
                        dtRow[col.Name] = DBNull.Value;
                    }
                }

                dt.Rows.Add(dtRow);
            }

            return dt;
        }
        public static void DeleteAllImageBeforeDeleteData(DataTable table, string columnName)
        {
            try
            {
                List<string> imageFilePath = new List<string>();
                foreach (DataRow row in table.Rows)
                {
                    if (row[columnName].ToString().Contains(".jpg"))
                        imageFilePath.Add(row[columnName].ToString());
                }
                foreach (var img_path in imageFilePath)
                {
                    if (File.Exists(img_path))
                    {
                        File.Delete(img_path);
                    }
                }
            }
            catch (Exception ex)
            {
                TDMK_Message.MessageBoxTDMK_Error($"Error deleting images: {ex.Message}");
            }
        }
        public static DataTable ConvertColumnImageTypeToDataTable(DataTable dataTable, string itemCode, string lotNo, string tableName)
        {
            DataTable copyTable = dataTable.Clone();

            foreach (DataColumn col in copyTable.Columns)
            {
                if (col.DataType == typeof(byte[]))
                {
                    col.DataType = typeof(string);
                }
            }

            foreach (DataRow row in dataTable.Rows)
            {
                DataRow newRow = copyTable.NewRow();

                foreach (DataColumn col in dataTable.Columns)
                {
                    if (col.DataType == typeof(byte[]))
                    {
                        string region = null;
                        string type = null;

                        switch (tableName)
                        {
                            case "SFE_IMAGE":
                                region = row["Region"]?.ToString() ?? "";
                                type = row["Side"]?.ToString() ?? "";
                                break;
                            case "STACKUP_IMAGE":
                            case "IMPEDANCE_GRAPH":
                            case "TRACEWIDTH_GRAPH":
                            case "SOLDERMASK_IMAGE":
                            case "CQRA_HOT_OIL_Image":
                            case "THERMAL_STRESS_IMAGE":
                            case "CQRA_BHAST_IMAGE":
                            case "CQRA_SOLDERABILITY_IMAGE":
                            case "OQC_GENERAL_IMAGE":
                            case "CQRA_CHEMICAL_RESISTANCE_IMAGE":
                            case "CQRA_IR_VIA_TO_VIA_IMAGE":
                            case "CQRA_IR_TRACE_TO_TRACE_IMAGE":
                            case "CQRA_IR_LAYER_TO_LAYER_IMAGE":
                            case "CQRA_FLUX_RESIST_IMAGE":
                            case "CQRA_HIGH_SPEED_BALL_SHEAR_IMAGE":
                            case "CQRA_DIELECTRIC_WITHSTANDING_IMAGE":
                            case "ACF_CLEANING_IMAGE":
                            case "ACF_BEFORE_TAPE_TEST_IMAGE":
                            case "ACF_AFTER_TAPE_TEST_IMAGE":
                            case "ACF_GLASS_COUPON_IMAGE":
                            case "ACF_GRAPH_FORCE":
                                region = row["Region"]?.ToString() ?? "";
                                type = row["Pcs_No"]?.ToString() ?? "";
                                break;
                            case "BVH_WITH_BONDING_SHEET":
                            case "BVH_WITHOUT_BONDING_SHEET":
                            case "PLATED_THROUGH_HOLE":
                                region = itemCode + lotNo;
                                type = row["PCS_No"]?.ToString() ?? "";
                                break;
                            default:
                                region = "";
                                type = "";
                                break;
                        }
                        type += col.ColumnName;
                        newRow[col.ColumnName] = SaveImageToStorage(imageFolder, ConvertByteToBitmap(row[col] as byte[]), region, type, tableName, itemCode, lotNo);
                    }
                    else
                    {
                        newRow[col.ColumnName] = row[col];
                    }
                }

                copyTable.Rows.Add(newRow);
            }

            return copyTable;
        }

        public static DataTable ConvertTableJsonToTable(DataTable jsonTable) 
        {
            if (jsonTable == null || jsonTable.Rows.Count == 0 || !jsonTable.Columns.Contains("Data"))
                return jsonTable; 
            try
            {
                string[] columnNotConvert = new string[]
            {
                "ItemCode",
                "LotNo",
                "Operator"
            };
                //Thêm các cột cơ bản
                DataTable copyTable = new DataTable();

                foreach (DataColumn col in jsonTable.Columns)
                {
                    if (!columnNotConvert.Contains(col.ColumnName))
                    {
                        continue;
                    }
                    else
                    {
                        copyTable.Columns.Add(col.ColumnName, typeof(string));
                    }
                }

                string dataDecode = jsonTable.Rows[0]["Data"].ToString();
                if (string.IsNullOrEmpty(dataDecode))
                    return jsonTable;
                //convert list data từ json sang list thường
                List<Dictionary<string, string>> listData = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(dataDecode);
                foreach (var dataRow in listData)
                {
                    foreach (var kvp in dataRow)
                    {
                        string columnName = kvp.Key;

                        if (!copyTable.Columns.Contains(columnName))
                        {
                            copyTable.Columns.Add(columnName, typeof(string));
                        }
                    }
                    break;
                }
                foreach (var dataRow in listData)
                {
                    DataRow newRow = copyTable.NewRow();
                    foreach (DataColumn copyCol in copyTable.Columns)
                    {
                        if (jsonTable.Columns.Contains(copyCol.ColumnName))
                            newRow[copyCol.ColumnName] = jsonTable.Rows[0][copyCol.ColumnName].ToString();
                    }
                    foreach (var kvp in dataRow)
                    {
                        string columnName = kvp.Key;
                        string columnValue = kvp.Value;

                        newRow[columnName] = columnValue;
                    }
                    copyTable.Rows.Add(newRow);
                }
                return ConvertColumnStringTypeToDataTable(copyTable);
            }
            catch (Exception)
            {
                return jsonTable;
            }
        }

        public static DataTable ConvertColumnStringTypeToDataTable(DataTable dataTable)
        {
            List<string> columnExpect = new List<string>() { "NameImage" };
            if (dataTable.Rows.Count == 0)
                return dataTable;
            DataTable copyTable = dataTable.Clone();
            foreach (DataColumn col in copyTable.Columns)
            {
                if (dataTable.Rows[0][col.ColumnName].ToString().Contains(".jpg") || col.ColumnName.Contains("Image") && !columnExpect.Contains(col.ColumnName))
                {
                    col.DataType = typeof(byte[]);
                }
            }

            foreach (DataRow row in dataTable.Rows)
            {
                DataRow newRow = copyTable.NewRow();
                foreach (DataColumn col in dataTable.Columns)
                {
                    if (row[col].ToString().Contains(".jpg"))
                    {
                        newRow[col.ColumnName] = ConvertImageToByteArray(LoadImageSafely(row[col].ToString()));
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(row[col].ToString()))
                        {
                            newRow[col.ColumnName] = DBNull.Value;
                        }
                        else
                        {
                            newRow[col.ColumnName] = row[col];
                        }
                    }
                }
                copyTable.Rows.Add(newRow);
            }
            return copyTable;
        }
        #endregion

        #region Image Processing
        public static Bitmap ConvertByteToBitmap(byte[] data)
        {
            if (data == null) return null;
            using (MemoryStream ms = new MemoryStream(data))
            {
                Bitmap bitmap = new Bitmap(ms);
                return bitmap;
            }
        }

        public static byte[] ConvertImageToByteArray(Image image)
        {
            if (image == null) 
            {
                //TDMK_Message.MessageBoxTDMK_Error("Không có ảnh");
                return null;
            } 

            using (MemoryStream ms = new MemoryStream())
            {
                image.Save(ms, ImageFormat.Jpeg);
                return ms.ToArray();
            }
        }

        public static System.Drawing.Image LoadImageSafely(string filePath)
        {
            try
            {
                if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                    return null;

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    return new Bitmap(System.Drawing.Image.FromStream(stream));
                }
            }
            catch (Exception ex)
            {
                //TDMK_Message.MessageBoxTDMK_Error($"Error loading image {filePath}: {ex.Message}");
                return null;
            }
        }

        public static string SaveImageToStorage(String savedFolder, System.Drawing.Image image, string region, string type, string tableName, string itemcode, string lotno)
        {
            try
            {
                if (image == null) return "";

                string fileName = $"{itemcode}-{lotno}-{region}-{type}-{DateTime.Now:yyyyMMdd_HHmmss}.jpg";
                string folder = Path.Combine(savedFolder, tableName, itemcode + "_" + lotno);
                string filePath = Path.Combine(folder, fileName);

                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);

                ImageCodecInfo jpgEncoder = GetEncoder(ImageFormat.Jpeg);
                EncoderParameters encoderParams = new EncoderParameters(1);
                EncoderParameter qualityParam = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 60L);
                encoderParams.Param[0] = qualityParam;

                using (Bitmap bmp = new Bitmap(image))
                {
                    bmp.Save(filePath, jpgEncoder, encoderParams);
                }

                return filePath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving image: {ex.Message}");
                return "";
            }
        }

        private static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            return ImageCodecInfo.GetImageDecoders().FirstOrDefault(c => c.FormatID == format.Guid);
        }
        #endregion

        #region Excel Operations
        private static ExcelPackage open_excel(string file_name)
        {
            ExcelPackage result = null;
            FileInfo newFile = new FileInfo(file_name);
            if (File.Exists(file_name))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                result = new ExcelPackage(newFile);
            }

            return result;
        }

        public static string Find_Cell_Location(string search_key, ExcelWorksheet tar_wrksht, bool en_contain = false)
        {
            ExcelAddressBase dimension = tar_wrksht.Dimension;
            int rows = dimension.Rows;
            int columns = dimension.Columns;
            List<char> reject_chars = new List<char> { ' ', '-', '_', '\r', '\n', '&' };
            string result = "";
            string text = remove_special_chars(search_key, reject_chars).ToUpper();
            for (int i = 1; i < rows; i++)
            {
                for (int j = 1; j < columns; j++)
                {
                    ExcelRangeBase excelRangeBase = tar_wrksht.Cells[i, j];
                    string text2 = remove_special_chars(checkDBNull(excelRangeBase.Value), reject_chars).ToUpper();
                    if (en_contain)
                    {
                        if (text2.Contains(text))
                        {
                            return excelRangeBase.Address;
                        }
                    }
                    else if (text2 == text)
                    {
                        return excelRangeBase.Address;
                    }
                }
            }

            return result;
        }

        private static string checkDBNull(object src_str)
        {
            if (src_str != null)
            {
                return src_str.ToString();
            }

            return "";
        }

        private static string remove_special_chars(string src_string, List<char> reject_chars)
        {
            return new string((from x in src_string.ToArray()
                               where reject_chars.IndexOf(x) == -1
                               select x).ToArray());
        }

        public static DataTable ReadExcelFile(string filePath)
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.AddRange
                (
                    new DataColumn[] {
                        new DataColumn("Sample Name"),
                        new DataColumn("Serial Number")
                    }
                );

            using (ExcelPackage format_pack = open_excel(filePath))
            {
                ExcelWorksheet worksheet = format_pack.Workbook.Worksheets[0];
                string sampleNameAddress = Find_Cell_Location("Sample Name", worksheet);

                for (int i = 1; i < worksheet.Rows.EndRow; i++)
                {
                    string sampleValue = worksheet.Cells[sampleNameAddress].Offset(i, 0).Text;
                    string serialValue = worksheet.Cells[sampleNameAddress].Offset(i, 1).Text;
                    if (string.IsNullOrEmpty(sampleValue) && string.IsNullOrEmpty(serialValue))
                    {
                        break;
                    }
                    DataRow row = dataTable.NewRow();
                    row[0] = sampleValue;
                    row[1] = serialValue;
                    dataTable.Rows.Add(row);
                }

                return dataTable;
            }
        }

        public static ExcelPicture AddImageToMergedCell(
            ExcelWorksheet ws,
            byte[] imageData,
            string imageName,
            string targetCellAddress,
            double widthScale = 1.0,
            double heightScale = 1.0,
            bool keepAspectRatio = true,
            int positionX = 0,
            int positionY = 0,
            int jpegQuality = 90)
        {
            if (imageData == null || imageData.Length == 0)
                throw new ArgumentException("Dữ liệu ảnh không hợp lệ.");

            try
            {
                using (var inputStream = new MemoryStream(imageData))
                using (var skBitmap = SKBitmap.Decode(inputStream))
                {
                    if (skBitmap == null)
                        throw new InvalidOperationException("Không thể giải mã dữ liệu ảnh.");

                    using (var outputStream = new MemoryStream())
                    {
                        skBitmap.Encode(outputStream, SKEncodedImageFormat.Jpeg, jpegQuality);
                        outputStream.Seek(0, SeekOrigin.Begin);

                        var picture = ws.Drawings.AddPicture(imageName, outputStream);
                        var targetCell = ws.Cells[targetCellAddress];

                        ExcelRange mergedRange = null;
                        foreach (var mergedAddress in ws.MergedCells)
                        {
                            var range = ws.Cells[mergedAddress];
                            if (range.Start.Row <= targetCell.Start.Row &&
                                range.End.Row >= targetCell.End.Row &&
                                range.Start.Column <= targetCell.Start.Column &&
                                range.End.Column >= targetCell.End.Column)
                            {
                                mergedRange = range;
                                break;
                            }
                        }

                        if (mergedRange != null)
                        {
                            double mergedWidth = CalculateMergedWidth(ws, mergedRange);
                            double mergedHeight = CalculateMergedHeight(ws, mergedRange);

                            int initialImageWidth = (int)(mergedWidth * widthScale);
                            int initialImageHeight = (int)(mergedHeight * heightScale);

                            int imageWidth, imageHeight;
                            var scaledSize = CalculateScaledSize(
                                    new SKSize(skBitmap.Width, skBitmap.Height),
                                    new SKSize((float)mergedWidth, (float)mergedHeight),
                                    widthScale,
                                    heightScale
                                );

                            if (keepAspectRatio)
                            {
                                imageWidth = scaledSize.Width;
                                imageHeight = scaledSize.Height;
                            }
                            else
                            {
                                imageWidth = initialImageWidth;
                                imageHeight = initialImageHeight;
                            }

                            picture.SetSize(imageWidth, imageHeight);

                            double cellCenterX = mergedWidth / 2.0;
                            double cellCenterY = mergedHeight / 2.0;
                            double imageCenterX = imageWidth / 2.0;
                            double imageCenterY = imageHeight / 2.0;

                            double offsetX = cellCenterX - imageCenterX;
                            double offsetY = cellCenterY - imageCenterY;

                            if (keepAspectRatio)
                            {
                                int positionXNew = 0;
                                int positionYNew = 0;
                                double width = 0;
                                double height = 0;
                                int i = 0;
                                for (int col = mergedRange.Start.Column; col <= mergedRange.End.Column; col++)
                                {
                                    width += ws.Column(col).Width * 7.0;
                                    if (width <= offsetX)
                                    {
                                        i++;
                                    }
                                    else
                                    {
                                        offsetX = width - offsetX;
                                        positionXNew = i;
                                        break;
                                    }
                                }
                                for (int row = mergedRange.Start.Row; row <= mergedRange.End.Row; row++)
                                {
                                    height += ws.Row(row).Height * 1.33;
                                    if (height <= offsetY)
                                    {
                                        i++;
                                    }
                                    else
                                    {
                                        offsetY = height - offsetY;
                                        positionYNew = i;
                                        break;
                                    }
                                }
                                picture.SetPosition(
                                    mergedRange.Start.Row + positionYNew,
                                    (int)-offsetY + positionY,
                                    mergedRange.Start.Column + positionXNew,
                                    (int)-offsetX + positionX
                                );
                            }
                            else
                            {
                                picture.SetPosition(
                                    mergedRange.Start.Row - 1,
                                    (int)offsetY + (int)positionY,
                                    mergedRange.Start.Column - 1,
                                    (int)offsetX + (int)positionX
                                );
                            }
                        }
                        else
                        {
                            double colWidth = ws.Column(targetCell.Start.Column).Width * 7.0;
                            double rowHeight = ws.Row(targetCell.Start.Row).Height * 1.33;

                            int imageWidth, imageHeight;
                            var scaledSize = CalculateScaledSize(
                                new SKSize(skBitmap.Width, skBitmap.Height),
                                new SKSize((float)colWidth, (float)rowHeight),
                                widthScale,
                                heightScale
                            );

                            if (keepAspectRatio)
                            {
                                imageWidth = scaledSize.Width;
                                imageHeight = scaledSize.Height;
                            }
                            else
                            {
                                imageWidth = (int)(colWidth * widthScale);
                                imageHeight = (int)(rowHeight * heightScale);
                            }

                            picture.SetSize(imageWidth, imageHeight);

                            double offsetX = (colWidth - imageWidth) / 2.0;
                            double offsetY = (rowHeight - imageHeight) / 2.0;

                            picture.SetPosition(
                                targetCell.Start.Row - 1,
                                (int)offsetY + (int)positionY,
                                targetCell.Start.Column - 1,
                                (int)offsetX + (int)positionX
                            );
                        }

                        return picture;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi thêm ảnh vào Excel: {ex.Message}");
                return null;
            }
        }

        private static double CalculateMergedWidth(ExcelWorksheet ws, ExcelRangeBase mergedRange)
        {
            double width = 0;
            for (int col = mergedRange.Start.Column; col <= mergedRange.End.Column; col++)
            {
                width += ws.Column(col).Width * 7.0;
            }
            return width;
        }

        private static double CalculateMergedHeight(ExcelWorksheet ws, ExcelRangeBase mergedRange)
        {
            double height = 0;
            for (int row = mergedRange.Start.Row; row <= mergedRange.End.Row; row++)
            {
                height += ws.Row(row).Height * 1.333;
            }
            return height;
        }

        private static SKSizeI CalculateScaledSize(
            SKSize imageSize,
            SKSize containerSize,
            double widthScale,
            double heightScale)
        {
            double targetWidth = containerSize.Width * widthScale;
            double targetHeight = containerSize.Height * heightScale;

            double ratio = Math.Min(
                targetWidth / imageSize.Width,
                targetHeight / imageSize.Height
            );

            return new SKSizeI(
                (int)(imageSize.Width * ratio),
                (int)(imageSize.Height * ratio)
            );
        }
        #endregion

        #region File Operations
        public static DataTable ReadCsvFile(string filePath)
        {
            DataTable dataTable = new DataTable();
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)))
            {
                csv.Read();
                csv.ReadHeader();
                foreach (var header in csv.HeaderRecord)
                {
                    dataTable.Columns.Add(header.Trim());
                }

                while (csv.Read())
                {
                    DataRow row = dataTable.NewRow();
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        row[i] = csv.GetField(i);
                    }
                    dataTable.Rows.Add(row);
                }
            }
            return dataTable;
        }
        #endregion

        #region UI Utilities
        public static async Task RunWithLoadingAsync(Form parent, Func<Task> asyncAction)
        {
            Frm_Loading loadingForm = null;
            try
            {
                await Task.Run(() =>
                {
                    parent.Invoke(new Action(() =>
                    {
                        loadingForm = new Frm_Loading();
                        loadingForm.Show(parent);
                        loadingForm.BringToFront();
                        loadingForm.TopMost = true;
                    }));
                });

                await asyncAction();
            }
            finally
            {
                if (loadingForm != null)
                {
                    parent.Invoke(new Action(() =>
                    {
                        if (!loadingForm.IsDisposed && loadingForm.Visible)
                        {
                            loadingForm.Close();
                            loadingForm.Dispose();
                        }
                    }));
                }
            }
        }
        #endregion
    }
}