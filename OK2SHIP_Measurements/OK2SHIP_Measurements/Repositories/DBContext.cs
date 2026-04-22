using IniLibs;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TDMK_SQL;

namespace OK2SHIP_Measurements.Repositories
{
    class DBContext : IDisposable
    {
        private SqlConnection SqlConnection;
        IniFile TDMK_init;
        private string _nasAddress = "";
        private TDMK_SQL_Lib TDMK_SQL = new TDMK_SQL_Lib();
        /// <summary>
        /// Tạo một bảng mới trong database MSSQL với tên được cung cấp và cấu trúc cơ bản.
        /// Lưu ý: Hàm này tạo một cấu trúc bảng mẫu. Để tạo bảng với cấu trúc tùy chỉnh,
        /// bạn cần truyền vào định nghĩa cột dưới dạng tham số.
        /// </summary>
        /// <param name="connectionString">Chuỗi kết nối đến cơ sở dữ liệu MSSQL.</param>
        /// <param name="tableName">Tên của bảng mới muốn tạo.</param>
        /// <param name="schemaName">Tên của schema mà bảng sẽ thuộc về (mặc định là 'dbo').</param>
        /// <returns>True nếu bảng được tạo thành công, False nếu có lỗi hoặc bảng đã tồn tại.</returns>
        /// <exception cref="SqlException">Ném ra nếu có lỗi xảy ra trong quá trình thực thi SQL.</exception>
        /// <exception cref="Exception">Ném ra cho các lỗi khác.</exception>
        //public bool CreateNewTable(string tableName, string schemaName = "dbo")
        //{
        //    // Kiểm tra xem bảng đã tồn tại chưa để tránh lỗi
        //    if (DoesTableExist(tableName, schemaName))
        //    {
        //        Console.WriteLine($"Bảng '{schemaName}.{tableName}' đã tồn tại. Không tạo lại.");
        //        return false;
        //    }

        //    // Định nghĩa cấu trúc bảng mẫu.
        //    // Bạn có thể tùy chỉnh các cột ở đây hoặc truyền chúng vào như một tham số phức tạp hơn.
        //    string createTableSql = $@"
        //    CREATE TABLE [{schemaName}].[{tableName}] (
        //        Id INT PRIMARY KEY IDENTITY(1,1),
        //        Name NVARCHAR(255) NOT NULL,
        //        Description NVARCHAR(MAX) NULL,
        //        CreatedAt DATETIME DEFAULT GETDATE()
        //    );";

        //    try
        //    {
        //        // Thực thi lệnh CREATE TABLE
        //        ExecuteNonQueryCommand(createTableSql);
        //        Console.WriteLine($"Bảng '{schemaName}.{tableName}' đã được tạo thành công.");
        //        return true;
        //    }
        //    catch (SqlException ex)
        //    {
        //        Console.WriteLine($"Lỗi SQL khi tạo bảng '{schemaName}.{tableName}': {ex.Message}");
        //        throw; // Ném lại lỗi để caller có thể xử lý
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"Đã xảy ra lỗi chung khi tạo bảng '{schemaName}.{tableName}': {ex.Message}");
        //        throw; // Ném lại lỗi
        //    }
        //}
        public bool DoesColumnExist(string tableName, string columnName, string schemaName = "dbo")
        {
            // Sử dụng INFORMATION_SCHEMA.COLUMNS vì nó chuẩn ANSI SQL và dễ đọc
            string checkSql = @"
            SELECT COUNT(*)
            FROM INFORMATION_SCHEMA.COLUMNS
            WHERE TABLE_SCHEMA = @SchemaName
            AND TABLE_NAME = @TableName
            AND COLUMN_NAME = @ColumnName;";

            // Hoặc có thể dùng sys.columns cho SQL Server-specific nếu cần:
            // string checkSql = @"
            //     SELECT COUNT(*)
            //     FROM sys.columns AS C
            //     INNER JOIN sys.tables AS T ON C.object_id = T.object_id
            //     INNER JOIN sys.schemas AS S ON T.schema_id = S.schema_id
            //     WHERE S.name = @SchemaName
            //     AND T.name = @TableName
            //     AND C.name = @ColumnName;";

            SqlParameter[] parameters = new SqlParameter[]
            {
            new SqlParameter("@SchemaName", SqlDbType.NVarChar, 128) { Value = schemaName },
            new SqlParameter("@TableName", SqlDbType.NVarChar, 128) { Value = tableName },
            new SqlParameter("@ColumnName", SqlDbType.NVarChar, 128) { Value = columnName }
            };

            try
            {
                DataTable result = ExecuteSqlToDataTable(checkSql);

                // Kiểm tra kết quả: nếu Count > 0 thì cột tồn tại
                if (result != null && result.Rows.Count > 0)
                {
                    return Convert.ToInt32(result.Rows[0][0]) > 0;
                }
                return false; // Không có dòng nào hoặc DataTable rỗng
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi kiểm tra sự tồn tại của cột '{columnName}' trong bảng '{tableName}' (schema '{schemaName}'): {ex.Message}");
                throw; // Ném lại lỗi để xử lý ở tầng cao hơn
            }
        }

        public bool DoesTableExist(string tableName, string schemaName = "dbo")
        {
            // Sử dụng INFORMATION_SCHEMA.TABLES vì nó chuẩn ANSI SQL và dễ đọc
            string checkSql = @"
            SELECT COUNT(*)
            FROM INFORMATION_SCHEMA.TABLES
            WHERE TABLE_SCHEMA = @SchemaName
            AND TABLE_NAME = @TableName;";

            // Có thể dùng sys.tables cho SQL Server-specific nếu cần:
            // string checkSql = @"
            //     SELECT COUNT(*)
            //     FROM sys.tables AS T
            //     INNER JOIN sys.schemas AS S ON T.schema_id = S.schema_id
            //     WHERE S.name = @SchemaName
            //     AND T.name = @TableName;";

            SqlParameter[] parameters = new SqlParameter[]
            {
            new SqlParameter("@SchemaName", SqlDbType.NVarChar, 128) { Value = schemaName }, // Kích thước schema_name
            new SqlParameter("@TableName", SqlDbType.NVarChar, 128) { Value = tableName }   // Kích thước table_name
            };

            try
            {
                DataTable result = ExecuteSqlToDataTable(checkSql);

                // Kiểm tra kết quả: nếu Count > 0 thì bảng tồn tại
                if (result != null && result.Rows.Count > 0)
                {
                    return Convert.ToInt32(result.Rows[0][0]) > 0;
                }
                return false; // Không có dòng nào hoặc DataTable rỗng
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi kiểm tra sự tồn tại của bảng '{tableName}' trong schema '{schemaName}': {ex.Message}");
                throw; // Ném lại lỗi để xử lý ở tầng cao hơn
            }
        }
        public DataTable ExecuteSqlToDataTable(string sqlQueryOrCommand)
        {
            DataTable dataTable = new DataTable();

            // Sử dụng 'using' statement để đảm bảo các đối tượng được giải phóng đúng cách

            using (SqlCommand command = new SqlCommand(sqlQueryOrCommand, SqlConnection))
            {

                try
                {

                    // SqlDataAdapter dùng để điền dữ liệu từ SqlCommand vào DataTable
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        adapter.Fill(dataTable); // Điền dữ liệu vào DataTable
                    }
                }
                catch (SqlException ex)
                {
                    // Xử lý lỗi SQL cụ thể
                    Console.WriteLine($"Lỗi SQL xảy ra: {ex.Message}");
                    throw; // Ném lại lỗi để caller có thể xử lý
                }
                catch (Exception ex)
                {
                    // Xử lý các lỗi chung khác
                    Console.WriteLine($"Đã xảy ra lỗi: {ex.Message}");
                    throw; // Ném lại lỗi
                }
            }

            return dataTable;
        }

        public DBContext(string catalog = null)
        {
            try
            {
                if (catalog == null)
                {
                    catalog = "OK2SHIP_SMT";
                }
                //config file
                string app_path = System.Windows.Forms.Application.StartupPath;
                //app_path = @"\\10.212.6.212\Saomai\QA\TDMK_DATA\Test_Areas\FPCA OK2SHIP Auto System(temp2)\VHX-IMADA";
                string config_path = Path.Combine(app_path.Replace(@"\FPCA OK2SHIP Auto System\VHX-IMADA", ""), "Config.ini");
                //string config_path = Path.Combine(app_path.Replace(@"\VHX-IMADA", ""), "Config.ini");
                TDMK_init = new IniFile(config_path);
                //data_loc = TDMK_init.Read("Format_Folder", "SMT_Config");
                //data_loc = Path.Combine(System.Windows.Forms.Application.StartupPath.Replace(@"\VHX-IMADA", ""));
                string server_name = TDMK_init.Read("Server", "SMT_Config");
                string server_acc = TDMK_init.Read("Account", "SMT_Config");
                string server_pass = TDMK_init.Read("Password", "SMT_Config");
                _nasAddress = TDMK_init.Read("NasAddress", "SMT_Config");
                SqlConnection = initial_data($";Connection Timeout=6000;Data Source={server_name};Initial Catalog={catalog};User ID={server_acc};Password='{server_pass}';Encrypt=True;TrustServerCertificate=True;");
                ///////////////////////

                SqlConnection.Open();
            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi kết nối sql: {ex.Message}");
            }
        }
        /// <summary>
        /// Hàm update
        /// Cột nào update phải giữ nguyên
        /// nếu giá trị là "" thì vẫn update
        /// nếu giá trị là null thì không update
        /// </summary>
        /// <param name="TableName"></param>
        /// <param name="dt_set"></param>
        /// <param name="listColCondition"></param>
        /// <param name="valueCondition"></param>
        /// <returns></returns>
        public int Update(string TableName, DataTable dt_set, string[] nameColCondition)
        {
            // Tạo câu lệnh SQL UPDATE
            List<string> listColCondition = new List<string>();
            foreach (DataRow row in dt_set.Rows)
            {
                string st = $"WHEN ";
                foreach (var item in nameColCondition)
                {
                    st += $" {item} = '{row[item]}' THEN  ";
                }
                listColCondition.Add(st);
            }

            string commandText = $"UPDATE [{TableName}] SET ";
            int c = 0;
            foreach (DataColumn Column in dt_set.Columns)
            {
                if (nameColCondition.Contains(Column.ColumnName))
                {

                }
                else
                {
                    commandText += $"{Column} = CASE ";
                    int i = 0;
                    foreach (DataRow row in dt_set.Rows)
                    {
                        commandText += $"{listColCondition[i]} '{row[Column]}'";
                    }
                    commandText += $" ELSE {Column} END ";
                    c++;
                    if (dt_set.Columns.Count - 1 != c)
                    {
                        commandText += " , ";
                    }
                }

            }
            using (SqlCommand command = new SqlCommand($"{commandText}", SqlConnection))
            {
                if (SqlConnection.State != ConnectionState.Open)
                {
                    SqlConnection.Open();
                }
                return command.ExecuteNonQuery();
            }
        }
        public int GetID(string tableName, string id = "id")
        {
            using (SqlCommand command = new SqlCommand($"SELECT MAX({id}) FROM [{tableName}]", SqlConnection))
            {
                object result = command.ExecuteScalar();
                int maxId = 0;
                if (result != null && result != DBNull.Value)
                {
                    maxId = Convert.ToInt32(result);
                }
                return maxId;
            }

        }
        public int InsertImageAndGetId(byte[] image, string tableName)
        {
            int newImageId = GetID(tableName) + 1;
            string query = @"
            INSERT INTO [dbo].[PACKAGING_IMAGE] (ID, Image )
            VALUES (@ID, @Image);";

            using (SqlCommand command = new SqlCommand(query, SqlConnection))
            {
                command.Parameters.AddWithValue("@ID", newImageId);
                command.Parameters.Add("@Image", SqlDbType.VarBinary, -1).Value = image;
                command.ExecuteScalar();
            }
            return newImageId;
        }

        public DataTable GetTableStructure(string tableName)
        {
            using (SqlCommand command = new SqlCommand($"SELECT TOP (0) * FROM [{tableName}]", SqlConnection))
            {
                using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                {
                    DataTable tableStructure = new DataTable();
                    adapter.Fill(tableStructure);
                    return tableStructure;
                }
            }
        }
        public void Dispose()
        {
            SqlConnection.Close();
        }
        public SqlConnection initial_data(string sqlConnection)
        {
            return new SqlConnection(sqlConnection);
        }
        #region Generate Command
        private static string GenerateInsertCommand(DataTable dataTable, string tableName)
        {
            string columns = string.Join(", ", dataTable.Columns.Cast<DataColumn>().Select(column => "[" + column.ColumnName + "]"));
            string parameters = string.Join(", ", dataTable.Columns.Cast<DataColumn>().Select(column => "@" + column.ColumnName.Replace("-", "")));

            return $"INSERT INTO [{tableName}] ({columns}) VALUES ({parameters.Replace(" ", "")})";
        }

        private static string GenerateDeleteCommand(string tableName, string[] fillter, string[] value)
        {
            string command = $"DELETE FROM [{tableName}] WHERE ";
            for (int i = 0; i < fillter.Length; i++)
            {
                command += fillter[i] + $" = \'{value[i]}\'";
                if (i < fillter.Length - 1)
                {
                    command += " AND ";
                }
            }
            return command;
        }

        public string AddCommandPaging(string command, string[] orderBy, int pageNumber = 1, int pageSize = 20)
        {

            if (orderBy != null)
            {
                string commandOrder = " ORDER BY ";
                foreach (string column in orderBy)
                {
                    commandOrder += $"{column}";
                }
                command += commandOrder;
            }

            return $"{command} OFFSET ({pageNumber} - 1) * {pageSize} ROWS FETCH NEXT {pageSize} ROWS ONLY;";
        }
        public string GenerateSelectCommand(string tableName, string[] colName = null, string[] selectColumn = null)
        {
            string command = $"FROM [{tableName}] ";
            if (selectColumn == null)
            {
                command = $"SELECT * {command}";
            }
            else
            {
                string selectClause = "";
                for (int i = 0; i < selectColumn.Length; i++)
                {
                    selectClause += $"[{selectColumn[i]}]"; // Using parameterized queries
                    if (i < selectColumn.Length - 1)
                    {
                        selectClause += ", ";
                    }
                }
                command = $"SELECT {selectClause} {command}";
            }
            if (colName != null && colName.Count() > 0)
            {
                string whereClause = "";

                whereClause = "WHERE ";
                for (int i = 0; i < colName.Length; i++)
                {
                    whereClause += $"[{colName[i]}] = @{colName[i]}"; // Using parameterized queries
                    if (i < colName.Length - 1)
                    {
                        whereClause += " AND ";
                    }
                }

                command = $"{command} {whereClause} ";
            }
            return command;
        }
        #endregion
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataTable">Dữ liệu cần luuw</param>
        /// <param name="TableSql">Bảng sẽ lưu</param>
        /// <param name="colCompare">cột mình sẽ so sánh để thay thế</param>
        /// <param name="mappingColName">cột trong sql để mình so sánh</param>
        /// <param name="colID">Đánh dấu cột id</param>
        /// <returns></returns>
        public int BuckDataTable(DataTable dataTable, string TableSql, string[] colCompare, string[] mappingColName = null, string colID = null)
        {
            if (dataTable.Rows.Count <= 0)
            {
                return 0;
            }
            int id = GetID(TableSql) + 1;
            if (colID != null)
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    row[colID] = id++;
                }
            }
            using (SqlTransaction transaction = SqlConnection.BeginTransaction())
            {
                IList<string> valueCompare = new List<string>();
                foreach (string col in colCompare)
                {
                    valueCompare.Add(dataTable.Rows[0][col].ToString());
                }
                string deleteCommand = GenerateDeleteCommand(TableSql, colCompare, valueCompare.ToArray());
                using (SqlCommand command = new SqlCommand())
                {
                    command.Connection = SqlConnection;
                    command.Transaction = transaction;
                    command.CommandText = deleteCommand;
                    command.ExecuteNonQuery();
                }
                if (mappingColName != null)
                {
                    for (int i = 0; i < mappingColName.Length; i++)
                    {
                        dataTable.Columns[i].ColumnName = mappingColName[i];
                    }
                }
                try
                {

                    //foreach (DataRow row in dataTable.Rows)
                    //{
                    //    using (SqlCommand command = new SqlCommand())
                    //    {
                    //        command.Connection = SqlConnection;
                    //        command.Transaction = transaction;
                    //        command.CommandText = GenerateInsertCommand(dataTable, TableSql);

                    //        // Thêm tham số cho từng cột
                    //        foreach (DataColumn column in dataTable.Columns)
                    //        {
                    //            command.Parameters.AddWithValue("@" + column.ColumnName.Replace("-", ""), row[column.ColumnName] ?? DBNull.Value);
                    //        }

                    //        command.ExecuteNonQuery();
                    //    }
                    //}
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(SqlConnection, SqlBulkCopyOptions.Default | SqlBulkCopyOptions.FireTriggers, transaction))
                    {


                        // Đặt tên bảng đích trong cơ sở dữ liệu
                        bulkCopy.DestinationTableName = TableSql;

                        // Thiết lập thời gian chờ (nếu cần)
                        bulkCopy.BulkCopyTimeout = 600; // Ví dụ: 60 giây

                        // Map các cột từ DataTable đến các cột trong bảng SQL Server
                        foreach (DataColumn col in dataTable.Columns)
                        {
                            bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName);
                        }

                        // Thực hiện bulk insert
                        bulkCopy.WriteToServer(dataTable);

                    }
                    transaction.Commit();
                    return dataTable.Rows.Count;
                }

                catch (Exception ex)
                {
                    transaction.Rollback();
                    throw ex;
                }
            }
        }
        public static void InsertDataTableBatch(string connectionString, string tableName, DataTable dataTable)
        {
            // Kiểm tra xem DataTable có dữ liệu hay không
            if (dataTable == null || dataTable.Rows.Count == 0)
            {
                Console.WriteLine("DataTable không có dữ liệu để insert.");
                return;
            }

            // Tạo đối tượng SqlBulkCopy
            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connectionString))
            {
                try
                {
                    // Đặt tên bảng đích trong cơ sở dữ liệu
                    bulkCopy.DestinationTableName = tableName;

                    // Thiết lập thời gian chờ (nếu cần)
                    // Map các cột từ DataTable đến các cột trong bảng SQL Server
                    // Quan trọng: Tên cột trong DataTable phải trùng với tên cột trong bảng SQL Server
                    foreach (DataColumn col in dataTable.Columns)
                    {
                        bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName);
                    }

                    // Thực hiện bulk insert
                    bulkCopy.WriteToServer(dataTable);

                    Console.WriteLine($"Đã insert thành công {dataTable.Rows.Count} bản ghi vào bảng '{tableName}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Lỗi trong quá trình bulk insert: {ex.Message}");
                }
            }
        }

        public string[] checkListIsExist(string[] listCheck, string TabltName, string colname, string selectColumn = "")
        {
            if (listCheck == null || listCheck.Count() == 0)
            {
                return null; // Return empty list if input is empty or null
            }
            string itemCodeList = string.Join("','", listCheck);
            string join = string.Join(",", new[] { colname, selectColumn });
            join = join.Trim().TrimEnd(',');
            string sql = $"SELECT {join} FROM {TabltName} WHERE {colname} IN ('{itemCodeList}')";
            List<string> existingItemCodes = new List<string>();

            try
            {
                using (SqlCommand command = new SqlCommand(sql, SqlConnection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string z = reader.GetValue(join.Split(',').Count() - 1).ToString();

                            existingItemCodes.Add(z); // Assuming ItemCode is the first column and a string
                        }
                    }
                }

            }
            catch (SqlException ex)
            {
                // Handle database exceptions (log, throw, etc.)
                throw new Exception($"Database error: {ex.Message}");
                // Consider throwing a more specific exception or logging the error.
            }
            catch (Exception ex)
            {
                throw new Exception($"General error: {ex.Message}");
            }

            return existingItemCodes.ToArray();
        }
        public int DeleteData(string tableName, string[] conditionColumn, string[] conditionValue)
        {
            int rowsAffected = 0;

            try
            {
                string query = $"DELETE FROM {tableName} WHERE ";
                for (int i = 0; i < conditionColumn.Count(); i++)
                {
                    query += $" {conditionColumn[i].Trim()} = '{conditionValue[i].Trim()}'";
                    if (i != conditionValue.Count() - 1)
                    {
                        query += " AND ";
                    }
                }

                using (SqlCommand command = new SqlCommand(query, SqlConnection))
                {
                    rowsAffected = command.ExecuteNonQuery();
                }

            }
            catch (SqlException ex)
            {
                // Xử lý lỗi SQL (ghi log, ném ngoại lệ, v.v.)
                throw new Exception($"Lỗi SQL: {ex.Message}");
                // Cân nhắc ném một ngoại lệ cụ thể hơn hoặc ghi log lỗi.
            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi chung: {ex.Message}");
            }
            return rowsAffected;
        }
        public int DeleteData(string tableName, string conditionColumn, string[] conditionValue)
        {
            int rowsAffected = 0;

            try
            {
                string query = $"DELETE FROM {tableName} WHERE ";
                for (int i = 0; i < conditionValue.Count(); i++)
                {
                    query += $" {conditionColumn} = '{conditionValue[i].Trim()}'";
                    if (i != conditionValue.Count() - 1)
                    {
                        query += " OR ";
                    }
                }

                using (SqlCommand command = new SqlCommand(query, SqlConnection))
                {
                    rowsAffected = command.ExecuteNonQuery();
                }

            }
            catch (SqlException ex)
            {
                // Xử lý lỗi SQL (ghi log, ném ngoại lệ, v.v.)
                throw new Exception($"Lỗi SQL: {ex.Message}");
                // Cân nhắc ném một ngoại lệ cụ thể hơn hoặc ghi log lỗi.
            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi chung: {ex.Message}");
            }

            return rowsAffected;
        }
        public int SaveDataTable(DataTable dataTable, string TableSql, string[] mappingColName = null, string colID = "NONE")
        {
            if (!colID.Equals("NONE"))
            {

                int id = -1;
                try
                {
                    id = GetID(TableSql, colID) + 1;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                if (id != -1)
                {
                    foreach (DataRow row in dataTable.Rows)
                    {
                        row[colID] = id++;
                    }
                }
            }
            using (SqlTransaction transaction = SqlConnection.BeginTransaction())
            {
                if (mappingColName != null)
                {
                    for (int i = 0; i < mappingColName.Length; i++)
                    {
                        dataTable.Columns[i].ColumnName = mappingColName[i];
                    }
                }
                //try
                //{
                foreach (DataRow row in dataTable.Rows)
                {
                    using (SqlCommand command = new SqlCommand())
                    {
                        command.Connection = SqlConnection;
                        command.Transaction = transaction;
                        command.CommandText = GenerateInsertCommand(dataTable, TableSql);

                        // Thêm tham số cho từng cột
                        foreach (DataColumn column in dataTable.Columns)
                        {
                            command.Parameters.AddWithValue("@" + column.ColumnName, row[column.ColumnName] ?? (object)DBNull.Value);
                            var ks = typeof(DBNull);
                        }

                        command.ExecuteNonQuery();
                    }
                }

                transaction.Commit();
                return dataTable.Rows.Count;
                //}

                //catch (Exception ex)
                //{
                //    transaction.Rollback();
                //    throw ex;
                //}
            }
        }

        public DataTable LoadDataTable(string tableName, string[] colName, string[] valueName, string[] selectColumn = null)
        {
            DataTable dataTable = new DataTable();
            try
            {
                string commandSql = GenerateSelectCommand(tableName, colName, selectColumn);
                using (SqlCommand command = new SqlCommand(commandSql, SqlConnection))
                {

                    if (colName != null)
                    {
                        for (int i = 0; i < colName.Length; i++)
                        {
                            command.Parameters.AddWithValue($"@{colName[i]}", valueName[i]);
                        }
                    }


                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        adapter.Fill(dataTable);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dataTable;

        }

        public DataTable LoadDataTableOfPath(string tableName, string[] colName, string[] valueName, string[] selectColumn = null, int pageNumber = 1, int pageSize = 20, string[] orderBy = null)
        {
            DataTable dataTable = new DataTable();
            try
            {
                if (orderBy == null)
                {
                    orderBy = new string[] { "ID" };
                }
                string commandSql = AddCommandPaging(GenerateSelectCommand(tableName, colName, selectColumn), orderBy, pageNumber, pageSize);

                using (SqlCommand command = new SqlCommand(commandSql, SqlConnection))
                {
                    if (colName != null)
                    {
                        for (int i = 0; i < colName.Length; i++)
                        {
                            command.Parameters.AddWithValue($"@{colName[i]}", valueName[i]);
                        }
                    }


                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        adapter.Fill(dataTable);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dataTable;

        }
    }
}
