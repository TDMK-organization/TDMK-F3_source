using IniLibs;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TDMK_SQL;

namespace OK2SHIP_SMT.Repositories
{
    class DBContext : IDisposable
    {
        private SqlConnection SqlConnection;
        IniFile TDMK_init;

        private TDMK_SQL_Lib TDMK_SQL = new TDMK_SQL_Lib();
        public DBContext()
        {
            //config file
            string app_path = System.Windows.Forms.Application.StartupPath;
            //app_path = @"\\10.212.6.212\Saomai\QA\TDMK_DATA\Test_Areas\FPCA OK2SHIP Auto System(temp2)\VHX-IMADA";
            string config_path = Path.Combine(app_path.Replace(@"\VHX-IMADA", ""), "Config.ini");
            //string config_path = Path.Combine(app_path.Replace(@"\VHX-IMADA", ""), "Config.ini");
            TDMK_init = new IniFile(config_path);
            //data_loc = TDMK_init.Read("Format_Folder", "SMT_Config");
            //data_loc = Path.Combine(System.Windows.Forms.Application.StartupPath.Replace(@"\VHX-IMADA", ""));
            string server_name = TDMK_init.Read("Server", "SMT_Config");
            string server_acc = TDMK_init.Read("Account", "SMT_Config");
            string server_pass = TDMK_init.Read("Password", "SMT_Config");
            SqlConnection = initial_data($"Data Source={server_name};Initial Catalog=OK2SHIP_SMT;User ID={server_acc};Password='{server_pass}';Encrypt=True;TrustServerCertificate=True;");
            ///////////////////////
            SqlConnection.Open();

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
                                command.Parameters.AddWithValue("@" + column.ColumnName.Replace("-", ""), row[column.ColumnName] ?? DBNull.Value);
                            }

                            command.ExecuteNonQuery();
                        }
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
        public string[] checkListIsExist(string[] listCheck, string TabltName, string colname)
        {
            if (listCheck == null || listCheck.Count() == 0)
            {
                return null; // Return empty list if input is empty or null
            }
            string itemCodeList = string.Join("','", listCheck);
            string sql = $"SELECT ItemCode FROM {TabltName} WHERE ItemCode IN ('{itemCodeList}')";
            List<string> existingItemCodes = new List<string>();

            try
            {
                using (SqlCommand command = new SqlCommand(sql, SqlConnection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            existingItemCodes.Add(reader.GetString(0)); // Assuming ItemCode is the first column and a string
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
        public int DeleteData(string tableName, string conditionColumn, string[] conditionValue)
        {
            int rowsAffected = 0;

            try
            {
                string query = $"DELETE FROM {tableName} WHERE ";
                for (int i = 0; i < conditionValue.Count(); i++)
                {
                    query += $" {conditionColumn} = '{conditionValue[i].Trim()}'";
                    if(i != conditionValue.Count() - 1)
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
        public int SaveDataTable(DataTable dataTable, string TableSql, string[] mappingColName = null, string colID = "Id")
        {
            int id = -1;
            try
            {
                id = GetID(TableSql, colID) + 1;
            }
            catch (Exception ex)
            {

            }
            if (id != -1)
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    row[colID] = id++;
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
                try
                {
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
                                command.Parameters.AddWithValue("@" + column.ColumnName, row[column.ColumnName] ?? DBNull.Value);
                            }

                            command.ExecuteNonQuery();
                        }
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
