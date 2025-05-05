using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OK2SHIP_SMT.Services
{

    public static class ConverterService
    {

        public static int GetNumberFromString(string str)
        {
            Match match = Regex.Match(str, @"\d+"); // Matches one or more digits

            if (match.Success)
            {
                if (int.TryParse(match.Value, out int number))
                {
                    return number; // Output: 5
                }
                else
                {
                    throw new Exception("Invalid number format");
                }

            }
            else
            {
                throw new Exception("No number found.");
            }
        }
        /// <summary>
        /// Converts a JSON array into a DataTable structure for easier data manipulation.  
        /// </summary>
        /// <param name="json">The input string containing JSON data to be converted into a DataTable.</param>
        /// <returns>Returns a DataTable populated with data from the JSON array or null if the input is invalid.</returns>
        public static DataTable JsonToDataTable(string json, bool convertImage = false)
        {
            if (string.IsNullOrEmpty(json))
            {
                return null; // Or throw an exception
            }

            try
            {
                DataTable dataTable = new DataTable();
                JArray jArray = JArray.Parse(json);

                if (jArray.Count == 0)
                {
                    return dataTable; // Return empty DataTable if JSON array is empty
                }

                // Create columns based on the first object in the array
                JObject firstObject = (JObject)jArray[0];
                foreach (JProperty property in firstObject.Properties())
                {
                    dataTable.Columns.Add(property.Name, typeof(string)); // Default to string, refine if needed
                }

                // Populate the DataTable with data from the JSON array
                foreach (JObject jObject in jArray)
                {
                    DataRow dataRow = dataTable.NewRow();
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        dataRow[column.ColumnName] = jObject[column.ColumnName]?.ToString(); // Handle nulls
                    }
                    dataTable.Rows.Add(dataRow);
                }
                if (convertImage)
                {
                    DataTable dataResult = new DataTable();
                    Dictionary<string, string> dic = new Dictionary<string, string>();
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        if (column.ColumnName.Contains("$CONVERTERIMAGE"))
                        {
                            string nameNew = column.ColumnName.Trim().Split('$')[0];
                            dataResult.Columns.Add(nameNew, typeof(byte[]));
                            dic.Add(nameNew, column.ColumnName);
                        }
                        else
                        {
                            dataResult.Columns.Add(column.ColumnName, column.GetType());
                        }
                    }
                    foreach (DataRow row in dataTable.Rows)
                    {
                        DataRow newRow = dataResult.NewRow();
                        foreach (DataColumn column in dataResult.Columns)
                        {
                            if (dic.TryGetValue(column.ColumnName, out string nameCol))
                            {
                                newRow[column.ColumnName] = TDMK_ImageConverter.ByteArrayToImage(GetBytesUTF8(row[nameCol].ToString()));
                            }
                            else
                            {
                                newRow[column.ColumnName] = row[column.ColumnName];
                            }
                        }
                        dataResult.Rows.Add(newRow);
                    }
                    return dataResult;
                }
                return dataTable;
            }
            catch (JsonReaderException ex)
            {
                // Handle invalid JSON format
                Console.WriteLine($"Error parsing JSON: {ex.Message}");
                return null; // Or throw a more specific exception
            }
            catch (Exception ex)
            {
                // Handle other exceptions (e.g., incorrect data types)
                throw new Exception($"An error occurred: {ex.Message}");
            }
        }
        /// <summary>
        ///  
        /// </summary>
        /// <param name="datatable"></param>
        /// <returns></returns>
        public static string DataTableToJson(this DataTable table, bool convertImage = false)
        {
            if (convertImage)
            {
                DataTable dt = DataTableConvertImageToString(table);
                return JsonConvert.SerializeObject(dt);
            }
            if (table == null)
            {
                return null; // Or throw an exception, depending on your needs.
            }

            return JsonConvert.SerializeObject(table);
        }
        public static byte[] GetBytesUTF8(this string str)
        {
            if (string.IsNullOrEmpty(str))
            {
                return new byte[0];
            }
            return Encoding.UTF8.GetBytes(str);
        }
        // Chuyển đổi string thành byte[] sử dụng Encoding ASCII
        public static byte[] GetBytesASCII(this string str)
        {
            if (string.IsNullOrEmpty(str))
            {
                return new byte[0];
            }
            return Encoding.ASCII.GetBytes(str);
        }
        public static string GetStringUTF8(this byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
            {
                return string.Empty;
            }
            return Encoding.UTF8.GetString(bytes);
        }
        // Chuyển đổi byte[] thành string sử dụng Encoding ASCII
        public static string GetStringASCII(this byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
            {
                return string.Empty;
            }
            return Encoding.ASCII.GetString(bytes);
        }
        //public static DataTable StringConvertImageToDataTable(DataTable table)
        public static DataTable DataTableConvertImageToString(DataTable table)
        {
            DataTable dataTable = new DataTable();
            foreach (DataColumn column in table.Columns)
            {
                if (column.DataType == typeof(Image))
                {

                    dataTable.Columns.Add($"{column.ColumnName}$CONVERTERIMAGE", typeof(string));

                }
                else
                {
                    dataTable.Columns.Add($"{column.ColumnName}", typeof(string));

                }
            }
            foreach (DataRow row in table.Rows)
            {
                DataRow rowz = dataTable.NewRow();
                foreach (DataColumn column in dataTable.Columns)
                {
                    string namez = column.ColumnName.ToString();
                    if (namez.Split('$').Count() >= 2)
                    {
                        namez = namez.Split('$')[0];
                    }
                    byte[] image = TDMK_ImageConverter.ImageToByteArray((Image)row[$"{namez}"], ImageFormat.Jpeg);
                    rowz[$"{column.ColumnName}"] = GetStringUTF8(image);
                }
                dataTable.Rows.Add(rowz);
            }
            return dataTable;
        }
    }
}
