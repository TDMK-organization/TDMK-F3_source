using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OK2SHIP_SMT.Services
{
    public static class ConverterService
    {
        /// <summary>
        /// Converts a JSON array into a DataTable structure for easier data manipulation.  
        /// </summary>
        /// <param name="json">The input string containing JSON data to be converted into a DataTable.</param>
        /// <returns>Returns a DataTable populated with data from the JSON array or null if the input is invalid.</returns>
        public static DataTable JsonToDataTable(string json)
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
                Console.WriteLine($"An error occurred: {ex.Message}");
                return null; // Or throw a more specific exception
            }
        }
        /// <summary>
        ///  
        /// </summary>
        /// <param name="datatable"></param>
        /// <returns></returns>
        public static string DataTableToJson(this DataTable table)
        {
            if (table == null)
            {
                return null; // Or throw an exception, depending on your needs.
            }

            return JsonConvert.SerializeObject(table);
        }
    }
}
