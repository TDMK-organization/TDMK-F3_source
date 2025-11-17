using IniLibs;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OK2SHIP_SMT.Repositories
{
    public static class FileFolderRepository
    {
        public static bool checkLocationIsValid(string location)
        {
            try
            {
                // Check if the location exists and is accessible
                return Directory.Exists(location);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                // Log or handle exceptions if needed
                return false;
            }
        }
        public static string[] GetFileByExtension(string folderPath, string extension)
        {
            // Lấy danh sách các tệp có phần mở rộng cụ thể trong thư mục
            return Directory.GetFiles(folderPath, $"*{extension}");
        }
        public static string GetFileName(string filePath)
        {
            // Lấy tên tệp từ đường dẫn
            return Path.GetFileName(filePath);
        }
        public static string GetFileNameWithoutExtension(string filePath)
        {
            // Lấy tên tệp không có phần mở rộng từ đường dẫn
            return Path.GetFileNameWithoutExtension(filePath);
        }
        public static string GetFolderName(string filePath)
        {
            // Lấy tên thư mục chứa tệp
            return Path.GetFileName(filePath);
        }
        public static IList<string> ListAllFileInFolder(string folderPath, string extension)
        {
            List<string> csvFiles = new List<string>();

            try
            {
                // Ensure the folder exists.
                if (Directory.Exists(folderPath))
                {
                    // Get all files in the folder.
                    string[] files = Directory.GetFiles(folderPath);

                    // Iterate through the files and add CSV files to the list.
                    foreach (string file in files)
                    {
                        if (file.EndsWith(extension, StringComparison.OrdinalIgnoreCase)) // Case-insensitive check
                        {
                            csvFiles.Add(file);
                        }
                    }
                }
                else
                {
                    Console.WriteLine($"Folder not found: {folderPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            return csvFiles;
        }
        
        /// <summary>
        /// Get picture in a folder
        /// </summary>
        /// <param name="locationFolder"></param>
        /// <returns></returns>
        public static List<KeyValuePair<Image, string>> ListAllPictureInAFolder(string locationFolder, string Extension = null)
        {
            List<KeyValuePair<Image, string>> listPicture = new List<KeyValuePair<Image, string>>();
            string[] folder = Directory.GetFiles(locationFolder);
            foreach (string file in folder)
            {
                string[] fileSplit = file.Split('\\');
                string fileName = fileSplit[fileSplit.Length - 1];
                if (Extension == null)
                {
                    if (fileName.ToLower().Contains(".jpg") || fileName.Contains(".png"))
                    {
                        listPicture.Add(new KeyValuePair<Image, string>(Image.FromFile(file), fileName));
                    }
                }
                else
                {
                    if (fileName.Contains(Extension))
                    {
                        listPicture.Add(new KeyValuePair<Image, string>(Image.FromFile(file), fileName));
                    }
                }
            }
            return listPicture;
        }

        /// <summary>
        /// Get all picture in many subfolder in folder
        /// </summary>
        /// <param name="location"></param>
        /// <returns></returns>
        public static Dictionary<string, IList<KeyValuePair<Image, string>>> ListAllPictureInFolder(string location)
        {
            Dictionary<string, IList<KeyValuePair<Image, string>>> listPicture = new Dictionary<string, IList<KeyValuePair<Image, string>>>();
            string[] folders = Directory.GetDirectories(location);
            foreach (string folder in folders)
            {
                listPicture.Add(folder, ListAllPictureInAFolder(folder));
            }
            return listPicture;
        }
        public static DataTable ConvertCsvToDataTable(string filePath, int countRow = -1)
        {
            DataTable dataTable = new DataTable();
            try
            {
                using (StreamReader reader = new StreamReader(filePath))
                {
                    // Đọc dòng đầu tiên để lấy tiêu đề cột
                    string headerLine = reader.ReadLine();
                    string[] headers = headerLine.Split(','); // Giả sử dấu phẩy là dấu phân tách

                    for (int i = 0; i < headers.Count(); i++)
                    {
                        try
                        {
                            dataTable.Columns.Add(headers[i].TrimEnd('"').TrimStart('"'));
                        }
                        catch
                        {
                            dataTable.Columns.Add($"{headers[i - 1]} {headers[i]}");
                        }
                    }

                    // Đọc các dòng dữ liệu còn lại
                    while (!reader.EndOfStream)
                    {
                        string dataLine = reader.ReadLine();
                        string[] values = dataLine.Split(','); // Giả sử dấu phẩy là dấu phân tách
                        for (int i = 0; i < values.Count(); i++)
                        {
                            values[i] = values[i].TrimEnd('"').TrimStart('"');
                        }
                        dataTable.Rows.Add(values);
                    }
                }
            }
            catch (FileNotFoundException)
            {
                Debugger.Break();
                throw new Exception($"Không tìm thấy tệp CSV: {filePath}");
            }
            catch (Exception ex)
            {
                Debugger.Break();
                throw new Exception($"Đã xảy ra lỗi: {ex.Message}");
            }

            return dataTable;
        }
        public static string[] GetSubFolders(string folderPath)
        {
            try
            {
                // Lấy danh sách các thư mục con
                string[] subFolders = Directory.GetDirectories(folderPath);
                return subFolders;
            }
            catch (DirectoryNotFoundException ex)
            {
                Console.WriteLine($"Thư mục không tồn tại: {ex.Message}");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi lấy thư mục con: {ex.Message}");
                return null;
            }
        }
    }
}
