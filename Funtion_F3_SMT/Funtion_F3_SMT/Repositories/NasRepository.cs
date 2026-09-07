
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using IniLibs;
using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Diagnostics;
using OK2SHIP_SMT.Services;

namespace OK2SHIP_SMT.Repositories
{
    public class NasRepository
    {
        IniFile TDMK_init;
        public string _nasAddress = "";
        public NasRepository()
        {
            //config file
            string app_path = System.Windows.Forms.Application.StartupPath;
            string config_path = Path.Combine(app_path.Replace(@"\FPCA OK2SHIP Auto System\VHX-IMADA", ""), "Config.ini");
            TDMK_init = new IniFile(config_path);
            _nasAddress = TDMK_init.Read("NasAddress", "SMT_Config");
            ///////////////////////
            try
            {
                if (FileFolderRepository.checkLocationIsValid(_nasAddress))
                {
                    _nasAddress = $"{_nasAddress}";
                }
                else
                {
                    throw new Exception("Không thể truy cập vào địa chỉ nas!");
                }

            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi kết nối Nas: {ex.Message}");
            }
        }
        public bool IsImageColumn(DataColumn column)
        {
            return column.DataType == typeof(Image);
        }

        public List<KeyValuePair<string, byte[]>> GetImageColumn(DataTable dataTable)
        {
            if (dataTable == null || dataTable.Rows.Count == 0)
            {
                throw new Exception("Datatable rỗng hoặc null");
            }

            if (!dataTable.Columns.Contains("ID"))
            {
                throw new Exception("Không có cột ID trong datatable");
            }



            List<KeyValuePair<string, byte[]>> listImage = new List<KeyValuePair<string, byte[]>>();
            List<string> strings = new List<string>();


            foreach (DataColumn column in dataTable.Columns)
            {
                if (IsImageColumn(column))
                {
                    strings.Add(column.ColumnName);
                }
                if (column.DataType == typeof(byte[]))
                {
                    strings.Add(column.ColumnName);
                }
            }


            foreach (string item in strings)
            {
                dataTable.Columns.Add(item + "$Image");
            }
            foreach (DataRow row in dataTable.Rows)
            {
                foreach (string item in strings)
                {
                    var r = row[item];
                    try
                    {

                        if (row[item] != null)
                        {
                            string name = $"{row["ID"]}-{Guid.NewGuid()}";
                            dataTable.Rows[row.Table.Rows.IndexOf(row)][item + "$Image"] = name;
                            if (dataTable.Columns[item].DataType == typeof(byte[]))
                            {

                                listImage.Add(new KeyValuePair<string, byte[]>(name, (byte[])row[item]));
                            }
                            else
                            {
                                listImage.Add(new KeyValuePair<string, byte[]>(name, TDMK_ImageConverter.ImageToByteArray((Image)row[item], ImageFormat.Jpeg)));
                            }
                        }
                    }
                    catch
                    {

                    }
                }
            }

            foreach (string item in strings)
            {
                dataTable.Columns.Remove(item);
                //dataTable.Columns[item + "$Image"].ColumnName = item;

            }
            return listImage;
        }

        public bool EnsureDirectoryExists(string directoryPath)
        {
            // 1. Kiểm tra sự tồn tại của thư mục
            if (!Directory.Exists(directoryPath))
            {
                try
                {
                    // 2. Nếu chưa tồn tại, sử dụng CreateDirectory để tạo.
                    // Phương thức này có thể tạo toàn bộ chuỗi thư mục nếu cần.
                    Directory.CreateDirectory(directoryPath);
                    return true;
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error init nas folder: {ex.Message}");
                    // Xử lý các lỗi có thể xảy ra (ví dụ: không có quyền, đường dẫn không hợp lệ).
                }
            }
            else
            {
                new DirectoryInfo(directoryPath).Delete(true);
                Directory.CreateDirectory(directoryPath);
                return false;
            }
        }
        public int SaveImage(List<KeyValuePair<string, byte[]>> listIMG, string location)
        {
            int res = 0;
            EnsureDirectoryExists(location);
            foreach (KeyValuePair<string, byte[]> item in listIMG)
            {
                //Image image = ConvertToJpeg(item.Value);

                SaveImageObjectToPath(item.Value, $"{location}\\{item.Key}.jpg");
            }
            return res;
        }
        public int SaveImageObjectToPath(byte[] imageObject, string destinationPath)
        {
            if (imageObject == null)
            {
                throw new ArgumentNullException(nameof(imageObject), "Đối tượng ảnh không được null.");
            }

            try
            {
                // Đảm bảo thư mục đích tồn tại
                string dir = Path.GetDirectoryName(destinationPath);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                // Lưu ảnh dưới định dạng JPEG
                //ImageCodecInfo jpegCodec = GetEncoder(ImageFormat.Jpeg);
                //if (jpegCodec == null)
                //    throw new InvalidOperationException("JPEG encoder not found.");

                //EncoderParameters encoderParams = new EncoderParameters(1);
                //encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 85L);

                //imageObject.Save(destinationPath, jpegCodec, encoderParams);
                File.WriteAllBytes(destinationPath, imageObject);
                return 1;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }
        public void MergeDataTable(DataTable dataTable, string category, string itemCode, string lotNo, string address)
        {
            Dictionary<string, byte[]> images = GetImages(address);
            List<string> columnImage = new List<string>();
            foreach (DataColumn column in dataTable.Columns)
            {
                if (column.ColumnName.Contains($"$Image"))
                {
                    columnImage.Add(column.ColumnName);
                }
            }
            foreach (string item in columnImage)
            {
                string name = item.Replace("$Image", "");
                dataTable.Columns.Add(name, typeof(byte[]));
            }
            foreach (DataRow item in dataTable.Rows)
            {
                foreach (string column in columnImage)
                {
                    if (images.TryGetValue(item[column].ToString(), out byte[] image))
                    {
                        item[$"{column.Replace("$Image", "")}"] = image;
                    }
                    //Debugger.Break();
                }
            }
            foreach (string item in columnImage)
            {
                dataTable.Columns.Remove(item);
            }
        }

        public Dictionary<string, byte[]> GetImages(string path)
        {
            string[] folder = Directory.GetFiles(path);
            Dictionary<string, byte[]> listPicture = new Dictionary<string, byte[]>();
            foreach (string item in folder)
            {
                if (item.Contains(".jpg"))
                {
                    // Debugger.Break();
                    string nameFile = FileFolderRepository.GetFileName(item).Replace(".jpg", "");
                    listPicture.Add(nameFile, File.ReadAllBytes(item));
                }
            }
            return listPicture;
        }
        public string HandleImageDataTable(DataTable dataTable, string category, string itemCode, string lotNo)
        {
            List<KeyValuePair<string, byte[]>> list = GetImageColumn(dataTable);
            string location = $"{_nasAddress}\\{category}\\{itemCode}-{lotNo}\\";
            SaveImage(list, location);
            return location;
        }
        /// <summary>
        /// Chuyển đổi một file ảnh bất kỳ (PNG, BMP, GIF, v.v.) thành JPEG.
        /// </summary>
        /// <param name="inputFilePath">Đường dẫn đến file ảnh đầu vào.</param>
        /// <param name="outputFilePath">Đường dẫn để lưu file JPEG đầu ra.</param>
        /// <param name="quality">Chất lượng nén JPEG (từ 0 đến 100). Mặc định là 85.</param>
        public Image ConvertToJpeg(Image image, long quality = 85L)
        {
            if (image == null)
                throw new ArgumentNullException(nameof(image));

            // Lấy codec cho JPEG
            ImageCodecInfo jpegCodec = GetEncoder(ImageFormat.Jpeg);
            if (jpegCodec == null)
                throw new InvalidOperationException("JPEG encoder not found.");

            // Thiết lập thông số chất lượng
            EncoderParameters encoderParams = new EncoderParameters(1);
            encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);

            using (var ms = new MemoryStream())
            {
                image.Save(ms, jpegCodec, encoderParams);
                ms.Position = 0;
                return Image.FromStream(ms);
            }
        }

        /// <summary>
        /// Helper method: Lấy ImageCodecInfo cho định dạng ảnh đã cho.
        /// </summary>
        private ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return null;
        }

    }
}
