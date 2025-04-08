using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Export_FPCA_OK2ship_Auto_System.Services
{
    public static class TDMK_ImageConverter
    {
        public static ImageFormat GetImageFormat(string imagePath)
        {
            try
            {
                using (Image image = Image.FromFile(imagePath))
                {
                    return image.RawFormat;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi: {ex.Message}");
                return null;
            }
        }
        public static string GetImageFormatString(string imagePath)
        {
            ImageFormat format = GetImageFormat(imagePath);
            if (format != null)
            {
                if (format.Equals(ImageFormat.Jpeg))
                    return "JPEG";
                if (format.Equals(ImageFormat.Png))
                    return "PNG";
                if (format.Equals(ImageFormat.Gif))
                    return "GIF";
                if (format.Equals(ImageFormat.Bmp))
                    return "BMP";
                if (format.Equals(ImageFormat.Tiff))
                    return "TIFF";
                if (format.Equals(ImageFormat.Icon))
                    return "ICON";
                if (format.Equals(ImageFormat.Wmf))
                    return "WMF";
                if (format.Equals(ImageFormat.Emf))
                    return "EMF";
                else return "Unknown";
            }
            return "Error or null";
        }
        public static byte[] ImageToByteArray(Image image, ImageFormat format)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                image.Save(ms, format);
                return ms.ToArray();
            }
        }

        public static Image ByteArrayToImage(byte[] byteArray)
        {
            using (MemoryStream ms = new MemoryStream(byteArray))
            {
                Image image = Image.FromStream(ms);
                return image;
            }
        }

        public static string ImageToBase64(Image image, ImageFormat format)
        {
            byte[] imageBytes = ImageToByteArray(image, format);
            return Convert.ToBase64String(imageBytes);
        }

        public static Image Base64ToImage(string base64String)
        {
            byte[] imageBytes = Convert.FromBase64String(base64String);
            return ByteArrayToImage(imageBytes);
        }

        public static byte[] ImageFileToByteArray(string filePath)
        {
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    byte[] buffer = new byte[fs.Length];
                    fs.Read(buffer, 0, (int)fs.Length);
                    return buffer;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading image file: {ex.Message}");
                return null;
            }

        }

        public static void ByteArrayToFile(byte[] byteArray, string filePath)
        {
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(byteArray, 0, byteArray.Length);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing image file: {ex.Message}");
            }
        }
    }
}
