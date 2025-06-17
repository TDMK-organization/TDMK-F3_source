using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.Services
{
    public class ArtistFramerService : IDisposable
    {

        public Image Image { get; set; }
        private string Location { get; set; }
        public ArtistFramerService(Image image)
        {
            Image = image;
            //Location = System.Windows.Forms.Application.StartupPath + $"\\PICTURE\\{Guid.NewGuid()}.png";

            //SaveImage();
            OpenMainFormWithAddress();

        }

        private void OpenMainFormWithAddress()
        {
            //using (SPixel.MainForm mainForm = new SPixel.MainForm(Location))
            //{
            //    mainForm.ShowDialog();
            //    Image = mainForm.bitmap;
            //    //Image = LoadImageFromFile(Location);
            //}
        }



        private void SaveImage()
        {
            SaveImageToFile(Image, Location, ImageFormat.Png);
        }

        /// <summary>
        /// Lưu một đối tượng Image vào một file với định dạng cụ thể.
        /// </summary>
        /// <param name="imageToSave">Đối tượng Image cần lưu.</param>
        /// <param name="filePath">Đường dẫn đầy đủ nơi bạn muốn lưu file ảnh.</param>
        /// <param name="format">Định dạng ảnh (ví dụ: Jpeg, Png, Bmp).</param>
        private void SaveImageToFile(Image imageToSave, string filePath, ImageFormat format)
        {
            if (imageToSave == null)
            {
                throw new Exception("Đối tượng ảnh không được rỗng.");
            }


            // Đảm bảo thư mục tồn tại
            string directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            imageToSave.Save(filePath, format);

        }

        public void Dispose()
        {
            return;
        }
    }
}
