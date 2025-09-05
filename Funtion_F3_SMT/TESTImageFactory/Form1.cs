using OK2SHIP_SMT.Repositories;
using System.Data;
using System.Diagnostics;
using System.Net.NetworkInformation;

namespace TESTImageFactory
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<KeyValuePair<Image, string>> list = FileFolderRepository.ListAllPictureInAFolder($"{textBox1.Text}", ".jpg");
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("FileName", typeof(string));
            dt.Columns.Add("Image", typeof(Image));
            foreach (var item in list)
            {
                DataRow row = dt.NewRow();
                row["ID"] = dt.Rows.Count + 1; // Tạo ID tự động    
                row["FileName"] = item.Value; // Lưu tên file
                row["Image"] = item.Key; // Lưu hình ảnh    
                dt.Rows.Add(row);
            }
            dataGridView1.DataSource = dt;
        }

        private async void button3_Click(object sender, EventArgs e)
        {
            tb_statusAPI.Text = "Checking..";
            tb_statusAPI.ForeColor = Color.DarkOrange;

            if (await CheckingAPI()) // Use 'await' to handle the Task<bool> returned by CheckingAPI  
            {
                tb_statusAPI.Text = "Connected";
                tb_statusAPI.ForeColor = Color.Green;
            }
            else
            {
                tb_statusAPI.Text = "Connection Failed";
                tb_statusAPI.ForeColor = Color.Red;
            }
        }
        private async Task<bool> CheckingAPI()
        {
            try
            {
                string s = await APIRepository.GetProductAsync("https://localhost:7286/HelloWorld");
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Đã xảy ra lỗi: {ex.Message}");
                return false;
            }

        }

        private async void button2_Click(object sender, EventArgs e)
        {
            List<byte[]> images = new List<byte[]>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells["Image"].Value is Image img)
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                        images.Add(ms.ToArray());
                    }
                }
            }
            try
            {

                await APIRepository.CallUploadImageApi("https://localhost:7286/api/Image", images, "SEM&BSE");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}");
            }
        }
    }
}
