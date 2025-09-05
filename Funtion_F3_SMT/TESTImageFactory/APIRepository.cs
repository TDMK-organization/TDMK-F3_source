using Microsoft.VisualBasic.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TESTImageFactory
{
    public class APIRepository
    {
        public async static Task CallUploadImageApi(string apiUrl, List<byte[]> images, string bucket)
        {
            using (HttpClient client = new HttpClient())
            {

                using (var formData = new MultipartFormDataContent())
                {
                    // Thêm chuỗi bucket vào formData
                    formData.Add(new StringContent(bucket), "bucket");

                    // Thêm các mảng byte của ảnh vào formData
                    int fileIndex = 0;
                    foreach (var imageBytes in images)
                    {
                        // Tạo MemoryStream từ mảng byte
                        var stream = new MemoryStream(imageBytes);
                        var streamContent = new StreamContent(stream);

                        // Gán tên file tạm thời
                        string fileName = $"image_{fileIndex}.jpg"; // Bạn có thể tùy chỉnh tên file
                        formData.Add(streamContent, "files", fileName);
                        fileIndex++;
                    }

                    try
                    {
                        HttpResponseMessage response = await client.PostAsync($"{apiUrl}?bucket={bucket}", formData);

                        if (response.IsSuccessStatusCode)
                        {
                            string jsonResult = await response.Content.ReadAsStringAsync();
                            // Deserialize và xử lý kết quả
                            throw new Exception(jsonResult);
                        }
                        else
                        {
                            string errorContent = await response.Content.ReadAsStringAsync();
                            throw new Exception($"Lỗi khi upload: {response.StatusCode} - {response.ReasonPhrase}\n{errorContent}");
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"{ex.Message}");
                    }
                }
            }
        }
        public async static Task<string> GetProductAsync(string path)
        {
            using (HttpClient client = new HttpClient())
            {
                // Địa chỉ API
                string apiUrl = $"{path}";

                try
                {
                    // Gửi yêu cầu GET bất đồng bộ và chờ phản hồi
                    HttpResponseMessage response = await client.GetAsync(apiUrl);
                    // Kiểm tra xem yêu cầu có thành công không (mã 200 OK)
                    if (response.IsSuccessStatusCode)
                    {
                        // Đọc nội dung JSON từ phản hồi và chuyển đổi thành đối tượng List<User>
                        string jsonString = await response.Content.ReadAsStringAsync();
                        return jsonString;
                    }
                    else
                    {
                        throw new Exception($"Lỗi: {response.StatusCode} - {response.ReasonPhrase}");
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Đã xảy ra lỗi khi gọi API: {ex.Message}");
                }
            }
        }
    }
}
