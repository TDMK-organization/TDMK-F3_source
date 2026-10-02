using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using AntdUI;
using Export_FPCA_OK2ship_Auto_System.Services;
using OfficeOpenXml;

namespace Export_FPCA_OK2ship_Auto_System.Views
{
    public partial class RunProgess : UserControl
    {
        private DataTable _dataTable = new DataTable();
        private string _location;
        private TDMK_ReportManagerService _service;

        public RunProgess(DataTable dt, string location, TDMK_ReportManagerService service)
        {
            InitializeComponent();
            _service = service;
            _dataTable = dt;
            _location = location;
            InitTable();
        }

        private void InitTable()
        {
            if (!_dataTable.Columns.Contains("Status"))
            {
                _dataTable.Columns.Add("Status", typeof(object));
            }

            foreach (DataRow row in _dataTable.Rows)
            {
                // Lấy một giá trị định danh duy nhất của dòng, ví dụ cột "Sheet Name"
                // Đảm bảo DataTable của bạn có cột này, nếu tên khác hãy sửa lại cho đúng
                string uniqueKey = row["SheetName"].ToString();

                // Truyền uniqueKey vào tham số đầu tiên (Id) của CellButton
                var tagBtn = new AntdUI.CellButton(uniqueKey, "Sẵn sàng", AntdUI.TTypeMini.Primary);

                // (Tuỳ chọn) Nếu bạn muốn lưu thêm dữ liệu ẩn phức tạp hơn, dùng thuộc tính Tag:
                // tagBtn.Tag = row; 

                row["Status"] = tagBtn;
            }

            var columnCollection = new AntdUI.ColumnCollection();

            foreach (DataColumn col in _dataTable.Columns)
            {
                if (col.ColumnName == "Status")
                {
                    columnCollection.Add(new AntdUI.Column("Status", "Trạng thái", AntdUI.ColumnAlign.Center)
                    {
                        Width = "120",
                        Fixed = true
                    });
                }
                else
                {
                    columnCollection.Add(new AntdUI.Column(col.ColumnName, col.ColumnName, AntdUI.ColumnAlign.Left));
                }
            }

            dgv.Columns = columnCollection;
            dgv.DataSource = _dataTable;
        }

        // =================================================================================
        // PHƯƠNG THỨC MỚI: Cập nhật Tag từ bên ngoài dựa vào Tên Sheet (hoặc Key định danh)
        // =================================================================================
        public void UpdateTagStatus(string sheetName, bool isCompleted)
        {
            bool isUpdated = false;

            // Duyệt qua DataTable để tìm dòng tương ứng
            foreach (DataRow row in _dataTable.Rows)
            {
                if (row["SheetName"].ToString() == sheetName) // So khớp Key
                {
                    var cellTag = row["Status"] as AntdUI.CellButton;
                    if (cellTag != null)
                    {
                        if (isCompleted)
                        {
                            cellTag.Text = "Hoàn thành";
                            cellTag.Type = AntdUI.TTypeMini.Success;
                        }
                        else
                        {
                            cellTag.Text = "Sẵn sàng";
                            cellTag.Type = AntdUI.TTypeMini.Primary;
                        }

                        // Nếu cần, bạn có thể thay đổi trạng thái Row
                        row["Status"] = cellTag;
                        isUpdated = true;
                    }

                    break; // Tìm thấy thì dừng vòng lặp
                }
            }

            // Vẽ lại bảng nếu có sự thay đổi
            if (isUpdated)
            {
                dgv.Invalidate();
            }
        }

        public Dictionary<string, string> GetGroupedSelectedSheets()
        {
            var result = new Dictionary<string, string>();

            // Duyệt qua từng dòng trong DataTable
            foreach (DataRow row in _dataTable.Rows)
            {
                // 1. Kiểm tra cột Select (Xử lý an toàn cho cả kiểu dữ liệu bool hoặc chuỗi "True")
                bool isSelected = false;
                var selectValue = row["Select"];

                if (selectValue != null && selectValue != DBNull.Value)
                {
                    if (selectValue is bool b)
                    {
                        isSelected = b;
                    }
                    else if (selectValue.ToString().Trim().Equals("True", StringComparison.OrdinalIgnoreCase))
                    {
                        isSelected = true;
                    }
                }

                // 2. Nếu dòng được chọn, tiến hành gom nhóm
                if (isSelected)
                {
                    string sheetName = row["SheetName"].ToString();
                    string sheetDic = row["SheetDic"].ToString();

                    // Nếu Dictionary đã chứa SheetDic (ví dụ: "Coverpage"), cộng dồn thêm chuỗi
                    if (result.ContainsKey(sheetDic))
                    {
                        result[sheetDic] += ";" + sheetName;
                    }
                    else
                    {
                        // Nếu chưa có, thêm phần tử mới
                        result.Add(sheetDic, sheetName);
                    }
                }
            }

            return result;
        }

        private async void btn_run_Click(object sender, EventArgs e)
        {
            string text = btn_run.Text;
            if (text == "Close")
            {
                Form parentForm = this.FindForm();

                if (parentForm != null)
                {
                    parentForm.Close();
                }
            }

            // Tạm thời vô hiệu hóa nút bấm để tránh người dùng click nhiều lần khi đang chạy
            btn_run.Enabled = false;


            try
            {
                if (!string.IsNullOrEmpty(_location))
                {
                    string destinationFilePath = new ExportProcess().getExportLocation();
                    string extension = Path.GetExtension(_location);
                    Dictionary<string, string> dic = GetGroupedSelectedSheets();

                    // 1. Dùng Path.Combine thay vì nối chuỗi thủ công "\\", đồng thời xác định thư mục đích
                    string targetFolder = Path.Combine(destinationFilePath,
                        $"{_service.__ItemName}_{_service.__Build}_{_service.__ItemName}");

                    // 2. Kiểm tra và TẠO THƯ MỤC nếu nó chưa tồn tại (Rất quan trọng, nếu không sẽ bị lỗi DirectoryNotFound)
                    if (!Directory.Exists(targetFolder))
                    {
                        Directory.CreateDirectory(targetFolder);
                    }

                    foreach (string key in dic.Keys)
                    {
                        string destination = Path.Combine(targetFolder, $"{key}{extension}");
                        List<string> list = dic[key].Split(';').ToList();

                        // 3. Đưa tiến trình xử lý Excel nặng xuống Background Thread bằng Task.Run
                        // Lệnh await sẽ chờ xuất file xong nhưng KHÔNG làm đơ giao diện (UI)
                        await Task.Run(() => { ExportSpecificSheets(_location, list, destination); });

                        // 4. Cập nhật trạng thái Tag (Chạy trên UI Thread do đã await xong Task phía trước)
                        foreach (string item in list)
                        {
                            UpdateTagStatus(item, true);
                            if (_service.__STATUS.TryGetValue(item, out string status))
                            {
                                _service.__STATUS[item] = "Nas";
                            }
                            else
                            {
                                _service.__STATUS.Add(key, $"Nas");
                            }
                        }
                    }

                    _service.Save(false);
                    MessageBox.Show("Hoàn tất quá trình xuất file!", "Thành công", MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                else
                {
                    // Qúa trình cài đặt itemcode lotno
                }
            }
            catch (Exception ex)
            {
                // Hiển thị thông báo nếu có lỗi (file đang mở, lỗi đường dẫn,...)
                MessageBox.Show($"Đã xảy ra lỗi trong quá trình xử lý:\n{ex.Message}", "Lỗi", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                // 5. Đảm bảo nút được bật lại bất kể chương trình chạy thành công hay gặp lỗi
                btn_run.Enabled = true;
                btn_run.Text = "Close";
            }
        }

        public void ExportSpecificSheets(string sourceFilePath, List<string> sheetNames, string destinationFilePath)
        {
            // 1. Kiểm tra file gốc có tồn tại không
            if (!File.Exists(sourceFilePath))
            {
                throw new FileNotFoundException($"Không tìm thấy file gốc tại: {sourceFilePath}");
            }

            // 2. Xóa file đích nếu đã tồn tại
            if (File.Exists(destinationFilePath))
            {
                File.Delete(destinationFilePath);
            }

            // (Bắt buộc cho EPPlus version 5 trở lên) Cấu hình bản quyền
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // 3. Mở file gốc và tạo file mới
            using (var sourcePackage = new ExcelPackage(new FileInfo(sourceFilePath)))
            using (var destPackage = new ExcelPackage())
            {
                foreach (string sheetName in sheetNames)
                {
                    // Kiểm tra sheet có tồn tại trong file gốc không
                    var sourceSheet = sourcePackage.Workbook.Worksheets[sheetName];
                    if (sourceSheet != null)
                    {
                        // Copy toàn bộ dữ liệu & định dạng sang file mới
                        destPackage.Workbook.Worksheets.Add(sheetName, sourceSheet);
                    }
                    else
                    {
                        Console.WriteLine($"Cảnh báo: Không tìm thấy sheet '{sheetName}' trong file gốc.");
                    }
                }

                // 4. Lưu file mới (chỉ lưu khi có ít nhất 1 sheet được copy thành công)
                if (destPackage.Workbook.Worksheets.Count > 0)
                {
                    destPackage.SaveAs(new FileInfo(destinationFilePath));
                }
                else
                {
                    throw new Exception("Không có sheet nào hợp lệ để trích xuất. File mới không được tạo.");
                }
            }
        }
    }
}