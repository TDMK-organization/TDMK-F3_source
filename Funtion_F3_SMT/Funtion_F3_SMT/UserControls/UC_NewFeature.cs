using OfficeOpenXml;
using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar;

namespace OK2SHIP_SMT.UserControls
{
    public partial class UC_NewFeature : UserControl
    {
        public UC_NewFeature()
        {
            InitializeComponent();
        }

        private void tb_locationFolder_DoubleClick(object sender, EventArgs e)
        {
            OpenFolderDialog();
        }
        private void btn_runProcess_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }
                row.Cells["Status"].Value = RenameFile(row.Cells["File Path"].Value.ToString(), $"{Path.GetDirectoryName(row.Cells["File Path"].Value.ToString()) + '\\' + row.Cells["File Name New"].Value.ToString()}");
                //object firstCellValue = row.Cells[0].Value;
                //dataTable.Columns.Add("File Name New", typeof(string));
                //dataTable.Columns.Add("File Name", typeof(string));
                //dataTable.Columns.Add("File Path", typeof(string));
                //dataTable.Columns.Add("Status", typeof(string));
                //dataTable.Columns.Add("Pick", typeof(bool));

            }
        }
        public string RenameFolder(string oldFolderPath, string newFolderName)
        {
            try
            {

                string parentDirectory = Path.GetDirectoryName(oldFolderPath);


                string newFolderPath = Path.Combine(parentDirectory, newFolderName);

                if (!Directory.Exists(oldFolderPath))
                {
                    return $"Lỗi: Thư mục cũ không tồn tại tại đường dẫn: {oldFolderPath}";
                }

                // Kiểm tra xem thư mục mới đã tồn tại chưa để tránh lỗi
                if (Directory.Exists(newFolderPath))
                {
                    return $"Lỗi: Thư mục với tên mới '{newFolderName}' đã tồn tại.";
                }

                // 3. Thực hiện việc đổi tên/di chuyển
                Directory.Move(oldFolderPath, newFolderPath);
                return "Success";
            }
            catch (Exception ex)
            {
                return "Đã xảy ra lỗi khi đổi tên thư mục: {ex.Message}";
            }
        }
        public string RenameFile(string oldFilePath, string newFileName)
        {
            try
            {
                // 1. Lấy đường dẫn thư mục cha từ đường dẫn tệp cũ
                // Ví dụ: từ C:\Docs\OldReport.txt -> C:\Docs
                string parentDirectory = Path.GetDirectoryName(oldFilePath);

                // 2. Tạo đường dẫn tệp mới bằng cách kết hợp thư mục cha và tên tệp mới
                // Ví dụ: C:\Docs\NewReport.txt
                string newFilePath = Path.Combine(parentDirectory, newFileName);

                // Kiểm tra xem tệp cũ có tồn tại không
                if (!File.Exists(oldFilePath))
                {
                    return $"Lỗi: Tệp cũ không tồn tại tại đường dẫn: {oldFilePath}";
                }

                // Kiểm tra xem tệp với tên mới đã tồn tại chưa để tránh ghi đè ngoài ý muốn
                if (File.Exists(newFilePath))
                {
                    return $"Lỗi: Tệp với tên mới '{newFileName}' đã tồn tại.";
                }

                // 3. Thực hiện việc đổi tên/di chuyển
                // Lưu ý: Tên mới phải bao gồm cả phần mở rộng (ví dụ: .txt, .pdf)
                File.Move(oldFilePath, newFilePath);
                return "Success";
            }
            catch (Exception ex)
            {
                return $"Đã xảy ra lỗi khi đổi tên tệp: {ex.Message}";
            }
        }
        private void btnGetData_Click(object sender, EventArgs e)
        {
            scanFolder(tb_locationFolder.Text.Trim());
        }

        private void scanFolder(string location)
        {
            location = location.Trim();
            if (string.IsNullOrEmpty(location))
            {
                MessageBox.Show("Please select a folder to scan.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (!System.IO.Directory.Exists(location))
            {
                MessageBox.Show("The specified folder does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            List<string> list = list_extension.CheckedItems.Cast<string>().ToList();
            List<string> xlsxFiles = GetAllXlsxFiles(location, list);
            if (xlsxFiles.Count <= 0)
            {
                MessageBox.Show("No files found in the selected folder.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("File Name New", typeof(string));
            dataTable.Columns.Add("File Name", typeof(string));
            dataTable.Columns.Add("File Path", typeof(string));
            dataTable.Columns.Add("Status", typeof(string));
            dataTable.Columns.Add("Pick", typeof(bool));

            foreach (string item in xlsxFiles)
            {
                dataTable.Rows.Add(Path.GetFileName(item), Path.GetFileName(item), item, "Not Processed");
            }
            dataGridView.DataSource = dataTable;
            dataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
        }

        private bool solveFileProcess(string locationFile)
        {
            try
            {
                ExportProcess exportProcess = new ExportProcess();
                using (ExcelPackage package = new ExcelPackage(locationFile))
                {
                    using (ExcelWorksheet worksheet = package.Workbook.Worksheets[0])
                    {
                        string[] sheet = new[] { "ĐK thêm dòng", "Vlo", "Tháng làm hố sơ", "Công ty mua", "STT in BKLS", "STT lô làm hs" };
                        IDictionary<string, string> dic = ExportProcess.FindAddressByText(worksheet, sheet);
                        if (dic.TryGetValue("STT lô làm hs", out string addZ))
                        {

                        }
                        else
                        {
                            MessageBox.Show("Không tìm thấy cột STT lô làm hs");
                        }
                        if (dic.TryGetValue("ĐK thêm dòng", out string address))
                        {
                            address = ExportProcess.AddRow(address, 1);
                            while (true)
                            {
                                string value = worksheet.Cells[address].Text;
                                if (string.IsNullOrEmpty(value))
                                {
                                    break;
                                }
                                if (int.TryParse(value, out int colAdd))
                                {
                                    worksheet.InsertRow(worksheet.Cells[address].End.Row + 1, colAdd - 1);
                                    string addRoot = $"{address}:{worksheet.Cells[worksheet.Cells[address].End.Row, worksheet.Cells[addZ].End.Column]}";
                                    address = ExportProcess.AddRow(address, 1);
                                    for (int i = 1; i < colAdd; i++)
                                    {
                                        exportProcess.CopyAndInsert(worksheet, addRoot, ref address);
                                    }
                                }
                            }
                            FileInfo file = new FileInfo($"{Path.GetDirectoryName(locationFile)}\\{Path.GetFileNameWithoutExtension(locationFile)}-PROCESS.xlsx");
                            package.SaveAs(file);
                            return true;
                        }
                        else
                        {
                            MessageBox.Show("Không tìm thấy địa chỉ ĐK thêm dòng", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return false;
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Debugger.Break();
                return false;
            }
        }

        public static List<string> GetAllXlsxFiles(string folderPath, List<string> list)
        {
            List<string> result = new List<string>();
            foreach (string item in list)
            {
                if (Directory.Exists(folderPath))
                {
                    SearchOption option = false ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                    result.AddRange(Directory.GetFiles(folderPath, $"*{item}", option));
                }
            }
            return result;
        }
        private void OpenFolderDialog()
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.Description = "Select a folder";
            folderBrowserDialog.ShowNewFolderButton = true;
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                tb_locationFolder.Text = folderBrowserDialog.SelectedPath;
            }

        }
        public List<string> GetLowercaseUniqueFileExtensionsInFolder(string folderPath)
        {
            // 1. Kiểm tra sự tồn tại của thư mục
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Lỗi: Thư mục '{folderPath}' không tồn tại.");
                return new List<string>();
            }

            // 2. Lấy tất cả đường dẫn tệp trong thư mục
            //    Sử dụng SearchOption.TopDirectoryOnly để tìm kiếm trong thư mục gốc.
            //    Dùng SearchOption.AllDirectories nếu muốn tìm kiếm đệ quy.
            string[] filePaths = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);

            // 3. Sử dụng LINQ để xử lý, lấy đuôi mở rộng và chuyển sang chữ thường
            List<string> extensions = filePaths
                // Lấy đuôi mở rộng của từng tệp (ví dụ: ".TXT", ".JpG")
                .Select(filePath => Path.GetExtension(filePath))

                // Lọc ra các đuôi mở rộng trống (tệp không có đuôi)
                .Where(extension => !string.IsNullOrEmpty(extension))

                // CHUYỂN TẤT CẢ SANG CHỮ THƯỜNG để đảm bảo không phân biệt hoa thường
                .Select(extension => extension.ToLowerInvariant())

                // Lấy các đuôi mở rộng DUY NHẤT (ví dụ: .jpg và .JPG chỉ còn lại .jpg)
                .Distinct()

                // Chuyển kết quả sang List<string>
                .ToList();

            return extensions;
        }
        private void tb_locationFolder_TextChanged(object sender, EventArgs e)
        {
            list_extension.Items.Clear();
            list_extension.Items.AddRange(GetLowercaseUniqueFileExtensionsInFolder(tb_locationFolder.Text).ToArray());
        }

        private void btn_Custom_Click(object sender, EventArgs e)
        {
            try
            {

                string t = tb_custom.Text;
                if (string.IsNullOrEmpty(t))
                {
                    MessageBox.Show("Plse input custom text");
                    return;
                }
                List<DataGridViewRow> rows = GetPickRow();
                foreach (DataGridViewRow row in rows)
                {
                    string v = t.Replace("*", Path.GetFileNameWithoutExtension(row.Cells["File Name New"].Value.ToString()));
                    row.Cells["File Name New"].Value = $"{v}{Path.GetExtension(row.Cells["File Name New"].Value.ToString())}";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}");
            }
        }

        private List<DataGridViewRow> GetPickRow(string columnName = "Pick")
        {
            List<DataGridViewRow> pickedRows = new List<DataGridViewRow>();

            // 1. Kiểm tra xem DataGridView có hàng nào hay không
            if (dataGridView.Rows.Count > 0)
            {
                // 2. Lặp qua tất cả các hàng trong DataGridView
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    // 3. Đảm bảo hàng không phải là hàng mới đang được chỉnh sửa (New Row/Template)
                    if (row.IsNewRow)
                    {
                        continue;
                    }

                    // 4. Kiểm tra giá trị của ô trong cột "pick"
                    // Lấy giá trị của ô
                    object cellValue = row.Cells[columnName].Value;

                    // Kiểm tra nếu giá trị là true (phổ biến cho DataGridViewCheckBoxColumn)
                    // Cần xử lý giá trị null và kiểu dữ liệu
                    if (cellValue != null && cellValue is bool && (bool)cellValue == true)
                    {
                        // Hoặc, nếu dữ liệu của bạn là kiểu số 1/0 hoặc string "True":
                        // if (cellValue != null && cellValue.ToString() == "True") 

                        // Thêm hàng thỏa mãn điều kiện vào danh sách
                        pickedRows.Add(row);
                    }
                }

                if (pickedRows.Count <= 0)
                {
                    throw new Exception("Pick some row");
                }
            }
            return pickedRows;
        }
        private void btn_oneToN_Click(object sender, EventArgs e)
        {
            try
            {

                List<DataGridViewRow> rows = GetPickRow();
                int i = 1;
                foreach (DataGridViewRow row in rows)
                {
                    row.Cells["File Name New"].Value = $"{i++}{Path.GetExtension(row.Cells["File Name"].Value.ToString())}";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}");
            }
        }

        private void btn_pick_Click(object sender, EventArgs e)
        {
            if (dataGridView.SelectedCells.Count > 0)
            {
                foreach (DataGridViewCell selectedCell in dataGridView.SelectedCells)
                {

                    int rowIndex = selectedCell.RowIndex;
                    dataGridView.Rows[rowIndex].Cells["Pick"].Value = true;
                }
            }
            else
            {
                MessageBox.Show("No rows are currently selected.");
            }
        }

        private void btn_unselect_Click(object sender, EventArgs e)
        {
            if (dataGridView.SelectedCells.Count > 0)
            {
                foreach (DataGridViewCell selectedCell in dataGridView.SelectedCells)
                {

                    int rowIndex = selectedCell.RowIndex;
                    dataGridView.Rows[rowIndex].Cells["Pick"].Value = false;
                }
            }
            else
            {
                MessageBox.Show("No rows are currently selected.");
            }
        }
    }
}
