using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using AntdUI;
using OK2SHIP_SMT.Services;

namespace OK2SHIP_SMT.UserControls.Build2Ship
{
    public partial class UC_Mainscreen : UserControl
    {
        private Build2ShipService _service = new Build2ShipService();

        public UC_Mainscreen()
        {
            InitializeComponent();
            LoadTableData();
            display_data();
        }

        private void tb_logfile_DoubleClick(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                // Cấu hình lọc: Chỉ hiển thị các file Excel
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.Title = "Chọn file dữ liệu Excel";

                // Cho phép chọn nhiều file (nếu muốn)
                openFileDialog.Multiselect = false;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    tb_logfile.Text = filePath;
                }
            }
        }

        private void LoadTableData()
        {
            // 1. Định nghĩa các cột cho Table
            // Lưu ý: Cú pháp khởi tạo cột có thể thay đổi nhẹ tùy theo phiên bản AntdUI bạn đang sử dụng.
            table.Columns = new AntdUI.ColumnCollection
            {
                new AntdUI.Column("STT", "STT") { Width = "60", Align = AntdUI.ColumnAlign.Center },
                new AntdUI.Column("SheetName", "Tên Sheet") { Width = "150" },
                new AntdUI.Column("Status", "Trạng thái") { Width = "120" },
                new AntdUI.Column("Action", "Thao tác") { Width = "100", Align = AntdUI.ColumnAlign.Center }
            };
            table.CellButtonClick += Table_base_CellButtonClick;
        }

        private void Table_base_CellButtonClick(object sender, AntdUI.TableButtonEventArgs e)
        {
            if (e.Column != null && e.Column.Title == "Thao tác" && e.Btn is AntdUI.CellButton btn)
            {
                var rowData = e.Record as Dictionary<string, object>;
                if (rowData == null) return;

                string sheetName = rowData["SheetName"]?.ToString();

                // Xử lý nút Xóa dựa vào ID hoặc Text của nút
                if (btn.Text == "Xóa" || btn.Id.StartsWith("delete_"))
                {
                    var confirm = MessageBox.Show($"Bạn có chắc chắn muốn xóa sheet '{sheetName}'?", "Xác nhận",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (confirm == DialogResult.Yes)
                    {
                        if (_service._DATA.ContainsKey(sheetName))
                        {
                            _service._DATA.Remove(sheetName);
                        }

                        display_data(); // Load lại bảng
                    }
                }
                // Xử lý nút Upload
                else if (btn.Text == "Upload" || btn.Id.StartsWith("upload_"))
                {
                    try
                    {

                        _service.Upload(tb_logfile.Text, btn.Id, tb_itemCode.Text.Trim(), tb_lotNo.Text.Trim());
                        MessageBox.Show("Cài đặt thành công");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }


        private void display_data()
        {
            List<Dictionary<string, object>> dataList = new List<Dictionary<string, object>>();
            int stt = 1;
            foreach (string key in _service._DATA.Keys)
            {
                dataList.Add(
                    new Dictionary<string, object>
                    {
                        { "STT", stt++ },
                        { "SheetName", key },
                        { "Status", _service._DATA[key] },
                        {
                            "Action", new CellLink[]
                            {
                                // Đặt ID rõ ràng cho từng nút để nhận diện khi click
                                new CellButton("upload_" + key, "Upload", TTypeMini.Primary),
                                new CellButton("delete_" + key, "Delete", TTypeMini.Primary),
                            }
                        }
                    }
                );
            }

            table.DataSource = dataList;
        }

        private void button_loading_Click(object sender, EventArgs e)
        {
            try
            {
                string location = tb_logfile.Text.Trim();
                _service.GetStatus(location);
                display_data();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        TAlignFrom align = TAlignFrom.Top;


        private void tb_logfile_TextChanged(object sender, EventArgs e)
        {
        }

        private void btn_loadData_Click(object sender, EventArgs e)
        {
            try
            {
                string itemCode = tb_itemCode.Text;
                string lotNo = tb_lotNo.Text;
                _service.Load(itemCode, lotNo);
                display_data();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}