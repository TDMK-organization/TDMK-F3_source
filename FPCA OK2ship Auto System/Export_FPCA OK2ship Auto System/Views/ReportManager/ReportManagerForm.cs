using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using AntdUI;
using Export_FPCA_OK2ship_Auto_System.Services;

namespace Export_FPCA_OK2ship_Auto_System.Views
{
    public partial class ReportManagerForm : AntdUI.Window
    {
        private TDMK_ReportManagerService _service = new TDMK_ReportManagerService("", "", "");
        TAlignFrom align = TAlignFrom.BL;

        public ReportManagerForm()
        {
            InitializeComponent();
            dgv.CellButtonClick += Table_base_CellButtonClick;
            dgv.Enabled = false;
            Reset();
        }

        #region Event

        private void btn_delete_Click(object sender, EventArgs e)
        {
            MessageNotification("Warining", "Chức năng chưa phát triển");
        }

        private void ip_itemname_Leave(object sender, EventArgs e)
        {
            ip_itemname.Text = ip_itemname.Text.Trim();
            if (ip_itemname.Text == _service.__ItemName)
            {
                initService();
            }
        }

        private void ip_build_Leave(object sender, EventArgs e)
        {
            ip_build.Text = ip_build.Text.Trim();
            if (ip_build.Text == _service.__Build)
            {
                initService();
            }
        }


        private void ip_config_Leave(object sender, EventArgs e)
        {
            ip_config.Text = ip_config.Text.Trim();
            if (ip_config.Text == _service.__Config)
            {
                initService();
            }
        }


        private void btn_load_Click(object sender, EventArgs e)
        {
            if (btn_load.Text == "Close")
            {
                CloseAction();
            }
            else if (btn_load.Text == "Load")
            {
                LoadAction();
            }
        }


        private void btn_save_Click(object sender, EventArgs e)
        {
            if (btn_save.Text == "Create New")
            {
                Create();
            }
            else if (btn_save.Text == "Save")
            {
                Save();
            }
        }


        private void Table_base_CellButtonClick(object sender, AntdUI.TableButtonEventArgs e)
        {
            if (e.Column != null && e.Column.Title == "Action" && e.Btn is AntdUI.CellButton btn)
            {
                // Ép kiểu về Class Model thay vì Dictionary
                if (e.Record is ReportManagerModel rowData)
                {
                    string sheetName = rowData.SheetName;

                    if (btn.Text == "Xóa" || btn.Id.StartsWith("delete_"))
                    {
                        // Logic xóa...
                    }
                    else if (btn.Text == "Setup" || btn.Id.StartsWith("setup_"))
                    {
                        SetupForm frm = new SetupForm(_service);
                        frm.ShowDialog();
                        DisplayData();
                    }
                }
            }
        }

        private void btn_check_Click(object sender, EventArgs e)
        {
            CheckStatus();
        }

        #endregion

        #region method

        private void ChangeColumnWidth()
        {
            // Kiểm tra xem bảng đã có cột nào chưa
            if (dgv.Columns != null && dgv.Columns.Count > 0)
            {
                // Thay đổi cột "STT" (index 0) thành 100px
                dgv.Columns[3].Width = "500";


                // Yêu cầu Control vẽ lại để cập nhật giao diện mới nhất
                dgv.Invalidate();
            }
        }

        private async void CheckStatus()
        {
            // 1. Lấy dữ liệu từ bảng (Sử dụng BindingList Model)
            var dataList = dgv.DataSource as System.ComponentModel.BindingList<ReportManagerModel>;
            if (dataList == null || dataList.Count == 0) return;

            try
            {
                // 3. Gom nhóm theo SheetName
                // Chỉ lấy những dòng có SheetName hợp lệ
                var groups = dataList
                    .Where(row => !string.IsNullOrEmpty(row.SheetName))
                    .GroupBy(row =>
                    {
                        // Ưu tiên 0: Nếu DataForm chứa chữ "None" (không phân biệt hoa thường)
                        if (!string.IsNullOrEmpty(row.DataForm) &&
                            row.DataForm.Equals("None", StringComparison.OrdinalIgnoreCase))
                        {
                            return "NONE_GROUP";
                        }

                        // Ưu tiên 1: Nếu DataForm chứa chữ "nas" (không phân biệt hoa thường)
                        if (!string.IsNullOrEmpty(row.DataForm) &&
                            row.DataForm.Equals("nas", StringComparison.OrdinalIgnoreCase))
                        {
                            return "NAS_GROUP"; // Tất cả các dòng này sẽ dùng chung key là "NAS_GROUP"
                        }

                        // Ưu tiên 2: Nếu không phải nas, gom nhóm theo SharedKey như cũ
                        return TDMK_ReportManagerService.GetSharedKey(row.SheetName);
                    })
                    .ToList();

                if (groups.Count == 0) return;

                // Cập nhật trạng thái chờ hiển thị tạm thời
                // Do dùng BindingList và NotifyProperty, giao diện sẽ tự động hiển thị ngay lập tức
                foreach (var row in dataList)
                {
                    if (!string.IsNullOrEmpty(row.SheetName))
                    {
                        row.Status = "Đang kiểm tra...";
                    }
                }

                // 4. Khởi tạo các Task chạy song song
                var tasks = groups.Select(async group =>
                {
                    string sharedKey = group.Key;
                    // 2. Tạo Dictionary Input từ các TextBox của ReportManagerForm
                    Dictionary<string, string> dic = new Dictionary<string, string>();
                    string dtFrom = group.First().DataForm.ToString();
                    string[] valueZ = dtFrom.Split('(', '-', ')');
                    try
                    {
                        dic.Add("ItemName", ip_itemname.Text.Trim());
                        dic.Add("Build", ip_build.Text.Trim());
                        dic.Add("Config", ip_config.Text.Trim());
                        dic.Add("ItemCode", valueZ[1]);
                        dic.Add("LotNo", valueZ[2]);
                        dic.Add("Maker", valueZ[3]);
                    }
                    catch
                    {
                    }

                    var (isSuccess, status, attact) = await Task.Run(() =>
                    {
                        bool success = TDMK_ReportManagerService.checkCategory(sharedKey, dic, out string statusValue,
                            out string valueAttact);
                        return (success, statusValue, valueAttact);
                    });

                    return new { Group = group, IsSuccess = isSuccess, Status = status, Collect = attact };
                });

                // Chờ tất cả truy vấn nền hoàn thành
                var results = await Task.WhenAll(tasks);

                // Ngăn crash nếu user đã tắt Form trong lúc đang chờ Task
                if (this.IsDisposed || this.Disposing) return;

                // Tạm tắt sự kiện vẽ lại để gán dữ liệu hàng loạt cho mượt, không bị giật UI
                dataList.RaiseListChangedEvents = false;

                // 5. Bóc tách dữ liệu và cập nhật ngược lại vào UI
                foreach (var result in results)
                {
                    string sharedKey = result.Group.Key;

                    if (sharedKey == "ORT" || sharedKey == "NAS_GROUP" || sharedKey == "ITEMCODE_LOTNO_CHECK" ||
                        sharedKey == "ONPRODUCT")
                    {
                        if (result.IsSuccess && !string.IsNullOrEmpty(result.Status))
                        {
                            var statusDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                            string[] pairs = result.Status.Split(new[] { '\u200F' },
                                StringSplitOptions.RemoveEmptyEntries);

                            foreach (string pair in pairs)
                            {
                                string[] parts = pair.Split(new[] { '\u200B' }, StringSplitOptions.RemoveEmptyEntries);
                                if (parts.Length == 2)
                                {
                                    statusDict[parts[0].Trim()] = parts[1].Trim();
                                }
                            }

                            foreach (var row in result.Group)
                            {
                                string sheetName = row.SheetName;
                                string lookupKey = switchORT(sheetName);

                                if (statusDict.TryGetValue(lookupKey, out string specificStatus))
                                {
                                    row.Status = specificStatus;
                                }
                                else
                                {
                                    row.Status = "No Data";
                                }
                            }
                        }
                        else
                        {
                            foreach (var row in result.Group)
                            {
                                row.Status = "Lỗi / Không tìm thấy";
                            }
                        }
                    }
                    else
                    {
                        // Các mục tra cứu bình thường (Single Check)
                        foreach (var row in result.Group)
                        {
                            row.Status = result.IsSuccess ? result.Status : "Lỗi / Không tìm thấy";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageNotification("Lỗi", $"Có lỗi xảy ra trong quá trình kiểm tra:\n{ex.Message}", "error");
            }
            finally
            {
                // 6. Khôi phục trạng thái UI
                if (!this.IsDisposed)
                {
                    if (dataList != null)
                    {
                        dataList.RaiseListChangedEvents = true; // Bật lại event
                        dataList.ResetBindings(); // Báo cho UI vẽ lại 1 lần duy nhất toàn bộ bảng
                    }

                    MessageNotification("Thông báo", "Hoàn tất kiểm tra dữ liệu!", "info");
                }
            }
        }


        // Bổ sung thêm hàm switchORT vào form này nếu bạn chưa copy từ form cũ sang
        private string switchORT(string sheetName)
        {
            sheetName = sheetName.ToLower().Replace(" - ict data", "");

            switch (sheetName.ToLower().Trim())
            {
                case "heat soak & bending":
                case "hs & bending":
                    return "heat soak and flex bend";

                case "heat soak and recovery":
                case "heat soak & recovery":
                    return "heat soak";

                case "thermal cycling & bending":
                case "tc & bending":
                    return "thermal cycling and flex bending";

                default:
                    return sheetName;
            }
        }

        private void MessageNotification(string title, string message, string type = "error")
        {
            switch (type)
            {
                case "error":
                    AntdUI.Notification.error(this, title, message, autoClose: 10, align: align);
                    break;
                default:
                    AntdUI.Notification.info(this, title, message, autoClose: 10, align: align);
                    break;
            }
        }

        private void Reset()
        {
            initService();
            DisplayData();
        }

        private void initService()
        {
            string itemName = ip_itemname.Text.Trim();
            string build = ip_build.Text.Trim();
            string config = ip_config.Text.Trim();
            _service = new TDMK_ReportManagerService(itemName, build, config);
        }

        private void DisplayData()
        {
            dgv.Columns = new AntdUI.ColumnCollection
            {
                new AntdUI.Column("STT", "STT") { Width = "60", Align = AntdUI.ColumnAlign.Center },
                new AntdUI.Column("SheetName", "Sheet Name") { Width = "200" },
                new AntdUI.Column("DataForm", "Data Form") { Width = "150" },
                new AntdUI.Column("Status", "Status") { Width = "300" },
                new AntdUI.Column("Action", "Action") { Width = "200", Align = AntdUI.ColumnAlign.Center }
            };
            var dataList = new System.ComponentModel.BindingList<ReportManagerModel>();
            int stt = 1;

            if (_service.listcategory.Count <= 0)
            {
                dataList.Add(new ReportManagerModel
                {
                    STT = "", SheetName = "", DataForm = "", Status = "", Action = null
                });
            }

            foreach (string key in _service.listcategory)
            {
                dataList.Add(new ReportManagerModel
                {
                    STT = (stt++).ToString(),
                    SheetName = key,
                    DataForm = _service.__STATUS[key],
                    Status = "None",
                    Action = new AntdUI.CellLink[]
                    {
                        new AntdUI.CellButton("delete_" + key, "Delete", AntdUI.TTypeMini.Error),
                    }
                });
            }

            dgv.DataSource = dataList;
        }

        public void Create()
        {
            try
            {
                while (true)
                {
                    try
                    {
                        string itemName = ip_itemname.Text.Trim();
                        string build = ip_build.Text.Trim();
                        string config = ip_config.Text.Trim();
                        _service.CreateNew(itemName, build, config);
                        MessageNotification("Create", "Success", "INFOR");
                        break;
                    }
                    catch (DataException ex)
                    {
                        if (MessageBox.Show(ex.Message, "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) ==
                            DialogResult.Yes)
                        {
                            _service.Save(false);
                            MessageNotification("Create", "Success", "INFOR");
                            break;
                        }
                        else
                        {
                            throw new Exception("Cancel Create Data!");
                        }
                    }
                }

                DisplayData();
                btn_save.Text = "Save";
                btn_load.Text = "Close";
                dgv.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageNotification("Create", ex.Message);
            }
        }

        private void CloseAction()
        {
            try
            {
                ip_itemname.Text = "";
                ip_build.Text = "";
                ip_config.Text = "";
                _service.EmptyData();

                dgv.Enabled = false;
                DisplayData();

                btn_save.Text = "Create New";
                btn_load.Text = "Load";
            }
            catch (Exception ex)
            {
                MessageNotification("Create", ex.Message);
            }
        }

        private void LoadAction()
        {
            try
            {
                string itemName = ip_itemname.Text.Trim();
                string build = ip_build.Text.Trim();
                string config = ip_config.Text.Trim();
                _service.Load(itemName, build, config);
                DisplayData();
                dgv.Enabled = true;
                btn_load.Text = "Close";
                btn_save.Text = "Save";
            }
            catch (Exception ex)
            {
                MessageNotification("Load", ex.Message);
            }
        }

        private void Save()
        {
            try
            {
                try
                {
                    _service.Save();
                    MessageNotification("Save", "Success", "INFOR");
                }
                catch (DataException ex)
                {
                    if (MessageBox.Show(ex.Message, "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) ==
                        DialogResult.Yes)
                    {
                        _service.Save(false);
                        MessageNotification("Save", "Success", "INFOR");
                    }
                    else
                    {
                        throw new Exception("Cancel Create Data!");
                    }
                }


                DisplayData();
                btn_save.Text = "Save";
                btn_load.Text = "Close";
                dgv.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageNotification("Save", ex.Message);
            }
        }

        #endregion

        // Biến toàn cục để giữ tín hiệu hủy và trạng thái chạy
        private CancellationTokenSource _cancellationTokenSource;
        private bool _isExporting = false;

        private async void button3_Click(object sender, EventArgs e)
        {
            // ==================================================
            // NẾU ĐANG CHẠY -> HÀNH ĐỘNG HỦY
            // ==================================================
            if (_isExporting)
            {
                if (_cancellationTokenSource != null && !_cancellationTokenSource.IsCancellationRequested)
                {
                    _cancellationTokenSource.Cancel();

                    button3.Text = "Đang hủy...";
                    button3.Enabled = false; // Tạm khóa để không bị bấm hủy liên tục nhiều lần

                    MessageNotification("Đang Hủy", "Hệ thống đang dừng quá trình, vui lòng chờ...", "info");
                }

                return; // Thoát hàm, không chạy code bên dưới
            }

            // ==================================================
            // NẾU CHƯA CHẠY -> HÀNH ĐỘNG BẮT ĐẦU XUẤT
            // ==================================================
            _isExporting = true;

            // Đổi giao diện nút thành Hủy (Bạn có thể đổi màu nút sang Đỏ nếu dùng AntdUI)
            button3.Text = "Hủy (Cancel)";

            // Gọi hàm phụ để KHÓA các Textbox nhưng VẪN ĐỂ bảng dgv hoạt động
            LockControls(true);

            progressExport.Visible = true;
            progressExport.Value = 0f;

            var dataList = dgv.DataSource as System.ComponentModel.BindingList<ReportManagerModel>;
            var progressReporter = new Progress<TDMK_ReportManagerService.ExportProgressReport>(reportInfo =>
            {
                progressExport.Value = reportInfo.PercentComplete / 100f;
                if (!string.IsNullOrEmpty(reportInfo.SheetName) && dataList != null)
                {
                    var rowToUpdate = dataList.FirstOrDefault(x => x.SheetName == reportInfo.SheetName);
                    if (rowToUpdate != null) rowToUpdate.Status = reportInfo.StatusResult;
                }
            });

            _cancellationTokenSource = new CancellationTokenSource();

            try
            {
                List<string> listCateogry = dataList?.Select(x => x.SheetName).ToList() ?? new List<string>();

                await _service.ExportAll(listCateogry, progressReporter, token: _cancellationTokenSource.Token);

                MessageNotification("Thông báo", "Xuất file hoàn tất!", "info");
            }
            catch (OperationCanceledException)
            {
                MessageNotification("Đã Hủy", "Quá trình xuất file đã bị hủy bỏ.", "warning");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
            finally
            {
                // Khôi phục lại trạng thái ban đầu sau khi Xong hoặc Hủy
                _isExporting = false;
                progressExport.Value = 1f;

                button3.Text = "Export All"; // Đổi lại tên nút
                button3.Enabled = true; // Mở khóa lại nút (nếu đã bị khóa lúc bấm Hủy)

                // Mở khóa lại các Textbox
                LockControls(false);

                if (_cancellationTokenSource != null)
                {
                    _cancellationTokenSource.Dispose();
                    _cancellationTokenSource = null;
                }
            }
        }

// Hàm phụ trợ để khóa/mở khóa giao diện thông minh
        private void LockControls(bool isLocked)
        {
            // Khóa các ô nhập liệu và các nút khác để người dùng không sửa giữa chừng
            ip_itemname.Enabled = !isLocked;
            ip_build.Enabled = !isLocked;
            ip_config.Enabled = !isLocked;
            btn_save.Enabled = !isLocked;
            btn_load.Enabled = !isLocked;

            // ĐẢM BẢO bảng dgv LUÔN MỞ để có thể cuộn (scroll)
            dgv.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SetupForm frm = new SetupForm(_service);
            frm.ShowDialog();
            DisplayData();
        }

        private void btn_checknas_Click(object sender, EventArgs e)
        {
            try
            {
                MessageBox.Show( _service.checkNas(), "Thông báo") ;
            }
            catch (Exception ex)
            {
                if (MessageBox.Show($"{ex.Message} Bạn có muốn cập nhật", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                     _service.updateNas();
                     DisplayData();
                }
            }
        }
    }

    public class ReportManagerModel : AntdUI.NotifyProperty
    {
        public string STT { get; set; }
        public string SheetName { get; set; }
        public string DataForm { get; set; }

        private string _status;

        public string Status
        {
            get => _status;
            set
            {
                _status = value;
                OnPropertyChanged("Status"); // Tự động báo UI vẽ lại ô này
            }
        }

        public AntdUI.CellLink[] Action { get; set; }
    }
}