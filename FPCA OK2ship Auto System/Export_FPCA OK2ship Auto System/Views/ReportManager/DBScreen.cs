using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using System.Data;
using System.Linq;
using Export_FPCA_OK2ship_Auto_System.Services;

namespace Export_FPCA_OK2ship_Auto_System.Views
{
    public partial class DBScreen : UserControl
    {
        public DBScreen()
        {
            InitializeComponent();

            // Đăng ký sự kiện click nút trong bảng
            dgv.CellButtonClick += Dgv_CellButtonClick;

            initData();
        }

        private void Dgv_CellButtonClick(object sender, AntdUI.TableButtonEventArgs e)
        {
            // Lấy đối tượng (dòng) mà người dùng vừa click vào
            if (e.Record is ReportCategoryModel rowModel)
            {
                // Kiểm tra xem có đúng là click vào cột Action không
                if (e.Column.Key == "Action")
                {
                    // Đảo ngược trạng thái Lock <-> Unlock
                    rowModel.IsLocked = !rowModel.IsLocked;
                }
            }
        }

        private void initData()
        {
            // 1. Khai báo cột cho bảng AntUI
            dgv.Columns = new AntdUI.ColumnCollection()
            {
                // AntUI sẽ tự động vẽ một ô checkbox đơn giản trên Header.
                // Khi người dùng click vào ô này, toàn bộ các dòng bên dưới sẽ tự động được Tick/Bỏ tick.
                new AntdUI.ColumnCheck("Select") { Fixed = true },

                new AntdUI.Column("SheetName", "Sheet Name", AntdUI.ColumnAlign.Left),
                new AntdUI.Column("CollectFrom", "Collect From", AntdUI.ColumnAlign.Left),
                new AntdUI.Column("Status", "Status", AntdUI.ColumnAlign.Center),
                new AntdUI.Column("Action", "Action", AntdUI.ColumnAlign.Center) { Width = "100" }
            };

            // 2. Dùng BindingList thay thế DataTable
            var dataList = new BindingList<ReportCategoryModel>();

            foreach (string item in TDMK_ReportManagerService.categories)
            {
                string cllec = TDMK_ReportManagerService.getCollectFrom(item);
                if (!string.IsNullOrEmpty(cllec))
                {
                    dataList.Add(new ReportCategoryModel
                    {
                        Select = false, // Mặc định bỏ tick
                        IsLocked = true, // Mặc định nút ở chế độ Lock
                        SheetName = item,
                        CollectFrom = cllec,
                        Status = "Sẵn sàng",
                    });
                }
            }

            // Gán danh sách đối tượng thẳng vào bảng
            dgv.DataSource = dataList;
        }

        public DataTable GetTableData()
        {
            // 1. Lấy DataSource của bảng và ép kiểu về BindingList<SheetRowData>
            var dataList = dgv.DataSource as System.ComponentModel.BindingList<ReportCategoryModel>;

            // Kiểm tra xem bảng có dữ liệu hay không
            if (dataList == null || dataList.Count == 0)
            {
                return new DataTable();
            }

            DataTable dt = new DataTable();
            dt.Columns.Add("Select", typeof(bool));
            dt.Columns.Add("SheetName", typeof(string));
            dt.Columns.Add("SheetDic", typeof(string));

            // 2. Duyệt qua từng dòng dữ liệu trong danh sách
            foreach (var row in dataList)
            {
                bool select = row.Select;
                if (select)
                {
                    DataRow dr = dt.NewRow();
                    dr["Select"] =select;
                    dr["SheetName"] = row.SheetName;
                    dr["SheetDic"] = row.CollectFrom;
                    dt.Rows.Add(dr);
                }
            }

            return dt;
        }

        private void InitTableColumns()
        {
            dgv.Columns = new AntdUI.ColumnCollection()
            {
                new AntdUI.ColumnCheck("Selected") { Fixed = true },
                new AntdUI.Column("Name", "Sheet Name", AntdUI.ColumnAlign.Center)
                {
                    Width = "120"
                },
                new AntdUI.Column("Action", "Action", AntdUI.ColumnAlign.Center)
            };
        }

        private async void btn_load_Click(object sender, EventArgs e)
        {
            // 1. Kiểm tra đầu vào TRƯỚC khi khóa nút (Fix P0.1)
            var dataList = dgv.DataSource as BindingList<ReportCategoryModel>;
            if (dataList == null || dataList.Count == 0) return;

            btn_load.Enabled = false;

            try
            {
                // 2. Tạo Dictionary Read-only truyền vào Task (Fix P1.2)
                Dictionary<string, string> dic = new Dictionary<string, string>
                {
                    { "ItemCode", ip_itemCode.Text.Trim() },
                    { "LotNo", ip_lotno.Text.Trim() },
                    { "Maker", ip_maker.Text.Trim() }
                };

                // 3. LỌC VÀ Gom nhóm (Fix P2.1, P1.2)
                // Dùng .Where(row => row.IsLocked) để CHỈ LẤY những dòng đang khóa (nút hiện chữ "Lock")
                // Bỏ qua hoàn toàn những dòng có nút "Unlock"
                var groups = dataList
                    .Where(row => row.IsLocked)
                    .GroupBy(row => TDMK_ReportManagerService.GetSharedKey(row.SheetName))
                    .ToList();

                // Nếu tất cả các dòng đều bị Unlock (không có nhóm nào để chạy), thì thoát luôn
                if (groups.Count == 0)
                {
                    MessageBox.Show("Không có hạng mục nào cần kiểm tra (tất cả đều đang Unlock)!", "Thông báo",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Khởi tạo các Task chạy song song
                var tasks = groups.Select(async group =>
                {
                    string sharedKey = group.Key;

                    // Đẩy tác vụ truy vấn DB xuống background thread
                    var (isSuccess, status) = await Task.Run(() =>
                    {
                        bool success = TDMK_ReportManagerService.checkCategory(sharedKey, dic, out string statusValue);
                        return (success, statusValue);
                    });

                    return new { Group = group, IsSuccess = isSuccess, Status = status };
                });

                // Chờ tất cả DB queries hoàn thành trên background
                var results = await Task.WhenAll(tasks);

                // 4. Ngăn crash nếu user đã tắt Form (Fix P0.3)
                if (this.IsDisposed || this.Disposing) return;

                // 5. Tạm tắt RaiseListChangedEvents để tránh DataGridView vẽ lại liên tục (Fix P1.1)
                dataList.RaiseListChangedEvents = false;

                // Áp dụng dữ liệu vào dataList ngay trên luồng UI (an toàn, không cần Invoke)
                foreach (var result in results)
                {
                    string sharedKey = result.Group.Key;

                    if (sharedKey == "ORT" || sharedKey == "ITEMCODE_LOTNO_CHECK" || sharedKey == "ONPRODUCT")
                    {
                        if (result.IsSuccess && !string.IsNullOrEmpty(result.Status))
                        {
                            // Parse chuỗi kết quả
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

                            // Map lại vào nhóm row tương ứng
                            foreach (var row in result.Group)
                            {
                                string lookupKey = switchORT(row.SheetName); // (Fix P3.1)
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
                        // Các mục tra cứu bình thường (Single)
                        foreach (var row in result.Group)
                        {
                            row.Status = result.IsSuccess ? result.Status : "Lỗi / Không tìm thấy";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 6. Bắt lỗi async void để tránh Crash Process toàn ứng dụng (Fix P0.2)
                MessageBox.Show($"Có lỗi xảy ra trong quá trình tải dữ liệu:\n{ex.Message}", "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 7. Khôi phục trạng thái UI (Fix P0.1)
                if (!this.IsDisposed)
                {
                    if (dataList != null)
                    {
                        dataList.RaiseListChangedEvents = true; // Bật lại event
                        dataList.ResetBindings(); // Báo cho UI vẽ lại 1 lần duy nhất
                    }

                    btn_load.Enabled = true;
                    MessageBox.Show("Hoàn tất kiểm tra dữ liệu!", "Thông báo", MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
        }

        private string switchORT(string sheetName)
        {
            sheetName = sheetName.Replace(" - ICT DATA", "");

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
                    break;
            }
        }
    }

    public class ReportCategoryModel : AntdUI.NotifyProperty
    {
        // 1. Cột Select (Checkbox)
        private bool _select;

        public bool Select
        {
            get => _select;
            set
            {
                _select = value;
                OnPropertyChanged("Select");
            }
        }

        public string SheetName { get; set; }
        public string CollectFrom { get; set; }

        private string _status;

        public string Status
        {
            get => _status;
            set
            {
                _status = value;
                OnPropertyChanged("Status");
            }
        }

        // 2. Thuộc tính quản lý trạng thái Lock/Unlock
        private bool _isLocked = true; // Mặc định là Lock

        public bool IsLocked
        {
            get => _isLocked;
            set
            {
                _isLocked = value;
                OnPropertyChanged("IsLocked");
                OnPropertyChanged("Action"); // Khi trạng thái Lock đổi, báo cho cột Action đổi text/màu theo
            }
        }

        // 3. Cột Action (Hiển thị nút bấm giao diện AntUI)
        public AntdUI.CellButton[] Action => new AntdUI.CellButton[]
        {
            new AntdUI.CellButton("btnLock") // Đặt một ID cố định cho nút (nếu cần phân biệt nhiều nút)
            {
                Text = _isLocked ? "Lock" : "Unlock", // BẮT BUỘC: Gán giá trị chữ hiển thị tại đây
                Type = _isLocked ? AntdUI.TTypeMini.Primary : AntdUI.TTypeMini.Default
            }
        };
    }
}