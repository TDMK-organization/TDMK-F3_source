using System;
using System.ComponentModel; // Bắt buộc thêm để dùng BindingList
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using AntdUI;
using Export_FPCA_OK2ship_Auto_System.Services;

namespace Export_FPCA_OK2ship_Auto_System.Views
{
    public partial class WorksheetScreen : UserControl
    {
        private string[] categories;

        private string mode = "MANUAL";
        public string _location = "";

        public WorksheetScreen(string[] category, string location = "")
        {
            InitializeComponent();
            categories = category;
            if (string.IsNullOrEmpty(location))
            {
                mode = "FORMAT";
            }

            ip.Text = location;
            
            init();
        }

        #region Event

        private void ip_DoubleClick(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.csv|All files (*.*)|*.*";
                openFileDialog.Title = "Chọn file Excel";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    ip.Text = filePath;
                }
            }
        }

        private void btn_scan_Click(object sender, EventArgs e)
        {
            scan();
        }

        #endregion

        #region Method

        public DataTable GetTableData()
        {
            // 1. Lấy DataSource của bảng và ép kiểu về BindingList<SheetRowData>
            var dataList = dgv.DataSource as System.ComponentModel.BindingList<SheetRowData>;

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
                DataRow dr = dt.NewRow();
                dr["Select"] = row.Selected;
                dr["SheetName"] = row.Name;
                dr["SheetDic"] = row.Action.Text;
                dt.Rows.Add(dr);

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

        private void scan()
        {
            string location = ip.Text;
            string[] arr = SetupService.ScanWorkSheet(location);

            display(arr);
        }

        private void display(string[] arr)
        {
            if (arr == null || arr.Length == 0)
            {
                dgv.DataSource = null;
                return;
            }

            var dataList = new BindingList<SheetRowData>();

            foreach (var item in arr)
            {
                var rowData = new SheetRowData
                {
                    Selected = true,
                    Name = item
                };

                List<AntdUI.ISelectItem> items = new List<AntdUI.ISelectItem>();
                foreach (string c in categories)
                {
                    items.Add(new AntdUI.SelectItem(c));
                }

                // ==========================================
                // 1. TÌM GIÁ TRỊ MẶC ĐỊNH CHO TỪNG DÒNG
                // ==========================================
                string defaultActionText = "Chọn thao tác";

                // Kiểm tra an toàn xem mảng categories có dữ liệu không
                // (Dùng .Length nếu categories là string[], dùng .Count nếu là List<string>)
                if (categories != null && categories.Length > 0)
                {
                    if (categories.Contains(item))
                    {
                        // Nếu sheet name (item) có trong categories -> Gán mặc định là tên sheet
                        defaultActionText = item;
                    }
                    else
                    {
                        // Nếu không có -> Gán mặc định là phần tử đầu tiên của categories
                        defaultActionText = categories[0];
                    }
                }

                // ==========================================
                // 2. KHỞI TẠO NÚT ACTION VỚI GIÁ TRỊ VỪA TÌM
                // ==========================================
                var actionBtn =
                    new AntdUI.CellButton(Guid.NewGuid().ToString(), defaultActionText, AntdUI.TTypeMini.Primary)
                    {
                        DropDownItems = items.ToArray()
                    };

                // Tuỳ chọn: Cập nhật màu sắc cho đúng với giá trị mặc định ban đầu
                switch (defaultActionText)
                {
                    case "Default": actionBtn.Type = AntdUI.TTypeMini.Default; break;
                    case "Primary": actionBtn.Type = AntdUI.TTypeMini.Primary; break;
                    case "Success": actionBtn.Type = AntdUI.TTypeMini.Success; break;
                    case "Error": actionBtn.Type = AntdUI.TTypeMini.Error; break;
                    case "Warn": actionBtn.Type = AntdUI.TTypeMini.Warn; break;
                    case "Processing": actionBtn.Type = AntdUI.TTypeMini.Primary; break;
                }

                // Xử lý sự kiện khi người dùng chọn thủ công trong dropdown
                actionBtn.DropDownValueChanged = (value) =>
                {
                    string badge = value.ToString();
                    actionBtn.Text = badge;

                    switch (badge)
                    {
                        case "Default": actionBtn.Type = AntdUI.TTypeMini.Default; break;
                        case "Primary": actionBtn.Type = AntdUI.TTypeMini.Primary; break;
                        case "Success": actionBtn.Type = AntdUI.TTypeMini.Success; break;
                        case "Error": actionBtn.Type = AntdUI.TTypeMini.Error; break;
                        case "Warn": actionBtn.Type = AntdUI.TTypeMini.Warn; break;
                        case "Processing": actionBtn.Type = AntdUI.TTypeMini.Primary; break;
                    }

                    rowData.Action = actionBtn;
                    dgv.Invalidate();
                };

                rowData.Action = actionBtn;
                dataList.Add(rowData);
            }

            dgv.DataSource = dataList;
        }

        // Hàm này dư thừa vì bạn đang dùng List Model, có thể xoá nếu không cần thiết
        private DataTable getDataTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Select", typeof(bool));
            dt.Columns.Add("Sheet Name", typeof(string));
            dt.Columns.Add("Action", typeof(string));
            return dt;
        }

        private void init()
        {
            InitTableColumns();
            if (mode == "MANUAL")
            {
                scan();
            }
        }

        #endregion
    }

    public class SheetRowData : AntdUI.NotifyProperty
    {
        private bool selected;

        public bool Selected
        {
            get => selected;
            set
            {
                selected = value;
                OnPropertyChanged("Selected");
            }
        }

        private string name;

        public string Name
        {
            get => name;
            set
            {
                name = value;
                OnPropertyChanged("Name");
            }
        }

        private AntdUI.CellButton action;

        public AntdUI.CellButton Action
        {
            get => action;
            set
            {
                action = value;
                OnPropertyChanged("Action");
            }
        }
    }
}