using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using AntdUI;
using Export_FPCA_OK2ship_Auto_System.Services;

namespace Export_FPCA_OK2ship_Auto_System.Views
{
    public partial class SetupForm : Form
    {
        public string SHEET_NAME { get; set; }

        private TDMK_ReportManagerService _service;

        public SetupForm(string sheetName, TDMK_ReportManagerService service)
        {
            InitializeComponent();
            SHEET_NAME = sheetName;
            this._service = service;
            InitView();
        }

        private string namecategory = "";

        #region Event

        private void sl_type_SelectedIndexChanged(object sender, IntEventArgs e)
        {
            namecategory = sl_type.Text;
        }

        private void btn_addCateogry_Click(object sender, EventArgs e)
        {
            string text = ip.Text;
            foreach (string i in selectMultiple.Items)
            {
                if (i.Equals(text))
                {
                    MessageBox.Show("Category đã tồn tại!");
                    return;
                }
            }

            selectMultiple.Items.Add(text);
        }

        private void step_TabIndexChanged(object sender, EventArgs e)
        {
            Debugger.Break();
        }

        private void btn_next_Click(object sender, EventArgs e)
        {
            int i = step.Current;
            step.Current = i + 1;
            setupBtn(1);
        }

        private void btn_back_Click(object sender, EventArgs e)
        {
            int i = step.Current;
            step.Current = i - 1;
            setupBtn(-1);
        }

        #endregion

        #region Method

        private void setupBtn(int stepz = 0)
        {
            int i = step.Current;
            if (i == 0)
            {
                btn_back.Enabled = false;
            }
            else
            {
                btn_back.Enabled = true;
            }

            if (i == 2)
            {
                btn_next.Enabled = false;
            }
            else
            {
                btn_next.Enabled = true;
            }

            Control currentControl = tableLayoutPanel1.GetControlFromPosition(0, 1);

            if (currentControl != null)
            {
                tableLayoutPanel1.Controls.Remove(currentControl);
            }

            if (i == 0)
            {
                tableLayoutPanel2.Dock = DockStyle.Fill;
                tableLayoutPanel1.Controls.Add(tableLayoutPanel2, 0, 1);
            }

            if (i == 1)
            {
                switch (namecategory)
                {
                    case "Database + Sheet":
                    case "Database":
                        dbScreen = new DBScreen() { Dock = DockStyle.Fill };
                        tableLayoutPanel1.Controls.Add(dbScreen, 0, 1);
                        break;
                    case "Format":
                        if (stepz == 1)
                        {
                            string[] values = ((object[])selectMultiple.SelectedValue).Select(x => x.ToString())
                                .ToArray();
                            string s = new ExportProcess().getFormatLocation();
                            using (OpenFileDialog openFileDialog = new OpenFileDialog())
                            {
                                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.csv|All files (*.*)|*.*";
                                openFileDialog.Title = "Chọn file Excel";

                                // Gán đường dẫn mặc định bạn muốn mở (ví dụ lấy từ ip.Text hoặc truyền biến location vào đây)
                                string location = s;

                                // Kiểm tra đường dẫn để gán InitialDirectory
                                if (!string.IsNullOrEmpty(location))
                                {
                                    if (System.IO.Directory.Exists(location))
                                    {
                                        // Nếu location là một thư mục hợp lệ
                                        openFileDialog.InitialDirectory = location;
                                    }
                                    else if (System.IO.File.Exists(location))
                                    {
                                        // Nếu location là đường dẫn một file (ví dụ người dùng đã chọn file trước đó)
                                        // -> Lấy thư mục chứa file đó
                                        openFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(location);
                                    }
                                }

                                if (openFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    string filePath = openFileDialog.FileName;
                                    s = filePath;
                                }
                            }

                            wsScreen = new WorksheetScreen(values, s) { Dock = DockStyle.Fill };
                        }

                        tableLayoutPanel1.Controls.Add(wsScreen, 0, 1);
                        break;
                    case "Worksheet":
                        if (stepz == 1)
                        {
                            string[] values = ((object[])selectMultiple.SelectedValue).Select(x => x.ToString())
                                .ToArray();
                            wsScreen = new WorksheetScreen(values) { Dock = DockStyle.Fill };
                        }

                        tableLayoutPanel1.Controls.Add(wsScreen, 0, 1);
                        break;
                    default:
                        break;
                }
            }

            if (i == 2)
            {
                DataTable dt;
                string location = "";
                if (!namecategory.Contains("Database"))
                {
                    dt = wsScreen.GetTableData();
                    location = wsScreen.ip.Text;
                }
                else
                {
                    dt = dbScreen.GetTableData();
                }

                runScreen = new RunProgess(dt, location, _service);
                tableLayoutPanel1.Controls.Add(runScreen, 0, 1);
            }
        }

        private DBScreen dbScreen = null;
        private WorksheetScreen wsScreen = null;
        private RunProgess runScreen = null;

        private void InitView()
        {
            setupBtn();
            sl_type.Items.Add("Database");
            sl_type.Items.Add("Format");
            sl_type.Items.Add("Worksheet");
            sl_type.Items.Add("Database + Sheet");
            selectMultiple.Items.AddRange(TDMK_ReportManagerService.categories);
        }

        #endregion
    }
}