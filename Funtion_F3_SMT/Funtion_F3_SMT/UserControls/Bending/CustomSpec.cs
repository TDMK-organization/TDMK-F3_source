using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using OK2SHIP_SMT.Services;

namespace OK2SHIP_SMT.UserControls.Bending
{
    public partial class CustomSpec : AntdUI.Window
    {
        public string _json { get; set; }

        private string[] cycles = new[]
            { "L1", "L2", "L3", "L4", "L5", "L10", "L15", "L20", "L25", "L30", "L40", "L50" };

        public CustomSpec(string json)
        {
            InitializeComponent();
            _json = json;
        }

        #region Event

        private void btn_cf_Click(object sender, EventArgs e)
        {
            this.Close();
            List<string> result = new List<string>();
            DataTable dt = dgv.DataSource as DataTable;
            foreach (DataRow row in dt.Rows)
            {
                string nameCycle = row["Name Cycle"].ToString();
                string value = row["Custom Spec"].ToString();
                if (!string.IsNullOrEmpty(value))
                {
                    result.Add($"{nameCycle}:{value}");
                }
            }

            _json = string.Join(";", result);
        }

        private void btn_add_Click(object sender, EventArgs e)
        {
            // 1. Lấy dữ liệu từ bảng hiện hành
            DataTable dt = dgv.DataSource as DataTable;
            if (dt == null || dt.Rows.Count == 0)
            {
                MessageBox.Show("Bảng chưa có dữ liệu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 2. Lấy giá trị từ TextBox
            string fileValue = tb_logfile.Text.Trim();
            if (string.IsNullOrEmpty(fileValue))
            {
                MessageBox.Show("Vui lòng chọn hoặc nhập giá trị vào ô log file trước!", "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            string targetColumn = "Custom Spec";

            if (!dt.Columns.Contains(targetColumn))
            {
                MessageBox.Show($"Không tìm thấy cột '{targetColumn}' trong bảng dữ liệu!", "Lỗi", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            int updatedCount = 0;

            // 3. Duyệt qua từng dòng trong DataTable để tìm các dòng được Select
            foreach (string value in select.SelectedValue)
            {
                int rowIndex = Array.FindIndex(cycles, n => n == value);

                if (rowIndex != -1 && rowIndex < dt.Rows.Count)
                {
                    dt.Rows[rowIndex]["Custom Spec"] = fileValue;
                }

                updatedCount++;
            }

            // 4. Thông báo kết quả
            if (updatedCount > 0)
            {
                display(dt);
                AntdUI.Notification.success(this, "Thành công", $"Đã gán giá trị cho {updatedCount} dòng được chọn!",
                    autoClose: 3, align: AntdUI.TAlignFrom.Bottom);
            }
            else
            {
                AntdUI.Notification.info(this, "Thông báo", "Bạn chưa tick chọn dòng nào trong bảng!", autoClose: 3,
                    align: AntdUI.TAlignFrom.Bottom);
            }
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void CustomSpec_Load(object sender, EventArgs e)
        {
            select.Items.AddRange(cycles);
            display(getDataFormString(_json));
        }

        #endregion

        #region Method

        private DataTable getDBStructor()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("NO", typeof(int));
            dt.Columns.Add("Name Cycle", typeof(string));
            dt.Columns.Add("Custom Spec", typeof(string));
            foreach (string s in cycles)
            {
                DataRow dr = dt.NewRow();
                dr["NO"] = dt.Rows.Count + 1;
                dr["Name Cycle"] = s.ToString();
                dt.Rows.Add(dr);
            }

            return dt;
        }

        private DataTable getDataFormString(string json)
        {
            DataTable dt = getDBStructor();
            string[] split = json.Split(';');
            foreach (string s in split)
            {
                string[] key = s.Split(':');
                int rowIndex = Array.FindIndex(cycles, n => n == key[0].ToString());

                if (rowIndex != -1 && rowIndex < dt.Rows.Count)
                {
                    dt.Rows[rowIndex]["Custom Spec"] = key[1];
                }
            }

            return dt;
        }

        private void display(DataTable dt)
        {
            dgv.DataSource = dt;
        }

        #endregion
    }
}