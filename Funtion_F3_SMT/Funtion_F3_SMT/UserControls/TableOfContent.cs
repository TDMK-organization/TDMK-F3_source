using OK2SHIP_SMT.Repositories;
using OK2SHIP_SMT.Services;
using OK2SHIP_SMT.Views;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;

namespace OK2SHIP_SMT.UserControls
{
    public partial class TableOfContent : UserControl
    {

        private TableOfContentService tableOfContentService = new TableOfContentService();
        public TableOfContent()
        {
            InitializeComponent();
            SetUpDataGridView();
        }
        #region Event
        private void btn_load_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = tableOfContentService.loadTOCSItemCode(tb_itemCodeView.Text.Trim());
                if (dt.Rows.Count <= 0)
                {
                    throw new Exception("ItemCode chưa được cài đặt!");
                }
                DataRow dataRow = dt.Rows[0];
                tb_Shipping.Text = dataRow["ShippingFrom"].ToString();
                tb_programname.Text = dataRow["ProgramName"].ToString();
                tb_MCO.Text = dataRow["MCORevision"].ToString();
                tb_ODB.Text = dataRow["ODBRevision"].ToString();
                tb_buildconfig.Text = dataRow["Build"].ToString();
                tb_xOut.Text = dataRow["XOUTRate"].ToString();
                tb_EECODE.Text = dataRow["EEEECode"].ToString();
                dataGridView.DataSource = tableOfContentService.MakeTable(dtp_send.Text);
                checkTAR();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btn_UpdateItemName_Click(object sender, EventArgs e)
        {
            string itemName = tb_ItemName.Text.Trim();
            int counting = listItemCode.Count();
            try
            {
                if (counting <= 0)
                {
                    throw new Exception("Hãy nhập ít nhất một itemcode!");
                }
                if (string.IsNullOrEmpty(itemName))
                {
                    throw new Exception("Hãy nhập itemName!");
                }
                DataTable dataTable = (DataTable)dgv_ItemName.DataSource;
                if (dataTable != null)
                {
                    dataTable.Rows.Clear();
                }
                else
                {
                    throw new Exception("Chưa có dataTable");
                }
                DataRow dr = dataTable.NewRow();
                dr["ItemName"] = itemName;
                dr["ProgramName"] = tb_Program.Text.Trim();
                dr["ODBRevision"] = tb_ODBRe.Text.Trim();
                dr["MCORevision"] = tb_MCORe.Text.Trim();
                dr["Build"] = tb_Buildz.Text.Trim();
                dr["XOUTRate"] = tb_xoutz.Text.Trim();
                dr["ShippingFrom"] = tb_shippingz.Text.Trim();
                dr["FactoryCode"] = tb_FactoryCode.Text.Trim();
                foreach (var item in listItemCode)
                {
                    DataRow drs = ExportProcess.CloneDataRow(dr);
                    drs["ID"] = dataTable.Rows.Count + 1;
                    drs["ItemCode"] = item;
                    dataTable.Rows.Add(drs);

                }
                dgv_ItemName.DataSource = dataTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btn_createItemName_Click(object sender, EventArgs e)
        {
            string itemName = tb_ItemName.Text.Trim();
            int counting = listItemCode.Count();
            try
            {
                if (counting <= 0)
                {
                    throw new Exception("Hãy nhập ít nhất một itemcode!");
                }
                if (string.IsNullOrEmpty(itemName))
                {
                    throw new Exception("Hãy nhập itemName!");
                }
                DataTable dataTable = tableOfContentService.getDataTableByItemName(itemName, (DataTable)dgv_ItemName.DataSource);
                if (dataTable.Rows.Count > 0)
                {
                    throw new Exception("ItemCode đã tồn tại!");
                }
                DataRow dr = dataTable.NewRow();
                dr["ID"] = dataTable.Rows.Count + 1;
                dr["ItemName"] = itemName;
                dr["ProgramName"] = tb_Program.Text.Trim();
                dr["ODBRevision"] = tb_Program.Text.Trim();
                dr["MCORevision"] = tb_Program.Text.Trim();
                dr["Build"] = tb_Program.Text.Trim();
                dr["XOUTRate"] = tb_Program.Text.Trim();
                dr["ShippingFrom"] = tb_Program.Text.Trim();
                foreach (var item in listItemCode)
                {
                    DataRow drs = ExportProcess.CloneDataRow(dr);
                    drs["ItemCode"] = item;
                    dataTable.Rows.Add(drs);
                }
                dgv_ItemName.DataSource = dataTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btn_removeItemCode_Click(object sender, EventArgs e)
        {
            int indexS = listBox_ItemCode.SelectedIndex;
            if (indexS < 0)
            {
                MessageBox.Show("Hãy chọn itemCode muốn xóa!");
                return;
            }
            listItemCode.Remove(listItemCode[indexS]);
            listBox_ItemCode.Items.Clear();
            listBox_ItemCode.Items.AddRange(listItemCode.ToArray());
        }
        private void btn_additemCode_Click(object sender, EventArgs e)
        {
            string itemCode = tb_ItemCode.Text.Trim();
            if (string.IsNullOrEmpty(itemCode))
            {
                MessageBox.Show("Hãy nhập itemcode");
                return;
            }
            foreach (var item in listItemCode)
            {
                if (item.Trim().Equals(itemCode))
                {
                    MessageBox.Show("Itemcode đã tồn tại!");
                    return;
                }
            }
            listItemCode.Add(itemCode);
            listBox_ItemCode.Items.Clear();
            listBox_ItemCode.Items.AddRange(listItemCode.ToArray());
            tb_ItemCode.Text = "";
        }
        private void btn_UpdateItemCode_Click(object sender, EventArgs e)
        {
            string itemCode = tb_ItemCode.Text;
            int indexS = listBox_ItemCode.SelectedIndex;
            if (indexS < 0 || string.IsNullOrEmpty(itemCode))
            {
                MessageBox.Show("Hãy chọn itemCode muốn update hoặc nhập itemcode");
                return;
            }
            listItemCode[indexS] = itemCode;
            listBox_ItemCode.Items.Clear();
            listBox_ItemCode.Items.AddRange(listItemCode.ToArray());
        }
        private void listBox_ItemCode_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = listBox_ItemCode.SelectedIndex;
            if (index != -1) // Kiểm tra xem có mục nào được chọn không
            {
                tb_ItemCode.Text = listItemCode[index];
            }
        }
        private void btn_loadItemName_Click(object sender, EventArgs e)
        {
            try
            {
                _DATABASEMODE = true;
                DataTable dt = tableOfContentService.getDataTableByItemName(tb_ItemName.Text);
                if (dt != null || dt.Rows.Count > 0)
                {
                    clearFormz();

                    dgv_ItemName.DataSource = dt;

                    fillDataRow(((DataTable)dgv_ItemName.DataSource).Rows[0]);
                }
                MessageBox.Show($"Lấy dữ liệu của Item {tb_ItemName.Text.Trim()} Thành công!", "Thành công");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}");
            }
        }
        private void btn_removeItemName_Click(object sender, EventArgs e)
        {
            removeByItemName(tb_ItemName.Text);
        }

        private void tb_ItemName_TextChanged(object sender, EventArgs e)
        {
            string str = tb_ItemName.Text;
            btn_removeItemName.Enabled = !string.IsNullOrEmpty(str);
        }

        private void btn_ClearDGV_Click(object sender, EventArgs e)
        {
            clearFormz();
        }

        private void fillDataRow(DataRow row)
        {
            tb_shippingz.Text = row["ShippingFrom"].ToString();
            tb_xoutz.Text = row["XOUTRate"].ToString();
            tb_Buildz.Text = row["Build"].ToString();
            tb_ODBRe.Text = row["ODBRevision"].ToString();
            tb_MCORe.Text = row["MCORevision"].ToString();
            tb_Program.Text = row["ProgramName"].ToString();
            tb_ItemName.Text = row["ItemName"].ToString();
            tb_FactoryCode.Text = row["FactoryCode"].ToString();
            getItemCodeByItemName(tb_ItemName.Text);
        }
        private void dgv_ItemName_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataRow row = ((DataTable)dgv_ItemName.DataSource).Rows[e.RowIndex];
                fillDataRow(row);
            }
        }



        private void btn_saveItemname_Click(object sender, EventArgs e)
        {
            bool prime = false;
        Ifinity_Loop:
            try
            {
                DataTable dt = (DataTable)dgv_ItemName.DataSource;
                if (dt == null)
                {
                    throw new Exception("Không có dữ liệu");
                }
                int res = tableOfContentService.SaveDataItemName(dt, prime);
                MessageBox.Show($"Lưu thành công {res} bản ghi!");
            }
            catch (Exception ex)
            {
                if (ex.Message.Split('-')[0].Contains("1234"))
                {
                    DialogResult result = MessageBox.Show($"{ex.Message.Split('-')[1]} \n Bạn có chắc chắn muốn tiếp tục lưu?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    // Kiểm tra kết quả
                    if (result == DialogResult.Yes)
                    {
                        prime = true;
                        goto Ifinity_Loop;
                    }
                }
                else
                {
                    MessageBox.Show($"Lỗi trong quá trình lưu {ex.Message}", "Thông báo lỗi");
                }
                return;
            }
            clearFormz();
        }
        private bool _DATABASEMODE = false;
        private void btn_getData_Click(object sender, EventArgs e)
        {
            try
            {
                lb_itemNameEditor.Text = "List of File";
                _DATABASEMODE = false;
                TableOfContentService tb = new TableOfContentService();
                DataTable datatable = tb.GetDataFormFile(tb_locationFile.Text);
                dgv_ItemName.DataSource = datatable;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"TDMK Error: {ex}");
            }
        }
        private void tb_locationFile_DoubleClick(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Thiết lập các thuộc tính của OpenFileDialog (tùy chọn)
            openFileDialog.Filter = "All files (*.*)|*.*";
            openFileDialog.FilterIndex = 1; // Chọn bộ lọc đầu tiên theo mặc định
            openFileDialog.RestoreDirectory = true; // Khôi phục thư mục đã chọn lần trước

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Lấy đường dẫn tệp đã chọn
                tb_locationFile.Text = openFileDialog.FileName;
            }
        }
        private void btn_Clear_Click(object sender, EventArgs e)
        {
            clearForm();
        }
        private void btn_loadDataView_Click(object sender, EventArgs e)
        {
            ProcessLoadData();
        }
        private void tb_lotnoView_TextChanged(object sender, EventArgs e)
        {
            tb_lotnoView.Text = ValidateService.lotNoHandle(tb_lotnoView.Text);
        }
        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
            if (e.RowIndex < 1 || e.ColumnIndex < 0)
                return;
            if (IsTheSameCellValue(e.ColumnIndex, e.RowIndex))
            {
                e.AdvancedBorderStyle.Top = DataGridViewAdvancedCellBorderStyle.None;

            }
            else
            {
                e.AdvancedBorderStyle.Top = dataGridView.AdvancedCellBorderStyle.Top;

            }
        }
        private void dataGridView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex == 0)
                return;
            if (IsTheSameCellValue(e.ColumnIndex, e.RowIndex))
            {
                e.Value = "";
                e.FormattingApplied = true;
            }
        }
        private void btn_saveDataView_Click(object sender, EventArgs e)
        {
            SaveProcess();

        }

        private void btn_Export_Click(object sender, EventArgs e)
        {
            Export_Process();
        }
        #endregion

        #region Action
        List<string> listItemCode = new List<string>();
        private void getItemCodeByItemName(string itemName)
        {
            listBox_ItemCode.Items.Clear();
            try
            {
                if (_DATABASEMODE)
                {
                    listItemCode = tableOfContentService.GetListItemCodeByItemName(itemName);
                }
                else
                {
                    DataTable dt = (DataTable)dgv_ItemName.DataSource;
                    listItemCode = tableOfContentService.GetListItemCodeByItemName(itemName, dt);
                }
                listBox_ItemCode.Items.AddRange(listItemCode.ToArray());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}");
            }
        }
        private void clearForm()
        {
            tb_itemCodeView.Text = "";
            tb_lotnoView.Text = "";
            tb_programname.Text = "";
            tb_MCO.Text = "";
            tb_ODB.Text = "";
            tb_buildconfig.Text = "";
            tb_xOut.Text = "";
            tb_Shipping.Text = "";
            tb_shippingTo.Text = "";
            tb_qty.Text = "";
        }
        void Export_Process()
        {
            string itemCode = tb_itemCodeView.Text, lotno = tb_lotnoView.Text;
            try
            {
                tableOfContentService.Export(itemCode, lotno);
                MessageBox.Show("Export thành công");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Có lỗi xảy ra trong quá trình export {ex.Message}");
            }
        }
        void SaveProcess()
        {
            try
            {
                string itemCode = tb_itemCodeView.Text, lotno = tb_lotnoView.Text;
                if (!tableOfContentService.checkItemCodeLotNo(itemCode, lotno))
                {
                    throw new Exception("Hãy nhập ItemCode LotNo");
                }
                DataTable dt = tableOfContentService.getDataTable();
                var row = dt.NewRow();
                row["ItemCode"] = itemCode;
                row["LotNo"] = lotno;
                row["BuildDate"] = dtp_build.Value;
                row["SendDate"] = dtp_send.Value;
                row["DeliveryQuatity"] = tb_qty.Text;
                row["ShippingTo"] = tb_shippingTo.Text;
                dt.Rows.Add(row);
                DataTable dt2 = dataGridView.DataSource as DataTable;
                int res = 0;
                try
                {
                    res = tableOfContentService.SaveData(dt, dt2, false);
                }
                catch (Exception ex)
                {
                    string s = ex.Message;
                    if (s.Split('-')[0].Trim().Equals("1412"))
                    {
                        if (MessageBox.Show(s.Split('-')[1], "Xác nhận", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {

                            res = tableOfContentService.SaveData(dt, dt2, true);
                        }
                    }
                    else
                    {
                        throw ex;
                    }
                }
                MessageBox.Show($"Đã lưu thành công {res} bản ghi!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Có lỗi trong quá trình lưu: {ex.Message}");
            }
        }
        bool IsTheSameCellValue(int column, int row)
        {
            if (column > 1)
            {
                return false;
            }
            DataGridViewCell cell1 = dataGridView[column, row];
            DataGridViewCell cell2 = dataGridView[column, row - 1];
            if (cell1.Value == null || cell2.Value == null)
            {
                return false;
            }

            return cell1.Value.ToString() == cell2.Value.ToString();
        }
        private void ProcessLoadData()
        {
            try
            {
                string itemCode = tb_itemCodeView.Text.Trim(), lotno = tb_lotnoView.Text.Trim();
                if (!tableOfContentService.checkItemCodeLotNo(itemCode, lotno))
                {
                    throw new Exception("Hãy Nhập ItemCode LotNo");
                }
                DataTable dt = tableOfContentService.getAllTOC(new string[] { "ItemCode", "LotNo" }, new string[] { itemCode, lotno });
                if (dt.Rows.Count <= 0)
                {
                    throw new Exception($"Dữ liệu {itemCode} - {lotno} chưa tồn tại!");
                }
                var rowRes = dt.Rows[0];
                string date = rowRes["BuildDate"].ToString();
                string format = "M/d/yyyy h:mm:ss tt";
                if (DateTime.TryParseExact(rowRes["BuildDate"].ToString(), format, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime parsedDateTime))
                {
                    dtp_build.Value = parsedDateTime;
                }
                else
                {
                    dtp_build.Value = DateTime.Today;
                }
                if (DateTime.TryParseExact(rowRes["SendDate"].ToString(), format, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out parsedDateTime))
                {
                    dtp_send.Value = parsedDateTime;
                }
                else
                {
                    dtp_send.Value = DateTime.Today;
                }
                tb_qty.Text = rowRes["DeliveryQuatity"].ToString();
                tb_shippingTo.Text = rowRes["ShippingTo"].ToString();
                string content = rowRes["ContentTable"].ToString();
                DataTable contentDt = ConverterService.JsonToDataTable(content);
                dataGridView.DataSource = contentDt;
                if (rowRes != null)
                {
                    MessageBox.Show("Lấy dữ liệu thành công!", "Thông báo");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void removeByItemName(string text)
        {
            throw new NotImplementedException();
        }
        private void SetUpDataGridView(DataTable datatable = null)
        {
            tabPage3.Controls.Add(new Packaging() { Dock = DockStyle.Fill });
            //dataGridView.DataSource = datatable;
            //dataGridView.DataSource = tableOfContentService.MakeTable();
        }
        private void clearFormz()
        {
            dgv_ItemName.DataSource = new DataTable();
            listBox_ItemCode.Items.Clear();
            tb_locationFile.Text = "";
            tb_Buildz.Text = "";
            tb_ODBRe.Text = "";
            tb_MCORe.Text = "";
            tb_Program.Text = "";
            tb_ItemName.Text = "";
            tb_ItemCode.Text = "";
        }













        #endregion


        private CommonForm commonDialog;

        private void close()
        {
            commonDialog.Close();
        }
        private void btn_pick_Click(object sender, EventArgs e)
        {
            PickData pick = new PickData(close);
            commonDialog = new CommonForm("Pick time", pick, true);
            commonDialog.WindowState = FormWindowState.Normal;
            commonDialog.Size = new Size(508, 300);
            commonDialog.ShowDialog();
            string str = pick.ShouldClose();
            Dictionary<int, string> dic = new Dictionary<int, string>();
            foreach (DataGridViewCell cell in dataGridView.SelectedCells)
            {
                if (dic.TryGetValue(cell.RowIndex, out string value))
                {

                }
                else
                {
                    dic.Add(cell.RowIndex, str);
                }
            }
            DataTable dt = dataGridView.DataSource as DataTable;
            foreach (int item in dic.Keys)
            {
                string strz = dt.Rows[item]["Target date Request"].ToString();
                if (strz.Trim().ToUpper().Contains("MAX AFTER SHIPPING"))
                {
                    if (strz.Contains("30"))
                    {
                        dt.Rows[item]["Target date Submission"] = str;
                        int com = compareDate(str, 30);
                    }
                    else if (strz.Contains("5"))
                    {
                        dt.Rows[item]["Target date Submission"] = str;
                        int com = compareDate(str, 5);
                    }
                }
            }
            checkTAR();
        }
        public static DateTime ConvertStringToDateTime(string dateString)
        {
            string format = "dd-MMM";
            CultureInfo provider = CultureInfo.InvariantCulture;
            DateTime dateTime;

            if (DateTime.TryParseExact(dateString, format, provider, DateTimeStyles.None, out dateTime))
            {
                return dateTime;
            }
            else
            {
                Console.WriteLine("Invalid date format.");
                return DateTime.MinValue; // Or handle the error as needed
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="date">date form</param>
        /// <param name="num"> after send (area)</param>
        /// <returns>
        /// 1: bigger
        /// 0, -1: in area 
        /// -1: smaller
        /// </returns>
        private int compareDate(string date, int num = 0)
        {
            int today = dtp_send.Value.Day;
            if (string.IsNullOrEmpty(date))
            {
                return -2;
            }
            int dateCompare = ConvertStringToDateTime(date).Day;
            double numDate = dateCompare - today;
            if (numDate < 0)
            {
                return -3;
            }
            if (numDate == 0)
            {
                return 0;
            }
            if (numDate <= num)
            {
                return 0;
            }
            return -2;
        }

        private void checkTAR()
        {
            for (int rowIndex = 0; rowIndex < ((DataTable)dataGridView.DataSource).Rows.Count; rowIndex++)
            {
                DataRow row = ((DataTable)dataGridView.DataSource).Rows[rowIndex];
                if (row["Target date Submission"].ToString().Contains("NA") || row["Target date Submission"].ToString().Contains("N/A"))
                {
                    row["Status"] = row["Target date Submission"];
                }
                else if (row["Target date Request"].ToString().Contains("Prior to ship"))
                {
                    string a1 = row["Target date Submission"].ToString().Replace(" ", "");
                    string a = dtp_send.Value.ToString("dd-MMM");
                    if (row["Target date Submission"].ToString().Replace(" ", "").Equals(dtp_send.Value.ToString("dd-MMM")))
                    {
                        row["Status"] = "Pass";
                    }
                }
                else if (row["Target date Request"].ToString().Contains("max after shipping"))
                {
                    int? com = null;
                    if (row["Target date Request"].ToString().Contains("30"))
                    {
                        com = compareDate(row["Target date Submission"].ToString(), 30);
                    }
                    else if (row["Target date Request"].ToString().Contains("5"))
                    {
                        com = compareDate(row["Target date Submission"].ToString(), 5);
                    }
                    string status = "";
                    switch (com)
                    {
                        case 0:
                            status = "On going";
                            break;
                        case -1:
                            status = "On going";
                            break;
                        default:
                            break;
                    }
                    row["Status"] = status;
                    if (new[] { "Heat Soak and Recovery", "Thermal Cycling", "Thermal shock", "Bending after thermal cycling", "Bending after heat soak " }.Contains(row["Test"].ToString().Trim()))
                    {
                        row["Status"] = "On going";
                    }
                }

            }
        }

        private void dtp_send_ValueChanged(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dataGridView.DataSource;
            if (dt == null)
            {
                return;
            }

            string str = dtp_send.Text;
            foreach (DataRow row in dt.Rows)
            {
                string z = row["Target date Request"].ToString();
                if (row["Target date Request"].ToString().Contains("Prior to ship") && !row["Target date Submission"].ToString().Contains("NA"))
                {
                    row["Target date Submission"] = str;
                }
            }
        }
    }
}
