using Funtion_F3_SMT;
using OK2SHIP_SMT.Repositories;
using OK2SHIP_SMT.Services;
using OK2SHIP_SMT.ToolBoxs;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
namespace OK2SHIP_SMT.UserControls
{
    public partial class SEM : UserControl
    {
        private bool browseStatusFile = false;
        private PictureBox pictureBox = new PictureBox();
        public string PROCESS { get; set; }
        private int MODE = 1, MAX_MODE = 2;
        private SEMServices semServices = new SEMServices();
        UC_EnvironmentTable UC_env = new UC_EnvironmentTable() { Dock = DockStyle.Fill };
        Dictionary<string, Dictionary<string, DataTable>> dic_UC_Env = new Dictionary<string, Dictionary<string, DataTable>>();

        public SEM(string pROCESS)
        {
            InitializeComponent();
            SetUpProcessScreen(pROCESS);
        }
        #region Event
        private void CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Kiểm tra xem có phải là cột "Check Result" không
            if (dgv_Combobox.Columns[e.ColumnIndex].Name == "Judgement") // Thay "CheckResult" bằng tên cột thực tế
            {
                // Kiểm tra xem giá trị ô có phải là "NG" không
                if (e.Value != null && e.Value.ToString() == "NG")
                {
                    e.CellStyle.BackColor = Color.Red; // Tô màu chữ đỏ
                                                       // Hoặc bạn có thể tô màu nền đỏ
                                                       // e.CellStyle.BackColor = Color.Red;
                }
            }
        }
        private void DataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Kiểm tra xem có phải là cột "Check Result" không
            if (dgv_Combobox.Columns[e.ColumnIndex].Name == "CheckResults") // Thay "CheckResult" bằng tên cột thực tế
            {
                // Kiểm tra xem giá trị ô có phải là "NG" không
                if (e.Value != null && e.Value.ToString() == "NG")
                {
                    e.CellStyle.BackColor = Color.Red; // Tô màu chữ đỏ
                                                       // Hoặc bạn có thể tô màu nền đỏ
                                                       // e.CellStyle.BackColor = Color.Red;
                }
            }
        }
        private void btn_checkBin_Click(object sender, EventArgs e)
        {
            try
            {
                switch (PROCESS)
                {
                    case "SEM BSE & Binarization":
                        DataTable dataTable = (DataTable)dgv_Combobox.DataSource;
                        semServices.JudgementCheck(dataTable);
                        dgv_Combobox.DataSource = dataTable;
                        break;
                    case "Bar Code Verification":
                        DataTable dataTablez = (DataTable)dataGridView.DataSource;
                        BarCodeVertification barCodeVertification = new BarCodeVertification();
                        dgv_left.DataSource = barCodeVertification.CheckSum(dataTablez);
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}");
            }
        }
        /// <summary>
        /// Open browser dialog
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tb_loactionFolder_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (!browseStatusFile)
            {

                FolderBrowserDialog dialog = new FolderBrowserDialog();
                dialog.ShowDialog();
            }
            else
            {
                OpenFileDialog fileDialog = new OpenFileDialog();
                fileDialog.Filter = "All files (*.*)|*.*";
                fileDialog.FilterIndex = 1;
                fileDialog.RestoreDirectory = true;
                fileDialog.Title = "Chọn tệp";
                fileDialog.ShowDialog();
                tb_locationFolder.Text = fileDialog.FileName;
            }
        }

        private void tb_Lotno_Leave(object sender, EventArgs e)
        {
            tb_Lotno.Text = ValidateService.lotNoHandle(tb_Lotno.Text);
        }
        private void btn_GetData_Click(object sender, EventArgs e)
        {
            getDataFormFile(tb_locationFolder.Text.Trim(), tb_ItemCode.Text.Trim(), tb_Lotno.Text.Trim(), PROCESS);
        }
        private void btn_checkSum_Click(object sender, EventArgs e)
        {

            try
            {
                string ItemCode = tb_ItemCode.Text.Trim();
                string location = tb_locationFolder.Text.Trim();
                BarCodeVertification barCodeVertification = new BarCodeVertification();
                string st = barCodeVertification.GetCheckSumData(location);
                if (string.IsNullOrEmpty(st))
                {
                    throw new Exception("Hãy đặt file phù hợp");
                }
                string[] str = st.Split('-');
                string ItemName = str[0];
                string FactoryCode = str[1];
                string ECode = str[2];
                tb_FactoryCode.Text = FactoryCode;
                tb_EEEEECode.Text = ECode;

                int num = barCodeVertification.UpdateCodeByItemName(ItemName, FactoryCode, ECode);
                MessageBox.Show($"Update {num} code với ItemName {ItemName} thành công!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}");
            }
        }
        private void btn_saveData_Click(object sender, EventArgs e)
        {
            if (!UserSession.Instance.IsLoggedIn)
            {
                MessageBox.Show("Bạn chưa đăng nhập! Hãy đăng nhập ngay!");
                Login login = new Login();
                login.Show();
                return;
            }
            //try
            //{
            switch (this.PROCESS)
            {

                case "Thermal cycling, Heat soak, Thermal shock":
                    int res = 0;
                    TCHSTSService service = new TCHSTSService();
                    if (CHANGE)
                    {
                        MessageBox.Show("Có thay đổi chưa được update!");
                        return;
                    }
                    try
                    {
                        res = service.Save((DataTable)dataGridView.DataSource, false);
                    }
                    catch (Exception ex)
                    {
                        string[] exStr = ex.Message.Split('-');
                        bool strz = exStr[0].Trim().Equals("2267");
                        if (exStr.Length >= 2 && exStr[0].Trim().Equals("2267"))
                        {
                            if (MessageBox.Show("Đã tồn tại SN trong cơ sở dữ liệu bạn có muốn ghi đè", "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            {
                                res = service.Save((DataTable)dataGridView.DataSource, true);
                            }
                        }
                        else
                        {
                            throw ex;
                        }
                    }
                    MessageBox.Show($"Đã lưu {res} row thành công");
                    dataGridView.DataSource = new DataTable();
                    break;
                case "SEM BSE & Binarization":
                    DataTable dataTablez = (DataTable)dgv_Combobox.DataSource;
                    SEMServices sem = new SEMServices();
                    bool prime = false;
                    try
                    {
                        sem.SaveProcess(dataTablez, prime);

                    }
                    catch (Exception ex)
                    {
                        string[] exStr = ex.Message.Split('-');
                        bool strz = exStr[0].Trim().Equals("2267");
                        if (exStr.Length >= 2 && exStr[0].Trim().Equals("2267"))
                        {
                            prime = MessageBox.Show("Đã tồn tại SN trong cơ sở dữ liệu bạn có muốn ghi đè", "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
                            if (prime)
                            {
                                sem.SaveProcess(dataTablez, prime);
                            }
                        }
                        else
                        {
                            throw ex;
                        }
                    }
                    int i = 1;
                    foreach (DataRow item in dataTablez.Rows)
                    {
                        item["ID"] = i++;
                        item["Judgement"] = "NG";
                    }
                    break;
                case "Bar Code Verification":
                    DataTable dataTable = (DataTable)dataGridView.DataSource;
                    string itemCode = dataTable.Rows[0]["ItemCode"].ToString(), lotNo = dataTable.Rows[0]["LotNo"].ToString();
                    BarCodeVertification barCodeVertification = new BarCodeVertification();
                    try
                    {
                        barCodeVertification.SaveProcess(dataTable, false, itemCode, lotNo);
                        MessageBox.Show("Lưu dữ liệu thành công");
                    }
                    catch (SqlException ex)
                    {
                        if (ex.Number == 2601) // Lỗi vi phạm ràng buộc khóa chính (SQL Server)
                        {
                            DialogResult dr = MessageBox.Show($"Đã tồn tại ItemCode LotNo trong cơ sở dữ liệu", "Thông báo", MessageBoxButtons.YesNoCancel);
                            if (dr == DialogResult.Yes)
                            {
                                barCodeVertification.SaveProcess(dataTable, true, itemCode, lotNo);

                            }
                        }
                        else
                        {
                            throw ex;
                        }
                    }
                    break;
                case "Environment en-durance":
                    string itemCodez = tb_ItemCode.Text.Trim();
                    string lotNoz = tb_Lotno.Text.Trim();
                    DataTable dt = (DataTable)UC_env.dataGridView.DataSource;
                    Dictionary<string, Dictionary<string, DataTable>> dic = UC_env.dictionary_Data;
                    EEDService eD = new EEDService();
                    int resz = 0;
                    try
                    {
                        resz = eD.save(itemCodez, lotNoz, dt, dic);
                    }
                    catch (Exception ex)
                    {
                        string[] ap = ex.Message.Split('-');
                        if (ap.Count() > 1 && ap[0].Contains("1234"))
                        {
                            resz = eD.save(itemCodez, lotNoz, dt, dic, MessageBox.Show("Đã tồn tại bạn muốn ghi đè?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes);
                        }
                    }
                    MessageBox.Show($"Lưu thành công {resz} row!");
                    break;
                case "OQC B2B Mating-Unmating":
                    DataTable dataTables = new DataTable();
                    if (dgv_Combobox != null)
                    {
                        dataTables = (DataTable)dgv_Combobox.DataSource;
                    }
                    if (dataTables.Rows.Count == 0)
                    {
                        MessageBox.Show("Không có gì để lưu!");
                        return;
                    }
                    OQCB2BMatingUnmatting oqc = new OQCB2BMatingUnmatting();
                    try
                    {
                        oqc.SaveProcess(dataTables);
                        MessageBox.Show("Lưu dữ liệu thành công");
                    }
                    catch (Exception ex)
                    {
                        string[] s = ex.Message.ToString().Split('-');
                        if (s.Length > 1)
                        {
                            DialogResult dialogResult = MessageBox.Show("Dữ liệu thừa bạn có muốn tiếp tục lưu", "Thông báo", MessageBoxButtons.OK);
                            if (dialogResult == DialogResult.OK)
                            {
                                oqc.SaveProcess(dataTables, true);
                            }
                            break;
                        }
                        MessageBox.Show($"Có lỗi xảy ra{ex.Message}");
                    }
                    break;
                default:
                    break;
            }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show($"Lỗi quá trình lưu, {ex.Message}");
            //    return;
            //}
        }

        private void btn_Export_Click(object sender, EventArgs e)
        {
            try
            {
                string status = "";
                switch (this.PROCESS)
                {
                    case "Environment en-durance":
                        status = new EEDService().Export(tb_ItemCode.Text, tb_Lotno.Text);
                        break;
                    case "Thermal cycling, Heat soak, Thermal shock":
                        status = new TCHSTSService().Export(tb_ItemCode.Text, tb_Lotno.Text, comboBox.Text);
                        break;
                    case "Bar Code Verification":
                        status = new BarCodeVertification().Export(tb_ItemCode.Text, tb_Lotno.Text);
                        break;
                    case "OQC B2B Mating-Unmating":
                        status = new OQCB2BMatingUnmatting().Export(tb_ItemCode.Text, tb_Lotno.Text);
                        break;
                    case "SEM BSE & Binarization":
                        try
                        {
                            status = new SEMServices().Export(tb_ItemCode.Text, tb_Lotno.Text, false);
                        }
                        catch (Exception ex)
                        {
                            string[] exStr = ex.Message.Split('-');
                            if (exStr.Count() >= 2 && exStr[0].Trim().Equals("ATPX4869"))
                            {
                                DialogResult dialogResult = MessageBox.Show("Dữ liệu thiếu bạn có muốn tiếp tục xuất dữ liệu", "Thông báo", MessageBoxButtons.YesNoCancel);
                                if (dialogResult == DialogResult.Yes)
                                {
                                    status = new SEMServices().Export(tb_ItemCode.Text, tb_Lotno.Text, true);
                                }
                            }
                            else
                            {
                                throw ex;
                            }
                        }
                        break;
                    default:
                        throw new Exception("Process not found");
                }
                MessageBox.Show(status);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return;
        }
        private void btn_switchMode_Click(object sender, EventArgs e)
        {
            slCtn_Mode.Panel2.Controls.Clear();
            MODE++;
            MODE = MODE % MAX_MODE;
            switch (MODE)
            {
                case 0:
                    tbl_Function.Dock = DockStyle.Fill;
                    slCtn_Mode.Panel2.Controls.Add(tbl_Function);
                    tbl_Function.Show();
                    break;
                case 1:
                    slctn_Location.Dock = DockStyle.Fill;
                    slCtn_Mode.Panel2.Controls.Add(slctn_Location);
                    break;
                case 2:
                    slctn_Location.Dock = DockStyle.Fill;
                    break;
                default:
                    MessageBox.Show("Mode not found");
                    break;
            }

        }
        private void splitContainer2_Panel1_Resize(object sender, EventArgs e)
        {
            resizeHeader();
        }
        private void btn_LoadData_Click(object sender, EventArgs e)
        {
            try
            {
                switch (PROCESS)
                {
                    case "Environment en-durance":
                        EEDService EED = new EEDService();
                        DataTable before = new DataTable();
                        UC_env.dictionary_Data = EED.Load(tb_ItemCode.Text, tb_Lotno.Text, out before);

                        UC_env.dataGridView.RowTemplate.Height = 300;
                        UC_env.dataGridView.DataSource = before;
                        foreach (DataGridViewColumn column in UC_env.dataGridView.Columns)
                        {
                            ((DataGridViewImageColumn)UC_env.dataGridView.Columns[column.Name]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                            ((DataGridViewImageColumn)UC_env.dataGridView.Columns[column.Name]).Width = 500; // Đặt chiều rộng cột
                        }
                        UC_env.listBox.Items.Clear();
                        UC_env.listBox.Items.AddRange(UC_env.dictionary_Data.Keys.ToArray());
                        //UC_env;
                        break;
                    case "Thermal cycling, Heat soak, Thermal shock":
                        string type = comboBox.Text.ToString().Trim();
                        TCHSTSService service = new TCHSTSService();
                        DataTable dtT = service.Load(tb_ItemCode.Text, tb_Lotno.Text, type);
                        if (dtT.Rows.Count <= 0)
                        {
                            MessageBox.Show($"{tb_ItemCode} {tb_Lotno.Text} không có dữ liệu!");
                            return;
                        }
                        dataGridView.DataSource = dtT;
                        dgv_left.DataSource = new DataTable();
                        break;
                    case "Bar Code Verification":
                        BarCodeVertification barCodeVertification = new BarCodeVertification();
                        DataTable dtBar = barCodeVertification.LoadProcess(tb_ItemCode.Text, tb_Lotno.Text);
                        dataGridView.DataSource = dtBar;
                        dataGridView.Columns["ID"].DisplayIndex = 0;
                        checkFactoryCode(tb_ItemCode.Text.Trim());
                        break;

                    case "SEM BSE & Binarization":
                        SEMServices sem = new SEMServices();
                        DataTable dtz = sem.LoadDataProcess(tb_ItemCode.Text, tb_Lotno.Text);
                        if (dtz.Rows.Count <= 0)
                        {
                            MessageBox.Show($"{tb_ItemCode} {tb_Lotno.Text} không có dữ liệu!");
                            return;
                        }
                        semServices.JudgementCheck(dtz);
                        dgv_Combobox = new CustomDataGridView(dtz, new Dictionary<string, string[]> { { "Judgement", new string[] { "", "Level 1", "Level 2", "Level 3" } } }) { Name = "", Dock = DockStyle.Fill };
                        dgv_Combobox.CellClick += cellContentClick;
                        MakeGood();
                        splitContainer2.Panel2.Controls.Clear();
                        splitContainer2.Panel2.Controls.Add(dgv_Combobox);
                        break;
                    case "OQC B2B Mating-Unmating":
                        OQCB2BMatingUnmatting oqc = new OQCB2BMatingUnmatting();
                        DataTable dt = oqc.LoadProcess(tb_ItemCode.Text, tb_Lotno.Text);
                        if (dt.Rows.Count <= 0)
                        {
                            MessageBox.Show($"{tb_ItemCode} {tb_Lotno.Text} không có dữ liệu!");
                            return;
                        }
                        dgv_Combobox = new CustomDataGridView(dt, new Dictionary<string, string[]> { { "FailureMode", new string[] { "", "Level 1", "Level 2", "Level 3" } } }) { Name = "", Dock = DockStyle.Fill };
                        dgv_Combobox.CellClick += cellContentClick;
                        splitContainer2.Panel2.Controls.Clear();
                        splitContainer2.Panel2.Controls.Add(dgv_Combobox);
                        MakeGood();
                        break;
                    default:
                        MessageBox.Show("Process not found!");
                        return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void cellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                DataGridViewCell cell = dgv_Combobox.Rows[e.RowIndex].Cells[e.ColumnIndex];

                if (cell is DataGridViewImageCell)
                {
                    Image image = TDMK_ImageConverter.ByteArrayToImage((byte[])cell.Value);

                    // Hiển thị hình ảnh trong PictureBox
                    pictureBox.Image = image;
                    pictureBox.SizeMode = PictureBoxSizeMode.Zoom; // Tùy chỉnh kích thước hình ảnh

                }
            }
        }
        #endregion

        #region Action

        private void resizeHeader()
        {
            int w = splitContainer2.Panel1.Width;
            int h = splitContainer2.Panel1.Height;
            lb_headerTable.Width = w;
            lb_headerTable.Height = h;
        }
        TableLayoutPanel sc = new TableLayoutPanel() { BackColor = Color.Aqua, Dock = DockStyle.Fill };
        DataGridView dataGridViewz = new DataGridView() { Dock = DockStyle.Fill };
        Button button = new Button();
        ComboBox comboBox = new ComboBox() { Dock = DockStyle.Fill };
        private void SetUpProcessScreen(string pROCESS)
        {
            btn_checkBin.Hide();
            resizeHeader();
            tbl_Function.Hide();
            PROCESS = pROCESS;
            lbHeadername.Text = PROCESS;
            btn_checkBin.Visible = false;
            switch (this.PROCESS)
            {
                case "Impedance":
                    button.BackColor = Color.Green;
                    button.Dock = DockStyle.Fill;
                    button.Text = "Load Form F1 DB";
                    button.Click += ButtonLoadImpedanceFormDB_Click;
                    tbl_Function.Controls.Remove(tlp_fillter);
                    tbl_Function.Controls.Add(button, 0, 0);
                    tb_datagridview.Controls.Remove(tb_datagridview);
                    break;
                case "Thermal cycling, Heat soak, Thermal shock":

                    Dictionary<string, string[]> dic = new Dictionary<string, string[]>();
                    dic.Add("Comestic", new[] { "OK", "NG" });
                    dgv_left.columnDropdowns = dic;
                    button.Click += Button_Click;
                    lb_headerTable.Text = "List LogFile";
                    tb_datagridview.ColumnCount = 2;
                    tb_datagridview.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
                    TDMK_Label label = new TDMK_Label() { AutoSize = false, TextAlign = ContentAlignment.MiddleCenter, Text = "Detail", Dock = DockStyle.Fill, Font = new Font("Arial", 12, FontStyle.Bold) };
                    TDMK_Label label2 = new TDMK_Label() { AutoSize = false, TextAlign = ContentAlignment.MiddleCenter, Text = "MiniTable", Dock = DockStyle.Fill, Font = new Font("Arial", 12, FontStyle.Bold) };
                    TDMK_Label label3 = new TDMK_Label() { AutoSize = false, TextAlign = ContentAlignment.MiddleCenter, Text = "Type", Dock = DockStyle.Fill, Font = new Font("Arial", 12, FontStyle.Bold) };
                    dgv_left.CellValueChanged += DGV_LEFT_CHANGE_VALUE_CELL;
                    sc.RowCount = 5;
                    sc.Controls.Add(label, 0, 0);
                    sc.Controls.Add(label2, 0, 3);
                    sc.Controls.Add(dataGridViewz, 0, 4);
                    comboBox.Items.AddRange(new object[] { "All", "Thermal Cycling", "Heat Soak", "Thermal Shock" });
                    tlp_fillter.Controls.Add(label3, 0, 0);
                    tlp_fillter.Controls.Add(comboBox, 1, 0);
                    btn_Disable();
                    sc.Controls.Add(button, 0, 1);
                    sc.Controls.Add(dgv_left, 0, 2);
                    sc.RowStyles.Add(new RowStyle(SizeType.Percent, 5));
                    sc.RowStyles.Add(new RowStyle(SizeType.Percent, 10));
                    sc.RowStyles.Add(new RowStyle(SizeType.Percent, 45));
                    sc.RowStyles.Add(new RowStyle(SizeType.Percent, 5));
                    sc.RowStyles.Add(new RowStyle(SizeType.Percent, 15));
                    dgv_left.Dock = DockStyle.Fill;
                    tb_datagridview.Controls.Add(sc, 1, 0);

                    break;
                case "SEM BSE & Binarization":
                    btn_checkBin.Visible = true;
                    DataTable dtz = new DataTable();
                    splitContainer2.Panel2.Controls.Clear();
                    dgv_Combobox = new CustomDataGridView(dtz, new Dictionary<string, string[]> { { "Judgement", new string[] { "", "Level 1", "Level 2", "Level 3" } } }) { Name = "", Dock = DockStyle.Fill };
                    splitContainer2.Panel2.Controls.Add(dgv_Combobox);
                    lb_headerTable.Text = "Logfile";
                    tb_datagridview.ColumnCount = 2;
                    tb_datagridview.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
                    // Add picture box to the second column
                    pictureBox.BackgroundImage = new Bitmap(10, 20); // Replace with your image path
                    pictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
                    pictureBox.Dock = DockStyle.Fill;
                    pictureBox.BackColor = Color.Green;
                    pictureBox.Padding = new Padding(3);
                    tb_datagridview.Controls.Add(pictureBox, 1, 0);
                    break;
                case "OQC B2B Mating-Unmating":
                    DataTable dt = new DataTable();
                    splitContainer2.Panel2.Controls.Clear();
                    dgv_Combobox = new CustomDataGridView(dt, new Dictionary<string, string[]> { { "FailureMode", new string[] { "", "Level 1", "Level 2", "Level 3" } } }) { Name = "", Dock = DockStyle.Fill };
                    splitContainer2.Panel2.Controls.Add(dgv_Combobox);
                    lb_headerTable.Text = "Logfile";
                    tb_datagridview.ColumnCount = 2;
                    tb_datagridview.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
                    // Add picture box to the second column
                    pictureBox.BackgroundImage = new Bitmap(10, 20); // Replace with your image path
                    pictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
                    pictureBox.Dock = DockStyle.Fill;
                    tb_datagridview.Controls.Add(pictureBox, 1, 0);
                    break;
                case "Bar Code Verification":
                    SetUpBarCodeScreen();
                    break;
                case "Environment en-durance":
                    tb_datagridview.Controls.Clear();
                    tb_datagridview.Controls.Add(UC_env, 0, 0);
                    break;
                default:
                    throw new Exception("Process not found");
            }
        }
        private void ButtonLoadImpedanceFormDB_Click(object sender, EventArgs e)
        {
            Debugger.Break();
        }
        private void Button_Click(object sender, EventArgs e)
        {
            DataGridViewCell selectedCell = dataGridView.SelectedCells[0];
            int rowIndex = selectedCell.RowIndex;
            DataTable dt = (DataTable)dgv_left.DataSource;
            string json = ConverterService.DataTableToJson(dt);
            dataGridView.Rows[rowIndex].Cells["DataLog"].Value = json;
            dgv_left.DataSource = new DataTable();
            btn_Disable();
            CHANGE = !CHANGE;
        }

        private CustomDataGridView dgv_left = new CustomDataGridView(new DataTable(), new Dictionary<string, string[]>());
        private Button btn_SetupSum = new Button() { Text = "Setup CheckSum", Dock = DockStyle.Fill };
        private void SetUpBarCodeScreen()
        {
            btn_checkBin.Visible = true;
            btn_checkBin.Text = "Check SUM";

            browseStatusFile = true;
            lb_headerTable.Text = "Logfile";
            tb_datagridview.ColumnCount = 2;
            tb_datagridview.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            TableLayoutPanel tableLayoutPanel2 = new TableLayoutPanel() { Dock = DockStyle.Fill };
            tableLayoutPanel2.RowCount = 2;
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            TDMK_Label label = new TDMK_Label() { AutoSize = false, TextAlign = ContentAlignment.MiddleCenter, Text = "Check Sum", Dock = DockStyle.Fill, Font = new Font("Arial", 12, FontStyle.Bold) };
            tableLayoutPanel2.Controls.Add(label, 0, 0);
            tableLayoutPanel2.Controls.Add(dgv_left, 0, 1);
            tb_datagridview.Controls.Add(tableLayoutPanel2, 1, 0);
            dgv_left.Dock = DockStyle.Fill;
            ////
            btn_SetupSum.Click += btn_checkSum_Click;
            /////
            TableLayoutPanel tableLayoutPanel = new TableLayoutPanel();
            // Thiết lập số lượng hàng và cột
            tableLayoutPanel.RowCount = 2;
            tableLayoutPanel.ColumnCount = 1; // Chỉ cần 1 cột nếu bạn muốn 2 hàng chiếm toàn bộ chiều rộng

            //// Thiết lập kiểu kích thước hàng
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel.Dock = DockStyle.Fill;
            tableLayoutPanel.Controls.Add(btnGetData, 0, 0);
            tableLayoutPanel.Controls.Add(btn_SetupSum, 0, 1);
            slctn_Location.Panel2.Controls.Add(tableLayoutPanel);
        }

        private bool checkData(string location, string itemCode, string lotNo, string process)
        {
            //Check itemcode lotno validate
            if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo))
            {
                throw new Exception("ItemCode or LotNo is empty");
            }
            ///////////////
            // check file location validate

            if (!browseStatusFile && !Directory.Exists(location))
            {
                throw new Exception("Folder not found");
            }
            if (browseStatusFile && !File.Exists(location))
            {
                throw new Exception("File not found");
            }
            /////////////////
            // check itemcode lot no of filename is match
            string[] fileName = Path.GetFileNameWithoutExtension(location.Replace("\r\n", "").Trim()).Split('-');
            string ItemCode = "", LotNo = "";
            switch (process)
            {
                case "Bar Code Verification":
                    string[] item = fileName[0].Split('_')[0].Split(new[] { "00000B" }, StringSplitOptions.None);
                    ItemCode = item[0];
                    LotNo = item[1];
                    break;
                /// for SEM BSE & Binarization
                case "SEM BSE & Binarization":
                    ItemCode = fileName[1];
                    LotNo = fileName[2];
                    if (!LotNo.Contains("_"))
                    {
                        string no = fileName[3].Split('_')[0];
                        LotNo += '-' + no;
                    }
                    else
                    {
                        LotNo = LotNo.Split('_')[0];
                    }
                    LotNo = ValidateService.lotNoHandle(LotNo);
                    break;
                case "Thermal cycling, Heat soak, Thermal shock":
                    break;
                case "OQC B2B Mating-Unmating":
                    string[] s = fileName;
                    if (s.Count() < 3)
                    {
                        throw new Exception("Not match format location!");
                    }
                    ItemCode = fileName[1];
                    LotNo = fileName[2];
                    if (!LotNo.Contains("_"))
                    {
                        string no = fileName[3].Split('_')[0];
                        LotNo += '-' + no;
                    }
                    else
                    {
                        LotNo = LotNo.Split('_')[0];
                    }
                    LotNo = ValidateService.lotNoHandle(LotNo);
                    //Debugger.Break();
                    break;
                default:
                    throw new Exception("Process not found");
            }
            itemCode = itemCode.Replace(" ", "").Replace("\n", "").Replace("\r", "");
            lotNo = lotNo.Replace(" ", "").Replace("\n", "").Replace("\r", "");
            if (!(itemCode.Equals(ItemCode) && lotNo.Equals(LotNo)))
            {
                throw new Exception("ItemCode or LotNo is not match");
            }
            //////////////////////////
            return true;
        }

        private void MakeGood()
        {
            try
            {
                switch (PROCESS)
                {
                    case "OQC B2B Mating-Unmating":
                        dgv_Combobox.CellFormatting += CellFormatting;
                        dgv_Combobox.CellValueChanged += dataGridView_CellValueChanged;
                        break;
                    case "SEM BSE & Binarization":
                        ((DataGridViewImageColumn)dgv_Combobox.Columns["SEM200250"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                        ((DataGridViewImageColumn)dgv_Combobox.Columns["SEM500700"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                        ((DataGridViewImageColumn)dgv_Combobox.Columns["SEM5K"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                        ((DataGridViewImageColumn)dgv_Combobox.Columns["Binarization200250"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                        ((DataGridViewImageColumn)dgv_Combobox.Columns["Binarization500700"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                        dgv_Combobox.Columns["Black200250"].DefaultCellStyle.Format = "0.00";
                        dgv_Combobox.Columns["Black500700"].DefaultCellStyle.Format = "0.00";
                        dgv_Combobox.CellFormatting += DataGridView1_CellFormatting;
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void checkFactoryCode(string itemCode)
        {
            ///Get EEECode
            TableOfContentService tableOfContentService = new TableOfContentService();
            DataTable dtzs = tableOfContentService.getDataTableByItemCode(itemCode);
            if (dtzs.Rows.Count > 0)
            {
                tb_FactoryCode.Text = dtzs.Rows[0]["FactoryCode"].ToString();
                tb_EEEEECode.Text = dtzs.Rows[0]["EEEECode"].ToString();
            }
            else
            {
                MessageBox.Show("ItemCode chưa được cài TABLE OF CONTENT");
            }
        }
        private void getDataFormFile(string location, string itemCode, string lotNo, string process)
        {

            //check input
            location = location.Replace("\r\n", "").Trim();
            DataTable dt = new DataTable();
            try
            {
                switch (process)
                {
                    case "Environment en-durance":
                        UC_env.dataGridView.RowTemplate.Height = 300;
                        EEDService eEDService = new EEDService();
                        dic_UC_Env.Clear();
                        string[] array = eEDService.ReadFile(location, itemCode, lotNo);
                        foreach (var item in array)
                        {
                            DataTable dtz = eEDService.SolveFolder(item, out string processz);
                            switch (processz)
                            {
                                case "LINERTRUOCKEO":
                                    UC_env.dataGridView.DataSource = dtz;
                                    break;
                                case "LINER":
                                case "PSA":
                                    eEDService.SolveFolderPSALiner(item, dic_UC_Env);
                                    break;
                                default:
                                    throw new Exception($"Folder {item} không phù hợp!");

                            }
                        }
                        UC_env.listBox.Items.Clear();
                        UC_env.listBox.Items.AddRange(dic_UC_Env.Keys.ToArray());
                        UC_env.dictionary_Data = dic_UC_Env;
                        foreach (DataGridViewColumn column in UC_env.dataGridView.Columns)
                        {
                            ((DataGridViewImageColumn)UC_env.dataGridView.Columns[column.Name]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                            ((DataGridViewImageColumn)UC_env.dataGridView.Columns[column.Name]).Width = 500; // Đặt chiều rộng cột
                        }
                        break;
                    case "Thermal cycling, Heat soak, Thermal shock":
                        TCHSTSService tsNew = new TCHSTSService();
                        DataTable dts = tsNew.ReadFile(location, itemCode, lotNo, process);
                        dataGridView.DataSource = dts;

                        break;
                    case "SEM BSE & Binarization":
                        checkData(location, itemCode, lotNo, process);
                        try
                        {
                            dt = semServices.SEMProcessRead(location, itemCode, lotNo, false);
                        }
                        catch (Exception ex)
                        {
                            if (ex.Message.Split('-')[0].Equals("12344"))
                            {
                                DialogResult dr = MessageBox.Show("Dữ liệu đã tồn tại bạn có muốn tiếp tục", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                if (dr == DialogResult.Yes)
                                {
                                    dt = semServices.SEMProcessRead(location, itemCode, lotNo, true);

                                }
                            }
                        }
                        dt = semServices.JudgementCheck(dt);
                        dgv_Combobox = new CustomDataGridView(dt, new Dictionary<string, string[]> { { "Judgement", new string[] { "", "Level 1", "Level 2", "Level 3" } } }) { Name = "", Dock = DockStyle.Fill };
                        dgv_Combobox.CellClick += cellContentClick;
                        MakeGood();
                        splitContainer2.Panel2.Controls.Clear();
                        splitContainer2.Panel2.Controls.Add(dgv_Combobox);

                        break;
                    case "Bar Code Verification":
                        checkData(location, itemCode, lotNo, process);
                        BarCodeVertification barCodeVertification = new BarCodeVertification();
                        dt = barCodeVertification.ReadProcess(location, itemCode, lotNo);
                        if (dt.Rows.Count < 100)
                        {
                            throw new Exception($"Chỉ có {dt.Rows.Count} phần tử hãy import đúng dữ liệu");
                        }
                        dataGridView.DataSource = dt;
                        dataGridView.Columns["ID"].DisplayIndex = 0;
                        checkFactoryCode(itemCode);
                        break;
                    case "OQC B2B Mating-Unmating":
                        checkData(location, itemCode, lotNo, process);
                        OQCB2BMatingUnmatting oqc = new OQCB2BMatingUnmatting();
                        dt = oqc.ProcessRead(location, itemCode, lotNo);
                        dgv_Combobox = new CustomDataGridView(dt, new Dictionary<string, string[]> { { "FailureMode", new string[] { "", "Level 1", "Level 2", "Level 3" } } }) { Name = "", Dock = DockStyle.Fill };
                        dgv_Combobox.CellClick += cellContentClick;
                        splitContainer2.Panel2.Controls.Clear();
                        splitContainer2.Panel2.Controls.Add(dgv_Combobox);
                        MakeGood();
                        break;
                    default:
                        throw new Exception("Process not found");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            MessageBox.Show("Lấy dữ liệu thành công");

        }
        private CustomDataGridView dgv_Combobox = new CustomDataGridView(new DataTable(), new Dictionary<string, string[]>());

        private ToolTip toolTip = new ToolTip();

        #endregion
        private string GetAttribute(DataGridViewCell cell)
        {
            if (cell != null && cell.Value != null)
            {
                return cell.Value.ToString(); // Lấy giá trị ô làm thuộc tính
            }
            return null;
        }
        private void dataGridView1_CellMouseMove(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                DataGridViewCell cell = dgv_Combobox.Rows[e.RowIndex].Cells[e.ColumnIndex];
                toolTip.Show("attribute", dgv_Combobox, dgv_Combobox.PointToClient(Cursor.Position));
            }
        }
        //when mouse is leaving cell
        private void dataGridView1_CellMouseLeave(object sender, DataGridViewCellEventArgs e)
        {
            toolTip.Hide(dgv_Combobox);
        }
        private void dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            switch (PROCESS)
            {
                case "Thermal cycling, Heat soak, Thermal shock":
                    CellClick_TCHSTS(e.RowIndex, e.ColumnIndex);
                    break;
                default:
                    break;
            }
        }
        private bool CHANGE = false;
        private void DGV_LEFT_CHANGE_VALUE_CELL(object sender, DataGridViewCellEventArgs e)
        {
            btn_Enable();
        }
        void btn_Disable()
        {
            button.Enabled = false;
            button.BackColor = Color.Red;
            button.ForeColor = Color.White;
            button.Dock = DockStyle.Fill;
            button.Font = new Font(button.Font.FontFamily, 12);
            button.Text = "Every change is updated";
        }
        void btn_Enable()
        {
            CHANGE = true;
            button.Enabled = true;
            button.BackColor = Color.Green;
            button.ForeColor = Color.White;
            button.Dock = DockStyle.Fill;
            button.Font = new Font(button.Font.FontFamily, 12);
            button.Text = "<< -- UPDATE CHANGE";
        }
        private void CellClick_TCHSTS(int row, int column)
        {
            DataTable dataTable = (DataTable)dataGridView.DataSource;
            string value;
            try
            {
                value = dataTable.Rows[row]["DataLog"].ToString();
            }
            catch
            {
                return;
            }
            DataTable data = ConverterService.JsonToDataTable(value.Split('@')[0]);
            DataTable data1 = ConverterService.JsonToDataTable(value.Split('@')[1]);

            dgv_left.DataSource = data1;

            dataGridViewz.DataSource = data;
            dgv_left.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            DataGridViewColumn columnToMove = dgv_left.Columns["Comestic"];

            // Kiểm tra xem DisplayIndex hiện tại có khác với vị trí mục tiêu không
            if (columnToMove.DisplayIndex != 3)
            {
                int currentDisplayIndex = columnToMove.DisplayIndex;
                columnToMove.DisplayIndex = 3;
            }
            columnToMove = dgv_left.Columns["Function test"];
            if (columnToMove.DisplayIndex != 4)
            {
                int currentDisplayIndex = columnToMove.DisplayIndex;
                columnToMove.DisplayIndex = 4;
            }
            columnToMove = dgv_left.Columns["ID"];
            if (columnToMove.DisplayIndex != 0)
            {
                int currentDisplayIndex = columnToMove.DisplayIndex;
                columnToMove.DisplayIndex = 00;
            }
            columnToMove = dgv_left.Columns["Content"];
            if (columnToMove.DisplayIndex != 1)
            {
                int currentDisplayIndex = columnToMove.DisplayIndex;
                columnToMove.DisplayIndex = 1;
            }
            //Debugger.Break();
        }
        private void btn_Export_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                string itemCode = tb_ItemCode.Text.ToString().Trim();
                string lotNo = tb_Lotno.Text.ToString().Trim();
                switch (PROCESS)
                {
                    case "Thermal cycling, Heat soak, Thermal shock":
                        if (MessageBox.Show("Bạn muốn export logfile?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            string type = comboBox.Text.ToString().Trim();
                            TCHSTSService service = new TCHSTSService();
                            service.ExportLogFile(itemCode, lotNo, type);
                            MessageBox.Show("Thành Công");
                        }
                        break;
                    default:
                        break;
                }
                #region HIDE
                COUNTING++;
                COUNTING = COUNTING % 7;
                if (COUNTING > 5)
                {
                    MessageBox.Show("Đừng thao tác quá nhiều vào nút này :<");
                }
                #endregion
            }


        }
        private int COUNTING = 0;
        private void dataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataTable dataTable = (DataTable)dgv_Combobox.DataSource;
                // Lấy DataRow của hàng đã thay đổi
                DataRow changedRow = dataTable.Rows[e.RowIndex];
                switch (PROCESS)
                {
                    case "OQC B2B Mating-Unmating":
                        changedRow["Judgement"] = OQCB2BMatingUnmatting.checkARow(changedRow) ? "OK" : "NG";
                        break;
                    case "":
                    default: break;
                }
                dgv_Combobox.Refresh();
            }
        }
    }
}
