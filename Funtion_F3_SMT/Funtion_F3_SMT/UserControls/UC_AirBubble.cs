using Funtion_F3_SMT;
using OK2SHIP_SMT.Repositories;
using OK2SHIP_SMT.Services;
using OK2SHIP_SMT.ToolBoxs;
using OK2SHIP_SMT.UserControls.Logins;
using OK2SHIP_SMT.Views;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZedGraph;

namespace OK2SHIP_SMT.UserControls
{
    public partial class UC_AirBubble : UserControl
    {
        private bool _PRIME = false;
        AirBubbleService _service = new AirBubbleService();
        PeelTestOnProductService _service2 = new PeelTestOnProductService();
        private DataGridView dgv = new DataGridView() { Dock = DockStyle.Fill };
        private CustomDataGridView dgv_custom = new CustomDataGridView(new DataTable(), new Dictionary<string, string[]>()) { Dock = DockStyle.Fill };
        public UC_AirBubble(bool prime = false)
        {
            _PRIME = prime;
            InitializeComponent();
            CheckStatus();
            if (_PRIME)
            {
                tdmK_Label8.Text = "Peel Test \n (On Product)";
                _service2 = new PeelTestOnProductService();
                tableLayoutPanel10.Controls.Clear();
                tableLayoutPanel10.Controls.Add(dgv, 0, 1);
                dgv.CellValueChanged += dataGridView1_CellValueChanged;
                tableLayoutPanel11.Controls.Clear();
                tableLayoutPanel11.Controls.Add(dgv_custom, 0, 0);

            }
            else
            {
                _service = new AirBubbleService();
                tableLayoutPanel11.Controls.Clear();
                tableLayoutPanel11.Controls.Add(dgv_custom, 0, 0);
            }
        }

        #region Event
        private void checkMeansuar(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dgvSpec = (DataGridView)sender;
            if (dgvSpec.Columns[e.ColumnIndex].Name.Contains("Measure(mm)"))
            {
                try
                {
                    // Debugger.Break();
                    if (double.TryParse(dgvSpec.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(), out double result))
                    {
                        if (result != 0)
                        {
                            dgvSpec.Rows[e.RowIndex].Cells["Adhesive"].Value = "Adhesive Squeeze Out";
                        }
                        else
                        {
                            dgvSpec.Rows[e.RowIndex].Cells["Adhesive"].Value = "No Adhesive Squeeze Out";
                        }
                    }
                }
                catch
                {

                }
            }
        }
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //Debugger.Break();
            DataGridView dgvSpec = (DataGridView)sender;
            if (dgvSpec.Columns[e.ColumnIndex].Name.Contains("(N)") || dgvSpec.Columns[e.ColumnIndex].Name.Contains("(Gf)"))
            {
                string value = dgvSpec[e.ColumnIndex, e.RowIndex].Value.ToString().Replace(" ", "");
                if (value.Contains('~'))
                {
                    return;
                }
                int i = 0;
                while (i < value.Length)
                {
                    try
                    {
                        if (value[i] == '.')
                        {

                        }
                        else
                        {
                            ValidateService.isDigit(value[i].ToString());
                        }
                    }
                    catch
                    {
                        break;
                    }
                    i++;
                }
                if (i != value.Length)
                {
                    string[] c = value.Split(value[i]);
                    string result = $"{Math.Round(double.Parse(c[0]), 2)} ~ {Math.Round(double.Parse(c[1]), 2)}";
                    dgvSpec[e.ColumnIndex, e.RowIndex].Value = result;
                    // convert 
                    if (dgvSpec.Columns[e.ColumnIndex].Name.Contains("Peak(N)"))
                    {
                        string num1 = Math.Round(double.Parse(c[0]) * 101.97, 2).ToString();
                        string num2 = Math.Round(double.Parse(c[1]) * 101.97, 2).ToString();

                        dgvSpec[7, e.RowIndex].Value = $"{num1} ~ {num2}";

                    }
                    if (dgvSpec.Columns[e.ColumnIndex].Name.Contains("Average(N)"))
                    {
                        string num1 = Math.Round(double.Parse(c[0]) * 101.97, 2).ToString();
                        string num2 = Math.Round(double.Parse(c[1]) * 101.97, 2).ToString();

                        dgvSpec[8, e.RowIndex].Value = $"{num1} ~ {num2}";

                    }
                    if (dgvSpec.Columns[e.ColumnIndex].Name.Contains("Average(Gf)"))
                    {
                        string num1 = Math.Round(double.Parse(c[0]) * 0.0098, 2).ToString();
                        string num2 = Math.Round(double.Parse(c[1]) * 0.0098, 2).ToString();

                        dgvSpec[6, e.RowIndex].Value = $"{num1} ~ {num2}";

                    }
                    if (dgvSpec.Columns[e.ColumnIndex].Name.Contains("Peak(Gf)"))
                    {
                        string num1 = Math.Round(double.Parse(c[0]) * 0.0098, 2).ToString();
                        string num2 = Math.Round(double.Parse(c[1]) * 0.0098, 2).ToString();

                        dgvSpec[5, e.RowIndex].Value = $"{num1} ~ {num2}";

                    }
                    foreach (string key in _service2._DIC.Keys)
                    {
                        string tape = dgvSpec[3, e.RowIndex].Value.ToString();
                        string type = dgvSpec[4, e.RowIndex].Value.ToString();
                        if (!key.Contains("Refer") && key.Contains(tape) && key.Contains(type))
                        {
                            _service2.CheckSpec(key);
                        }

                    }
                }
            }
        }
        private void btn_MakeAirBubble_Click(object sender, EventArgs e)
        {
            CommentItemCodeAirBubble cm = new CommentItemCodeAirBubble();
            cm.ShowDialog();
        }
        private void tdmK_Button3_Click(object sender, EventArgs e)
        {
            if (_PRIME)
            {
                _service2 = new PeelTestOnProductService(txt_ItemCode.Text, txt_lotNo.Text);
                _service2.LoadData(Legacy.Checked, true);
            }
            else
            {
                _service = new AirBubbleService(txt_ItemCode.Text, txt_lotNo.Text);
                _service.LoadDataRefer();

            }
            FillData();
        }

        private void tdmK_Button1_Click_1(object sender, EventArgs e)
        {
            if (_PRIME)
            {
                try
                {
                    _service2 = new PeelTestOnProductService(txt_ItemCode.Text, txt_lotNo.Text);
                    _service2.Export(Legacy.Checked);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Export AirBubble: {ex.Message}");
                }
            }
            else
            {
                try
                {

                    _service = new AirBubbleService(txt_ItemCode.Text, txt_lotNo.Text);
                    _service.Export();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Export AirBubble: {ex.Message}");
                }

            }

            FillData();
        }
        private void btn_setupRefer_Click(object sender, EventArgs e)
        {
            try
            {

                _service = new AirBubbleService(txt_ItemCode.Text.Trim(), txt_lotNo.Text.Trim());


                ReferForm referForm = new ReferForm(_service);
                referForm.ShowDialog();
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message);
            }
        }
        private void CheckNGAllData()
        {

        }
        private void btn_load_Click(object sender, EventArgs e)
        {
            loadData(txt_ItemCode.Text, txt_lotNo.Text, Legacy.Checked);
            checkingDataGridView(new DataGridView());

        }
        private void txt_lotNo_Leave(object sender, EventArgs e)
        {
            txt_lotNo.Text = ValidateService.lotNoHandle(txt_lotNo.Text);
        }


        private void btn_SaveData_Click(object sender, EventArgs e)
        {
            try
            {
                if (_PRIME)
                {
                    if (!_service2.checkNG())
                    {
                        if (MessageBox.Show("Dữ liệu NG bạn có muốn lưu không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.No)
                        {
                            return;
                        }
                    }
                    _service2.SaveData();
                }
                else
                {
                    _service.SaveData();
                }
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                if (msg.Split('-')[0].Contains("1402"))
                {
                    try
                    {
                        if (MessageBox.Show(msg.Split('-')[1].Trim(), "AIRBUBBLE", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {

                            if (_PRIME)
                            {
                                _service2.SaveData(true);
                            }
                            else
                            {
                                _service.SaveData(null, true);
                            }
                        }
                    }
                    catch (Exception ex2)
                    {
                        MessageBox.Show(ex2.Message, "AIR_BUBBLE");
                    }
                }
                else
                {
                    MessageBox.Show(msg);
                }
            }
        }
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string z = listBox1.Items[listBox1.SelectedIndex].ToString();
                PickDataTable(z);
            }
            catch
            {

            }
        }
        private void tdmK_Button1_Click(object sender, EventArgs e)
        {
            GetData(txt_ItemCode.Text, txt_lotNo.Text, tb_location.Text, tb_PIDLocation.Text);
            checkingDataGridView(new DataGridView());


        }

        private void checkingDataGridView(DataGridView dt)
        {
            bool? primeNG = false;
            primeERROR.Text = "";
            primeERROR.BackColor = SystemColors.Control;
            foreach (DataGridViewRow row in dt.Rows)
            {
                if (_PRIME)
                {

                    if (!row.Cells["JudgementForce"].Value.ToString().Equals("OK"))
                    {
                        row.Cells["JudgementForce"].Style.BackColor = Color.Red;
                        primeNG = true;
                    }
                }
                else
                {

                    if (!row.Cells["Judgement"].Value.ToString().Trim().Equals("OK"))
                    {
                        row.Cells["Judgement"].Style.BackColor = Color.Red;
                        primeNG = true;
                    }
                }
            }
            if (primeNG == false)
            {
                primeERROR.Text = "OK";
                primeERROR.BackColor = Color.Green;
            }
            if (primeNG == true)
            {
                primeERROR.Text = "NG";
                primeERROR.BackColor = Color.Red;
            }
        }
        private void tb_location_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(tb_location.Text))
                {
                    btn_getData.Enabled = false;
                }
                else
                {
                    //"G:\\GD 2 F3\\DATA 3.1\\DATA 3.1\\9.Air\\5CCEV-720211-00012-CPL5-MIYAGI-JPOND-DO KHI-78888"
                    btn_getData.Enabled = true;
                    string folderName = tb_location.Text.Split('\\')[tb_location.Text.Split('\\').Count() - 1];
                    KeyValuePair<string, string> dic = new KeyValuePair<string, string>();
                    if (_PRIME)
                    {
                        dic = _service2.getItemCodeLotNo(folderName);

                    }
                    else
                    {
                        dic = _service.getItemCodeLotNo(folderName);
                    }
                    txt_ItemCode.Text = dic.Key;
                    txt_lotNo.Text = dic.Value;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, _NAME);
            }
        }
        private const string _NAME = "Airbubble";
        private void btn_login_Click(object sender, EventArgs e)
        {
            if (UserSession.Instance.IsLoggedIn)
            {
                UserSession.Instance.Logout();
            }
            else
            {
                Form loginForm = new Login();
                loginForm.ShowDialog();
            }
            CheckStatus();
        }

        #endregion
        #region Action
        private void PickDataTable(string key)
        {
            if (_PRIME)
            {
                Dictionary<string, string[]> dicZ = new Dictionary<string, string[]>();
                string[] s = new[] { "" };
                if (key.Contains("LINER"))
                {
                    s = new[] { "No adhesive stick on liner" };

                }
                if (key.Contains("PSA"))
                {

                    s = new[] { "PSA go with reingorcement" };
                }
                dicZ.Add("JudgementMode", s);
                //Debugger.Break();
                if (key.Contains("REFER"))
                {
                    dgv_Before.DataSource = _service2._BEFOREIMAGE[$"{key.Split('(')[1].Replace(")", "").Replace("-", " - ")}"];

                }
                else
                {
                    dgv_Before.DataSource = _service2._BEFOREIMAGE[$"{_service2._ITEMCODE} - {_service2._LOTNO}"];
                }
                FillID(dgv_Before);
                dgv_custom = new CustomDataGridView(_service2._DIC[key], dicZ);

                tableLayoutPanel11.Controls.Clear();
                tableLayoutPanel11.Controls.Add(dgv_custom, 0, 0);
                Makeup(dgv_custom);

                FillID(dgv_custom);

            }
            else
            {
                Dictionary<string, string[]> dicZ = new Dictionary<string, string[]>();

                dicZ.Add("Measure(mm)", new[] { "0" });
                dicZ.Add("Adhesive", new[] { "No Adhesive Squeeze Out" });
                dgv_custom = new CustomDataGridView(_service._dic[key], dicZ);
                tableLayoutPanel11.Controls.Clear();
                tableLayoutPanel11.Controls.Add(dgv_custom, 0, 0);
                FillID(dgv_custom);
                Makeup(dgv_custom);
                dgv_custom.CellValueChanged += checkMeansuar;
                string itemCode = "";
                if (key.Contains("Refer"))
                {
                    itemCode = key.Split(':')[1].Replace(" ", "").Replace("-", " - ");
                }
                else
                {
                    itemCode = $"{_service._itemCode} - {_service._lotNo}";
                }
                if (_service._beforeImage.TryGetValue(itemCode, out DataTable beforeData))
                {
                    dgv_Before.DataSource = beforeData;
                    Makeup(dgv_Before);
                    FillID(dgv_Before);
                }
            }
            try
            {

                pictureBox1.Image = TDMK_ImageConverter.ByteArrayToImage(_service._dicTONG[key].Value);
                tdmK_Label10.Text = $"{_service._dicTONG[key].Key} mm2";
            }
            catch
            {
                pictureBox1.Image = null;
                tdmK_Label10.Text = "NA mm2";
            }

            checkingDataGridView(dgv_custom);
        }
        private void FillID(DataGridView dgv)
        {
            if (dgv == null || dgv.Columns == null)
                return;

            // Tìm cột "ID" (không phân biệt hoa thường)
            DataGridViewColumn idColumn = null;
            foreach (DataGridViewColumn col in dgv.Columns)
            {
                if (string.Equals(col.Name, "ID", StringComparison.OrdinalIgnoreCase))
                {
                    idColumn = col;
                    break;
                }
            }

            if (idColumn == null)
                return;

            int idColIndex = idColumn.Index;
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                if (!dgv.Rows[i].IsNewRow) // Bỏ qua dòng mới (nếu có)
                {
                    dgv.Rows[i].Cells[idColIndex].Value = (i + 1).ToString();
                }
            }
        }
        private void FillData()
        {

            dgv_Before.DataSource = new DataTable();
            pictureBox1.Image = new Bitmap(1, 1);
            tdmK_Label10.Text = "NA";
            dgv_custom.DataSource = new DataTable();
            listBox1.Items.Clear();
            if (_PRIME)
            {
                listBox1.Items.AddRange(_service2._DIC.Keys.ToArray());
                if (_service2._BEFOREIMAGE.TryGetValue($"{_service2._ITEMCODE} - {_service2._LOTNO}", out DataTable dz))
                {
                    dgv_Before.DataSource = dz;
                }
                txt_ItemCode.Text = _service2._ITEMCODE;
                txt_lotNo.Text = _service2._LOTNO;
                try
                {

                    dgv.DataSource = _service2._SPEC[$"{_service2._ITEMCODE} - {_service2._LOTNO}"];
                }
                catch
                {

                }

            }
            else
            {
                txt_ItemCode.Text = _service._itemCode;
                txt_lotNo.Text = _service._lotNo;
                listBox1.Items.AddRange(_service._dic.Keys.ToArray());
                dgv_Before.DataSource = _service._beforeImage;
            }
            Makeup(dgv_Before);
        }
        private void Makeup(DataGridView dgv)
        {

            foreach (DataGridViewColumn col in dgv.Columns)
            {
                if (!col.Name.Contains("JudgementMode") && !col.Name.Contains("Measure(mm)") && !col.Name.Contains("Adhesive"))
                {
                    ((DataGridViewColumn)col).ReadOnly = true;
                }
                if (col is DataGridViewImageColumn)
                {
                    ((DataGridViewImageColumn)col).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)col).Width = 150;
                }
            }
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.Height = 100;
            }
        }
        private void GetData(string itemCode, string lotNo, string location, string pid)
        {
            try
            {
                if (!_PRIME)
                {

                    _service = new AirBubbleService(itemCode, lotNo);
                    _service.ReadData(location, pid);
                }
                else
                {
                    _service2 = new PeelTestOnProductService(itemCode, lotNo);
                    _service2.ReadData(location, pid);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, _NAME);
            }
            FillData();
        }
        private void CheckStatus()
        {
            if (UserSession.Instance.IsLoggedIn)
            {
                btn_login.Text = "LogOut";
                tb_Operator.Text = UserSession.Instance.User_ID;
            }
            else
            {
                btn_login.Text = "Login";
                tb_Operator.Text = "";
            }
        }

        private void loadData(string itemCode, string lotNo, bool legacy)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo))
            {
                MessageBox.Show("Item Code and Lot No cannot be empty.", _NAME);
                return;
            }
            if (_PRIME)
            {
                _service2 = new PeelTestOnProductService(itemCode, lotNo);
                _service2.LoadData(legacy);
            }
            else
            {

                _service = new AirBubbleService(itemCode, lotNo);
                _service.LoadData();
            }
            FillData();
        }







        #endregion


    }




}