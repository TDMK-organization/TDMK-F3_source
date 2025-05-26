using Funtion_F3_SMT;
using OK2SHIP_SMT.Services;
using OK2SHIP_SMT.Views;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Security.Authentication;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.UserControls
{
    public partial class Packaging : UserControl
    {
        private CustomDataGridView _datagridview = new CustomDataGridView() { Dock = DockStyle.Fill };
        private Dictionary<string, Image> _image = new Dictionary<string, Image>();
        private Dictionary<string, KeyValuePair<Dictionary<string, Image>, string>> _dic = new Dictionary<string, KeyValuePair<Dictionary<string, Image>, string>>();
        private bool CHANGING = false;

        public Packaging()
        {
            InitializeComponent();
            setUp();
        }
        #region Action
        private void setUp()
        {
            textBox.DoubleClick += tb_ItemCodePack_DoubleClick;
            UpdateStage();
            areaDG.Controls.Add(_datagridview, 1, 1);
            initDataGirdView();
            pictureBox1.DoubleClick += pictureBox_DoubleClick;
            pictureBox1.Click += pictureBox_Click;
            pictureBox2.DoubleClick += pictureBox_DoubleClick;
            pictureBox2.Click += pictureBox_Click;
            pictureBox3.DoubleClick += pictureBox_DoubleClick;
            pictureBox3.Click += pictureBox_Click;
            pictureBox4.DoubleClick += pictureBox_DoubleClick;
            pictureBox4.Click += pictureBox_Click;
            pictureBox5.DoubleClick += pictureBox_DoubleClick;
            pictureBox5.Click += pictureBox_Click;
            pictureBox6.DoubleClick += pictureBox_DoubleClick;
            pictureBox6.Click += pictureBox_Click;
        }
        private void UpdateStage()
        {
            if (UserSession.Instance.IsLoggedIn)
            {
                tb_operator.Text = UserSession.Instance.User_ID;
                btn_Login.Visible = false;
                btn_Logout.Visible = true;
                lb_name.Text = $"Hello, {UserSession.Instance.Username}!";
            }
            else
            {
                tb_operator.Text = "Hello, Guest";
                btn_Login.Visible = true;
                btn_Logout.Visible = false;
                lb_name.Text = $"Hello, Guest!";
            }
        }
        private void resetChoice()
        {
            pictureBox1.BackColor = Color.Transparent;
            pictureBox2.BackColor = Color.Transparent;
            pictureBox3.BackColor = Color.Transparent;
            pictureBox4.BackColor = Color.Transparent;
            pictureBox5.BackColor = Color.Transparent;
            pictureBox6.BackColor = Color.Transparent;
        }
        private void initDataGirdView()
        {
            #region Create datagridview

            Dictionary<string, string[]> data = new Dictionary<string, string[]>();
            data.Add("Result", new string[] { "OK", "NG" });

            _datagridview.columnDropdowns = data;
            _datagridview.DataSource = PackagingService.createDataVote();

            #endregion

        }
        private void PictureClear()
        {
            pictureBox.Image = null;
            pictureBox1.Image = null;
            pictureBox2.Image = null;
            pictureBox3.Image = null;
            pictureBox4.Image = null;
            pictureBox5.Image = null;
            pictureBox6.Image = null;
        }
        private void ResetData()
        {
            resetChoice();
            _dic = new Dictionary<string, KeyValuePair<Dictionary<string, Image>, string>>();
            _image = new Dictionary<string, Image>();
            PictureClear();
            initDataGirdView();
            tb_shippingto.Text = "";
            tb_traycode.Text = "";
        }

        private void SaveAction()
        {
            string itemCode = tb_ItemCodePack.Text.Trim();
            int res = 0;
            try
            {
                res = new PackagingService().Save(itemCode, _dic, false);
            }
            catch (AuthenticationException exz)
            {
                MessageBox.Show(exz.Message);
                Login login = new Login();
                login.ShowDialog();
                return;
            }
            catch (Exception ex)
            {
                if (ex.Message.Split('-')[0].Trim().Equals("1234"))
                {
                    if (MessageBox.Show("Dữ liệu đã có bạn có muốn lưu đè?", "thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        res = new PackagingService().Save(itemCode, _dic, true);
                    }
                }
                else
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
            if (res == 0)
            {
                MessageBox.Show("Lưu thất bại");
            }
            else
            {
                ResetData();
                MessageBox.Show($"Lưu Thành Công!");
                listBox.Items.Clear();
            }
        }

        #endregion
        #region Event
        private void btn_clear_Click(object sender, EventArgs e)
        {
            ResetData();
        }
        private void btn_Login_Click(object sender, EventArgs e)
        {
            Login login = new Login();
            login.ShowDialog();
            UpdateStage();
        }
        private void btn_Logout_Click(object sender, EventArgs e)
        {
            UserSession.Instance.Logout();
            UpdateStage();
        }
        private void pictureBox_DoubleClick(object sender, EventArgs e)
        {
            string name = ((PictureBox)sender).Name.ToString();
            if (_image.TryGetValue($"{name[10]}", out Image image))
            {
                using (ArtistFramerService edit = new ArtistFramerService((Image)image))
                {
                    _image[$"{name[10]}"] = edit.Image;
                    FillImage();
                }

            }
            else
            {
                OpenFileDialog open = new OpenFileDialog();
                open.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp";
                if (open.ShowDialog() == DialogResult.OK)
                {
                    Image img = Image.FromFile(open.FileName);
                    _image[$"{name[10]}"] = img;
                    FillImage();
                }
            }


        }
        private void FillImage()
        {
            if (_image.TryGetValue("1", out Image image))
            {
                pictureBox1.Image = image;
            }
            if (_image.TryGetValue("2", out image))
            {
                pictureBox2.Image = image;
            }
            if (_image.TryGetValue("3", out image))
            {
                pictureBox3.Image = image;
            }
            if (_image.TryGetValue("4", out image))
            {
                pictureBox4.Image = image;
            }
            if (_image.TryGetValue("5", out image))
            {
                pictureBox5.Image = image;
            }
            if (_image.TryGetValue("6", out image))
            {
                pictureBox6.Image = image;
            }

        }
        private void btn_save_Click(object sender, EventArgs e)
        {
            SaveAction();
        }
        #endregion

        private void tableLayoutPanel6_Paint(object sender, PaintEventArgs e)
        {

        }

        private void pictureBox_Click(object sender, EventArgs e)
        {

        }
        private void updateListBox()
        {
            listBox.Items.Clear();
            listBox.Items.AddRange(_dic.Keys.ToArray());
        }
        private void tdmK_Button1_Click(object sender, EventArgs e)
        {

            listBox.Items.Clear();
            string shippingTo = tb_shippingto.Text.Trim();
            string trayCode = tb_traycode.Text.Trim();
            if (string.IsNullOrEmpty(shippingTo) || string.IsNullOrEmpty(trayCode))
            {
                MessageBox.Show("Vui lòng nhập Shipping To và Tray Code");
                return;
            }
            //if (_dic.ContainsKey($"{shippingTo}_{trayCode}"))
            //{
            //    if (MessageBox.Show("Dữ liệu đã có bạn có muốn update dữ liệu", "Thông báo", MessageBoxButtons.YesNo) != DialogResult.Yes)
            //    {
            //        return;
            //    }
            //}
            DataTable dataTable = (DataTable)_datagridview.DataSource;
            string json = ConverterService.DataTableToJson(dataTable);
            KeyValuePair<Dictionary<string, Image>, string> data = new KeyValuePair<Dictionary<string, Image>, string>(_image, json);
            if (_dic.TryGetValue($"{shippingTo}_{trayCode}", out var z))
            {
                _dic[$"{shippingTo}_{trayCode}"] = data;
            }
            else
            {
                _dic.Add($"{shippingTo}_{trayCode}", data);
            }
            updateListBox();
            tb_shippingto.Text = null;
            tb_traycode.Text = null;
            initDataGirdView();
            _image = new Dictionary<string, Image>();
            PictureClear();
        }

        private void btn_load_Click(object sender, EventArgs e)
        {
            listBox.Items.Clear();
            PictureClear();
            ResetData();
            string st = tb_ItemCodePack.Text;
            if (string.IsNullOrEmpty(st))
            {
                MessageBox.Show("Vui lòng nhập ItemCode");
                return;
            }
            try
            {
                if (CHANGING)
                {
                    PackagingService service = new PackagingService();
                    string location = textBox.Text;
                    string itemCode = tb_ItemCodePack.Text.Trim();
                    _dic = service.GetData(location, itemCode);
                    listBox.Items.AddRange(_dic.Keys.ToArray());
                    MessageBox.Show("Get data successfully!");
                }
                else
                {
                    PackagingService service = new PackagingService();
                    _dic = service.LoadDictionary(st);
                    listBox.Items.AddRange(_dic.Keys.ToArray());
                    MessageBox.Show("Load data successfully!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void tdmK_Button2_Click(object sender, EventArgs e)
        {
            string shippingTo = tb_shippingto.Text.Trim();
            string trayCode = tb_traycode.Text.Trim();
            if (string.IsNullOrEmpty(shippingTo) || string.IsNullOrEmpty(trayCode))
            {
                MessageBox.Show("Vui lòng nhập Shipping To và Tray Code");
                return;
            }
            if (_dic.ContainsKey($"{shippingTo}_{trayCode}"))
            {
                if (MessageBox.Show("Bạn có muốn xóa dữ liệu", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    _dic.Remove($"{shippingTo}_{trayCode}");
                }
            }
            updateListBox();
            tb_shippingto.Text = null;
            tb_traycode.Text = null;
            initDataGirdView();
            _image = new Dictionary<string, Image>();
            PictureClear();
        }

        private void listBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            _image = new Dictionary<string, Image>();
            PictureClear();
            if (listBox.SelectedItem == null)
                return;

            string selectedKey = listBox.SelectedItem.ToString().Trim();
            if (_dic.TryGetValue(selectedKey, out KeyValuePair<Dictionary<string, Image>, string> value))
            {
                string str1 = selectedKey.Split('_')[0];
                string str2 = selectedKey.Split('_')[1];
                tb_shippingto.Text = str1;
                tb_traycode.Text = str2;
                string json = value.Value;
                Dictionary<string, Image> dic = value.Key;
                DataTable dataTable = ConverterService.JsonToDataTable(json);
                _datagridview.DataSource = dataTable;
                _image = dic;
                FillImage();
            }
        }

        private void btn_Export_Click(object sender, EventArgs e)
        {
            string itemCode = tb_ItemCodePack.Text.Trim();
            PackagingService packagingService = new PackagingService();
            try
            {
                packagingService.ExportToExcel(itemCode, false);
            }
            catch (AuthenticationException exz)
            {
                MessageBox.Show(exz.Message);
                Login login = new Login();
                login.ShowDialog();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            MessageBox.Show("Export sucssefully!");
        }

        private void btn_GetData_Click(object sender, EventArgs e)
        {

        }

        private void tb_ItemCodePack_TextChanged(object sender, EventArgs e)
        {
            string itemCode = tb_ItemCodePack.Text.Trim();
            if (itemCode.Contains('\\'))
            {
                groupBox1.Text = "Location";
                btn_load.Text = "Get Data";
            }
        }

        private void tb_ItemCodePack_DoubleClick(object sender, EventArgs e)
        {
            if (CHANGING)
            {
                FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
                folderBrowserDialog.Description = "Select a folder";
                folderBrowserDialog.ShowDialog();
                textBox.Text = folderBrowserDialog.SelectedPath;
            }
        }

        private void tableLayoutPanel3_Paint(object sender, PaintEventArgs e)
        {

        }
        private TextBox textBox = new TextBox() { Dock = DockStyle.Fill, Multiline = true };
        private GroupBox groupBox = new GroupBox() { Text = "Location", Dock = DockStyle.Fill };
        private void tdmK_Button3_Click(object sender, EventArgs e)
        {
            CHANGING = !CHANGING;
            if (CHANGING)
            {
                btn_load.Text = "Get Data";
                tableLayoutPanel8.ColumnCount = 2;
                tableLayoutPanel8.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
                groupBox.Controls.Add(textBox);
                tableLayoutPanel8.Controls.Add(groupBox, 1, 0);
            }
            else
            {
                tableLayoutPanel8.ColumnCount = 1;
                tableLayoutPanel8.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

                btn_load.Text = "Load Data";
            }
        }


    }
}
