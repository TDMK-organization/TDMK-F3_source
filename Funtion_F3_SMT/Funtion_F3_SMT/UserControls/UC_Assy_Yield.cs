using Funtion_F3_SMT;
using OK2SHIP_SMT.Services;
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
    public partial class UC_Assy_Yield : UserControl
    {
        private Image imageMAIN = null;
        private Dictionary<string, DataTable> _DIC = new Dictionary<string, DataTable>();
        public UC_Assy_Yield()
        {
            InitializeComponent();
            Inti();
        }
        #region Event
        private void textBox2_Leave(object sender, EventArgs e)
        {
            textBox2.Text = ValidateService.lotNoHandle(textBox2.Text.Trim());
        }
        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            PackData();
            string sa = comboBox.Text;
            if (double.TryParse(textBox4.Text, out double i))
            {
                DataTable dataTable = _DIC["Process"];
                foreach (DataRow item in dataTable.Rows)
                {
                    if (item["Station"].Equals(sa))
                    {
                        if (double.TryParse(item["Input"].ToString(), out double z))
                        {
                            if (z != 0)
                            {
                                double res = (i / z) * 100;
                                textBox6.Text = $"{Math.Round(res, 2)}" + "%";
                            }
                        }
                    }
                }
            }
        }
        private void tdmK_Button1_Click(object sender, EventArgs e)
        {
            LoadDataProcess();
            MakeColor();
            
        }

        private void tdmK_Button4_Click(object sender, EventArgs e)
        {
            // export
            string itemCode = textBox1.Text.Trim();
            string lotNo = textBox2.Text.Trim();
            try
            {
                new ASSY_YIELDService().Export(itemCode, lotNo);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void tdmK_Button7_Click(object sender, EventArgs e)
        {
            string itemCode = textBox1.Text.Trim();
            string lotNo = textBox2.Text.Trim();
            //save
            try
            {
                PackData();
                int a = new ASSY_YIELDService().Save(itemCode, lotNo, _DIC, Guid.Parse("00000000-0000-0000-0000-000000000000"));
                MessageBox.Show($"Lưu thành công {a} row");
                ClearData();
            }
            catch (AuthenticationException ex)
            {
                MessageBox.Show("Hãy đăng nhập ngay!");
            }
            catch (Exception ex)
            {
                if (ex.Message.Split('&')[0].Contains("Area"))
                {
                    string guid = ex.Message.Split('&')[0].Trim().Split('=')[1].Trim();
                    if (MessageBox.Show($"Đã có dữ liệu bạn có muốn ghi đè!", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        int a = new ASSY_YIELDService().Save(itemCode, lotNo, _DIC, Guid.Parse(guid));
                        MessageBox.Show($"Lưu thành công {a} row");
                        ClearData();
                    }
                    return;
                }
                MessageBox.Show(ex.Message);
            }
        }



        private void tdmK_Button6_Click(object sender, EventArgs e)
        {
            LoginProcess();
        }


        private void ClearForm()
        {
            comboBox.SelectedItem = null;
            pictureBox1.Image = null;
            textBox4.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox5.Text = "";
        }
        //Create button
        private void tdmK_Button2_Click(object sender, EventArgs e)
        {
            string value = comboBox.Text.ToString();
            if (!string.IsNullOrEmpty(value))
            {
                DataTable dataTable = new DataTable();
                switch (value)
                {
                    case "SMT":
                    case "X-Ray":
                        dataTable = (DataTable)dataGridView1.DataSource;
                        break;
                    case "ICT/FCT":
                    case "FQC":
                        dataTable = (DataTable)dataGridView2.DataSource;
                        break;
                    case "OQC":
                    case "Packaging":
                        dataTable = (DataTable)dataGridView3.DataSource;
                        break;
                }
                DataRow row = dataTable.NewRow();
                row["Station"] = value;
                row["Defect Description"] = imageMAIN;
                row["Defect Name"] = string.IsNullOrEmpty(textBox5.Text) ? "" : textBox5.Text;
                row["Defects Qty"] = string.IsNullOrEmpty(textBox4.Text) ? "0" : textBox4.Text;
                row["Defect Rate"] = string.IsNullOrEmpty(textBox6.Text) ? "0" : textBox6.Text;
                row["Root Cause"] = string.IsNullOrEmpty(textBox7.Text) ? "0" : textBox7.Text;
                row["Corrective Action"] = string.IsNullOrEmpty(textBox8.Text) ? "0" : textBox8.Text;
                dataTable.Rows.Add(row);
                switch (value)
                {
                    case "SMT":
                    case "X-Ray":
                        dataGridView1.DataSource = dataTable;
                        break;
                    case "ICT/FCT":
                    case "FQC":
                        dataGridView2.DataSource = dataTable;
                        break;
                    case "OQC":
                    case "Packaging":
                        dataGridView3.DataSource = dataTable;
                        break;
                }
            }
            ClearForm();
        }
        private void tdmK_Button5_Click(object sender, EventArgs e)
        {
            ClearData();
        }
        private void dataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView.DataSource = CaculateDataTable((DataTable)dataGridView.DataSource);
        }
        private void pictureBox1_DoubleClick(object sender, EventArgs e)
        {
            if (pictureBox1.Image == null)
            {
                OpenFileDialog open = new OpenFileDialog();
                open.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp";
                if (open.ShowDialog() == DialogResult.OK)
                {
                    imageMAIN = Image.FromFile(open.FileName);
                    pictureBox1.Image = imageMAIN;
                }
            }
            else
            {
                using (ArtistFramerService edit = new ArtistFramerService((Image)pictureBox1.Image))
                {
                    pictureBox1.Image = edit.Image;
                }
            }
        }
        #endregion
        #region Action
        private void DisplayData()
        {
            Inti();
            if (_DIC.TryGetValue("Process", out DataTable dataTable))
            {
                dataGridView.DataSource = dataTable;
            }
            
            if (_DIC.TryGetValue("Top_SMT", out dataTable))
            {
                dataGridView1.DataSource = dataTable;
            }
            if (_DIC.TryGetValue("Top_Backend", out dataTable))
            {
                dataGridView2.DataSource = dataTable;
            }
            if (_DIC.TryGetValue("OQC", out dataTable))
            {
                dataGridView3.DataSource = dataTable;
            }

        }
        private void LoadDataProcess()
        {
            try
            {
                string itemCode = textBox1.Text.Trim();
                string lotNo = textBox2.Text.Trim();
                _DIC = new ASSY_YIELDService().Load(itemCode, lotNo);
                DisplayData();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



        private void PackData()
        {
            _DIC = new Dictionary<string, DataTable>();
            _DIC.Add("Process", (DataTable)dataGridView.DataSource);
            _DIC.Add("Top_SMT", (DataTable)dataGridView1.DataSource);
            _DIC.Add("Top_Backend", (DataTable)dataGridView2.DataSource);
            _DIC.Add("OQC", (DataTable)dataGridView3.DataSource);
        }
        private void UpdateStatus()
        {
            if (UserSession.Instance.IsLoggedIn)
            {
                tdmK_Label9.Text = $"Hello {UserSession.Instance.Username}";
                textBox3.Text = UserSession.Instance.User_ID;
                tdmK_Button6.Text = "Logout";
            }
            else
            {
                tdmK_Label9.Text = $"Hello guest";
                textBox3.Text = "";
                tdmK_Button6.Text = "Login";
            }
        }
        private void LoginProcess()
        {
            if (!UserSession.Instance.IsLoggedIn)
            {
                Login login = new Login();
                login.ShowDialog();
            }
            else
            {
                UserSession.Instance.Logout();
            }
            UpdateStatus();
        }
        private DataTable CaculateDataTable(DataTable dataTable)
        {
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                int sumStation = 0, sumError = 0;
                if (i == 0)
                {
                    sumStation = Convert.ToInt32(dataTable.Rows[0][1]);
                }
                else
                {
                    sumStation = Convert.ToInt32(dataTable.Rows[i - 1][2]);
                }
                dataTable.Rows[i][1] = sumStation;
                for (int j = dataTable.Columns.Count - 1; j > 2; j--)
                {
                    int z = Convert.ToInt32(dataTable.Rows[i][j]);
                    sumError += Convert.ToInt32(dataTable.Rows[i][j]);
                }
                dataTable.Rows[i][2] = sumStation - sumError;
            }
            double sumALL = double.Parse(dataTable.Rows[0][1].ToString());
            label6.Text = sumALL.ToString();
            double elu = 0, IPQC = 0, ORT = 0, WIP = 0, Others = 0, reject = 0;
            foreach (DataRow row in dataTable.Rows)
            {
                elu += double.Parse(row["Evaluation"].ToString());
                IPQC += double.Parse(row["IPQC"].ToString());
                ORT += double.Parse(row["ORT"].ToString());
                WIP += double.Parse(row["WIP"].ToString());
                Others += double.Parse(row["Others"].ToString());
                reject += double.Parse(row["Rejected"].ToString());
            }
            label9.Text = (sumALL - elu - IPQC - ORT - WIP - Others).ToString();
            label10.Text = (sumALL - elu - IPQC - ORT - WIP - Others - reject).ToString();
            label7.Text = (sumALL - reject - elu - IPQC - ORT - WIP - Others).ToString();
            label11.Text = (Math.Round((sumALL - reject - elu - IPQC - ORT - WIP - Others) / sumALL * 100, 2)).ToString() + "%";
            label12.Text = (Math.Round((sumALL - elu - IPQC - ORT - WIP - Others - reject) / (sumALL - elu - IPQC - ORT - WIP - Others) * 100, 2)).ToString() + "%";
            return dataTable;
        }
        private DataTable Yield_HittersDataTable()
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Station", typeof(string));
            dataTable.Columns.Add("Defect Description", typeof(Image));
            dataTable.Columns.Add("Defect Name", typeof(string));
            dataTable.Columns.Add("Defects Qty", typeof(int));
            dataTable.Columns.Add("Defect Rate", typeof(string));
            dataTable.Columns.Add("Root Cause", typeof(string));
            dataTable.Columns.Add("Corrective Action", typeof(string));
            return dataTable;
        }
        private DataTable YieldDataTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Station", typeof(string));
            dt.Columns.Add("Input", typeof(int));
            dt.Columns.Add("Passed and shipped to next process", typeof(int));
            dt.Columns.Add("Rejected", typeof(int));
            dt.Columns.Add("Evaluation", typeof(int));
            dt.Columns.Add("IPQC", typeof(int));
            dt.Columns.Add("ORT", typeof(int));
            dt.Columns.Add("WIP", typeof(int));
            dt.Columns.Add("Others", typeof(int));
            string[] str = { "SMT", "X-Ray", "ICT/FCT", "FQC", "OQC", "Packaging" };
            foreach (string item in str)
            {
                DataRow row = dt.NewRow();
                row["Station"] = item;
                row["Input"] = 0;
                row["Passed and shipped to next process"] = 0;
                row["Rejected"] = 0;
                row["Evaluation"] = 0;
                row["IPQC"] = 0;
                row["ORT"] = 0;
                row["WIP"] = 0;
                row["Others"] = 0;
                dt.Rows.Add(row);
            }

            return dt;
        }
        private void Inti()
        {
            dataGridView.DataSource = YieldDataTable();
            dataGridView1.DataSource = Yield_HittersDataTable();
            dataGridView2.DataSource = Yield_HittersDataTable();
            dataGridView3.DataSource = Yield_HittersDataTable();
        }
        private void ClearData()
        {
            Inti();
        }









        #endregion
        private void MakeColor()
        {
            ((DataGridViewImageColumn)dataGridView1.Columns["Defect Description"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            ((DataGridViewImageColumn)dataGridView2.Columns["Defect Description"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            ((DataGridViewImageColumn)dataGridView3.Columns["Defect Description"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
        }
        private void dataGridView_CellValueChanged(object sender, EventArgs e)
        {
            dataGridView.DataSource = CaculateDataTable((DataTable)dataGridView.DataSource);
        }
    }
}
