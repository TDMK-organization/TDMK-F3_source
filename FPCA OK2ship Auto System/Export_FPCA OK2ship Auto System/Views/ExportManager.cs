using Export_FPCA_OK2ship_Auto_System.Repositories;
using Export_FPCA_OK2ship_Auto_System.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Export_FPCA_OK2ship_Auto_System.Views
{
    public partial class ExportManager : Form
    {
        ExportService _service = new ExportService();
        public ExportManager()
        {

            InitializeComponent();
            setupView();

        }
        private void setupView()
        {
            this.WindowState = FormWindowState.Maximized;
            DGV_Main.DataSource = _service.setupDGV();
        }

        private void UpdateStatus(string category, string type, string status)
        {
            switch (type)
            {
                case "StatusDB":
                    if (_service._dicCategory.TryGetValue(category, out int row))
                    {

                        DGV_Main.Rows[row - 1].Cells[type].Value = status.ToString();
                        switch (status.ToString())
                        {
                            case "OK":
                                DGV_Main.Rows[((int)_service._dicCategory[category]) - 1].Cells[type].Style.BackColor = Color.Green;
                                break;
                            case "No Data":
                                DGV_Main.Rows[((int)_service._dicCategory[category]) - 1].Cells[type].Style.BackColor = Color.Red;
                                break;

                        }
                    }
                    break;
                case "ExportLE":
                    if (_service._dicCategory.TryGetValue(category, out int rowZ))
                    {
                        DGV_Main.Rows[rowZ - 1].Cells["Export status"].Value = $"{status.ToString()}";
                        switch (status.ToString())
                        {
                            case "OK":
                                DGV_Main.Rows[((int)_service._dicCategory[category]) - 1].Cells["Export status"].Style.BackColor = Color.Green;
                                break;
                            default:
                                DGV_Main.Rows[((int)_service._dicCategory[category]) - 1].Cells["Export status"].Style.BackColor = Color.Red;
                                break;

                        }
                    }
                    break;
                case "CheckManual":
                    if (_service._dicCategory.TryGetValue(category, out int rowM))
                    {
                        DGV_Main.Rows[rowM - 1].Cells["Status Manual"].Value = $"{status.ToString()}";
                        switch (status.ToString())
                        {
                            case "OK":
                                DGV_Main.Rows[((int)_service._dicCategory[category]) - 1].Cells["Status Manual"].Style.BackColor = Color.Green;
                                break;
                            default:
                                DGV_Main.Rows[((int)_service._dicCategory[category]) - 1].Cells["Status Manual"].Style.BackColor = Color.Red;
                                break;

                        }
                    }
                    break;
            }
        }
        private void btn_CHECKDB_Click(object sender, System.EventArgs e)
        {
            string itemCode = tb_itemCode.Text.ToString();
            string lotNo = tb_lotNo.Text.ToString();
            if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo))
            {
                MessageBox.Show("ItemCode, LotNo cant null");
                return;
            }
            try
            {
                List<string> s = _service.checkingDatabase(itemCode, lotNo);
                foreach (string item in s)
                {
                    //Debugger.Break();
                    string[] str = item.Split('-');
                    UpdateStatus(str[0], str[1], str[2]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"CHECK DB: {ex.Message}");
            }
        }

        private void tb_lotNo_TextChanged(object sender, EventArgs e)
        {
            tb_lotNo.Text = Lotno_Formated(tb_lotNo.Text);
        }
        public string Lotno_Formated(string lotno)
        {
            string result = "";
            if (lotno.All(char.IsDigit))
            {
                try
                {
                    result = string.Format("{0:00000}", Convert.ToInt32(lotno));
                    result = result.Substring(0, 5);
                }
                catch { }
            }
            else
            {
                if (lotno.Contains('-'))
                {
                    string lotno1 = lotno.Split('-')[0];
                    string cutno = lotno.Split('-')[1];
                    if (lotno1.All(char.IsDigit) && cutno.All(char.IsDigit))
                    {
                        string lotno2 = string.Format("{0:00000}", Convert.ToInt32(lotno1));
                        string cutno2 = string.Format("{0:00}", Convert.ToInt32(cutno));
                        result = lotno2.Substring(0, 5) + "-" + cutno2.Substring(0, 2);
                    }
                }
                if (lotno.Contains('('))
                {
                    string lotno1 = lotno.Split('(')[0];
                    string cutno = lotno.Split('(')[1].Replace(")", "");
                    if (lotno1.All(char.IsDigit) && cutno.All(char.IsDigit))
                    {
                        string lotno2 = string.Format("{0:00000}", Convert.ToInt32(lotno1));
                        string cutno2 = string.Format("{0:00}", Convert.ToInt32(cutno));
                        result = lotno2.Substring(0, 5) + "-" + cutno2.Substring(0, 2);
                    }
                }
            }
            return result;
        }


        private void Export(object sender, EventArgs e)
        {
            try
            {
                string itemCode = tb_itemCode.Text.ToString();
                string lotNo = tb_lotNo.Text.ToString();
                if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo))
                {
                    throw new Exception("NHAP ITEM CODE LOT NO");
                }

                // Lấy category xuất lẻ
                List<string> listCategory1 = new List<string>();
                DataTable dt = (DataTable)DGV_Main.DataSource;

                foreach (DataRow row in dt.Rows)
                {
                    if ((bool)row["Export form DB"] == true)
                    {
                        listCategory1.Add(row["Category"].ToString());
                    }
                }




                List<string> list = _service.ExportOneByOne(itemCode, lotNo, listCategory1.ToArray());

                foreach (string item in list)
                {
                    //Debugger.Break();
                    string[] str = item.Split('-');
                    UpdateStatus(str[0], str[1], str[2]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"EXPORT: {ex.Message}");
            }
        }


        private void btn_check_manual_Click(object sender, EventArgs e)
        {
            try
            {

                string itemCode = tb_itemCode.Text.ToString();
                string lotNo = tb_lotNo.Text.ToString();
                if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo))
                {
                    throw new Exception("NHAP ITEM CODE LOT NO");
                }
                List<string> list = _service.CheckManual(itemCode, lotNo);
                foreach (string item in list)
                {
                    //Debugger.Break();
                    string[] str = item.Split('-');
                    if (str.Count() < 3)
                    {
                        break;
                    }
                    UpdateStatus(str[0], str[1], str[2]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"CHECK: {ex.Message}");
            }
        }

        private void DGV_Main_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            string columnName = DGV_Main.Columns[e.ColumnIndex].Name;
            if (columnName.Contains("Export form DB") || columnName.Contains("Export manual"))
            {
                bool value = (bool)DGV_Main.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                //Debugger.Break();
                if (value)
                {
                    if (columnName.Contains("Export form DB"))
                    {
                        DGV_Main.Rows[e.RowIndex].Cells["Export manual"].Value = false;
                    }
                    else
                    {
                        DGV_Main.Rows[e.RowIndex].Cells["Export form DB"].Value = false;

                    }
                }
            }
        }
        /// <summary>
        /// Uu tien xuat manual
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_UT3_Click(object sender, EventArgs e)
        {
            try
            {
                int i = 0;
                while (true)
                {
                    DGV_Main.Rows[i].Cells["Export manual"].Value = true;
                    i++;
                }
            }
            catch
            {

            }
        }

        private void btn_UT4_Click(object sender, EventArgs e)
        {
            try
            {
                int i = 0;
                while (true)
                {
                    DGV_Main.Rows[i].Cells["Export form DB"].Value = true;
                    i++;
                }
            }
            catch
            {

            }
        }

        private void btn_ExportALL_Click(object sender, EventArgs e)
        {
            try
            {
                string itemCode = tb_itemCode.Text.ToString();
                string lotNo = tb_lotNo.Text.ToString();
                if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo))
                {
                    throw new Exception("NHAP ITEM CODE LOT NO");
                }

                // Lấy category xuất lẻ
                List<string> listCategory = new List<string>();
                DataTable dt = (DataTable)DGV_Main.DataSource;

                foreach (DataRow row in dt.Rows)
                {
                    if ((bool)row["Export form DB"] == true)
                    {
                        listCategory.Add(row["Category"].ToString());
                    }
                }
                // Update trạng thái xuất lẻ
                List<string> list = _service.ExportOneByOne(itemCode, lotNo, listCategory.ToArray());

                foreach (string item in list)
                {
                    //Debugger.Break();
                    string[] str = item.Split('-');
                    UpdateStatus(str[0], str[1], str[2]);
                }

                // Lấy category xuất manual

                list.Clear();

                foreach (DataRow row in dt.Rows)
                {
                    if ((bool)row["Export manual"] == true)
                    {
                        listCategory.Add(row["Category"].ToString());
                    }
                }

                list = _service.ExportAll(itemCode, lotNo, listCategory.ToArray());
                foreach (string item in list)
                {
                    //Debugger.Break();
                    string[] str = item.Split('-');
                    if (str.Count() < 3)
                    {
                        break;
                    }
                    UpdateStatus(str[0], str[1], str[2]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"EXPORT: {ex.Message}");
            }
        }

        private void btn_setting_Click(object sender, EventArgs e)
        {
            SettingManual stn = new SettingManual(_service.getCategories());
            stn.ShowDialog();
            try
            {

                _service.SetupManual(tb_itemCode.Text, tb_lotNo.Text, stn._LOCATION, stn._CATEGORY, stn._TAKEALL);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
