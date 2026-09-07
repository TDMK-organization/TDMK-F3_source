using Funtion_F3_SMT;
using OK2SHIP_SMT.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AntdUI;

namespace OK2SHIP_SMT.UserControls
{
    public partial class UC_IQCPeelingTest : UserControl
    {
        private IQCPeelingTest _service = new IQCPeelingTest();

        public UC_IQCPeelingTest()
        {
            InitializeComponent();
            CheckStatus();
        }

        public void CheckStatus()
        {
            if (UserSession.Instance.IsLoggedIn)
            {
                btn_login.Text = "Logout";
                btn_login.Type = AntdUI.TTypeMini.Error;
                tb_username.Text = UserSession.Instance.Username;
            }
            else
            {
                btn_login.Text = "Login";
                btn_login.Type = AntdUI.TTypeMini.Success;
                tb_username.Text = "";
            }
        }

        private void tb_logfile_Leave(object sender, EventArgs e)
        {
        }

        private void btn_login_Click(object sender, EventArgs e)
        {
            if (UserSession.Instance.IsLoggedIn)
            {
                UserSession.Instance.Logout();
            }
            else
            {
                Login fr1 = new Login();
                fr1.ShowDialog();
            }

            CheckStatus();
        }

        private void tb_locationlogfile_TextChanged(object sender, EventArgs e)
        {
            try
            {
                string location = tb_locationlogfile.Text;
                if (!string.IsNullOrEmpty(location))
                {
                    string msg = _service.GetInfor(location, out string itemName, out string partName, out string lotNo,
                        out string Invoice, out string tape, out string makername);
                    tb_itemname.Text = itemName;
                    tb_lotNo.Text = lotNo;
                    tb_invoice.Text = Invoice;
                    tb_partName.Text = partName;

                    throw new Exception(msg);
                }
            }
            catch (Exception ex)
            {
            }
        }


        private void btn_getData_Click(object sender, EventArgs e)
        {
            string itemName = tb_itemname.Text;
            string lotNo = tb_lotNo.Text;
            string location = tb_locationlogfile.Text;
            string invoice = tb_invoice.Text;
            string partName = tb_partName.Text;
            try
            {
                _service = new IQCPeelingTest();
                _service.GetData(location, itemName, partName, lotNo, invoice);
                displayData();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Get infor");
            }
        }

        private void displayData()
        {
            dgv.DataSource = new DataTable();
            MenuOption.Items.Clear();

            foreach (string item in _service._DATA.Keys.ToList())
            {
                MenuOption.Items.Add(new AntdUI.MenuItem(item, "Homepage"));
            }

            tb_itemcodeE.Text = _service._ITEMCODE_E;
        }

        private void FillData(string key)
        {
            DataTable dt = _service._DATA[key];
            dgv.DataSource = dt;
            Makeup(dgv);
        }

        private void Makeup(DataGridView dgv)
        {
            foreach (DataGridViewColumn col in dgv.Columns)
            {
                if (!col.Name.Contains("JudgementMode") && !col.Name.Contains("Measure(mm)") &&
                    !col.Name.Contains("Adhesive"))
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

        private void MenuOption_SelectChanged(object sender, MenuSelectEventArgs e)
        {
            // e.Value hoặc e.Item.Id chứa ID định danh của mục mà bạn vừa đặt ở trên
            string selectedId = e.Value.ToString();
            FillData(selectedId);
        }


        private void tb_save_Click(object sender, EventArgs e)
        {
            string itemName = tb_itemname.Text;
            string partname = tb_partName.Text;
            string lotNo = tb_lotNo.Text;
            string invoice = tb_invoice.Text;
            try
            {
                _service.Save(itemName, partname, lotNo, invoice);
            }
            catch (DataException ex)
            {
                try
                {
                    if (MessageBox.Show(ex.Message, "Save Error", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        _service.Save(itemName, partname, lotNo, invoice, true);
                    }
                }
                catch (Exception exZ)
                {
                    MessageBox.Show(exZ.Message, "Message");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message");
            }
            finally
            {
                displayData();
            }
        }


        private void btn_load_Click(object sender, EventArgs e)
        {
            string itemName = tb_itemname.Text;
            string lotNo = tb_lotNo.Text;
            string partname = tb_partName.Text;
            string invoice = tb_invoice.Text;
            try
            {
                _service.Load(itemName, partname, lotNo, invoice);
                displayData();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Load Error");
            }
        }

        private void btn_export_Click(object sender, EventArgs e)
        {
            string itemName = tb_itemname.Text;
            string lotNo = tb_lotNo.Text;
            string partname = tb_partName.Text;
            string invoice = tb_invoice.Text;
            try
            {
                _service.Export(itemName, partname, lotNo, invoice);
                displayData();
                MessageBox.Show("Export Complete");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Load Error");
            }
        }


        private void btn_saveI_Click(object sender, EventArgs e)
        {
            string itemName = tb_itemname.Text;
            string lotNo = tb_lotNo.Text;
            string partname = tb_partName.Text;
            string invoice = tb_invoice.Text;
            string sitemcode = tb_itemcodeE.Text;
            try
            {
                _service.SaveItem(itemName, partname, lotNo, invoice, sitemcode);
                displayData();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Save Error");
            }
        }
    }
}