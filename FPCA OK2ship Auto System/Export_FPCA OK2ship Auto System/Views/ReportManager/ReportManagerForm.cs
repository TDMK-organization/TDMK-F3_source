using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using AntdUI;
using Export_FPCA_OK2ship_Auto_System.Services;

namespace Export_FPCA_OK2ship_Auto_System.Views
{
    public partial class ReportManagerForm : AntdUI.Window
    {
        private TDMK_ReportManagerService _service = new TDMK_ReportManagerService("", "", "");
        TAlignFrom align = TAlignFrom.BL;

        public ReportManagerForm()
        {
            InitializeComponent();
            dgv.CellButtonClick += Table_base_CellButtonClick;
            dgv.Enabled = false;
            Reset();
        }

        #region Event

        private void btn_delete_Click(object sender, EventArgs e)
        {
            MessageNotification("Warining", "Chức năng chưa phát triển");
        }

        private void ip_itemname_Leave(object sender, EventArgs e)
        {
            ip_itemname.Text = ip_itemname.Text.Trim();
            if (ip_itemname.Text == _service.__ItemName)
            {
                initService();
            }
        }

        private void ip_build_Leave(object sender, EventArgs e)
        {
            ip_build.Text = ip_build.Text.Trim();
            if (ip_build.Text == _service.__Build)
            {
                initService();
            }
        }


        private void ip_config_Leave(object sender, EventArgs e)
        {
            ip_config.Text = ip_config.Text.Trim();
            if (ip_config.Text == _service.__Config)
            {
                initService();
            }
        }


        private void btn_load_Click(object sender, EventArgs e)
        {
            if (btn_load.Text == "Close")
            {
                CloseAction();
            }
            else if (btn_load.Text == "Load")
            {
                LoadAction();
            }
        }


        private void btn_save_Click(object sender, EventArgs e)
        {
            if (btn_save.Text == "Create New")
            {
                Create();
            }
            else if (btn_save.Text == "Save")
            {
                Save();
            }
        }


        private void Table_base_CellButtonClick(object sender, AntdUI.TableButtonEventArgs e)
        {
            if (e.Column != null && e.Column.Title == "Action" && e.Btn is AntdUI.CellButton btn)
            {
                var rowData = e.Record as Dictionary<string, object>;
                if (rowData == null) return;

                string sheetName = rowData["SheetName"]?.ToString();

                if (btn.Text == "Xóa" || btn.Id.StartsWith("delete_"))
                {
                    // var confirm = MessageBox.Show($"Bạn có chắc chắn muốn xóa sheet '{sheetName}'?", "Xác nhận",
                    //     MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    // if (confirm == DialogResult.Yes)
                    // {
                    //     if (_service._DATA.ContainsKey(sheetName))
                    //     {
                    //         _service._DATA.Remove(sheetName);
                    //     }
                    //
                    //     display_data(); // Load lại bảng
                    // }
                }
                // Xử lý nút Upload
                else if (btn.Text == "Setup" || btn.Id.StartsWith("setup_"))
                {
                    string itemName = ip_itemname.Text;
                    string build = ip_build.Text;
                    string config = ip_config.Text;
                    SetupForm frm = new SetupForm(sheetName, _service);
                    frm.ShowDialog();
                    DisplayData();
                    
                }
            }
        }

        #endregion

        #region method

        private void MessageNotification(string title, string message, string type = "error")
        {
            switch (type)
            {
                case "error":
                    AntdUI.Notification.error(this, title, message, autoClose: 10, align: align);
                    break;
                default:
                    AntdUI.Notification.info(this, title, message, autoClose: 10, align: align);
                    break;
            }
        }

        private void Reset()
        {
            initService();
            DisplayData();
        }

        private void initService()
        {
            string itemName = ip_itemname.Text.Trim();
            string build = ip_build.Text.Trim();
            string config = ip_config.Text.Trim();
            _service = new TDMK_ReportManagerService(itemName, build, config);
        }

        private void DisplayData()
        {
            List<Dictionary<string, object>> dataList = new List<Dictionary<string, object>>();
            int stt = 1;
            if (_service.listcategory.Count <= 0)
            {
                dataList.Add(
                    new Dictionary<string, object>
                    {
                        { "STT", "" },
                        { "SheetName", "" },
                        { "DataForm", "" },
                        { "Status", "" },
                        {
                            "Action", ""
                        }
                    }
                );
            }

            foreach (string key in _service.listcategory)
            {
                dataList.Add(
                    new Dictionary<string, object>
                    {
                        { "STT", stt++ },
                        { "SheetName", key },
                        { "DataForm", _service.__STATUS[key] },
                        { "Status", _service.__STATUS[key] },
                        {
                            "Action", new CellLink[]
                            {
                                // Đặt ID rõ ràng cho từng nút để nhận diện khi click
                                new CellButton("setup_" + key, "Setup", TTypeMini.Primary),
                                // Đặt ID rõ ràng cho từng nút để nhận diện khi click
                                new CellButton("delete_" + key, "Delete", TTypeMini.Error),
                            }
                        }
                    }
                );
            }

            dgv.DataSource = dataList;
        }

        public void Create()
        {
            try
            {
                while (true)
                {
                    try
                    {
                        string itemName = ip_itemname.Text.Trim();
                        string build = ip_build.Text.Trim();
                        string config = ip_config.Text.Trim();
                        _service.CreateNew(itemName, build, config);
                        MessageNotification("Create", "Success", "INFOR");
                        break;
                    }
                    catch (DataException ex)
                    {
                        if (MessageBox.Show(ex.Message, "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) ==
                            DialogResult.Yes)
                        {
                            _service.Save(false);
                            MessageNotification("Create", "Success", "INFOR");
                            break;
                        }
                        else
                        {
                            throw new Exception("Cancel Create Data!");
                        }
                    }
                }

                DisplayData();
                btn_save.Text = "Save";
                btn_load.Text = "Close";
                dgv.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageNotification("Create", ex.Message);
            }
        }

        private void CloseAction()
        {
            try
            {
                ip_itemname.Text = "";
                ip_build.Text = "";
                ip_config.Text = "";
                _service.EmptyData();

                dgv.Enabled = false;
                DisplayData();

                btn_save.Text = "Create New";
                btn_load.Text = "Load";
            }
            catch (Exception ex)
            {
                MessageNotification("Create", ex.Message);
            }
        }

        private void LoadAction()
        {
            try
            {
                string itemName = ip_itemname.Text.Trim();
                string build = ip_build.Text.Trim();
                string config = ip_config.Text.Trim();
                _service.Load(itemName, build, config);
                DisplayData();
                dgv.Enabled = true;
                btn_load.Text = "Close";
                btn_save.Text = "Save";
            }
            catch (Exception ex)
            {
                MessageNotification("Load", ex.Message);
            }
        }

        private void Save()
        {
            try
            {
                try
                {
                    _service.Save();
                    MessageNotification("Save", "Success", "INFOR");
                }
                catch (DataException ex)
                {
                    if (MessageBox.Show(ex.Message, "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) ==
                        DialogResult.Yes)
                    {
                        _service.Save(false);
                        MessageNotification("Save", "Success", "INFOR");
                    }
                    else
                    {
                        throw new Exception("Cancel Create Data!");
                    }
                }


                DisplayData();
                btn_save.Text = "Save";
                btn_load.Text = "Close";
                dgv.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageNotification("Save", ex.Message);
            }
        }

        #endregion
    }
}