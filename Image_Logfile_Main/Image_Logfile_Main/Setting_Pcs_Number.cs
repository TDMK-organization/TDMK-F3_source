using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AForge.Video;
using AForge.Video.DirectShow;
using TDMK_SQL;
using System.Data.SqlClient;
using System.Diagnostics.Tracing;
using System.Runtime.InteropServices;
using System.IO;
using TDMK_SEEV_DLL;
using myExcel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using ACF_Process;
using VHX;
using System.Drawing.Drawing2D;

namespace OK2SHIP
{
    public partial class Setting_Pcs_Number : Form
    {

        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SqlConnection sqlcon;
        public Export_Image_Class diem = new Export_Image_Class();
        SEI_Lib myCode = new SEI_Lib();
        myVar exp_proc = new myVar();
        private FilterInfoCollection cameras;
        public VideoCaptureDevice cam;
        DataTable dgv_dt = new DataTable();
        public string folder_name1;
        public string folder_name;
        public static Label lb1;
        public string data_loc;
       
        
        public Setting_Pcs_Number()
        {
            InitializeComponent();
        }

       

        private void Setting_Parameter_Load(object sender, EventArgs e)
        {
            this.Padding = new Padding(10, 0, 10, 10);
            sqlcon = exp_proc.initial_data(myVar.sel_DB, true);
            data_loc = myVar.data_loc;
            try
            {
                sqlcon.Open();
                label8.Text = /*" Server Name: " + server_name + " -*/ "DataBase: " + myVar.sel_DB;
                label11.Text = " STATE :CONNECTED";
                label11.ForeColor = Color.Green;
                sqlcon.Close();
            }
            catch
            {
                label8.Text = /* Server Name: " + server_name + " -- */ "DataBase: " + myVar.sel_DB;
                label11.Text = " STATE :NO CONNECTION";
                label11.ForeColor = Color.Red;
            }
            //Lam_dep(label1);
            //Lam_dep(label2);
            //Lam_dep(label3);
            //Lam_dep(label4);
            //Lam_dep(label9);
            //Lam_dep(label8);
        }
        private void btn_load_setting_Click(object sender, EventArgs e)
        {
            if (tb_itemcode.Text != "")
            {
                if (DGV_setting.DataSource != null)
                {
                    Clear_DGV(DGV_setting);
                }
                DGV_setting.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "SETTING_PCS", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { tb_itemcode.Text }));
                DGV_setting.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            else
            {
                LoadTable("SETTING_PCS", DGV_setting, sqlcon);
                DGV_setting.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }    
        }
        private void btn_save_setting_Click(object sender, EventArgs e)
        {
            CopyDataToTable( DGV_setting, "SETTING_PCS", sqlcon);
            DGV_setting = new DataGridView();
            btn_load_setting.PerformClick();
        }
        private void btn_insert_Click(object sender, EventArgs e)
        {           
            if (tb_itemcode.Text != "")
            {
                if (cb_data_for.Text != "")
                {
                    if (cb_region.Text != "")
                    {
                        if (cb_pcs.Text != "")
                        {
                start_label: string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Data_For", "Region_Zone"}, new string[] { tb_itemcode.Text, cb_data_for.Text, cb_region.Text});
                            if(cb_region.Text=="All")
                            {
                                filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Data_For" }, new string[] { tb_itemcode.Text, cb_data_for.Text });
                            }
                            DataTable sel_dt = TDMK_Code.Datatable_Filter(sqlcon, "SETTING_PCS", filter_str);
                            if(sel_dt.Rows.Count==0)
                            {
                                if (cb_region.Text == "All")
                                {
                                    for (int i = 0; i < cb_region.Items.Count - 1; i++)
                                    {
                                        int ID = TDMK_Code.SQL_MAX("SETTING_PCS", "ID", sqlcon) + 1;
                                        TDMK_Code.insert_val_arr("SETTING_PCS", sqlcon, new string[] { "ID", "ItemCode", "Data_For", "Region_Zone", "Pcs_setting" }, new string[] { ID.ToString(), tb_itemcode.Text, cb_data_for.Text, cb_region.Items[i].ToString(), cb_pcs.Text });

                                    }
                                    if (DGV_setting.DataSource != null)
                                    {
                                        Clear_DGV(DGV_setting);
                                    }
                                    DGV_setting.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "SETTING_PCS", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { tb_itemcode.Text }));

                                }
                                else
                                {
                                    int ID = TDMK_Code.SQL_MAX("SETTING_PCS", "ID", sqlcon) + 1;
                                    TDMK_Code.insert_val_arr("SETTING_PCS", sqlcon, new string[] { "ID", "ItemCode", "Data_For", "Region_Zone", "Pcs_setting" }, new string[] { ID.ToString(), tb_itemcode.Text, cb_data_for.Text, cb_region.Text, cb_pcs.Text });
                                    if (DGV_setting.DataSource != null)
                                    {
                                        Clear_DGV(DGV_setting);
                                        
                                    }
                                    DGV_setting.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "SETTING_PCS", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { tb_itemcode.Text }));
                                }
                            }
                            else
                            {
                                if(MessageBox.Show("Pcs Number is existed. Do you want to update?","Warning",MessageBoxButtons.YesNo)==DialogResult.Yes)
                                {
                                    TDMK_Code.Delelte_FilteredItem_arr("SETTING_PCS", sqlcon, filter_str);
                                    goto start_label;
                                }
                            }
                        }
                        else
                            MessageBox.Show("You need to input pcs number", "Warming", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                        MessageBox.Show("You need to input region/zone", "Warming", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                    MessageBox.Show("You need to input data for", "Warming", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
                MessageBox.Show("You need to input itemcode first", "Warming", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        private void btn_delete_all_Click(object sender, EventArgs e)
        {
            DeleteAllData("SETTING_PCS",sqlcon);
            MessageBox.Show("OK");
        }


        public void Clear_DGV(DataGridView DGV)
        {
            DGV.DataSource = null;
            DGV.Columns.Clear();
            DGV.Rows.Clear();
        }
        public void CheckAndOpenConnection(SqlConnection connection)
        {
            if (connection.State == System.Data.ConnectionState.Closed)
            {
                connection.Open();
            }
        }
        public void LoadTable(string tableName, DataGridView dataGridView,SqlConnection sqlcon)
        {
            try
            {
                CheckAndOpenConnection(sqlcon);
                string selectQuery = $"SELECT * FROM {tableName}";
                SqlDataAdapter dataAdapter = new SqlDataAdapter(selectQuery, sqlcon);
                DataTable dataTable = new DataTable();
                dataAdapter.Fill(dataTable);
                if (dataTable.Rows.Count > 0)
                {
                    dataGridView.DataSource = dataTable;
                }
                else
                    MessageBox.Show("No Data!","Message");

            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }
        public void CopyDataToTable(DataGridView dataGridView, string destinationTable,SqlConnection sqlcon)
        {
            try
            {
                CheckAndOpenConnection(sqlcon);
                // Xóa dữ liệu hiện có trong bảng đích
                DeleteAllData(destinationTable,sqlcon);
                // Tạo DataTable để lưu dữ liệu từ DataGridView
                DataTable dataTable = new DataTable();
                foreach (DataGridViewColumn column in dataGridView.Columns)
                {
                    dataTable.Columns.Add(column.HeaderText);
                }

                // Lưu dữ liệu từ DataGridView vào DataTable
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        DataRow dataRow = dataTable.NewRow();
                        for (int i = 0; i < dataGridView.Columns.Count; i++)
                        {
                            dataRow[i] = row.Cells[i].Value;
                        }
                        dataTable.Rows.Add(dataRow);
                    }
                }

                // Sao chép dữ liệu từ DataTable vào bảng đích
                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlcon))
                {
                    bulkCopy.DestinationTableName = destinationTable;
                    bulkCopy.WriteToServer(dataTable);
                }

                MessageBox.Show("Dữ liệu đã được sao chép vào bảng mới!");

            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }
        public void DeleteAllData(string tableName,SqlConnection sqlcon)
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn xóa tất cả dữ liệu trong bảng?", "Xác nhận xóa dữ liệu", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                try
                {
                    CheckAndOpenConnection(sqlcon);
                    string deleteQuery = $"DELETE FROM {tableName}";
                    SqlCommand command = new SqlCommand(deleteQuery, sqlcon);
                    MessageBox.Show("Đã xóa tất cả dữ liệu trong bảng!");                   
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.Message);
                }
            }
        }

        private void cb_data_for_TextChanged(object sender, EventArgs e)
        {
            switch (cb_data_for.Text)
            {
                case "STACKUP":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("PTH_X-section");
                    cb_region.Items.Add("PTH_X-section_A");
                    cb_region.Items.Add("PTH_X-section_B");
                    cb_region.Items.Add("BVH_X-section");
                    cb_region.Items.Add("BVH_X-section_A");
                    cb_region.Items.Add("BVH_X-section_B");
                    cb_region.Items.Add("Zone_A");
                    cb_region.Items.Add("Zone_B");
                    cb_region.Items.Add("Zone_C");
                    cb_region.Items.Add("Zone_D");
                    cb_region.Items.Add("Zone_E");
                    cb_region.Items.Add("Zone_F");
                    cb_region.Items.Add("Zone_G");
                    cb_region.Items.Add("Zone_H");
                    cb_region.Items.Add("Zone_I");
                    cb_region.Items.Add("Zone_K");
                    cb_region.Items.Add("Zone_J");
                    cb_region.Items.Add("All");
                    break;
                case "IMPEDANCE":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Picture Sample 1");
                    cb_region.Items.Add("Picture Sample 2");
                    cb_region.Items.Add("Picture Sample 3");
                    cb_region.Items.Add("All");
                    break;
                case "BVH_PTH":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Inner layer");
                    cb_region.Items.Add("Outer layer");
                    cb_region.Items.Add("All");
                    break;
                case "SOLDERMASK":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("BGA");
                    cb_region.Items.Add("Hot Bar");
                    cb_region.Items.Add("Connector");
                    cb_region.Items.Add("Trace to trace");
                    cb_region.Items.Add("All");
                    break;
                case "ACF_BEFORE_TAPE_TEST":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("TRAI");
                    cb_region.Items.Add("GIUA");
                    cb_region.Items.Add("PHAI");
                    cb_region.Items.Add("All");
                    break;
                case "ACF_AFTER_TAPE_TEST":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("TRAI");
                    cb_region.Items.Add("GIUA");
                    cb_region.Items.Add("PHAI");
                    cb_region.Items.Add("All");
                    break;


                case "ACF_GLASS_COUPON":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("TRAI");
                    cb_region.Items.Add("GIUA");
                    cb_region.Items.Add("PHAI");
                    cb_region.Items.Add("All");
                    break;

                case "ACF_CLEANING":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    //cb_region.Items.Add("Picture");
                    //cb_region.Text = "Picture";
                    cb_region.Items.Add("Photo before cleaning");
                    cb_region.Items.Add("Photo after cleaning");
                    cb_region.Items.Add("Photo after ACF peel test");
                    cb_region.Items.Add("Photo after OQC testing");
                    break;
                case "ACF_GRAPH_FORCE":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Graph");
                    cb_region.Items.Add("All");

                    break;
                case "CQRA_BHAST":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Picture");
                    cb_region.Text = "Picture";


                    break;
                case "CQRA_HIGH_SPEED_BALL_SHEAR":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("SEM");
                    cb_region.Items.Add("Graph");
                    cb_region.Items.Add("All");

                    break;
                case "OTHER":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Other");
                   
                    break;
                case "CQRA_SOLDERABILITY":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Picture");
                 
                    break;
                case "THERMAL_STRESS":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("BVH");
                    cb_region.Items.Add("BVH_A_Side");
                    cb_region.Items.Add("BVH_B_Side");
                    cb_region.Items.Add("BVH_Inner");
                    cb_region.Items.Add("PTH");
                    cb_region.Items.Add("All");
                    break;
                case "HOT_BAR_LOOP_TEST":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Picture");
                    cb_region.Text = "Picture";
                    break;
                case "OQC_TEST":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Solder Mask adhesion");
                    cb_region.Items.Add("PI & SUS Stiffener adhesion");
                    cb_region.Items.Add("Coverlay adhesion");
                    cb_region.Items.Add("Gold adhesion");
                    cb_region.Items.Add("EMI shield adhesion");
                    cb_region.Items.Add("All");
                    break;
                case "CQRA_CHEMICAL_RESISTANCE":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Picture before");
                    cb_region.Items.Add("Picture after");
                    cb_region.Items.Add("All");
                    break;
                case "CQRA_IR_VIA_TO_VIA":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Picture");
                    break;
                case "CQRA_IR_TRACE_TO_TRACE":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Picture");
                    cb_region.Text = "Picture";
                    break;
                case "CQRA_IR_LAYER_TO_LAYER":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Picture");
                    break;
                case "CQRA_DIELECTRIC_WITHSTANDING":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Before");
                    cb_region.Items.Add("After");
                    cb_region.Items.Add("All");
                    break;
                case "CQRA_FLUX_RESIST":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Without_FR4_Bottom side_B2B zone");
                    cb_region.Items.Add("Without_FR4_Bottom side_BGA pads zone");
                    cb_region.Items.Add("Without_FR4_Bottom side_Large Cu plane / GND zone");
                    cb_region.Items.Add("Without_FR4_Bottom side_Wire bonding pads zone");

                    cb_region.Items.Add("Without_FR4_Top side_B2B zone");
                    cb_region.Items.Add("Without_FR4_Top side_BGA pads zone");
                    cb_region.Items.Add("Without_FR4_Top side_Large Cu plane / GND zone");
                    cb_region.Items.Add("Without_FR4_Top side_Wire bonding pads zone");

                    cb_region.Items.Add("With_FR4_Bottom side_B2B zone");
                    cb_region.Items.Add("With_FR4_Bottom side_BGA pads zone");
                    cb_region.Items.Add("With_FR4_Bottom side_Large Cu plane / GND zone");
                    cb_region.Items.Add("With_FR4_Bottom side_Wire bonding pads zone");

                    cb_region.Items.Add("With_FR4_Top side_B2B zone");
                    cb_region.Items.Add("With_FR4_Top side_BGA pads zone");
                    cb_region.Items.Add("With_FR4_Top side_Large Cu plane / GND zone");
                    cb_region.Items.Add("With_FR4_Top side_Wire bonding pads zone");
                    cb_region.Items.Add("All");
                    break;

                default:
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Picture");
                    cb_region.Text = "Picture";
                    break;
            }
        }

        private void updateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (DGV_setting.SelectedRows.Count > 0)
            {
                foreach (DataGridViewRow row in DGV_setting.SelectedRows)
                {
                    string itemcode = row.Cells[1].Value.ToString();
                    string data_for = row.Cells[2].Value.ToString();
                    string region = row.Cells[3].Value.ToString();
                    string pcs_number = row.Cells[4].Value.ToString();
                    TDMK_Code.update_item_val_filter("SETTING_PCS", sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode", "Data_For", "Region_Zone" }, new string[] { itemcode, data_for, region }), "Pcs_setting", pcs_number);


                }

                MessageBox.Show("Update Completed!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                btn_load_setting.PerformClick();
            }
            else
                MessageBox.Show("You need to select row to update!","Message", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (DGV_setting.SelectedRows.Count > 0)
            {
                foreach (DataGridViewRow row in DGV_setting.SelectedRows)
                {
                    string ID = row.Cells[0].Value.ToString();
                    string itemcode = row.Cells[1].Value.ToString();
                    string data_for = row.Cells[2].Value.ToString();
                    string region = row.Cells[3].Value.ToString();
                    string pcs_number = row.Cells[4].Value.ToString();
                    TDMK_Code.Delelte_FilteredItem_arr("SETTING_PCS", sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode", "Data_For", "Region_Zone" ,"Pcs_setting","ID"}, new string[] { itemcode, data_for, region,pcs_number ,ID}));
                }
                MessageBox.Show("Deleted!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                btn_load_setting.PerformClick();
            }
            else
                MessageBox.Show("You need to select row to delete!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void btn_copy_all_Click(object sender, EventArgs e)
        {
            //try
            //{
                if (DGV_setting.DataSource != null)
                {
                    Clear_DGV(DGV_setting);
                }
                DGV_setting.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "SETTING_PCS", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { cb_itemcode_refer.Text }));
                DGV_setting.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                TDMK_Code.Delelte_FilteredItem_arr("SETTING_PCS", sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { tb_itemcode.Text }));

                if (DGV_setting.Rows.Count > 0)
                {
                    for (int x = 0; x < DGV_setting.Rows.Count; x++)
                    {
                        DGV_setting.Rows[x].Cells["ItemCode"].Value = tb_itemcode.Text;
                        int ID = TDMK_Code.SQL_MAX("SETTING_PCS", "ID", sqlcon) + 1;
                        TDMK_Code.insert_val_arr("SETTING_PCS", sqlcon, new string[] { "ID", "ItemCode", "Data_For", "Region_Zone", "Pcs_setting" }, new string[] { ID.ToString(), DGV_setting.Rows[x].Cells["ItemCode"].Value.ToString(), DGV_setting.Rows[x].Cells["Data_For"].Value.ToString(), DGV_setting.Rows[x].Cells["Region_Zone"].Value.ToString(), DGV_setting.Rows[x].Cells["Pcs_setting"].Value.ToString() });

                    }
                    MessageBox.Show("Updated all setting for itemcode " + tb_itemcode.Text + " successfully!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    btn_load_setting.PerformClick();
                }
                else
                    MessageBox.Show("No Data", "Warming", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //}
            //catch
            //{

            //}
           
        }

        private void tb_itemcode_MouseDown(object sender, MouseEventArgs e)
        {
         
        }

        private void tb_itemcode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btn_load_setting.PerformClick();
            }
        }

        // Làm đẹp

        public void Lam_dep(Label label1)
        {
            // Tạo một hình ảnh gradient
            Bitmap gradientImage = new Bitmap(label1.Width, label1.Height);
            using (Graphics g = Graphics.FromImage(gradientImage))
            {
                Color startColor = Color.FromArgb(255, 10, 100, 100); // Màu bắt đầu
                Color endColor = Color.FromArgb(255, 20, 200, 200); // Màu kết thúc
                // Vẽ gradient từ màu bắt đầu đến màu kết thúc
                using (LinearGradientBrush brush = new LinearGradientBrush(label1.ClientRectangle, startColor, endColor, LinearGradientMode.BackwardDiagonal))
                {
                    g.FillRectangle(brush, label1.ClientRectangle);
                }
            }
            // Gán hình ảnh gradient cho BackgroundImage của Label
            label1.BackgroundImage = gradientImage;
            // Đặt BorderStyle là None để loại bỏ đường viền mặc định
            label1.BorderStyle = BorderStyle.None;          
            label1.AutoSize = false;
           
        }

        private void cb_data_for_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void tb_itemcode_TextChanged(object sender, EventArgs e)
        {

        }
    }


}
