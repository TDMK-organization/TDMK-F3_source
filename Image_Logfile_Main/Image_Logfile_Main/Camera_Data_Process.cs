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
using Echeck_LogFile_Process;

using Microsoft.VisualBasic.FileIO;
using System.Drawing.Drawing2D;
using Resize_Lib;
using System.Security.Cryptography;
using OfficeOpenXml;
using System.Drawing.Imaging;

namespace OK2SHIP
{
    public partial class FrmCamera : Form
    {

        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SqlConnection sqlcon;
        public Funtion_export F_export_new = new Funtion_export();
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
        public myExcel.Workbook wb;
        bool state_setting = false;
        public int selected_row;
        public int selected_cell;
        public bool update_dielectric = false;
        public ExportEPPlus F_exportEPPlus = new ExportEPPlus();

        public struct High_Speed_Ball_Shear_Logfile
        {
            public byte[] graph_data { get; set; }
            public string force_data { get; set; }
            public High_Speed_Ball_Shear_Logfile(byte[] in_graph, string in_force)
            {
                graph_data = in_graph;
                force_data = in_force;
            }
        }
        public FrmCamera()
        {
            InitializeComponent();
            cameras = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            foreach (FilterInfo info in cameras)
            {
                cb_usb.Items.Add(info.Name);

            }
            if (cb_usb.Items.Count > 0)
            {
                cb_usb.SelectedIndex = 0;
            }

        }
        public void Lam_dep(Label label1)
        {
            // Tạo một hình ảnh gradient
            Bitmap gradientImage = new Bitmap(label1.Width, label1.Height);
            using (Graphics g = Graphics.FromImage(gradientImage))
            {
                Color startColor = Color.FromArgb(255, 10, 100, 100); // Màu bắt đầu
                Color endColor = Color.FromArgb(255, 20, 200, 200); // Màu kết thúc
                // Vẽ gradient từ màu bắt đầu đến màu kết thúc
                using (LinearGradientBrush brush = new LinearGradientBrush(label1.ClientRectangle, startColor, endColor, LinearGradientMode.Vertical))
                {
                    g.FillRectangle(brush, label1.ClientRectangle);
                }
            }
            // Gán hình ảnh gradient cho BackgroundImage của Label
            label1.BackgroundImage = gradientImage;
            // Đặt BorderStyle là None để loại bỏ đường viền mặc định          
            label1.AutoSize = false;

        }
        public void Lam_dep(Button button)
        {
            // Tạo một hình ảnh gradient
            Bitmap gradientImage = new Bitmap(button.Width, button.Height);
            using (Graphics g = Graphics.FromImage(gradientImage))
            {
                Color startColor = Color.FromArgb(255, 10, 10, 100); // Màu bắt đầu
                Color endColor = Color.FromArgb(255, 20, 20, 200); // Màu kết thúc
                // Vẽ gradient từ màu bắt đầu đến màu kết thúc
                using (LinearGradientBrush brush = new LinearGradientBrush(button.ClientRectangle, startColor, endColor, LinearGradientMode.Vertical))
                {
                    g.FillRectangle(brush, button.ClientRectangle);
                }
            }
            // Gán hình ảnh gradient cho BackgroundImage của Label
            button.BackgroundImage = gradientImage;
            // Đặt BorderStyle là None để loại bỏ đường viền mặc định          
            button.AutoSize = false;

        }
        public static bool IsConnectionOpen(SqlConnection connection)
        {
            return connection.State == System.Data.ConnectionState.Open;
        }
        public void Insert_SQL_Data(string tableName, string[] columnNames, object[] values, SqlConnection sqlcon)
        {
            if (IsConnectionOpen(sqlcon) == false)
            {
                sqlcon.Open();
            }
            // Tạo câu lệnh SQL INSERT
            string sql = "INSERT INTO " + tableName + " (";
            for (int i = 0; i < columnNames.Length; i++)
            {
                sql += columnNames[i];
                if (i < columnNames.Length - 1)
                {
                    sql += ", ";
                }
            }
            sql += ") VALUES (";
            for (int i = 0; i < columnNames.Length; i++)
            {
                sql += "@" + columnNames[i];
                if (i < columnNames.Length - 1)
                {
                    sql += ", ";
                }
            }
            sql += ")";

            // Tạo đối tượng SqlCommand và đặt tham số giá trị vào câu lệnh SQL
            using (SqlCommand command = new SqlCommand(sql, sqlcon))
            {
                for (int i = 0; i < columnNames.Length; i++)
                {
                    if (values[i] is byte[])
                    {
                        command.Parameters.AddWithValue("@" + columnNames[i], values[i]);
                    }
                    else
                    {
                        command.Parameters.AddWithValue("@" + columnNames[i], values[i].ToString());
                    }
                }
                // Thực thi câu lệnh SQL để chèn giá trị vào bảng
                command.ExecuteNonQuery();
            }

        }
        public void Form1_Load(object sender, EventArgs e)
        {
            this.Padding = new Padding(8, 0, 8, 8);
            tb_pcs.Text = "1";
            tb_login.Hide();
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

            try
            {
                cb_usb.SelectedIndex = 0;
            }
            catch
            {

            }
            
        }
        public SqlConnection initial_data(string DB_name, bool sa_en)
        {
            SqlConnection _sqlcon_OK2SHIP;
            string app_path = Application.StartupPath;
            string config_file = Path.Combine(app_path, "Config", "config.txt");
            string[] my_config = myCode.read_config_arr(config_file);
            string server_name = "";
            string server_acc = "";
            string server_pass = "";
            foreach (string c in my_config)
            {
                if (c.Contains("Server"))
                {
                    server_name = c.Split(':')[1].Trim();
                }
                if (c.Contains("Account"))
                {
                    server_acc = c.Split(':')[1].Trim();
                }
                if (c.Contains("Password"))
                {
                    server_pass = c.Split(':')[1].Trim();
                }
                if (c.Contains("Data_Location"))
                {
                    data_loc = c.Split('#')[1].Trim();
                }
            }
            string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, DB_name, server_acc, server_pass).ConnectionString;
            if (!sa_en)
            {
                connstr_OK2SHIP = TDMK_Code.data_connection2(server_name, DB_name).ConnectionString;
            }
            _sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);
            try
            {
                _sqlcon_OK2SHIP.Open();
                label8.Text = " Server Name: " + server_name + " - DataBase: " + DB_name;
                label11.Text = " STATE :CONNECTED";
                label11.ForeColor = Color.Green;
                _sqlcon_OK2SHIP.Close();
            }
            catch
            {
                label8.Text = " Server Name: " + server_name + " -- DataBase: " + DB_name;
                label11.Text = " STATE :NO CONNECTION";
                label11.ForeColor = Color.Red;
            }
            return _sqlcon_OK2SHIP;
        }
        private void btn_start_Click(object sender, EventArgs e)
        {
            if (btn_start.Text == "START")
            {
                if (cam != null && cam.IsRunning)
                {
                    cam.Stop();
                }
                cam = new VideoCaptureDevice(cameras[cb_usb.SelectedIndex].MonikerString);
                cam.NewFrame += Cam_NewFrame;
                cam.Start();
                btn_start.Text = "STOP";
                btn_start.BackColor = Color.Red;
            }
            else
            {
                if (cam != null && cam.IsRunning)
                {
                    cam.Stop();
                    btn_start.Text = "START";
                    btn_start.BackColor = Color.White;
                    pictureBox1.Image = null;
                }
            }

        }

        public void Image_Table_Excel(DataTable dt, string Image_Col_name, myExcel.Worksheet tar_wrksht, string tar_rgn)
        {
            int r_inx = 0;
            int x = dt.Rows.Count;
            DataRow dr;
            for (int i = 0; i < x; i++)
            {
                dr = dt.Rows[i];
                //byte[] img_byte = Encoding.ASCII.GetBytes(dr[Image_Col_name].ToString());

                byte[] img_byte = (byte[])(dr[Image_Col_name]);

                Image temp = byteArrayToImage(img_byte);
                Clipboard.SetDataObject(temp, false);
                myExcel.Range sel_rgn = tar_wrksht.Range[tar_rgn].Offset[r_inx, 0];
                int shape_count = tar_wrksht.Shapes.Count;
                tar_wrksht.Paste(sel_rgn, temp);
                myExcel.Shape sel_pic = tar_wrksht.Shapes.Item(shape_count + 1);
                //sel_pic.Name = dr["Image_Name"].ToString();
                sel_pic.LockAspectRatio = MsoTriState.msoFalse;
                float pic_ratio = sel_pic.Width / sel_pic.Height;
                float rgn_ratio = (float)sel_rgn.Width / (float)sel_rgn.RowHeight;
                var rate = pic_ratio / rgn_ratio;
                float margin = (float)(sel_rgn.RowHeight * 0.05);
                if (rate > 1)
                {
                    sel_pic.Width = (float)sel_rgn.Width - 2 * margin;
                    sel_pic.Height = (float)(sel_rgn.Width / rgn_ratio) - 2 * margin;
                }
                else
                {
                    sel_pic.Height = (float)(sel_rgn.Height) - 2 * margin;
                    sel_pic.Width = (float)(sel_rgn.Height * rgn_ratio) - 2 * margin;
                }
                sel_pic.Top = (float)sel_rgn.Top + margin;
                sel_pic.Left = (float)sel_rgn.Left + margin;
                sel_pic.LockAspectRatio = MsoTriState.msoTrue;
                r_inx++;
            }
        }
        private void Cam_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            Bitmap bitmap = (Bitmap)eventArgs.Frame.Clone();
            pictureBox1.Image = bitmap;
        }
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            if (cam != null && cam.IsRunning)
            {
                cam.Stop();
            }

        }
        // các hàm để lý load hình ảnh xuống picturebox
        public byte[] imgToByteConverter(Image inImg)
        {
            ImageConverter imgCon = new ImageConverter();
            return (byte[])imgCon.ConvertTo(inImg, typeof(byte[]));
        }

        public Image byteArrayToImage(byte[] byteArrayIn)
        {
            using (MemoryStream mStream = new MemoryStream(byteArrayIn))
            {
                return Image.FromStream(mStream);
            }
        }
        // Hàm xóa gridview

        public void Clear_DGV(DataGridView DGV)
        {
            DGV.DataSource = null;
            DGV.Columns.Clear();
            DGV.Rows.Clear();
        }
        //Hàm để lưu dữ liệu vào database
        public void Save_To_DataBase(DataGridView dgv, string table_name, string ItemCode, string Lotno, string region)
        {
            if (cb_data_for.Text == "CQRA_DIELECTRIC_WITHSTANDING" && update_dielectric == true)
            {
                // Xóa dữ liệu cũ để cập nhật cái mới
                bool emty_resistane = false;
                foreach (DataGridViewRow row in DGV_Image.Rows)
                {
                    if (row.Cells["Resistance_Data"].Value.ToString() == "")
                    {

                        emty_resistane = true;
                    }
                }
                if (emty_resistane == false) // Cho phép update
                {
                    TDMK_Code.Delelte_FilteredItem_arr(cb_data_for.Text + "_IMAGE", sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { tb_itemcode.Text, tb_lotno.Text }));

                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, Lotno, region });
                start_lable: DataTable tar_dt = TDMK_Code.Datatable_Filter(sqlcon, table_name, filter_str);
                    if (tar_dt.Rows.Count == 0)
                    {
                        List<string> item_lst = new List<string>();

                        foreach (DataGridViewColumn dgv_c in dgv.Columns)
                        {
                            item_lst.Add(dgv_c.Name);
                        }
                        int ID = TDMK_Code.SQL_MAX(table_name, "ID", sqlcon);
                        foreach (DataGridViewRow dgv_r in dgv.Rows)
                        {
                            List<object> item_val_lst = new List<object>();
                            ID += 1;
                            foreach (DataGridViewColumn dgv_c in dgv.Columns)
                            {
                                if (dgv_c.Name == "ID")
                                {
                                    item_val_lst.Add(ID.ToString());
                                }
                                else
                                {
                                    item_val_lst.Add(dgv_r.Cells[dgv_c.Name].Value);
                                }
                            }
                            insert_image_arr(table_name, sqlcon, item_lst, item_val_lst);
                        }
                        MessageBox.Show("Insert completed", "Warning");
                    }
                    else
                    {
                        if (MessageBox.Show("Data is existed!\r\nDo you want to update all?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            TDMK_Code.Delelte_FilteredItem_arr(table_name, sqlcon, filter_str);
                            goto start_lable;
                        }
                    }

                    Clear_DGV(DGV_Image);
                }
                else
                {
                    MessageBox.Show("Chưa đủ dữ liệu điện trở", "Lỗi", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                }
            }
            else
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, Lotno, region });
            start_lable: DataTable tar_dt = TDMK_Code.Datatable_Filter(sqlcon, table_name, filter_str);
                if (tar_dt.Rows.Count == 0)
                {
                    List<string> item_lst = new List<string>();

                    foreach (DataGridViewColumn dgv_c in dgv.Columns)
                    {
                        item_lst.Add(dgv_c.Name);
                    }
                    int ID = TDMK_Code.SQL_MAX(table_name, "ID", sqlcon);
                    foreach (DataGridViewRow dgv_r in dgv.Rows)
                    {
                        List<object> item_val_lst = new List<object>();
                        ID += 1;
                        foreach (DataGridViewColumn dgv_c in dgv.Columns)
                        {
                            if (dgv_c.Name == "ID")
                            {
                                item_val_lst.Add(ID.ToString());
                            }
                            else
                            {
                                item_val_lst.Add(dgv_r.Cells[dgv_c.Name].Value);
                            }
                        }
                        insert_image_arr(table_name, sqlcon, item_lst, item_val_lst);
                    }
                    MessageBox.Show("Insert completed", "Warning");
                }
                else
                {
                    if (MessageBox.Show("Data is existed!\r\nDo you want to update all?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        TDMK_Code.Delelte_FilteredItem_arr(table_name, sqlcon, filter_str);
                        goto start_lable;
                    }
                }
                Clear_DGV(DGV_Image);
            }

        }
        public void Save_To_DataBase_qty(DataGridView dgv, string table_name, string ItemCode, string Lotno, string region, int qty)
        {
            if (cb_data_for.Text == "CQRA_DIELECTRIC_WITHSTANDING" && update_dielectric == true)
            {
                // Xóa dữ liệu cũ để cập nhật cái mới
                bool emty_resistane = false;
                foreach (DataGridViewRow row in DGV_Image.Rows)
                {
                    if (row.Cells["Resistance_Data"].Value.ToString() == "")
                    {
                        emty_resistane = true;
                        break;
                    }
                }
                if (emty_resistane == false) // Cho phép update
                {
                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, Lotno });
                start_lable: DataTable tar_dt = TDMK_Code.Datatable_Filter(sqlcon, table_name, filter_str);
                    if (tar_dt.Rows.Count == 0)
                    {

                        List<string> item_lst = new List<string>();

                        foreach (DataGridViewColumn dgv_c in dgv.Columns)
                        {
                            item_lst.Add(dgv_c.Name);
                        }
                        int ID = TDMK_Code.SQL_MAX(table_name, "ID", sqlcon);
                        int row_count = 0;
                        foreach (DataGridViewRow dgv_r in dgv.Rows)
                        {
                            if (row_count < qty)
                            {
                                List<object> item_val_lst = new List<object>();
                                ID += 1;
                                foreach (DataGridViewColumn dgv_c in dgv.Columns)
                                {
                                    if (dgv_c.Name == "ID")
                                    {
                                        item_val_lst.Add(ID.ToString());
                                    }
                                    else
                                    {
                                        item_val_lst.Add(dgv_r.Cells[dgv_c.Name].Value);
                                    }
                                }
                                insert_image_arr(table_name, sqlcon, item_lst, item_val_lst);
                            }
                            row_count++;
                        }

                        MessageBox.Show("Insert completed", "Warning");
                    }
                    else
                    {
                        if (MessageBox.Show("Data is existed!\r\nDo you want to update all?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            TDMK_Code.Delelte_FilteredItem_arr(table_name, sqlcon, filter_str);
                            goto start_lable;
                        }
                    }

                    Clear_DGV(DGV_Image);
                }
                else
                {
                    MessageBox.Show("Chưa đủ dữ liệu điện trở", "Lỗi", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                }
            }
            else
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, Lotno, region });

                if (cb_data_for.Text == "CQRA_HIGH_SPEED_BALL_SHEAR" /*|| cb_data_for.Text == "CQRA_CHEMICAL_RESISTANCE"*/)
                {
                    filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, Lotno });
                }
            start_lable: DataTable tar_dt = TDMK_Code.Datatable_Filter(sqlcon, table_name, filter_str);
                if (tar_dt.Rows.Count == 0)
                {
                    if (cb_data_for.Text == "ACF_GRAPH_FORCE")
                    {
                        DataTable tbl_data = (DataTable)dgv.DataSource;
                        int i = TDMK_Code.SQL_MAX("ACF_GRAPH_FORCE_IMAGE", "ID", sqlcon) + 1;
                        foreach (DataRow dr in tbl_data.Rows)
                        {
                            dr[0] = i;
                            i++;
                        }
                        BatchBulkCopy(sqlcon, tbl_data, "ACF_GRAPH_FORCE_IMAGE");
                    }
                    else
                    {
                        List<string> item_lst = new List<string>();
                        foreach (DataGridViewColumn dgv_c in dgv.Columns)
                        {
                            item_lst.Add(dgv_c.Name);
                        }
                        int ID = TDMK_Code.SQL_MAX(table_name, "ID", sqlcon);
                        int r_count = 0;
                        //if (cb_data_for.Text == "CQRA_CHEMICAL_RESISTANCE")
                        //    qty = 2 * qty;
                        foreach (DataGridViewRow dgv_r in dgv.Rows)
                        {
                            if (r_count < qty)
                            {
                                List<object> item_val_lst = new List<object>();
                                ID += 1;
                                foreach (DataGridViewColumn dgv_c in dgv.Columns)
                                {
                                    if (dgv_c.Name == "ID")
                                    {
                                        item_val_lst.Add(ID.ToString());
                                    }
                                    else
                                    {
                                        item_val_lst.Add(dgv_r.Cells[dgv_c.Name].Value);
                                    }
                                }
                                insert_image_arr(table_name, sqlcon, item_lst, item_val_lst);
                            }
                            r_count++;
                        }
                    }
                    MessageBox.Show("Insert completed", "Warning");
                }
                else
                {
                    if (MessageBox.Show("Data is existed!\r\nDo you want to update all?", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        TDMK_Code.Delelte_FilteredItem_arr(table_name, sqlcon, filter_str);
                        goto start_lable;
                    }
                }
                Clear_DGV(DGV_Image);
            }

        }
        public void insert_image_arr(string tar_table, SqlConnection in_sqlcon, List<string> item, List<object> item_val)
        {
            string col_name = "";
            string col_val = "";
            int qty = Math.Min(item.Count, item_val.Count);
            for (int i = 0; i < qty; i++)
            {
                if (!(item_val[i] is DBNull))
                {
                    col_name = col_name + item[i] + ",";
                    col_val = col_val + "@" + item[i] + ",";
                }
            }
            col_name = col_name.TrimEnd(',');
            col_val = col_val.TrimEnd(',');
            string query = "Insert into " + tar_table + "(" + col_name + " ) values (" + col_val + ")";
            using (SqlCommand cmd = new SqlCommand(query))
            {
                cmd.Connection = sqlcon;
                for (int i = 0; i < qty; i++)
                {
                    if (!(item_val[i] is DBNull))
                    {
                        cmd.Parameters.AddWithValue("@" + item[i], item_val[i]);
                    }
                }
                sqlcon.Open();
                cmd.ExecuteNonQuery();
                sqlcon.Close();
            }
        }

        public void Save_To_DataBase_Deleted(DataGridView dgv)
        {
            // Byte[] data = new Byte[0];
            int ID = TDMK_Code.SQL_MAX("IMAGE_DELETED", "ID", sqlcon) + 1;
            string data_for = cb_data_for.Text + "_IMAGE";
            string itemcode = dgv.CurrentRow.Cells["ItemCode"].Value.ToString();
            string lotno = dgv.CurrentRow.Cells["LotNo"].Value.ToString();

            byte[] data = (byte[])DGV_Image.CurrentRow.Cells["Image_Data"].Value;

            string pcs_no = dgv.CurrentRow.Cells["Pcs_No"].Value.ToString();
            string Region = dgv.CurrentRow.Cells["Region"].Value.ToString();
            string time = dgv.CurrentRow.Cells["Time_Update"].Value.ToString();
            string op = dgv.CurrentRow.Cells["Operator"].Value.ToString();

            string query = "Insert into " + "IMAGE_DELETED" + "(ID,Data_For,ItemCode,LotNo,Region,Image_Data,Pcs_No,Time_Update,Operator ) values (@ID,@Data_For, @ItemCode, @LotNo,@Region,@Image_Data,@Pcs_No,@Update_Time,@Operator)";

            //AutoCompleteStringCollection list = TDMK_Code.Load_Item_Filter_str(sqlcon, cb_data_for.Text, "ID", TDMK_Code.filter_str(item, item_val));
            using (SqlCommand cmd = new SqlCommand(query))
            {
                try
                {
                    cmd.Connection = sqlcon;
                    cmd.Parameters.AddWithValue("@ID", ID);
                    cmd.Parameters.AddWithValue("@Data_For", data_for);
                    cmd.Parameters.AddWithValue("@ItemCode", itemcode);
                    cmd.Parameters.AddWithValue("@LotNo", lotno);
                    cmd.Parameters.AddWithValue("@Region", Region);
                    cmd.Parameters.AddWithValue("@Image_Data", data);
                    cmd.Parameters.AddWithValue("@Pcs_No", pcs_no);
                    cmd.Parameters.AddWithValue("@Update_Time", time);
                    cmd.Parameters.AddWithValue("@Operator", op);
                    sqlcon.Open();
                    cmd.ExecuteNonQuery();
                    sqlcon.Close();
                }
                catch
                {
                    MessageBox.Show("Nhấp chọn dòng cần xóa trước!");
                }

            }

        }
        private void btn_save_Click(object sender, EventArgs e)
        {
            int pcs_num = -1;
            if (TDMK_Code.IsNumeric(tb_pcs_set.Text))
            {
                pcs_num = Convert.ToInt32(tb_pcs_set.Text);
            }
            if (pcs_num > 0)
            {
                if (DGV_Image.Rows.Count >= pcs_num)
                {
                    int qty = pcs_num;
                    string tar_tbl = cb_data_for.Text + "_IMAGE";
                    Save_To_DataBase_qty(DGV_Image, tar_tbl, tb_itemcode.Text, tb_lotno.Text, cb_region.Text.ToString(), qty);
                    tb_pcs.Text = "1";
                }
                else
                {
                    MessageBox.Show("Bạn cần chụp đúng đủ " + tb_pcs_set.Text + " pcs để lưu dữ liệu", "Chưa đủ dữ liệu!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                } 
            }
            else
            {
                MessageBox.Show("Chưa cài đặt số lượng", "Cảnh báo!");
            }

        }
        private void Add_Columns_To_DataGridView(string[] columnNames, Type[] columnTypes, DataGridView dataGridView)
        {
            // Clear existing columns
            dataGridView.Columns.Clear();

            // Add columns from input arrays
            for (int i = 0; i < columnNames.Length; i++)
            {
                DataGridViewColumn column = (DataGridViewColumn)Activator.CreateInstance(typeof(DataGridViewTextBoxColumn));
                column.Name = columnNames[i];
                column.DataPropertyName = columnNames[i];
                column.ValueType = columnTypes[i];
                column.HeaderText = columnNames[i];
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dataGridView.Columns.Add(column);
            }
        }


        public void Config_DGV(DataGridView DGV, string data_for)
        {
            /*
             
            CQRA_BHAST
            CQRA_SOLDERABILITY
            OQC_TEST
            CQRA_CHEMICAL_RESISTANCE
            CQRA_IR_VIA_TO_VIA
            CQRA_IR_TRACE_TO_TRACE
            CQRA_IR_LAYER_TO_LAYER
            CQRA_FLUX_RESIST
            CQRA_HIGH_SPEED_BALL_SHEAR
            CQRA_DIELECTRIC_WITHSTANDING
            ACF_CLEANING
            ACF_BEFORE_TAPE_TEST
            ACF_AFTER_TAPE_TEST
            ACF_GRAPH_FORCE
            OTHER
             
             */
            Type text = typeof(DataGridTextBoxColumn);
            Type image = typeof(DataGridViewImageColumn);
            if (data_for == "CQRA_BHAST")
            {
                string[] item = new string[] { "ID", "ItemCode", "LotNo", "Region", "Pcs_No", "Image_Data", "Remark", "Operator", "Time_Update" };
                Type[] dataTypes = new Type[] { text, text, text, text, text, image, text, text, text, };
                Add_Columns_To_DataGridView(item, dataTypes, DGV);
            }

        }
        public void Load_Image_To_DGV(DataGridView DGV_Image)
        {
            // Load to Gridview
            if (cb_data_for.Text == "CQRA_BHAST")
            {
                int row_count = DGV_Image.Rows.Count;
                DGV_Image.Rows.Add();
                DGV_Image.Rows[row_count].Cells["ID"].Value = row_count + 1;
                DGV_Image.Rows[row_count].Cells["ItemCode"].Value = tb_itemcode.Text;
                DGV_Image.Rows[row_count].Cells["Lotno"].Value = tb_lotno.Text;
                DGV_Image.Rows[row_count].Cells["Region"].Value = cb_region.Text;
                DGV_Image.Rows[row_count].Cells["Pcs_No"].Value = tb_pcs.Text;
                DGV_Image.Rows[row_count].Cells["Image_Data"].Value = CapturedImage.Image;
                DGV_Image.Rows[row_count].Cells["Remark"].Value = " ";
                DGV_Image.Rows[row_count].Cells["Operator"].Value = tb_operator.Text;
                DGV_Image.Rows[row_count].Cells["Time_Update"].Value = DateTime.Now;
            }

        }
        private void btn_capture_Click(object sender, EventArgs e)
        {
            if (DGV_Image.Rows.Count == 0 && DGV_Image.Columns.Count > 14)
            {
                btn_clear.PerformClick();
            }
            if (tb_operator.Text != "")
            {
                if (pictureBox1.Image != null)
                {
                    if (cb_data_for.Text != "")
                    {
                        if (tb_itemcode.Text != "")
                        {
                            if (tb_lotno.Text != "")
                            {
                                if (cb_region.Text != "")
                                {
                                    if (cb_mode.Checked == false)
                                    {
                                        try
                                        {
                                            CapturedImage.Image = pictureBox1.Image;
                                            DataTable dgv_img = new DataTable();
                                            string data_for = cb_data_for.SelectedItem.ToString();
                                            string tar_tbl = cb_data_for.SelectedItem.ToString() + "_IMAGE";
                                            string region = cb_region.Text;
                                            string img_col = "Image_Data";


                                            if (data_for == "CQRA_DIELECTRIC_WITHSTANDING") // Cập nhật ngày 2.4.2024
                                            {
                                                if (cb_region.Text == "After" && DGV_Image.Rows.Count < 1)
                                                {
                                                    MessageBox.Show("Bạn cần chụp ảnh Before trước, sau đó chụp ảnh After", "Cảnh báo!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                                    cb_region.Text = "Before";
                                                }
                                                else
                                                {
                                                    region = "";
                                                    img_col = "Image_" + cb_region.Text;
                                                    dgv_img = (DataTable)DGV_Image.DataSource;
                                                }
                                            }
                                            else
                                            {
                                                dgv_img = (DataTable)DGV_Image.DataSource;
                                            }

                                            if (dgv_img == null)
                                            {
                                                string filter_str = "";
                                                filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Lotno", "Region" }, new string[] { tb_itemcode.Text, tb_lotno.Text, region });
                                                dgv_img = TDMK_Code.Datatable_Filter(sqlcon, tar_tbl, filter_str);
                                            }
                                            List<string> pcs_lst = dgv_img.AsEnumerable().Select(x => x.Field<string>("Pcs_No")).ToList();
                                            int sel_row = DGV_Image.Rows.Count;
                                            byte[] image_data = imgToByteConverter(CapturedImage.Image);
                                            if (pcs_lst.IndexOf(tb_pcs.Text) != -1)
                                            {
                                                sel_row = pcs_lst.IndexOf(tb_pcs.Text);
                                                dgv_img.Rows[sel_row][img_col] = image_data;
                                            }
                                            else
                                            {
                                                dgv_img.Rows.Add();
                                                dgv_img.Rows[sel_row]["ID"] = sel_row + 1;
                                                dgv_img.Rows[sel_row]["ItemCode"] = tb_itemcode.Text;
                                                dgv_img.Rows[sel_row]["Lotno"] = tb_lotno.Text;
                                                dgv_img.Rows[sel_row]["Pcs_No"] = tb_pcs.Text;
                                                dgv_img.Rows[sel_row][img_col] = image_data;
                                                dgv_img.Rows[sel_row]["Region"] = region;
                                                if (data_for == "CQRA_SOLDERABILITY" || data_for == "CQRA_CHEMICAL_RESISTANCE")
                                                {
                                                    dgv_img.Rows[sel_row]["Remark"] = "PASS";
                                                }
                                                else
                                                {
                                                    dgv_img.Rows[sel_row]["Remark"] = " ";
                                                }
                                                dgv_img.Rows[sel_row]["Operator"] = tb_operator.Text;
                                                dgv_img.Rows[sel_row]["Time_Update"] = DateTime.Now;
                                            }
                                            DGV_Image.DataSource = dgv_img;
                                            if (tb_pcs.Text != "")
                                            {
                                                int pcs = int.Parse(tb_pcs.Text) + 1;
                                                tb_pcs.Text = pcs.ToString();
                                            }
                                        }
                                        catch (Exception err)
                                        {
                                            MessageBox.Show(err.Message, "Warming", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                        }
                                    }
                                    else
                                    {
                                        try
                                        {
                                            if (DGV_Image.Columns.Count < 1)
                                            {
                                                DGV_Image.Columns.Add("ID", "ID");
                                                DGV_Image.Columns.Add("ItemCode", "ItemCode");
                                                DGV_Image.Columns.Add("LotNo", "LotNo");
                                                DGV_Image.Columns.Add("Region", "Region");
                                                DGV_Image.Columns.Add("Pcs_No", "Pcs_No");
                                                DGV_Image.Columns.Add("Image_Data", "Image_Data");
                                                DGV_Image.Columns["Image_Data"].ValueType = typeof(DataGridViewImageColumn);
                                                DGV_Image.Columns.Add("Operator", "Operator");
                                            }
                                            CapturedImage.Image = pictureBox1.Image;
                                            // Load to Gridview
                                            int row_count = DGV_Image.Rows.Count;
                                            DGV_Image.Rows.Add();
                                            DGV_Image.Rows[row_count].Cells["ID"].Value = row_count + 1;
                                            DGV_Image.Rows[row_count].Cells["ItemCode"].Value = tb_itemcode.Text;
                                            DGV_Image.Rows[row_count].Cells["Lotno"].Value = tb_lotno.Text;
                                            DGV_Image.Rows[row_count].Cells["Region"].Value = cb_region.Text;
                                            DGV_Image.Rows[row_count].Cells["Pcs_No"].Value = tb_pcs.Text;
                                            byte[] image_data = imgToByteConverter(CapturedImage.Image);
                                            DGV_Image.Rows[row_count].Cells["Image_Data"].Value = image_data;
                                            DGV_Image.Rows[row_count].Cells["Operator"].Value = tb_operator.Text;

                                            if (tb_pcs.Text != "")
                                            {
                                                int pcs = int.Parse(tb_pcs.Text) + 1;
                                                tb_pcs.Text = pcs.ToString();
                                            }
                                        }
                                        catch (Exception errr)
                                        {
                                            MessageBox.Show(errr.Message, "Warming", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                        }
                                    }
                                }
                                else
                                    MessageBox.Show("Chọn region!");
                            }
                            else
                                MessageBox.Show("Nhập lot number!");
                        }
                        else
                            MessageBox.Show("Nhập ItemCode!");
                    }
                    else
                        MessageBox.Show("Chọn mục Data For!");
                }
                else
                    MessageBox.Show("Nhấn nút START hoặc kết nối CAMERA để bắt đầu chụp ảnh!");
            }
            else
                MessageBox.Show("Nhập tên người thao tác (Operator)");

            
            Auto_Resize_DGV();
        }
        private void DGV_Image_RowHeaderMouseDoubleClick_1(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                byte[] data;
                string data_for = cb_data_for.Text;
                string region = cb_region.Text;
                if ((data_for == "CQRA_DIELECTRIC_WITHSTANDING") || (data_for == "CQRA_HIGH_SPEED_BALL_SHEAR") || (data_for == "ACF_GRAPH_FORCE"))
                {
                    if (region != "")
                    {
                        data = (byte[])DGV_Image.CurrentRow.Cells["Image_" + region].Value;
                    }
                    else
                    {
                        MessageBox.Show("Please, choose region", "Warning");
                        return;
                    }
                }
                else
                {
                    data = (byte[])DGV_Image.CurrentRow.Cells["Image_Data"].Value;
                }
                using (MemoryStream ms = new MemoryStream(data))
                {
                    CapturedImage.Image = Image.FromStream(ms);
                }
            }
            catch
            {
                MessageBox.Show("Please, choose correct Region / Data For", "Warning");
            }

        }
        public void save_image(string region, string col_name_in_DGV, string folder_path)
        {
            try
            {
                if (DGV_Image.Rows.Count > 0)
                {
                    string directory = folder_path;
                    folder_name1 = Path.Combine(directory, DGV_Image.CurrentRow.Cells["ItemCode"].Value.ToString() + "_" + DGV_Image.CurrentRow.Cells["LotNo"].Value.ToString());
                    folder_name = Path.Combine(folder_name1, cb_data_for.Text, region);
                    ImageFormat img_format = ImageFormat.Jpeg;
                    if(tb_out_image.Text==".png")
                    {
                        img_format = ImageFormat.Png;
                    }    
                    for (int i = 0; i < DGV_Image.Rows.Count; i++)
                    {
                        bool exists = System.IO.Directory.Exists(folder_name);
                        if (!exists)
                        {
                            System.IO.Directory.CreateDirectory(folder_name);
                        }
                        string file_dic = folder_name + "\\" + DGV_Image.Rows[i].Cells["Pcs_No"].Value.ToString() + tb_out_image.Text; //Path.Combine(folder_name, cb_data_for.Text, DGV_Image.Rows[i].Cells["ItemCode"].Value.ToString(), DGV_Image.Rows[i].Cells["LotNo"].Value.ToString(), DGV_Image.Rows[i].Cells["Region"].Value.ToString(), DGV_Image.Rows[i].Cells["Pcs_No"].Value.ToString(), ".png");
                        if (cb_mode.Checked == false)
                        {
                            if (DGV_Image.DataSource == null)
                            {
                                Image image = (Bitmap)DGV_Image.Rows[i].Cells[col_name_in_DGV].Value;
                                Bitmap objBitmap = new Bitmap(image, new Size(image.Width / (Convert.ToInt32(cb_size.Text.Replace("÷", ""))), image.Height / (Convert.ToInt32(cb_size.Text.Replace("÷", ""))))); // = 360x270
                                objBitmap.Save(file_dic,img_format);
                            }
                            else
                            {
                                byte[] data = (byte[])DGV_Image.Rows[i].Cells[col_name_in_DGV].Value;
                                using (MemoryStream ms = new MemoryStream(data))
                                {
                                    Image image = Image.FromStream(ms);
                                    Bitmap objBitmap = new Bitmap(image, new Size(image.Width / (Convert.ToInt32(cb_size.Text.Replace("÷", ""))), image.Height / (Convert.ToInt32(cb_size.Text.Replace("÷", "")))));
                                    objBitmap.Save(file_dic, img_format);
                                }    
                            }
                        }
                        else
                        {
                            byte[] data = (byte[])DGV_Image.Rows[i].Cells[col_name_in_DGV].Value;
                            using (MemoryStream ms = new MemoryStream(data))
                            {
                                Image image = Image.FromStream(ms);
                                Bitmap objBitmap = new Bitmap(image, new Size(image.Width / (Convert.ToInt32(cb_size.Text.Replace("÷", ""))), image.Height / (Convert.ToInt32(cb_size.Text.Replace("÷", "")))));
                                objBitmap.Save(file_dic, img_format);
                            }    
                        }
                    }
                    MessageBox.Show("Saved all to: " + folder_name);
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString());
            }

        }

        private void btn_save_folder_Click(object sender, EventArgs e)
        {
            if (DGV_Image.Rows.Count > 0)
            {
                try
                {
                    var fd = new SaveFileDialog();
                    FolderBrowserDialog fbd = new FolderBrowserDialog();
                    if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string folder_path = fbd.SelectedPath;
                        try
                        {
                            if ((cb_data_for.Text == "CQRA_DIELECTRIC_WITHSTANDING") || (cb_data_for.Text == "CQRA_HIGH_SPEED_BALL_SHEAR"))
                            {
                                if (cb_data_for.Text == "CQRA_DIELECTRIC_WITHSTANDING")
                                {
                                    save_image("Image_Before", "Image_Before", folder_path);
                                    save_image("Image_After", "Image_After", folder_path);
                                }
                                if (cb_data_for.Text == "CQRA_HIGH_SPEED_BALL_SHEAR")
                                {
                                    save_image("Photo", "Image_Photo", folder_path);
                                    // save_image("Graph", "Image_Graph", folder_path);
                                }
                            }
                            else
                            {
                                save_image(cb_region.Text, "Image_Data", folder_path);
                            }
                            Clear_DGV(DGV_Image);
                        }
                        catch (Exception err)
                        {
                            MessageBox.Show(err.Message, "Warming", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
                catch (Exception exx)
                {
                    MessageBox.Show(exx.Message, "Warming", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
                MessageBox.Show("No data to save", "Warming", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        }
        public string[] GetFilesWithFilters(string folderPath, string[] filters)
        {
            // Tạo danh sách tạm thời để chứa kết quả
            var fileList = new List<string>();

            // Lặp qua từng bộ lọc
            foreach (string filter in filters)
            {
                // Lấy danh sách các file với bộ lọc hiện tại
                string[] files = Directory.GetFiles(folderPath, filter);

                // Thêm các file vào danh sách tạm thời
                fileList.AddRange(files);
            }

            // Chuyển đổi danh sách tạm thời thành mảng và trả về
            return fileList.ToArray();
        }
        private void btn_load_Click(object sender, EventArgs e)
        {
            if ((tb_itemcode.Text != "") && (tb_lotno.Text != ""))
            {
                string _sel_folder = "";
                if (tb_folder.Text == "")
                {
                    FolderBrowserDialog fold_open = new FolderBrowserDialog();
                    fold_open.SelectedPath = Application.StartupPath;
                    if (fold_open.ShowDialog() == DialogResult.OK)
                    {
                        if (fold_open.SelectedPath != "")
                        {
                            _sel_folder = fold_open.SelectedPath;
                            tb_folder.Text = _sel_folder;
                        }
                    }
                }
                else
                {
                    _sel_folder = tb_folder.Text;
                }
                if (_sel_folder != "")
                {
                    if (_sel_folder.Contains(tb_itemcode.Text) && _sel_folder.Contains(tb_lotno.Text))
                    {
                        if (cb_data_for.SelectedItem.ToString() == "CQRA_HIGH_SPEED_BALL_SHEAR")
                        {
                            string sel_folder = Path.Combine(_sel_folder, cb_region.Text);
                            if (Directory.Exists(sel_folder))
                            {
                                DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, "CQRA_HIGH_SPEED_BALL_SHEAR_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { tb_itemcode.Text, tb_lotno.Text }));
                                if (DGV_Image.DataSource != null)
                                {
                                    src_dt = (DataTable)DGV_Image.DataSource;
                                }
                                if (cb_region.Text == "Graph")
                                {
                                    SortedDictionary<int, High_Speed_Ball_Shear_Logfile> my_HighSpeed = High_speed_Image_LogFile_process(tb_itemcode.Text, tb_lotno.Text, sel_folder, sqlcon);
                                    if (src_dt.Rows.Count == 0)
                                    {
                                        int r_inx = 0;
                                        int id_max = TDMK_Code.SQL_MAX("CQRA_HIGH_SPEED_BALL_SHEAR_IMAGE", "ID", sqlcon);
                                        foreach (var t in my_HighSpeed)
                                        {
                                            src_dt.Rows.Add(t.Key + id_max, tb_itemcode.Text, tb_lotno.Text, null, t.Key, null, t.Value.graph_data, t.Value.force_data, tb_operator.Text, DateTime.Now.ToString(), null);
                                            r_inx++;
                                        }
                                        DGV_Image.DataSource = src_dt;
                                    }
                                    else
                                    {
                                        if (MessageBox.Show("Đã đủ dữ liệu Graph / Force. Bạn muốn cập nhật dữ liệu?", "Cảnh báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                                        {
                                            List<string> pcs_lst = src_dt.AsEnumerable().Select(x => x.Field<string>("Pcs_No")).ToList();
                                            foreach (var pcs in pcs_lst)
                                            {
                                                int pcs_no = Convert.ToInt32(pcs);
                                                if (my_HighSpeed.Keys.ToList().IndexOf(pcs_no) != -1)
                                                {
                                                    int r_inx = pcs_lst.IndexOf(pcs);
                                                    src_dt.Rows[r_inx]["Image_Graph"] = my_HighSpeed[pcs_no].graph_data;
                                                    src_dt.Rows[r_inx]["Force_Data"] = my_HighSpeed[pcs_no].force_data;
                                                }
                                            }
                                            DGV_Image.DataSource = src_dt;
                                        }
                                    }
                                }
                                else
                                {
                                    if (src_dt.Rows.Count > 0)
                                    {
                                        List<byte[]> SEM_lst = src_dt.AsEnumerable().Where(x => x.Field<byte[]>("Image_Photo") != null).Select(x => x.Field<byte[]>("Image_Photo")).ToList();
                                        if (SEM_lst.Count > 0)
                                        {
                                            if (MessageBox.Show("Đã đủ dữ liệu SEM. Bạn muốn cập nhật dữ liệu?", "Cảnh báo", MessageBoxButtons.YesNo) == DialogResult.No)
                                            {
                                                DGV_Image.DataSource = src_dt;
                                                return;
                                            }
                                        }
                                        SortedDictionary<int, byte[]> photo_lst = High_speed_Photo_process(tb_itemcode.Text, tb_lotno.Text, sel_folder, sqlcon);
                                        List<string> pcs_lst = src_dt.AsEnumerable().Select(x => x.Field<string>("Pcs_No")).ToList();
                                        foreach (var pcs in pcs_lst)
                                        {
                                            int pcs_no = Convert.ToInt32(pcs);
                                            if (photo_lst.Keys.ToList().IndexOf(pcs_no) != -1)
                                            {
                                                int r_inx = pcs_lst.IndexOf(pcs);
                                                src_dt.Rows[r_inx]["Image_Photo"] = photo_lst[pcs_no];
                                            }
                                        }
                                        MessageBox.Show("Cập nhật dữ liệu SEM thành công. Ấn nút Update Data để lưu", "Thông báo");
                                        DGV_Image.DataSource = src_dt;
                                    }
                                    else
                                    {
                                        MessageBox.Show("Lấy dữ liệu Graph / Force trước", "Cảnh báo");
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Thư mục: " + sel_folder + " không tồn tại. Hãy chọn đúng thư mục", "Cảnh báo");
                            }

                        }
                        if (cb_data_for.SelectedItem.ToString() == "ACF_GRAPH_FORCE")
                        {
                            DGV_Image.DataSource = diem.Load_ACF_Graph_Force(_sel_folder, tb_itemcode.Text, tb_lotno.Text, sqlcon);
                        }
                    }
                    else
                    {
                        MessageBox.Show("ItemCode / LotNo is not matched. Please select corrected folder", "Warning");
                    }
                }
            }
            else
            {
                MessageBox.Show("Please, enter ItemCode / LotNo", "Warning");
            }
        }
        public List<string> Read_High_Speed_Excel_File(string filePath)
        {
            List<string> resultList = new List<string>();

            using (var reader = new StreamReader(filePath))
            using (var csvParser = new TextFieldParser(reader))
            {
                csvParser.TextFieldType = FieldType.Delimited;
                csvParser.SetDelimiters(","); // Có thể thay đổi dấu phân cách nếu cần

                while (!csvParser.EndOfData)
                {
                    string[] fields = csvParser.ReadFields();

                    if (fields.Length >= 6 && fields[0] == "TEST") // Kiểm tra giá trị cột A
                    {
                        string value = fields[5]; // Lấy giá trị cột F tương ứng
                        resultList.Add(value);
                    }
                }
            }

            return resultList;
        }
        public void Create_DGV_High_Speed_Ball_Shear()
        {
            try
            {
                DGV_Image.Columns.Clear();
                DGV_Image.Rows.Clear();
                DGV_Image.Columns.Add("ID", "ID");
                DGV_Image.Columns.Add("ItemCode", "ItemCode");
                DGV_Image.Columns.Add("LotNo", "LotNo");
                DGV_Image.Columns.Add("Region", "Region");
                DGV_Image.Columns.Add("Pcs_No", "Pcs_No");
                DGV_Image.Columns.Add("Image_Photo", "Image_Photo");
                DGV_Image.Columns.Add("Image_Graph", "Image_Graph");
                DGV_Image.Columns.Add("Force_Data", "Force_Data");
                DGV_Image.Columns.Add("Operator", "Operator");
                DGV_Image.Columns.Add("Time_Update", "Time_Update");
                DGV_Image.Columns.Add("Remark", "Remark");
                DGV_Image.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                DGV_Image.Columns["Image_Photo"].ValueType = typeof(DataGridViewImageColumn);
                DGV_Image.Columns["Image_Graph"].ValueType = typeof(DataGridViewImageColumn);
                btn_capture.Enabled = false;
            }
            catch
            {

            }
        }
        public void data_for_change()
        {
            //if (cb_data_for.Text == "CQRA_HIGH_SPEED_BALL_SHEAR")
            //{
            //    Create_DGV_High_Speed_Ball_Shear();
            //}
            if (DGV_Image.Rows.Count > 1)
            {
                if (btn_load_database.Text != "LOADED")
                {
                    DialogResult result = MessageBox.Show("Bạn có muốn lưu bảng nhớ tạm không ?", "Cảnh báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (result == DialogResult.Yes)
                    {
                        Save_To_DataBase(DGV_Image, cb_data_for.Text + "_IMAGE", tb_itemcode.Text, tb_lotno.Text, cb_region.SelectedItem.ToString());
                        Clear_DGV(DGV_Image);
                    }
                    else
                    {
                        Clear_DGV(DGV_Image);
                    }
                }
            }
            btn_load.Enabled = false;
            btn_load_database.Text = "LOAD FROM DATABASE";
            cb_region.Text = "";
            btn_capture.Enabled = true;
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
                    break;
                case "IMPEDANCE":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Picture Sample 1");
                    cb_region.Items.Add("Picture Sample 2");
                    cb_region.Items.Add("Picture Sample 3");
                    break;
                case "BVH_PTH":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Inner layer");
                    cb_region.Items.Add("Outer layer");
                    break;
                case "SOLDERMASK":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("BGA");
                    cb_region.Items.Add("Hot Bar");
                    cb_region.Items.Add("Connector");
                    cb_region.Items.Add("Trace to trace");
                    break;
                case "ACF_BEFORE_TAPE_TEST":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("TRAI");
                    cb_region.Items.Add("GIUA");
                    cb_region.Items.Add("PHAI");
                    break;
                case "ACF_AFTER_TAPE_TEST":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("TRAI");
                    cb_region.Items.Add("GIUA");
                    cb_region.Items.Add("PHAI");
                    break;

                case "ACF_GLASS_COUPON":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("TRAI");
                    cb_region.Items.Add("GIUA");
                    cb_region.Items.Add("PHAI");
                    break;
                case "ACF_CLEANING":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Photo before cleaning");
                    cb_region.Items.Add("Photo after cleaning");
                    cb_region.Items.Add("Photo after ACF peel test");
                    cb_region.Items.Add("Photo after OQC testing");
                    break;
                case "ACF_GRAPH_FORCE":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Graph");
                    btn_capture.Enabled = false;
                    btn_load.Enabled = true;
                    break;
                case "CQRA_BHAST":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Picture");
                    break;
                case "CQRA_HIGH_SPEED_BALL_SHEAR":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("SEM");
                    cb_region.Items.Add("Graph");
                    btn_capture.Enabled = false;
                    //btn_load.Enabled = false;
                    MessageBox.Show("Ở hạng mục này bạn cần input Graph và lực trước rồi lưu vào DataBase sau đó input photo sau", "Lưu ý");
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
                    break;
                case "HOT_BAR_LOOP_TEST":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Picture");
                    break;
                case "OQC_TEST":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Solder Mask adhesion");
                    cb_region.Items.Add("PI & SUS Stiffener adhesion");
                    cb_region.Items.Add("Coverlay adhesion");
                    cb_region.Items.Add("Gold adhesion");
                    cb_region.Items.Add("EMI shield adhesion");
                    break;
                case "CQRA_CHEMICAL_RESISTANCE":
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Picture before");
                    cb_region.Items.Add("Picture after");
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
                    break;
                case "CQRA_FLUX_RESIST":
                    cb_region.Text = "";

                    DataTable dt_setting = TDMK_Code.Datatable_Filter(sqlcon, "SETTING_PCS", TDMK_Code.filter_str(new string[] { "ItemCode", "Data_For" }, new string[] { tb_itemcode.Text, "CQRA_FLUX_RESIST" }));
                    if (dt_setting.Rows.Count > 0)
                    {
                        List<string> lst_region = dt_setting.AsEnumerable().Select(x => x.Field<string>("Region_Zone")).Distinct().ToList();
                        cb_region.Items.Clear();
                        foreach (string reg in lst_region)
                        {
                            cb_region.Items.Add(reg);
                        }
                    }
                    //cb_region.Items.Add("Without_FR4_Bottom side_B2B zone");
                    //cb_region.Items.Add("Without_FR4_Bottom side_BGA pads zone");
                    //cb_region.Items.Add("Without_FR4_Bottom side_Large Cu plane / GND zone");
                    //cb_region.Items.Add("Without_FR4_Bottom side_Wire bonding pads zone");

                    //cb_region.Items.Add("Without_FR4_Top side_B2B zone");
                    //cb_region.Items.Add("Without_FR4_Top side_BGA pads zone");
                    //cb_region.Items.Add("Without_FR4_Top side_Large Cu plane / GND zone");
                    //cb_region.Items.Add("Without_FR4_Top side_Wire bonding pads zone");

                    //cb_region.Items.Add("With_FR4_Bottom side_B2B zone");
                    //cb_region.Items.Add("With_FR4_Bottom side_BGA pads zone");
                    //cb_region.Items.Add("With_FR4_Bottom side_Large Cu plane / GND zone");
                    //cb_region.Items.Add("With_FR4_Bottom side_Wire bonding pads zone");

                    //cb_region.Items.Add("With_FR4_Top side_B2B zone");
                    //cb_region.Items.Add("With_FR4_Top side_BGA pads zone");
                    //cb_region.Items.Add("With_FR4_Top side_Large Cu plane / GND zone");
                    //cb_region.Items.Add("With_FR4_Top side_Wire bonding pads zone");
                    break;

                default:
                    cb_region.Text = "";
                    cb_region.Items.Clear();
                    cb_region.Items.Add("Picture");
                    break;
            }
            cb_region.SelectedIndex = 0;

        }
        private void cb_data_for_TextChanged(object sender, EventArgs e)
        {
            //data_for_change();
            btn_save.Enabled = true;
            btn_capture.Enabled = true;
        }
        public void BatchBulkCopy(SqlConnection sqlcon_OK2SHIP, DataTable dataTable, string DestinationTbl)
        {
            DataTable dtInsertRows = dataTable;
            using (SqlBulkCopy sbc = new SqlBulkCopy(sqlcon_OK2SHIP))
            {
                if (sqlcon_OK2SHIP.State != ConnectionState.Open)
                {
                    sqlcon_OK2SHIP.Open();
                }
                sbc.DestinationTableName = DestinationTbl;
                foreach (DataColumn dc in dataTable.Columns)
                {
                    sbc.ColumnMappings.Add(dc.ColumnName, dc.ColumnName);
                }
                sbc.WriteToServer(dtInsertRows);
                sqlcon_OK2SHIP.Close();
            }
        }
        private void btn_load_database_Click(object sender, EventArgs e)
        {
            bool data_exist = true;
            //if (DGV_Image.DataSource == null)
            //{
            //    Clear_DGV(DGV_Image);
            //}
            Clear_DGV(DGV_Image);
            try
            {
                string[] item = { "ItemCode", "LotNo", "Region" };
                string[] item_val = { tb_itemcode.Text, tb_lotno.Text, cb_region.Text };
                string tar_tbl = cb_data_for.Text + "_IMAGE";
                if (cb_data_for.Text == "CQRA_HIGH_SPEED_BALL_SHEAR" || cb_data_for.Text == "CQRA_DIELECTRIC_WITHSTANDING" || cb_data_for.Text == "ACF_GRAPH_FORCE" || cb_data_for.Text == "CQRA_CHEMICAL_RESISTANCE")
                {
                    item_val[2] = "";
                }
                DGV_Image.DataSource = TDMK_Code.Datatable_Filter(sqlcon, tar_tbl, TDMK_Code.filter_str(item, item_val));
                btn_load_database.Text = "LOADED";
            }
            catch (Exception)
            {
                MessageBox.Show("Lỗi truy xuất dữ liệu!", "Chú ý!");
                btn_load_database.Text = "LOAD FROM DATABASE";
                return;
            }
            if (DGV_Image.Rows.Count < 1)
            {
                MessageBox.Show("No data!", "Thông báo");
                btn_load_database.Text = "LOAD FROM DATABASE";
                data_exist = false;
                //cb_region.Items.Clear();
                //cb_region.Items.Add("Before");
                //cb_region.Items.Add("After");
                Clear_DGV(DGV_Image);
            }
            else
            {
                btn_load_database.Text = "LOADED";
                Auto_Resize_DGV();
                btn_save.Enabled = false;
                btn_capture.Enabled = false;
            }
            if (cb_data_for.Text == "CQRA_DIELECTRIC_WITHSTANDING" && data_exist == true)
            {
                btn_capture.Enabled = true; btn_save.Enabled = true;
                cb_region.Items.Remove("Before");
                cb_region.SelectedIndex = 0;
                update_dielectric = true;

            }
            //else
            //{
            //    cb_region.Items.Clear();
            //    cb_region.Items.Add("Before");
            //    cb_region.Items.Add("After");
            //}    
        }
        public DataTable GetContentAsDataTable(DataGridView dgv)
        {
            try
            {
                if (dgv.ColumnCount == 0) return null;
                DataTable dtSource = new DataTable();
                foreach (DataGridViewColumn col in dgv.Columns)
                {
                    if (col.Name == string.Empty) continue;
                    if (col.CellType.Name.Contains("Image"))
                    {
                        dtSource.Columns.Add(col.Name, System.Type.GetType("System.Byte[]"));
                    }
                    else
                    {
                        dtSource.Columns.Add(col.Name);
                    }
                    dtSource.Columns[col.Name].Caption = col.HeaderText;
                }
                if (dtSource.Columns.Count == 0) return null;
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    DataRow drNewRow = dtSource.NewRow();
                    foreach (DataColumn col in dtSource.Columns)
                    {
                        if (col.DataType == System.Type.GetType("System.Byte[]"))
                        {
                            drNewRow[col.ColumnName] = imgToByteConverter((Bitmap)row.Cells[col.ColumnName].Value);
                        }
                        else
                        {
                            drNewRow[col.ColumnName] = row.Cells[col.ColumnName].Value;
                        }
                    }
                    dtSource.Rows.Add(drNewRow);
                }
                return dtSource;
            }
            catch
            {
                return null;
            }
        }
        private void resetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btn_start.Enabled = true;
            if (btn_start.Text == "STOP")
            {
                cam.Stop();
                btn_start.Text = "START";
                btn_start.BackColor = Color.White;
            }
            btn_save.Enabled = true;
            btn_capture.Enabled = true;

            if ((DGV_Image.Rows.Count > 0) && (btn_load_database.Text != "LOADED"))
            {
                DialogResult result = MessageBox.Show("Bạn có muốn lưu bảng nhớ tạm không ?", "Cảnh báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    Save_To_DataBase(DGV_Image, cb_data_for.Text + "_IMAGE", tb_itemcode.Text, tb_lotno.Text, cb_region.SelectedItem.ToString());
                    Clear_DGV(DGV_Image);
                }
                else
                {
                    Clear_DGV(DGV_Image);
                }
            }
            btn_load_database.Text = "LOAD FROM DATABASE";
            //tb_itemcode.Text = "";
            tb_lotno.Text = "";
            cb_data_for.Text = "";
            cb_region.Text = "";
            tb_pcs_set.Clear();
            tb_pcs.Text = "1";
            pictureBox1.Image = null;
            CapturedImage.Image = null;
            DGV_Image.DataSource = null;
            btn_load.Enabled = false;
            update_dielectric = false;
        }

        private void DGV_Image_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DGV_Image.ContextMenuStrip = contextMenuStrip2;
        }


        private void Update_menu_Click(object sender, EventArgs e)
        {
            try
            {
                bool en_update = false;
                if (DGV_Image.Rows.Count > 0)
                {
                    if (login.Text == "Login Successful!")
                    {
                        en_update = true;
                    }
                    else
                    {
                        if (cb_data_for.Text == "CQRA_HIGH_SPEED_BALL_SHEAR")
                        {
                            en_update = true;
                        }
                        else
                        {
                            MessageBox.Show("Cần đăng nhập để cập nhật dữ liệu!");
                            tb_login.Show();
                            btn_ok1.Show();
                            tb_login.Focus();
                            return;
                        }
                    }
                    if(en_update)
                    {
                        DataTable temp_dt = (DataTable)DGV_Image.DataSource;
                        string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { tb_itemcode.Text, tb_lotno.Text, cb_region.Text });
                        if (cb_data_for.Text == "CQRA_HIGH_SPEED_BALL_SHEAR")
                        {
                            filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { tb_itemcode.Text, tb_lotno.Text });
                        }
                        TDMK_Code.Delelte_FilteredItem_arr(cb_data_for.Text + "_IMAGE", sqlcon, filter_str);
                        BatchBulkCopy(sqlcon, temp_dt, cb_data_for.Text + "_IMAGE");
                        MessageBox.Show("Data has been updated!");
                    }
                }
                else
                {
                    MessageBox.Show("Chưa có dữ liệu!");
                }
                    
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Error");
            }

        }

        private void insert_menu_Click(object sender, EventArgs e)
        {

        }
        private void eXITToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Do you want to exit!", "Warming!", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                this.Dispose();
            }
        }
        private void DGV_Image_RowEnter(object sender, DataGridViewCellEventArgs e)
        {

            DGV_Image.ContextMenu = menuStrip2.ContextMenu;
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        { 
            //try
            //{
                if (btn_load_database.Text != "LOADED")
                {
                    int r_inx = DGV_Image.CurrentCell.RowIndex;
                    DGV_Image.Rows.RemoveAt(r_inx);
                }
                else
                {
                    if (login.Text == "Login Successful!")
                    {
                        if (MessageBox.Show("Do you want to delete", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            // XÓA
                            string[] item = { "ItemCode", "LotNo", "ID" };
                            string[] item_val = { DGV_Image.CurrentRow.Cells["ItemCode"].Value.ToString(), DGV_Image.CurrentRow.Cells["LotNo"].Value.ToString(), DGV_Image.CurrentRow.Cells["ID"].Value.ToString() };
                            // Save_To_DataBase_Deleted(DGV_Image);
                            TDMK_Code.Delelte_FilteredItem_arr(cb_data_for.Text + "_IMAGE", sqlcon, TDMK_Code.filter_str(item, item_val));
                            DGV_Image.DataSource = TDMK_Code.Datatable_Filter(sqlcon, cb_data_for.Text + "_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { item_val[0], item_val[1] }));

                        }
                    }
                    else
                    {
                        MessageBox.Show("Cần đăng nhập để cập nhật dữ liệu!");
                        tb_login.Show();
                        btn_ok1.Show();
                        tb_login.Focus();
                    }

                }
            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("Chọn dòng cần xóa trước!");
            //}





            //if (DGV_Image.DataSource != null)
            //    {
            //        if (login.Text == "Login Successful!")
            //        {
            //            if (MessageBox.Show("Do you want to delete", "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
            //            {
            //                // XÓA
            //                string[] item = { "ItemCode", "LotNo", "ID" };
            //                string[] item_val = { DGV_Image.CurrentRow.Cells["ItemCode"].Value.ToString(), DGV_Image.CurrentRow.Cells["LotNo"].Value.ToString(), DGV_Image.CurrentRow.Cells["ID"].Value.ToString() };
            //                Save_To_DataBase_Deleted(DGV_Image);
            //                TDMK_Code.Delelte_FilteredItem_arr(cb_data_for.Text + "_IMAGE", sqlcon, TDMK_Code.filter_str(item, item_val));
            //                DGV_Image.DataSource = TDMK_Code.Datatable_Filter(sqlcon, cb_data_for.Text + "_IMAGE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { item_val[0], item_val[1] }));
            //                //Int32 rowToDelete = DGV_Image.Rows.GetFirstRow(DataGridViewElementStates.Selected);
            //                //DGV_Image.Rows.RemoveAt(rowToDelete);
            //                //DGV_Image.ClearSelection();
            //            }
            //        }
            //        else
            //            MessageBox.Show("Cần đăng nhập để xóa dữ liệu!");                
            //    }
            //    else
            //    {
            //        Int32 rowToDelete = DGV_Image.Rows.GetFirstRow(DataGridViewElementStates.Selected);
            //        DGV_Image.Rows.RemoveAt(rowToDelete);
            //        DGV_Image.ClearSelection();
            //    }
        }
        private void iNSERTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (pictureBox1.Image != null)
            {
                if (DGV_Image.DataSource != null)
                {
                    if (DGV_Image.Rows.Count > 0)
                    {
                        byte[] image_data = imgToByteConverter(pictureBox1.Image);
                        //DGV_Image.CurrentRow.Cells["Image_Data"].Value = image_data;
                        DGV_Image.CurrentCell.Value = image_data;
                        MessageBox.Show("Inserted!");
                    }
                }
            }
            else
            {
                MessageBox.Show("Nhấn START để bắt đầu lấy hình ảnh!");
            }
        }

        private void cLEARLISTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clear_DGV(DGV_Image);
            tb_pcs.Text = "1";
        }

        private void login_Click(object sender, EventArgs e)
        {
            if (login.Text == "Admin Login")
            {
                btn_ok1.Show();
                tb_login.Show();
                MessageBox.Show("Nhập mật khẩu để đăng nhập!");
                tb_login.Focus();
                //   lb_login.Text = "Password:";
                login.Text = "Admin Login";


            }
            else
            {
                login.Text = "Admin Login";
                tb_login.Text = "";
                // lb_login.Text = "";
            }

        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            string app_path = Application.StartupPath;
            string file_pass = Path.Combine(app_path, "Config", "Pass_Camera.txt");
            string pass = File.ReadLines(file_pass).Skip(0).First().Split(':')[1].Trim();
            if (tb_login.Text == "TDMK" || tb_login.Text == pass)
            {
                login.Text = "Login Successful!";
                tb_login.Text = "";
                //lb_login.Text = "";
                btn_ok1.Hide();
                tb_login.Hide();
                //  lb_login.Hide();

                if (state_setting == true)
                {
                    btn_setting.PerformClick();
                }
                tb_pcs_set.Enabled = true;
            }
            else
            {
                MessageBox.Show("Mật khẩu sai!");
                tb_login.Focus();
            }
        }

        private void cb_region_SelectedIndexChanged(object sender, EventArgs e)
        {
            tb_pcs.Text = "1";
            string data_for = cb_data_for.Text;
            if (btn_load_database.Text != "LOADED")
            {
                if (data_for == "CQRA_HIGH_SPEED_BALL_SHEAR")
                {
                    btn_load.Enabled = true;
                }
            }
            else
            {
                if ((data_for != "CQRA_DIELECTRIC_WITHSTANDING") && (data_for != "CQRA_HIGH_SPEED_BALL_SHEAR"))
                {
                    DGV_Image.DataSource = null;
                }
            }

            if (tb_itemcode.Text.Trim() != "")
            {
                try
                {
                    int pcs_num = Get_setting_pcs(tb_itemcode.Text, cb_data_for.Text, cb_region.Text, sqlcon);
                    if (pcs_num > 0)
                    {
                        tb_pcs_set.Text = pcs_num.ToString();
                        tb_pcs_set.Enabled = false;
                        btn_save.Enabled = true;
                    }
                    else
                    {
                        if (cb_region.Text != "Graph")
                        {
                            tb_pcs_set.Enabled = true;
                            MessageBox.Show("Chưa có dữ liệu cài đặt cho hạng mục này ", "Warming");
                            tb_pcs_set.Text = "";
                            btn_save.Enabled = false;
                        }
                        else
                        {
                            tb_pcs_set.Enabled = false;
                            btn_save.Enabled = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error");
                }
            }
        }
        public int Get_setting_pcs(string ItemCode, string data_for, string region, SqlConnection sqlcon)
        {
            int result = -1;
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Data_For", "Region_Zone" }, new string[] { ItemCode, data_for, region });
            AutoCompleteStringCollection pcs_lst = TDMK_Code.Load_Item_Filter_str(sqlcon, "SETTING_PCS", "Pcs_setting", filter_str);
            if (pcs_lst.Count > 0)
            {
                string temp = pcs_lst[0];
                if (TDMK_Code.IsNumeric(temp))
                {
                    result = Convert.ToInt32(temp);
                }
            }

            return result;
        }

        private void cb_data_for_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                cb_region.Focus();
            }
        }

        private void cb_region_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                tb_pcs.Focus();
            }
        }

        private void tb_pcs_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                tb_operator.Focus();
            }
        }
        private void tb_operator_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btn_capture.Focus();
            }
        }
        private void tb_login_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btn_ok1.Focus();
                btn_ok1.PerformClick();
            }
        }
        private void DGV_Image_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                int col_inx = DGV_Image.CurrentCell.ColumnIndex;
                if (DGV_Image.Columns[col_inx].Name.Contains("Image"))
                {
                    DGV_Image.ContextMenuStrip = contextMenuStrip2;
                    //contextMenuStrip2.Show();
                }
                else
                {
                    DGV_Image.ContextMenuStrip = null;
                }    

            }
        }
        public myExcel.Workbook create_export_wrk_old(string format_file, string tar_sheet, string process_name, string ItemCode, string Lotno)
        {
            //string format_file = @"D:\Customer Projects\SEEV\Giai Doan 2\TDMK Source Code\TDMK-Diem\08Apr23\Get_Export_Image\TestCam\bin\Debug\Format\821-04380-03 718780.xlsm";
            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            myExcel.Workbook wb = TDMK_Code.open_excel_file(format_file, "", "");
            string export_path = Path.Combine(Application.StartupPath, "Export", process_name);
            if (!Directory.Exists(export_path))
            {
                Directory.CreateDirectory(export_path);
            }
            myExcel.Workbook save_wb = TDMK_Code.Create_workbook();
            wb.Sheets[tar_sheet].Copy(After: save_wb.Sheets[1]);
            save_wb.SaveAs(Path.Combine(export_path, ItemCode + "-" + Lotno));
            wb.Close();
            return save_wb;
        }
        public myExcel.Workbook create_export_wrk(string format_file, string process_name, string ItemCode, string Lotno)
        {
            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            myExcel.Workbook wb = TDMK_Code.open_excel_file(format_file, "", "");
            string _process = process_name.Replace("_", "");
            string mySheet = "";
            foreach (myExcel.Worksheet tg_sht in wb.Worksheets)
            {
                string cur_sht_name = tg_sht.Name.ToUpper().Replace("-", "").Replace(" ", "");
                if (cur_sht_name == _process)
                {
                    mySheet = tg_sht.Name;
                    break;
                }
            }
            if (mySheet != "")
            {
                string export_path = Path.Combine(data_loc, diem.Find_Export_Path(format_file, process_name));
                if (!Directory.Exists(export_path))
                {
                    Directory.CreateDirectory(export_path);
                }
                myExcel.Workbook save_wb = TDMK_Code.Create_workbook();
                wb.Sheets[mySheet].Copy(After: save_wb.Sheets[1]);
                save_wb.SaveAs(Path.Combine(export_path, ItemCode + "-" + Lotno));
                wb.Close();
                return save_wb;
            }
            else
            {
                wb.Close();
                return null;
            }
        }

        private void tsmExport_Click(object sender, EventArgs e)
        {
            btn_load_database.PerformClick();
            string sheet = cb_data_for.SelectedItem.ToString();
            if (sheet.Contains("ACF"))
                sheet = "OQC ACF";
          // Export_Data();
            string format_file = exp_proc.find_format(myVar.data_loc, tb_itemcode.Text);
            if(format_file != "")
            {
                string export_path = F_exportEPPlus.get_reportpath(tb_itemcode.Text, tb_lotno.Text, tb_operator.Text, sheet , myVar.data_loc);
                if(export_path != "")
                {
                    F_exportEPPlus.Export_EPPlus(sqlcon, format_file, export_path, tb_itemcode.Text, tb_lotno.Text, sheet, null);
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Folder " + sheet +  " not found!", "Warning");
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Format not found!", "Warning");
            }
           
        }

       
        public void save_report_new(string ItemCode, string LotNo, string id, myExcel.Workbook curr_wrkbook, string sheet, string fpath)
        {
            string report_path = "";
            DirectoryInfo tar_parent = new DirectoryInfo(fpath);
            DirectoryInfo[] arr_dic_child = tar_parent.GetDirectories();
            foreach (DirectoryInfo folder_sheet in arr_dic_child)
            {
                string f_name = folder_sheet.Name;
                if (f_name.Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper().Contains(sheet.Replace("_", "").Replace("-", "").Replace(" ", "").ToUpper()))
                {
                    report_path = Path.Combine(fpath, f_name, ItemCode + "-" + LotNo + "_" + id);
                    if (File.Exists(report_path))
                        File.Delete(report_path);


                    curr_wrkbook.SaveAs(report_path);
                    break;
                }
            }

        }

        private void FrmCamera_FormClosed(object sender, FormClosedEventArgs e)
        {
            myVar._frmMain.Show();
        }
        private void btn_load_database_TextChanged(object sender, EventArgs e)
        {
            if (btn_load_database.Text == "LOAD FROM DATABASE")
            {
                btn_save.Enabled = true;
            }
        }

        private void cb_data_for_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Config_DGV(DGV_Image, "CQRA_BHAST");    
            if (cb_data_for.SelectedIndex != -1)
            {
                tb_pcs.Text = "1";
                if (cb_data_for.Text == "OTHER")
                {
                    cb_region.Text = "Image";
                }
                else
                {
                    try
                    {
                        cb_region.SelectedIndex = 0;
                    }
                    catch { }
                }
                data_for_change();
                if (cb_data_for.SelectedItem.ToString() == "CQRA_SOLDERABILITY" || cb_data_for.SelectedItem.ToString() == "CQRA_CHEMICAL_RESISTANCE")
                {
                    btn_Judgement.Visible = true;
                }
                else
                {
                    btn_Judgement.Visible = false;
                }

                btn_load_database.Text = "LOAD FROM DATABASE";
            }

            //btn_save.Enabled = true;
        }
        private void DGV_Image_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                byte[] data = (byte[])DGV_Image.CurrentCell.Value;
                using (MemoryStream ms = new MemoryStream(data))
                {
                    CapturedImage.Image = Image.FromStream(ms);
                }
            }
            catch (Exception)
            {

            }
        }
        public void Auto_Resize_DGV()
        {
            // Đặt chế độ tự động điều chỉnh độ rộng của các cột
            DGV_Image.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            // Đặt chế độ tự động điều chỉnh độ cao của các hàng
            //DGV_Image.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
        }



        public void Export_Data()
        {
            string format_file = exp_proc.find_format(myVar.data_loc, tb_itemcode.Text);
            if (format_file != "")
            {
                List<string> search_item = new List<string>() { };
                switch (cb_data_for.SelectedItem.ToString())
                {
                    case "CQRA_BHAST":
                        wb = exp_proc.create_export_wrk(format_file, cb_data_for.SelectedItem.ToString(), tb_itemcode.Text, tb_lotno.Text);
                        //diem.Export_CQRA_bHast_Image(sqlcon, wb, tb_itemcode.Text, tb_lotno.Text);
                        F_export_new.Export_CQRA_bHast_Image(sqlcon, wb, tb_itemcode.Text, tb_lotno.Text);
                        // wb.Close();
                        break;
                    case "CQRA_SOLDERABILITY":
                        wb = exp_proc.create_export_wrk(format_file, cb_data_for.SelectedItem.ToString(), tb_itemcode.Text, tb_lotno.Text);
                        F_export_new.export_CQRA_Solderability_Image(sqlcon, wb, tb_itemcode.Text, tb_lotno.Text);
                        // diem.Export_CQRA_Solderability_Image(sqlcon, wb, tb_itemcode.Text, tb_lotno.Text);
                        // wb.Close();
                        break;
                    case "CQRA_DIELECTRIC_WITHSTANDING":
                        wb = exp_proc.create_export_wrk(format_file, cb_data_for.SelectedItem.ToString(), tb_itemcode.Text, tb_lotno.Text);
                        if (wb == null)
                        {
                            wb = exp_proc.create_export_wrk(format_file, "CQRA-Dielectric withstand", tb_itemcode.Text, tb_lotno.Text);
                        }
                        // wb = TDMK_Code.open_excel_file(@"D:\\OK2SHIP\\Build\\0.OK2SHIP report format\\Ok2ship Request SEEV D38 Sphinx SP-F Flex_P2(C4.0) 22B0125.xlsm", "", "");
                        // diem.Export_CQRA_Dielectric_withstanding(sqlcon, wb, tb_itemcode.Text, tb_lotno.Text
                        F_export_new.Export_CQRA_Dielectric_withstanding(sqlcon, wb, tb_itemcode.Text, tb_lotno.Text);
                        // wb.Close();
                        break;
                    case "CQRA_HIGH_SPEED_BALL_SHEAR":
                        wb = exp_proc.create_export_wrk(format_file, cb_data_for.SelectedItem.ToString(), tb_itemcode.Text, tb_lotno.Text);
                        diem.Export_High_Speed_Ball_Shear_image_new(sqlcon, wb, tb_itemcode.Text, tb_lotno.Text);
                        // wb.Close();
                        break;
                    case "OQC_TEST":
                        wb = exp_proc.create_export_wrk(format_file, cb_data_for.SelectedItem.ToString(), tb_itemcode.Text, tb_lotno.Text);
                        diem.Export_OQC_Test_Image(sqlcon, wb, tb_itemcode.Text, tb_lotno.Text);
                        //wb.Close();
                        break;
                    case "CQRA_CHEMICAL_RESISTANCE":
                        wb = exp_proc.create_export_wrk(format_file, cb_data_for.SelectedItem.ToString(), tb_itemcode.Text, tb_lotno.Text);
                        if (wb == null)
                        {
                            wb = exp_proc.create_export_wrk(format_file, "CQRA - Chemical resist", tb_itemcode.Text, tb_lotno.Text);
                        }
                        F_export_new.Export_Chemical_resist_Image(sqlcon, wb, tb_itemcode.Text, tb_lotno.Text); 

                        break;
                    case "CQRA_IR_VIA_TO_VIA":
                        wb = exp_proc.create_export_wrk(format_file, cb_data_for.Text, tb_itemcode.Text, tb_lotno.Text);
                        diem.Export_IR_via_to_via_Image(sqlcon, wb, tb_itemcode.Text, tb_lotno.Text);
                        //wb.Close();
                        break;
                    case "CQRA_IR_TRACE_TO_TRACE":
                        wb = exp_proc.create_export_wrk(format_file, cb_data_for.SelectedItem.ToString(), tb_itemcode.Text, tb_lotno.Text);
                        diem.Export_IR_trace_to_trace_Image(sqlcon, wb, tb_itemcode.Text, tb_lotno.Text);
                        // wb.Close();
                        break;
                    case "CQRA_IR_LAYER_TO_LAYER":
                        wb = exp_proc.create_export_wrk(format_file, cb_data_for.SelectedItem.ToString(), tb_itemcode.Text, tb_lotno.Text);
                        diem.Export_IR_layer_to_layer_Image(sqlcon, wb, tb_itemcode.Text, tb_lotno.Text);
                        //wb.Close();
                        break;
                    case "CQRA_FLUX_RESIST":
                        wb = exp_proc.create_export_wrk(format_file, "CQRA - Flux resistance", tb_itemcode.Text, tb_lotno.Text);
                        if (wb == null)
                        {
                            wb = exp_proc.create_export_wrk(format_file, "CQRA - Flux resist", tb_itemcode.Text, tb_lotno.Text);
                        }
                        diem.Export_Flux_Resist_Image(sqlcon, wb, tb_itemcode.Text, tb_lotno.Text);
                        //wb.Close();
                        break;
                    case "ACF_CLEANING":
                        wb = exp_proc.create_export_wrk(format_file, "ACF", tb_itemcode.Text, tb_lotno.Text);
                        if (wb == null)
                        {
                            wb = exp_proc.create_export_wrk(format_file, "OQC ACF", tb_itemcode.Text, tb_lotno.Text);
                        }
                        F_export_new.export_ACF_ALL_CAMERA(wb, tb_itemcode.Text, tb_lotno.Text, sqlcon);
                        //search_item = new List<string>() { "Photo before cleaning", "Photo after cleaning", "Photo after OQC testing" };
                        //Export_ACF_select(sqlcon, wb, tb_itemcode.Text, tb_lotno.Text, search_item);
                        //
                        // New format - new code
                        //diem.Export_ACF_Cleaning_Edit_22_9(sqlcon, wb, tb_itemcode.Text, tb_lotno.Text);
                        //F_export_new.Export_ACF_Cleaning(sqlcon, wb, tb_itemcode.Text, tb_lotno.Text);
                        break;

                    case "ACF_BEFORE_TAPE_TEST":
                        wb = exp_proc.create_export_wrk(format_file, "ACF", tb_itemcode.Text, tb_lotno.Text);
                        if (wb == null)
                        {
                            wb = exp_proc.create_export_wrk(format_file, "OQC ACF", tb_itemcode.Text, tb_lotno.Text);
                        }
                        F_export_new.export_ACF_ALL_CAMERA(wb, tb_itemcode.Text, tb_lotno.Text, sqlcon);
                        //search_item = new List<string>() { "Before tape test" };
                        //Export_ACF_select(sqlcon, wb, tb_itemcode.Text, tb_lotno.Text, search_item);
                        break;

                    case "ACF_AFTER_TAPE_TEST":
                        wb = exp_proc.create_export_wrk(format_file, "ACF", tb_itemcode.Text, tb_lotno.Text);
                        if (wb == null)
                        {
                            wb = exp_proc.create_export_wrk(format_file, "OQC ACF", tb_itemcode.Text, tb_lotno.Text);
                        }
                        F_export_new.export_ACF_ALL_CAMERA(wb, tb_itemcode.Text, tb_lotno.Text, sqlcon);
                        //search_item = new List<string>() { "After tape test" };
                        //Export_ACF_select(sqlcon, wb, tb_itemcode.Text, tb_lotno.Text, search_item);
                        break;

                    case "ACF_GLASS_COUPON":
                        wb = exp_proc.create_export_wrk(format_file, "ACF", tb_itemcode.Text, tb_lotno.Text);
                        if (wb == null)
                        {
                            wb = exp_proc.create_export_wrk(format_file, "OQC ACF", tb_itemcode.Text, tb_lotno.Text);
                        }
                        F_export_new.export_ACF_ALL_CAMERA(wb, tb_itemcode.Text, tb_lotno.Text, sqlcon);
                        break;

                    case "ACF_GRAPH_FORCE":
                        wb = exp_proc.create_export_wrk(format_file, "ACF", tb_itemcode.Text, tb_lotno.Text);
                        if (wb == null)
                        {
                            wb = exp_proc.create_export_wrk(format_file, "OQC ACF", tb_itemcode.Text, tb_lotno.Text);
                        }
                        F_export_new.export_ACF_ALL_CAMERA(wb, tb_itemcode.Text, tb_lotno.Text, sqlcon);
                        break;
                }

                save_report_new(tb_itemcode.Text, tb_lotno.Text, tb_operator.Text, wb, cb_data_for.SelectedItem.ToString(), myVar.data_loc);
                MessageBox.Show(new Form { TopMost = true }, "Export to checksheet complete !", "Warning");
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Format not found!", "Warning");
            }

        }
        public void Export_ACF_select(SqlConnection sqlcon, myExcel.Workbook wb, string Itemcode, string Lotno, List<string> search_item)
        {
            myExcel.Worksheet ws = wb.Sheets["ACF"];

            myExcel.Range sel_rgn = ws.Range["A16"];
            foreach (var item in search_item)
            {
                int num_insert = 5;
                string tar_table = "";
                string tar_col_data = "Image_Data";
                if (item.Contains("cleaning") || (item.Contains("OQC")))
                {
                    tar_table = "ACF_CLEANING_IMAGE";
                    num_insert = 10;
                }
                else
                {
                    if (item.Contains("picture"))
                    {
                        tar_table = "ACF_GRAPH_FORCE_IMAGE";
                        tar_col_data = "Image_Graph";
                    }
                    else
                    {
                        tar_table = "ACF_" + item.Trim().Replace(" ", "_").ToUpper() + "_IMAGE";
                    }
                }
                myExcel.Range data_rgn = ws.Range[diem.Find_addr(item, sel_rgn)].Offset[0, 1];
                DataTable data_tbl = TDMK_Code.Datatable_Filter(sqlcon, tar_table, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { Itemcode, Lotno }));
                List<DataTable> src_lst_tbl = new List<DataTable>();
                diem.Get_ListTable(-1, data_tbl, new string[] { "ItemCode", "LotNo", "Region" }, ref src_lst_tbl, tar_col_data);
                foreach (var dt in src_lst_tbl)
                {
                    if (dt.Rows.Count > 0)
                    {
                        string region = dt.Rows[0]["Region"].ToString();
                        int off_set = 0;
                        switch (region)
                        {
                            case "Picture":
                                off_set = 0;
                                break;
                            case "TRAI":
                                off_set = 0;
                                break;
                            case "GIUA":
                                off_set = 1;
                                break;
                            case "PHAI":
                                off_set = 2;
                                break;
                            default:
                                off_set = 0;
                                break;
                        }
                        int qty = Math.Min(dt.Rows.Count, num_insert);
                        for (int i = 0; i < qty; i++)
                        {
                            myExcel.Range cur_rgn = data_rgn.Offset[0, i + 5 * off_set];
                            byte[] data = (byte[])dt.Rows[i][tar_col_data];
                            using (MemoryStream ms = new MemoryStream(data))
                            {
                                string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                                if (!Directory.Exists(file_folder))
                                {
                                    Directory.CreateDirectory(file_folder);
                                }
                                string file_dic = Path.Combine(file_folder, "1.jpg");
                                Image image = Image.FromStream(ms);
                                image.Save(file_dic);
                                diem.InsertPicture_Name(ws, cur_rgn, file_dic, 10);
                            }
                            if (tar_col_data.Contains("Graph"))
                            {
                                myExcel.Range force_rgn = data_rgn.Offset[1, i + 5 * off_set];
                                force_rgn.Value = dt.Rows[i]["Force_Data"];
                            }
                        }
                    }
                }
                sel_rgn = data_rgn.Offset[0, -1];
            }
        }

        public void export_ACF_GRAPH_FORCE_old(string table_name, myExcel.Workbook wb, string Itemcode, string Lotno, SqlConnection sqlcon)
        {
            myExcel.Worksheet ws = wb.Sheets[1];
            int count = 0;
            string[] lst_region = new string[] { "TRAI", "GIUA", "PHAI" };


            for (int i = 50; i < 100; i++)
            {
                for (int j = 1; j < 15; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).ToString().Replace(" ", "").ToUpper().Contains("Force-displacement profile".Replace(" ", "").ToUpper()))
                    {
                        string zone = lst_region[count];
                        string[] item = { "ItemCode", "LotNo", "Region" };
                        string[] item_val = { Itemcode, Lotno, zone };
                        System.Data.DataTable data_image = TDMK_Code.Datatable_Filter(sqlcon, table_name, TDMK_Code.filter_str(item, item_val));

                        int number_image = new int[] { 5, data_image.Rows.Count }.Min();

                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        string file_dic = Path.Combine(file_folder, "1.jpg");

                        myExcel.Range curr_rgn = ws.Cells[i + 1, j];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image.Rows[m]["Image_Graph"];
                            string value = data_image.Rows[m]["Force_Data"].ToString();
                            MemoryStream ms = new MemoryStream(data);
                            Image image = Image.FromStream(ms);
                            image.Save(file_dic);
                            diem.InsertPicture_Name(ws, curr_rgn, file_dic, 6);
                            curr_rgn.Offset[0, 1].Value = value;
                            curr_rgn = curr_rgn.Offset[1, 0];
                        }

                        MessageBox.Show(new Form { TopMost = true }, "Hoàn thành!");

                        count++;
                    }
                    if (count == 3)
                        return;
                }

            }
        }

        public void export_ACF_GRAPH_FORCE(myExcel.Workbook wb, string Itemcode, string Lotno, SqlConnection sqlcon)
        {
            myExcel.Worksheet ws = wb.Sheets[1];
            int count = 0;
            string[] lst_region = new string[] { "TRAI", "GIUA", "PHAI" };


            for (int i = 50; i < 100; i++)
            {
                for (int j = 1; j < 15; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).ToString().Replace(" ", "").ToUpper().Contains("Force-displacement profile".Replace(" ", "").ToUpper()))
                    {
                        string zone = lst_region[count];
                        string[] item = { "ItemCode", "LotNo", "Region" };
                        string[] item_val = { Itemcode, Lotno, zone };
                        DataTable data_image_force = TDMK_Code.Datatable_Filter(sqlcon, "ACF_GRAPH_FORCE_IMAGE", TDMK_Code.filter_str(item, item_val));
                        DataTable data_image_tape_test = TDMK_Code.Datatable_Filter(sqlcon, "ACF_AFTER_TAPE_TEST_IMAGE", TDMK_Code.filter_str(item, item_val));

                        int number_image = new int[] { 5, data_image_force.Rows.Count }.Min();

                        string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
                        string file_dic = Path.Combine(file_folder, "1.jpg");

                        myExcel.Range curr_rgn = ws.Cells[i + 1, j];
                        for (int m = 0; m < number_image; m++)
                        {
                            byte[] data = (byte[])data_image_force.Rows[m]["Image_Graph"];
                            byte[] data_tape_test = (byte[])data_image_tape_test.Rows[m]["Image_Data"];
                            string value = data_image_force.Rows[m]["Force_Data"].ToString();

                            // export graph
                            MemoryStream ms = new MemoryStream(data);
                            Image image = Image.FromStream(ms);
                            image.Save(file_dic);
                            diem.InsertPicture_Name(ws, curr_rgn, file_dic, 6);


                            // export tape test
                            MemoryStream ms2 = new MemoryStream(data_tape_test);
                            Image image2 = Image.FromStream(ms2);
                            image2.Save(file_dic);
                            diem.InsertPicture_Name(ws, curr_rgn.Offset[0, 4], file_dic, 6);



                            curr_rgn.Offset[0, 1].Value = value;
                            curr_rgn = curr_rgn.Offset[1, 0];
                        }

                        MessageBox.Show(new Form { TopMost = true }, "Hoàn thành!");

                        count++;
                    }
                    if (count == 3)
                        return;
                }

            }
        }

        public void Message_Done(object sender, DoWorkEventArgs e)
        {
            MessageBox.Show("Hoàn thành!", "Thông báo!");
        }

        private void cb_mode_CheckedChanged(object sender, EventArgs e)
        {
            tb_pcs.Text = "1";
        }

        private void btn_delete_all_Click(object sender, EventArgs e)
        {

            TDMK_Code.Delelte_FilteredItem_arr(cb_data_for.Text + "_IMAGE", sqlcon, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { tb_itemcode.Text, tb_lotno.Text }));
            Clear_DGV(DGV_Image);
            MessageBox.Show("Deleted all images of item " + tb_itemcode.Text);

        }

        private void splitContainer42_SplitterMoved(object sender, SplitterEventArgs e)
        {
        }
        public void DGV_Image_Fitting(DataGridView src_DGV)
        {
            foreach (DataGridViewColumn dc in src_DGV.Columns)
            {
                string col_name = dc.Name;
                if (col_name.ToUpper().Contains("IMAGE"))
                {
                    ((DataGridViewImageColumn)dc).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    dc.Width = 200;
                }
            }
            foreach (DataGridViewRow dr in src_DGV.Rows)
            {
                dr.Height = 150;
            }
        }

        private void tb_lotno_Validated(object sender, EventArgs e)
        {
            tb_lotno.Text = exp_proc.Lotno_Formated(tb_lotno.Text);
        }

        private void btn_setting_Click_1(object sender, EventArgs e)
        {
            state_setting = true;
            if (login.Text == "Login Successful!")
            {
                Form form = new Setting_Pcs_Number();
                form.Show();
            }
            else
            {
                login.PerformClick();
            }

        }

        private void btn_replace_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Chọn một tệp ảnh";
            openFileDialog.Filter = "Tệp ảnh (*.jpg; *.jpeg; *.png; *.gif) | *.jpg; *.jpeg; *.png; *.gif";

            // Mở hộp thoại và xử lý kết quả
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string fileName = openFileDialog.FileName;
                    byte[] image = diem.File_image_to_byte(fileName);
                    DGV_Image.CurrentCell.Value = image;
                    MessageBox.Show("Replace image successfully!", "Message");
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message, "Error");
                }
            }
        }

        private void menuStrip2_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void menu_replace_image_Opening(object sender, CancelEventArgs e)
        {

        }

        private void replaceFromFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Chọn một tệp ảnh";
            openFileDialog.Filter = "Tệp ảnh (*.jpg; *.jpeg; *.png; *.gif) | *.jpg; *.jpeg; *.png; *.gif";

            // Mở hộp thoại và xử lý kết quả
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string fileName = openFileDialog.FileName;
                    byte[] image = diem.File_image_to_byte(fileName);
                    DGV_Image.CurrentCell.Value = image;
                    MessageBox.Show("Replace image successfully!", "Message");
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message, "Error");
                }
            }
        }

        private void tb_itemcode_KeyDown(object sender, KeyEventArgs e)
        {
            if (tb_itemcode.Text.Trim() != "")
            {
                if (e.KeyCode == Keys.Enter)
                {
                    tb_lotno.Focus();
                }
            }
        }

        private void DGV_Image_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        public SortedDictionary<int, High_Speed_Ball_Shear_Logfile> High_speed_Image_LogFile_process(string Itemcode, string Lot, string folder, SqlConnection sqlcon)
        {
            SortedDictionary<int, High_Speed_Ball_Shear_Logfile> result = new SortedDictionary<int, High_Speed_Ball_Shear_Logfile>();
            string[] file_image = Directory.GetFiles(folder, "*.JPG");
            string[] file_data = Directory.GetFiles(folder, "*.csv");
            Dictionary<string, byte[]> image_data_lst = new Dictionary<string, byte[]>();
            Dictionary<string, string> data_val_lst = new Dictionary<string, string>();
            foreach (string file in file_image)
            {
                Byte[] data = new Byte[0];
                Image myImg = Image.FromFile(file);
                ImageConverter imgCon = new ImageConverter();
                data = (byte[])imgCon.ConvertTo(myImg, typeof(byte[]));
                string pcs_no = (Path.GetFileName(file)).Split('.')[0];
                if (image_data_lst.Keys.ToList().IndexOf(pcs_no) == -1)
                {
                    image_data_lst.Add(pcs_no, data);
                }
            }
            foreach (string f_data in file_data)
            {
                myExcel.Workbook tar_wrk = TDMK_Code.open_excel_file(f_data, "", "");
                myExcel.Worksheet tar_sht = tar_wrk.Sheets[1];
                myExcel.Range id_rgn = tar_sht.Range["E6"];
                myExcel.Range data_rgn = tar_sht.Range["F6"];
                int r_inx = 0;
                while (myCode.checkDBNull(id_rgn.Offset[r_inx, 0].Value) != "")
                {
                    string inx = Convert.ToString(id_rgn.Offset[r_inx, 0].Value);
                    List<string> temp = data_val_lst.Keys.ToList();
                    if (data_val_lst.Keys.ToList().IndexOf(inx) == -1)
                    {
                        data_val_lst.Add(inx, myCode.checkDBNull(data_rgn.Offset[r_inx, 0].Value));
                    }
                    r_inx++;
                }
                tar_wrk.Close();
            }
            foreach (var img in image_data_lst.Keys)
            {
                if (data_val_lst.Keys.ToList().IndexOf(img) != -1)
                {
                    result.Add(Convert.ToInt32(img), new High_Speed_Ball_Shear_Logfile(image_data_lst[img], data_val_lst[img]));
                }
            }
            return result;
        }
        public SortedDictionary<int, byte[]> High_speed_Photo_process(string Itemcode, string Lot, string folder, SqlConnection sqlcon)
        {
            SortedDictionary<int, byte[]> result = new SortedDictionary<int, byte[]>();
            string[] file_image = Directory.GetFiles(folder, "*.bmp");
            SortedDictionary<int, byte[]> image_data_lst = new SortedDictionary<int, byte[]>();
            foreach (string file in file_image)
            {
                Byte[] data = new Byte[0];
                Image myImg = Image.FromFile(file);
                ImageConverter imgCon = new ImageConverter();
                data = (byte[])imgCon.ConvertTo(myImg, typeof(byte[]));
                string pcs_no = (Path.GetFileName(file)).Split('.')[0];
                int pcs_inx = Convert.ToInt32(pcs_no);
                if (image_data_lst.Keys.ToList().IndexOf(pcs_inx) == -1)
                {
                    image_data_lst.Add(pcs_inx, data);
                }
            }
            int item_key = 1;
            foreach (var temp in image_data_lst)
            {
                List<Bitmap> split_img = SplitImages(byteArrayToImage(temp.Value), 2);
                foreach (var t in split_img)
                {
                    Byte[] temp_data = new Byte[0];
                    ImageConverter imgCon = new ImageConverter();
                    temp_data = (byte[])imgCon.ConvertTo(t, typeof(byte[]));
                    result.Add(item_key, temp_data);
                    item_key++;

                }
            }
            return result;
        }
        public List<Bitmap> SplitImages(Image src_img, int num)
        {
            List<Bitmap> result = new List<Bitmap>();
            int w = src_img.Width;
            int h = (int)(src_img.Height * 0.9);
            int off_set = (int)(w / num);
            Rectangle myrec = new Rectangle(0, 0, off_set, h);
            for (int i = 0; i < num; i++)
            {
                Bitmap bmp = new Bitmap(myrec.Width, myrec.Height);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.DrawImage(src_img, myrec, new Rectangle(i * off_set, 0, off_set, h), GraphicsUnit.Pixel);
                }
                result.Add(bmp);
            }
            return result;
        }

        private void DGV_Image_DataSourceChanged(object sender, EventArgs e)
        {
            if (DGV_Image.DataSource == null)
            {
                btn_save.Enabled = false;
            }
            else
            {
                if ((cb_data_for.Text == "CQRA_HIGH_SPEED_BALL_SHEAR") && (cb_region.Text == "Photo"))
                {
                    btn_save.Enabled = false;
                }
                else
                {
                    btn_save.Enabled = true;
                }
            }
        }

        private void tsmExport_Click_1(object sender, EventArgs e)
        {

        }

        private void splitContainer29_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }
        public void judgement_image(string value)
        {
            if (value != "")
            {
                foreach (DataGridViewRow dgv_r in DGV_Image.Rows)
                {
                    foreach (string item in value.TrimEnd(';').Split(';'))
                    {
                        if (dgv_r.Cells["Pcs_No"].Value.ToString() == item.Split('-')[0])
                        {
                            dgv_r.Cells["Remark"].Value = item.Split('-')[1];

                            if (sqlcon.State != ConnectionState.Open)
                            {
                                sqlcon.Open();
                            }

                            SqlCommand cmd = new SqlCommand();
                            cmd.Connection = sqlcon;
                            cmd.CommandText = "Update " + cb_data_for.SelectedItem.ToString() + "_IMAGE " + "SET Remark = '" + item.Split('-')[1] + "' WHERE ItemCode = '" + tb_itemcode.Text + "' and LotNo = '" + tb_lotno.Text + "' and Pcs_No = '" + item.Split('-')[0] +  "' ;";
                            cmd.ExecuteNonQuery();
                            sqlcon.Close();
                            break;
                        }
                    }
                }
            }
        }

        private void btn_Judgement_Click(object sender, EventArgs e)
        {
            Judgement_Image frm = new Judgement_Image(judgement_image);

            DataTable dt1 = new DataTable();
            if (DGV_Image.DataSource != null)
            {
                dt1 = (DataTable)DGV_Image.DataSource;
            }

            if (dt1.Rows.Count > 0)
            {
                List<string> lst_reg = dt1.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToList();
                if (cb_data_for.SelectedItem.ToString() == "CQRA_CHEMICAL_RESISTANCE" && (!lst_reg.Contains("Picture before") || !lst_reg.Contains("Picture after")))
                {
                    MessageBox.Show(new Form { TopMost = true }, "Chưa đủ ảnh before/after", "Thông báo");
                    return;
                }
                else
                {
                    frm.data_tbl_ = dt1;
                    frm.sheet_ = cb_data_for.SelectedItem.ToString();
                    frm.Show();
                }
            }

        }
    }
}

