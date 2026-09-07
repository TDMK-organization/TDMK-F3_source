using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SEEV_DLL;
using TDMK_SQL;
using myExcel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using Microsoft.Office.Core;
using System.Security.Cryptography;
using Patagames.Ocr;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip;
using System.Web;
using System.Reflection;
using System.Runtime.InteropServices;
//using Retangle = System.Drawing.Rectangle;
using OK2SHIP_SMT;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
//using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using System.Windows.Forms.VisualStyles;
using System.Diagnostics.Eventing.Reader;
//using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using System.Drawing.Drawing2D;
using System.Net.NetworkInformation;
using System.Diagnostics;
using Bending_Export;
using System.Drawing.Printing;
using System.ComponentModel.Design;
using Patagames.Ocr.InputFilters;
using System.Xml.Linq;
using IniLibs;
using static System.Net.Mime.MediaTypeNames;
using Image = System.Drawing.Image;
using System.Runtime.ConstrainedExecution;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using OfficeOpenXml.Style;
using System.Xml;
using OK2SHIP_SMT.Libary;
using OK2SHIP_SMT.Services;
using ZedGraph;
using OK2SHIP_SMT.Repositories;
using System.Drawing.Imaging;

//using TDMK_EPPLUS_7;


namespace Funtion_F3_SMT
{
    public partial class View_Data_3 : Form
    {
        //  public Bending_Export_Lib Bending_Exp = new Bending_Export_Lib();
        public Bending_Export_Lib Bending_Exp = new Bending_Export_Lib();
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();

        public TDMK_EPPLUS TDMK_EPPLUS = new TDMK_EPPLUS();

        public SEI_Lib myCode = new SEI_Lib();
        Funtion_SMT my_SMT = new Funtion_SMT();
        Funtion_export_file_le_EPPlus F_export_EPPlus = new Funtion_export_file_le_EPPlus();
        public string sheet = "";

        IniFile TDMK_init;

        //public string admin_mode = "";
        string[] arr_ignored = new string[2];
        string[] arr_onproduct = new string[2];
        string[] arr_comment_2 = new string[2];
        string tbl_name_comment3 = "COMMENT_3_NEW";
        string admin_mode = "LOGIN";
        string data_loc = "";
        byte[] img_null = null;
        Image image_null = new Bitmap(AppContext.BaseDirectory + "\\null.png");
        string col_name_click = "";
        string server_name = "";
        string server_acc = "";
        string server_pass = "";

        bool edit_mode = false;
        bool hide_mode_logfile = true;
        bool hide_mode_analysis = true;
        bool judge_all = true;

        int set_chan = 0;


        public SqlConnection sqlcon = null;
        public string strcon = "";
        string DB_name = "OK2SHIP_SMT";

        public View_Data_3()
        {
            InitializeComponent();
            checkSession();
        }


        public string sheet_
        {
            get { return sheet; }
            set { sheet = value; }
        }


        public string admin_mode_
        {
            get { return admin_mode; }
            set { admin_mode = value; }
        }


        public void BatchBulkCopy(SqlConnection sqlcon, DataTable dataTable, string tablename)
        {
            DataTable dtInsertRows = dataTable;
            using (SqlBulkCopy sbc = new SqlBulkCopy(sqlcon))
            {
                if (sqlcon.State != ConnectionState.Open)
                {
                    sqlcon.Open();
                }

                sbc.DestinationTableName = tablename;
                foreach (DataColumn dc in dataTable.Columns)
                {
                    sbc.ColumnMappings.Add(dc.ColumnName, dc.ColumnName);
                }

                sbc.WriteToServer(dtInsertRows);
                sqlcon.Close();
            }
        }

        public void resize_column_image(DataGridView dgv)
        {
            if (dgv.Rows.Count > 0)
            {
                if (sheet == "CROSS_SECTION")
                {
                    // Duyệt qua tất cả các cột của DataGridView
                    foreach (DataGridViewColumn col in dgv.Columns)
                    {
                        // Kiểm tra nếu cột là cột ảnh (DataGridViewImageColumn)
                        if (col is DataGridViewImageColumn imgCol)
                        {
                            imgCol.ImageLayout = DataGridViewImageCellLayout.Zoom;
                            imgCol.Width = 100;
                        }
                    }

                    // Đặt chiều cao cho các dòng
                    foreach (DataGridViewRow dr in dgv.Rows)
                    {
                        dr.Height = 70;
                    }
                }
                else if (sheet == "GAP_CONNECTOR")
                {
                    // Duyệt qua tất cả các cột của DataGridView
                    foreach (DataGridViewColumn col in dgv.Columns)
                    {
                        // Kiểm tra nếu cột là cột ảnh (DataGridViewImageColumn)
                        if (col is DataGridViewImageColumn imgCol)
                        {
                            imgCol.ImageLayout = DataGridViewImageCellLayout.Zoom;
                            imgCol.Width = 100;
                        }
                    }

                    // Đặt chiều cao cho các dòng
                    foreach (DataGridViewRow dr in dgv.Rows)
                    {
                        dr.Height = 70;
                    }
                }
                else
                {
                    ((DataGridViewImageColumn)dgv.Columns["Image"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)dgv.Columns["Image"]).Width = 100;

                    ((DataGridViewImageColumn)dgv.Columns["Graph"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                    ((DataGridViewImageColumn)dgv.Columns["Graph"]).Width = 100;

                    foreach (DataGridViewRow dr in dgv.Rows)
                    {
                        dr.Height = 70;
                    }
                }
            }
        }

        public void resize_column_image_cross_section(DataGridView dgv)
        {
            if (dgv.Rows.Count > 0)
            {
                ((DataGridViewImageColumn)dgv.Columns["Image1"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)dgv.Columns["Image1"]).Width = 100;

                ((DataGridViewImageColumn)dgv.Columns["Image2"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)dgv.Columns["Image2"]).Width = 100;

                foreach (DataGridViewRow dr in dgv.Rows)
                {
                    dr.Height = 70;
                }
            }
        }

        public SortedDictionary<int, string> Get_csv_data(string in_src_file)
        {
            SortedDictionary<int, string> result_lst = new SortedDictionary<int, string>();
            string[] Lines = System.IO.File.ReadAllLines(in_src_file);
            foreach (var line in Lines)
            {
                string str = new string(line.Where(s => s != '"').ToArray());
                string[] temp = str.Split(',');
                if (temp[0] != "")
                {
                    if (myCode.IsNumeric(temp[0]))
                    {
                        result_lst.Add(Convert.ToInt32(temp[0]), temp[2]);
                    }
                }
                else
                {
                    break;
                }
            }

            return result_lst;
        }

        public void PasteClipboardValue(bool _transpose, DataGridView tar_DGV)
        {
            if (tar_DGV.SelectedCells.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn ô", "Paste", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }


            Dictionary<string, string> dic_mode = new Dictionary<string, string> { };
            dic_mode.Add("Mode 1#", "Mode 1: Solder joint crack");
            dic_mode.Add("Mode 2#", "Mode 2: Pad lift");
            dic_mode.Add("Mode 3#", "Mode 3: Solder joint lift");
            dic_mode.Add("Mode 4#", "Mode 4: Intermetallic break");
            dic_mode.Add("Mode 5#", "Mode 5: Component damage");
            dic_mode.Add("Mode 6#", "Mode 6: Component detached");
            dic_mode.Add("Mode 7#", "Mode 7: Flex torn");


            List<string> lst_col_name = new List<string>
            {
                "Solder joint crack", "Pad lift", "Solder joint lift", "Intermetallic break", "Component damage",
                "Component detached", "Flex torn"
            };

            DataGridViewCell startCell = myCode.GetStartCell(tar_DGV);
            Dictionary<int, Dictionary<int, string>> dictionary = myCode.ClipBoardValues(Clipboard.GetText());
            if (dictionary.Count > 0)
            {
                if (!_transpose)
                {
                    int num = startCell.RowIndex;
                    if (num != -1)
                    {
                        foreach (int key in dictionary.Keys)
                        {
                            int num2 = startCell.ColumnIndex;
                            foreach (int key2 in dictionary[key].Keys)
                            {
                                if (num2 <= tar_DGV.Columns.Count - 1 && num <= tar_DGV.Rows.Count - 1)
                                {
                                    DataGridViewCell dataGridViewCell = tar_DGV[num2, num];
                                    if (dataGridViewCell.Selected)
                                    {
                                        dataGridViewCell.Value = dictionary[key][key2];

                                        double data_col = 0;
                                        for (int i = 2; i <= 7; i++)
                                        {
                                            if (myCode.IsNumeric(dgv_logfile.Rows[num]
                                                    .Cells[dic_mode["Mode " + i + "#"]].Value.ToString()
                                                    .Replace("%", "")))
                                            {
                                                data_col += double.Parse(dgv_logfile.Rows[num]
                                                    .Cells[dic_mode["Mode " + i + "#"]].Value.ToString()
                                                    .Replace("%", ""));
                                            }
                                        }

                                        tar_DGV.Rows[num].Cells[dic_mode["Mode 1#"]].Value =
                                            (100 - data_col).ToString() + "%";
                                    }
                                }

                                num2++;
                            }

                            num++;
                        }
                    }


                    return;
                }

                int num5 = startCell.ColumnIndex;
                foreach (int key3 in dictionary.Keys)
                {
                    int num6 = startCell.RowIndex;
                    foreach (int key4 in dictionary[key3].Keys)
                    {
                        if (num5 <= tar_DGV.Columns.Count - 1 && num6 <= tar_DGV.Rows.Count - 1)
                        {
                            DataGridViewCell dataGridViewCell2 = tar_DGV[num5, num6];
                            if (dataGridViewCell2.Selected)
                            {
                                dataGridViewCell2.Value = dictionary[key3][key4];
                            }
                        }

                        num6++;
                    }

                    num5++;
                }
            }
            else
            {
                MessageBox.Show("Vui lòng chọn dữ liệu!", "Thông báo");
            }
        }

        public Image get_pic(ExcelWorksheet wrksht, string pic_name)
        {
            Image result = null;
            List<string> list = new List<string>();
            foreach (ExcelPicture item in (IEnumerable<ExcelDrawing>)wrksht.Drawings)
            {
                list.Add(item.Name);
            }

            int num = list.IndexOf(pic_name);
            if (num != -1)
            {
                ExcelPicture excelPicture2 = wrksht.Drawings[list[num]] as ExcelPicture;
                byte[] imageBytes = excelPicture2.Image.ImageBytes;
                result = byteArrayToImage(imageBytes);
            }

            return result;
        }

        public Image byteArrayToImage(byte[] byteArrayIn)
        {
            MemoryStream stream = new MemoryStream(byteArrayIn);
            return Image.FromStream(stream);
        }

        public string Extract_data(Bitmap src_img)
        {
            string result = "";

            using (var api = OcrApi.Create())
            {
                api.Init(); // Languages.English);
                string plainText = api.GetTextFromImage(split_Image_rec(src_img));
                var tg = new string(plainText.Where(x => char.IsDigit(x)).ToArray());
                result = tg;
            }

            return result;
        }


        public Bitmap split_Image_rec(Image src_img)
        {
            Bitmap result = (Bitmap)src_img;
            int w = 60; // (int)src_img.Width * 7 / 100;
            int h = 30; // (int)src_img.Height * 4 / 100;
            System.Drawing.Rectangle src_rec = new System.Drawing.Rectangle(0, 0, w, h);
            Bitmap bmp = new Bitmap(w, h);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.DrawImage(src_img, src_rec, src_rec, GraphicsUnit.Pixel);
            }

            return bmp;
        }

        public byte[] File_Image_To_DGV_Value(string file)
        {
            Byte[] data = new Byte[0];
            Image myImg = System.Drawing.Image.FromFile(file);
            ImageConverter imgCon = new ImageConverter();
            data = (byte[])imgCon.ConvertTo(myImg, typeof(byte[]));
            return data;
        }

        public void show_messageform()
        {
            //if (message_error != "")
            //{
            //    Message_Error frm = new Message_Error();
            //    frm.message_ = message_error;
            //    frm.ShowDialog();
            //}
        }

        private bool _PRIME_PEEL_TEST = false;

        private void View_Data_3_Load(object sender, EventArgs e)
        {
            if (sheet.Equals("PEEL_TEST WITHOUT SUS"))
            {
                sheet = "PEEL_TEST";
                _PRIME_PEEL_TEST = true;
            }

            if (sheet.Equals("PULL_TEST"))
            {
                sheet = "MATING_PULL_TEST";
            }

            //Debugger.Break();
            string program_loc =
                F_export_EPPlus.find_config_path(System.Windows.Forms.Application.StartupPath, "TDMK Program");
            string config_path = Path.Combine(program_loc, "Config.ini");
            TDMK_init = new IniFile(config_path);
            //data_loc = F_export_EPPlus.find_config_path(System.Windows.Forms.Application.StartupPath, "SEEV Data");
            data_loc = TDMK_init.Read("Format_Folder", "SMT_Config") + "\\SEEV Data";
            server_name = TDMK_init.Read("Server", "SMT_Config");
            server_acc = TDMK_init.Read("Account", "SMT_Config");
            server_pass = TDMK_init.Read("Password", "SMT_Config");
            sqlcon = initial_data(DB_name, true);

            lbltitle.Text = sheet;
            arr_ignored = new string[]
            {
                "SHEARTEST", "IQC Liner peeling (Coupon)".Replace(" ", "").ToUpper(),
                "IQC PSA peeling (Coupon)".Replace(" ", "").ToUpper()
            };
            arr_onproduct = new string[]
            {
                "Liner peel test On product".Replace(" ", "_").ToUpper(),
                "PSA peel test On product".Replace(" ", "_").ToUpper()
            };
            arr_comment_2 = new string[]
                { "Cross section".Replace(" ", "_").ToUpper(), "GAP Connector".Replace(" ", "_").ToUpper() };
            var img = new Bitmap(10, 20);
            //var img = Bitmap.FromFile(Path.Combine(System.Windows.Forms.Application.StartupPath, "Img_null", "Img_null.jpg"));
            // var img = Bitmap.FromFile(Path.Combine(program_loc, "Img_null", "Img_null.jpg"));
            ImageConverter imgcon = new ImageConverter();
            img_null = (byte[])imgcon.ConvertTo(img, typeof(byte[]));
            btn_checkall.Visible = false;

            if (sheet.Contains("UNMATING") || sheet.Contains("COUPON"))
            {
                txt_line.Enabled = false;
                txt_ca.Enabled = false;
                txt_itemcode_nvl.Enabled = true;
                txt_lotno_nvl.Enabled = true;
                txtLotNo.Text = "NA";
                txtLotNo.BackColor = Color.Yellow;
                txtLotNo.Enabled = false;
            }
            else if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST" || sheet.Contains("ON_PRODUCT"))
            {
                lbl_itemcode_nvl.Text = "ItemCode(Refer)";
                lbl_lotno_nvl.Text = "LotNo(Refer)";
                lbl_itemcode_nvl.Enabled = false;
                lbl_lotno_nvl.Enabled = false;
                txt_itemcode_nvl.Enabled = false;
                txt_lotno_nvl.Enabled = false;
            }
            else
            {
                lbl_itemcode_nvl.Enabled = false;
                lbl_lotno_nvl.Enabled = false;
            }

            if (sheet == "CROSS_SECTION" || sheet == "GAP_CONNECTOR")
            {
                // 1. Khởi tạo TableLayoutPanel với 1 hàng và 2 cột
                TableLayoutPanel tlpStatus = new TableLayoutPanel();
                tlpStatus.RowCount = 1;
                tlpStatus.ColumnCount = 2;
                tlpStatus.Dock =
                    DockStyle.Top; // Dock lên trên cùng của Panel (bạn có thể đổi thành Fill nếu muốn lấp đầy)
                tlpStatus.Height = 35; // Chiều cao dự kiến vừa vặn cho ComboBox
                tlpStatus.Padding = new Padding(5); // Căn lề một chút cho đẹp

                // Thiết lập kích thước cho 2 cột: Cột 1 tự động vừa chữ, Cột 2 chiếm phần còn lại
                tlpStatus.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
                tlpStatus.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

                // 2. Khởi tạo Label "Status"
                System.Windows.Forms.Label lblStatus = new System.Windows.Forms.Label();
                lblStatus.Text = "Status";
                lblStatus.AutoSize = true;
                lblStatus.Anchor = AnchorStyles.Left; // Để Label canh giữa theo chiều dọc của hàng
                lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;

                // 3. Khởi tạo ComboBox
                cbStatus.DropDownStyle = ComboBoxStyle.DropDownList; // Chỉ cho phép chọn, không cho phép gõ text tự do
                if (sheet == "CROSS_SECTION")
                {
                    cbStatus.Items.Add("Normal");
                    cbStatus.Items.Add("Shield b2b");
                    cbStatus.Items.Add("Clip");
                }

                if (sheet == "GAP_CONNECTOR")
                {
                    cbStatus.Items.Add("Normal");
                    cbStatus.Items.Add("Shield b2b");
                }

                cbStatus.SelectedIndex = 0; // Đặt giá trị mặc định là "Normal" (vị trí index 0)
                cbStatus.Dock = DockStyle.Fill; // Để ComboBox trải dài hết cột thứ 2

                // 4. Thêm Label và ComboBox vào TableLayoutPanel
                // Cú pháp: Add(Control, columnIndex, rowIndex)
                tlpStatus.Controls.Add(lblStatus, 0, 0);
                tlpStatus.Controls.Add(cbStatus, 1, 0);

                // 5. Thêm TableLayoutPanel vào splitContainer2.Controls[1]
                // (Lưu ý: trong WinForms, splitContainer2.Panel2 chính là splitContainer2.Controls[1])
                splitContainer2.Panel2.Controls.Add(tlpStatus);
            }

            if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST" || sheet == "SHEAR_TEST")
            {
                txt_setchan.Visible = true;
                lbl_sochan.Visible = true;
            }
            else
            {
                txt_setchan.Visible = false;
                lbl_sochan.Visible = false;
            }
            //  data_loc = data_loc.Replace(@"\VHX-IMADA", "");
        }

        public SqlConnection initial_data_old(string DB_name, bool sa_en)
        {
            SqlConnection _sqlcon_OK2SHIP;
            string app_path = System.Windows.Forms.Application.StartupPath;
            string config_file = Path.Combine(app_path, "Config", "config.txt");
            string[] my_config = myCode.read_config_arr(config_file);
            // string server_name = "";
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

            if (sa_en)
            {
                string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, DB_name, server_acc, server_pass)
                    .ConnectionString;
                _sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);
            }
            else
            {
                string _strcon = TDMK_Code.data_connection2(server_name, DB_name).ConnectionString;
                _sqlcon_OK2SHIP = new SqlConnection(_strcon);
            }

            return _sqlcon_OK2SHIP;
        }

        ComboBox cbStatus = new ComboBox();

        public SqlConnection initial_data(string DB_name, bool sa_en)
        {
            SqlConnection _sqlcon_OK2SHIP;
            if (sa_en)
            {
                string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, DB_name, server_acc, server_pass)
                    .ConnectionString;
                _sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);
            }
            else
            {
                string _strcon = TDMK_Code.data_connection2(server_name, DB_name).ConnectionString;
                _sqlcon_OK2SHIP = new SqlConnection(_strcon);
            }

            return _sqlcon_OK2SHIP;
        }

        public Image get_image_excel(ExcelWorksheet wrk_sheet)
        {
            //myExcel.Shape cur_image = wrk_sheet.Shapes.Item("Picture 1");
            //cur_image.Copy();
            Image myImg = TDMK_EPPLUS.get_pic(wrk_sheet, "Picture 1");
            return myImg;
        }

        public List<Image> get_image_excel_ShearTest(myExcel.Worksheet wrk_sheet, ref List<string> lst_data)
        {
            // Tạo một list tạm để chứa: [Số thứ tự (cột A), Ảnh, Giá trị Data]
            var tempData = new List<Tuple<int, Image, string>>();

            foreach (myExcel.Shape cur_image in wrk_sheet.Shapes)
            {
                int indexNumber = int.MaxValue; // Số thứ tự, mặc định gán số lớn nhất nếu không tìm thấy

                // 1. TÌM SỐ THỨ TỰ BÊN GÓC TRÁI (CỘT A)
                try
                {
                    myExcel.Range topLeft = cur_image.TopLeftCell;
                    int row = topLeft.Row;

                    // Quét từ dòng của ảnh ngược lên trên tối đa 5 dòng để tìm số (phòng trường hợp ảnh bị lệch dòng)
                    for (int i = row; i >= Math.Max(1, row - 5); i--)
                    {
                        var cellValue = ((myExcel.Range)wrk_sheet.Cells[i, 1]).Value2; // Cột 1 = Cột A

                        if (cellValue != null)
                        {
                            // Trường hợp 1: Excel đã nhận diện sẵn ô này là số (double)
                            if (cellValue is double exactNum)
                            {
                                indexNumber = (int)exactNum;
                                break;
                            }
                            // Trường hợp 2: Ô chứa chuỗi (ví dụ "1", "2") thì ta ép kiểu an toàn bằng Convert.ToString
                            else if (double.TryParse(Convert.ToString(cellValue), out double parsedNum))
                            {
                                indexNumber = (int)parsedNum;
                                break;
                            }
                        }
                    }
                }
                catch
                {
                    // Bỏ qua nếu có lỗi đọc cell
                }

                // 2. COPY VÀ XỬ LÝ ẢNH NHƯ CŨ
                try
                {
                    cur_image.Copy();
                }
                catch
                {
                }

                // Mẹo nhỏ: Khi dùng Clipboard trong vòng lặp, nên có độ trễ nhỏ để tránh lỗi không copy kịp
                System.Threading.Thread.Sleep(100);

                Image myImg = Clipboard.GetImage();
                string val = "";

                if (myImg != null)
                {
                    Byte[] data = new Byte[0];
                    ImageConverter imgCon = new ImageConverter();
                    data = (byte[])imgCon.ConvertTo(myImg, typeof(byte[]));
                    try
                    {
                        Bitmap bmp;
                        using (var ms = new MemoryStream(data))
                        {
                            bmp = new Bitmap(ms);
                        }

                        val = Extract_data(bmp);
                        if (myCode.IsNumeric(val))
                        {
                            val = (double.Parse(val) / 1000).ToString();
                        }
                    }
                    catch
                    {
                    }
                }

                // Đưa dữ liệu 1 Shape vào list tạm
                tempData.Add(new Tuple<int, Image, string>(indexNumber, myImg, val));
            }

            // 3. SẮP XẾP LẠI DỮ LIỆU THEO SỐ THỨ TỰ (indexNumber)
            var sortedData = tempData.OrderBy(x => x.Item1).ToList();

            // 4. ĐỔ DỮ LIỆU ĐÃ SẮP XẾP RA ĐỂ TRẢ VỀ
            List<Image> lst_grp = new List<Image>();
            lst_data.Clear(); // Xóa danh sách cũ (nếu có) trước khi add dữ liệu đã sort

            foreach (var item in sortedData)
            {
                lst_grp.Add(item.Item2);
                lst_data.Add(item.Item3); // Dù lỗi hay ko (chuỗi rỗng) vẫn add vào cho khớp index
            }

            return lst_grp;
        }

        public List<Image> get_image_excel_ShearTest_OLD(myExcel.Worksheet wrk_sheet, ref List<string> lst_data)
        {
            List<Image> lst_grp = new List<Image> { };
            foreach (myExcel.Shape cur_image in wrk_sheet.Shapes)
            {
                try
                {
                    cur_image.Copy();
                }
                catch
                {
                }

                Byte[] data = new Byte[0];
                Image myImg = Clipboard.GetImage();
                ImageConverter imgCon = new ImageConverter();
                data = (byte[])imgCon.ConvertTo(myImg, typeof(byte[]));
                try
                {
                    Bitmap bmp;
                    using (var ms = new MemoryStream(data))
                    {
                        bmp = new Bitmap(ms);
                    }

                    string val = Extract_data(bmp);
                    if (myCode.IsNumeric(val))
                    {
                        val = (double.Parse(val) / 1000).ToString();
                    }
                    else
                    {
                        val = "";
                    }

                    lst_data.Add(val);
                }
                catch
                {
                    lst_data.Add("");
                }

                lst_grp.Add(myImg);
            }

            return lst_grp;
        }

        public void Hide_column(DataGridView dgv, ref bool en_)
        {
            string[] col_hide = { "ItemCode", "LotNo", "Sheet", "Operator", "Time_Update", "Remark" };
            if (en_)
            {
                for (int i = 0; i < dgv.Columns.Count; i++)
                {
                    if (col_hide.Contains(dgv.Columns[i].Name))
                    {
                        dgv.Columns[i].Visible = false;
                    }
                }
            }
            else
            {
                for (int i = 0; i < dgv.Columns.Count; i++)
                {
                    if (col_hide.Contains(dgv.Columns[i].Name))
                    {
                        dgv.Columns[i].Visible = true;
                    }
                }
            }

            en_ = !en_;
        }

        public string get_data_val_comment3_onproduct(ExcelWorksheet ws)
        {
            string val = "";
            for (int i = 5; i < 30; i++)
            {
                for (int j = 4; j < 15; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Replace(" ", string.Empty).ToUpper()
                        .Contains("Max per width".Replace(" ", string.Empty).ToUpper()))
                    {
                        string cur_addr = ws.Cells[i, j].Address;
                        string tar_addr = get_offset_addr(cur_addr, ws, 1, true);

                        string val_max = myCode.checkDBNull(ws.Cells[tar_addr].Value);
                        if (!myCode.IsNumeric(val_max))
                        {
                            string xml = @"<root>" + myCode.checkDBNull(ws.Cells[tar_addr].Value) + @"</root>";

                            XmlDocument xmlDoc = new XmlDocument();
                            xmlDoc.LoadXml(xml);
                            XmlNodeList textNodes = xmlDoc.GetElementsByTagName("t");

                            foreach (XmlNode textNode in textNodes)
                            {
                                string textValue = textNode.InnerText.Trim();
                                if (myCode.IsNumeric(textValue))
                                {
                                    val_max = textValue;
                                    break;
                                }
                            }
                        }

                        val += "Max:" + val_max + "_";

                        // val += "Max:" + myCode.checkDBNull(ws.Cells[tar_addr].Value).Replace(" ", "") + "_";
                        for (int k = 1; k < 5; k++)
                        {
                            string sel_addr = get_offset_addr(ws.Cells[cur_addr].Address, ws, k, false);
                            if (myCode.checkDBNull(ws.Cells[sel_addr].Value).Replace(" ", string.Empty).ToUpper()
                                .Contains("Average".Replace(" ", string.Empty).ToUpper()))
                            {
                                string ave_addr = get_offset_addr(ws.Cells[sel_addr].Address, ws, 1, true);
                                val += "Average:" + myCode.checkDBNull(ws.Cells[ave_addr].Value).Replace(" ", "");
                                return val;
                            }
                        }
                    }
                }
            }

            return val;
        }


        public string get_data_val_comment3(ExcelWorksheet ws, string textfind)
        {
            string val = "";

            for (int i = 5; i < 30; i++)
            {
                for (int j = 4; j < 15; j++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Replace(" ", string.Empty).ToUpper()
                        .Contains(textfind.Replace(" ", string.Empty).ToUpper()))
                    {
                        string cur_addr = ws.Cells[i, j].Address;
                        string tar_addr = get_offset_addr(cur_addr, ws, 1, true);
                        val = myCode.checkDBNull(ws.Cells[tar_addr].Value);
                        if (!myCode.IsNumeric(val))
                        {
                            string xml = @"<root>" + myCode.checkDBNull(ws.Cells[tar_addr].Value) + @"</root>";

                            XmlDocument xmlDoc = new XmlDocument();
                            xmlDoc.LoadXml(xml);
                            XmlNodeList textNodes = xmlDoc.GetElementsByTagName("t");

                            foreach (XmlNode textNode in textNodes)
                            {
                                string textValue = textNode.InnerText.Trim();
                                if (myCode.IsNumeric(textValue))
                                {
                                    val = textValue;
                                    break;
                                }
                            }
                        }

                        return val;
                    }
                }
            }

            return val;
        }


        public Image resizeImage(double nPercent, Image src_img)
        {
            int sourceWidth = src_img.Width;
            int sourceHeight = src_img.Height;
            int sourceX = 0, sourceY = 0, destX = 0, destY = 0;
            int destWidth = (int)(sourceWidth * nPercent);
            int destHeight = (int)(sourceHeight * nPercent);
            Bitmap bmPhoto = new Bitmap(destWidth, destHeight, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            bmPhoto.SetResolution(src_img.HorizontalResolution, src_img.VerticalResolution);
            Graphics grPhoto = Graphics.FromImage(bmPhoto);
            grPhoto.Clear(Color.Black);
            grPhoto.InterpolationMode = InterpolationMode.HighQualityBicubic;
            grPhoto.DrawImage(src_img, new System.Drawing.Rectangle(destX, destY, destWidth, destHeight),
                new System.Drawing.Rectangle(sourceX, sourceY, sourceWidth, sourceHeight), GraphicsUnit.Pixel);
            grPhoto.Dispose();
            src_img.Dispose();
            return bmPhoto;
        }

        public void Get_Image_comment3(string in_src, ref SortedDictionary<int, Image> lst_result_sorted)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            SortedDictionary<int, Image> lst_result = new SortedDictionary<int, Image> { };
            FileInfo[] temp_lst = tar_d.GetFiles("*.jpg").Concat(tar_d.GetFiles("*.jpeg")).ToArray();
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    Image sel_img = Image.FromFile(temp_lst[i].FullName);

                    ImageConverter imgcon = new ImageConverter();
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                        .TrimEnd(new char[] { ',', '.', ' ' });

                    char[] ch_arr = f_na.ToCharArray();
                    for (int t = 0; t < ch_arr.Length; t++)
                    {
                        string a = ch_arr[t].ToString();
                        if (!myCode.IsNumeric(a))
                        {
                            f_na = f_na.Replace(a, "");
                        }
                    }

                    if (!lst_result.ContainsKey(Convert.ToInt32(f_na)))
                    {
                        lst_result.Add(Convert.ToInt32(f_na), sel_img);
                    }
                }

                int k = 0;
                foreach (var item in lst_result)
                {
                    lst_result_sorted.Add(k, item.Value);
                    k++;
                }
            }
        }

        public void Get_comment3_logfile_Multi_onproduct(string in_src,
            ref SortedDictionary<string, Funtion_SMT.Peeltest_data> dic_lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
            Get_Image_comment3(in_src, ref dic_image);
            FileInfo[] temp_lst = tar_d.GetFiles("*.xlsx");
            if (temp_lst.Length > 0)
            {
                Dictionary<string, SortedDictionary<int, string>> lst_result =
                    new Dictionary<string, SortedDictionary<int, string>>();
                int i = 0;
                foreach (var img in dic_image.Values)
                {
                    if (i < temp_lst.Length)
                    {
                        //myExcel.Application xlsApp = TDMK_Code.StartExcel();
                        ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                        ExcelWorksheet wrksht = wrkbk.Worksheets[0];

                        Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                        data.data_val = get_data_val_comment3_onproduct(wrksht);
                        //data.data_val = Get_LogFile_Data(temp_lst[i].FullName, "Min"); 
                        data.grap_data = get_image_excel(wrksht);
                        data.image_data = img;
                        string f_inx = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                            .TrimEnd(new char[] { ',', '.', ' ' });
                        if (!dic_lst_result.ContainsKey(f_inx))
                        {
                            dic_lst_result.Add(f_inx, data);
                        }

                        i++;
                    }
                }
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_comment3_logfile_Multi_onproduct(inter_lst.FullName, ref dic_lst_result);
                    }
                }
            }
        }

        public void insert_tbl_old(DirectoryInfo tar_d, ref DataTable tbl_in, string str_parent, string in_src,
            string infor, bool d)
        {
            SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_lst_result =
                new SortedDictionary<int, Funtion_SMT.Peeltest_data>();
            SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
            Get_Image_comment3(tar_d.FullName, ref dic_image);
            FileInfo[] temp_lst = tar_d.GetFiles("*.xlsx");
            if (temp_lst.Length > 0)
            {
                myExcel.Application xlsApp = TDMK_Code.StartExcel();
                xlsApp.DisplayAlerts = false;

                for (int i = 0; i < temp_lst.Length; i++)
                {
                    ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                    ExcelWorksheet wrksht = wrkbk.Worksheets[0];

                    Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                    data.data_val = get_data_val_comment3_onproduct(wrksht);
                    data.grap_data = get_image_excel(wrksht);

                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                        .TrimEnd(new char[] { ',', '.', ' ' });
                    f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                    if (myCode.IsNumeric(f_na))
                    {
                        int f_inx = Convert.ToInt32(f_na);
                        if (!dic_lst_result.ContainsKey(f_inx))
                        {
                            dic_lst_result.Add(f_inx, data);
                        }
                    }

                    TDMK_Code.releaseObject(wrkbk);
                }

                xlsApp.DisplayAlerts = true;
            }

            if (str_parent.ToUpper().Contains("PSA") || str_parent.ToUpper().Contains("LINER") ||
                str_parent.Contains(txtItemCode.Text))
            {
                str_parent = "";
            }

            string region = "_";
            int k = 0;
            int ID = tbl_in.Rows.Count + 1;
            foreach (var data in dic_lst_result)
            {
                if (k < dic_image.Count)
                {
                    Image img = image_null;
                    if (dic_image.ContainsKey(k))
                    {
                        img = dic_image[k];
                    }

                    Image graph = data.Value.grap_data;
                    string val = data.Value.data_val;

                    if (d)
                    {
                        region = str_parent + "_" + tar_d.Name;
                    }

                    if (sheet.Contains("COUPON"))
                    {
                        tbl_in.Rows.Add(ID, txt_itemcode_nvl.Text, txt_lotno_nvl.Text, sheet + infor, region, data.Key,
                            img, graph, val, txtOperator.Text, DateTime.Now.ToString(), in_src, true, true);
                    }
                    else
                    {
                        tbl_in.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor, region, data.Key, img,
                            graph, val, txtOperator.Text, DateTime.Now.ToString(), in_src, true, true);
                    }

                    k++;
                    ID++;
                }
            }

            insert_data_refer(ref tbl_in, ID, region);
        }

        public void insert_tbl(DirectoryInfo tar_d, ref DataTable tbl_in, string str_parent, string in_src,
            string infor, bool d)
        {
            SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_lst_result =
                new SortedDictionary<int, Funtion_SMT.Peeltest_data>();
            SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
            Get_Image_comment3(tar_d.FullName, ref dic_image);
            FileInfo[] temp_lst = tar_d.GetFiles("*.xlsx");
            if (temp_lst.Length > 0)
            {
                myExcel.Application xlsApp = TDMK_Code.StartExcel();
                xlsApp.DisplayAlerts = false;

                for (int i = 0; i < temp_lst.Length; i++)
                {
                    ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                    ExcelWorksheet wrksht = wrkbk.Worksheets[0];

                    Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                    data.data_val = get_data_val_comment3_onproduct(wrksht);
                    data.grap_data = get_image_excel(wrksht);

                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                        .TrimEnd(new char[] { ',', '.', ' ' });
                    f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                    if (myCode.IsNumeric(f_na))
                    {
                        int f_inx = Convert.ToInt32(f_na);
                        if (!dic_lst_result.ContainsKey(f_inx))
                        {
                            dic_lst_result.Add(f_inx, data);
                        }
                    }

                    TDMK_Code.releaseObject(wrkbk);
                }

                xlsApp.DisplayAlerts = true;
            }

            if (str_parent.ToUpper().Contains("PSA") || str_parent.ToUpper().Contains("LINER") ||
                str_parent.Contains(txtItemCode.Text))
            {
                str_parent = "";
            }

            string region = "_";
            int k = 0;
            int ID = tbl_in.Rows.Count + 1;
            foreach (var data in dic_lst_result)
            {
                //if (k < dic_image.Count)
                //{
                Image img = image_null;
                if (dic_image.ContainsKey(k))
                {
                    img = dic_image[k];
                }

                Image graph = data.Value.grap_data;
                string val = data.Value.data_val;

                if (d)
                {
                    region = str_parent + "_" + tar_d.Name;
                }

                if (sheet.Contains("COUPON"))
                {
                    tbl_in.Rows.Add(ID, txt_itemcode_nvl.Text, txt_lotno_nvl.Text, sheet + infor, region, data.Key, img,
                        graph, val, txtOperator.Text, DateTime.Now.ToString(), in_src, true, true);
                }
                else
                {
                    tbl_in.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor, region, data.Key, img, graph,
                        val, txtOperator.Text, DateTime.Now.ToString(), in_src, true, true);
                }

                k++;
                ID++;
                // }
            }

            insert_data_refer(ref tbl_in, ID, region);
        }


        public void insert_tbl_folder_ANH(DirectoryInfo tar_d, ref DataTable tbl_in, string str_parent, string in_src,
            string infor, bool d)
        {
            SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_lst_result =
                new SortedDictionary<int, Funtion_SMT.Peeltest_data>();
            SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
            Get_Image_comment3(tar_d.FullName, ref dic_image);
            FileInfo[] temp_lst = new DirectoryInfo(tar_d.Parent.FullName).GetFiles("*.xlsx");
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                    ExcelWorksheet wrksht = wrkbk.Worksheets[0];

                    Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                    data.data_val = get_data_val_comment3_onproduct(wrksht);
                    data.grap_data = get_image_excel(wrksht);

                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                        .TrimEnd(new char[] { ',', '.', ' ' });
                    f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                    if (myCode.IsNumeric(f_na))
                    {
                        int f_inx = Convert.ToInt32(f_na);
                        if (!dic_lst_result.ContainsKey(f_inx))
                        {
                            dic_lst_result.Add(f_inx, data);
                        }
                    }

                    TDMK_Code.releaseObject(wrkbk);
                }
            }

            if (str_parent.ToUpper().Contains("PSA") || str_parent.ToUpper().Contains("LINER") ||
                str_parent.Contains(txtItemCode.Text))
            {
                str_parent = "";
            }

            string region = "_";
            int k = 0;
            int ID = tbl_in.Rows.Count + 1;
            foreach (var data in dic_lst_result)
            {
                if (k < dic_image.Count)
                {
                    Image img = image_null;
                    if (dic_image.ContainsKey(k))
                    {
                        img = dic_image[k];
                    }

                    Image graph = data.Value.grap_data;
                    string val = data.Value.data_val;

                    if (d)
                    {
                        region = str_parent + "_" + tar_d.Parent.Name;
                    }

                    if (sheet.Contains("COUPON"))
                    {
                        tbl_in.Rows.Add(ID, txt_itemcode_nvl.Text, txt_lotno_nvl.Text, sheet + infor, region, data.Key,
                            img, graph, val, txtOperator.Text, DateTime.Now.ToString(), in_src, true, true);
                    }
                    else
                    {
                        tbl_in.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor, region, data.Key, img,
                            graph, val, txtOperator.Text, DateTime.Now.ToString(), in_src, true, true);
                    }

                    k++;
                    ID++;
                }
            }

            insert_data_refer(ref tbl_in, ID, region);
        }

        public DataTable load_data_logfile_onproduct_old(string in_src)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
            DataTable Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, sheet, filter_str).Clone();
            Data_tbl.Columns.Add("Select_Img", typeof(bool));
            Data_tbl.Columns.Add("Select_Grp", typeof(bool));

            DirectoryInfo tar_parent = new DirectoryInfo(in_src);
            DirectoryInfo[] arr_dic_child = tar_parent.GetDirectories();

            if (arr_dic_child.Length > 0)
            {
                foreach (DirectoryInfo tar_d in arr_dic_child)
                {
                    SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_lst_result =
                        new SortedDictionary<int, Funtion_SMT.Peeltest_data>();
                    SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
                    Get_Image_comment3(tar_d.FullName, ref dic_image);
                    FileInfo[] temp_lst = tar_d.GetFiles("*.xlsx");
                    if (temp_lst.Length > 0)
                    {
                        Dictionary<string, SortedDictionary<int, string>> lst_result =
                            new Dictionary<string, SortedDictionary<int, string>>();
                        for (int i = 0; i < temp_lst.Length; i++)
                        {
                            ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                            ExcelWorksheet wrksht = wrkbk.Worksheets[0];

                            Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                            data.data_val = get_data_val_comment3_onproduct(wrksht);
                            data.grap_data = get_image_excel(wrksht);

                            string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                                .TrimEnd(new char[] { ',', '.', ' ' });
                            f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                            if (myCode.IsNumeric(f_na))
                            {
                                int f_inx = Convert.ToInt32(f_na);
                                if (!dic_lst_result.ContainsKey(f_inx))
                                {
                                    dic_lst_result.Add(f_inx, data);
                                }
                            }
                        }
                    }

                    int k = 0;
                    int ID = 1;
                    foreach (var data in dic_lst_result)
                    {
                        if (k < dic_image.Count)
                        {
                            Image img = dic_image[k];
                            Image graph = data.Value.grap_data;
                            string val = data.Value.data_val;
                            Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet, tar_d.Name, data.Key, img,
                                graph, val, txtOperator.Text, DateTime.Now.ToString(), txtLogfile.Text, true, true);

                            k++;
                            ID++;
                        }
                    }
                }
            }
            else
            {
                SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_lst_result =
                    new SortedDictionary<int, Funtion_SMT.Peeltest_data>();
                SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
                Get_Image_comment3(tar_parent.FullName, ref dic_image);
                FileInfo[] temp_lst = tar_parent.GetFiles("*.xlsx");
                if (temp_lst.Length > 0)
                {
                    Dictionary<string, SortedDictionary<int, string>> lst_result =
                        new Dictionary<string, SortedDictionary<int, string>>();
                    int i = 0;
                    foreach (var img in dic_image.Values)
                    {
                        if (i < temp_lst.Length)
                        {
                            ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                            ExcelWorksheet wrksht = wrkbk.Worksheets[0];

                            Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                            data.data_val = get_data_val_comment3_onproduct(wrksht);
                            data.grap_data = get_image_excel(wrksht);
                            data.image_data = img;

                            string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                                .TrimEnd(new char[] { ',', '.', ' ' });
                            f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                            if (myCode.IsNumeric(f_na))
                            {
                                int f_inx = Convert.ToInt32(f_na);
                                if (!dic_lst_result.ContainsKey(f_inx))
                                {
                                    dic_lst_result.Add(f_inx, data);
                                }
                            }

                            i++;
                        }
                    }
                }

                int k = 0;
                int ID = 1;
                foreach (var data in dic_lst_result)
                {
                    if (k < dic_image.Count)
                    {
                        //  byte[] img = data.Value.image_data;
                        Image img = dic_image[k];
                        Image graph = data.Value.grap_data;
                        string val = data.Value.data_val;
                        Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet, tar_parent.Name, data.Key, img,
                            graph, val, txtOperator.Text, DateTime.Now.ToString(), txtLogfile.Text, true, true);
                        ID++;
                        k++;
                    }
                }
            }

            return Data_tbl;
        }

        public void get_multi_onproduct(DirectoryInfo tar_parent, ref DataTable Data_tbl, string in_src, string infor)
        {
            DirectoryInfo[] arr_dic_child = tar_parent.GetDirectories();

            if (arr_dic_child.Length > 0)
            {
                DirectoryInfo tar_sheet = null;
                foreach (DirectoryInfo tar_d in arr_dic_child)
                {
                    if (sheet.Contains(tar_d.Name.ToUpper()))
                    {
                        tar_sheet = tar_d;
                    }
                }

                if (tar_sheet != null)
                {
                    get_multi_onproduct(tar_sheet, ref Data_tbl, in_src, infor);
                }
                else
                {
                    foreach (DirectoryInfo tar_d2 in arr_dic_child)
                    {
                        if (tar_d2.Name.Replace(" ", "").ToUpper().Contains("ANH"))
                        {
                            bool d = true;
                            if (tar_parent.FullName == in_src)
                            {
                                d = false;
                            }

                            insert_tbl_folder_ANH(tar_d2, ref Data_tbl, tar_parent.Parent.Name, in_src, infor, d);
                            break;
                        }
                        else
                        {
                            if (tar_d2.GetDirectories().Length > 0)
                            {
                                get_multi_onproduct(tar_d2, ref Data_tbl, in_src, infor);
                            }
                            else
                            {
                                insert_tbl(tar_d2, ref Data_tbl, tar_parent.Name, in_src, infor, true);
                            }
                        }
                    }
                }
            }
            else
            {
                insert_tbl(tar_parent, ref Data_tbl, "", in_src, infor, false);
            }
        }

        public DataTable load_data_logfile_onproduct_old(string in_src, string infor)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
            DataTable Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, sheet, filter_str).Clone();
            Data_tbl.Columns.Add("Select_Img", typeof(bool));
            Data_tbl.Columns.Add("Select_Grp", typeof(bool));


            DirectoryInfo tar_parent = new DirectoryInfo(in_src);
            DirectoryInfo[] arr_dic_child = tar_parent.GetDirectories();


            // int ID = 1;
            if (arr_dic_child.Length > 0)
            {
                foreach (DirectoryInfo tar_d2 in arr_dic_child)
                {
                    if (sheet.Contains(tar_d2.Name.ToUpper()))
                    {
                        if (tar_d2.GetDirectories().Length > 0)
                        {
                            foreach (DirectoryInfo tar_d1 in tar_d2.GetDirectories())
                            {
                                if (tar_d1.GetDirectories().Length > 0)
                                {
                                    foreach (DirectoryInfo tar_d in tar_d1.GetDirectories())
                                    {
                                        insert_tbl(tar_d, ref Data_tbl, tar_d1.Name, in_src, infor, true);
                                    }
                                }
                                else
                                {
                                    insert_tbl(tar_d1, ref Data_tbl, tar_d2.Name, in_src, infor, true);
                                }
                            }
                        }
                        else
                        {
                            insert_tbl(tar_d2, ref Data_tbl, "", in_src, infor, false);
                        }

                        break;
                    }
                }
            }
            else
            {
                insert_tbl(tar_parent, ref Data_tbl, "", in_src, infor, false);
            }

            return Data_tbl;
        }

        private DataTable GetConstructorLinerOnProduct()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("ItemCode", typeof(string));
            dt.Columns.Add("LotNo", typeof(string));
            dt.Columns.Add("Sheet", typeof(string));
            dt.Columns.Add("Region", typeof(string));
            dt.Columns.Add("Sample", typeof(string));
            dt.Columns.Add("Image", typeof(Image));
            dt.Columns.Add("Graph", typeof(Image));
            dt.Columns.Add("Data", typeof(string));
            dt.Columns.Add("Operator", typeof(string));
            dt.Columns.Add("Time_Update", typeof(string));
            dt.Columns.Add("Remark", typeof(string));
            return dt;
        }

        public DataTable load_data_logfile_onproduct(string in_src, string infor)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" },
                new string[] { txtItemCode.Text, txtLotNo.Text });
            DataTable Data_tbl = GetConstructorLinerOnProduct();
            Data_tbl.Columns.Add("Select_Img", typeof(bool));
            Data_tbl.Columns.Add("Select_Grp", typeof(bool));

            DirectoryInfo tar_parent = new DirectoryInfo(in_src);
            get_multi_onproduct(tar_parent, ref Data_tbl, in_src, infor);
            return Data_tbl;
        }

        public DataTable load_data_logfile_cross_section_ngang(string in_src)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
            DataTable Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, "CROSS_SECTION", filter_str).Clone();
            Data_tbl.Columns.Add("Select", typeof(bool));

            DirectoryInfo tar_parent = new DirectoryInfo(in_src);
            DirectoryInfo[] arr_dir_child = tar_parent.GetDirectories();
            foreach (DirectoryInfo dir_child in arr_dir_child)
            {
                if (dir_child.Name.ToUpper() == "NGANG")
                {
                    SortedDictionary<int, SortedDictionary<int, string>> logfile_result =
                        new SortedDictionary<int, SortedDictionary<int, string>>();
                    Get_logfile_Multi(dir_child.FullName, ref logfile_result);
                    SortedDictionary<int, Image> Image_result = new SortedDictionary<int, Image>();
                    SortedDictionary<int, Image> Image_result_2 = new SortedDictionary<int, Image>();
                    Get_Image_Multi(dir_child.FullName, ref Image_result);
                    Get_Image_Multi_2(dir_child.FullName, ref Image_result_2);


                    int log_inx = 0;
                    foreach (var log in logfile_result)
                    {
                        Image data_image = null;
                        if (Image_result.ContainsKey(log.Key))
                        {
                            data_image = Image_result[log.Key];
                        }

                        Image data_image_2 = null;
                        if (Image_result_2.ContainsKey(log.Key))
                        {
                            data_image_2 = Image_result_2[log.Key];
                        }

                        Data_tbl.Rows.Add(log_inx + 1, txtItemCode.Text, txtLotNo.Text,
                            sheet.ToUpper().Replace("-", "").Replace(" ", ""), "NGANG", Image_result[log.Key],
                            Image_result_2[log.Key], log.Value[1] + " ; " + log.Value[2], txtOperator.Text,
                            DateTime.Now.ToString(), dir_child.FullName, false);
                        log_inx++;
                    }
                }
            }

            return Data_tbl;
        }

        public DataTable load_data_logfile_gap_connector_old(string in_src)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
            DataTable Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, "GAP_CONNECTOR", filter_str).Clone();
            Data_tbl.Columns.Add("Select", typeof(bool));

            DirectoryInfo tar_parent = new DirectoryInfo(in_src);
            DirectoryInfo[] arr_dir_child = tar_parent.GetDirectories();
            foreach (DirectoryInfo dir_child in arr_dir_child)
            {
                if (dir_child.Name.ToUpper().Contains("GAP"))
                {
                    SortedDictionary<int, SortedDictionary<int, string>> logfile_result1 =
                        new SortedDictionary<int, SortedDictionary<int, string>>();
                    SortedDictionary<int, SortedDictionary<int, string>> logfile_result2 =
                        new SortedDictionary<int, SortedDictionary<int, string>>();

                    Get_logfile_Multi_GAP1(dir_child.FullName, ref logfile_result1);
                    Get_logfile_Multi_GAP2(dir_child.FullName, ref logfile_result2);
                    SortedDictionary<int, Image> Image_result = new SortedDictionary<int, Image>();
                    SortedDictionary<int, Image> Image_result_1 = new SortedDictionary<int, Image>();
                    SortedDictionary<int, Image> Image_result_2 = new SortedDictionary<int, Image>();
                    Get_Image_Multi_gap(dir_child.FullName, ref Image_result);
                    Get_Image_Multi_gap1(dir_child.FullName, ref Image_result_1);
                    Get_Image_Multi_gap2(dir_child.FullName, ref Image_result_2);

                    int log_inx = 0;
                    foreach (var log in logfile_result1)
                    {
                        Image data_image = image_null;
                        if (Image_result.ContainsKey(log.Key))
                        {
                            data_image = Image_result[log.Key];
                        }


                        Image data_image_1 = image_null;
                        if (Image_result_1.ContainsKey(log.Key))
                        {
                            data_image_1 = Image_result_1[log.Key];
                        }

                        Image data_image_2 = image_null;
                        if (Image_result_2.ContainsKey(log.Key))
                        {
                            data_image_2 = Image_result_2[log.Key];
                        }

                        string data = "";
                        foreach (var log_val in log.Value)
                        {
                            data += log_val.Value + " ; ";
                        }

                        if (logfile_result2.ContainsKey(log.Key))
                        {
                            SortedDictionary<int, string> log2 = logfile_result2[log.Key];

                            data += "/";
                            foreach (var log_val in log2)
                            {
                                data += log_val.Value + " ; ";
                            }
                        }

                        if (data_image != image_null && data_image_1 != image_null && data_image_2 != image_null)
                        {
                            Data_tbl.Rows.Add(log_inx + 1, txtItemCode.Text, txtLotNo.Text, sheet, dir_child.Name,
                                log.Key, data_image, data_image_1, data_image_2, data, txtOperator.Text,
                                DateTime.Now.ToString(), dir_child.FullName, true);
                        }


                        log_inx++;
                    }
                }
                //else if (dir_child.Name.ToUpper().Contains("TRU") && dir_child.Name.ToUpper().Contains("GAP"))
                //{
                //    foreach (DirectoryInfo tar_d in dir_child.GetDirectories())
                //    {
                //        SortedDictionary<int, SortedDictionary<int, string>> logfile_result1 = new SortedDictionary<int, SortedDictionary<int, string>>();
                //        SortedDictionary<int, SortedDictionary<int, string>> logfile_result2 = new SortedDictionary<int, SortedDictionary<int, string>>();
                //        Get_logfile_Multi_GAP1(tar_d.FullName, ref logfile_result1);
                //        Get_logfile_Multi_GAP2(tar_d.FullName, ref logfile_result2);
                //        SortedDictionary<int, byte[]> Image_result = new SortedDictionary<int, byte[]>();
                //        SortedDictionary<int, byte[]> Image_result_1 = new SortedDictionary<int, byte[]>();
                //        SortedDictionary<int, byte[]> Image_result_2 = new SortedDictionary<int, byte[]>();
                //        Get_Image_Multi_gap(tar_d.FullName, ref Image_result);
                //        Get_Image_Multi_gap1(tar_d.FullName, ref Image_result_1);
                //        Get_Image_Multi_gap2(tar_d.FullName, ref Image_result_2);

                //        int log_inx = 0;
                //        foreach (var log in logfile_result1)
                //        {
                //            byte[] data_image = img_null;
                //            if (Image_result.ContainsKey(log.Key))
                //            {
                //                data_image = Image_result[log.Key];
                //            }

                //            byte[] data_image_1 = img_null;
                //            if (Image_result_1.ContainsKey(log.Key))
                //            {
                //                data_image_1 = Image_result_1[log.Key];
                //            }

                //            byte[] data_image_2 = img_null;
                //            if (Image_result_2.ContainsKey(log.Key))
                //            {
                //                data_image_2 = Image_result_2[log.Key];
                //            }

                //            string data = "";
                //            foreach (var log_val in log.Value)
                //            {
                //                data += log_val.Value + " ; ";
                //            }

                //            if (logfile_result2.ContainsKey(log.Key))
                //            {
                //                SortedDictionary<int, string> log2 = logfile_result2[log.Key];

                //                data += "/";
                //                foreach (var log_val in log2)
                //                {
                //                    data += log_val.Value + " ; ";
                //                }
                //            }

                //            if (data_image != img_null && data_image_1 != img_null && data_image_2 != img_null)
                //            {
                //                Data_tbl.Rows.Add(log_inx + 1, txtItemCode.Text, txtLotNo.Text, sheet, dir_child.Name + "_" + tar_d.Name, log.Key, data_image, data_image_1, data_image_2, data, txtOperator.Text, DateTime.Now.ToString(), dir_child.FullName, true);
                //            }

                //            log_inx++;


                //        }
                //    }
                //}
            }

            return Data_tbl;
        }

        public void get_data_GAP()
        {
        }

        public DataTable load_data_logfile_gap_connector(string in_src, string infor, string prime)
        {
            if (prime == "Shield b2b")
            {
                return GAPConnectorService.ReadLogfile(in_src, infor);
            }

            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" },
                new string[] { txtItemCode.Text, txtLotNo.Text });
            //THUY: THAY DOI CACH LAY CAU TRUC 
            DataTable Data_tbl = GAPConnectorService.getConstructor();
            Data_tbl.Columns.Add("Select", typeof(bool));
            bool primePID = false;
            Dictionary<string, List<string>> listProductID = new Dictionary<string, List<string>>();
            try
            {
                listProductID =
                    GAPConnectorService.Get_ProductID(textBox1.Text.ToString(), txtItemCode.Text, txtLotNo.Text);
                if (listProductID.Count > 0)
                {
                    Data_tbl.Columns.Add("ProductID");
                    primePID = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Khong co product ID", "Product ID");
            }

            DirectoryInfo tar_parent = new DirectoryInfo(in_src);
            DirectoryInfo[] arr_dir_child = tar_parent.GetDirectories();
            foreach (DirectoryInfo dir_child in arr_dir_child)
            {
                if (dir_child.Name.ToUpper().Contains("GAP"))
                {
                    if (dir_child.GetDirectories().Length > 0)
                        foreach (DirectoryInfo dir_child_2 in dir_child.GetDirectories())
                        {
                            {
                                if (dir_child_2.Name.Replace(" ", "") == "1" ||
                                    dir_child_2.Name.Replace(" ", "") == "2")
                                {
                                    SortedDictionary<int, SortedDictionary<int, string>> logfile_result1 =
                                        new SortedDictionary<int, SortedDictionary<int, string>>();
                                    SortedDictionary<int, SortedDictionary<int, string>> logfile_result2 =
                                        new SortedDictionary<int, SortedDictionary<int, string>>();

                                    Get_logfile_Multi_GAP1(dir_child_2.FullName, ref logfile_result1);
                                    Get_logfile_Multi_GAP2(dir_child_2.FullName, ref logfile_result2);
                                    SortedDictionary<int, Image> Image_result = new SortedDictionary<int, Image>();
                                    SortedDictionary<int, Image> Image_result_1 = new SortedDictionary<int, Image>();
                                    SortedDictionary<int, Image> Image_result_2 = new SortedDictionary<int, Image>();
                                    Get_Image_Multi_gap(dir_child_2.FullName, ref Image_result);
                                    Get_Image_Multi_gap1(dir_child_2.FullName, ref Image_result_1);
                                    Get_Image_Multi_gap2(dir_child_2.FullName, ref Image_result_2);

                                    int log_inx = 0;
                                    foreach (var log in logfile_result1)
                                    {
                                        Image data_image = image_null;
                                        if (Image_result.ContainsKey(log.Key))
                                        {
                                            data_image = Image_result[log.Key];
                                        }

                                        Image data_image_1 = image_null;
                                        if (Image_result_1.ContainsKey(log.Key))
                                        {
                                            data_image_1 = Image_result_1[log.Key];
                                        }

                                        Image data_image_2 = image_null;
                                        if (Image_result_2.ContainsKey(log.Key))
                                        {
                                            data_image_2 = Image_result_2[log.Key];
                                        }

                                        string data = "";
                                        foreach (var log_val in log.Value)
                                        {
                                            data += log_val.Value + " ; ";
                                        }

                                        if (logfile_result2.ContainsKey(log.Key))
                                        {
                                            SortedDictionary<int, string> log2 = logfile_result2[log.Key];
                                            data += "/";
                                            foreach (var log_val in log2)
                                            {
                                                data += log_val.Value + " ; ";
                                            }
                                        }

                                        if (data_image != image_null && data_image_1 != image_null &&
                                            data_image_2 != image_null)
                                        {
                                            Data_tbl.Rows.Add(log_inx + 1, txtItemCode.Text, txtLotNo.Text,
                                                sheet + infor, dir_child.Name + "_" + dir_child_2.Name, log.Key,
                                                data_image, data_image_1, data_image_2, data, txtOperator.Text,
                                                DateTime.Now.ToString(), in_src, true);
                                        }

                                        log_inx++;
                                    }
                                }
                                else
                                {
                                    Debugger.Break();
                                }
                            }
                        }
                    else
                    {
                        SortedDictionary<int, SortedDictionary<int, string>> logfile_result1 =
                            new SortedDictionary<int, SortedDictionary<int, string>>();
                        SortedDictionary<int, SortedDictionary<int, string>> logfile_result2 =
                            new SortedDictionary<int, SortedDictionary<int, string>>();

                        Get_logfile_Multi_GAP1(dir_child.FullName, ref logfile_result1);
                        Get_logfile_Multi_GAP2(dir_child.FullName, ref logfile_result2);
                        SortedDictionary<int, Image> Image_result = new SortedDictionary<int, Image>();
                        SortedDictionary<int, Image> Image_result_1 = new SortedDictionary<int, Image>();
                        SortedDictionary<int, Image> Image_result_2 = new SortedDictionary<int, Image>();
                        Get_Image_Multi_gap(dir_child.FullName, ref Image_result);
                        Get_Image_Multi_gap1(dir_child.FullName, ref Image_result_1);
                        Get_Image_Multi_gap2(dir_child.FullName, ref Image_result_2);
                        int log_inx = 0;

                        foreach (var log in logfile_result1)
                        {
                            Image data_image = image_null;
                            if (Image_result.ContainsKey(log.Key))
                            {
                                data_image = Image_result[log.Key];
                            }


                            Image data_image_1 = image_null;
                            if (Image_result_1.ContainsKey(log.Key))
                            {
                                data_image_1 = Image_result_1[log.Key];
                            }

                            Image data_image_2 = image_null;
                            if (Image_result_2.ContainsKey(log.Key))
                            {
                                data_image_2 = Image_result_2[log.Key];
                            }

                            string data = "";
                            foreach (var log_val in log.Value)
                            {
                                data += log_val.Value + " ; ";
                            }

                            if (logfile_result2.ContainsKey(log.Key))
                            {
                                SortedDictionary<int, string> log2 = logfile_result2[log.Key];

                                data += "/";
                                foreach (var log_val in log2)
                                {
                                    data += log_val.Value + " ; ";
                                }
                            }

                            if (data_image != image_null && data_image_1 != image_null && data_image_2 != image_null)
                            {
                                string GAP = dir_child.Name.ToString().Replace("GAP", "").Trim();
                                if (primePID)
                                {
                                    string productid = "";
                                    if (listProductID[GAP].Count > log_inx)
                                    {
                                        productid = listProductID[GAP][log_inx];
                                    }

                                    Data_tbl.Rows.Add(log_inx + 1, txtItemCode.Text, txtLotNo.Text, sheet + infor,
                                        dir_child.Name, log.Key, data_image, data_image_1, data_image_2, data,
                                        txtOperator.Text, DateTime.Now.ToString(), in_src, true, productid);
                                }
                                else
                                {
                                    Data_tbl.Rows.Add(log_inx + 1, txtItemCode.Text, txtLotNo.Text, sheet + infor,
                                        dir_child.Name, log.Key, data_image, data_image_1, data_image_2, data,
                                        txtOperator.Text, DateTime.Now.ToString(), in_src, true);
                                }
                            }

                            log_inx++;
                        }
                    }
                }
            }

            return Data_tbl;
        }

        public DataTable load_data_logfile_gap_connector_F4(string in_src, string infor)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
            DataTable Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, "GAP_CONNECTOR", filter_str).Clone();
            Data_tbl.Columns.Add("Select", typeof(bool));

            DirectoryInfo tar_parent = new DirectoryInfo(in_src);
            DirectoryInfo[] arr_dir_child = tar_parent.GetDirectories();
            foreach (DirectoryInfo dir_child in arr_dir_child)
            {
                if (dir_child.Name.ToUpper().Contains("GAP"))
                {
                    if (dir_child.GetDirectories().Length > 0)
                    {
                        foreach (DirectoryInfo dir_child_2 in dir_child.GetDirectories())
                        {
                            if (dir_child_2.Name.Replace(" ", "") == "1" || dir_child_2.Name.Replace(" ", "") == "2")
                            {
                                SortedDictionary<int, SortedDictionary<int, string>> logfile_result1 =
                                    new SortedDictionary<int, SortedDictionary<int, string>>();
                                SortedDictionary<int, SortedDictionary<int, string>> logfile_result2 =
                                    new SortedDictionary<int, SortedDictionary<int, string>>();

                                Get_logfile_Multi_GAP1(dir_child_2.FullName, ref logfile_result1);
                                Get_logfile_Multi_GAP2(dir_child_2.FullName, ref logfile_result2);
                                SortedDictionary<int, Image> Image_result = new SortedDictionary<int, Image>();
                                SortedDictionary<int, Image> Image_result_1 = new SortedDictionary<int, Image>();
                                SortedDictionary<int, Image> Image_result_2 = new SortedDictionary<int, Image>();
                                Get_Image_Multi_gap(dir_child_2.FullName, ref Image_result);
                                Get_Image_Multi_gap1(dir_child_2.FullName, ref Image_result_1);
                                Get_Image_Multi_gap2(dir_child_2.FullName, ref Image_result_2);

                                int log_inx = 0;
                                foreach (var log in logfile_result1)
                                {
                                    Image data_image = image_null;
                                    if (Image_result.ContainsKey(log.Key))
                                    {
                                        data_image = Image_result[log.Key];
                                    }


                                    Image data_image_1 = image_null;
                                    if (Image_result_1.ContainsKey(log.Key))
                                    {
                                        data_image_1 = Image_result_1[log.Key];
                                    }

                                    Image data_image_2 = image_null;
                                    if (Image_result_2.ContainsKey(log.Key))
                                    {
                                        data_image_2 = Image_result_2[log.Key];
                                    }

                                    string data = "";
                                    foreach (var log_val in log.Value)
                                    {
                                        data += log_val.Value + " ; ";
                                    }

                                    if (logfile_result2.ContainsKey(log.Key))
                                    {
                                        SortedDictionary<int, string> log2 = logfile_result2[log.Key];

                                        data += "/";
                                        foreach (var log_val in log2)
                                        {
                                            data += log_val.Value + " ; ";
                                        }
                                    }

                                    if (data_image != image_null && data_image_1 != image_null &&
                                        data_image_2 != image_null)
                                    {
                                        Data_tbl.Rows.Add(log_inx + 1, txtItemCode.Text, txtLotNo.Text, sheet + infor,
                                            dir_child.Name + "_" + dir_child_2.Name, log.Key, data_image, data_image_1,
                                            data_image_2, data, txtOperator.Text, DateTime.Now.ToString(), in_src,
                                            true);
                                    }

                                    log_inx++;
                                }
                            }
                        }
                    }
                    else
                    {
                        SortedDictionary<int, SortedDictionary<int, string>> logfile_result1 =
                            new SortedDictionary<int, SortedDictionary<int, string>>();
                        SortedDictionary<int, SortedDictionary<int, string>> logfile_result2 =
                            new SortedDictionary<int, SortedDictionary<int, string>>();

                        Get_logfile_Multi_GAP1(dir_child.FullName, ref logfile_result1);
                        Get_logfile_Multi_GAP2(dir_child.FullName, ref logfile_result2);
                        SortedDictionary<int, Image> Image_result = new SortedDictionary<int, Image>();
                        SortedDictionary<int, Image> Image_result_1 = new SortedDictionary<int, Image>();
                        SortedDictionary<int, Image> Image_result_2 = new SortedDictionary<int, Image>();
                        Get_Image_Multi_gap(dir_child.FullName, ref Image_result);
                        Get_Image_Multi_gap1(dir_child.FullName, ref Image_result_1);
                        Get_Image_Multi_gap2(dir_child.FullName, ref Image_result_2);

                        int log_inx = 0;
                        foreach (var log in logfile_result1)
                        {
                            Image data_image = image_null;
                            if (Image_result.ContainsKey(log.Key))
                            {
                                data_image = Image_result[log.Key];
                            }


                            Image data_image_1 = image_null;
                            if (Image_result_1.ContainsKey(log.Key))
                            {
                                data_image_1 = Image_result_1[log.Key];
                            }

                            Image data_image_2 = image_null;
                            if (Image_result_2.ContainsKey(log.Key))
                            {
                                data_image_2 = Image_result_2[log.Key];
                            }

                            string data = "";
                            foreach (var log_val in log.Value)
                            {
                                data += log_val.Value + " ; ";
                            }

                            if (logfile_result2.ContainsKey(log.Key))
                            {
                                SortedDictionary<int, string> log2 = logfile_result2[log.Key];

                                data += "/";
                                foreach (var log_val in log2)
                                {
                                    data += log_val.Value + " ; ";
                                }
                            }

                            if (data_image != image_null && data_image_1 != image_null && data_image_2 != image_null)
                            {
                                Data_tbl.Rows.Add(log_inx + 1, txtItemCode.Text, txtLotNo.Text, sheet + infor,
                                    dir_child.Name, log.Key, data_image, data_image_1, data_image_2, data,
                                    txtOperator.Text, DateTime.Now.ToString(), in_src, true);
                            }

                            log_inx++;
                        }
                    }
                }
            }

            return Data_tbl;
        }

        /// <summary>
        /// chỉ sử dụng để get dữ liệu từ logfile
        /// </summary>
        /// <param name="in_src"></param>
        /// <param name="infor"></param>
        /// <returns></returns>
        public DataTable load_data_logfile_cross_section(string in_src, string infor, string prime = "None")
        {
            if (prime == "Shield b2b" || prime == "Clip")
            {
                return CrossSectionService.ReadLogfile(in_src, infor, prime == "Shield b2b");
            }


            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" },
                new string[] { txtItemCode.Text, txtLotNo.Text });
            DataTable Data_tbl = new CrossSectionService().getStructorTable();
            Data_tbl.Columns.Add("Select", typeof(bool));
            DataTable sort_dt = Data_tbl.Clone();

            DirectoryInfo tar_parent = new DirectoryInfo(in_src);
            DirectoryInfo[] arr_dir_child = tar_parent.GetDirectories();

            string[] ListFolderName = new string[arr_dir_child.Length];

            for (int i = 0; i < arr_dir_child.Length; i++)
            {
                ListFolderName[i] = arr_dir_child[i].Name.ToUpper().Replace(" ", "");
            }

            int ID = 1;

            if (ListFolderName.Contains("1") && ListFolderName.Contains("2"))
            {
                foreach (DirectoryInfo dir_child_2 in arr_dir_child)
                {
                    if (dir_child_2.Name.Replace(" ", "") == "1" || dir_child_2.Name.Replace(" ", "") == "2")
                    {
                        Data_tbl = sort_dt.Clone();
                        foreach (DirectoryInfo dir_child in dir_child_2.GetDirectories())
                        {
                            if (dir_child.Name.ToUpper().Contains("NGANG") && !dir_child.Name.ToUpper().Contains("TRU"))
                            {
                                SortedDictionary<int, SortedDictionary<int, string>> logfile_result =
                                    new SortedDictionary<int, SortedDictionary<int, string>>();
                                Get_logfile_Multi_ngang(dir_child.FullName, ref logfile_result);
                                SortedDictionary<int, Image> Image_result = new SortedDictionary<int, Image>();
                                SortedDictionary<int, Image> Image_result_2 = new SortedDictionary<int, Image>();
                                Get_Image_Multi(dir_child.FullName, ref Image_result);
                                Get_Image_Multi_2(dir_child.FullName, ref Image_result_2);

                                //int log_inx = 0;

                                foreach (var log in logfile_result)
                                {
                                    Image data_image = new Bitmap(1, 1);
                                    if (Image_result.ContainsKey(log.Key))
                                    {
                                        data_image = Image_result[log.Key];
                                    }

                                    Image data_image_2 = new Bitmap(1, 1);
                                    if (Image_result_2.ContainsKey(log.Key))
                                    {
                                        data_image_2 = Image_result_2[log.Key];
                                    }

                                    string data = "";

                                    foreach (var log_val in log.Value)
                                    {
                                        data += log_val.Value + " ; ";
                                    }

                                    if (data != "")
                                    {
                                        Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor,
                                            dir_child_2.Name.Replace(" ", "").ToUpper() + "_" +
                                            dir_child.Name.Replace(" ", "").ToUpper(), log.Key, data_image,
                                            data_image_2, data, txtOperator.Text, DateTime.Now.ToString(), in_src,
                                            true);
                                        ID++;
                                    }
                                }
                            }
                            else if (dir_child.Name.ToUpper().Replace(" ", "").Replace("_", "").Contains("TRUNGANG"))
                            {
                                foreach (DirectoryInfo tar_d in dir_child.GetDirectories())
                                {
                                    SortedDictionary<int, SortedDictionary<int, string>> logfile_result =
                                        new SortedDictionary<int, SortedDictionary<int, string>>();
                                    Get_logfile_Multi_Tru_ngang(tar_d.FullName, ref logfile_result);
                                    SortedDictionary<int, Image> Image_result = new SortedDictionary<int, Image>();
                                    SortedDictionary<int, Image> Image_result_2 = new SortedDictionary<int, Image>();
                                    Get_Image_Multi_Tru_ngang(tar_d.FullName, ref Image_result);


                                    foreach (var log in logfile_result)
                                    {
                                        Image data_image = new Bitmap(1, 1);
                                        if (Image_result.ContainsKey(log.Key))
                                        {
                                            data_image = Image_result[log.Key];
                                        }

                                        string data = "";

                                        foreach (var log_val in log.Value)
                                        {
                                            data += log_val.Value + " ; ";
                                        }

                                        if (data != "")
                                        {
                                            Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor,
                                                dir_child_2.Name.Replace(" ", "").ToUpper() + "_" +
                                                dir_child.Name.Replace(" ", "").ToUpper() + "_" +
                                                tar_d.Name.Replace(" ", "").ToUpper(), log.Key, data_image, img_null,
                                                data, txtOperator.Text, DateTime.Now.ToString(), in_src, true);
                                            ID++;
                                        }
                                    }
                                }
                            }
                            else if (!dir_child.Name.ToUpper().Contains("GAP") &&
                                     (dir_child.Name.ToUpper().Contains("TRU") ||
                                      dir_child.Name.ToUpper().Contains("DOC")))
                            {
                                foreach (DirectoryInfo tar_d in dir_child.GetDirectories())
                                {
                                    SortedDictionary<int, SortedDictionary<int, string>> logfile_result1 =
                                        new SortedDictionary<int, SortedDictionary<int, string>>();
                                    SortedDictionary<int, SortedDictionary<int, string>> logfile_result2 =
                                        new SortedDictionary<int, SortedDictionary<int, string>>();
                                    Get_logfile_Multi_TRU_DOC1(tar_d.FullName, ref logfile_result1);
                                    Get_logfile_Multi_TRU_DOC2(tar_d.FullName, ref logfile_result2);
                                    SortedDictionary<int, Image> Image_result = new SortedDictionary<int, Image>();
                                    SortedDictionary<int, Image> Image_result_2 = new SortedDictionary<int, Image>();
                                    Get_Image_Multi_Tru_ngang(tar_d.FullName, ref Image_result);
                                    Get_Image_Multi_2(tar_d.FullName, ref Image_result_2);


                                    foreach (var log in logfile_result1)
                                    {
                                        Image data_image = new Bitmap(1, 1);
                                        if (Image_result.ContainsKey(log.Key))
                                        {
                                            data_image = Image_result[log.Key];
                                        }

                                        Image data_image_2 = new Bitmap(1, 1);
                                        if (Image_result_2.ContainsKey(log.Key))
                                        {
                                            data_image_2 = Image_result_2[log.Key];
                                        }

                                        string data = "";
                                        foreach (var log_val in log.Value)
                                        {
                                            data += log_val.Value + " ; ";
                                        }

                                        if (logfile_result2.ContainsKey(log.Key))
                                        {
                                            SortedDictionary<int, string> log2 = logfile_result2[log.Key];

                                            data += "/";
                                            foreach (var log_val in log2)
                                            {
                                                data += log_val.Value + " ; ";
                                            }


                                            Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor,
                                                dir_child_2.Name.Replace(" ", "").ToUpper() + "_" +
                                                dir_child.Name.Replace(" ", "").ToUpper() + "_" +
                                                tar_d.Name.Replace(" ", "").ToUpper(), log.Key, data_image,
                                                data_image_2, data, txtOperator.Text, DateTime.Now.ToString(), in_src,
                                                true);
                                            ID++;
                                        }
                                    }
                                }
                            }
                        }

                        sort_table_crosssection(Data_tbl, ref sort_dt);
                    }
                }
            }
            else
            {
                foreach (DirectoryInfo dir_child in arr_dir_child)
                {
                    if (dir_child.Name.ToUpper().Contains("NGANG") && !dir_child.Name.ToUpper().Contains("TRU"))
                    {
                        SortedDictionary<int, SortedDictionary<int, string>> logfile_result =
                            new SortedDictionary<int, SortedDictionary<int, string>>();
                        Get_logfile_Multi_ngang(dir_child.FullName, ref logfile_result);
                        SortedDictionary<int, Image> Image_result = new SortedDictionary<int, Image>();
                        SortedDictionary<int, Image> Image_result_2 = new SortedDictionary<int, Image>();
                        Get_Image_Multi(dir_child.FullName, ref Image_result);
                        Get_Image_Multi_2(dir_child.FullName, ref Image_result_2);


                        foreach (var log in logfile_result)
                        {
                            Image data_image = new Bitmap(1, 1);
                            if (Image_result.ContainsKey(log.Key))
                            {
                                data_image = Image_result[log.Key];
                            }

                            Image data_image_2 = new Bitmap(1, 1);
                            if (Image_result_2.ContainsKey(log.Key))
                            {
                                data_image_2 = Image_result_2[log.Key];
                            }

                            string data = "";

                            foreach (var log_val in log.Value)
                            {
                                data += log_val.Value + " ; ";
                            }

                            if (data != "")
                            {
                                Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor,
                                    dir_child.Name.Replace(" ", "").ToUpper(), log.Key, data_image, data_image_2, data,
                                    txtOperator.Text, DateTime.Now.ToString(), in_src, true);
                            }

                            ID++;
                        }

                        //}
                    }
                    else if (dir_child.Name.ToUpper().Contains("NGANG") && dir_child.Name.ToUpper().Contains("TRU"))
                    {
                        foreach (DirectoryInfo tar_d in dir_child.GetDirectories())
                        {
                            SortedDictionary<int, SortedDictionary<int, string>> logfile_result =
                                new SortedDictionary<int, SortedDictionary<int, string>>();
                            Get_logfile_Multi_Tru_ngang(tar_d.FullName, ref logfile_result);
                            SortedDictionary<int, Image> Image_result = new SortedDictionary<int, Image>();
                            SortedDictionary<int, Image> Image_result_2 = new SortedDictionary<int, Image>();
                            Get_Image_Multi_Tru_ngang(tar_d.FullName, ref Image_result);

                            int log_inx = 0;
                            foreach (var log in logfile_result)
                            {
                                Image data_image = new Bitmap(1, 1);
                                if (Image_result.ContainsKey(log.Key))
                                {
                                    data_image = Image_result[log.Key];
                                }

                                string data = "";

                                foreach (var log_val in log.Value)
                                {
                                    data += log_val.Value + " ; ";
                                }

                                if (data != "")
                                {
                                    Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor,
                                        dir_child.Name.Replace(" ", "").ToUpper() + "_" +
                                        tar_d.Name.Replace(" ", "").ToUpper(), log.Key, data_image, new Bitmap(1, 1),
                                        data, txtOperator.Text, DateTime.Now.ToString(), in_src, true);
                                    ID++;
                                }

                                log_inx++;
                            }
                        }
                    }
                    else if (!dir_child.Name.ToUpper().Contains("GAP") && (dir_child.Name.ToUpper().Contains("TRU") ||
                                                                           dir_child.Name.ToUpper().Contains("DOC")))
                    {
                        foreach (DirectoryInfo tar_d in dir_child.GetDirectories())
                        {
                            SortedDictionary<int, SortedDictionary<int, string>> logfile_result1 =
                                new SortedDictionary<int, SortedDictionary<int, string>>();
                            SortedDictionary<int, SortedDictionary<int, string>> logfile_result2 =
                                new SortedDictionary<int, SortedDictionary<int, string>>();
                            Get_logfile_Multi_TRU_DOC1(tar_d.FullName, ref logfile_result1);
                            Get_logfile_Multi_TRU_DOC2(tar_d.FullName, ref logfile_result2);
                            SortedDictionary<int, Image> Image_result = new SortedDictionary<int, Image>();
                            SortedDictionary<int, Image> Image_result_2 = new SortedDictionary<int, Image>();
                            Get_Image_Multi_Tru_ngang(tar_d.FullName, ref Image_result);
                            Get_Image_Multi_2(tar_d.FullName, ref Image_result_2);

                            int log_inx = 0;
                            foreach (var log in logfile_result1)
                            {
                                Image data_image = new Bitmap(1, 1);
                                if (Image_result.ContainsKey(log.Key))
                                {
                                    data_image = Image_result[log.Key];
                                }

                                Image data_image_2 = new Bitmap(1, 1);
                                if (Image_result_2.ContainsKey(log.Key))
                                {
                                    data_image_2 = Image_result_2[log.Key];
                                }

                                string data = "";
                                foreach (var log_val in log.Value)
                                {
                                    data += log_val.Value + " ; ";
                                }

                                if (logfile_result2.ContainsKey(log.Key))
                                {
                                    SortedDictionary<int, string> log2 = logfile_result2[log.Key];

                                    data += "/";
                                    foreach (var log_val in log2)
                                    {
                                        data += log_val.Value + " ; ";
                                    }


                                    Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor,
                                        dir_child.Name.Replace(" ", "").ToUpper() + "_" +
                                        tar_d.Name.Replace(" ", "").ToUpper(), log.Key, data_image, data_image_2, data,
                                        txtOperator.Text, DateTime.Now.ToString(), in_src, true);
                                    ID++;
                                }

                                log_inx++;
                            }
                        }
                    }
                }

                sort_table_crosssection(Data_tbl, ref sort_dt);
            }

            // kiểm tra PRODUCT ID
            //Debugger.Break();
            try
            {
                DataTable dtZ = sort_dt;
                string region = "";
                int iz = 1;
                foreach (DataRow row in dtZ.Rows)
                {
                    if (region.Contains(row["Region"].ToString()))
                    {
                    }
                    else
                    {
                        region = row["Region"].ToString();
                        iz = 1;
                    }

                    row["Sample"] = iz++;
                }

                ProductIDService productService = new ProductIDService(txtItemCode.Text, txtLotNo.Text, textBox1.Text,
                    new[] { "Xsection" }, new[] { txtItemCode.Text });
                if (productService._listFile.TryGetValue(txtItemCode.Text, out string location))
                {
                    List<string> list = productService.getListProductID(location);
                    if (list.Count > 0)
                    {
                        sort_dt.Columns.Add("ProductID");
                        int i = 0;
                        Dictionary<string, string> dic = new Dictionary<string, string>();
                        foreach (DataRow row in sort_dt.Rows)
                        {
                            string r = "";
                            if (int.TryParse(row["Sample"].ToString(), out int sam))
                            {
                                switch (row["Region"])
                                {
                                    case "NGANG":
                                    case "TRUNGANG_P":
                                    case "TRUNGANG_T":
                                    case "TRUNGANG_PHAI":
                                    case "TRUNGANG_TRAI":
                                        if (sam <= 20)
                                        {
                                            if (sam <= 10)
                                            {
                                                r = $"NPT<10{sam}";
                                            }

                                            if (sam > 10)
                                            {
                                                r = $"NPT>10{sam}";
                                            }
                                        }

                                        break;
                                    case "TRU_T":
                                    case "TRU_TRAI":
                                    case "DOC_T":
                                    case "DOC_TRAI":
                                        if (sam <= 10)
                                        {
                                            r = $"TRUDOC_T{sam}";
                                        }

                                        break;
                                    case "TRU_P":
                                    case "TRU_PHAI":
                                    case "DOC_P":
                                    case "DOC_PHAI":
                                        if (sam <= 10)
                                        {
                                            r = $"TRUDOC_P{sam}";
                                        }

                                        break;
                                    default:
                                        Debugger.Break();
                                        break;
                                }
                            }

                            if (!string.IsNullOrEmpty(r))
                            {
                                if (dic.TryGetValue(r, out string PID))
                                {
                                }
                                else
                                {
                                    if (i < list.Count)
                                    {
                                        dic.Add(r, list[i]);
                                        PID = list[i++];
                                    }
                                }

                                row["ProductID"] = PID;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi không tìm thấy file productID");
            }

            return sort_dt;
        }

        public void sort_table_crosssection(DataTable Data_tbl, ref DataTable sort_dt)
        {
            List<string> lst_region =
                Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToList();
            List<string> lst_sorted = new List<string> { };

            string[] arr_sort = new string[]
                { "NGANG", "TRU_T", "TRU_P", "TRU1_T", "TRU1_P", "TRU2_T", "TRU2_P", "DOC_T", "DOC_P" };

            foreach (string key in arr_sort)
            {
                foreach (var region in lst_region)
                {
                    if (region.ToUpper().Contains(key) && lst_sorted.IndexOf(region) == -1)
                    {
                        lst_sorted.Add(region);
                    }
                }
            }

            int ID = 1;

            foreach (string region in lst_sorted)
            {
                DataView dv = Data_tbl.AsDataView();
                string str_filter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { region });
                dv.RowFilter = str_filter;

                for (int i = 0; i < dv.Count; i++)
                {
                    DataRow dr = sort_dt.NewRow();
                    dr[0] = ID;
                    for (int k = 1; k < sort_dt.Columns.Count; k++)
                    {
                        dr[k] = dv.ToTable().Rows[i][k];
                    }

                    sort_dt.Rows.Add(dr);
                    ID++;
                }
            }
        }


        public void Get_logfile_Multi_ngang(string in_src,
            ref SortedDictionary<int, SortedDictionary<int, string>> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            try
            {
                FileInfo[] temp_lst = tar_d.GetFiles("*.csv");
                if (temp_lst.Length > 0)
                {
                    for (int i = 0; i < temp_lst.Length; i++)
                    {
                        SortedDictionary<int, string> temp = Get_csv_Data(temp_lst[i].FullName);
                        string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                            .TrimEnd(new char[] { ',', '.', ' ' });

                        //if (f_na.Replace(" ", "").Contains("-1"))
                        //{

                        if (myCode.IsNumeric(f_na.Replace(" ", "").Replace("-1", "")))
                        {
                            int f_inx = Convert.ToInt32(f_na.Replace(" ", "").Replace("-1", ""));
                            if (!lst_result.ContainsKey(f_inx))
                            {
                                lst_result.Add(f_inx, temp);
                            }
                        }
                    }
                    //}
                }
                else
                {
                    DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                    if (inter_type_lst.Length > 0)
                    {
                        foreach (var inter_lst in inter_type_lst)
                        {
                            Get_logfile_Multi_ngang(inter_lst.FullName, ref lst_result);
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("Không tìm được vị trí đường dẫn " + in_src, "Thông báo");
            }
        }

        public void Get_logfile_Multi(string in_src,
            ref SortedDictionary<int, SortedDictionary<int, string>> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            try
            {
                FileInfo[] temp_lst = tar_d.GetFiles("*.csv");
                if (temp_lst.Length > 0)
                {
                    for (int i = 0; i < temp_lst.Length; i++)
                    {
                        SortedDictionary<int, string> temp = Get_csv_Data(temp_lst[i].FullName);
                        string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                            .TrimEnd(new char[] { ',', '.', ' ' });

                        if (f_na.Contains("-") == false && f_na.Contains("+") == false)
                        {
                            if (myCode.IsNumeric(f_na))
                            {
                                int f_inx = Convert.ToInt32(f_na);
                                if (!lst_result.ContainsKey(f_inx))
                                {
                                    lst_result.Add(f_inx, temp);
                                }
                            }
                        }
                    }
                }
                else
                {
                    DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                    if (inter_type_lst.Length > 0)
                    {
                        foreach (var inter_lst in inter_type_lst)
                        {
                            Get_logfile_Multi(inter_lst.FullName, ref lst_result);
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("Không tìm được ví trí đường dẫn  " + in_src, "Thông báo");
            }
        }

        public void Get_logfile_Multi_GAP1(string in_src,
            ref SortedDictionary<int, SortedDictionary<int, string>> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);

            FileInfo[] temp_lst = tar_d.GetFiles("*.csv");
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    SortedDictionary<int, string> temp = Get_csv_Data(temp_lst[i].FullName);
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                        .TrimEnd(new char[] { ',', '.', ' ' });

                    if (f_na.Replace(" ", "").Contains("-1"))
                    {
                        if (myCode.IsNumeric(f_na.Replace(" ", "").Replace("-1", "")))
                        {
                            int f_inx = Convert.ToInt32(f_na.Replace(" ", "").Replace("-1", ""));
                            if (!lst_result.ContainsKey(f_inx))
                            {
                                lst_result.Add(f_inx, temp);
                            }
                        }
                    }
                }
            }
            //else
            //{
            //    DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
            //    if (inter_type_lst.Length > 0)
            //    {
            //        foreach (var inter_lst in inter_type_lst)
            //        {
            //            Get_logfile_Multi_GAP_DOC1(inter_lst.FullName, ref lst_result);
            //        }
            //    }
            //}
        }

        public void Get_logfile_Multi_GAP2(string in_src,
            ref SortedDictionary<int, SortedDictionary<int, string>> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);

            FileInfo[] temp_lst = tar_d.GetFiles("*.csv");
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    SortedDictionary<int, string> temp = Get_csv_Data(temp_lst[i].FullName);
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                        .TrimEnd(new char[] { ',', '.', ' ' });

                    if (f_na.Replace(" ", "").Contains("-2"))
                    {
                        if (myCode.IsNumeric(f_na.Replace(" ", "").Replace("-2", "")))
                        {
                            int f_inx = Convert.ToInt32(f_na.Replace(" ", "").Replace("-2", ""));
                            if (!lst_result.ContainsKey(f_inx))
                            {
                                lst_result.Add(f_inx, temp);
                            }
                        }
                    }
                }
            }
            //else
            //{
            //    DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
            //    if (inter_type_lst.Length > 0)
            //    {
            //        foreach (var inter_lst in inter_type_lst)
            //        {
            //            Get_logfile_Multi_GAP_DOC1(inter_lst.FullName, ref lst_result);
            //        }
            //    }
            //}
        }

        public void Get_logfile_Multi_TRU_DOC1(string in_src,
            ref SortedDictionary<int, SortedDictionary<int, string>> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);

            FileInfo[] temp_lst = tar_d.GetFiles("*.csv");
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    SortedDictionary<int, string> temp = Get_csv_Data(temp_lst[i].FullName);
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                        .TrimEnd(new char[] { ',', '.', ' ' });

                    if (f_na.Contains("-") == false && f_na.Contains("+") == false)
                    {
                        if (myCode.IsNumeric(f_na))
                        {
                            int f_inx = Convert.ToInt32(f_na);
                            if (!lst_result.ContainsKey(f_inx))
                            {
                                lst_result.Add(f_inx, temp);
                            }
                        }
                    }
                }
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_logfile_Multi_TRU_DOC1(inter_lst.FullName, ref lst_result);
                    }
                }
            }
        }

        public void Get_logfile_Multi_TRU_DOC2(string in_src,
            ref SortedDictionary<int, SortedDictionary<int, string>> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            FileInfo[] temp_lst = tar_d.GetFiles("*.csv");
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    SortedDictionary<int, string> temp = Get_csv_Data(temp_lst[i].FullName);
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                        .TrimEnd(new char[] { ',', '.', ' ' });
                    if (f_na.Replace(" ", "").Contains("-1"))
                    {
                        if (myCode.IsNumeric(f_na.Replace(" ", "").Replace("-1", "")))
                        {
                            int f_inx = Convert.ToInt32(f_na.Replace(" ", "").Replace("-1", ""));
                            if (!lst_result.ContainsKey(f_inx))
                            {
                                lst_result.Add(f_inx, temp);
                            }
                        }
                    }
                }
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_logfile_Multi_TRU_DOC2(inter_lst.FullName, ref lst_result);
                    }
                }
            }
        }

        public void Get_logfile_Multi_Tru_ngang(string in_src,
            ref SortedDictionary<int, SortedDictionary<int, string>> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);

            FileInfo[] temp_lst = tar_d.GetFiles("*.csv");
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    SortedDictionary<int, string> temp = Get_csv_Data(temp_lst[i].FullName);
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                        .TrimEnd(new char[] { ',', '.', ' ' });

                    if (f_na.Contains("-") == false && f_na.Contains("+") == false)
                    {
                        if (myCode.IsNumeric(f_na))
                        {
                            int f_inx = Convert.ToInt32(f_na);
                            if (!lst_result.ContainsKey(f_inx))
                            {
                                lst_result.Add(f_inx, temp);
                            }
                        }
                    }
                }
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_logfile_Multi_Tru_ngang(inter_lst.FullName, ref lst_result);
                    }
                }
            }
        }

        public void Get_Image_Multi(string in_src, ref SortedDictionary<int, Image> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);

            FileInfo[] temp_lst = tar_d.GetFiles("*.jpg");
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    var sel_img = Bitmap.FromFile(temp_lst[i].FullName);
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                        .TrimEnd(new char[] { ',', '.', ' ' });
                    if (f_na.Contains("-") == false && f_na.Contains("+") == false)
                    {
                        if (myCode.IsNumeric(f_na))
                        {
                            int f_inx = Convert.ToInt32(f_na);
                            lst_result.Add(f_inx, sel_img);
                        }
                    }
                }
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_Image_Multi(inter_lst.FullName, ref lst_result);
                    }
                }
            }
        }

        public void Get_Image_Multi_gap(string in_src, ref SortedDictionary<int, Image> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);

            FileInfo[] temp_lst = tar_d.GetFiles("*.jpg").Concat(tar_d.GetFiles("*.png")).ToArray();
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    var sel_img = Bitmap.FromFile(temp_lst[i].FullName);
                    ImageConverter imgcon = new ImageConverter();
                    Image img_data = sel_img;
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                        .TrimEnd(new char[] { ',', '.', ' ' });
                    if (f_na.Contains("@"))
                    {
                        if (myCode.IsNumeric(f_na.Replace(" ", "").Replace("@", "")))
                        {
                            int f_inx = Convert.ToInt32(f_na.Replace(" ", "").Replace("@", ""));
                            if (!lst_result.ContainsKey(f_inx))
                            {
                                lst_result.Add(f_inx, img_data);
                            }
                        }
                    }
                    else if (!f_na.Contains("+") && !f_na.Contains("-"))
                    {
                        if (myCode.IsNumeric(f_na.Replace(" ", "")))
                        {
                            int f_inx = Convert.ToInt32(f_na.Replace(" ", ""));
                            if (!lst_result.ContainsKey(f_inx))
                            {
                                lst_result.Add(f_inx, img_data);
                            }
                        }
                    }
                }
            }
            //else
            //{
            //    DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
            //    if (inter_type_lst.Length > 0)
            //    {
            //        foreach (var inter_lst in inter_type_lst)
            //        {
            //            Get_Image_Multi_gap1(inter_lst.FullName, ref lst_result);
            //        }
            //    }
            //}
        }

        public void Get_Image_Multi_gap1(string in_src, ref SortedDictionary<int, Image> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);

            FileInfo[] temp_lst = tar_d.GetFiles("*.jpg").Concat(tar_d.GetFiles("*.png")).ToArray();
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    var sel_img = Bitmap.FromFile(temp_lst[i].FullName);
                    ImageConverter imgcon = new ImageConverter();
                    Image img_data = sel_img;
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                        .TrimEnd(new char[] { ',', '.', ' ' });
                    if (f_na.Replace(" ", "").Contains("-1"))
                    {
                        if (myCode.IsNumeric(f_na.Replace(" ", "").Replace("-1", "")))
                        {
                            int f_inx = Convert.ToInt32(f_na.Replace(" ", "").Replace("-1", ""));
                            if (!lst_result.ContainsKey(f_inx))
                            {
                                lst_result.Add(f_inx, img_data);
                            }
                        }
                    }
                }
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_Image_Multi_gap1(inter_lst.FullName, ref lst_result);
                    }
                }
            }
        }

        public void Get_Image_Multi_gap2(string in_src, ref SortedDictionary<int, Image> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);

            FileInfo[] temp_lst = tar_d.GetFiles("*.jpg").Concat(tar_d.GetFiles("*.png")).ToArray();
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    var sel_img = Bitmap.FromFile(temp_lst[i].FullName);
                    ImageConverter imgcon = new ImageConverter();
                    Image img_data = sel_img;
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                        .TrimEnd(new char[] { ',', '.', ' ' });
                    if (f_na.Replace(" ", "").Contains("-2"))
                    {
                        if (myCode.IsNumeric(f_na.Replace(" ", "").Replace("-2", "")))
                        {
                            int f_inx = Convert.ToInt32(f_na.Replace(" ", "").Replace("-2", ""));
                            if (!lst_result.ContainsKey(f_inx))
                            {
                                lst_result.Add(f_inx, img_data);
                            }
                        }
                    }
                }
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_Image_Multi_gap2(inter_lst.FullName, ref lst_result);
                    }
                }
            }
        }

        public void Get_Image_Multi_Tru_ngang(string in_src, ref SortedDictionary<int, Image> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);

            FileInfo[] temp_lst = tar_d.GetFiles("*.jpg");
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    var sel_img = Bitmap.FromFile(temp_lst[i].FullName);
                    Image img_data = sel_img;
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                        .TrimEnd(new char[] { ',', '.', ' ' });
                    //if (f_na.Contains("-") == false && f_na.Contains("+") == false)
                    //{

                    //  f_na = f_na.Replace("-1", "").Replace("+", "");
                    if (myCode.IsNumeric(f_na))
                    {
                        int f_inx = Convert.ToInt32(f_na);
                        if (!lst_result.ContainsKey(f_inx))
                        {
                            lst_result.Add(f_inx, img_data);
                        }
                    }
                    //}
                }
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_Image_Multi_Tru_ngang(inter_lst.FullName, ref lst_result);
                    }
                }
            }
        }

        public void Get_Image_Multi_2(string in_src, ref SortedDictionary<int, Image> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            FileInfo[] temp_lst = tar_d.GetFiles("*.jpg");
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    var sel_img = Bitmap.FromFile(temp_lst[i].FullName);
                    ImageConverter imgcon = new ImageConverter();
                    Image img_data = sel_img;
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                        .TrimEnd(new char[] { ',', '.', ' ' });
                    if (f_na.Replace(" ", "").Contains("-1"))
                    {
                        f_na = f_na.Replace(" ", "").Replace("-1", string.Empty);
                        if (myCode.IsNumeric(f_na))
                        {
                            int f_inx = Convert.ToInt32(f_na);
                            if (!lst_result.ContainsKey(f_inx))
                            {
                                lst_result.Add(f_inx, img_data);
                            }
                        }
                    }
                }
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_Image_Multi_2(inter_lst.FullName, ref lst_result);
                    }
                }
            }
        }
        //public Dictionary<int, string> Get_LogFile_Data(string in_src_file)
        //{
        //    Dictionary<int, string> result_lst = new Dictionary<int, string>();
        //    string[] Lines = File.ReadAllLines(in_src_file);
        //    foreach (var line in Lines)
        //    {
        //        string str = new string(line.Where(s => s != '"').ToArray());
        //        string[] temp = str.Split(',');
        //        if (temp[0] != "")
        //        {
        //            if (myCode.IsNumeric(temp[0]))
        //            {
        //                result_lst.Add(Convert.ToInt32(temp[0]), temp[2]);
        //            }
        //        }
        //        else
        //        {
        //            break;
        //        }
        //    }
        //    return result_lst;
        //}

        public SortedDictionary<int, string> Get_LogFile_data(string in_src_file)
        {
            SortedDictionary<int, string> result_lst = new SortedDictionary<int, string>();
            string[] Lines = System.IO.File.ReadAllLines(in_src_file);
            foreach (var line in Lines)
            {
                string str = new string(line.Where(s => s != '"').ToArray());
                string[] temp = str.Split(',');
                if (temp[0] != "")
                {
                    if (myCode.IsNumeric(temp[0]))
                    {
                        result_lst.Add(Convert.ToInt32(temp[0]), temp[2]);
                    }
                }
                else
                {
                    break;
                }
            }

            return result_lst;
        }
        //public void Get_comment3_logfile_Multi_onproduct_new(string in_src, ref SortedDictionary<string, Funtion_SMT.Peeltest_data> dic_lst_result)
        //{
        //    DirectoryInfo tar_parent = new DirectoryInfo(in_src);
        //    DirectoryInfo[] arr_dic_child = tar_parent.GetDirectories();

        //    foreach (DirectoryInfo tar_d in arr_dic_child)
        //    {
        //        //Dictionary<string, Funtion_SMT.Peeltest_data> dic_region = new Dictionary<string, Funtion_SMT.Peeltest_data> { };  
        //        SortedDictionary<string, byte[]> dic_image = new SortedDictionary<string, byte[]>();
        //        Get_Image_comment3(in_src, ref dic_image);
        //        FileInfo[] temp_lst = tar_d.GetFiles("*.xlsx");
        //        if (temp_lst.Length > 0)
        //        {
        //            Dictionary<string, SortedDictionary<int, string>> lst_result = new Dictionary<string, SortedDictionary<int, string>>();
        //            int i = 0;
        //            foreach (var img in dic_image.Values)
        //            {
        //                if (i < temp_lst.Length)
        //                {
        //                    myExcel.Application xlsApp = TDMK_Code.StartExcel();
        //                    myExcel.Workbook wrkbk = xlsApp.Workbooks.Open(temp_lst[i].FullName);
        //                    myExcel.Worksheet wrksht = wrkbk.Sheets[1];

        //                    Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
        //                    data.data_val = get_data_val_comment3_onproduct(wrksht);
        //                    //data.data_val = Get_LogFile_Data(temp_lst[i].FullName, "Min"); 
        //                    data.grap_data = get_image_excel(wrksht);
        //                    data.image_data = img;
        //                    string f_inx = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
        //                    if (!dic_lst_result.ContainsKey(tar_d.Name + "_" + f_inx))
        //                    {
        //                        dic_lst_result.Add(tar_d.Name + "_" + f_inx, data);
        //                    }

        //                    i++;
        //                    wrkbk.Close();
        //                }
        //            }
        //        }
        //        //else
        //        //{
        //        //    DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
        //        //    if (inter_type_lst.Length > 0)
        //        //    {
        //        //        foreach (var inter_lst in inter_type_lst)
        //        //        {
        //        //            Get_comment3_logfile_Multi_onproduct(inter_lst.FullName, ref dic_region);
        //        //        }
        //        //    }
        //        //}
        //        //dic_lst_result.Add(tar_d.Name, dic_region);
        //    }
        //}

        public DataTable load_data_logfile_Peel_Pull_old(string in_src, string tb_name)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
            DataTable Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, tb_name, filter_str).Clone();
            Data_tbl.Columns.Add("Select", typeof(bool));

            DirectoryInfo tar_parent = new DirectoryInfo(in_src);
            DirectoryInfo[] arr_dic_child = tar_parent.GetDirectories();
            int ID = 1;
            if (arr_dic_child.Length > 0)
            {
                foreach (DirectoryInfo tar_d in arr_dic_child)
                {
                    SortedDictionary<string, Funtion_SMT.Peeltest_data> dic_lst_result =
                        new SortedDictionary<string, Funtion_SMT.Peeltest_data>();
                    SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
                    Get_Image_comment3(tar_d.FullName, ref dic_image);
                    FileInfo[] temp_lst = tar_d.GetFiles("*.xlsx");
                    if (temp_lst.Length > 0)
                    {
                        Dictionary<string, SortedDictionary<int, string>> lst_result =
                            new Dictionary<string, SortedDictionary<int, string>>();
                        int i = 0;
                        foreach (var img in dic_image.Values)
                        {
                            if (i < temp_lst.Length)
                            {
                                //myExcel.Application xlsApp = TDMK_Code.StartExcel();
                                ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                                ExcelWorksheet wrksht = wrkbk.Worksheets[0];

                                Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                                data.data_val = get_data_val_comment3(wrksht, "Max");
                                //data.data_val = Get_LogFile_Data(temp_lst[i].FullName, "Min"); 
                                data.grap_data = get_image_excel(wrksht);
                                data.image_data = img;
                                string f_inx = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                                    .TrimEnd(new char[] { ',', '.', ' ' });
                                if (!dic_lst_result.ContainsKey(tar_d.Name + "_" + f_inx))
                                {
                                    dic_lst_result.Add(tar_d.Name + "_" + f_inx, data);
                                }

                                i++;
                            }
                        }
                    }

                    foreach (var data in dic_lst_result)
                    {
                        Image img = data.Value.image_data;
                        Image graph = data.Value.grap_data;
                        string val = data.Value.data_val;
                        Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text,
                            sheet.ToUpper().Replace("-", "").Replace(" ", ""), data.Key.Split('_')[0],
                            data.Key.Split('_')[1], img, graph, val, "100%", "0.00%", "0.00%", "0.00%", "0.00%",
                            "0.00%", "0.00%", txtOperator.Text, DateTime.Now.ToString(), txtLogfile.Text, true);
                        ID++;
                    }
                }
            }
            else
            {
                SortedDictionary<string, Funtion_SMT.Peeltest_data> dic_lst_result =
                    new SortedDictionary<string, Funtion_SMT.Peeltest_data>();
                SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
                Get_Image_comment3(tar_parent.FullName, ref dic_image);
                FileInfo[] temp_lst = tar_parent.GetFiles("*.xlsx");
                if (temp_lst.Length > 0)
                {
                    Dictionary<string, SortedDictionary<int, string>> lst_result =
                        new Dictionary<string, SortedDictionary<int, string>>();
                    int i = 0;
                    foreach (var img in dic_image.Values)
                    {
                        if (i < temp_lst.Length)
                        {
                            //myExcel.Application xlsApp = TDMK_Code.StartExcel();
                            ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                            ExcelWorksheet wrksht = wrkbk.Worksheets[0];

                            Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                            data.data_val = get_data_val_comment3(wrksht, "Max");
                            //data.data_val = Get_LogFile_Data(temp_lst[i].FullName, "Min"); 
                            data.grap_data = get_image_excel(wrksht);
                            data.image_data = img;
                            string f_inx = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                                .TrimEnd(new char[] { ',', '.', ' ' });
                            if (!dic_lst_result.ContainsKey(f_inx))
                            {
                                dic_lst_result.Add(f_inx, data);
                            }

                            i++;
                        }
                    }
                }

                foreach (var data in dic_lst_result)
                {
                    Image img = data.Value.image_data;
                    Image graph = data.Value.grap_data;
                    string val = data.Value.data_val;
                    Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet, 1, data.Key, img, graph, val, "100%",
                        "0.00%", "0.00%", "0.00%", "0.00%", "0.00%", "0.00%", txtOperator.Text, DateTime.Now.ToString(),
                        txtLogfile.Text, true);
                    ID++;
                }
            }

            return Data_tbl;
        }

        public void insert_data_refer_old(ref DataTable Data_tbl, int ID, string region)
        {
            if (myCode.IsNumeric(txt_pcs_begin.Text) && myCode.IsNumeric(txt_pcs_end.Text) &&
                txt_itemcode_nvl.Text != "" && txt_lotno_nvl.Text != "")
            {
                DataTable dt_ok2build_saved = TDMK_Code.Datatable_Filter(sqlcon, sheet,
                    TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" },
                        new string[] { txt_itemcode_nvl.Text, txt_lotno_nvl.Text, region }));

                DataTable dt_ok2build = dt_ok2build_saved.Clone();
                dt_ok2build.Columns.Add("Select_Img", typeof(bool));
                dt_ok2build.Columns.Add("Select_Grp", typeof(bool));
                foreach (DataRow dr in dt_ok2build_saved.Rows)
                {
                    DataRow dr_new = dt_ok2build.NewRow();
                    for (int i = 0; i < dt_ok2build_saved.Columns.Count; i++)
                    {
                        dr_new[i] = dr[i];
                    }

                    dr_new["Select_Img"] = true;
                    dr_new["Select_Grp"] = true;
                    dt_ok2build.Rows.Add(dr_new);
                }

                int pcs_begin = int.Parse(txt_pcs_begin.Text);
                int pcs_after = int.Parse(txt_pcs_end.Text);

                for (int i = pcs_begin - 1; i < pcs_after; i++)
                {
                    if (i < dt_ok2build.Rows.Count)
                    {
                        DataRow dr = Data_tbl.NewRow();
                        dr[0] = ID;
                        dr[1] = txtItemCode.Text;
                        dr[2] = txtLotNo.Text;
                        for (int c = 3; c < Data_tbl.Columns.Count; c++)
                        {
                            dr[c] = dt_ok2build.Rows[i][c];
                        }

                        Data_tbl.Rows.Add(dr);
                        ID++;
                    }
                }
            }
        }

        public void insert_data_refer(ref DataTable Data_tbl, int ID, string region)
        {
            if (!myCode.IsNumeric(txt_pcs_begin.Text) || !myCode.IsNumeric(txt_pcs_end.Text) ||
                !(txt_itemcode_nvl.Text != "") || !(txt_lotno_nvl.Text != ""))
                return;
            DataTable dataTable1 = TDMK_Code.Datatable_Filter(sqlcon, sheet,
                TDMK_Code.filter_str(new string[4] { "ItemCode", "LotNo", "Sheet", "Region" },
                    new string[4]
                    {
                        txt_itemcode_nvl.Text, txt_lotno_nvl.Text,
                        sheet + ("/" + txt_ItemName.Text + "_" + txt_line.Text + "_" + txt_ca.Text + "_" +
                                 txt_date.Text + "_" + txt_worker.Text + "_" + cb_Type.SelectedItem.ToString()),
                        region
                    }));
            DataTable dataTable2 = dataTable1.Clone();
            dataTable2.Columns.Add("Select_Img", typeof(bool));
            dataTable2.Columns.Add("Select_Grp", typeof(bool));
            foreach (DataRow row1 in (InternalDataCollectionBase)dataTable1.Rows)
            {
                DataRow row2 = dataTable2.NewRow();
                for (int columnIndex = 0; columnIndex < dataTable1.Columns.Count; ++columnIndex)
                    row2[columnIndex] = row1[columnIndex];
                row2["Select_Img"] = (object)true;
                row2["Select_Grp"] = (object)true;
                dataTable2.Rows.Add(row2);
            }

            int num1 = int.Parse(txt_pcs_begin.Text);
            int num2 = int.Parse(txt_pcs_end.Text);
            for (int index = num1 - 1; index < num2; ++index)
            {
                if (index < dataTable2.Rows.Count)
                {
                    DataRow row = Data_tbl.NewRow();
                    row[0] = (object)ID;
                    row[1] = (object)txtItemCode.Text;
                    row[2] = (object)txtLotNo.Text;
                    for (int columnIndex = 3; columnIndex < Data_tbl.Columns.Count; ++columnIndex)
                        row[columnIndex] = dataTable2.Rows[index][columnIndex];
                    Data_tbl.Rows.Add(row);
                    ++ID;
                }
            }
        }

        //Lưu ảnh vào folder "ANH"
        public void load_peelpull(DirectoryInfo d, string infor, string in_src, ref DataTable Data_tbl)
        {
            DirectoryInfo[] arr_dic_child = d.GetDirectories();
            int ID = 1;
            if (arr_dic_child.Length > 0)
            {
                foreach (DirectoryInfo tar_d in arr_dic_child)
                {
                    if (tar_d.Name.Replace(" ", "").ToUpper().Contains("ANH"))
                    {
                        SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_lst_result =
                            new SortedDictionary<int, Funtion_SMT.Peeltest_data>();
                        SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
                        Get_Image_comment3(tar_d.FullName, ref dic_image);
                        FileInfo[] temp_lst = d.GetFiles("*.xlsx");

                        if (temp_lst.Length > 0)
                        {
                            Dictionary<string, SortedDictionary<int, string>> lst_result =
                                new Dictionary<string, SortedDictionary<int, string>>();
                            for (int i = 0; i < temp_lst.Length; i++)
                            {
                                ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                                ExcelWorksheet wrksht = wrkbk.Worksheets[0];
                                Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                                data.data_val = get_data_val_comment3(wrksht, "Max");
                                data.grap_data = get_image_excel(wrksht);
                                string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                                    .TrimEnd(new char[] { ',', '.', ' ' });
                                f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                                if (myCode.IsNumeric(f_na))
                                {
                                    int f_inx = Convert.ToInt32(f_na);
                                    if (!dic_lst_result.ContainsKey(f_inx))
                                    {
                                        dic_lst_result.Add(f_inx, data);
                                    }
                                }
                            }
                        }

                        int k = 0;
                        foreach (var data in dic_lst_result)
                        {
                            if (k < dic_image.Count)
                            {
                                Image img = dic_image[k];
                                Image graph = data.Value.grap_data;
                                string val = data.Value.data_val;
                                Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor, "1", data.Key,
                                    img, graph, val, "0.00%", "0.00%", "0.00%", "0.00%", "100%", "0.00%", "0.00%",
                                    txtOperator.Text, DateTime.Now.ToString(), in_src, true, true);
                                ID++;
                                k++;
                            }
                        }

                        insert_data_refer(ref Data_tbl, ID, "1");
                    }
                    else
                    {
                        SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_lst_result =
                            new SortedDictionary<int, Funtion_SMT.Peeltest_data>();
                        SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();

                        DirectoryInfo[] arr_dic_child_2 = tar_d.GetDirectories();
                        if (arr_dic_child_2.Length > 0)
                        {
                            foreach (DirectoryInfo dic_child_2 in arr_dic_child_2)
                            {
                                if (dic_child_2.Name.Replace(" ", "").ToUpper().Contains("ANH"))
                                {
                                    Get_Image_comment3(dic_child_2.FullName, ref dic_image);
                                    break;
                                }
                            }
                        }
                        else
                        {
                            Get_Image_comment3(tar_d.FullName, ref dic_image);
                        }

                        FileInfo[] temp_lst = tar_d.GetFiles("*.xlsx");

                        if (temp_lst.Length > 0)
                        {
                            Dictionary<string, SortedDictionary<int, string>> lst_result =
                                new Dictionary<string, SortedDictionary<int, string>>();
                            for (int i = 0; i < temp_lst.Length; i++)
                            {
                                ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                                //ExcelWorksheet wrksht = wrkbk.Worksheets[1];
                                ExcelWorksheet wrksht = wrkbk.Worksheets[0];
                                Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                                data.data_val = get_data_val_comment3(wrksht, "Max");
                                data.grap_data = get_image_excel(wrksht);
                                string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                                    .TrimEnd(new char[] { ',', '.', ' ' });
                                f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                                if (myCode.IsNumeric(f_na))
                                {
                                    int f_inx = Convert.ToInt32(f_na);
                                    if (!dic_lst_result.ContainsKey(f_inx))
                                    {
                                        dic_lst_result.Add(f_inx, data);
                                    }
                                }
                            }
                        }

                        int k = 0;
                        foreach (var data in dic_lst_result)
                        {
                            if (k < dic_image.Count)
                            {
                                Image img = dic_image[k];
                                Image graph = data.Value.grap_data;
                                string val = data.Value.data_val;
                                Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor, tar_d.Name,
                                    data.Key, img, graph, val, "0.00%", "0.00%", "0.00%", "0.00%", "100%", "0.00%",
                                    "0.00%", txtOperator.Text, DateTime.Now.ToString(), in_src, true, true);
                                ID++;
                                k++;
                            }
                        }

                        insert_data_refer(ref Data_tbl, ID, tar_d.Name);
                    }
                }
            }
            else
            {
                SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_lst_result =
                    new SortedDictionary<int, Funtion_SMT.Peeltest_data>();
                SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
                Get_Image_comment3(d.FullName, ref dic_image);
                FileInfo[] temp_lst = d.GetFiles("*.xlsx");

                if (temp_lst.Length > 0)
                {
                    Dictionary<string, SortedDictionary<int, string>> lst_result =
                        new Dictionary<string, SortedDictionary<int, string>>();
                    int i = 0;
                    foreach (var img in dic_image.Values)
                    {
                        if (i < temp_lst.Length)
                        {
                            ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                            //ExcelWorksheet wrksht = wrkbk.Worksheets[1];
                            ExcelWorksheet wrksht = wrkbk.Worksheets[0];
                            Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                            data.data_val = get_data_val_comment3(wrksht, "Max");
                            data.grap_data = get_image_excel(wrksht);
                            string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                                .TrimEnd(new char[] { ',', '.', ' ' });

                            f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                            if (myCode.IsNumeric(f_na))
                            {
                                int f_inx = Convert.ToInt32(f_na);
                                if (!dic_lst_result.ContainsKey(f_inx))
                                {
                                    dic_lst_result.Add(f_inx, data);
                                }
                            }

                            i++;
                        }
                    }
                }

                int k = 0;
                foreach (var data in dic_lst_result)
                {
                    if (k < dic_image.Count)
                    {
                        Image img = dic_image[k];
                        Image graph = data.Value.grap_data;
                        string val = data.Value.data_val;
                        Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor, 1, data.Key, img, graph,
                            val, "0.00%", "0.00%", "0.00%", "0.00%", "100%", "0.00%", "0.00%", txtOperator.Text,
                            DateTime.Now.ToString(), in_src, true, true);
                        ID++;
                        k++;
                    }
                }

                insert_data_refer(ref Data_tbl, ID, "1");
            }
        }


        public void load_peelpull_old(DirectoryInfo d, string infor, string in_src, ref DataTable Data_tbl)
        {
            DirectoryInfo[] arr_dic_child = d.GetDirectories();
            int ID = 1;
            if (arr_dic_child.Length > 0)
            {
                foreach (DirectoryInfo tar_d in arr_dic_child)
                {
                    SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_lst_result =
                        new SortedDictionary<int, Funtion_SMT.Peeltest_data>();
                    SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
                    Get_Image_comment3(tar_d.FullName, ref dic_image);
                    FileInfo[] temp_lst = tar_d.GetFiles("*.xlsx");

                    if (temp_lst.Length > 0)
                    {
                        Dictionary<string, SortedDictionary<int, string>> lst_result =
                            new Dictionary<string, SortedDictionary<int, string>>();
                        for (int i = 0; i < temp_lst.Length; i++)
                        {
                            ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                            //ExcelWorksheet wrksht = wrkbk.Worksheets[1];
                            ExcelWorksheet wrksht = wrkbk.Worksheets[0];
                            Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                            data.data_val = get_data_val_comment3(wrksht, "Max");
                            data.grap_data = get_image_excel(wrksht);
                            string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                                .TrimEnd(new char[] { ',', '.', ' ' });
                            f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                            if (myCode.IsNumeric(f_na))
                            {
                                int f_inx = Convert.ToInt32(f_na);
                                if (!dic_lst_result.ContainsKey(f_inx))
                                {
                                    dic_lst_result.Add(f_inx, data);
                                }
                            }
                        }
                    }

                    int k = 0;
                    foreach (var data in dic_lst_result)
                    {
                        if (k < dic_image.Count)
                        {
                            Image img = dic_image[k];
                            Image graph = data.Value.grap_data;
                            string val = data.Value.data_val;
                            Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor, tar_d.Name, data.Key,
                                img, graph, val, "100%", "0.00%", "0.00%", "0.00%", "0.00%", "0.00%", "0.00%",
                                txtOperator.Text, DateTime.Now.ToString(), in_src, true, true);
                            ID++;
                            k++;
                        }
                    }

                    insert_data_refer(ref Data_tbl, ID, tar_d.Name);
                }
            }
            else
            {
                SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_lst_result =
                    new SortedDictionary<int, Funtion_SMT.Peeltest_data>();
                SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
                Get_Image_comment3(d.FullName, ref dic_image);
                FileInfo[] temp_lst = d.GetFiles("*.xlsx");


                if (temp_lst.Length > 0)
                {
                    Dictionary<string, SortedDictionary<int, string>> lst_result =
                        new Dictionary<string, SortedDictionary<int, string>>();
                    int i = 0;
                    foreach (var img in dic_image.Values)
                    {
                        if (i < temp_lst.Length)
                        {
                            ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                            //ExcelWorksheet wrksht = wrkbk.Worksheets[1];
                            ExcelWorksheet wrksht = wrkbk.Worksheets[0];
                            Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                            data.data_val = get_data_val_comment3(wrksht, "Max");
                            data.grap_data = get_image_excel(wrksht);
                            string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                                .TrimEnd(new char[] { ',', '.', ' ' });

                            f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                            if (myCode.IsNumeric(f_na))
                            {
                                int f_inx = Convert.ToInt32(f_na);
                                if (!dic_lst_result.ContainsKey(f_inx))
                                {
                                    dic_lst_result.Add(f_inx, data);
                                }
                            }

                            i++;
                        }
                    }
                }

                int k = 0;
                foreach (var data in dic_lst_result)
                {
                    if (k < dic_image.Count)
                    {
                        Image img = dic_image[k];
                        Image graph = data.Value.grap_data;
                        string val = data.Value.data_val;
                        Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor, 1, data.Key, img, graph,
                            val, "100%", "0.00%", "0.00%", "0.00%", "0.00%", "0.00%", "0.00%", txtOperator.Text,
                            DateTime.Now.ToString(), in_src, true, true);
                        ID++;
                        k++;
                    }
                }

                insert_data_refer(ref Data_tbl, ID, "1");
            }
        }

        public DataTable load_data_logfile_Peel_Pull(string in_src, string tb_name, string infor,
            string locationProductID = null)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" },
                new string[] { txtItemCode.Text, txtLotNo.Text });
            DataTable Data_tbl = PeelTestWOSUSService.getStructor();
            Data_tbl.Columns.Add("Select_Img", typeof(bool));
            Data_tbl.Columns.Add("Select_Grp", typeof(bool));
            DirectoryInfo tar_parent = new DirectoryInfo(in_src);

            DirectoryInfo[] arr_dic_parent = tar_parent.GetDirectories();
            bool b_ = false;
            foreach (DirectoryInfo tar_d in arr_dic_parent)
            {
                if (tar_d.Name.ToUpper().Contains(tb_name.Replace("_TEST", "").Replace("MATING_", "")))
                {
                    b_ = true;
                    break;
                }
            }

            if (arr_dic_parent.Length > 0)
            {
                if (b_)
                {
                    foreach (DirectoryInfo d in arr_dic_parent)
                    {
                        string f_sheet = tb_name.Replace("_TEST", "").Replace("MATING_", "");
                        if (d.Name.ToUpper().Contains(f_sheet))
                        {
                            load_peelpull(d, infor, in_src, ref Data_tbl);
                            break;
                        }
                    }
                }
                else
                {
                    load_peelpull(tar_parent, infor, in_src, ref Data_tbl);
                }
            }
            else
            {
                load_peelpull(tar_parent, infor, in_src, ref Data_tbl);
            }

            if (!string.IsNullOrEmpty(locationProductID))
            {
                string find = "";
                switch (sheet)
                {
                    case "MATING_PULL_TEST":
                        find = "Pulling";
                        break;
                    case "PEEL_TEST":
                        find = "Peeling";
                        break;
                    case "PEEL_TEST_WITHOUT_SUS":
                        find = "without SUS";
                        break;
                }

                try
                {
                    ProductIDService productIDService = new ProductIDService(txtItemCode.Text, txtLotNo.Text,
                        locationProductID, new[] { find }, new[] { txtItemCode.Text.Trim() });
                    List<string> list = productIDService.getListProductID(productIDService._listFile[txtItemCode.Text]);
                    if (list.Count > 0)
                    {
                        Data_tbl.Columns.Add("ProductID");
                    }

                    // Debugger.Break();
                    //Data_tbl.Columns.Add("ProductID");
                    for (int i = 0; i < list.Count(); i++)
                    {
                        Data_tbl.Rows[i]["ProductID"] = list[i];
                    }
                }
                catch
                {
                    MessageBox.Show("Không tìm thấy productID");
                }
            }

            return Data_tbl;
        }

        public DataTable load_data_logfile_unmating_pulltest_old(string in_src, string tb_name, string infor)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
            DataTable Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, tb_name, filter_str).Clone();
            Data_tbl.Columns.Add("Select_Img", typeof(bool));
            Data_tbl.Columns.Add("Select_Grp", typeof(bool));

            DirectoryInfo tar_parent = new DirectoryInfo(in_src);
            DirectoryInfo[] arr_dic_child = tar_parent.GetDirectories();
            int ID = 1;
            if (arr_dic_child.Length > 0)
            {
                foreach (DirectoryInfo tar_d in arr_dic_child)
                {
                    SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_lst_result =
                        new SortedDictionary<int, Funtion_SMT.Peeltest_data>();
                    SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
                    Get_Image_comment3(tar_d.FullName, ref dic_image);
                    FileInfo[] temp_lst = tar_d.GetFiles("*.xlsx");

                    if (temp_lst.Length > 0)
                    {
                        Dictionary<string, SortedDictionary<int, string>> lst_result =
                            new Dictionary<string, SortedDictionary<int, string>>();

                        for (int i = 0; i < temp_lst.Length; i++)
                        {
                            ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                            ExcelWorksheet wrksht = wrkbk.Worksheets[0];

                            Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                            data.data_val = get_data_val_comment3(wrksht, "Max");
                            data.grap_data = get_image_excel(wrksht);

                            string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                                .TrimEnd(new char[] { ',', '.', ' ' });
                            f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                            if (myCode.IsNumeric(f_na))
                            {
                                int f_inx = Convert.ToInt32(f_na);
                                if (!dic_lst_result.ContainsKey(f_inx))
                                {
                                    dic_lst_result.Add(f_inx, data);
                                }
                            }
                        }
                    }

                    int k = 0;
                    foreach (var data in dic_lst_result)
                    {
                        if (k < dic_image.Count)
                        {
                            Image img = dic_image[k];
                            Image graph = data.Value.grap_data;
                            string val = data.Value.data_val;
                            Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor, tar_d.Name, data.Key,
                                img, graph, val, txtOperator.Text, DateTime.Now.ToString(), in_src, true, true);
                            ID++;
                            k++;
                        }
                    }
                }
            }

            else
            {
                SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_lst_result =
                    new SortedDictionary<int, Funtion_SMT.Peeltest_data>();
                SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
                Get_Image_comment3(tar_parent.FullName, ref dic_image);
                FileInfo[] temp_lst = tar_parent.GetFiles("*.xlsx");


                if (temp_lst.Length > 0)
                {
                    Dictionary<string, SortedDictionary<int, string>> lst_result =
                        new Dictionary<string, SortedDictionary<int, string>>();
                    int i = 0;
                    foreach (var img in dic_image.Values)
                    {
                        if (i < temp_lst.Length)
                        {
                            ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                            ExcelWorksheet wrksht = wrkbk.Worksheets[0];
                            Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                            data.data_val = get_data_val_comment3(wrksht, "Max");
                            data.grap_data = get_image_excel(wrksht);
                            string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                                .TrimEnd(new char[] { ',', '.', ' ' });

                            f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                            if (myCode.IsNumeric(f_na))
                            {
                                int f_inx = Convert.ToInt32(f_na);
                                if (!dic_lst_result.ContainsKey(f_inx))
                                {
                                    dic_lst_result.Add(f_inx, data);
                                }
                            }

                            i++;
                        }
                    }
                }

                int k = 0;
                foreach (var data in dic_lst_result)
                {
                    if (k < dic_image.Count)
                    {
                        Image img = dic_image[k];
                        Image graph = data.Value.grap_data;
                        string val = data.Value.data_val;
                        Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor, 1, data.Key, img, graph,
                            val, txtOperator.Text, DateTime.Now.ToString(), in_src, true, true);
                        ID++;
                        k++;
                    }
                }
            }

            return Data_tbl;
        }

        public DataTable load_data_logfile_unmating_pulltest(string in_src, string tb_name, string infor)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" },
                new string[] { txtItemCode.Text, txtLotNo.Text });
            DataTable Data_tbl = IQCUmatingPullTestService.getStructor();
            Data_tbl.Columns.Add("Select_Img", typeof(bool));
            Data_tbl.Columns.Add("Select_Grp", typeof(bool));

            DirectoryInfo tar_parent = new DirectoryInfo(in_src);
            int ID = 1;

            SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_lst_result =
                new SortedDictionary<int, Funtion_SMT.Peeltest_data>();
            SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
            /****/
            FileInfo[] temp_lst_image = tar_parent.GetFiles("*.jpg").Concat(tar_parent.GetFiles("*.jpeg")).ToArray();
            if (temp_lst_image.Length > 0)
            {
                Get_Image_comment3(tar_parent.FullName, ref dic_image);
            }
            else
            {
                foreach (DirectoryInfo dir in tar_parent.GetDirectories())
                {
                    if (dir.Name.Replace(" ", "").ToUpper().Contains("ANH"))
                    {
                        Get_Image_comment3(dir.FullName, ref dic_image);
                        break;
                    }
                }
            }

            /****/

            FileInfo[] temp_lst = tar_parent.GetFiles("*.xlsx");

            if (temp_lst.Length > 0)
            {
                int i = 0;
                int k = 1;
                foreach (var img in dic_image.Values)
                {
                    if (i < temp_lst.Length)
                    {
                        string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                            .TrimEnd(new char[] { ',', '.', ' ' });
                        if (myCode.IsNumeric(f_na))
                        {
                            ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                            ExcelWorksheet wrksht = wrkbk.Worksheets[0];
                            Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                            data.data_val = get_data_val_comment3(wrksht, "Max");
                            data.grap_data = get_image_excel(wrksht);


                            //f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                            //if (myCode.IsNumeric(f_na))
                            //{
                            //    int f_inx = Convert.ToInt32(f_na);
                            //    if (!dic_lst_result.ContainsKey(f_inx))
                            //    {
                            //        dic_lst_result.Add(f_inx, data);
                            //    }
                            //}


                            if (!dic_lst_result.ContainsKey(int.Parse(f_na)))
                            {
                                dic_lst_result.Add(int.Parse(f_na), data);
                                k++;
                            }
                        }

                        i++;
                    }
                }
            }

            for (int k = 1; k <= dic_image.Count; k++)
            {
                Image img = dic_image[k - 1];
                Image graph = image_null;
                string val = "";
                if (k <= dic_lst_result.Count)
                {
                    if (dic_lst_result.ContainsKey(k))
                    {
                        graph = dic_lst_result[k].grap_data;
                        val = dic_lst_result[k].data_val;
                    }
                }

                Data_tbl.Rows.Add(ID, txt_itemcode_nvl.Text, txt_lotno_nvl.Text, sheet + infor, 1, k, img, graph, val,
                    txtOperator.Text, DateTime.Now.ToString(), in_src, true, true);
                ID++;
            }

            return Data_tbl;
        }

        public void get_data_logifle_coupon_old(DirectoryInfo tar_d, int ID, ref DataTable Data_tbl, string in_src,
            string infor)
        {
            SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_lst_result =
                new SortedDictionary<int, Funtion_SMT.Peeltest_data>();
            SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
            Get_Image_comment3(tar_d.FullName, ref dic_image);
            FileInfo[] temp_lst = tar_d.GetFiles("*.xlsx");
            if (temp_lst.Length > 0)
            {
                Dictionary<string, SortedDictionary<int, string>> lst_result =
                    new Dictionary<string, SortedDictionary<int, string>>();

                for (int i = 0; i < temp_lst.Length; i++)
                {
                    ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                    ExcelWorksheet wrksht = wrkbk.Worksheets[0];
                    Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                    data.data_val = get_data_val_comment3(wrksht, "Average");
                    data.grap_data = get_image_excel(wrksht);

                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                        .TrimEnd(new char[] { ',', '.', ' ' });
                    f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                    if (myCode.IsNumeric(f_na))
                    {
                        int f_inx = Convert.ToInt32(f_na);
                        if (!dic_lst_result.ContainsKey(f_inx))
                        {
                            dic_lst_result.Add(f_inx, data);
                        }
                    }
                }
            }


            int k = 0;
            foreach (var data in dic_lst_result)
            {
                if (k < dic_image.Count)
                {
                    Image img = dic_image[k];
                    Image graph = data.Value.grap_data;
                    string val = data.Value.data_val;
                    Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor, tar_d.Name, data.Key, img,
                        graph, val, txtOperator.Text, DateTime.Now.ToString(), in_src, true, true);
                    ID++;
                    k++;
                }
            }
        }

        public void get_multi_coupon(DirectoryInfo tar_parent, ref DataTable Data_tbl, string in_src, string infor,
            ref int ID)
        {
            DirectoryInfo[] arr_dic_child = tar_parent.GetDirectories();

            if (arr_dic_child.Length > 0)
            {
                DirectoryInfo tar_sheet = null;
                foreach (DirectoryInfo tar_d in arr_dic_child)
                {
                    if (sheet.Contains(tar_d.Name.ToUpper()))
                    {
                        tar_sheet = tar_d;
                    }
                }

                if (tar_sheet != null)
                {
                    get_multi_coupon(tar_sheet, ref Data_tbl, in_src, infor, ref ID);
                }
                else
                {
                    foreach (DirectoryInfo tar_d2 in arr_dic_child)
                    {
                        if (tar_d2.Name.Replace(" ", "").ToUpper().Contains("ANH"))
                        {
                            bool d = true;
                            if (tar_parent.FullName == in_src)
                            {
                                d = false;
                            }

                            get_data_logifle_coupon_folder_ANH(tar_d2, ID, ref Data_tbl, tar_parent.Parent.Name, in_src,
                                infor, d);
                            break;
                        }
                        else
                        {
                            if (tar_d2.GetDirectories().Length > 0)
                            {
                                get_multi_coupon(tar_d2, ref Data_tbl, in_src, infor, ref ID);
                            }
                            else
                            {
                                get_data_logifle_coupon(tar_d2, ID, ref Data_tbl, tar_parent.Name, in_src, infor, true);
                            }
                        }
                    }
                }
            }
            else
            {
                //insert_tbl(tar_parent, ref Data_tbl, "", in_src, infor, false);
                get_data_logifle_coupon(tar_parent, ID, ref Data_tbl, tar_parent.Name, in_src, infor, false);
            }
        }


        public void get_data_logifle_coupon(DirectoryInfo tar_d, int ID, ref DataTable Data_tbl, string str_parent,
            string in_src, string infor, bool d)
        {
            SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_lst_result =
                new SortedDictionary<int, Funtion_SMT.Peeltest_data>();
            SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
            Get_Image_comment3(tar_d.FullName, ref dic_image);
            FileInfo[] temp_lst = tar_d.GetFiles("*.xlsx");
            if (temp_lst.Length > 0)
            {
                Dictionary<string, SortedDictionary<int, string>> lst_result =
                    new Dictionary<string, SortedDictionary<int, string>>();

                for (int i = 0; i < temp_lst.Length; i++)
                {
                    ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                    ExcelWorksheet wrksht = wrkbk.Worksheets[0];
                    Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                    data.data_val = get_data_val_comment3(wrksht, "Average");
                    data.grap_data = get_image_excel(wrksht);

                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                        .TrimEnd(new char[] { ',', '.', ' ' });
                    f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                    if (myCode.IsNumeric(f_na))
                    {
                        int f_inx = Convert.ToInt32(f_na);
                        if (!dic_lst_result.ContainsKey(f_inx))
                        {
                            dic_lst_result.Add(f_inx, data);
                        }
                    }
                }
            }

            if (str_parent.ToUpper() == "PSA" || str_parent.ToUpper() == "LINER" ||
                str_parent.Contains(txtItemCode.Text))
            {
                str_parent = "";
            }

            int k = 0;
            foreach (var data in dic_lst_result)
            {
                //if (k < dic_image.Count)
                //{
                Image img = image_null;
                if (dic_image.ContainsKey(k))
                {
                    img = dic_image[k];
                }

                Image graph = data.Value.grap_data;
                string val = data.Value.data_val;
                string region = "_";
                if (d)
                {
                    region = str_parent + "_" + tar_d.Name;
                }

                byte[] img_byte = TDMK_ImageConverter.ImageToByteArray(img, ImageFormat.Jpeg);
                byte[] graph_byte = TDMK_ImageConverter.ImageToByteArray(graph, ImageFormat.Jpeg);
                // Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor, region, data.Key, img, graph, val, txtOperator.Text, DateTime.Now.ToString(), in_src, true, true);
                Data_tbl.Rows.Add(ID, txt_itemcode_nvl.Text, txt_lotno_nvl.Text, sheet + infor, region, data.Key,
                    img_byte, graph_byte, val, txtOperator.Text, DateTime.Now.ToString(), in_src, true, true);
                ID++;
                k++;
                //}
            }
        }

        public void get_data_logifle_coupon_folder_ANH(DirectoryInfo tar_d, int ID, ref DataTable Data_tbl,
            string str_parent, string in_src, string infor, bool d)
        {
            SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_lst_result =
                new SortedDictionary<int, Funtion_SMT.Peeltest_data>();
            SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
            Get_Image_comment3(tar_d.FullName, ref dic_image);
            FileInfo[] temp_lst = new DirectoryInfo(tar_d.Parent.FullName).GetFiles("*.xlsx");
            if (temp_lst.Length > 0)
            {
                Dictionary<string, SortedDictionary<int, string>> lst_result =
                    new Dictionary<string, SortedDictionary<int, string>>();

                for (int i = 0; i < temp_lst.Length; i++)
                {
                    ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                    ExcelWorksheet wrksht = wrkbk.Worksheets[0];
                    Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                    data.data_val = get_data_val_comment3(wrksht, "Average");
                    data.grap_data = get_image_excel(wrksht);

                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                        .TrimEnd(new char[] { ',', '.', ' ' });
                    f_na = f_na.Replace("+", string.Empty).Replace("-", String.Empty);
                    if (myCode.IsNumeric(f_na))
                    {
                        int f_inx = Convert.ToInt32(f_na);
                        if (!dic_lst_result.ContainsKey(f_inx))
                        {
                            dic_lst_result.Add(f_inx, data);
                        }
                    }
                }
            }

            if (str_parent.ToUpper() == "PSA" || str_parent.ToUpper() == "LINER" ||
                str_parent.Contains(txtItemCode.Text))
            {
                str_parent = "";
            }

            int k = 0;
            foreach (var data in dic_lst_result)
            {
                if (k < dic_image.Count)
                {
                    Image img = image_null;
                    if (dic_image.ContainsKey(k))
                    {
                        img = dic_image[k];
                    }

                    Image graph = data.Value.grap_data;
                    string val = data.Value.data_val;
                    string region = "_";
                    if (d)
                    {
                        region = str_parent + "_" + tar_d.Parent.Name;
                    }

                    // Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor, region, data.Key, img, graph, val, txtOperator.Text, DateTime.Now.ToString(), in_src, true, true);
                    Data_tbl.Rows.Add(ID, txt_itemcode_nvl.Text, txt_lotno_nvl.Text, sheet + infor, region, data.Key,
                        img, graph, val, txtOperator.Text, DateTime.Now.ToString(), in_src, true, true);
                    ID++;
                    k++;
                }
            }
        }

        public DataTable load_data_logfile_coupon_OLD(string in_src, string tb_name, string infor)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text });
            DataTable Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, tb_name, filter_str).Clone();

            Data_tbl.Columns.Add("Select_Img", typeof(bool));
            Data_tbl.Columns.Add("Select_Grp", typeof(bool));

            DirectoryInfo tar_parent = new DirectoryInfo(in_src);
            DirectoryInfo[] arr_dic_child = tar_parent.GetDirectories();
            int ID = 1;

            if (arr_dic_child.Length > 0)
            {
                foreach (DirectoryInfo tar_d2 in arr_dic_child)
                {
                    if (sheet.Contains(tar_d2.Name.ToUpper()))
                    {
                        if (tar_d2.GetDirectories().Length > 0)
                        {
                            foreach (DirectoryInfo tar_d1 in tar_d2.GetDirectories())
                            {
                                if (tar_d1.GetDirectories().Length > 0)
                                {
                                    foreach (DirectoryInfo tar_d in tar_d1.GetDirectories())
                                    {
                                        get_data_logifle_coupon(tar_d, ID, ref Data_tbl, tar_d1.Name, in_src, infor,
                                            true);
                                    }
                                }
                                else
                                {
                                    get_data_logifle_coupon(tar_d1, ID, ref Data_tbl, tar_d2.Name, in_src, infor, true);
                                }
                            }
                        }
                        else
                        {
                            get_data_logifle_coupon(tar_d2, ID, ref Data_tbl, "", in_src, infor, false);
                        }

                        break;
                    }
                }
            }
            else
            {
                get_data_logifle_coupon(tar_parent, ID, ref Data_tbl, "", in_src, infor, false);
            }

            return Data_tbl;
        }

        public DataTable load_data_logfile_coupon(string in_src, string tb_name, string infor)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" },
                new string[] { txtItemCode.Text, txtLotNo.Text });
            DataTable Data_tbl = new DataTable();
            Data_tbl.Columns.Add("ID", typeof(int));
            Data_tbl.Columns.Add("ItemCode", typeof(string));
            Data_tbl.Columns.Add("LotNo", typeof(string));
            Data_tbl.Columns.Add("Sheet", typeof(string));
            Data_tbl.Columns.Add("Region", typeof(string));
            Data_tbl.Columns.Add("Sample", typeof(string));
            Data_tbl.Columns.Add("Image", typeof(byte[]));
            Data_tbl.Columns.Add("Graph", typeof(byte[]));
            Data_tbl.Columns.Add("Data", typeof(string));
            Data_tbl.Columns.Add("Operator", typeof(string));
            Data_tbl.Columns.Add("Time_Update", typeof(string));
            Data_tbl.Columns.Add("Remark", typeof(string));

            Data_tbl.Columns.Add("Select_Img", typeof(bool));
            Data_tbl.Columns.Add("Select_Grp", typeof(bool));
            DirectoryInfo tar_parent = new DirectoryInfo(in_src);
            DirectoryInfo[] arr_dic_child = tar_parent.GetDirectories();
            int ID = 1;
            get_multi_coupon(tar_parent, ref Data_tbl, in_src, infor, ref ID);


            return Data_tbl;
        }


        public void Get_Sheartest_logfile_Multi(string in_src,
            ref SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_lst_result, string textfind)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
            List<Image> lst_graph = new List<Image> { };

            Get_Image_comment3(in_src, ref dic_image);
            FileInfo[] temp_lst = tar_d.GetFiles("*.xlsx");
            if (temp_lst.Length > 0)
            {
                myExcel.Application xlsApp = TDMK_Code.StartExcel();
                myExcel.Workbook wrkbk = xlsApp.Workbooks.Open(temp_lst[0].FullName);
                myExcel.Worksheet wrksht = wrkbk.Sheets[1];
                List<string> lst_data = new List<string> { };
                lst_graph = get_image_excel_ShearTest(wrksht, ref lst_data);
                wrkbk.Close();

                Dictionary<string, SortedDictionary<int, string>> lst_result =
                    new Dictionary<string, SortedDictionary<int, string>>();
                int i = 0;
                foreach (var img in dic_image)
                {
                    if (i < lst_graph.Count)
                    {
                        Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                        //data.data_val = get_data_val_comment3(wrksht, textfind); 
                        data.data_val = lst_data[i];
                        data.grap_data = lst_graph[i];
                        data.image_data = img.Value;
                        int f_inx = img.Key;
                        if (!dic_lst_result.ContainsKey(f_inx))
                        {
                            dic_lst_result.Add(f_inx, data);
                        }

                        i++;
                    }
                }
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_Sheartest_logfile_Multi(inter_lst.FullName, ref dic_lst_result, textfind);
                    }
                }
            }
        }

        public DataTable load_data_logfile_sheartest(string in_src, string infor)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" },
                new string[] { txtItemCode.Text, txtLotNo.Text });
            DataTable Data_tbl = ShearTestServices.getStructor();
            Data_tbl.Columns.Add("Select_Img", typeof(bool));
            Data_tbl.Columns.Add("Select_Grp", typeof(bool));

            DirectoryInfo tar_parent = new DirectoryInfo(in_src);
            DirectoryInfo[] arr_dic_child = tar_parent.GetDirectories();
            int ID = 1;

            if (arr_dic_child.Length > 0)
            {
                foreach (DirectoryInfo tar_d in arr_dic_child)
                {
                    SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_lst_result =
                        new SortedDictionary<int, Funtion_SMT.Peeltest_data>();
                    SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
                    List<Image> lst_graph = new List<Image> { };
                    Get_Image_comment3(tar_d.FullName, ref dic_image);
                    FileInfo[] temp_lst = tar_d.GetFiles("*.xlsx");
                    if (temp_lst.Length > 0)
                    {
                        myExcel.Application xlsApp = TDMK_Code.StartExcel();
                        myExcel.Workbook wrkbk = xlsApp.Workbooks.Open(temp_lst[0].FullName);
                        myExcel.Worksheet wrksht = wrkbk.Sheets[1];
                        List<string> lst_data = new List<string> { };
                        lst_graph = get_image_excel_ShearTest(wrksht, ref lst_data);
                        wrkbk.Close();

                        Dictionary<string, SortedDictionary<int, string>> lst_result =
                            new Dictionary<string, SortedDictionary<int, string>>();
                        int i = 0;
                        foreach (var img in dic_image)
                        {
                            if (i < lst_graph.Count)
                            {
                                Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                                data.data_val = lst_data[i];
                                data.grap_data = lst_graph[i];
                                data.image_data = img.Value;
                                int f_inx = img.Key;
                                if (!dic_lst_result.ContainsKey(f_inx))
                                {
                                    dic_lst_result.Add(f_inx, data);
                                }

                                i++;
                            }
                        }
                    }

                    int sample = 1;
                    foreach (var data in dic_lst_result)
                    {
                        Image img = data.Value.image_data;
                        Image graph = data.Value.grap_data;
                        string val = data.Value.data_val;
                        Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor, tar_d.Name, sample, img,
                            graph, val, "0.00%", "0.00%", "0.00%", "0.00%", "100%", "0.00%", "0.00%", txtOperator.Text,
                            DateTime.Now.ToString(), in_src, true, true);
                        ID++;
                        sample++;
                    }
                }
            }
            else
            {
                SortedDictionary<int, Funtion_SMT.Peeltest_data> dic_lst_result =
                    new SortedDictionary<int, Funtion_SMT.Peeltest_data>();
                SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
                List<Image> lst_graph = new List<Image> { };
                Get_Image_comment3(tar_parent.FullName, ref dic_image);
                FileInfo[] temp_lst = tar_parent.GetFiles("*.xlsx");
                if (temp_lst.Length > 0)
                {
                    myExcel.Application xlsApp = TDMK_Code.StartExcel();
                    myExcel.Workbook wrkbk = xlsApp.Workbooks.Open(temp_lst[0].FullName);
                    myExcel.Worksheet wrksht = wrkbk.Sheets[1];
                    List<string> lst_data = new List<string> { };
                    lst_graph = get_image_excel_ShearTest(wrksht, ref lst_data);
                    wrkbk.Close();

                    Dictionary<string, SortedDictionary<int, string>> lst_result =
                        new Dictionary<string, SortedDictionary<int, string>>();
                    int i = 0;
                    foreach (var img in dic_image)
                    {
                        if (i < lst_graph.Count)
                        {
                            Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                            data.data_val = lst_data[i];
                            data.grap_data = lst_graph[i];
                            data.image_data = img.Value;

                            int f_inx = img.Key;
                            if (!dic_lst_result.ContainsKey(f_inx))
                            {
                                dic_lst_result.Add(f_inx, data);
                            }

                            i++;
                        }
                    }
                }

                int sample = 1;
                foreach (var data in dic_lst_result)
                {
                    Image img = data.Value.image_data;
                    Image graph = data.Value.grap_data;
                    string val = data.Value.data_val;
                    Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text, sheet + infor, "1", sample, img, graph, val,
                        "0.00%", "0.00%", "0.00%", "0.00%", "100%", "0.00%", "0.00%", txtOperator.Text,
                        DateTime.Now.ToString(), in_src, true, true);
                    ID++;
                    sample++;
                }
            }

            try
            {
                ProductIDService productIDService = new ProductIDService(txtItemCode.Text, txtLotNo.Text, textBox1.Text,
                    new[] { "OQC", "B2B", "Shearing" }, new[] { txtItemCode.Text });
                if (productIDService._listFile.Count > 0)
                {
                    Data_tbl.Columns.Add("ProductID");
                    int i = 0;
                    List<string> listPID =
                        productIDService.getListProductID(productIDService._listFile[txtItemCode.Text]);
                    foreach (string item in listPID)
                    {
                        try
                        {
                            Data_tbl.Rows[i]["ProductID"] = item;
                        }
                        catch
                        {
                            break;
                        }

                        i++;
                    }
                }
            }
            catch
            {
                MessageBox.Show("ProductID: Không lấy được PID");
            }

            return Data_tbl;
        }


        public void Get_comment3_logfile_Multi(string in_src,
            ref SortedDictionary<string, Funtion_SMT.Peeltest_data> dic_lst_result, string textfind)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            SortedDictionary<int, Image> dic_image = new SortedDictionary<int, Image>();
            Get_Image_comment3(in_src, ref dic_image);
            FileInfo[] temp_lst = tar_d.GetFiles("*.xlsx");
            if (temp_lst.Length > 0)
            {
                Dictionary<string, SortedDictionary<int, string>> lst_result =
                    new Dictionary<string, SortedDictionary<int, string>>();
                int i = 0;
                foreach (var img in dic_image.Values)
                {
                    if (i < temp_lst.Length)
                    {
                        ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(temp_lst[i].FullName);
                        ExcelWorksheet wrksht = wrkbk.Worksheets[0];
                        Funtion_SMT.Peeltest_data data = new Funtion_SMT.Peeltest_data();
                        data.data_val = get_data_val_comment3(wrksht, textfind);
                        //data.data_val = Get_LogFile_Data(temp_lst[i].FullName, "Min"); 
                        data.grap_data = get_image_excel(wrksht);
                        data.image_data = img;
                        string f_inx = Path.GetFileNameWithoutExtension(temp_lst[i].Name)
                            .TrimEnd(new char[] { ',', '.', ' ' });
                        if (!dic_lst_result.ContainsKey(f_inx))
                        {
                            dic_lst_result.Add(f_inx, data);
                        }

                        i++;
                    }
                }
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_comment3_logfile_Multi(inter_lst.FullName, ref dic_lst_result, textfind);
                    }
                }
            }
        }
        //public void get_data_peeltest(string file_path)
        //{

        //                Get_Impedance_logfile_Multi2(file_path, ref Impedance_result);
        //                string sel_graph = Impedance_result.Keys.ToArray()[0];
        //                SortedDictionary<int, myVar.Impedance_data> graph_data = Impedance_result["1"];
        //                int r_inx = 0;
        //                foreach (var item in Impedance_result)
        //                {
        //                    var cur_val = item.Value;
        //                    int pcs_no = 0;
        //                    foreach (var t in cur_val)
        //                    {
        //                        Imp_data_dt.Rows.Add(r_inx + 1, txtItemCode.Text, fill_name(txtLotNo.Text), t.Key, no_region + cur_coupon, t.Value.data_val, zone, src_path);
        //                        pcs_no++;
        //                        r_inx++;
        //                    }
        //                    no_region++;
        //                }
        //                DGV_Impedance.DataSource = Imp_data_dt;
        //                // DGV_Graph.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text }));
        //                r_inx = DGV_Graph.Rows.Count + 1;
        //                //if (DGV_Graph.Rows.Count > 0)
        //                //    Graph_dt = (DataTable)DGV_Graph.DataSource;

        //                foreach (var t in graph_data)
        //                {
        //                    if (t.Value.grap_data != null)
        //                    {
        //                        //Graph_dt.Rows.Add(r_inx, txtItemCode.Text, txtLotNo.Text, t.Key.ToString(), 1, t.Value.grap_data, txtOperator.Text, zone, src_path);
        //                        cur_graph_dt.Rows.Add(r_inx, txtItemCode.Text, txtLotNo.Text, t.Key.ToString(), 1, t.Value.grap_data, txtOperator.Text, zone, src_path);
        //                    }
        //                    else
        //                    {
        //                        break;
        //                    }
        //                    r_inx++;

        //                }
        //                //sort_data_impedance(DGV_Graph, "IMPEDANCE_GRAPH", Graph_dt);
        //                DGV_Graph.DataSource = cur_graph_dt;

        //                ((DataGridViewImageColumn)DGV_Graph.Columns["Image_Graph"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
        //                ((DataGridViewImageColumn)DGV_Graph.Columns["Image_Graph"]).Width = 100;
        //                foreach (DataGridViewRow dr in DGV_Graph.Rows)
        //                {
        //                    dr.Height = 70;
        //                }

        //                btnSummary.Enabled = false;

        //}
        public string Get_LogFile_Data(string in_src_file, string textfind)
        {
            string value = "";
            SortedDictionary<int, string> result_lst = new SortedDictionary<int, string>();
            string[] Lines = System.IO.File.ReadAllLines(in_src_file);

            foreach (var line in Lines)
            {
                string str = new string(line.Where(s => s != '"').ToArray());
                string[] temp = str.Split(',');
                foreach (string val in temp)
                {
                    if (myCode.checkDBNull(val) == textfind)
                    {
                        value = val;
                        return value;
                    }
                }
            }

            return value;
        }

        public SortedDictionary<int, string> Get_csv_Data(string in_src_file)
        {
            SortedDictionary<int, string> result_lst = new SortedDictionary<int, string>();
            string[] Lines = File.ReadAllLines(in_src_file);
            foreach (var line in Lines)
            {
                string str = new string(line.Where(s => s != '"').ToArray());
                string[] temp = str.Split(',');
                if (temp[0] != "")
                {
                    if (myCode.IsNumeric(temp[0]))
                    {
                        result_lst.Add(Convert.ToInt32(temp[0]), temp[2]);
                    }
                }
                else
                {
                    break;
                }
            }

            return result_lst;
        }

        private void btn_loaddata_Click(object sender, EventArgs e)
        {
            string textfind = "Max";
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" },
                new string[] { txtItemCode.Text, txtLotNo.Text });
            DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", filter_str).Clone();
            DataTable Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, tbl_name_comment3, filter_str).Clone();
            SortedDictionary<string, Funtion_SMT.Peeltest_data> dic_result =
                new SortedDictionary<string, Funtion_SMT.Peeltest_data>();

            Data_tbl.Columns.Add("Select", typeof(bool));


            Get_comment3_logfile_Multi(txtLogfile.Text, ref dic_result, textfind);
            int ID = TDMK_Code.SQL_MAX(tbl_name_comment3, "ID", sqlcon) + 1;
            foreach (var data in dic_result)
            {
                Image img = data.Value.image_data;
                Image graph = data.Value.grap_data;
                string val = data.Value.data_val;

                Data_tbl.Rows.Add(ID, txtItemCode.Text, txtLotNo.Text,
                    sheet.ToUpper().Replace("-", "").Replace(" ", ""), img, graph, val, "", "", "", false);
                ID++;
            }

            dgv_logfile.DataSource = Data_tbl;
            //DataGridViewCheckBoxColumn dgvcCheckBox = new DataGridViewCheckBoxColumn();
            //dgvcCheckBox.HeaderText = "Select"; 
            //dgv_logfile.Columns.Add(dgvcCheckBox);


            resize_column_image(dgv_logfile);

            // BatchBulkCopy(sqlcon, Data_tbl, "COMMENT_3");
        }

        private void dgv_Analysis_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //int col_inx = e.ColumnIndex;
            //int r_inx = e.RowIndex;
            //DataGridViewCell cur_cell = dgv_Analysis.CurrentCell;
            //if (dgv_Analysis.Columns[col_inx].Name.Contains("Image") || dgv_Analysis.Columns[col_inx].Name.Contains("Graph"))
            //{
            //    if (myCode.checkDBNull(cur_cell.Value) != "")
            //    {
            //        View_detail_Image fr1 = new View_detail_Image() { TopMost = true };
            //        fr1.data = (byte[])cur_cell.Value;
            //        fr1.Show();

            //    }

            //}
        }

        public string count_selected()
        {
            string str_selected = "     Selected" + Environment.NewLine + Environment.NewLine;

            if (sheet == "CROSS_SECTION" || sheet == "GAP_CONNECTOR")
            {
                DataTable Data_tbl = (DataTable)dgv_logfile.DataSource;
                List<DataTable> lst_Table = new List<DataTable> { };
                Get_ListTable(-1, Data_tbl, new string[] { "Region" }, ref lst_Table, "Data");
                foreach (DataTable dt in lst_Table)
                {
                    string region = dt.Rows[0]["Region"].ToString();
                    int count_select = 0;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if ((bool)dt.Rows[i]["Select"] == true)
                        {
                            count_select++;
                        }
                    }

                    str_selected += region + " : " + count_select.ToString() + Environment.NewLine +
                                    Environment.NewLine;
                }
            }
            else if (sheet.Contains("UNMATING"))
            {
                DataTable Data_tbl = (DataTable)dgv_logfile.DataSource;
                List<DataTable> lst_Table = new List<DataTable> { };
                Get_ListTable(-1, Data_tbl, new string[] { "Region" }, ref lst_Table, "Data");
                foreach (DataTable dt in lst_Table)
                {
                    int count_select_Img = 0;
                    int count_select_Grp = 0;
                    string region = dt.Rows[0]["Region"].ToString();
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if ((bool)dt.Rows[i]["Select_Img"] == true)
                        {
                            count_select_Img++;
                        }

                        if ((bool)dt.Rows[i]["Select_Grp"] == true && myCode.checkDBNull(dt.Rows[i]["Data"]) != "")
                        {
                            count_select_Grp++;
                        }
                    }

                    str_selected += "LK" + region + Environment.NewLine + "Image: " + count_select_Img.ToString() +
                                    Environment.NewLine + "Graph: " + count_select_Grp.ToString() +
                                    Environment.NewLine + Environment.NewLine;
                }
            }
            else
            {
                DataTable Data_tbl = (DataTable)dgv_logfile.DataSource;
                List<DataTable> lst_Table = new List<DataTable> { };
                Get_ListTable(-1, Data_tbl, new string[] { "Region" }, ref lst_Table, "Data");
                foreach (DataTable dt in lst_Table)
                {
                    int count_select_Img = 0;
                    int count_select_Grp = 0;
                    string region = dt.Rows[0]["Region"].ToString();
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if ((bool)dt.Rows[i]["Select_Img"] == true)
                        {
                            count_select_Img++;
                        }

                        if ((bool)dt.Rows[i]["Select_Grp"] == true)
                        {
                            count_select_Grp++;
                        }
                    }

                    str_selected += "LK" + region + Environment.NewLine + "Image: " + count_select_Img.ToString() +
                                    Environment.NewLine + "Graph: " + count_select_Grp.ToString() +
                                    Environment.NewLine + Environment.NewLine;
                }
            }

            return str_selected;
        }
        //public int count_selected()
        //{
        //    int count_select = 0;
        //    if (sheet == "CROSS_SECTION" || sheet == "GAP_CONNECTOR")
        //    {
        //        for (int i = 0; i < dgv_logfile.Rows.Count; i++)
        //        {
        //            if ((bool)dgv_logfile.Rows[i].Cells["Select"].Value == true)
        //            {
        //                count_select++;
        //            }
        //        }
        //    }
        //    else
        //    {
        //        int count_select_Img = 0;
        //        int count_select_Grp = 0;
        //        for (int i = 0; i < dgv_logfile.Rows.Count; i++)
        //        {
        //            if ((bool)dgv_logfile.Rows[i].Cells["Select_Img"].Value == true)
        //            {
        //                count_select_Img++;
        //            }
        //            if ((bool)dgv_logfile.Rows[i].Cells["Select_Grp"].Value == true)
        //            {
        //                count_select_Grp++;
        //            }
        //        }

        //        count_select = new int[] { count_select_Img, count_select_Grp }.Min();
        //    }
        //    return count_select;
        //}

        public void select_data(ref bool select)
        {
            if (dgv_logfile.Columns.Contains("Select_Img") && dgv_logfile.Columns.Contains("Select_Grp"))
            {
                SortedDictionary<int, int> dic_img = new SortedDictionary<int, int> { };
                SortedDictionary<int, int> dic_graph = new SortedDictionary<int, int> { };
                int indx_img = 1;
                int indx_grp = 1;
                for (int i = 0; i < dgv_logfile.Rows.Count; i++)
                {
                    bool checkedCell_Img = (bool)dgv_logfile.Rows[i].Cells["Select_Img"].Value;
                    bool checkedCell_Grp = (bool)dgv_logfile.Rows[i].Cells["Select_Grp"].Value;

                    if (checkedCell_Img == true)
                    {
                        dic_img.Add(indx_img, i);
                        indx_img++;
                    }

                    if (checkedCell_Grp == true)
                    {
                        if (dgv_logfile.Rows[i].Cells["Data"].Style.BackColor != Color.Red)
                        {
                            dic_graph.Add(indx_grp, i);
                            indx_grp++;
                        }
                        else
                        {
                            if (UserSession.Instance.IsLoggedIn)
                            {
                                dic_graph.Add(indx_grp, i);
                                indx_grp++;
                            }
                            else
                            {
                                MessageBox.Show(new Form { TopMost = true }, "Vui lòng để đăng nhập để chọn dữ liệu NG",
                                    "Thông báo");
                                select = false;
                                return;
                            }
                        }
                    }
                }


                DataTable dt = new DataTable();

                if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST" || sheet == "SHEAR_TEST")
                {
                    DataTable z = (DataTable)dgv_logfile.DataSource;
                    if (z.Columns.Contains("ProductID"))
                    {
                        dt = ((DataTable)dgv_logfile.DataSource).AsDataView().ToTable(false,
                            new string[]
                            {
                                "ID", "ProductID", "ItemCode", "LotNo", "Sheet", "Region", "Sample", "Image", "Graph",
                                "Data", "Mode 1: Solder joint crack", "Mode 2: Pad lift", "Mode 3: Solder joint lift",
                                "Mode 4: Intermetallic break", "Mode 5: Component damage", "Mode 6: Component detached",
                                "Mode 7: Flex torn", "Operator", "Time_Update", "Remark"
                            });
                    }
                    else
                    {
                        dt = ((DataTable)dgv_logfile.DataSource).AsDataView().ToTable(false,
                            new string[]
                            {
                                "ID", "ItemCode", "LotNo", "Sheet", "Region", "Sample", "Image", "Graph", "Data",
                                "Mode 1: Solder joint crack", "Mode 2: Pad lift", "Mode 3: Solder joint lift",
                                "Mode 4: Intermetallic break", "Mode 5: Component damage", "Mode 6: Component detached",
                                "Mode 7: Flex torn", "Operator", "Time_Update", "Remark"
                            });
                    }
                }
                else
                {
                    dt = ((DataTable)dgv_logfile.DataSource).AsDataView().ToTable(false,
                        new string[]
                        {
                            "ID", "ItemCode", "LotNo", "Sheet", "Region", "Sample", "Image", "Graph", "Data",
                            "Operator", "Time_Update", "Remark"
                        });
                }


                DataTable dt_select = dt.Clone();
                int ID = 1;
                int min_select = new int[2] { indx_img - 1, indx_grp - 1 }.Min();

                for (int i = 1; i <= min_select; i++)
                {
                    DataRow dr = dt_select.NewRow();

                    for (int k = 1; k < dt.Columns.Count; k++)
                    {
                        dr["ID"] = ID;
                        if (k != 5 && k != 6 && k != 7 && k != 8)
                        {
                            //dr[k] = dt.Rows[i - 1][k].ToString();
                            dr[k] = dt.Rows[dic_img[i]][k].ToString();
                        }

                        dr["Sample"] = i;
                        dr["Region"] = dt.Rows[dic_img[i]]["Region"];
                        dr["Image"] = dt.Rows[dic_img[i]]["Image"];
                        dr["Graph"] = dt.Rows[dic_graph[i]]["Graph"];
                        dr["Data"] = dt.Rows[dic_graph[i]]["Data"].ToString();
                    }

                    dt_select.Rows.Add(dr);
                    ID++;
                }


                List<string> lst_region =
                    dt_select.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToList();
                if (lst_region.Count > 1)
                {
                    foreach (string s in lst_region)
                    {
                        int smp = 1;
                        foreach (DataRow dr in dt_select.Rows)
                        {
                            if (dr["Region"].ToString() == s)
                            {
                                dr["Sample"] = smp;
                                smp++;
                            }
                        }
                    }
                }

                dgv_Analysis.DataSource = dt_select;
            }
        }


        private void btn_select_Click(object sender, EventArgs e)
        {
            if (dgv_logfile.Rows.Count > 0)
            {
                bool select_OK = true;

                if (sheet == "SHEAR_TEST")
                {
                    select_data(ref select_OK);
                    if (select_OK)
                    {
                        resize_column_image(dgv_Analysis);
                        myCode.Disable_Sort_DGV(dgv_Analysis);
                        if (cb_Type.SelectedIndex != -1)
                        {
                            Check_spec_ShearTest_new(dgv_Analysis);
                        }
                    }
                }
                else
                {
                    if (sheet == "CROSS_SECTION" || sheet == "GAP_CONNECTOR")
                    {
                        DataTable dt_select;
                        if (cbStatus.Text == "Shield b2b" || cbStatus.Text == "Clip")
                        {
                            DataTable dt_logfile = (DataTable)dgv_logfile.DataSource;
                            dt_select = CrossSectionService.SelectData(dt_logfile);
                        }
                        else
                        {
                            dt_select = CrossSectionService.getStructorTableByte("GAP");
                            int ID = 1;
                            //for (int i = 0; i < dgv_logfile.Rows.Count; i++)
                            //{
                            DataTable dt_logfile = (DataTable)dgv_logfile.DataSource;
                            List<string> lst_region = dt_logfile.AsEnumerable().Select(x => x.Field<string>("Region"))
                                .Distinct().ToList();
                            foreach (string s in lst_region)
                            {
                                DataView dv = dt_logfile.AsDataView();
                                dv.RowFilter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { s });
                                for (int inx = 0; inx < dv.Count; inx++)
                                {
                                    DataRow row = dv[inx].Row;
                                    int i = dt_logfile.Rows.IndexOf(row);
                                    bool checkedCell = (bool)dgv_logfile.Rows[i].Cells["Select"].Value;
                                    DataTable dataTable = (DataTable)dgv_logfile.DataSource;
                                    if (checkedCell == true)
                                    {
                                        if (dgv_logfile.Rows[i].Cells["Data"].Style.BackColor != Color.Red)
                                        {
                                            DataRow dr = dt_select.NewRow();
                                            dr["ID"] = ID;
                                            dr["Region"] = dataTable.Rows[i]["Region"];
                                            dr["ItemCode"] = dataTable.Rows[i]["ItemCode"];
                                            dr["LotNo"] = dataTable.Rows[i]["LotNo"];
                                            dr["Sheet"] = dataTable.Rows[i]["Sheet"];
                                            dr["Sample"] = dataTable.Rows[i]["Sample"];
                                            try
                                            {
                                                if (dataTable.Rows[i]["Image"] is byte[])
                                                {
                                                    dr["Image"] = dataTable.Rows[i]["Image"];
                                                }

                                                if (dataTable.Rows[i]["Image"] is Image)
                                                {
                                                    dr["Image"] =
                                                        TDMK_ImageConverter.ImageToByteArray(
                                                            (Image)dataTable.Rows[i]["Image"], ImageFormat.Jpeg);
                                                }
                                            }
                                            catch
                                            {
                                            }

                                            try
                                            {
                                                if (dataTable.Rows[i]["Image1"] is byte[])
                                                {
                                                    dr["Image1"] = dataTable.Rows[i]["Image1"];
                                                }

                                                if (dataTable.Rows[i]["Image1"] is Image)
                                                {
                                                    dr["Image1"] =
                                                        TDMK_ImageConverter.ImageToByteArray(
                                                            (Image)dataTable.Rows[i]["Image1"], ImageFormat.Jpeg);
                                                }
                                            }
                                            catch
                                            {
                                            }

                                            try
                                            {
                                                if (dataTable.Rows[i]["Image2"] is byte[])
                                                {
                                                    dr["Image2"] = dataTable.Rows[i]["Image2"];
                                                }

                                                if (dataTable.Rows[i]["Image2"] is Image)
                                                {
                                                    dr["Image2"] =
                                                        TDMK_ImageConverter.ImageToByteArray(
                                                            (Image)dataTable.Rows[i]["Image2"], ImageFormat.Jpeg);
                                                }
                                            }
                                            catch
                                            {
                                            }

                                            dr["Data"] = dataTable.Rows[i]["Data"];
                                            dr["Operator"] = dataTable.Rows[i]["Operator"];
                                            dr["Time_Update"] = dataTable.Rows[i]["Time_Update"];
                                            dr["Remark"] = dataTable.Rows[i]["Remark"];
                                            try
                                            {
                                                dr["ProductID"] = dataTable.Rows[i]["ProductID"];
                                            }
                                            catch
                                            {
                                            }

                                            dt_select.Rows.Add(dr);
                                            ID++;
                                        }
                                        else
                                        {
                                            if (admin_mode == "Admin mode")
                                            {
                                                DataRow dr = dt_select.NewRow();
                                                dr["ID"] = ID;
                                                dr[5] = inx + 1;
                                                for (int k = 1; k < dgv_logfile.Columns.Count - 1; k++)
                                                {
                                                    if (k != 5)
                                                    {
                                                        dr[k] = dgv_logfile.Rows[i].Cells[k].Value;
                                                    }
                                                }

                                                dt_select.Rows.Add(dr);
                                                ID++;
                                            }
                                            else
                                            {
                                                // MessageBox.Show("Please login to select NG data");
                                                MessageBox.Show(new Form { TopMost = true },
                                                    "Vui lòng để đăng nhập để chọn dữ liệu NG", "Thông báo");
                                                select_OK = false;
                                                return;
                                            }
                                        }
                                    }
                                }
                            }

                            //} 
                        }

                        dgv_Analysis.DataSource = dt_select;
                    }


                    else
                    {
                        select_data(ref select_OK);
                    }

                    if (select_OK)
                    {
                        resize_column_image(dgv_Analysis);
                    }
                }


                if (dgv_Analysis.Rows.Count > 0)
                {
                    if (cb_Type.SelectedIndex != -1)
                    {
                        switch (sheet)
                        {
                            case "PEEL_TEST":
                                Check_spec_Peel_Pull(dgv_Analysis, sheet);
                                break;

                            case "MATING_PULL_TEST":
                                Check_spec_Peel_Pull(dgv_Analysis, sheet);
                                break;

                            case "SHEAR_TEST":
                                Check_spec_ShearTest_new(dgv_Analysis);
                                break;

                            case "IQC_UNMATING_PULL_TEST":
                                Check_spec_unmatingpull(dgv_Analysis, sheet);
                                break;

                            case "IQC_LINER_PEELING_COUPON":
                                check_spec_coupon(dgv_Analysis, sheet);
                                break;

                            case "IQC_PSA_PEELING_COUPON":
                                check_spec_coupon(dgv_Analysis, sheet);
                                break;

                            case "LINER_PEEL_TEST_ON_PRODUCT":
                                check_spec_onproduct(dgv_Analysis, sheet);
                                break;

                            case "PSA_PEEL_TEST_ON_PRODUCT":
                                check_spec_onproduct(dgv_Analysis, sheet);
                                break;

                            case "CROSS_SECTION":
                                Check_spec_crossection(dgv_Analysis, sheet);
                                break;

                            case "GAP_CONNECTOR":
                                Check_spec_GAP(dgv_Analysis, sheet);
                                break;
                        }
                    }

                    Check_Alldata(dgv_Analysis);
                    bool chk_ = true;
                    for (int i = 0; i < dgv_Analysis.Rows.Count; i++)
                    {
                        if (dgv_Analysis.Rows[i].Cells["Data"].Style.BackColor == Color.Red)
                        {
                            chk_ = false;
                            break;
                        }
                    }

                    if (check_enough_data(dgv_Analysis) && chk_)
                    {
                        lbl_judge.Text = "OK";
                        lbl_judge.BackColor = Color.Green;
                    }
                    else
                    {
                        lbl_judge.Text = "NG";
                        lbl_judge.BackColor = Color.Red;
                    }

                    myCode.Disable_Sort_DGV(dgv_Analysis);
                }
                else
                {
                    lbl_judge.Text = "";
                    lbl_judge.BackColor = Color.Transparent;
                }
            }
        }

        private void btnLoadB_Click(object sender, EventArgs e)
        {
            GC.Collect();
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && cb_Type.SelectedIndex != -1)
            {
                string itemCode = "", lotNo = "";
                lbl_judge.BackColor = Color.Transparent;
                lbl_judge.Text = "";
                lbl_judge_logfile.BackColor = Color.Transparent;
                string infor = "/" + txt_ItemName.Text + "_" + txt_line.Text + "_" + txt_ca.Text + "_" + txt_date.Text +
                               "_" + txt_worker.Text + "_" + cb_Type.SelectedItem.ToString();
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Sheet" },
                    new string[]
                    {
                        LegacyMode.Checked ? txtItemCode.Text : txtItemCode.Text.PadRight(10),
                        LegacyMode.Checked ? txtLotNo.Text : txtLotNo.Text.PadRight(10),
                        (LegacyMode.Checked ? sheet + infor : cb_Type.SelectedItem.ToString())
                    });
                itemCode = txtItemCode.Text;
                lotNo = txtLotNo.Text;
                if (sheet.Contains("UNMATING") || sheet.Contains("COUPON"))
                {
                    if (txt_itemcode_nvl.Text != "" && txt_lotno_nvl.Text != "")
                    {
                        itemCode = txt_itemcode_nvl.Text;
                        lotNo = txt_lotno_nvl.Text;
                        if (sheet == "IQC_PSA_PEELING_COUPON" || sheet == "IQC_UNMATING_PULL_TEST" ||
                            sheet == "IQC_LINER_PEELING_COUPON")
                        {
                            infor = "";
                        }

                        filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Sheet" },
                            new string[]
                            {
                                txt_itemcode_nvl.Text.PadRight(10), txt_lotno_nvl.Text.PadRight(30),
                                (LegacyMode.Checked ? sheet + infor : cb_Type.SelectedItem.ToString())
                            });
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Chưa nhập đủ ItemCode/lotNo", "Thông báo");
                        return;
                    }
                }

                if (sheet == "CROSS_SECTION" && (cbStatus.Text == "Shield b2b" || cbStatus.Text == "Clip"))
                {
                    filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Sheet" },
                        new string[]
                        {
                            LegacyMode.Checked ? txtItemCode.Text : txtItemCode.Text.PadRight(10),
                            LegacyMode.Checked ? txtLotNo.Text : txtLotNo.Text.PadRight(10),
                            cbStatus.Text
                        });
                }


                DataTable dt_analysis = TDMK_Code.Datatable_Filter(sqlcon,
                    sheet + (_PRIME_PEEL_TEST ? "_WITHOUT_SUS" : "") + (!LegacyMode.Checked == true ? "_NAS" : ""),
                    filter_str);

                if (!LegacyMode.Checked && dt_analysis.Rows.Count > 0)
                {
                    string data = dt_analysis.Rows[0]["Data"].ToString();
                    string location = dt_analysis.Rows[0]["LocationImg"].ToString();
                    dt_analysis = ConverterService.JsonToDataTable(data);
                    NasRepository nas = new NasRepository();
                    nas.MergeDataTable(dt_analysis, sheet, itemCode, lotNo, location);
                }

                string sheetZ = sheet + (_PRIME_PEEL_TEST ? "_WITHOUT_SUS" : "");

                try
                {
                    ProductIDService.FillProductID(dt_analysis, itemCode, lotNo, sheetZ);
                }
                catch
                {
                }

                int st = 1;
                foreach (DataRow dr in dt_analysis.Rows)
                {
                    dr["ID"] = st;
                    st++;
                }

                dgv_Analysis.DataSource = dt_analysis;
                myCode.Disable_Sort_DGV(dgv_Analysis);
                dgv_logfile.DataSource = null;


                if (dgv_Analysis.Rows.Count > 0)
                {
                    resize_column_image(dgv_logfile);
                    resize_column_image(dgv_Analysis);
                    if (cb_Type.SelectedIndex != -1)
                    {
                        switch (sheet)
                        {
                            case "PEEL_TEST":
                                Check_spec_Peel_Pull(dgv_Analysis, sheet);
                                break;

                            case "MATING_PULL_TEST":
                                Check_spec_Peel_Pull(dgv_Analysis, sheet);
                                break;

                            case "SHEAR_TEST":
                                Check_spec_ShearTest_new(dgv_Analysis);
                                break;

                            case "IQC_UNMATING_PULL_TEST":
                                Check_spec_unmatingpull(dgv_Analysis, sheet);
                                break;

                            case "IQC_LINER_PEELING_COUPON":
                                check_spec_coupon(dgv_Analysis, sheet);
                                break;

                            case "IQC_PSA_PEELING_COUPON":
                                check_spec_coupon(dgv_Analysis, sheet);
                                break;

                            case "LINER_PEEL_TEST_ON_PRODUCT":
                                check_spec_onproduct(dgv_Analysis, sheet);
                                break;

                            case "PSA_PEEL_TEST_ON_PRODUCT":
                                check_spec_onproduct(dgv_Analysis, sheet);
                                break;

                            case "CROSS_SECTION":
                                Check_spec_crossection(dgv_Analysis, sheet);
                                break;

                            case "GAP_CONNECTOR":
                                try
                                {

                                    Check_spec_GAP(dgv_Analysis, sheet);
                                }
                                catch
                                {
                                    MessageBox.Show("Error Check spec");
                                }
                                break;
                        }
                    }

                    Check_Alldata(dgv_Analysis);
                    bool chk_ = true;
                    for (int i = 0; i < dgv_Analysis.Rows.Count; i++)
                    {
                        if (dgv_Analysis.Rows[i].Cells["Data"].Style.BackColor == Color.Red)
                        {
                            chk_ = false;
                            break;
                        }
                    }

                    if (check_enough_data(dgv_Analysis) && chk_)
                    {
                        lbl_judge.Text = "OK";
                        lbl_judge.BackColor = Color.Green;
                    }
                    else
                    {
                        lbl_judge.Text = "NG";
                        lbl_judge.BackColor = Color.Red;
                    }

                    btnEdit.Visible = true;
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Không có dữ liệu", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Vui lòng điền đầy đủ ItemCode / LotNo / Type",
                    "Thông báo");
            }
        }

        public void Get_ListTable(int col_inx, DataTable myDt, string[] src_arr, ref List<DataTable> src_lst_tbl,
            string tar_item)
        {
            DataRow[] temp_dr;
            if (col_inx < src_arr.Length - 1)
            {
                if (src_arr[col_inx + 1] != "")
                {
                    string[] sel_val = myDt.AsEnumerable().Select(x => x.Field<string>(src_arr[col_inx + 1])).Distinct()
                        .ToArray();
                    if (sel_val.Length != 0)
                    {
                        foreach (string sv in sel_val)
                        {
                            if (sv != null)
                            {
                                DataTable curTbl = myDt.AsEnumerable()
                                    .Where(r => r.Field<string>(src_arr[col_inx + 1]) == sv).CopyToDataTable();
                                Get_ListTable(col_inx + 1, curTbl, src_arr, ref src_lst_tbl, tar_item);
                            }
                        }
                    }
                    else
                    {
                        temp_dr = myDt.AsEnumerable().Where(x => x.Field<string>(tar_item) != null).ToArray();
                        if (temp_dr.Length > 0)
                        {
                            src_lst_tbl.Add(myDt);
                        }
                    }
                }
            }
            else
            {
                temp_dr = myDt.AsEnumerable().Where(x => x.Field<string>(tar_item) != null).ToArray();
                if (temp_dr.Length > 0)
                {
                    src_lst_tbl.Add(myDt);
                }
            }
        }

        public string spec_coupon(string mysheet, string region)
        {
            string str_spec = "";
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet" },
                new string[] { txtItemCode.Text, mysheet });
            DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", filter_str);

            string spec_region = dt_spec.Rows[0]["Location"].ToString().Split('_')[0];
            if (dt_spec.Rows.Count > 0)
            {
                string spec = "";

                if (mysheet == "IQC_LINER_PEELING_COUPON")
                {
                    spec = spec_region.Split('+')[3].Split(';')[0].Split('(')[1].Split(')')[0]
                        .Replace(" ", string.Empty).Split('/')[0].Replace("gf", "");

                    if (spec.Contains("-"))
                    {
                        str_spec = spec_region.Split('+')[3].Split(';')[0];
                    }
                }

                if (mysheet == "IQC_PSA_PEELING_COUPON")
                {
                    spec = spec_region.Split('+')[3].Split(';')[0].Split('(')[1].Split('N')[0]
                        .Replace(" ", string.Empty).Replace("≥", string.Empty).Replace(">", string.Empty)
                        .Replace("<", string.Empty);


                    if (TDMK_Code.IsNumeric(spec))
                    {
                        str_spec = spec_region.Split('+')[3].Split(';')[0];
                    }
                }
            }

            return str_spec;
        }


        public void check_spec_coupon(DataGridView dgv, string mysheet)
        {
            //string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet" }, new string[] { txtItemCode.Text, mysheet });
            DataTable dt_spec = new DataTable();
            //DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, sheet, cb_Type.SelectedItem.ToString() }));
            if (cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, sheet, "MASS" }));
            }
            else if (cb_Type.SelectedItem.ToString() == "LQ" || cb_Type.SelectedItem.ToString() == "NPI")
            {
                dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, sheet, "NPI" }));
            }

            if (dt_spec.Rows.Count > 0)
            {
                int reg = 0;
                List<DataTable> lst_Table = new List<DataTable> { };
                Get_ListTable(-1, (DataTable)dgv.DataSource, new string[] { "Region" }, ref lst_Table, "Data");
                foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                {
                    if (reg < lst_Table.Count)
                    {
                        DataTable dt_region = lst_Table[reg];
                        string type = spec_region.Split('+')[0];
                        if (type == "A")
                        {
                            string spec = "";
                            if (mysheet == "IQC_LINER_PEELING_COUPON")
                            {
                                if (spec_region.Split('+')[3].Contains("Judgement") &&
                                    spec_region.Split('+')[3].Contains("/") && spec_region.Split('+')[3].Contains("("))
                                {
                                    spec = spec_region.Split('+')[3].Split(';')[0].Split('(')[1].Split(')')[0]
                                        .Replace(" ", string.Empty).Split('/')[0].Replace("gf", "");
                                }
                                else if (spec_region.Split('+')[4].Contains("Judgement") &&
                                         spec_region.Split('+')[4].Contains("/") &&
                                         spec_region.Split('+')[4].Contains("("))
                                {
                                    spec = spec_region.Split('+')[4].Split(';')[0].Split('(')[1].Split(')')[0]
                                        .Replace(" ", string.Empty).Split('/')[0].Replace("gf", "");
                                }


                                if (spec.Contains("-"))
                                {
                                    string ll = spec.Split('-')[0];
                                    string ul = spec.Split('-')[1];

                                    if (myCode.IsNumeric(ul) && myCode.IsNumeric(ll))
                                    {
                                        for (int i = 0; i < dt_region.Rows.Count; i++)
                                        {
                                            string val = dt_region.Rows[i]["Data"].ToString();
                                            if (myCode.IsNumeric(val))
                                            {
                                                int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                                if (Double.Parse(val) < Double.Parse(ll) ||
                                                    Double.Parse(val) > Double.Parse(ul))
                                                {
                                                    dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            if (mysheet == "IQC_PSA_PEELING_COUPON")
                            {
                                //spec = spec_region.Split('+')[3].Split(';')[0].Split('(')[1].Split('N')[0].Replace(" ", string.Empty).Replace("≥", string.Empty).Replace(">", string.Empty).Replace("<", string.Empty);
                                if (spec_region.Split('+')[3].Contains("Judgement") &&
                                    spec_region.Split('+')[3].Contains("/") && spec_region.Split('+')[3].Contains("("))
                                {
                                    spec = spec_region.Split('+')[3].Split(';')[0].Split('(')[1].Split('N')[0]
                                        .Replace(" ", string.Empty).Replace("≥", string.Empty)
                                        .Replace(">", string.Empty).Replace("<", string.Empty);
                                }
                                else if (spec_region.Split('+')[4].Contains("Judgement") &&
                                         spec_region.Split('+')[4].Contains("/") &&
                                         spec_region.Split('+')[4].Contains("("))
                                {
                                    spec = spec_region.Split('+')[3].Split(';')[0].Split('(')[1].Split('N')[0]
                                        .Replace(" ", string.Empty).Replace("≥", string.Empty)
                                        .Replace(">", string.Empty).Replace("<", string.Empty);
                                }

                                if (TDMK_Code.IsNumeric(spec))
                                {
                                    for (int i = 0; i < dt_region.Rows.Count; i++)
                                    {
                                        string val = dt_region.Rows[i]["Data"].ToString();
                                        if (myCode.IsNumeric(val))
                                        {
                                            int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                            if (Double.Parse(val) < Double.Parse(spec))
                                            {
                                                dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else if (type == "B")
                        {
                            double UCL = 0;
                            double LCL = 0;
                            try
                            {
                                if (spec_region.Split('+')[6].Split(';')[1] != "")
                                {
                                    UCL = double.Parse(spec_region.Split('+')[6].Split(';')[1]);
                                }

                                if (spec_region.Split('+')[6].Split(';')[2] != "")
                                {
                                    LCL = double.Parse(spec_region.Split('+')[6].Split(';')[2]);
                                }
                            }
                            catch
                            {
                                if (spec_region.Split('+')[5].Split(';')[1] != "")
                                {
                                    UCL = double.Parse(spec_region.Split('+')[5].Split(';')[1]);
                                }

                                if (spec_region.Split('+')[5].Split(';')[2] != "")
                                {
                                    LCL = double.Parse(spec_region.Split('+')[5].Split(';')[2]);
                                }
                            }

                            SortedDictionary<int, string> dic_data = new SortedDictionary<int, string> { };

                            List<double> lst_data = new List<double> { };
                            for (int i = 0; i < dt_region.Rows.Count; i++)
                            {
                                string val = dt_region.Rows[i]["Data"].ToString().Split('_')[1].Replace("Average:", "");
                                if (myCode.IsNumeric(val))
                                {
                                    int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                    dic_data.Add(r, val);
                                    lst_data.Add(double.Parse(val));

                                    if (Double.Parse(val) > UCL && UCL != 0)
                                    {
                                        dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                    }

                                    if (Double.Parse(val) < LCL && LCL != 0)
                                    {
                                        dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                    }
                                }
                            }

                            string r_spec = "";

                            try
                            {
                                r_spec = spec_region.Split('+')[6].Split(';')[0];
                            }
                            catch
                            {
                                r_spec = spec_region.Split('+')[5].Split(';')[0];
                            }


                            if (r_spec != "")
                            {
                                double R = double.Parse(r_spec);
                                double tb = lst_data.ToArray().Average();

                                foreach (int r1 in dic_data.Keys)
                                {
                                    foreach (int r2 in dic_data.Keys)
                                    {
                                        if (r2 < r1)
                                        {
                                            double sub_data = Math.Abs(double.Parse(dic_data[r1]) -
                                                                       double.Parse(dic_data[r2]));

                                            if (sub_data > R && tb != 0)
                                            {
                                                double a1 = Math.Abs(double.Parse(dic_data[r1]) - tb);
                                                double a2 = Math.Abs(double.Parse(dic_data[r2]) - tb);
                                                if (a1 > a2)
                                                {
                                                    dgv.Rows[r1 - 1].Cells["Data"].Style.BackColor = Color.Red;
                                                }
                                                else
                                                {
                                                    dgv.Rows[r2 - 1].Cells["Data"].Style.BackColor = Color.Red;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        reg++;
                    }
                }
            }
        }

        //public void CloseExcel(string xlFilePath)
        //{
        //    workbook.Close(false, xlFilePath, null); // Close the connection to workbook
        //    Marshal.FinalReleaseComObject(workbook); // Release unmanaged object references.
        //    workbook = null;

        //    workbooks.Close();
        //    Marshal.FinalReleaseComObject(workbooks);
        //    workbooks = null;

        //    xlApp.Quit();
        //    Marshal.FinalReleaseComObject(xlApp);
        //    xlApp = null;
        //}

        public void check_spec_onproduct(DataGridView dgv, string mysheet)
        {
            //string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet" }, new string[] { txtItemCode.Text, mysheet });
            DataTable dt_spec = new DataTable();
            // DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, sheet, cb_Type.SelectedItem.ToString() }));
            if (cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, sheet, "MASS" }));
            }
            else if (cb_Type.SelectedItem.ToString() == "LQ" || cb_Type.SelectedItem.ToString() == "NPI")
            {
                dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, sheet, "NPI" }));
            }

            if (dt_spec.Rows.Count > 0)
            {
                int reg = 0;
                List<DataTable> lst_Table = new List<DataTable> { };
                Get_ListTable(-1, (DataTable)dgv.DataSource, new string[] { "Region" }, ref lst_Table, "Data");
                foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                {
                    if (reg < lst_Table.Count)
                    {
                        //DataTable dt_region = lst_Table[reg];
                        DataTable Data_all = (DataTable)dgv.DataSource;
                        List<string> lst_region = Data_all.AsEnumerable().Select(x => x.Field<string>("Region"))
                            .Distinct().ToList();

                        string cpn = "";
                        foreach (string ar in lst_region)
                        {
                            if (spec_region.Replace(" ", "").ToUpper().Replace("COMPONENT", "TAPE")
                                    .Contains(ar.Split('_')[0].Replace(" ", "").ToUpper()) && spec_region
                                    .Replace(" ", "").ToUpper().Replace("COMPONENT", "TAPE")
                                    .Contains(ar.Split('_')[1].Replace(" ", "").ToUpper()))
                            {
                                cpn = ar;
                                break;
                            }
                        }

                        DataView dv = Data_all.AsDataView();
                        string str_filter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { cpn });
                        dv.RowFilter = str_filter;
                        DataTable dt_region = dv.ToTable();


                        string type = spec_region.Split('+')[0];
                        if (type == "A")
                        {
                            string spec_max = spec_region.Split('+')[2].Split(';')[0].Replace(" ", "")
                                .Replace("(gf)", "").Replace("(N)", "");
                            string ll_max = "";
                            string ul_max = "";
                            string ll_ave = "";
                            string ul_ave = "";

                            if (spec_max.Contains("(") && spec_max.Contains(")"))
                            {
                                spec_max = spec_max.Split('(')[1].Split(')')[0].Replace(" ", ")").Replace("gf", "")
                                    .Replace("N", "");
                                if (spec_max.Contains("-"))
                                {
                                    ll_max = spec_max.Split('-')[0];
                                    ul_max = spec_max.Split('-')[1];
                                }
                            }

                            string spec_average = spec_region.Split('+')[4].Split(';')[0].Replace(" ", "")
                                .Replace("(gf)", "").Replace("(N)", "");
                            if (spec_average.Contains("(") && spec_average.Contains(")"))
                            {
                                spec_average = spec_average.Split('(')[1].Split(')')[0].Replace(" ", ")")
                                    .Replace("gf", "").Replace("N", "");
                                if (spec_average.Contains("-"))
                                {
                                    ll_ave = spec_average.Split('-')[0];
                                    ul_ave = spec_average.Split('-')[1];
                                }
                            }

                            if (myCode.IsNumeric(ul_max) && myCode.IsNumeric(ll_max))
                            {
                                for (int i = 0; i < dt_region.Rows.Count; i++)
                                {
                                    string val = dt_region.Rows[i]["Data"].ToString().Split('_')[0].Replace("Max:", "");
                                    if (myCode.IsNumeric(val))
                                    {
                                        int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                        if (Double.Parse(val) < Double.Parse(ll_max) ||
                                            Double.Parse(val) > Double.Parse(ul_max))
                                        {
                                            dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                        }
                                    }
                                }
                            }

                            if (myCode.IsNumeric(ul_ave) && myCode.IsNumeric(ll_ave))
                            {
                                for (int i = 0; i < dt_region.Rows.Count; i++)
                                {
                                    string val = dt_region.Rows[i]["Data"].ToString().Split('_')[1]
                                        .Replace("Average:", "");
                                    if (myCode.IsNumeric(val))
                                    {
                                        int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                        if (Double.Parse(val) < Double.Parse(ll_ave) ||
                                            Double.Parse(val) > Double.Parse(ul_ave))
                                        {
                                            dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                        }
                                    }
                                }
                            }
                        }
                        else if (type == "B")
                        {
                            double UCL = 0;
                            double LCL = 0;
                            try
                            {
                                if (spec_region.Split('+')[6].Split(';')[1] != "")
                                {
                                    UCL = double.Parse(spec_region.Split('+')[6].Split(';')[1]);
                                }

                                if (spec_region.Split('+')[6].Split(';')[2] != "")
                                {
                                    LCL = double.Parse(spec_region.Split('+')[6].Split(';')[2]);
                                }
                            }
                            catch
                            {
                                if (spec_region.Split('+')[5].Split(';')[1] != "")
                                {
                                    UCL = double.Parse(spec_region.Split('+')[5].Split(';')[1]);
                                }

                                if (spec_region.Split('+')[5].Split(';')[2] != "")
                                {
                                    LCL = double.Parse(spec_region.Split('+')[5].Split(';')[2]);
                                }
                            }


                            SortedDictionary<int, string> dic_data = new SortedDictionary<int, string> { };
                            List<double> lst_data = new List<double> { };
                            for (int i = 0; i < dt_region.Rows.Count; i++)
                            {
                                string val = dt_region.Rows[i]["Data"].ToString().Split('_')[1].Replace("Average:", "");
                                if (myCode.IsNumeric(val))
                                {
                                    int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                    dic_data.Add(r, val);
                                    lst_data.Add(double.Parse(val));

                                    if (Double.Parse(val) > UCL && UCL != 0)
                                    {
                                        dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                    }

                                    if (Double.Parse(val) < LCL && LCL != 0)
                                    {
                                        dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                    }
                                }
                            }


                            string r_spec = "";
                            try
                            {
                                r_spec = spec_region.Split('+')[6].Split(';')[0];
                            }
                            catch
                            {
                                r_spec = spec_region.Split('+')[5].Split(';')[0];
                            }


                            if (myCode.IsNumeric(r_spec))
                            {
                                double R = double.Parse(r_spec);
                                double tb = lst_data.ToArray().Average();

                                foreach (int r1 in dic_data.Keys)
                                {
                                    foreach (int r2 in dic_data.Keys)
                                    {
                                        if (r2 < r1)
                                        {
                                            double sub_data = Math.Abs(double.Parse(dic_data[r1]) -
                                                                       double.Parse(dic_data[r2]));

                                            if (sub_data > R && tb != 0)
                                            {
                                                double a1 = Math.Abs(double.Parse(dic_data[r1]) - tb);
                                                double a2 = Math.Abs(double.Parse(dic_data[r2]) - tb);
                                                if (a1 > a2)
                                                {
                                                    dgv.Rows[r1 - 1].Cells["Data"].Style.BackColor = Color.Red;
                                                }
                                                else
                                                {
                                                    dgv.Rows[r2 - 1].Cells["Data"].Style.BackColor = Color.Red;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // } 
                        reg++;
                    }
                }
            }
        }

        public void Check_spec_unmatingpull(DataGridView dgv, string mysheet)
        {
            //string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet" }, new string[] { txtItemCode.Text, mysheet });
            DataTable dt_spec = new DataTable();
            //DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, sheet, cb_Type.SelectedItem.ToString() }));
            if (cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, sheet, "MASS" }));
            }
            else if (cb_Type.SelectedItem.ToString() == "LQ" || cb_Type.SelectedItem.ToString() == "NPI")
            {
                dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, sheet, "NPI" }));
            }

            if (dt_spec.Rows.Count > 0)
            {
                string a = dt_spec.Rows[0]["Location"].ToString();
                //string spec = dt_spec.Rows[0]["Location"].ToString().Split('+')[3].Split(';')[0].Split('(')[1].Split(')')[0].Replace(" ", string.Empty).Replace("N", string.Empty).Replace("≥", string.Empty).Replace(">", string.Empty).Replace("<", string.Empty);
                string spec = "";
                try
                {
                    spec = get_number_spec2(
                        dt_spec.Rows[0]["Location"].ToString().Split('+')[1].Split(';')[0].Split('(')[1].Split(')')[0]
                            .Replace(" ", string.Empty).Replace("N", ""));
                }
                catch
                {
                    spec = "4";
                }

                if (TDMK_Code.IsNumeric(spec))
                {
                    for (int i = 0; i < dgv.Rows.Count; i++)
                    {
                        string val = dgv.Rows[i].Cells["Data"].Value.ToString();
                        if (myCode.IsNumeric(val))
                        {
                            if (Double.Parse(val) < Double.Parse(spec))
                            {
                                dgv.Rows[i].Cells["Data"].Style.BackColor = Color.Red;
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Vui lòng cài đặt spec", "Thông báo");
            }
        }

        //public void Check_spec_comment3(DataGridView dgv, string mysheet)
        //{

        //    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet" }, new string[] { txtItemCode.Text, mysheet });
        //    DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", filter_str);
        //    if (dt_spec.Rows.Count > 0)
        //    {
        //        if (mysheet == "IQC_LINER_PEELING_COUPON" || mysheet == "IQC_PSA_PEELING_COUPON")
        //        {
        //            int reg = 0;
        //            List<DataTable> lst_Table = new List<DataTable> { };
        //            Get_ListTable(-1, (DataTable)dgv.DataSource, new string[] { "Region" }, ref lst_Table, "Data");
        //            foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
        //            {
        //                if (reg < lst_Table.Count)
        //                {
        //                    DataTable dt_region = lst_Table[reg];
        //                    string type = spec_region.Split('+')[0];
        //                    if (type == "A")
        //                    {
        //                        string spec = "";

        //                        if (mysheet == "IQC_LINER_PEELING_COUPON")
        //                        {
        //                            spec = spec_region.Split('+')[3].Split(';')[0].Split('(')[1].Split(')')[0].Replace(" ", string.Empty).Split('/')[0].Replace("gf", "");

        //                        }
        //                        else
        //                        {
        //                            //spec = dt_spec.Rows[0]["Location"].ToString().Split('+')[3].Split(';')[0].Split('(')[1].Split('N')[0].Replace(" ", string.Empty).Replace("≥", string.Empty).Replace(">", string.Empty).Replace("<", string.Empty);
        //                        }

        //                        if (TDMK_Code.IsNumeric(spec))
        //                        {

        //                            for (int i = 0; i < dt_region.Rows.Count; i++)
        //                            {
        //                                string val = dt_region.Rows[i]["Data"].ToString();
        //                                if (myCode.IsNumeric(val))
        //                                {
        //                                    int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
        //                                    if (Double.Parse(val) < Double.Parse(spec))
        //                                    {
        //                                        dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
        //                                    }
        //                                }

        //                            }
        //                        }
        //                    }
        //                    else if (type == "B")
        //                    {
        //                        //    string spec = spec_region.Split('+')[4];
        //                        //    if(spec != "")
        //                        //{
        //                        //}

        //                        double UCL = double.Parse(spec_region.Split('+')[5].Split(';')[1]);
        //                        double LCL = double.Parse(spec_region.Split('+')[5].Split(';')[2]);

        //                        SortedDictionary<int, string> dic_data = new SortedDictionary<int, string> { };
        //                        for (int i = 0; i < dt_region.Rows.Count; i++)
        //                        {
        //                            string val = dt_region.Rows[i]["Data"].ToString();
        //                            if (myCode.IsNumeric(val))
        //                            {
        //                                int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
        //                                dic_data.Add(r, val);
        //                                if (Double.Parse(val) > UCL || Double.Parse(val) < LCL)
        //                                {
        //                                    dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;

        //                                }
        //                            }
        //                        }
        //                        double R = double.Parse(spec_region.Split('+')[5].Split(';')[0]);


        //                        foreach (int r1 in dic_data.Keys)
        //                        {
        //                            foreach (int r2 in dic_data.Keys)
        //                            {
        //                                if (r2 < r1)
        //                                {
        //                                    double sub_data = Math.Abs(double.Parse(dic_data[r1]) - double.Parse(dic_data[r2]));
        //                                    if (sub_data > R)
        //                                    {

        //                                        dgv.Rows[r1 - 1].Cells["Data"].Style.BackColor = Color.Red;
        //                                        dgv.Rows[r2 - 1].Cells["Data"].Style.BackColor = Color.Red;

        //                                    }
        //                                }
        //                            }
        //                        }
        //                    }

        //                    reg++;
        //                }
        //            }


        //        }
        //        else if (!arr_onproduct.Contains(mysheet))
        //        {
        //            string spec = dt_spec.Rows[0]["Location"].ToString().Split('+')[3].Split(';')[0].Split('(')[1].Split(')')[0].Replace(" ", string.Empty).Replace("N", string.Empty).Replace("≥", string.Empty).Replace(">", string.Empty).Replace("<", string.Empty);
        //            if (TDMK_Code.IsNumeric(spec))
        //            {
        //                for (int i = 0; i < dgv.Rows.Count; i++)
        //                {
        //                    string val = dgv.Rows[i].Cells["Data"].Value.ToString();
        //                    if (Double.Parse(val) < Double.Parse(spec))
        //                    {
        //                        for (int k = 0; k < dgv.Columns.Count; k++)
        //                        {
        //                            dgv.Rows[i].Cells[k].Style.BackColor = Color.Red;
        //                        }
        //                    }
        //                }
        //            }

        //        }
        //    }
        //    else
        //    {
        //        MessageBox.Show(new Form { TopMost = true }, "Vui lòng cài đặt spec", "Thông báo");
        //    }
        //}


        public void Check_Alldata_old(DataGridView dgv)
        {
            if (sheet != "CROSS_SECTION" && sheet != "GAP_CONNECTOR")
            {
                List<double> lst_data = new List<double> { };
                SortedDictionary<int, string> dic_data = new SortedDictionary<int, string> { };
                if (dgv.Rows.Count > 0)
                {
                    double tb = 0;
                    if (sheet != "LINER_PEEL_TEST_ON_PRODUCT" && sheet != "PSA_PEEL_TEST_ON_PRODUCT")
                    {
                        for (int i = 0; i < dgv.Rows.Count; i++)
                        {
                            string val = myCode.checkDBNull(dgv.Rows[i].Cells["Data"].Value);
                            if (val != "")
                            {
                                dic_data.Add(i, val);
                                if (myCode.IsNumeric(val))
                                {
                                    lst_data.Add(double.Parse(val));
                                }
                            }
                        }

                        if (lst_data.Count > 0)
                        {
                            tb = lst_data.ToArray().Average();
                        }
                    }


                    double R = 0;
                    switch (sheet)
                    {
                        case "PEEL_TEST":
                            R = 10;
                            break;

                        case "MATING_PULL_TEST":
                            R = 15;
                            break;

                        case "SHEAR_TEST":
                            R = 5;
                            break;

                        case "IQC_UNMATING_PULL_TEST":
                            R = 5;
                            break;

                        case "IQC_LINER_PEELING_COUPON":
                            R = 0;
                            break;

                        case "IQC_PSA_PEELING_COUPON":
                            R = 0;
                            break;

                        case "LINER_PEEL_TEST_ON_PRODUCT":
                            R = 20;
                            break;

                        case "PSA_PEEL_TEST_ON_PRODUCT":
                            R = 20;
                            break;

                        default:
                            break;
                    }

                    if (R != 0)
                    {
                        foreach (int r1 in dic_data.Keys)
                        {
                            foreach (int r2 in dic_data.Keys)
                            {
                                if (r2 > r1)
                                {
                                    double sub_data = 0;
                                    if (sheet == "LINER_PEEL_TEST_ON_PRODUCT" || sheet == "PSA_PEEL_TEST_ON_PRODUCT")
                                    {
                                        sub_data = Math.Abs(
                                            double.Parse(dic_data[r1].Split('_')[1].Replace("Average:", "")) -
                                            double.Parse(dic_data[r2].Split('_')[1].Replace("Average:", "")));
                                    }
                                    else
                                    {
                                        sub_data = Math.Abs(double.Parse(dic_data[r1]) - double.Parse(dic_data[r2]));
                                    }

                                    if (sub_data > R && tb != 0)
                                    {
                                        double a1 = Math.Abs(double.Parse(dic_data[r1]) - tb);
                                        double a2 = Math.Abs(double.Parse(dic_data[r2]) - tb);
                                        if (a1 > a2)
                                        {
                                            dgv.Rows[r1].Cells["Data"].Style.BackColor = Color.Red;
                                        }
                                        else
                                        {
                                            dgv.Rows[r2].Cells["Data"].Style.BackColor = Color.Red;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        public void Check_Alldata(DataGridView dgv)
        {
            if (dgv.DataSource != null)
            {
                if (sheet != "CROSS_SECTION" && sheet != "GAP_CONNECTOR")
                {
                    List<DataTable> lst_Table = new List<DataTable> { };
                    Get_ListTable(-1, (DataTable)dgv.DataSource, new string[] { "Region" }, ref lst_Table, "Data");
                    foreach (DataTable dt_region in lst_Table)
                    {
                        List<double> lst_data = new List<double> { };
                        SortedDictionary<int, string> dic_data = new SortedDictionary<int, string> { };

                        double tb = 0;
                        if (sheet == "LINER_PEEL_TEST_ON_PRODUCT" || sheet == "PSA_PEEL_TEST_ON_PRODUCT")
                        {
                            for (int i = 0; i < dt_region.Rows.Count; i++)
                            {
                                string val = myCode.checkDBNull(dt_region.Rows[i]["Data"]).Split('_')[1]
                                    .Replace("Average:", "");
                                if (myCode.IsNumeric(val))
                                {
                                    dic_data.Add(i, val);
                                    lst_data.Add(double.Parse(val));
                                }
                            }
                        }
                        else
                        {
                            for (int i = 0; i < dt_region.Rows.Count; i++)
                            {
                                string val = myCode.checkDBNull(dt_region.Rows[i]["Data"]);
                                if (val != "")
                                {
                                    dic_data.Add(i, val);
                                    if (myCode.IsNumeric(val))
                                    {
                                        lst_data.Add(double.Parse(val));
                                    }
                                }
                            }
                        }

                        if (lst_data.Count > 0)
                        {
                            tb = lst_data.ToArray().Average();
                        }


                        double R = 0;
                        switch (sheet)
                        {
                            case "PEEL_TEST":
                                R = 10;
                                break;

                            case "MATING_PULL_TEST":
                                R = 15;
                                break;

                            case "SHEAR_TEST":
                                R = 5;
                                break;

                            case "IQC_UNMATING_PULL_TEST":
                                R = 10;
                                break;

                            case "IQC_LINER_PEELING_COUPON":
                                R = 0;
                                break;

                            case "IQC_PSA_PEELING_COUPON":
                                R = 0;
                                break;

                            case "LINER_PEEL_TEST_ON_PRODUCT":
                                R = 20;
                                break;

                            case "PSA_PEEL_TEST_ON_PRODUCT":
                                R = 20;
                                break;

                            default:
                                break;
                        }

                        if (R != 0)
                        {
                            foreach (int r1 in dic_data.Keys)
                            {
                                foreach (int r2 in dic_data.Keys)
                                {
                                    if (r2 > r1)
                                    {
                                        double sub_data = 0;
                                        sub_data = Math.Abs(double.Parse(dic_data[r1]) - double.Parse(dic_data[r2]));


                                        if (sub_data > R && tb != 0)
                                        {
                                            double a1 = Math.Abs(double.Parse(dic_data[r1]) - tb);
                                            double a2 = Math.Abs(double.Parse(dic_data[r2]) - tb);
                                            if (a1 > a2)
                                            {
                                                int ID = int.Parse(dt_region.Rows[r1]["ID"].ToString());
                                                dgv.Rows[ID - 1].Cells["Data"].Style.BackColor = Color.Red;
                                            }
                                            else
                                            {
                                                int ID = int.Parse(dt_region.Rows[r2]["ID"].ToString());
                                                dgv.Rows[ID - 1].Cells["Data"].Style.BackColor = Color.Red;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        public void Check_spec_ShearTest_new(DataGridView dgv)
        {
            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                for (int j = 0; j < dgv.Rows.Count; j++)
                {
                    dgv.Rows[j].Cells[i].Style.BackColor = Color.White;
                }
            }

            DataTable dt_spec = new DataTable();
            // DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, sheet, cb_Type.SelectedItem.ToString() }));
            if (cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, sheet, "MASS" }));
            }
            else if (cb_Type.SelectedItem.ToString() == "LQ" || cb_Type.SelectedItem.ToString() == "NPI")
            {
                dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, sheet, "NPI" }));
            }

            if (dt_spec.Rows.Count > 0)
            {
                int reg = 0;
                List<DataTable> lst_Table = new List<DataTable> { };
                Get_ListTable(-1, (DataTable)dgv.DataSource, new string[] { "Region" }, ref lst_Table, "Data");
                foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                {
                    if (reg < lst_Table.Count)
                    {
                        DataTable dt_region = lst_Table[reg];
                        string type = spec_region.Split('+')[0];
                        if (type == "A")
                        {
                            //string spec = spec_region.Split('+')[3].Split(';')[0].Split('(')[1].Split(')')[0].Replace(" ", string.Empty).Replace("N", string.Empty).Replace("≥", string.Empty).Replace(">", string.Empty).Replace("<", string.Empty);
                            //string spec = get_number_spec(spec_region.Split('+')[3].Split(';')[0]);
                            string spec =
                                get_number_spec2(spec_region.Split('+')[3].Split(';')[0].Split('(')[1].Split(')')[0]
                                    .Replace(" ", "").Replace("N", string.Empty));
                            if (myCode.IsNumeric(spec))
                            {
                                for (int i = 0; i < dt_region.Rows.Count; i++)
                                {
                                    string val = dt_region.Rows[i]["Data"].ToString();
                                    if (myCode.IsNumeric(val))
                                    {
                                        int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                        if (Double.Parse(spec) >= Double.Parse(val) * 9.81)
                                        {
                                            dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                        }
                                    }
                                }
                            }
                        }
                        else if (type == "B")
                        {
                            double UCL = double.Parse(spec_region.Split('+')[4].Split(';')[1]);
                            double LCL = double.Parse(spec_region.Split('+')[4].Split(';')[2]);


                            SortedDictionary<int, string> dic_data = new SortedDictionary<int, string> { };
                            List<double> lst_data = new List<double> { };
                            for (int i = 0; i < dt_region.Rows.Count; i++)
                            {
                                string val = dt_region.Rows[i]["Data"].ToString();
                                if (myCode.IsNumeric(val))
                                {
                                    int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                    dic_data.Add(r, val);
                                    lst_data.Add(double.Parse(val));
                                    if (Double.Parse(val) > UCL || Double.Parse(val) < LCL)
                                    {
                                        dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                    }
                                }
                            }

                            double R = double.Parse(spec_region.Split('+')[4].Split(';')[0]);

                            double tb = lst_data.Average();

                            foreach (int r1 in dic_data.Keys)
                            {
                                foreach (int r2 in dic_data.Keys)
                                {
                                    if (r2 < r1)
                                    {
                                        double sub_data =
                                            Math.Abs(double.Parse(dic_data[r1]) - double.Parse(dic_data[r2]));
                                        if (sub_data > R && tb != 0)
                                        {
                                            double a1 = Math.Abs(double.Parse(dic_data[r1]) - tb);
                                            double a2 = Math.Abs(double.Parse(dic_data[r2]) - tb);
                                            if (a1 > a2)
                                            {
                                                dgv.Rows[r1 - 1].Cells["Data"].Style.BackColor = Color.Red;
                                            }
                                            else
                                            {
                                                dgv.Rows[r2 - 1].Cells["Data"].Style.BackColor = Color.Red;
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        reg++;
                    }
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Vui lòng cài đặt spec", "Thông báo");
            }
        }


        public Boolean check_cross_mass(string data, string _ucl, string _lcl)
        {
            bool chk = true;
            if (myCode.IsNumeric(data))
            {
                double UCL = 0;
                double LCL = 0;
                if (myCode.IsNumeric(_ucl))
                {
                    UCL = double.Parse(_ucl);
                }

                if (myCode.IsNumeric(_lcl))
                {
                    LCL = double.Parse(_lcl);
                }

                if (Double.Parse(data) > UCL && UCL != 0)
                {
                    chk = false;
                }

                if (Double.Parse(data) < LCL && LCL != 0)
                {
                    chk = false;
                }
            }


            return chk;
        }

        public void Check_spec_crossection(DataGridView dgv, string mysheet)
        {
            try
            {
                if (cbStatus.Text == "Shield b2b" || cbStatus.Text == "Clip")
                {
                    DataTable dtz = (DataTable)dgv.DataSource;
                    List<string> check = CrossSectionService.check_SPEC(dtz, txtItemCode.Text);
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet" },
                new string[] { txtItemCode.Text, mysheet });
            DataTable Data_tbl = (DataTable)dgv.DataSource;
            DataTable dt_spec = new DataTable();
            //DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, sheet, cb_Type.SelectedItem.ToString() }));
            if (cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, sheet, "MASS" }));
            }
            else if (cb_Type.SelectedItem.ToString() == "LQ" || cb_Type.SelectedItem.ToString() == "NPI")
            {
                dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, sheet, "NPI" }));
            }

            if (dt_spec.Rows.Count > 0)
            {
                string type = dt_spec.Rows[0]["Remark"].ToString();
                string spec = dt_spec.Rows[0]["Location"].ToString().Split('+')[0].Split(';')[2];
                if (spec.Contains("<"))
                {
                    spec = spec.Split('<')[1].Replace(" ", string.Empty).Replace("µm", string.Empty)
                        .Replace(")", string.Empty);
                    if (TDMK_Code.IsNumeric(spec))
                    {
                        for (int i = 0; i < Data_tbl.Rows.Count; i++)
                        {
                            string val = Data_tbl.Rows[i]["Data"].ToString();
                            if (val != "")
                            {
                                if (Double.Parse(val.Split(';')[0]) >= Double.Parse(spec))
                                {
                                    dgv.Rows[i].Cells["Data"].Style.BackColor = Color.Red;
                                }

                                if (val.Contains("/"))
                                {
                                    foreach (string data in val.Replace(" ", "").Split('/')[1].Split(';'))
                                    {
                                        if (TDMK_Code.IsNumeric(data))
                                        {
                                            if (Double.Parse(data) > 4.5 || Double.Parse(data) < 1.5)
                                            {
                                                dgv.Rows[i].Cells["Data"].Style.BackColor = Color.Red;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                string[] strArray = dt_spec.Rows[0]["Location"].ToString().Split('_');
                string str4 = strArray[strArray.Length - 2];
                if (str4.Contains("^"))
                {
                    DataTable dataSource2 = (DataTable)dgv.DataSource;
                    int val1 = int.Parse(txt_qty.Text);
                    List<string> list1 = ((IEnumerable<string>)strArray[strArray.Length - 1].TrimEnd(';').Split(';'))
                        .ToList<string>();
                    int index1 = 0;
                    List<string> list2 = dataSource2.AsEnumerable()
                        .Select<DataRow, string>((System.Func<DataRow, string>)(x => x.Field<string>("Region")))
                        .Distinct<string>().ToList<string>();
                    Dictionary<string, SortedDictionary<int, string>> dictionary =
                        new Dictionary<string, SortedDictionary<int, string>>();
                    foreach (string str5 in list2)
                    {
                        DataView dataView = Data_tbl.AsDataView();
                        string str6 = this.TDMK_Code.filter_str(new string[1] { "Region" }, new string[1] { str5 });
                        dataView.RowFilter = str6;
                        DataTable table = dataView.ToTable();
                        if (str5.ToUpper().Contains("NGANG") && !str5.ToUpper().Contains("TRU"))
                        {
                            for (int index2 = 0; index2 < 2; ++index2)
                            {
                                SortedDictionary<int, string> sortedDictionary1 = new SortedDictionary<int, string>();
                                for (int index3 = index2 * val1;
                                     index3 < Math.Min(val1 + index2 * val1, table.Rows.Count);
                                     ++index3)
                                {
                                    string str7 = table.Rows[index3]["Data"].ToString().Replace(" ", "");
                                    if (this.myCode.IsNumeric(str7.Split(';')[0]))
                                        sortedDictionary1.Add(int.Parse(table.Rows[index3]["ID"].ToString()) - 1,
                                            str7.Split(';')[0]);
                                }

                                if (index1 < list1.Count)
                                {
                                    dictionary.Add(list1[index1], sortedDictionary1);
                                    ++index1;
                                }

                                SortedDictionary<int, string> sortedDictionary2 = new SortedDictionary<int, string>();
                                for (int index4 = index2 * val1;
                                     index4 < Math.Min(val1 + index2 * val1, table.Rows.Count);
                                     ++index4)
                                {
                                    string str8 = table.Rows[index4]["Data"].ToString();
                                    if (this.myCode.IsNumeric(str8.Split(';')[1]))
                                        sortedDictionary2.Add(int.Parse(table.Rows[index4]["ID"].ToString()) - 1,
                                            str8.Split(';')[1]);
                                }

                                if (index1 < list1.Count)
                                {
                                    dictionary.Add(list1[index1], sortedDictionary2);
                                    ++index1;
                                }
                            }
                        }
                        else if (!str5.ToUpper().Contains("NGANG"))
                        {
                            for (int index5 = 0; index5 < 6; ++index5)
                            {
                                SortedDictionary<int, string> sortedDictionary = new SortedDictionary<int, string>();
                                for (int index6 = 0; index6 < Math.Min(val1, table.Rows.Count); ++index6)
                                {
                                    string str9 = table.Rows[index6]["Data"].ToString().Replace(" ", "")
                                        .Replace("/", "");
                                    string str10 = str9.Split(';')[index5];
                                    switch (index5)
                                    {
                                        case 3:
                                            str10 = str9.Split(';')[4];
                                            break;
                                        case 4:
                                            str10 = str9.Split(';')[3];
                                            break;
                                    }

                                    if (this.myCode.IsNumeric(str10))
                                        sortedDictionary.Add(int.Parse(table.Rows[index6]["ID"].ToString()) - 1, str10);
                                }

                                if (index1 < list1.Count)
                                {
                                    dictionary.Add(list1[index1], sortedDictionary);
                                    ++index1;
                                }
                            }
                        }
                    }

                    string str11 = str4.Split('#')[1];
                    List<string> stringList1 = new List<string>();
                    List<string> stringList2 = new List<string>();
                    List<string> stringList3 = new List<string>();
                    string str12 = str11;
                    char[] chArray = new char[1] { '^' };
                    foreach (string str13 in str12.Split(chArray))
                    {
                        if (str13.Contains("R"))
                            stringList1 = ((IEnumerable<string>)str13.Remove(str13.LastIndexOf(";")).Replace("R:", "")
                                .Replace(" ", "").Split(';')).ToList<string>();
                        if (str13.Contains("UCL"))
                            stringList2 = ((IEnumerable<string>)str13.Remove(str13.LastIndexOf(";")).Replace("UCL:", "")
                                .Replace(" ", "").Split(';')).ToList<string>();
                        if (str13.Contains("LCL"))
                            stringList3 = ((IEnumerable<string>)str13.Remove(str13.LastIndexOf(";")).Replace("LCL:", "")
                                .Replace(" ", "").Split(';')).ToList<string>();
                    }

                    List<string> list3 = ((IEnumerable<string>)str4.Split('#')[0].TrimEnd('@').Split('@'))
                        .ToList<string>();
                    int index7 = 0;
                    foreach (string str14 in list3)
                    {
                        List<string> list4 = ((IEnumerable<string>)str14.Replace(" ", "").Split(',')).ToList<string>();
                        foreach (KeyValuePair<string, SortedDictionary<int, string>> keyValuePair1 in dictionary)
                        {
                            if (list4.Contains(keyValuePair1.Key))
                            {
                                string _r = stringList1[index7];
                                string _ucl = stringList2[index7];
                                string _lcl = stringList3[index7];
                                foreach (KeyValuePair<int, string> keyValuePair2 in keyValuePair1.Value)
                                {
                                    if (!this.check_cross_mass(keyValuePair2.Value, _ucl, _lcl))
                                        dgv.Rows[keyValuePair2.Key].Cells["Data"].Style.BackColor = Color.Red;
                                }

                                check_R_crosscut(_r, keyValuePair1.Value, dgv);
                            }
                        }

                        ++index7;
                    }
                }
            }
        }

        public void Check_spec_crossection_old(DataGridView dgv, string mysheet)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet" },
                new string[] { txtItemCode.Text, mysheet });
            DataTable Data_tbl = (DataTable)dgv.DataSource;
            DataTable dt_spec = new DataTable();
            //DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, sheet, cb_Type.SelectedItem.ToString() }));
            if (cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, sheet, "MASS" }));
            }
            else if (cb_Type.SelectedItem.ToString() == "LQ" || cb_Type.SelectedItem.ToString() == "NPI")
            {
                dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, sheet, "NPI" }));
            }

            if (dt_spec.Rows.Count > 0)
            {
                string type = dt_spec.Rows[0]["Remark"].ToString();
                string spec = dt_spec.Rows[0]["Location"].ToString().Split('+')[0].Split(';')[2];
                if (spec.Contains("<"))
                {
                    spec = spec.Split('<')[1].Replace(" ", string.Empty).Replace("µm", string.Empty)
                        .Replace(")", string.Empty);
                    if (TDMK_Code.IsNumeric(spec))
                    {
                        for (int i = 0; i < Data_tbl.Rows.Count; i++)
                        {
                            string val = Data_tbl.Rows[i]["Data"].ToString();
                            if (val != "")
                            {
                                if (Double.Parse(val.Split(';')[0]) >= Double.Parse(spec))
                                {
                                    dgv.Rows[i].Cells["Data"].Style.BackColor = Color.Red;
                                }

                                if (val.Contains("/"))
                                {
                                    foreach (string data in val.Replace(" ", "").Split('/')[1].Split(';'))
                                    {
                                        if (TDMK_Code.IsNumeric(data))
                                        {
                                            if (Double.Parse(data) > 4.5 || Double.Parse(data) < 1.5)
                                            {
                                                dgv.Rows[i].Cells["Data"].Style.BackColor = Color.Red;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                string[] a = dt_spec.Rows[0]["Location"].ToString().Split('_');
                string spec_mass = a[a.Length - 1];
                //   arr_sheet_setup = arr_sheet_setup.Remove(arr_sheet_setup.LastIndexOf(','));
                if (spec_mass.Contains("^"))
                {
                    string[] arr_R = new string[4];
                    string[] arr_UCL = new string[4];
                    string[] arr_LCL = new string[4];

                    foreach (string s in spec_mass.Split('^'))
                    {
                        if (s.Contains("R"))
                        {
                            arr_R = s.Remove(s.LastIndexOf(";")).Replace("R:", "").Replace(" ", "").Split(';');
                        }

                        if (s.Contains("UCL"))
                        {
                            arr_UCL = s.Remove(s.LastIndexOf(";")).Replace("UCL:", "").Replace(" ", "").Split(';');
                        }

                        if (s.Contains("LCL"))
                        {
                            arr_LCL = s.Remove(s.LastIndexOf(";")).Replace("LCL:", "").Replace(" ", "").Split(';');
                        }
                    }


                    SortedDictionary<int, string> dic_data0 = new SortedDictionary<int, string> { };
                    SortedDictionary<int, string> dic_data1 = new SortedDictionary<int, string> { };
                    SortedDictionary<int, string> dic_data2 = new SortedDictionary<int, string> { };
                    SortedDictionary<int, string> dic_data3 = new SortedDictionary<int, string> { };
                    SortedDictionary<int, string> dic_data4 = new SortedDictionary<int, string> { };
                    SortedDictionary<int, string> dic_data5 = new SortedDictionary<int, string> { };


                    for (int i = 0; i < Data_tbl.Rows.Count; i++)
                    {
                        bool chk = true;
                        string val = Data_tbl.Rows[i]["Data"].ToString();
                        string region = Data_tbl.Rows[i]["Region"].ToString();
                        string data0 = val.Split(';')[0];
                        dic_data0.Add(i, data0);
                        if (!check_cross_mass(data0, arr_UCL[0], arr_LCL[0]))
                        {
                            chk = false;
                            goto lbl_highline;
                        }

                        string data1 = val.Split(';')[1];
                        dic_data1.Add(i, data1);
                        if (!check_cross_mass(data1, arr_UCL[1], arr_LCL[1]))
                        {
                            chk = false;
                            goto lbl_highline;
                        }

                        if (val.Contains("/"))
                        {
                            string data2 = val.Split('/')[1].Split(';')[0];
                            dic_data2.Add(i, data2);
                            string data3 = val.Split('/')[1].Split(';')[2];
                            dic_data3.Add(i, data3);
                            string data4 = val.Split('/')[1].Split(';')[1];
                            dic_data4.Add(i, data4);
                            string data5 = val.Split('/')[1].Split(';')[3];
                            dic_data5.Add(i, data5);
                            if (!check_cross_mass(data2, arr_UCL[2], arr_LCL[2]) ||
                                !check_cross_mass(data3, arr_UCL[2], arr_LCL[2]))
                            {
                                chk = false;
                                goto lbl_highline;
                            }

                            if (!check_cross_mass(data4, arr_UCL[3], arr_LCL[3]) ||
                                !check_cross_mass(data5, arr_UCL[3], arr_LCL[3]))
                            {
                                chk = false;
                                goto lbl_highline;
                            }
                        }

                        lbl_highline:

                        if (!chk)
                        {
                            dgv.Rows[i].Cells["Data"].Style.BackColor = Color.Red;
                        }
                    }

                    check_R_crosscut(arr_R[0], dic_data0, dgv);
                    check_R_crosscut(arr_R[1], dic_data1, dgv);
                    check_R_crosscut(arr_R[2], dic_data2, dgv);
                    check_R_crosscut(arr_R[2], dic_data3, dgv);
                    check_R_crosscut(arr_R[3], dic_data4, dgv);
                    check_R_crosscut(arr_R[3], dic_data5, dgv);
                }
            }
        }

        public void check_R_crosscut(string _r, SortedDictionary<int, string> dic_data, DataGridView dgv)
        {
            if (myCode.IsNumeric(_r))
            {
                List<double> lst_data = new List<double>();
                foreach (string val in dic_data.Values)
                {
                    lst_data.Add(double.Parse(val));
                }

                double R = double.Parse(_r);
                double tb = lst_data.ToArray().Average();

                foreach (int r1 in dic_data.Keys)
                {
                    foreach (int r2 in dic_data.Keys)
                    {
                        if (r2 < r1)
                        {
                            double sub_data = Math.Abs(double.Parse(dic_data[r1]) - double.Parse(dic_data[r2]));

                            if (sub_data > R && tb != 0)
                            {
                                double a1 = Math.Abs(double.Parse(dic_data[r1]) - tb);
                                double a2 = Math.Abs(double.Parse(dic_data[r2]) - tb);
                                if (a1 > a2)
                                {
                                    dgv.Rows[r1].Cells["Data"].Style.BackColor = Color.Red;
                                }
                                else
                                {
                                    dgv.Rows[r2].Cells["Data"].Style.BackColor = Color.Red;
                                }
                            }
                        }
                    }
                }
            }
        }

        public string get_number_spec(string str_in)
        {
            string v = "";
            char[] ch_arr = str_in.ToCharArray();
            for (int t = 0; t < ch_arr.Length; t++)
            {
                string a = ch_arr[t].ToString();
                if (myCode.IsNumeric(a))
                {
                    v += a;
                }
            }

            return v;
        }

        public string get_number_spec2(string str_in)
        {
            List<char> lst_csplit = new List<char>() { '>', '<', '≥', '≤' };
            string v = "";
            foreach (char c in lst_csplit)
            {
                if (str_in.Contains(c.ToString()))
                {
                    v = str_in.Split(c)[1];
                    break;
                }
            }

            return v;
        }

        public void Check_spec_Peel_Pull(DataGridView dgv, string mysheet)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet" },
                new string[] { txtItemCode.Text, mysheet });
            DataTable dt_spec = new DataTable();
            //DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, mysheet, cb_Type.SelectedItem.ToString() }));
            if (cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, mysheet, "MASS" }));
            }
            else if (cb_Type.SelectedItem.ToString() == "LQ" || cb_Type.SelectedItem.ToString() == "NPI")
            {
                dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, mysheet, "NPI" }));
            }

            if (dt_spec.Rows.Count > 0)
            {
                int reg = 0;
                List<DataTable> lst_Table = new List<DataTable> { };
                Get_ListTable(-1, (DataTable)dgv.DataSource, new string[] { "Region" }, ref lst_Table, "Data");
                foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                {
                    if (reg < lst_Table.Count)
                    {
                        DataTable dt_region = lst_Table[reg];
                        string type = spec_region.Split('+')[0];
                        if (type == "A")
                        {
                            //string spec = spec_region.Split('+')[3].Split(';')[0].Split('(')[1].Split(')')[0].Replace(" ", string.Empty).Replace("N", string.Empty).Replace("≥", string.Empty).Replace(">", string.Empty).Replace("<", string.Empty);
                            //string spec = get_number_spec(spec_region.Split('+')[3].Split(';')[0]);
                            string spec =
                                get_number_spec2(spec_region.Split('+')[3].Split(';')[0].Split('(')[1].Split(')')[0]
                                    .Replace(" ", string.Empty).Replace("N", ""));
                            if (TDMK_Code.IsNumeric(spec))
                            {
                                for (int i = 0; i < dt_region.Rows.Count; i++)
                                {
                                    string val = dt_region.Rows[i]["Data"].ToString();
                                    if (myCode.IsNumeric(val))
                                    {
                                        int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                        if (Double.Parse(val) < Double.Parse(spec))
                                        {
                                            dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                        }
                                    }
                                }
                            }
                        }
                        else if (type == "B")
                        {
                            string spec = spec_region.Split('+')[4];
                            List<double> lst_data = new List<double> { };

                            double UCL = 0;
                            double LCL = 0;

                            if (myCode.IsNumeric(spec_region.Split('+')[5].Split(';')[1]))
                            {
                                UCL = double.Parse(spec_region.Split('+')[5].Split(';')[1]);
                            }

                            if (myCode.IsNumeric(spec_region.Split('+')[5].Split(';')[2]))
                            {
                                LCL = double.Parse(spec_region.Split('+')[5].Split(';')[2]);
                            }

                            SortedDictionary<int, string> dic_data = new SortedDictionary<int, string> { };
                            for (int i = 0; i < dt_region.Rows.Count; i++)
                            {
                                string val = dt_region.Rows[i]["Data"].ToString();
                                if (myCode.IsNumeric(val))
                                {
                                    int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                    dic_data.Add(r, val);
                                    lst_data.Add(double.Parse(val));
                                    if (myCode.IsNumeric(spec))
                                    {
                                        if (Double.Parse(val) < Double.Parse(spec))
                                        {
                                            dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                        }
                                    }

                                    if ((UCL != 0 && Double.Parse(val) > UCL) || (LCL != 0 && Double.Parse(val) < LCL))
                                    {
                                        dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;
                                    }
                                }
                            }

                            if (myCode.IsNumeric(spec_region.Split('+')[5].Split(';')[0]))
                            {
                                double R = double.Parse(spec_region.Split('+')[5].Split(';')[0]);
                                double tb = lst_data.ToArray().Average();

                                foreach (int r1 in dic_data.Keys)
                                {
                                    foreach (int r2 in dic_data.Keys)
                                    {
                                        if (r2 < r1)
                                        {
                                            double sub_data = Math.Abs(double.Parse(dic_data[r1]) -
                                                                       double.Parse(dic_data[r2]));

                                            if (sub_data > R && tb != 0)
                                            {
                                                double a1 = Math.Abs(double.Parse(dic_data[r1]) - tb);
                                                double a2 = Math.Abs(double.Parse(dic_data[r2]) - tb);
                                                if (a1 > a2)
                                                {
                                                    dgv.Rows[r1 - 1].Cells["Data"].Style.BackColor = Color.Red;
                                                }
                                                else
                                                {
                                                    dgv.Rows[r2 - 1].Cells["Data"].Style.BackColor = Color.Red;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        reg++;
                    }
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Vui lòng cài đặt spec", "Thông báo");
            }
        }

        public void Check_spec_GAP(DataGridView dgv, string mysheet)
        {
            try
            {
                if (cbStatus.Text == "Shield b2b")
                {
                    DataTable dtz = (DataTable)dgv.DataSource;
                    List<string> check = GAPConnectorService.check_SPEC(dtz, txtItemCode.Text);
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            DataTable dt_spec = new DataTable();
            //DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, sheet, cb_Type.SelectedItem.ToString() }));
            if (cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, sheet, "MASS" }));
            }
            else if (cb_Type.SelectedItem.ToString() == "LQ" || cb_Type.SelectedItem.ToString() == "NPI")
            {
                dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, sheet, "NPI" }));
            }

            string str_filter = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" },
                new string[] { txtItemCode.Text, txtLotNo.Text });
            if (dt_spec.Rows.Count > 0)
            {
                for (int i = 0; i < dgv.Rows.Count; i++)
                {
                    string val = myCode.checkDBNull(dgv.Rows[i].Cells["Data"].Value.ToString());
                    if (val.Contains("/"))
                    {
                        //if (val.Split('/')[0].Split(';').Length > 2 && val.Split('/')[1].Split(';').Length > 2)
                        //{
                        //    if (myCode.IsNumeric(val.Split('/')[0].Split(';')[2]) && myCode.IsNumeric(val.Split('/')[1].Split(';')[2]))
                        //    {
                        //        double data1 = double.Parse(val.Split('/')[0].Split(';')[2]);
                        //        double data2 = double.Parse(val.Split('/')[1].Split(';')[2]);

                        //        double sub_data = Math.Abs(data1 - data2);
                        //        if (sub_data > 30)
                        //        {
                        //            dgv.Rows[i].Cells["Data"].Style.BackColor = Color.Red;
                        //            break;
                        //        }
                        //        else if (val.Split('/')[1].Split(';').Length > 3)
                        //        {
                        //            if (myCode.IsNumeric(val.Split('/')[0].Split(';')[3]) && myCode.IsNumeric(val.Split('/')[1].Split(';')[3]))
                        //            {
                        //                double data1_1 = double.Parse(val.Split('/')[0].Split(';')[3]);
                        //                double data2_2 = double.Parse(val.Split('/')[1].Split(';')[3]);

                        //                double sub_data_ = Math.Abs(data1_1 - data2_2);
                        //                if (sub_data_ > 30)
                        //                {
                        //                    dgv.Rows[i].Cells["Data"].Style.BackColor = Color.Red;
                        //                    break;
                        //                }

                        //            }
                        //        }

                        //    }
                        //}


                        if (myCode.IsNumeric(val.Split('/')[0].TrimEnd(' ', ';').Split(';').Last()) &&
                            myCode.IsNumeric(val.Split('/')[1].TrimEnd(' ', ';').Split(';').Last()))
                        {
                            double data1 = double.Parse(val.Split('/')[0].TrimEnd(' ', ';').Split(';').Last());
                            double data2 = double.Parse(val.Split('/')[1].TrimEnd(' ', ';').Split(';').Last());

                            double sub_data = Math.Abs(data1 - data2);
                            if (sub_data > 30)
                            {
                                dgv.Rows[i].Cells["Data"].Style.BackColor = Color.Red;
                                break;
                            }
                        }
                    }
                }

                string[] region = dt_spec.Rows[0]["Location"].ToString().Split('+')[1].Split('_');
                DataTable Data_tbl = (DataTable)dgv.DataSource;

                List<string> lst_region_db = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).ToList();

                foreach (string reg in region)
                {
                    if (reg != "")
                    {
                        int code = 1;
                        foreach (string pos in reg.Split('^'))
                        {
                            if (pos != "")
                            {
                                int row_begin = int.Parse(pos.Split(';')[1]);
                                int count_r_offset = int.Parse(pos.Split(';')[2]);
                                string spec_b = pos.Split(';')[3];
                                if (spec_b.Contains("<") && spec_b.Contains("µm"))
                                {
                                    if (myCode.IsNumeric(spec_b.Split('<')[1].Replace(" ", "").Replace("µm", "")
                                            .Replace(")", "")))
                                    {
                                        double spec = double.Parse(spec_b.Split('<')[1].Replace(" ", "")
                                            .Replace("µm", "").Replace(")", ""));

                                        string find_text = "";
                                        if (reg.Split('^').Length == 3)
                                        {
                                            find_text = code.ToString();
                                        }

                                        string filter_reg = "";

                                        if (pos.Split(';')[0] == "IOPIN")
                                        {
                                            foreach (string i in lst_region_db)
                                            {
                                                if (i.ToUpper().Contains("DOC") && i.Contains(find_text))
                                                {
                                                    filter_reg = i;
                                                    break;
                                                }
                                            }
                                        }
                                        else if (pos.Split(';')[0] == "RIGHT" || pos.Split(';')[0] == "LEFT")
                                        {
                                            foreach (string i in lst_region_db)
                                            {
                                                if (i.ToUpper().Contains("TRU") && i.Contains(find_text))
                                                {
                                                    filter_reg = i;
                                                    break;
                                                }
                                            }
                                        }

                                        DataView dv = Data_tbl.AsDataView();
                                        dv.RowFilter = TDMK_Code.filter_str(
                                            new string[] { "ItemCode", "LotNo", "Region" },
                                            new string[] { txtItemCode.Text, txtLotNo.Text, filter_reg });
                                        int min = new int[] { dv.Count, dgv.Rows.Count - 1 }.Min();

                                        // List<double> lst_1 = new List<double> { };
                                        // List<double> lst_2 = new List<double> { };


                                        for (int inx = 0; inx < min; inx++)
                                        {
                                            DataRow dr = dv[inx].Row;
                                            int r_x = Data_tbl.Rows.IndexOf(dr);

                                            string val = myCode.checkDBNull(Data_tbl.Rows[r_x]["Data"]);
                                            if (val.Contains("/"))
                                            {
                                                if (myCode.IsNumeric(val.Split('/')[0].Split(';')[1]) &&
                                                    myCode.IsNumeric(val.Split('/')[1].Split(';')[1]))
                                                {
                                                    int min_r = new int[]
                                                        { count_r_offset, val.Split('/')[0].Split(';').Length }.Min();
                                                    //lst_1.Add(double.Parse(val.Split('/')[0].Split(';')[min_r - 1]));
                                                    //lst_2.Add(double.Parse(val.Split('/')[1].Split(';')[min_r - 1]));

                                                    double data1 = double.Parse(val.Split('/')[0].Split(';')[1]);
                                                    double data2 = double.Parse(val.Split('/')[1].Split(';')[1]);

                                                    if (data1 > spec || data2 > spec)
                                                    {
                                                        dgv.Rows[r_x].Cells["Data"].Style.BackColor = Color.Red;
                                                    }
                                                }
                                            }
                                        }

                                        //if (lst_1.Max() - lst_1.Min() > 30 || lst_2.Max() - lst_2.Min() > 30)
                                        //{
                                        //    for (int inx = 0; inx < min; inx++)
                                        //    {
                                        //        DataRow dr = dv[inx].Row;
                                        //        int r_x = Data_tbl.Rows.IndexOf(dr);
                                        //        dgv.Rows[r_x].Cells["Data"].Style.BackColor = Color.Red;
                                        //    } 
                                        //    if (pos.Split(';')[0] == "IOPIN")
                                        //    {
                                        //        msg += "GAP DOC ;";
                                        //    }
                                        //    else if (pos.Split(';')[0] == "RIGHT")
                                        //    {
                                        //        msg += "GAP TRU PHAI ;";
                                        //    }
                                        //    else if (pos.Split(';')[0] == "LEFT")
                                        //    {
                                        //        msg += "GAP TRU TRAI ;";
                                        //    }
                                        //}
                                    }
                                }
                            }

                            code++;
                        }
                    }
                }


                //if (msg != "" && dgv == dgv_logfile)
                //{
                //    MessageBox.Show(new Form { TopMost = true }, msg.TrimEnd(';') + " NG : Max - Min > 30");
                //}
            }
        }

        public Boolean check_path()
        {
            if (System.IO.File.Exists(txtLogfile.Text) || System.IO.Directory.Exists(txtLogfile.Text))
            {
                string f_path = txtLogfile.Text.Replace(" ", "").Replace("-", "").Replace("_", "").ToUpper();
                bool chk = true;

                if (chk)
                {
                    if (sheet == "IQC_UNMATING_PULL_TEST" || sheet.Contains("COUPON"))
                    {
                        chk = true;
                    }
                    else
                    {
                        string f_name = Path.GetFileNameWithoutExtension(txtLogfile.Text)
                            .TrimEnd(new char[] { ',', '.', ' ' });
                        bool chk_name = false;

                        if (f_name.Split('-')[1].Trim() == txtItemCode.Text)
                        {
                            if (txtLotNo.Text.Contains("-") && f_name.Split('_')[0].Split('-').Length == 4)
                            {
                                if (Convert.ToDecimal(f_name.Split('-')[2]) ==
                                    Convert.ToDecimal(txtLotNo.Text.ToString().Split('-')[0]) &&
                                    Convert.ToDecimal(f_name.Split('-')[3].Split('_')[0]) ==
                                    Convert.ToDecimal(txtLotNo.Text.ToString().Split('-')[1]))
                                    chk_name = true;
                            }
                            else if (!txtLotNo.Text.Contains("-") && f_name.Split('_')[0].Split('-').Length == 3)
                            {
                                if (Convert.ToDecimal(f_name.Split('-')[2].Split('_')[0].Trim()) ==
                                    Convert.ToDecimal(txtLotNo.Text.ToString()))
                                {
                                    chk_name = true;
                                }
                            }
                        }

                        chk = chk_name;
                    }
                }

                return chk;
            }
            else
            {
                return false;
            }
        }

        public Boolean check_path_()
        {
            if (System.IO.File.Exists(txtLogfile.Text) || System.IO.Directory.Exists(txtLogfile.Text))
            {
                bool chk = true;
                if (chk)
                {
                    if (sheet == "IQC_UNMATING_PULL_TEST" || sheet.Contains("COUPON"))
                    {
                        chk = true;
                    }
                    else
                    {
                        string f_name = Path.GetFileNameWithoutExtension(txtLogfile.Text)
                            .TrimEnd(new char[] { ',', '.', ' ' }).Split('_')[0];

                        bool chk_name = false;

                        if (f_name.Contains(txtItemCode.Text) && f_name.Split('-').Length > 2)
                        {
                            if (myCode.IsNumeric(f_name.Split('-')[2].Trim()) &&
                                myCode.IsNumeric(txtLotNo.Text.ToString()))
                            {
                                if (f_name.Split('-')[1].Trim() == txtItemCode.Text && !txtLotNo.Text.Contains("-") &&
                                    Convert.ToDecimal(f_name.Split('-')[2].Trim()) ==
                                    Convert.ToDecimal(txtLotNo.Text.ToString()))
                                {
                                    chk_name = true;
                                }
                                else if (f_name.Split('-')[1].Trim() == txtItemCode.Text &&
                                         txtLotNo.Text.Contains("-") &&
                                         Convert.ToDecimal(f_name.Split('-')[2].Trim()) ==
                                         Convert.ToDecimal(txtLotNo.Text.ToString().Split('-')[0]) &&
                                         Convert.ToDecimal(f_name.Split('-')[3].Trim()) ==
                                         Convert.ToDecimal(txtLotNo.Text.ToString().Split('-')[1]))
                                {
                                    chk_name = true;
                                }
                            }
                        }

                        chk = chk_name;
                    }
                }

                return chk;
            }
            else
            {
                return false;
            }
        }

        public Boolean check_path_update()
        {
            if (System.IO.File.Exists(txtLogfile.Text) || System.IO.Directory.Exists(txtLogfile.Text))
            {
                bool chk = true;
                if (chk)
                {
                    if (sheet == "IQC_UNMATING_PULL_TEST" || sheet.Contains("COUPON"))
                    {
                        chk = true;
                    }
                    else if ((sheet == "CROSS_SECTION" || sheet == "GAP_CONNECTOR") && (cbStatus.Text == "Shield b2b" || cbStatus.Text == "Clip"))
                    {
                        string[] f_name =
                            Path.GetFileNameWithoutExtension(txtLogfile.Text).TrimEnd(new char[] { ',', '.', ' ' })
                                .Split('-');
                        if (txtItemCode.Text.ToString() == f_name[0].Trim() &&
                            txtLotNo.Text.ToString() == f_name[1].Trim())
                        {
                            chk = true;
                        }
                    }
                    else
                    {
                        string f_name =
                            Path.GetFileNameWithoutExtension(txtLogfile.Text).TrimEnd(new char[] { ',', '.', ' ' })
                                .Split('_')[0].Replace(" ", "");
                        bool chk_name = false;

                        if (f_name.Contains(txtItemCode.Text) && f_name.Split('-').Length > 2)
                        {
                            //if (myCode.IsNumeric(f_name.Split('-')[2].Trim()) && myCode.IsNumeric(txtLotNo.Text.ToString()))
                            //{
                            if (f_name.Split('-')[1].Trim() == txtItemCode.Text && !txtLotNo.Text.Contains("-") &&
                                Convert.ToDecimal(f_name.Split('-')[2].Trim()) ==
                                Convert.ToDecimal(txtLotNo.Text.ToString()))
                            {
                                chk_name = true;
                            }
                            else if (f_name.Split('-')[1].Trim() == txtItemCode.Text && txtLotNo.Text.Contains("-"))
                            {
                                if (f_name.Split('-')[2].Contains("(") && f_name.Split('-')[2].Contains(")"))
                                {
                                    if (Convert.ToDecimal(f_name.Split('-')[2].Split('(')[0].Trim()) ==
                                        Convert.ToDecimal(txtLotNo.Text.ToString().Split('-')[0]) &&
                                        Convert.ToDecimal(f_name.Split('-')[2].Split('(')[1].Split(')')[0].Trim()) ==
                                        Convert.ToDecimal(txtLotNo.Text.ToString().Split('-')[1]))
                                    {
                                        chk_name = true;
                                    }
                                }
                                else if (f_name.Split('-').Length > 3)
                                {
                                    if (Convert.ToDecimal(f_name.Split('-')[2].Trim()) ==
                                        Convert.ToDecimal(txtLotNo.Text.ToString().Split('-')[0]) &&
                                        myCode.IsNumeric(f_name.Split('-')[3]) &&
                                        Convert.ToDecimal(f_name.Split('-')[3]) ==
                                        Convert.ToDecimal(txtLotNo.Text.ToString().Split('-')[1]))
                                    {
                                        chk_name = true;
                                    }
                                }
                            }
                            //}
                        }

                        chk = chk_name;
                    }
                }

                return chk;
            }
            else
            {
                return false;
            }
        }

        public Boolean check_path_new()
        {
            string f_path = Path.GetFileNameWithoutExtension(txtLogfile.Text).Replace(" ", "").Replace("-", "")
                .Replace("_", "").ToUpper();
            bool chk = true;
            switch (sheet)
            {
                case "PEEL_TEST":
                    if (!f_path.Contains("PEEL") && !f_path.Contains("PULL"))
                        chk = false;
                    break;

                case "MATING_PULL_TEST":
                    if (!f_path.Contains("PEEL") && !f_path.Contains("PULL"))
                        chk = false;
                    break;

                case "SHEAR_TEST":
                    if (!f_path.Contains("DAYLINHKIEN"))
                        chk = false;
                    break;

                case "IQC_UNMATING_PULL_TEST":
                    //if (!f_path.Contains("PEEL") && !f_path.Contains("PULL"))
                    //    chk = false;
                    break;

                case "IQC_LINER_PEELING_COUPON":
                    if (!f_path.Contains("COUPON"))
                        chk = false;
                    break;

                case "IQC_PSA_PEELING_COUPON":
                    if (!f_path.Contains("COUPON"))
                        chk = false;
                    break;

                case "LINER_PEEL_TEST_ON_PRODUCT":
                    if (!f_path.Contains("DIECUT"))
                        chk = false;
                    break;

                case "PSA_PEEL_TEST_ON_PRODUCT":
                    if (!f_path.Contains("DIECUT"))
                        chk = false;
                    break;

                case "CROSS_SECTION":
                    if (!f_path.Contains("CROSSCUT"))
                        chk = false;
                    break;

                case "GAP_CONNECTOR":
                    if (!f_path.Contains("CROSSCUT"))
                        chk = false;
                    break;
            }

            return chk;
        }

        public myExcel.Range find_cell(myExcel.Worksheet ws, string find_item)
        {
            myExcel.Range cur_cell = null;
            if (find_item == "IPQC")
                find_item = "SMT";

            for (int j = 1; j < 14; j++)
            {
                for (int i = 1; i < 14; i++)
                {
                    if (myCode.checkDBNull(ws.Cells[i, j].Value).ToUpper().Replace(" ", "").Contains(find_item))
                    {
                        cur_cell = ws.Cells[i, j];
                        return cur_cell;
                    }
                }
            }

            return cur_cell;
        }


        public void export_info_mass_old(string logfile_path, myExcel.Worksheet ws, string itemCode, string lotNo)
        {
            string f_name = Path.GetFileName(logfile_path).Replace(" ", "");
            string[] arr_info = f_name.Split('_')[0].Split('-');
            string[] arr_item = new string[6] { "ITEMNAME:", "ITEMCODE:", "LOTNO:", "LINE:", "SHIFT/CA:", "DATE:" };
            Dictionary<string, string> dic_item = new Dictionary<string, string>();

            for (int indx = 0; indx < arr_info.Length; indx++)
            {
                if (indx < arr_item.Length)
                {
                    if (arr_item[indx] == "DATE:")
                    {
                        if (indx + 1 < arr_info.Length)
                            dic_item.Add(arr_item[indx], arr_info[indx] + "-" + arr_info[indx + 1]);
                    }
                    else
                    {
                        dic_item.Add(arr_item[indx], arr_info[indx]);
                    }
                }
            }

            foreach (var val in dic_item)
            {
                myExcel.Range cur_rgn = find_cell(ws, val.Key);
                if (cur_rgn != null)
                {
                    if (val.Key == "LOTNO:")
                    {
                        cur_rgn.Offset[0, 1].FormulaR1C1 = "'" + val.Value;
                    }
                    else
                    {
                        cur_rgn.Offset[0, 1].Value = val.Value;
                    }
                }
            }

            DataTable Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, sheet,
                TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { itemCode, lotNo }));
            string worker = Data_tbl.Rows[0]["Operator"].ToString();
            string leader = txtOperator.Text;
            string judgement = lbl_judge.Text;

            myExcel.Range cur_rgn_worker = find_cell(ws, "WORKER");
            if (cur_rgn_worker != null)
                cur_rgn_worker.Offset[1, 0].Value = worker;

            myExcel.Range cur_rgn_leader = find_cell(ws, "LEADER");
            if (cur_rgn_leader != null)
                cur_rgn_leader.Offset[1, 0].Value = leader;

            //myExcel.Range cur_rgn_judge = find_cell(ws, "JUDGEMENT");
            //if (cur_rgn_judge != null)
            //{
            //    cur_rgn_judge.Offset[1, 0].Value = judgement;
            //    ws.Cells[1, 1].Value = judgement;

            //}
        }

        public void export_info_mass(myExcel.Worksheet ws, List<string> lst_infor)
        {
            string[] arr_item = new string[6] { "ITEMNAME:", "ITEMCODE:", "LOTNO:", "LINE:", "SHIFT/CA:", "DATE:" };
            Dictionary<string, string> dic_item = new Dictionary<string, string>();

            int i = 0;
            foreach (string item in arr_item)
            {
                dic_item.Add(item, lst_infor[i]);
                i++;
            }

            if (!sheet.Contains("COUPON") && !sheet.Contains("PRODUCT"))
            {
                foreach (var val in dic_item)
                {
                    myExcel.Range cur_rgn = find_cell(ws, val.Key);
                    if (cur_rgn != null)
                    {
                        if (val.Key == "LOTNO:")
                        {
                            cur_rgn.Offset[0, 1].FormulaR1C1 = "'" + val.Value;
                        }
                        else
                        {
                            cur_rgn.Offset[0, 1].Value = val.Value;
                        }
                    }
                }
            }


            string leader = txtOperator.Text;

            myExcel.Range cur_rgn_worker = find_cell(ws, "WORKER");
            if (cur_rgn_worker != null)
                cur_rgn_worker.Offset[1, 0].Value = lst_infor[6];

            myExcel.Range cur_rgn_leader = find_cell(ws, "LEADER");
            if (cur_rgn_leader != null)
                cur_rgn_leader.Offset[1, 0].Value = leader;
        }

        public int count_mergcell_crosscut(ExcelRangeBase tar_rgn, ExcelWorksheet tar_wrksht)
        {
            int count = 0;

            int r_count = tar_rgn.End.Row;
            int col_count = tar_rgn.End.Column;
            var idx = tar_wrksht.GetMergeCellId(r_count, col_count);
            string mergedCellAddress = tar_rgn.Address;
            if (idx > 0)
            {
                mergedCellAddress = tar_wrksht.MergedCells[idx - 1];
            }

            count = tar_wrksht.Cells[mergedCellAddress].Rows;

            return count;
        }


        public string get_offset_addr(string start_addr, ExcelWorksheet tar_wrksht, int offset_val, bool left_to_right)
        {
            string result = start_addr;
            ExcelRangeBase tar_rgn = tar_wrksht.Cells[start_addr];
            int inx = 0;
            while (inx < offset_val)
            {
                int r_count = tar_rgn.End.Row;
                int col_count = tar_rgn.End.Column;
                var idx = tar_wrksht.GetMergeCellId(r_count, col_count);
                string mergedCellAddress = tar_rgn.Address;
                if (idx > 0)
                {
                    mergedCellAddress = tar_wrksht.MergedCells[idx - 1];
                }

                int r = tar_wrksht.Cells[mergedCellAddress].Rows;
                int c = tar_wrksht.Cells[mergedCellAddress].Columns;
                if (left_to_right)
                {
                    tar_rgn = tar_rgn.Offset(0, c);
                }
                else
                {
                    tar_rgn = tar_rgn.Offset(r, 0);
                }

                inx++;
            }

            result = tar_rgn.Address;
            return result;
        }

        private void btn_load_Click(object sender, EventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtOperator.Text != "" && txtLogfile.Text != "" &&
                cb_Type.SelectedIndex != -1)
            {
                dgv_logfile.DataSource = dgv_Analysis.DataSource = null;
                txt_selected.Text = "";
                if ((sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST") &&
                    (!myCode.IsNumeric(txt_pcs_begin.Text) || !myCode.IsNumeric(txt_pcs_end.Text)) &&
                    txt_itemcode_nvl.Text != "" && txt_lotno_nvl.Text != "")
                {
                    MessageBox.Show(new Form() { TopMost = true }, "Chưa nhập pcs cần refer", "Thông báo");
                }
                else
                {
                    if (cb_Type.SelectedItem.ToString() == "NPI")
                    {
                        txt_line.Text = "";
                        txt_ca.Text = "";
                        txt_date.Text = "";
                        txt_worker.Text = "";
                    }

                    string infor = "/" + txt_ItemName.Text + "_" + txt_line.Text + "_" + txt_ca.Text + "_" +
                                   txt_date.Text + "_" + txt_worker.Text + "_" + cb_Type.SelectedItem.ToString();
                    bool check_ = check_path_update();

                    if (check_)
                    {
                        string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Sheet" },
                            new string[] { txtItemCode.Text, txtLotNo.Text, sheet + infor });

                        if (sheet.Contains("UNMATING") || sheet.Contains("COUPON"))
                        {
                            if (txt_itemcode_nvl.Text != "" && txt_lotno_nvl.Text != "")
                            {
                                filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Sheet" },
                                    new string[] { txt_itemcode_nvl.Text, txt_lotno_nvl.Text, sheet + infor });
                            }
                            else
                            {
                                MessageBox.Show(new Form { TopMost = true }, "Chưa nhập đủ ItemCode/lotNo",
                                    "Thông báo");
                                return;
                            }
                        }

                        btnEdit.Visible = false;
                        edit_mode = false;
                        string sheetZ = sheet;
                        if (sheet.Equals("PEEL_TEST") && _PRIME_PEEL_TEST)
                        {
                            sheetZ = "PEEL_TEST_WITHOUT_SUS";
                        }

                        DataTable dt_data = TDMK_Code.Datatable_Filter(sqlcon, sheetZ + "_NAS", filter_str);
                        bool chk = true;
                        if (dt_data.Rows.Count > 0)
                            chk = false;
                        DataTable Data_tbl = new DataTable();
                        DataTable dt_load = new DataTable();
                        string f_folder = txtLogfile.Text;

                        lbl_getdata:
                        if (chk)
                        {
                            SortedDictionary<string, Funtion_SMT.Peeltest_data> dic_result =
                                new SortedDictionary<string, Funtion_SMT.Peeltest_data>();
                            switch (sheet)
                            {
                                case "PEEL_TEST":
                                    if (_PRIME_PEEL_TEST)
                                    {
                                        if (!new PeelTestWOSUSService().checkPeelTest(txtItemCode.Text, txtLotNo.Text))
                                        {
                                            if (MessageBox.Show("Peel test chưa tồn tại có tiếp tục đẩy dữ liệu!",
                                                    "Thông báo!", MessageBoxButtons.YesNo) == DialogResult.No)
                                            {
                                                return;
                                            }
                                        }
                                    }

                                    dt_load = load_data_logfile_Peel_Pull(f_folder,
                                        "PEEL_TEST" + (_PRIME_PEEL_TEST ? "_WITHOUT_SUS" : ""), infor, textBox1.Text);
                                    if (dt_load.Columns.Contains("ProductID"))
                                    {
                                        Data_tbl = dt_load.AsDataView().ToTable(false,
                                            new string[]
                                            {
                                                "ID", "ProductID", "ItemCode", "LotNo", "Sheet", "Region", "Sample",
                                                "Image", "Select_Img", "Graph", "Select_Grp", "Data",
                                                "Mode 1: Solder joint crack", "Mode 2: Pad lift",
                                                "Mode 3: Solder joint lift", "Mode 4: Intermetallic break",
                                                "Mode 5: Component damage", "Mode 6: Component detached",
                                                "Mode 7: Flex torn", "Operator", "Time_Update", "Remark"
                                            });
                                    }
                                    else
                                    {
                                        Data_tbl = dt_load.AsDataView().ToTable(false,
                                            new string[]
                                            {
                                                "ID", "ItemCode", "LotNo", "Sheet", "Region", "Sample", "Image",
                                                "Select_Img", "Graph", "Select_Grp", "Data",
                                                "Mode 1: Solder joint crack", "Mode 2: Pad lift",
                                                "Mode 3: Solder joint lift", "Mode 4: Intermetallic break",
                                                "Mode 5: Component damage", "Mode 6: Component detached",
                                                "Mode 7: Flex torn", "Operator", "Time_Update", "Remark"
                                            });
                                    }

                                    break;

                                case "MATING_PULL_TEST":
                                    dt_load = load_data_logfile_Peel_Pull(f_folder, "MATING_PULL_TEST", infor,
                                        textBox1.Text);
                                    if (dt_load.Columns.Contains("ProductID"))
                                    {
                                        Data_tbl = dt_load.AsDataView().ToTable(false,
                                            new string[]
                                            {
                                                "ID", "ProductID", "ItemCode", "LotNo", "Sheet", "Region", "Sample",
                                                "Image", "Select_Img", "Graph", "Select_Grp", "Data",
                                                "Mode 1: Solder joint crack", "Mode 2: Pad lift",
                                                "Mode 3: Solder joint lift", "Mode 4: Intermetallic break",
                                                "Mode 5: Component damage", "Mode 6: Component detached",
                                                "Mode 7: Flex torn", "Operator", "Time_Update", "Remark"
                                            });
                                    }
                                    else
                                    {
                                        Data_tbl = dt_load.AsDataView().ToTable(false,
                                            new string[]
                                            {
                                                "ID", "ItemCode", "LotNo", "Sheet", "Region", "Sample", "Image",
                                                "Select_Img", "Graph", "Select_Grp", "Data",
                                                "Mode 1: Solder joint crack", "Mode 2: Pad lift",
                                                "Mode 3: Solder joint lift", "Mode 4: Intermetallic break",
                                                "Mode 5: Component damage", "Mode 6: Component detached",
                                                "Mode 7: Flex torn", "Operator", "Time_Update", "Remark"
                                            });
                                    }

                                    break;

                                case "SHEAR_TEST":
                                    dt_load = load_data_logfile_sheartest(f_folder, infor);
                                    if (dt_load.Columns.Contains("ProductID"))
                                    {
                                        Data_tbl = dt_load.AsDataView().ToTable(false,
                                            new string[]
                                            {
                                                "ID", "ProductID", "ItemCode", "LotNo", "Sheet", "Region", "Sample",
                                                "Image", "Select_Img", "Graph", "Select_Grp", "Data",
                                                "Mode 1: Solder joint crack", "Mode 2: Pad lift",
                                                "Mode 3: Solder joint lift", "Mode 4: Intermetallic break",
                                                "Mode 5: Component damage", "Mode 6: Component detached",
                                                "Mode 7: Flex torn", "Operator", "Time_Update", "Remark"
                                            });
                                    }
                                    else
                                    {
                                        Data_tbl = dt_load.AsDataView().ToTable(false,
                                            new string[]
                                            {
                                                "ID", "ItemCode", "LotNo", "Sheet", "Region", "Sample", "Image",
                                                "Select_Img", "Graph", "Select_Grp", "Data",
                                                "Mode 1: Solder joint crack", "Mode 2: Pad lift",
                                                "Mode 3: Solder joint lift", "Mode 4: Intermetallic break",
                                                "Mode 5: Component damage", "Mode 6: Component detached",
                                                "Mode 7: Flex torn", "Operator", "Time_Update", "Remark"
                                            });
                                    }

                                    break;

                                case "IQC_UNMATING_PULL_TEST":
                                    if (txt_itemcode_nvl.Text != "" && txt_lotno_nvl.Text != "")
                                    {
                                        dt_load = load_data_logfile_unmating_pulltest(f_folder,
                                            "IQC_UNMATING_PULL_TEST_NAS", infor);
                                        Data_tbl = dt_load.AsDataView().ToTable(false,
                                            new string[]
                                            {
                                                "ID", "ItemCode", "LotNo", "Sheet", "Region", "Sample", "Image",
                                                "Select_Img", "Graph", "Select_Grp", "Data", "Operator", "Time_Update",
                                                "Remark"
                                            });
                                    }
                                    else
                                    {
                                        MessageBox.Show(new Form { TopMost = true },
                                            "Nhập đầy đủ ItemCode(NVL) và LotNo(NVL)", "Thông báo");
                                    }

                                    break;

                                case "IQC_LINER_PEELING_COUPON":
                                    if (txt_itemcode_nvl.Text != "" && txt_lotno_nvl.Text != "")
                                    {
                                        if (cb_Type.SelectedItem.ToString() == "NPI")
                                        {
                                            dt_load = load_data_logfile_coupon(f_folder, "IQC_LINER_PEELING_COUPON_NAS",
                                                infor);
                                        }
                                        else
                                        {
                                            dt_load = load_data_logfile_onproduct(f_folder, infor);
                                        }

                                        Data_tbl = dt_load.AsDataView().ToTable(false,
                                            new string[]
                                            {
                                                "ID", "ItemCode", "LotNo", "Sheet", "Region", "Sample", "Image",
                                                "Select_Img", "Graph", "Select_Grp", "Data", "Operator", "Time_Update",
                                                "Remark"
                                            });
                                    }
                                    else
                                    {
                                        MessageBox.Show(new Form { TopMost = true },
                                            "Nhập đầy đủ ItemCode(NVL) và LotNo(NVL)", "Thông báo");
                                    }

                                    break;

                                case "IQC_PSA_PEELING_COUPON":
                                    if (cb_Type.SelectedItem.ToString() == "NPI")
                                    {
                                        dt_load = load_data_logfile_coupon(f_folder, "IQC_PSA_PEELING_COUPON", infor);
                                    }
                                    else
                                    {
                                        dt_load = load_data_logfile_onproduct(f_folder, infor);
                                    }

                                    Data_tbl = dt_load.AsDataView().ToTable(false,
                                        new string[]
                                        {
                                            "ID", "ItemCode", "LotNo", "Sheet", "Region", "Sample", "Image",
                                            "Select_Img", "Graph", "Select_Grp", "Data", "Operator", "Time_Update",
                                            "Remark"
                                        });
                                    break;

                                case "LINER_PEEL_TEST_ON_PRODUCT":
                                    dt_load = load_data_logfile_onproduct(f_folder, infor);
                                    Data_tbl = dt_load.AsDataView().ToTable(false,
                                        new string[]
                                        {
                                            "ID", "ItemCode", "LotNo", "Sheet", "Region", "Sample", "Image",
                                            "Select_Img", "Graph", "Select_Grp", "Data", "Operator", "Time_Update",
                                            "Remark"
                                        });
                                    break;

                                case "PSA_PEEL_TEST_ON_PRODUCT":
                                    dt_load = load_data_logfile_onproduct(f_folder, infor);
                                    Data_tbl = dt_load.AsDataView().ToTable(false,
                                        new string[]
                                        {
                                            "ID", "ItemCode", "LotNo", "Sheet", "Region", "Sample", "Image",
                                            "Select_Img", "Graph", "Select_Grp", "Data", "Operator", "Time_Update",
                                            "Remark"
                                        });
                                    break;

                                case "CROSS_SECTION":
                                    try
                                    {
                                        Data_tbl = load_data_logfile_cross_section(f_folder, infor,
                                            cbStatus.Text);
                                    }
                                    catch (Exception exception)
                                    {
                                        MessageBox.Show($"Error Read cross section: {exception}!");
                                    }

                                    break;

                                case "GAP_CONNECTOR":

                                    Data_tbl = load_data_logfile_gap_connector(f_folder, infor,
                                        cbStatus.Text);
                                    break;
                            }

                            dgv_logfile.DataSource = Data_tbl;
                            btn_checkall.Visible = true;
                            myCode.Disable_Sort_DGV(dgv_logfile);

                            if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST" || sheet == "SHEAR_TEST")
                            {
                                dgv_logfile.Columns["Mode 1: Solder joint crack"].ReadOnly = true;
                                dgv_logfile.Columns["Mode 3: Solder joint lift"].ReadOnly = true;
                                dgv_logfile.Columns["Mode 4: Intermetallic break"].ReadOnly = true;
                                dgv_logfile.Columns["Mode 6: Component detached"].ReadOnly = true;
                                dgv_logfile.Columns["Mode 7: Flex torn"].ReadOnly = true;
                            }

                            if (sheet != "SHEAR_TEST")
                            {
                                dgv_logfile.Columns["Data"].ReadOnly = true;
                            }

                            for (int i = 0; i < dgv_logfile.Rows.Count; i++)
                            {
                                dgv_logfile.Rows[i].Cells["Data"].Style.BackColor = Color.White;
                            }


                            if (cb_Type.SelectedIndex != -1)
                            {
                                switch (sheet)
                                {
                                    case "PEEL_TEST":
                                        Check_spec_Peel_Pull(dgv_logfile, sheet);
                                        break;

                                    case "MATING_PULL_TEST":
                                        Check_spec_Peel_Pull(dgv_logfile, sheet);
                                        break;

                                    case "SHEAR_TEST":
                                        Check_spec_ShearTest_new(dgv_logfile);
                                        break;

                                    case "IQC_UNMATING_PULL_TEST":
                                        Check_spec_unmatingpull(dgv_logfile, sheet);
                                        break;

                                    case "IQC_LINER_PEELING_COUPON":
                                        check_spec_coupon(dgv_logfile, sheet);
                                        break;

                                    case "IQC_PSA_PEELING_COUPON":
                                        check_spec_coupon(dgv_logfile, sheet);
                                        break;

                                    case "LINER_PEEL_TEST_ON_PRODUCT":
                                        check_spec_onproduct(dgv_logfile, sheet);
                                        break;

                                    case "PSA_PEEL_TEST_ON_PRODUCT":
                                        check_spec_onproduct(dgv_logfile, sheet);
                                        break;

                                    case "CROSS_SECTION":
                                        Check_spec_crossection(dgv_logfile, sheet);
                                        break;

                                    case "GAP_CONNECTOR":
                                        Check_spec_GAP(dgv_logfile, sheet);
                                        break;
                                }

                                Check_Alldata(dgv_logfile);
                            }

                            lbl_judge_logfile.BackColor = Color.Transparent;
                            for (int i = 0; i < dgv_logfile.Rows.Count; i++)
                            {
                                if (dgv_logfile.Rows[i].Cells["Data"].Style.BackColor == Color.Red)
                                {
                                    lbl_judge_logfile.BackColor = Color.Red;
                                    break;
                                }
                            }

                            resize_column_image(dgv_logfile);
                            // txt_selected.Text = count_selected();
                            txt_selected.Text = count_selected();
                        }
                        else
                        {
                            if (MessageBox.Show(new Form { TopMost = true },
                                    "Dữ liệu đã tồn tại. Bạn có muốn thay thế toàn bộ dữ liệu ?", "Thông báo",
                                    MessageBoxButtons.YesNo) == DialogResult.Yes)
                            {
                                if (UserSession.Instance.IsLoggedIn == true)
                                {
                                    chk = true;
                                    goto lbl_getdata;
                                }
                                else
                                {
                                    MessageBox.Show(new Form { TopMost = true },
                                        "Vui lòng đăng nhập để cập nhật dữ liệu", "Thông báo");
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Đường dẫn file không đúng", "Thông báo");
                    }
                }
            }
            else if (txtLogfile.Text == "" && cb_Type.SelectedIndex != -1 && txtItemCode.Text != "" &&
                     txtLotNo.Text != "")
            {
                if (!(cb_Type.SelectedItem.ToString() == "NPI"))
                    return;
                DataTable Data_tbl = new DataTable();
                string str1 = "/" + txt_ItemName.Text + "_" + txt_line.Text + "_" + txt_ca.Text + "_" + txt_date.Text +
                              "_" + txt_worker.Text + "_" + cb_Type.SelectedItem.ToString();
                string str2 = TDMK_Code.filter_str(new string[3] { "ItemCode", "LotNo", "Sheet" },
                    new string[3] { txt_itemcode_nvl.Text, txt_lotno_nvl.Text, sheet + str1 });
                string str3 = TDMK_Code.filter_str(new string[3] { "ItemCode", "LotNo", "Sheet" },
                    new string[3] { txtItemCode.Text, txtLotNo.Text, sheet + str1 });
                bool flag = true;
                if (TDMK_Code.Datatable_Filter(sqlcon, sheet, str3).Rows.Count > 0)
                    flag = false;
                for (; !flag; flag = true)
                {
                    if (MessageBox.Show((IWin32Window)new Form() { TopMost = true },
                            "Dữ liệu đã tồn tại. Bạn có muốn thay thế toàn bộ dữ liệu ?", "Thông báo",
                            MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        if (admin_mode == "Admin mode")
                            continue;
                        int num7 = (int)MessageBox.Show((IWin32Window)new Form() { TopMost = true },
                            "Vui lòng đăng nhập để cập nhật dữ liệu", "Thông báo");
                    }

                    goto label_127;
                }

                if (sheet.Contains("CROSS") || sheet.Contains("GAP"))
                {
                    Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, sheet, str2);
                    Data_tbl.Columns.Add("Select", typeof(bool));
                    foreach (DataRow row in (InternalDataCollectionBase)Data_tbl.Rows)
                    {
                        row["ItemCode"] = (object)txtItemCode.Text;
                        row["LotNo"] = (object)txtLotNo.Text;
                        row["Select"] = (object)true;
                    }
                }
                else if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST" || sheet.Contains("ON_PRODUCT"))
                {
                    DataTable source = TDMK_Code.Datatable_Filter(sqlcon, sheet, str2);
                    Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, sheet,
                        TDMK_Code.filter_str(new string[1] { "ItemCode" }, new string[1] { txtItemCode.Text })).Clone();
                    Data_tbl.Columns.Add("Select_Img", typeof(bool));
                    Data_tbl.Columns.Add("Select_Grp", typeof(bool));
                    foreach (string region in source.AsEnumerable()
                                 .Select<DataRow, string>(
                                     (System.Func<DataRow, string>)(x => x.Field<string>("Region"))).Distinct<string>()
                                 .ToList<string>())
                        insert_data_refer(ref Data_tbl, Data_tbl.Rows.Count + 1, region);
                }

                dgv_logfile.DataSource = (object)Data_tbl;
                btn_checkall.Visible = true;
                myCode.Disable_Sort_DGV(dgv_logfile);
                if (cb_Type.SelectedIndex != -1)
                {
                    switch (sheet)
                    {
                        case "PEEL_TEST":
                            Check_spec_Peel_Pull(dgv_logfile, sheet);
                            break;
                        case "MATING_PULL_TEST":
                            Check_spec_Peel_Pull(dgv_logfile, sheet);
                            break;
                        case "LINER_PEEL_TEST_ON_PRODUCT":
                            check_spec_onproduct(dgv_logfile, sheet);
                            break;
                        case "PSA_PEEL_TEST_ON_PRODUCT":
                            check_spec_onproduct(dgv_logfile, sheet);
                            break;
                        case "CROSS_SECTION":
                            Check_spec_crossection(dgv_logfile, sheet);
                            break;
                        case "GAP_CONNECTOR":
                            Check_spec_GAP(dgv_logfile, sheet);
                            break;
                    }

                    Check_Alldata(dgv_logfile);
                }

                lbl_judge_logfile.BackColor = Color.Transparent;
                for (int index = 0; index < dgv_logfile.Rows.Count; ++index)
                {
                    if (dgv_logfile.Rows[index].Cells["Data"].Style.BackColor == Color.Red)
                    {
                        lbl_judge_logfile.BackColor = Color.Red;
                        break;
                    }
                }

                resize_column_image(dgv_logfile);
                //  txt_selected.Text = count_selected();
                txt_selected.Text = count_selected();
                label_127: ;
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },
                    "Vui lòng điền đầy đủ thông tin ItemCode / LotNo / Operator / Type / Logfile Location",
                    "Thông báo");
            }
        }


        private void Setmode(string value)
        {
            this.lbl_Login.Text = value;
            admin_mode = value;
            if (value == "Admin mode")
            {
                //this.lbl_Login.BackColor = Color.GreenYellow;
                this.lbl_Login.ForeColor = Color.Green;
                txt_qty.Enabled = true;
            }
            else
            {
                //this.lbl_Login.BackColor = Color.Yellow;
                this.lbl_Login.ForeColor = Color.SteelBlue;
                txt_qty.Enabled = false;
                edit_mode = false;
                btnEdit.BackColor = Color.Transparent;
            }
        }


        private void check_mode(string value)
        {
            if (sheet == "CROSS_SECTION" || sheet == "GAP_CONNECTOR")
            {
                if (value.Contains("false"))
                {
                    int row = int.Parse(value.Split('_')[0]);
                    this.dgv_logfile.Rows[row].Cells["Select"].Value = false;
                }
                else if (value.Contains("true"))
                {
                    int row = int.Parse(value.Split('_')[0]);
                    this.dgv_logfile.Rows[row].Cells["Select"].Value = true;
                }
            }
            else
            {
                string col_select = "";
                if (col_name_click == "Image")
                {
                    col_select = "Select_Img";
                }
                else if (col_name_click == "Graph")
                {
                    col_select = "Select_Grp";
                }

                if (col_name_click != "")
                {
                    if (value.Contains("false"))
                    {
                        int row = int.Parse(value.Split('_')[0]);
                        this.dgv_logfile.Rows[row].Cells[col_select].Value = false;
                    }
                    else if (value.Contains("true"))
                    {
                        int row = int.Parse(value.Split('_')[0]);
                        this.dgv_logfile.Rows[row].Cells[col_select].Value = true;
                    }
                }
            }
        }

        private void judge_mode(string value)
        {
            if (value != "")
            {
                int r_inx = int.Parse(value.Split('_')[0]);
                dgv_logfile.Rows[r_inx].Cells["Mode 2: Pad lift"].Value = value.Split('_')[1];
                dgv_logfile.Rows[r_inx].Cells["Mode 5: Component damage"].Value = value.Split('_')[2];

                string per =
                    (100 - double.Parse(value.Split('_')[1].Split('%')[0]) -
                     double.Parse(value.Split('_')[2].Split('%')[0])).ToString() + "%";
                int a_col = 0;
                int b = 0;

                if (value.Split('_')[1].Contains("(") && value.Split('_')[1].Contains("/"))
                {
                    b = int.Parse(value.Split('_')[1].Split('(')[1].Split(')')[0].Split('/')[1]);
                    a_col += int.Parse(value.Split('_')[1].Split('(')[1].Split(')')[0].Split('/')[0]);
                }

                if (value.Split('_')[2].Contains("(") && value.Split('_')[2].Contains("/"))
                {
                    b = int.Parse(value.Split('_')[2].Split('(')[1].Split(')')[0].Split('/')[1]);
                    a_col += int.Parse(value.Split('_')[2].Split('(')[1].Split(')')[0].Split('/')[0]);
                }

                if (b != 0 && a_col != 0)
                {
                    dgv_logfile.Rows[r_inx].Cells["Mode 1: Solder joint crack"].Value =
                        per + "(" + (b - a_col).ToString() + "/" + b.ToString() + ")";
                }


                //    if ((value.Split('_')[1].Contains("(") && value.Split('_')[1].Contains("/")) || (value.Split('_')[2].Contains("(") && value.Split('_')[2].Contains("/")))
                //{
                //    try
                //    {
                //        b = int.Parse(value.Split('_')[1].Split('(')[1].Split(')')[0].Split('/')[1]);
                //    }
                //    catch
                //    {
                //        b = int.Parse(value.Split('_')[2].Split('(')[1].Split(')')[0].Split('/')[1]);
                //    }


                //    if (myCode.IsNumeric(value.Split('_')[1].Split('(')[1].Split(')')[0].Split('/')[1]))
                //    { 
                //       a_col += int.Parse(value.Split('_')[1].Split('(')[1].Split(')')[0].Split('/')[1]);
                //    }
                //    if (myCode.IsNumeric(value.Split('_')[2].Split('(')[1].Split(')')[0].Split('/')[1]))
                //    { 
                //       a_col += int.Parse(value.Split('_')[2].Split('(')[1].Split(')')[0].Split('/')[1]);
                //    } 

                //    //if (value.Split('_')[1].Contains("("))
                //    //    {
                //    //        a_col += int.Parse(value.Split('_')[1].Split('(')[1].Split('/')[0]);
                //    //    }
                //    //    if (value.Split('_')[2].Contains("("))
                //    //    {
                //    //        a_col += int.Parse(value.Split('_')[2].Split('(')[1].Split('/')[0]);
                //    //    } 

                //    //}
                //}
            }
        }

        private void shear_data(string value)
        {
            if (myCode.IsNumeric(value.Split('_')[1]))
            {
                int r_inx = int.Parse(value.Split('_')[0]);
                dgv_logfile.Rows[r_inx].Cells["Data"].Value = value.Split('_')[1];
            }
        }

        private void setup_sochan(int value)
        {
            set_chan = value;
            txt_setchan.Text = set_chan.ToString();
        }

        public Boolean check_enough(DataTable Data_tbl)
        {
            bool chk = true;
            if (myCode.IsNumeric(txt_qty.Text))
            {
                List<string> lst_region = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).ToList();

                foreach (string region in lst_region)
                {
                    DataView dv = Data_tbl.AsDataView();
                    string filter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { region });
                    dv.RowFilter = filter;
                    if (dv.Count < int.Parse(txt_qty.Text))
                    {
                        chk = false;
                        return chk;
                    }
                }
            }

            return chk;
        }

        public Boolean check_enough_data(DataGridView dgv)
        {
            DataTable Data_tbl = (DataTable)dgv.DataSource;
            bool chk = true;
            switch (sheet)
            {
                case "PEEL_TEST":
                    chk = check_enough(Data_tbl);
                    break;

                case "MATING_PULL_TEST":
                    chk = check_enough(Data_tbl);
                    break;

                case "SHEAR_TEST":
                    chk = check_enough(Data_tbl);
                    break;

                case "IQC_UNMATING_PULL_TEST":
                    chk = check_enough(Data_tbl);
                    break;

                case "IQC_LINER_PEELING_COUPON":
                    chk = check_enough(Data_tbl);
                    break;

                case "IQC_PSA_PEELING_COUPON":
                    chk = check_enough(Data_tbl);
                    break;

                case "LINER_PEEL_TEST_ON_PRODUCT":
                    chk = check_enough(Data_tbl);
                    break;

                case "PSA_PEEL_TEST_ON_PRODUCT":
                    chk = check_enough(Data_tbl);
                    break;

                case "CROSS_SECTION":
                    //DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, sheet, cb_Type.SelectedItem.ToString() }));
                    //if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                    //{
                    //    dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, sheet, "MASS" }));
                    //}
                    DataTable dt_spec = new DataTable();
                    if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                    {
                        dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                            TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                                new string[] { txtItemCode.Text, sheet, "MASS" }));
                    }
                    else if (cb_Type.SelectedItem.ToString() == "LQ" || cb_Type.SelectedItem.ToString() == "NPI")
                    {
                        dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                            TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                                new string[] { txtItemCode.Text, sheet, "NPI" }));
                    }

                    if (dt_spec.Rows.Count > 0)
                    {
                        string location_spec = dt_spec.Rows[0]["Location"].ToString();

                        int col_begin = int.Parse(location_spec.Split('+')[0].Split(';')[1]);
                        string[] arr_region = location_spec.Split('+')[1].Split('_');
                        int count_lk_format = (arr_region.Length - 1) / 2;

                        // var lst_ngang = Data_tbl.AsEnumerable().Where(x => x.Field<string>("Region").Contains("NGANG") && x.Field<string>("Region").Contains("TRU") == false).Select(s => s).Distinct();
                        List<string> lst_region_cross = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region"))
                            .Distinct().ToList();
                        List<string> lst_ngang = lst_region_cross.Where(s => s.Contains("NGANG") && !s.Contains("TRU"))
                            .ToList();
                        int count_lk_data = lst_ngang.Count;
                        if (count_lk_data < count_lk_format)
                        {
                            chk = false;
                        }
                        else
                        {
                            foreach (string region in lst_region_cross)
                            {
                                DataView dv = Data_tbl.AsDataView();
                                string filter =
                                    TDMK_Code.filter_str(new string[] { "Region" }, new string[] { region });
                                dv.RowFilter = filter;

                                if (region.Contains("NGANG"))
                                {
                                    if (dv.Count < 2 * int.Parse(txt_qty.Text))
                                    {
                                        chk = false;
                                        return chk;
                                    }
                                }
                                else
                                {
                                    if (dv.Count < int.Parse(txt_qty.Text))
                                    {
                                        chk = false;
                                        return chk;
                                    }
                                }
                            }
                        }
                    }

                    break;

                case "GAP_CONNECTOR":
                    List<string> lst_region = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).ToList();

                    foreach (string region in lst_region)
                    {
                        DataView dv = Data_tbl.AsDataView();
                        string filter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { region });
                        dv.RowFilter = filter;

                        if (region.Contains("DOC"))
                        {
                            if (dv.Count < int.Parse(txt_qty.Text))
                            {
                                chk = false;
                                return chk;
                            }
                        }
                        else if (region.Contains("TRU"))
                        {
                            if (dv.Count < 2 * int.Parse(txt_qty.Text))
                            {
                                chk = false;
                                return chk;
                            }
                        }
                    }

                    break;
            }

            switch (sheet)
            {
                case "PEEL_TEST":
                case "MATING_PULL_TEST":
                case "SHEAR_TEST":
                    try
                    {
                        DataTable dt = (DataTable)dgv.DataSource;
                        foreach (DataRow row in dt.Rows)
                        {
                            int tong = 0;
                            foreach (DataColumn column in dt.Columns)
                            {
                                if (column.ColumnName.Contains("Mode"))
                                {
                                    string value = row[column].ToString().Split('%')[1].Split('(')[1].Replace(")", "")
                                        .Replace(" ", "").Replace(")", "");
                                    tong += int.Parse(value.Split('/')[0]);
                                }
                            }

                            if (tong != int.Parse(txt_setchan.Text))
                            {
                                MessageBox.Show("Số lượng chân không phù hợp [ID] = " + row["ID"]);
                                return false;
                            }
                        }
                    }
                    catch
                    {
                        return false;
                    }

                    break;
                default:
                    Debugger.Break();
                    break;
            }

            return chk;
        }


        private void btn_save_Click(object sender, EventArgs e)
        {
            if (sheet == "CROSS_SECTION" && (cbStatus.Text == "Shield b2b" || cbStatus.Text == "Clip"))
            {
                try
                {
                    string itemCode = txtItemCode.Text;
                    string lotNo = txtLotNo.Text;
                    string cbStatusZ = cbStatus.Text;
                    DataTable dt = (DataTable)dgv_Analysis.DataSource;
                    DataTable crossSctionDT = CrossSectionService.load(itemCode, lotNo, cbStatusZ);
                    if (crossSctionDT.Rows.Count > 0)
                    {
                        if (MessageBox.Show("Data is valid, do you want overwrite", "Warning",
                                MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            crossSctionDT.Rows.Clear();
                        }
                        else
                        {
                            return;
                        }
                    }

                    CrossSectionService.Save(dt, itemCode, lotNo, cbStatusZ, crossSctionDT);
                    dgv_Analysis.DataSource = new DataTable();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Save cross section: {ex.Message}");
                }

                return;
            }

            if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST" || sheet == "SHEAR_TEST")
            {
                if (txt_setchan.Text == "0" || !myCode.IsNumeric(txt_setchan.Text))
                {
                    MessageBox.Show(new Form { TopMost = true }, "Vui lòng nhập số chân linh kiện", "Thông báo");
                    txt_setchan.BackColor = Color.Yellow;
                    return;
                }
                else if (dgv_Analysis.Rows.Count > 0)
                {
                    if (dgv_Analysis.Rows[0].Cells["Mode 1: Solder joint crack"].Value.ToString() == "100%")
                        dgv_Analysis.Rows[0].Cells["Mode 1: Solder joint crack"].Value =
                            "100%(" + txt_setchan.Text + "/" + txt_setchan.Text + ")";
                }

                switch (sheet)
                {
                    case "PEEL_TEST":
                    case "MATING_PULL_TEST":
                    case "SHEAR_TEST":
                        try
                        {
                            DataTable dt = (DataTable)dgv_Analysis.DataSource;
                            foreach (DataRow row in dt.Rows)
                            {
                                int tong = 0;
                                foreach (DataColumn column in dt.Columns)
                                {
                                    if (column.ColumnName.Contains("Mode"))
                                    {
                                        string value = row[column].ToString().Split('%')[1].Split('(')[1]
                                            .Replace(")", "").Replace(" ", "").Replace(")", "");
                                        tong += int.Parse(value.Split('/')[0]);
                                    }
                                }

                                if (tong != int.Parse(txt_setchan.Text))
                                {
                                    if (MessageBox.Show(
                                            "Số lượng chân không phù hợp [ID] = " + row["ID"] + "Bạn có muốn tiếp tục?",
                                            "", MessageBoxButtons.YesNo) == DialogResult.No)
                                    {
                                        return;
                                    }
                                }
                            }
                        }
                        catch
                        {
                        }

                        break;
                    default:
                        Debugger.Break();
                        break;
                }
            }

            if (txtItemCode.Text != "" && txtLotNo.Text != "" && cb_Type.SelectedIndex != -1 && txt_qty.Text != "")
            {
                string itemCode = txtItemCode.Text, lotNo = txtLotNo.Text;
                string infor = "";
                string filter_str = "";
                infor = "/" + txt_ItemName.Text + "_" + txt_line.Text + "_" + txt_ca.Text + "_" + txt_date.Text + "_" +
                        txt_worker.Text + "_" + cb_Type.SelectedItem.ToString();
                filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Sheet" },
                    new string[] { txtItemCode.Text, txtLotNo.Text, sheet + infor });
                if (sheet.Contains("UNMATING") || sheet.Contains("COUPON"))
                {
                    if (txt_itemcode_nvl.Text != "" && txt_lotno_nvl.Text != "")
                    {
                        filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Sheet" },
                            new string[] { txt_itemcode_nvl.Text, txt_lotno_nvl.Text, sheet + infor });
                        itemCode = txt_itemcode_nvl.Text;
                        lotNo = txt_lotno_nvl.Text;
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Chưa nhập đủ ItemCode/lotNo", "Thông báo");
                    }
                }

                lblsave:
                DataTable dt = TDMK_Code.Datatable_Filter(sqlcon,
                    sheet + (_PRIME_PEEL_TEST ? "_WITHOUT_SUS" : "") + "_NAS", filter_str);
                //DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, sheet, cb_Type.SelectedItem.ToString() }));
                if (dt.Rows.Count == 0)
                {
                    if (dgv_Analysis.Rows.Count > 0)
                    {
                        if (check_enough_data(dgv_Analysis))
                        {
                            bool chk = true;
                            for (int i = 0; i < dgv_Analysis.Columns.Count; i++)
                            {
                                for (int k = 0; k < dgv_Analysis.Rows.Count; k++)
                                {
                                    if (dgv_Analysis.Rows[k].Cells[i].Style.BackColor == Color.Red)
                                    {
                                        chk = false;
                                        break;
                                    }
                                }
                            }

                            save_lbl:
                            if (chk)
                            {
                                DataTable tbl_data_analysis = (DataTable)dgv_Analysis.DataSource;
                                int i = TDMK_Code.SQL_MAX(
                                    sheet + (_PRIME_PEEL_TEST ? "_WITHOUT_SUS" : "") +
                                    (!LegacyMode.Checked == true ? "_Nas" : ""), "ID", sqlcon) + 1;
                                //foreach (DataRow dr in tbl_data_analysis.Rows)
                                //{
                                //    dr[0] = i;
                                //    i++;
                                //}
                                if (tbl_data_analysis.Columns.Contains("ProductID"))
                                {
                                    string str = ProductIDService.ConverterProductID(tbl_data_analysis);

                                    string process = sheet + (_PRIME_PEEL_TEST ? "_WITHOUT_SUS" : "");
                                    try
                                    {
                                        ProductIDService.InsertProductID(itemCode, lotNo, process, str);
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show($" {ex.Message}", "Lưu productID thất bại",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        return;
                                    }

                                    tbl_data_analysis.Columns.Remove("ProductID");
                                }

                                NasRepository nas = new NasRepository();
                                string location = nas.HandleImageDataTable(tbl_data_analysis, sheet, itemCode, lotNo);
                                string json = ConverterService.DataTableToJson(tbl_data_analysis);
                                DataRow row = dt.NewRow();
                                string sheetZ = cb_Type.Text.ToString().Trim();
                                int s = location.Length;
                                //row["ID"] = i;

                                row["ItemCode"] = itemCode;
                                row["LotNo"] = lotNo;
                                row["Data"] = json;
                                row["Sheet"] = sheetZ;
                                row["LocationImg"] = location;
                                dt.Rows.Add(row);
                                try
                                {
                                    new DBContext().BuckDataTable(dt,
                                        sheet + (_PRIME_PEEL_TEST ? "_WITHOUT_SUS" : "") + "_NAS",
                                        new[] { "ItemCode", "LotNo" }, null, "ID");
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show($"Error: {ex.Message}");
                                }
                                //BatchBulkCopy(sqlcon, dt, sheet + (_PRIME_PEEL_TEST ? "_WITHOUT_SUS" : "") + "_NAS");

                                int st = 1;
                                foreach (DataRow dr in tbl_data_analysis.Rows)
                                {
                                    dr["ID"] = st;
                                    st++;
                                }

                                dgv_Analysis.DataSource = tbl_data_analysis;


                                MessageBox.Show(new Form { TopMost = true }, "Lưu dữ liệu thành công!", "Thông báo");
                                txt_selected.Text = "";
                                btn_checkall.Visible = false;
                            }
                            else
                            {
                                if (MessageBox.Show(new Form { TopMost = true },
                                        "Dữ liệu ngoài chuẩn. Bạn có muốn tiếp tục lưu không ?", "Thông báo",
                                        MessageBoxButtons.YesNo) == DialogResult.Yes)
                                {
                                    if (UserSession.Instance.IsLoggedIn)
                                    {
                                        chk = true;
                                        goto save_lbl;
                                    }
                                    else
                                    {
                                        // MessageBox.Show("Please, login to save data", "Thông báo");
                                        MessageBox.Show("Vui lòng đăng nhập để lưu dữ liệu", "Thông báo");
                                    }
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Không đủ dữ liệu. Không thể lưu!",
                                "Thông báo");
                            return;
                        }
                    }
                    else

                    {
                        MessageBox.Show(new Form { TopMost = true }, "Không có dữ liệu. Không thể lưu", "Thông báo");
                    }
                }
                else
                {
                    if (!edit_mode)
                    {
                        if (MessageBox.Show(new Form { TopMost = true },
                                "Dữ liệu đã tồn tại. Bạn có muốn thay thế không", "Thông báo",
                                MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            if (UserSession.Instance.IsLoggedIn)
                            {
                                TDMK_Code.Delelte_FilteredItem_arr(
                                    sheet + (_PRIME_PEEL_TEST ? "_WITHOUT_SUS" : "") + "_NAS", sqlcon, filter_str);
                                //   TDMK_Code.Delelte_FilteredItem_arr(sheet + "_LOGFILE", sqlcon, filter_str);
                                goto lblsave;
                            }
                            else
                            {
                                MessageBox.Show(new Form { TopMost = true }, "Vui lòng đăng nhập để cập nhật dữ liệu",
                                    "Thông báo");
                            }
                        }
                    }
                    else
                    {
                        if (MessageBox.Show(new Form { TopMost = true },
                                "Dữ liệu đã tồn tại. Bạn có muốn thay thế không", "Thông báo",
                                MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            if (admin_mode == "Admin mode")
                            {
                                TDMK_Code.Delelte_FilteredItem_arr(
                                    sheet + (_PRIME_PEEL_TEST ? "_WITHOUT_SUS" : "") + "_NAS", sqlcon, filter_str);

                                btnEdit.BackColor = Color.GreenYellow;
                                goto lblsave;
                            }
                            else
                            {
                                MessageBox.Show(new Form { TopMost = true }, "Vui lòng đăng nhập để cập nhật dữ liệu",
                                    "Thông báo");
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Chưa cài format!");
            }
        }

        private void dgv_logfile_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int col_inx = e.ColumnIndex;
            if (col_inx != -1)
            {
                if (sheet == "CROSS_SECTION" || sheet == "GAP_CONNECTOR")
                {
                    if (dgv_logfile.Columns[col_inx].Name.Contains("Image") ||
                        dgv_logfile.Columns[col_inx].Name.Contains("Graph"))
                    {
                        if (dgv_logfile.Columns.Contains("Select"))
                        {
                            int r_inx = e.RowIndex;
                            Select_Image fr1 = new Select_Image(check_mode, judge_mode, shear_data, setup_sochan);
                            fr1.sheet = sheet;
                            fr1.dt_image_ = (DataTable)dgv_logfile.DataSource;
                            fr1.r_inx_ = r_inx;
                            fr1.Show();
                        }
                    }
                }
                else
                {
                    if (dgv_logfile.Columns[col_inx].Name.Contains("Image") ||
                        dgv_logfile.Columns[col_inx].Name.Contains("Graph"))
                    {
                        //foreach(string col_name in lst_col_select)
                        //{
                        //if (dgv_logfile.Columns.Contains(col_name))
                        //{

                        if (dgv_logfile.Columns.Contains("Select_Img") && dgv_logfile.Columns.Contains("Select_Grp"))
                        {
                            int r_inx = e.RowIndex;
                            col_name_click = dgv_logfile.Columns[col_inx].Name;
                            Select_Image fr1 = new Select_Image(check_mode, judge_mode, shear_data, setup_sochan)
                                { TopMost = true };
                            fr1.sheet = sheet;
                            fr1.dt_image_ = (DataTable)dgv_logfile.DataSource;
                            fr1.r_inx_ = r_inx;
                            fr1.column_name_ = col_name_click;
                            if (set_chan != 0)
                            {
                                fr1.sochan_ = set_chan;
                            }

                            fr1.Show();
                        }
                        else
                        {
                            DataGridViewCell cur_cell = dgv_Analysis.CurrentCell;

                            if (myCode.checkDBNull(cur_cell.Value) != "")
                            {
                                View_detail_Image fr1 = new View_detail_Image() { TopMost = true };
                                fr1.data = (byte[])cur_cell.Value;
                                fr1.Show();
                            }
                        }
                        //}
                        //} 
                    }
                }
            }
        }

        public myExcel.Workbook create_export_wrk(string format_file, string process_name)
        {
            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            xlsApp.DisplayAlerts = false;
            myExcel.Workbook wb = TDMK_Code.open_excel_file(format_file, "", "");
            if (wb != null)
            {
                string _process = process_name.Replace("_", "").Replace("-", "").Replace(" ", "").Replace("(", "")
                    .Replace(")", "").ToUpper();
                string mySheet = "";
                foreach (myExcel.Worksheet tg_sht in wb.Worksheets)
                {
                    string cur_sht_name = tg_sht.Name.ToUpper().Replace("_", "").Replace("-", "").Replace(" ", "")
                        .Replace("(", "").Replace(")", "");
                    if (_process == cur_sht_name)
                    {
                        mySheet = tg_sht.Name;
                        break;
                    }
                }

                if (mySheet != "")
                {
                    myExcel.Workbook save_wb = TDMK_Code.Create_workbook();
                    wb.Sheets[mySheet].Copy(After: save_wb.Sheets[1]);
                    myExcel.Worksheet del_sht = save_wb.Sheets[1];
                    del_sht.Delete();
                    wb.Close();
                    xlsApp.DisplayAlerts = true;
                    return save_wb;
                }
                else
                {
                    wb.Close();
                    xlsApp.DisplayAlerts = true;
                    return null;
                }
            }
            else
            {
                return null;
            }
        }

        //public myExcel.Workbook copy_new_wrk(string format_file)
        //{
        //    myExcel.Application xlsApp = TDMK_Code.StartExcel();
        //    xlsApp.DisplayAlerts = false;
        //    myExcel.Workbook wb = TDMK_Code.open_excel_file(format_file, "", "");
        //    if (wb != null)
        //    {
        //        myExcel.Workbook save_wb = TDMK_Code.Create_workbook();
        //        wb.Sheets[1].Copy(After: save_wb.Sheets[1]);
        //        save_wb.Sheets.
        //        wb.Close();
        //        xlsApp.DisplayAlerts = true;
        //        return save_wb; 
        //    }
        //    else
        //    {
        //        return null;
        //    }
        //}


        public void Export_PeelTest(SqlConnection sqlcon, myExcel.Worksheet ws, string Itemcode, string Lotno)
        {
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet" },
                new string[] { Itemcode, "Peel Test" });
            DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", filter_str);
            DataTable Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, tbl_name_comment3,
                TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Sheet" },
                    new string[] { Itemcode, Lotno, "Peel Test" }));

            if (Data_tbl.Rows.Count > 0)
            {
                List<List<string>> lst_coldata = new List<List<string>>() { };
                string[] arr_name = { "Image", "Graph", "Data" };
                foreach (var item in arr_name)
                {
                    lst_coldata.Add(Data_tbl.AsEnumerable().Select(x => x.Field<string>(item)).ToList());
                }


                if (dt_spec.Rows.Count > 0)
                {
                    string[] str_location = dt_spec.Rows[0]["Location"].ToString().Split('+');
                    myExcel.Range cell_Pic = ws.Cells[str_location[0].Split(';')[1], str_location[0].Split(';')[2]];
                    myExcel.Range cell_graph = ws.Cells[str_location[1].Split(';')[1], str_location[0].Split(';')[2]];
                    myExcel.Range cell_data = ws.Cells[str_location[2].Split(';')[1], str_location[0].Split(';')[2]];

                    int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
                    for (int i = 0; i < count_sample; i++)
                    {
                        cell_Pic.Offset[0, i].Value = lst_coldata[0][i];
                        cell_graph.Offset[0, i].Value = lst_coldata[0][i];
                        cell_data.Offset[0, i].Value = lst_coldata[0][i];
                    }


                    MessageBox.Show("Xuất dữ liệu thành công!", "Thông báo");
                }
            }
        }

        public void InsertPicture_Name(myExcel.Worksheet tar_wrksht, myExcel.Range tar_range, string picFile,
            int margin)
        {
            float left;
            float top;
            float width;
            float height;
            string pic_name = Path.GetFileName(picFile);

            if (tar_range.MergeCells)
            {
                myExcel.Range refer_range = tar_range.MergeArea; //Range["G16"];
                left = (float)(tar_range.Left) + margin;
                top = (float)(tar_range.Top) + margin;
                width = (float)(refer_range.Width) - 2 * margin;
                height = (float)(refer_range.Height) - 2 * margin;
                //width = (float)(refer_range.Width * 2.8) + 2 * margin;
                //height = (float)(refer_range.Height * 8.9) + 2 * margin;
            }
            else
            {
                left = (float)tar_range.Left + margin;
                top = (float)tar_range.Top + margin;
                width = (float)tar_range.Width - 2 * margin;
                height = (float)tar_range.Height - 2 * margin;
            }

            //myExcel.Shape sel_picture = tar_wrksht.Shapes.AddPicture(picFile, MsoTriState.msoFalse, MsoTriState.msoTrue, left, top, width, height);
            myExcel.Shape sel_picture = tar_wrksht.Shapes.AddPicture2(picFile, MsoTriState.msoFalse,
                MsoTriState.msoTrue, left, top, width, height, MsoPictureCompress.msoPictureCompressFalse);
            //tar_wrksht.Paste(tar_range, sel_picture);
            sel_picture.LockAspectRatio = MsoTriState.msoTrue;
            sel_picture.Placement = myExcel.XlPlacement.xlMoveAndSize;
            sel_picture.Name = pic_name;
        }

        public void Export_DatatableImage_Excel_ngang(DataTable dt, string Image_Col_name, myExcel.Worksheet tar_wrksht,
            myExcel.Range sel_rgn, bool row_offset, int count_sample)
        {
            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = Path.Combine(file_folder, DateTime.Now.ToString("ddMMMyy_HHmmss") + ".jpg");
            int r_inx = 0;

            myExcel.Range rgn_begin = sel_rgn;

            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (count < 2 * count_sample)
                {
                    byte[] img_byte = (byte[])(dr[Image_Col_name]);
                    using (MemoryStream ms = new MemoryStream(img_byte))
                    {
                        Image temp = Image.FromStream(ms);
                        temp.Save(file_dic);
                    }

                    InsertPicture_Name(tar_wrksht, sel_rgn, file_dic, 5);
                    int col_offset = sel_rgn.Columns.Count;
                    int row_off = sel_rgn.Rows.Count;
                    if (row_offset)
                    {
                        sel_rgn = sel_rgn.Offset[row_off, 0];
                    }
                    else
                    {
                        sel_rgn = sel_rgn.Offset[0, col_offset];
                    }

                    r_inx++;
                    count++;


                    if (count == count_sample)
                    {
                        sel_rgn = sel_rgn.Offset[4, -count_sample];
                        //count = 0;
                    }
                }
                else
                {
                    break;
                }
            }

            try
            {
                File.Delete(file_dic);
            }
            catch
            {
            }
        }

        public void Export_DatatableImage_Excel_ngang_mass(DataTable dt, string Image_Col_name,
            myExcel.Worksheet tar_wrksht, myExcel.Range sel_rgn, bool row_offset, int count_sample)
        {
            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = Path.Combine(file_folder, DateTime.Now.ToString("ddMMMyy_HHmmss") + ".jpg");
            int r_inx = 0;

            myExcel.Range rgn_begin = sel_rgn;

            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (count < 2 * count_sample)
                {
                    byte[] img_byte = (byte[])(dr[Image_Col_name]);
                    using (MemoryStream ms = new MemoryStream(img_byte))
                    {
                        Image temp = Image.FromStream(ms);
                        temp.Save(file_dic);
                    }

                    InsertPicture_Name(tar_wrksht, sel_rgn, file_dic, 5);
                    int col_offset = sel_rgn.Columns.Count;
                    int row_off = sel_rgn.Rows.Count;
                    if (row_offset)
                    {
                        sel_rgn = sel_rgn.Offset[row_off, 0];
                    }
                    else
                    {
                        sel_rgn = sel_rgn.Offset[0, col_offset];
                    }

                    r_inx++;
                    count++;

                    if (count == count_sample)
                    {
                        sel_rgn = sel_rgn.Offset[3, -count_sample];
                    }
                }
                else
                {
                    break;
                }
            }

            try
            {
                File.Delete(file_dic);
            }
            catch
            {
            }
        }


        public void Export_DatatableImage_Excel_unmating(DataTable dt, string Image_Col_name,
            myExcel.Worksheet tar_wrksht, myExcel.Range sel_rgn, bool row_offset, int count_sample)
        {
            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = Path.Combine(file_folder, DateTime.Now.ToString("ddMMMyy_HHmmss") + ".jpg");
            int r_inx = 0;

            myExcel.Range rgn_begin = sel_rgn;

            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (count < 2 * count_sample)
                {
                    byte[] img_byte = (byte[])(dr[Image_Col_name]);
                    using (MemoryStream ms = new MemoryStream(img_byte))
                    {
                        Image temp = Image.FromStream(ms);
                        temp.Save(file_dic);
                    }

                    InsertPicture_Name(tar_wrksht, sel_rgn, file_dic, 5);
                    int col_offset = sel_rgn.Columns.Count;
                    int row_off = sel_rgn.Rows.Count;
                    if (row_offset)
                    {
                        sel_rgn = sel_rgn.Offset[row_off, 0];
                    }
                    else
                    {
                        sel_rgn = sel_rgn.Offset[0, col_offset];
                    }

                    r_inx++;
                    count++;

                    if (count == count_sample)
                    {
                        sel_rgn = sel_rgn.Offset[1, -count_sample];
                    }
                }
                else
                {
                    break;
                }
            }

            try
            {
                File.Delete(file_dic);
            }
            catch
            {
            }
        }

        public void Export_DatatableImage_Excel_trungang(DataTable dt, string Image_Col_name,
            myExcel.Worksheet tar_wrksht, myExcel.Range sel_rgn, bool row_offset, int count_sample, int r_offset)
        {
            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = Path.Combine(file_folder, DateTime.Now.ToString("ddMMMyy_HHmmss") + ".jpg");
            int r_inx = 0;
            // myExcel.Range sel_rgn = tar_wrksht.Range[tar_rgn];
            myExcel.Range rgn_begin = sel_rgn;

            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (count < 2 * count_sample)
                {
                    byte[] img_byte = (byte[])(dr[Image_Col_name]);
                    using (MemoryStream ms = new MemoryStream(img_byte))
                    {
                        Image temp = Image.FromStream(ms);
                        temp.Save(file_dic);
                    }

                    InsertPicture_Name(tar_wrksht, sel_rgn, file_dic, 5);
                    int col_offset = sel_rgn.Columns.Count;
                    int row_off = sel_rgn.Rows.Count;
                    if (row_offset)
                    {
                        sel_rgn = sel_rgn.Offset[row_off, 0];
                    }
                    else
                    {
                        sel_rgn = sel_rgn.Offset[0, col_offset];
                    }

                    r_inx++;
                    count++;


                    if (count == count_sample)
                    {
                        sel_rgn = sel_rgn.Offset[r_offset, -count_sample];
                        //count = 0;
                    }
                }
            }

            try
            {
                File.Delete(file_dic);
            }
            catch
            {
            }
        }

        public void Export_DatatableImage_Excel(DataTable dt, string Image_Col_name, myExcel.Worksheet tar_wrksht,
            myExcel.Range sel_rgn, bool row_offset, int count_sample)
        {
            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = Path.Combine(file_folder, DateTime.Now.ToString("ddMMMyy_HHmmss") + ".jpg");
            int r_inx = 0;
            // myExcel.Range sel_rgn = tar_wrksht.Range[tar_rgn];
            myExcel.Range rgn_begin = sel_rgn;

            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (count < count_sample)
                {
                    byte[] img_byte = (byte[])(dr[Image_Col_name]);
                    using (MemoryStream ms = new MemoryStream(img_byte))
                    {
                        Image temp = Image.FromStream(ms);
                        temp.Save(file_dic);
                    }

                    InsertPicture_Name(tar_wrksht, sel_rgn, file_dic, 5);
                    int col_offset = sel_rgn.Columns.Count;
                    int row_off = sel_rgn.Rows.Count;
                    if (row_offset)
                    {
                        sel_rgn = sel_rgn.Offset[row_off, 0];
                    }
                    else
                    {
                        sel_rgn = sel_rgn.Offset[0, col_offset];
                    }

                    r_inx++;
                }

                count++;
            }

            try
            {
                File.Delete(file_dic);
            }
            catch
            {
            }
        }


        public void Export_DatatableImage_Excel_Graph(DataTable dt, string Image_Col_name, myExcel.Worksheet tar_wrksht,
            myExcel.Range sel_rgn, bool row_offset, int count_sample)
        {
            string file_folder = Path.Combine(System.Windows.Forms.Application.StartupPath, "Temp");
            string file_dic = Path.Combine(file_folder, DateTime.Now.ToString("ddMMMyy_HHmmss") + ".jpg");
            int r_inx = 0;
            // myExcel.Range sel_rgn = tar_wrksht.Range[tar_rgn];
            myExcel.Range rgn_begin = sel_rgn;
            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (count < count_sample)
                {
                    byte[] img_byte = (byte[])(dr[Image_Col_name]);
                    using (MemoryStream ms = new MemoryStream(img_byte))
                    {
                        Image _temp = Image.FromStream(ms);
                        //Image temp = ChangeColor((Bitmap)_temp, Color.FromArgb(243, 219, 203));
                        _temp.Save(file_dic);
                    }

                    InsertPicture_Name(tar_wrksht, sel_rgn, file_dic, 5);
                    int col_offset = sel_rgn.Columns.Count;
                    int row_off = sel_rgn.Rows.Count;
                    if (row_offset)
                    {
                        sel_rgn = sel_rgn.Offset[row_off, 0];
                    }
                    else
                    {
                        sel_rgn = sel_rgn.Offset[0, col_offset];
                    }

                    r_inx++;
                }

                count++;
            }

            try
            {
                File.Delete(file_dic);
            }
            catch
            {
            }
        }

        public void insert_columns_other(myExcel.Worksheet ws, string ItemCode, int c_setup, int row_insert)
        {
            DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                    new string[] { txtItemCode.Text, sheet, cb_Type.SelectedItem.ToString() }));
            if (cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, sheet, "MASS" }));
            }

            if (cb_Type.SelectedItem.ToString() == "NPI" || cb_Type.SelectedItem.ToString() == "LQ")
            {
                dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, sheet, "NPI" }));
            }

            if (dt_spec.Rows.Count > 0)
            {
                int c_format = int.Parse(dt_spec.Rows[0]["Count_sample"].ToString());
                if (c_setup > c_format)
                {
                    myExcel.Range cell_end = null;
                    int k_offset = -2;
                    if (cb_Type.SelectedItem.ToString() == "NPI")
                    {
                        k_offset = -1;
                    }

                    if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST" || sheet == "SHEAR_TEST")
                    {
                        foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                        {
                            if (spec_region != "")
                            {
                                string[] str_location = spec_region.Split('+');
                                myExcel.Range cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[1]),
                                    int.Parse(str_location[1].Split(';')[2]) + 1];
                                if (cell_end == null)
                                    cell_end = cell_Pic.Offset[row_insert, c_setup - 1];
                                double width_col = cell_Pic.ColumnWidth;
                                for (int i = 0; i < c_setup - c_format; i++)
                                {
                                    myExcel.Range cell_begin = cell_Pic.Offset[k_offset, c_format + i];
                                    string col_name = cell_begin.AddressLocal.Split('$')[2];
                                    int col = cell_begin.Column;
                                    //ws.Columns[int.Parse(col_name) + i - 1].ColumnWidth = width_col;
                                    ws.Columns[col].ColumnWidth = width_col;
                                    double size_text = cell_Pic.Offset[k_offset, c_format - 1].Font.Size;
                                    cell_begin.Value = "Sample " + (c_format + i + 1).ToString();
                                    for (int j = 0; j < 13; j++)
                                    {
                                        cell_begin.Offset[j, 0].Borders.LineStyle = myExcel.XlLineStyle.xlContinuous;
                                        cell_begin.Offset[j, 0].HorizontalAlignment = myExcel.XlHAlign.xlHAlignCenter;
                                        cell_begin.Offset[j, 0].VerticalAlignment = myExcel.XlHAlign.xlHAlignCenter;
                                        cell_begin.Offset[j, 0].Font.Size = size_text;
                                    }
                                    //myExcel.Range rgn_from = ws.Range[cell_begin.Offset[0, -1], cell_begin.Offset[row_insert, -1]];
                                    //myExcel.Range rgn_to = ws.Range[cell_begin, cell_begin.Offset[row_insert, 0]];
                                    //rgn_from.Copy(rgn_to);
                                    //cell_begin.Value = "Sample " + (c_format + i + 1).ToString();
                                }
                            }
                        }

                        if (cell_end != null)
                        {
                            ws.PageSetup.PrintArea = "$A$1:" + cell_end.AddressLocal;
                        }
                    }
                    else if (sheet == "CROSS_SECTION")
                    {
                        int row_pic = int.Parse(dt_spec.Rows[0]["Location"].ToString().Split('+')[1].Split('_')[0]
                            .Split(';')[2]);
                        int col_pic = int.Parse(dt_spec.Rows[0]["Location"].ToString().Split('+')[0].Split(';')[1]);
                        myExcel.Range cell_pic = ws.Cells[row_pic, col_pic];
                        if (cell_end == null)
                            cell_end = cell_pic.Offset[row_insert, c_setup - 1];
                        double width_col = cell_pic.ColumnWidth;
                        for (int i = 0; i < c_setup - c_format; i++)
                        {
                            myExcel.Range cell_begin = cell_pic.Offset[k_offset, c_format + i];
                            string col_name = cell_begin.AddressLocal.Split('$')[2];
                            ws.Columns[int.Parse(col_name) + i - 1].ColumnWidth = width_col;
                            //cell_begin.Value = "Sample " + (c_format + i + 1).ToString();
                            //for (int j = 0; j < row_insert + 12; j++)
                            //{
                            //    cell_begin.Offset[j, 0].Borders.LineStyle = myExcel.XlLineStyle.xlContinuous;
                            //    cell_begin.Offset[j, 0].HorizontalAlignment = myExcel.XlHAlign.xlHAlignCenter;
                            //    cell_begin.Offset[j, 0].VerticalAlignment = myExcel.XlHAlign.xlHAlignCenter;
                            //    cell_begin.Offset[j, 0].Font.Size = 24;
                            //}
                            myExcel.Range rgn_from =
                                ws.Range[cell_begin.Offset[0, -1], cell_begin.Offset[row_insert, -1]];
                            myExcel.Range rgn_to = ws.Range[cell_begin, cell_begin.Offset[row_insert, 0]];
                            rgn_from.Copy(rgn_to);
                            cell_begin.Value = "Sample " + (c_format + i + 1).ToString();
                        }

                        if (cell_end != null)
                        {
                            ws.PageSetup.PrintArea = "$A$1:" + cell_end.AddressLocal;
                        }
                    }
                    else if (sheet.Contains("ON_PRODUCT"))
                    {
                        foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                        {
                            if (spec_region != "")
                            {
                                string[] str_location = spec_region.Split('+');
                                myExcel.Range cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[2]),
                                    int.Parse(str_location[1].Split(';')[3]) + 1];
                                double size_text = cell_Pic.Offset[k_offset, c_format - 1].Font.Size;
                                if (cell_end == null)
                                    cell_end = cell_Pic.Offset[row_insert, c_setup - 1];
                                double width_col = cell_Pic.ColumnWidth;
                                for (int i = 0; i < c_setup - c_format; i++)
                                {
                                    myExcel.Range cell_begin = cell_Pic.Offset[k_offset, c_format + i];
                                    string col_name = cell_begin.AddressLocal.Split('$')[2];
                                    int col = cell_begin.Column;
                                    ws.Columns[col].ColumnWidth = width_col;

                                    cell_begin.Value = "Sample " + (c_format + i + 1).ToString();
                                    for (int j = 0; j < 8; j++)
                                    {
                                        cell_begin.Offset[j, 0].Borders.LineStyle = myExcel.XlLineStyle.xlContinuous;
                                        cell_begin.Offset[j, 0].HorizontalAlignment = myExcel.XlHAlign.xlHAlignCenter;
                                        cell_begin.Offset[j, 0].VerticalAlignment = myExcel.XlHAlign.xlHAlignCenter;
                                        cell_begin.Offset[j, 0].Font.Size = size_text;
                                    }
                                }
                            }
                        }

                        if (cell_end != null)
                        {
                            ws.PageSetup.PrintArea = "$A$1:" + cell_end.AddressLocal;
                        }
                    }
                    else if (sheet == "GAP_CONNECTOR")
                    {
                        int col_begin = int.Parse(dt_spec.Rows[0]["Location"].ToString().Split('+')[0]);
                        for (int i = 0; i < c_setup - c_format; i++)
                        {
                            ws.Columns[col_begin + c_format - 1].Insert();
                        }

                        foreach (string reg in dt_spec.Rows[0]["Location"].ToString().Split('+')[1].Split('_'))
                        {
                            if (reg != "")
                            {
                                foreach (string pos in reg.Split('^'))
                                {
                                    if (pos != "")
                                    {
                                        int row_begin = int.Parse(pos.Split(';')[1]);
                                        int count_r_offset = int.Parse(pos.Split(';')[2]);
                                        myExcel.Range cell_Pic = ws.Cells[row_begin, col_begin];
                                        double width_col = cell_Pic.ColumnWidth;
                                        //cell_Pic.Offset[-1, c_format-1].Value = "Sample " + c_format.ToString();
                                        //int size_text = cell_Pic.Offset[-1, c_format - 1].Font.Size;
                                        for (int i = -1; i < c_setup - c_format; i++)
                                        {
                                            myExcel.Range cell_begin = cell_Pic.Offset[k_offset, c_format + i];
                                            int col = cell_begin.Column;
                                            ws.Columns[col].ColumnWidth = width_col;
                                            cell_begin.Value = "Sample " + (c_format + i + 1).ToString();
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                        {
                            if (spec_region != "")
                            {
                                string[] str_location = spec_region.Split('+');
                                myExcel.Range cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[1]),
                                    int.Parse(str_location[1].Split(';')[2]) + 1];
                                if (cell_end == null)
                                    cell_end = cell_Pic.Offset[row_insert, c_setup - 1];
                                double width_col = cell_Pic.ColumnWidth;
                                for (int i = 0; i < c_setup - c_format; i++)
                                {
                                    myExcel.Range cell_begin = cell_Pic.Offset[k_offset, c_format + i];
                                    string col_name = cell_begin.AddressLocal.Split('$')[2];
                                    int col = cell_begin.Column;
                                    //ws.Columns[int.Parse(col_name) + i - 1].ColumnWidth = width_col;
                                    ws.Columns[col].ColumnWidth = width_col;
                                    double size_text = cell_Pic.Offset[k_offset, c_format - 1].Font.Size;
                                    cell_begin.Value = "Sample " + (c_format + i + 1).ToString();
                                    for (int j = 0; j < 8; j++)
                                    {
                                        cell_begin.Offset[j, 0].Borders.LineStyle = myExcel.XlLineStyle.xlContinuous;
                                        cell_begin.Offset[j, 0].HorizontalAlignment = myExcel.XlHAlign.xlHAlignCenter;
                                        cell_begin.Offset[j, 0].VerticalAlignment = myExcel.XlHAlign.xlHAlignCenter;
                                        cell_begin.Offset[j, 0].Font.Size = size_text;
                                    }
                                }
                            }
                        }

                        if (cell_end != null)
                        {
                            ws.PageSetup.PrintArea = "$A$1:" + cell_end.AddressLocal;
                        }
                    }
                }
            }
        }


        public void complete_sheet_other_old(myExcel.Worksheet ws, string ItemCode, int c_setup, int row_delete)
        {
            DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                    new string[] { ItemCode, sheet, cb_Type.SelectedItem.ToString() }));

            if (dt_spec.Rows.Count > 0)
            {
                int c_format = int.Parse(dt_spec.Rows[0]["Count_sample"].ToString());
                if (c_setup < c_format)
                {
                    myExcel.Range cell_begin = find_cell(ws, "SAMPLE" + (c_setup + 1).ToString());

                    for (int j = 0; j < c_format - c_setup; j++)
                    {
                        for (int i = 0; i < row_delete; i++)
                        {
                            cell_begin.Offset[i, j].Value = "";
                            cell_begin.Offset[i, j].Borders.LineStyle = myExcel.XlLineStyle.xlLineStyleNone;
                        }
                    }
                    // myExcel.Range cell_end = cell_begin.Offset[row_delete, 0];
                }
            }
        }

        public void complete_sheet_other(myExcel.Worksheet ws, string ItemCode, int c_setup, int r_ofset_del)
        {
            DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                    new string[] { txtItemCode.Text, sheet, cb_Type.SelectedItem.ToString() }));
            if (cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, sheet, "MASS" }));
            }

            if (cb_Type.SelectedItem.ToString() == "NPI" || cb_Type.SelectedItem.ToString() == "LQ")
            {
                dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, sheet, "NPI" }));
            }

            if (dt_spec.Rows.Count > 0)
            {
                int c_format = int.Parse(dt_spec.Rows[0]["Count_sample"].ToString());
                if (c_setup < c_format)
                {
                    //myExcel.Range cell_begin = find_cell(ws, "SAMPLE" + (c_setup + 1).ToString()); 
                    //for (int j = 0; j < c_format - c_setup; j++)
                    //{
                    //    for (int i = 0; i < row_delete; i++)
                    //    {
                    //        cell_begin.Offset[i, j].Value = "";
                    //        cell_begin.Offset[i, j].Borders.LineStyle = myExcel.XlLineStyle.xlLineStyleNone;
                    //    }
                    //}

                    if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST" || sheet == "SHEAR_TEST")
                    {
                        foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                        {
                            if (spec_region != "")
                            {
                                string[] str_location = spec_region.Split('+');
                                myExcel.Range cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[1]),
                                    int.Parse(str_location[1].Split(';')[2]) + 1];

                                double width_col = cell_Pic.ColumnWidth;
                                for (int j = 0; j < c_format - c_setup; j++)
                                {
                                    myExcel.Range cell_begin = cell_Pic.Offset[-2, c_setup + j];
                                    for (int i = 0; i < r_ofset_del; i++)
                                    {
                                        cell_begin.Offset[i, 0].Value = "";
                                        cell_begin.Offset[i, 0].Borders.LineStyle = myExcel.XlLineStyle.xlLineStyleNone;
                                    }
                                }
                            }
                        }
                    }
                    else if (sheet == "CROSS_SECTION")
                    {
                        int row_pic = int.Parse(dt_spec.Rows[0]["Location"].ToString().Split('+')[1].Split('_')[0]
                            .Split(';')[2]);
                        int col_pic = int.Parse(dt_spec.Rows[0]["Location"].ToString().Split('+')[0].Split(';')[1]);
                        myExcel.Range cell_pic = ws.Cells[row_pic, col_pic];

                        for (int j = 0; j < c_format - c_setup; j++)
                        {
                            myExcel.Range cell_begin = cell_pic.Offset[-2, c_setup + j];
                            for (int i = 0; i < r_ofset_del; i++)
                            {
                                cell_begin.Offset[i, 0].Value = "";
                                cell_begin.Offset[i, 0].Borders.LineStyle = myExcel.XlLineStyle.xlLineStyleNone;
                            }
                        }
                    }
                    else if (sheet.Contains("ON_PRODUCT"))
                    {
                        foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                        {
                            if (spec_region != "")
                            {
                                string[] str_location = spec_region.Split('+');
                                myExcel.Range cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[2]),
                                    int.Parse(str_location[1].Split(';')[3]) + 1];

                                double width_col = cell_Pic.ColumnWidth;
                                for (int j = 0; j < c_format - c_setup; j++)
                                {
                                    myExcel.Range cell_begin = cell_Pic.Offset[-2, c_setup + j];
                                    for (int i = 0; i < r_ofset_del; i++)
                                    {
                                        cell_begin.Offset[i, 0].Value = "";
                                        cell_begin.Offset[i, 0].Borders.LineStyle = myExcel.XlLineStyle.xlLineStyleNone;
                                    }
                                }
                            }
                        }
                    }
                    else if (sheet == "GAP_CONNECTOR")
                    {
                        int col_begin = int.Parse(dt_spec.Rows[0]["Location"].ToString().Split('+')[0]);
                        for (int i = 0; i < c_format - c_setup; i++)
                        {
                            // ws.Columns.Delete(ws.Columns[col_begin + c_setup + i]); 
                            // ws.DeleteCells(ws.Cells["B2"], DeleteMode.EntireRow);
                            ws.Columns[col_begin + c_setup].Delete();
                        }
                    }
                    else
                    {
                        foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                        {
                            if (spec_region != "")
                            {
                                string[] str_location = spec_region.Split('+');
                                myExcel.Range cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[1]),
                                    int.Parse(str_location[1].Split(';')[2]) + 1];

                                for (int j = 0; j < c_format - c_setup; j++)
                                {
                                    myExcel.Range cell_begin = cell_Pic.Offset[-2, c_setup + j];
                                    for (int i = 0; i < r_ofset_del; i++)
                                    {
                                        cell_begin.Offset[i, 0].Value = "";
                                        cell_begin.Offset[i, 0].Borders.LineStyle = myExcel.XlLineStyle.xlLineStyleNone;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        public void export_excel_unmating(string ItemCode, string LotNo, string file_format, string mysheet,
            DataTable Data_tbl, DataTable dt_spec)
        {
            myExcel.Workbook curr_wrkbook = null;
            if (cb_Type.SelectedItem.ToString() == "NPI" || cb_Type.SelectedItem.ToString() == "LQ")
            {
                curr_wrkbook = create_export_wrk(file_format, mysheet);
            }
            else if (cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                curr_wrkbook = TDMK_Code.open_excel_file(file_format, "", "");
            }

            if (curr_wrkbook != null)
            {
                myExcel.Worksheet ws = curr_wrkbook.Sheets[1];

                if (cb_Type.SelectedItem.ToString().Contains("MASS") || cb_Type.SelectedItem.ToString() == "Other")
                {
                    myExcel.Range cur_rgn_type = find_cell(ws,
                        cb_Type.SelectedItem.ToString().Replace("MASS(", "").Replace(")", ""));
                    if (cur_rgn_type != null)
                        cur_rgn_type.Value = " v" + cur_rgn_type.Value.ToString();

                    string str_infor = Data_tbl.Rows[0]["Sheet"].ToString().Replace(sheet, "").Replace("/", "");
                    if (str_infor.Contains("_"))
                    {
                        string itemname = str_infor.Split('_')[0];
                        string line = str_infor.Split('_')[1];
                        string ca = str_infor.Split('_')[2];
                        string date = str_infor.Split('_')[3];
                        string worker = str_infor.Split('_')[4];
                        List<string> lst_infor = new List<string> { itemname, ItemCode, LotNo, line, ca, date, worker };
                        export_info_mass(ws, lst_infor);
                    }
                }

                int count_reg = dt_spec.Rows[0]["Location"].ToString().Split('_').Length - 1;
                int r_end = 10 * count_reg;
                insert_columns_other(ws, ItemCode, int.Parse(txt_qty.Text), r_end);

                int reg = 0;
                List<DataTable> lst_Table = new List<DataTable> { };
                Get_ListTable(-1, Data_tbl, new string[] { "Region" }, ref lst_Table, "Data");
                string[] region_data =
                    Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();
                foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                {
                    if (spec_region != "")
                    {
                        if (reg < lst_Table.Count)
                        {
                            DataTable dt_region = lst_Table[reg];
                            string[] str_location = spec_region.Split('+');
                            myExcel.Range cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[1]),
                                int.Parse(str_location[1].Split(';')[2]) + 1];
                            myExcel.Range cell_graph = ws.Cells[int.Parse(str_location[2].Split(';')[1]),
                                int.Parse(str_location[1].Split(';')[2]) + 1];
                            myExcel.Range cell_data = ws.Cells[int.Parse(str_location[3].Split(';')[1]),
                                int.Parse(str_location[1].Split(';')[2]) + 1];
                            // int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
                            int count_sample = int.Parse(txt_qty.Text);
                            Export_DatatableImage_Excel_unmating(dt_region, "Image", ws, cell_Pic, false, count_sample);
                            Export_DatatableImage_Excel_Graph(dt_region, "Graph", ws, cell_graph, false, count_sample);

                            List<string> lst_data =
                                dt_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                            List<string> lst_judge = new List<string> { };

                            string type = dt_spec.Rows[0]["Location"].ToString().Split('+')[0];
                            if (type == "A")
                            {
                                for (int i = 0; i < count_sample; i++)
                                {
                                    if (i < dt_region.Rows.Count)
                                    {
                                        cell_data.Offset[0, i].Value = lst_data[i];
                                    }
                                }
                            }
                            else if (type == "B")
                            {
                                for (int i = 0; i < count_sample; i++)
                                {
                                    if (i < dt_region.Rows.Count)
                                    {
                                        cell_data.Offset[1, i].Value = lst_data[i];
                                    }
                                }
                            }

                            for (int i = 1; i < 13; i++)
                            {
                                if (myCode.checkDBNull(cell_data.Offset[i, -1].Value).Replace(" ", "").ToUpper()
                                    .Contains("MINFORCE"))
                                {
                                    myExcel.Range cell_min = cell_data.Offset[i, 0];
                                    myExcel.Range cell_max = cell_data.Offset[i + 1, 0];
                                    myExcel.Range cell_ave = cell_data.Offset[i + 2, 0];
                                    int c_offset = count_sample;
                                    cell_min.FormulaR1C1 = "=MIN(R[-" + i.ToString() + "]C[0]:R[-" + i.ToString() +
                                                           "]C[" + (c_offset - 1).ToString() + "])";
                                    cell_max.FormulaR1C1 = "=MAX(R[-" + (i + 1).ToString() + "]C[0]:R[-" +
                                                           (i + 1).ToString() + "]C[" + (c_offset - 1).ToString() +
                                                           "])";
                                    cell_ave.FormulaR1C1 = "=AVERAGE(R[-" + (i + 2).ToString() + "]C[0]:R[-" +
                                                           (i + 2).ToString() + "]C[" + (c_offset - 1).ToString() +
                                                           "])";
                                    break;
                                }
                            }

                            for (int i = 0; i < count_sample; i++)
                            {
                                if (i < dt_region.Rows.Count)
                                {
                                    cell_data.Offset[0, i].Value = lst_data[i];
                                    int ID = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                    if (dgv_Analysis.Rows[ID - 1].Cells["Data"].Style.BackColor == Color.Red)
                                    {
                                        lst_judge.Add("Fail");
                                        judge_all = false;
                                    }
                                    else
                                    {
                                        lst_judge.Add("Pass");
                                    }
                                }
                            }

                            for (int i = 0; i < count_sample; i++)
                            {
                                if (i < lst_judge.Count)
                                    cell_data.Offset[2, i].Value = lst_judge[i];
                            }
                        }

                        reg++;
                    }
                }


                int r_offset_del = 8;
                complete_sheet_other(ws, ItemCode, int.Parse(txt_qty.Text), r_offset_del);
                if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                {
                    myExcel.Range cur_rgn_judge = find_cell(ws, "JUDGEMENT");
                    if (cur_rgn_judge != null)
                    {
                        if (judge_all)
                        {
                            cur_rgn_judge.Offset[1, 0].Value = "OK";
                            ws.Cells[1, 1].Value = "OK";
                        }
                        else
                        {
                            cur_rgn_judge.Offset[1, 0].Value = "NG";
                            ws.Cells[1, 1].Value = "NG";
                        }
                    }
                }

                MessageBox.Show(new Form { TopMost = true }, "Xuất dữ liệu thành công!", "Thông báo");
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Không tìm thấy sheet: " + sheet, "Thông báo");
            }
        }

        public void export_excel_coupon_old(string ItemCode, string LotNo, string file_format, string mysheet,
            DataTable Data_tbl, DataTable dt_spec)
        {
            myExcel.Workbook curr_wrkbook = null;
            if (cb_Type.SelectedItem.ToString() == "NPI" || cb_Type.SelectedItem.ToString() == "Other")
            {
                curr_wrkbook = create_export_wrk(file_format, mysheet);
            }
            else if (cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                curr_wrkbook = TDMK_Code.open_excel_file(file_format, "", "");
            }

            if (curr_wrkbook != null)
            {
                myExcel.Worksheet ws = curr_wrkbook.Sheets[1];

                //int st = 1;
                //foreach (DataRow dr in Data_tbl.Rows)
                //{
                //    dr["ID"] = st;
                //    st++;
                //}

                if (cb_Type.SelectedItem.ToString().Contains("MASS") || cb_Type.SelectedItem.ToString() == "Other")
                {
                    myExcel.Range cur_rgn_type = find_cell(ws,
                        cb_Type.SelectedItem.ToString().Replace("MASS(", "").Replace(")", ""));
                    if (cur_rgn_type != null)
                        cur_rgn_type.Value = " v" + cur_rgn_type.Value.ToString();


                    string str_infor = Data_tbl.Rows[0]["Sheet"].ToString().Replace(sheet, "").Replace("/", "");
                    if (str_infor.Contains("_"))
                    {
                        string itemname = str_infor.Split('_')[0];
                        string line = str_infor.Split('_')[1];
                        string ca = str_infor.Split('_')[2];
                        string date = str_infor.Split('_')[3];
                        string worker = str_infor.Split('_')[4];
                        List<string> lst_infor = new List<string> { itemname, ItemCode, LotNo, line, ca, date, worker };
                        export_info_mass(ws, lst_infor);
                    }
                }

                int count_reg = dt_spec.Rows[0]["Location"].ToString().Split('_').Length - 1;
                int r_end = 10 * count_reg;
                insert_columns_other(ws, ItemCode, int.Parse(txt_qty.Text), r_end);

                //int reg = 0;
                List<DataTable> lst_Table = new List<DataTable> { };
                Get_ListTable(-1, Data_tbl, new string[] { "Region" }, ref lst_Table, "Data");
                string[] region_data =
                    Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();
                foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                {
                    if (spec_region != "")
                    {
                        //if (reg < lst_Table.Count)
                        //{

                        // DataTable dt_region = lst_Table[reg];

                        List<string> lst_region = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region"))
                            .Distinct().ToList();

                        string cpn = "";

                        foreach (string r in lst_region)
                        {
                            string key1 = r.Split('_')[0].Replace(" ", "").ToUpper();
                            string key2 = r.Split('_')[1].Replace(" ", "").ToUpper();

                            //if (reg.Replace("Component", "Tape").Replace(" ", "").ToUpper().Contains(key1) && reg.Replace("Component", "Tape").Replace(" ", "").ToUpper().Contains(key2))
                            //{
                            //    cpn = r;
                            //    break;
                            //}
                        }

                        DataTable dt_region = Data_tbl.AsEnumerable().Where(s => s.Field<string>("Region") == cpn)
                            .CopyToDataTable();

                        string[] str_location = spec_region.Split('+');
                        myExcel.Range cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[1]),
                            int.Parse(str_location[1].Split(';')[2]) + 1];
                        myExcel.Range cell_graph = ws.Cells[int.Parse(str_location[2].Split(';')[1]),
                            int.Parse(str_location[1].Split(';')[2]) + 1];
                        myExcel.Range cell_data = ws.Cells[int.Parse(str_location[3].Split(';')[1]),
                            int.Parse(str_location[1].Split(';')[2]) + 1];
                        // int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
                        int count_sample = int.Parse(txt_qty.Text);
                        Export_DatatableImage_Excel(dt_region, "Image", ws, cell_Pic, false, count_sample);
                        Export_DatatableImage_Excel_Graph(dt_region, "Graph", ws, cell_graph, false, count_sample);

                        List<string> lst_data = dt_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                        List<string> lst_judge = new List<string> { };

                        string type = dt_spec.Rows[0]["Location"].ToString().Split('+')[0];
                        if (type == "A")
                        {
                            for (int i = 0; i < count_sample; i++)
                            {
                                if (i < dt_region.Rows.Count)
                                {
                                    cell_data.Offset[0, i].Value = lst_data[i];
                                }
                            }
                        }
                        else if (type == "B")
                        {
                            for (int i = 0; i < count_sample; i++)
                            {
                                if (i < dt_region.Rows.Count)
                                {
                                    cell_data.Offset[1, i].Value = lst_data[i];
                                }
                            }
                        }

                        for (int i = 1; i < 13; i++)
                        {
                            if (myCode.checkDBNull(cell_data.Offset[i, -1].Value).Replace(" ", "").ToUpper()
                                .Contains("MINFORCE"))
                            {
                                myExcel.Range cell_min = cell_data.Offset[i, 0];
                                myExcel.Range cell_max = cell_data.Offset[i + 1, 0];
                                myExcel.Range cell_ave = cell_data.Offset[i + 2, 0];
                                int c_offset = count_sample;
                                cell_min.FormulaR1C1 = "=MIN(R[-" + i.ToString() + "]C[0]:R[-" + i.ToString() + "]C[" +
                                                       c_offset.ToString() + "])";
                                cell_max.FormulaR1C1 = "=MAX(R[-" + (i + 1).ToString() + "]C[0]:R[-" +
                                                       (i + 1).ToString() + "]C[" + c_offset.ToString() + "])";
                                cell_ave.FormulaR1C1 = "=AVERAGE(R[-" + (i + 2).ToString() + "]C[0]:R[-" +
                                                       (i + 2).ToString() + "]C[" + c_offset.ToString() + "])";
                                break;
                            }
                        }

                        for (int i = 0; i < count_sample; i++)
                        {
                            if (i < dt_region.Rows.Count)
                            {
                                cell_data.Offset[0, i].Value = lst_data[i];
                                int ID = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                if (dgv_Analysis.Rows[ID - 1].Cells["Data"].Style.BackColor == Color.Red)
                                {
                                    lst_judge.Add("Fail");
                                    judge_all = false;
                                }
                                else
                                {
                                    lst_judge.Add("Pass");
                                }
                            }
                        }

                        for (int i = 0; i < count_sample; i++)
                        {
                            cell_data.Offset[2, i].Value = lst_judge[i];
                        }
                    }
                    // reg++;
                    //}
                }


                int r_offset_del = 7;
                complete_sheet_other(ws, ItemCode, int.Parse(txt_qty.Text), r_offset_del);
                MessageBox.Show(new Form { TopMost = true }, "Xuất dữ liệu thành công!", "Thông báo");
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Không tìm thấy sheet: " + sheet, "Thông báo");
            }
        }

        public void export_excel_coupon(string ItemCode, string LotNo, string file_format, string mysheet,
            DataTable Data_all, DataTable dt_spec)
        {
            myExcel.Workbook curr_wrkbook = null;
            if (cb_Type.SelectedItem.ToString() == "NPI" || cb_Type.SelectedItem.ToString() == "Other")
            {
                curr_wrkbook = create_export_wrk(file_format, mysheet);
            }
            else if (cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                curr_wrkbook = TDMK_Code.open_excel_file(file_format, "", "");
            }

            if (curr_wrkbook != null)
            {
                myExcel.Worksheet ws = curr_wrkbook.Sheets[1];
                int count_sample = int.Parse(txt_qty.Text);
                if (dt_spec.Rows.Count > 0)
                {
                    List<string> lst_region = Data_all.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct()
                        .ToList();
                    string[] region = dt_spec.Rows[0]["Location"].ToString().Split('_');

                    if (cb_Type.SelectedItem.ToString().Contains("MASS") || cb_Type.SelectedItem.ToString() == "Other")
                    {
                        myExcel.Range cur_rgn_type = find_cell(ws,
                            cb_Type.SelectedItem.ToString().Replace("MASS(", "").Replace(")", ""));
                        if (cur_rgn_type != null)
                            cur_rgn_type.Value = " v" + cur_rgn_type.Value.ToString();


                        string str_infor = Data_all.Rows[0]["Sheet"].ToString().Replace(sheet, "").Replace("/", "");
                        if (str_infor.Contains("_"))
                        {
                            string itemname = str_infor.Split('_')[0];
                            string line = str_infor.Split('_')[1];
                            string ca = str_infor.Split('_')[2];
                            string date = str_infor.Split('_')[3];
                            string worker = str_infor.Split('_')[4];
                            List<string> lst_infor = new List<string>
                                { itemname, ItemCode, LotNo, line, ca, date, worker };
                            export_info_mass(ws, lst_infor);
                        }

                        for (int i = 1; i < 20; i++)
                        {
                            for (int j = 1; j < 4; j++)
                            {
                                if (myCode.checkDBNull(ws.Cells[i, j].Value).ToUpper().Replace(" ", "")
                                    .Contains("ITEM-LOT"))
                                {
                                    ws.Cells[i, j].Offset[0, 1].Value = ItemCode + "-" + LotNo;
                                    break;
                                }
                            }
                        }
                    }

                    int r_end = 0;
                    if (region.Length == 2)
                    {
                        r_end = 25;
                    }
                    else if (region.Length == 3)
                    {
                        r_end = 40;
                    }
                    else if (region.Length == 4)
                    {
                        r_end = 52;
                    }

                    insert_columns_other(ws, ItemCode, int.Parse(txt_qty.Text), r_end);

                    foreach (string reg in region)
                    {
                        if (reg != "")
                        {
                            string[] str_location = reg.Split('+');
                            string cpn = "";

                            foreach (string r in lst_region)
                            {
                                bool chk = true;
                                foreach (string key in r.Split('_'))
                                {
                                    if (!reg.Replace(" ", "").ToUpper().Contains(key.Replace(" ", "").ToUpper()))
                                    {
                                        chk = false;
                                        break;
                                    }
                                }

                                if (chk)
                                {
                                    cpn = r;
                                    break;
                                }

                                //string key1 = r.Split('_')[0].Replace(" ", "").ToUpper();
                                //string key2 = r.Split('_')[1].Replace(" ", "").ToUpper();

                                //if (reg.Replace(" ", "").ToUpper().Contains(key1) && reg.Replace(" ", "").ToUpper().Contains(key2))
                                //{
                                //    cpn = r;
                                //    break;
                                //}
                            }


                            if (cpn != "")
                            {
                                DataTable dt_region = Data_all.AsEnumerable()
                                    .Where(s => s.Field<string>("Region") == cpn).CopyToDataTable();
                                if (dt_region.Rows.Count > 0)
                                {
                                    //if (type == "A")
                                    //{

                                    //myExcel.Range cell_Pic = null;
                                    //myExcel.Range cell_graph = null;
                                    //myExcel.Range cell_data = null;
                                    List<string> lst_data = dt_region.AsEnumerable()
                                        .Select(x => x.Field<string>("Data")).ToList();
                                    List<string> lst_judge = new List<string> { };


                                    myExcel.Range cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[2]),
                                        int.Parse(str_location[1].Split(';')[3]) + 1];
                                    myExcel.Range cell_graph = ws.Cells[int.Parse(str_location[2].Split(';')[1]),
                                        int.Parse(str_location[1].Split(';')[3]) + 1];
                                    myExcel.Range cell_data = ws.Cells[int.Parse(str_location[3].Split(';')[1]),
                                        int.Parse(str_location[1].Split(';')[3]) + 1];

                                    Export_DatatableImage_Excel(dt_region, "Image", ws, cell_Pic, false, count_sample);
                                    Export_DatatableImage_Excel_Graph(dt_region, "Graph", ws, cell_graph, false,
                                        count_sample);


                                    if (cb_Type.SelectedItem.ToString() == "NPI")
                                    {
                                        // string[] str_location = spec_region.Split('+');
                                        //cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[1]), int.Parse(str_location[1].Split(';')[2]) + 1];
                                        //cell_graph = ws.Cells[int.Parse(str_location[2].Split(';')[1]), int.Parse(str_location[1].Split(';')[2]) + 1];
                                        //cell_data = ws.Cells[int.Parse(str_location[3].Split(';')[1]), int.Parse(str_location[1].Split(';')[2]) + 1];

                                        //Export_DatatableImage_Excel(dt_region, "Image", ws, cell_Pic, false, count_sample);
                                        //Export_DatatableImage_Excel_Graph(dt_region, "Graph", ws, cell_graph, false, count_sample);


                                        for (int i = 0; i < count_sample; i++)
                                        {
                                            if (i < dt_region.Rows.Count)
                                            {
                                                cell_data.Offset[0, i].Value = lst_data[i];
                                            }
                                        }
                                    }
                                    else
                                    {
                                        //cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[2]), int.Parse(str_location[1].Split(';')[3]) + 1];
                                        //cell_graph = ws.Cells[int.Parse(str_location[2].Split(';')[1]), int.Parse(str_location[1].Split(';')[3]) + 1];
                                        //cell_data = ws.Cells[int.Parse(str_location[3].Split(';')[1]), int.Parse(str_location[1].Split(';')[3]) + 1];

                                        //Export_DatatableImage_Excel(dt_region, "Image", ws, cell_Pic, false, count_sample);
                                        //Export_DatatableImage_Excel_Graph(dt_region, "Graph", ws, cell_graph, false, count_sample);


                                        for (int i = 0; i < count_sample; i++)
                                        {
                                            if (i < dt_region.Rows.Count)
                                            {
                                                cell_data.Offset[0, i].Value =
                                                    lst_data[i].Split('_')[0].Replace("Max:", "");
                                                cell_data.Offset[1, i].Value =
                                                    lst_data[i].Split('_')[1].Replace("Average:", "");
                                            }
                                        }
                                    }

                                    //}
                                    //else if (type == "B")
                                    //{
                                    //    for (int i = 0; i < count_sample; i++)
                                    //    {
                                    //        if (i < dt_region.Rows.Count)
                                    //        {
                                    //            cell_data.Offset[1, i].Value = lst_data[i];
                                    //        }
                                    //    }
                                    //}

                                    for (int i = 1; i < 13; i++)
                                    {
                                        if (myCode.checkDBNull(cell_data.Offset[i, -1].Value).Replace(" ", "").ToUpper()
                                            .Contains("MINFORCE"))
                                        {
                                            myExcel.Range cell_min = cell_data.Offset[i, 0];
                                            myExcel.Range cell_max = cell_data.Offset[i + 1, 0];
                                            myExcel.Range cell_ave = cell_data.Offset[i + 2, 0];
                                            int c_offset = count_sample;
                                            cell_min.FormulaR1C1 = "=MIN(R[-" + i.ToString() + "]C[0]:R[-" +
                                                                   i.ToString() + "]C[" + (c_offset - 1).ToString() +
                                                                   "])";
                                            cell_max.FormulaR1C1 = "=MAX(R[-" + (i + 1).ToString() + "]C[0]:R[-" +
                                                                   (i + 1).ToString() + "]C[" +
                                                                   (c_offset - 1).ToString() + "])";
                                            cell_ave.FormulaR1C1 = "=AVERAGE(R[-" + (i + 2).ToString() + "]C[0]:R[-" +
                                                                   (i + 2).ToString() + "]C[" +
                                                                   (c_offset - 1).ToString() + "])";
                                            break;
                                        }
                                    }


                                    if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                                    {
                                        for (int i = 0; i < count_sample; i++)
                                        {
                                            if (i < dt_region.Rows.Count)
                                            {
                                                int ID = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                                if (dgv_Analysis.Rows[ID - 1].Cells["Data"].Style.BackColor ==
                                                    Color.Red)
                                                {
                                                    lst_judge.Add("Fail");
                                                    judge_all = false;
                                                }
                                                else
                                                {
                                                    lst_judge.Add("Pass");
                                                }
                                            }
                                        }

                                        for (int i = 0; i < count_sample; i++)
                                        {
                                            cell_data.Offset[3, i].Value = lst_judge[i];
                                        }
                                    }
                                }
                            }
                        }
                    }


                    int r_offset_del = 8;
                    complete_sheet_other(ws, ItemCode, int.Parse(txt_qty.Text), r_offset_del);


                    if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                    {
                        myExcel.Range cur_rgn_judge = find_cell(ws, "JUDGEMENT");
                        if (cur_rgn_judge != null)
                        {
                            if (judge_all)
                            {
                                cur_rgn_judge.Offset[1, 0].Value = "OK";
                                ws.Cells[1, 1].Value = "OK";
                            }
                            else
                            {
                                cur_rgn_judge.Offset[1, 0].Value = "NG";
                                ws.Cells[1, 1].Value = "NG";
                            }
                        }
                    }

                    MessageBox.Show(new Form { TopMost = true }, "Xuất dữ liệu thành công!", "Thông báo");
                }
                // }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "KHông tìm thấy sheet: " + sheet, "Thông báo");
            }
        }

        public void export_excel_peel_pull_shear(string ItemCode, string LotNo, string file_format, string mysheet,
            DataTable Data_tbl, DataTable dt_spec)
        {
            myExcel.Workbook curr_wrkbook = null;
            if (cb_Type.SelectedItem.ToString() == "NPI" || cb_Type.SelectedItem.ToString() == "Other")
            {
                curr_wrkbook = create_export_wrk(file_format, mysheet);
            }
            else if (cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                curr_wrkbook = TDMK_Code.open_excel_file(file_format, "", "");
            }

            if (curr_wrkbook != null)
            {
                myExcel.Worksheet ws = curr_wrkbook.Sheets[1];
                // ws.Cells.Interior.Color = myExcel.XlRgbColor.rgbWhite;
                ws.Cells.FormatConditions.Delete();
                //int st = 1;
                //foreach (DataRow dr in Data_tbl.Rows)
                //{
                //    dr["ID"] = st;
                //    st++;
                //}

                if (cb_Type.SelectedItem.ToString().Contains("MASS") || cb_Type.SelectedItem.ToString() == "Other")
                {
                    myExcel.Range cur_rgn_type = find_cell(ws,
                        cb_Type.SelectedItem.ToString().Replace("MASS(", "").Replace(")", ""));
                    if (cur_rgn_type != null)
                        cur_rgn_type.Value = " v" + cur_rgn_type.Value.ToString();

                    string str_infor = Data_tbl.Rows[0]["Sheet"].ToString().Replace(sheet, "").Replace("/", "");
                    if (str_infor.Contains("_"))
                    {
                        string itemname = str_infor.Split('_')[0];
                        string line = str_infor.Split('_')[1];
                        string ca = str_infor.Split('_')[2];
                        string date = str_infor.Split('_')[3];
                        string worker = str_infor.Split('_')[4];
                        List<string> lst_infor = new List<string> { itemname, ItemCode, LotNo, line, ca, date, worker };
                        export_info_mass(ws, lst_infor);
                    }
                }

                int reg = 0;
                List<DataTable> lst_Table = new List<DataTable> { };
                Get_ListTable(-1, Data_tbl, new string[] { "Region" }, ref lst_Table, "Data");
                string[] region_data =
                    Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();

                Dictionary<string, string> dic_mode = new Dictionary<string, string> { };
                dic_mode.Add("Mode 1#", "Mode 1: Solder joint crack");
                dic_mode.Add("Mode 2#", "Mode 2: Pad lift");
                dic_mode.Add("Mode 3#", "Mode 3: Solder joint lift");
                dic_mode.Add("Mode 4#", "Mode 4: Intermetallic break");
                dic_mode.Add("Mode 5#", "Mode 5: Component damage");
                dic_mode.Add("Mode 6#", "Mode 6: Component detached");
                dic_mode.Add("Mode 7#", "Mode 7: Flex torn");

                int r_end;
                if (lst_Table.Count == 1)
                {
                    r_end = 22;
                }
                else
                {
                    r_end = 42;
                }

                insert_columns_other(ws, ItemCode, int.Parse(txt_qty.Text), r_end);
                foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                {
                    if (spec_region != "")
                    {
                        if (reg < lst_Table.Count)
                        {
                            DataTable dt_region = lst_Table[reg];

                            string[] str_location = spec_region.Split('+');
                            myExcel.Range cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[1]),
                                int.Parse(str_location[1].Split(';')[2]) + 1];
                            myExcel.Range cell_graph = ws.Cells[int.Parse(str_location[2].Split(';')[1]),
                                int.Parse(str_location[1].Split(';')[2]) + 1];
                            myExcel.Range cell_data = ws.Cells[int.Parse(str_location[3].Split(';')[1]),
                                int.Parse(str_location[1].Split(';')[2]) + 1];
                            //int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
                            int count_sample = int.Parse(txt_qty.Text);
                            Export_DatatableImage_Excel(dt_region, "Image", ws, cell_Pic, false, count_sample);
                            Export_DatatableImage_Excel_Graph(dt_region, "Graph", ws, cell_graph, false, count_sample);

                            List<string> lst_data =
                                dt_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                            List<string> lst_judge = new List<string> { };

                            for (int i = 0; i < count_sample; i++)
                            {
                                if (i < dt_region.Rows.Count)
                                {
                                    //cell_data.Offset[0, i].Value = lst_data[i];
                                    if (myCode.IsNumeric(lst_data[i]))
                                    {
                                        cell_data.Offset[0, i].Value = Math.Round(double.Parse(lst_data[i]), 2);
                                    }

                                    int ID = int.Parse(dt_region.Rows[i]["ID"].ToString());
                                    if (dgv_Analysis.Rows[ID - 1].Cells["Data"].Style.BackColor == Color.Red)
                                    {
                                        lst_judge.Add("Fail");
                                        judge_all = false;
                                    }
                                    else
                                    {
                                        lst_judge.Add("Pass");
                                    }
                                }
                            }

                            for (int i = 1; i < 13; i++)
                            {
                                if (myCode.checkDBNull(cell_data.Offset[i, -1].Value).Replace(" ", "").ToUpper()
                                    .Contains("MINFORCE"))
                                {
                                    myExcel.Range cell_min = cell_data.Offset[i, 0];
                                    myExcel.Range cell_max = cell_data.Offset[i + 1, 0];
                                    myExcel.Range cell_ave = cell_data.Offset[i + 2, 0];
                                    int c_offset = count_sample;
                                    cell_min.FormulaR1C1 = "=MIN(R[-" + i.ToString() + "]C[0]:R[-" + i.ToString() +
                                                           "]C[" + (c_offset - 1).ToString() + "])";
                                    cell_max.FormulaR1C1 = "=MAX(R[-" + (i + 1).ToString() + "]C[0]:R[-" +
                                                           (i + 1).ToString() + "]C[" + (c_offset - 1).ToString() +
                                                           "])";
                                    cell_ave.FormulaR1C1 = "=AVERAGE(R[-" + (i + 2).ToString() + "]C[0]:R[-" +
                                                           (i + 2).ToString() + "]C[" + (c_offset - 1).ToString() +
                                                           "])";
                                    break;
                                }
                            }

                            int k1 = 1;
                            int k2 = 7;
                            string type = dt_spec.Rows[0]["Location"].ToString().Split('+')[0];
                            int sochan = 0;
                            string str13 = dt_region.AsEnumerable()
                                .Select<DataRow, string>(
                                    (System.Func<DataRow, string>)(x => x.Field<string>(dic_mode["Mode 1#"])))
                                .ToList<string>()[0];
                            if (str13.Contains("("))
                            {
                                if (myCode.IsNumeric(str13.Split('/')[1].Split(')')[0]))
                                    sochan = int.Parse(str13.Split('/')[1].Split(')')[0]);
                            }

                            if (sheet == "SHEAR_TEST" && type == "A")
                            {
                                for (int i = 0; i < count_sample; i++)
                                {
                                    if (i < dt_region.Rows.Count)
                                    {
                                        if (myCode.IsNumeric(lst_data[i]))
                                        {
                                            //cell_data.Offset[0, i].Value = Math.Round(double.Parse(lst_data[i]), 2);
                                            cell_data.Offset[1, i].Value =
                                                Math.Round(double.Parse(lst_data[i]) * 9.81, 2);
                                        }
                                    }
                                }

                                k1 = 2;
                                k2 = 8;

                                for (int i = k1; i <= k2; i++)
                                {
                                    List<string> lst_mode = dt_region.AsEnumerable().Select(x =>
                                        x.Field<string>(dic_mode["Mode " + (i - 1).ToString() + "#"])).ToList();
                                    for (int j = 0; j < count_sample; j++)
                                    {
                                        if (j < lst_mode.Count)
                                        {
                                            if (lst_mode[j].Contains("("))
                                            {
                                                cell_data.Offset[i, j].FormulaR1C1 =
                                                    "=" + lst_mode[j].Split('(')[1].Split(')')[0];
                                            }
                                            else if (sochan != 0)
                                            {
                                                //cell_data.Offset[i, j].Value = lst_mode[j];
                                                cell_data.Offset[i, j].Value = "=0/" + sochan.ToString();
                                            }
                                            else
                                            {
                                                cell_data.Offset[i, j].Value = "'" + lst_mode[j];
                                            }

                                            // cell_data.Offset[i, j].NumberFormat = "Percentage";
                                            cell_data.Offset[i, j].Style = "Percent";
                                        }
                                    }
                                }

                                for (int i = 0; i < count_sample; i++)
                                {
                                    cell_data.Offset[9, i].Value = lst_judge[i];
                                    if (lst_judge[i] == "Pass")
                                    {
                                        cell_data.Offset[9, i].Interior.Color = Color.Green;
                                    }
                                    else
                                    {
                                        cell_data.Offset[9, i].Interior.Color = Color.Red;
                                    }
                                }
                            }
                            else
                            {
                                for (int i = k1; i <= k2; i++)
                                {
                                    List<string> lst_mode = dt_region.AsEnumerable()
                                        .Select(x => x.Field<string>(dic_mode["Mode " + i.ToString() + "#"])).ToList();
                                    for (int j = 0; j < count_sample; j++)
                                    {
                                        if (j < lst_mode.Count)
                                        {
                                            if (lst_mode[j].Contains("("))
                                            {
                                                cell_data.Offset[i, j].FormulaR1C1 =
                                                    "=" + lst_mode[j].Split('(')[1].Split(')')[0];
                                            }
                                            else if (sochan != 0)
                                            {
                                                //cell_data.Offset[i, j].Value = "'" + lst_mode[j];
                                                cell_data.Offset[i, j].Value = "=0/" + sochan.ToString();
                                            }
                                            else
                                            {
                                                cell_data.Offset[i, j].Value = "'" + lst_mode[j];
                                            }

                                            // cell_data.Offset[i, j].NumberFormat = "Percentage";
                                            cell_data.Offset[i, j].Style = "Percent";
                                        }
                                    }
                                }

                                for (int i = 0; i < Math.Min(count_sample, lst_judge.Count); i++)
                                {
                                    cell_data.Offset[8, i].Value = lst_judge[i];
                                    if (lst_judge[i] == "Pass")
                                    {
                                        cell_data.Offset[8, i].Interior.Color = Color.Green;
                                    }
                                    else
                                    {
                                        cell_data.Offset[8, i].Interior.Color = Color.Red;
                                    }
                                }
                            }
                        }

                        reg++;
                    }
                }

                int r_offset_del = 14;
                complete_sheet_other(ws, ItemCode, int.Parse(txt_qty.Text), r_offset_del);


                if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                {
                    myExcel.Range cur_rgn_judge = find_cell(ws, "JUDGEMENT");
                    if (cur_rgn_judge != null)
                    {
                        if (judge_all)
                        {
                            cur_rgn_judge.Offset[1, 0].Value = "OK";
                            ws.Cells[1, 1].Value = "OK";
                        }
                        else
                        {
                            cur_rgn_judge.Offset[1, 0].Value = "NG";
                            ws.Cells[1, 1].Value = "NG";
                        }
                    }
                }


                //if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                //{
                //    string report_folder = Path.Combine(data_loc, "Report", cb_Type.SelectedItem.ToString(), sheet);
                //    if (!System.IO.Directory.Exists(report_folder))
                //        System.IO.Directory.CreateDirectory(report_folder);

                //    // string _fname = Path.GetFileNameWithoutExtension(file_format).TrimEnd(new char[] { ',', '.' });
                //    string _fname = Path.GetFileNameWithoutExtension(file_format);
                //    string name = Path.GetFileNameWithoutExtension(Data_all.Rows[0]["Remark"].ToString());

                //    _fname = cb_Type.SelectedItem.ToString().Split('(')[1].Split(')')[0] + "_" + _fname.Replace(name.Split('-')[0], "").Replace(txtItemCode.Text, "").Replace("-", "");
                //    export_path = Path.Combine(report_folder, _fname + "-" + name + ".xlsx");
                //    //    report_format.SaveAs(export_path);

                //}

                MessageBox.Show(new Form { TopMost = true }, "Xuất dữ liệu thành công!", "Thông báo");
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Không tìm thấy sheet: " + sheet, "Thông báo");
            }
        }

        public Bitmap ChangeColor(Bitmap scrBitmap, System.Drawing.Color newColor)
        {
            System.Drawing.Color actualColor;
            Color ref_color = Color.FromArgb(255, 255, 255);
            Bitmap newBitmap = new Bitmap(scrBitmap.Width, scrBitmap.Height);

            for (int i = 0; i < scrBitmap.Width; i++)
            {
                for (int j = 0; j < scrBitmap.Height; j++)
                {
                    actualColor = scrBitmap.GetPixel(i, j);

                    if (actualColor.R > 240 && actualColor.G > 240 && actualColor.B > 240)
                    {
                        newBitmap.SetPixel(i, j, newColor);
                    }
                    else
                    {
                        newBitmap.SetPixel(i, j, actualColor);
                    }
                }
            }

            return newBitmap;
        }
        //   3+Cross section B2B MALE & FEMALE_IO PIN;11_Cross section B2B MALE & FEMALE_GROUNDING PIN_LEFT;27_Cross section B2B MALE & FEMALE_GROUNDING PIN_RIGHT;47

        public void export_excel_onproduct_old(string ItemCode, string LotNo, string file_format, string mysheet)
        {
            myExcel.Workbook curr_wrkbook = create_export_wrk(file_format, mysheet);
            if (curr_wrkbook != null)
            {
                myExcel.Worksheet ws = curr_wrkbook.Sheets[1];
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet" },
                    new string[] { ItemCode, mysheet });
                DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, mysheet, cb_Type.SelectedItem.ToString() }));
                // DataTable Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, tbl_name_comment3, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Sheet" }, new string[] { ItemCode, LotNo, mysheet.ToUpper().Replace("-", "").Replace(" ", "") }));

                //if (Data_tbl.Rows.Count > 0)
                //{

                if (dt_spec.Rows.Count > 0)
                {
                    string[] region = dt_spec.Rows[0]["Location"].ToString().Split('_');
                    foreach (string reg in region)
                    {
                        if (reg != "")
                        {
                            string[] str_location = reg.Split('+');

                            string cpn = "";
                            foreach (char c in str_location[0].Split(';')[0])
                            {
                                if (myCode.IsNumeric(c.ToString()))
                                    cpn += c;
                            }

                            cpn += "deg";
                            // string cpn = str_location[0].Split(';')[0].Replace(" ", "").Replace("JPOND(Peeling", "").Replace(")", "");

                            myExcel.Range cell_Pic = ws.Cells[int.Parse(str_location[0].Split(';')[2]),
                                int.Parse(str_location[0].Split(';')[3]) + 1];
                            myExcel.Range cell_graph = ws.Cells[int.Parse(str_location[1].Split(';')[1]),
                                int.Parse(str_location[0].Split(';')[3]) + 1];
                            myExcel.Range cell_data = ws.Cells[int.Parse(str_location[2].Split(';')[1]),
                                int.Parse(str_location[0].Split(';')[3]) + 1];


                            DataTable Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, sheet,
                                TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" },
                                    new string[] { ItemCode, LotNo, cpn }));

                            if (Data_tbl.Rows.Count > 0)
                            {
                                //int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
                                int count_sample = int.Parse(txt_qty.Text);
                                Export_DatatableImage_Excel(Data_tbl, "Image", ws, cell_Pic, false, count_sample);
                                Export_DatatableImage_Excel_Graph(Data_tbl, "Graph", ws, cell_graph, false,
                                    count_sample);


                                //List<byte[]> lst_img = Data_tbl.AsEnumerable().Select(x => x.Field<byte[]>("Image")).ToList();
                                //List<byte[]> lst_grp = Data_tbl.AsEnumerable().Select(x => x.Field<byte[]>("Graph")).ToList();
                                List<string> lst_data = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Data"))
                                    .ToList();

                                for (int i = 0; i < count_sample; i++)
                                {
                                    if (i < Data_tbl.Rows.Count)
                                    {
                                        cell_data.Offset[0, i].Value = lst_data[i].Split('_')[0].Replace("Max:", "");
                                        cell_data.Offset[1, i].Value =
                                            lst_data[i].Split('_')[1].Replace("Average:", "");
                                    }
                                }
                            }
                        }
                    }

                    MessageBox.Show(new Form { TopMost = true }, "Xuất dữ liệu thành công!", "Thông báo");
                }
                // }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Không tìm thấy sheet: " + sheet, "Thông báo");
            }
        }


        public void export_excel_onproduct(string ItemCode, string LotNo, string file_format, string mysheet,
            DataTable Data_all, DataTable dt_spec)
        {
            myExcel.Workbook curr_wrkbook = null;
            if (cb_Type.SelectedItem.ToString() == "NPI" || cb_Type.SelectedItem.ToString() == "Other")
            {
                curr_wrkbook = create_export_wrk(file_format, mysheet);
            }
            else if (cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                curr_wrkbook = TDMK_Code.open_excel_file(file_format, "", "");
            }

            if (curr_wrkbook != null)
            {
                myExcel.Worksheet ws = curr_wrkbook.Sheets[1];
                int count_sample = int.Parse(txt_qty.Text);
                if (dt_spec.Rows.Count > 0)
                {
                    List<string> lst_region = Data_all.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct()
                        .ToList();
                    string[] region = dt_spec.Rows[0]["Location"].ToString().Split('_');

                    if (cb_Type.SelectedItem.ToString().Contains("MASS") || cb_Type.SelectedItem.ToString() == "Other")
                    {
                        //myExcel.Range cur_rgn_type = find_cell(ws, cb_Type.SelectedItem.ToString().Replace("MASS(", "").Replace(")", ""));
                        //if (cur_rgn_type != null)
                        //    cur_rgn_type.Value = " v" + cur_rgn_type.Value.ToString();


                        string str_infor = Data_all.Rows[0]["Sheet"].ToString().Replace(sheet, "").Replace("/", "");
                        if (str_infor.Contains("_"))
                        {
                            string itemname = str_infor.Split('_')[0];
                            string line = str_infor.Split('_')[1];
                            string ca = str_infor.Split('_')[2];
                            string date = str_infor.Split('_')[3];
                            string worker = str_infor.Split('_')[4];
                            List<string> lst_infor = new List<string>
                                { itemname, ItemCode, LotNo, line, ca, date, worker };
                            export_info_mass(ws, lst_infor);
                        }


                        for (int i = 1; i < 20; i++)
                        {
                            for (int j = 1; j < 4; j++)
                            {
                                if (myCode.checkDBNull(ws.Cells[i, j].Value).ToUpper().Replace(" ", "")
                                    .Contains("ITEM-LOT"))
                                {
                                    ws.Cells[i, j].Offset[0, 1].Value = ItemCode + "-" + LotNo;
                                    break;
                                }
                            }
                        }
                    }

                    int r_end = 0;
                    if (region.Length == 2)
                    {
                        r_end = 25;
                    }
                    else if (region.Length == 3)
                    {
                        r_end = 40;
                    }
                    else if (region.Length == 4)
                    {
                        r_end = 52;
                    }

                    insert_columns_other(ws, ItemCode, int.Parse(txt_qty.Text), r_end);

                    foreach (string reg in region)
                    {
                        if (reg != "")
                        {
                            string[] str_location = reg.Split('+');
                            string cpn = "";

                            foreach (string r in lst_region)
                            {
                                string key1 = r.Split('_')[0].Replace(" ", "").ToUpper();
                                string key2 = r.Split('_')[1].Replace(" ", "").ToUpper();

                                if (reg.Replace("Component", "Tape").Replace(" ", "").ToUpper().Contains(key1) &&
                                    reg.Replace("Component", "Tape").Replace(" ", "").ToUpper().Contains(key2))
                                {
                                    cpn = r;
                                    break;
                                }
                            }

                            myExcel.Range cell_Pic = ws.Cells[int.Parse(str_location[1].Split(';')[2]),
                                int.Parse(str_location[1].Split(';')[3]) + 1];
                            myExcel.Range cell_graph = ws.Cells[int.Parse(str_location[2].Split(';')[1]),
                                int.Parse(str_location[1].Split(';')[3]) + 1];
                            myExcel.Range cell_data = ws.Cells[int.Parse(str_location[3].Split(';')[1]),
                                int.Parse(str_location[1].Split(';')[3]) + 1];

                            if (cpn != "")
                            {
                                DataTable Data_tbl = Data_all.AsEnumerable()
                                    .Where(s => s.Field<string>("Region") == cpn).CopyToDataTable();
                                if (Data_tbl.Rows.Count > 0)
                                {
                                    Export_DatatableImage_Excel(Data_tbl, "Image", ws, cell_Pic, false, count_sample);
                                    Export_DatatableImage_Excel_Graph(Data_tbl, "Graph", ws, cell_graph, false,
                                        count_sample);

                                    List<string> lst_data = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Data"))
                                        .ToList();
                                    List<string> lst_judge = new List<string> { };
                                    for (int i = 0; i < count_sample; i++)
                                    {
                                        if (i < Data_tbl.Rows.Count)
                                        {
                                            cell_data.Offset[0, i].Value =
                                                lst_data[i].Split('_')[0].Replace("Max:", "");
                                            cell_data.Offset[1, i].Value =
                                                lst_data[i].Split('_')[1].Replace("Average:", "");
                                        }
                                    }

                                    for (int i = 1; i < 13; i++)
                                    {
                                        if (myCode.checkDBNull(cell_data.Offset[i, -1].Value).Replace(" ", "").ToUpper()
                                            .Contains("MINFORCE"))
                                        {
                                            myExcel.Range cell_min = cell_data.Offset[i, 0];
                                            myExcel.Range cell_max = cell_data.Offset[i + 1, 0];
                                            myExcel.Range cell_ave = cell_data.Offset[i + 2, 0];
                                            int c_offset = count_sample;
                                            cell_min.FormulaR1C1 = "=MIN(R[-" + (i - 1).ToString() + "]C[0]:R[-" +
                                                                   (i - 1).ToString() + "]C[" +
                                                                   (c_offset - 1).ToString() + "])";
                                            cell_max.FormulaR1C1 = "=MAX(R[-" + i.ToString() + "]C[0]:R[-" +
                                                                   i.ToString() + "]C[" + (c_offset - 1).ToString() +
                                                                   "])";
                                            cell_ave.FormulaR1C1 = "=AVERAGE(R[-" + (i + 1).ToString() + "]C[0]:R[-" +
                                                                   (i + 1).ToString() + "]C[" +
                                                                   (c_offset - 1).ToString() + "])";


                                            if (cb_Type.SelectedItem.ToString() == "NPI")
                                            {
                                                myExcel.Range cell_min1 = cell_data.Offset[i, 1];
                                                myExcel.Range cell_max1 = cell_data.Offset[i + 1, 1];
                                                myExcel.Range cell_ave1 = cell_data.Offset[i + 2, 1];

                                                cell_min1.FormulaR1C1 = "=MIN(R[-" + i.ToString() + "]C[-1]:R[-" +
                                                                        i.ToString() + "]C[" +
                                                                        (c_offset - 2).ToString() + "])";
                                                cell_max1.FormulaR1C1 = "=MAX(R[-" + (i + 1).ToString() + "]C[-1]:R[-" +
                                                                        (i + 1).ToString() + "]C[" +
                                                                        (c_offset - 2).ToString() + "])";
                                                cell_ave1.FormulaR1C1 = "=AVERAGE(R[-" + (i + 2).ToString() +
                                                                        "]C[-1]:R[-" + (i + 2).ToString() + "]C[" +
                                                                        (c_offset - 2).ToString() + "])";
                                            }

                                            break;
                                        }
                                    }


                                    for (int i = 0; i < count_sample; i++)
                                    {
                                        if (i < Data_tbl.Rows.Count)
                                        {
                                            int ID = int.Parse(Data_tbl.Rows[i]["ID"].ToString());
                                            if (dgv_Analysis.Rows[ID - 1].Cells["Data"].Style.BackColor == Color.Red)
                                            {
                                                lst_judge.Add("Fail");
                                                judge_all = false;
                                            }
                                            else
                                            {
                                                lst_judge.Add("Pass");
                                            }
                                        }
                                    }

                                    for (int t = 1; t < 5; t++)
                                    {
                                        if (myCode.checkDBNull(cell_data.Offset[t, -1].Value).ToUpper()
                                                .Contains("JUDGEMENT") &&
                                            !myCode.checkDBNull(cell_data.Offset[t, -1].Value).ToUpper()
                                                .Replace(" ", "").Contains("FAILUREMODE"))
                                        {
                                            for (int i = 0; i < lst_judge.Count; i++)
                                            {
                                                cell_data.Offset[t, i].Value = lst_judge[i];
                                            }

                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                    {
                        myExcel.Range cur_rgn_judge = find_cell(ws, "JUDGEMENT");
                        if (cur_rgn_judge != null)
                        {
                            if (judge_all)
                            {
                                cur_rgn_judge.Offset[1, 0].Value = "OK";
                                ws.Cells[1, 1].Value = "OK";
                            }
                            else
                            {
                                cur_rgn_judge.Offset[1, 0].Value = "NG";
                                ws.Cells[1, 1].Value = "NG";
                            }
                        }
                    }

                    int r_offset_del = 8;
                    complete_sheet_other(ws, ItemCode, int.Parse(txt_qty.Text), r_offset_del);

                    MessageBox.Show(new Form { TopMost = true }, "Xuất dữ liệu thành công!", "Thông báo");
                }
                // }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Không tìm thấy sheet: " + sheet, "Thông báo");
            }
        }

        public void export_excel_gap_connector(string ItemCode, string LotNo, string file_format, string mysheet,
            DataTable Data_tbl, DataTable dt_spec)
        {
            myExcel.Workbook curr_wrkbook = null;
            if (cb_Type.SelectedItem.ToString() == "NPI" || cb_Type.SelectedItem.ToString() == "Other")
            {
                curr_wrkbook = create_export_wrk(file_format, mysheet);
            }
            else if (cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                curr_wrkbook = TDMK_Code.open_excel_file(file_format, "", "");
            }

            if (curr_wrkbook != null)
            {
                myExcel.Worksheet ws = curr_wrkbook.Sheets[1];

                int col_begin = int.Parse(dt_spec.Rows[0]["Location"].ToString().Split('+')[0]);
                string[] region = dt_spec.Rows[0]["Location"].ToString().Split('+')[1].Split('_');
                int count_sample = int.Parse(txt_qty.Text);

                List<string> lst_region_db =
                    Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToList();
                int r_end = 0;

                insert_columns_other(ws, ItemCode, int.Parse(txt_qty.Text), r_end);

                foreach (string reg in region)
                {
                    if (reg != "")
                    {
                        if (reg.Split(';')[0] == "IOPIN")
                        {
                            int code = 1;
                            foreach (string pos in reg.Split('^'))
                            {
                                if (pos != "")
                                {
                                    int row_begin = int.Parse(pos.Split(';')[1]);
                                    int count_r_offset = int.Parse(pos.Split(';')[2]);
                                    myExcel.Range cell_Pic = ws.Cells[row_begin, col_begin];
                                    myExcel.Range cell_Pic1 = ws.Cells[row_begin + 1, col_begin];
                                    myExcel.Range cell_Pic2 = ws.Cells[row_begin + count_r_offset + 2, col_begin];
                                    myExcel.Range cell_data = ws.Cells[row_begin + 2, col_begin];
                                    string find_text = "";

                                    if (reg.Split('^').Length == 3)
                                    {
                                        find_text = code.ToString();
                                    }

                                    string filter_reg = "";
                                    foreach (string i in lst_region_db)
                                    {
                                        if (i.ToUpper().Contains("DOC") && i.Contains(find_text))
                                        {
                                            filter_reg = i;
                                            break;
                                        }
                                    }
                                    // DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo, filter_reg }));

                                    if (filter_reg != "")
                                    {
                                        DataTable tbl_region = Data_tbl.AsEnumerable()
                                            .Where(s => s.Field<string>("Region") == filter_reg).CopyToDataTable();
                                        if (tbl_region.Rows.Count > 0)
                                        {
                                            // int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());

                                            Export_DatatableImage_Excel(tbl_region, "Image", ws, cell_Pic, false,
                                                count_sample);
                                            Export_DatatableImage_Excel(tbl_region, "Image1", ws, cell_Pic1, false,
                                                count_sample);
                                            Export_DatatableImage_Excel(tbl_region, "Image2", ws, cell_Pic2, false,
                                                count_sample);

                                            List<string> lst_data = tbl_region.AsEnumerable()
                                                .Select(x => x.Field<string>("Data")).ToList();

                                            for (int i = 0; i < count_sample; i++)
                                            {
                                                if (i < tbl_region.Rows.Count)
                                                {
                                                    int min_r = new int[]
                                                        {
                                                            count_r_offset, lst_data[i].Split('/')[0].Split(';').Length
                                                        }
                                                        .Min();
                                                    for (int k = 0; k < min_r; k++)
                                                    {
                                                        cell_data.Offset[k, i].Value =
                                                            lst_data[i].Split('/')[0].Split(';')[k];
                                                        cell_data.Offset[k + count_r_offset + 1, i].Value =
                                                            lst_data[i].Split('/')[1].Split(';')[k];
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                code++;
                            }
                        }
                        else if (reg.Split(';')[0] == "LEFT" || reg.Split(';')[0] == "RIGHT")
                        {
                            int code = 1;
                            foreach (string pos in reg.Split('^'))
                            {
                                if (pos != "")
                                {
                                    int row_begin = int.Parse(pos.Split(';')[1]);
                                    int count_r_offset = int.Parse(pos.Split(';')[2]);
                                    myExcel.Range cell_Pic = ws.Cells[row_begin, col_begin];
                                    myExcel.Range cell_Pic1 = ws.Cells[row_begin + 1, col_begin];
                                    myExcel.Range cell_Pic2 = ws.Cells[row_begin + count_r_offset + 2, col_begin];
                                    myExcel.Range cell_data = ws.Cells[row_begin + 2, col_begin];

                                    string find_text = "";

                                    if (reg.Split('^').Length == 3)
                                    {
                                        find_text = code.ToString();
                                    }

                                    string filter_reg = "";
                                    foreach (string i in lst_region_db)
                                    {
                                        if (i.ToUpper().Contains("TRU") && i.Contains(find_text))
                                        {
                                            filter_reg = i;
                                            break;
                                        }
                                    }

                                    //DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo, filter_reg }));
                                    if (filter_reg != "")
                                    {
                                        DataTable tbl_region = Data_tbl.AsEnumerable()
                                            .Where(s => s.Field<string>("Region") == filter_reg).CopyToDataTable();
                                        if (tbl_region.Rows.Count > 0)
                                        {
                                            //int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());


                                            if (reg.Contains("RIGHT"))
                                            {
                                                for (int r = 0; r < count_sample; r++)
                                                {
                                                    tbl_region.Rows.Remove(tbl_region.Rows[0]);
                                                }
                                            }

                                            Export_DatatableImage_Excel(tbl_region, "Image", ws, cell_Pic, false,
                                                count_sample);
                                            Export_DatatableImage_Excel(tbl_region, "Image1", ws, cell_Pic1, false,
                                                count_sample);
                                            Export_DatatableImage_Excel(tbl_region, "Image2", ws, cell_Pic2, false,
                                                count_sample);


                                            List<string> lst_data = tbl_region.AsEnumerable()
                                                .Select(x => x.Field<string>("Data")).ToList();

                                            for (int i = 0; i < count_sample; i++)
                                            {
                                                if (i < tbl_region.Rows.Count)
                                                {
                                                    int min_r = new int[]
                                                        {
                                                            count_r_offset, lst_data[i].Split('/')[0].Split(';').Length
                                                        }
                                                        .Min();
                                                    for (int k = 0; k < min_r; k++)
                                                    {
                                                        cell_data.Offset[k, i].Value =
                                                            lst_data[i].Split('/')[0].Split(';')[k];
                                                        cell_data.Offset[k + count_r_offset + 1, i].Value =
                                                            lst_data[i].Split('/')[1].Split(';')[k];
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                code++;
                            }
                        }
                    }
                }

                int r_offset_del = 0;

                complete_sheet_other(ws, ItemCode, int.Parse(txt_qty.Text), r_offset_del);

                MessageBox.Show(new Form { TopMost = true }, "Xuất dữ liệu thành công!", "Thông báo");
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Không tìm thấy sheet " + sheet, "Thông báo");
            }
        }

        public void export_Hottizontal_NPI(string reg, myExcel.Worksheet ws, int row_begin, int col_begin, int no_lk,
            List<string> lst_region, int count_row, DataTable Data_tbl, DataTable dt_spec, string ItemCode,
            string LotNo)
        {
            int count_sample = int.Parse(txt_qty.Text);
            myExcel.Range cell_Pic1 = ws.Cells[row_begin, col_begin];
            myExcel.Range cell_Pic2 = ws.Cells[row_begin + 1, col_begin];
            myExcel.Range cell_data = ws.Cells[row_begin + 2, col_begin];
            string filter_reg = "";

            string txt_find = "";
            if (reg.ToUpper().Contains("MALE"))
            {
                txt_find = no_lk.ToString() + "_";
            }

            foreach (string i in lst_region)
            {
                if (i.ToUpper().Contains("NGANG") && !i.ToUpper().Contains("TRU") && i.ToUpper().Contains(txt_find))
                {
                    filter_reg = i;
                    break;
                }
            }

            if (filter_reg != "")
            {
                // DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo, filter_reg }));
                DataView dv = Data_tbl.AsDataView();
                dv.RowFilter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { filter_reg });
                DataTable tbl_region = dv.ToTable();

                if (tbl_region.Rows.Count > 0)
                {
                    // int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
                    //int count_sample = int.Parse(txt_qty.Text);
                    Export_DatatableImage_Excel_ngang(tbl_region, "Image1", ws, cell_Pic1, false, count_sample);
                    Export_DatatableImage_Excel_ngang(tbl_region, "Image2", ws, cell_Pic2, false, count_sample);

                    List<string> lst_data = tbl_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                    int k = 0;
                    for (int i = 0; i < tbl_region.Rows.Count; i++)
                    {
                        if (i < 2 * count_sample)
                        {
                            cell_data.Offset[0, k].Value = lst_data[i].Split(';')[0];
                            cell_data.Offset[1, k].Value = lst_data[i].Split(';')[1];
                            k++;

                            if (i == count_sample - 1)
                            {
                                cell_data = cell_data.Offset[4, 0];
                                k = 0;
                            }
                        }

                        if (i == 2 * count_sample - 1)
                        {
                            break;
                        }
                    }
                }
            }

            switch (count_row)
            {
                case 20:

                    row_begin = row_begin + 8;
                    Dictionary<string, int> _dic = new Dictionary<string, int> { };
                    _dic.Add("_T", 9);
                    _dic.Add("_P", 3);
                    foreach (var item in _dic)
                    {
                        myExcel.Range cell_Pic1_2 = ws.Cells[row_begin, col_begin];
                        myExcel.Range cell_data_2 = ws.Cells[row_begin + 1, col_begin];

                        string filter_reg_2 = "";
                        // string txt_find_2 = "";
                        if (reg.ToUpper().Contains("MALE"))
                        {
                            txt_find = "MALE";
                        }
                        else if (reg.ToUpper().Contains("FEMALE"))
                        {
                            txt_find = "FEMALE";
                        }

                        foreach (string i in lst_region)
                        {
                            if (i.Replace(" ", "").ToUpper().Contains("TRUNGANG") &&
                                i.Replace(" ", "").ToUpper().Contains(item.Key) && i.ToUpper().Contains(txt_find))
                            {
                                filter_reg_2 = i;
                                break;
                            }
                        }

                        if (filter_reg_2 != "")
                        {
                            //DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo, filter_reg_2 }));

                            DataView dv = Data_tbl.AsDataView();
                            dv.RowFilter =
                                TDMK_Code.filter_str(new string[] { "Region" }, new string[] { filter_reg_2 });
                            DataTable tbl_region = dv.ToTable();

                            if (Data_tbl.Rows.Count > 0)
                            {
                                // int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());

                                Export_DatatableImage_Excel_trungang(tbl_region, "Image1", ws, cell_Pic1_2, false,
                                    count_sample, item.Value);
                                //Export_DatatableImage_Excel_trungang(tbl_region, "Image2", ws, cell_Pic2, false, count_sample, item.Value);

                                List<string> lst_data = tbl_region.AsEnumerable().Select(x => x.Field<string>("Data"))
                                    .ToList();
                                int k = 0;
                                for (int i = 0; i < tbl_region.Rows.Count; i++)
                                {
                                    if (i < 2 * count_sample)
                                    {
                                        cell_data_2.Offset[0, k].Value = lst_data[i].Split(';')[0];
                                        cell_data_2.Offset[1, k].Value = lst_data[i].Split(';')[1];
                                        k++;
                                        if (i == count_sample - 1)
                                        {
                                            cell_data_2 = cell_data_2.Offset[item.Value, 0];
                                            k = 0;
                                        }
                                    }

                                    if (i == 2 * count_sample - 1)
                                    {
                                        break;
                                    }
                                }
                            }
                        }

                        row_begin = row_begin + 3;
                    }

                    break;

                case 26:
                    row_begin = row_begin + 8;
                    Dictionary<string, int> _dic1 = new Dictionary<string, int> { };
                    _dic1.Add("1_T", 9);
                    _dic1.Add("1_P", 3);
                    _dic1.Add("2", 3);

                    foreach (var item in _dic1)
                    {
                        myExcel.Range cell_Pic1_2 = ws.Cells[row_begin, col_begin];
                        myExcel.Range cell_data_2 = ws.Cells[row_begin + 1, col_begin];

                        string filter_reg_2 = "";
                        if (reg.ToUpper().Contains("MALE"))
                        {
                            txt_find = "MALE";
                        }
                        else if (reg.ToUpper().Contains("FEMALE"))
                        {
                            txt_find = "FEMALE";
                        }

                        foreach (string i in lst_region)
                        {
                            if (i.Replace(" ", "").ToUpper().Contains("TRUNGANG") &&
                                i.Replace(" ", "").ToUpper().Contains(item.Key) && i.ToUpper().Contains(txt_find))
                            {
                                filter_reg_2 = i;
                                break;
                            }
                        }

                        if (filter_reg_2 != "")
                        {
                            //DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo, filter_reg_2 }));
                            DataView dv = Data_tbl.AsDataView();
                            dv.RowFilter =
                                TDMK_Code.filter_str(new string[] { "Region" }, new string[] { filter_reg_2 });
                            DataTable tbl_region = dv.ToTable();

                            if (tbl_region.Rows.Count > 0)
                            {
                                //int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
                                Export_DatatableImage_Excel_trungang(tbl_region, "Image1", ws, cell_Pic1_2, false,
                                    count_sample, item.Value);

                                List<string> lst_data = tbl_region.AsEnumerable().Select(x => x.Field<string>("Data"))
                                    .ToList();
                                int k = 0;
                                for (int i = 0; i < tbl_region.Rows.Count; i++)
                                {
                                    if (i < 2 * count_sample)
                                    {
                                        cell_data_2.Offset[0, k].Value = lst_data[i].Split(';')[0];
                                        cell_data_2.Offset[1, k].Value = lst_data[i].Split(';')[1];
                                        k++;
                                        if (i == count_sample - 1)
                                        {
                                            cell_data_2 = cell_data_2.Offset[item.Value, 0];
                                            k = 0;
                                        }
                                    }

                                    if (i == 2 * count_sample - 1)
                                    {
                                        break;
                                    }
                                }
                            }
                        }

                        if (item.Key == "1_T")
                        {
                            row_begin = row_begin + 3;
                        }
                        else if (item.Key == "1_P")
                        {
                            row_begin = row_begin + 9;
                        }
                    }

                    break;

                case 32:
                    row_begin = row_begin + 8;
                    Dictionary<string, int> _dic2 = new Dictionary<string, int> { };
                    _dic2.Add("1_T", 9);
                    _dic2.Add("1_P", 3);
                    _dic2.Add("2_T", 9);
                    _dic2.Add("2_P", 3);
                    foreach (var item in _dic2)
                    {
                        myExcel.Range cell_Pic1_2 = ws.Cells[row_begin, col_begin];
                        myExcel.Range cell_data_2 = ws.Cells[row_begin + 1, col_begin];

                        string filter_reg_2 = "";
                        if (reg.ToUpper().Contains("MALE"))
                        {
                            txt_find = "MALE";
                        }
                        else if (reg.ToUpper().Contains("FEMALE"))
                        {
                            txt_find = "FEMALE";
                        }

                        foreach (string i in lst_region)
                        {
                            if (i.Replace(" ", "").ToUpper().Contains("TRUNGANG") &&
                                i.Replace(" ", "").ToUpper().Contains(item.Key) && i.ToUpper().Contains(txt_find))
                            {
                                filter_reg_2 = i;
                                break;
                            }
                        }

                        if (filter_reg_2 != "")
                        {
                            //DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo, filter_reg_2 }));
                            DataView dv = Data_tbl.AsDataView();
                            dv.RowFilter =
                                TDMK_Code.filter_str(new string[] { "Region" }, new string[] { filter_reg_2 });
                            DataTable tbl_region = dv.ToTable();

                            if (tbl_region.Rows.Count > 0)
                            {
                                //int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
                                Export_DatatableImage_Excel_trungang(tbl_region, "Image1", ws, cell_Pic1_2, false,
                                    count_sample, item.Value);

                                List<string> lst_data = tbl_region.AsEnumerable().Select(x => x.Field<string>("Data"))
                                    .ToList();
                                int k = 0;
                                for (int i = 0; i < tbl_region.Rows.Count; i++)
                                {
                                    if (i < 2 * count_sample)
                                    {
                                        cell_data_2.Offset[0, k].Value = lst_data[i].Split(';')[0];
                                        cell_data_2.Offset[1, k].Value = lst_data[i].Split(';')[1];
                                        k++;
                                        if (i == count_sample - 1)
                                        {
                                            cell_data_2 = cell_data_2.Offset[item.Value, 0];
                                            k = 0;
                                        }
                                    }

                                    if (i == 2 * count_sample - 1)
                                    {
                                        break;
                                    }
                                }
                            }
                        }

                        if (item.Key.Contains("T"))
                        {
                            row_begin = row_begin + 3;
                        }
                        else
                        {
                            row_begin = row_begin + 9;
                        }
                    }

                    break;


                default:
                    break;
            }
        }

        public void export_Hottizontal_NPI_28(string reg, myExcel.Worksheet ws, int row_begin, int col_begin, int no_lk,
            List<string> lst_region, int count_row, DataTable Data_tbl, DataTable dt_spec, string ItemCode,
            string LotNo)
        {
            int count_sample = int.Parse(txt_qty.Text);
            myExcel.Range cell_Pic1 = ws.Cells[row_begin, col_begin];
            myExcel.Range cell_Pic2 = ws.Cells[row_begin + 1, col_begin];
            myExcel.Range cell_data = ws.Cells[row_begin + 2, col_begin];

            List<string> lst_stt_ngang = new List<string> { "NGANG_1", "NGANG_4", "NGANG_2", "NGANG_3" };

            foreach (string filter_reg in lst_stt_ngang)
            {
                DataView dv = Data_tbl.AsDataView();
                dv.RowFilter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { filter_reg });
                DataTable tbl_region = dv.ToTable();

                if (tbl_region.Rows.Count > 0)
                {
                    Export_DatatableImage_Excel_ngang(tbl_region, "Image1", ws, cell_Pic1, false, count_sample);
                    Export_DatatableImage_Excel_ngang(tbl_region, "Image2", ws, cell_Pic2, false, count_sample);

                    List<string> lst_data = tbl_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                    int k = 0;
                    for (int i = 0; i < tbl_region.Rows.Count; i++)
                    {
                        if (i < count_sample)
                        {
                            cell_data.Offset[0, k].Value = lst_data[i].Split(';')[0];
                            cell_data.Offset[1, k].Value = lst_data[i].Split(';')[1];
                            k++;
                        }
                    }

                    cell_Pic1 = cell_Pic1.Offset[4, 0];
                    cell_Pic2 = cell_Pic2.Offset[4, 0];
                    cell_data = cell_data.Offset[4, 0];
                }
            }

            switch (count_row)
            {
                case 28:
                    row_begin = row_begin + 16;
                    Dictionary<string, int> _dic2 = new Dictionary<string, int> { };
                    _dic2.Add("1_T", 9);
                    _dic2.Add("1_P", 3);
                    _dic2.Add("2_T", 9);
                    _dic2.Add("2_P", 3);
                    foreach (var item in _dic2)
                    {
                        myExcel.Range cell_Pic1_2 = ws.Cells[row_begin, col_begin];
                        myExcel.Range cell_data_2 = ws.Cells[row_begin + 1, col_begin];

                        string filter_reg_2 = "";

                        foreach (string i in lst_region)
                        {
                            if (i.Replace(" ", "").ToUpper().Contains("TRUNGANG") &&
                                i.Replace(" ", "").ToUpper().Contains(item.Key))
                            {
                                filter_reg_2 = i;
                                break;
                            }
                        }

                        if (filter_reg_2 != "")
                        {
                            //DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo, filter_reg_2 }));
                            DataView dv = Data_tbl.AsDataView();
                            dv.RowFilter =
                                TDMK_Code.filter_str(new string[] { "Region" }, new string[] { filter_reg_2 });
                            DataTable tbl_region = dv.ToTable();

                            if (tbl_region.Rows.Count > 0)
                            {
                                //int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
                                Export_DatatableImage_Excel_trungang(tbl_region, "Image1", ws, cell_Pic1_2, false,
                                    count_sample, item.Value);

                                List<string> lst_data = tbl_region.AsEnumerable().Select(x => x.Field<string>("Data"))
                                    .ToList();
                                int k = 0;
                                for (int i = 0; i < tbl_region.Rows.Count; i++)
                                {
                                    if (i < 2 * count_sample)
                                    {
                                        cell_data_2.Offset[0, k].Value = lst_data[i].Split(';')[0];
                                        cell_data_2.Offset[1, k].Value = lst_data[i].Split(';')[1];
                                        k++;
                                        if (i == count_sample - 1)
                                        {
                                            cell_data_2 = cell_data_2.Offset[item.Value, 0];
                                            k = 0;
                                        }
                                    }

                                    if (i == 2 * count_sample - 1)
                                    {
                                        break;
                                    }
                                }
                            }
                        }

                        if (item.Key.Contains("T"))
                        {
                            row_begin = row_begin + 3;
                        }
                        else
                        {
                            row_begin = row_begin + 9;
                        }
                    }

                    break;


                default:
                    break;
            }
        }

        public void export_Hottizontal_MASS(string reg, myExcel.Worksheet ws, int row_begin, int col_begin, int no_lk,
            List<string> lst_region, int count_row, DataTable Data_tbl, DataTable dt_spec, string ItemCode,
            string LotNo)
        {
            myExcel.Range cell_Pic = ws.Cells[row_begin, col_begin];
            myExcel.Range cell_data = ws.Cells[row_begin + 1, col_begin];
            string filter_reg = "";
            string txt_find = "";
            if (reg.ToUpper().Contains("MALE"))
            {
                txt_find = no_lk.ToString() + "_";
            }


            foreach (string i in lst_region)
            {
                if (i.ToUpper().Contains("NGANG") && !i.ToUpper().Contains("TRU") && i.ToUpper().Contains(txt_find))
                {
                    filter_reg = i;
                    break;
                }
            }

            if (filter_reg != "")
            {
                //DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo, filter_reg }));
                DataView dv = Data_tbl.AsDataView();
                dv.RowFilter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { filter_reg });
                DataTable tbl_region = dv.ToTable();

                if (tbl_region.Rows.Count > 0)
                {
                    // int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
                    int count_sample = int.Parse(txt_qty.Text);
                    Export_DatatableImage_Excel_ngang_mass(tbl_region, "Image1", ws, cell_Pic, false, count_sample);

                    List<string> lst_data = tbl_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();
                    int k = 0;
                    for (int i = 0; i < tbl_region.Rows.Count; i++)
                    {
                        if (i < 2 * count_sample)
                        {
                            cell_data.Offset[0, k].Value = lst_data[i].Split(';')[0];
                            cell_data.Offset[1, k].Value = lst_data[i].Split(';')[1];
                            k++;

                            if (i == count_sample - 1)
                            {
                                cell_data = cell_data.Offset[3, 0];
                                k = 0;
                            }
                        }

                        if (i == 2 * count_sample - 1)
                        {
                            break;
                        }
                    }
                }
            }
        }

        public SortedDictionary<int, string> dic_judge_crosscut(int count_sample, DataTable dt_data)
        {
            List<DataTable> lst_Table = new List<DataTable> { };
            Get_ListTable(-1, dt_data, new string[] { "Region" }, ref lst_Table, "Data");
            SortedDictionary<int, string> dic_judge = new SortedDictionary<int, string> { };


            foreach (DataTable dt in lst_Table)
            {
                string reggion = dt.Rows[0]["Region"].ToString();
                for (int i = 0; i < count_sample; i++)
                {
                    if (reggion.Contains("NGANG"))
                    {
                        int ID1 = int.Parse(dt.Rows[i]["ID"].ToString());
                        int ID2 = int.Parse(dt.Rows[i + count_sample]["ID"].ToString());
                        if (dgv_Analysis.Rows[ID1 - 1].Cells["Data"].Style.BackColor == Color.Red ||
                            dgv_Analysis.Rows[ID2 - 1].Cells["Data"].Style.BackColor == Color.Red)
                        {
                            if (!dic_judge.ContainsKey(i + 1))
                            {
                                dic_judge.Add(i + 1, "Fail");
                                judge_all = false;
                            }
                        }
                    }
                    else
                    {
                        int ID = int.Parse(dt.Rows[i]["ID"].ToString());
                        if (dgv_Analysis.Rows[ID - 1].Cells["Data"].Style.BackColor == Color.Red)
                        {
                            if (!dic_judge.ContainsKey(i + 1))
                            {
                                dic_judge.Add(i + 1, "Fail");
                                judge_all = false;
                            }
                        }
                    }
                }
            }

            return dic_judge;
        }


        public void export_excel_cross_section_old(string ItemCode, string LotNo, string file_format, string mysheet)
        {
            myExcel.Workbook curr_wrkbook = null;
            if (cb_Type.SelectedItem.ToString() == "NPI")
            {
                curr_wrkbook = create_export_wrk(file_format, mysheet);
            }
            else if (cb_Type.SelectedItem.ToString() == "MASS")
            {
                curr_wrkbook = TDMK_Code.open_excel_file(file_format, "", "");
            }

            if (curr_wrkbook != null)
            {
                myExcel.Worksheet ws = curr_wrkbook.Sheets[1];
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet" },
                    new string[] { ItemCode, mysheet });
                DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                    TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                        new string[] { txtItemCode.Text, mysheet, cb_Type.SelectedItem.ToString() }));

                if (dt_spec.Rows.Count > 0)
                {
                    string location_spec = dt_spec.Rows[0]["Location"].ToString();


                    int col_begin = int.Parse(location_spec.Split('+')[0].Split(';')[1]);
                    string[] region = location_spec.Split('+')[1].Split('_');

                    DataTable Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, sheet,
                        TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo }));
                    List<string> lst_region = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct()
                        .ToList();
                    string txt_find = "TTT";
                    if (region.Length == 3)
                    {
                        txt_find = "";
                    }

                    foreach (string reg in region)
                    {
                        if (reg != "")
                        {
                            int count_row = int.Parse(reg.Split(';')[1]);
                            int row_begin = int.Parse(reg.Split(';')[2]);
                            if (reg.ToUpper().Contains("Hottizontal".ToUpper()))
                            {
                                myExcel.Range cell_Pic1 = ws.Cells[row_begin, col_begin];
                                myExcel.Range cell_Pic2 = ws.Cells[row_begin + 1, col_begin];
                                myExcel.Range cell_data = ws.Cells[row_begin + 2, col_begin];
                                string filter_reg = "";
                                //string txt_find = "";
                                if (reg.ToUpper().Contains("MALE"))
                                {
                                    txt_find = "MALE";
                                }
                                else if (reg.ToUpper().Contains("FEMALE"))
                                {
                                    txt_find = "FEMALE";
                                }

                                foreach (string i in lst_region)
                                {
                                    if (i.ToUpper().Contains("NGANG") && !i.ToUpper().Contains("TRU") &&
                                        i.ToUpper().Contains(txt_find))
                                    {
                                        filter_reg = i;
                                        break;
                                    }
                                }

                                if (filter_reg != "")
                                {
                                    DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet,
                                        TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" },
                                            new string[] { ItemCode, LotNo, filter_reg }));
                                    if (tbl_region.Rows.Count > 0)
                                    {
                                        // int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
                                        int count_sample = int.Parse(txt_qty.Text);
                                        Export_DatatableImage_Excel_ngang(tbl_region, "Image1", ws, cell_Pic1, false,
                                            count_sample);
                                        Export_DatatableImage_Excel_ngang(tbl_region, "Image2", ws, cell_Pic2, false,
                                            count_sample);

                                        List<string> lst_data = tbl_region.AsEnumerable()
                                            .Select(x => x.Field<string>("Data")).ToList();
                                        int k = 0;
                                        for (int i = 0; i < tbl_region.Rows.Count; i++)
                                        {
                                            if (i < 2 * count_sample)
                                            {
                                                cell_data.Offset[0, k].Value = lst_data[i].Split(';')[0];
                                                cell_data.Offset[1, k].Value = lst_data[i].Split(';')[1];
                                                k++;

                                                if (i == count_sample - 1)
                                                {
                                                    cell_data = cell_data.Offset[4, 0];
                                                    k = 0;
                                                }
                                            }

                                            if (i == 2 * count_sample - 1)
                                            {
                                                break;
                                            }
                                        }
                                    }
                                }

                                switch (count_row)
                                {
                                    case 20:

                                        row_begin = row_begin + 8;
                                        Dictionary<string, int> _dic = new Dictionary<string, int> { };
                                        _dic.Add("_T", 9);
                                        _dic.Add("_P", 3);
                                        foreach (var item in _dic)
                                        {
                                            myExcel.Range cell_Pic1_2 = ws.Cells[row_begin, col_begin];
                                            myExcel.Range cell_data_2 = ws.Cells[row_begin + 1, col_begin];

                                            string filter_reg_2 = "";
                                            // string txt_find_2 = "";
                                            if (reg.ToUpper().Contains("MALE"))
                                            {
                                                txt_find = "MALE";
                                            }
                                            else if (reg.ToUpper().Contains("FEMALE"))
                                            {
                                                txt_find = "FEMALE";
                                            }

                                            foreach (string i in lst_region)
                                            {
                                                if (i.Replace(" ", "").ToUpper().Contains("TRUNGANG") &&
                                                    i.Replace(" ", "").ToUpper().Contains(item.Key) &&
                                                    i.ToUpper().Contains(txt_find))
                                                {
                                                    filter_reg_2 = i;
                                                    break;
                                                }
                                            }

                                            if (filter_reg_2 != "")
                                            {
                                                DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet,
                                                    TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" },
                                                        new string[] { ItemCode, LotNo, filter_reg_2 }));
                                                if (Data_tbl.Rows.Count > 0)
                                                {
                                                    int count_sample =
                                                        int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
                                                    Export_DatatableImage_Excel_trungang(tbl_region, "Image1", ws,
                                                        cell_Pic1_2, false, count_sample, item.Value);
                                                    //Export_DatatableImage_Excel_trungang(tbl_region, "Image2", ws, cell_Pic2, false, count_sample, item.Value);

                                                    List<string> lst_data = tbl_region.AsEnumerable()
                                                        .Select(x => x.Field<string>("Data")).ToList();
                                                    int k = 0;
                                                    for (int i = 0; i < tbl_region.Rows.Count; i++)
                                                    {
                                                        if (i < 2 * count_sample)
                                                        {
                                                            cell_data_2.Offset[0, k].Value = lst_data[i].Split(';')[0];
                                                            cell_data_2.Offset[1, k].Value = lst_data[i].Split(';')[1];
                                                            k++;
                                                            if (i == count_sample - 1)
                                                            {
                                                                cell_data_2 = cell_data_2.Offset[item.Value, 0];
                                                                k = 0;
                                                            }
                                                        }

                                                        if (i == 2 * count_sample - 1)
                                                        {
                                                            break;
                                                        }
                                                    }
                                                }
                                            }

                                            row_begin = row_begin + 3;
                                        }

                                        break;

                                    case 26:
                                        row_begin = row_begin + 8;
                                        Dictionary<string, int> _dic1 = new Dictionary<string, int> { };
                                        _dic1.Add("1_T", 9);
                                        _dic1.Add("1_P", 3);
                                        _dic1.Add("2", 3);

                                        foreach (var item in _dic1)
                                        {
                                            myExcel.Range cell_Pic1_2 = ws.Cells[row_begin, col_begin];
                                            myExcel.Range cell_data_2 = ws.Cells[row_begin + 1, col_begin];

                                            string filter_reg_2 = "";
                                            if (reg.ToUpper().Contains("MALE"))
                                            {
                                                txt_find = "MALE";
                                            }
                                            else if (reg.ToUpper().Contains("FEMALE"))
                                            {
                                                txt_find = "FEMALE";
                                            }

                                            foreach (string i in lst_region)
                                            {
                                                if (i.Replace(" ", "").ToUpper().Contains("TRUNGANG") &&
                                                    i.Replace(" ", "").ToUpper().Contains(item.Key) &&
                                                    i.ToUpper().Contains(txt_find))
                                                {
                                                    filter_reg_2 = i;
                                                    break;
                                                }
                                            }

                                            if (filter_reg_2 != "")
                                            {
                                                DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet,
                                                    TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" },
                                                        new string[] { ItemCode, LotNo, filter_reg_2 }));
                                                if (tbl_region.Rows.Count > 0)
                                                {
                                                    int count_sample =
                                                        int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
                                                    Export_DatatableImage_Excel_trungang(tbl_region, "Image1", ws,
                                                        cell_Pic1_2, false, count_sample, item.Value);

                                                    List<string> lst_data = tbl_region.AsEnumerable()
                                                        .Select(x => x.Field<string>("Data")).ToList();
                                                    int k = 0;
                                                    for (int i = 0; i < tbl_region.Rows.Count; i++)
                                                    {
                                                        if (i < 2 * count_sample)
                                                        {
                                                            cell_data_2.Offset[0, k].Value = lst_data[i].Split(';')[0];
                                                            cell_data_2.Offset[1, k].Value = lst_data[i].Split(';')[1];
                                                            k++;
                                                            if (i == count_sample - 1)
                                                            {
                                                                cell_data_2 = cell_data_2.Offset[item.Value, 0];
                                                                k = 0;
                                                            }
                                                        }

                                                        if (i == 2 * count_sample - 1)
                                                        {
                                                            break;
                                                        }
                                                    }
                                                }
                                            }

                                            if (item.Key == "1_T")
                                            {
                                                row_begin = row_begin + 3;
                                            }
                                            else if (item.Key == "1_P")
                                            {
                                                row_begin = row_begin + 9;
                                            }
                                        }

                                        break;

                                    case 32:
                                        row_begin = row_begin + 8;
                                        Dictionary<string, int> _dic2 = new Dictionary<string, int> { };
                                        _dic2.Add("1_T", 9);
                                        _dic2.Add("1_P", 3);
                                        _dic2.Add("2_T", 9);
                                        _dic2.Add("2_P", 3);
                                        foreach (var item in _dic2)
                                        {
                                            myExcel.Range cell_Pic1_2 = ws.Cells[row_begin, col_begin];
                                            myExcel.Range cell_data_2 = ws.Cells[row_begin + 1, col_begin];

                                            string filter_reg_2 = "";
                                            if (reg.ToUpper().Contains("MALE"))
                                            {
                                                txt_find = "MALE";
                                            }
                                            else if (reg.ToUpper().Contains("FEMALE"))
                                            {
                                                txt_find = "FEMALE";
                                            }

                                            foreach (string i in lst_region)
                                            {
                                                if (i.Replace(" ", "").ToUpper().Contains("TRUNGANG") &&
                                                    i.Replace(" ", "").ToUpper().Contains(item.Key) &&
                                                    i.ToUpper().Contains(txt_find))
                                                {
                                                    filter_reg_2 = i;
                                                    break;
                                                }
                                            }

                                            if (filter_reg_2 != "")
                                            {
                                                DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet,
                                                    TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" },
                                                        new string[] { ItemCode, LotNo, filter_reg_2 }));
                                                if (tbl_region.Rows.Count > 0)
                                                {
                                                    int count_sample =
                                                        int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
                                                    Export_DatatableImage_Excel_trungang(tbl_region, "Image1", ws,
                                                        cell_Pic1_2, false, count_sample, item.Value);

                                                    List<string> lst_data = tbl_region.AsEnumerable()
                                                        .Select(x => x.Field<string>("Data")).ToList();
                                                    int k = 0;
                                                    for (int i = 0; i < tbl_region.Rows.Count; i++)
                                                    {
                                                        if (i < 2 * count_sample)
                                                        {
                                                            cell_data_2.Offset[0, k].Value = lst_data[i].Split(';')[0];
                                                            cell_data_2.Offset[1, k].Value = lst_data[i].Split(';')[1];
                                                            k++;
                                                            if (i == count_sample - 1)
                                                            {
                                                                cell_data_2 = cell_data_2.Offset[item.Value, 0];
                                                                k = 0;
                                                            }
                                                        }

                                                        if (i == 2 * count_sample - 1)
                                                        {
                                                            break;
                                                        }
                                                    }
                                                }
                                            }

                                            if (item.Key.Contains("T"))
                                            {
                                                row_begin = row_begin + 3;
                                            }
                                            else
                                            {
                                                row_begin = row_begin + 9;
                                            }
                                        }

                                        break;
                                    default:
                                        break;
                                }
                            }

                            if (reg.ToUpper().Contains("Vertical".ToUpper()))
                            {
                                //int count_row = int.Parse(reg.Split(';')[1]);
                                //int row_begin = int.Parse(reg.Split(';')[2]);
                                List<string> lst_key = new List<string> { };
                                if (count_row == 32)
                                {
                                    lst_key = new List<string> { "TRU_T", "TRU_P", "DOC_T", "DOC_P" };
                                }
                                else if (count_row == 48)
                                {
                                    lst_key = new List<string>
                                        { "TRU1_T", "TRU1_P", "TRU2_T", "TRU2_P", "DOC_T", "DOC_P" };
                                }


                                foreach (string item in lst_key)
                                {
                                    myExcel.Range cell_Pic1_2 = ws.Cells[row_begin, col_begin];
                                    myExcel.Range cell_Pic2_2 = ws.Cells[row_begin + 3, col_begin];
                                    myExcel.Range cell_data_2 = ws.Cells[row_begin + 1, col_begin];

                                    string filter_reg_2 = "";

                                    if (reg.ToUpper().Contains("MALE"))
                                    {
                                        txt_find = "MALE";
                                    }
                                    else if (reg.ToUpper().Contains("FEMALE"))
                                    {
                                        txt_find = "FEMALE";
                                    }

                                    foreach (string i in lst_region)
                                    {
                                        if (i.Replace(" ", "").ToUpper().Contains(item) &&
                                            i.ToUpper().Contains(txt_find))
                                        {
                                            filter_reg_2 = i;
                                            break;
                                        }
                                    }

                                    if (filter_reg_2 != "")
                                    {
                                        DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet,
                                            TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" },
                                                new string[] { ItemCode, LotNo, filter_reg_2 }));
                                        if (tbl_region.Rows.Count > 0)
                                        {
                                            int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
                                            Export_DatatableImage_Excel(tbl_region, "Image1", ws, cell_Pic1_2, false,
                                                count_sample);
                                            Export_DatatableImage_Excel(tbl_region, "Image2", ws, cell_Pic2_2, false,
                                                count_sample);

                                            List<string> lst_data = tbl_region.AsEnumerable()
                                                .Select(x => x.Field<string>("Data")).ToList();

                                            for (int i = 0; i < tbl_region.Rows.Count; i++)
                                            {
                                                if (i < count_sample)
                                                {
                                                    cell_data_2.Offset[0, i].Value =
                                                        lst_data[i].Split('/')[0].Split(';')[0];
                                                    cell_data_2.Offset[1, i].Value =
                                                        lst_data[i].Split('/')[0].Split(';')[1];

                                                    //for (int j = 0; j < 4; j++)
                                                    //{
                                                    //    if (j < lst_data[i].Split('/')[1].Split(';').Length)
                                                    //    {
                                                    //        cell_data_2.Offset[3 + j, i].Value = lst_data[i].Split('/')[1].Split(';')[j];
                                                    //    }
                                                    //}

                                                    cell_data_2.Offset[3, i].Value =
                                                        lst_data[i].Split('/')[1].Split(';')[0];
                                                    cell_data_2.Offset[4, i].Value =
                                                        lst_data[i].Split('/')[1].Split(';')[2];
                                                    cell_data_2.Offset[5, i].Value =
                                                        lst_data[i].Split('/')[1].Split(';')[1];
                                                    cell_data_2.Offset[6, i].Value =
                                                        lst_data[i].Split('/')[1].Split(';')[3];
                                                }
                                            }
                                        }
                                    }

                                    row_begin = row_begin + 8;
                                }
                            }
                        }
                    }

                    MessageBox.Show(new Form { TopMost = true }, "Xuất dữ liệu thành công!", "Thông báo");
                }
                // }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Không tìm thấy sheet: " + sheet, "Thông báo");
            }
        }

        public void export_excel_cross_section(string ItemCode, string LotNo, string file_format, string mysheet,
            DataTable Data_tbl, DataTable dt_spec)
        {
            myExcel.Workbook curr_wrkbook = null;
            if (cb_Type.SelectedItem.ToString() == "NPI" || cb_Type.SelectedItem.ToString() == "Other")
            {
                curr_wrkbook = create_export_wrk(file_format, mysheet);
            }
            else if (cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                curr_wrkbook = TDMK_Code.open_excel_file(file_format, "", "");
            }

            if (curr_wrkbook != null)
            {
                myExcel.Worksheet ws = curr_wrkbook.Sheets[1];

                string location_spec = dt_spec.Rows[0]["Location"].ToString();

                int col_begin = int.Parse(location_spec.Split('+')[0].Split(';')[1]);
                string[] region = location_spec.Split('+')[1].Split('_');

                if (cb_Type.SelectedItem.ToString().Contains("MASS") || cb_Type.SelectedItem.ToString() == "Other")
                {
                    string str_infor = Data_tbl.Rows[0]["Sheet"].ToString().Replace(sheet, "").Replace("/", "");
                    if (str_infor.Contains("_"))
                    {
                        string itemname = str_infor.Split('_')[0];
                        string line = str_infor.Split('_')[1];
                        string ca = str_infor.Split('_')[2];
                        string date = str_infor.Split('_')[3];
                        string worker = str_infor.Split('_')[4];
                        List<string> lst_infor = new List<string> { itemname, ItemCode, LotNo, line, ca, date, worker };
                        export_info_mass(ws, lst_infor);
                    }
                }

                List<string> lst_region =
                    Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToList();

                int r_end = 0;
                foreach (string reg in region)
                {
                    if (!reg.Contains("^") && reg != "")
                    {
                        r_end += int.Parse(reg.Split(';')[1]);
                    }
                }

                if (r_end != 0)
                {
                    insert_columns_other(ws, ItemCode, int.Parse(txt_qty.Text), r_end + 2);
                }


                int no_lk = 1;
                int count_sample = int.Parse(txt_qty.Text);
                foreach (string reg in region)
                {
                    if (!reg.Contains("^") && reg != "")
                    {
                        int count_row = int.Parse(reg.Split(';')[1]);
                        int row_begin = int.Parse(reg.Split(';')[2]);
                        if (reg.ToUpper().Contains("Horizontal".ToUpper()))
                        {
                            if (cb_Type.SelectedItem.ToString() == "NPI")
                            {
                                if (count_row == 28)
                                {
                                }
                                else
                                {
                                    export_Hottizontal_NPI(reg, ws, row_begin, col_begin, no_lk, lst_region, count_row,
                                        Data_tbl, dt_spec, ItemCode, LotNo);
                                }
                            }
                            else
                            {
                                export_Hottizontal_MASS(reg, ws, row_begin, col_begin, no_lk, lst_region, count_row,
                                    Data_tbl, dt_spec, ItemCode, LotNo);
                            }
                        }

                        if (reg.ToUpper().Contains("Vertical".ToUpper()))
                        {
                            List<string> lst_key = new List<string> { };
                            if (count_row == 32)
                            {
                                lst_key = new List<string> { "TRU_T", "TRU_P", "DOC_T", "DOC_P" };
                            }
                            else if (count_row == 48)
                            {
                                lst_key = new List<string> { "TRU1_T", "TRU1_P", "TRU2_T", "TRU2_P", "DOC_T", "DOC_P" };
                            }

                            foreach (string item in lst_key)
                            {
                                myExcel.Range cell_Pic1_2 = ws.Cells[row_begin, col_begin];
                                myExcel.Range cell_Pic2_2 = ws.Cells[row_begin + 3, col_begin];
                                myExcel.Range cell_data_2 = ws.Cells[row_begin + 1, col_begin];

                                string filter_reg_2 = "";
                                //if (reg.ToUpper().Contains("MALE"))
                                //{
                                //    txt_find = "MALE"; 
                                //}
                                //else if (reg.ToUpper().Contains("FEMALE"))
                                //{
                                //    txt_find = "FEMALE";
                                //}

                                string txt_find = "";
                                if (reg.ToUpper().Contains("MALE"))
                                {
                                    txt_find = no_lk.ToString() + "_";
                                }

                                foreach (string i in lst_region)
                                {
                                    if (i.Replace(" ", "").ToUpper().Contains(item) && i.ToUpper().Contains(txt_find))
                                    {
                                        filter_reg_2 = i;
                                        break;
                                    }
                                }

                                if (filter_reg_2 != "")
                                {
                                    //DataTable tbl_region = TDMK_Code.Datatable_Filter(sqlcon, sheet, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Region" }, new string[] { ItemCode, LotNo, filter_reg_2 }));


                                    DataView dv = Data_tbl.AsDataView();
                                    dv.RowFilter = TDMK_Code.filter_str(new string[] { "Region" },
                                        new string[] { filter_reg_2 });
                                    DataTable tbl_region = dv.ToTable();


                                    if (tbl_region.Rows.Count > 0)
                                    {
                                        //int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
                                        Export_DatatableImage_Excel(tbl_region, "Image1", ws, cell_Pic1_2, false,
                                            count_sample);
                                        Export_DatatableImage_Excel(tbl_region, "Image2", ws, cell_Pic2_2, false,
                                            count_sample);

                                        List<string> lst_data = tbl_region.AsEnumerable()
                                            .Select(x => x.Field<string>("Data")).ToList();

                                        for (int i = 0; i < tbl_region.Rows.Count; i++)
                                        {
                                            if (i < count_sample)
                                            {
                                                cell_data_2.Offset[0, i].Value =
                                                    lst_data[i].Split('/')[0].Split(';')[0];
                                                cell_data_2.Offset[1, i].Value =
                                                    lst_data[i].Split('/')[0].Split(';')[1];

                                                //for (int j = 0; j < 4; j++)
                                                //{
                                                //    if (j < lst_data[i].Split('/')[1].Split(';').Length)
                                                //    {
                                                //        cell_data_2.Offset[3 + j, i].Value = lst_data[i].Split('/')[1].Split(';')[j];
                                                //    }
                                                //}

                                                cell_data_2.Offset[3, i].Value =
                                                    lst_data[i].Split('/')[1].Split(';')[0];
                                                cell_data_2.Offset[4, i].Value =
                                                    lst_data[i].Split('/')[1].Split(';')[2];
                                                cell_data_2.Offset[5, i].Value =
                                                    lst_data[i].Split('/')[1].Split(';')[1];
                                                cell_data_2.Offset[6, i].Value =
                                                    lst_data[i].Split('/')[1].Split(';')[3];
                                            }
                                        }
                                    }
                                }

                                row_begin = row_begin + 8;
                            }

                            no_lk++;
                        }
                    }
                }

                int r_offset_del = r_end + 3;
                complete_sheet_other(ws, ItemCode, int.Parse(txt_qty.Text), r_offset_del);


                SortedDictionary<int, string> dic_judge = dic_judge_crosscut(count_sample, Data_tbl);


                for (int j = 1; j < 5; j++)
                {
                    for (int i = 1; i < 120; i++)
                    {
                        if (myCode.checkDBNull(ws.Cells[i, j].Value).ToUpper().Replace(" ", "") == "RESULT")
                        {
                            for (int k = 0; k < count_sample; k++)
                            {
                                if (dic_judge.ContainsKey(k + 1))
                                {
                                    ws.Cells[i, j + 2 + k].Value = "Fail";
                                }
                                else
                                {
                                    ws.Cells[i, j + 2 + k].Value = "Pass";
                                }
                            }

                            goto lbl_complete;
                        }
                    }
                }

                lbl_complete:
                if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                {
                    myExcel.Range cur_rgn_judge = find_cell(ws, "JUDGEMENT");
                    if (cur_rgn_judge != null)
                    {
                        if (judge_all)
                        {
                            cur_rgn_judge.Offset[1, 0].Value = "OK";
                            ws.Cells[1, 1].Value = "OK";
                        }
                        else
                        {
                            cur_rgn_judge.Offset[1, 0].Value = "NG";
                            ws.Cells[1, 1].Value = "NG";
                        }
                    }
                }

                MessageBox.Show(new Form { TopMost = true }, "Xuất dữ liệu thành công!", "Thông báo");
            }
            // }

            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Không tìm thấy sheet: " + sheet, "Thông báo");
            }
        }

        //public void export_excel_new(string ItemCode, string LotNo, string file_format, string mysheet)
        //{

        //    if (file_format != "" && txtItemCode.Text != "" && txtLotNo.Text != "")
        //    {
        //        myExcel.Workbook curr_wrkbook = create_export_wrk(file_format, mysheet);
        //        myExcel.Worksheet ws = curr_wrkbook.Sheets[1];

        //        string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet" }, new string[] { ItemCode, mysheet });
        //        DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", filter_str);
        //        DataTable Data_tbl = TDMK_Code.Datatable_Filter(sqlcon,  mysheet, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { ItemCode, LotNo, mysheet }));

        //        //if (Data_tbl.Rows.Count > 0)
        //        //{


        //        if (dt_spec.Rows.Count > 0)
        //        {
        //            string[] region = dt_spec.Rows[0]["Location"].ToString().Split('_');
        //            string[] region_data = Data_tbl.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToArray();
        //            foreach (string reg in region)
        //            {
        //                if (reg != "")
        //                {
        //                    string[] str_location = reg.Split('+');

        //                    // string cpn = str_location[0].Split(';')[0].Replace(" ", "").Replace("JPOND(Peeling", "").Replace(")", "");

        //                    myExcel.Range cell_Pic = ws.Cells[int.Parse(str_location[0].Split(';')[2]), int.Parse(str_location[0].Split(';')[3]) + 1];
        //                    myExcel.Range cell_graph = ws.Cells[int.Parse(str_location[1].Split(';')[1]), int.Parse(str_location[0].Split(';')[3]) + 1];
        //                    myExcel.Range cell_data = ws.Cells[int.Parse(str_location[2].Split(';')[1]), int.Parse(str_location[0].Split(';')[3]) + 1];


        //                    DataTable dt_region = TDMK_Code.Datatable_Filter(sqlcon, mysheet , TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo",  "Region" }, new string[] { ItemCode, LotNo , reg }));

        //                    if (dt_region.Rows.Count > 0)
        //                    {
        //                        Export_DatatableImage_Excel(dt_region, "Image", ws, cell_Pic, false);
        //                        Export_DatatableImage_Excel(dt_region, "Graph", ws, cell_graph, false);

        //                        //List<byte[]> lst_img = Data_tbl.AsEnumerable().Select(x => x.Field<byte[]>("Image")).ToList();
        //                        //List<byte[]> lst_grp = Data_tbl.AsEnumerable().Select(x => x.Field<byte[]>("Graph")).ToList();
        //                        List<string> lst_data = dt_region.AsEnumerable().Select(x => x.Field<string>("Data")).ToList();


        //                        int count_sample = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
        //                        for (int i = 0; i < count_sample; i++)
        //                        {
        //                            if (i < Data_tbl.Rows.Count)
        //                            {
        //                                cell_data.Offset[0, i].Value = lst_data[i].Split('_')[0].Replace("Max:", "");
        //                                cell_data.Offset[1, i].Value = lst_data[i].Split('_')[1].Replace("Average:", "");
        //                            }
        //                        }
        //                    }

        //                }
        //            }
        //            MessageBox.Show(new Form { TopMost = true }, "Export to Checksheet completed!", "Thông báo");
        //        }
        //        // }
        //    }
        //    else
        //    {
        //        MessageBox.Show(new Form { TopMost = true }, "Format of " + txtItemCode.Text + " not found", "Thông báo");

        //    }
        //}
        public string find_format(string in_data_loc, string ItemCode)
        {
            string result = "";
            string folder_format = cb_Type.SelectedItem.ToString();
            if (cb_Type.SelectedItem.ToString().Contains("MASS"))
            {
                folder_format = "MASS";
            }

            if (cb_Type.SelectedItem.ToString() == "NPI" || cb_Type.SelectedItem.ToString() == "LQ")
            {
                folder_format = "NPI";
            }

            if (System.IO.Directory.Exists(Path.Combine(in_data_loc, "Format", folder_format)))
            {
                string path = "";
                if (cb_Type.SelectedItem.ToString() == "NPI" || cb_Type.SelectedItem.ToString() == "LQ")
                {
                    path = Path.Combine(in_data_loc, "Format", folder_format);
                }
                else
                {
                    path = Path.Combine(in_data_loc, "Format", folder_format, sheet);
                }

                if (System.IO.Directory.Exists(path))
                {
                    string[] file_xlsm = Directory.GetFiles(path, "*" + ItemCode + "*.xlsm");

                    if (file_xlsm.Length > 0)
                    {
                        result = file_xlsm[0];
                    }
                    else
                    {
                        string[] file_xlsx = Directory.GetFiles(path, "*" + ItemCode + "*.xlsx");
                        if (file_xlsx.Length > 0)
                        {
                            result = file_xlsx[0];
                            foreach (string f in file_xlsx)
                            {
                                if (f.ToUpper().Replace(" ", "_").Contains(sheet))
                                {
                                    result = f;
                                }
                            }
                        }
                    }
                }
            }

            result = result.Replace("~$", "");
            return result;
        }

        public string find_Folder_Logfile(string in_data_loc, string ItemCode, string LotNo)
        {
            string result = "";
            if (System.IO.Directory.Exists(Path.Combine(in_data_loc)))
            {
                string[] arr_path = Directory.GetDirectories(in_data_loc, "*" + ItemCode + "-" + LotNo + "*");
                if (arr_path.Length > 0)
                {
                    result = arr_path[0];
                }
            }

            return result;
        }

        public void export_old()
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtOperator.Text != "")
            {
                if (dgv_Analysis.DataSource == null)
                {
                    btnLoadb.PerformClick();
                }

                if (cb_Type.SelectedIndex != -1)
                {
                    DataTable dt_spec = new DataTable();
                    if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                    {
                        dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                            TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                                new string[] { txtItemCode.Text, sheet, "MASS" }));
                    }
                    else if (cb_Type.SelectedItem.ToString() == "LQ" || cb_Type.SelectedItem.ToString() == "NPI")
                    {
                        dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                            TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                                new string[] { txtItemCode.Text, sheet, "NPI" }));
                    }

                    if (dt_spec.Rows.Count > 0)
                    {
                        string infor = "/" + txt_ItemName.Text + "_" + txt_line.Text + "_" + txt_ca.Text + "_" +
                                       txt_date.Text + "_" + txt_worker.Text + "_" + cb_Type.SelectedItem.ToString();
                        string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Sheet" },
                            new string[] { txtItemCode.Text, txtLotNo.Text, sheet + infor });

                        if (sheet.Contains("UNMATING") || sheet.Contains("COUPON"))
                        {
                            if (txt_itemcode_nvl.Text != "" && txt_lotno_nvl.Text != "")
                            {
                                filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Sheet" },
                                    new string[] { txt_itemcode_nvl.Text, txt_lotno_nvl.Text, sheet + infor });
                            }
                            else
                            {
                                MessageBox.Show(new Form { TopMost = true }, "Chưa nhập đủ ItemCode/lotNo",
                                    "Thông báo");
                                return;
                            }
                        }

                        DataTable Data_all = TDMK_Code.Datatable_Filter(sqlcon, sheet, filter_str);
                        if (Data_all.Rows.Count > 0)
                        {
                            int st = 1;
                            foreach (DataRow dr in Data_all.Rows)
                            {
                                dr["ID"] = st;
                                st++;
                            }

                            if (txt_qty.Text == "")
                            {
                                txt_qty.Text = dt_spec.Rows[0]["Count_sample"].ToString();
                            }

                            if (txt_qty.Text != "")
                            {
                                lbl_create:

                                string file_format = find_format(data_loc, txtItemCode.Text);
                                string f = file_format;
                                string export_path = "";
                                if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                                {
                                    string report_folder = Path.Combine(data_loc, "Report",
                                        cb_Type.SelectedItem.ToString(), sheet);
                                    if (!System.IO.Directory.Exists(report_folder))
                                        System.IO.Directory.CreateDirectory(report_folder);

                                    // string _fname = Path.GetFileNameWithoutExtension(file_format).TrimEnd(new char[] { ',', '.' });
                                    string _fname = Path.GetFileNameWithoutExtension(file_format);
                                    string name =
                                        Path.GetFileNameWithoutExtension(Data_all.Rows[0]["Remark"].ToString());

                                    _fname = cb_Type.SelectedItem.ToString().Split('(')[1].Split(')')[0] + "_" +
                                             _fname.Replace(name.Split('-')[0], "").Replace(txtItemCode.Text, "")
                                                 .Replace("-", "");
                                    export_path = Path.Combine(report_folder, _fname + "-" + name + ".xlsx");

                                    if (!System.IO.File.Exists(export_path))
                                    {
                                        myExcel.Workbook report_format = TDMK_Code.open_excel_file(file_format, "", "");
                                        report_format.SaveAs(export_path);
                                        report_format.Close();
                                        file_format = export_path;
                                    }
                                    else
                                    {
                                        file_format = "";
                                    }
                                }

                                if (file_format != "")
                                {
                                    judge_all = true;
                                    if (sheet == "GAP_CONNECTOR")
                                    {
                                        export_excel_gap_connector(txtItemCode.Text, txtLotNo.Text, file_format, sheet,
                                            Data_all, dt_spec);
                                    }
                                    else if (sheet == "CROSS_SECTION")
                                    {
                                        export_excel_cross_section(txtItemCode.Text, txtLotNo.Text, file_format, sheet,
                                            Data_all, dt_spec);
                                    }
                                    else if (arr_onproduct.Contains(sheet))
                                    {
                                        export_excel_onproduct(txtItemCode.Text, txtLotNo.Text, file_format, sheet,
                                            Data_all, dt_spec);
                                    }
                                    else if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST" ||
                                             sheet == "SHEAR_TEST")
                                    {
                                        export_excel_peel_pull_shear(txtItemCode.Text, txtLotNo.Text, file_format,
                                            sheet, Data_all, dt_spec);
                                    }
                                    else if (sheet == "IQC_UNMATING_PULL_TEST")
                                    {
                                        export_excel_unmating(txtItemCode.Text, txtLotNo.Text, file_format, sheet,
                                            Data_all, dt_spec);
                                    }
                                    else
                                    {
                                        export_excel_coupon(txtItemCode.Text, txtLotNo.Text, file_format, sheet,
                                            Data_all, dt_spec);
                                    }
                                }
                                else
                                {
                                    if (cb_Type.SelectedItem.ToString() == "NPI" ||
                                        cb_Type.SelectedItem.ToString() == "LQ")
                                    {
                                        MessageBox.Show(new Form { TopMost = true },
                                            "Không tìm thấy format của ItemCode: " + txtItemCode.Text, "");
                                    }
                                    else if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                                    {
                                        if (MessageBox.Show(new Form { TopMost = true },
                                                "Báo cáo đã tồn tại. Bạn có muốn thay thế không?", "Thông báo",
                                                MessageBoxButtons.YesNo) == DialogResult.Yes)
                                        {
                                            try
                                            {
                                                System.IO.File.Delete(export_path);
                                                goto lbl_create;
                                            }
                                            catch
                                            {
                                                MessageBox.Show(new Form { TopMost = true },
                                                    "File excel đang được mở. Vui lòng đóng file !", "Thông báo");
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show(new Form { TopMost = true }, "Chưa cài đặt Qty", "Thông báo");
                            }
                        }

                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Không có dữ liệu", "Thông báo");
                        }
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Chưa cài đặt format cho type này", "Thông báo");
                    }
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Vui lòng chọn type để xuất báo cáo", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },
                    "Vui lòng điền đầy đủ thông tin ItemCode / LotNo / Operator", "Thông báo");
            }
        }

        public void export_new()
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtOperator.Text != "")
            {
                if (cbStatus.Text == "Shield b2b")
                {
                    if (sheet == "CROSS_SECTION")
                    {
                        CrossSectionService.Export(txtItemCode.Text, txtLotNo.Text);
                    return;
                    }

                }
                if (cbStatus.Text == "Shield b2b")
                {
                    if (sheet == "GAP_CONNECTOR")
                    {
                        GAPConnectorService.Export(txtItemCode.Text, txtLotNo.Text);
                    return;
                    }

                }
                if (cbStatus.Text == "Clip")
                {
                    if (sheet == "CROSS_SECTION")
                    {
                        CrossSectionService.ExportClip(txtItemCode.Text, txtLotNo.Text);
                    }

                    return;
                }

                if (dgv_Analysis.DataSource == null)
                {
                    btnLoadb.PerformClick();
                }

                if (cb_Type.SelectedIndex != -1)
                {
                    DataTable dt_spec = new DataTable();
                    if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                    {
                        dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                            TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                                new string[] { txtItemCode.Text, sheet, "MASS" }));
                    }
                    else if (cb_Type.SelectedItem.ToString() == "LQ" || cb_Type.SelectedItem.ToString() == "NPI")
                    {
                        dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                            TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                                new string[] { txtItemCode.Text, sheet, "NPI" }));
                    }

                    if (dt_spec.Rows.Count > 0)
                    {
                        string infor = "/" + txt_ItemName.Text + "_" + txt_line.Text + "_" + txt_ca.Text + "_" +
                                       txt_date.Text + "_" + txt_worker.Text + "_" + cb_Type.SelectedItem.ToString();
                        string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Sheet" },
                            new string[] { txtItemCode.Text.PadRight(10), txtLotNo.Text.PadRight(10), cb_Type.Text });

                        string itemCodeZ = txtItemCode.Text, lotNoZ = txtLotNo.Text;
                        if (sheet.Contains("UNMATING") || sheet.Contains("COUPON"))
                        {
                            if (txt_itemcode_nvl.Text != "" && txt_lotno_nvl.Text != "")
                            {
                                filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Sheet" },
                                    new string[]
                                    {
                                        txt_itemcode_nvl.Text.PadRight(10), txt_lotno_nvl.Text.PadRight(30),
                                        cb_Type.Text
                                    });

                                itemCodeZ = txt_itemcode_nvl.Text;
                                lotNoZ = txt_lotno_nvl.Text;
                            }
                            else
                            {
                                MessageBox.Show(new Form { TopMost = true }, "Chưa nhập đủ ItemCode/lotNo",
                                    "Thông báo");
                                return;
                            }
                        }


                        DataTable Data_all = TDMK_Code.Datatable_Filter(sqlcon,
                            sheet + (_PRIME_PEEL_TEST ? "_WITHOUT_SUS" : "") +
                            (!LegacyMode.Checked == true ? "_NAS" : ""), filter_str);
                        if (!LegacyMode.Checked && Data_all.Rows.Count > 0)
                        {
                            string data = Data_all.Rows[0]["Data"].ToString();
                            string location = Data_all.Rows[0]["LocationImg"].ToString();
                            Data_all = ConverterService.JsonToDataTable(data);
                            NasRepository nas = new NasRepository();
                            nas.MergeDataTable(Data_all, sheet, itemCodeZ, lotNoZ, location);
                        }

                        try
                        {
                            ProductIDService.FillProductID(Data_all, txtItemCode.Text, txtLotNo.Text,
                                sheet + (_PRIME_PEEL_TEST ? "_WITHOUT_SUS" : ""));
                        }
                        catch
                        {
                        }

                        if (Data_all.Rows.Count > 0)
                        {
                            int st = 1;
                            foreach (DataRow dr in Data_all.Rows)
                            {
                                dr["ID"] = st;
                                st++;
                            }

                            if (txt_qty.Text == "")
                            {
                                txt_qty.Text = dt_spec.Rows[0]["Count_sample"].ToString();
                            }

                            if (txt_qty.Text != "")
                            {
                                //Debugger.Break();
                                string file_format = find_format(data_loc, txtItemCode.Text);
                                string f = file_format;
                                string export_path = "";
                                string report_folder = Path.Combine(data_loc, "Report", cb_Type.SelectedItem.ToString(),
                                    sheet);
                                if (!System.IO.Directory.Exists(report_folder))
                                    System.IO.Directory.CreateDirectory(report_folder);
                                if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                                {
                                    string _fname = Path.GetFileNameWithoutExtension(file_format);
                                    string name =
                                        Path.GetFileNameWithoutExtension(Data_all.Rows[0]["Remark"].ToString());

                                    _fname = cb_Type.SelectedItem.ToString().Split('(')[1].Split(')')[0] + "_" +
                                             _fname.Replace(name.Split('-')[0], "").Replace(txtItemCode.Text, "")
                                                 .Replace("-", "");
                                    export_path = Path.Combine(report_folder, _fname + "-" + name + ".xlsx");
                                }
                                else if (cb_Type.SelectedItem.ToString() == "NPI")
                                {
                                    if (sheet.Contains("UNMATING") || sheet.Contains("COUPON"))
                                    {
                                        export_path = Path.Combine(report_folder,
                                            txt_itemcode_nvl.Text + "-" + txt_lotno_nvl.Text + ".xlsx");
                                    }
                                    else
                                    {
                                        export_path = Path.Combine(report_folder,
                                            txtItemCode.Text + "-" + txtLotNo.Text + ".xlsx");
                                    }
                                }

                                if (file_format != "")
                                {
                                    F_export_EPPlus._PRIME_WITHOUT_SUS = _PRIME_PEEL_TEST;
                                    F_export_EPPlus.export_new_viewdata3(file_format, report_folder, export_path,
                                        txtItemCode.Text, txtLotNo.Text, cb_Type.SelectedItem.ToString(), Data_all,
                                        dt_spec, sheet, dgv_Analysis, txtOperator.Text, int.Parse(txt_qty.Text),
                                        txt_itemcode_nvl.Text, txt_lotno_nvl.Text);
                                }
                                else
                                {
                                    MessageBox.Show(new Form { TopMost = true },
                                        "Không tìm thấy format của ItemCode: " + txtItemCode.Text, "");
                                }
                            }
                            else
                            {
                                MessageBox.Show(new Form { TopMost = true }, "Chưa cài đặt Qty", "Thông báo");
                            }
                        }

                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Không có dữ liệu", "Thông báo");
                        }
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Chưa cài đặt format cho type này", "Thông báo");
                    }
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Vui lòng chọn type để xuất báo cáo", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true },
                    "Vui lòng điền đầy đủ thông tin ItemCode / LotNo / Operator", "Thông báo");
            }
        }


        private void btn_export_Click(object sender, EventArgs e)
        {
            //export_old();
            try
            {
                export_new();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtLogfile_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (txtItemCode.Text != "" && txtLotNo.Text != "" && txtOperator.Text != "" && sheet != "")
            {
                FolderBrowserDialog f_open = new FolderBrowserDialog();
                f_open.SelectedPath = System.Windows.Forms.Application.StartupPath;
                if (f_open.ShowDialog() == DialogResult.OK)
                {
                    txtLogfile.Text = f_open.SelectedPath;
                }
            }
        }

        private void txtLogfile_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string str_infor = Path.GetFileName(txtLogfile.Text);
                if (sheet.Contains("UNMATING") || sheet.Contains("COUPON"))
                {
                    if (str_infor.Contains("_"))
                    {
                        //Itemname_itemcode_lotno_ngay-thang_id
                        if (str_infor.Split('_').Length == 5)
                        {
                            //txtLotNo.Text = str_infor.Split('_')[2];
                            txt_itemcode_nvl.Text = str_infor.Split('_')[1].Split('-')[0];
                            txt_lotno_nvl.Text = str_infor.Split('_')[2];

                            if (cb_Type.SelectedIndex != -1)
                            {
                                if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                                {
                                    txt_ItemName.Text = str_infor.Split('_')[0];
                                    txt_date.Text = str_infor.Split('_')[3];
                                    txt_worker.Text = str_infor.Split('_')[4];
                                }
                            }
                        }
                    }
                }
                else if (cb_Type.SelectedIndex != -1)
                {
                    if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                    {
                        if (str_infor.Contains("_"))
                            str_infor = str_infor.Split('_')[0];

                        string[] arr_infor = str_infor.Split('-');

                        //if (sheet.Contains("ON_PRODUCT"))
                        //{
                        //    if (cb_Type.SelectedItem.ToString() == "MASS(IQC)")
                        //    {
                        //        txtLotNo.Text = arr_infor[2];
                        //    }
                        //}
                        if (arr_infor.Length == 8)
                        {
                            txt_ItemName.Text = arr_infor[0];
                            txt_line.Text = arr_infor[3];
                            txt_ca.Text = arr_infor[4];
                            txt_date.Text = arr_infor[5] + "-" + arr_infor[6];
                            txt_worker.Text = arr_infor[7];
                        }
                        else if (arr_infor.Length == 9)
                        {
                            txt_ItemName.Text = arr_infor[0];
                            txt_line.Text = arr_infor[3] + "-" + arr_infor[4];
                            txt_ca.Text = arr_infor[5];
                            txt_date.Text = arr_infor[6] + "-" + arr_infor[7];
                            txt_worker.Text = arr_infor[8];
                        }
                        else if (arr_infor.Length == 6)
                        {
                            txt_ItemName.Text = arr_infor[0];
                            txt_date.Text = arr_infor[3] + "-" + arr_infor[4];
                            txt_worker.Text = arr_infor[5];
                        }
                    }
                }
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {
        }

        private void txtLotNo_TextChanged(object sender, EventArgs e)
        {
            dgv_Analysis.DataSource = dgv_logfile.DataSource = null;
            lbl_judge_logfile.BackColor = Color.Transparent;
            lbl_judge.BackColor = Color.Transparent;
            lbl_judge.Text = "";
            txt_selected.Text = "";
            // txtLogfile.Text = "";
        }

        private void label4_Click(object sender, EventArgs e)
        {
        }

        private void txtOperator_TextChanged(object sender, EventArgs e)
        {
        }

        private void tableLayoutPanel3_Paint(object sender, PaintEventArgs e)
        {
        }

        private void txtItemCode_TextChanged(object sender, EventArgs e)
        {
            if (!sheet.Contains("UNMATING") && !sheet.Contains("COUPON") && !sheet.Contains("GAP") &&
                !sheet.Contains("CROSS"))
            {
                txtLotNo.Text = "";
            }

            txtLogfile.Text = "";
            txt_qty.Text = "";
            cb_Type.SelectedIndex = -1;
            txt_ItemName.Text = "";
            txt_line.Text = "";
            txt_ca.Text = "";
            txt_date.Text = "";
            txt_itemcode_nvl.Text = "";
            txt_lotno_nvl.Text = "";
            set_chan = 0;
            txt_setchan.Text = "";


            dgv_Analysis.DataSource = dgv_logfile.DataSource = null;
            lbl_judge_logfile.BackColor = Color.Transparent;
            lbl_judge.Text = "";
            lbl_judge.BackColor = Color.Transparent;
            txt_selected.Text = "";
        }

        private void checkSession()
        {
            if (UserSession.Instance.IsLoggedIn)
            {
                txtOperator.Text = UserSession.Instance.User_ID;
                lbl_Login.Text = "Logout";
                lbl_Login.ForeColor = Color.Red;
            }
            else
            {
                lbl_Login.ForeColor = Color.Blue;
                lbl_Login.Text = "Login";
                txtOperator.Text = "";
            }
        }

        private void lbl_Login_Click(object sender, EventArgs e)
        {
            if (UserSession.Instance.IsLoggedIn)
            {
                UserSession.Instance.Logout();
            }
            else
            {
                Login fr1 = new Login(Setmode);
                fr1.ShowDialog();
            }

            checkSession();
        }


        private void dgv_logfile_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            string col_name = dgv_logfile.Columns[e.ColumnIndex].Name;

            if (col_name != "Select" || col_name != "Select_Img" || col_name != "Select_Grp")
            {
                for (int i = 0; i < dgv_logfile.Rows.Count; i++)
                {
                    dgv_logfile.Rows[i].Cells["Data"].Style.BackColor = Color.White;
                }
            }

            lbl_judge_logfile.BackColor = Color.Transparent;
            Dictionary<string, string> dic_mode = new Dictionary<string, string> { };
            dic_mode.Add("Mode 1#", "Mode 1: Solder joint crack");
            dic_mode.Add("Mode 2#", "Mode 2: Pad lift");
            dic_mode.Add("Mode 3#", "Mode 3: Solder joint lift");
            dic_mode.Add("Mode 4#", "Mode 4: Intermetallic break");
            dic_mode.Add("Mode 5#", "Mode 5: Component damage");
            dic_mode.Add("Mode 6#", "Mode 6: Component detached");
            dic_mode.Add("Mode 7#", "Mode 7: Flex torn");


            List<string> lst_col_name = new List<string>
            {
                "Mode 1: Solder joint crack", "Mode 2: Pad lift", "Mode 3: Solder joint lift",
                "Mode 4: Intermetallic break", "Mode 5: Component damage", "Mode 6: Component detached",
                "Mode 7: Flex torn"
            };
            if (lst_col_name.Contains(col_name))
            {
                dgv_logfile.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.White;

                if (!dgv_logfile.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Contains("%"))
                {
                    double data = 0;

                    if (dgv_logfile.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Contains("/"))
                    {
                        try
                        {
                            string val_change = dgv_logfile.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()
                                .Replace(" ", "");
                            int a = int.Parse(val_change.Split('/')[0]);
                            int b = int.Parse(val_change.Split('/')[1]);
                            data = Math.Round((double)a / b, 4);

                            if (data <= 1)
                            {
                                dgv_logfile.Rows[e.RowIndex].Cells[e.ColumnIndex].Value =
                                    (data * 100).ToString() + "%" + "( " + val_change + " )";
                                int a_col = 0;
                                double data_col = 0;
                                for (int i = 2; i <= 7; i++)
                                {
                                    string val = dgv_logfile.Rows[e.RowIndex].Cells[dic_mode["Mode " + i + "#"]].Value
                                        .ToString();
                                    if (myCode.IsNumeric(val.Split('%')[0]) && val.Contains("("))
                                    {
                                        data_col += double.Parse(val.Split('%')[0]);
                                        a_col += int.Parse(val.Split('(')[1].Split('/')[0]);
                                    }
                                }

                                dgv_logfile.Rows[e.RowIndex].Cells[dic_mode["Mode 1#"]].Value =
                                    (100 - data_col).ToString() + "%" + "(" + (b - a_col).ToString() + "/" +
                                    b.ToString() + ")";
                            }
                            else
                            {
                                dgv_logfile.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Red;
                            }
                        }
                        catch
                        {
                            dgv_logfile.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Red;
                        }
                    }
                    else
                    {
                        dgv_logfile.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "";
                        dgv_logfile.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Red;
                    }

                    //if (data > 1)
                    //{
                    //    dgv_logfile.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Red;
                    //}
                }
            }
            else if (col_name.Contains("Data"))
            {
                dgv_logfile.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.White;
                if (cb_Type.SelectedIndex != -1)
                {
                    switch (sheet)
                    {
                        case "PEEL_TEST":
                            Check_spec_Peel_Pull(dgv_logfile, sheet);
                            break;

                        case "MATING_PULL_TEST":
                            Check_spec_Peel_Pull(dgv_logfile, sheet);
                            break;

                        case "SHEAR_TEST":
                            Check_spec_ShearTest_new(dgv_logfile);
                            break;

                        case "IQC_UNMATING_PULL_TEST":
                            Check_spec_unmatingpull(dgv_logfile, sheet);
                            break;

                        case "IQC_LINER_PEELING_COUPON":
                            check_spec_coupon(dgv_logfile, sheet);
                            break;

                        case "IQC_PSA_PEELING_COUPON":
                            check_spec_coupon(dgv_logfile, sheet);
                            break;

                        case "LINER_PEEL_TEST_ON_PRODUCT":
                            check_spec_onproduct(dgv_logfile, sheet);
                            break;

                        case "PSA_PEEL_TEST_ON_PRODUCT":
                            check_spec_onproduct(dgv_logfile, sheet);
                            break;

                        case "CROSS_SECTION":
                            Check_spec_crossection(dgv_logfile, sheet);
                            break;

                        case "GAP_CONNECTOR":
                            Check_spec_GAP(dgv_logfile, sheet);

                            break;
                    }

                    Check_Alldata(dgv_logfile);
                }
            }
            else if (col_name == "Select" || col_name == "Select_Img" || col_name == "Select_Grp")
            {
                txt_selected.Text = count_selected();
                if (cb_Type.SelectedIndex != -1)
                {
                    switch (sheet)
                    {
                        case "PEEL_TEST":
                            Check_spec_Peel_Pull(dgv_logfile, sheet);
                            break;

                        case "MATING_PULL_TEST":
                            Check_spec_Peel_Pull(dgv_logfile, sheet);
                            break;

                        case "SHEAR_TEST":
                            Check_spec_ShearTest_new(dgv_logfile);
                            break;

                        case "IQC_UNMATING_PULL_TEST":
                            Check_spec_unmatingpull(dgv_logfile, sheet);
                            break;

                        case "IQC_LINER_PEELING_COUPON":
                            check_spec_coupon(dgv_logfile, sheet);
                            break;

                        case "IQC_PSA_PEELING_COUPON":
                            check_spec_coupon(dgv_logfile, sheet);
                            break;

                        case "LINER_PEEL_TEST_ON_PRODUCT":
                            check_spec_onproduct(dgv_logfile, sheet);
                            break;

                        case "PSA_PEEL_TEST_ON_PRODUCT":
                            check_spec_onproduct(dgv_logfile, sheet);
                            break;

                        case "CROSS_SECTION":
                            Check_spec_crossection(dgv_logfile, sheet);
                            break;

                        case "GAP_CONNECTOR":
                            Check_spec_GAP(dgv_logfile, sheet);
                            break;
                    }

                    Check_Alldata(dgv_logfile);
                }
            }


            for (int i = 0; i < dgv_logfile.Rows.Count; i++)
            {
                if (dgv_logfile.Rows[i].Cells["Data"].Style.BackColor == Color.Red)
                {
                    lbl_judge_logfile.BackColor = Color.Red;
                    break;
                }
            }
        }

        private void btn_checkall_Click(object sender, EventArgs e)
        {
            string[] lst_col_select = new string[3] { "Select", "Select_Img", "Select_Grp" };
            if (dgv_logfile.Rows.Count > 0)
            {
                if (btn_checkall.Text == "Check All")
                {
                    foreach (string col_name in lst_col_select)
                    {
                        if (dgv_logfile.Columns.Contains(col_name))
                        {
                            foreach (DataGridViewRow dr in dgv_logfile.Rows)
                            {
                                dr.Cells[col_name].Value = true;
                            }
                        }
                    }

                    btn_checkall.Text = "Uncheck All";
                }
                else
                {
                    foreach (string col_name in lst_col_select)
                    {
                        if (dgv_logfile.Columns.Contains(col_name))
                        {
                            foreach (DataGridViewRow dr in dgv_logfile.Rows)
                            {
                                dr.Cells[col_name].Value = false;
                            }
                        }
                    }

                    btn_checkall.Text = "Check All";
                }
            }
        }

        private void dgv_logfile_MouseClick(object sender, MouseEventArgs e)
        {
            //if (e.Button == MouseButtons.Right)
            //{  
            //    cmsPaste.Show(dgv_logfile, e.Location);
            //}
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PasteClipboardValue(false, dgv_logfile);
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            //DataTable dt_update = (DataTable)dgv_Analysis.DataSource;
            //dt_update.Columns.Add("Select", typeof(bool));
            //for(int i = 0; i < dgv_Analysis.Rows.Count; i++)
            //{
            //    dgv_Analysis.Rows[0].Cells["Select"].Value = false;
            //}
        }

        private void btnEdit_Click_1(object sender, EventArgs e)
        {
            if (admin_mode == "Admin mode")
            {
                if (!edit_mode)
                {
                    edit_mode = true;
                    btnEdit.BackColor = Color.GreenYellow;
                    //  dgv_Analysis.ReadOnly = false;

                    if (sheet != "SHEAR_TEST")
                    {
                        dgv_Analysis.Columns["Data"].ReadOnly = true;
                    }

                    if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST" || sheet == "SHEAR_TEST")
                    {
                        if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST" || sheet == "SHEAR_TEST")
                        {
                            dgv_Analysis.Columns["Mode 1: Solder joint crack"].ReadOnly = true;
                            dgv_Analysis.Columns["Mode 3: Solder joint lift"].ReadOnly = true;
                            dgv_Analysis.Columns["Mode 4: Intermetallic break"].ReadOnly = true;
                            dgv_Analysis.Columns["Mode 6: Component detached"].ReadOnly = true;
                            dgv_Analysis.Columns["Mode 7: Flex torn"].ReadOnly = true;
                        }
                    }
                }
                else
                {
                    edit_mode = false;
                    btnEdit.BackColor = Color.Transparent;
                    dgv_Analysis.ReadOnly = true;
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Vui lòng đăng nhập để chỉnh sửa dữ liệu", "Thông báo");
            }
        }

        private void dgv_Analysis_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //int col = dgv_Analysis.Columns.Count;
            //if(e.ColumnIndex == col - 1)
            //{
            //    bool checkedCell = (bool)dgv_Analysis.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            //    if (checkedCell == true)
            //    {
            //        for (int k = 0; k < dgv_Analysis.Columns.Count; k++)
            //        {
            //            dgv_Analysis.Rows[e.RowIndex].Cells[k].Style.BackColor = Color.LightGray;
            //        }

            //    }
            //    else
            //    {
            //        for (int k = 0; k < dgv_Analysis.Columns.Count; k++)
            //        {
            //            dgv_Analysis.Rows[e.RowIndex].Cells[k].Style.BackColor = Color.White;
            //        }
            //    }
            //}
        }

        private void dgv_Analysis_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            //int col = dgv_Analysis.Columns.Count;
            //if (e.ColumnIndex == col - 1)
            //{
            //    bool checkedCell = (bool)dgv_Analysis.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            //    if (checkedCell == true)
            //    {
            //        for (int k = 0; k < dgv_Analysis.Columns.Count; k++)
            //        {
            //            dgv_Analysis.Rows[e.RowIndex].Cells[k].Style.BackColor = Color.LightGray;
            //        }

            //    }
            //    else
            //    {
            //        for (int k = 0; k < dgv_Analysis.Columns.Count; k++)
            //        {
            //            dgv_Analysis.Rows[e.RowIndex].Cells[k].Style.BackColor = Color.White;
            //        }
            //    }
            //}
        }

        private void dgv_Analysis_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            //int col = dgv_Analysis.Columns.Count;
            //if (e.ColumnIndex == col - 1)
            //{
            //    bool checkedCell = (bool)dgv_Analysis.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            //    if (checkedCell == true)
            //    {
            //        for (int k = 0; k < dgv_Analysis.Columns.Count; k++)
            //        {
            //            dgv_Analysis.Rows[e.RowIndex].Cells[k].Style.BackColor = Color.LightGray;
            //        }

            //    }
            //    else
            //    {
            //        for (int k = 0; k < dgv_Analysis.Columns.Count; k++)
            //        {
            //            dgv_Analysis.Rows[e.RowIndex].Cells[k].Style.BackColor = Color.White;
            //        }
            //    }
            //}
        }

        private void dgv_Analysis_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int col_inx = e.ColumnIndex;
            int r_inx = e.RowIndex;
            DataGridViewCell cur_cell = dgv_Analysis.CurrentCell;
            if (dgv_Analysis.Columns[col_inx].Name.Contains("Image") ||
                dgv_Analysis.Columns[col_inx].Name.Contains("Graph"))
            {
                if (myCode.checkDBNull(cur_cell.Value) != "")
                {
                    View_detail_Image fr1 = new View_detail_Image() { TopMost = true };
                    fr1.data = (byte[])cur_cell.Value;
                    fr1.Show();
                }
            }
        }

        private void dgv_Analysis_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //int col_inx = e.ColumnIndex;
            //int r_inx = e.RowIndex;
            //if (r_inx != -1)
            //{
            //    // int col = dgv_Analysis.Columns.Count;
            //    if (edit_mode)
            //    {
            //        if (!dgv_Analysis.Columns[col_inx].Name.Contains("Image") && !dgv_Analysis.Columns[col_inx].Name.Contains("Select"))
            //        {
            //            bool checkedCell = (bool)dgv_Analysis.Rows[r_inx].Cells["Select"].Value;
            //            if (checkedCell == false)
            //            {
            //                for (int k = 0; k < dgv_Analysis.Columns.Count; k++)
            //                {
            //                    dgv_Analysis.Rows[r_inx].Cells[k].Style.BackColor = Color.LightGray;
            //                    dgv_Analysis.Rows[r_inx].Cells["Select"].Value = true;
            //                }
            //            }
            //            else
            //            {
            //                for (int k = 0; k < dgv_Analysis.Columns.Count; k++)
            //                {
            //                    dgv_Analysis.Rows[r_inx].Cells[k].Style.BackColor = Color.White;
            //                    dgv_Analysis.Rows[r_inx].Cells["Select"].Value = false;
            //                }
            //            }
            //        }
            //    }
            //}
        }


        private void dgv_Analysis_CellValueChanged_1(object sender, DataGridViewCellEventArgs e)
        {
            if (edit_mode)
            {
                for (int i = 0; i < dgv_Analysis.Rows.Count; i++)
                {
                    dgv_Analysis.Rows[i].Cells["Data"].Style.BackColor = Color.White;
                }

                string colname = dgv_Analysis.Columns[e.ColumnIndex].Name;
                if (colname.Contains("Data") && cb_Type.SelectedIndex != -1)
                {
                    switch (sheet)
                    {
                        case "PEEL_TEST":
                            Check_spec_Peel_Pull(dgv_Analysis, sheet);
                            break;

                        case "MATING_PULL_TEST":
                            Check_spec_Peel_Pull(dgv_Analysis, sheet);
                            break;

                        case "SHEAR_TEST":
                            Check_spec_ShearTest_new(dgv_Analysis);
                            break;

                        case "IQC_UNMATING_PULL_TEST":
                            Check_spec_unmatingpull(dgv_Analysis, sheet);
                            break;

                        case "IQC_LINER_PEELING_COUPON":
                            check_spec_coupon(dgv_Analysis, sheet);
                            break;

                        case "IQC_PSA_PEELING_COUPON":
                            check_spec_coupon(dgv_Analysis, sheet);
                            break;

                        case "LINER_PEEL_TEST_ON_PRODUCT":
                            check_spec_onproduct(dgv_Analysis, sheet);
                            break;

                        case "PSA_PEEL_TEST_ON_PRODUCT":
                            check_spec_onproduct(dgv_Analysis, sheet);
                            break;

                        case "CROSS_SECTION":
                            Check_spec_crossection(dgv_Analysis, sheet);
                            break;

                        case "GAP_CONNECTOR":
                            Check_spec_GAP(dgv_Analysis, sheet);
                            break;
                    }

                    Check_Alldata(dgv_Analysis);
                }

                Dictionary<string, string> dic_mode = new Dictionary<string, string> { };
                dic_mode.Add("Mode 1#", "Mode 1: Solder joint crack");
                dic_mode.Add("Mode 2#", "Mode 2: Pad lift");
                dic_mode.Add("Mode 3#", "Mode 3: Solder joint lift");
                dic_mode.Add("Mode 4#", "Mode 4: Intermetallic break");
                dic_mode.Add("Mode 5#", "Mode 5: Component damage");
                dic_mode.Add("Mode 6#", "Mode 6: Component detached");
                dic_mode.Add("Mode 7#", "Mode 7: Flex torn");


                List<string> lst_col_name = new List<string>
                {
                    "Mode 1: Solder joint crack", "Mode 2: Pad lift", "Mode 3: Solder joint lift",
                    "Mode 4: Intermetallic break", "Mode 5: Component damage", "Mode 6: Component detached",
                    "Mode 7: Flex torn"
                };
                if (lst_col_name.Contains(colname))
                {
                    for (int i = 0; i < dgv_Analysis.Rows.Count; i++)
                    {
                        dgv_Analysis.Rows[i].Cells[colname].Style.BackColor = Color.White;
                    }

                    if (!dgv_Analysis.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Contains("%"))
                    {
                        double data = 0;
                        if (dgv_Analysis.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Contains("/"))
                        {
                            try
                            {
                                string val_change = dgv_Analysis.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()
                                    .Replace(" ", "");
                                int a = int.Parse(val_change.Split('/')[0]);
                                int b = int.Parse(val_change.Split('/')[1]);
                                data = Math.Round((double)a / b, 4);

                                if (data <= 1)
                                {
                                    dgv_Analysis.Rows[e.RowIndex].Cells[e.ColumnIndex].Value =
                                        (data * 100).ToString() + "%" + "( " + val_change + " )";
                                    int a_col = 0;
                                    double data_col = 0;
                                    for (int i = 2; i <= 7; i++)
                                    {
                                        string val = dgv_Analysis.Rows[e.RowIndex].Cells[dic_mode["Mode " + i + "#"]]
                                            .Value.ToString();
                                        if (myCode.IsNumeric(val.Split('%')[0]) && val.Contains("("))
                                        {
                                            data_col += double.Parse(val.Split('%')[0]);
                                            a_col += int.Parse(val.Split('(')[1].Split('/')[0]);
                                        }
                                    }

                                    dgv_Analysis.Rows[e.RowIndex].Cells[dic_mode["Mode 1#"]].Value =
                                        (100 - data_col).ToString() + "%" + "(" + (b - a_col).ToString() + "/" +
                                        b.ToString() + ")";
                                }
                                else
                                {
                                    dgv_Analysis.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Red;
                                }
                            }
                            catch
                            {
                                dgv_Analysis.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Red;
                            }
                        }
                        else
                        {
                            dgv_Analysis.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "";
                            dgv_Analysis.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Red;
                        }
                    }
                }
            }
        }

        private void dgv_logfile_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void txtLotNo_Validated(object sender, EventArgs e)
        {
            //if (!sheet.Contains("UNMATING"))
            //{
            //if (cb_Type.SelectedIndex == -1)
            //{ 
            txtLotNo.Text = Lotno_Formated(txtLotNo.Text);
            // }
            //else if (!sheet.Contains("ON_PRODUCT") || cb_Type.SelectedItem.ToString() != "MASS(IQC)")
            //{
            //    txtLotNo.Text = Lotno_Formated(txtLotNo.Text);
            //}
            // }
        }

        public string Lotno_Formated_old(string lotno)
        {
            string result = "";
            if (lotno.All(char.IsDigit))
            {
                try
                {
                    result = string.Format("{0:00000}", Convert.ToInt32(lotno));
                    result = result.Substring(0, 5);
                }
                catch
                {
                }
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
            }

            return result;
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
                catch
                {
                }
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

        private void txtItemCode_Validated(object sender, EventArgs e)
        {
            txtItemCode.Text = txtItemCode.Text.Replace(" ", "").ToUpper();
        }

        private void dgv_logfile_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
        }

        private void dgv_logfile_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            Hide_column(dgv_logfile, ref hide_mode_logfile);
        }

        private void dgv_Analysis_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            Hide_column(dgv_Analysis, ref hide_mode_analysis);
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {
        }

        private void dgv_logfile_CellMouseMove(object sender, DataGridViewCellMouseEventArgs e)
        {
        }

        private void dgv_logfile_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.RowIndex != -1 && e.ColumnIndex != -1)
            //{
            //    string region = dgv_logfile.Rows[e.RowIndex].Cells["Region"].Value.ToString();
            //    string filter = TDMK_Code.filter_str(new string[] {"Region"}, new string[] {region});
            //    DataView dv = ((DataTable)dgv_logfile.DataSource).AsDataView();
            //    dv.RowFilter = filter;
            //    DataTable dt_filter = dv.ToTable();
            //    List<double> lst_data = new List<double> { };
            //    foreach(DataRow dr in dt_filter.Rows)
            //    {
            //        if (myCode.IsNumeric(dr["Data"].ToString()))
            //        {
            //            lst_data.Add(Double.Parse(dr["Data"].ToString()));
            //        }
            //    }

            //    string max = lst_data.ToArray().Max().ToString();
            //    string min = lst_data.ToArray().Min().ToString();
            //    string Average = Math.Round( lst_data.ToArray().Average(), 2).ToString();


            //    if (dgv_logfile.Columns[e.ColumnIndex].Name == "Data")
            //    {
            //        MessageBox.Show("Min: " + min + "\n" + "Max: " + max + "\n" + "Average: " + Average);
            //    }
            //}
        }

        private void dgv_logfile_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            if (e.RowIndex != -1 && e.ColumnIndex != -1)
            {
                string region = dgv_logfile.Rows[e.RowIndex].Cells["Region"].Value.ToString();
                string filter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { region });
                DataView dv = ((DataTable)dgv_logfile.DataSource).AsDataView();
                dv.RowFilter = filter;
                DataTable dt_filter = dv.ToTable();
                List<double> lst_data = new List<double> { };

                DataTable dt_spec = new DataTable();
                // DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, sheet, cb_Type.SelectedItem.ToString() }));
                if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                {
                    dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                        TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                            new string[] { txtItemCode.Text, sheet, "MASS" }));
                }
                else if (cb_Type.SelectedItem.ToString() == "LQ" || cb_Type.SelectedItem.ToString() == "NPI")
                {
                    dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                        TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                            new string[] { txtItemCode.Text, sheet, "NPI" }));
                }

                if (dt_spec.Rows.Count > 0)
                {
                    if (sheet == "CROSS_SECTION")
                    {
                        if (dgv_logfile.Rows[e.RowIndex].Cells["Data"].Style.BackColor == Color.Red)
                            dgv_logfile.CurrentCell.ToolTipText =
                                msg_ToolTipneeeded_crosssection(dt_spec, dgv_logfile, e.RowIndex);
                    }
                    else if (sheet == "GAP_CONNECTOR")
                    {
                        if (dgv_logfile.Rows[e.RowIndex].Cells["Data"].Style.BackColor == Color.Red)
                            dgv_logfile.CurrentCell.ToolTipText =
                                msg_ToolTipneeeded_GAP(dt_spec, dgv_logfile, e.RowIndex);
                    }
                    else if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST")
                    {
                        if (dgv_logfile.Rows[e.RowIndex].Cells["Data"].Style.BackColor == Color.Red)
                            dgv_logfile.CurrentCell.ToolTipText =
                                msg_ToolTipneeeded_peelpull(dt_spec, dgv_logfile, e.RowIndex);
                    }
                    else if (sheet.Contains("ON_PRODUCT"))
                    {
                        foreach (DataRow dr in dt_filter.Rows)
                        {
                            if (myCode.IsNumeric(dr["Data"].ToString().Split('_')[1].Replace("Average:", "")))
                            {
                                lst_data.Add(Double.Parse(dr["Data"].ToString().Split('_')[1].Replace("Average:", "")));
                            }
                        }
                    }
                    else if (sheet.Contains("COUPON"))
                    {
                        if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                        {
                            foreach (DataRow dr in dt_filter.Rows)
                            {
                                if (dr["Data"].ToString().Contains("_"))
                                {
                                    if (myCode.IsNumeric(dr["Data"].ToString().Split('_')[1].Replace("Average:", "")))
                                    {
                                        lst_data.Add(Double.Parse(dr["Data"].ToString().Split('_')[1]
                                            .Replace("Average:", "")));
                                    }
                                }
                            }
                        }
                        else
                        {
                            foreach (DataRow dr in dt_filter.Rows)
                            {
                                if (myCode.IsNumeric(dr["Data"].ToString()))
                                {
                                    lst_data.Add(Double.Parse(dr["Data"].ToString()));
                                }
                            }
                        }
                    }
                    else if (sheet != "GAP_CONNECTOR")
                    {
                        foreach (DataRow dr in dt_filter.Rows)
                        {
                            if (myCode.IsNumeric(dr["Data"].ToString()))
                            {
                                lst_data.Add(Double.Parse(dr["Data"].ToString()));
                            }
                        }
                    }


                    if (lst_data.Count > 0)
                    {
                        string max = lst_data.ToArray().Max().ToString();
                        string min = lst_data.ToArray().Min().ToString();
                        string Average = Math.Round(lst_data.ToArray().Average(), 2).ToString();
                        string spec_type2 = "";
                        string spec_cp = "";
                        if (sheet.Contains("COUPON"))
                        {
                            if (cb_Type.SelectedItem.ToString() == "NPI")
                            {
                                spec_cp = spec_coupon(sheet, region);
                            }
                            else
                            {
                                if (dt_spec.Rows.Count > 0)
                                {
                                    foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                                    {
                                        string key1 = region.Split('_')[0].Replace(" ", "").ToUpper();
                                        string key2 = region.Split('_')[1].Replace(" ", "").ToUpper();

                                        if (spec_region.Replace(" ", "").ToUpper().Contains(key1) &&
                                            spec_region.Replace(" ", "").ToUpper().Contains(key2))
                                        {
                                            string R = "";
                                            string UCL = "";
                                            string LCL = "";

                                            try
                                            {
                                                R = spec_region.Split('+')[6].Split(';')[0];
                                                UCL = spec_region.Split('+')[6].Split(';')[1];
                                                LCL = spec_region.Split('+')[6].Split(';')[2];
                                            }
                                            catch
                                            {
                                                R = spec_region.Split('+')[5].Split(';')[0];
                                                UCL = spec_region.Split('+')[5].Split(';')[1];
                                                LCL = spec_region.Split('+')[5].Split(';')[2];
                                            }


                                            spec_type2 += "R:" + R + "\n";
                                            spec_type2 += "UCL:" + UCL + "\n";
                                            spec_type2 += "LCL:" + LCL;

                                            break;
                                        }
                                    }
                                }
                            }
                        }

                        if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST" || sheet == "SHEAR_TEST")
                        {
                            if (cb_Type.SelectedIndex != -1)
                            {
                                List<DataTable> lst_Table = new List<DataTable> { };

                                if (dt_spec.Rows.Count > 0)
                                {
                                    string spec_region =
                                        dt_spec.Rows[0]["Location"].ToString().Split('_')[int.Parse(region) - 1];
                                    string type = spec_region.Split('+')[0];
                                    if (type == "B")
                                    {
                                        if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST")
                                        {
                                            string spec = spec_region.Split('+')[4];
                                            if (myCode.IsNumeric(spec))
                                            {
                                                spec_type2 += "R:" + spec_region.Split('+')[5].Split(';')[0] + "\n";
                                                spec_type2 += "UCL:" + spec_region.Split('+')[5].Split(';')[1] + "\n";
                                                spec_type2 += "LCL:" + spec_region.Split('+')[5].Split(';')[2];
                                            }
                                        }
                                        else if (sheet == "SHEAR_TEST")
                                        {
                                            spec_type2 += "R:" + spec_region.Split('+')[4].Split(';')[0] + "\n";
                                            spec_type2 += "UCL:" + spec_region.Split('+')[4].Split(';')[1] + "\n";
                                            spec_type2 += "LCL:" + spec_region.Split('+')[4].Split(';')[2];
                                        }
                                    }
                                }
                            }
                        }

                        if (sheet.Contains("ON_PRODUCT"))
                        {
                            if (dt_spec.Rows.Count > 0)
                            {
                                foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
                                {
                                    if (spec_region.Replace(" ", "").ToUpper().Replace("COMPONENT", "TAPE")
                                            .Contains(region.Split('_')[0].Replace(" ", "").ToUpper()) &&
                                        spec_region.Replace(" ", "").ToUpper().Replace("COMPONENT", "TAPE")
                                            .Contains(region.Split('_')[1].Replace(" ", "").ToUpper()))
                                    {
                                        string type = spec_region.Split('+')[0];
                                        if (type == "B")
                                        {
                                            string R = "";
                                            string UCL = "";
                                            string LCL = "";
                                            try
                                            {
                                                R = spec_region.Split('+')[6].Split(';')[0];
                                                UCL = spec_region.Split('+')[6].Split(';')[1];
                                                LCL = spec_region.Split('+')[6].Split(';')[2];
                                            }
                                            catch
                                            {
                                                R = spec_region.Split('+')[5].Split(';')[0];
                                                UCL = spec_region.Split('+')[5].Split(';')[1];
                                                LCL = spec_region.Split('+')[5].Split(';')[2];
                                            }


                                            spec_type2 += "R:" + R + "\n";
                                            spec_type2 += "UCL:" + UCL + "\n";
                                            spec_type2 += "LCL:" + LCL;
                                        }
                                        else if (type == "A")
                                        {
                                            string spec_max = spec_region.Split('+')[3].Split(';')[0];
                                            string spec_ave = spec_region.Split('+')[4].Split(';')[0];
                                            spec_type2 += spec_max + "\n" + spec_ave;
                                        }

                                        break;
                                    }
                                }
                            }
                        }

                        if (sheet == "IQC_UNMATING_PULL_TEST")
                        {
                            if (dt_spec.Rows.Count > 0)
                            {
                                spec_type2 += dt_spec.Rows[0]["Location"].ToString().Split('+')[3].Split(';')[0];
                            }
                        }

                        if (dgv_logfile.Columns[e.ColumnIndex].Name == "Data")
                        {
                            if (dgv_logfile.CurrentCell != null)
                            {
                                dgv_logfile.CurrentCell.ToolTipText = "Min: " + min + "\n" + "Max: " + max + "\n" +
                                                                      "Average: " + Average + "\n" + spec_type2 + "\n" +
                                                                      spec_cp;
                            }
                        }
                    }
                }
            }
        }
        //public string msg_ToolTipneeeded_onproduct(DataTable dt_spec, DataGridView dgv_data, int r_indx)
        //{
        //    string msg = "";

        //    int reg = 0;
        //    List<DataTable> lst_Table = new List<DataTable> { };
        //    Get_ListTable(-1, (DataTable)dgv_data.DataSource, new string[] { "Region" }, ref lst_Table, "Data");
        //    foreach (string spec_region in dt_spec.Rows[0]["Location"].ToString().Split('_'))
        //    {
        //        if (reg < lst_Table.Count)
        //        {
        //            DataTable Data_all = (DataTable)dgv_data.DataSource;
        //            List<string> lst_region = Data_all.AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToList();

        //            string cpn = "";
        //            foreach (string ar in lst_region)
        //            {
        //                if (spec_region.Replace(" ", "").ToUpper().Replace("COMPONENT", "TAPE").Contains(ar.Split('_')[0].Replace(" ", "").ToUpper()) && spec_region.Replace(" ", "").ToUpper().Replace("COMPONENT", "TAPE").Contains(ar.Split('_')[1].Replace(" ", "").ToUpper()))
        //                {
        //                    cpn = ar;
        //                    break;
        //                }
        //            }

        //            DataView dv = Data_all.AsDataView();
        //            string str_filter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { cpn });
        //            dv.RowFilter = str_filter;
        //            DataTable dt_region = dv.ToTable();


        //            string type = spec_region.Split('+')[0];
        //            if (type == "A")
        //            {
        //                string spec_max = spec_region.Split('+')[3].Split(';')[0].Replace("(gf)", "");
        //                msg += spec_region.Split('+')[3].Split(';')[0] + "\n";
        //                string ll_max = "";
        //                string ul_max = "";
        //                string ll_ave = "";
        //                string ul_ave = "";

        //                if (spec_max.Contains("(") && spec_max.Contains(")"))
        //                {
        //                    spec_max = spec_max.Split('(')[1].Split(')')[0].Replace(" ", ")").Replace("gf", "");
        //                    if (spec_max.Contains("-"))
        //                    {
        //                        ll_max = spec_max.Split('-')[0];
        //                        ul_max = spec_max.Split('-')[1];
        //                    }

        //                }

        //                string spec_average = spec_region.Split('+')[4].Split(';')[0].Replace("(gf)", "");

        //                if (spec_max.Contains("(") && spec_max.Contains(")"))
        //                {
        //                    spec_average = spec_average.Split('(')[1].Split(')')[0].Replace(" ", ")").Replace("gf", "");
        //                    if (spec_average.Contains("-"))
        //                    {
        //                        ll_ave = spec_average.Split('-')[0];
        //                        ul_ave = spec_average.Split('-')[1];
        //                    }
        //                }

        //                if (myCode.IsNumeric(ul_max) && myCode.IsNumeric(ll_max))
        //                {
        //                    for (int i = 0; i < dt_region.Rows.Count; i++)
        //                    {
        //                        string val = dt_region.Rows[i]["Data"].ToString().Split('_')[0].Replace("Max:", "");
        //                        if (myCode.IsNumeric(val))
        //                        {
        //                            int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
        //                            if (Double.Parse(val) < Double.Parse(ll_max))
        //                            {
        //                                msg += val + " < " + ll_max + "\n";
        //                            }
        //                            if (Double.Parse(val) > Double.Parse(ul_max))
        //                            {
        //                                msg += val + "> " + ul_max + "\n";
        //                            }
        //                        }
        //                    }
        //                }
        //                msg += spec_region.Split('+')[4].Split(';')[0] + "\n";
        //                if (myCode.IsNumeric(ul_ave) && myCode.IsNumeric(ll_ave))
        //                {
        //                    for (int i = 0; i < dt_region.Rows.Count; i++)
        //                    {
        //                        string val = dt_region.Rows[i]["Data"].ToString().Split('_')[1].Replace("Average:", "");
        //                        if (myCode.IsNumeric(val))
        //                        {
        //                            int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
        //                            if (Double.Parse(val) < Double.Parse(ll_ave))
        //                            {
        //                                msg += val + " < " + ll_ave + "\n";
        //                            }
        //                            if (Double.Parse(val) > Double.Parse(ul_ave))
        //                            {
        //                                msg += val + " < " + ll_ave + "\n";
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //            else if (type == "B")
        //            {
        //                double UCL = 0;
        //                double LCL = 0;
        //                try
        //                {
        //                    if (spec_region.Split('+')[6].Split(';')[1] != "")
        //                    {
        //                        UCL = double.Parse(spec_region.Split('+')[6].Split(';')[1]);
        //                    }
        //                    if (spec_region.Split('+')[6].Split(';')[2] != "")
        //                    {
        //                        LCL = double.Parse(spec_region.Split('+')[6].Split(';')[2]);
        //                    }
        //                }
        //                catch
        //                {
        //                    if (spec_region.Split('+')[5].Split(';')[1] != "")
        //                    {
        //                        UCL = double.Parse(spec_region.Split('+')[5].Split(';')[1]);
        //                    }
        //                    if (spec_region.Split('+')[5].Split(';')[2] != "")
        //                    {
        //                        LCL = double.Parse(spec_region.Split('+')[5].Split(';')[2]);
        //                    }
        //                }


        //                SortedDictionary<int, string> dic_data = new SortedDictionary<int, string> { };
        //                List<double> lst_data = new List<double> { };
        //                for (int i = 0; i < dt_region.Rows.Count; i++)
        //                {
        //                    string val = dt_region.Rows[i]["Data"].ToString().Split('_')[1].Replace("Average:", "");
        //                    if (myCode.IsNumeric(val))
        //                    {
        //                        int r = int.Parse(dt_region.Rows[i]["ID"].ToString());
        //                        dic_data.Add(r, val);
        //                        lst_data.Add(double.Parse(val));

        //                        if (Double.Parse(val) > UCL && UCL != 0)
        //                        {
        //                            dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;

        //                        }
        //                        if (Double.Parse(val) < LCL && LCL != 0)
        //                        {
        //                            dgv.Rows[r - 1].Cells["Data"].Style.BackColor = Color.Red;

        //                        }
        //                    }
        //                }


        //                string r_spec = "";
        //                try
        //                {
        //                    r_spec = spec_region.Split('+')[6].Split(';')[0];
        //                }
        //                catch
        //                {
        //                    r_spec = spec_region.Split('+')[5].Split(';')[0];
        //                }


        //                if (myCode.IsNumeric(r_spec))
        //                {
        //                    double R = double.Parse(r_spec);
        //                    double tb = lst_data.ToArray().Average();

        //                    foreach (int r1 in dic_data.Keys)
        //                    {
        //                        foreach (int r2 in dic_data.Keys)
        //                        {
        //                            if (r2 < r1)
        //                            {
        //                                double sub_data = Math.Abs(double.Parse(dic_data[r1]) - double.Parse(dic_data[r2]));

        //                                if (sub_data > R && tb != 0)
        //                                {
        //                                    double a1 = Math.Abs(double.Parse(dic_data[r1]) - tb);
        //                                    double a2 = Math.Abs(double.Parse(dic_data[r2]) - tb);
        //                                    if (a1 > a2)
        //                                    {
        //                                        dgv.Rows[r1 - 1].Cells["Data"].Style.BackColor = Color.Red;
        //                                    }
        //                                    else
        //                                    {
        //                                        dgv.Rows[r2 - 1].Cells["Data"].Style.BackColor = Color.Red;
        //                                    }
        //                                }
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //        reg++;
        //    }
        //}

        //    return msg;
        //}


        public string msg_ToolTipneeeded_crosssection(DataTable dt_spec, DataGridView dgv_data, int r_indx)
        {
            string str1 = "";
            string[] strArray = dt_spec.Rows[0]["Location"].ToString().Split('_');
            string str2 = strArray[strArray.Length - 2];
            if (str2.Contains("^"))
            {
                DataTable dataSource = (DataTable)dgv_data.DataSource;
                int val1 = int.Parse(txt_qty.Text);
                List<string> list1 = ((IEnumerable<string>)strArray[strArray.Length - 1].TrimEnd(';').Split(';'))
                    .ToList<string>();
                int index1 = 0;
                Dictionary<string, SortedDictionary<int, string>> dictionary =
                    new Dictionary<string, SortedDictionary<int, string>>();
                string str3 = dgv_data.Rows[r_indx].Cells["Region"].Value.ToString();
                DataView dataView = dataSource.AsDataView();
                string str4 = TDMK_Code.filter_str(new string[1] { "Region" }, new string[1] { str3 });
                dataView.RowFilter = str4;
                DataTable table = dataView.ToTable();
                if (str3.ToUpper().Contains("NGANG") && !str3.ToUpper().Contains("TRU"))
                {
                    for (int index2 = 0; index2 < 2; ++index2)
                    {
                        SortedDictionary<int, string> sortedDictionary1 = new SortedDictionary<int, string>();
                        for (int index3 = index2 * val1;
                             index3 < Math.Min(val1 + index2 * val1, table.Rows.Count);
                             ++index3)
                        {
                            string str5 = table.Rows[index3]["Data"].ToString().Replace(" ", "");
                            if (myCode.IsNumeric(str5.Split(';')[0]))
                                sortedDictionary1.Add(int.Parse(table.Rows[index3]["ID"].ToString()) - 1,
                                    str5.Split(';')[0]);
                        }

                        if (index1 < list1.Count)
                        {
                            dictionary.Add(list1[index1], sortedDictionary1);
                            ++index1;
                        }

                        SortedDictionary<int, string> sortedDictionary2 = new SortedDictionary<int, string>();
                        for (int index4 = index2 * val1;
                             index4 < Math.Min(val1 + index2 * val1, table.Rows.Count);
                             ++index4)
                        {
                            string str6 = table.Rows[index4]["Data"].ToString();
                            if (myCode.IsNumeric(str6.Split(';')[1]))
                                sortedDictionary2.Add(int.Parse(table.Rows[index4]["ID"].ToString()) - 1,
                                    str6.Split(';')[1]);
                        }

                        if (index1 < list1.Count)
                        {
                            dictionary.Add(list1[index1], sortedDictionary2);
                            ++index1;
                        }
                    }
                }
                else if (!str3.ToUpper().Contains("NGANG"))
                {
                    for (int index5 = 0; index5 < 6; ++index5)
                    {
                        SortedDictionary<int, string> sortedDictionary = new SortedDictionary<int, string>();
                        for (int index6 = 0; index6 < Math.Min(val1, table.Rows.Count); ++index6)
                        {
                            string str7 = table.Rows[index6]["Data"].ToString().Replace(" ", "").Replace("/", "");
                            string str8 = str7.Split(';')[index5];
                            switch (index5)
                            {
                                case 3:
                                    str8 = str7.Split(';')[4];
                                    break;
                                case 4:
                                    str8 = str7.Split(';')[3];
                                    break;
                            }

                            if (myCode.IsNumeric(str8))
                                sortedDictionary.Add(int.Parse(table.Rows[index6]["ID"].ToString()) - 1, str8);
                        }

                        if (index1 < list1.Count)
                        {
                            dictionary.Add(list1[index1], sortedDictionary);
                            ++index1;
                        }
                    }
                }

                string str9 = str2.Split('#')[1];
                List<string> stringList1 = new List<string>();
                List<string> stringList2 = new List<string>();
                List<string> stringList3 = new List<string>();
                string str10 = str9;
                char[] chArray = new char[1] { '^' };
                foreach (string str11 in str10.Split(chArray))
                {
                    if (str11.Contains("R"))
                        stringList1 = ((IEnumerable<string>)str11.Remove(str11.LastIndexOf(";")).Replace("R:", "")
                            .Replace(" ", "").Split(';')).ToList<string>();
                    if (str11.Contains("UCL"))
                        stringList2 = ((IEnumerable<string>)str11.Remove(str11.LastIndexOf(";")).Replace("UCL:", "")
                            .Replace(" ", "").Split(';')).ToList<string>();
                    if (str11.Contains("LCL"))
                        stringList3 = ((IEnumerable<string>)str11.Remove(str11.LastIndexOf(";")).Replace("LCL:", "")
                            .Replace(" ", "").Split(';')).ToList<string>();
                }

                List<string> list2 = ((IEnumerable<string>)str2.Split('#')[0].TrimEnd('@').Split('@')).ToList<string>();
                int index7 = 0;
                foreach (string str12 in list2)
                {
                    List<string> list3 = ((IEnumerable<string>)str12.Replace(" ", "").Split(',')).ToList<string>();
                    foreach (KeyValuePair<string, SortedDictionary<int, string>> keyValuePair1 in dictionary)
                    {
                        if (list3.Contains(keyValuePair1.Key))
                        {
                            string _r = stringList1[index7];
                            string _ucl = stringList2[index7];
                            string _lcl = stringList3[index7];
                            foreach (KeyValuePair<int, string> keyValuePair2 in keyValuePair1.Value)
                            {
                                if (keyValuePair2.Key == r_indx)
                                {
                                    string str13 = message_check_cross_mass(keyValuePair2.Value, _ucl, _lcl);
                                    if (str13 != "")
                                    {
                                        str1 = str1 + str13 + "\n\n";
                                        break;
                                    }

                                    break;
                                }
                            }

                            string str_i = message_check_R_crosscut(_r, keyValuePair1.Value, r_indx);
                            if (str_i != "")
                                str1 += str_i;
                        }
                    }

                    ++index7;
                }
            }

            if (str1 == "")
            {
                string str14 = dt_spec.Rows[0]["Location"].ToString().Split('+')[0].Split(';')[2];
                if (str14.Contains("<") && dgv_data.CurrentCell != null)
                    str1 = str14 + "\n1.5 < B < 4.5";
            }

            return str1;
        }

        public string message_check_cross_mass(string data, string _ucl, string _lcl)
        {
            string str = "";
            if (this.myCode.IsNumeric(data))
            {
                double num1 = 0.0;
                double num2 = 0.0;
                if (this.myCode.IsNumeric(_ucl))
                    num1 = double.Parse(_ucl);
                if (this.myCode.IsNumeric(_lcl))
                    num2 = double.Parse(_lcl);
                if (double.Parse(data) > num1 && num1 != 0.0)
                    str = data + " > UCL(" + num1.ToString() + ")";
                if (double.Parse(data) < num2 && num2 != 0.0)
                    str = data + " < LCL(" + num2.ToString() + ")";
            }

            return str;
        }

        public string message_check_R_crosscut(string _r, SortedDictionary<int, string> dic_data, int r1)
        {
            string str = "";
            if (myCode.IsNumeric(_r) && dic_data.ContainsKey(r1))
            {
                List<double> doubleList = new List<double>();
                foreach (string s in dic_data.Values)
                    doubleList.Add(double.Parse(s));
                double num1 = double.Parse(_r);
                double num2 = ((IEnumerable<double>)doubleList.ToArray()).Average();
                foreach (int key in dic_data.Keys)
                {
                    if (r1 != key && Math.Abs(double.Parse(dic_data[r1]) - double.Parse(dic_data[key])) > num1 &&
                        num2 != 0.0 && Math.Abs(double.Parse(dic_data[r1]) - num2) >
                        Math.Abs(double.Parse(dic_data[key]) - num2))
                    {
                        if (!str.Contains("R ="))
                            str += "R = " + num1.ToString() + "\n";
                        str += "|" + dic_data[r1] + "-" + dic_data[key] + "(ID:" + key.ToString() + ") | > R" + "\n";
                    }
                }
            }

            return str;
        }

        public string msg_ToolTipneeeded_GAP(DataTable dt_spec, DataGridView dgv, int r_indx)
        {
            string str1 = "";
            string val = this.myCode.checkDBNull(dgv.Rows[r_indx].Cells["Data"].Value.ToString());
            if (myCode.IsNumeric(val.Split('/')[0].TrimEnd(' ', ';').Split(';').Last()) &&
                myCode.IsNumeric(val.Split('/')[1].TrimEnd(' ', ';').Split(';').Last()))
            {
                double data1 = double.Parse(val.Split('/')[0].TrimEnd(' ', ';').Split(';').Last());
                double data2 = double.Parse(val.Split('/')[1].TrimEnd(' ', ';').Split(';').Last());

                double sub_data = Math.Abs(data1 - data2);
                if (sub_data > 30)
                {
                    str1 = "|" + data1.ToString() + "-" + data2.ToString() + "| > 30\n";
                }
            }


            string[] strArray = dt_spec.Rows[0]["Location"].ToString().Split('+')[1].Split('_');
            string str3 = dgv.Rows[r_indx].Cells["Region"].Value.ToString();
            foreach (string str4 in strArray)
            {
                if (str4 != "")
                {
                    int num6 = 1;
                    string str5 = str4;
                    char[] chArray = new char[1] { '^' };
                    foreach (string str6 in str5.Split(chArray))
                    {
                        if (str6 != "")
                        {
                            int.Parse(str6.Split(';')[1]);
                            int.Parse(str6.Split(';')[2]);
                            string str7 = str6.Split(';')[3].Replace("Left", "").Replace("Right", "").Replace("B1", "")
                                .Replace("B2 ", "");
                            if (str7.Contains("<") && str7.Contains("µm"))
                            {
                                if (this.myCode.IsNumeric(str7.Split('<')[1].Replace(" ", "").Replace("µm", "")
                                        .Replace(")", "")))
                                {
                                    double num7 = double.Parse(str7.Split('<')[1].Replace(" ", "").Replace("µm", "")
                                        .Replace(")", ""));
                                    string str8 = "";
                                    if (str4.Split('^').Length == 3)
                                        str8 = num6.ToString();
                                    // string str9 = "";
                                    string str10 = "";
                                    if (str6.Split(';')[0] == "IOPIN")
                                    {
                                        if (str3.ToUpper().Contains("DOC") && str3.Contains(str8))
                                        {
                                            str10 = str3;
                                            // break;
                                        }
                                    }
                                    else
                                    {
                                        int num8;
                                        if (!(str6.Split(';')[0] == "RIGHT"))
                                            num8 = str6.Split(';')[0] == "LEFT" ? 1 : 0;
                                        else
                                            num8 = 1;
                                        if (num8 != 0 && str3.ToUpper().Contains("TRU") && str3.Contains(str8))
                                        {
                                            str10 = str3;
                                            //  break;
                                        }
                                    }

                                    if (str10 != "")
                                    {
                                        this.myCode.checkDBNull((object)dgv.Rows[r_indx].Cells["Data"].Value
                                            .ToString());
                                        double num9 = double.Parse(val.Split('/')[0].Split(';')[1]);
                                        double num10 = double.Parse(val.Split('/')[1].Split(';')[1]);

                                        str1 += str7 + "\n";
                                        if (num9 > num7)
                                            str1 += num9.ToString() + " > " + num7.ToString() + "\n";
                                        if (num10 > num7)
                                            str1 += num10.ToString() + " > " + num7.ToString() + "\n";
                                        return str1;
                                    }
                                }
                            }
                        }

                        ++num6;
                    }
                }
            }

            return str1;
        }

        public string msg_ToolTipneeeded_peelpull(DataTable dt_spec, DataGridView dgv, int r_indx)
        {
            string str1 = "";
            if (this.myCode.IsNumeric(dgv.Rows[r_indx].Cells["Region"].Value.ToString()))
            {
                int index1 = int.Parse(dgv.Rows[r_indx].Cells["Region"].Value.ToString()) - 1;
                List<DataTable> src_lst_tbl = new List<DataTable>();
                this.Get_ListTable(-1, (DataTable)dgv.DataSource, new string[1] { "Region" }, ref src_lst_tbl, "Data");
                string str2 = dt_spec.Rows[0]["Location"].ToString().Split('_')[index1];
                DataTable dataTable = src_lst_tbl[index1];
                switch (str2.Split('+')[0])
                {
                    case "A":
                        string s1 = str2.Split('+')[3].Split(';')[0].Split('(')[1].Split(')')[0]
                            .Replace(" ", string.Empty).Replace("N", string.Empty).Replace("≥", string.Empty)
                            .Replace(">", string.Empty).Replace("<", string.Empty);
                        if (this.TDMK_Code.IsNumeric(s1))
                        {
                            string s2 = dataTable.Rows[r_indx]["Data"].ToString();
                            if (this.myCode.IsNumeric(s2) && double.Parse(s2) < double.Parse(s1))
                                str1 = str2.Split('+')[3].Split(';')[0] + "   " + s2 + " < " + s1 + "\n\n";
                        }

                        string str3 = this.msg_checkAlldata_peelpullshear(dgv, r_indx);
                        if (str3 != "")
                        {
                            str1 = str1 + str3 + "\n";
                            break;
                        }

                        break;
                    case "B":
                        string s3 = str2.Split('+')[4];
                        List<double> doubleList = new List<double>();
                        if (this.TDMK_Code.IsNumeric(s3))
                        {
                            double num1 = double.Parse(str2.Split('+')[5].Split(';')[1]);
                            double num2 = double.Parse(str2.Split('+')[5].Split(';')[2]);
                            SortedDictionary<int, string> sortedDictionary = new SortedDictionary<int, string>();
                            for (int index2 = 0; index2 < dataTable.Rows.Count; ++index2)
                            {
                                string s4 = dataTable.Rows[index2]["Data"].ToString().Replace(" ", "");
                                if (this.myCode.IsNumeric(s4))
                                {
                                    int key = int.Parse(dataTable.Rows[index2]["ID"].ToString());
                                    sortedDictionary.Add(key, s4);
                                    doubleList.Add(double.Parse(s4));
                                    if (key - 1 == r_indx)
                                    {
                                        if (double.Parse(s4) < double.Parse(s3))
                                            str1 = "Judgement: ≥" + s3 + "\n" + s4 + " < " + s3 + "\n\n";
                                        else if (double.Parse(s4) > num1)
                                            str1 = str1 + s4 + " > UCL: " + num1.ToString() + "\n\n";
                                        else if (double.Parse(s4) < num2)
                                            str1 = str1 + s4 + " < LCL: " + num2.ToString() + "\n\n";
                                    }
                                }
                            }

                            double num3 = double.Parse(str2.Split('+')[5].Split(';')[0]);
                            double num4 = ((IEnumerable<double>)doubleList.ToArray()).Average();
                            int key1 = r_indx + 1;
                            foreach (int key2 in sortedDictionary.Keys)
                            {
                                if (key2 != key1 &&
                                    Math.Abs(
                                        double.Parse(sortedDictionary[key1]) - double.Parse(sortedDictionary[key2])) >
                                    num3 && num4 != 0.0 && Math.Abs(double.Parse(sortedDictionary[key1]) - num4) >
                                    Math.Abs(double.Parse(sortedDictionary[key2]) - num4))
                                {
                                    if (!str1.Contains("R ="))
                                        str1 = str1 + "R = " + num3.ToString() + "\n";
                                    str1 = str1 + "| " + sortedDictionary[key1] + "-" + sortedDictionary[key2] +
                                           " | > R\n";
                                }
                            }
                        }

                        break;
                }
            }

            return str1;
        }

        public string msg_checkAlldata_peelpullshear(DataGridView dgv, int r_indx)
        {
            string str = "";
            if (dgv.DataSource != null)
            {
                List<DataTable> src_lst_tbl = new List<DataTable>();
                this.Get_ListTable(-1, (DataTable)dgv.DataSource, new string[1] { "Region" }, ref src_lst_tbl, "Data");
                if (this.myCode.IsNumeric(dgv.Rows[r_indx].Cells["Region"].Value.ToString()))
                {
                    int index1 = int.Parse(dgv.Rows[r_indx].Cells["Region"].Value.ToString()) - 1;
                    DataTable dataTable = src_lst_tbl[index1];
                    List<double> doubleList = new List<double>();
                    SortedDictionary<int, string> sortedDictionary = new SortedDictionary<int, string>();
                    double num1 = 0.0;
                    for (int index2 = 0; index2 < dataTable.Rows.Count; ++index2)
                    {
                        string s = this.myCode.checkDBNull(dataTable.Rows[index2]["Data"]);
                        int key = int.Parse(dataTable.Rows[index2]["ID"].ToString()) - 1;
                        if (s != "")
                        {
                            sortedDictionary.Add(key, s);
                            if (this.myCode.IsNumeric(s))
                                doubleList.Add(double.Parse(s));
                        }
                    }

                    if (doubleList.Count > 0)
                        num1 = ((IEnumerable<double>)doubleList.ToArray()).Average();
                    double num2 = 0.0;
                    switch (this.sheet)
                    {
                        case "PEEL_TEST":
                            num2 = 10.0;
                            break;
                        case "MATING_PULL_TEST":
                            num2 = 15.0;
                            break;
                        case "SHEAR_TEST":
                            num2 = 5.0;
                            break;
                    }

                    if (num2 != 0.0)
                    {
                        int key1 = r_indx;
                        str = "R = " + num2.ToString() + "\n";
                        foreach (int key2 in sortedDictionary.Keys)
                        {
                            if (key1 != key2)
                            {
                                if (Math.Abs(
                                        double.Parse(sortedDictionary[key1]) - double.Parse(sortedDictionary[key2])) >
                                    num2 && num1 != 0.0 && Math.Abs(double.Parse(sortedDictionary[key1]) - num1) >
                                    Math.Abs(double.Parse(sortedDictionary[key2]) - num1))
                                    str = str + "| " + sortedDictionary[key1] + " - " + sortedDictionary[key2] +
                                          "(ID:" + key2.ToString() + ") | > R\n";
                            }
                        }
                    }
                }
            }

            return str;
        }

        public string str_round(string s)
        {
            string str_ref = "";
            if (myCode.IsNumeric(s))
            {
                str_ref = Math.Round(double.Parse(s), 2).ToString();
            }

            return str_ref;
        }

        private void dgv_Analysis_CellClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void cmsPaste_Opening(object sender, CancelEventArgs e)
        {
        }

        public void edit_image(int r_inx, int c_indx, DataGridView dgv)
        {
            string f_path = "";
            OpenFileDialog f_open = new OpenFileDialog();
            f_open.Filter = "Excel Files|*.jpg;*.png;*.jpeg;";
            f_open.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            if (f_open.ShowDialog() == DialogResult.OK)
            {
                if (f_open.FileName != "")
                {
                    f_path = f_open.FileName;
                    dgv.Rows[r_inx].Cells[c_indx].Value = File_Image_To_DGV_Value(f_path);
                }
            }
        }

        public void edit_graph(int r_inx, DataGridView dgv)
        {
            string f_path = "";
            OpenFileDialog f_open = new OpenFileDialog();
            f_open.Filter = "Excel Files|*.xlsx;";
            f_open.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            if (f_open.ShowDialog() == DialogResult.OK)
            {
                if (f_open.FileName != "")
                {
                    f_path = f_open.FileName;

                    ExcelWorkbook wrkbk = TDMK_EPPLUS.open_excel_file(f_path);
                    ExcelWorksheet wrksht = wrkbk.Worksheets[0];
                    switch (sheet)
                    {
                        case "PEEL_TEST":
                            dgv.Rows[r_inx].Cells["Data"].Value = get_data_val_comment3(wrksht, "Max");
                            break;

                        case "MATING_PULL_TEST":
                            dgv.Rows[r_inx].Cells["Data"].Value = get_data_val_comment3(wrksht, "Max");
                            break;

                        case "SHEAR_TEST":
                            break;

                        case "IQC_UNMATING_PULL_TEST":
                            dgv.Rows[r_inx].Cells["Data"].Value = get_data_val_comment3(wrksht, "Max");
                            break;

                        case "IQC_LINER_PEELING_COUPON":
                            dgv.Rows[r_inx].Cells["Data"].Value = get_data_val_comment3(wrksht, "Average");
                            break;

                        case "IQC_PSA_PEELING_COUPON":
                            dgv.Rows[r_inx].Cells["Data"].Value = get_data_val_comment3(wrksht, "Average");
                            break;

                        case "LINER_PEEL_TEST_ON_PRODUCT":
                            dgv.Rows[r_inx].Cells["Data"].Value = get_data_val_comment3_onproduct(wrksht);
                            break;

                        case "PSA_PEEL_TEST_ON_PRODUCT":
                            dgv.Rows[r_inx].Cells["Data"].Value = get_data_val_comment3_onproduct(wrksht);
                            break;

                        case "CROSS_SECTION":
                            break;

                        case "GAP_CONNECTOR":
                            break;
                    }

                    dgv.Rows[r_inx].Cells["Graph"].Value = get_image_excel(wrksht);
                }
            }
        }


        private void dgv_Analysis_MouseClick(object sender, MouseEventArgs e)
        {
        }

        private void dgv_Analysis_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (edit_mode)
            {
                if (admin_mode == "Admin mode")
                {
                    if (e.Button == MouseButtons.Right)
                    {
                        int r_indx = e.RowIndex;
                        int c_indx = e.ColumnIndex;

                        if (dgv_Analysis.Columns[c_indx].Name.Contains("Image"))
                        {
                            edit_image(r_indx, c_indx, dgv_Analysis);
                        }
                        else if (dgv_Analysis.Columns[c_indx].Name == "Graph")
                        {
                            edit_graph(r_indx, dgv_Analysis);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Vui lòng đăng nhập để chỉnh sửa", "Thông báo");
                }
            }
        }

        private void cb_Type_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_Type.SelectedIndex != -1)
            {
                DataTable dt_spec = new DataTable();
                //DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, sheet, cb_Type.SelectedItem.ToString() }));
                if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                {
                    dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                        TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                            new string[] { txtItemCode.Text, sheet, "MASS" }));
                }
                else if (cb_Type.SelectedItem.ToString() == "NPI" || cb_Type.SelectedItem.ToString() == "LQ")
                {
                    dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                        TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                            new string[] { txtItemCode.Text, sheet, "NPI" }));
                }

                txt_qty.Text = "";
                if (dt_spec.Rows.Count > 0)
                {
                    txt_qty.Text = dt_spec.Rows[0]["Count_sample"].ToString();
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Chưa cài đặt format cho Type này", "Thông báo");
                }

                if (cb_Type.SelectedItem.ToString() != "NPI")
                {
                    txt_qty.Enabled = true;
                    txt_ItemName.Enabled = true;
                    txt_worker.Enabled = true;
                    txt_date.Enabled = true;
                    if (!sheet.Contains("UNMATING") && !sheet.Contains("COUPON"))
                    {
                        txt_line.Enabled = true;
                        txt_ca.Enabled = true;
                    }

                    lbl_refer.Visible = false;
                    lbl_.Visible = false;
                    txt_pcs_begin.Visible = false;
                    txt_pcs_end.Visible = false;
                    if (sheet.Contains("CROSS") || sheet.Contains("GAP") || sheet == "PEEL_TEST" ||
                        sheet == "MATING_PULL_TEST" || sheet.Contains("ON_PRODUCT"))
                    {
                        lbl_itemcode_nvl.Enabled = false;
                        lbl_lotno_nvl.Enabled = false;
                        txt_itemcode_nvl.Enabled = false;
                        txt_lotno_nvl.Enabled = false;
                        txt_itemcode_nvl.Text = "";
                        txt_lotno_nvl.Text = "";
                    }
                }
                else
                {
                    txt_ItemName.Text = txt_worker.Text = txt_line.Text = txt_ca.Text = txt_date.Text = "";
                    txt_qty.Enabled = false;
                    txt_ItemName.Enabled = false;
                    txt_worker.Enabled = false;
                    txt_line.Enabled = false;
                    txt_ca.Enabled = false;
                    txt_date.Enabled = false;
                    if (sheet.Contains("UNMATING") || sheet.Contains("COUPON"))
                        txt_date.Enabled = true;
                    if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST" || sheet.Contains("ON_PRODUCT"))
                    {
                        lbl_refer.Visible = true;
                        lbl_.Visible = true;
                        txt_pcs_begin.Visible = true;
                        txt_pcs_end.Visible = true;
                        lbl_itemcode_nvl.Enabled = true;
                        lbl_lotno_nvl.Enabled = true;
                        txt_itemcode_nvl.Enabled = true;
                        txt_lotno_nvl.Enabled = true;
                    }

                    if (sheet.Contains("CROSS") || sheet.Contains("GAP"))
                    {
                        lbl_itemcode_nvl.Enabled = true;
                        lbl_lotno_nvl.Enabled = true;
                        lbl_itemcode_nvl.Text = "ItemCode(Refer)";
                        lbl_lotno_nvl.Text = "LotNo(Refer)";
                        txt_itemcode_nvl.Enabled = true;
                        txt_lotno_nvl.Enabled = true;
                    }
                }
            }
        }

        private void txt_qty_TextChanged(object sender, EventArgs e)
        {
            if (!myCode.IsNumeric(txt_qty.Text))
            {
                txt_qty.BackColor = Color.Yellow;
            }
            else
            {
                txt_qty.BackColor = Color.White;
            }
        }

        private void dgv_logfile_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void dgv_Analysis_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            if (e.RowIndex != -1 && e.ColumnIndex != -1)
            {
                string region = dgv_Analysis.Rows[e.RowIndex].Cells["Region"].Value.ToString();
                string filter = TDMK_Code.filter_str(new string[] { "Region" }, new string[] { region });
                DataView dv = ((DataTable)dgv_Analysis.DataSource).AsDataView();
                dv.RowFilter = filter;
                DataTable dt_filter = dv.ToTable();
                List<double> lst_data = new List<double> { };
                DataTable dt_spec = new DataTable();
                // DataTable dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3", TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" }, new string[] { txtItemCode.Text, sheet, cb_Type.SelectedItem.ToString() }));
                if (cb_Type.SelectedItem.ToString().Contains("MASS"))
                {
                    dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                        TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                            new string[] { txtItemCode.Text, sheet, "MASS" }));
                }
                else if (cb_Type.SelectedItem.ToString() == "LQ" || cb_Type.SelectedItem.ToString() == "NPI")
                {
                    dt_spec = TDMK_Code.Datatable_Filter(sqlcon, "SPEC_COMMENT_3",
                        TDMK_Code.filter_str(new string[] { "ItemCode", "Sheet", "Remark" },
                            new string[] { txtItemCode.Text, sheet, "NPI" }));
                }

                if (dt_spec.Rows.Count > 0)
                {
                    if (sheet == "CROSS_SECTION")
                    {
                        if (dgv_Analysis.Rows[e.RowIndex].Cells["Data"].Style.BackColor == Color.Red)
                            dgv_Analysis.CurrentCell.ToolTipText =
                                msg_ToolTipneeeded_crosssection(dt_spec, dgv_Analysis, e.RowIndex);
                    }
                    else if (sheet == "GAP_CONNECTOR")
                    {
                        if (dgv_Analysis.Rows[e.RowIndex].Cells["Data"].Style.BackColor == Color.Red)
                            dgv_Analysis.CurrentCell.ToolTipText =
                                msg_ToolTipneeeded_GAP(dt_spec, dgv_Analysis, e.RowIndex);
                    }
                    else if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST")
                    {
                        if (dgv_Analysis.Rows[e.RowIndex].Cells["Data"].Style.BackColor == Color.Red)
                            dgv_Analysis.CurrentCell.ToolTipText =
                                msg_ToolTipneeeded_peelpull(dt_spec, dgv_Analysis, e.RowIndex);
                    }
                    else if (!sheet.Contains("ON_PRODUCT"))
                    {
                        foreach (DataRow dr in dt_filter.Rows)
                        {
                            if (myCode.IsNumeric(dr["Data"].ToString()))
                            {
                                lst_data.Add(Double.Parse(dr["Data"].ToString()));
                            }
                        }
                    }

                    if (lst_data.Count > 0)
                    {
                        string max = lst_data.ToArray().Max().ToString();
                        string min = lst_data.ToArray().Min().ToString();
                        string Average = Math.Round(lst_data.ToArray().Average(), 2).ToString();
                        string spec_type2 = "";
                        if (sheet == "IQC_UNMATING_PULL_TEST")
                        {
                            spec_type2 += dt_spec.Rows[0]["Location"].ToString().Split('+')[3].Split(';')[0];
                        }

                        if (dgv_Analysis.Columns[e.ColumnIndex].Name == "Data")
                        {
                            dgv_Analysis.CurrentCell.ToolTipText = "Min: " + min + "\n" + "Max: " + max + "\n" +
                                                                   "Average: " + Average + "\n" + spec_type2;
                        }
                    }
                }
            }
        }

        private void lbl_select_Click(object sender, EventArgs e)
        {
        }

        private void txtLogfile_TextChanged(object sender, EventArgs e)
        {
            txtLogfile.Text = txtLogfile.Text.Replace("\r", "").Replace("\n", "");

            if (!sheet.Contains("CROSS") && !sheet.Contains("GAP") || !(txtLogfile.Text != ""))
                return;
            txt_itemcode_nvl.Text = "";
            txt_lotno_nvl.Text = "";
        }

        private void txt_date_TextChanged(object sender, EventArgs e)
        {
        }

        private void cb_Type_Click(object sender, EventArgs e)
        {
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {
        }

        private void dgv_logfile_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                int r_indx = e.RowIndex;
                int c_indx = e.ColumnIndex;

                if (dgv_logfile.Columns[c_indx].Name.Contains("Image"))
                {
                    edit_image(r_indx, c_indx, dgv_logfile);
                }
                else if (dgv_logfile.Columns[c_indx].Name == "Graph")
                {
                    edit_graph(r_indx, dgv_logfile);
                }
                else if (dgv_logfile.Columns[c_indx].Name == "Mode")
                {
                    cmsPaste.Show(dgv_logfile, e.Location);
                }
            }
        }

        private void lbl_itemcode_nvl_Click(object sender, EventArgs e)
        {
        }

        private void btn_B2B_Click(object sender, EventArgs e)
        {
            if (txt_itemcode_nvl.Text != "" && txt_lotno_nvl.Text != "")
            {
                DataTable dt_ok2ship = (DataTable)dgv_logfile.DataSource;
                if (dgv_logfile.DataSource != null)
                {
                    DataTable dt_ok2build_saved = TDMK_Code.Datatable_Filter(sqlcon, sheet,
                        TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" },
                            new string[] { txt_itemcode_nvl.Text, txt_lotno_nvl.Text }));
                    List<string> lst_lot = dt_ok2build_saved.AsEnumerable().Select(x => x.Field<string>("Region"))
                        .Distinct().ToList();
                    if (lst_lot.Count == 1)
                    {
                        DataTable dt_ok2build = TDMK_Code.Datatable_Filter(sqlcon, sheet,
                            TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" },
                                new string[] { txt_itemcode_nvl.Text, txt_lotno_nvl.Text })).Clone();
                        dt_ok2build.Columns.Add("Select_Img", typeof(bool));
                        dt_ok2build.Columns.Add("Select_Grp", typeof(bool));
                        foreach (DataRow dr in dt_ok2build_saved.Rows)
                        {
                            DataRow dr_new = dt_ok2build.NewRow();
                            for (int i = 0; i < dt_ok2build_saved.Columns.Count; i++)
                            {
                                dr_new[i] = dr[i];
                            }

                            dr_new["Select_Img"] = true;
                            dr_new["Select_Grp"] = true;
                            dt_ok2build.Rows.Add(dr_new);
                        }

                        dt_ok2build = dt_ok2build.AsDataView().ToTable(false,
                            new string[]
                            {
                                "ID", "ItemCode", "LotNo", "Sheet", "Region", "Sample", "Image", "Select_Img", "Graph",
                                "Select_Grp", "Data", "Mode 1: Solder joint crack", "Mode 2: Pad lift",
                                "Mode 3: Solder joint lift", "Mode 4: Intermetallic break", "Mode 5: Component damage",
                                "Mode 6: Component detached", "Mode 7: Flex torn", "Operator", "Time_Update", "Remark"
                            });

                        for (int i = 10; i < 32; i++)
                        {
                            if (i < dt_ok2build.Rows.Count)
                            {
                                DataRow dr = dt_ok2ship.NewRow();
                                dr[0] = i + 1;
                                dr[1] = txtItemCode.Text;
                                dr[2] = txtLotNo.Text;
                                for (int c = 3; c < dt_ok2ship.Columns.Count; c++)
                                {
                                    dr[c] = dt_ok2build.Rows[i][c];
                                }

                                dt_ok2ship.Rows.Add(dr);
                            }
                        }

                        int st = 1;
                        foreach (DataRow dr in dt_ok2ship.Rows)
                        {
                            dr["ID"] = st;
                            st++;
                        }
                    }
                    else
                    {
                    }

                    dgv_logfile.DataSource = dt_ok2ship;
                    resize_column_image(dgv_logfile);
                    txt_selected.Text = count_selected();
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Chưa nhập ItemCode/Lot OK2Build", "Thông báo!");
            }
        }

        private void txt_lotno_nvl_TextChanged(object sender, EventArgs e)
        {
            if (!sheet.Contains("CROSS") && !sheet.Contains("GAP") || !(txt_itemcode_nvl.Text != "") ||
                !(txt_lotno_nvl.Text != ""))
                return;
            txtLogfile.Text = "";
        }

        private void txt_lotno_nvl_Validated(object sender, EventArgs e)
        {
            //if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST")
            //{
            //    txt_lotno_nvl.Text = Lotno_Formated(txt_lotno_nvl.Text);
            //}
            if (sheet.Contains("UNMATING") || sheet.Contains("COUPON"))
                return;
            txt_lotno_nvl.Text = Lotno_Formated(txt_lotno_nvl.Text);
        }

        private void txt_pcs_begin_Validated(object sender, EventArgs e)
        {
        }

        private void txt_pcs_after_Validated(object sender, EventArgs e)
        {
        }

        private void txt_pcs_begin_TextChanged(object sender, EventArgs e)
        {
            txt_pcs_begin.Text = txt_pcs_begin.Text.Replace(" ", "");

            if (!myCode.IsNumeric(txt_pcs_begin.Text) && txt_pcs_begin.Text != "")
            {
                txt_pcs_begin.BackColor = Color.Yellow;
            }
            else
            {
                txt_pcs_begin.BackColor = Color.White;
            }
        }

        private void txt_pcs_after_TextChanged(object sender, EventArgs e)
        {
            txt_pcs_end.Text = txt_pcs_end.Text.Replace(" ", "");

            if (!myCode.IsNumeric(txt_pcs_end.Text) && txt_pcs_end.Text != "")
            {
                txt_pcs_end.BackColor = Color.Yellow;
            }
            else
            {
                txt_pcs_end.BackColor = Color.White;
            }
        }

        private void txt_itemcode_nvl_TextChanged(object sender, EventArgs e)
        {
            if (!sheet.Contains("CROSS") && !sheet.Contains("GAP") || !(txt_itemcode_nvl.Text != "") ||
                !(txt_lotno_nvl.Text != ""))
                return;
            txtLogfile.Text = "";
        }

        private void tableLayoutPanel6_Paint(object sender, PaintEventArgs e)
        {
        }

        private void txt_itemcode_nvl_TextChanged_1(object sender, EventArgs e)
        {
        }

        private void txt_selected_TextChanged(object sender, EventArgs e)
        {
        }

        private void txt_setchan_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (txt_setchan.Text == "0" || !myCode.IsNumeric(txt_setchan.Text))
                {
                    txt_setchan.BackColor = Color.Yellow;
                }
                else
                {
                    txt_setchan.BackColor = Color.White;
                    set_chan = int.Parse(txt_setchan.Text);
                    DataTable dataTable = (DataTable)dgv_logfile.DataSource;
                    foreach (DataRow row in dataTable.Rows)
                    {
                        foreach (DataColumn col in dataTable.Columns)
                        {
                            if (col.ColumnName.Contains("Mode"))
                            {
                                if (row[col.ColumnName].ToString().Contains("100%"))
                                {
                                    row[col.ColumnName] = $"100% ({set_chan}/{set_chan})";
                                }
                                else if (row[col.ColumnName].ToString().Contains("0%"))
                                {
                                    row[col.ColumnName] = $"0% ({0}/{set_chan})";
                                }
                                else
                                {
                                    if (int.TryParse(row[col.ColumnName].ToString().TrimEnd('%'), out int a))
                                    {
                                        a = 100 / (a * set_chan);
                                        row[col.ColumnName] = $"{row[col.ColumnName]} ({a}/{set_chan})";
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
            }
        }

        private void dgv_logfile_DataSourceChanged(object sender, EventArgs e)
        {
            //DataTable dataTable = (DataTable)dgv_logfile.DataSource;
            //DataColumn sttColumn = new DataColumn("ID", typeof(int));
            //dataTable.Columns.Add(sttColumn);

            //// 2. Duyệt qua các dòng và gán giá trị số thứ tự
            //for (int i = 0; i < dgv_logfile.Rows.Count; i++)
            //{
            //    dataTable.Rows[i]["ID"] = i + 1;
            //}
            //// Di chuyển cột 'STT' lên index 1
            //if (dataTable.Columns.Contains("ID"))
            //{
            //    dataTable.Columns["ID"].SetOrdinal(1);
            //}
        }

        private void tableLayoutPanel8_Paint(object sender, PaintEventArgs e)
        {
        }

        private void textBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "All files (*.*)|*.*";
            fileDialog.FilterIndex = 1;
            fileDialog.RestoreDirectory = true;
            fileDialog.Title = "Chọn tệp";
            fileDialog.ShowDialog();
            textBox1.Text = fileDialog.FileName;
        }

        private void dgv_Analysis_DataSourceChanged(object sender, EventArgs e)
        {
            switch (sheet)
            {
                case "PEEL_TEST":
                case "MATING_PULL_TEST":
                case "SHEAR_TEST":
                    try
                    {
                        DataTable dt = (DataTable)((DataGridView)sender).DataSource;
                        foreach (DataRow row in dt.Rows)
                        {
                            int tong = 0;

                            foreach (DataColumn column in dt.Columns)
                            {
                                if (column.ColumnName.Contains("Mode"))
                                {
                                    string value = row[column].ToString().Split('%')[1].Split('(')[1].Replace(")", "")
                                        .Replace(" ", "").Replace(")", "");
                                    tong += int.Parse(value.Split('/')[0]);
                                    txt_setchan.Text = value.Split('/')[1];
                                }
                            }

                            if (tong != int.Parse(txt_setchan.Text))
                            {
                                lbl_judge.Text = "NG";
                                lbl_judge.BackColor = Color.Red;
                                return;
                            }
                            else
                            {
                                lbl_judge.Text = "OK";
                                lbl_judge.BackColor = Color.Green;
                            }
                        }
                    }
                    catch
                    {
                        return;
                    }

                    break;
                default:
                    Debugger.Break();
                    break;
            }
        }
    }
}