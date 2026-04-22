using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OK2SHIP_Measurements;
using TDMK_SQL;
using OK2SHIP_Lib;

namespace OK2SHIP_Measurements
{
    public partial class Load_Test_Data : Form
    {
        //public SEI_Lib myCode = new SEI_Lib();
        public TDMK_OK2SHIP myCode = new TDMK_OK2SHIP();
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        
        public static SqlConnection sqlcon_FAI;
        public static SqlConnection sqlcon_Declare;
        public static SqlConnection sqlcon_Materials;
        public static SqlConnection sqlcon_Recycle;
        public static SqlConnection sqlcon_IPQC;
        public static SqlConnection sqlcon_OK2SHIP;
        public static SqlConnection sel_sqlcon;
        public string server_name;
        public string server_acc;
        public string server_pass;
        public string DB_name;
        public string app_path;
        public string Data_Location;
        public string format_folder;
        public string log_folder;
        public string f_ext;
        public Load_Test_Data()
        {
            InitializeComponent();
        }

        private void Load_Test_Data_Load(object sender, EventArgs e)
        {
            initial_data();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //string fil_loc = @"D:\Customer Projects\SEEV\TestAreas\TestData";
            //string ItemCode = "22B0473";
            //string LotNo = "00004";
            //myCode.Load_Test_Data(fil_loc, ItemCode, LotNo, ".xlsm", DGV_DataView);
            DataTable mytbl = new DataTable();
            DataTable tbl_spec = new DataTable();
            DataTable FAI_tbl = new DataTable();
            DataTable non_FAI = new DataTable();
            mytbl = myCode.Mitutoyo_Get_raw_data2(txtDataFolder.Text);
            DGV_CPK.DataSource = mytbl;
            string[] col_name;
            string tar_Col = mytbl.Columns[1].ColumnName;
            string vitri_col_name = mytbl.Columns[0].ColumnName;
            string vitri_act_val = mytbl.Columns[4].ColumnName;
            string vitri_Set_val = mytbl.Columns[5].ColumnName;
            string vitri_UL_val = mytbl.Columns[7].ColumnName;
            string vitri_LL_val = mytbl.Columns[8].ColumnName;
            col_name = mytbl.AsEnumerable().Select(r => r.Field<string>(tar_Col)).Distinct().ToArray();
            //myCode.fill_datatable_DGV_Testing(mytbl, DGV_DataView, DGV_SpecView, DGV_Data_Plus);

        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Save_FAI_data();
        }

        private void btnLoadSpec_Click(object sender, EventArgs e)
        {
            string fil_loc = @"D:\Customer Projects\SEEV\TestAreas\TestData";
            string ItemCode = txtItemCode.Text;//"22B0473";
            string LotNo = txtLotNo.Text;// "00004";
            myCode.Load_Spec(sqlcon_FAI, fil_loc, ItemCode, LotNo, ".xlsm", DGV_SpecView, "");
        }
        public void initial_data()
        {
            app_path = Application.StartupPath;
            string config_file = Path.Combine(app_path, "Config", "config.txt");
            string[] my_config = myCode.read_config_arr(config_file);
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
                    string[] temp = c.Split(':');
                    string tg = "";
                    if (temp.Length > 2)
                    {

                        for (int t = 1; t < temp.Length; t++)
                        {
                            tg = tg + temp[t] + ":";
                        }
                    }
                    Data_Location = tg.TrimEnd(':').Trim();
                }
                if (c.Contains("Format_Folder"))
                {
                    format_folder = c.Split('#')[1].Trim();
                }
                if (c.Contains("Log_folder"))
                {
                    log_folder = c.Split('#')[1].Trim();
                }

            }
            string connstr_Declare = TDMK_Code.data_connection(server_name, "Declaration", server_acc, server_pass).ConnectionString;
            string connstr_Materials = TDMK_Code.data_connection(server_name, "Materials", server_acc, server_pass).ConnectionString;
            string connstr_Recycle = TDMK_Code.data_connection(server_name, "Recycled_PGC_Copper", server_acc, server_pass).ConnectionString;
            //string connstr_FAI = TDMK_Code.data_connection(server_name, "SEI_DB", server_acc, server_pass).ConnectionString;
            string connstr_FAI = TDMK_Code.data_connection(server_name, "SEI_FAI", server_acc, server_pass).ConnectionString;
            string connstr_IPQC = TDMK_Code.data_connection(server_name, "IPQC_Data", server_acc, server_pass).ConnectionString;
            string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, "OK2SHIP_Items", server_acc, server_pass).ConnectionString;
            sqlcon_Declare = new SqlConnection(connstr_Declare);
            sqlcon_Materials = new SqlConnection(connstr_Materials);
            sqlcon_Recycle = new SqlConnection(connstr_Recycle);
            sqlcon_FAI = new SqlConnection(connstr_FAI);
            sqlcon_IPQC = new SqlConnection(connstr_IPQC);
            sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);

        }
        public void Save_FAI_data()
        {
            bool data_rec_en = true;
            string[] items = new string[9];
            string[] item_vals = new string[9];
            items[0] = "ID";
            items[1] = "ItemCode";
            items[2] = "LotNo";
            items[3] = "Operator";
            items[4] = "Machine";
            items[5] = "MDate";
            items[6] = "FAI_No";
            items[7] = "FAI_Data";
            item_vals[1] = txtItemCode.Text;//"ItemCode";
            item_vals[2] = txtLotNo.Text;//"LotNo";
            item_vals[3] = txtOperator.Text;//"Operator";
            item_vals[4] = cbMachine.Text;//"Machine";
            item_vals[5] = DateTime.Now.ToString();//"MDate";
            DataTable tbl_spec = myCode.DGV_To_Table(DGV_SpecView);
            foreach (Control c in GBInfo.Controls)
            {
                if ((c.Text == "") && (c.GetType().ToString().Contains("TextBox")))
                {
                    c.BackColor = Color.Red;
                    data_rec_en = false;
                }
            }
            if (data_rec_en)
            {
                if (DGV_DataView.Rows.Count > 0)
                {
                    for (int i = 0; i < DGV_DataView.ColumnCount; i++)
                    {
                        string curr_col_name = DGV_DataView.Columns[i].Name;
                        if (myCode.check_columns_existed(tbl_spec, curr_col_name))
                        {
                            item_vals[6] = curr_col_name;
                            for (int j = 0; j < DGV_DataView.RowCount; j++)
                            {
                                if (myCode. checkDBNull(DGV_DataView.Rows[j].Cells[i].Value) != "")
                                {
                                    item_vals[0] = (TDMK_Code.SQL_MAX("FAI_Auto", "ID", sqlcon_FAI) + 1).ToString();//"ID";
                                    item_vals[7] = myCode. checkDBNull(DGV_DataView.Rows[j].Cells[i].Value);
                                    TDMK_Code.insert_val_arr("FAI_Auto", sqlcon_FAI, items, item_vals);
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }
                    }
                    DGV_DataView.Columns.Clear();
                    DGV_CPK.Columns.Clear();
                    MessageBox.Show("Ghi dữ liệu thành công", "Thông báo");
                }
                else
                {
                    MessageBox.Show("Không có dữ liệu!", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show("Nhập dữ liệu thông tin!", "Chú ý");
            }
        }

        private void txtDataFolder_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            OpenFileDialog f_file = new OpenFileDialog();
            if (f_file.ShowDialog() == DialogResult.OK)
            {
                string f_name = f_file.FileName;
                txtDataFolder.Text = f_name;
            }
        }
    }
}
