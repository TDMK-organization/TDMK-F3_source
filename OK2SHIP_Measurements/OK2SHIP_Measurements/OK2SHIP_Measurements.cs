using IniLibs;
using OK2SHIP_Lib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SQL;

namespace OK2SHIP_Measurements
{
    public partial class OK2SHIP_Measurements : Form
    {
        public static SqlConnection sqlcon_FAI;
        public static SqlConnection sqlcon_Declare;
        public static SqlConnection sqlcon_Materials;
        public static SqlConnection sqlcon_Recycle;
        public static SqlConnection sqlcon_IPQC;
        public static SqlConnection sqlcon_OK2SHIP;
        public static SqlConnection sel_sqlcon;
        public static SqlConnection sqlcon_SMT;
        public static SqlConnection sqlcon_OK2SHIP_Period2;
        public string server_name;
        public string server_acc;
        public string server_pass;
        public string DB_name;
        public string app_path;
        public string Data_Location;
        public string format_folder;
        public string log_folder;
        public string f_ext;
        public static TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        //public static SEI_Lib myCode = new SEI_Lib();
        TDMK_OK2SHIP myCode = new TDMK_OK2SHIP();
        public string tar_DB;
        public string tar_mode;
        public string tar_Dev;
        //public FrmFAI frm_FAI;///= new Form1("Manual", "SEI_FAI") { Dock = DockStyle.Fill, TopLevel = false, TopMost = true };
        //public FrmIPQC_Man frm_IPQC_Man;
        //public FrmHistogram frm_IPQC;
        public List<string> FAI_Dev_Lst = new List<string>() { "Mitutoyo", "Nikon", "Keyence" };
        public List<string> IPQC_Dev_Lst = new List<string>() { "Roughness", "AU_NI_Thickness" };
        public List<string> sel_Dev_lst = new List<string>();
        public IniFile TDMK_init = new IniFile("Config.ini");
        public string report_location;
        public OK2SHIP_Measurements()
        {
            InitializeComponent();
        }

        private void OK2SHIP_Measurements_Load(object sender, EventArgs e)
        {
            Process[] p;
            Process p1;
            p = Process.GetProcesses();
            p1 = Process.GetCurrentProcess();
            int p_count = p.Count(x => x.ProcessName == p1.ProcessName);
            if (p_count > 1)
            {
                MessageBox.Show(new Form { TopMost = true }, "OK2SHIP_Measurements đang hoạt động", "Cảnh báo");
                this.Close();
            }
            else
            {
                initial_data();
                tsmCB_Items.SelectedIndex = 0;
            }
        }
        private void tsmLoad_Click(object sender, EventArgs e)
        {
            //if ((tsmCB_Mode.Text !="") && (tsmCB_Items.Text !=""))
            //{
            //    if ((tsmCB_Mode.Text =="Manual") && (tsmCB_Items.Text == "IPQC"))
            //    {
            //        if (SEI_Lib.IPQC_Man_form ==null)
            //        {
            //            //SEI_Lib.IPQC_Man_form = new FrmIPQC_Man() { Dock = DockStyle.Fill, TopLevel = false, TopMost = true };
            //            SEI_Lib.IPQC_Man_form = new FrmIPQC_Man(log_folder, sel_sqlcon) { Dock = DockStyle.Fill, TopLevel = false, TopMost = true };
            //            this.pMain.Controls.Add(SEI_Lib.IPQC_Man_form);
            //            SEI_Lib.IPQC_Man_form.Show();
            //        }
            //    }
            //    else
            //    {
            //        if (SEI_Lib.FAI_form  == null)
            //        {
            //            SEI_Lib.FAI_form = new FrmFAI(tar_Dev, tar_DB, sel_sqlcon, Data_Location, format_folder, report_location, f_ext, log_folder) { Dock = DockStyle.Fill, TopLevel = false, TopMost = true };
            //            this.pMain.Controls.Add(SEI_Lib.FAI_form);
            //            SEI_Lib.FAI_form.Show();
            //        }
            //    }
            //    tsmLoad.Enabled = false;
            //    tsmCB_Mode.Enabled = false;
            //    tsmCB_Items.Enabled = false;

            //}
            //else
            //{
            //    MessageBox.Show(new Form { TopMost = true },"Select operation mode","Warning");
            //}
            if ((tsmCB_Mode.Text != "") && (tsmCB_Items.Text != ""))
            {
                if ((tsmCB_Mode.Text == "Manual") && (tsmCB_Items.Text == "IPQC"))
                {
                    if (myVar.IPQC_Man_form == null)
                    {
                        //TDMK_OK2SHIP.IPQC_Man_form = new FrmIPQC_Man() { Dock = DockStyle.Fill, TopLevel = false, TopMost = true };
                        myVar.IPQC_Man_form = new FrmIPQC_Man(log_folder, sel_sqlcon) { Dock = DockStyle.Fill, TopLevel = false, TopMost = true };
                        this.pMain.Controls.Add(myVar.IPQC_Man_form);
                        myVar.IPQC_Man_form.Show();
                    }
                }
                else
                {
                    if ((tsmCB_Items.Text == "IPQC") && (tsmDevice.Text == "Roughness"))
                    {
                        sel_sqlcon = sqlcon_OK2SHIP_Period2;
                        tar_DB = "OK2SHIP_Period2";
                    }

                    if (myVar.FAI_form == null)
                    {
                        myVar.FAI_form = new FrmFAI(tar_Dev, tar_DB, sel_sqlcon, Data_Location, format_folder, report_location, f_ext, log_folder) { Dock = DockStyle.Fill, TopLevel = false, TopMost = true };
                        this.pMain.Controls.Add(myVar.FAI_form);
                        myVar.FAI_form.Show();
                    }
                }
                tsmLoad.Enabled = false;
                tsmCB_Mode.Enabled = false;
                tsmCB_Items.Enabled = false;

            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Select operation mode", "Warning");
            }
        }

        private void tsmCB_Items_Click(object sender, EventArgs e)
        {

        }

        private void tsmCB_Items_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tsmCB_Items.SelectedIndex > -1)
            {
                int cur_inx = tsmCB_Items.SelectedIndex;
                switch (cur_inx)
                {
                    case 0:
                        //tar_DB = "SEI_FAI";
                        //sel_sqlcon = sqlcon_FAI;
                        tar_DB = "OK2SHIP_SMT";
                        sel_sqlcon = sqlcon_SMT;
                        sel_Dev_lst = FAI_Dev_Lst;
                        f_ext = "*.csv";
                        break;
                    case 1:
                        tar_DB = "IPQC_DATA";
                        sel_sqlcon = sqlcon_IPQC;
                        sel_Dev_lst = IPQC_Dev_Lst;
                        f_ext = "*.xls";
                        break;
                    case 2:
                        tar_DB = "OK2SHIP_Period2";
                        sel_sqlcon = sqlcon_OK2SHIP_Period2;
                        sel_Dev_lst = FAI_Dev_Lst;
                        f_ext = "*.csv";
                        break;
                    default:
                        tar_DB = "";
                        break;
                }
                if (tsmCB_Mode.SelectedIndex > -1)
                {
                    if (tsmCB_Mode.SelectedItem.ToString() == "Auto")
                    {
                        load_items(sel_Dev_lst, tsmDevice);
                    }
                    else
                    {
                        tsmDevice.Items.Clear();
                        tsmDevice.Text = "";
                        tsmDevice.Enabled = false;
                    }
                }
            }
        }
        private void tsmCB_Mode_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tsmCB_Mode.SelectedIndex > -1)
            {
                string sel_mode = tsmCB_Mode.SelectedItem.ToString();

                if (sel_mode == "Auto")
                {
                    load_items(sel_Dev_lst, tsmDevice);
                    TDMK_OK2SHIP.operate_mode = "Auto";
                }
                else
                {
                    tar_mode = sel_mode;
                    tar_Dev = "Manual";
                    TDMK_OK2SHIP.operate_mode = "Manual";
                    tsmDevice.Text = "";
                    tsmDevice.Items.Clear();
                    tsmDevice.Enabled = false;
                }
            }
        }

        private void tsmExit_Click(object sender, EventArgs e)
        {
            this.Close();

        }
        public void load_items(List<string> src_list, ToolStripComboBox tar_TSM)
        {
            tar_TSM.Text = "";
            tar_TSM.Items.Clear();
            foreach (string t in src_list)
            {
                tar_TSM.Items.Add(t);
            }
            tar_TSM.SelectedIndex = 0;
            tar_TSM.Enabled = true;
        }
        public void initial_data1()
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
                if (c.Contains("Data_LocationData_Location"))
                {
                    Data_Location = c.Split('#')[1].Trim();
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
        public void initial_data()
        {
            app_path = Application.StartupPath;

            // lấy config ngoài folder đích
            int lastSlashIndex = app_path.LastIndexOf('\\');
            if (lastSlashIndex != -1)
            {
                app_path = app_path.Substring(0, lastSlashIndex);
            }


            string config_file = Path.Combine(app_path,"config.ini");

            IniFile TDMK_init = new IniFile(config_file);


            if (!File.Exists(config_file))
            {
                TDMK_init.Write("Server", "10.212.1.243", "SMT_Config");
                TDMK_init.Write("Account", "sa", "SMT_Config");
                TDMK_init.Write("Password", "seev@123;", "SMT_Config");
                TDMK_init.Write("Report_Location", @"\\10.212.1.200\rgo\OK2SHIP automation system", "SMT_Config");
                TDMK_init.Write("Data_Location", @"\\10.212.1.200\rgo\OK2SHIP automation system\TDMK_Data\TestAreas\OK2SHIP_Measurements\OMM", "SMT_Config");
                TDMK_init.Write("Format_Folder", @"\\10.212.1.200\rgo\OK2SHIP automation system\0.OK2SHIP report format", "SMT_Config");
                TDMK_init.Write("Log_folder", @"\\10.212.1.200\rgo\OK2SHIP automation system\TDMK_Data\TestAreas\OK2SHIP_Measurements", "SMT_Config");
            }
            server_name = TDMK_init.Read("Server", "SMT_Config");
            server_acc = TDMK_init.Read("Account", "SMT_Config");
            server_pass = TDMK_init.Read("Password", "SMT_Config");
            Data_Location = TDMK_init.Read("Data_Location", "SMT_Config");
            format_folder = TDMK_init.Read("Format_Folder", "SMT_Config");
            log_folder = TDMK_init.Read("Log_folder", "SMT_Config");
            report_location = TDMK_init.Read("Report_Location", "SMT_Config");

            string connstr_Declare = TDMK_Code.data_connection(server_name, "Declaration", server_acc, server_pass).ConnectionString;
            string connstr_Materials = TDMK_Code.data_connection(server_name, "Materials", server_acc, server_pass).ConnectionString;
            string connstr_Recycle = TDMK_Code.data_connection(server_name, "Recycled_PGC_Copper", server_acc, server_pass).ConnectionString;
            string connstr_FAI = TDMK_Code.data_connection(server_name, "SEI_FAI", server_acc, server_pass).ConnectionString;
            string connstr_IPQC = TDMK_Code.data_connection(server_name, "IPQC_Data", server_acc, server_pass).ConnectionString;
            string connstr_SMT = TDMK_Code.data_connection(server_name, "OK2SHIP_SMT", server_acc, server_pass).ConnectionString;
            string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, "OK2SHIP_Items", server_acc, server_pass).ConnectionString;
            string connstr_OK2SHIP_Period2 = TDMK_Code.data_connection(server_name, "OK2SHIP_Period2", server_acc, server_pass).ConnectionString;
            sqlcon_Declare = new SqlConnection(connstr_Declare);
            sqlcon_Materials = new SqlConnection(connstr_Materials);
            sqlcon_Recycle = new SqlConnection(connstr_Recycle);
            sqlcon_FAI = new SqlConnection(connstr_FAI);
            sqlcon_IPQC = new SqlConnection(connstr_IPQC);
            sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);
            sqlcon_SMT = new SqlConnection(connstr_SMT);
            sqlcon_OK2SHIP_Period2 = new SqlConnection(connstr_OK2SHIP_Period2);
        }
        public SqlConnection initial_data(string DB_name, bool sa_en)
        {
            SqlConnection _sqlcon_OK2SHIP;
            if (!File.Exists("Config.ini"))
            {
                TDMK_init.Write("Server", "10.212.1.243", "OK2SHIP_Config");
                TDMK_init.Write("Account", "sa", "OK2SHIP_Config");
                TDMK_init.Write("Password", "seev@123;", "OK2SHIP_Config");
                TDMK_init.Write("Report_Location", @"\\10.212.1.200\rgo\OK2SHIP automation system", "OK2SHIP_Config");
                TDMK_init.Write("Data_Location", @"\\10.212.1.200\rgo\OK2SHIP automation system\TDMK_Data\TestAreas\OK2SHIP_Measurements\OMM", "OK2SHIP_Config");
                TDMK_init.Write("Format_Folder", @"\\10.212.1.200\rgo\OK2SHIP automation system\0.OK2SHIP report format", "OK2SHIP_Config");
                TDMK_init.Write("Log_folder", @"\\10.212.1.200\rgo\OK2SHIP automation system\TDMK_Data\TestAreas\OK2SHIP_Measurements", "OK2SHIP_Config");
            }
            server_name = TDMK_init.Read("Server", "OK2SHIP_Config");
            server_acc = TDMK_init.Read("Account", "OK2SHIP_Config");
            server_pass = TDMK_init.Read("Password", "OK2SHIP_Config");
            Data_Location = TDMK_init.Read("Data_Location", "OK2SHIP_Config");
            format_folder = TDMK_init.Read("Format_Folder", "OK2SHIP_Config");
            log_folder = TDMK_init.Read("Log_folder", "OK2SHIP_Config");
            report_location = TDMK_init.Read("Report_Location", "OK2SHIP_Config");
            if (sa_en)
            {
                string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, DB_name, server_acc, server_pass).ConnectionString;
                _sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);
            }
            else
            {
                string _strcon = TDMK_Code.data_connection2(server_name, DB_name).ConnectionString;
                _sqlcon_OK2SHIP = new SqlConnection(_strcon);
            }
            return _sqlcon_OK2SHIP;
        }
        private void tsmDevice_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tsmDevice.SelectedIndex > -1)
            {
                tar_Dev = tsmDevice.SelectedItem.ToString();
            }
            else
            {
                tar_Dev = "Manual";
            }
        }

        private void tsmReset_Click(object sender, EventArgs e)
        {
            if (myVar.FAI_form != null)
            {
                myVar.FAI_form.Close();
                myVar.FAI_form = null;
            }
            if (myVar.IPQC_Man_form != null)
            {
                myVar.IPQC_Man_form.Close();
                myVar.IPQC_Man_form = null;
            }
            tsmLoad.Enabled = true;
            tsmLoad.Enabled = true;
            tsmCB_Mode.Enabled = true;
            tsmCB_Items.Enabled = true;
            TDMK_OK2SHIP.admin_mode = false;
        }
    }
}
