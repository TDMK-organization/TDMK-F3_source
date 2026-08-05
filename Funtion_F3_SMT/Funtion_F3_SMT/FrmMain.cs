
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SEEV_DLL;
using TDMK_SQL;
using System.Data.SqlClient;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Application = System.Windows.Forms.Application;
using OK2SHIP_SMT;
using OK2SHIP_SMT.UserControls;

namespace Funtion_F3_SMT
{
    public partial class FrmMain : Form
    {

        public static TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SEI_Lib myCode = new SEI_Lib();

        public FrmMain()
        {
            InitializeComponent();
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {

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
            string Data_Location = "";
            string format_folder = "";
            string log_folder = "";
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
            if (sa_en)
            {
                string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, DB_name, server_acc, server_pass).ConnectionString;
                _sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);
            }
            else
            {
                string _strcon = data_connection2(server_name, DB_name).ConnectionString;
                _sqlcon_OK2SHIP = new SqlConnection(_strcon);
            }

            return _sqlcon_OK2SHIP;
        }
        public SqlConnectionStringBuilder data_connection2(string _servername, string _databasename)
        {
            return new SqlConnectionStringBuilder
            {
                DataSource = _servername,
                InitialCatalog = _databasename,
                IntegratedSecurity = true
            };
        }
        public string[] read_config_arr(string src_config_file)
        {
            string[] _result = new string[1000];
            StreamReader reader;
            int inx = 0;
            reader = new StreamReader(src_config_file);
            while ((!reader.EndOfStream) && (inx < 1000))
            {
                string temp = reader.ReadLine();
                if (temp != "")
                {
                    _result[inx] = temp;
                    inx++;
                }
            }
            Array.Resize<string>(ref _result, inx);
            reader.Close();
            reader.Dispose();
            return _result;
        }

        private void btnreset_Click(object sender, EventArgs e)
        {
            if (pMain != null)
            {
                this.pMain.Controls.Clear();
            }
            //tsmLoad.Enabled = true;
            //tsmLoad.Enabled = true;
            //tsmCB_Mode.Enabled = true;
            //tsmCB_Items.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void mnuLoad_Click(object sender, EventArgs e)
        {
            string sheet = mnuProcess.SelectedText;
            this.WindowState = FormWindowState.Maximized;
            switch (sheet)
            {
                case "Assy Yield":
                    pMain.Controls.Clear();
                    pMain.Controls.Add(new UC_Assy_Yield() { Dock = DockStyle.Fill });
                    break;
                case "ACF":
                    ACF frm2 = new ACF() { Dock = DockStyle.Fill, TopLevel = false, TopMost = true };
                    this.pMain.Controls.Add(frm2);
                    frm2.Show();
                    mnuProcess.Enabled = false;
                    break;
                case "Impedance":
                    pMain.Controls.Clear();
                    pMain.Controls.Add(new UC_Impedance() { Dock = DockStyle.Fill });
                    break;
                case "SEM BSE & Binarization":
                case "Environment en-durance":
                case "OQC B2B Mating-Unmating":
                case "Bar Code Verification":
                case "X-Ray picture":
                case "Thermal cycling, Heat soak, Thermal shock":
                case "Air Bubble":
                case "Peel Test (On Product)":
                    pMain.Controls.Clear();
                    pMain.Controls.Add(new SEM(sheet) { Dock = DockStyle.Fill });
                    break;
                case "IQC Peeling Test":
                    pMain.Controls.Clear();
                    pMain.Controls.Add(new UC_IQCPeelingTest() { Dock = DockStyle.Fill });
                    break;
                default:
                    if (sheet != "")
                    {
                        View_Data_3 frm1 = new View_Data_3() { Dock = DockStyle.Fill, TopLevel = false, TopMost = true };
                        this.pMain.Controls.Add(frm1);
                        frm1.sheet_ = sheet;
                        // frm1.admin_mode_ = mnu_Login.Text;
                        frm1.Show();
                        mnuProcess.Enabled = false;
                    }
                    break;
            }
        }

        private void mnuReset_Click(object sender, EventArgs e)
        {
            if (pMain != null)
            {
                this.pMain.Controls.Clear();
            }
            mnuProcess.Enabled = true;
        }

        private void mnuExit_Click(object sender, EventArgs e)
        {
            Close();
            Dispose();
        }

        private void FrmMainVHX_FormClosed(object sender, FormClosedEventArgs e)
        {
            //myVar._frmMain.Show();
            MainScreen frm = new MainScreen();
            frm.Show();
        }

        private void mnuProcess_Click(object sender, EventArgs e)
        {

        }

        //private void Setmode(string value)
        //{
        //    this.mnu_Login.Text = value;
        //    if (value == "Admin mode")
        //    {
        //        this.mnu_Login.BackColor = Color.GreenYellow;
        //    }
        //    else
        //    {
        //        this.mnu_Login.BackColor = Color.Yellow;
        //    }

        //}

        //private void mnu_Login_Click_1(object sender, EventArgs e)
        //{
        //    Login fr1 = new Login(Setmode);
        //    if( mnu_Login.Text == "Admin mode")
        //    {
        //        fr1.mode = true;
        //    }
        //    else
        //    {
        //        fr1.mode = false;
        //    }

        //    fr1.Show();

        //    //if (admin_mode == "Admin mode")
        //    //{
        //    //    mnu_Login.Text = "Admin mode";
        //    //    mnu_Login.BackColor = Color.GreenYellow;
        //    //}
        //    //else
        //    //{
        //    //    mnu_Login.Text = "LOGIN";
        //    //    mnu_Login.BackColor = Color.Yellow;
        //    //}


        //}

        private void mnu_Login_TextChanged(object sender, EventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }
    }
}
