using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SEEV_DLL;
using TDMK_SQL;
using System.Data.SqlClient;
using System.IO;

namespace Bending_Items
{
    public class myVar_ECheck
    {
        public static SEI_Lib myCode = new SEI_Lib();
        public static TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public static bool confirm_mode = false;
        public static Bending_Main frmMain;
        public static bool CSV_en = false;
        public struct edit_cell
        {
            public string col { get; set; }
            public string row { get; set; }
            public string before_val { get; set; }
            public string after_val { get; set; }
            public edit_cell(string incol, string inrow, string in_beforeval)
            {
                col = incol;
                row = inrow;
                before_val = in_beforeval;
                after_val = "";
            }
        }
        public struct item_info
        {
            public string cur_ItemCode { get; set; }
            public string cur_LotNo { get; set; }
            public string cur_Process { get; set; }
            public string cur_Cycles { get; set; }
            public item_info(string inItemCode, string inLotNo, string inProcess, string inCycles)
            {
                cur_ItemCode = inItemCode;
                cur_LotNo = inLotNo;
                cur_Process = inProcess;
                cur_Cycles = inCycles;
            }
        }
        public static string cur_User;
        public static SqlConnection initial_data(string DB_name)
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
            string connstr_OK2SHIP = TDMK_Code.data_connection(server_name, DB_name, server_acc, server_pass).ConnectionString;
            _sqlcon_OK2SHIP = new SqlConnection(connstr_OK2SHIP);
            return _sqlcon_OK2SHIP;
        }

    }
}
