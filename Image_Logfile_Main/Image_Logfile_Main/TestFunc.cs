using Echeck_LogFile_Process;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SEEV_DLL;
using TDMK_SQL;
using VHX;
using myExcel = Microsoft.Office.Interop.Excel;

namespace OK2SHIP
{
    public partial class TestFunc : Form
    {
        SEI_Lib myCode = new SEI_Lib();
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        ECheck_Process proc_data = new ECheck_Process();
        myVar exp_proc = new myVar();
        SqlConnection sqlcon_OK2SHIP;
        myVar _myvar = new myVar();
        SqlConnection sqlcon;
        public TestFunc()
        {
            InitializeComponent();
        }

        private void TestFunc_Load(object sender, EventArgs e)
        {
            sqlcon = _myvar.initial_data(myVar.sel_DB, true);
        }

        private void txtFile_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtFile_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            OpenFileDialog f_open = new OpenFileDialog();
            if(f_open.ShowDialog()==DialogResult.OK)
            {
                if(f_open.FileName!="")
                {
                    txtFile.Text = f_open.FileName;
                }
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //myExcel.Workbook tar_wrkbk = TDMK_Code.open_excel_file(txtFile.Text, "", "");
            //myExcel.Worksheet tar_wrksht = tar_wrkbk.Sheets[0];
            myExcel.Application xlsApp = TDMK_Code.StartExcel();
            myExcel.Workbook tar_wrkbk = xlsApp.ActiveWorkbook;
            myExcel.Worksheet tar_wrksht = tar_wrkbk.ActiveSheet;
            //string Echeck_start_rgn = proc_data.Get_start_range("Test item", "A10", tar_wrksht);
            //myExcel.Range Echeck_cycle_rgn = tar_wrksht.Range[Echeck_start_rgn].Offset[0,1];
            //txtResult.Text = Echeck_cycle_rgn.Value;

            string data_addr = proc_data.Get_start_range("Sample no.", "A10", tar_wrksht);
            myExcel.Range PTH_rgn = tar_wrksht.Range[data_addr].Offset[0, 1];
            myExcel.Range BVH_rgn = tar_wrksht.Range[data_addr].Offset[0, 33];
            PTH_rgn.Value = "PTH";
            BVH_rgn.Value = "BVH";

        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            DGV_Logfile.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "BVH_PTH_LOGFILE", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Process" }, new string[] { "71937E", "00002-99", "BVH_with_Bonding_Sheet" }));
        }
    }
}
