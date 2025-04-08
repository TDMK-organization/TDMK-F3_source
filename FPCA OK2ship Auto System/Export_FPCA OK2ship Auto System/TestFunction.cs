using TDMK_EPPLUS_7;
using Bending_Export;
using FAI_Export;
using OK2SHIP_Lib;
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

namespace Export_FPCA_OK2ship_Auto_System
{
    public partial class TestFunction : Form
    {
        public Bending_Export_EPPLUS_Lib Bending_Exp = new Bending_Export_EPPLUS_Lib();
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public SEI_Lib myCode = new SEI_Lib();
        public TDMK_EPPLUS7_lib TDMK_Code2 = new TDMK_EPPLUS7_lib();
        public TDMK_OK2SHIP myCode2 = new TDMK_OK2SHIP();
        FAI_EPPLUS_Lib FAI_lib = new FAI_EPPLUS_Lib();
        public SqlConnection sqlcon = null;
        public TestFunction()
        {
            InitializeComponent();
        }

        private void TestFunction_Load(object sender, EventArgs e)
        {

            sqlcon = Bending_Exp.initial_data("OK2SHIP_SMT", true);
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            string tar_ItemCode = txtItemCode.Text;
            string tar_LotNo = txtLotNo.Text;
            DataTable spec_dt = FAI_lib.Load_FAI_Spec_ToTable(sqlcon, tar_ItemCode, "NPI");
            DataTable FAI_tbl = myCode2.Load_FAI_ToTable(sqlcon, tar_ItemCode, tar_LotNo, "NPI");
            DataTable CPK_tbl = FAI_lib.Load_CPK_Table(FAI_tbl, spec_dt);
            DataTable FAI_Spec_tbl = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Spec", "ItemCode = '" + tar_ItemCode + "'");
            DGV_Spec.DataSource = spec_dt;
            DGV_CPK.DataSource = CPK_tbl;
            DGV_FAI.DataSource = FAI_tbl;
        }
    }
}
