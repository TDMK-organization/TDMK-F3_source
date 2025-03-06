using Bending_Export;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;
using TDMK_SEEV_DLL;
using TDMK_SQL;
using static OfficeOpenXml.ExcelErrorValue;

namespace Export_FPCA_OK2ship_Auto_System
{
    public partial class NVL_SP : Form
    {
        public SqlConnection sqlcon = null;
        public Bending_Export_EPPLUS_Lib Bending_Exp = new Bending_Export_EPPLUS_Lib();
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public SEI_Lib myCode = new SEI_Lib();
        public List<string> lstLotNo_unmating = new List<string> { };

        public string ItemCodeSMT;
        public string LotNoSMT;
        public string infor_item_lot;


        public string ItemCodeSMT_
        {
            get { return ItemCodeSMT; }
            set { ItemCodeSMT = value; }
        }

        public string LotNoSMT_
        {
            get { return LotNoSMT; }
            set { LotNoSMT = value; }
        }
        public string infor_item_lot_
        {
            get { return infor_item_lot; }
            set { infor_item_lot = value; }
        }

        public NVL_SP(infor a)
        {
            InitializeComponent();
            this.str_infor = a;

        }

        public delegate void infor(string str_info);
        public infor str_infor;


        //public NVL_SP()
        //{
        //    InitializeComponent();
        //}

        private void ReferNVL_SP_Load(object sender, EventArgs e)
        {
            sqlcon = Bending_Exp.initial_data("OK2SHIP_SMT", true);
            lblItemCodeSMT.Text = ItemCodeSMT;
            lblLotNoSMT.Text = LotNoSMT; 

            string[] arr_info = infor_item_lot.Split('^'); 
            txtItemCode_unmating.Text = arr_info[0];
            cbLotNo_unmating.Text = arr_info[1];
            txtItemCode_linercoupon.Text = arr_info[2];
            cbLotNo_linercoupon.Text = arr_info[3];
            txtitemcode_psacoupon.Text = arr_info[4];
            cbLotNo_psacoupon.Text = arr_info[5]; 
        }

        private void txtItemCode_unmating_Validated(object sender, EventArgs e)
        {
            txtItemCode_unmating.Text = txtItemCode_unmating.Text.Replace(" ", "");
        }

        private void cbLotNo_unmating_SelectedIndexChanged(object sender, EventArgs e)
        {
            //cbLotNo_unmating.DataSource = lstLotNo_unmating;
        }

        private void cbLotNo_unmating_TextChanged(object sender, EventArgs e)
        {
            //string sel_lot = cbLotNo_unmating.Text;
            //cbLotNo_unmating.DataSource = null;
            //List<string> lst_lot_select = new List<string> { };
            //foreach (string lot in lstLotNo_unmating)
            //{
            //    if (lot.Contains(sel_lot))
            //    {
            //        lst_lot_select.Add(lot);
            //    }
            //}

            //cbLotNo_unmating.DataSource = lst_lot_select;
        }

        private void txtItemCode_unmating_TextChanged(object sender, EventArgs e)
        {
            cbLotNo_unmating.DataSource = null;
            cbLotNo_unmating.Text = "";
        }

        private void txtItemCode_linercoupon_Validated(object sender, EventArgs e)
        {
            txtItemCode_linercoupon.Text = txtItemCode_linercoupon.Text.Replace(" ", "");

        }

        private void txtitemcode_psacoupon_Validated(object sender, EventArgs e)
        {
            txtitemcode_psacoupon.Text = txtitemcode_psacoupon.Text.Replace(" ", "");


        }

        private void btn_OK_Click(object sender, EventArgs e)
        {
            this.str_infor( txtItemCode_unmating.Text + "^" + cbLotNo_unmating.Text + "^" + txtItemCode_linercoupon.Text + "^" + cbLotNo_linercoupon.Text + "^" + txtitemcode_psacoupon.Text + "^" + cbLotNo_psacoupon.Text);
            this.Close();
        }

        private void cbLotNo_unmating_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void txtItemCode_unmating_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (txtItemCode_unmating.Text != "")
                {
                    cbLotNo_unmating.Items.Clear();
                    DataTable dt_item = TDMK_Code.Datatable_Filter(sqlcon, "IQC_UNMATING_PULL_TEST", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode_unmating.Text }));
                    if (dt_item.Rows.Count > 0)
                    {
                        lstLotNo_unmating = dt_item.AsEnumerable().Select(x => x.Field<string>("LotNo")).Distinct().ToList();
                        cbLotNo_unmating.DataSource = lstLotNo_unmating;
                    }
                }
            }

        }

        private void txtItemCode_linercoupon_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (txtItemCode_linercoupon.Text != "")
                {
                    cbLotNo_linercoupon.Items.Clear();
                    DataTable dt_item = TDMK_Code.Datatable_Filter(sqlcon, "IQC_LINER_PEELING_COUPON", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode_linercoupon.Text }));
                    if (dt_item.Rows.Count > 0)
                    {
                        // cbLotNo_unmating.DataSource = dt_item.AsEnumerable().Select(x => x.Field<string>("LotNo")).Distinct().ToList();
                        List<string> lstLotNo_liner_coupon = dt_item.AsEnumerable().Select(x => x.Field<string>("LotNo")).Distinct().ToList();
                        foreach (string lot in lstLotNo_liner_coupon)
                        {
                            cbLotNo_linercoupon.Items.Add(lot);
                        }

                    }
                }
            }

        }

        private void txtitemcode_psacoupon_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (txtitemcode_psacoupon.Text != "")
                {
                    cbLotNo_psacoupon.Items.Clear();
                    DataTable dt_item = TDMK_Code.Datatable_Filter(sqlcon, "IQC_PSA_PEELING_COUPON", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtitemcode_psacoupon.Text }));
                    if (dt_item.Rows.Count > 0)
                    {

                        List<string> lstLotNo_psacoupon = dt_item.AsEnumerable().Select(x => x.Field<string>("LotNo")).Distinct().ToList();
                        foreach (string lot in lstLotNo_psacoupon)
                        {
                            cbLotNo_psacoupon.Items.Add(lot);
                        }
                    }
                }
            }

        }

        private void txtItemCode_Crosscut_KeyDown(object sender, KeyEventArgs e)
        {
             
        }
    }
}
