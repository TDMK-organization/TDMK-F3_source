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
using TDMK_SQL;

namespace OK2SHIP
{
    public partial class Judgement_Image : Form
    {

        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public SqlConnection sqlcon = null;
        public Judgement_Image(judge str_)
        {
            InitializeComponent();
            this.str_judge = str_;
        }
        DataTable data_tbl;
        string sheet;
        int pcs_setup;

        public DataTable data_tbl_
        {
            get { return data_tbl; }
            set { data_tbl = value; }
        }

        public string sheet_
        {
            get { return sheet; }
            set { sheet = value; }
        }

        public int pcs_setup_
        {
            get { return pcs_setup; }
            set { pcs_setup = value; }
        }
        public delegate void judge(string str_judgement);

        public judge str_judge;



        private void Judgement_Image_Load(object sender, EventArgs e)
        {
            if (data_tbl != null)
            {
                DataTable tar_dt = new DataTable();
                switch (sheet)
                {
                    case "CQRA_SOLDERABILITY":
                       
                        tar_dt.Columns.Add("Unit S/N", typeof(int));
                        tar_dt.Columns.Add("Image_Data", typeof(byte[]));
                        tar_dt.Columns.Add("Judgement", typeof(bool));

                        foreach (DataRow dr in data_tbl.Rows)
                        {
                            DataRow r = tar_dt.NewRow();
                            r["Unit S/N"] = dr["Pcs_No"];
                            r["Image_Data"] = (byte[])dr["Image_Data"];
                            if (dr["Remark"].ToString() == "PASS")
                            {
                                r["Judgement"] = true;
                            }
                            else if (dr["Remark"].ToString() == "FAIL")
                            {
                                r["Judgement"] = false;
                            }
                            tar_dt.Rows.Add(r);
                        }
                        dgv.DataSource = tar_dt;
                        ((DataGridViewImageColumn)dgv.Columns["Image_Data"]).ImageLayout = DataGridViewImageCellLayout.Zoom;

                        break;

                    case "CQRA_CHEMICAL_RESISTANCE":
                        
                        tar_dt.Columns.Add("Unit S/N", typeof(int));
                        tar_dt.Columns.Add("Picture_before", typeof(byte[]));
                        tar_dt.Columns.Add("Picture_after", typeof(byte[]));
                        tar_dt.Columns.Add("Judgement", typeof(bool));


                        DataTable dt_before = data_tbl.AsEnumerable().Where(x => x.Field<string>("Region") == "Picture before").CopyToDataTable();
                        DataTable dt_after = data_tbl.AsEnumerable().Where(x => x.Field<string>("Region") == "Picture after").CopyToDataTable();
                        int min_pcs = new int[] { dt_before.Rows.Count, dt_after.Rows.Count }.Min();

                        for (int i = 0; i < min_pcs; i++)
                        {
                            DataRow r = tar_dt.NewRow();
                            r["Unit S/N"] = dt_before.Rows[i]["Pcs_No"].ToString();
                            r["Picture_before"] = (byte[])dt_before.Rows[i]["Image_Data"];
                            r["Picture_after"] = (byte[])dt_after.Rows[i]["Image_Data"];
                            if (dt_before.Rows[i]["Remark"].ToString() == "PASS")
                            {
                                r["Judgement"] = true;
                            }
                            else if (dt_before.Rows[i]["Remark"].ToString() == "FAIL")
                            {
                                r["Judgement"] = false;
                            }
                            tar_dt.Rows.Add(r);
                        }

                        dgv.DataSource = tar_dt;
                        ((DataGridViewImageColumn)dgv.Columns["Picture_before"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                        ((DataGridViewImageColumn)dgv.Columns["Picture_after"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                        break;

                    default:
                        break;
                }

              
                dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; 
                dgv.Columns["Unit S/N"].Width = 50;
                dgv.Columns["Judgement"].Width = 100; 
                foreach (DataGridViewRow dr in dgv.Rows)
                {
                    dr.Height = 100;
                    dr.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
            }

        }

        private void dgv_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Judgement_Image_FormClosed(object sender, FormClosedEventArgs e)
        {
            string str = "";
            foreach (DataGridViewRow r in dgv.Rows)
            {
                if ((bool)r.Cells["Judgement"].Value == true)
                {
                    str += r.Cells[0].Value.ToString() + "-PASS;";
                }
                else if ((bool)r.Cells["Judgement"].Value == false)
                {
                    str += r.Cells[0].Value.ToString() + "-FAIL;";
                }
            }

            this.str_judge(str);

        }
    }
}
