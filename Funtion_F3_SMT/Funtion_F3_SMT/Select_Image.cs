using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Interop;
using TDMK_SEEV_DLL;
using TDMK_SQL;
using static Funtion_F3_SMT.Select_Image;
using DataTable = System.Data.DataTable;

namespace Funtion_F3_SMT
{
    public partial class Select_Image : Form
    {
        public SEI_Lib myCode = new SEI_Lib();
        public byte[] data_image = null;
        public byte[] data_image1 = null;
        public byte[] data_image2 = null;
        public DataTable dt_image = new DataTable();

        public string region = "";
        public string pcs = "";
        public int r_inx = 0;
        public string sheet = "";
        string column_name = "";
        int sochan = 0;


        byte[] img_null = null;


        public Select_Image(Check_en sender, judge b, shear_data shear, setchan s)
        {
            InitializeComponent();
            this.mode = sender;
            this.str_judge = b;
            this.str_sheardata = shear;
            this.s_chan = s;
        }
        public delegate void Check_en(string mode);

        public delegate void judge(string str_judge);

        public delegate void shear_data(string str_sheardata);

        public delegate void setchan(int s);

        public Check_en mode;

        public judge str_judge;

        public shear_data str_sheardata;

        public setchan s_chan;



        public DataTable dt_image_
        {
            get { return dt_image; }
            set { dt_image = value; }
        }

        public int r_inx_
        {
            get { return r_inx; }
            set { r_inx = value; }
        }

        public byte[] data
        {
            get { return data_image; }
            set { data_image = value; }
        }

        //public byte[] data1
        //{
        //    get { return data_image1; }
        //    set { data_image1 = value; }
        //}

        //public byte[] data2
        //{
        //    get { return data_image2; }
        //    set { data_image2 = value; }
        //}

        public string region_
        {
            get { return region; }
            set { region = value; }
        }

        public string column_name_
        {
            get { return column_name; }
            set { column_name = value; }
        }

        public string pcs_
        {
            get { return pcs; }
            set { pcs = value; }
        }

        public int sochan_
        {
            get { return sochan; }
            set { sochan = value; }
        }


        public Select_Image()
        {
            InitializeComponent();
        }

        public void display_image_row()
        {
            dgv_Image.Rows.Clear();
            List<string> lst_col = new List<string>();
            foreach (DataColumn dc in dt_image.Columns)
            {
                if (dc.ColumnName.Contains("Image") || dc.ColumnName.Contains("Graph"))
                {
                    lst_col.Add(dc.ColumnName);
                }
            }

            List<byte[]> lst_data_image = new List<byte[]> { };

            foreach (string col_name in lst_col)
            {
                lst_data_image.Add((byte[])dt_image.Rows[r_inx][col_name]);

            }

            //data_image = (byte[])dt_image.Rows[r_inx]["Image1"];
            //data_image1 = (byte[])dt_image.Rows[r_inx]["Image2"];


            bool select_mode = (bool)dt_image.Rows[r_inx]["Select"];
            cb_NG.Checked = !select_mode;

            lblImage_Graph.Text = "Region: " + dt_image.Rows[r_inx]["Region"].ToString() + " ----" + "Sample: " + dt_image.Rows[r_inx]["Sample"].ToString();
            using (MemoryStream ms = new MemoryStream(lst_data_image[0]))
            {
                image_detail.Image = Image.FromStream(ms);

            }

            dgv_Image.Rows.Add(lst_data_image.ToArray());

            if (lst_data_image.Count == 1)
            {
                ((DataGridViewImageColumn)dgv_Image.Columns["Column1"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)dgv_Image.Columns["Column1"]).Width = 100;

                dgv_Image.Width = 110;
                dgv_Image.Columns[1].Visible = false;
                dgv_Image.Columns[2].Visible = false;
            }
            else if (lst_data_image.Count == 2)
            {
                ((DataGridViewImageColumn)dgv_Image.Columns["Column1"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)dgv_Image.Columns["Column1"]).Width = 100;

                ((DataGridViewImageColumn)dgv_Image.Columns["Column2"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)dgv_Image.Columns["Column2"]).Width = 100;


                dgv_Image.Width = 210;
                dgv_Image.Columns[2].Visible = false;
            }
            else if (lst_data_image.Count == 3)
            {
                ((DataGridViewImageColumn)dgv_Image.Columns["Column1"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)dgv_Image.Columns["Column1"]).Width = 100;

                ((DataGridViewImageColumn)dgv_Image.Columns["Column2"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)dgv_Image.Columns["Column2"]).Width = 100;

                ((DataGridViewImageColumn)dgv_Image.Columns["Column3"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)dgv_Image.Columns["Column3"]).Width = 100;

                dgv_Image.Width = 310;
            }


            foreach (DataGridViewRow dr in dgv_Image.Rows)
            {
                dr.Height = 100;
            }
            dgv_Image.ColumnHeadersVisible = false;
            dgv_Image.RowHeadersVisible = false;

        }

        public void judgement_mode()
        {
            dgv_judgement.DataSource = null;
            if (sheet == "PEEL_TEST" || sheet == "MATING_PULL_TEST" || sheet == "SHEAR_TEST")
            {
                //  lbl_sochan.Enabled = true; 
                dgv_judgement.Visible = true;
                lbl_sochan.Visible = true;
                sumChan.Visible = true;

                if (sochan != 0)
                {
                    sumChan.Text = sochan.ToString();
                }

                DataTable dt_judge = new DataTable();
                dt_judge.Columns.Add("Mode 1: Solder joint crack");
                dt_judge.Columns.Add("Mode 2: Pad lift");
                dt_judge.Columns.Add("Mode 3: Solder joint lift");
                dt_judge.Columns.Add("Mode 4: Intermetallic break");
                dt_judge.Columns.Add("Mode 5: Component damage");
                dt_judge.Columns.Add("Mode 6: Component detached");
                dt_judge.Columns.Add("Mode 7: Flex torn");

                DataRow dr = dt_judge.NewRow();

                dr[0] = dt_image.Rows[r_inx]["Mode 1: Solder joint crack"];
                dr[1] = dt_image.Rows[r_inx]["Mode 2: Pad lift"];
                dr[2] = dt_image.Rows[r_inx]["Mode 3: Solder joint lift"];
                dr[3] = dt_image.Rows[r_inx]["Mode 4: Intermetallic break"];
                dr[4] = dt_image.Rows[r_inx]["Mode 5: Component damage"];
                dr[5] = dt_image.Rows[r_inx]["Mode 6: Component detached"];
                dr[6] = dt_image.Rows[r_inx]["Mode 7: Flex torn"];

                dt_judge.Rows.Add(dr);
                dgv_judgement.DataSource = dt_judge;
            }
            else
            {
                dgv_judgement.Visible = false;
                lbl_sochan.Visible = false;
                sumChan.Visible = false;
            }
        }

        private void Select_Image_Load(object sender, EventArgs e)
        {
            Bitmap img = new Bitmap(100, 50, PixelFormat.Format32bppArgb);
            //var img = Bitmap.FromFile(Path.Combine(System.Windows.Forms.Application.StartupPath, "Img_null", "Img_null.jpg"));
            ImageConverter imgcon = new ImageConverter();
            img_null = (byte[])imgcon.ConvertTo(img, typeof(byte[]));
            if (sheet == "SHEAR_TEST")
            {
                lbl_data.Enabled = true;
            }
            else
            {
                lbl_data.Enabled = false;
            }

            if (sheet == "CROSS_SECTION" || sheet == "GAP_CONNECTOR")
            {
                lbl_data.Visible = false;
                display_image_row();

                if ((bool)dt_image.Rows[r_inx]["Select"] == false)
                {
                    cb_NG.BackColor = Color.Red;

                }
                else
                {
                    cb_NG.BackColor = Color.White;
                }
            }

            else
            {
                dgv_Image.Visible = false;
                lblImage_Graph.Text = "Region: " + dt_image.Rows[r_inx]["Region"].ToString() + " ----" + "Sample: " + dt_image.Rows[r_inx]["Sample"].ToString();

                using (MemoryStream ms = new MemoryStream((byte[])dt_image.Rows[r_inx][column_name]))
                {
                    image_detail.Image = Image.FromStream(ms);
                }
                string col_select = "";
                if (column_name == "Image")
                {
                    col_select = "Select_Img";
                    lbl_data.Visible = false;
                    judgement_mode();
                }
                else if (column_name == "Graph")
                {
                    col_select = "Select_Grp";
                    lbl_data.Visible = true;
                    lbl_data.Text = dt_image.Rows[r_inx]["Data"].ToString();
                }

                if (col_select != "")
                {
                    bool select_mode = (bool)dt_image.Rows[r_inx][col_select];
                    cb_NG.Checked = !select_mode;

                    if ((bool)dt_image.Rows[r_inx][col_select] == false)
                    {
                        cb_NG.BackColor = Color.Red;

                    }
                    else
                    {
                        cb_NG.BackColor = Color.White;
                    }
                }
            }


            if (r_inx == 0)
            {
                btn_prev.Enabled = false;
            }
            if (r_inx == dt_image.Rows.Count - 1)
            {
                btn_next.Enabled = false;
            }



            //if(myCode.IsNumeric(lbl_sochan.Text.Replace("Tổng số chân:", "").Replace(" ", "")))
            //{
            //    sochan = int.Parse(lbl_sochan.Text.Replace("Tổng số chân:", "").Replace(" ", ""));
            //}



        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if ((byte[])dgv_Image.CurrentCell.Value != img_null)
            {
                using (MemoryStream ms = new MemoryStream((byte[])dgv_Image.CurrentCell.Value))
                {
                    image_detail.Image = Image.FromStream(ms);

                }
            }

        }
        private void UpdateData()
        {
            DataTable dataTable = (DataTable)dgv_judgement.DataSource;
         
            foreach(DataColumn col in dataTable.Columns)
            {
                dt_image.Rows[r_inx][col.ColumnName] = dataTable.Rows[0][col];
            }
        }
        private void btn_next_Click(object sender, EventArgs e)
        {
            UpdateData();
            btn_prev.Enabled = true;
            btn_next.Enabled = true;
            if (r_inx < dt_image.Rows.Count - 1)
            {
                r_inx++;
                if (sheet == "CROSS_SECTION" || sheet == "GAP_CONNECTOR")
                {
                    display_image_row();
                    if ((bool)dt_image.Rows[r_inx]["Select"] == false)
                    {
                        cb_NG.BackColor = Color.Red;
                    }
                    else
                    {
                        cb_NG.BackColor = Color.White;
                    }
                }
                else
                {
                    dgv_Image.Visible = false;
                    lblImage_Graph.Text = "Region: " + dt_image.Rows[r_inx]["Region"].ToString() + " ----" + "Sample: " + dt_image.Rows[r_inx]["Sample"].ToString();

                    using (MemoryStream ms = new MemoryStream((byte[])dt_image.Rows[r_inx][column_name]))
                    {
                        image_detail.Image = Image.FromStream(ms);
                    }

                    string col_select = "";
                    if (column_name == "Image")
                    {
                        col_select = "Select_Img";
                        judgement_mode();
                    }
                    else if (column_name == "Graph")
                    {
                        col_select = "Select_Grp";
                        lbl_data.Text = dt_image.Rows[r_inx]["Data"].ToString();
                    }

                    if (col_select != "")
                    {
                        bool select_mode = (bool)dt_image.Rows[r_inx][col_select];
                        cb_NG.Checked = !select_mode;

                        if ((bool)dt_image.Rows[r_inx][col_select] == false)
                        {
                            cb_NG.BackColor = Color.Red;

                        }
                        else
                        {
                            cb_NG.BackColor = Color.White;
                        }
                    }
                }
            }
            else
            {
                btn_next.Enabled = false;
            }


        }

        private void btn_prev_Click(object sender, EventArgs e)
        {

            UpdateData();
            btn_prev.Enabled = true;
            btn_next.Enabled = true;
            if (r_inx > 0)
            {

                r_inx--;
                if (sheet == "CROSS_SECTION" || sheet == "GAP_CONNECTOR")
                {
                    display_image_row();
                    if ((bool)dt_image.Rows[r_inx]["Select"] == false)
                    {
                        cb_NG.BackColor = Color.Red;
                        ;
                    }
                    else
                    {
                        cb_NG.BackColor = Color.White;
                    }
                }
                else
                {
                    dgv_Image.Visible = false;
                    lblImage_Graph.Text = "Region: " + dt_image.Rows[r_inx]["Region"].ToString() + " ----" + "Sample: " + dt_image.Rows[r_inx]["Sample"].ToString();

                    using (MemoryStream ms = new MemoryStream((byte[])dt_image.Rows[r_inx][column_name]))
                    {
                        image_detail.Image = Image.FromStream(ms);
                    }

                    string col_select = "";
                    if (column_name == "Image")
                    {
                        col_select = "Select_Img";
                        judgement_mode();
                    }
                    else if (column_name == "Graph")
                    {
                        col_select = "Select_Grp";
                        lbl_data.Text = dt_image.Rows[r_inx]["Data"].ToString();
                    }

                    if (col_select != "")
                    {
                        bool select_mode = (bool)dt_image.Rows[r_inx][col_select];
                        cb_NG.Checked = !select_mode;

                        if ((bool)dt_image.Rows[r_inx][col_select] == false)
                        {
                            cb_NG.BackColor = Color.Red;

                        }
                        else
                        {
                            cb_NG.BackColor = Color.White;
                        }
                    }

                }
            }
            else
            {
                btn_prev.Enabled = false;
            }
        }

        private void cb_NG_CheckedChanged(object sender, EventArgs e)
        {
            if (sheet == "CROSS_SECTION" || sheet == "GAP_CONNECTOR")
            {
                if (cb_NG.Checked)
                {
                    this.mode(r_inx.ToString() + "_false");
                    cb_NG.BackColor = Color.Red;
                    dt_image.Rows[r_inx]["Select"] = false;

                }
                else
                {
                    this.mode(r_inx.ToString() + "_true");
                    cb_NG.BackColor = Color.White;
                    dt_image.Rows[r_inx]["Select"] = true;
                }
            }
            else
            {
                string col_select = "";
                if (column_name == "Image")
                {
                    col_select = "Select_Img";
                }
                else if (column_name == "Graph")
                {
                    col_select = "Select_Grp";
                }

                int r_ins = r_inx;
                if (col_select != "")
                {
                    if (cb_NG.Checked)
                    {
                        this.mode(r_inx.ToString() + "_false");
                        cb_NG.BackColor = Color.Red;
                        dt_image.Rows[r_inx][col_select] = false;

                    }
                    else
                    {
                        this.mode(r_inx.ToString() + "_true");
                        cb_NG.BackColor = Color.White;
                        dt_image.Rows[r_inx][col_select] = true;
                    }
                }
            }
        }

        private void dgv_judgement_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (sumChan.Text != "")
            {
                int col = e.ColumnIndex;
                int row = e.RowIndex;
                string value = dgv_judgement.Rows[row].Cells[col].Value.ToString();
                if (double.TryParse(value, out double num) && num > 0)
                {
                    if (!(value.Contains("%") && value.Contains("(")))
                    {
                        dgv_judgement.Rows[row].Cells[col].Value = $"{(num * 100 / (double)sochan).ToString("#.##")}%({num}/{sochan})";
                        Calculate();
                    }

                }
                else
                {
                    if (!value.Contains("%") || !value.Contains("("))
                    {
                        dgv_judgement.Rows[row].Cells[col].Value = $"0.00%(0/{sochan})";
                    }
                }
            }
        }

        private void tableLayoutPanel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void lbl_data_TextChanged(object sender, EventArgs e)
        {
            this.str_sheardata(r_inx.ToString() + "_" + lbl_data.Text);
        }

        private void btn_next_Resize(object sender, EventArgs e)
        {


        }

        private void txt_sochan_Validated(object sender, EventArgs e)
        {
            if (myCode.IsNumeric(sumChan.Text))
            {
                this.s_chan(int.Parse(sumChan.Text));
            }
        }

        private void txt_sochan_TextChanged(object sender, EventArgs e)
        {
            if (!myCode.IsNumeric(sumChan.Text))
            {
                sumChan.BackColor = Color.Yellow;
            }
            else
            {
                sumChan.BackColor = Color.White;
            }
            string st = sumChan.Text;
            if (!string.IsNullOrEmpty(st))
            {

            }
        }

        private void dgv_judgement_DataSourceChanged(object sender, EventArgs e)
        {
            Calculate();
        }
        private void Calculate()
        {
            int sum = 0;
            DataTable dataTable = (DataTable)dgv_judgement.DataSource;
            if (dgv_judgement.DataSource != null)
            {
                foreach (DataColumn col in dataTable.Columns)
                {
                    string value = dataTable.Rows[0][col].ToString();
                    if (value.Contains("/") || value.Contains('('))
                    {
                        string z = value.Split('/')[0].Split('(')[1];
                        if (int.TryParse(z, out int res))
                        {
                            sum += res;
                        }
                    }
                }
            }
            sumChan.Text = sum.ToString();
            if (sum > sochan || sum < sochan)
            {
                sumChan.BackColor = Color.Red;
            }
            else
            {
                sumChan.BackColor = Color.Green;
            }
        }
    }
}
