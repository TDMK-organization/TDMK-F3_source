using Funtion_F3_SMT;
using OfficeOpenXml.Style.XmlAccess;
using OK2SHIP_SMT.Services;
using Patagames.Ocr.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Security.Authentication;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.UserControls
{
    public partial class UC_Impedance : UserControl
    {

        public UC_Impedance()
        {
            InitializeComponent();

            checkLogin();
        }
        private void ClearData()
        {
            DataTable data = new DataTable();
            List<string> list = _dic.Keys.ToList();
            foreach (string item in list)
            {
                _dic[item] = data;
            }
            fillData();
        }
        private void checkLogin()
        {
            btn_logout.Visible = false;
            btn_Login.Visible = false;
            if (UserSession.Instance.IsLoggedIn == true)
            {
                btn_logout.Click += btn_Logout_Click;
                txtOperator.Text = UserSession.Instance.User_ID;
                btn_logout.Visible = true;

            }
            else
            {
                btn_logout.Click += btn_Login_Click;
                txtOperator.Text = "GUEST";

                btn_Login.Visible = true;
            }

        }
        private void btn_Logout_Click(object sender, EventArgs e)
        {
            UserSession.Instance.Logout();
            checkLogin();
        }
        private void btn_Login_Click(object sender, EventArgs e)
        {
            Login login = new Login();
            login.ShowDialog();
            checkLogin();
        }
        private void fillData()
        {
            DGV_Impedance_Summary.DataSource = _dic["IMPEDANCE_VAL"];
            DGV_Graph.DataSource = _dic["IMPEDANCE_GRAPH"];
            DGV_Impedance_Spec.DataSource = _dic["IMPEDANCE_SPEC"];
            DGV_Tracewidth_spec.DataSource = _dic["TRACEWIDTH_SPEC"];
            DGV_VHX_Data.DataSource = _dic["TRACEWIDTH_VAL"];
            DGV_Image_Tracewidth.DataSource = _dic["TRACEWIDTH_IMAGE"];
            if (DGV_Graph.Columns.Count > 0)
            {

                ((DataGridViewImageColumn)DGV_Graph.Columns["Image_Graph"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)DGV_Graph.Columns["Image_Graph"]).Width = 100;
            }
            foreach (DataGridViewRow dr in DGV_Graph.Rows)
            {
                dr.Height = 70;
            }
            if (DGV_Image_Tracewidth.Columns.Count > 0)
            {
                ((DataGridViewImageColumn)DGV_Image_Tracewidth.Columns["Image_Tracewidth"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)DGV_Image_Tracewidth.Columns["Image_Tracewidth"]).Width = 100;
            }
            foreach (DataGridViewRow dr in DGV_Image_Tracewidth.Rows)
            {
                dr.Height = 70;
            }
            DGV_Impedance.DataSource = null;



            if (DGV_Impedance_Summary.Rows.Count > 0)
            {
                check_OOS_Impedance(DGV_Impedance_Summary, DGV_Impedance_Spec);
            }
            if (DGV_VHX_Data.Rows.Count > 0)
            {
                check_OOS_Trw(DGV_VHX_Data, DGV_Tracewidth_spec);
            }
        }
        Dictionary<string, DataTable> _dic = new Dictionary<string, DataTable>();
        private void btnLoadLogFile_Click(object sender, EventArgs e)
        {
            string itemCode = tb_itemCoderefer.Text.Trim();
            string lotNo = tb_lotnorefer.Text.Trim();
            string itemCodeNew = txt_itemcode.Text.Trim();
            string lotNoNew = txt_lotNo.Text.Trim();
            try
            {
                _dic = ImpedanceService.LoadOldImpedanceData(itemCode, lotNo, itemCodeNew, lotNoNew);
                fillData();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }


        }
        public void check_OOS_Impedance(DataGridView dgv_data, DataGridView dgv_spec)
        {
            if (dgv_data.Rows.Count > 0)
            {
                if (dgv_spec.Rows.Count > 0)
                {
                    for (int i = 0; i < dgv_data.Rows.Count; i++)
                    {
                        int region = int.Parse(dgv_data.Rows[i].Cells["Region"].Value.ToString());
                        if (region <= dgv_spec.Rows.Count)
                        {
                            Double max_target = Double.Parse(dgv_spec.Rows[region - 1].Cells["USL"].Value.ToString());
                            Double min_target = Double.Parse(dgv_spec.Rows[region - 1].Cells["LSL"].Value.ToString());
                            Double data = Double.Parse(dgv_data.Rows[i].Cells["Data"].Value.ToString());
                            if (data > max_target || data < min_target)
                            {
                                for (int j = 0; j < dgv_data.Columns.Count; j++)
                                {
                                    dgv_data.Rows[i].Cells[j].Style.BackColor = Color.Red;
                                }
                            }
                            else
                            {
                                for (int j = 0; j < dgv_data.Columns.Count; j++)
                                {
                                    dgv_data.Rows[i].Cells[j].Style.BackColor = Color.White;
                                }
                            }
                        }

                    }

                }
                //else
                //{
                //    MessageBox.Show(new Form { TopMost = true }, "Format does not contain Tracewidth", "Warning");
                //}

            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, dgv_data.Name + " No data", "Warning");
            }
        }
        public void check_OOS_Trw(DataGridView dgv_data, DataGridView dgv_spec)
        {
            if (dgv_data.Rows.Count > 0)
            {
                if (dgv_spec.Rows.Count > 0)
                {
                    List<string> lst_region = ((DataTable)dgv_spec.DataSource).AsEnumerable().Select(x => x.Field<string>("Region")).Distinct().ToList();
                    for (int i = 0; i < dgv_data.Rows.Count; i++)
                    {
                        string reg = ((DataTable)dgv_data.DataSource).Rows[i]["Data_For"].ToString();
                        int region = 5;
                        int k = 0;
                        foreach (string item in lst_region)
                        {
                            if (reg.Contains(item))
                            {
                                region = k;
                                break;
                            }
                            k++;

                        }
                        if (region < dgv_spec.Rows.Count)
                        {
                            Double max_target = Double.Parse(dgv_spec.Rows[region].Cells["USL"].Value.ToString());
                            Double min_target = Double.Parse(dgv_spec.Rows[region].Cells["LSL"].Value.ToString());
                            Double data = Double.Parse(dgv_data.Rows[i].Cells["Data"].Value.ToString());
                            if (data > max_target || data < min_target)
                            {
                                for (int j = 0; j < dgv_data.Columns.Count; j++)
                                {
                                    dgv_data.Rows[i].Cells[j].Style.BackColor = Color.Red;
                                }
                            }
                            else
                            {
                                for (int j = 0; j < dgv_data.Columns.Count; j++)
                                {
                                    dgv_data.Rows[i].Cells[j].Style.BackColor = Color.White;
                                }
                            }
                        }

                        // int region = int.Parse(dgv_data.Rows[i].Cells["Data_For"].Value.ToString().Replace("Tracewidth_", ""));
                        //string x = dgv_spec.Rows[region - 1].Cells["USL"].ToString();


                    }
                }
                //else
                //{
                //    MessageBox.Show(new Form { TopMost = true }, "Format does not contain Tracewidth", "Warning");
                //}

            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, dgv_data.Name + " No data", "Warning");
            }
        }

        private void btnSummary_Click(object sender, EventArgs e)
        {
            try
            {
                _dic = new ImpedanceService().GetData(txt_itemcode.Text, txt_lotNo.Text, txtLogfile_Imp.Text);
                fillData();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}");
            }
        }

        private void btn_save_all_Click(object sender, EventArgs e)
        {
            int res = 0;
            try
            {
                ImpedanceService impedanceService = new ImpedanceService();
                res = impedanceService.Save(txt_itemcode.Text, txt_lotNo.Text, _dic);
            }
            catch (AuthenticationException ath)
            {
                MessageBox.Show(ath.Message);
                new Login().ShowDialog();
                checkLogin();
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            if (res > 0)
            {
                MessageBox.Show($"Save thành công!");
                ClearData();
            }
            else
            {
                MessageBox.Show($"lưu lỗi");
            }
        }

        private void btn_export_Click(object sender, EventArgs e)
        {
            try
            {
                new ImpedanceService().Export(txt_itemcode.Text, txt_lotNo.Text);
                MessageBox.Show("Xuất thành công");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }


        }

        private void txtLotNo_Validated(object sender, EventArgs e)
        {
            txt_lotNo.Text = Lotno_Formated(txt_lotNo.Text);
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
                catch { }
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

        private void button1_Click(object sender, EventArgs e)
        {
            string itemCode = txt_itemcode.Text.Trim();
            string lotNo = txt_lotNo.Text.Trim();
            try
            {
                _dic = new ImpedanceService().Load(itemCode, lotNo);
                fillData();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            ClearData();
        }

        private void txtLogfile_Imp_DoubleClick(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*";
            openFileDialog.Title = "Select an Excel file";
            openFileDialog.ShowDialog();
        }

        private void tableLayoutPanel21_Leave(object sender, EventArgs e)
        {

        }

        private void tb_lotnorefer_Leave(object sender, EventArgs e)
        {
            tb_lotnorefer.Text = ValidateService.lotNoHandle(tb_lotnorefer.Text);
        }

        private void txt_itemcode_Leave(object sender, EventArgs e)
        {
            txt_lotNo.Text = ValidateService.lotNoHandle(txt_lotNo.Text);

        }
    }
}
