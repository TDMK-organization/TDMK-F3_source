using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using myExcel = Microsoft.Office.Interop.Excel;
using TDMK_SEEV_DLL;
using TDMK_SQL;

namespace OK2SHIP
{
    public partial class Spec_Setup : Form
    {
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SEI_Lib myCode = new SEI_Lib();
        public struct Stackup_Spec
        {
            string SetVal { get; set; }
            string Tol_U { get; set; }
            string Tol_L { get; set; }
            public Stackup_Spec(string inSetval, string inTol_U, string inTol_L)
            {
                SetVal = inSetval;
                Tol_U = inTol_U;
                Tol_L = inTol_L;
            }
        }
        public Spec_Setup()
        {
            InitializeComponent();
        }

        private void Spec_Setup_Load(object sender, EventArgs e)
        {

        }
        public void Load_Spec_StackUp_To_DGV(string file_addr, DataGridView DGV, string process_name)
        {
            char[] remove_char = new char[] { ' ', '-', '_' };
            bool format_found = false;
            myExcel.Workbook wb = TDMK_Code.open_excel_file(file_addr, "", "");
            foreach (myExcel.Worksheet wrksht in wb.Worksheets)
            {
                string sht_name = RemoveChars(wrksht.Name.ToUpper(),remove_char);
                string pro_name = RemoveChars(process_name,remove_char).ToUpper();
                if (sht_name==pro_name)
                {
                    myExcel.Worksheet ws = wrksht;
                    DGV.Rows.Add("USL", "USL");
                    DGV.Rows.Add("LSL", "LSL");
                    for (int i = 5; i < 400; i++)
                    {
                        if (ws.Cells[i, 1].value != null)
                        {
                            if (ws.Cells[i, 1].value.ToString() == "Zone" && ws.Cells[i + 1, 1].value != null)
                            {
                                DGV.Columns.Add(ws.Cells[i + 1, 1].value.ToString(), ws.Cells[i + 1, 1].value.ToString());
                                for (int j = i; j < i + 20; j++)
                                {
                                    if (ws.Cells[j, 4].Value != null)
                                    {
                                        if (ws.Cells[j, 4].Value.ToString() == "USL")
                                        {
                                            DGV.Rows[0].Cells[DGV.Columns.Count - 1].Value = ws.Cells[j, 5].Value;
                                            DGV.Rows[1].Cells[DGV.Columns.Count - 1].Value = ws.Cells[j + 1, 5].Value;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    format_found = true;
                    break;
                }
            }
            if (!format_found)
            {
                MessageBox.Show("Format not found!\r\nPlease, choose other format", "Warning");
            }
        }
        public string RemoveChars(string input, params char[] chars)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < input.Length; i++)
            {
                if (!chars.Contains(input[i]))
                    sb.Append(input[i]);
            }
            return sb.ToString();
        }
        private void btnLoad_Click(object sender, EventArgs e)
        {
            //Load_Spec_StackUp_To_DGV(txtFileLocation.Text, DGV_Spec,cbProcess.Text);

            string file_addr = txtFileLocation.Text;
            string process_name = cbProcess.Text;
            DataGridView DGV = DGV_Spec;

            char[] remove_char = new char[] { ' ', '-', '_' };
            bool format_found = false;
            myExcel.Workbook wb = TDMK_Code.open_excel_file(file_addr, "", "");
            foreach (myExcel.Worksheet wrksht in wb.Worksheets)
            {
                string sht_name = RemoveChars(wrksht.Name.ToUpper(), remove_char);
                string pro_name = RemoveChars(process_name, remove_char).ToUpper();
                if (sht_name == pro_name)
                {
                    myExcel.Worksheet ws = wrksht;
                    int r_inx = 0;
                    string rgn_start = "";
                    string rgn_stop = "";
                    myExcel.Range ref_rgn = ws.Range["B8"];
                    Dictionary<string, Stackup_Spec> result_spec = new Dictionary<string, Stackup_Spec>();
                    bool en_search = false;
                    int r_offset = 0;
                    while (r_inx<30)
                    {
                        string cur_rgn_val = myCode.checkDBNull(ref_rgn.Offset[r_inx].Value);
                        if(cur_rgn_val.ToUpper().Contains("MATERIAL"))
                        {
                            rgn_start = ref_rgn.Offset[r_inx].AddressLocal;
                        }
                        if (cur_rgn_val.ToUpper().Contains("SPEC"))
                        {
                            rgn_stop = ref_rgn.Offset[r_inx].AddressLocal;
                        }
                        if((rgn_start!="")&& (rgn_stop!=""))
                        {
                            en_search = true;
                            r_offset = ws.Range[rgn_stop].Row - ws.Range[rgn_start].Row;
                            break;
                        }
                        r_inx++;
                    }
                    if(en_search)
                    {
                        DataTable result_dt = new DataTable();
                        List<string> col_name = new List<string>() { "Zone","SetVal", "Tol_U", "Tol_L" };
                        foreach (var c in col_name)
                        {
                            result_dt.Columns.Add(c);
                        }
                        int col_inx = 0;
                        myExcel.Range start_rgn = ws.Range[rgn_start];
                        while (myCode.checkDBNull(start_rgn.Offset[0,col_inx].Value)!="")
                        {
                            string sel_rgn_val = myCode.checkDBNull(start_rgn.Offset[0, col_inx].Value);
                            if (sel_rgn_val.ToUpper().Contains("ZONE"))
                            {
                                string spec_rgn = myCode.checkDBNull(start_rgn.Offset[r_offset,col_inx].Value);
                                if(spec_rgn.Contains("±"))
                                {
                                    string[] sv_tol = spec_rgn.Split('±');
                                    result_spec.Add(sel_rgn_val, new Stackup_Spec (sv_tol[0], sv_tol[1],sv_tol[1]));
                                    result_dt.Rows.Add(sel_rgn_val,sv_tol[0], sv_tol[1], sv_tol[1]);
                                }
                                else
                                {
                                    string[] sv_tol = spec_rgn.Split('/');
                                    string sv = sv_tol[0].Split('+')[0];
                                    string tol_u = sv_tol[0].Split('+')[1];
                                    string tol_l = sv_tol[1].Replace("-","");
                                    result_spec.Add(sel_rgn_val, new Stackup_Spec(sv, tol_u, tol_l));
                                    result_dt.Rows.Add(sel_rgn_val, sv, tol_u, tol_l);
                                }
                            }
                            col_inx++;
                        }
                        DGV_Spec.DataSource = result_dt;
                    }
                    format_found = true;
                    break;
                }
            }
            if (!format_found)
            {
                MessageBox.Show("Format not found!\r\nPlease, choose other format", "Warning");
            }

        }

        private void txtFileLocation_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            OpenFileDialog f_open = new OpenFileDialog();
            f_open.InitialDirectory = Application.StartupPath;
            f_open.Filter = "Excel xlsm(*.xlsm)|*.xlsm|Excel xlsx(*.xlsx)|*.xlsx";
            if(f_open.ShowDialog()==DialogResult.OK)
            {
                if (f_open.FileName!="")
                {
                    txtFileLocation.Text = f_open.FileName;
                }
            }
        }
    }
}
