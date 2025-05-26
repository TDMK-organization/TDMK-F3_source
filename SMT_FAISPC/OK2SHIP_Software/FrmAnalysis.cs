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
using OK2SHIP_Lib;
using TDMK_SQL;
using System.IO;

namespace OK2SHIP_Software
{
    public partial class FrmAnalysis : Form
    {
        myVar _myvar = new myVar();
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        myVar myCode = new myVar();
        
        public FrmAnalysis()
        {
            InitializeComponent();
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
start_lbl:  if (splMain.Panel2.Controls.Count==0)
            {
                FAI_Data frmFAI = new FAI_Data(txtItemCode.Text, txtLotNo.Text) { Dock = DockStyle.Fill, TopLevel = false, TopMost = true };
                frmFAI.Show();
                splMain.Panel2.Controls.Add(frmFAI);
            }
            else
            {
                splMain.Panel2.Controls.Clear();
                goto start_lbl;
            }
        }

        private void FrmAnalysis_Load(object sender, EventArgs e)
        {
            myVar.sqlcon_SMT = _myvar.initial_data("OK2SHIP_SMT", true);
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            string format_file = "";
            string _itemcode = txtItemCode.Text;
            string _lotno = txtLotNo.Text;
            List<string> format_lst = get_multiple_format(myVar.format_loc, new List<string> { "*.xlsx", "*.xlsm" }, _itemcode, "FAI");
            if(format_lst.Count>0)
            {
                DataTable FAI_Data_tbl = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { _itemcode, _lotno }));
                if(FAI_Data_tbl.Rows.Count>0)
                {
                    format_file = format_lst[0];
                    string export_file = Path.Combine(myVar.report_loc, Path.GetFileNameWithoutExtension(format_file) + "-" + _lotno + Path.GetExtension(format_file));
                    myExcel.Workbook report_saved = null;
                    if (System.IO.File.Exists(export_file))
                    {
                        report_saved = TDMK_Code.open_excel_file(export_file, "", "");
                    }
                    else
                    {
                        report_saved = TDMK_Code.open_excel_file(format_file, "", "");
                        report_saved.SaveAs(export_file);
                    }

                    myExcel.Workbook exp_file;
                    exp_file = TDMK_Code.Create_workbook();
                    myCode.Export_FAI_Batch(_itemcode, _lotno, exp_file);
                    myCode.Export_To_FAI2(report_saved, _itemcode, _lotno, exp_file);
                }  
                else
                {
                    MessageBox.Show("Không tìm thấy dữ liệu của ItemCode / Lotno: " + _itemcode + "/" + _lotno, "Thông báo");
                }    
            }
            else
            {
                MessageBox.Show("Không tìm thấy Format", "Cảnh báo");
            }
        }
        public List<string> get_multiple_files(string src_path, List<string> extensions, string tar_ItemCode, string tar_process)
        {
            List<string> result = new List<string>();
            List<char> remove_char = new List<char> { ' ', '-', '_' };
            string process = new string(tar_process.Where(x => remove_char.IndexOf(x) == -1).ToArray());
            DirectoryInfo directory = new DirectoryInfo(src_path);
            var files = extensions.SelectMany(e => directory.EnumerateFiles(e, SearchOption.AllDirectories)).Where(x => x.FullName.Contains(tar_ItemCode) && remove_special_char(Path.GetFileNameWithoutExtension(x.Name), remove_char).ToUpper() == process.ToUpper());
            foreach (var item in files)
            {
                result.Add(item.FullName);
            }
            return result;
        }
        public List<string> get_multiple_format(string src_path, List<string> extensions, string tar_ItemCode, string tar_process)
        {
            List<string> result = new List<string>();
            List<char> remove_char = new List<char> { ' ', '-', '_' };
            string process = new string(tar_process.Where(x => remove_char.IndexOf(x) == -1).ToArray());
            DirectoryInfo directory = new DirectoryInfo(src_path);

            var files = extensions.SelectMany(e => directory.EnumerateFiles(e, SearchOption.AllDirectories)).Where(x =>x.FullName.Contains(tar_ItemCode) && x.FullName.Contains(tar_process));
            foreach (var item in files)
            {
                result.Add(item.FullName);
            }
            return result;
        }
        public List<string> get_multiple_format(string src_path, List<string> extensions, string tar_ItemCode)
        {
            List<string> result = new List<string>();
            List<char> remove_char = new List<char> { ' ', '-', '_' };
            DirectoryInfo directory = new DirectoryInfo(src_path);
            var files = extensions.SelectMany(e => directory.EnumerateFiles(e, SearchOption.AllDirectories)).Where(x => x.FullName.Contains(tar_ItemCode));
            foreach (var item in files)
            {
                result.Add(item.FullName);
            }
            return result;
        }
        public string remove_special_char(string src_str, List<char> remove_char)
        {
            return new string(src_str.Where(x => remove_char.IndexOf(x) == -1).ToArray());
        }
    }
}
