using Export_FPCA_OK2ship_Auto_System.Repositories;
using FAI_Export;
using IniLibs;
using OfficeOpenXml;
using OK2SHIP_Lib;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SEEV_DLL;
using TDMK_SQL;


namespace Export_FPCA_OK2ship_Auto_System.Services
{
    public class FAIService
    {
        #region Properties
        private readonly string NAME_TABLE_SQL = "TC_HS_TS";
        private DBContext _dbContext;
        private TDMK_OK2SHIP myCode;
        private TDMK_SQL_Lib TDMK_Code;
        private FAI_EPPLUS_Lib FAI_Lib;
        //private string format_folder;
        //private string report_location;
        private SqlConnection sqlcon;
        private string FORMAT_LOACTION = null;
        private string EXPORT_LOACTION = null;
        private string EXPORT_LOACTION_LAST = null;
        IniFile TDMK_init = new IniFile();

        #endregion
        #region Constructor
        public FAIService()
        {
            _dbContext = new DBContext();
            myCode = new TDMK_OK2SHIP();
            TDMK_Code = new TDMK_SQL_Lib();
            FAI_Lib = new FAI_EPPLUS_Lib();
            sqlcon = _dbContext.SqlConnection;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }
        #endregion
        #region Methods
        public string find_config_path(string src_string, string f_name)
        {
            DirectoryInfo di = new DirectoryInfo(src_string);
            var folder_lst = di.GetDirectories().Select(x => x.FullName).ToList();
            var temp2 = folder_lst.Where(x => new DirectoryInfo(x).Name == f_name).ToList();
            if (temp2.Count != 0)
            {
                return temp2.FirstOrDefault();
            }
            else
            {
                if (di.Parent != null)
                {
                    return find_config_path(di.Parent.FullName, f_name);
                }
                else
                {
                    return string.Empty;
                }
            }
        }
        public List<string> get_multiple_format(string src_path, List<string> extensions, string tar_ItemCode)
        {
            List<string> result = new List<string>();
            List<char> remove_char = new List<char> { ' ', '-', '_' };
            DirectoryInfo directory = new DirectoryInfo(src_path);
            var files = extensions.SelectMany(e => directory.EnumerateFiles(e, SearchOption.TopDirectoryOnly)).Where(x => x.FullName.Contains(tar_ItemCode));
            foreach (var item in files)
            {
                result.Add(item.FullName);
            }
            return result;
        }
        public string Export_FAI_Data_byEPPLUS(string _itemcode, string _lotno, string _type="NPI")
        {
            //string format_loc = Path.Combine(format_folder, "MASS","FAI");
            string msg = "";
            //format_folder = Path.Combine(find_config_path(Application.StartupPath, "SEEV Data"), "Format");
            //report_location = Path.Combine(find_config_path(Application.StartupPath, "SEEV Data"), "Report");
            //ExportProcess exportProcess = new ExportProcess();
            string app_path = System.Windows.Forms.Application.StartupPath.Replace(@"\FPCA OK2SHIP Auto System", "");
            string config_path = Path.Combine(app_path, "config.ini");
            TDMK_init = new IniFile(config_path);
            FORMAT_LOACTION = TDMK_init.Read("Format_Folder", "SMT_Config") + $"\\SEEV Data\\Format\\{_type}";
            EXPORT_LOACTION = TDMK_init.Read("NasAddress", "SMT_Config").Replace("\\ImageF3", "") + $"\\Report\\{_type}\\SOFTWARE(NOTOUCH)";
            //EXPORT_LOACTION_LAST = TDMK_init.Read("Report_Location", "SMT_Config") + $"\\SEEV Data\\Report\\{_type}";
            string format_loc = FORMAT_LOACTION;
            string report_location = EXPORT_LOACTION;
            //if (_type == "NPI")
            //{
            //    format_loc = Path.Combine(format_folder, "NPI");
            //}
            //else
            //{
            //    format_loc = Path.Combine(format_folder, "MASS", "FAI");
            //}

            DataTable FAI_dt = myCode.Load_FAI_ToTable(sqlcon, _itemcode, _lotno, _type);
            DataTable FAI_Spec = myCode.Load_FAI_Spec_ToTable(sqlcon, _itemcode, _type);
            bool export_en = false;
            if (myCode.check_FAIdata_inSpec(FAI_dt, FAI_Spec))
            {
                export_en = true;
            }
            else
            {
                if (MessageBox.Show("Dữ liệu NG. Tiếp tục xuất dữ liệu?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    export_en = true;
                    msg = "Cảnh báo dữ liệu NG\r\n";
                }
                else
                {
                    export_en = false;
                    msg = "Dữ liệu NG";
                }

            }
            if (export_en)
            {
                List<string> format_lst = get_multiple_format(format_loc, new List<string> { "*.xlsx", "*.xlsm" }, _itemcode);
                if (format_lst.Count > 0)
                {
                    string report_type = _type;
                    DataTable FAI_Data_tbl = TDMK_Code.Datatable_Filter(sqlcon, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Remark" }, new string[] { _itemcode, _lotno, _type }));
                    if (FAI_Data_tbl.Rows.Count > 0)
                    {
                        string format_file = format_lst[0];
                        string export_file = "";
                        if (_type == "MASS")
                        {
                            string report_path = Path.Combine(report_location, "OMM_Dimension");
                            if (!Directory.Exists(report_path))
                            {
                                Directory.CreateDirectory(report_path);
                            }
                            export_file = Path.Combine(report_path, Path.GetFileNameWithoutExtension(format_file) + "-" + _lotno + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(format_file));

                        }
                        else
                        {
                            //string report_path = Path.Combine(report_location, "NPI", "FAI");
                            string report_path = Path.Combine(report_location, "FAI");
                            if (!Directory.Exists(report_path))
                            {
                                Directory.CreateDirectory(report_path);
                            }
                            export_file = Path.Combine(report_path, _itemcode + "_" + _lotno + Path.GetExtension(format_file));
                        }
                        ExcelPackage Report_Pack = new ExcelPackage();
                        ExcelWorkbook report_saved = null;
                        bool file_existed = false;
                        if (System.IO.File.Exists(export_file))
                        {
                            Report_Pack = FAI_Lib.open_excel(export_file);
                            file_existed = true;
                        }
                        else
                        {
                            Report_Pack = FAI_Lib.open_excel(format_file);
                            file_existed = false;
                            //Report_Pack.SaveAs(new FileInfo(export_file));
                        }
                        report_saved = Report_Pack.Workbook;
                        int error_count = 0;
                        while (report_saved.Names.Count > 0 && error_count < 5)
                        {
                            for (int i = 0; i < report_saved.Names.Count; i++)
                            {
                                try
                                {
                                    string cur_name = report_saved.Names[i].Name;
                                    report_saved.Names.Remove(cur_name);
                                }
                                catch
                                {
                                    error_count++;
                                    continue;
                                }
                            }
                        }
                        try
                        {
                            report_saved.ExternalLinks.Clear();
                        }
                        catch
                        {

                        }
                        Dictionary<string, DataTable> dic_data = FAI_Lib.Export_FAI_Batch(sqlcon, report_saved, _itemcode, _lotno, report_type);
                        FAI_Lib.Export_To_FAI(sqlcon, Report_Pack, _itemcode, _lotno, dic_data, report_type);
                        if (file_existed)
                        {
                            Report_Pack.Save();
                        }
                        else
                        {
                            Report_Pack.SaveAs(new FileInfo(export_file));
                        }
                        Report_Pack.Dispose();
                        msg += "OK";
                        //MessageBox.Show(new Form { TopMost = true }, "Hoàn thành xuất dữ liệu", "Thông báo");
                        //ProcessStartInfo pi = new ProcessStartInfo(export_file);
                        //Process.Start(pi);
                    }
                    else
                    {
                        msg += "Không tìm thấy dữ liệu của ItemCode / Lotno / Shift : " + _itemcode + " / " + _lotno;
                        //MessageBox.Show(new Form { TopMost = true }, "Không tìm thấy dữ liệu của ItemCode / Lotno / Shift : " + _itemcode + " / " + _lotno, "Thông báo");
                    }
                }
                else
                {
                    //MessageBox.Show(new Form { TopMost = true }, "Không tìm thấy Format", "Cảnh báo");
                    msg += "Không tìm thấy Format";
                }
            }
            return msg;
        }
        #endregion
    }
}
