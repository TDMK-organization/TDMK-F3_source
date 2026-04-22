using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FAI_Export_EPPLUS;
using OK2SHIP_Lib;
using TDMK_EPPLUS_7;
using TDMK_SQL;
using OfficeOpenXml;
using System.IO;

namespace OK2SHIP_Measurements
{
    public class PTH_Diameter_Process
    {
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public TDMK_OK2SHIP myCode = new TDMK_OK2SHIP();
        public TDMK_EPPLUS7_lib TDMK_EPPLUS7 = new TDMK_EPPLUS7_lib();
        public bool PTH_Diameter_Spec_Setup(string format_file, string _ItemCode, SqlConnection sqlcon, string DB_table_name)
        {
            bool result = false;
            DataTable format_spec_dt = PTH_Diameter_Load_Spec_fromFile(sqlcon, format_file, _ItemCode, new List<string> { "*.xlsx", "*.xlsm" });
            if (format_spec_dt.Rows.Count != 0)
            {
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { _ItemCode });
                TDMK_Code.Delelte_FilteredItem_arr(DB_table_name, sqlcon, filter_str);
                DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon, DB_table_name, filter_str);
                int id = Convert.ToInt32(TDMK_Code.SQL_MAX(DB_table_name, "ID", sqlcon)) + 1;
                foreach (DataColumn dc in format_spec_dt.Columns)
                {
                    DataRow dr = spec_dt.NewRow();
                    dr[0] = id++;
                    dr[1] = _ItemCode;
                    dr[2] = TDMK_EPPLUS7.checkDBNull(format_spec_dt.Rows[0][dc]);
                    dr[3] = TDMK_EPPLUS7.checkDBNull(format_spec_dt.Rows[1][dc]);
                    dr[4] = TDMK_EPPLUS7.checkDBNull(format_spec_dt.Rows[2][dc]);
                    spec_dt.Rows.Add(dr);
                }
                BatchBulkCopy(sqlcon, spec_dt, DB_table_name);
                result = true;
            }
            return result;
        }
        public DataTable Load_PTH_Diameter_Spec(string _ItemCode, SqlConnection sqlcon, string DB_table_name)
        {
            DataTable result = new DataTable();
            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { _ItemCode });
            DataTable spec_dt = TDMK_Code.Datatable_Filter(sqlcon, DB_table_name, filter_str);
            foreach (DataRow dr in spec_dt.Rows)
            {
                string setval = TDMK_EPPLUS7.checkDBNull(dr["Normdim"]);
                string USL = TDMK_EPPLUS7.checkDBNull(dr["USL"]);
                string LSL = TDMK_EPPLUS7.checkDBNull(dr["LSL"]);
                string col_name = "PTH-Diameter_" + setval;
                if (!result.Columns.Contains(col_name))
                {
                    result.Columns.Add(col_name);
                    if (result.Rows.Count == 0)
                    {
                        for (int i = 0; i < 3; i++)
                        {
                            result.Rows.Add();
                        }
                    }
                    result.Rows[0][col_name] = setval;
                    result.Rows[1][col_name] = USL;
                    result.Rows[2][col_name] = LSL;
                    myCode.testFAI_spec.Add(new TDMK_OK2SHIP.FAI_Spec(col_name, "NA", setval, USL, LSL, "NA", "VHX"));
                }
            }
            return result;
        }
        public bool Load_PTH_Diameter_Spec(string _ItemCode, SqlConnection sqlcon, string DB_table_name, ref DataGridView tar_DGV)
        {
            DataTable result = Load_PTH_Diameter_Spec(_ItemCode, sqlcon, DB_table_name);
            if (result.Rows.Count == 0)
            {
                return false;
            }
            else
            {
                tar_DGV.DataSource = result;
                List<string> header_list = new List<string>() { "Normdim", "USL", "LSL" };
                for (int i = 0; i < 3; i++)
                {
                    tar_DGV.Rows[i].HeaderCell.Value = header_list[i];
                }
                tar_DGV.AutoResizeColumns();
                tar_DGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                return true;
            }
        }
        public DataTable PTH_Diameter_Load_Spec_fromFile(SqlConnection sqlcon, string format_loc, string tar_ItemCode, List<string> extensions)
        {
            DataTable spec_dt = new DataTable();
            string tar_format_file;
            List<string> file_format_lst = new List<string>();
            DirectoryInfo directory = new DirectoryInfo(format_loc);
            var files = extensions.SelectMany(e => directory.EnumerateFiles(e, SearchOption.AllDirectories)).Where(x => x.FullName.Contains(tar_ItemCode));
            foreach (var item in files)
            {
                file_format_lst.Add(item.FullName);
            }
            if (file_format_lst.Count > 0)
            {
                tar_format_file = file_format_lst[0];
                FileInfo excel_file = new FileInfo(tar_format_file);
                using (ExcelPackage myexcel = new ExcelPackage(excel_file))
                {
                    ExcelWorkbook tar_wkbook = myexcel.Workbook;
                    List<char> rejected_char_lst = new List<char> { ' ', '-', '&', '_', '\r', '\n' };
                    List<string> sht_keys = new List<string> { "BVH-PTH" };
                    int wrksheet_num = tar_wkbook.Worksheets.Count;
                    foreach (ExcelWorksheet tg in tar_wkbook.Worksheets)
                    {
                        string sht_name = TDMK_EPPLUS7.remove_special_chars(tg.Name, rejected_char_lst).ToUpper();
                        if (sht_keys.Any(x => sht_name.Contains(TDMK_EPPLUS7.remove_special_chars(x, rejected_char_lst))))
                        {
                            string spec_location_address = TDMK_EPPLUS7.Find_Cell_Location("Diameter seen from top", tg, true);
                            if (spec_location_address != "")
                            {
                                var cell_info = TDMK_EPPLUS7.Get_Cells_Info(tg, tg.Cells[spec_location_address]);
                                int row_offset = cell_info.row_qty;
                                int col_offset = cell_info.col_qty;
                                string search_item_address = tg.Cells[spec_location_address].Offset(row_offset, -col_offset).Address;
                                string normdim_addr = TDMK_EPPLUS7.Find_Cell_Addr("Nominal Dim.", search_item_address, tg, false);
                                string tol_max_addr = TDMK_EPPLUS7.Find_Cell_Addr("Tol. Max. (+)", search_item_address, tg, false);
                                string tol_min_addr = TDMK_EPPLUS7.Find_Cell_Addr("Tol. Min. (-)", search_item_address, tg, false);
                                string instrument_addr = TDMK_EPPLUS7.Find_Cell_Addr("instrument", search_item_address, tg, false);
                                string USL_addr = TDMK_EPPLUS7.Find_Cell_Addr("USL", search_item_address, tg, false);
                                string LSL_addr = TDMK_EPPLUS7.Find_Cell_Addr("LSL", search_item_address, tg, false);
                                try
                                {
                                    ExcelRangeBase sel_rgn = tg.Cells[normdim_addr].Offset(0, col_offset);// tg.Range["D19"];    //tg.Range["C17"];
                                    ExcelRangeBase dev_rgn = tg.Cells[instrument_addr].Offset(0, col_offset);// tg.Range["D23"];    //tg.Range["C21"]
                                    int sel_inx = 0;
                                    while (myCode.checkDBNull(sel_rgn.Offset(0, sel_inx).Value) != "")
                                    {
                                        double USL = 0;
                                        double LSL = 0;
                                        string t_FAIName = "PTH-Dimension";
                                        string t_FAI_Setval = myCode.checkDBNull(sel_rgn.Offset(0, sel_inx).Value);
                                        string t_FAI_UL = myCode.checkDBNull(tg.Cells[USL_addr].Offset(0, col_offset + sel_inx).Value); ;// myCode.checkDBNull(sel_rgn.Offset[2, sel_inx].Value);
                                        string t_FAI_LL = myCode.checkDBNull(tg.Cells[LSL_addr].Offset(0, col_offset + sel_inx).Value); ;// myCode.checkDBNull(sel_rgn.Offset[3, sel_inx].Value);
                                        string t_FAI_Tol_plus = myCode.checkDBNull(tg.Cells[tol_max_addr].Offset(0, col_offset + sel_inx).Value);
                                        string t_FAI_Tol_minus = myCode.checkDBNull(tg.Cells[tol_min_addr].Offset(0, col_offset + sel_inx).Value);
                                        if ((t_FAI_Setval != "NA") && (t_FAI_Setval != ""))
                                        {
                                            double sv = Convert.ToDouble(t_FAI_Setval);
                                            if (myCode.IsNumeric(t_FAI_Tol_plus))
                                            {
                                                USL = sv + Convert.ToDouble(t_FAI_Tol_plus);
                                                t_FAI_UL = USL.ToString();
                                            }
                                            else
                                            {
                                                t_FAI_UL = t_FAI_Tol_plus;
                                            }
                                            if (myCode.IsNumeric(t_FAI_Tol_minus))
                                            {
                                                LSL = sv - Convert.ToDouble(t_FAI_Tol_minus);
                                                t_FAI_LL = LSL.ToString();
                                            }
                                            else
                                            {
                                                t_FAI_LL = t_FAI_Tol_minus;
                                            }
                                        }
                                        else
                                        {
                                            t_FAI_LL = t_FAI_Tol_minus;
                                            t_FAI_UL = t_FAI_Tol_plus;
                                        }
                                        string col_name = t_FAIName + "_" + t_FAI_Setval;
                                        string t_instrument = myCode.checkDBNull(dev_rgn.Offset(0, sel_inx).Value);
                                        if (!myCode.check_columns_existed(spec_dt, col_name))
                                        {
                                            spec_dt.Columns.Add(col_name);
                                            if (spec_dt.Rows.Count == 0)
                                            {
                                                for (int i = 0; i < 3; i++)
                                                {
                                                    spec_dt.Rows.Add();
                                                }
                                            }
                                            spec_dt.Rows[0][col_name] = t_FAI_Setval;
                                            spec_dt.Rows[1][col_name] = t_FAI_UL;
                                            spec_dt.Rows[2][col_name] = t_FAI_LL;
                                        }
                                        sel_inx++;
                                    }
                                }
                                catch
                                {

                                }

                            }
                        }
                    }
                }
            }
            else
            {
                //MessageBox.Show(new Form { TopMost = true }, "File format not found");
            }
            return spec_dt;
        }
        public void BatchBulkCopy(SqlConnection sqlcon_OK2SHIP, DataTable dataTable, string DestinationTbl)
        {
            DataTable dtInsertRows = dataTable;
            using (SqlBulkCopy sbc = new SqlBulkCopy(sqlcon_OK2SHIP))
            {
                if (sqlcon_OK2SHIP.State != ConnectionState.Open)
                {
                    sqlcon_OK2SHIP.Open();
                }
                sbc.DestinationTableName = DestinationTbl;
                foreach (DataColumn dc in dataTable.Columns)
                {
                    sbc.ColumnMappings.Add(dc.ColumnName, dc.ColumnName);
                }
                sbc.WriteToServer(dtInsertRows);
                sqlcon_OK2SHIP.Close();
            }
        }
    }
}
