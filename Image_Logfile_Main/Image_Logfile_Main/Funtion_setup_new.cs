using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OK2SHIP_SMT;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SEEV_DLL;
using TDMK_SQL;
using VHX;
using DataTable = System.Data.DataTable;

namespace OK2SHIP
{
    public class Funtion_setup_new
    {
        public EPPlus_Lib TDMK_Code2 = new EPPlus_Lib();
        public TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        SEI_Lib myCode = new SEI_Lib();
        myVar exp_proc = new myVar();

        public void setup_Impedance_new(string ItemCode, ExcelWorkbook wb, SqlConnection sqlcon)
        { 
            string sheet_name = "";
            foreach (ExcelWorksheet tg_sht in wb.Worksheets)
            {
                if (tg_sht.Name.ToUpper().Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "").Contains("IMPEDANCE"))
                {
                    sheet_name = tg_sht.Name;
                    break;
                }
            }
            if (sheet_name != "")
            { 
                ExcelWorksheet ws = wb.Worksheets[sheet_name]; 
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { ItemCode });
                DataTable dt_impedance_spec = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", filter_str);
                DataTable dt_tracewidth_spec = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", filter_str);
                int i = 0;
                int id_imp = TDMK_Code.SQL_MAX("IMPEDANCE_SPEC", "ID", sqlcon);
                AutoCompleteStringCollection lst_info_imp = Get_impedance_infor_to_input_value_2(ws);

                foreach (string item in lst_info_imp)
                {
                    DataRow dr = dt_impedance_spec.NewRow();
                    dr["ID"] = id_imp + i + 1;
                    dr["ItemCode"] = ItemCode;
                    dr["Region"] = "Impedance_" + (i + 1).ToString();
                    dr["Range"] = item.Split('+')[0];
                    dr["USL"] = item.Split('+')[1];
                    dr["LSL"] = item.Split('+')[2];
                    dt_impedance_spec.Rows.Add(dr);
                    i++;
                }
                i = 0;
                int id_trace = TDMK_Code.SQL_MAX("TRACEWIDTH_SPEC", "ID", sqlcon);

                AutoCompleteStringCollection lst_info_trw = Get_tracewidth_infor_to_input_value_2(ws);
                foreach (string item in lst_info_trw)
                {
                    DataRow dr = dt_tracewidth_spec.NewRow();
                    dr["ID"] = id_trace + i + 1;
                    dr["ItemCode"] = ItemCode;
                    dr["Region"] = item.Split('+')[3];
                    dr["Range"] = item.Split('+')[0];
                    dr["USL"] = item.Split('+')[1];
                    dr["LSL"] = item.Split('+')[2];
                    dt_tracewidth_spec.Rows.Add(dr);
                    i++;
                }
                exp_proc.BatchBulkCopy(sqlcon, dt_impedance_spec, "IMPEDANCE_SPEC");
                exp_proc.BatchBulkCopy(sqlcon, dt_tracewidth_spec, "TRACEWIDTH_SPEC");
                 
            }

        }

        public void setup_Stackup_new(ExcelWorkbook wb, string itemcode, SqlConnection sqlcon)
        {
            SortedDictionary<string, string> dic_spec_detail = new SortedDictionary<string, string>();
            string sheet_name = ""; 
            foreach (ExcelWorksheet tg_sht in wb.Worksheets)
            { 
                if (tg_sht.Name.ToUpper().Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "").Contains("STACKUP"))
                {
                    sheet_name = tg_sht.Name;
                    break;
                }
            }
            if (sheet_name != "")
            {
                string[] item = new string[] { "ID", "ItemCode", "Zone", "USL", "LSL", "Normal", "Spec_detail" };
                ExcelWorksheet ws = wb.Worksheets[sheet_name];
                int col_usl = 4;

                bool found = false;
                for (int row = 50; row <= 120; row++)
                {
                    for (int col = 3; col <= 7; col++)
                    {
                        ExcelRange cell = ws.Cells[row, col];
                        if (cell.Value != null && (cell.Value.ToString() == "USL"))
                        {
                            col_usl = col;
                            found = true;
                            break;
                        }
                    }
                    if (found)
                    {
                        break;
                    }
                }

                for (int i = 5; i < 400; i++)
                {
                    if (ws.Cells[i, 1].Value != null)
                    {
                        if (ws.Cells[i, 1].Value.ToString() == "Zone" && ws.Cells[i + 1, 1].Value != null)
                        {

                            string zone_ = ws.Cells[i + 1, 1].Value.ToString();
                            string lst_spec_detail = "";
                            string[] value = new string[7];

                            value[0] = (TDMK_Code.SQL_MAX("STACKUP_SPEC", "ID", sqlcon) + 1).ToString();
                            value[1] = itemcode;
                            value[2] = ws.Cells[i + 1, 1].Value.ToString();
                            for (int j = i + 1; j < i + 20; j++)
                            {
                                if (myCode.checkDBNull(ws.Cells[j, 6]) != "")
                                {
                                    if (myCode.checkDBNull(ws.Cells[j, 2].Value).ToUpper() == "ADHESIVE")
                                    {
                                        string usl = "";
                                        string lsl = "";
                                        string val_post = myCode.checkDBNull(ws.Cells[j, 6].Value);
                                        string val_pre = myCode.checkDBNull(ws.Cells[j, 5].Value);
                                        if (myCode.IsNumeric(val_pre))
                                        {
                                            usl = (Math.Round(double.Parse(val_pre) * 1.1, 2)).ToString();
                                        }
                                        if (myCode.IsNumeric(val_post))
                                        {
                                            lsl = (Math.Round(double.Parse(val_post) * 0.9)).ToString();
                                        }

                                        lst_spec_detail += usl + "/" + lsl + ";";

                                    }
                                    else
                                    {
                                        lst_spec_detail += myCode.checkDBNull(ws.Cells[j, 6].Value) + ";";
                                    }

                                }
                                if (myCode.checkDBNull(ws.Cells[j + 1, 2].Value).Replace(" ", string.Empty).ToUpper() == "Total thickness".Replace(" ", string.Empty).ToUpper())
                                {
                                    break;
                                }
                            }
                            if (!dic_spec_detail.ContainsKey(zone_))
                            {
                                dic_spec_detail.Add(zone_, lst_spec_detail);
                                value[6] = lst_spec_detail;
                            }

                            for (int j = i; j < i + 20; j++)
                            {
                                if (ws.Cells[j, col_usl].Value != null)
                                {
                                    if (ws.Cells[j, col_usl].Value.ToString() == "USL")
                                    {
                                        value[3] = ws.Cells[j, col_usl + 1].Value.ToString();
                                        value[4] = ws.Cells[j + 1, col_usl + 1].Value.ToString();

                                        if (ws.Cells[j - 2, 2].Value != null && ws.Cells[j - 2, 2].Value.ToString() != "")
                                        {
                                            if (ws.Cells[j - 2, 2].Value.ToString() == "Total thickness")
                                            {
                                                value[5] = ((ws.Cells[j - 2, 6].Value.ToString()).Split('+')[0]).Split('±')[0];
                                            }
                                        }
                                        if (ws.Cells[j - 3, 2].Value != null && ws.Cells[j - 3, 2].Value.ToString() != "")
                                        {
                                            if (ws.Cells[j - 3, 2].Value.ToString() == "Total thickness")
                                            {
                                                value[5] = ((ws.Cells[j - 3, 6].Value.ToString()).Split('+')[0]).Split('±')[0];
                                            }
                                        }
                                    }
                                }
                            }
                         //   TDMK_Code.Delelte_FilteredItem_arr("STACKUP_SPEC", sqlcon, item, value)
                            TDMK_Code.insert_val_arr("STACKUP_SPEC", sqlcon, item, value);
                        }
                    }
                }

                int count_sample = 0;
                for (int row = 2; row <= 100; row++)
                {

                    ExcelRange cell = wb.Worksheets[sheet_name].Cells[row, 1];
                    if (cell.Value != null && (cell.Value.ToString().ToUpper().Contains("ZONE")))
                    {
                        for (int col = 1; col < 20; col++)
                        {
                            string cell_val = myCode.checkDBNull(wb.Worksheets[sheet_name].Cells[row, col].Value);
                            if (cell_val.Contains("#"))
                            {
                                count_sample++;
                            }
                            else if (cell_val == "")
                            {
                                break;
                            }
                        }
                        break;
                    }
                }
                int ID = TDMK_Code.SQL_MAX("SETTING_PCS", "ID", sqlcon) + 1;
                TDMK_Code.insert_val_arr("SETTING_PCS", sqlcon, new string[] { "ID", "ItemCode", "Data_For", "Region_Zone", "Pcs_setting" }, new string[] { ID.ToString(), itemcode, "Stack-up", "", count_sample.ToString() });
            }
        } 

        public AutoCompleteStringCollection Get_impedance_infor_to_input_value_2(ExcelWorksheet ws)
        {
            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 24; i < 400; i++)
            {
                if (ws.Cells[i, 5].Value != null)
                {
                    if (ws.Cells[i, 5].Value.ToString().Replace(" ", "").ToUpper() == "Actual Impedance".Replace(" ", "").ToUpper())
                    {
                        if (ws.Cells[i + 1, 3].Value != null && ws.Cells[i + 1, 4].Value != null)
                        {
                            string max_imp = ws.Cells[i + 1, 3].Value.ToString();
                            string min_imp = ws.Cells[i + 1, 4].Value.ToString();

                            int count = 1;
                            while (ws.Cells[i + count, 2].Value != null)
                            {
                                count++;
                            }


                            list.Add("E" + (i + 1).ToString() + "-" + (count - 1).ToString() + "+" + max_imp + "+" + min_imp);
                            i = i + count;
                        }

                    }

                }
            }
            return list;
        }

        public AutoCompleteStringCollection Get_tracewidth_infor_to_input_value_2(ExcelWorksheet ws)
        {
            AutoCompleteStringCollection list = new AutoCompleteStringCollection();
            for (int i = 24; i < 300; i++)
            {
                for (int j = 10; j < 15; j++)
                {

                    if (myCode.checkDBNull(ws.Cells[i, j].Value).Replace(" ", "").ToUpper().Contains("Actual Trace width".Replace(" ", "").ToUpper()))
                    {
                        if (ws.Cells[i + 1, j - 1].Value != null && ws.Cells[i + 1, j - 2].Value != null)
                        {
                            int count = 1;
                            while (ws.Cells[i + count, j - 1].Value != null)
                            {
                                count++;
                            }

                            string max_tracewidth = ws.Cells[i + 1, j - 2].Value.ToString();
                            string min_tracewidth = ws.Cells[i + 1, j - 1].Value.ToString();
                            string region = "";
                            if (myCode.checkDBNull(ws.Cells[i, j].Value).Contains("(") && myCode.checkDBNull(ws.Cells[i, j].Value).Contains(")"))
                                region = myCode.checkDBNull(ws.Cells[i, j].Value).Split('(')[1].Split(')')[0];

                            list.Add((i + 1).ToString() + ":" + j.ToString() + ":" + (count - 1).ToString() + "+" + max_tracewidth + "+" + min_tracewidth + "+" + region) ;
                             
                           
                            i = i + count;
                        }
                        break;

                    }

                }
            }
            return list;
        }

    }
}
