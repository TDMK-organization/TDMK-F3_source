using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SQL;
using OK2SHIP_Lib;

namespace OK2SHIP_Measurements
{
    public class IPQC_LogFile
    {
        //SEI_Lib myCode = new SEI_Lib();
        TDMK_OK2SHIP myCode = new TDMK_OK2SHIP();
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public Dictionary<int, string> Get_LogFile_Data(string in_src_file)
        {
            Dictionary<int, string> result_lst = new Dictionary<int, string>();
            string[] Lines = File.ReadAllLines(in_src_file);
            foreach (var line in Lines)
            {

                //string[] temp = line.Select(s => s != '"').ToString().Split(',');
                string str = new string(line.Where(s => s != '"').ToArray());
                string[] temp = str.Split(',');
                if (temp[0] != "")
                {
                    if (myCode.IsNumeric(temp[0]))
                    {
                        result_lst.Add(Convert.ToInt32(temp[0]), temp[2]);
                    }
                }
                else
                {
                    break;
                }
            }
            return result_lst;
        }
        public string Get_Number_String(string src_str, char split_chr)
        {
            string _result = "";
            string[] temp = src_str.Split(split_chr);
            foreach (var t in temp)
            {
                if (myCode.IsNumeric(t))
                {
                    _result = t;
                    break;
                }
            }
            return _result;
        }
        public void Get_List_data2(int col_inx, DataTable myDt, string[] src_arr, ref List<string> src_lst_data, string tar_item, bool distinct_en)
        {
            if (col_inx < src_arr.Length - 1)
            {
                if (src_arr[col_inx + 1] != "")
                {
                    string[] sel_val = myDt.AsEnumerable().Select(x => x.Field<string>(src_arr[col_inx + 1])).Distinct().ToArray();
                    if (sel_val.Length != 0)
                    {
                        foreach (string sv in sel_val)
                        {
                            if (sv != null)
                            {
                                DataTable curTbl = myDt.AsEnumerable().Where(r => r.Field<string>(src_arr[col_inx + 1]) == sv).CopyToDataTable();
                                Get_List_data2(col_inx + 1, curTbl, src_arr, ref src_lst_data, tar_item, distinct_en);
                            }
                        }
                    }
                    else
                    {
                        if (distinct_en)
                        {
                            //src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Distinct().ToArray());
                            src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Where(r => r != null).Distinct().ToArray());
                        }
                        else
                        {
                            //src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).ToArray());
                            src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Where(r => r != null).ToArray());
                        }
                    }
                }
            }
            else
            {
                if (distinct_en)
                {
                    //src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Distinct().ToArray());
                    src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Where(r => r != null).Distinct().ToArray());
                }
                else
                {
                    //src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).ToArray());
                    src_lst_data.AddRange(myDt.AsEnumerable().Select(r => r.Field<string>(tar_item)).Where(r => r != null).ToArray());
                }
            }
        }
        public void Fill_LogFile_Data_old(string log_path, SqlConnection sqlcon, string ItemCode, string Process_name, DataGridView tar_DGV)
        {
            if (log_path != "")
            {
                DirectoryInfo tar_d = new DirectoryInfo(log_path);
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                foreach (var inter_type in inter_type_lst)
                {
                    DirectoryInfo sel_dir = new DirectoryInfo(inter_type.FullName);
                    DirectoryInfo[] inter_point = sel_dir.GetDirectories();
                    foreach (var p in inter_point)
                    {
                        string inter_p = p.Name + "_" + sel_dir.Name;
                        List<FileInfo> f_list = p.GetFiles("*.csv").ToList();
                        string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark", "Internal_Point" }, new string[] { ItemCode, Process_name, inter_p });
                        DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, "SpecList", filter_str);
                        List<string> Logfile_lst = new List<string>();
                        Get_List_data2(-1, src_dt, new string[] { "ItemCode", "Remark", "Internal_Point" }, ref Logfile_lst, "LogFile_Point", true);
                        Dictionary<int, Dictionary<int, string>> result_lst = new Dictionary<int, Dictionary<int, string>>();
                        for (int i = 0; i < f_list.Count; i++)
                        {
                            Dictionary<int, string> temp = Get_LogFile_Data(f_list[i].FullName);
                            string f_na = Path.GetFileNameWithoutExtension(f_list[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
                            int f_inx = Convert.ToInt32(f_na);
                            result_lst.Add(f_inx, temp);
                            if (tar_DGV.Rows.Count < i + 1)
                            {
                                tar_DGV.Rows.Add();
                                tar_DGV.Rows[i].HeaderCell.Value = (i + 1).ToString();
                            }
                        }
                        foreach (var _res in result_lst)
                        {
                            int r_inx = _res.Key;
                            Dictionary<int, string> cur_dic = _res.Value;
                            foreach (var sel in cur_dic)
                            {
                                if (TDMK_Code.check_exist_list_index(sel.Key.ToString(), Logfile_lst) != -1)
                                {
                                    string tar_col = src_dt.AsEnumerable().Where(x => x.Field<string>("LogFile_Point") == sel.Key.ToString()).Select(x => x.Field<string>("Col_Name")).ToList()[0];
                                    tar_DGV.Rows[r_inx - 1].Cells[tar_col].Value = sel.Value;
                                }
                            }
                        }

                    }
                }
            }
        }
        public void Get_logfile(string in_src, ref List<string> result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            FileInfo[] temp_lst = tar_d.GetFiles("*.csv");
            if (temp_lst.Length > 0)
            {
                foreach (var t in temp_lst)
                {
                    result.Add(t.FullName);
                }
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_logfile(inter_lst.FullName, ref result);
                    }
                }
            }
        }
        public void Get_logfile2(string in_src, ref Dictionary<int, Dictionary<int, string>> lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            FileInfo[] temp_lst = tar_d.GetFiles("*.csv");
            if (temp_lst.Length > 0)
            {
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    Dictionary<int, string> temp = Get_LogFile_Data(temp_lst[i].FullName);
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
                    int f_inx = Convert.ToInt32(f_na);
                    lst_result.Add(f_inx, temp);
                }
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_logfile2(inter_lst.FullName, ref lst_result);
                    }
                }
            }
        }
        public void Get_logfile_Multi(string in_src, ref Dictionary<string, Dictionary<int, Dictionary<int, string>>> dic_lst_result)
        {
            DirectoryInfo tar_d = new DirectoryInfo(in_src);
            FileInfo[] temp_lst = tar_d.GetFiles("*.csv")?.OrderBy(f => int.Parse(Path.GetFileNameWithoutExtension(f.Name)))?.ToArray();
            if (temp_lst.Length > 0)
            {
                Dictionary<int, Dictionary<int, string>> lst_result = new Dictionary<int, Dictionary<int, string>>();
                for (int i = 0; i < temp_lst.Length; i++)
                {
                    Dictionary<int, string> temp = Get_LogFile_Data(temp_lst[i].FullName);
                    string f_na = Path.GetFileNameWithoutExtension(temp_lst[i].Name).TrimEnd(new char[] { ',', '.', ' ' });
                    int f_inx = Convert.ToInt32(f_na);
                    lst_result.Add(f_inx, temp);
                }
                dic_lst_result.Add(tar_d.Name, lst_result);
            }
            else
            {
                DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
                if (inter_type_lst.Length > 0)
                {
                    foreach (var inter_lst in inter_type_lst)
                    {
                        Get_logfile_Multi(inter_lst.FullName, ref dic_lst_result);
                    }
                }
            }
        }
        public void Fill_LogFile_Data(string log_path, SqlConnection sqlcon, string ItemCode, string Process_name, DataGridView tar_DGV)
        {
            if (log_path != "")
            {
                DirectoryInfo tar_d = new DirectoryInfo(log_path);
                DirectoryInfo[] inter_type_lst;
                if (Process_name =="Etching_Process")
                {
                    inter_type_lst = tar_d.GetDirectories();
                }
                else
                {
                    inter_type_lst = new DirectoryInfo[] { tar_d };
                }
                foreach (var inter_type in inter_type_lst)
                {
                    DirectoryInfo sel_dir = new DirectoryInfo(inter_type.FullName);
                    Dictionary<string, Dictionary<int, Dictionary<int, string>>> dic_result = new Dictionary<string, Dictionary<int, Dictionary<int, string>>>();
                    Get_logfile_Multi(inter_type.FullName, ref dic_result);
                    foreach (var _result in dic_result)
                    {
                        string inter_p;
                        if (Process_name == "Etching_Process")
                        {
                            inter_p = _result.Key + "_" + sel_dir.Name;
                        }
                        else
                        {
                            inter_p = _result.Key;
                        }
                        string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark", "Internal_Point" }, new string[] { ItemCode, Process_name, inter_p });
                        DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, "SpecList", filter_str);
                        List<string> Logfile_lst = new List<string>();
                        Get_List_data2(-1, src_dt, new string[] { "ItemCode", "Remark", "Internal_Point" }, ref Logfile_lst, "LogFile_Point", true);
                        Dictionary<int, Dictionary<int, string>> result_lst = _result.Value;
                        foreach (var _res in result_lst)
                        {
                            int r_inx = _res.Key;
                            Dictionary<int, string> cur_dic = _res.Value;
                            foreach (var sel in cur_dic)
                            {
                                if (TDMK_Code.check_exist_list_index(sel.Key.ToString(), Logfile_lst) != -1)
                                {
                                    string tar_col = src_dt.AsEnumerable().Where(x => x.Field<string>("LogFile_Point") == sel.Key.ToString()).Select(x => x.Field<string>("Col_Name")).ToList()[0];
                                    try
                                    {
                                        tar_DGV.Rows[r_inx - 1].Cells[tar_col].Value = sel.Value;
                                    }
                                    catch 
                                    { 

                                    }
                                    
                                }
                            }
                        }
                    }
                }
            }
        }
        public void Fill_LogFile_Data_CheckSpec(string log_path, SqlConnection sqlcon, string ItemCode, string Process_name, DataGridView tar_DGV, DataGridView in_DGV_Spec)
        {
            if (log_path != "")
            {
                DirectoryInfo tar_d = new DirectoryInfo(log_path);
                DirectoryInfo[] inter_type_lst;
                if (Process_name == "Etching_Process")
                {
                    inter_type_lst = tar_d.GetDirectories();
                }
                else
                {
                    inter_type_lst = new DirectoryInfo[] { tar_d };
                }
                DataTable spec_tbl = myCode.DGV_To_Table(in_DGV_Spec);
                foreach (var inter_type in inter_type_lst)
                {
                    DirectoryInfo sel_dir = new DirectoryInfo(inter_type.FullName);
                    Dictionary<string, Dictionary<int, Dictionary<int, string>>> dic_result = new Dictionary<string, Dictionary<int, Dictionary<int, string>>>();
                    Get_logfile_Multi(inter_type.FullName, ref dic_result);
                    foreach (var _result in dic_result)
                    {
                        List<string> inter_p_lst = new List<string>();
                        char[] split_char = new char[] { '_', '+','-' };
                        if (Process_name == "Etching_Process")
                        {
                            string[] p_total = _result.Key.Split(split_char);
                            foreach(string p in p_total)
                            {
                                string _inter_p; ;
                                string region = sel_dir.Name.Trim();
                                if((region=="IL")||(region=="OL"))
                                {
                                    _inter_p = p.Trim() + "_" + region;
                                }
                                else
                                {
                                    _inter_p = p.Trim();
                                }
                                //string _inter_p = p.Trim() + "_" + sel_dir.Name.Trim();
                                inter_p_lst.Add(_inter_p);
                            }
                        }
                        else
                        {
                            inter_p_lst.Add(_result.Key.Trim());
                        }
                        foreach(string inter_p in inter_p_lst)
                        {
                            string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark", "Internal_Point" }, new string[] { ItemCode, Process_name, inter_p });
                            DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, "SpecList", filter_str);
                            List<string> Logfile_lst = new List<string>();
                            Get_List_data2(-1, src_dt, new string[] { "ItemCode", "Remark", "Internal_Point" }, ref Logfile_lst, "LogFile_Point", true);
                            if(Logfile_lst.Count==0)
                            {
                                filter_str = TDMK_Code.filter_str(new string[] { "Process_Name", "Internal_Point" }, new string[] { Process_name, inter_p });
                                src_dt = TDMK_Code.Datatable_Filter(sqlcon, "All_Items", filter_str);
                                Get_List_data2(-1, src_dt, new string[] { "Process_Name", "Internal_Point" }, ref Logfile_lst, "LogFile_Point", true);
                            }
                            Dictionary<int, Dictionary<int, string>> result_lst = _result.Value;
                            foreach (var _res in result_lst)
                            {
                                int r_inx = _res.Key;
                                Dictionary<int, string> cur_dic = _res.Value;
                                foreach (var sel in cur_dic)
                                {
                                    if (TDMK_Code.check_exist_list_index(sel.Key.ToString(), Logfile_lst) != -1)
                                    {
                                        string tar_col = src_dt.AsEnumerable().Where(x => x.Field<string>("LogFile_Point") == sel.Key.ToString()).Select(x => x.Field<string>("Col_Name")).ToList()[0];
                                        if (myCode.check_columns_existed(spec_tbl, tar_col))
                                        {
                                            string UL = myCode.checkDBNull(in_DGV_Spec.Rows[1].Cells[tar_col].Value);
                                            string LL = myCode.checkDBNull(in_DGV_Spec.Rows[2].Cells[tar_col].Value);
                                            string SV = myCode.checkDBNull(in_DGV_Spec.Rows[0].Cells[tar_col].Value);
                                            string act_val = sel.Value;
                                            tar_DGV.Rows[r_inx - 1].Cells[tar_col].Style.BackColor = myCode.check_in_limit_Color(UL, LL, act_val, SV);
                                            tar_DGV.Rows[r_inx - 1].Cells[tar_col].Value = sel.Value;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        public void Fill_LogFile_Data_CheckSpec_backup(string log_path, SqlConnection sqlcon, string ItemCode, string Process_name, DataGridView tar_DGV, DataGridView in_DGV_Spec) 
        {
            if (log_path != "")
            {
                DirectoryInfo tar_d = new DirectoryInfo(log_path);
                DirectoryInfo[] inter_type_lst;
                if (Process_name == "Etching_Process")
                {
                    inter_type_lst = tar_d.GetDirectories();
                }
                else
                {
                    inter_type_lst = new DirectoryInfo[] { tar_d };
                }
                DataTable spec_tbl = myCode.DGV_To_Table(in_DGV_Spec);
                foreach (var inter_type in inter_type_lst)
                {
                    DirectoryInfo sel_dir = new DirectoryInfo(inter_type.FullName);
                    Dictionary<string, Dictionary<int, Dictionary<int, string>>> dic_result = new Dictionary<string, Dictionary<int, Dictionary<int, string>>>();
                    Get_logfile_Multi(inter_type.FullName, ref dic_result);
                    foreach (var _result in dic_result)
                    {
                        string inter_p;
                        if (Process_name == "Etching_Process")
                        {
                            inter_p = _result.Key + "_" + sel_dir.Name;
                        }
                        else
                        {
                            inter_p = _result.Key;
                        }
                        string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "Remark", "Internal_Point" }, new string[] { ItemCode, Process_name, inter_p });
                        DataTable src_dt = TDMK_Code.Datatable_Filter(sqlcon, "SpecList", filter_str);
                        List<string> Logfile_lst = new List<string>();
                        Get_List_data2(-1, src_dt, new string[] { "ItemCode", "Remark", "Internal_Point" }, ref Logfile_lst, "LogFile_Point", true);
                        Dictionary<int, Dictionary<int, string>> result_lst = _result.Value;
                        foreach (var _res in result_lst)
                        {
                            int r_inx = _res.Key;
                            Dictionary<int, string> cur_dic = _res.Value;
                            foreach (var sel in cur_dic)
                            {
                                if (TDMK_Code.check_exist_list_index(sel.Key.ToString(), Logfile_lst) != -1)
                                {
                                    string tar_col = src_dt.AsEnumerable().Where(x => x.Field<string>("LogFile_Point") == sel.Key.ToString()).Select(x => x.Field<string>("Col_Name")).ToList()[0];
                                    if (myCode.check_columns_existed(spec_tbl, tar_col))
                                    {
                                        string UL = myCode.checkDBNull(in_DGV_Spec.Rows[1].Cells[tar_col].Value);
                                        string LL = myCode.checkDBNull(in_DGV_Spec.Rows[2].Cells[tar_col].Value);
                                        string SV = myCode.checkDBNull(in_DGV_Spec.Rows[0].Cells[tar_col].Value);
                                        string act_val = sel.Value;
                                        tar_DGV.Rows[r_inx - 1].Cells[tar_col].Style.BackColor = myCode.check_in_limit_Color(UL, LL, act_val, SV);
                                        tar_DGV.Rows[r_inx - 1].Cells[tar_col].Value = sel.Value;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
