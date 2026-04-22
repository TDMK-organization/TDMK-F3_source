using OK2SHIP_Lib;
//using OK2SHIP_Software;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SQL;

namespace OK2SHIP_Measurements
{
    public partial class FAI_Edit : Form
    {
        public string ItemCode { get; set; }
        public string LotNo { get; set; }
        public string type { get; set; }
        public string shift { get; set; }
        public string userid { get; set; }
        public string machine { get; set; }
        public bool FAI_mode { get; set; }
        TDMK_OK2SHIP myCode = new TDMK_OK2SHIP();
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        //public List<string> edit_fai_lst = new List<string>();
        public Dictionary<string, List<int>> dic_edit_fai = new Dictionary<string, List<int>>();
        public Dictionary<string, List<int>> dic_del_fai = new Dictionary<string, List<int>>();
        public FAI_Edit()
        {
            InitializeComponent();
        }
        public FAI_Edit(string itemcode, string lotno, string _type, string _shift, string _userid, string _machine, bool _FAI_mode=true)
        {
            InitializeComponent();
            ItemCode = itemcode;
            LotNo = lotno;
            type = _type;
            shift = _shift;
            userid = _userid;
            machine = _machine;
            FAI_mode = _FAI_mode;
        }
        private void txtDataFolder_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            FolderBrowserDialog folder_data = new FolderBrowserDialog();
            if (folder_data.ShowDialog() == DialogResult.OK)
            {
                txtDataFolder.Text = folder_data.SelectedPath;
            }
        }

        private void txtDataFolder_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                if (txtDataFolder.Text != "")
                {
                    string tar_loc = txtDataFolder.Text.Replace(Environment.NewLine, "");
                    DirectoryInfo di = new DirectoryInfo(tar_loc);
                    if (di.Exists)
                    {
                        FileInfo[] files;
                        files = di.GetFiles("*.csv");
                        List<string> _lstLogfile = files.Select(x => x.FullName).ToList();
                        DataTable result_dt = new DataTable();
                        foreach(var f_log in _lstLogfile)
                        {
                            result_dt = Logfile_data_process(f_log, ref result_dt);
                        }
                        DGV_LogFile.DataSource = result_dt;
                        myCode.check_FAIdata_inSpec2(DGV_LogFile, DGV_Spec);
                        myCode.DGV_Auto_Resize(DGV_LogFile);
                    }
                    else
                    {
                        MessageBox.Show("Folder not existed!", "Warning");
                    }
                }
            }
        }
        private void FAI_Edit_Load(object sender, EventArgs e)
        {
            if(FAI_mode)
            {
                if (myCode.Load_Spec(myVar.sqlcon_SMT, "", ItemCode, LotNo, ".xlsm", DGV_Spec, "", "NPI"))
                {
                    myCode.Load_FAI_DGV(DGV_Data, myVar.sqlcon_SMT, "FAI_Auto", "FAI_No", "FAI_Data", ItemCode, LotNo, "",false);
                    myCode.Check_Data_inSpec2(DGV_Data, DGV_Spec);
                    myCode.Load_CPK_DGV_Histogram(DGV_Data, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
                    myCode.DGV_Auto_Resize(DGV_Data);
                }
                else
                {
                    MessageBox.Show("Không tìm thấy Spec", "Thông báo");
                }
            }
            else
            {
                if (myCode.Load_PTH_Diameter_Spec(ItemCode, myVar.sqlcon_SMT, "SPEC_PTH_Diameter", ref DGV_Spec))
                {
                    Load_PTH_Diameter_DGV(DGV_Data, myVar.sqlcon_SMT, "PTH_Diameter", ItemCode, LotNo);
                    myCode.Load_CPK_DGV_Histogram(DGV_Data, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
                    myCode.Check_Data_inSpec2(DGV_Data, DGV_Spec);
                    for (int r = 0; r < DGV_Data.Rows.Count; r++)
                    {
                        DGV_Data.Rows[r].HeaderCell.Value = (r + 1).ToString();
                    }
                    DGV_Data.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                }
            }

        }
        private void DGV_Data_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                List<string> col_list = new List<string>();
                foreach (DataGridViewCell c in DGV_Data.SelectedCells)
                {
                    int k = c.ColumnIndex;
                    if (!col_list.Contains(DGV_Data.Columns[k].Name))
                    {
                        col_list.Add(DGV_Data.Columns[k].Name);
                    }
                }

                DataTable test_tbl = myCode.DGV_To_Table(DGV_Data);
                int arr_num = col_list.Count;
                TDMK_OK2SHIP.Calculate_CPK[] myCalc_CPK = new TDMK_OK2SHIP.Calculate_CPK[arr_num];
                TDMK_OK2SHIP.FAI_Histogram_Data[] myHistogram_data = new TDMK_OK2SHIP.FAI_Histogram_Data[arr_num];
                int col_list_inx = 0;
                foreach (string t in col_list)
                {
                    TDMK_OK2SHIP.FAI_Spec sel_FAI_test;
                    foreach (TDMK_OK2SHIP.FAI_Spec ref_spec in myCode.testFAI_spec)
                    {
                        if (ref_spec.FAI_Name == t)
                        {
                            sel_FAI_test = ref_spec;
                            string[] temp_FAI_data = test_tbl.AsEnumerable().Select(r => r.Field<string>(t)).ToArray();
                            Double[] data_arr = new double[temp_FAI_data.Length];
                            int inx = 0;
                            foreach (string c in temp_FAI_data)
                            {
                                if (myCode.checkDBNull(c) != "")
                                {
                                    data_arr[inx] = Convert.ToDouble(c);
                                    inx++;
                                }
                            }
                            Array.Resize(ref data_arr, inx);
                            if (inx > 0)
                            {
                                myCode.Calcul_CPK_FAI2(ref_spec, data_arr, ref myCalc_CPK[col_list_inx], ref myHistogram_data[col_list_inx]);
                                FrmHistogram chart_form = new FrmHistogram();
                                chart_form.src_bin_data = myHistogram_data[col_list_inx]._Bin_data;// bin_data;
                                chart_form.src_freq_bin_data = myHistogram_data[col_list_inx]._Freq_bin_data;//freq_bin_data;
                                chart_form.src_FAI_Data = myHistogram_data[col_list_inx]._FAI_Data;
                                chart_form.FAI_Spec_val = ref_spec;
                                chart_form.src_CPK_result = myCalc_CPK[col_list_inx];
                                chart_form.src_modified_NormDist = myHistogram_data[col_list_inx]._Modified_NormDist_data; //modified_NormDist;
                                chart_form.Show();
                            }
                            break;
                        }
                    }
                    col_list_inx++;
                }
            }
            catch
            {

            }
        }
        public DataTable Logfile_data_process(string src_logfile, ref DataTable logfile_tbl)
        {
            DataTable FAI_logfile_tbl = new DataTable();
            if(FAI_mode)
            {
                DataTable logfile_dt = myCode.Mitutoyo_Get_raw_data2(src_logfile);
                FAI_logfile_tbl=myCode.FAI_raw_tbl(logfile_dt);
            }
            else
            {
                FAI_logfile_tbl = myCode.PTH_Diameter_Logfile(src_logfile);
            }
            DataTable tbl_spec = (DataTable)DGV_Spec.DataSource;
            DataTable FAI_tbl = new DataTable();
            DataTable non_FAI = new DataTable();
            DataTable FAI_arranged_tbl = myCode. filter_table(FAI_logfile_tbl, tbl_spec, ref FAI_tbl, ref non_FAI);
            DataTable dest_FAI_tbl = myCode.result_add(logfile_tbl, FAI_arranged_tbl, true);
            return dest_FAI_tbl;
        }

        private void DGV_LogFile_MouseClick(object sender, MouseEventArgs e)
        {
            if ((e.Button == MouseButtons.Right) && DGV_LogFile.Columns.Count > 0)
            {
                cmsAction.Show(DGV_LogFile, e.Location);
            }
        }

        private void tsmInsert_Click(object sender, EventArgs e)
        {
            if (insert_FAI_data_DGV(DGV_LogFile, DGV_Data, myCode.DGV_To_Table(DGV_Spec)))
            {
                myCode.check_FAIdata_inSpec2(DGV_Data, DGV_Spec);
                myCode.Load_CPK_DGV_Histogram(DGV_Data, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
                myCode.DGV_Auto_Resize(DGV_Data);
            }
        }

        private void tsmRemove_Click(object sender, EventArgs e)
        {
            if (DGV_LogFile.SelectedRows.Count > 0)
            {
                foreach (DataGridViewRow dgv_r in DGV_LogFile.SelectedRows)
                {
                    DGV_LogFile.Rows.Remove(dgv_r);
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Please, select rows to remove");
            }
        }

        private void tsmReplace_Click(object sender, EventArgs e)
        {
            if (DGV_Data.SelectedRows.Count > 0)
            {
                int sel_row_inx = DGV_Data.SelectedRows[0].Index;
                foreach (DataGridViewCell c in DGV_LogFile.SelectedCells)
                {
                    int src_col_inx = c.ColumnIndex;
                    string src_col_name = DGV_LogFile.Columns[src_col_inx].Name;
                    try
                    {
                        if (myCode.check_columns_existed(myCode.DGV_To_Table(DGV_Spec), src_col_name))
                        {
                            string SetVal = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_Spec.Rows[0].Cells[src_col_name].Value));
                            string UL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_Spec.Rows[1].Cells[src_col_name].Value));
                            string LL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_Spec.Rows[2].Cells[src_col_name].Value));
                            string Act_val = myCode.checkDBNull(c.Value);
                            DGV_Data.Rows[sel_row_inx].Cells[src_col_name].Value = c.Value;
                            DGV_Data.Rows[sel_row_inx].Cells[src_col_name].Style.BackColor = myCode.check_in_limit_Color(UL, LL, Act_val, SetVal.ToString());
                            c.Value = "";
                            if(dic_edit_fai.ContainsKey(src_col_name))
                            {
                                dic_edit_fai[src_col_name].Add(sel_row_inx);
                            }
                            else
                            {
                                dic_edit_fai.Add(src_col_name, new List<int> { sel_row_inx });
                            }
                        }
                    }
                    catch
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Sai vi tri", "Thong bao");
                    }
                }
                myCode.Load_CPK_DGV_Histogram(DGV_Data, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Chọn dòng");
            }
        }

        private void tsmSelectedColumn_Click(object sender, EventArgs e)
        {
            int sel_col_inx = DGV_LogFile.SelectedCells[0].ColumnIndex;
            int src_col_inx = DGV_Spec.SelectedCells[0].ColumnIndex;
            string src_col_name = DGV_Spec.Columns[src_col_inx].Name;
            string src_col_header = src_col_name.Split('_')[0];
            if (!myCode.check_columns_existed(myCode.DGV_To_Table(DGV_LogFile), src_col_name))
            {
                DGV_LogFile.Columns[sel_col_inx].HeaderText = src_col_name;
                DGV_LogFile.Columns[sel_col_inx].Name = src_col_name;
                double SetVal = Convert.ToDouble(DGV_Spec.Rows[0].Cells[src_col_name].Value);
                string UL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_Spec.Rows[1].Cells[src_col_name].Value));
                string LL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_Spec.Rows[2].Cells[src_col_name].Value));
                foreach (DataGridViewRow r in DGV_LogFile.Rows)
                {
                    string act_val = myCode. checkDBNull(r.Cells[src_col_name].Value);
                    r.Cells[src_col_name].Style.BackColor = myCode.check_in_limit_Color(UL, LL, act_val, SetVal.ToString());
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Cannot rename");
            }
        }

        private void tsmDefault_Click(object sender, EventArgs e)
        {
            int sel_col_inx = DGV_LogFile.SelectedCells[0].ColumnIndex;
            string default_name = "Default" + sel_col_inx.ToString();
            if (!myCode.check_columns_existed(myCode.DGV_To_Table(DGV_LogFile), default_name))
            {
                DGV_LogFile.Columns[sel_col_inx].HeaderText = default_name;
                DGV_LogFile.Columns[sel_col_inx].Name = default_name;
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Cannot rename");
            }
        }
        private void PasteClipboardValue(bool _transpose, DataGridView tar_DGV)
        {
            //Show Error if no cell is selected
            if (tar_DGV.SelectedCells.Count == 0)
            {
                MessageBox.Show(new Form { TopMost = true }, "Please select a cell", "Paste", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            //Get the satring Cell
            DataGridViewCell startCell = GetStartCell(tar_DGV);
            //Get the clipboard value in a dictionary
            Dictionary<int, Dictionary<int, string>> cbValue = ClipBoardValues(Clipboard.GetText());
            if (cbValue.Count > 0)
            {
                if (!_transpose)
                {
                    int iRowIndex = startCell.RowIndex;
                    foreach (int rowKey in cbValue.Keys)
                    {
                        int iColIndex = startCell.ColumnIndex;
                        foreach (int cellKey in cbValue[rowKey].Keys)
                        {
                            if (iColIndex <= tar_DGV.Columns.Count - 1 && iRowIndex <= tar_DGV.Rows.Count - 1)
                            {
                                DataGridViewCell cell = tar_DGV[iColIndex, iRowIndex];
                                if (cell.Selected)
                                {
                                    cell.Value = cbValue[rowKey][cellKey];
                                    if (myCode.checkDBNull(cell.Value) != "")
                                    {
                                        string curr_col_name = tar_DGV.Columns[iColIndex].Name.ToString();
                                        if (myCode.check_columns_existed(myCode.DGV_To_Table(DGV_Spec), curr_col_name))
                                        {
                                            string SetVal = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_Spec.Rows[0].Cells[curr_col_name].Value));
                                            string UL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_Spec.Rows[1].Cells[curr_col_name].Value));
                                            string LL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_Spec.Rows[2].Cells[curr_col_name].Value));
                                            string act_val = myCode.checkDBNull(cell.Value);
                                            cell.Style.BackColor = myCode.check_in_limit_Color(UL, LL, act_val, SetVal.ToString());
                                        }
                                    }

                                }

                            }
                            iColIndex++;
                        }
                        iRowIndex++;
                    }
                }
                else
                {
                    int iColIndex = startCell.ColumnIndex;
                    foreach (int rowKey in cbValue.Keys)
                    {
                        int iRowIndex = startCell.RowIndex;
                        foreach (int cellKey in cbValue[rowKey].Keys)
                        {
                            if (iColIndex <= tar_DGV.Columns.Count - 1 && iRowIndex <= tar_DGV.Rows.Count - 1)
                            {
                                DataGridViewCell cell = tar_DGV[iColIndex, iRowIndex];
                                if (cell.Selected)
                                {
                                    cell.Value = cbValue[rowKey][cellKey];
                                    if (myCode.checkDBNull(cell.Value) != "")
                                    {
                                        string curr_col_name = tar_DGV.Columns[iColIndex].Name.ToString();
                                        if (myCode.check_columns_existed(myCode.DGV_To_Table(DGV_Spec), curr_col_name))
                                        {

                                            string SetVal = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_Spec.Rows[0].Cells[curr_col_name].Value));
                                            string UL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_Spec.Rows[1].Cells[curr_col_name].Value));
                                            string LL = myCode.IsNumeric_Val(myCode.checkDBNull(DGV_Spec.Rows[2].Cells[curr_col_name].Value));

                                            string act_val = myCode.checkDBNull(cell.Value);
                                            cell.Style.BackColor = myCode.check_in_limit_Color(UL, LL, act_val, SetVal);
                                        }
                                    }

                                }
                            }
                            iRowIndex++;
                        }
                        iColIndex++;
                    }
                }
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "PLease, select data!", "Warning");
            }

        }
        private DataGridViewCell GetStartCell(DataGridView dgView)
        {
            //get the smallest row,column index
            if (dgView.SelectedCells.Count == 0)
                return null;

            int rowIndex = dgView.Rows.Count - 1;
            int colIndex = dgView.Columns.Count - 1;

            foreach (DataGridViewCell dgvCell in dgView.SelectedCells)
            {
                if (dgvCell.RowIndex < rowIndex)
                    rowIndex = dgvCell.RowIndex;
                if (dgvCell.ColumnIndex < colIndex)
                    colIndex = dgvCell.ColumnIndex;
            }

            return dgView[colIndex, rowIndex];
        }
        private Dictionary<int, Dictionary<int, string>> ClipBoardValues(string clipboardValue)
        {
            Dictionary<int, Dictionary<int, string>> copyValues = new Dictionary<int, Dictionary<int, string>>();

            String[] lines = clipboardValue.Split('\n');

            for (int i = 0; i <= lines.Length - 1; i++)
            {
                copyValues[i] = new Dictionary<int, string>();
                String[] lineContent = lines[i].Split('\t');

                //if an empty cell value copied, then set the dictionay with an empty string
                //else Set value to dictionary
                if (lineContent.Length == 0)
                    copyValues[i][0] = string.Empty;
                else
                {
                    for (int j = 0; j <= lineContent.Length - 1; j++)
                        copyValues[i][j] = lineContent[j];
                }
            }
            return copyValues;
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyToClipboard(DGV_LogFile);
        }
        private void CopyToClipboard(DataGridView tar_DGV)
        {
            //Copy to clipboard
            DataObject dataObj = tar_DGV.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void selectedRowsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PasteClipboardValue(false, DGV_LogFile);
        }

        private void selectedColumnsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            PasteClipboardValue(true, DGV_LogFile);
        }

        private void tsmReplaceAll_Click(object sender, EventArgs e)
        {

        }
        public DataTable DGV_To_Table(DataGridView src_DGV)
        {
            DataTable dt = new DataTable();
            int col_inx = src_DGV.Columns.Count;
            string[] col_name = new string[col_inx];
            for (int i = 0; i < col_inx; i++)
            {
                col_name[i] = src_DGV.Columns[i].Name;
                dt.Columns.Add(col_name[i]);
            }

            for (int i = 0; i < src_DGV.Rows.Count; i++)
            {
                DataRow dr = dt.NewRow();
                for (int j = 0; j < col_inx; j++)
                {
                    dr[j] = src_DGV.Rows[i].Cells[j].Value;
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        private void tsmReplaceAllData_Click(object sender, EventArgs e)
        {
            Dictionary<string, List<int>> sel_FAI_dic = new Dictionary<string, List<int>>();
            var sel_cells = DGV_LogFile.SelectedCells;
            foreach (DataGridViewCell sel_cell in sel_cells)
            {
                int col = sel_cell.ColumnIndex;
                int row = sel_cell.RowIndex;
                string fai_no = DGV_LogFile.Columns[col].Name;
                if (sel_FAI_dic.ContainsKey(fai_no))
                {
                    sel_FAI_dic[fai_no].Add(row);
                }
                else
                {
                    sel_FAI_dic.Add(fai_no, new List<int> { row });
                }
            }
            DataTable logfile_tbl = DGV_To_Table(DGV_LogFile);
            DataTable src_dt = myCode.DGV_To_Table(DGV_Data);
            foreach (var fai in sel_FAI_dic)
            {
                if (myCode.check_columns_existed(src_dt, fai.Key))
                {
                    if (!dic_edit_fai.ContainsKey(fai.Key))
                    {
                        dic_edit_fai.Add(fai.Key, new List<int>());
                    }
                    List<string> cur_fai_lst = logfile_tbl.AsEnumerable().Select(x => x.Field<string>(fai.Key)).Where(x => x != null).ToList();
                    for (int i = 0; i < cur_fai_lst.Count; i++)
                    {
                        if (src_dt.Rows.Count <= i)
                        {
                            src_dt.Rows.Add();
                        }
                        src_dt.Rows[i][fai.Key] = cur_fai_lst[i];
                        dic_edit_fai[fai.Key].Add(i);
                    }
                    if(cur_fai_lst.Count< src_dt.Rows.Count)
                    {
                        if (!dic_del_fai.ContainsKey(fai.Key))
                        {
                            dic_del_fai.Add(fai.Key, new List<int>());
                        }
                        for(int i=cur_fai_lst.Count;i<src_dt.Rows.Count;i++)
                        {
                            dic_del_fai[fai.Key].Add(i);
                            src_dt.Rows[i][fai.Key] = "";
                        }
                    }
                    DGV_LogFile.Columns.Remove(fai.Key);
                }
            }
            myCode.Check_Data_inSpec2(DGV_Data, DGV_Spec);
            myCode.Load_CPK_DGV_Histogram(DGV_Data, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
            myCode.DGV_Auto_Resize(DGV_Data);

        }
        private void tsmReplaceSelection_Click(object sender, EventArgs e)
        {
            Dictionary<string, List<int>> sel_FAI_dic = new Dictionary<string, List<int>>();
            var sel_cells = DGV_LogFile.SelectedCells;
            foreach (DataGridViewCell sel_cell in sel_cells)
            {
                int col = sel_cell.ColumnIndex;
                int row = sel_cell.RowIndex;
                string fai_no = DGV_LogFile.Columns[col].Name;
                if (sel_FAI_dic.ContainsKey(fai_no))
                {
                    sel_FAI_dic[fai_no].Add(row);
                }
                else
                {
                    sel_FAI_dic.Add(fai_no, new List<int> { row });
                }
            }
            DataTable logfile_tbl = DGV_To_Table(DGV_LogFile);
            DataTable src_dt = myCode.DGV_To_Table(DGV_Data);
            foreach (var fai in sel_FAI_dic)
            {
                if (myCode.check_columns_existed(src_dt, fai.Key))
                {
                    if(!dic_edit_fai.ContainsKey(fai.Key))
                    {
                        dic_edit_fai.Add(fai.Key, new List<int>());
                    }
                    foreach (var inx in fai.Value)
                    {
                        if (src_dt.Rows.Count > inx)
                        {
                            src_dt.Rows[inx][fai.Key] = logfile_tbl.Rows[inx][fai.Key];
                            DGV_LogFile.Rows[inx].Cells[fai.Key].Value = "";
                            dic_edit_fai[fai.Key].Add(inx);
                        }
                    }
                }
            }
            myCode.Check_Data_inSpec2(DGV_Data, DGV_Spec);
            myCode.Load_CPK_DGV_Histogram(DGV_Data, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
            for (int r = 0; r < DGV_Data.Rows.Count; r++)
            {
                DGV_Data.Rows[r].HeaderCell.Value = (r + 1).ToString();
            }
            DGV_Data.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        private void DGV_Data_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if( MessageBox.Show("Bạn muốn cập nhật lại dữ liệu?","Thông báo",MessageBoxButtons.YesNo)==DialogResult.Yes)
            {
                int id = TDMK_Code.SQL_MAX("FAI_Auto", "ID", myVar.sqlcon_SMT);
                string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Remark", "Shift" }, new string[] { ItemCode, LotNo, type, shift });
                DataTable src_dt = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "FAI_Auto", filter_str);
                DataView dv = src_dt.AsDataView();
                foreach(var fai in dic_edit_fai)
                {
                    string UL = DGV_Spec.Rows[1].Cells[fai.Key].Value.ToString();
                    string LL = DGV_Spec.Rows[2].Cells[fai.Key].Value.ToString();
                    string SV = DGV_Spec.Rows[0].Cells[fai.Key].Value.ToString();
                    dv.RowFilter = "FAI_No ='" + fai.Key + "'";
                    List<int> fai_val = fai.Value.Distinct().ToList();
                    fai_val.Sort();
                    foreach (int r_inx in fai_val)
                    {
                        if(myCode.check_in_limit(UL,LL, DGV_Data.Rows[r_inx].Cells[fai.Key].Value.ToString(),SV))
                        {
                            if (r_inx < dv.Count)
                            {
                                DataRow dr = dv[r_inx].Row;
                                int r_x = src_dt.Rows.IndexOf(dr);
                                if (r_x != -1)
                                {
                                    src_dt.Rows[r_x]["FAI_Data"] = DGV_Data.Rows[r_inx].Cells[fai.Key].Value;
                                }
                            }
                            else
                            {
                                src_dt.Rows.Add(id++, ItemCode, LotNo, userid, machine, DateTime.Now.ToString(), fai.Key, DGV_Data.Rows[r_inx].Cells[fai.Key].Value, type, shift);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Dữ liệu " + fai.Key + "/ pcs: "+ (r_inx+1).ToString() + "= " + DGV_Data.Rows[r_inx].Cells[fai.Key].Value.ToString() + " : NG --> Không được phép cập nhật", "Thông báo");
                        }
                    }
                }
                foreach (var fai in dic_del_fai)
                {
                    dv.RowFilter = "FAI_No ='" + fai.Key + "'";
                    List<DataRow> del_row = new List<DataRow>();
                    List<int> fai_val = fai.Value.Distinct().ToList();
                    foreach (int r_inx in fai_val)
                    {
                        DataRow dr = dv[r_inx].Row;
                        del_row.Add(dr);
                    }
                    foreach(var _dr in del_row)
                    {
                        src_dt.Rows.Remove(_dr);
                    }
                }
                TDMK_Code.Delelte_FilteredItem_arr("FAI_Auto", myVar.sqlcon_SMT, filter_str);
                myCode.BatchBulkCopy(myVar.sqlcon_SMT, src_dt, "FAI_Auto");
                dic_del_fai = new Dictionary<string, List<int>>();
                dic_edit_fai = new Dictionary<string, List<int>>();
                FAI_Edit_Load(null, null);
            }    
        }
        public bool insert_FAI_data_DGV(DataGridView src_DGV, DataGridView tar_DGV, DataTable src_tbl_spec)
        {
            bool _result = false;
            int sel_col_inx = src_DGV.SelectedCells[0].ColumnIndex;
            string src_col_name = src_DGV.Columns[sel_col_inx].Name;
            DataTable tar_tbl = myCode. DGV_To_Table(tar_DGV);
            if (myCode.check_columns_existed(src_tbl_spec, src_col_name))
            {
                if (!myCode.check_columns_existed(tar_tbl, src_col_name))
                {
                    tar_tbl.Columns.Add(src_col_name);
                    for (int i = 0; i < src_DGV.Rows.Count; i++)
                    {
                        if (tar_tbl.Rows.Count <= i)
                        {
                            tar_tbl.Rows.Add();
                        }
                        tar_tbl.Rows[i][src_col_name] = src_DGV.Rows[i].Cells[sel_col_inx].Value;
                    }
                    _result = true;
                }
                else
                {
                    if (MessageBox.Show(new Form { TopMost = true }, src_col_name + " đã có, bạn muốn thêm dữ liệu?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        DataView dv = tar_tbl.AsDataView();
                        dv.RowFilter = "[" + src_col_name + "] = null";
                        if (dv.Count > 0)
                        {
                            MessageBox.Show("Dữ liệu đã tồn tại","Thông báo");
                        }
                        else
                        {
                            if(!dic_edit_fai.ContainsKey(src_col_name))
                            {
                                dic_edit_fai.Add(src_col_name, new List<int>());
                            }
                            int cur_row = tar_tbl.AsEnumerable().Where(x => x.Field<string>(src_col_name) != null).Select(x => x.Field<string>(src_col_name)).ToList().Count;
                            for (int i = 0; i < src_DGV.Rows.Count; i++)
                            {
                                if (tar_tbl.Rows.Count <= i + cur_row)
                                {
                                    tar_tbl.Rows.Add();
                                }
                                tar_tbl.Rows[i + cur_row][src_col_name] = src_DGV.Rows[i].Cells[sel_col_inx].Value;
                                dic_edit_fai[src_col_name].Add(i + cur_row);
                            }
                        }
                        _result = true;
                    }

                }
                tar_DGV.DataSource = null;
                tar_DGV.DataSource = myCode.Order_table_inSpec(tar_tbl, src_tbl_spec);
                src_DGV.Columns.Remove(src_col_name);
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Chọn FAI phù hợp với Spec!", "Thông báo");
            }
            return _result;
        }

        private void DGV_Data_MouseClick(object sender, MouseEventArgs e)
        {
            if ((e.Button == MouseButtons.Right) && DGV_Data.Columns.Count > 0)
            {
                cmsDelete.Show(DGV_Data, e.Location);
            }
        }
        public Dictionary<string, List<int>> Get_SelectedCells_DGV(DataGridView tar_DGV)
        {
            Dictionary<string, List<int>> sel_FAI_dic = new Dictionary<string, List<int>>();
            var sel_cells = tar_DGV.SelectedCells;
            foreach (DataGridViewCell sel_cell in sel_cells)
            {
                int col = sel_cell.ColumnIndex;
                int row = sel_cell.RowIndex;
                string fai_no = tar_DGV.Columns[col].Name;
                if (sel_FAI_dic.ContainsKey(fai_no))
                {
                    sel_FAI_dic[fai_no].Add(row);
                }
                else
                {
                    sel_FAI_dic.Add(fai_no, new List<int> { row });
                }
            }
            return sel_FAI_dic;
        }

        private void tsmDelSelect_Click(object sender, EventArgs e)
        {
            Dictionary<string, List<int>> sel_FAI_dic = Get_SelectedCells_DGV(DGV_Data);
            if(sel_FAI_dic.Count!=0)
            {
                if (MessageBox.Show("Bạn muốn xóa dữ liệu đã chọn?", "Cảnh báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Remark", "Shift" }, new string[] { ItemCode, LotNo, type, shift });
                    DataTable src_dt = TDMK_Code.Datatable_Filter(myVar.sqlcon_SMT, "FAI_Auto", filter_str);
                    DataView dv = src_dt.AsDataView();
                    foreach (var fai in sel_FAI_dic)
                    {
                        dv.RowFilter = "FAI_No ='" + fai.Key + "'";
                        List<DataRow> del_row = new List<DataRow>();
                        List<int> fai_val = fai.Value.Distinct().ToList();
                        fai_val.Sort();
                        foreach (int r_inx in fai_val)
                        {
                            DataRow dr = dv[r_inx].Row;
                            del_row.Add(dr);
                        }
                        foreach (var _dr in del_row)
                        {
                            src_dt.Rows.Remove(_dr);
                        }
                    }
                    TDMK_Code.Delelte_FilteredItem_arr("FAI_Auto", myVar.sqlcon_SMT, filter_str);
                    myCode.BatchBulkCopy(myVar.sqlcon_SMT, src_dt, "FAI_Auto");
                    myCode.Load_FAI_DGV2(DGV_Data, myVar.sqlcon_SMT, "FAI_Auto", "FAI_No", "FAI_Data", ItemCode, LotNo, type, shift);
                    myCode.Check_Data_inSpec2(DGV_Data, DGV_Spec);
                    myCode.Load_CPK_DGV_Histogram(DGV_Data, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
                    myCode.DGV_Auto_Resize(DGV_Data);
                }
            }    
        }

        private void tsmRemoveFAI_Click(object sender, EventArgs e)
        {
            Dictionary<string, List<int>> sel_FAI_dic = Get_SelectedCells_DGV(DGV_Data);
            if (sel_FAI_dic.Count != 0)
            {
                if (MessageBox.Show("Bạn muốn xóa tất cả dữ liệu của FAI đã chọn?", "Cảnh báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    foreach (var fai in sel_FAI_dic)
                    {
                        string fai_no = fai.Key;
                        string del_filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Remark", "Shift", "FAI_No" }, new string[] { ItemCode, LotNo, type, shift, fai_no});
                        TDMK_Code.Delelte_FilteredItem_arr("FAI_Auto", myVar.sqlcon_SMT, del_filter_str);
                    }
                    myCode.Load_FAI_DGV2(DGV_Data, myVar.sqlcon_SMT, "FAI_Auto", "FAI_No", "FAI_Data", ItemCode, LotNo, type, shift);
                    myCode.Check_Data_inSpec2(DGV_Data, DGV_Spec);
                    myCode.Load_CPK_DGV_Histogram(DGV_Data, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
                    myCode.DGV_Auto_Resize(DGV_Data);
                }
            }
        }

        private void tsmSel_pcs_Click(object sender, EventArgs e)
        {
            Dictionary<string, List<int>> sel_FAI_dic = new Dictionary<string, List<int>>();
            var sel_cells = DGV_LogFile.SelectedCells;
            foreach (DataGridViewCell sel_cell in sel_cells)
            {
                int col = sel_cell.ColumnIndex;
                int row = sel_cell.RowIndex;
                string fai_no = DGV_LogFile.Columns[col].Name;
                if (sel_FAI_dic.ContainsKey(fai_no))
                {
                    sel_FAI_dic[fai_no].Add(row);
                }
                else
                {
                    sel_FAI_dic.Add(fai_no, new List<int> { row });
                }
            }
            Dictionary<string, List<int>> tar_FAI_dic = new Dictionary<string, List<int>>();
            var tar_cells = DGV_Data.SelectedCells;
            foreach (DataGridViewCell sel_cell in tar_cells)
            {
                int col = sel_cell.ColumnIndex;
                int row = sel_cell.RowIndex;
                string fai_no = DGV_Data.Columns[col].Name;
                if (tar_FAI_dic.ContainsKey(fai_no))
                {
                    tar_FAI_dic[fai_no].Add(row);
                }
                else
                {
                    tar_FAI_dic.Add(fai_no, new List<int> { row });
                }
            }

            DataTable logfile_tbl = DGV_To_Table(DGV_LogFile);
            DataTable src_dt = myCode.DGV_To_Table(DGV_Data);
            foreach (var fai in sel_FAI_dic)
            {
                List<int> fai_val_lst = fai.Value.ToList();
                fai_val_lst.Sort();
                string tar_col_name = fai.Key;
                if (tar_FAI_dic.Keys.ToList().IndexOf(tar_col_name) !=-1)
                {
                    if (!dic_edit_fai.ContainsKey(tar_col_name))
                    {
                        dic_edit_fai.Add(tar_col_name, new List<int>());
                    }
                    List<int> tar_fai_lst = tar_FAI_dic[tar_col_name];
                    tar_fai_lst.Sort();
                    for (int inx=0;inx< fai_val_lst.Count; inx++)
                    {
                        if(inx<tar_FAI_dic[tar_col_name].Count)
                        {
                            int log_inx = fai_val_lst[inx];
                            int r_inx = tar_fai_lst[inx];
                            src_dt.Rows[r_inx][tar_col_name] = logfile_tbl.Rows[log_inx][tar_col_name];
                            DGV_LogFile.Rows[log_inx].Cells[tar_col_name].Value = "";
                            dic_edit_fai[tar_col_name].Add(r_inx);
                        }
                    }
                }
            }
            myCode.Check_Data_inSpec2(DGV_Data, DGV_Spec);
            myCode.Load_CPK_DGV_Histogram(DGV_Data, DGV_CPK, myCode.testFAI_spec.ToArray(), DGV_Histogram);
            for (int r = 0; r < DGV_Data.Rows.Count; r++)
            {
                DGV_Data.Rows[r].HeaderCell.Value = (r + 1).ToString();
            }
            DGV_Data.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }
        public void Load_FAI_DGV_old(DataGridView tar_DGV, SqlConnection tar_sqlcon, string tar_tbl, string FAI_No_name, string FAI_Data_Col_name, string tar_ItemCode, string tar_LotNo, string format_type)
        {
            string[] flt_items = new string[5];
            string[] flt_item_vals = new string[5];
            AutoCompleteStringCollection FAI_lst;
            flt_items[0] = "ItemCode";
            flt_items[1] = "LotNo";
            flt_items[2] = "FAI_No";
            flt_items[3] = "SheetNo";
            flt_items[4] = "Remark";
            flt_item_vals[0] = tar_ItemCode;
            DataTable FAI_spec_dt = TDMK_Code.Datatable_Filter(tar_sqlcon, "FAI_Spec", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { tar_ItemCode, "NPI" }));
            List<string> search_key_lst = new List<string>() { "FAI", "Dimension" };
            List<string> SheetNo_lst = FAI_spec_dt.AsEnumerable().Where(x => search_key_lst.Any(y => x.Field<string>("SheetNo").Contains(y))).Select(x => x.Field<string>("SheetNo")).Distinct().ToList();
            DataTable tbl_data = new DataTable();
            foreach (string t_ShtNo in SheetNo_lst)
            {
                flt_item_vals[1] = "";
                flt_item_vals[2] = "";
                flt_item_vals[3] = t_ShtNo;
                flt_item_vals[4] = "NPI";
                FAI_lst = TDMK_Code.Load_Item_Filter_str(tar_sqlcon, "FAI_Spec", "FAI_No", TDMK_Code.filter_str(flt_items, flt_item_vals));
                foreach (string c in FAI_lst)
                {
                    AutoCompleteStringCollection FAI_vals;
                    flt_item_vals[1] = tar_LotNo;
                    flt_item_vals[2] = c;
                    flt_item_vals[3] = "";
                    flt_item_vals[4] = format_type;
                    FAI_vals = TDMK_Code.Load_Item_Filter_str(tar_sqlcon, tar_tbl, FAI_Data_Col_name, TDMK_Code.filter_str(flt_items, flt_item_vals));
                    if (!myCode.check_columns_existed(tbl_data, c))
                    {
                        tbl_data.Columns.Add(c);
                    }
                    int tbl_row_count = tbl_data.Rows.Count;
                    if (tbl_row_count < FAI_vals.Count)
                    {
                        for (int i = 0; i < FAI_vals.Count - tbl_row_count; i++)
                        {
                            tbl_data.Rows.Add();
                        }
                    }
                    int inx = 0;
                    foreach (string t in FAI_vals)
                    {
                        tbl_data.Rows[inx][c] = t;
                        inx++;
                    }
                }
            }
            tar_DGV.DataSource = tbl_data;
            foreach (DataGridViewColumn t in tar_DGV.Columns)
            {
                t.HeaderText = t.HeaderText.Split('_')[0];
            }
            myCode.Disable_Sort_DGV(tar_DGV);
        }
        public void Load_FAI_DGV(DataGridView tar_DGV, SqlConnection tar_sqlcon, string tar_tbl, string FAI_No_name, string FAI_Data_Col_name, string tar_ItemCode, string tar_LotNo, string format_type)
        {
            string[] flt_items = new string[5];
            string[] flt_item_vals = new string[5];
            //AutoCompleteStringCollection FAI_lst;
            flt_items[0] = "ItemCode";
            flt_items[1] = "LotNo";
            flt_items[2] = "FAI_No";
            flt_items[3] = "SheetNo";
            flt_items[4] = "Remark";
            flt_item_vals[0] = tar_ItemCode;
            DataTable FAI_spec_dt = TDMK_Code.Datatable_Filter(tar_sqlcon, "FAI_Spec", TDMK_Code.filter_str(new string[] { "ItemCode", "Remark" }, new string[] { tar_ItemCode, format_type }));
            DataTable FAI_data_dt = TDMK_Code.Datatable_Filter(tar_sqlcon, "FAI_Auto", TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo", "Remark" }, new string[] { tar_ItemCode, tar_LotNo, format_type }));
            List<string> temp_lst = FAI_spec_dt.AsEnumerable().Select(x => x.Field<string>("FAI_No")).Distinct().ToList();
            List<string> main_FAI = new List<string>();
            char[] split_chars = new char[] { '_', '/', '\\' };
            foreach (string t in temp_lst)
            {
                string setval = t.Split(split_chars).LastOrDefault().Trim();
                string fai_name = t.Split(split_chars).FirstOrDefault().Trim();
                string sel_FAI = fai_name + "_" + setval;
                if (main_FAI.IndexOf(sel_FAI) == -1)
                {
                    main_FAI.Add(sel_FAI);
                }
            }
            DataTable tbl_data = new DataTable();
            foreach (string c in main_FAI)
            {
                List<string> FAI_vals = FAI_data_dt.AsEnumerable().Where(x => x.Field<string>(FAI_No_name) == c).Select(x => x.Field<string>(FAI_Data_Col_name)).ToList();

                if (!myCode.check_columns_existed(tbl_data, c))
                {
                    tbl_data.Columns.Add(c);
                }
                int tbl_row_count = tbl_data.Rows.Count;
                if (tbl_row_count < FAI_vals.Count)
                {
                    for (int i = 0; i < FAI_vals.Count - tbl_row_count; i++)
                    {
                        tbl_data.Rows.Add();
                    }
                }
                int inx = 0;
                foreach (string t in FAI_vals)
                {
                    tbl_data.Rows[inx][c] = t;
                    inx++;
                }
            }
            tar_DGV.DataSource = tbl_data;
            foreach (DataGridViewColumn t in tar_DGV.Columns)
            {
                t.HeaderText = t.HeaderText.Split('_')[0];
            }
            myCode.Disable_Sort_DGV(tar_DGV);
        }
        public void Load_PTH_Diameter_DGV(DataGridView tar_DGV, SqlConnection tar_sqlcon, string tar_tbl, string tar_ItemCode, string tar_LotNo)
        {
            DataTable PTH_Diameter_dt = TDMK_Code.Datatable_Filter(tar_sqlcon, tar_tbl, TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { tar_ItemCode, tar_LotNo }));
            DataTable result = new DataTable();
            List<DataTable> normdim_dt_lst = new List<DataTable>();
            myCode.Get_ListTable(-1, PTH_Diameter_dt, new string[] { "Normdim" }, ref normdim_dt_lst, "Data");
            foreach (DataTable dt in normdim_dt_lst)
            {
                string setval = myCode.checkDBNull(dt.Rows[0]["Normdim"]);
                string col_name = "PTH-Diameter_" + setval;
                if (!result.Columns.Contains(col_name))
                {
                    result.Columns.Add(col_name);
                    if (result.Rows.Count < dt.Rows.Count)
                    {
                        int row_add = dt.Rows.Count - result.Rows.Count;
                        for (int i = 0; i < row_add; i++)
                        {
                            result.Rows.Add();
                        }
                    }
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        result.Rows[j][col_name] = dt.Rows[j]["Data"];
                    }
                }
            }
            tar_DGV.DataSource = result;
            foreach (DataGridViewColumn t in tar_DGV.Columns)
            {
                t.HeaderText = t.HeaderText.Split('_')[0];
            }
            myCode.Disable_Sort_DGV(tar_DGV);
        }
    }
}
