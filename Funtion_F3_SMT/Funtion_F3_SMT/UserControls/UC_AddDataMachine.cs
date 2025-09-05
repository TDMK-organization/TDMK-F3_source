using OK2SHIP_SMT.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.UserControls
{
    public partial class UC_AddDataMachine : UserControl
    {
        public DataTable DATATABLE = new DataTable();
        CustomDataGridView dgv_LogData = new CustomDataGridView();
        public Dictionary<string, DataTable> DICTIONARY = new Dictionary<string, DataTable>();
        public UC_AddDataMachine()
        {
            InitializeComponent();
            setUp();
        }
        public Dictionary<string, DataTable> GetData()
        {
            Dictionary<string, DataTable> dic = DICTIONARY;

            foreach (DataRow row in DATATABLE.Rows)
            {
                string machine = row["Machine"].ToString().Replace(" ", "");
                if (dic.TryGetValue(machine, out DataTable dataTable))
                {

                }
                else
                {
                    dataTable = MakeDataTableResult();
                    dic.Add(machine, dataTable);
                }
                SetValue(row, dataTable);
            }
            return dic;
        }
        private void SetValue(DataRow row, DataTable dataTable)
        {
            int prime = row["Side"].ToString().Contains("ACF") ? 0 : 1;
            prime += row["WettingAngle"].ToString().Contains("After") ? 0 : 2;
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                if (!new[] { "WettingAngle", "Side", "JUDGE" }.Contains(dataTable.Columns[i].ColumnName))
                {
                    if (double.TryParse(row[i + 1].ToString(), out double z))
                    {

                    }
                    dataTable.Rows[prime][i] = z;
                }
                else
                {
                    dataTable.Rows[prime][i] = row[i + 1];
                }
            }
        }
        private DataTable MakeDataTableResult()
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("WettingAngle");
            dataTable.Columns.Add("Side");
            dataTable.Columns.Add("Spec", typeof(double));
            dataTable.Columns.Add("JUDGE");
            for (int i = 1; i <= 32; i++)
            {
                dataTable.Columns.Add($"Sample {i}", typeof(double));
            }
            DataRow row = dataTable.NewRow();
            row["WettingAngle"] = "After plasma cleaning";
            row["Side"] = "ACF";
            dataTable.Rows.Add(row);
            DataRow row1 = dataTable.NewRow();
            row1["WettingAngle"] = "After plasma cleaning";
            row1["Side"] = "GND";
            dataTable.Rows.Add(row1);
            DataRow row2 = dataTable.NewRow();
            row2["WettingAngle"] = "Before packing";
            row2["Side"] = "ACF";
            dataTable.Rows.Add(row2);
            DataRow row3 = dataTable.NewRow();
            row3["WettingAngle"] = "Before packing";
            row3["Side"] = "GND";
            dataTable.Rows.Add(row3);
            return dataTable;
        }
        private DataTable MakeDataTable()
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Machine");
            dataTable.Columns.Add("WettingAngle");
            dataTable.Columns.Add("Side");
            dataTable.Columns.Add("Spec");
            dataTable.Columns.Add("Judge");
            for (int i = 1; i <= 32; i++)
            {
                dataTable.Columns.Add($"Sample {i}");
            }
            return dataTable;
        }
        private void setUp()
        {
            TLP.Controls.Add(dgv_LogData, 0, 0);
            DATATABLE = MakeDataTable();
            Dictionary<string, string[]> dicz = new Dictionary<string, string[]>();
            dicz.Add("Machine", new[] { "1", "2", "3", "4", "5", "6" });
            dicz.Add("WettingAngle", new[] { "After plasma cleaning", "Before packing" });
            dicz.Add("Side", new[] { "ACF", "GND" });
            dgv_LogData.columnDropdowns = dicz;
            dgv_LogData.DataSource = DATATABLE;
            dgv_LogData.SetUpWidth(new[] { 50, 200, 50, 50 });
            dgv_LogData.CellLeave += dataGridView_CellLeave;
        }
        private void dataGridView_CellLeave(object sender, DataGridViewCellEventArgs e)
        {

            //Debugger.Break();
            int row = e.RowIndex;
            int col = e.ColumnIndex;
            if (dgv_LogData.Columns[col].Name.Contains("Spec"))
            {
                string z = dgv_LogData.Rows[row].Cells[col].Value.ToString();
                if (double.TryParse(z, out double a))
                {
                    for (int i = 5; i < dgv_LogData.Columns.Count; i++)
                    {
                        var cel = dgv_LogData.Rows[row].Cells[i];
                        if (double.TryParse(cel.Value.ToString(), out double value))
                        {
                            if (a < value)
                            {
                                cel.Style.BackColor = Color.Red;
                                cel.Style.ForeColor = Color.White;
                            }
                            else
                            {
                                cel.Style.BackColor = Color.White;
                                cel.Style.ForeColor = Color.Black;
                            }
                        }
                    }
                }
            }
        }
        private void btn_upload_Click(object sender, EventArgs e)
        {
            DataTable dataTable = (DataTable)dgv_LogData.DataSource;
            dgv_LogData.DataSource = ACFService.ReadWettingContactAngle(dataTable, content.Text);
            content.Text = "";
        }

        private void tdmK_Label10_TextChanged(object sender, EventArgs e)
        {

        }
        private void updateDataRow()
        {


        }
        private void dgv_LogData_DataSourceChanged(object sender, EventArgs e)
        {
            updateDataRow();
        }


        private void tdmK_Button1_Click(object sender, EventArgs e)
        {


        }

        private void button1_Click(object sender, EventArgs e)
        {
            DICTIONARY = GetData();
        }
    }
}
