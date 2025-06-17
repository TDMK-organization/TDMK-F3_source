using OK2SHIP_SMT.Repositories;
using OK2SHIP_SMT.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.UserControls
{
    public partial class DatabaseTool : UserControl
    {
        public DatabaseTool()
        {
            InitializeComponent();
        }
        DataTable dataTable = new DataTable();
        private void button1_Click(object sender, EventArgs e)
        {
            DBContext dBContext = new DBContext();
            string sql = $"SELECT DB_NAME() AS CurrentDatabaseName, S.name AS SchemaName, T.name AS TableName, C.name AS ColumnName, C.column_id AS ColumnOrder, TYPE_NAME(C.user_type_id) AS DataType, C.max_length AS MaxLength, C.precision AS NumericPrecision, C.scale AS NumericScale, C.is_nullable AS IsNullable, ISNULL(DC.definition, '') AS DefaultValue, C.is_identity AS IsIdentity, C.is_computed AS IsComputed FROM OK2SHIP_SMT.sys.tables AS T INNER JOIN OK2SHIP_SMT.sys.schemas AS S ON T.schema_id = S.schema_id INNER JOIN OK2SHIP_SMT.sys.columns AS C ON T.object_id = C.object_id LEFT JOIN OK2SHIP_SMT.sys.default_constraints AS DC ON C.default_object_id = DC.object_id WHERE T.is_ms_shipped = 0 ORDER BY SchemaName, TableName, ColumnOrder;";
            dataTable = dBContext.ExecuteSqlToDataTable(sql);

            dataGridView1.DataSource = dataTable;
            textBox2.Text = ConverterService.DataTableToJson(dataTable);
        }

        private void button2_Click(object sender, EventArgs e)
        {

            DBContext dBContext = new DBContext();

            foreach (DataRow item in dataTable.Rows)
            {
                string tableName = item["TableName"].ToString();
                if (!dBContext.DoesTableExist(tableName))
                {
                    MessageBox.Show($"Bảng {tableName} chưa tồn tại!");
                }
                if(!dBContext.DoesColumnExist(tableName, item["ColumnName"].ToString()))
                {
                    MessageBox.Show($"Cột {item["ColumnName"]} trong bảng {tableName} chưa tồn tại!");
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = new DataTable();
            dataGridView1.DataSource = ConverterService.JsonToDataTable(textBox2.Text);
        }
    }
}
