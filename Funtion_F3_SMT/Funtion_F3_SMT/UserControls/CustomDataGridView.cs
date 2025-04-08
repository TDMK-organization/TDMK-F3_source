using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.UserControls
{
    public partial class CustomDataGridView : DataGridView
    {
        #region Fields
        private Dictionary<string, string[]> columnDropdowns;
        private TextBox editingTextBox;
        private DataTable dtable;
        private readonly Dictionary<string, ContextMenuStrip> headerDropdownMenus = new Dictionary<string, ContextMenuStrip>();
        private const int ArrowWidth = 15;
        #endregion

        #region Constructor
        public CustomDataGridView(DataTable dataTable, Dictionary<string, string[]> dropdownItems)
        {
            columnDropdowns = dropdownItems ?? new Dictionary<string, string[]>();
            dtable = dataTable;
            InitializeComponents();
            SetupDataGridView();
            this.DataSource = dataTable;
        }
        #endregion

        #region Initialization
        private void InitializeComponents()
        {
            this.Dock = DockStyle.Fill;
            this.AllowUserToAddRows = false;
            this.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            this.DataSourceChanged += CustomDataGridView_DataSourceChanged;
            this.DataBindingComplete += CustomDataGridView_DataBindingComplete;
            this.CellPainting += CustomDataGridView_CellPainting;
            this.CellMouseDown += CustomDataGridView_CellMouseDown;
            this.DataError += CustomDataGridView_DataError;
            this.CellDoubleClick += CustomDataGridView_CellDoubleClick;
            this.CellEndEdit += CustomDataGridView_CellEndEdit;

            editingTextBox = new TextBox
            {
                Visible = false,
                BorderStyle = BorderStyle.FixedSingle
            };
            this.Controls.Add(editingTextBox);
        }

        private void SetupDataGridView()
        {
            this.EnableHeadersVisualStyles = false;
            this.ColumnHeadersDefaultCellStyle.BackColor = Color.LightBlue;
            this.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            this.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            this.ColumnHeadersHeight = 80;
        }
        #endregion

        #region Data Binding
        private void CustomDataGridView_DataSourceChanged(object sender, EventArgs e)
        {
            if (this.DataSource is DataTable dataTable)
            {
                InitializeColumnsFromDataTable(dtable);
                ReplaceColumns();
            }
        }

        private void CustomDataGridView_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            RestoreOldValues();
            //ReplaceColumns();
            foreach (DataGridViewRow row in this.Rows)
            {
                row.Height = 70;
            }
            this.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            SetupHeaderDropdownMenus();
        }

        private void InitializeColumnsFromDataTable(DataTable dataTable)
        {
            this.Columns.Clear();
            foreach (DataColumn column in dataTable.Columns)
            {
                this.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = column.ColumnName,
                    HeaderText = column.ColumnName,
                    DataPropertyName = column.ColumnName,
                    ValueType = column.DataType
                });
            }
        }

        private void ReplaceColumns()
        {
            var columnsToReplace = new List<(string Name, DataGridViewColumn NewColumn, int Index, object[] OldValues)>();

            foreach (DataGridViewColumn column in this.Columns)
            {
                int columnIndex = column.Index;
                object[] oldValues = new object[dtable.Rows.Count];
                for (int i = 0; i < dtable.Rows.Count; i++)
                {
                    oldValues[i] = dtable.Rows[i][column.Name];
                }

                if (columnDropdowns.ContainsKey(column.Name))
                {
                    DataGridViewComboBoxColumn comboColumn = new DataGridViewComboBoxColumn
                    {
                        HeaderText = column.HeaderText,
                        Name = column.Name,
                        FlatStyle = FlatStyle.Flat,
                        DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox,
                        DataPropertyName = column.DataPropertyName
                    };
                    comboColumn.Items.AddRange(columnDropdowns[column.Name]);
                    columnsToReplace.Add((column.Name, comboColumn, columnIndex, oldValues));
                }
                else if (column.ValueType == typeof(Image) || column.ValueType == typeof(byte[]))
                {
                    DataGridViewImageColumn imageColumn = new DataGridViewImageColumn
                    {
                        HeaderText = column.HeaderText,
                        Name = column.Name,
                        DataPropertyName = column.DataPropertyName,
                        ImageLayout = DataGridViewImageCellLayout.Zoom
                    };
                    columnsToReplace.Add((column.Name, imageColumn, columnIndex, oldValues));
                }
            }

            foreach (var columnInfo in columnsToReplace)
            {
                this.Columns.Remove(columnInfo.Name);
                this.Columns.Insert(columnInfo.Index, columnInfo.NewColumn);
            }

            this.Tag = columnsToReplace;
        }

        private void RestoreOldValues()
        {
            if (this.Tag is List<(string Name, DataGridViewColumn NewColumn, int Index, object[] OldValues)> columnsToReplace)
            {
                foreach (var columnInfo in columnsToReplace)
                {
                    for (int i = 0; i < this.Rows.Count && i < columnInfo.OldValues.Length; i++)
                    {
                        if (!this.Rows[i].IsNewRow && columnInfo.OldValues[i] != null)
                        {
                            if (columnInfo.NewColumn is DataGridViewComboBoxColumn comboColumn)
                            {
                                object oldValue = columnInfo.OldValues[i];
                                string oldValueStr = oldValue?.ToString();

                                if (!string.IsNullOrEmpty(oldValueStr) && !comboColumn.Items.Contains(oldValue))
                                {
                                    comboColumn.Items.Add(oldValue);
                                    string columnName = columnInfo.Name;
                                    var updatedItems = new List<string>(columnDropdowns[columnName]) { oldValueStr };
                                    columnDropdowns[columnName] = updatedItems.ToArray();
                                    UpdateComboBoxItems(columnName, oldValueStr);
                                    this.Rows[i].Cells[columnInfo.Name].Value = oldValue;
                                }

                            }
                            //else
                            //{
                            //    this.Rows[i].Cells[columnInfo.Name].Value = columnInfo.OldValues[i];
                            //}
                        }
                    }
                }
            }
        }
        #endregion

        #region Header Dropdown
        private void SetupHeaderDropdownMenus()
        {
            headerDropdownMenus.Clear();
            foreach (var pair in columnDropdowns)
            {
                string columnName = pair.Key;
                string[] items = pair.Value;

                if (!this.Columns.Contains(columnName))
                    continue;

                ContextMenuStrip dropDownMenu = new ContextMenuStrip();
                foreach (string item in items)
                {
                    ToolStripMenuItem menuItem = new ToolStripMenuItem(item);
                    int columnIndex = this.Columns[columnName].Index;
                    menuItem.Click += (s, e) => HeaderMenuItem_Click(columnIndex, item);
                    dropDownMenu.Items.Add(menuItem);
                }
                headerDropdownMenus[columnName] = dropDownMenu;
            }
        }

        private void HeaderMenuItem_Click(int columnIndex, string selectedValue)
        {
            for (int i = 0; i < this.Rows.Count; i++)
            {
                this.Rows[i].Cells[columnIndex].Value = selectedValue;
            }
        }

        private void CustomDataGridView_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex == -1 && e.ColumnIndex >= 0 && columnDropdowns.ContainsKey(this.Columns[e.ColumnIndex].Name))
            {
                e.PaintBackground(e.CellBounds, true);

                string headerText = this.Columns[e.ColumnIndex].HeaderText;
                TextFormatFlags flags = TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter | TextFormatFlags.WordBreak;
                Rectangle textRect = new Rectangle(e.CellBounds.X, e.CellBounds.Y, e.CellBounds.Width - ArrowWidth, e.CellBounds.Height);
                TextRenderer.DrawText(e.Graphics, headerText, this.ColumnHeadersDefaultCellStyle.Font, textRect, Color.Black, flags);

                const int FixedArrowWidth = 10;
                const int FixedArrowHeight = 6;
                int arrowMargin = 8;

                int arrowX = e.CellBounds.Right - FixedArrowWidth - arrowMargin;
                int arrowY = e.CellBounds.Y + (e.CellBounds.Height - FixedArrowHeight) / 2;

                Rectangle arrowRect = new Rectangle(arrowX, arrowY, FixedArrowWidth, FixedArrowHeight);
                using (Brush brush = new SolidBrush(Color.Black))
                {
                    Point[] arrowPoints = new Point[]
                    {
                        new Point(arrowRect.X, arrowRect.Y),
                        new Point(arrowRect.X + FixedArrowWidth, arrowRect.Y),
                        new Point(arrowRect.X + FixedArrowWidth / 2, arrowRect.Y + FixedArrowHeight)
                    };
                    e.Graphics.FillPolygon(brush, arrowPoints);
                }

                e.Handled = true;
            }
        }

        private void CustomDataGridView_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex == -1 && e.ColumnIndex >= 0 && columnDropdowns.ContainsKey(this.Columns[e.ColumnIndex].Name))
            {
                Rectangle cellRect = this.GetCellDisplayRectangle(e.ColumnIndex, -1, false);
                Rectangle arrowRect = new Rectangle(cellRect.Right - ArrowWidth, cellRect.Y, ArrowWidth, cellRect.Height);

                if (arrowRect.Contains(e.X + cellRect.X, e.Y + cellRect.Y))
                {
                    string columnName = this.Columns[e.ColumnIndex].Name;
                    if (headerDropdownMenus.ContainsKey(columnName))
                    {
                        headerDropdownMenus[columnName].Show(this, new Point(cellRect.Right, cellRect.Bottom));
                    }
                }
            }
        }
        #endregion

        #region Cell Editing
        private void CustomDataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 &&
                columnDropdowns.ContainsKey(this.Columns[e.ColumnIndex].Name))
            {
                Rectangle cellRect = this.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, false);
                editingTextBox.Size = new Size(cellRect.Width, cellRect.Height);
                editingTextBox.Location = cellRect.Location;
                editingTextBox.Text = this[e.ColumnIndex, e.RowIndex].Value?.ToString() ?? "";
                editingTextBox.Visible = true;
                editingTextBox.Focus();
                this.BeginEdit(false);
            }
        }

        private void CustomDataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (editingTextBox.Visible &&
                columnDropdowns.ContainsKey(this.Columns[e.ColumnIndex].Name))
            {
                string newValue = editingTextBox.Text.Trim();
                string columnName = this.Columns[e.ColumnIndex].Name;

                this[e.ColumnIndex, e.RowIndex].Value = newValue;

                if (!string.IsNullOrEmpty(newValue) &&
                    !columnDropdowns[columnName].Contains(newValue))
                {
                    var updatedItems = new List<string>(columnDropdowns[columnName]) { newValue };
                    columnDropdowns[columnName] = updatedItems.ToArray();
                    UpdateComboBoxItems(columnName, newValue);
                    UpdateHeaderDropdownMenu(columnName, newValue);
                }

                editingTextBox.Visible = false;
            }
        }

        private void UpdateComboBoxItems(string columnName, string newValue)
        {
            foreach (DataGridViewColumn column in this.Columns)
            {
                if (column.Name == columnName && column is DataGridViewComboBoxColumn comboColumn)
                {
                    if (!comboColumn.Items.Contains(newValue))
                    {
                        comboColumn.Items.Add(newValue);
                    }
                }
            }
        }

        private void UpdateHeaderDropdownMenu(string columnName, string newValue)
        {
            if (headerDropdownMenus.ContainsKey(columnName))
            {
                ContextMenuStrip menu = headerDropdownMenus[columnName];
                if (!menu.Items.Cast<ToolStripMenuItem>().Any(item => item.Text == newValue))
                {
                    ToolStripMenuItem menuItem = new ToolStripMenuItem(newValue);
                    int columnIndex = this.Columns[columnName].Index;
                    menuItem.Click += (s, e) => HeaderMenuItem_Click(columnIndex, newValue);
                    menu.Items.Add(menuItem);
                }
            }
        }
        #endregion

        #region Event Handlers
        private void CustomDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (Columns[e.ColumnIndex].ValueType == typeof(Image) ||
                Columns[e.ColumnIndex].ValueType == typeof(byte[]))
            {
                e.Cancel = true;
            }
        }
        #endregion

        #region Public Methods
        public DataTable ConvertToDataTable()
        {
            DataTable resultTable = new DataTable();

            foreach (DataGridViewColumn column in this.Columns)
            {
                Type columnType = column.ValueType ?? typeof(string);
                resultTable.Columns.Add(column.Name, columnType);
            }

            foreach (DataGridViewRow row in this.Rows)
            {
                if (row.IsNewRow) continue;

                DataRow dataRow = resultTable.NewRow();
                foreach (DataGridViewColumn column in this.Columns)
                {
                    object cellValue = row.Cells[column.Name].Value;
                    dataRow[column.Name] = cellValue ?? DBNull.Value;
                }
                resultTable.Rows.Add(dataRow);
            }

            return resultTable;
        }

        public DataGridViewRowCollection Row()
        {
            return this.Rows;
        }

        public DataGridViewColumnCollection Column()
        {
            return this.Columns;
        }
        #endregion

        #region Dispose
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                foreach (var menu in headerDropdownMenus.Values)
                {
                    menu.Dispose();
                }
                editingTextBox?.Dispose();
            }
            base.Dispose(disposing);
        }
        #endregion
    }
}
