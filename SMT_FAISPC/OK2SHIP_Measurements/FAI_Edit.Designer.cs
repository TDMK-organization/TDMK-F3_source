
namespace OK2SHIP_Measurements
{
    partial class FAI_Edit
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            this.DGV_Spec = new System.Windows.Forms.DataGridView();
            this.DGV_CPK = new System.Windows.Forms.DataGridView();
            this.DGV_Histogram = new System.Windows.Forms.DataGridView();
            this.DGV_Data = new System.Windows.Forms.DataGridView();
            this.DGV_LogFile = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.txtDataFolder = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.tblMain = new System.Windows.Forms.TableLayoutPanel();
            this.cmsAction = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.tsmInsert = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmRemove = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmReplace = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmReplaceAll = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmReplaceSelection = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmReplaceAllData = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmRename = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmSelectedColumn = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmDefault = new System.Windows.Forms.ToolStripMenuItem();
            this.copyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pasteToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.selectedRowsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.selectedColumnsToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.cmsDelete = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.tsmDelSelect = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmRemoveFAI = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmSel_pcs = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Spec)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_CPK)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Histogram)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_LogFile)).BeginInit();
            this.tblMain.SuspendLayout();
            this.cmsAction.SuspendLayout();
            this.cmsDelete.SuspendLayout();
            this.SuspendLayout();
            // 
            // DGV_Spec
            // 
            this.DGV_Spec.AllowUserToAddRows = false;
            this.DGV_Spec.AllowUserToDeleteRows = false;
            this.DGV_Spec.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_Spec.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV_Spec.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.DGV_Spec.Location = new System.Drawing.Point(597, 233);
            this.DGV_Spec.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.DGV_Spec.Name = "DGV_Spec";
            this.DGV_Spec.Size = new System.Drawing.Size(588, 145);
            this.DGV_Spec.TabIndex = 2;
            // 
            // DGV_CPK
            // 
            this.DGV_CPK.AllowUserToAddRows = false;
            this.DGV_CPK.AllowUserToDeleteRows = false;
            this.DGV_CPK.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_CPK.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV_CPK.Location = new System.Drawing.Point(3, 232);
            this.DGV_CPK.Name = "DGV_CPK";
            this.DGV_CPK.ReadOnly = true;
            this.DGV_CPK.Size = new System.Drawing.Size(588, 147);
            this.DGV_CPK.TabIndex = 6;
            // 
            // DGV_Histogram
            // 
            this.DGV_Histogram.AllowUserToAddRows = false;
            this.DGV_Histogram.AllowUserToDeleteRows = false;
            this.DGV_Histogram.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.Format = "N3";
            dataGridViewCellStyle1.NullValue = null;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.DGV_Histogram.DefaultCellStyle = dataGridViewCellStyle1;
            this.DGV_Histogram.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV_Histogram.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.DGV_Histogram.Location = new System.Drawing.Point(3, 42);
            this.DGV_Histogram.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.DGV_Histogram.Name = "DGV_Histogram";
            this.DGV_Histogram.Size = new System.Drawing.Size(588, 145);
            this.DGV_Histogram.TabIndex = 7;
            // 
            // DGV_Data
            // 
            this.DGV_Data.AllowUserToAddRows = false;
            this.DGV_Data.AllowUserToDeleteRows = false;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.DGV_Data.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.DGV_Data.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.Format = "N3";
            dataGridViewCellStyle3.NullValue = null;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.DGV_Data.DefaultCellStyle = dataGridViewCellStyle3;
            this.DGV_Data.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV_Data.Location = new System.Drawing.Point(3, 424);
            this.DGV_Data.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.DGV_Data.Name = "DGV_Data";
            this.DGV_Data.Size = new System.Drawing.Size(588, 337);
            this.DGV_Data.TabIndex = 4;
            this.DGV_Data.ColumnHeaderMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.DGV_Data_ColumnHeaderMouseDoubleClick);
            this.DGV_Data.RowHeaderMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.DGV_Data_RowHeaderMouseDoubleClick);
            this.DGV_Data.MouseClick += new System.Windows.Forms.MouseEventHandler(this.DGV_Data_MouseClick);
            // 
            // DGV_LogFile
            // 
            this.DGV_LogFile.AllowUserToAddRows = false;
            this.DGV_LogFile.AllowUserToDeleteRows = false;
            this.DGV_LogFile.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle4.Format = "N3";
            dataGridViewCellStyle4.NullValue = null;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.DGV_LogFile.DefaultCellStyle = dataGridViewCellStyle4;
            this.DGV_LogFile.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV_LogFile.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.DGV_LogFile.Location = new System.Drawing.Point(597, 424);
            this.DGV_LogFile.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.DGV_LogFile.Name = "DGV_LogFile";
            this.DGV_LogFile.Size = new System.Drawing.Size(588, 337);
            this.DGV_LogFile.TabIndex = 5;
            this.DGV_LogFile.MouseClick += new System.Windows.Forms.MouseEventHandler(this.DGV_LogFile_MouseClick);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Font = new System.Drawing.Font("Arial", 15.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(597, 191);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(588, 38);
            this.label1.TabIndex = 8;
            this.label1.Text = "Specification";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label3.Font = new System.Drawing.Font("Arial", 15.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(597, 382);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(588, 38);
            this.label3.TabIndex = 9;
            this.label3.Text = "Confirmed Logfile Data";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label4.Font = new System.Drawing.Font("Arial", 15.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(3, 382);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(588, 38);
            this.label4.TabIndex = 10;
            this.label4.Text = "Database: FAI Data";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Font = new System.Drawing.Font("Arial", 15.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(3, 191);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(588, 38);
            this.label2.TabIndex = 11;
            this.label2.Text = "Analysis Results";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label7.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label7.Font = new System.Drawing.Font("Arial", 15.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(3, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(588, 38);
            this.label7.TabIndex = 12;
            this.label7.Text = "Histogram";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtDataFolder
            // 
            this.txtDataFolder.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtDataFolder.Location = new System.Drawing.Point(597, 41);
            this.txtDataFolder.Multiline = true;
            this.txtDataFolder.Name = "txtDataFolder";
            this.txtDataFolder.Size = new System.Drawing.Size(588, 147);
            this.txtDataFolder.TabIndex = 13;
            this.txtDataFolder.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDataFolder_KeyDown);
            this.txtDataFolder.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.txtDataFolder_MouseDoubleClick);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label5.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(597, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(588, 38);
            this.label5.TabIndex = 14;
            this.label5.Text = "Chọn thư mục chứa logfile (*.csv)";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tblMain
            // 
            this.tblMain.ColumnCount = 2;
            this.tblMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tblMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tblMain.Controls.Add(this.label7, 0, 0);
            this.tblMain.Controls.Add(this.DGV_LogFile, 1, 5);
            this.tblMain.Controls.Add(this.label3, 1, 4);
            this.tblMain.Controls.Add(this.txtDataFolder, 1, 1);
            this.tblMain.Controls.Add(this.label1, 1, 2);
            this.tblMain.Controls.Add(this.DGV_Spec, 1, 3);
            this.tblMain.Controls.Add(this.label5, 1, 0);
            this.tblMain.Controls.Add(this.DGV_Histogram, 0, 1);
            this.tblMain.Controls.Add(this.label2, 0, 2);
            this.tblMain.Controls.Add(this.label4, 0, 4);
            this.tblMain.Controls.Add(this.DGV_CPK, 0, 3);
            this.tblMain.Controls.Add(this.DGV_Data, 0, 5);
            this.tblMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tblMain.Location = new System.Drawing.Point(0, 0);
            this.tblMain.Name = "tblMain";
            this.tblMain.RowCount = 6;
            this.tblMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5F));
            this.tblMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tblMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5F));
            this.tblMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tblMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5F));
            this.tblMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 45F));
            this.tblMain.Size = new System.Drawing.Size(1188, 765);
            this.tblMain.TabIndex = 15;
            // 
            // cmsAction
            // 
            this.cmsAction.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmInsert,
            this.tsmRemove,
            this.tsmReplace,
            this.tsmReplaceAll,
            this.tsmRename,
            this.copyToolStripMenuItem,
            this.pasteToolStripMenuItem1});
            this.cmsAction.Name = "cmsAction";
            this.cmsAction.Size = new System.Drawing.Size(181, 180);
            // 
            // tsmInsert
            // 
            this.tsmInsert.Name = "tsmInsert";
            this.tsmInsert.Size = new System.Drawing.Size(180, 22);
            this.tsmInsert.Text = "Insert";
            this.tsmInsert.Click += new System.EventHandler(this.tsmInsert_Click);
            // 
            // tsmRemove
            // 
            this.tsmRemove.Name = "tsmRemove";
            this.tsmRemove.Size = new System.Drawing.Size(180, 22);
            this.tsmRemove.Text = "Remove";
            this.tsmRemove.Click += new System.EventHandler(this.tsmRemove_Click);
            // 
            // tsmReplace
            // 
            this.tsmReplace.Name = "tsmReplace";
            this.tsmReplace.Size = new System.Drawing.Size(180, 22);
            this.tsmReplace.Text = "Replace";
            this.tsmReplace.Click += new System.EventHandler(this.tsmReplace_Click);
            // 
            // tsmReplaceAll
            // 
            this.tsmReplaceAll.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmReplaceSelection,
            this.tsmReplaceAllData,
            this.tsmSel_pcs});
            this.tsmReplaceAll.Name = "tsmReplaceAll";
            this.tsmReplaceAll.Size = new System.Drawing.Size(180, 22);
            this.tsmReplaceAll.Text = "Replace FAI";
            this.tsmReplaceAll.Click += new System.EventHandler(this.tsmReplaceAll_Click);
            // 
            // tsmReplaceSelection
            // 
            this.tsmReplaceSelection.Name = "tsmReplaceSelection";
            this.tsmReplaceSelection.Size = new System.Drawing.Size(180, 22);
            this.tsmReplaceSelection.Text = "Selection";
            this.tsmReplaceSelection.Click += new System.EventHandler(this.tsmReplaceSelection_Click);
            // 
            // tsmReplaceAllData
            // 
            this.tsmReplaceAllData.Name = "tsmReplaceAllData";
            this.tsmReplaceAllData.Size = new System.Drawing.Size(180, 22);
            this.tsmReplaceAllData.Text = "All Data";
            this.tsmReplaceAllData.Click += new System.EventHandler(this.tsmReplaceAllData_Click);
            // 
            // tsmRename
            // 
            this.tsmRename.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmSelectedColumn,
            this.tsmDefault});
            this.tsmRename.Name = "tsmRename";
            this.tsmRename.Size = new System.Drawing.Size(180, 22);
            this.tsmRename.Text = "Rename";
            // 
            // tsmSelectedColumn
            // 
            this.tsmSelectedColumn.Name = "tsmSelectedColumn";
            this.tsmSelectedColumn.Size = new System.Drawing.Size(164, 22);
            this.tsmSelectedColumn.Text = "Selected Column";
            this.tsmSelectedColumn.Click += new System.EventHandler(this.tsmSelectedColumn_Click);
            // 
            // tsmDefault
            // 
            this.tsmDefault.Name = "tsmDefault";
            this.tsmDefault.Size = new System.Drawing.Size(164, 22);
            this.tsmDefault.Text = "default";
            this.tsmDefault.Click += new System.EventHandler(this.tsmDefault_Click);
            // 
            // copyToolStripMenuItem
            // 
            this.copyToolStripMenuItem.Name = "copyToolStripMenuItem";
            this.copyToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.copyToolStripMenuItem.Text = "Copy";
            this.copyToolStripMenuItem.Click += new System.EventHandler(this.copyToolStripMenuItem_Click);
            // 
            // pasteToolStripMenuItem1
            // 
            this.pasteToolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.selectedRowsToolStripMenuItem,
            this.selectedColumnsToolStripMenuItem1});
            this.pasteToolStripMenuItem1.Name = "pasteToolStripMenuItem1";
            this.pasteToolStripMenuItem1.Size = new System.Drawing.Size(180, 22);
            this.pasteToolStripMenuItem1.Text = "Paste";
            // 
            // selectedRowsToolStripMenuItem
            // 
            this.selectedRowsToolStripMenuItem.Name = "selectedRowsToolStripMenuItem";
            this.selectedRowsToolStripMenuItem.Size = new System.Drawing.Size(169, 22);
            this.selectedRowsToolStripMenuItem.Text = "Selected Rows";
            this.selectedRowsToolStripMenuItem.Click += new System.EventHandler(this.selectedRowsToolStripMenuItem_Click);
            // 
            // selectedColumnsToolStripMenuItem1
            // 
            this.selectedColumnsToolStripMenuItem1.Name = "selectedColumnsToolStripMenuItem1";
            this.selectedColumnsToolStripMenuItem1.Size = new System.Drawing.Size(169, 22);
            this.selectedColumnsToolStripMenuItem1.Text = "Selected Columns";
            this.selectedColumnsToolStripMenuItem1.Click += new System.EventHandler(this.selectedColumnsToolStripMenuItem1_Click);
            // 
            // cmsDelete
            // 
            this.cmsDelete.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmDelSelect,
            this.tsmRemoveFAI});
            this.cmsDelete.Name = "cmsDelete";
            this.cmsDelete.Size = new System.Drawing.Size(184, 48);
            // 
            // tsmDelSelect
            // 
            this.tsmDelSelect.Name = "tsmDelSelect";
            this.tsmDelSelect.Size = new System.Drawing.Size(183, 22);
            this.tsmDelSelect.Text = "Delete Selection";
            this.tsmDelSelect.Click += new System.EventHandler(this.tsmDelSelect_Click);
            // 
            // tsmRemoveFAI
            // 
            this.tsmRemoveFAI.Name = "tsmRemoveFAI";
            this.tsmRemoveFAI.Size = new System.Drawing.Size(183, 22);
            this.tsmRemoveFAI.Text = "Remove Selected FAI";
            this.tsmRemoveFAI.Click += new System.EventHandler(this.tsmRemoveFAI_Click);
            // 
            // tsmSel_pcs
            // 
            this.tsmSel_pcs.Name = "tsmSel_pcs";
            this.tsmSel_pcs.Size = new System.Drawing.Size(180, 22);
            this.tsmSel_pcs.Text = "Selected_Pcs";
            this.tsmSel_pcs.Click += new System.EventHandler(this.tsmSel_pcs_Click);
            // 
            // FAI_Edit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1188, 765);
            this.Controls.Add(this.tblMain);
            this.Name = "FAI_Edit";
            this.Text = "FAI_Edit";
            this.Load += new System.EventHandler(this.FAI_Edit_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Spec)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_CPK)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Histogram)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_LogFile)).EndInit();
            this.tblMain.ResumeLayout(false);
            this.tblMain.PerformLayout();
            this.cmsAction.ResumeLayout(false);
            this.cmsDelete.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView DGV_Spec;
        private System.Windows.Forms.DataGridView DGV_CPK;
        private System.Windows.Forms.DataGridView DGV_Histogram;
        private System.Windows.Forms.DataGridView DGV_Data;
        private System.Windows.Forms.DataGridView DGV_LogFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtDataFolder;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TableLayoutPanel tblMain;
        private System.Windows.Forms.ContextMenuStrip cmsAction;
        private System.Windows.Forms.ToolStripMenuItem tsmInsert;
        private System.Windows.Forms.ToolStripMenuItem tsmRemove;
        private System.Windows.Forms.ToolStripMenuItem tsmReplace;
        private System.Windows.Forms.ToolStripMenuItem tsmRename;
        private System.Windows.Forms.ToolStripMenuItem tsmSelectedColumn;
        private System.Windows.Forms.ToolStripMenuItem tsmDefault;
        private System.Windows.Forms.ToolStripMenuItem copyToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem pasteToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem selectedRowsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem selectedColumnsToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem tsmReplaceAll;
        private System.Windows.Forms.ToolStripMenuItem tsmReplaceSelection;
        private System.Windows.Forms.ToolStripMenuItem tsmReplaceAllData;
        private System.Windows.Forms.ContextMenuStrip cmsDelete;
        private System.Windows.Forms.ToolStripMenuItem tsmDelSelect;
        private System.Windows.Forms.ToolStripMenuItem tsmRemoveFAI;
        private System.Windows.Forms.ToolStripMenuItem tsmSel_pcs;
    }
}