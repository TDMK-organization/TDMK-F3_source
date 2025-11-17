namespace OK2SHIP_SMT.UserControls
{
    partial class UC_NewFeature
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.slctn_Location = new System.Windows.Forms.SplitContainer();
            this.tbl_Loaction = new System.Windows.Forms.TableLayoutPanel();
            this.label1 = new System.Windows.Forms.Label();
            this.tb_locationFolder = new System.Windows.Forms.TextBox();
            this.lb_loactionFile = new System.Windows.Forms.Label();
            this.list_extension = new System.Windows.Forms.CheckedListBox();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.btnGetData = new System.Windows.Forms.Button();
            this.btn_runProcess = new System.Windows.Forms.Button();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.btn_oneToEnd = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.tb_custom = new System.Windows.Forms.TextBox();
            this.btn_oneToN = new OK2SHIP_SMT.ToolBoxs.TDMK_Button();
            this.btn_unselect = new OK2SHIP_SMT.ToolBoxs.TDMK_Button();
            this.btn_pick = new OK2SHIP_SMT.ToolBoxs.TDMK_Button();
            this.btn_Custom = new OK2SHIP_SMT.ToolBoxs.TDMK_Button();
            ((System.ComponentModel.ISupportInitialize)(this.slctn_Location)).BeginInit();
            this.slctn_Location.Panel1.SuspendLayout();
            this.slctn_Location.Panel2.SuspendLayout();
            this.slctn_Location.SuspendLayout();
            this.tbl_Loaction.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).BeginInit();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            this.btn_oneToEnd.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // slctn_Location
            // 
            this.slctn_Location.Dock = System.Windows.Forms.DockStyle.Fill;
            this.slctn_Location.Location = new System.Drawing.Point(0, 0);
            this.slctn_Location.Name = "slctn_Location";
            // 
            // slctn_Location.Panel1
            // 
            this.slctn_Location.Panel1.Controls.Add(this.tbl_Loaction);
            // 
            // slctn_Location.Panel2
            // 
            this.slctn_Location.Panel2.Controls.Add(this.splitContainer2);
            this.slctn_Location.Size = new System.Drawing.Size(1044, 77);
            this.slctn_Location.SplitterDistance = 823;
            this.slctn_Location.TabIndex = 2;
            // 
            // tbl_Loaction
            // 
            this.tbl_Loaction.ColumnCount = 2;
            this.tbl_Loaction.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tbl_Loaction.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 375F));
            this.tbl_Loaction.Controls.Add(this.label1, 1, 0);
            this.tbl_Loaction.Controls.Add(this.tb_locationFolder, 0, 1);
            this.tbl_Loaction.Controls.Add(this.lb_loactionFile, 0, 0);
            this.tbl_Loaction.Controls.Add(this.list_extension, 1, 1);
            this.tbl_Loaction.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbl_Loaction.Location = new System.Drawing.Point(0, 0);
            this.tbl_Loaction.Name = "tbl_Loaction";
            this.tbl_Loaction.RowCount = 2;
            this.tbl_Loaction.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 23.28767F));
            this.tbl_Loaction.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 76.71233F));
            this.tbl_Loaction.Size = new System.Drawing.Size(823, 77);
            this.tbl_Loaction.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Location = new System.Drawing.Point(451, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(369, 17);
            this.label1.TabIndex = 6;
            this.label1.Text = "Extension";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // tb_locationFolder
            // 
            this.tb_locationFolder.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tb_locationFolder.Location = new System.Drawing.Point(3, 20);
            this.tb_locationFolder.Multiline = true;
            this.tb_locationFolder.Name = "tb_locationFolder";
            this.tb_locationFolder.Size = new System.Drawing.Size(442, 54);
            this.tb_locationFolder.TabIndex = 5;
            this.tb_locationFolder.TextChanged += new System.EventHandler(this.tb_locationFolder_TextChanged);
            this.tb_locationFolder.DoubleClick += new System.EventHandler(this.tb_locationFolder_DoubleClick);
            // 
            // lb_loactionFile
            // 
            this.lb_loactionFile.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lb_loactionFile.Location = new System.Drawing.Point(3, 0);
            this.lb_loactionFile.Name = "lb_loactionFile";
            this.lb_loactionFile.Size = new System.Drawing.Size(442, 17);
            this.lb_loactionFile.TabIndex = 0;
            this.lb_loactionFile.Text = "Loaction ";
            this.lb_loactionFile.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // list_extension
            // 
            this.list_extension.Dock = System.Windows.Forms.DockStyle.Fill;
            this.list_extension.FormattingEnabled = true;
            this.list_extension.Location = new System.Drawing.Point(451, 20);
            this.list_extension.Name = "list_extension";
            this.list_extension.Size = new System.Drawing.Size(369, 54);
            this.list_extension.TabIndex = 7;
            // 
            // splitContainer2
            // 
            this.splitContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer2.Location = new System.Drawing.Point(0, 0);
            this.splitContainer2.Name = "splitContainer2";
            this.splitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.Controls.Add(this.btnGetData);
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.btn_runProcess);
            this.splitContainer2.Size = new System.Drawing.Size(217, 77);
            this.splitContainer2.SplitterDistance = 38;
            this.splitContainer2.TabIndex = 6;
            // 
            // btnGetData
            // 
            this.btnGetData.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnGetData.Location = new System.Drawing.Point(0, 0);
            this.btnGetData.Name = "btnGetData";
            this.btnGetData.Size = new System.Drawing.Size(217, 38);
            this.btnGetData.TabIndex = 5;
            this.btnGetData.Text = "Scan Folder";
            this.btnGetData.UseVisualStyleBackColor = true;
            this.btnGetData.Click += new System.EventHandler(this.btnGetData_Click);
            // 
            // btn_runProcess
            // 
            this.btn_runProcess.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_runProcess.Location = new System.Drawing.Point(0, 0);
            this.btn_runProcess.Name = "btn_runProcess";
            this.btn_runProcess.Size = new System.Drawing.Size(217, 35);
            this.btn_runProcess.TabIndex = 6;
            this.btn_runProcess.Text = "Commit";
            this.btn_runProcess.UseVisualStyleBackColor = true;
            this.btn_runProcess.Click += new System.EventHandler(this.btn_runProcess_Click);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.slctn_Location);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.tableLayoutPanel1);
            this.splitContainer1.Size = new System.Drawing.Size(1044, 538);
            this.splitContainer1.SplitterDistance = 77;
            this.splitContainer1.TabIndex = 3;
            // 
            // dataGridView
            // 
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView.Location = new System.Drawing.Point(3, 3);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.Size = new System.Drawing.Size(822, 451);
            this.dataGridView.TabIndex = 0;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 79.31035F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.68966F));
            this.tableLayoutPanel1.Controls.Add(this.dataGridView, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.btn_oneToEnd, 1, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1044, 457);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // btn_oneToEnd
            // 
            this.btn_oneToEnd.ColumnCount = 1;
            this.btn_oneToEnd.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.btn_oneToEnd.Controls.Add(this.btn_oneToN, 0, 0);
            this.btn_oneToEnd.Controls.Add(this.btn_unselect, 0, 6);
            this.btn_oneToEnd.Controls.Add(this.btn_pick, 0, 5);
            this.btn_oneToEnd.Controls.Add(this.tableLayoutPanel2, 0, 2);
            this.btn_oneToEnd.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_oneToEnd.Location = new System.Drawing.Point(831, 3);
            this.btn_oneToEnd.Name = "btn_oneToEnd";
            this.btn_oneToEnd.RowCount = 7;
            this.btn_oneToEnd.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571F));
            this.btn_oneToEnd.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571F));
            this.btn_oneToEnd.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 17.73836F));
            this.btn_oneToEnd.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.64302F));
            this.btn_oneToEnd.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571F));
            this.btn_oneToEnd.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571F));
            this.btn_oneToEnd.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571F));
            this.btn_oneToEnd.Size = new System.Drawing.Size(210, 451);
            this.btn_oneToEnd.TabIndex = 1;
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble;
            this.tableLayoutPanel2.ColumnCount = 1;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.Controls.Add(this.btn_Custom, 0, 1);
            this.tableLayoutPanel2.Controls.Add(this.tb_custom, 0, 0);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(3, 131);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 2;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(204, 74);
            this.tableLayoutPanel2.TabIndex = 4;
            // 
            // tb_custom
            // 
            this.tb_custom.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tb_custom.Location = new System.Drawing.Point(6, 6);
            this.tb_custom.Multiline = true;
            this.tb_custom.Name = "tb_custom";
            this.tb_custom.Size = new System.Drawing.Size(192, 26);
            this.tb_custom.TabIndex = 2;
            // 
            // btn_oneToN
            // 
            this.btn_oneToN.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_oneToN.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold);
            this.btn_oneToN.Location = new System.Drawing.Point(3, 3);
            this.btn_oneToN.Name = "btn_oneToN";
            this.btn_oneToN.Size = new System.Drawing.Size(204, 58);
            this.btn_oneToN.TabIndex = 0;
            this.btn_oneToN.Text = "1 - N";
            this.btn_oneToN.UseVisualStyleBackColor = true;
            this.btn_oneToN.Click += new System.EventHandler(this.btn_oneToN_Click);
            // 
            // btn_unselect
            // 
            this.btn_unselect.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_unselect.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold);
            this.btn_unselect.Location = new System.Drawing.Point(3, 387);
            this.btn_unselect.Name = "btn_unselect";
            this.btn_unselect.Size = new System.Drawing.Size(204, 61);
            this.btn_unselect.TabIndex = 3;
            this.btn_unselect.Text = "Unselect";
            this.btn_unselect.UseVisualStyleBackColor = true;
            this.btn_unselect.Click += new System.EventHandler(this.btn_unselect_Click);
            // 
            // btn_pick
            // 
            this.btn_pick.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_pick.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold);
            this.btn_pick.Location = new System.Drawing.Point(3, 323);
            this.btn_pick.Name = "btn_pick";
            this.btn_pick.Size = new System.Drawing.Size(204, 58);
            this.btn_pick.TabIndex = 2;
            this.btn_pick.Text = "Select";
            this.btn_pick.UseVisualStyleBackColor = true;
            this.btn_pick.Click += new System.EventHandler(this.btn_pick_Click);
            // 
            // btn_Custom
            // 
            this.btn_Custom.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_Custom.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold);
            this.btn_Custom.Location = new System.Drawing.Point(6, 41);
            this.btn_Custom.Name = "btn_Custom";
            this.btn_Custom.Size = new System.Drawing.Size(192, 27);
            this.btn_Custom.TabIndex = 1;
            this.btn_Custom.Text = "Custom";
            this.btn_Custom.UseVisualStyleBackColor = true;
            this.btn_Custom.Click += new System.EventHandler(this.btn_Custom_Click);
            // 
            // UC_NewFeature
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.splitContainer1);
            this.Name = "UC_NewFeature";
            this.Size = new System.Drawing.Size(1044, 538);
            this.slctn_Location.Panel1.ResumeLayout(false);
            this.slctn_Location.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.slctn_Location)).EndInit();
            this.slctn_Location.ResumeLayout(false);
            this.tbl_Loaction.ResumeLayout(false);
            this.tbl_Loaction.PerformLayout();
            this.splitContainer2.Panel1.ResumeLayout(false);
            this.splitContainer2.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).EndInit();
            this.splitContainer2.ResumeLayout(false);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.btn_oneToEnd.ResumeLayout(false);
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer slctn_Location;
        private System.Windows.Forms.Button btnGetData;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.SplitContainer splitContainer2;
        private System.Windows.Forms.Button btn_runProcess;
        private System.Windows.Forms.TableLayoutPanel tbl_Loaction;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tb_locationFolder;
        private System.Windows.Forms.Label lb_loactionFile;
        private System.Windows.Forms.CheckedListBox list_extension;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TableLayoutPanel btn_oneToEnd;
        private ToolBoxs.TDMK_Button btn_oneToN;
        private ToolBoxs.TDMK_Button btn_Custom;
        private ToolBoxs.TDMK_Button btn_pick;
        private ToolBoxs.TDMK_Button btn_unselect;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.TextBox tb_custom;
    }
}
