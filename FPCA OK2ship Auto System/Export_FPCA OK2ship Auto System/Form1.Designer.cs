namespace Export_FPCA_OK2ship_Auto_System
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.cbAll = new System.Windows.Forms.CheckBox();
            this.cbl_sheet = new System.Windows.Forms.CheckedListBox();
            this.btn_check = new System.Windows.Forms.Button();
            this.txtLotNo = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtItemCode = new System.Windows.Forms.TextBox();
            this.lblItemCode = new System.Windows.Forms.Label();
            this.dgv_progress = new System.Windows.Forms.DataGridView();
            this.lbltitle = new System.Windows.Forms.Label();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.dgv_View = new System.Windows.Forms.DataGridView();
            this.tabMain = new System.Windows.Forms.TabControl();
            this.tabSelFunc = new System.Windows.Forms.TabPage();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.lblSetting = new System.Windows.Forms.Label();
            this.lblFAI = new System.Windows.Forms.Label();
            this.lblBending = new System.Windows.Forms.Label();
            this.lblType3 = new System.Windows.Forms.Label();
            this.tabExport = new System.Windows.Forms.TabPage();
            this.spMain = new System.Windows.Forms.SplitContainer();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btn_convert = new System.Windows.Forms.Button();
            this.btn_export = new System.Windows.Forms.Button();
            this.spDataView = new System.Windows.Forms.SplitContainer();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_progress)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_View)).BeginInit();
            this.tabMain.SuspendLayout();
            this.tabSelFunc.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            this.tabExport.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.spMain)).BeginInit();
            this.spMain.Panel1.SuspendLayout();
            this.spMain.Panel2.SuspendLayout();
            this.spMain.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.spDataView)).BeginInit();
            this.spDataView.Panel1.SuspendLayout();
            this.spDataView.Panel2.SuspendLayout();
            this.spDataView.SuspendLayout();
            this.SuspendLayout();
            // 
            // cbAll
            // 
            this.cbAll.AutoSize = true;
            this.cbAll.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbAll.Location = new System.Drawing.Point(191, 31);
            this.cbAll.Name = "cbAll";
            this.cbAll.Size = new System.Drawing.Size(56, 22);
            this.cbAll.TabIndex = 16;
            this.cbAll.Text = "ALL";
            this.cbAll.UseVisualStyleBackColor = true;
            this.cbAll.CheckedChanged += new System.EventHandler(this.cbAll_CheckedChanged);
            // 
            // cbl_sheet
            // 
            this.cbl_sheet.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cbl_sheet.Cursor = System.Windows.Forms.Cursors.Hand;
            this.cbl_sheet.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbl_sheet.FormattingEnabled = true;
            this.cbl_sheet.Items.AddRange(new object[] {
            "Coverpage",
            "Rev History",
            "User Guideline",
            "Low CPK Action",
            "Declaration",
            "Table of Contents",
            "Deviation summary",
            "Assy Yield",
            "FAI",
            "OQC Test",
            "Cross section",
            "GAP Connector",
            "Peel Test",
            "(Mating) Pull Test",
            "IQC Liner peeling (Coupon)",
            "IQC PSA peeling (Coupon)",
            "Shear test",
            "Liner peel test (On product)",
            "Air bubble btw Liner-PSA",
            "PSA peel test (On product)",
            "Air bubble btw PSA-FPC",
            "(IQC Unmating) Pull Test",
            "OQC B2B Mating-Unmating",
            "ORT-Assy",
            "Flex bending",
            "Thermal Cycling & bending",
            "Heat Soak & bending",
            "X-Ray picture",
            "Heat Soak and Recovery",
            "Thermal Cycling",
            "Thermal Shock",
            "Environment en-durance",
            "Impedance",
            "Switch Quality",
            "ACF",
            "SEM BSE & Binarization",
            "Bar Code Verification",
            "Packaging",
            "Mishandling test",
            "Process flow",
            "Process Comparison"});
            this.cbl_sheet.Location = new System.Drawing.Point(3, 59);
            this.cbl_sheet.Name = "cbl_sheet";
            this.cbl_sheet.Size = new System.Drawing.Size(244, 340);
            this.cbl_sheet.TabIndex = 12;
            // 
            // btn_check
            // 
            this.btn_check.BackColor = System.Drawing.Color.LightSkyBlue;
            this.btn_check.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_check.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_check.Location = new System.Drawing.Point(122, 117);
            this.btn_check.Name = "btn_check";
            this.btn_check.Size = new System.Drawing.Size(110, 40);
            this.btn_check.TabIndex = 43;
            this.btn_check.Text = "CHECK";
            this.btn_check.UseVisualStyleBackColor = false;
            this.btn_check.Click += new System.EventHandler(this.btn_check_Click);
            // 
            // txtLotNo
            // 
            this.txtLotNo.Location = new System.Drawing.Point(85, 75);
            this.txtLotNo.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtLotNo.Name = "txtLotNo";
            this.txtLotNo.Size = new System.Drawing.Size(147, 26);
            this.txtLotNo.TabIndex = 18;
            this.txtLotNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtLotNo.TextChanged += new System.EventHandler(this.txtLotNo_TextChanged);
            this.txtLotNo.Validated += new System.EventHandler(this.txtLotNo_Validated);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(3, 78);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(52, 20);
            this.label3.TabIndex = 17;
            this.label3.Text = "LotNo";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // txtItemCode
            // 
            this.txtItemCode.Location = new System.Drawing.Point(85, 28);
            this.txtItemCode.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtItemCode.Name = "txtItemCode";
            this.txtItemCode.Size = new System.Drawing.Size(147, 26);
            this.txtItemCode.TabIndex = 1;
            this.txtItemCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtItemCode.TextChanged += new System.EventHandler(this.txtItemCode_TextChanged);
            // 
            // lblItemCode
            // 
            this.lblItemCode.AutoSize = true;
            this.lblItemCode.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblItemCode.Location = new System.Drawing.Point(3, 31);
            this.lblItemCode.Name = "lblItemCode";
            this.lblItemCode.Size = new System.Drawing.Size(79, 20);
            this.lblItemCode.TabIndex = 0;
            this.lblItemCode.Text = "ItemCode";
            // 
            // dgv_progress
            // 
            this.dgv_progress.AllowUserToAddRows = false;
            this.dgv_progress.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgv_progress.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_progress.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_progress.Location = new System.Drawing.Point(3, 3);
            this.dgv_progress.Name = "dgv_progress";
            this.dgv_progress.ReadOnly = true;
            this.dgv_progress.RowHeadersVisible = false;
            this.dgv_progress.RowHeadersWidth = 82;
            this.dgv_progress.Size = new System.Drawing.Size(344, 657);
            this.dgv_progress.TabIndex = 17;
            this.dgv_progress.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_progress_CellContentDoubleClick);
            // 
            // lbltitle
            // 
            this.lbltitle.BackColor = System.Drawing.Color.RosyBrown;
            this.lbltitle.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lbltitle.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbltitle.Font = new System.Drawing.Font("Arial", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbltitle.Location = new System.Drawing.Point(0, 0);
            this.lbltitle.Name = "lbltitle";
            this.lbltitle.Size = new System.Drawing.Size(613, 52);
            this.lbltitle.TabIndex = 43;
            this.lbltitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lbltitle.Click += new System.EventHandler(this.lbltitle_Click);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 350F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.dgv_progress, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.dgv_View, 1, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(613, 663);
            this.tableLayoutPanel1.TabIndex = 44;
            // 
            // dgv_View
            // 
            this.dgv_View.AllowUserToAddRows = false;
            this.dgv_View.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_View.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_View.Location = new System.Drawing.Point(353, 3);
            this.dgv_View.Name = "dgv_View";
            this.dgv_View.ReadOnly = true;
            this.dgv_View.RowHeadersVisible = false;
            this.dgv_View.RowHeadersWidth = 82;
            this.dgv_View.Size = new System.Drawing.Size(257, 657);
            this.dgv_View.TabIndex = 18;
            this.dgv_View.ColumnHeaderMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dgv_View_ColumnHeaderMouseDoubleClick);
            // 
            // tabMain
            // 
            this.tabMain.Controls.Add(this.tabSelFunc);
            this.tabMain.Controls.Add(this.tabExport);
            this.tabMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabMain.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabMain.Location = new System.Drawing.Point(0, 0);
            this.tabMain.Name = "tabMain";
            this.tabMain.SelectedIndex = 0;
            this.tabMain.Size = new System.Drawing.Size(890, 756);
            this.tabMain.TabIndex = 45;
            // 
            // tabSelFunc
            // 
            this.tabSelFunc.Controls.Add(this.tableLayoutPanel3);
            this.tabSelFunc.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabSelFunc.Location = new System.Drawing.Point(4, 27);
            this.tabSelFunc.Name = "tabSelFunc";
            this.tabSelFunc.Padding = new System.Windows.Forms.Padding(3);
            this.tabSelFunc.Size = new System.Drawing.Size(882, 725);
            this.tabSelFunc.TabIndex = 0;
            this.tabSelFunc.Text = "Select Function";
            this.tabSelFunc.UseVisualStyleBackColor = true;
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.ColumnCount = 2;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel3.Controls.Add(this.lblSetting, 0, 1);
            this.tableLayoutPanel3.Controls.Add(this.lblFAI, 0, 0);
            this.tableLayoutPanel3.Controls.Add(this.lblBending, 1, 0);
            this.tableLayoutPanel3.Controls.Add(this.lblType3, 1, 1);
            this.tableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel3.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 2;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.16077F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 49.83923F));
            this.tableLayoutPanel3.Size = new System.Drawing.Size(876, 719);
            this.tableLayoutPanel3.TabIndex = 4;
            // 
            // lblSetting
            // 
            this.lblSetting.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblSetting.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblSetting.Font = new System.Drawing.Font("Arial", 26.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSetting.Image = ((System.Drawing.Image)(resources.GetObject("lblSetting.Image")));
            this.lblSetting.Location = new System.Drawing.Point(3, 360);
            this.lblSetting.Name = "lblSetting";
            this.lblSetting.Size = new System.Drawing.Size(432, 359);
            this.lblSetting.TabIndex = 3;
            this.lblSetting.Text = "Setting";
            this.lblSetting.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.lblSetting.Click += new System.EventHandler(this.lblSetting_Click);
            // 
            // lblFAI
            // 
            this.lblFAI.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblFAI.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblFAI.Font = new System.Drawing.Font("Arial", 26.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFAI.Image = ((System.Drawing.Image)(resources.GetObject("lblFAI.Image")));
            this.lblFAI.Location = new System.Drawing.Point(3, 0);
            this.lblFAI.Name = "lblFAI";
            this.lblFAI.Size = new System.Drawing.Size(432, 360);
            this.lblFAI.TabIndex = 0;
            this.lblFAI.Text = "FAI/SPC";
            this.lblFAI.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.lblFAI.Click += new System.EventHandler(this.lblFAI_Click);
            // 
            // lblBending
            // 
            this.lblBending.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblBending.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblBending.Font = new System.Drawing.Font("Arial", 26.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblBending.Image = ((System.Drawing.Image)(resources.GetObject("lblBending.Image")));
            this.lblBending.Location = new System.Drawing.Point(441, 0);
            this.lblBending.Name = "lblBending";
            this.lblBending.Size = new System.Drawing.Size(432, 360);
            this.lblBending.TabIndex = 1;
            this.lblBending.Text = "Bending";
            this.lblBending.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.lblBending.Click += new System.EventHandler(this.lblBending_Click);
            // 
            // lblType3
            // 
            this.lblType3.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblType3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblType3.Font = new System.Drawing.Font("Arial", 26.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblType3.Image = ((System.Drawing.Image)(resources.GetObject("lblType3.Image")));
            this.lblType3.Location = new System.Drawing.Point(441, 360);
            this.lblType3.Name = "lblType3";
            this.lblType3.Size = new System.Drawing.Size(432, 359);
            this.lblType3.TabIndex = 2;
            this.lblType3.Text = "VHX-IMADA";
            this.lblType3.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.lblType3.Click += new System.EventHandler(this.lblType3_Click);
            // 
            // tabExport
            // 
            this.tabExport.Controls.Add(this.spMain);
            this.tabExport.Location = new System.Drawing.Point(4, 27);
            this.tabExport.Name = "tabExport";
            this.tabExport.Padding = new System.Windows.Forms.Padding(3);
            this.tabExport.Size = new System.Drawing.Size(882, 725);
            this.tabExport.TabIndex = 1;
            this.tabExport.Text = "Export Data";
            this.tabExport.UseVisualStyleBackColor = true;
            // 
            // spMain
            // 
            this.spMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.spMain.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.spMain.Location = new System.Drawing.Point(3, 3);
            this.spMain.Name = "spMain";
            // 
            // spMain.Panel1
            // 
            this.spMain.Panel1.Controls.Add(this.tableLayoutPanel2);
            // 
            // spMain.Panel2
            // 
            this.spMain.Panel2.Controls.Add(this.spDataView);
            this.spMain.Size = new System.Drawing.Size(876, 719);
            this.spMain.SplitterDistance = 259;
            this.spMain.TabIndex = 0;
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 1;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.Controls.Add(this.groupBox3, 0, 1);
            this.tableLayoutPanel2.Controls.Add(this.groupBox1, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.btn_export, 0, 2);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 3;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 180F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 81.22172F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 18.77828F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(259, 719);
            this.tableLayoutPanel2.TabIndex = 20;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.cbl_sheet);
            this.groupBox3.Controls.Add(this.cbAll);
            this.groupBox3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox3.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox3.Location = new System.Drawing.Point(3, 183);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(253, 431);
            this.groupBox3.TabIndex = 21;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = " Lựa chọn sheet xuất dữ liệu";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btn_convert);
            this.groupBox1.Controls.Add(this.btn_check);
            this.groupBox1.Controls.Add(this.lblItemCode);
            this.groupBox1.Controls.Add(this.txtItemCode);
            this.groupBox1.Controls.Add(this.txtLotNo);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(3, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(253, 174);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Nhập thông tin";
            // 
            // btn_convert
            // 
            this.btn_convert.BackColor = System.Drawing.Color.Plum;
            this.btn_convert.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_convert.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_convert.Location = new System.Drawing.Point(7, 117);
            this.btn_convert.Name = "btn_convert";
            this.btn_convert.Size = new System.Drawing.Size(109, 40);
            this.btn_convert.TabIndex = 44;
            this.btn_convert.Text = "Cài đặt Code/Lot NVL";
            this.btn_convert.UseVisualStyleBackColor = false;
            this.btn_convert.Click += new System.EventHandler(this.btn_convert_Click);
            // 
            // btn_export
            // 
            this.btn_export.BackColor = System.Drawing.Color.MediumAquamarine;
            this.btn_export.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_export.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_export.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_export.Location = new System.Drawing.Point(3, 620);
            this.btn_export.Name = "btn_export";
            this.btn_export.Size = new System.Drawing.Size(253, 96);
            this.btn_export.TabIndex = 42;
            this.btn_export.Text = "EXPORT";
            this.btn_export.UseVisualStyleBackColor = false;
            this.btn_export.Click += new System.EventHandler(this.btn_export_Click);
            // 
            // spDataView
            // 
            this.spDataView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.spDataView.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.spDataView.Location = new System.Drawing.Point(0, 0);
            this.spDataView.Name = "spDataView";
            this.spDataView.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // spDataView.Panel1
            // 
            this.spDataView.Panel1.Controls.Add(this.lbltitle);
            // 
            // spDataView.Panel2
            // 
            this.spDataView.Panel2.Controls.Add(this.tableLayoutPanel1);
            this.spDataView.Size = new System.Drawing.Size(613, 719);
            this.spDataView.SplitterDistance = 52;
            this.spDataView.TabIndex = 45;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(890, 756);
            this.Controls.Add(this.tabMain);
            this.Name = "Form1";
            this.Text = "FPCA OK2SHIP Auto System";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_progress)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_View)).EndInit();
            this.tabMain.ResumeLayout(false);
            this.tabSelFunc.ResumeLayout(false);
            this.tableLayoutPanel3.ResumeLayout(false);
            this.tabExport.ResumeLayout(false);
            this.spMain.Panel1.ResumeLayout(false);
            this.spMain.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.spMain)).EndInit();
            this.spMain.ResumeLayout(false);
            this.tableLayoutPanel2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.spDataView.Panel1.ResumeLayout(false);
            this.spDataView.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.spDataView)).EndInit();
            this.spDataView.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.CheckBox cbAll;
        private System.Windows.Forms.CheckedListBox cbl_sheet;
        private System.Windows.Forms.TextBox txtLotNo;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtItemCode;
        private System.Windows.Forms.Label lblItemCode;
        private System.Windows.Forms.Button btn_check;
        private System.Windows.Forms.DataGridView dgv_progress;
        private System.Windows.Forms.Label lbltitle;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.DataGridView dgv_View;
        private System.Windows.Forms.TabControl tabMain;
        private System.Windows.Forms.TabPage tabSelFunc;
        private System.Windows.Forms.TabPage tabExport;
        private System.Windows.Forms.SplitContainer spMain;
        private System.Windows.Forms.SplitContainer spDataView;
        private System.Windows.Forms.Label lblSetting;
        private System.Windows.Forms.Label lblFAI;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button btn_convert;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;
        private System.Windows.Forms.Button btn_export;
        private System.Windows.Forms.Label lblBending;
        private System.Windows.Forms.Label lblType3;
    }
}

